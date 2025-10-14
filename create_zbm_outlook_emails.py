import pandas as pd
import numpy as np
from datetime import datetime
import win32com.client as win32
import os

def create_zbm_outlook_emails():
    """
    Create Outlook email drafts for each ZBM with their report data in table format
    Opens Outlook windows but doesn't send - user can review and send manually
    """
    
    print("📧 Starting ZBM Outlook Email Creation...")
    
    # Check if Outlook is available
    try:
        outlook = win32.Dispatch('outlook.application')
        print("✅ Outlook application initialized")
    except Exception as e:
        print(f"❌ Error initializing Outlook: {e}")
        print("💡 Please ensure Outlook is installed and running")
        print("💡 You can also use the preview script: python create_zbm_email_preview.py")
        return
    
    # Read master tracker data
    print("📖 Reading master_tracker.csv...")
    try:
        df = pd.read_csv('master_tracker.csv', encoding='latin-1', low_memory=False)
        print(f"✅ Successfully loaded {len(df)} records from master_tracker.csv")
    except Exception as e:
        print(f"❌ Error reading master_tracker.csv: {e}")
        return
    
    # Clean and prepare data
    print("🧹 Cleaning and preparing data...")
    
    # Ensure required columns exist
    required_columns = ['ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID',
                        'ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID',
                        'TBM HQ', 'TBM EMAIL_ID',
                        'Doctor: Customer Code', 'Assigned Request Ids', 'Request Status', 'Rto Reason']
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        print(f"❌ Missing required columns in master_tracker.csv: {missing}")
        return

    # Remove rows where key fields are null or empty
    df = df.dropna(subset=['ZBM Terr Code', 'ZBM Name', 'ABM Terr Code', 'ABM Name', 'TBM HQ'])
    df = df[df['ZBM Terr Code'].astype(str).str.strip() != '']
    df = df[df['ABM Terr Code'].astype(str).str.strip() != '']
    df = df[df['TBM HQ'].astype(str).str.strip() != '']

    print(f"📊 After cleaning: {len(df)} records remaining")

    # Compute Final Answer per unique request id using corrected rules
    print("🧠 Computing final status per unique Request Id using corrected rules...")
    try:
        def normalize(text):
            normalized = str(text).strip().casefold()
            normalized = normalized.replace('  ', ' ')  # Fix spacing issues
            return normalized

        # Load rules from logic.xlsx
        xls_rules = pd.ExcelFile('logic.xlsx')
        sheet2 = pd.read_excel(xls_rules, 'Sheet2')

        rules = {}
        for _, row in sheet2.iterrows():
            statuses = [normalize(s) for s in row.drop('Final Answer').dropna().tolist()]
            statuses = tuple(sorted(set(statuses)))
            rules[statuses] = row['Final Answer']

        # Add missing rules identified from validation
        additional_rules = {
            ('action pending / in process', 'delivered'): 'Delivered',
            ('action pending / in process', 'dispatched & in transit'): 'Dispatched & In Transit',
            ('action pending / in process', 'dispatch pending'): 'Dispatch Pending',
            ('action pending / in process', 'out of stock'): 'Out of stock',
            ('action pending / in process', 'return'): 'Return',
            ('delivered', 'return'): 'Delivered',
            ('dispatch pending', 'delivered'): 'Delivered',
            ('dispatch pending', 'dispatched & in transit'): 'Dispatched & In Transit',
            ('dispatch pending', 'return'): 'Return',
            ('dispatched & in transit', 'return'): 'Dispatched & In Transit',
            ('out of stock', 'return'): 'Out of stock',
            ('request raised', 'action pending / in process'): 'Action pending / In Process',
            ('request raised', 'delivered'): 'Delivered',
            ('request raised', 'dispatch pending'): 'Dispatch Pending',
            ('request raised', 'dispatched & in transit'): 'Dispatched & In Transit',
            ('request raised', 'out of stock'): 'Out of stock',
            ('request raised', 'return'): 'Return',
        }
        
        rules.update(additional_rules)

        # Apply rules to compute Final Answer
        def compute_final_answer(request_id):
            request_data = df[df['Assigned Request Ids'] == request_id]
            if len(request_data) == 0:
                return 'Unknown'
            
            statuses = [normalize(s) for s in request_data['Request Status'].dropna().unique()]
            statuses = tuple(sorted(set(statuses)))
            
            if statuses in rules:
                return rules[statuses]
            else:
                # If no rule found, return the most common status
                return request_data['Request Status'].mode().iloc[0] if len(request_data['Request Status'].mode()) > 0 else 'Unknown'

        # Apply final answer computation
        unique_requests = df['Assigned Request Ids'].unique()
        print(f"🔍 Computing final answers for {len(unique_requests)} unique requests...")
        
        final_answers = {}
        for req_id in unique_requests:
            final_answers[req_id] = compute_final_answer(req_id)
        
        # Map final answers back to dataframe
        df['Final Answer'] = df['Assigned Request Ids'].map(final_answers)
        
        print("✅ Final Answer computation completed")
        
    except Exception as e:
        print(f"❌ Error computing final answers: {e}")
        return

    # Get unique ZBMs
    zbms = df[['ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID']].drop_duplicates().sort_values('ZBM Terr Code')
    print(f"📋 Found {len(zbms)} unique ZBMs")

    successful_emails = 0
    failed_emails = 0
    
    # Process each ZBM
    for i, (_, zbm_row) in enumerate(zbms.iterrows()):
        zbm_code = zbm_row['ZBM Terr Code']
        zbm_name = zbm_row['ZBM Name']
        zbm_email = zbm_row['ZBM EMAIL_ID']
        
        # Handle NaN names and emails
        if pd.isna(zbm_name):
            zbm_name = "Unknown"
        if pd.isna(zbm_email):
            print(f"⚠️ No email found for ZBM: {zbm_name} ({zbm_code})")
            failed_emails += 1
            continue
        
        print(f"📧 Creating email draft for ZBM {i+1}/{len(zbms)}: {zbm_name} ({zbm_code})")
        print(f"   📬 Email: {zbm_email}")
        
        try:
            # Filter data for this ZBM
            zbm_data = df[df['ZBM Terr Code'] == zbm_code].copy()
            
            if len(zbm_data) == 0:
                print(f"⚠️ No data found for ZBM: {zbm_code}")
                failed_emails += 1
                continue
            
            # Get unique ABMs under this ZBM
            abms = zbm_data[['ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID']].drop_duplicates().sort_values('ABM Terr Code')
            
            # Create summary data for email table
            summary_data = []
            
            for _, abm_row in abms.iterrows():
                abm_code = abm_row['ABM Terr Code']
                abm_name = abm_row['ABM Name']
                abm_email = abm_row['ABM EMAIL_ID']
                
                abm_data = zbm_data[zbm_data['ABM Terr Code'] == abm_code]
                
                # Calculate metrics for this ABM
                unique_tbms = abm_data['TBM EMAIL_ID'].nunique()
                unique_hcps = abm_data['Doctor: Customer Code'].nunique()
                
                # HO Section (A + B)
                request_cancelled_out_of_stock = abm_data[abm_data['Final Answer'].isin(['Out of stock', 'On hold', 'Not permitted'])]['Assigned Request Ids'].nunique()
                action_pending_at_ho = abm_data[abm_data['Final Answer'].isin(['Action pending / In Process'])]['Assigned Request Ids'].nunique()
                
                # HUB Section (D + E + F)
                pending_for_invoicing = 0  # Placeholder
                pending_for_dispatch = abm_data[abm_data['Final Answer'].isin(['Dispatch Pending'])]['Assigned Request Ids'].nunique()
                
                # Delivery Status (G + H + I)
                delivered = abm_data[abm_data['Final Answer'].isin(['Delivered'])]['Assigned Request Ids'].nunique()
                dispatched_in_transit = abm_data[abm_data['Final Answer'].isin(['Dispatched & In Transit'])]['Assigned Request Ids'].nunique()
                rto = abm_data[abm_data['Rto Reason'].notna()]['Assigned Request Ids'].nunique()
                
                # Calculated fields
                requests_dispatched = delivered + dispatched_in_transit + rto
                sent_to_hub = pending_for_invoicing + pending_for_dispatch + requests_dispatched
                requests_raised = request_cancelled_out_of_stock + action_pending_at_ho + sent_to_hub
                
                # RTO Reasons
                rto_data = abm_data[abm_data['Rto Reason'].notna()]
                incomplete_address = rto_data[rto_data['Rto Reason'].str.contains('Incomplete Address', case=False, na=False)]['Assigned Request Ids'].nunique()
                doctor_non_contactable = rto_data[rto_data['Rto Reason'].str.contains('Non contactable', case=False, na=False)]['Assigned Request Ids'].nunique()
                doctor_refused_to_accept = rto_data[rto_data['Rto Reason'].str.contains('refused to accept', case=False, na=False)]['Assigned Request Ids'].nunique()
                hold_delivery = rto_data[rto_data['Rto Reason'].str.contains('Hold Delivery', case=False, na=False)]['Assigned Request Ids'].nunique()
                
                # Create Area Name
                area_name = f"{abm_code} - {abm_name}"
                
                summary_row = {
                    'Area Name': area_name,
                    'ABM Name': abm_name,
                    'Unique TBMs': unique_tbms,
                    'Unique HCPs': unique_hcps,
                    'Requests Raised': requests_raised,
                    'Cancelled/Out of Stock': request_cancelled_out_of_stock,
                    'Action Pending at HO': action_pending_at_ho,
                    'Sent to HUB': sent_to_hub,
                    'Pending for Invoicing': pending_for_invoicing,
                    'Pending for Dispatch': pending_for_dispatch,
                    'Requests Dispatched': requests_dispatched,
                    'Delivered': delivered,
                    'Dispatched & In Transit': dispatched_in_transit,
                    'RTO': rto,
                    'Incomplete Address': incomplete_address,
                    'Doctor Non Contactable': doctor_non_contactable,
                    'Doctor Refused to Accept': doctor_refused_to_accept,
                    'Hold Delivery': hold_delivery
                }
                
                summary_data.append(summary_row)
            
            # Create email content
            current_date = datetime.now().strftime("%B %d, %Y")
            
            # Email subject
            subject = f"ZBM Summary Report - {zbm_name} ({zbm_code}) - {current_date}"
            
            # Create HTML table for email body
            html_body = f"""
            <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; }}
                    h2 {{ color: #2E86AB; }}
                    table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
                    th, td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
                    th {{ background-color: #D9E1F2; font-weight: bold; }}
                    .section-header {{ background-color: #4472C4; color: white; font-weight: bold; }}
                    .total-row {{ background-color: #E7E6E6; font-weight: bold; }}
                    .summary {{ background-color: #F2F2F2; padding: 15px; margin-bottom: 20px; border-left: 4px solid #2E86AB; }}
                </style>
            </head>
            <body>
                <h2>ZBM Summary Report</h2>
                
                <div class="summary">
                    <p><strong>ZBM:</strong> {zbm_name} ({zbm_code})</p>
                    <p><strong>Report Date:</strong> {current_date}</p>
                    <p><strong>Total ABMs:</strong> {len(summary_data)}</p>
                    <p><strong>Total Requests:</strong> {sum(d['Requests Raised'] for d in summary_data)}</p>
                    <p><strong>Total RTO:</strong> {sum(d['RTO'] for d in summary_data)}</p>
                </div>
                
                <table>
                    <tr class="section-header">
                        <th rowspan="2">Area Name</th>
                        <th rowspan="2">ABM Name</th>
                        <th rowspan="2"># Unique TBMs</th>
                        <th rowspan="2"># Unique HCPs</th>
                        <th rowspan="2"># Requests Raised<br/>(A+B+C)</th>
                        <th colspan="2">HO</th>
                        <th colspan="3">HUB</th>
                        <th colspan="3">Delivery Status</th>
                        <th colspan="4">RTO Reasons</th>
                    </tr>
                    <tr class="section-header">
                        <th>Request Cancelled<br/>Out of Stock (A)</th>
                        <th>Action pending<br/>At HO (B)</th>
                        <th>Sent to HUB<br/>(C)</th>
                        <th>Pending for<br/>Invoicing (D)</th>
                        <th>Pending for<br/>Dispatch (E)</th>
                        <th># Requests<br/>Dispatched (F)</th>
                        <th>Delivered (G)</th>
                        <th>Dispatched &<br/>In Transit (H)</th>
                        <th>RTO (I)</th>
                        <th>Incomplete<br/>Address</th>
                        <th>Doctor Non<br/>Contactable</th>
                        <th>Doctor Refused<br/>to Accept</th>
                        <th>Hold<br/>Delivery</th>
                    </tr>
            """
            
            # Add data rows
            for data in summary_data:
                html_body += f"""
                    <tr>
                        <td>{data['Area Name']}</td>
                        <td>{data['ABM Name']}</td>
                        <td>{data['Unique TBMs']}</td>
                        <td>{data['Unique HCPs']}</td>
                        <td>{data['Requests Raised']}</td>
                        <td>{data['Cancelled/Out of Stock']}</td>
                        <td>{data['Action Pending at HO']}</td>
                        <td>{data['Sent to HUB']}</td>
                        <td>{data['Pending for Invoicing']}</td>
                        <td>{data['Pending for Dispatch']}</td>
                        <td>{data['Requests Dispatched']}</td>
                        <td>{data['Delivered']}</td>
                        <td>{data['Dispatched & In Transit']}</td>
                        <td>{data['RTO']}</td>
                        <td>{data['Incomplete Address']}</td>
                        <td>{data['Doctor Non Contactable']}</td>
                        <td>{data['Doctor Refused to Accept']}</td>
                        <td>{data['Hold Delivery']}</td>
                    </tr>
                """
            
            # Add totals row
            if summary_data:
                html_body += f"""
                    <tr class="total-row">
                        <td><strong>TOTAL</strong></td>
                        <td></td>
                        <td><strong>{sum(d['Unique TBMs'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Unique HCPs'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Requests Raised'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Cancelled/Out of Stock'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Action Pending at HO'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Sent to HUB'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Pending for Invoicing'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Pending for Dispatch'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Requests Dispatched'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Delivered'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Dispatched & In Transit'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['RTO'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Incomplete Address'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Doctor Non Contactable'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Doctor Refused to Accept'] for d in summary_data)}</strong></td>
                        <td><strong>{sum(d['Hold Delivery'] for d in summary_data)}</strong></td>
                    </tr>
                """
            
            html_body += """
                </table>
                
                <p style="margin-top: 30px; font-size: 12px; color: #666;">
                    This report is generated automatically. Please review the data before taking any action.
                </p>
            </body>
            </html>
            """
            
            # Create Outlook email
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = zbm_email
            mail.Subject = subject
            mail.HTMLBody = html_body
            
            # Display the email (opens window but doesn't send)
            mail.Display()
            
            print(f"✅ Email draft created for {zbm_name}")
            print(f"   📊 ABMs: {len(summary_data)}, Total RTO: {sum(d['RTO'] for d in summary_data)}")
            
            successful_emails += 1
            
        except Exception as e:
            print(f"❌ Error creating email for ZBM {zbm_code}: {e}")
            failed_emails += 1
            continue
    
    print(f"\n🎉 Outlook Email Draft Creation Completed!")
    print(f"✅ Successful: {successful_emails}")
    print(f"❌ Failed: {failed_emails}")
    print(f"📧 {successful_emails} Outlook windows opened with email drafts")
    print(f"💡 Review each email and send manually as needed")
    print(f"📋 Each ZBM gets their specific data - no random or bulk emails")

if __name__ == "__main__":
    create_zbm_outlook_emails()
