#!/usr/bin/env python3
"""
ZBM Email Display
Displays professional emails in Outlook for ZBMs with precise data matching
USES OUTLOOK TO DISPLAY EMAILS (NOT SEND AUTOMATICALLY)
"""

import pandas as pd
import os
from datetime import datetime
import warnings
import win32com.client
from openpyxl import load_workbook

# Suppress pandas warnings
warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')

def send_zbm_emails():
    """Display emails in Outlook for review without sending"""
    
    print("🚀 Starting ZBM Email Display...")
    print("📧 This will DISPLAY emails in Outlook for review - NOT SEND automatically")
    
    # Read Sample Master Tracker data
    print("📖 Reading Sample Master Tracker.xlsx...")
    try:
        df = pd.read_excel('Sample Master Tracker.xlsx')
        print(f"✅ Successfully loaded {len(df)} records from Sample Master Tracker.xlsx")
    except Exception as e:
        print(f"❌ Error reading Sample Master Tracker.xlsx: {e}")
        return
    
    # Required columns
    required_columns = [
        'ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID', 'ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID',
        'Assigned Request Ids', 'Doctor: Customer Code', 'Request Status', 'TBM EMAIL_ID', 'TBM HQ'
    ]
    
    # Check for missing columns
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        print(f"❌ Missing required columns: {missing}")
        return
    
    # Clean and filter data
    df = df.dropna(subset=['ZBM Terr Code', 'ZBM Name', 'ABM Terr Code', 'ABM Name'])
    df = df[df['ZBM Terr Code'].astype(str).str.strip() != '']
    df = df[df['ABM Terr Code'].astype(str).str.strip() != '']
    df = df[df['ZBM Terr Code'].astype(str).str.startswith('ZN')]
    
    print(f"📊 After cleaning: {len(df)} records remaining")
    
    # Compute Final Status using logic.xlsx
    print("🧠 Computing final status...")
    try:
        xls_rules = pd.ExcelFile('logic.xlsx')
        
        # Check available sheet names
        sheet_names = xls_rules.sheet_names
        print(f"   📋 Available sheets in logic.xlsx: {sheet_names}")
        
        # Try to find the rules sheet (case-insensitive)
        rules_sheet = None
        for sheet in sheet_names:
            if 'rule' in sheet.lower():
                rules_sheet = sheet
                break
        
        if rules_sheet:
            print(f"   📖 Using sheet: {rules_sheet}")
            rules_df = pd.read_excel(xls_rules, rules_sheet)
        else:
            # Use the first sheet if no rules sheet found
            rules_sheet = sheet_names[0]
            print(f"   📖 Using first sheet: {rules_sheet}")
            rules_df = pd.read_excel(xls_rules, rules_sheet)
        
        # Check if required columns exist
        required_rule_columns = ['Request Status', 'Final Answer']
        missing_rule_columns = [col for col in required_rule_columns if col not in rules_df.columns]
        
        if missing_rule_columns:
            print(f"   ⚠️ Missing columns in rules sheet: {missing_rule_columns}")
            print(f"   📋 Available columns: {list(rules_df.columns)}")
            # Use alternative column names if available
            status_col = None
            answer_col = None
            
            for col in rules_df.columns:
                if 'request' in col.lower() and 'status' in col.lower():
                    status_col = col
                if 'final' in col.lower() and 'answer' in col.lower():
                    answer_col = col
            
            if status_col and answer_col:
                print(f"   🔄 Using alternative columns: {status_col} -> {answer_col}")
                status_mapping = {}
                for _, row in rules_df.iterrows():
                    if pd.notna(row[status_col]) and pd.notna(row[answer_col]):
                        status_mapping[row[status_col]] = row[answer_col]
            else:
                raise Exception("Cannot find suitable columns for status mapping")
        else:
            status_mapping = {}
            for _, row in rules_df.iterrows():
                if pd.notna(row['Request Status']) and pd.notna(row['Final Answer']):
                    status_mapping[row['Request Status']] = row['Final Answer']
        
        df['Final Status'] = df['Request Status'].map(status_mapping)
        df['Final Status'] = df['Final Status'].fillna(df['Request Status'])
        print("✅ Final status computed successfully")
        
    except Exception as e:
        print(f"❌ Error computing final status: {e}")
        print("   🔄 Using Request Status as Final Status")
        df['Final Status'] = df['Request Status']
    
    # Get unique ZBMs
    zbms = df[['ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID']].drop_duplicates().sort_values('ZBM Terr Code')
    print(f"📋 Found {len(zbms)} unique ZBMs")
    
    # Initialize Outlook with robust error handling
    print("📧 Initializing Outlook...")
    outlook = None
    
    # Try different Outlook initialization methods
    outlook_methods = [
        "Outlook.Application",
        "Outlook.Application.16",  # Office 2016/2019/365
        "Outlook.Application.15",  # Office 2013
        "Outlook.Application.14",  # Office 2010
        "Outlook.Application.12",  # Office 2007
    ]
    
    for method in outlook_methods:
        try:
            print(f"   🔄 Trying: {method}")
            outlook = win32com.client.Dispatch(method)
            print(f"✅ Outlook initialized successfully using: {method}")
            break
        except Exception as e:
            print(f"   ❌ Failed with {method}: {e}")
            continue
    
    if outlook is None:
        print("❌ Could not initialize Outlook with any method")
        print("🔧 Troubleshooting steps:")
        print("   1. Ensure Outlook is installed on this computer")
        print("   2. Try opening Outlook manually first")
        print("   3. Check if Outlook is running in the background")
        print("   4. Try running as administrator")
        print("   5. Install Microsoft Office/Outlook if not present")
        
        # Fallback: Create HTML email files
        print("\n🔄 Creating HTML email files as fallback...")
        create_html_email_files(df, zbms)
        return
    
    # Process each ZBM
    success_count = 0
    error_count = 0
    
    for _, zbm_row in zbms.iterrows():
        zbm_code = zbm_row['ZBM Terr Code']
        zbm_name = zbm_row['ZBM Name']
        zbm_email = zbm_row['ZBM EMAIL_ID']
        
        print(f"\n🔄 Processing ZBM: {zbm_code} - {zbm_name}")
        
        try:
            # Filter data for this specific ZBM ONLY
            zbm_data = df[df['ZBM Terr Code'] == zbm_code]
            
            if len(zbm_data) == 0:
                print(f"⚠️ No data found for ZBM: {zbm_code}")
                continue
            
            # Get unique ABMs under this ZBM
            abms = zbm_data.groupby(['ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID']).agg({
                'TBM HQ': 'first'
            }).reset_index()
            
            # Create summary data
            summary_data = create_summary_data(zbm_data, abms)
            summary_df = pd.DataFrame(summary_data)
            
            # Generate email content
            email_content, cc_emails = generate_email_content(zbm_name, zbm_email, abms, summary_df)
            
            # Display email in Outlook (without sending)
            display_single_email(outlook, zbm_email, cc_emails, email_content, zbm_code, zbm_name)
            
            success_count += 1
            print(f"   ✅ Email displayed in Outlook for {zbm_name}")
            
        except Exception as e:
            error_count += 1
            print(f"   ❌ Error displaying email for {zbm_name}: {e}")
            continue
    
    print(f"\n🎉 Email display completed!")
    print(f"✅ Successfully displayed: {success_count} emails")
    print(f"❌ Failed to display: {error_count} emails")
    print(f"\n📧 All emails are now open in Outlook for your review and manual sending")

def create_summary_data(zbm_data, abms):
    """Create summary data for email body"""
    
    summary_data = []
    
    for _, abm_row in abms.iterrows():
        abm_code = abm_row['ABM Terr Code']
        abm_name = abm_row['ABM Name']
        tbm_hq = abm_row['TBM HQ']
        
        # Filter data for this specific ABM under this ZBM
        abm_data = zbm_data[(zbm_data['ABM Terr Code'] == abm_code) & (zbm_data['ABM Name'] == abm_name)]
        
        # Calculate metrics
        unique_tbms = abm_data['TBM EMAIL_ID'].nunique() if 'TBM EMAIL_ID' in abm_data.columns else 0
        unique_hcps = abm_data['Doctor: Customer Code'].nunique()
        unique_requests = abm_data['Assigned Request Ids'].nunique()
        
        # Status counts using Request Status field (matching summary report logic)
        request_cancelled_out_of_stock = abm_data[abm_data['Request Status'].isin(['Out of stock', 'On hold', 'Not permitted'])]['Assigned Request Ids'].nunique()
        action_pending_at_ho = abm_data[abm_data['Request Status'].isin(['Request Raised'])]['Assigned Request Ids'].nunique()
        pending_for_invoicing = abm_data[abm_data['Request Status'].isin(['Action pending / In Process'])]['Assigned Request Ids'].nunique()
        pending_for_dispatch = abm_data[abm_data['Request Status'].isin(['Dispatch  Pending'])]['Assigned Request Ids'].nunique()
        delivered = abm_data[abm_data['Request Status'].isin(['Delivered'])]['Assigned Request Ids'].nunique()
        dispatched_in_transit = abm_data[abm_data['Request Status'].isin(['Dispatched & In Transit'])]['Assigned Request Ids'].nunique()
        rto = abm_data[abm_data['Request Status'].isin(['RTO'])]['Assigned Request Ids'].nunique()
        
        # RTO Reasons using string contains (matching summary report logic)
        incomplete_address = abm_data[abm_data['Rto Reason'].str.contains('Incomplete Address', na=False)]['Assigned Request Ids'].nunique()
        doctor_non_contactable = abm_data[abm_data['Rto Reason'].str.contains('Dr. Non contactable', na=False)]['Assigned Request Ids'].nunique()
        doctor_refused_to_accept = abm_data[abm_data['Rto Reason'].str.contains('Doctor Refused to Accept', na=False)]['Assigned Request Ids'].nunique()
        
        # Calculated fields
        requests_dispatched = delivered + dispatched_in_transit + rto
        sent_to_hub = pending_for_invoicing + pending_for_dispatch + requests_dispatched
        requests_raised = request_cancelled_out_of_stock + action_pending_at_ho + sent_to_hub
        
        # Create Area Name
        area_name = f"{abm_code} and {tbm_hq}"
        
        summary_data.append({
            'Area Name': area_name,
            'ABM Name': abm_name,
            'Unique TBMs': unique_tbms,
            'Unique HCPs': unique_hcps,
            'Unique Requests': unique_requests,
            'Requests Raised': requests_raised,
            'Request Cancelled Out of Stock': request_cancelled_out_of_stock,
            'Action Pending at HO': action_pending_at_ho,
            'Sent to HUB': sent_to_hub,
            'Pending for Invoicing': pending_for_invoicing,
            'Pending for Dispatch': pending_for_dispatch,
            'Requests Dispatched': requests_dispatched,
            'Delivered': delivered,
            'Dispatched In Transit': dispatched_in_transit,
            'RTO': rto,
            'Incomplete Address': incomplete_address,
            'Doctor Non Contactable': doctor_non_contactable,
            'Doctor Refused to Accept': doctor_refused_to_accept,
            'Hold Delivery': 0
        })
    
    return summary_data

def generate_email_content(zbm_name, zbm_email, abms, summary_df):
    """Generate professional email content"""
    
    current_date = datetime.now().strftime('%B %d, %Y')
    
    # Get ABM emails for CC
    abm_emails = abms['ABM EMAIL_ID'].dropna().unique().tolist()
    cc_emails = ', '.join(abm_emails)
    
    # Create summary table HTML
    table_html = create_summary_table_html(summary_df)
    
    email_content = f"""
Hi {zbm_name},

Please refer the status Sample requests raised in Abbworld for your area.

{table_html}

You can track your sample request at the following link with the Docket Number:

DTDC: <a href="https://www.dtdc.com/tracking">Click here</a>

Speed Post: <a href="https://www.indiapost.gov.in/vas/Pages/IndiaPostHome.aspx">Click Here</a>

In case of any query, please contact 1Point.

Regards,
Umesh Pawar.
"""
    
    return email_content, cc_emails

def create_summary_table_html(summary_df):
    """Create HTML table for summary data with all columns from summary report"""
    
    if summary_df.empty:
        return "<p>No data available</p>"
    
    # Create HTML table with complete column structure matching the summary report
    html = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse; width: 100%; font-size: 12px;'>"
    
    # Header row with all columns from summary report
    html += "<tr style='background-color: #f0f0f0; font-weight: bold;'>"
    html += "<th>Area Name</th>"
    html += "<th>ABM Name</th>"
    html += "<th># Unique TBMs</th>"
    html += "<th># Unique HCPs</th>"
    html += "<th># Requests Raised<br/>(A+B+C)</th>"
    html += "<th>Request Cancelled /<br/>Out of Stock (A)</th>"
    html += "<th>Action pending /<br/>In Process At HO (B)</th>"
    html += "<th>Sent to HUB (C)<br/>(D+E+F)</th>"
    html += "<th>Pending for<br/>Invoicing (D)</th>"
    html += "<th>Pending for<br/>Dispatch (E)</th>"
    html += "<th># Requests Dispatched (F)<br/>(G+H+I)</th>"
    html += "<th>Delivered (G)</th>"
    html += "<th>Dispatched &<br/>In Transit (H)</th>"
    html += "<th>RTO (I)</th>"
    html += "<th>Incomplete Address</th>"
    html += "<th>Doctor Non Contactable</th>"
    html += "<th>Doctor Refused to Accept</th>"
    html += "<th>Hold Delivery</th>"
    html += "</tr>"
    
    # Data rows
    for _, row in summary_df.iterrows():
        html += "<tr>"
        html += f"<td>{row.get('Area Name', '')}</td>"
        html += f"<td>{row.get('ABM Name', '')}</td>"
        html += f"<td>{row.get('Unique TBMs', 0)}</td>"
        html += f"<td>{row.get('Unique HCPs', 0)}</td>"
        html += f"<td>{row.get('Requests Raised', 0)}</td>"
        html += f"<td>{row.get('Request Cancelled Out of Stock', 0)}</td>"
        html += f"<td>{row.get('Action Pending at HO', 0)}</td>"
        html += f"<td>{row.get('Sent to HUB', 0)}</td>"
        html += f"<td>{row.get('Pending for Invoicing', 0)}</td>"
        html += f"<td>{row.get('Pending for Dispatch', 0)}</td>"
        html += f"<td>{row.get('Requests Dispatched', 0)}</td>"
        html += f"<td>{row.get('Delivered', 0)}</td>"
        html += f"<td>{row.get('Dispatched In Transit', 0)}</td>"
        html += f"<td>{row.get('RTO', 0)}</td>"
        html += f"<td>{row.get('Incomplete Address', 0)}</td>"
        html += f"<td>{row.get('Doctor Non Contactable', 0)}</td>"
        html += f"<td>{row.get('Doctor Refused to Accept', 0)}</td>"
        html += f"<td>{row.get('Hold Delivery', 0)}</td>"
        html += "</tr>"
    
    # Total row
    html += "<tr style='background-color: #e0e0e0; font-weight: bold;'>"
    html += "<td>TOTAL</td>"
    html += "<td></td>"
    html += f"<td>{summary_df.get('Unique TBMs', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Unique HCPs', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Requests Raised', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Request Cancelled Out of Stock', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Action Pending at HO', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Sent to HUB', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Pending for Invoicing', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Pending for Dispatch', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Requests Dispatched', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Delivered', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Dispatched In Transit', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('RTO', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Incomplete Address', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Doctor Non Contactable', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Doctor Refused to Accept', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Hold Delivery', pd.Series()).sum()}</td>"
    html += "</tr>"
    
    html += "</table>"
    
    return html

def display_single_email(outlook, zbm_email, cc_emails, email_content, zbm_code, zbm_name):
    """Display a single email in Outlook for review (without sending)"""
    
    # Create new mail item
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    
    # Set recipients
    mail.To = zbm_email
    
    # Set CC recipients
    if cc_emails:
        mail.CC = cc_emails
    
    # Set subject
    current_date = datetime.now().strftime('%B %d, %Y')
    mail.Subject = f"Sample Direct Dispatch to Doctors - Request Status as of {current_date}"
    
    # Set body
    mail.HTMLBody = email_content
    
    # Add attachment (consolidated file)
    consolidated_file = find_consolidated_file(zbm_code, zbm_name)
    if consolidated_file and os.path.exists(consolidated_file):
        mail.Attachments.Add(consolidated_file)
        print(f"   📎 Attached: {os.path.basename(consolidated_file)}")
    
    # Display email (don't send)
    mail.Display()
    
    print(f"   📧 Email displayed for: {zbm_email}")
    if cc_emails:
        print(f"   📧 CC'd to: {cc_emails}")
    print(f"   ⚠️  Review the email and send manually from Outlook")

def find_consolidated_file(zbm_code, zbm_name):
    """Find the consolidated file for this ZBM"""
    
    # Look for consolidated files in current directory and subdirectories
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.startswith(f"ZBM_Consolidated_{zbm_code}_") and file.endswith('.xlsx'):
                return os.path.join(root, file)
    
    return None

def create_html_email_files(df, zbms):
    """Create HTML email files as fallback when Outlook is not available"""
    
    print("📧 Creating HTML email files...")
    
    # Create output directory
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_dir = f"ZBM_HTML_Emails_{timestamp}"
    os.makedirs(output_dir, exist_ok=True)
    print(f"📁 Created output directory: {output_dir}")
    
    success_count = 0
    
    for _, zbm_row in zbms.iterrows():
        zbm_code = zbm_row['ZBM Terr Code']
        zbm_name = zbm_row['ZBM Name']
        zbm_email = zbm_row['ZBM EMAIL_ID']
        
        print(f"\n🔄 Processing ZBM: {zbm_code} - {zbm_name}")
        
        try:
            # Filter data for this specific ZBM ONLY
            zbm_data = df[df['ZBM Terr Code'] == zbm_code]
            
            if len(zbm_data) == 0:
                print(f"⚠️ No data found for ZBM: {zbm_code}")
                continue
            
            # Get unique ABMs under this ZBM
            abms = zbm_data.groupby(['ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID']).agg({
                'TBM HQ': 'first'
            }).reset_index()
            
            # Create summary data
            summary_data = create_summary_data(zbm_data, abms)
            summary_df = pd.DataFrame(summary_data)
            
            # Generate email content
            email_content, cc_emails = generate_email_content(zbm_name, zbm_email, abms, summary_df)
            
            # Create HTML email file
            create_single_html_email(zbm_code, zbm_name, zbm_email, cc_emails, email_content, output_dir)
            
            success_count += 1
            print(f"   ✅ HTML email created for {zbm_name}")
            
        except Exception as e:
            print(f"   ❌ Error creating HTML email for {zbm_name}: {e}")
            continue
    
    print(f"\n🎉 HTML email creation completed!")
    print(f"✅ Successfully created: {success_count} HTML email files")
    print(f"📁 Files saved in: {output_dir}")
    print(f"📧 You can open these HTML files in your browser and copy content to Outlook")

def create_single_html_email(zbm_code, zbm_name, zbm_email, cc_emails, email_content, output_dir):
    """Create a single HTML email file"""
    
    current_date = datetime.now().strftime('%B %d, %Y')
    
    # Create full HTML email
    html_email = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Sample Direct Dispatch to Doctors - Request Status as of {current_date}</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        .total-row {{ background-color: #e0e0e0; font-weight: bold; }}
        .header {{ background-color: #f0f0f0; padding: 10px; margin-bottom: 20px; }}
    </style>
</head>
<body>
    <div class="header">
        <h3>Email Details:</h3>
        <p><strong>To:</strong> {zbm_email}</p>
        <p><strong>CC:</strong> {cc_emails}</p>
        <p><strong>Subject:</strong> Sample Direct Dispatch to Doctors - Request Status as of {current_date}</p>
    </div>
    
    <div class="email-content">
        <p>Hi {zbm_name},</p>
        
        <p>Please refer the status Sample requests raised in Abbworld for your area.</p>
        
        {email_content}
        
        <p>You can track your sample request at the following link with the Docket Number:</p>
        <p>DTDC: <a href="https://www.dtdc.com/tracking">Click here</a></p>
        <p>Speed Post: <a href="https://www.indiapost.gov.in/vas/Pages/IndiaPostHome.aspx">Click Here</a></p>
        
        <p>In case of any query, please contact 1Point.</p>
        
        <p>Regards,<br>Umesh Pawar.</p>
    </div>
</body>
</html>
"""
    
    # Save HTML file
    safe_zbm_name = str(zbm_name).replace(' ', '_').replace('/', '_').replace('\\', '_')
    filename = f"Email_{zbm_code}_{safe_zbm_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    filepath = os.path.join(output_dir, filename)
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(html_email)
    
    print(f"   📧 HTML email saved: {filename}")

if __name__ == "__main__":
    send_zbm_emails()
