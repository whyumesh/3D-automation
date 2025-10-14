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
            print(f"To: {zbm_email}, CC: {cc_emails}")
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
        
        # Status counts
        request_cancelled_out_of_stock = abm_data[abm_data['Final Status'].isin(['Out of stock', 'On hold', 'Not permitted'])]['Assigned Request Ids'].nunique()
        action_pending_at_ho = abm_data[abm_data['Final Status'].isin(['Action pending / In Process'])]['Assigned Request Ids'].nunique()
        pending_for_invoicing = abm_data[abm_data['Final Status'].isin(['Dispatch Pending'])]['Assigned Request Ids'].nunique()
        pending_for_dispatch = abm_data[abm_data['Final Status'].isin(['Dispatch Pending'])]['Assigned Request Ids'].nunique()
        delivered = abm_data[abm_data['Final Status'].isin(['Delivered'])]['Assigned Request Ids'].nunique()
        dispatched_in_transit = abm_data[abm_data['Final Status'].isin(['Dispatched & In Transit'])]['Assigned Request Ids'].nunique()
        rto = abm_data[abm_data['Final Status'].isin(['RTO'])]['Assigned Request Ids'].nunique()
        
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
            'RTO': rto
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
    """Create HTML table for summary data"""
    
    if summary_df.empty:
        return "<p>No data available</p>"
    
    # Create HTML table
    html = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse; width: 100%;'>"
    
    # Header row
    html += "<tr style='background-color: #f0f0f0; font-weight: bold;'>"
    html += "<th>Area Name</th>"
    html += "<th>ABM Name</th>"
    html += "<th># Unique TBMs</th>"
    html += "<th># Unique HCPs</th>"
    html += "<th># Requests Raised</th>"
    html += "<th>Request Cancelled Out of Stock</th>"
    html += "<th>Action Pending at HO</th>"
    html += "<th>Sent to HUB</th>"
    html += "<th>Pending for Invoicing</th>"
    html += "<th>Pending for Dispatch</th>"
    html += "<th>Requests Dispatched</th>"
    html += "<th>Delivered</th>"
    html += "<th>Dispatched In Transit</th>"
    html += "<th>RTO</th>"
    html += "</tr>"
    
    # Data rows
    for _, row in summary_df.iterrows():
        html += "<tr>"
        html += f"<td>{row['Area Name']}</td>"
        html += f"<td>{row['ABM Name']}</td>"
        html += f"<td>{row['Unique TBMs']}</td>"
        html += f"<td>{row['Unique HCPs']}</td>"
        html += f"<td>{row['Requests Raised']}</td>"
        html += f"<td>{row['Request Cancelled Out of Stock']}</td>"
        html += f"<td>{row['Action Pending at HO']}</td>"
        html += f"<td>{row['Sent to HUB']}</td>"
        html += f"<td>{row['Pending for Invoicing']}</td>"
        html += f"<td>{row['Pending for Dispatch']}</td>"
        html += f"<td>{row['Requests Dispatched']}</td>"
        html += f"<td>{row['Delivered']}</td>"
        html += f"<td>{row['Dispatched In Transit']}</td>"
        html += f"<td>{row['RTO']}</td>"
        html += "</tr>"
    
    # Total row
    html += "<tr style='background-color: #e0e0e0; font-weight: bold;'>"
    html += "<td>TOTAL</td>"
    html += "<td></td>"
    html += f"<td>{summary_df['Unique TBMs'].sum()}</td>"
    html += f"<td>{summary_df['Unique HCPs'].sum()}</td>"
    html += f"<td>{summary_df['Requests Raised'].sum()}</td>"
    html += f"<td>{summary_df['Request Cancelled Out of Stock'].sum()}</td>"
    html += f"<td>{summary_df['Action Pending at HO'].sum()}</td>"
    html += f"<td>{summary_df['Sent to HUB'].sum()}</td>"
    html += f"<td>{summary_df['Pending for Invoicing'].sum()}</td>"
    html += f"<td>{summary_df['Pending for Dispatch'].sum()}</td>"
    html += f"<td>{summary_df['Requests Dispatched'].sum()}</td>"
    html += f"<td>{summary_df['Delivered'].sum()}</td>"
    html += f"<td>{summary_df['Dispatched In Transit'].sum()}</td>"
    html += f"<td>{summary_df['RTO'].sum()}</td>"
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

import os

def find_consolidated_file(zbm_code, zbm_name):
    # Assuming files are stored in a folder named 'consolidated_files' in the root directory
    root_dir = os.getcwd()  # Gets the current working directory
    folder_name = "consolidated_files"
    
    # Construct filename (adjust this logic based on your actual naming convention)
    filename = f"{zbm_code}_{zbm_name}_consolidated.xlsx"
    
    # Build full path
    full_path = os.path.join(root_dir, folder_name, filename)
    
    # Return absolute path
    return os.path.abspath(full_path)

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
