#!/usr/bin/env python3
"""
ZBM Email Preview Generator
Creates professional email content for each ZBM with precise data matching
DISPLAYS EMAIL CONTENT - DOES NOT SEND
"""

import pandas as pd
import os
from datetime import datetime
import warnings
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from copy import copy as copy_style

# Suppress pandas warnings
warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')

def create_email_preview():
    """Create email preview for each ZBM with precise data matching"""
    
    print("🚀 Starting ZBM Email Preview Generation...")
    print("📧 This will DISPLAY email content - NOT SEND emails")
    
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
        rules_df = pd.read_excel(xls_rules, 'Rules')
        
        status_mapping = {}
        for _, row in rules_df.iterrows():
            if pd.notna(row['Request Status']) and pd.notna(row['Final Answer']):
                status_mapping[row['Request Status']] = row['Final Answer']
        
        df['Final Status'] = df['Request Status'].map(status_mapping)
        df['Final Status'] = df['Final Status'].fillna(df['Request Status'])
        print("✅ Final status computed successfully")
    except Exception as e:
        print(f"❌ Error computing final status: {e}")
        df['Final Status'] = df['Request Status']
    
    # Get unique ZBMs
    zbms = df[['ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID']].drop_duplicates().sort_values('ZBM Terr Code')
    print(f"📋 Found {len(zbms)} unique ZBMs")
    
    # Create output directory
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_dir = f"ZBM_Email_Previews_{timestamp}"
    os.makedirs(output_dir, exist_ok=True)
    print(f"📁 Created output directory: {output_dir}")
    
    # Process each ZBM
    for _, zbm_row in zbms.iterrows():
        zbm_code = zbm_row['ZBM Terr Code']
        zbm_name = zbm_row['ZBM Name']
        zbm_email = zbm_row['ZBM EMAIL_ID']
        
        print(f"\n🔄 Processing ZBM: {zbm_code} - {zbm_name}")
        
        # Filter data for this specific ZBM ONLY
        zbm_data = df[df['ZBM Terr Code'] == zbm_code]
        
        if len(zbm_data) == 0:
            print(f"⚠️ No data found for ZBM: {zbm_code}")
            continue
        
        print(f"   📊 Found {len(zbm_data)} records for this ZBM")
        
        # Get unique ABMs under this ZBM
        abms = zbm_data.groupby(['ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID']).agg({
            'TBM HQ': 'first'
        }).reset_index()
        
        print(f"   📋 Found {len(abms)} ABMs under this ZBM")
        
        # Create summary data for email body
        summary_data = []
        
        for _, abm_row in abms.iterrows():
            abm_code = abm_row['ABM Terr Code']
            abm_name = abm_row['ABM Name']
            abm_email = abm_row['ABM EMAIL_ID']
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
            
            # RTO reasons
            incomplete_address = 0
            doctor_non_contactable = 0
            doctor_refused_to_accept = 0
            hold_delivery = 0
            
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
                'Hold Delivery': hold_delivery
            })
        
        # Create DataFrame
        summary_df = pd.DataFrame(summary_data)
        
        # Generate email content
        email_content = generate_email_content(zbm_name, zbm_email, abms, summary_df)
        
        # Save email preview
        save_email_preview(zbm_code, zbm_name, email_content, summary_df, output_dir)
        
        print(f"   ✅ Email preview created for {zbm_name}")
    
    print(f"\n🎉 Successfully created {len(zbms)} email previews in directory: {output_dir}")
    print("📧 Review the email content before sending")

def generate_email_content(zbm_name, zbm_email, abms, summary_df):
    """Generate professional email content"""
    
    current_date = datetime.now().strftime('%B %d, %Y')
    
    # Get ABM emails for CC
    abm_emails = abms['ABM EMAIL_ID'].dropna().unique().tolist()
    cc_emails = ', '.join(abm_emails)
    
    # Create summary table HTML
    table_html = create_summary_table_html(summary_df)
    
    email_content = f"""
To: {zbm_email}
Cc: {cc_emails}

Subject: Sample Direct Dispatch to Doctors - Request Status as of {current_date}

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
    
    return email_content

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

def save_email_preview(zbm_code, zbm_name, email_content, summary_df, output_dir):
    """Save email preview to file"""
    
    # Create filename
    safe_zbm_name = str(zbm_name).replace(' ', '_').replace('/', '_').replace('\\', '_')
    filename = f"Email_Preview_{zbm_code}_{safe_zbm_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    filepath = os.path.join(output_dir, filename)
    
    # Save email content
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(email_content)
    
    # Also save summary data as Excel for reference
    excel_filename = f"Summary_Data_{zbm_code}_{safe_zbm_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    excel_filepath = os.path.join(output_dir, excel_filename)
    summary_df.to_excel(excel_filepath, index=False)
    
    print(f"   📧 Email preview saved: {filename}")
    print(f"   📊 Summary data saved: {excel_filename}")

if __name__ == "__main__":
    create_email_preview()