def display_single_email(outlook, zbm_email, cc_emails, email_content, zbm_code, zbm_name):
    """Display a single email in Outlook for review (without sending)"""
    
    try:
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
        
        # Find and attach ZBM Summary file
        zbm_summary_file = find_latest_zbm_summary_file(zbm_code, zbm_name)
        if zbm_summary_file and os.path.exists(zbm_summary_file):
            try:
                mail.Attachments.Add(zbm_summary_file)
                print(f"   📎 Attached ZBM Summary: {os.path.basename(zbm_summary_file)}")
            except Exception as e:
                print(f"   ⚠️  Failed to attach ZBM Summary: {e}")
        else:
            print(f"   ⚠️  ZBM Summary file not found for {zbm_code}")
        
        # Find and attach Consolidated file
        consolidated_file = find_latest_consolidated_file(zbm_code, zbm_name)
        if consolidated_file and os.path.exists(consolidated_file):
            try:
                mail.Attachments.Add(consolidated_file)
                print(f"   📎 Attached Consolidated: {os.path.basename(consolidated_file)}")
            except Exception as e:
                print(f"   ⚠️  Failed to attach Consolidated file: {e}")
        else:
            print(f"   ⚠️  Consolidated file not found for {zbm_code}")
        
        # Display email (don't send)
        mail.Display()
        
        print(f"   📧 Email displayed for: {zbm_email}")
        if cc_emails:
            print(f"   📧 CC'd to: {cc_emails}")
        print(f"   ⚠️  Review the email and send manually from Outlook")
        
    except Exception as e:
        print(f"   ❌ Error displaying email: {e}")
        import traceback
        traceback.print_exc()
        raise

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
            
            # Find attachments
            zbm_summary_file = find_latest_zbm_summary_file(zbm_code, zbm_name)
            consolidated_file = find_latest_consolidated_file(zbm_code, zbm_name)
            
            # Create HTML email file
            create_single_html_email(zbm_code, zbm_name, zbm_email, cc_emails, email_content, 
                                    zbm_summary_file, consolidated_file, output_dir)
            
            success_count += 1
            print(f"   ✅ HTML email created for {zbm_name}")
            
        except Exception as e:
            print(f"   ❌ Error creating HTML email for {zbm_name}: {e}")
            continue
    
    print(f"\n🎉 HTML email creation completed!")
    print(f"✅ Successfully created: {success_count} HTML email files")
    print(f"📁 Files saved in: {output_dir}")
    print(f"📧 You can open these HTML files in your browser and copy content to Outlook")

def create_single_html_email(zbm_code, zbm_name, zbm_email, cc_emails, email_content, 
                             zbm_summary_file, consolidated_file, output_dir):
    """Create a single HTML email file"""
    
    current_date = datetime.now().strftime('%B %d, %Y')
    
    # Create attachment info
    attachments_info = "<p><strong>Attachments:</strong></p><ul>"
    if zbm_summary_file:
        attachments_info += f"<li>ZBM Summary: {os.path.basename(zbm_summary_file)}</li>"
    if consolidated_file:
        attachments_info += f"<li>Consolidated File: {os.path.basename(consolidated_file)}</li>"
    if not zbm_summary_file and not consolidated_file:
        attachments_info += "<li>No attachments found</li>"
    attachments_info += "</ul>"
    
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
        th {{ background-color: #4472C4; color: white; font-weight: bold; }}
        .total-row {{ background-color: #D9E1F2; font-weight: bold; }}
        .header {{ background-color: #f0f0f0; padding: 10px; margin-bottom: 20px; border: 1px solid #ddd; }}
        .attachments {{ background-color: #fff3cd; padding: 10px; margin: 20px 0; border: 1px solid #ffc107; }}
    </style>
</head>
<body>
    <div class="header">
        <h3>Email Details:</h3>
        <p><strong>To:</strong> {zbm_email}</p>
        <p><strong>CC:</strong> {cc_emails}</p>
        <p><strong>Subject:</strong> Sample Direct Dispatch to Doctors - Request Status as of {current_date}</p>
    </div>
    
    <div class="attachments">
        {attachments_info}
    </div>
    
    <div class="email-content">
        {email_content}
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
