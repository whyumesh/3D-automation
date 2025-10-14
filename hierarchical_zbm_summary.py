import pandas as pd
import numpy as np
from datetime import datetime

def create_hierarchical_zbm_summary():
    """
    Create hierarchical ZBM → ABM → TBM summary report from master_tracker.csv
    """
    
    print("🔄 Starting Hierarchical ZBM Summary Automation...")
    
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
                        'Doctor: Customer Code', 'Assigned Request Ids', 'Request Status']
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
        # Use the corrected logic from our validation
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
            ('not permitted',): 'Not Permitted',
            ('delivered', 'out of stock', 'return'): 'Delivered',
            ('action pending / in process', 'dispatch pending', 'out of stock'): 'Dispatch Pending',
            ('dispatch pending', 'not permitted'): 'Not Permitted',
        }
        rules.update(additional_rules)

        # Group statuses by request id from master data
        grouped = df.groupby('Assigned Request Ids')['Request Status'].apply(list).reset_index()

        def get_final_answer(status_list):
            key = tuple(sorted(set(normalize(s) for s in status_list)))
            return rules.get(key, '❌ No matching rule')

        grouped['Request Status'] = grouped['Request Status'].apply(lambda lst: sorted(set(lst), key=str))
        grouped['Final Answer'] = grouped['Request Status'].apply(get_final_answer)

        # Merge Final Answer back to main dataframe
        df = df.merge(grouped[['Assigned Request Ids', 'Final Answer']], on='Assigned Request Ids', how='left')
        
        print(f"✅ Final status calculated for {len(grouped)} unique requests")
        
    except Exception as e:
        print(f"❌ Error computing final status: {e}")
        return

    # Create hierarchical summary
    print("📈 Creating hierarchical ZBM → ABM → TBM summary...")
    
    # Define status categories for aggregation
    status_categories = {
        'out_of_stock_on_hold': ['Out of stock', 'On hold', 'Not permitted'],
        'request_raised': ['Request Raised'],
        'action_pending': ['Action pending / In Process'],
        'dispatch_pending': ['Dispatch Pending'],
        'delivered': ['Delivered'],
        'dispatched_in_transit': ['Dispatched & In Transit'],
        'rto': ['RTO']
    }

    # Create hierarchical aggregation
    hierarchical_summary = []
    
    # Get unique ZBMs
    zbms = df[['ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID']].drop_duplicates().sort_values('ZBM Terr Code')
    
    for _, zbm_row in zbms.iterrows():
        zbm_code = zbm_row['ZBM Terr Code']
        zbm_name = zbm_row['ZBM Name']
        zbm_email = zbm_row['ZBM EMAIL_ID']
        
        zbm_data = df[df['ZBM Terr Code'] == zbm_code]
        
        # ZBM Level Summary
        zbm_summary = {
            'Level': 'ZBM',
            'ZBM_Code': zbm_code,
            'ZBM_Name': zbm_name,
            'ZBM_Email': zbm_email,
            'ABM_Code': '',
            'ABM_Name': '',
            'ABM_Email': '',
            'TBM_HQ': '',
            'TBM_Email': '',
            'Unique_TBMs': zbm_data['TBM EMAIL_ID'].nunique(),
            'Unique_HCPs': zbm_data['Doctor: Customer Code'].nunique(),
            'Unique_Requests': zbm_data['Assigned Request Ids'].nunique(),
        }
        
        # Calculate status metrics for ZBM
        for category_name, status_list in status_categories.items():
            count = zbm_data[zbm_data['Final Answer'].isin(status_list)]['Assigned Request Ids'].nunique()
            zbm_summary[f'count_{category_name}'] = count
        
        # Calculate derived metrics
        zbm_summary['Request_Cancelled_Out_of_Stock'] = zbm_summary['count_out_of_stock_on_hold']
        zbm_summary['Action_Pending_at_HO'] = zbm_summary['count_action_pending']
        zbm_summary['Pending_for_Invoicing'] = 0  # Placeholder
        zbm_summary['Pending_for_Dispatch'] = zbm_summary['count_dispatch_pending']
        zbm_summary['Delivered'] = zbm_summary['count_delivered']
        zbm_summary['Dispatched_In_Transit'] = zbm_summary['count_dispatched_in_transit']
        zbm_summary['RTO'] = zbm_summary['count_rto']
        zbm_summary['Requests_Dispatched'] = zbm_summary['Delivered'] + zbm_summary['Dispatched_In_Transit'] + zbm_summary['RTO']
        zbm_summary['Sent_to_HUB'] = zbm_summary['Pending_for_Invoicing'] + zbm_summary['Pending_for_Dispatch'] + zbm_summary['Requests_Dispatched']
        zbm_summary['Requests_Raised'] = zbm_summary['Request_Cancelled_Out_of_Stock'] + zbm_summary['Action_Pending_at_HO'] + zbm_summary['Sent_to_HUB']
        
        # RTO Reasons (placeholders)
        zbm_summary['Incomplete_Address'] = 0
        zbm_summary['Doctor_Non_Contactable'] = 0
        zbm_summary['Doctor_Refused_to_Accept'] = 0
        zbm_summary['Hold_Delivery'] = 0
        
        hierarchical_summary.append(zbm_summary)
        
        # Get unique ABMs under this ZBM
        abms = zbm_data[['ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID']].drop_duplicates().sort_values('ABM Terr Code')
        
        for _, abm_row in abms.iterrows():
            abm_code = abm_row['ABM Terr Code']
            abm_name = abm_row['ABM Name']
            abm_email = abm_row['ABM EMAIL_ID']
            
            abm_data = zbm_data[zbm_data['ABM Terr Code'] == abm_code]
            
            # ABM Level Summary
            abm_summary = {
                'Level': 'ABM',
                'ZBM_Code': zbm_code,
                'ZBM_Name': zbm_name,
                'ZBM_Email': zbm_email,
                'ABM_Code': abm_code,
                'ABM_Name': abm_name,
                'ABM_Email': abm_email,
                'TBM_HQ': '',
                'TBM_Email': '',
                'Unique_TBMs': abm_data['TBM EMAIL_ID'].nunique(),
                'Unique_HCPs': abm_data['Doctor: Customer Code'].nunique(),
                'Unique_Requests': abm_data['Assigned Request Ids'].nunique(),
            }
            
            # Calculate status metrics for ABM
            for category_name, status_list in status_categories.items():
                count = abm_data[abm_data['Final Answer'].isin(status_list)]['Assigned Request Ids'].nunique()
                abm_summary[f'count_{category_name}'] = count
            
            # Calculate derived metrics
            abm_summary['Request_Cancelled_Out_of_Stock'] = abm_summary['count_out_of_stock_on_hold']
            abm_summary['Action_Pending_at_HO'] = abm_summary['count_action_pending']
            abm_summary['Pending_for_Invoicing'] = 0  # Placeholder
            abm_summary['Pending_for_Dispatch'] = abm_summary['count_dispatch_pending']
            abm_summary['Delivered'] = abm_summary['count_delivered']
            abm_summary['Dispatched_In_Transit'] = abm_summary['count_dispatched_in_transit']
            abm_summary['RTO'] = abm_summary['count_rto']
            abm_summary['Requests_Dispatched'] = abm_summary['Delivered'] + abm_summary['Dispatched_In_Transit'] + abm_summary['RTO']
            abm_summary['Sent_to_HUB'] = abm_summary['Pending_for_Invoicing'] + abm_summary['Pending_for_Dispatch'] + abm_summary['Requests_Dispatched']
            abm_summary['Requests_Raised'] = abm_summary['Request_Cancelled_Out_of_Stock'] + abm_summary['Action_Pending_at_HO'] + abm_summary['Sent_to_HUB']
            
            # RTO Reasons (placeholders)
            abm_summary['Incomplete_Address'] = 0
            abm_summary['Doctor_Non_Contactable'] = 0
            abm_summary['Doctor_Refused_to_Accept'] = 0
            abm_summary['Hold_Delivery'] = 0
            
            hierarchical_summary.append(abm_summary)
            
            # Get unique TBMs under this ABM
            tbms = abm_data[['TBM HQ', 'TBM EMAIL_ID']].drop_duplicates().sort_values('TBM HQ')
            
            for _, tbm_row in tbms.iterrows():
                tbm_hq = tbm_row['TBM HQ']
                tbm_email = tbm_row['TBM EMAIL_ID']
                
                tbm_data = abm_data[abm_data['TBM EMAIL_ID'] == tbm_email]
                
                # TBM Level Summary
                tbm_summary = {
                    'Level': 'TBM',
                    'ZBM_Code': zbm_code,
                    'ZBM_Name': zbm_name,
                    'ZBM_Email': zbm_email,
                    'ABM_Code': abm_code,
                    'ABM_Name': abm_name,
                    'ABM_Email': abm_email,
                    'TBM_HQ': tbm_hq,
                    'TBM_Email': tbm_email,
                    'Unique_TBMs': 1,  # Each TBM row represents 1 TBM
                    'Unique_HCPs': tbm_data['Doctor: Customer Code'].nunique(),
                    'Unique_Requests': tbm_data['Assigned Request Ids'].nunique(),
                }
                
                # Calculate status metrics for TBM
                for category_name, status_list in status_categories.items():
                    count = tbm_data[tbm_data['Final Answer'].isin(status_list)]['Assigned Request Ids'].nunique()
                    tbm_summary[f'count_{category_name}'] = count
                
                # Calculate derived metrics
                tbm_summary['Request_Cancelled_Out_of_Stock'] = tbm_summary['count_out_of_stock_on_hold']
                tbm_summary['Action_Pending_at_HO'] = tbm_summary['count_action_pending']
                tbm_summary['Pending_for_Invoicing'] = 0  # Placeholder
                tbm_summary['Pending_for_Dispatch'] = tbm_summary['count_dispatch_pending']
                tbm_summary['Delivered'] = tbm_summary['count_delivered']
                tbm_summary['Dispatched_In_Transit'] = tbm_summary['count_dispatched_in_transit']
                tbm_summary['RTO'] = tbm_summary['count_rto']
                tbm_summary['Requests_Dispatched'] = tbm_summary['Delivered'] + tbm_summary['Dispatched_In_Transit'] + tbm_summary['RTO']
                tbm_summary['Sent_to_HUB'] = tbm_summary['Pending_for_Invoicing'] + tbm_summary['Pending_for_Dispatch'] + tbm_summary['Requests_Dispatched']
                tbm_summary['Requests_Raised'] = tbm_summary['Request_Cancelled_Out_of_Stock'] + tbm_summary['Action_Pending_at_HO'] + tbm_summary['Sent_to_HUB']
                
                # RTO Reasons (placeholders)
                tbm_summary['Incomplete_Address'] = 0
                tbm_summary['Doctor_Non_Contactable'] = 0
                tbm_summary['Doctor_Refused_to_Accept'] = 0
                tbm_summary['Hold_Delivery'] = 0
                
                hierarchical_summary.append(tbm_summary)

    # Convert to DataFrame
    hierarchical_df = pd.DataFrame(hierarchical_summary)
    
    # Save hierarchical summary
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    csv_output = f"hierarchical_zbm_summary_{timestamp}.csv"
    hierarchical_df.to_csv(csv_output, index=False)
    print(f"💾 Saved hierarchical summary to {csv_output}")
    
    # Create Excel output with proper formatting
    excel_output = f"hierarchical_zbm_summary_{timestamp}.xlsx"
    
    with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
        # Write hierarchical summary
        hierarchical_df.to_excel(writer, sheet_name='Hierarchical_Summary', index=False)
        
        # Create separate sheets for each level
        zbm_only = hierarchical_df[hierarchical_df['Level'] == 'ZBM'].copy()
        abm_only = hierarchical_df[hierarchical_df['Level'] == 'ABM'].copy()
        tbm_only = hierarchical_df[hierarchical_df['Level'] == 'TBM'].copy()
        
        zbm_only.to_excel(writer, sheet_name='ZBM_Level', index=False)
        abm_only.to_excel(writer, sheet_name='ABM_Level', index=False)
        tbm_only.to_excel(writer, sheet_name='TBM_Level', index=False)
    
    print(f"✅ Successfully created hierarchical Excel file: {excel_output}")
    
    # Print summary statistics
    print("\n📊 Hierarchical Summary Statistics:")
    print(f"   Total ZBMs: {len(zbm_only)}")
    print(f"   Total ABMs: {len(abm_only)}")
    print(f"   Total TBMs: {len(tbm_only)}")
    print(f"   Total Unique HCPs: {hierarchical_df['Unique_HCPs'].sum()}")
    print(f"   Total Unique Requests: {hierarchical_df['Unique_Requests'].sum()}")
    print(f"   Total Delivered: {hierarchical_df['Delivered'].sum()}")
    print(f"   Total RTO: {hierarchical_df['RTO'].sum()}")
    
    print("\n📋 Sample hierarchical data:")
    print(hierarchical_df.head(10).to_string(index=False))
    
    print("\n🎉 Hierarchical ZBM Summary automation completed successfully!")

if __name__ == "__main__":
    create_hierarchical_zbm_summary()


