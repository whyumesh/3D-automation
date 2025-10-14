import pandas as pd
import numpy as np
from datetime import datetime
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from copy import copy as copy_style
import warnings

# Suppress FutureWarning for groupby operations
warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')

def create_zbm_hierarchical_reports():
    """
    Create separate ZBM reports showing ABM hierarchy with perfect tallies
    Each ZBM gets a report showing all ABMs under them
    """
    
    print("🔄 Starting ZBM Hierarchical Reports Creation...")
    
    # Read master tracker data from Excel file
    print("📖 Reading Sample Master Tracker.xlsx...")
    try:
        df = pd.read_excel('Sample Master Tracker.xlsx')
        print(f"✅ Successfully loaded {len(df)} records from Sample Master Tracker.xlsx")
    except Exception as e:
        print(f"❌ Error reading Sample Master Tracker.xlsx: {e}")
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
        print(f"❌ Missing required columns in Sample Master Tracker.xlsx: {missing}")
        return

    # Remove rows where key fields are null or empty
    df = df.dropna(subset=['ZBM Terr Code', 'ZBM Name', 'ABM Terr Code', 'ABM Name', 'TBM HQ'])
    df = df[df['ZBM Terr Code'].astype(str).str.strip() != '']
    df = df[df['ABM Terr Code'].astype(str).str.strip() != '']
    df = df[df['TBM HQ'].astype(str).str.strip() != '']

    # Filter for ZBM codes that start with "ZN" (only restriction needed)
    df = df[df['ZBM Terr Code'].astype(str).str.startswith('ZN')]
    print(f"📊 After cleaning and ZBM filtering: {len(df)} records remaining")
    print(f"📊 Processing all ZBM codes starting with 'ZN' - no geographic restrictions")

    # Compute Final Answer per unique request id using rules from logic.xlsx
    print("🧠 Computing final status per unique Request Id using rules...")
    try:
        xls_rules = pd.ExcelFile('logic.xlsx')
        sheet2 = pd.read_excel(xls_rules, 'Sheet2')

        def normalize(text):
            return str(text).strip().casefold()

        rules = {}
        for _, row in sheet2.iterrows():
            statuses = [normalize(s) for s in row.drop('Final Answer').dropna().tolist()]
            statuses = tuple(sorted(set(statuses)))
            rules[statuses] = row['Final Answer']

        # Group statuses by request id from master data
        grouped = df.groupby('Assigned Request Ids')['Request Status'].apply(list).reset_index()

        def get_final_answer(status_list):
            key = tuple(sorted(set(normalize(s) for s in status_list)))
            return rules.get(key, '❌ No matching rule')

        grouped['Request Status'] = grouped['Request Status'].apply(lambda lst: sorted(set(lst), key=str))
        grouped['Final Answer'] = grouped['Request Status'].apply(get_final_answer)

        def has_action_pending(status_list):
            target = 'action pending / in process'
            return any(normalize(s) == target for s in status_list)
        grouped['Has D Pending'] = grouped['Request Status'].apply(has_action_pending)

        # Merge Final Answer back to main dataframe
        df = df.merge(grouped[['Assigned Request Ids', 'Final Answer', 'Has D Pending']], on='Assigned Request Ids', how='left')
        
        print("✅ Final status computed successfully")
        print(f"🔍 Final Answer distribution:")
        final_answer_counts = df.groupby('Assigned Request Ids')['Final Answer'].first().value_counts()
        for answer, count in final_answer_counts.items():
            print(f"   {answer}: {count}")
            
    except Exception as e:
        print(f"❌ Error computing final status from logic.xlsx: {e}")
        return
    
    # Get unique ZBMs
    zbms = df[['ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID']].drop_duplicates().sort_values('ZBM Terr Code')
    print(f"📋 Found {len(zbms)} unique ZBMs")
    
    # Create output directory
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_dir = f"ZBM_Reports_{timestamp}"
    os.makedirs(output_dir, exist_ok=True)
    print(f"📁 Created output directory: {output_dir}")
    
    # Process each ZBM
    for _, zbm_row in zbms.iterrows():
        zbm_code = zbm_row['ZBM Terr Code']
        zbm_name = zbm_row['ZBM Name']
        zbm_email = zbm_row['ZBM EMAIL_ID']
        
        print(f"\n🔄 Processing ZBM: {zbm_code} - {zbm_name}")
        
        # Filter data for this ZBM
        zbm_data = df[df['ZBM Terr Code'] == zbm_code].copy()
        
        if len(zbm_data) == 0:
            print(f"⚠️ No data found for ZBM: {zbm_code}")
            continue
        
        # Get unique ABMs under this ZBM
        abms = zbm_data.groupby(['ABM Terr Code', 'ABM Name']).agg({
            'ABM EMAIL_ID': 'first',
            'TBM HQ': 'first',
            'ABM HQ': 'first' if 'ABM HQ' in zbm_data.columns else lambda x: None
        }).reset_index()
        
        abms = abms.sort_values('ABM Terr Code')
        print(f"   📊 Found {len(abms)} ABMs under this ZBM")
        
        # Create summary data for this ZBM
        summary_data = []
        
        for _, abm_row in abms.iterrows():
            abm_code = abm_row['ABM Terr Code']
            abm_name = abm_row['ABM Name']
            abm_email = abm_row['ABM EMAIL_ID']
            tbm_hq = abm_row['TBM HQ']
            
            # Filter data for this specific ABM
            abm_data = zbm_data[(zbm_data['ABM Terr Code'] == abm_code) & (zbm_data['ABM Name'] == abm_name)].copy()
            
            print(f"      Processing {abm_name} ({abm_code}): {len(abm_data)} records")
            
            # CRITICAL: Work with UNIQUE REQUEST IDs only for accurate counts
            # Get unique requests for this ABM
            unique_requests_df = abm_data.groupby('Assigned Request Ids').agg({
                'Final Answer': 'first',
                'Has D Pending': 'first',
                'TBM EMAIL_ID': 'first',
                'Doctor: Customer Code': 'first'
            }).reset_index()
            
            print(f"         Unique Request IDs: {len(unique_requests_df)}")
            
            # Basic counts
            unique_tbms = abm_data['TBM EMAIL_ID'].nunique() if 'TBM EMAIL_ID' in abm_data.columns else 0
            unique_hcps = abm_data['Doctor: Customer Code'].nunique()
            unique_requests = len(unique_requests_df)  # This is the total unique requests
            
            print(f"         Unique TBMs: {unique_tbms}, Unique HCPs: {unique_hcps}, Unique Requests: {unique_requests}")
            
            # Calculate metrics based on UNIQUE REQUEST IDs and Final Answer
            # A: Request Cancelled / Out of Stock / On Hold / Not Permitted
            request_cancelled_out_of_stock = len(unique_requests_df[
                unique_requests_df['Final Answer'].isin(['Out of stock', 'On hold', 'Not permitted'])
            ])
            
            # B: Action Pending at HO (but NOT the ones with "Has D Pending" = True)
            # Those with Has D Pending = True should go to D (Pending for Invoicing)
            action_pending_at_ho = len(unique_requests_df[
                (unique_requests_df['Final Answer'].isin(['Action pending / In Process'])) &
                (unique_requests_df['Has D Pending'] != True)
            ])
            
            # D: Pending for Invoicing (Has D Pending = True)
            pending_for_invoicing = len(unique_requests_df[unique_requests_df['Has D Pending'] == True])
            
            # E: Pending for Dispatch
            pending_for_dispatch = len(unique_requests_df[
                unique_requests_df['Final Answer'].isin(['Dispatch Pending'])
            ])
            
            # G: Delivered
            delivered = len(unique_requests_df[
                unique_requests_df['Final Answer'].isin(['Delivered'])
            ])
            
            # H: Dispatched & In Transit
            dispatched_in_transit = len(unique_requests_df[
                unique_requests_df['Final Answer'].isin(['Dispatched & In Transit'])
            ])
            
            # I: RTO
            rto = len(unique_requests_df[
                unique_requests_df['Final Answer'].isin(['RTO'])
            ])
            
            # Calculated fields following the formulas
            # F = G + H + I (Requests Dispatched)
            requests_dispatched = delivered + dispatched_in_transit + rto
            
            # C = D + E + F (Sent to HUB)
            sent_to_hub = pending_for_invoicing + pending_for_dispatch + requests_dispatched
            
            # Requests Raised = A + B + C
            requests_raised = request_cancelled_out_of_stock + action_pending_at_ho + sent_to_hub
            
            # RTO Reasons (placeholders - would need specific data)
            incomplete_address = 0
            doctor_non_contactable = 0
            doctor_refused_to_accept = 0
            hold_delivery = 0
            
            # Create Area Name: "ABM Terr Code and ABM HQ" as per template
            if 'ABM HQ' in abm_row and pd.notna(abm_row['ABM HQ']):
                abm_hq = abm_row['ABM HQ']
            else:
                abm_hq = tbm_hq  # Fallback to TBM HQ
            area_name = f"{abm_code} and {abm_hq}"
            
            # Debug output
            print(f"         Breakdown:")
            print(f"           A (Cancelled/Out of Stock): {request_cancelled_out_of_stock}")
            print(f"           B (Action Pending at HO): {action_pending_at_ho}")
            print(f"           C (Sent to HUB): {sent_to_hub}")
            print(f"           D (Pending for Invoicing): {pending_for_invoicing}")
            print(f"           E (Pending for Dispatch): {pending_for_dispatch}")
            print(f"           F (Requests Dispatched): {requests_dispatched}")
            print(f"           G (Delivered): {delivered}")
            print(f"           H (Dispatched In Transit): {dispatched_in_transit}")
            print(f"           I (RTO): {rto}")
            print(f"         Total Requests Raised (A+B+C): {requests_raised}")
            
            # Verification: requests_raised should equal unique_requests
            if requests_raised != unique_requests:
                print(f"         ⚠️  WARNING: Mismatch! Requests Raised ({requests_raised}) != Unique Requests ({unique_requests})")
                print(f"         🔍 Checking for unmapped Final Answers:")
                all_final_answers = unique_requests_df['Final Answer'].unique()
                mapped_answers = ['Out of stock', 'On hold', 'Not permitted', 'Action pending / In Process', 
                                'Dispatch Pending', 'Delivered', 'Dispatched & In Transit', 'RTO']
                unmapped = [ans for ans in all_final_answers if ans not in mapped_answers]
                if unmapped:
                    print(f"         ⚠️  Unmapped Final Answers: {unmapped}")
                    for unmapped_answer in unmapped:
                        count = len(unique_requests_df[unique_requests_df['Final Answer'] == unmapped_answer])
                        print(f"            '{unmapped_answer}': {count} requests")
            
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
        
        # Create DataFrame for this ZBM
        zbm_summary_df = pd.DataFrame(summary_data)
        
        # Create Excel file for this ZBM
        create_zbm_excel_report(zbm_code, zbm_name, zbm_email, zbm_summary_df, output_dir)
    
    print(f"\n🎉 Successfully created {len(zbms)} ZBM reports in directory: {output_dir}")

def create_zbm_excel_report(zbm_code, zbm_name, zbm_email, summary_df, output_dir):
    """Create Excel report for a specific ZBM with perfect formatting"""
    
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
        from copy import copy as copy_style

        # Load template
        wb = load_workbook('zbm_summary.xlsx')
        ws = wb['ZBM']

        print(f"   📋 Creating Excel report for {zbm_code}...")

        # Clear data area (rows 4 onwards) - preserve headers
        data_start_row = 3  # Data starts from row 4 (index 3)
        max_clear_rows = max(len(summary_df) + 10, 100)
        
        # Handle merged cells properly
        merged_ranges_to_remove = []
        for merged_range in ws.merged_cells.ranges:
            if (merged_range.min_row >= data_start_row + 1 and 
                merged_range.min_col >= 2 and 
                merged_range.max_col <= 19):
                merged_ranges_to_remove.append(merged_range)
        
        # Remove merged cells in data area
        for merged_range in merged_ranges_to_remove:
            ws.unmerge_cells(str(merged_range))
        
        # Clear data area
        for r in range(data_start_row + 1, data_start_row + max_clear_rows):
            for c in range(2, 21):  # Columns B to T
                cell = ws.cell(row=r, column=c)
                cell.value = None

        # Define exact column mapping based on template structure
        column_mapping = {
            'Area Name': 2,           # Column B
            'ABM Name': 3,           # Column C  
            'Unique TBMs': 4,        # Column D
            'Unique HCPs': 5,        # Column E
            'Unique Requests': 6,     # Column F
            'Requests Raised': 7,     # Column G
            'Request Cancelled Out of Stock': 8,  # Column H
            'Action Pending at HO': 9,            # Column I
            'Sent to HUB': 10,                   # Column J
            'Pending for Invoicing': 11,         # Column K
            'Pending for Dispatch': 12,          # Column L
            'Requests Dispatched': 13,           # Column M
            'Delivered': 14,                     # Column N
            'Dispatched In Transit': 15,         # Column O
            'RTO': 16,                           # Column P
            'Incomplete Address': 17,            # Column Q
            'Doctor Non Contactable': 18,        # Column R
            'Doctor Refused to Accept': 19,      # Column S
            'Hold Delivery': 20                 # Column T
        }

        def copy_row_style(src_row_idx, dst_row_idx):
            """Copy formatting from source row to destination row"""
            for c in range(2, 21):  # Columns B to T
                src = ws.cell(row=src_row_idx, column=c)
                dst = ws.cell(row=dst_row_idx, column=c)
                
                if src.font:
                    dst.font = copy_style(src.font)
                if src.alignment:
                    dst.alignment = copy_style(src.alignment)
                if src.border:
                    dst.border = copy_style(src.border)
                if src.fill:
                    dst.fill = copy_style(src.fill)
                dst.number_format = src.number_format

        def write_to_cell_safely(row, col, value, formatting_func=None):
            """Write to a cell safely"""
            cell = ws.cell(row=row, column=col)
            cell.value = value
            
            if formatting_func:
                formatting_func(cell)
            
            return cell

        # Write data rows
        for i in range(len(summary_df)):
            target_row = data_start_row + 1 + i  # Start from row 4
            if target_row > ws.max_row:
                ws.insert_rows(target_row)
            
            # Copy formatting from template row 4
            copy_row_style(4, target_row)
            
            # Write data according to exact column mapping
            for col_name, col_num in column_mapping.items():
                if col_name in summary_df.columns:
                    value = summary_df.at[i, col_name]
                    
                    def apply_number_formatting(cell):
                        if isinstance(value, (int, float)) and not pd.isna(value):
                            cell.number_format = '0'  # Integer format
                    
                    write_to_cell_safely(target_row, col_num, value, apply_number_formatting)

        # Add total row
        total_row = data_start_row + 1 + len(summary_df)
        if total_row > ws.max_row:
            ws.insert_rows(total_row)
        
        # Copy formatting for total row
        copy_row_style(4, total_row)
        
        # Write totals
        def apply_total_formatting(cell):
            cell.font = Font(bold=True, name='Arial', size=10)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        write_to_cell_safely(total_row, 2, None)  # Empty first column
        write_to_cell_safely(total_row, 3, "Total", apply_total_formatting)
        
        # Calculate and write totals for each column
        for col_name, col_num in column_mapping.items():
            if col_name in summary_df.columns and col_name not in ['Area Name', 'ABM Name']:
                total_value = summary_df[col_name].sum()
                
                def apply_total_value_formatting(cell):
                    cell.font = Font(bold=True, name='Arial', size=10)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    if isinstance(total_value, (int, float)) and not pd.isna(total_value):
                        cell.number_format = '0'  # Integer format
                
                write_to_cell_safely(total_row, col_num, total_value, apply_total_value_formatting)

        # Save file
        safe_zbm_name = str(zbm_name).replace(' ', '_').replace('/', '_').replace('\\', '_')
        filename = f"ZBM_Summary_{zbm_code}_{safe_zbm_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        wb.save(filepath)
        print(f"   ✅ Created: {filename}")
        
        # Print summary statistics with verification
        total_unique_requests = summary_df['Unique Requests'].sum()
        total_requests_raised = summary_df['Requests Raised'].sum()
        
        print(f"   📊 Summary for {zbm_code}:")
        print(f"      Total ABMs: {len(summary_df)}")
        print(f"      Total Unique HCPs: {summary_df['Unique HCPs'].sum()}")
        print(f"      Total Unique Requests: {total_unique_requests}")
        print(f"      Total Requests Raised: {total_requests_raised}")
        
        if total_unique_requests != total_requests_raised:
            print(f"      ⚠️  WARNING: Unique Requests ({total_unique_requests}) != Requests Raised ({total_requests_raised})")
        else:
            print(f"      ✅ VERIFIED: Unique Requests = Requests Raised (Perfect Match!)")
        
        print(f"      Total Delivered: {summary_df['Delivered'].sum()}")
        print(f"      Total RTO: {summary_df['RTO'].sum()}")
        
    except Exception as e:
        print(f"   ❌ Error creating Excel report for {zbm_code}: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    create_zbm_hierarchical_reports()
