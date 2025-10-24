 Read data rows
        data = []
        empty_row_count = 0
        for row_idx in range(header_row + 1, ws.max_row + 1):
            row_data = []
            has_any_value = False

            for col_offset in range(len(headers)):
                col_idx = start_col + col_offset
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                
                # Check if cell has any meaningful value
                if cell_value is not None and str(cell_value).strip() != "":
                    has_any_value = True
                
                row_data.append(cell_value)

            # If row has at least one value, add it to data
            if has_any_value:
                data.append(row_data)
                empty_row_count = 0
            else:
                empty_row_count += 1
                # Stop only after 2 consecutive completely empty rows
                if empty_row_count >= 2:
                    break
