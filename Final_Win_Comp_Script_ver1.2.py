import openpyxl

def sort_table_by_float_numbers(filename, sheetname, sort_column):
    # Load the workbook
    workbook = openpyxl.load_workbook(filename)
    
    # Select the worksheet
    worksheet = workbook[sheetname]
    
    # Get the range of the entire table
    table_range = worksheet.dimensions
    
    # Get the column index from the column letter
    sort_column_index = openpyxl.utils.column_index_from_string(sort_column)
    
    # Get all the rows in the table (excluding the header row)
    rows = list(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=sort_column_index))
    
    # Sort the rows based on the floating-point values in the specified column
    sorted_rows = sorted(rows, key=lambda row: float(row[sort_column_index - 1].value) if row[sort_column_index - 1].value else float('inf'))
    
    # Clear the existing data in the worksheet
    worksheet.delete_rows(2, worksheet.max_row)
    
    # Write the sorted rows back to the worksheet
    for sorted_row in sorted_rows:
        worksheet.append([cell.value for cell in sorted_row])
    
    # Save the modified workbook
    workbook.save(filename)
    workbook.close()

# Usage example
filename = 'result_comp.xlsx'
sheetname = 'noncompliance'
sort_column = 'C'

sort_table_by_float_numbers(filename, sheetname, sort_column)

print("Successess")