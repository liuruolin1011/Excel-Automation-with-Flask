import win32com.client as win32
import os

def pivot_table(wb, ws1, pt_ws, ws_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields, pivot_position):

    # Adjust starting location based on pivot_position
    if pivot_position == 1:  # First pivot table
        pt_loc = 'R1C1'
    elif pivot_position == 2:  # Right of the first pivot table, assuming first pivot is about 5 columns wide
        pt_loc = 'R1C6'
    elif pivot_position == 3:  # Below the first pivot table, assuming first pivot is about 15 rows tall
        pt_loc = 'R16C1'
    elif pivot_position == 4:  # Right of the third pivot table, assuming third pivot is about 5 columns wide
        pt_loc = 'R16C6'

    pt_cache = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)
    pt_cache.CreatePivotTable(TableDestination=f'{ws_name}!{pt_loc}', TableName=pt_name)

    for field_list, field_r in ((pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField), (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]),
                                                field[1], field[2]).NumberFormat = field[3]
    
    # Add CalculatedFields to get monthly average amount for value and volume
    pt_table.CalculatedFields().Add('Monthly Average Value', "='TRXN_BASE_AM' / 12", True)
    pt_table.PivotFields('Monthly Average Value').Orientation = win32c.xlDataField
    pt_table.PivotFields('Monthly Average Value').NumberFormat = '$#,##0.00'

    pt_table.CalculatedFields().Add('Monthly Average Volume', "='Column 1' / 12", True)
    pt_table.PivotFields('Monthly Average Volume').Orientation = win32c.xlDataField
    pt_table.PivotFields('Monthly Average Volume').NumberFormat = '#,##0.00'
    
    pt_ws.PivotTables(pt_name).ShowValuesRow = True
    pt_ws.PivotTables(pt_name).ColumnGrand = True


def run_excel(filename):
    """
    Opens an Excel workbook, creates pivot tables, and saves the workbook.
    
    :param filename: Path to the source Excel file
    :param debug: Enables debug mode for additional output
    :return: Path to the saved Excel file with pivot tables
    """
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    # Uncomment the next line if you want to see Excel while the script runs
    # excel.Visible = True
    wb = excel.Workbooks.Open(filename)
    ws1 = wb.Sheets("Expected Transaction")
    
    # Setup for the first pivot table
    ws2_name = 'pivot_table'
    wb.Sheets.Add(Name=ws2_name)
    ws2 = wb.Sheets(ws2_name)

    pt_name_1 = 'Expected Activity'
    pt_rows_1 = ['TRN_TYPE', 'CREDIT_DEBIT_FLAG']
    pt_cols_1 = []
    pt_filters_1 = ['CUST_INTRL_ID']
    pt_fields_1 = [['TRXN_BASE_AM', 'Value', win32.constants.xlSum, '$#,##0.00'],
                   ['FO_TRXN_SEQ_ID', 'Volume', win32.constants.xlCount, '0']]
    pivot_table(wb, ws1, ws2, ws2_name, pt_name_1, pt_rows_1, pt_cols_1, pt_filters_1, pt_fields_1, 1)
    
    # Second Pivot Table
    ws3 = wb.Sheets('Actual Transaction')
    pt_name_2 = 'Actual Activity'
    pt_rows_2 = ['TRN_TYPE', 'CREDIT_DEBIT_FLAG']
    pt_cols_2 = []
    pt_filters_2 = ['CUST_INTRL_ID']
    pt_fields_2 = [['TRXN_BASE_AM', 'Value', win32.constants.xlSum, '$#,##0.00'],
                   ['FO_TRXN_SEQ_ID', 'Volume', win32.constants.xlCount, '0']]
    pivot_table(wb, ws3, ws2, ws2_name, pt_name_2, pt_rows_2, pt_cols_2, pt_filters_2, pt_fields_2, 2)

    # Third Pivot Table

    # Fourth Pivot Table

    output_file = os.path.join(os.path.dirname(filename), 'Transaction Pivot for_' + os.path.basename(filename))

    wb.SaveAs(output_file)
    wb.Close(True)
    excel.Quit()
    return output_file

if __name__ == '__main__':
    output_file = run_excel(filename)
    print(f"Pivot tables saved in: {output_file}")
