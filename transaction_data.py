import pandas as pd
import numpy as np
import datetime
pd.options.mode.chained_assignment = None
import sys
import os
from pathlib import Path
import datetime

def process_data(start_date, end_date, CIFs, filepath1, filepath2, dirpath_dst):

    print("Start Date:", start_date)
    print("End Date:", end_date)

    filename = '_' + f"_Transactions_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    print("CIFs:", CIFs)

    # Transactions here

    df1 = pd.read_csv(filepath1, dtype=str, na_filter=False)
    df2 = pd.read_csv(filepath2, dtype=str, na_filter=False)

    # Merge two csv files together
    df1 = pd.DataFrame(df1)
    df2 = pd.DataFrame(df2)
    df = pd.concat([df1, df2])
    df['TRXN_BASE_AM'] = pd.to_numeric(df['TRXN_BASE_AM'])
    
    # Check if all rows are adding up
    if len(df1.index) + len(df2.index) == len(df.index):
        print('Rows are matching!!!')
    else:
        print('Rows are not matching.')

    # Data cleaning
    ## Replace 'MI' with 'CHECK' in column 'TRXN_TYPE'
    ## Replace blanks and 'ZZ' in columns 'ORIG_CNTRY' and 'BENEF_CNTRY'
    df['TRXN_TYPE'].replace('MI', 'CHECK', inplace=True)
    df['ORIG_CNTRY'].replace(['', 'ZZ', np.nan], 'US', inplace=True)
    df['BENEF_CNTRY'].replace(['', 'ZZ', np.nan], 'US', inplace=True)

    # Drop column 'PRMRY_CUST_INTRL_ID'
    # df = df.drop(columns = ['PRMRY_CUST_INTRL_ID'])


    # Set the expected end date and actual start date
    ## Check the year is a leap year or not                  
    exp_start_date = datetime.datetime.strptime(start_date, '%Y-%m-%d')
    act_end_date = datetime.datetime.strptime(end_date, '%Y-%m-%d')

    next_year = exp_start_date.year + 1
    is_leap_year = next_year % 4 == 0 and (next_year % 100 != 0 or next_year % 400 == 0)
    days_to_add = 366 if is_leap_year else 365
    ## Calculate Expected End Date for Expected Transaction Activity
    exp_end_date = (exp_start_date + datetime.timedelta(days=days_to_add) - datetime.timedelta(days=1))

    current_year = act_end_date.year
    is_leap_year = current_year % 4 == 0 and (current_year % 100 != 0 or current_year % 400 == 0)
    days_to_subtract = 366 if is_leap_year else 365
    ## Calculate Actual Start Date for Actual Transaction Activity
    act_start_date = (act_end_date - datetime.timedelta(days=days_to_subtract) + datetime.timedelta(days=1))

    print("Actual Start Date:", act_start_date)
    print("Actual End Date:", end_date)
    print("Expected Start Date:", start_date)
    print("Expected End Date:", exp_end_date)

    df['TRXN_EXCTN_DT'] = pd.to_datetime(df['TRXN_EXCTN_DT'])


    # Sort the order by date
    output = df.sort_values(by='TRXN_EXCTN_DT')
    exp_output = output[(output['TRXN_EXCTN_DT'] >= exp_start_date) & (output['TRXN_EXCTN_DT'] <= exp_end_date)]
    act_output = output[(output['TRXN_EXCTN_DT'] >= act_start_date) & (output['TRXN_EXCTN_DT'] <= act_end_date)]


    # Look up ID in CIF list with transaction dataframe
    result_act = []
    for cus_id in CIFs:
        matching_rows = act_output[act_output['CUST_INTRL_ID'] == cus_id]
        result_act.append(matching_rows)
    act_output = pd.concat(result_act)

    result_exp = []
    for cus_id in CIFs:
        matching_rows = exp_output[exp_output['CUST_INTRL_ID'] == cus_id]
        result_exp.append(matching_rows)
    exp_output = pd.concat(result_exp)

    result_dir = Path(dirpath_dst)
    os.chdir(result_dir)

    act_output['Column 1'] = 1
    exp_output['Column 1'] = 1

    with pd.ExcelWriter(filename) as writer:
        act_output.to_excel(writer, sheet_name='Actual Transaction', index=False)
        exp_output.to_excel(writer, sheet_name='Expected Transaction', index=False)
    
    print("Excel file contains actual and expected transaction data has been created.")

    return filename


if __name__ == '__main__':
    # Get data from command-line arguments
    start_date = sys.argv[1]
    end_date = sys.argv[2]
    CIFs = sys.argv[3:]

    process_data(start_date, end_date, CIFs)
