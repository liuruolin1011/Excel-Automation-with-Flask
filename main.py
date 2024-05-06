from flask import Flask, render_template, request, send_file
import os
from transactions import process_data
from pivot import run_excel
import re
import pythoncom

app = Flask(__name__)

data = {
    'start_date': None,
    'end_date': None,
    'CIFs': []
    }


@app.route('/', methods=['GET', 'POST'])

def date_input():

    global data
    # Update the file every month
    filepath= r'\\9ushsnybfls102\usa_all$\Limited Authorized Access Folders\LCD\KYC Repapering Project\Transaction for all Customer\cif_all_trxn_202404_1.csv'

    filename =os.path.basename(filepath)
    filename_pattern = re.compile(r'_(\d{4})(\d{2})_')
    match = filename_pattern.search(filename)

    # Extract the year and month from the matched groups
    if match:
        year = match.group(1)
        month = match.group(2)

    return render_template('date_input.html', year = year, month = month)

dirpath_dst = r'C:\Users\RUOLINLIU\Desktop\Transaction Pivot\May 2024' # where the file saved on local of Ruolin's computer
# dirpath_dst = r'D:\tasks\EDD_transactions' # where the file saved on local when Brad is hosting the link

@app.route('/completed', methods=['GET', 'POST'])

def call_another_script():

    start_date = request.form.get('start_date')
    end_date = request.form.get('end_date')
    CIFs = request.form.get('CIFs')

    # Update the file every month
    filepath1 = r'\\9ushsnybfls102\usa_all$\Limited Authorized Access Folders\LCD\KYC Repapering Project\Transaction for all Customer\cif_all_trxn_202404_1.csv'
    filepath2 = r'\\9ushsnybfls102\usa_all$\Limited Authorized Access Folders\LCD\KYC Repapering Project\Transaction for all Customer\cif_all_trxn_202404_2.csv'

    # Convert the string of items to a list
    items_list = [item.strip() for item in re.findall('\d+', CIFs)]

    # Update the data dictionary
    data['start_date'] = start_date
    data['end_date'] = end_date
    data['cifs'] = items_list

    filename = process_data(data['start_date'], data['end_date'], data['cifs'], filepath1, filepath2, dirpath_dst)

    pythoncom.CoInitialize() # https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called

    filepath = os.path.join(dirpath_dst, filename)
    filename = run_excel(filepath)
    
    return render_template('result.html', ext_filename=filename, data=data)

@app.route('/<filename>')

def download_excel(filename):

    print(f"Congratulations! The Transaction Pivot has been created!")

    filepath = os.path.join(dirpath_dst, filename)

    print(filepath)

    if os.path.exists(filepath):
        print("Path found")
    
        return send_file(filepath, as_attachment=True)
    else:
        return f"File '{filepath}' not found in the current working directory."


if __name__ == '__main__':
    # app.run()
    app.run(debug=True, host='22.232.100.153', port=5000) # Ruolin
    # app.run(debug=True, host='21.232.104.67', port=5013) # Brad
