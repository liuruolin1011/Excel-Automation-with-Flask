from flask import Flask, render_template, request, send_file
import os
from transactions import process_data
from pivot import run_excel
import re
import pythoncom
import socket

debug = False
host = socket.gethostname(socket.gethostname())
port = 5000

app = Flask(__name__)

data = {
    'start_date': None,
    'end_date': None,
    'CIFs': []
    }


@app.route('/', methods=['GET', 'POST'])

def date_input():
    available_years = ['2021','2022','2023','2024','2025']
    available_months = {'01': 'January',
                        '02': 'February',
                        '03': 'March',
                        '04': 'April',
                        '05': 'May',
                        '06': 'June',
                        '07': 'July',
                        '08': 'August',
                        '09': 'September',
                        '10': 'October',
                        '11': 'November',
                        '12': 'December'}
    if request.method == 'POST':
        selected_year = request.form.get('year')
        selected_month = request.form.get('month')
        return f"Selected Year: {selected_year}, Selected Month: {available_months[selected_month]}"

    return render_template('date_input.html', available_years=available_years, available_months=availble_months)

dirpath_dst = r'C:\Users\Desktop\Transaction Pivot\May 2024' 

@app.route('/completed', methods=['GET', 'POST'])

def call_another_script():

    start_date = request.form.get('start_date')
    end_date = request.form.get('end_date')
    CIFs = request.form.get('CIFs')
    target_month = request.form.get('target_month')
    target_year = request.form.get('target_year')

    # Update the file every month
    filepath1 = r'\\C:\Users\Desktop\Transaction Data\transaction_data_{target_year}{target_month}_1.csv'
    filepath2 = r'\\C:\Users\Desktop\Transaction Data\transaction_data_{target_year}{target_month}_2.csv'

    # Convert the string of items to a list
    items_list = [item.strip() for item in re.findall('\d+', CIFs)]

    # Update the data dictionary
    data['start_date'] = start_date
    data['end_date'] = end_date
    data['cifs'] = items_list

    filename = process_data(data['start_date'], data['end_date'], data['cifs'], filepath1, filepath2, dirpath_dst)

    pythoncom.CoInitialize() # https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called
    try:
        filepath = os.path.join(dirpath_dst, filename)
        filename = run_excel(filepath)
    finally:
        pythoncom.CoUninitialize()
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
    app.run(debug=True, host=host, port=port)