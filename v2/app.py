from flask import Flask, request, send_file, send_from_directory
from flask_cors import CORS  # Import CORS
import os
from datetime import datetime, timedelta
import logging
from outer_registry import add_coordinators_outer, format_excel_outer
from inner_registry import format_excel_inner
import scripts as sts



def save_excel_on_server(workbook, base_dir):
    # Format the current date
    date_str = datetime.now().strftime('%d/%m/%Y')
    filename = f"Реестр от {date_str}.xlsx"

    # Ensure there's a directory to save the file in
    save_dir = os.path.join(base_dir, 'saves')
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    # Construct the full path for the file
    file_path = os.path.join(save_dir, filename)

    # Save the workbook
    workbook.save(file_path)

    # Return the relative path or any identifier you prefer
    return filename



app = Flask(__name__)
CORS(app)

@app.route('/form-excel', methods=['POST'])
def form_excel():
    # Assume the JSON data is sent in the body of the POST request
    json_data = request.get_json(force=True)

    payment_documents = json_data.get('request', [])

    sorted_payment_documents = sorted(payment_documents, key=lambda x: x.get('object_name', ''))

    json_data = {'request': sorted_payment_documents}
    logging.info('Fetched JSON data: %s', json_data)

    # Process the received JSON data into the excel
    workbook = format_excel_inner(json_data)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    filename = save_excel_on_server(workbook, base_dir)

     # Construct a download URL or simply return the filename
    download_url = f"http://192.168.30.19:25351/saves/{filename}"

    # Return the response
    return download_url

@app.route('/outer_registry', methods=['POST'])
def form_excel_outer():
    json_data = request.get_json(force = True)
    logging.info('Fetched JSON data: %s', json_data)
    workbook = format_excel_outer(json_data)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    filename = save_excel_on_server(workbook, base_dir)

     # Construct a download URL or simply return the filename
    download_url = f"http://192.168.30.19:25351/saves/{filename}"

    # Return the response
    return download_url


@app.route('/saves/<filename>', methods = ['GET'])
def download_file(filename):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    directory = os.path.join(base_dir, 'saves')
    return send_from_directory(directory, filename, as_attachment=True)

# New function
@app.route('/delete_files', methods = ['POST'])
def delete_file(directory, age_days = 14):
    '''
    Deletes files older than certain date in the given directory

    :param directory: Path to the directory
    :param age_days: difference between date of creation and current time of the files

    output: void
    '''

    cutoff_time = datetime.now() - timedelta(days = age_days)

    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)

        if os.path.isfile(file_path):
            file_modified_time = datetime.fromtimestamp(os.path.getmtime(file_path))

            if file_modified_time < cutoff_time:
                os.remove(file_path)
                print(f'Deleted {filename}.')



if __name__ == '__main__':
    app.run(debug=True, port=25351, host = '192.168.30.19')
    logging.basicConfig(level=logging.INFO)
