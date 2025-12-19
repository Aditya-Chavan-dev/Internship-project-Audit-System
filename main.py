import webbrowser
from datetime import datetime
import pyodbc
from flask import Flask, render_template, redirect, url_for, request, jsonify, flash
import base64
import os
import configparser
import win32com.client as win32


# Declaration of global parameters
date_stamp = datetime.now().strftime('%Y-%m-%d')
global laType, lineNo, dieNo, auditorName, auditeeName

# Read the config file
config_data = configparser.ConfigParser()
config_data.read("example.ini")

# File saving details
path = config_data["path"]
uploadPath = path.get('uploadPath')
excelPath = path.get('excelPath')
excelSave = path.get('excelSave')
pdfSave = path.get('pdfSave')

# Flask Setup
app = Flask(__name__)
app.secret_key = 'you_will_never_gueses'

# Create Target Directory
TARGET_DIR = os.path.join(uploadPath, date_stamp)
app.config['TARGET_DIR'] = TARGET_DIR
excelSave_date_wise = os.path.join(excelSave, date_stamp)


def create_directory_if_not_exists(directory_path):
    """Creates a directory if it does not exist."""
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        print(f"Directory created: {directory_path}")
    else:
        print(f"Directory already exists: {directory_path}")


# Create the directories
create_directory_if_not_exists(TARGET_DIR)
create_directory_if_not_exists(excelSave_date_wise)

# Database connection parameters
db = config_data["database_details"]
server = db.get('server')
database = db.get('database')
username = db.get('username')
password = db.get('password')

# Connection String
conn_str = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'


# Connect to the database
def db_connection():
    try:
        conn = pyodbc.connect(conn_str)
        return conn
    except pyodbc.Error as e:
        print(f'{e}')
        return None


def record_exists(date_stamp, point_number, laType, lineNo, dieNo):
    conn = db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT COUNT(*) FROM LAYER_AUDIT WHERE DATE_STAMP = ? AND AREA = ? AND LS_TYPE = ? AND LINE_NO = ? AND "
        "DIE_NO = ?",
        (date_stamp, point_number, laType, lineNo, dieNo))
    return cursor.fetchone()[0] > 0


@app.route('/', methods=['POST', 'GET'])
@app.route('/layermain', methods=['POST', 'GET'])
def layermain():
    global laType, lineNo, dieNo, auditorName, auditeeName

    if request.method == 'POST':
        laType = request.form.get("ls_type")
        lineNo = request.form.get("line_no")
        dieNo = request.form.get("die_no")
        auditorName = request.form.get("auditor")
        auditeeName = request.form.get("auditee")

        print(laType, lineNo, dieNo, auditorName, auditeeName)

        # Database connection for status check (Optimized or removed if unused in future)


        audit_query = "INSERT INTO LAYER_AUDIT_DATA (DATE_STAMP, LS_TYPE, LINE_NO, DIE_NO, AUDITOR_NAME, " \
                      "AUDITEE_NAME) VALUES (?,?,?,?,?,?) "
        audit_values = date_stamp, laType, lineNo, dieNo, auditorName, auditeeName
        db_connection().execute(audit_query, audit_values).commit()

        return redirect(url_for('rm_storage_l1'))

    return render_template('layer_audit_main.html', date=date_stamp)


@app.route('/camera', methods=['POST', 'GET'])
def camera():
    return render_template('camera.html')


@app.route('/submit', methods=['POST'])
def submit_image():
    global laType, lineNo, dieNo
    try:
        data = request.json
        image_data = data['image']
        image_name = data['name']

        image_data = base64.b64decode(image_data.split(',')[1])

        filename = os.path.join(app.config['TARGET_DIR'], f'LA_{date_stamp}_{laType}_{lineNo}_{dieNo}_{image_name}')
        with open(filename, 'wb') as f:
            f.write(image_data)
        return jsonify({'success': True})
    except Exception as e:
        print(e)
        return jsonify({'success': False}), 500


@app.route('/process_data', methods=['POST'])
def process_data():
    try:
        data = request.get_json()

        # Process the data points as needed
        for item in data:
            point_number = item.get('pointNumber')
            selected_value = item.get('selectedValue')
            remarks = item.get('remarks')

            if not record_exists(date_stamp, point_number, laType, lineNo, dieNo):
                values = date_stamp, point_number, laType, lineNo, dieNo, int(selected_value), remarks
                print(f'1st Insert {point_number, laType, lineNo, dieNo, selected_value, remarks}')
                insert_query = "INSERT INTO LAYER_AUDIT (DATE_STAMP, AREA, LS_TYPE, LINE_NO, DIE_NO, POINTS, REMARKS) " \
                               "VALUES (?,?,?,?,?,?,?)"
                db_connection().execute(insert_query, values).commit()
            else:
                print(f"Record already exists for {date_stamp}, {point_number}, {lineNo}, {dieNo}")


        return jsonify({"message": "Data received successfully"})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route('/rm_storage_l1', methods=['POST', 'GET'])
def rm_storage_l1():
    num_images = 5
    image_names = [f'RMStorage_{i}.png' for i in range(1, num_images + 1)]

    expected_files = [f'LA_{date_stamp}_{laType}_{lineNo}_{dieNo}_{image_name}' for image_name in image_names]

    missing_files = [file for file in expected_files if
                     not os.path.exists(os.path.join(app.config['TARGET_DIR'], file))]

    if request.method == 'POST':
        submit = request.form.get('submit-btn')
        print(f'Submit: {submit}')

        if missing_files:
            missing_file_names = ', '.join(missing_files)
            print('missing files')
            flash(f"Please generate the image(s) for the missing file(s): {missing_file_names}", 'error')
        else:
            return redirect(url_for("rm_cutting_l1"))

    return render_template('/rm_storage_l1.html', laType=laType)


@app.route('/rm_cutting_l1', methods=['POST', 'GET'])
def rm_cutting_l1():
    num_images = 12
    image_names = [f'RMCutting_{i}.png' for i in range(1, num_images + 1)]

    expected_files = [f'LA_{date_stamp}_{laType}_{lineNo}_{dieNo}_{image_name}' for image_name in image_names]

    missing_files = [file for file in expected_files if
                     not os.path.exists(os.path.join(app.config['TARGET_DIR'], file))]

    if request.method == 'POST':
        submit = request.form.get('submit-btn')
        print(f'Submit: {submit}')

        if missing_files:
            missing_file_names = ', '.join(missing_files)
            print('missing files')
            flash(f"Please generate the image(s) for the missing file(s): {missing_file_names}", 'error')
        else:
            return redirect(url_for("ibh_heating_l1"))

    return render_template('/rm_cutting_l1.html', laType=laType)


@app.route('/ibh_heating_l1', methods=['POST', 'GET'])
def ibh_heating_l1():
    num_images = 9
    image_names = [f'IBHHeating_{i}.png' for i in range(1, num_images + 1)]

    expected_files = [f'LA_{date_stamp}_{laType}_{lineNo}_{dieNo}_{image_name}' for image_name in image_names]

    missing_files = [file for file in expected_files if
                     not os.path.exists(os.path.join(app.config['TARGET_DIR'], file))]

    if request.method == 'POST':
        submit = request.form.get('submit-btn')
        print(f'Submit: {submit}')

        if missing_files:
            missing_file_names = ', '.join(missing_files)
            print('missing files')
            flash(f"Please generate the image(s) for the missing file(s): {missing_file_names}", 'error')
        else:
            return redirect(url_for("production_l1"))

    return render_template('/ibh_heating_l1.html', laType=laType)


@app.route('/production_l1', methods=['POST', 'GET'])
def production_l1():
    num_images = 18
    image_names = [f'Production_{i}.png' for i in range(1, num_images + 1)]

    expected_files = [f'LA_{date_stamp}_{laType}_{lineNo}_{dieNo}_{image_name}' for image_name in image_names]

    missing_files = [file for file in expected_files if
                     not os.path.exists(os.path.join(app.config['TARGET_DIR'], file))]

    if request.method == 'POST':
        submit = request.form.get('submit-btn')
        print(f'Submit: {submit}')

        if missing_files:
            missing_file_names = ', '.join(missing_files)
            print('missing files')
            flash(f"Please generate the image(s) for the missing file(s): {missing_file_names}", 'error')
        else:
            return redirect(url_for("hot_inspection_l1"))

    return render_template('/production_l1.html', laType=laType)


@app.route('/hot_inspection_l1', methods=['POST', 'GET'])
def hot_inspection_l1():
    num_images = 11
    image_names = [f'HotInspection_{i}.png' for i in range(1, num_images + 1)]

    expected_files = [f'LA_{date_stamp}_{laType}_{lineNo}_{dieNo}_{image_name}' for image_name in image_names]

    missing_files = [file for file in expected_files if
                     not os.path.exists(os.path.join(app.config['TARGET_DIR'], file))]

    if request.method == 'POST':
        submit = request.form.get('submit-btn')
        print(f'Submit: {submit}')

        if missing_files:
            missing_file_names = ', '.join(missing_files)
            print('missing files')
            flash(f"Please generate the image(s) for the missing file(s): {missing_file_names}", 'error')
        else:
            return redirect(url_for("sparck_spectra_l1"))

    return render_template('/hot_inspection_l1.html', laType=laType)


@app.route('/sparck_spectra_l1', methods=['POST', 'GET'])
def sparck_spectra_l1():
    num_images = 3
    image_names = [f'SparkSpectra_{i}.png' for i in range(1, num_images + 1)]

    expected_files = [f'LA_{date_stamp}_{laType}_{lineNo}_{dieNo}_{image_name}' for image_name in image_names]

    missing_files = [file for file in expected_files if
                     not os.path.exists(os.path.join(app.config['TARGET_DIR'], file))]

    if request.method == 'POST':
        submit = request.form.get('submit-btn')
        print(f'Submit: {submit}')

        if missing_files:
            missing_file_names = ', '.join(missing_files)
            print('missing files')
            flash(f"Please generate the image(s) for the missing file(s): {missing_file_names}", 'error')
        else:
            return redirect(url_for("heat_treatment_"))

    return render_template('/sparck_spectra_l1.html', laType=laType)


@app.route('/heat_treatment_', methods=['POST', 'GET'])
def heat_treatment_():
    num_images = 12
    image_names = [f'HeatTreatment_{i}.png' for i in range(1, num_images + 1)]

    expected_files = [f'LA_{date_stamp}_{laType}_{lineNo}_{dieNo}_{image_name}' for image_name in image_names]

    missing_files = [file for file in expected_files if
                     not os.path.exists(os.path.join(app.config['TARGET_DIR'], file))]

    if request.method == 'POST':
        submit = request.form.get('submit-btn')
        print(f'Submit: {submit}')

        if missing_files:
            missing_file_names = ', '.join(missing_files)
            print('missing files')
            flash(f"Please generate the image(s) for the missing file(s): {missing_file_names}", 'error')
        else:
            return redirect(url_for("score_board_new_l1"))

    return render_template('/heat_treatment_l1.html', laType=laType)


@app.route('/score_board_new_l1', methods=['POST', 'GET'])
def score_board_new_l1():
    global laType, dieNo, lineNo

    category_names = ['RM Storage', 'RM Cutting', 'IBH Heating', 'Production', 'Hot Inspection', 'Spark & Spectra',
                      'Heat Treatment']
    category_patterns = ['RMS%', 'RMC%', 'IBH%', 'Pro%', 'Hot%', 'Spark%', 'Heat%']
    scores = []

    conn = db_connection()
    if conn:
        cursor = conn.cursor()
        for name, pattern in zip(category_names, category_patterns):
            query = f"""
            SELECT SUM(POINTS)
            FROM LAYER_AUDIT
            WHERE DATE_STAMP='{date_stamp}' AND LS_TYPE='{laType}' 
            AND LINE_NO='{lineNo}' AND DIE_NO='{dieNo}' AND AREA LIKE '{pattern}'
            """
            cursor.execute(query)
            score = cursor.fetchone()[0] or 0
            scores.append({'name': name, 'score': score})
        conn.close()
    else:
        print("Failed to connect to the database")

    total_score = sum(item['score'] for item in scores)

    if request.method == 'POST':
        return redirect(url_for("lareport"))

    return render_template('/score_board_new_l1.html', laType=laType, scores=scores, total_score=total_score)


@app.route('/lareport', methods=['POST', 'GET'])
def lareport():
    global laType, lineNo, dieNo

    # Generate image paths
    image_paths = {
        'RMS': generate_image_paths('RMStorage', 5),
        'RMC': generate_image_paths('RMCutting', 12),
        'IBH': generate_image_paths('IBHHeating', 9),
        'Pro': generate_image_paths('Production', 18),
        'HotInsp': generate_image_paths('HotInspection', 11),
        'Spark': generate_image_paths('SparkSpectra', 3),
        'HeatTreat': generate_image_paths('HeatTreatment', 12)
    }

    # Open Excel and load workbook
    excel_app = win32.Dispatch('Excel.Application')
    wb = excel_app.Workbooks.Open(fr'{excelPath}\Layer Audit.xlsx')

    # Process each sheet
    for idx, (sheet_name, sheet_range) in enumerate(
            [('RM Storage', 'RMS'), ('RM Cutting', 'RMC'), ('IBH Heating', 'IBH'), ('Production', 'Pro'),
             ('Hot Inspection', 'HotInsp'), ('Spark Spectra', 'Spark'), ('Heat Treatment', 'HeatTreat')], start=1):
        sheet = wb.Sheets(sheet_name)
        if sheet_name == 'RM Storage':
            lineNo_cell = 'C5'
            dieNo_cell = 'C6'
            auditorName_cell = 'E5'
            auditeeName_cell = 'E6'
            sheet.Range(lineNo_cell).Value = lineNo
            sheet.Range(dieNo_cell).Value = dieNo
            sheet.Range(auditorName_cell).Value = auditorName
            sheet.Range(auditeeName_cell).Value = auditeeName

        insert_images_and_data(sheet, image_paths[sheet_range], idx)

    # Save and close Excel workbook
    saving_path = fr'{excelSave_date_wise}\LA_{date_stamp}_{laType}_{lineNo}_{dieNo}.xlsx'
    wb.SaveAs(saving_path)
    wb.Close()
    excel_app.Quit()

    # Convert Excel to PDF
    pdf_file_path = convert_to_pdf(saving_path)

    return render_template('report.html', lineNo=lineNo, dieNo=dieNo, pdf_file_path=pdf_file_path)


def generate_image_paths(suffix, count):
    print(suffix, count)
    return [rf"{TARGET_DIR}\LA_{date_stamp}_{laType}_{lineNo}_{dieNo}_{suffix}_{i}.png" for i in range(1, count + 1)]


def insert_images_and_data(sheet, image_paths, sheet_index):
    fives_type_suffix = {
        1: 'RMS%',
        2: 'RMC%',
        3: 'IBH%',
        4: 'Pro%',
        5: 'HotIn%',
        6: 'Spark%',
        7: 'Heat%'
    }

    query_suffix = fives_type_suffix.get(sheet_index, '')

    data_query = f"SELECT LS_TYPE, POINTS, REMARKS FROM LAYER_AUDIT WHERE DATE_STAMP='{date_stamp}' AND LS_TYPE = '{laType}' " \
                 f"AND LINE_NO='{lineNo}' AND DIE_NO='{dieNo}' AND AREA LIKE '{query_suffix}'"
    cursor = db_connection().execute(data_query)
    results = cursor.fetchall()

    for i, path in enumerate(image_paths, start=8):
        insert_image(sheet, path, i)
        if i - 8 < len(results):
            insert_data(sheet, results[i - 8], i)


def insert_image(sheet, path, index):
    cell = sheet.Range(f'D{index}')
    pic = sheet.Pictures().Insert(path)
    pic.Left = cell.Left + 1
    pic.Top = cell.Top + 1


def insert_data(sheet, data, index):
    score_cell, remark_cell = f'E{index}', f'F{index}'
    sheet.Range(score_cell).Value = data[1]
    sheet.Range(remark_cell).Value = data[2]


def convert_to_pdf(excel_file_path):
    excel = win32.Dispatch('Excel.Application')
    workbook = excel.Workbooks.Open(excel_file_path)

    pdf_filename = f'LA_{date_stamp}_{laType}_{lineNo}_{dieNo}.pdf'
    pdf_save = os.path.join(pdfSave, pdf_filename)

    # Export the entire workbook as PDF
    workbook.ExportAsFixedFormat(0, pdf_save)

    workbook.Close(False)
    excel.Quit()
    return rf'static/Image/{pdf_filename}'


if __name__ == "__main__":
    webbrowser.open_new('http://127.0.0.1:5080')
    app.run(debug=True, use_reloader=False, port=5080)
