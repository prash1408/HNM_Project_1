from flask import Flask, render_template, request, redirect, url_for, flash
import sqlite3
from datetime import datetime
import os
from openpyxl import load_workbook,Workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.styles import Font

app = Flask(__name__)

EXCEL_FOLDER = r"\\Server\d\Requirement_Excel"
os.makedirs(EXCEL_FOLDER, exist_ok=True)

EXCEL_PATHS = {
    'Nomura_JAVA' :os.path.join(EXCEL_FOLDER,"Nomura_Java.xlsx"),
    'Morgan_Stanley':os.path.join(EXCEL_FOLDER,"Morgan_Stanley.xlsx"),
    'Nomura_TECH':os.path.join(EXCEL_FOLDER,"Nomura_TECH.xlsx"),
    'Nomura_NonTECH':os.path.join(EXCEL_FOLDER,"Nomura_NonTECH.xlsx"),
    'Nomura_Senior':os.path.join(EXCEL_FOLDER,"Nomura_Senior    .xlsx"),
    'British_Petrolium':os.path.join(EXCEL_FOLDER,"British_Petrolium.xlsx"),
    'Russell':os.path.join(EXCEL_FOLDER,"Russell.xlsx"),
    'MUFG':os.path.join(EXCEL_FOLDER,"MUFG.xlsx"),
    'Chevron':os.path.join(EXCEL_FOLDER,"Chevron.xlsx"),
    'Lufthansa':os.path.join(EXCEL_FOLDER,"Lufthansa.xlsx"),
    'Interactive_Brokers':os.path.join(EXCEL_FOLDER,"Interactive_Brokers.xlsx"),
}

COLUMN_ORDER = {
    'sequence' :['Status','date','req_id','req_Name','Mandatory_skills','Desirable_Skills','exp_range','sal_budget','Notice_Period','Total_Upload','Location','Client_Spoc','Remarks']
}
def get_db_connection():
    conn = sqlite3.connect('Requirement.db')
    conn.row_factory = sqlite3.Row
    return conn

def append_to_excel(file_or_key, data_dict):

    file_path = EXCEL_PATHS.get(file_or_key, file_or_key)

    if not file_path.endswith('.xlsx'):
        raise ValueError(f" File must be .xlsx format: {file_path}")

    headers = list(data_dict.keys())
    values = list(data_dict.values())

    try:

        if not os.path.exists(file_path):
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            wb = Workbook()
            ws = wb.active
            ws.append(headers)
            for cell in ws[1]:
                cell.font = Font(bold=True)
            ws.append(values)
            wb.save(file_path)
            print(f" Created new Excel file and wrote: {file_path}")
        else:
            wb = load_workbook(file_path)
            ws = wb.active
            existing_headers = [cell.value for cell in ws[1]]

            if existing_headers != headers:
                print(" Header mismatch in Excel. Proceeding without updating headers.")

            ws.append(values)
            wb.save(file_path)
            print(f" Appended new row to: {file_path}")

    except InvalidFileException:
        print(f" Invalid Excel file format: {file_path}. Delete and recreate.")
        raise
    except Exception as e:
        print(f" Error writing to Excel: {e}")
        raise

def export_table_to_excel(table_name, file_path):
    column_order = COLUMN_ORDER.get('sequence')
    if not column_order:
        print(f" Column order not defined for table: {table_name}")
        return

    with sqlite3.connect('Requirement.db') as conn:
        c = conn.cursor()
        column_str = ', '.join(column_order)  # Don't use quotes around column names
        try:
            c.execute(f"SELECT {column_str} FROM {table_name}")
        except Exception as e:
            print(f" SQL Error: {e}")
            return

        rows = c.fetchall()

    wb = Workbook()
    ws = wb.active
    ws.append(column_order)  # write custom header
    for row in rows:
        ws.append(row)

    wb.save(file_path)
    print(f" Excel file updated for {table_name} → {file_path}")


@app.route('/')
def home():
    conn = sqlite3.connect('Requirement.db')
    conn.row_factory = sqlite3.Row    #'Nomura_Senior'
    clients = ['Nomura_JAVA','Nomura_TECH','Nomura_Senior','Nomura_NonTECH','British_Petrolium', 'Morgan_Stanley', 'Russell', 'MUFG','Chevron','Lufthansa','Interactive_Brokers']
    notices = []
    today_str = datetime.now().strftime('%d-%m-%Y')
    for table in clients:
        try:
            latest = conn.execute(f'''SELECT req_id, req_Name, date, Mandatory_skills, Remarks FROM [{table}] WHERE date=? ORDER BY id DESC''',(today_str,)).fetchone()

            if latest:
                entry_date = latest['date']
                try:
                    parsed_date = datetime.strptime(entry_date, '%d-%m-%Y')
                    if parsed_date.strftime('%d-%m-%Y') == today_str:
                        notices.append({
                            'client':table,
                            'req_id':latest['req_id'], 
                            'req_Name':latest['req_Name'], 
                            'date':latest['date'], 
                            'Mandatory_skills':latest['Mandatory_skills'],
                            'Remarks':latest['Remarks']
                            })
                except ValueError:
                    continue
        except sqlite3.OperationalError:
            continue
    conn.close()
    return render_template('home.html', clients=clients, notices=notices)

@app.route('/dashboard/<client_name>')
def dashboard(client_name):
    conn = get_db_connection()
    conn.execute(f'''
        CREATE TABLE IF NOT EXISTS {client_name} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            req_id TEXT,
            req_Name TEXT,
            date TEXT,
            sal_budget TEXT,
            Notice_Period TEXT,
            Mandatory_skills TEXT,
            Desirable_Skills TEXT,
            exp_range TEXT,
            Client_Spoc TEXT,
            Total_Upload TEXT,
            Remarks TEXT
        )
    ''')
    conn.commit()

    rows = conn.execute(f'SELECT * FROM {client_name} ORDER BY id DESC').fetchall()
    conn.close()
    return render_template('dashboard.html', client=client_name, rows=rows)
    


@app.route('/add/<client_name>')
def add_requirement(client_name):
    today = datetime.now().strftime('%d-%m-%Y')
    return render_template('form.html', client=client_name, today=today)

@app.route('/submit/<client_name>', methods=['POST'])
def submit(client_name):
    req_id = request.form['req_id']
    req_Name = request.form['req_Name']
    date = request.form['date']
    sal_budget = request.form['sal_budget']
    Notice_Period = request.form['Notice_Period']
    Mandatory_skills = request.form['Mandatory_skills']
    Desirable_Skills = request.form['Desirable_Skills']
    exp_range = request.form['exp_range']
    Client_Spoc = request.form['Client_Spoc']
    Total_Upload = request.form['Total_Upload']
    Remarks = request.form['Remarks']
    Location = request.form['Location']
    Status = request.form['Status']

    conn = get_db_connection()
    conn.execute(f'''
        INSERT INTO {client_name} (req_id, req_Name, date, sal_budget, Notice_Period, Mandatory_skills, Desirable_Skills, exp_range, Client_Spoc, Total_Upload, Remarks, Location, Status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (req_id, req_Name, date, sal_budget, Notice_Period, Mandatory_skills, Desirable_Skills, exp_range, Client_Spoc, Total_Upload, Remarks, Location, Status))
    conn.commit()
    conn.close()
    
    form_data = {
        'Status' :request.form.get('Status'),
        'date' :request.form.get('date'),
        'req_id' :request.form.get('req_id'),
        'req_Name' :request.form.get('req_Name'),
        'Mandatory_skills' :request.form.get('Mandatory_skills'),
        'Desirable_Skills' :request.form.get('Desirable_Skills'),
        'exp_range' :request.form.get('exp_range'),
        'sal_budget' :request.form.get('sal_budget'),
        'Notice_Period' :request.form.get('Notice_Period'),
        'Total_Upload':request.form.get('Total_Upload'),
        'Location' :request.form.get('Location'),
        'Client_Spoc' :request.form.get('Client_Spoc'),
        'Remarks':request.form.get('Remarks')
    }
    excel_data = form_data.copy()
    append_to_excel(EXCEL_PATHS[client_name] , excel_data)

    return redirect(url_for('dashboard', client_name=client_name))

@app.route('/edit/<client_name>/<int:entry_id>', methods=['GET', 'POST'])
def edit_entry(client_name, entry_id):
    conn = get_db_connection()
    c = conn.cursor()

    if request.method == 'POST':
        req_id = request.form['req_id']
        req_Name = request.form['req_Name']
        date = request.form['date']
        sal_budget = request.form['sal_budget']
        Notice_Period = request.form['Notice_Period']
        Mandatory_skills = request.form['Mandatory_skills']
        Desirable_Skills = request.form['Desirable_Skills']
        exp_range = request.form['exp_range']
        Client_Spoc = request.form['Client_Spoc']
        Total_Upload = request.form['Total_Upload']
        Remarks = request.form['Remarks']
        Location = request.form['Location']
        Status = request.form['Status']

        c.execute(f'''
            UPDATE {client_name}
            SET req_id=?, req_Name=?, date=?, sal_budget=?, Notice_Period=?, Mandatory_skills=?, Desirable_Skills=?, exp_range=?, Client_Spoc=?, Total_Upload=?, Remarks=?, Location=?, Status=?
            WHERE id=?
        ''', (req_id, req_Name, date, sal_budget, Notice_Period, Mandatory_skills, Desirable_Skills, exp_range, Client_Spoc, Total_Upload, Remarks, Location, Status, entry_id))
        conn.commit()
        conn.close()

        excel_file = EXCEL_PATHS.get(client_name)
        if excel_file:
            try:
                export_table_to_excel(client_name, excel_file)
            except Exception as e:
                print(f"Excel Update Failed: {e}")
        # flash("entry updated succesfully..")  
        return redirect(url_for('dashboard', client_name=client_name))

    c.execute(f"SELECT * FROM {client_name} WHERE id = ?", (entry_id,))
    row = c.fetchone()
    conn.close()

    return render_template('edit_entry.html', client=client_name, row=row)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)