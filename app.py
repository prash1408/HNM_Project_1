from flask import Flask, render_template, request, redirect, url_for, flash
import sqlite3
from datetime import datetime

app = Flask(__name__)

def get_db_connection():
    conn = sqlite3.connect('Requirement.db')
    conn.row_factory = sqlite3.Row
    return conn

@app.route('/')
def home():
    clients = ['Nomura_JAVA', 'British_Petrolium', 'Nomura_TECH', 'Morgan_Stanley', 'Russell', 'MUFG','Chevron',]
    return render_template('home.html', clients=clients)

@app.route('/dashboard/<client_name>')
def dashboard(client_name):
    conn = get_db_connection()
    conn.execute(f'''
        CREATE TABLE IF NOT EXISTS {client_name} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            req_id TEXT,
            date TEXT,
            sal_budget TEXT,
            Mandatory_skills TEXT,
            Desirable_Skills TEXT,
            exp_range TEXT,
            Remarks TEXT
        )
    ''')
    conn.commit()

    rows = conn.execute(f'SELECT * FROM {client_name} ORDER BY id DESC').fetchall()
    conn.close()
    return render_template('dashboard.html', client=client_name, rows=rows)


@app.route('/add/<client_name>')
def add_requirement(client_name):
    today = datetime.now().strftime('%d-%m-%Y %H:%M')
    return render_template('form.html', client=client_name, today=today)

@app.route('/submit/<client_name>', methods=['POST'])
def submit(client_name):
    req_id = request.form['req_id']
    date = request.form['date']
    sal_budget = request.form['sal_budget']
    Mandatory_skills = request.form['Mandatory_skills']
    Desirable_Skills = request.form['Desirable_Skills']
    exp_range = request.form['exp_range']
    Remarks = request.form['Remarks']

    conn = get_db_connection()
    conn.execute(f'''
        INSERT INTO {client_name} (req_id, date, sal_budget, Mandatory_skills, Desirable_Skills, exp_range, Remarks)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (req_id, date, sal_budget, Mandatory_skills, Desirable_Skills, exp_range, Remarks))
    conn.commit()
    conn.close()
    return redirect(url_for('dashboard', client_name=client_name))

@app.route('/edit/<client_name>/<int:entry_id>', methods=['GET', 'POST'])
def edit_entry(client_name, entry_id):
    conn = get_db_connection()
    c = conn.cursor()

    if request.method == 'POST':
        req_id = request.form['req_id']
        date = request.form['date']
        sal_budget = request.form['sal_budget']
        Mandatory_skills = request.form['Mandatory_skills']
        Desirable_Skills = request.form['Desirable_Skills']
        exp_range = request.form['exp_range']
        Remarks = request.form['Remarks']

        c.execute(f'''
            UPDATE {client_name}
            SET req_id=?, date=?, sal_budget=?, Mandatory_skills=?, Desirable_Skills=?, exp_range=?, Remarks=?
            WHERE id=?
        ''', (req_id, date, sal_budget, Mandatory_skills, Desirable_Skills, exp_range, Remarks, entry_id))

        conn.commit()
        conn.close()
        return redirect(url_for('dashboard', client_name=client_name))

    c.execute(f"SELECT * FROM {client_name} WHERE id = ?", (entry_id,))
    row = c.fetchone()
    conn.close()

    return render_template('edit_entry.html', client=client_name, row=row)

@app.route('/delete/<client_name>/<int:entry_id>', methods=['POST'])
def delete_entry(client_name, entry_id):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute(f"DELETE FROM {client_name} WHERE id = ?", (entry_id,))
    conn.commit()
    conn.close()
    return redirect(url_for('dashboard', client_name=client_name))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)