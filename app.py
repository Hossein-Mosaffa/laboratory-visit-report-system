from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, after_this_request
import tempfile
import pandas as pd
import openpyxl
from datetime import datetime
import sqlite3
import io
import os

app = Flask(__name__, static_folder='templates/static')
app.secret_key = 'your_secret_key'

# Function to initialize the database
def initialize_database():
    conn = sqlite3.connect('records.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS laboratories
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, laboratory_name TEXT, industry TEXT, certificate_number TEXT, first_issue_date TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS records
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, laboratory_id INTEGER, certificate_validity_date TEXT, evaluation_type TEXT, evaluation_date TEXT, evaluation_team_members TEXT, evaluation_result TEXT, non_conformance_count TEXT, non_conformance_deadline TEXT, follow_up_evaluation_date TEXT, file_submission_date TEXT, committee_meeting_date TEXT, committee_result TEXT, technical_manager TEXT, registrant TEXT,
                 FOREIGN KEY (laboratory_id) REFERENCES laboratories(id))''')
    conn.commit()
    conn.close()

##########Lab Functions#######################
def submit_laboratory(laboratory_name, industry, certificate_number, first_issue_date):
    conn = sqlite3.connect('records.db')
    c = conn.cursor()
    c.execute("INSERT INTO laboratories (laboratory_name, industry, certificate_number, first_issue_date) VALUES (?, ?, ?, ?)",
              (laboratory_name, industry, certificate_number, first_issue_date))
    conn.commit()
    conn.close()

def get_laboratories():
    conn = sqlite3.connect('records.db')
    c = conn.cursor()
    c.execute("SELECT id, laboratory_name, industry, certificate_number, first_issue_date FROM laboratories ORDER BY id DESC")
    laboratories = c.fetchall()
    conn.close()
    return laboratories

def get_laboratory(laboratory_id):
    conn = sqlite3.connect('records.db')
    c = conn.cursor()
    c.execute("SELECT id, laboratory_name, industry, certificate_number, first_issue_date FROM laboratories WHERE id = ?", (laboratory_id,))
    laboratory = c.fetchone()
    conn.close()
    return laboratory

def update_laboratory(laboratory_id, laboratory_name, industry, certificate_number, first_issue_date):
    conn = sqlite3.connect('records.db')
    c = conn.cursor()
    c.execute("UPDATE laboratories SET laboratory_name = ?, industry = ?, certificate_number = ?, first_issue_date = ? WHERE id = ?",
              (laboratory_name, industry, certificate_number, first_issue_date, laboratory_id))
    conn.commit()
    conn.close()

# Function to submit a record
def submit_record(laboratory_id, certificate_validity_date, evaluation_type, evaluation_date, evaluation_team_members, evaluation_result, non_conformance_count, non_conformance_deadline, follow_up_evaluation_date, file_submission_date, committee_meeting_date, committee_result, technical_manager, registrant):
    conn = sqlite3.connect('records.db')
    c = conn.cursor()
    c.execute("INSERT INTO records (laboratory_id, certificate_validity_date, evaluation_type, evaluation_date, evaluation_team_members, evaluation_result, non_conformance_count, non_conformance_deadline, follow_up_evaluation_date, file_submission_date, committee_meeting_date, committee_result, technical_manager, registrant) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
              (laboratory_id, certificate_validity_date, evaluation_type, evaluation_date, evaluation_team_members, evaluation_result, non_conformance_count, non_conformance_deadline, follow_up_evaluation_date, file_submission_date, committee_meeting_date, committee_result, technical_manager, registrant))
    conn.commit()
    conn.close()

# Function to get all records
def get_records():
    conn = sqlite3.connect('records.db')
    c = conn.cursor()
    c.execute('''
    SELECT
        r.id AS record_id,
        l.laboratory_name,
        l.industry,
        l.certificate_number,
        l.first_issue_date,
        r.certificate_validity_date,
        r.evaluation_type,
        r.evaluation_date,
        r.evaluation_team_members,
        r.evaluation_result,
        r.non_conformance_count,
        r.non_conformance_deadline,
        r.follow_up_evaluation_date,
        r.file_submission_date,
        r.committee_meeting_date,
        r.committee_result,
        r.technical_manager,
        r.registrant
    FROM
        records r
    JOIN
        laboratories l
    ON
        r.laboratory_id = l.id
    ORDER BY record_id DESC
''')

    records = c.fetchall()
    conn.close()

    return records

# Function to authenticate admin
def authenticate_admin(username, password):
    return username == 'admin' and password == '178'

# Function to generate report
def generate_report():
    conn = sqlite3.connect('records.db')
    c = conn.cursor()
    c.execute('''
    SELECT
        r.id AS record_id,
        l.laboratory_name,
        l.industry,
        l.certificate_number,
        l.first_issue_date,
        r.certificate_validity_date,
        r.evaluation_type,
        r.evaluation_date,
        r.evaluation_team_members,
        r.evaluation_result,
        r.non_conformance_count,
        r.non_conformance_deadline,
        r.follow_up_evaluation_date,
        r.file_submission_date,
        r.committee_meeting_date,
        r.committee_result,
        r.technical_manager,
        r.registrant
    FROM
        records r
    JOIN
        laboratories l
    ON
        r.laboratory_id = l.id
    ORDER BY record_id DESC
''')
    report = c.fetchall()
    conn.close()
    return report
# Edit records Functions
def get_record(record_id):
    conn = sqlite3.connect('records.db')
    c = conn.cursor()
    c.execute('''
    SELECT
        r.id AS record_id,
        r.laboratory_id,
        l.laboratory_name,
        l.industry,
        l.certificate_number,
        l.first_issue_date,
        r.certificate_validity_date,
        r.evaluation_type,
        r.evaluation_date,
        r.evaluation_team_members,
        r.evaluation_result,
        r.non_conformance_count,
        r.non_conformance_deadline,
        r.follow_up_evaluation_date,
        r.file_submission_date,
        r.committee_meeting_date,
        r.committee_result,
        r.technical_manager,
        r.registrant
    FROM
        records r
    JOIN
        laboratories l
    ON
        r.laboratory_id = l.id
    WHERE r.id = ?
    ''', (record_id,))
    record = c.fetchone()
    conn.close()
    return record

# Add this new function to update a record
def update_record(record_id, laboratory_id, certificate_validity_date, evaluation_type, evaluation_date, evaluation_team_members, evaluation_result, non_conformance_count, non_conformance_deadline, follow_up_evaluation_date, file_submission_date, committee_meeting_date, committee_result, technical_manager, registrant):
    conn = sqlite3.connect('records.db')
    c = conn.cursor()
    c.execute('''
    UPDATE records
    SET laboratory_id = ?, certificate_validity_date = ?, evaluation_type = ?, evaluation_date = ?,
        evaluation_team_members = ?, evaluation_result = ?, non_conformance_count = ?,
        non_conformance_deadline = ?, follow_up_evaluation_date = ?, file_submission_date = ?,
        committee_meeting_date = ?, committee_result = ?, technical_manager = ?, registrant = ?
    WHERE id = ?
    ''', (laboratory_id, certificate_validity_date, evaluation_type, evaluation_date, evaluation_team_members,
          evaluation_result, non_conformance_count, non_conformance_deadline, follow_up_evaluation_date,
          file_submission_date, committee_meeting_date, committee_result, technical_manager, registrant, record_id))
    conn.commit()
    conn.close()

# Add this new route for editing a record
@app.route('/edit_record/<int:record_id>', methods=['GET', 'POST'])
def edit_record(record_id):
    if 'logged_in' in session:
        if request.method == 'POST':
            laboratory_id = request.form['laboratory_id']
            certificate_validity_date = request.form['certificate_validity_date']
            evaluation_type = request.form['evaluation_type']
            evaluation_date = request.form['evaluation_date']
            evaluation_team_members = request.form['evaluation_team_members']
            evaluation_result = request.form['evaluation_result']
            non_conformance_count = request.form['non_conformance_count']
            non_conformance_deadline = request.form['non_conformance_deadline']
            follow_up_evaluation_date = request.form['follow_up_evaluation_date']
            file_submission_date = request.form['file_submission_date']
            committee_meeting_date = request.form['committee_meeting_date']
            committee_result = request.form['committee_result']
            technical_manager = request.form['technical_manager']
            registrant = request.form['registrant']
            
            update_record(record_id, laboratory_id, certificate_validity_date, evaluation_type, evaluation_date,
                          evaluation_team_members, evaluation_result, non_conformance_count, non_conformance_deadline,
                          follow_up_evaluation_date, file_submission_date, committee_meeting_date, committee_result,
                          technical_manager, registrant)
            
            #flash('Record updated successfully.', 'success')
            return redirect(url_for('admin_page'))
        else:
            record = get_record(record_id)
            laboratories = get_laboratories()
            return render_template('edit_record.html', record=record, laboratories=laboratories)
    else:
        return redirect(url_for('login'))

# Function to filter records by name and/or date
def filter_records(laboratory_name, industry, first_issue_date, certificate_validity_date, evaluation_date):
    conn = sqlite3.connect('records.db')
    c = conn.cursor()
    query = '''SELECT
        r.id AS record_id,
        l.laboratory_name,
        l.industry,
        l.certificate_number,
        l.first_issue_date,
        r.certificate_validity_date,
        r.evaluation_type,
        r.evaluation_date,
        r.evaluation_team_members,
        r.evaluation_result,
        r.non_conformance_count,
        r.non_conformance_deadline,
        r.follow_up_evaluation_date,
        r.file_submission_date,
        r.committee_meeting_date,
        r.committee_result,
        r.technical_manager,
        r.registrant
    FROM
        records r
    JOIN
        laboratories l
    ON
        r.laboratory_id = l.id
    WHERE 1=1'''

    params = []
    if laboratory_name:
        query += " AND l.laboratory_name LIKE ?"
        params.append(f"%{laboratory_name}%")
    if industry:
        query += " AND l.industry LIKE ?"
        params.append(f"%{industry}%")
    if first_issue_date:
        query += " AND l.first_issue_date = ?"
        params.append(first_issue_date)
    if certificate_validity_date:
        query += " AND r.certificate_validity_date = ?"
        params.append(certificate_validity_date)
    if evaluation_date:
        query += " AND r.evaluation_date = ?"
        params.append(evaluation_date)
    c.execute(query, params)
    records = c.fetchall()
    conn.close()
    return records

# Route for the login page
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if authenticate_admin(username, password):
            session['logged_in'] = True
            return redirect(url_for('admin_page'))
        else:
            flash('نام کاربری یا رمز عبور اشتباه است', 'alert')
            return render_template('login.html', error=True)
    return render_template('login.html', error=False)

# Route for logging out
@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

# Route for the admin page
@app.route('/', methods=['GET', 'POST'])
def admin_page():
    if 'logged_in' in session:
        records = get_records()
        return render_template('admin.html', records=records)
    else:
        return redirect(url_for('login'))

# Route for the report page
@app.route('/report', methods=['GET', 'POST'])
def report_page():
    if 'logged_in' in session:
        # Get all laboratory names from the database
        laboratories = get_laboratories()
        
        if request.method == 'POST':
            laboratory_name = request.form.get('laboratory_name')
            industry = request.form.get('industry')
            first_issue_date = request.form.get('first_issue_date')
            certificate_validity_date = request.form.get('certificate_validity_date')
            evaluation_date = request.form.get('evaluation_date')
            
            filtered_records = filter_records(laboratory_name, industry, first_issue_date, certificate_validity_date, evaluation_date)
            return render_template('report.html', filtered_records=filtered_records, laboratories=laboratories)
        else:
            return render_template('report.html', laboratories=laboratories)
    else:
        return redirect(url_for('login'))

# Route for downloading the filtered report as an Excel file
@app.route('/download_filtered_report')
def download_filtered_report():
    laboratory_name = request.args.get('laboratory_name')
    industry = request.args.get('industry')
    first_issue_date = request.args.get('first_issue_date')
    certificate_validity_date = request.args.get('certificate_validity_date')
    evaluation_date = request.args.get('evaluation_date')
    filtered_records = filter_records(laboratory_name, industry, first_issue_date, certificate_validity_date, evaluation_date)

    # Create a DataFrame
    df = pd.DataFrame(filtered_records, columns=['شناسه', 'نام آزمايشگاه', 'صنعت', 'شماره گواهينامه', 'تاريخ اولين صدور گواهينامه', 'تاريخ اعتبار گواهينامه', 'نوع ارزيابي', 'تاريخ ارزيابي', 'اعضاي تيم ارزيابي', 'نتيجه ارزيابي', 'تعداد موارد عدم انطباق', 'مهلت رفع عدم انطباق', 'تاريخ ارزيابي پيگيري', 'تاريخ ارسال پرونده به كميته', 'تاريخ برگزاري كميته فني', 'نتيجه كميته', 'مدير فني', 'ثبت كننده'])

    # Create a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        # Write DataFrame to Excel
        df.to_excel(tmp.name, index=False, engine='openpyxl')
        tmp_path = tmp.name

    @after_this_request
    def remove_file(response):
        try:
            os.remove(tmp_path)
        except Exception as error:
            app.logger.error("Error removing or closing downloaded file handle", error)
        return response
    
    return send_file(tmp_path, as_attachment=True, download_name='filtered_report.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
# Route for the laboratory form page
@app.route('/laboratory_show')
def laboratory_show():
    if 'logged_in' in session:
        laboratories = get_laboratories()  # Get the list of laboratories
        return render_template('show_labs.html', laboratories=laboratories)  # Pass the laboratories to the template
    else:
        return redirect(url_for('login'))

@app.route('/laboratory_form')
def laboratory_form():
    if 'logged_in' in session:
        return render_template('laboratory_form.html')
    else:
        return redirect(url_for('login'))

# Route for submitting the laboratory form
@app.route('/submit_laboratory', methods=['POST'])
def submit_laboratory_form():
    laboratory_name = request.form['laboratory_name']
    industry = request.form['industry']
    certificate_number = request.form['certificate_number']
    first_issue_date = request.form['first_issue_date']
    submit_laboratory(laboratory_name, industry, certificate_number, first_issue_date)
    return redirect(url_for('laboratory_show'))

# Route for editing a laboratory
@app.route('/edit_laboratory/<int:laboratory_id>', methods=['GET', 'POST'])
def edit_laboratory(laboratory_id):
    if 'logged_in' in session:
        if request.method == 'POST':
            laboratory_name = request.form['laboratory_name']
            industry = request.form['industry']
            certificate_number = request.form['certificate_number']
            first_issue_date = request.form['first_issue_date']
            update_laboratory(laboratory_id, laboratory_name, industry, certificate_number, first_issue_date)
            return redirect(url_for('laboratory_show'))
        else:
            laboratory = get_laboratory(laboratory_id)
            return render_template('edit_laboratory.html', laboratory=laboratory)
    else:
        return redirect(url_for('login'))

# Route for the records page
@app.route('/records')
def records_page():
    if 'logged_in' in session:
        laboratories = get_laboratories()
        records = get_records()
        return render_template('records.html', laboratories=laboratories, records=records)
    else:
        return redirect(url_for('login'))

@app.route('/record_form')
def record_form():
    if 'logged_in' in session:
        laboratories = get_laboratories()
        return render_template('record_form.html', laboratories=laboratories)
    return redirect(url_for('login'))

# Route for submitting a record
@app.route('/submit_record', methods=['POST'])
def submit_record_form():
    laboratory_id = request.form['laboratory_id']
    certificate_validity_date = request.form['certificate_validity_date']
    evaluation_type = request.form['evaluation_type']
    evaluation_date = request.form['evaluation_date']
    evaluation_team_members = request.form['evaluation_team_members']
    evaluation_result = request.form['evaluation_result']
    non_conformance_count = request.form['non_conformance_count']
    non_conformance_deadline = request.form['non_conformance_deadline']
    follow_up_evaluation_date = request.form['follow_up_evaluation_date']
    file_submission_date = request.form['file_submission_date']
    committee_meeting_date = request.form['committee_meeting_date']
    committee_result = request.form['committee_result']
    technical_manager = request.form['technical_manager']
    registrant = request.form['registrant']
    submit_record(laboratory_id, certificate_validity_date, evaluation_type, evaluation_date, evaluation_team_members, evaluation_result, non_conformance_count, non_conformance_deadline, follow_up_evaluation_date, file_submission_date, committee_meeting_date, committee_result, technical_manager, registrant)
    
    flash('رکورد با موفقیت ثبت شد', 'success')
    return redirect(url_for('records_page'))


if __name__ == '__main__':
    # Initialize the database
    initialize_database()
    app.run(host='0.0.0.0', port=5001)

