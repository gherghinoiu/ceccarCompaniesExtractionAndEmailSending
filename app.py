from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests
import io
import re
import threading
import uuid
import os

app = Flask(__name__)
app.secret_key = 'supersecretkey'

# --- Global Task Management ---
tasks = {}
TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_files')
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# --- Email Functionality (No Changes) ---
@app.route('/')
def index():
    return render_template('index.html')

def is_valid_email(email):
    if isinstance(email, str):
        return re.match(r"[^@]+@[^@]+\.[^@]+", email)
    return False

@app.route('/send-emails', methods=['POST'])
def send_emails():
    try:
        # SMTP details
        smtp_host = request.form['smtp_host']
        smtp_port = int(request.form['smtp_port'])
        smtp_user = request.form['smtp_user']
        smtp_pass = request.form['smtp_pass']
        smtp_secure = request.form['smtp_secure']

        # Email content
        subject = request.form['subject']
        body = request.form['body']

        # Excel file
        excel_file = request.files['excel_file']
        if not excel_file:
            flash('No Excel file provided.')
            return redirect(url_for('index'))

        df = pd.read_excel(excel_file)
        if 'email' not in df.columns:
            flash("Excel file must have a column named 'email'.")
            return redirect(url_for('index'))

        emails = df['email'].dropna().unique().tolist()
        valid_emails = [email for email in emails if is_valid_email(email)]

        if not valid_emails:
            flash('No valid emails found in the Excel file.')
            return redirect(url_for('index'))

        server = None
        if smtp_secure == 'smtps':
            server = smtplib.SMTP_SSL(smtp_host, smtp_port)
        else:
            server = smtplib.SMTP(smtp_host, smtp_port)
            server.starttls()

        server.login(smtp_user, smtp_pass)

        for email in valid_emails:
            msg = MIMEMultipart()
            msg['From'] = smtp_user
            msg['To'] = email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'html'))
            server.sendmail(smtp_user, email, msg.as_string())

        server.quit()

        flash(f'Successfully sent emails to {len(valid_emails)} recipients.')
        return redirect(url_for('index'))

    except Exception as e:
        flash(f'An error occurred: {e}')
        return redirect(url_for('index'))

# --- Asynchronous Data Extraction ---
def run_extraction_task(member_region, task_id):
    try:
        tasks[task_id] = {'status': 'running', 'progress': 'Initializing...'}
        api_url = 'https://raportare.ceccar.ro/api/search'
        headers = {
            'Accept': 'application/ld+json',
            'Content-Type': 'application/ld+json',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36',
        }

        all_items = []
        page = 1
        total_pages = 1

        while page <= total_pages:
            payload = {
                "page": page,
                "membersType": "companies",
                "memberLastName": "",
                "memberFirstName": "",
                "memberRegNumber": "",
                "memberRegion": member_region,
                "memberCurrentYearVisa": None
            }

            response = requests.post(api_url, headers=headers, json=payload)
            response.raise_for_status()
            data = response.json()

            all_items.extend(data.get('items', []))

            if page == 1:
                total_pages = data.get('pager', {}).get('pagination', {}).get('total_pages', 1)

            tasks[task_id]['progress'] = f'Processed page {page} of {total_pages}'
            page += 1

        if not all_items:
            tasks[task_id] = {'status': 'error', 'message': 'No data found for the selected region.'}
            return

        df = pd.DataFrame(all_items)
        df = df[['email', 'name', 'cui', 'region', 'phone', 'type']]

        # Save file to temp directory
        filename = f'ceccar_data_{task_id}.xlsx'
        filepath = os.path.join(TEMP_DIR, filename)
        df.to_excel(filepath, index=False, sheet_name='CECCAR Data')

        tasks[task_id] = {'status': 'complete', 'filepath': filepath}

    except Exception as e:
        tasks[task_id] = {'status': 'error', 'message': str(e)}

@app.route('/start-extraction', methods=['POST'])
def start_extraction():
    member_region_str = request.form['member_region']
    member_region = int(member_region_str) if member_region_str != "-1" else None

    task_id = str(uuid.uuid4())
    thread = threading.Thread(target=run_extraction_task, args=(member_region, task_id))
    thread.start()

    return jsonify({'task_id': task_id})

@app.route('/extraction-status/<task_id>')
def extraction_status(task_id):
    task = tasks.get(task_id, {})
    return jsonify(task)

@app.route('/download-file/<task_id>')
def download_file(task_id):
    task = tasks.get(task_id)
    if task and task.get('status') == 'complete':
        return send_file(task['filepath'], as_attachment=True, download_name=os.path.basename(task['filepath']))
    else:
        flash('File not ready or an error occurred.')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
