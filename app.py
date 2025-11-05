from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, jsonify
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

# --- Email Functionality ---
@app.route('/')
def index():
    return render_template('index.html')

def run_email_task(task_id, smtp_details, email_content, filepath):
    try:
        tasks[task_id] = {'status': 'running', 'progress': 'Initializing...'}

        df = pd.read_excel(filepath)
        emails = df['email'].dropna().unique().tolist()
        valid_emails = [email for email in emails if is_valid_email(email)]

        if not valid_emails:
            tasks[task_id] = {'status': 'error', 'message': 'No valid emails found in the Excel file.'}
            return

        server = None
        if smtp_details['secure'] == 'smtps':
            server = smtplib.SMTP_SSL(smtp_details['host'], smtp_details['port'])
        else:
            server = smtplib.SMTP(smtp_details['host'], smtp_details['port'])
            server.starttls()

        server.login(smtp_details['user'], smtp_details['pass'])

        total_emails = len(valid_emails)
        for i, email in enumerate(valid_emails):
            msg = MIMEMultipart()
            msg['From'] = smtp_details['user']
            msg['To'] = email
            msg['Subject'] = email_content['subject']
            msg.attach(MIMEText(email_content['body'], 'html'))
            recipients = [email, smtp_details['user']]
            server.sendmail(smtp_details['user'], recipients, msg.as_string())
            tasks[task_id]['progress'] = f'Sent {i+1} of {total_emails} emails'

        server.quit()

        tasks[task_id] = {'status': 'complete', 'message': f'Successfully sent emails to {total_emails} recipients.'}

    except Exception as e:
        tasks[task_id] = {'status': 'error', 'message': str(e)}

@app.route('/send-emails', methods=['POST'])
def send_emails():
    try:
        smtp_details = {
            'host': request.form['smtp_host'],
            'port': int(request.form['smtp_port']),
            'user': request.form['smtp_user'],
            'pass': request.form['smtp_pass'],
            'secure': request.form['smtp_secure']
        }

        email_content = {
            'subject': request.form['subject'],
            'body': request.form['body']
        }

        excel_file = request.files['excel_file']
        if not excel_file:
            return jsonify({'error': 'No Excel file provided.'}), 400

        # Save the uploaded file with a datetime prefix
        from datetime import datetime
        filename = f"{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_{excel_file.filename}"
        filepath = os.path.join(TEMP_DIR, filename)
        excel_file.save(filepath)

        task_id = str(uuid.uuid4())
        thread = threading.Thread(target=run_email_task, args=(task_id, smtp_details, email_content, filepath))
        thread.start()

        return jsonify({'task_id': task_id})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/send-emails-status/<task_id>')
def send_emails_status(task_id):
    task = tasks.get(task_id, {})
    return jsonify(task)

def is_valid_email(email):
    if isinstance(email, str):
        return re.match(r"[^@]+@[^@]+\.[^@]+", email)
    return False

# --- Asynchronous Data Extraction ---
def run_extraction_task(member_region, region_name, task_id):
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

        # First request to check for data
        payload = {
            "page": page,
            "membersType": "companies",
            "memberLastName": "",
            "memberFirstName": "",
            "memberRegNumber": "",
            "memberRegion": member_region,
            "memberCurrentYearVisa": None
        }
        response = requests.post(api_url, headers=headers, json=payload, timeout=30)
        response.raise_for_status()
        data = response.json()

        if not data.get('items'):
            tasks[task_id] = {'status': 'error', 'message': 'No data found for the selected region.'}
            return

        all_items.extend(data.get('items', []))
        total_pages = data.get('pager', {}).get('pagination', {}).get('total_pages', 1)
        tasks[task_id]['progress'] = f'Processed page 1 of {total_pages}'
        page = 2

        while page <= total_pages:
            payload['page'] = page
            response = requests.post(api_url, headers=headers, json=payload)
            response.raise_for_status()
            data = response.json()
            all_items.extend(data.get('items', []))
            tasks[task_id]['progress'] = f'Processed page {page} of {total_pages}'
            page += 1

        df = pd.DataFrame(all_items)
        df = df[['email', 'name', 'cui', 'region', 'phone', 'type']]

        # Save file to temp directory
        from datetime import datetime
        region_str = region_name.replace(" ", "_")
        filename = f"{region_str}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        filepath = os.path.join(TEMP_DIR, filename)
        df.to_excel(filepath, index=False, sheet_name='CECCAR Data')

        tasks[task_id] = {'status': 'complete', 'filepath': filepath, 'filename': filename}

    except Exception as e:
        tasks[task_id] = {'status': 'error', 'message': str(e)}

@app.route('/start-extraction', methods=['POST'])
def start_extraction():
    member_region_str = request.form['member_region']
    region_name = request.form['region_name']
    member_region = int(member_region_str) if member_region_str != "-1" else None

    task_id = str(uuid.uuid4())
    thread = threading.Thread(target=run_extraction_task, args=(member_region, region_name, task_id))
    thread.start()

    return jsonify({'task_id': task_id})

@app.route('/extraction-status/<task_id>')
def extraction_status(task_id):
    task = tasks.get(task_id, {})
    return jsonify(task)

@app.route('/download-file/<filename>')
def download_file(filename):
    return send_from_directory(TEMP_DIR, filename, as_attachment=True)

@app.route('/file-history')
def file_history():
    files = [f for f in os.listdir(TEMP_DIR) if f.endswith('.xlsx')]
    return jsonify(files)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
