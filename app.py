from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests
import io
import re

app = Flask(__name__)
app.secret_key = 'supersecretkey'

@app.route('/')
def index():
    return render_template('index.html')

def is_valid_email(email):
    """Simple regex for email validation."""
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


@app.route('/download-excel', methods=['POST'])
def download_excel():
    try:
        member_region_str = request.form['member_region']
        member_region = int(member_region_str) if member_region_str != "-1" else None

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

            page += 1

        if not all_items:
            flash('No data found for the selected region.')
            return redirect(url_for('index'))

        df = pd.DataFrame(all_items)
        df = df[['email', 'name', 'cui', 'region', 'phone', 'type']]

        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl')
        df.to_excel(writer, index=False, sheet_name='CECCAR Data')
        writer.close()
        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'ceccar_data_region_{"all" if member_region is None else member_region}.xlsx'
        )

    except Exception as e:
        flash(f'An error occurred while fetching data: {e}')
        return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
