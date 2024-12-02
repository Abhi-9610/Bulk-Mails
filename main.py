from flask import Flask, request, render_template, jsonify, redirect, url_for
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import threading
import time

app = Flask(__name__)

# Replace with your email configuration
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS', 'info@danielcraft.in')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD', 'dancraft@1718')

# Global variable to store email statuses
email_status = {
    'sent': [],
    'failed': []
}
def load_html_template(file_path="templates/body.html"):
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            return file.read()
    except FileNotFoundError:
        print(f"Error: The file {file_path} was not found.")
        return None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send_bulk_emails', methods=['POST'])
def send_bulk_emails():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file:
        try:
            # Read Excel file
            df = pd.read_excel(file)
            print("Excel file successfully read.")

            # Verify columns (adjust according to your Excel structure)
            required_columns = ['Email', 'Name', 'Company']
            if not all(col in df.columns for col in required_columns):
                return jsonify({'error': 'Excel file does not contain required columns'}), 400

            # Reset the email status dictionary
            global email_status
            email_status = {'sent': [], 'failed': []}

            # Create a thread to send emails
            thread = threading.Thread(target=process_emails, args=(df,))
            thread.start()

            return jsonify({'message': 'Email sending started. Check status updates on the frontend.'})

        except Exception as e:
            return jsonify({'error': f'Error processing the file: {e}'}), 500

def process_emails(df):
    for _, row in df.iterrows():
        recipient_email = row['Email']
        recipient_name = row['Name']
        company_name = row['Company']
        try:
            send_email(recipient_email, recipient_name, company_name)
            email_status['sent'].append({'Name': recipient_name, 'Email': recipient_email, 'Company': company_name})
        except Exception as e:
            email_status['failed'].append({'Name': recipient_name, 'Email': recipient_email, 'Company': company_name, 'Error': str(e)})
        
        # Simulate delay for better demonstration (remove in production)
        time.sleep(1)

@app.route('/get_email_status', methods=['GET'])
def get_email_status():
    return jsonify(email_status)

def send_email(recipient_email, recipient_name, company_name):
    # Load the HTML template
    html_template = load_html_template("templates/body.html")
    if html_template is None:
        raise ValueError("Failed to load HTML template.")

    # Customize the template with the recipient's data
    greeting_message = f"Dear {recipient_name},\n\nGreetings from Daniel Craft Inc, Bangalore, a leading supplier of corporate apparels to corporates and institutes."
    html_body = html_template.replace("{greeting_name}", recipient_name)
    html_body = html_body.replace("{greeting_message}", greeting_message)

    subject = f"Introduction of Daniel Craft Inc, Corporate Suppliers to {company_name}"
    message = MIMEMultipart()
    message['From'] = EMAIL_ADDRESS
    message['To'] = recipient_email
    message['Subject'] = subject

    # Attach the HTML body
    message.attach(MIMEText(html_body, 'html'))

    try:
        smtp_server = "smtpout.secureserver.net"
        smtp_port = 587
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, recipient_email, message.as_string())
            print(f"Email sent to {recipient_name} ({recipient_email})")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")
        raise e  # Raise the error to handle it in the main route


if __name__ == '__main__':
    app.run(debug=True)
