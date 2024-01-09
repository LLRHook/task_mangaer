import os
import datetime
import calendar
import time
import pandas as pd
import win32gui
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from password import email

def get_active_window_title():
    window = win32gui.GetForegroundWindow()
    return win32gui.GetWindowText(window)

def convert_seconds_to_hms(seconds):
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f'{int(hours)}:{int(minutes):02d}:{int(seconds):02d}'

def create_directory(year, month):
    month_name = calendar.month_name[month]
    directory = os.path.join(year, month_name)
    if not os.path.exists(directory):
        os.makedirs(directory)
    return directory

def get_file_path():
    now = datetime.datetime.now()
    year = str(now.year)
    month = now.month
    day = now.day
    directory = create_directory(year, month)
    # Updated file naming convention
    filename = f'{day}_{now.strftime("%H-%M-%S")}_usage.xlsx'
    return os.path.join(directory, filename)

def send_email_with_attachment(sender_email, receiver_email, subject, body, filename, server, port, login, app_password):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    with open(filename, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())

    encoders.encode_base64(part)

    part.add_header('Content-Disposition', f"attachment; filename= {filename}")

    msg.attach(part)

    with smtplib.SMTP(server, port) as server:
        server.starttls()  # Start a TLS session
        server.login(login, app_password)
        text = msg.as_string()
        server.sendmail(sender_email, receiver_email, text)

def track_app_usage():
    usage_records = []
    current_app_name = None
    current_app_start_time = None
    current_date = datetime.date.today()

    try:
        while True:
            new_app_name = get_active_window_title()
            new_date = datetime.date.today()
            
            if current_app_name and (new_app_name != current_app_name or new_date != current_date):
                usage_end_time = datetime.datetime.now()
                duration_seconds = (usage_end_time - current_app_start_time).total_seconds()
                
                # Format the start and end times to only include hours, minutes, and seconds
                formatted_start_time = current_app_start_time.strftime("%H:%M:%S")
                formatted_end_time = usage_end_time.strftime("%H:%M:%S")

                usage_records.append({
                    'Application': current_app_name,
                    'Start Time': formatted_start_time,
                    'End Time': formatted_end_time,
                    'Duration': convert_seconds_to_hms(duration_seconds)
                })

                if new_date != current_date:
                    file_path = get_file_path()
                    save_to_excel(usage_records, file_path)
                    time.sleep(5)    # Send email with attachment after waiting for 5 seconds
                    send_email_with_attachment(email['sender_email'], email['receiver_email'], 'Application Usage', 'Please find attached your application usage report.', file_path, 'smtp.gmail.com', 587, email['login'], email['app_password'])                    
                    usage_records = []
                    current_date = new_date

                current_app_name = new_app_name
                current_app_start_time = datetime.datetime.now()

            elif not current_app_name:
                current_app_name = new_app_name
                current_app_start_time = datetime.datetime.now()

            time.sleep(1)  # Check every 1 second

    except KeyboardInterrupt:
        print("Stopping application tracking...")
        if usage_records:
            save_to_excel(usage_records, get_file_path())

def save_to_excel(data, file_path):
    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)

def main():
    track_app_usage()

if __name__ == "__main__":
    main()
