import os
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

load_dotenv()

pdf_dir = os.getcwd() + "/ATTACH"
csv_file = "data.csv"

smtp_server = "smtp.gmail.com"
smtp_port = 587
smtp_username = os.getenv("SMTP_USERNAME")
smtp_password = os.getenv("SMTP_PASSWORD")
home_univ = os.getenv("HOME_UNIVERSITY")

smtp = smtplib.SMTP(smtp_server, smtp_port)
smtp.starttls()
smtp.login(smtp_username, smtp_password)

with open(csv_file, 'r', encoding='UTF-8-sig') as file:
    reader = csv.DictReader(file)
    for row in reader:
        name = row['NAME']
        univ = row['UNIV']
        lab1 = row['LAB1']
        lab2 = row['LAB2']
        email = row['EMAIL']

        msg = MIMEMultipart()
        msg['Subject'] = f"Request for Research Internship Opportunity at {univ}"
        msg['From'] = smtp_username
        msg['To'] = email

        body = f'''

Dear Prof. {name},
I hope this email finds you well. ...............
{home_univ}
'''

        msg.attach(MIMEText(body, 'plain'))

        for pdf_file in os.listdir(pdf_dir):
            if pdf_file.endswith(".pdf"):
                pdf_path = os.path.join(pdf_dir, pdf_file)
                with open(pdf_path, 'rb') as f:
                    attach = MIMEApplication(f.read(), _subtype="pdf")
                    attach.add_header('Content-Disposition', 'attachment', filename=str(pdf_file))
                    msg.attach(attach)

        print(f"Sending email to {email}...")
        smtp.send_message(msg)
        print(f"Email sent to {email} successfully!\n")

smtp.quit()
