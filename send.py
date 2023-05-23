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
I hope this email finds you well. My name is Arun Ashok Badri, and I am currently a third-year undergraduate student pursuing a Bachelor's degree in CSE with a specialization in Cybersecurity at {home_univ}
, India. I am writing to express my strong interest in obtaining a research internship opportunity at {univ}.
I have been following the research and academic activities at {univ} with great admiration, and I am eager to join your esteemed institution to gain invaluable research experience in the domains of Cybersecurity, Artificial Intelligence (AI),Machine Learning (ML), Internet of Things (IOT), and Robotics.
I am writing to request your guidance and approval in securing an internship at {univ} in your research lab. The labs I am interested to work for are,
1. {lab1}
2. {lab2}
3. any other Cybersecurity/ Info sec related labs
I am available for 3 to 6 months from September of this year.
I have attached my Resume, Transcripts, Letter of Recommendations,and Statement of Purpose. If there are any specific application procedures or other documents that I need to submit, please provide me with the necessary information. I am prepared to complete any required application forms or provide additional supporting documents to strengthen my application.
Thank you for considering my request, and I look forward to the possibility of joining {univ} as a research intern under your esteemed guidance.
Yours sincerely,
Arun Ashok Badri
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