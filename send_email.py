import smtplib
import os
import datetime as dt
from email.message import EmailMessage
from dotenv import load_dotenv
load_dotenv()


def send_email():
    current_datetime = dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    current_date = dt.datetime.now().strftime('%d/%m/%Y')
    prior_7_days_datetime = (dt.datetime.now() - dt.timedelta(days=7)).strftime('%d/%m/%Y %H:%M:%S')
    prior_7_days_date = (dt.datetime.now() - dt.timedelta(days=7)).strftime('%d/%m/%Y')

    sender = os.environ.get("EMAIL_USER")
    recipient = ["spyros.fragoulakis@mitsis.com", "konstantinos.pampris@mitsis.com", "manolis.arvanitis@mitsis.com"]

    email = EmailMessage()
    email["From"] = sender
    email["To"] = recipient
    email["Cc"] = ["lefteris.fthenos@mitsis.com", "petros.eritsos@mitsis.com"]
    email["Subject"] = f"Cloudtalk Automated Report - {prior_7_days_date} - {current_date}"
    # include attachment
    email.set_content(f"Καλησπέρα σας, \n\n Επισυνάπτεται το αρχείο Cloudtalk Report με την αναφορά του Call Center όπως έχει διαμορφωθεί έως τις {current_datetime}.\n\nΤο παρόν email και η αναφορά δημιουργήθηκαν αυτοματοποιημένα.\nΣε περίπτωση που αντιμετωπίσετε οποιοδήποτε πρόβλημα, παρακαλώ απαντήστε σε αυτό το email.\n\nΜε εκτίμηση,")
    #
    with open("Cloudtalk_Report.pptx", "rb") as file:
        file_data = file.read()
        file_name = file.name
    email.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

    smtp = smtplib.SMTP("smtp-mail.outlook.com", port=587)
    smtp.starttls()
    smtp.login(sender, os.environ.get("EMAIL_PASSWORD"))
    smtp.sendmail(sender, recipient, email.as_string())
    smtp.quit()
    
