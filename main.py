from email.utils import formataddr
from smtplib import SMTP_SSL
from email.message import EmailMessage
import json
from mimetypes import guess_type
import pandas as pd
from pretty_html_table import build_table

# put your email id and app password in the credentials.json file in current directory

file = open("credentials.json")

data = json.load(file)

SENDER_EMAIL = data["email"]
MAIL_PASSWORD = data["password"]

RECEIVER_EMAIL = "sayantankar01@gmail.com"
subject = "Test Multiple  Attachment subject"

file_names = ["Reddit.jpg", "Archive.zip", "sample_data.xlsx"]

df = pd.read_excel("sample_data.xlsx")


name = "Aditya"
body = """
<html>
<head>
</head>

<body>
Hi {1},

<br><br>

Please find below the Report 
        {0}
</body>

</html>
""".format(
    build_table(
        df,
        "blue_light",
        width="auto",
        font_family="Open Sans",
        font_size="13px",
        text_align="justify",
    ),
    name,
)


def send_mail(sender_email, receiver_email, mail_password, mail_subject, all_file_names, body):
    msg = EmailMessage()

    # msg["From"] = sender_email
    msg["From"] = formataddr(("Sender's Name", sender_email))
    msg["To"] = receiver_email
    msg["Cc"] = "sayantankar02@gmail.com"
    msg["Bcc"] = "contact@sayantankar.com"
    msg["Subject"] = mail_subject
    msg['Reply-to'] = "demo@gmail.com"
    msg.add_alternative(body, subtype="html")
    for file_name in all_file_names:
        mime_type, encoding = guess_type(file_name)
        app_type, sub_type = mime_type.split("/")[0], mime_type.split("/")[1]

        with open(file_name, "rb") as FILE:
            file_data = FILE.read()
            msg.add_attachment(
                file_data, maintype=app_type, subtype=sub_type, filename=file_name
            )
            FILE.close()
    # Sending mail via SMTP server
    # with SMTP_SSL("smtp.gmail.com", 465) as smtp:
    with SMTP_SSL("mail.dillkhush.com", 465) as smtp:
        smtp.login(sender_email, mail_password)
        smtp.send_message(msg)
        smtp.close()
    print("Mail Sent Successfully")


send_mail(
    SENDER_EMAIL, RECEIVER_EMAIL, MAIL_PASSWORD, subject, file_names, body
)
