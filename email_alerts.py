import smtplib
from email.mime.text import MIMEText

EMAIL_ENABLED = False  # TURN ON LATER

def send_email(to_email, subject, body):
    if not EMAIL_ENABLED:
        print(f"[EMAIL DISABLED] {subject} → {to_email}")
        return

    sender = "your_email@gmail.com"
    password = "APP_PASSWORD"

    msg = MIMEText(body)
    msg["From"] = sender
    msg["To"] = to_email
    msg["Subject"] = subject

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)