import os
import smtplib
import sys
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from datetime import date
from dotenv import load_dotenv

load_dotenv()


def send_report(pptx_path, recipient):
    sender   = os.getenv("GMAIL_SENDER_EMAIL")
    password = os.getenv("GMAIL_APP_PASSWORD")

    if not sender or not password:
        print("⚠  GMAIL_SENDER_EMAIL / GMAIL_APP_PASSWORD not set in .env — skipping email.")
        return False

    fname = Path(pptx_path).name
    msg = MIMEMultipart()
    msg["From"]    = sender
    msg["To"]      = recipient
    msg["Subject"] = f"YouTube Analysis Report — {date.today().strftime('%B %d, %Y')}"

    msg.attach(MIMEText(
        f"Hi,\n\nPlease find attached your YouTube space analysis report.\n\n"
        f"File: {fname}\nGenerated: {date.today().strftime('%B %d, %Y')}\n\nBest,\nYouTube Analyser",
        "plain",
    ))

    with open(pptx_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{fname}"')
    msg.attach(part)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, recipient, msg.as_string())

    print(f"✓ Report emailed to {recipient}")
    return True


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python tools/send_email.py <pptx_path> <recipient_email>")
        sys.exit(1)
    send_report(sys.argv[1], sys.argv[2])
