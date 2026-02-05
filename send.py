import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
import time
from dotenv import load_dotenv
import os

load_dotenv()

# Server information
OUTGOING_SERVER = os.getenv("SMTP_SERVER")
OUTGOING_PORT = os.getenv("SMTP_PORT")

# Login information
LOGIN_USERNAME = os.getenv("LOGIN_USERNAME")
LOGIN_PASSWORD = os.getenv("LOGIN_PASSWORD")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")

# Content information
EXCEL_FILE = "information.xlsx"
PDF_FOLDER = Path("./pdfs")

# Email information
EMAIL_SUBJECT = "Dein Betreff"
EMAIL_BODY = """Hallo {first_name} {last_name},

vielen Dank für dein Interesse.

Viele Grüße

"""

def send_email_with_attachment(recipient_email, first_name, last_name, pdf_path):
    """Sends email with attachment."""
    email = MIMEMultipart()
    email['From'] = SENDER_EMAIL
    email['To'] = recipient_email
    email['Subject'] = EMAIL_SUBJECT

    # Create body from constant above
    body = EMAIL_BODY.format(first_name=first_name, last_name=last_name)
    email.attach(MIMEText(body, 'plain', 'utf-8'))

    # Attach PDF
    with open(pdf_path, 'rb') as f:
        attach = MIMEApplication(f.read(), _subtype="pdf")
    attach.add_header('Content-Disposition', 'attachment', filename=pdf_path.name)
    email.attach(attach)
    return email


def main():
    # Read Excel
    df = pd.read_excel(EXCEL_FILE)

    # Expected columns: First name, Last name, email (additional columns are ignored here)
    if not {'Last name', 'First name', 'email'}.issubset(df.columns):
        print("ERROR: Columns don't match expectation")
        return

    # SMTP connect
    print(f"Connecting to {OUTGOING_SERVER}:{OUTGOING_PORT}...")
    server = smtplib.SMTP(OUTGOING_SERVER, OUTGOING_PORT)
    server.starttls()
    print("TLS encryption enabled.")

    # Login
    print(f"Logging in as {LOGIN_USERNAME}...")
    server.login(LOGIN_USERNAME, LOGIN_PASSWORD)
    print(f"Logged in successfully. Sending from: {SENDER_EMAIL}\n")

    success_count, fail_count = 0, 0

    for index, row in df.iterrows():
        first = row['First name']
        last = row['Last name']
        email = row['email']

        # PDF file pattern
        pdf_filename = f"{last.lower()}_{first.lower()}.pdf"
        pdf_path = PDF_FOLDER / pdf_filename

        if not pdf_path.exists():
            print(f"[SKIP] {first} {last} ({email}): PDF not found: {pdf_filename}")
            fail_count += 1
            continue

        try:
            msg = send_email_with_attachment(email, first, last, pdf_path)
            server.send_message(msg)
            print(f"[SENT] {first} {last} ({email})")
            success_count += 1
            time.sleep(1)
        except Exception as e:
            print(f"[FAIL] {first} {last} ({email}): {e}")
            fail_count += 1

    server.quit()
    print(f"\n=== Done: {success_count} sent, {fail_count} failed ===")

if __name__ == "__main__":
    main()
