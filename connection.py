import smtplib
from dotenv import load_dotenv
import os

load_dotenv()

LOGIN_USERNAME = os.getenv("LOGIN_USERNAME")
LOGIN_PASSWORD = os.getenv("LOGIN_PASSWORD")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")

try:
    server = smtplib.SMTP("mail-auth.th-brandenburg.de", 587)
    server.starttls()
    server.login(LOGIN_USERNAME, LOGIN_PASSWORD)
    print(f"Logged in as {LOGIN_USERNAME} successfully")
    print(f"Sending emails from {SENDER_EMAIL}")
    server.quit()
except Exception as e:
    print(f"✗ Login failed: {e}")
