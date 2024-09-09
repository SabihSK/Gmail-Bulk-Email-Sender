"""main.py"""

import logging
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd

# Set up logging
logging.basicConfig(
    filename="email_logs.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


SENDER_NAME = os.getenv("SENDER_NAME")
FROM_EMAIL = os.getenv("FROM_EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
PORTAL_URL = os.getenv("PORTAL_URL")


def read_student_data(file_path):
    """read_student_data"""
    if file_path.endswith(".csv"):
        return pd.read_csv(file_path)
    elif file_path.endswith(".xlsx"):
        return pd.read_excel(file_path)
    else:
        raise ValueError(
            "Unsupported file format. Please use a .csv or .xlsx file.",
        )


def send_email(server, to_email, sender_name, subject, body, from_email):
    """Function to send email"""
    msg = MIMEMultipart()
    msg["From"] = sender_name + " <" + from_email + ">"
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    try:
        server.sendmail(from_email, to_email, msg.as_string())
        logging.info("Email sent successfully to %s", to_email)
        print(f"Email sent to {to_email}")
    except Exception as e:
        logging.error("Failed to send email to %s. Error: %s", to_email, e)
        print(f"Failed to send email to {to_email}. Error: {e}")


# Function to prepare and send emails to all students
def send_emails_to_students(
    file_path, from_email, email_password, portal_url, sender_name
):
    """Read student data"""
    student_data = read_student_data(file_path)

    # Gmail's secure SMTP server
    server = smtplib.SMTP_SSL("smtp.gmail.com", 465)

    try:
        server.login(from_email, email_password)
        logging.info("Logged into the email server successfully.")

        for index, row in student_data.iterrows():
            name = row["Name"]
            email = row["email"]
            username = row["username"]
            password = row["password"]

            # Prepare the email content
            subject = "PPL - Login Details"
            body = f"""
Dear {name},

You are requested to attempt your test.

Below are your login details for the portal:

Portal URL: {portal_url}
Username: {username}
Password: {password}

Please use this information to access your account.

Test Duration: 1 Hour
No. of Questions: 30
All questions are important to answer

Date: 18th Sept 2024
Link availble from 2:00PM to 5:00PM

Best regards,
PPL
"""

            # Send the email
            send_email(
                server,
                email,
                sender_name,
                subject,
                body,
                from_email,
            )

        logging.info("All emails sent successfully.")
    except Exception as e:
        logging.error("Failed to log into the email server. Error: %s", e)
    finally:
        server.quit()
        logging.info("Logged out of the email server.")


# Main execution
if __name__ == "__main__":
    # Path to your CSV or Excel file
    FILE_PATH = "students_data.xlsx"  # or "students_data.csv"

    # Send emails to all students
    send_emails_to_students(
        FILE_PATH,
        FROM_EMAIL,
        EMAIL_PASSWORD,
        PORTAL_URL,
        SENDER_NAME,
    )
