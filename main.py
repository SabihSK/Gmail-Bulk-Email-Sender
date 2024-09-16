"""main.py"""

import logging
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time

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


def load_html_template(file_path):
    """Load HTML file as a template"""
    with open(file_path, "r", encoding="utf-8") as file:
        return file.read()


def send_email(server, to_email, sender_name, subject, body, from_email):
    """Function to send email"""
    msg = MIMEMultipart("alternative")
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

        # Load the HTML template once
        template_html = load_html_template("email_template.html")

        for index, row in student_data.iterrows():
            name = row["Name"]
            email = row["email"]
            username = row["username"]
            password = row["password"]

            # Prepare the email content
            subject = "DGF-PPL - Login Details"

            # Replace placeholders with actual data
            body = f"""
Dear {name},\n
We are pleased to inform you that your online entrance test is scheduled for September 18, 2024. Below are your personal login credentials. Kindly ensure you do not share them with anyone.\n
https://www.classmarker.com/\n
Login Name: {username}\n
Password: {password}\n
A detailed video on how to take the test has been uploaded. Please follow the link below to watch the video and familiarize yourself with the process:\n
https://www.youtube.com/watch?v=Jbl6w3XfC8g (youtube.com)\n
\n
Important Instructions:\n
The test will consist of 30 multiple-choice questions (MCQs), and the allotted time is 1 hour.\n
The test link will be live on September 18, from 2:00 PM to 5:00 PM. You may attempt the test within this 3-hour window.\n
\n
Note: Only one attempt will be allowed once you sign in on ClassMarker.com.\n
Once you start the test, it is timed and will automatically end after 60 minutes, regardless of how much you have completed. Please ensure you finish and submit your answers within one hour.\n
Cheating, screen toggling, or copy-pasting is strictly prohibited. Any such activity will result in disqualification, as we receive notifications for these violations.\n
If your internet connection is lost during the test, you can resume it once the connection is restored without impacting your score.\n
You can take the test on a laptop, PC, or any mobile device.\n
After completing the test, please ensure to submit it within the allotted 1 hour.\n
The results will be compiled, and qualified candidates will be notified within a week to proceed to an online interview. Only those who perform well on the test and pass the interview will be invited to join the training in Karachi.\n
If you do not get selected this time, we encourage you not to lose hope and try again in the future.\n
\n
During the 3-hour test window, if you face any issues with the testing system, please reach out to our support team immediately:\n
Ali: 0329 - 2014749\n
Sabih: 0331 - 8055712\n
Best regards\n
DGF-PPL Team\n
"""

            # (
            #     template_html.replace("{{name}}", name)
            #     .replace("{{portal_url}}", portal_url)
            #     .replace("{{username}}", username)
            #     .replace("{{password}}", password)
            # )

            # Send the email
            time.sleep(10)
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
    FILE_PATH = "updated.xlsx"  # or "students_data.csv"

    # Send emails to all students
    send_emails_to_students(
        FILE_PATH,
        FROM_EMAIL,
        EMAIL_PASSWORD,
        PORTAL_URL,
        SENDER_NAME,
    )
