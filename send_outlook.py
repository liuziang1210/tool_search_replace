import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# def send_email(host, port, username, password, recipient, subject, body):
#     # Create a MIME multipart message
#     message = MIMEMultipart()
#     message['From'] = username
#     message['To'] = recipient
#     message['Subject'] = subject
#
#     # Add the email body to the message
#     message.attach(MIMEText(body, 'plain'))
#
#     # Connect to the SMTP server and send the email
#     with smtplib.SMTP_SSL(host, port) as server:
#         server.login(username, password)
#         server.send_message(message)
#
#     print("Email sent successfully!")
#
# if __name__ == '__main__':
#     # Environment variables retrieved from GitHub Secrets
#     EMAIL_HOST = os.getenv('EMAIL_HOST')
#     EMAIL_PORT = int(os.getenv('EMAIL_PORT'))  # Ensure port is an integer
#     EMAIL_USER = os.getenv('EMAIL_USER')
#     EMAIL_PASS = os.getenv('EMAIL_PASS')
#     EMAIL_TO = os.getenv('EMAIL_TO')
#
#     # Email content
#     SUBJECT = "Notification from GitHub Actions"
#     BODY = "This is a test email sent by GitHub Actions on push to the main branch."
#
#     send_email(EMAIL_HOST, EMAIL_PORT, EMAIL_USER, EMAIL_PASS, EMAIL_TO, SUBJECT, BODY)


def send_email(host, port, username, password, recipient, subject, body):
    # Create a MIME multipart message
    message = MIMEMultipart()
    message['From'] = username
    message['To'] = recipient
    message['Subject'] = subject

    # Add the email body to the message
    message.attach(MIMEText(body, 'plain'))

    # Connect to the SMTP server and send the email using starttls
    with smtplib.SMTP(host, port) as server:
        server.starttls()  # Upgrade the connection to secure
        server.login(username, password)
        server.send_message(message)

    print("Email sent successfully!")

if __name__ == '__main__':
    # Environment variables retrieved from GitHub Secrets
    EMAIL_HOST = os.getenv('EMAIL_HOST')
    EMAIL_PORT = int(os.getenv('EMAIL_PORT'))  # Ensure port is an integer
    EMAIL_USER = os.getenv('EMAIL_USER')
    EMAIL_PASS = os.getenv('EMAIL_PASS')
    EMAIL_TO = os.getenv('EMAIL_TO')

    # Email content
    SUBJECT = "Notification from GitHub Actions"
    BODY = "This is a test email sent by GitHub Actions on push to the main branch."

    # Send the email
    send_email(EMAIL_HOST, EMAIL_PORT, EMAIL_USER, EMAIL_PASS, EMAIL_TO, SUBJECT, BODY)