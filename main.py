import smtplib
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Debugging: Print current working directory
print("Current Working Directory:", os.getcwd())

# Change the working directory to the script's directory
script_dir = r'C:\Users\admin\OneDrive\Documents\emails'
os.chdir(script_dir)
print("Updated Working Directory:", os.getcwd())

# SMTP server configuration
SMTP_SERVER = 'smtp.gmail.com'  # Replace with your SMTP server
SMTP_PORT = 587                 # Replace with your SMTP port
SENDER_EMAIL = 'kanishkaarde99@gmail.com'  # Replace with your email
SENDER_PASSWORD = 'ercx xxzg hlox ozqh'      # Replace with your email password

# Function to send email
def send_email(receiver_email, subject, body, attachment_path):
    # Create the email
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = receiver_email
    msg['Subject'] = subject

    # Attach the body
    msg.attach(MIMEText(body, 'plain'))

    # Attach the file
    if attachment_path:
        try:
            with open(attachment_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
                msg.attach(part)
        except FileNotFoundError:
            print(f"Attachment file not found: {attachment_path}")
            return

    # Send the email
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, receiver_email, msg.as_string())
        server.quit()
        print(f"Email sent to {receiver_email}")
    except Exception as e:
        print(f"Failed to send email to {receiver_email}: {e}")

# Function to read user details from CSV
def read_users_from_csv(csv_file):
    users = []
    try:
        with open(csv_file, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                users.append(row)
    except FileNotFoundError:
        print(f"CSV file not found: {csv_file}")
    return users

# Main function
def main():
    # Use absolute paths for files
    csv_file = r'C:\Users\admin\OneDrive\Documents\emails\users.csv'
    attachment_path = r'C:\Users\admin\OneDrive\Documents\emails\attachment.txt'

    # Read user details from CSV
    users = read_users_from_csv(csv_file)

    # Send personalized emails
    for user in users:
        receiver_email = user['email']
        subject = f"Hello {user['name']}, here is your report"
        body = f"Dear {user['name']},\n\nPlease find the attached report.\n\nBest regards,\nYour Team"
        send_email(receiver_email, subject, body, attachment_path)

if __name__ == "__main__":
    main()
