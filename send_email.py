import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
import datetime
from config import LEGACY_EMAIL_IP, THC_EMAIL_IP, SMTP_PORT, SOURCE_EMAIL_ADDRESS, DESTINATION_EMAIL_ADDRESS, BASE_DIR, OUTPUT_BASE_DIR

class send_eamil():
    def __init__(self):
        self.monthFormat = datetime.datetime(2025, 5, 10).strftime("%b")
        self.dateFromat = datetime.datetime(2025, 5, 10).strftime("%d")
        self.smtp_servers = [LEGACY_EMAIL_IP, THC_EMAIL_IP]
        self.smtp_port = SMTP_PORT 
        self.from_email = SOURCE_EMAIL_ADDRESS
        self.to_email = DESTINATION_EMAIL_ADDRESS
        self.subject = f"Daily VRD Operation Dashboard {self.dateFromat}-{self.monthFormat}  Success"
        self.html_file_path = rf"{BASE_DIR}\email_template.html"
        
    def _email_sent_func(self):
        with open(self.html_file_path, "r", encoding="utf-8") as f:
            html_body = f.read()

        message = MIMEMultipart()
        message["From"] = self.from_email
        message["To"] = ", " .join(self.to_email)
        message["Subject"] = self.subject
        message.attach(MIMEText(html_body, "html"))
        
        attachment_path = rf"{OUTPUT_BASE_DIR}\Daily_Report_{self.monthFormat}.xlsx"
        
        if os.path.isfile(attachment_path):
            with open(attachment_path, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(attachment_path)}"'
            message.attach(part)
        else:
            print(f"Warning: Attachment file not found: {attachment_path}")

        for smtp_server in self.smtp_servers:
            try:
                with smtplib.SMTP(smtp_server, self.smtp_port, timeout=60) as server:
                    server.sendmail(self.from_email, self.to_email, message.as_string())
                print(f"Email sent successfully via {smtp_server}")
                break  
            except Exception as e:
                print(f"Failed to send via {smtp_server}: {e}")

def send_email_main():
    app = send_eamil()
    app._email_sent_func()
    
if __name__ == '__main__':
    send_email_main()
    print("Sending email successfully.")