import os
import keyring

INPUT_BASE_DIR = f"{os.path.dirname(os.path.abspath(__file__))}\\Input"
OUTPUT_BASE_DIR = f"{os.path.dirname(os.path.abspath(__file__))}\\output"
BASE_DIR = f"{os.path.dirname(os.path.abspath(__file__))}"

# sftp server configuration
HOST='10.84.83.30'
USER_NAME='rpaproduser1'
PASSWORD=keyring.get_password('FTP_PASSWORD', 'rpaproduser1')
REMOTE_VRD_PATH='/Power_Automate_Reports/Daily_VRD_Call_Compilation/Input'
REMOTE_MIS_PATH_SALE_PATH='/Robotic Process Automation/Daily_RD_Operations_Dashboard/MIS_Report'

# email configuration
LEGACY_EMAIL_IP="10.95.62.43"
THC_EMAIL_IP="10.95.76.165"
SMTP_PORT=25
SOURCE_EMAIL_ADDRESS="powerautomate.user@aml.mobi"
DESTINATION_EMAIL_ADDRESS=["thanthtet.aung@aml.mobi", "ayechan.thu@aml.mobi"]