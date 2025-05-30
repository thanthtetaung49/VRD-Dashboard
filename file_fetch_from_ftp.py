import ftplib
import datetime
import os
from config import INPUT_BASE_DIR, HOST, USER_NAME, PASSWORD, REMOTE_VRD_PATH, REMOTE_MIS_PATH_SALE_PATH

class FTPFileFetcher:
    def __init__(self, host, username, password):
        self.host = host
        self.username = username
        self.password = password
        self.ftp = ftplib.FTP(host)
        self.datetime_vrd = datetime.datetime.now().strftime("%d-%b-%Y")
        self.datetime_mis = datetime.datetime.now().strftime("%d%m%Y")
    
    def _dir_create_in_local(self, local_path, status):
        if status == 'vrd':
            local_path = f'{local_path}\\{self.datetime_vrd}'
        else:
            local_path = f'{local_path}\\{self.datetime_mis}'
            
        if not os.path.exists(local_path):
            os.makedirs(local_path)
            print(f"Directory {local_path} created.")
        else:
            print(f"Directory {local_path} already exists.")

    def connect(self):
        self.ftp.login(user=self.username, passwd=self.password)

    def fetch_file(self, remote_path, local_path, status):
        if status == 'vrd':
            local_path = f'{local_path}\\{self.datetime_vrd}'
            remote_path = f'{remote_path}/{self.datetime_vrd}'
            self.ftp.cwd(remote_path)
        else:
            local_path = f'{local_path}\\{self.datetime_mis}'
            remote_path = f'{remote_path}/{self.datetime_mis}'
            self.ftp.cwd(remote_path)
            
        files = self.ftp.nlst()
        
        for file in files:
            with open(f'{local_path}\\{file}', 'wb') as local_file:
                self.ftp.retrbinary(f'RETR {file}', local_file.write)
    
    def list_files(self,remote_vrd_path):
        self.ftp.cwd(f'{remote_vrd_path}/{self.datetime_vrd}')
        files = self.ftp.nlst()
        for file in files:
            print(file)
        
        print(self.datetime)

    def close(self):
        self.ftp.quit()
        
def ftp_file_fatch_main():
    host = HOST
    username = USER_NAME
    password = PASSWORD
    remote_vrd_path = REMOTE_VRD_PATH
    local_vrd_path = f'{INPUT_BASE_DIR}\\VRD_Files'
    remote_mis_packsale_path = REMOTE_MIS_PATH_SALE_PATH
    local_mis_pack_path = f'{INPUT_BASE_DIR}\\MIS_Pack_Sale'

    fetcher = FTPFileFetcher(host, username, password)
    
    # create folder in local
    fetcher._dir_create_in_local(local_vrd_path, status='vrd')
    fetcher._dir_create_in_local(local_mis_pack_path, status='mis')
    
    try:
        fetcher.connect()
        print("Connected to FTP server.")
        
        fetcher.fetch_file(remote_vrd_path, local_vrd_path, status='vrd')
        fetcher.fetch_file(remote_mis_packsale_path, local_mis_pack_path, status='mis')
        
        print(f"File fetched successfully and saved to {local_vrd_path}.")
    except ftplib.all_errors as e:
        print(f"FTP error: {e}")
    finally:
        fetcher.close()
        print("Connection closed.")
        
if __name__ == "__main__":
    ftp_file_fatch_main()