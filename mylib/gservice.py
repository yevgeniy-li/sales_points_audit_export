from __future__ import print_function
import os
import time
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# import mylib.stdfunc as myfunc

class GService:
    def __init__(self, scopes):
        self.scopes = scopes
        self.creds = None

    def auth(self, token_json):
        creds = None
        token_pickle = os.path.join("token", "token.pickle")
        if os.path.exists(token_pickle):
            with open(token_pickle, "rb") as token:
                creds = pickle.load(token)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(token_json, self.scopes)
                creds = flow.run_local_server(port=0)
            with open(token_pickle, "wb") as token:
                pickle.dump(creds, token)
        self.creds = creds

    def get_spreadsheets(self):
        if self.creds is None:
            raise Exception("Authentication required!")
        service_sheets = build("sheets", "v4", credentials=self.creds)
        return service_sheets.spreadsheets()

    def get_service_drive(self):
        if self.creds is None:
            raise Exception("Authentication required!")
        service_drive = build("drive", "v3", credentials=self.creds)
        return service_drive

    def download_media(self, service_drive, file_id, dest_file_path):
        try_max_count = 4
        try_delay = 5
        try_count = 0
        done = False
        while (not done) and (try_count < try_max_count):
            try:
                request = service_drive.files().get_media(fileId=file_id)
                file_obj = open(dest_file_path, "wb")
                downloader = MediaIoBaseDownload(file_obj, request)
                while done is False:
                    done = downloader.next_chunk()
            except Exception as err:
                if try_count == try_max_count - 1: # была последняя попытка
                    raise err
                else:
                    time.sleep(try_delay)
                    print("ERROR: " + err)
                    print("Ошибка при скачивании файла. Повторная попытка (%d)..." % (try_count + 2)) # показываем номер следующей попытки
            try_count += 1
