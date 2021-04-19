import os
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file
from datetime import date
import datetime
import fnmatch
import mimetypes
from dotenv import load_dotenv
from pathlib import Path

env_path = Path('.') / 'credentials.env'
load_dotenv(dotenv_path=env_path)


class GoogleDrive:
    def __init__(self, credsPath: str = None):
        self.store = file.Storage(credsPath)
        self.credz = self.store.get()
        self.SCOPES = 'http://www.googleapis.com/auth/drive.readonly'
        self.drive_service = build('drive', 'v2', http=self.credz.authorize(Http()))

    def list_item_id_in_folder(self, folderId: str):
        itemIdList = []
        items = self.drive_service.children().list(folderId=folderId).execute()
        itemNum = 0
        keep_going = True
        while keep_going:
            try:
                itemIdList.append(items['items'][itemNum]['id'])
                itemNum += 1
            except:
                keep_going = False
        return itemIdList

    def list_item_names_in_folder(self, folderId: str):
        itemNameList = []
        itemNames = self.drive_service.files().list(q=f" '{folderId}' in parents", supportsAllDrives=True,
                                                    corpora='allDrives',
                                                    includeItemsFromAllDrives=True).execute()
        itemNum = 0
        keep_going = True
        while keep_going:
            try:
                itemNameList.append(itemNames['items'][itemNum]['title'])
                itemNum += 1
            except:
                keep_going = False
        return itemNameList

    def create_folder(self, folderId: str, folderName: str):
        body = {'title': folderName, 'mimeType': "application/vnd.google-apps.folder"}
        body['parents'] = [{'id': folderId}]
        self.drive_service.files().insert(body=body, supportsAllDrives=True).execute()

    def upload_csv(self, filepattern: str = None, localFolder: str = None, folderId: str = None):
        os.chdir(localFolder)
        allItems = os.listdir(localFolder)
        for item in allItems:
            itemPath = os.path.abspath(item)
            if fnmatch.fnmatch(item, '*' + filepattern + '*'):
                try:
                    self.drive_service.files().insert(convert=False, body={'title': os.path.basename(item), "parents": [
                        {"kind": "drive#fileLink", "id": folderId}],
                                                                           "mimeType": "text/csv"},
                                                      media_body=itemPath, fields='mimeType,exportLinks',
                                                      supportsAllDrives=True).execute()
                except:
                    mimetypes.add_type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.xlsx')
                    self.drive_service.files().insert(convert=True, body={'title': os.path.basename(item), "parents": [
                        {"kind": "drive#fileLink", "id": folderId}],
                                                                          "mimeType": "application/vnd.ms-excel"},
                                                      media_body=itemPath, fields='mimeType,exportLinks',
                                                      supportsAllDrives=True).execute()


if __name__ == '__main__':
    # <-example->

    drive_service = GoogleDrive(
        credsPath=r"C:\your path\mycreds.txt")
    # <-upload csv->
    drive_service.upload_csv(filepattern='Square',
                             localFolder=r'C:\Users\B\Desktop\Download From Email Automation\2021-03-21',
                             folderId='1vSD8K67Q5ElwW2fLHBYkKQgnbVN0MGAv')
    # <-upload csv->

    # <-list item names in folder->
    filenames = drive_service.list_item_names_in_folder('1vSD8K67Q5ElwW2fLHBYkKQgnbVN0MGAv')
    print(filenames)
    # <-list item names in folder->

    # <-create folder->
    drive_service.create_folder('1vSD8K67Q5ElwW2fLHBYkKQgnbVN0MGAv', 'test folder')
    # <-create folder->

    # <-example->
