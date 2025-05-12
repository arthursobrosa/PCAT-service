import os
import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from .modules.distributor_data import load_acronyms


def _get_drive_service():
    service_account_file = os.path.join(os.path.dirname(__file__), "data")

    credentials = service_account.Credentials.from_service_account_file(
        service_account_file,
        scopes=['https://www.googleapis.com/auth/drive']
    )

    return build('drive', 'v3', credentials=credentials)


def download_file_from_drive(file_id: str, destination_path: str):
    service = _get_drive_service()
    request = service.files().get_media(fileId=file_id)

    fh = io.FileIO(destination_path, 'wb')

    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()
        print(f"Download {int(status.progress() * 100)}%.")

    print(f"Arquivo salvo em {destination_path}")


def update_file_on_drive(file_id: str, local_file_path: str):
    service = _get_drive_service()

    media = MediaFileUpload(local_file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    updated_file = service.files().update(
        fileId=file_id,
        media_body=media
    ).execute()

    print(f"Arquivo atualizado no Google Drive: {updated_file['name']}")