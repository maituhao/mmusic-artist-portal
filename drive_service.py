import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import httplib2
import google_auth_httplib2

SCOPES = ['https://www.googleapis.com/auth/drive']
TOKEN_PATH = os.path.join(os.path.dirname(__file__), 'token.json')
CREDS_PATH = os.path.join(os.path.dirname(__file__), 'credentials.json')

def get_credentials():
    import json
    creds = None
    token_str = os.getenv("GOOGLE_TOKEN_JSON")
    
    if token_str:
        info = json.loads(token_str)
        creds = Credentials.from_authorized_user_info(info, SCOPES)
    elif os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise Exception("Chưa xác thực Drive. Hãy cấp quyền qua file token hoặc biến môi trường!")
    return creds

def _build_http(creds):
    """Create an authorized HTTP object with a 300-second timeout."""
    http = httplib2.Http(timeout=300)
    return google_auth_httplib2.AuthorizedHttp(creds, http=http)

def get_drive_service(creds=None):
    if creds is None:
        creds = get_credentials()
    return build('drive', 'v3', http=_build_http(creds))

def get_sheets_service(creds=None):
    if creds is None:
        creds = get_credentials()
    return build('sheets', 'v4', http=_build_http(creds))

def read_sheet_data(service, spreadsheet_id, range_name):
    try:
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        return result.get('values', [])
    except Exception as e:
        print(f"Error reading sheet {spreadsheet_id}: {e}")
        return []

def find_folder(service, name, parent_id):
    query = f"name='{name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=query, spaces='drive', fields='files(id, name, webViewLink)').execute()
    items = results.get('files', [])
    return items[0] if items else None

def create_folder(service, name, parent_id):
    folder_metadata = {
        'name': name,
        'parents': [parent_id],
        'mimeType': 'application/vnd.google-apps.folder'
    }
    folder = service.files().create(body=folder_metadata, fields='id, webViewLink').execute()
    return folder

def upload_file(service, file_path, name, parent_id, mimetype=None):
    file_metadata = {
        'name': name,
        'parents': [parent_id]
    }
    
    # Tự động chuyển đổi sang Google Sheets/Docs và bỏ đuôi mở rộng
    if name.endswith('.xlsx'):
        file_metadata['mimeType'] = 'application/vnd.google-apps.spreadsheet'
        file_metadata['name'] = name[:-5]
    elif name.endswith('.docx'):
        file_metadata['mimeType'] = 'application/vnd.google-apps.document'
        file_metadata['name'] = name[:-5]

    media = MediaFileUpload(file_path, mimetype=mimetype, resumable=True)
    file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    return file

def share_folder(service, file_id, domain, artist_email):
    # Domain sharing removed to keep default access as Restricted (Hạn chế)

    # Share with artist email(s) (Editor) — supports ; separator
    if artist_email:
        emails = [e.strip() for e in artist_email.split(';') if e.strip()]
        for email in emails:
            try:
                user_permission = {
                    'type': 'user',
                    'role': 'writer',
                    'emailAddress': email
                }
                service.permissions().create(
                    fileId=file_id,
                    body=user_permission,
                    sendNotificationEmail=False
                ).execute()
            except Exception as e:
                print(f"Could not set user permission for {email}: {e}")
