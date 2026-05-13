"""
Google Drive Integration – Smart Data Analyzer
===============================================
يراقب مجلداً محدداً في Google Drive ويُشغّل التحليل تلقائياً عند رصد ملف Excel جديد.

المتطلبات:
    pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client

الإعداد:
1. أنشئ مشروعاً في Google Cloud Console
2. فعّل Google Drive API
3. حمّل ملف credentials.json
4. شغّل البرنامج وأدخل رمز المصادقة
"""

import os, time, pickle, tempfile
from pathlib import Path
from googleapiclient.discovery import build
from googleapiclient.http      import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES         = ["https://www.googleapis.com/auth/drive.readonly"]
WATCHED_FOLDER = "YOUR_FOLDER_ID"   # ← ضع معرّف المجلد هنا
POLL_INTERVAL  = 60                  # فحص كل 60 ثانية
PROCESSED_FILE = "processed_files.txt"


def get_drive_service():
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as f:
            pickle.dump(creds, f)
    return build("drive", "v3", credentials=creds)


def list_xlsx_files(service, folder_id):
    query = (f"'{folder_id}' in parents "
             f"and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
             f"and trashed=false")
    results = service.files().list(q=query,
                                    fields="files(id,name,modifiedTime)",
                                    orderBy="modifiedTime desc").execute()
    return results.get("files", [])


def download_file(service, file_id, dest_path):
    request = service.files().get_media(fileId=file_id)
    with open(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()


def load_processed():
    if not os.path.exists(PROCESSED_FILE):
        return set()
    with open(PROCESSED_FILE) as f:
        return set(f.read().splitlines())


def save_processed(processed):
    with open(PROCESSED_FILE, "w") as f:
        f.write("\n".join(processed))


def watch_and_analyze():
    from smart_data_analyzer import SmartDataAnalyzer
    service   = get_drive_service()
    processed = load_processed()
    print(f"👁️  مراقبة المجلد: {WATCHED_FOLDER}")
    print(f"⏱️  الفاصل الزمني: {POLL_INTERVAL} ثانية")

    while True:
        try:
            files = list_xlsx_files(service, WATCHED_FOLDER)
            for f in files:
                fid = f["id"]
                if fid in processed:
                    continue
                print(f"\n🆕  ملف جديد: {f['name']}")
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                    tmp_path = tmp.name
                download_file(service, fid, tmp_path)
                analyzer = SmartDataAnalyzer(tmp_path)
                analyzer.run()
                processed.add(fid)
                save_processed(processed)
                os.unlink(tmp_path)
        except Exception as e:
            print(f"⚠️  خطأ: {e}")
        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    watch_and_analyze()
