#!/usr/bin/env python3
"""
Pulls the latest Promised Land Cloudbeds exports directly from Google Drive
using a service account, so the GitHub Actions run doesn't depend on a
manually-committed copy in the repo.

Looks files up by name within the PA-PL AssetManagement folder (not by a
hardcoded file ID), since replacing a file in Drive can mean a new file ID.

Auth: reads the service account JSON from the GDRIVE_SERVICE_ACCOUNT_JSON
env var (the full JSON content, not a file path).
"""

import json
import os
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

PA_PL_FOLDER_ID = "1zfvnQ0YhEd3W1RgZtOBJ6m0gNfY0OJLm"

FILES_TO_FETCH = [
    "PA-PL_occupancy_2026ytd.xlsx",
    "PA-PL_occupancy_pace_150day.xlsx",
    "PA-PL_occupancy_2025.xlsx",
]


def get_drive_service():
    creds_json = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")
    if not creds_json:
        print("❌ GDRIVE_SERVICE_ACCOUNT_JSON not set")
        sys.exit(1)
    info = json.loads(creds_json)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    return build("drive", "v3", credentials=creds)


def fetch_file_by_name(service, filename, folder_id, dest_dir):
    query = (
        f"name = '{filename}' and '{folder_id}' in parents and trashed = false"
    )
    results = service.files().list(
        q=query, fields="files(id, name, modifiedTime)", orderBy="modifiedTime desc"
    ).execute()
    files = results.get("files", [])
    if not files:
        print(f"  ⚠️  Not found in Drive: {filename}")
        return False

    file_id = files[0]["id"]
    request = service.files().get_media(fileId=file_id)
    dest_path = os.path.join(dest_dir, filename)
    fh = io.FileIO(dest_path, "wb")
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.close()
    print(f"  ✅ Fetched: {filename} (modified {files[0]['modifiedTime']})")
    return True


def main():
    dest_dir = os.path.dirname(os.path.abspath(__file__))
    service = get_drive_service()
    print(f"📥 Fetching Promised Land exports from Drive folder {PA_PL_FOLDER_ID}...")
    any_failed = False
    for filename in FILES_TO_FETCH:
        ok = fetch_file_by_name(service, filename, PA_PL_FOLDER_ID, dest_dir)
        any_failed = any_failed or not ok
    if any_failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
