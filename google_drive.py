import os
import re
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload


# ============================================================
# CONFIGURATION
# ============================================================

SCOPES = ["https://www.googleapis.com/auth/drive"]

ROOT_FOLDER_NAME = "Resume Applications"

SERVICE_ACCOUNT_FILE = os.environ.get(
    "GOOGLE_APPLICATION_CREDENTIALS"
)


# ============================================================
# GOOGLE DRIVE CONNECTION
# ============================================================

def get_drive_service():
    if not SERVICE_ACCOUNT_FILE:
        raise RuntimeError(
            "GOOGLE_APPLICATION_CREDENTIALS environment variable is not set."
        )

    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise RuntimeError(
            f"Google credentials file not found: {SERVICE_ACCOUNT_FILE}"
        )

    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES
    )

    return build(
        "drive",
        "v3",
        credentials=credentials
    )


# ============================================================
# SANITIZE FOLDER NAME
# ============================================================

def sanitize_folder_name(name):
    name = name.strip()

    # Remove characters that can cause problems
    name = re.sub(r'[\\/:*?"<>|]', '', name)

    # Replace multiple spaces with one
    name = re.sub(r'\s+', ' ', name)

    return name.strip()


# ============================================================
# EXTRACT COMPANY NAME FROM JD
# ============================================================

def extract_company_name(jd_path):
    if not os.path.exists(jd_path):
        raise FileNotFoundError(
            f"JD file not found: {jd_path}"
        )

    with open(
        jd_path,
        "r",
        encoding="utf-8"
    ) as file:
        jd_text = file.read()

    if not jd_text.strip():
        raise ValueError("JD file is empty.")

    # Try common JD formats first
    patterns = [
        r'^\s*\*\*(.+?)\*\*\s*$',
        r'^\s*Company\s*:\s*(.+?)\s*$',
        r'^\s*Company Name\s*:\s*(.+?)\s*$',
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            jd_text,
            re.MULTILINE | re.IGNORECASE
        )

        if match:
            company = match.group(1).strip()

            if company:
                return sanitize_folder_name(company)

    # Fallback:
    # Look at the first few non-empty lines
    lines = [
        line.strip()
        for line in jd_text.splitlines()
        if line.strip()
    ]

    for line in lines[:10]:
        cleaned = re.sub(r'^\[|\]$', '', line)
        cleaned = re.sub(r'\*\*', '', cleaned)
        cleaned = cleaned.strip()

        # Skip obvious labels
        if cleaned.lower() in {
            "job description",
            "job details",
            "description",
            "responsibilities",
            "requirements"
        }:
            continue

        if len(cleaned) <= 100:
            return sanitize_folder_name(cleaned)

    raise ValueError(
        "Could not determine company name from JD."
    )


# ============================================================
# FIND FOLDER
# ============================================================

def find_folder(drive_service, folder_name, parent_id=None):
    query_parts = [
        f"name = '{folder_name.replace(chr(39), chr(92) + chr(39))}'",
        "mimeType = 'application/vnd.google-apps.folder'",
        "trashed = false"
    ]

    if parent_id:
        query_parts.append(
            f"'{parent_id}' in parents"
        )

    query = " and ".join(query_parts)

    response = (
        drive_service.files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id, name)",
            pageSize=10
        )
        .execute()
    )

    files = response.get("files", [])

    if files:
        return files[0]

    return None


# ============================================================
# CREATE FOLDER
# ============================================================

def create_folder(drive_service, folder_name, parent_id=None):
    metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder"
    }

    if parent_id:
        metadata["parents"] = [parent_id]

    folder = (
        drive_service.files()
        .create(
            body=metadata,
            fields="id, name"
        )
        .execute()
    )

    print(
        f"Created folder: {folder['name']}"
    )

    return folder


# ============================================================
# FIND OR CREATE FOLDER
# ============================================================

def find_or_create_folder(
    drive_service,
    folder_name,
    parent_id=None
):
    folder = find_folder(
        drive_service,
        folder_name,
        parent_id
    )

    if folder:
        print(
            f"Using existing folder: {folder['name']}"
        )
        return folder

    return create_folder(
        drive_service,
        folder_name,
        parent_id
    )


# ============================================================
# FIND FILE INSIDE FOLDER
# ============================================================

def find_file(
    drive_service,
    file_name,
    folder_id
):
    escaped_name = file_name.replace(
        "'",
        "\\'"
    )

    query = (
        f"name = '{escaped_name}' "
        f"and '{folder_id}' in parents "
        f"and trashed = false"
    )

    response = (
        drive_service.files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id, name)",
            pageSize=10
        )
        .execute()
    )

    files = response.get("files", [])

    if files:
        return files[0]

    return None


# ============================================================
# UPLOAD OR UPDATE FILE
# ============================================================

def upload_or_update_file(
    drive_service,
    local_file,
    drive_file_name,
    folder_id,
    mime_type
):
    if not os.path.exists(local_file):
        raise FileNotFoundError(
            f"Local file not found: {local_file}"
        )

    existing_file = find_file(
        drive_service,
        drive_file_name,
        folder_id
    )

    media = MediaFileUpload(
        local_file,
        mimetype=mime_type,
        resumable=True
    )

    if existing_file:
        print(
            f"Updating existing file: {drive_file_name}"
        )

        updated_file = (
            drive_service.files()
            .update(
                fileId=existing_file["id"],
                media_body=media,
                fields="id, name, webViewLink"
            )
            .execute()
        )

        return updated_file

    print(
        f"Uploading new file: {drive_file_name}"
    )

    metadata = {
        "name": drive_file_name,
        "parents": [folder_id]
    }

    uploaded_file = (
        drive_service.files()
        .create(
            body=metadata,
            media_body=media,
            fields="id, name, webViewLink"
        )
        .execute()
    )

    return uploaded_file


# ============================================================
# MAIN
# ============================================================

def main():
    if len(sys.argv) != 4:
        print(
            "Usage:"
        )
        print(
            "python3 google_drive.py "
            "<jd.txt> <final.pdf> <root_folder_id>"
        )
        sys.exit(1)

    jd_path = sys.argv[1]
    pdf_path = sys.argv[2]
    root_folder_id = sys.argv[3]

    print("=" * 60)
    print("GOOGLE DRIVE UPLOAD")
    print("=" * 60)

    # --------------------------------------------------------
    # 1. Extract company name
    # --------------------------------------------------------

    company_name = extract_company_name(
        jd_path
    )

    print(
        f"Company: {company_name}"
    )

    # --------------------------------------------------------
    # 2. Connect to Google Drive
    # --------------------------------------------------------

    drive_service = get_drive_service()

    # --------------------------------------------------------
    # 3. Find/Create company folder
    # --------------------------------------------------------

    company_folder = find_or_create_folder(
        drive_service,
        company_name,
        root_folder_id
    )

    company_folder_id = company_folder["id"]

    print(
        f"Company Folder ID: {company_folder_id}"
    )

    # --------------------------------------------------------
    # 4. Upload/Update jd.txt
    # --------------------------------------------------------

    jd_file = upload_or_update_file(
        drive_service,
        jd_path,
        "jd.txt",
        company_folder_id,
        "text/plain"
    )

    print(
        f"JD uploaded: {jd_file['name']}"
    )

    # --------------------------------------------------------
    # 5. Upload/Update Final Resume PDF
    # --------------------------------------------------------

    pdf_file = upload_or_update_file(
        drive_service,
        pdf_path,
        "Final_Resume.pdf",
        company_folder_id,
        "application/pdf"
    )

    print(
        f"Resume uploaded: {pdf_file['name']}"
    )

    # --------------------------------------------------------
    # Done
    # --------------------------------------------------------

    print("=" * 60)
    print("GOOGLE DRIVE UPLOAD COMPLETED")
    print("=" * 60)

    if pdf_file.get("webViewLink"):
        print(
            f"Drive Resume Link: {pdf_file['webViewLink']}"
        )


if __name__ == "__main__":
    main()
