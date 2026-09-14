import os
import re
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload


# ============================================================
# CONFIGURATION
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/drive"
]

SERVICE_ACCOUNT_FILE = os.environ.get(
    "GOOGLE_APPLICATION_CREDENTIALS"
)


# ============================================================
# GOOGLE DRIVE CONNECTION
# ============================================================

def get_drive_service():
    if not SERVICE_ACCOUNT_FILE:
        raise RuntimeError(
            "GOOGLE_APPLICATION_CREDENTIALS is not set."
        )

    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise RuntimeError(
            f"Google credentials file not found: "
            f"{SERVICE_ACCOUNT_FILE}"
        )

    credentials = (
        service_account.Credentials
        .from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=SCOPES
        )
    )

    return build(
        "drive",
        "v3",
        credentials=credentials
    )


# ============================================================
# SANITIZE COMPANY NAME
# ============================================================

def sanitize_company_name(company_name):
    company_name = company_name.strip()

    company_name = re.sub(
        r'[\\/:*?"<>|]',
        '',
        company_name
    )

    company_name = re.sub(
        r'\s+',
        ' ',
        company_name
    )

    return company_name.strip()


# ============================================================
# EXTRACT COMPANY NAME FROM JD
# ============================================================

def extract_company_name(jd_path):
    with open(
        jd_path,
        "r",
        encoding="utf-8"
    ) as file:
        jd_text = file.read()

    if not jd_text.strip():
        raise ValueError(
            "JD file is empty."
        )

    # Common formats
    patterns = [
        r'^\s*\*\*(.+?)\*\*\s*$',
        r'^\s*Company\s*:\s*(.+?)\s*$',
        r'^\s*Company Name\s*:\s*(.+?)\s*$'
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            jd_text,
            re.MULTILINE | re.IGNORECASE
        )

        if match:
            company_name = match.group(1).strip()

            if company_name:
                return sanitize_company_name(
                    company_name
                )

    # Fallback: first meaningful lines
    lines = [
        line.strip()
        for line in jd_text.splitlines()
        if line.strip()
    ]

    ignored_lines = {
        "job description",
        "job details",
        "description",
        "responsibilities",
        "requirements"
    }

    for line in lines[:10]:
        cleaned = re.sub(
            r'\*\*',
            '',
            line
        ).strip()

        if cleaned.lower() in ignored_lines:
            continue

        if len(cleaned) <= 100:
            return sanitize_company_name(
                cleaned
            )

    raise ValueError(
        "Could not determine company name from JD."
    )


# ============================================================
# FIND FOLDER
# ============================================================

def find_folder(
    drive_service,
    folder_name,
    parent_id
):
    escaped_name = folder_name.replace(
        "'",
        "\\'"
    )

    query = (
        f"name = '{escaped_name}' "
        "and mimeType = "
        "'application/vnd.google-apps.folder' "
        "and trashed = false "
        f"and '{parent_id}' in parents"
    )

    response = (
        drive_service.files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id,name)",
            pageSize=10
        )
        .execute()
    )

    folders = response.get(
        "files",
        []
    )

    if folders:
        return folders[0]

    return None


# ============================================================
# CREATE FOLDER
# ============================================================

def create_folder(
    drive_service,
    folder_name,
    parent_id
):
    metadata = {
        "name": folder_name,
        "mimeType": (
            "application/vnd.google-apps.folder"
        ),
        "parents": [parent_id]
    }

    folder = (
        drive_service.files()
        .create(
            body=metadata,
            fields="id,name,webViewLink"
        )
        .execute()
    )

    print(
        f"Created company folder: "
        f"{folder['name']}"
    )

    return folder


# ============================================================
# FIND OR CREATE COMPANY FOLDER
# ============================================================

def find_or_create_company_folder(
    drive_service,
    company_name,
    root_folder_id
):
    folder = find_folder(
        drive_service,
        company_name,
        root_folder_id
    )

    if folder:
        print(
            f"Using existing company folder: "
            f"{folder['name']}"
        )

        return folder

    return create_folder(
        drive_service,
        company_name,
        root_folder_id
    )


# ============================================================
# FIND FILE INSIDE COMPANY FOLDER
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
        "and trashed = false"
    )

    response = (
        drive_service.files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id,name)",
            pageSize=10
        )
        .execute()
    )

    files = response.get(
        "files",
        []
    )

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
            f"File not found: {local_file}"
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

    # --------------------------------------------------------
    # Existing file → UPDATE
    # --------------------------------------------------------

    if existing_file:
        print(
            f"Updating existing file: "
            f"{drive_file_name}"
        )

        updated_file = (
            drive_service.files()
            .update(
                fileId=existing_file["id"],
                media_body=media,
                fields="id,name,webViewLink"
            )
            .execute()
        )

        return updated_file

    # --------------------------------------------------------
    # File doesn't exist → CREATE
    # --------------------------------------------------------

    print(
        f"Uploading new file: "
        f"{drive_file_name}"
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
            fields="id,name,webViewLink"
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
    final_pdf = sys.argv[2]
    root_folder_id = sys.argv[3]

    print("=" * 60)
    print("GOOGLE DRIVE RESUME UPLOAD")
    print("=" * 60)

    # --------------------------------------------------------
    # 1. Extract company name
    # --------------------------------------------------------

    company_name = extract_company_name(
        jd_path
    )

    print(
        f"Company Name: {company_name}"
    )

    # --------------------------------------------------------
    # 2. Connect to Google Drive
    # --------------------------------------------------------

    drive_service = get_drive_service()

    # --------------------------------------------------------
    # 3. Find/Create company folder
    # --------------------------------------------------------

    company_folder = (
        find_or_create_company_folder(
            drive_service,
            company_name,
            root_folder_id
        )
    )

    company_folder_id = company_folder["id"]

    print(
        f"Company Folder ID: "
        f"{company_folder_id}"
    )

    # --------------------------------------------------------
    # 4. Upload/Update JD
    # --------------------------------------------------------

    jd_file = upload_or_update_file(
        drive_service,
        jd_path,
        "jd.txt",
        company_folder_id,
        "text/plain"
    )

    print(
        f"JD saved as: {jd_file['name']}"
    )

    # --------------------------------------------------------
    # 5. Keep ORIGINAL PDF filename
    # --------------------------------------------------------

    pdf_filename = os.path.basename(
        final_pdf
    )

    print(
        f"PDF filename: {pdf_filename}"
    )

    # --------------------------------------------------------
    # 6. Upload/Update PDF
    # --------------------------------------------------------

    pdf_file = upload_or_update_file(
        drive_service,
        final_pdf,
        pdf_filename,
        company_folder_id,
        "application/pdf"
    )

    print(
        f"PDF saved as: {pdf_file['name']}"
    )

    # --------------------------------------------------------
    # DONE
    # --------------------------------------------------------

    print("=" * 60)
    print("GOOGLE DRIVE UPLOAD COMPLETED")
    print("=" * 60)

    if company_folder.get("webViewLink"):
        print(
            f"Company Folder: "
            f"{company_folder['webViewLink']}"
        )

    if pdf_file.get("webViewLink"):
        print(
            f"Resume PDF: "
            f"{pdf_file['webViewLink']}"
        )


if __name__ == "__main__":
    main()
