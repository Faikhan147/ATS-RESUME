import os
import re
import sys

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload


# ============================================================
# CONFIGURATION
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/drive"
]

TOKEN_FILE = os.environ.get(
    "GOOGLE_DRIVE_TOKEN_FILE"
)


# ============================================================
# GOOGLE DRIVE CONNECTION
# ============================================================

def get_drive_service():
    if not TOKEN_FILE:
        raise RuntimeError(
            "GOOGLE_DRIVE_TOKEN_FILE is not set."
        )

    if not os.path.exists(TOKEN_FILE):
        raise RuntimeError(
            f"Google OAuth token file not found: "
            f"{TOKEN_FILE}"
        )

    credentials = Credentials.from_authorized_user_file(
        TOKEN_FILE,
        SCOPES
    )

    if credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())

    return build(
        "drive",
        "v3",
        credentials=credentials
    )


# ============================================================
# SANITIZE NAME
# ============================================================

def sanitize_name(name):
    name = name.strip()

    name = re.sub(
        r'[\\/:*?"<>|]',
        '',
        name
    )

    name = re.sub(
        r'\s+',
        ' ',
        name
    )

    return name.strip()


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

    # --------------------------------------------------------
    # Common company-name formats
    # --------------------------------------------------------

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

                return sanitize_name(
                    company_name
                )

    # --------------------------------------------------------
    # Fallback: first meaningful line
    # --------------------------------------------------------

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

            return sanitize_name(
                cleaned
            )

    raise ValueError(
        "Could not determine company name from JD."
    )


# ============================================================
# EXTRACT JOB TITLE FROM JD
# ============================================================

def extract_job_title(jd_path):

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

    # --------------------------------------------------------
    # Common job-title formats
    # --------------------------------------------------------

    patterns = [
        r'^\s*Job Title\s*:\s*(.+?)\s*$',
        r'^\s*Job\s*Title\s*:\s*(.+?)\s*$',
        r'^\s*Position\s*:\s*(.+?)\s*$',
        r'^\s*Role\s*:\s*(.+?)\s*$'
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            jd_text,
            re.MULTILINE | re.IGNORECASE
        )

        if match:

            job_title = match.group(1).strip()

            if job_title:

                return sanitize_name(
                    job_title
                )

    # --------------------------------------------------------
    # LinkedIn-style JD fallback
    #
    # Example:
    #
    # **Palo Alto Networks**
    #
    # **DevOps Engineer**
    #
    # --------------------------------------------------------

    lines = [
        line.strip()
        for line in jd_text.splitlines()
        if line.strip()
    ]

    cleaned_lines = []

    for line in lines[:15]:

        cleaned = re.sub(
            r'\*\*',
            '',
            line
        ).strip()

        if cleaned:
            cleaned_lines.append(
                cleaned
            )

    # If first line is company,
    # second meaningful line is usually title.
    if len(cleaned_lines) >= 2:

        first_line = cleaned_lines[0].lower()

        ignored_lines = {
            "job description",
            "job details",
            "description",
            "responsibilities",
            "requirements"
        }

        if first_line not in ignored_lines:

            possible_title = (
                cleaned_lines[1]
            )

            if (
                len(possible_title) <= 100
                and possible_title.lower()
                not in ignored_lines
            ):

                return sanitize_name(
                    possible_title
                )

    raise ValueError(
        "Could not determine job title from JD."
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
            fields="files(id,name,webViewLink)",
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
        f"Created folder: {folder['name']}"
    )

    return folder


# ============================================================
# FIND OR CREATE FOLDER
# ============================================================

def find_or_create_folder(
    drive_service,
    folder_name,
    parent_id
):

    folder = find_folder(
        drive_service,
        folder_name,
        parent_id
    )

    if folder:

        print(
            f"Using existing folder: "
            f"{folder['name']}"
        )

        return folder

    return create_folder(
        drive_service,
        folder_name,
        parent_id
    )


# ============================================================
# FIND OR CREATE COMPANY FOLDER
# ============================================================

def find_or_create_company_folder(
    drive_service,
    company_name,
    root_folder_id
):

    return find_or_create_folder(
        drive_service,
        company_name,
        root_folder_id
    )


# ============================================================
# FIND OR CREATE JOB TITLE FOLDER
# ============================================================

def find_or_create_job_title_folder(
    drive_service,
    job_title,
    company_folder_id
):

    return find_or_create_folder(
        drive_service,
        job_title,
        company_folder_id
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
        "and trashed = false"
    )

    response = (
        drive_service.files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id,name,webViewLink)",
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
    # EXISTING FILE → UPDATE
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
    # NEW FILE → CREATE
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

        print("Usage:")
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
    # 1. Extract Company Name
    # --------------------------------------------------------

    company_name = extract_company_name(
        jd_path
    )

    print(
        f"Company Name: {company_name}"
    )

    # --------------------------------------------------------
    # 2. Extract Job Title
    # --------------------------------------------------------

    job_title = extract_job_title(
        jd_path
    )

    print(
        f"Job Title: {job_title}"
    )

    # --------------------------------------------------------
    # 3. Connect to Google Drive
    # --------------------------------------------------------

    drive_service = get_drive_service()

    # --------------------------------------------------------
    # 4. Find/Create Company Folder
    # --------------------------------------------------------

    company_folder = (
        find_or_create_company_folder(
            drive_service,
            company_name,
            root_folder_id
        )
    )

    company_folder_id = (
        company_folder["id"]
    )

    print(
        f"Company Folder ID: "
        f"{company_folder_id}"
    )

    # --------------------------------------------------------
    # 5. Find/Create Job Title Folder
    # --------------------------------------------------------

    job_title_folder = (
        find_or_create_job_title_folder(
            drive_service,
            job_title,
            company_folder_id
        )
    )

    job_title_folder_id = (
        job_title_folder["id"]
    )

    print(
        f"Job Title Folder ID: "
        f"{job_title_folder_id}"
    )

    # --------------------------------------------------------
    # 6. Upload/Update JD
    # --------------------------------------------------------

    jd_file = upload_or_update_file(
        drive_service,
        jd_path,
        "jd.txt",
        job_title_folder_id,
        "text/plain"
    )

    print(
        f"JD saved as: "
        f"{jd_file['name']}"
    )

    # --------------------------------------------------------
    # 7. KEEP ORIGINAL PDF FILENAME
    # --------------------------------------------------------

    pdf_filename = os.path.basename(
        final_pdf
    )

    print(
        f"PDF filename: "
        f"{pdf_filename}"
    )

    # --------------------------------------------------------
    # 8. Upload/Update PDF
    # --------------------------------------------------------

    pdf_file = upload_or_update_file(
        drive_service,
        final_pdf,
        pdf_filename,
        job_title_folder_id,
        "application/pdf"
    )

    print(
        f"PDF saved as: "
        f"{pdf_file['name']}"
    )

    # --------------------------------------------------------
    # 9. DONE
    # --------------------------------------------------------

    print("=" * 60)
    print("GOOGLE DRIVE UPLOAD COMPLETED")
    print("=" * 60)

    if company_folder.get(
        "webViewLink"
    ):

        print(
            f"Company Folder: "
            f"{company_folder['webViewLink']}"
        )

    if job_title_folder.get(
        "webViewLink"
    ):

        print(
            f"Job Title Folder: "
            f"{job_title_folder['webViewLink']}"
        )

    if pdf_file.get(
        "webViewLink"
    ):

        print(
            f"Resume PDF: "
            f"{pdf_file['webViewLink']}"
        )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()
