import streamlit as st
import imaplib
import email
import os
import zipfile
import shutil
from email.header import decode_header
from datetime import date

TEMP_FOLDER = "backup_data"
ZIP_FILE = "email_backup.zip"

# Clear temp folder
def clear_folder(folder):
    if os.path.exists(folder):
        shutil.rmtree(folder)
    os.makedirs(folder)

# Clean filename but preserve extension
def clean_filename(filename):
    if "." in filename:
        name, ext = filename.rsplit(".", 1)
        ext = "." + ext
    else:
        name = filename
        ext = ""
    clean_name = "".join(c if c.isalnum() else "_" for c in name)
    return f"{clean_name}{ext}"

# Fetch emails + attachments from a folder with date range
def fetch_and_save_emails(mail, folder, start_date, end_date):
    try:
        mail.select(f'"{folder}"')

        imap_from = start_date.strftime("%d-%b-%Y")
        imap_to = end_date.strftime("%d-%b-%Y")

        status, messages = mail.search(None, f'(SINCE "{imap_from}" BEFORE "{imap_to}")')
        email_ids = messages[0].split()
        total_emails = len(email_ids)

        if total_emails == 0:
            st.info(f"📭 No emails in folder `{folder}` for this date range.")
            return

        progress_bar = st.progress(0)
        display_limit = 10
        display_logs = []

        for i, eid in enumerate(email_ids, 1):
            _, msg_data = mail.fetch(eid, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject, encoding = decode_header(msg.get("Subject"))[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding or "utf-8", errors="ignore")
            subject = subject or "no_subject"

            filename_base = f"{clean_filename(folder)}_{i:04d}_{clean_filename(subject)}"
            eml_path = os.path.join(TEMP_FOLDER, f"{filename_base}.eml")
            with open(eml_path, "wb") as f:
                f.write(raw_email)

            # Attachments
            attachment_counter = 1
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition", "")).lower()
                filename = part.get_filename()

                if part.get_content_maintype() == 'multipart':
                    continue

                if filename or 'attachment' in content_disposition:
                    if not filename:
                        ext = content_type.split("/")[-1]
                        if not ext or len(ext) > 5:
                            ext = "bin"
                        filename = f"file.{ext}"

                    filename = clean_filename(filename)
                    unique_filename = f"{filename_base}_{attachment_counter:02d}_{filename}"
                    att_path = os.path.join(TEMP_FOLDER, unique_filename)
                    with open(att_path, "wb") as af:
                        af.write(part.get_payload(decode=True))
                    attachment_counter += 1

            # UI logs
            display_logs.append(subject)
            if len(display_logs) > display_limit:
                display_logs.pop(0)
            st.write(f"📂 Folder `{folder}` progress: {i}/{total_emails} emails processed")
            st.write("📩 Last processed emails:")
            for msg_log in display_logs:
                st.write(f" - {msg_log}")
            progress_bar.progress(i / total_emails)

        st.success(f"✅ Folder `{folder}` completed: {total_emails} emails processed")

    except Exception as e:
        st.warning(f"⚠️ Error in `{folder}`: {e}")

# Streamlit UI
st.set_page_config(page_title="Email Backup Tool", layout="centered")
st.title("📥 Email Backup (EML + Attachments + Date Range)")

with st.form("login_form"):
    email_user = st.text_input("✉️ Email Address")
    email_pass = st.text_input("🔒 Password / App Password", type="password")
    imap_server = st.text_input("🌐 IMAP Server", value="imap.gmail.com")

    # Date Range
    start_date = st.date_input("📅 Start Date", value=date(2025,1,1))
    end_date = st.date_input("📅 End Date", value=date(2025,2,28))

    submitted = st.form_submit_button("🚀 Start Backup")

if submitted:
    try:
        st.info("🔌 Connecting to mail server...")
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_user, email_pass)

        status, folders = mail.list()
        all_folders = []
        for folder_line in folders:
            parts = folder_line.decode().split(' "/" ')
            if len(parts) == 2:
                folder_name = parts[1].replace('"', '')
                all_folders.append(folder_name)

        if not all_folders:
            st.warning("📭 No folders found!")
        else:
            clear_folder(TEMP_FOLDER)
            st.success(f"✅ Connected. Found {len(all_folders)} folders.")
            for folder in all_folders:
                fetch_and_save_emails(mail, folder, start_date, end_date)

            mail.logout()

            # Create ZIP
            with zipfile.ZipFile(ZIP_FILE, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files in os.walk(TEMP_FOLDER):
                    for file in files:
                        zipf.write(os.path.join(root, file), arcname=file)

            with open(ZIP_FILE, "rb") as f:
                st.download_button("⬇️ Download ZIP", f, file_name="email_backup.zip")

            st.success("🎉 Backup completed successfully!")

    except Exception as e:
        st.error(f"❌ Error: {e}")