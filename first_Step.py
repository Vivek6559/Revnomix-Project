import win32com.client
import os
import subprocess
import sys

# ==========================================================
# CONFIGURATION
# ==========================================================

MAILBOX_NAME = "revseed@revnomix.com"
SENDER_EMAIL = "h5706@accor.com"

SAVE_FOLDER = (
    r"C:\Users\srishti sinha\Documents\H5706\H5706"
    r"\files (27)\stly_v2_pipeline\stly_v2"
)

WORKING_DIR = (
    r"C:\Users\srishti sinha\Documents\H5706\H5706"
    r"\files (27)\stly_v2_pipeline\stly_v2"
)

os.makedirs(SAVE_FOLDER, exist_ok=True)

# ==========================================================
# SUBJECT -> FILE NAME
# ==========================================================

SUBJECT_MAPPING = {
    "FOR01121 Past and Future Forecast - SFRS_HNF":
        "H5706_hnf",

    "FOR01109 Forecast report by Date - SFRS_MS_FORECAST":
        "H5706_ms_future_data",

    "FOR01109 Forecast report by Date - SFRS_RM_FORECAST":
        "H5706_rm_future_data",

    "STA01144 Reservation Statistics 1 - SFRS_MS_STATISTICS":
        "H5706_ms_history_data",

    "STA01144 Reservation Statistics 1 - SFRS_RM_STATISTICS":
        "H5706_rm_history_data"
}

# ==========================================================
# COLUMN REPLACEMENTS
# ==========================================================

FILES_TO_UPDATE = {
    "H5706_rm_history_data": {
        "MASTER_VALUE": "MARKET_CODE"
    },
    "H5706_ms_history_data": {
        "MASTER_VALUE": "MARKET_CODE"
    },
    "H5706_rm_future_data": {
        "MARKET_CODE_SEQ": "CF_MARKET_CODE_SEQ"
    },
    "H5706_ms_future_data": {
        "MARKET_CODE_SEQ": "CF_MARKET_CODE_SEQ"
    }
}

# ==========================================================
# CONNECT TO OUTLOOK
# ==========================================================

outlook = win32com.client.Dispatch(
    "Outlook.Application"
).GetNamespace("MAPI")

mailbox = None

for folder in outlook.Folders:
    if folder.Name.lower() == MAILBOX_NAME.lower():
        mailbox = folder
        break

if mailbox is None:
    raise Exception(f"Mailbox '{MAILBOX_NAME}' not found.")

print(f"Connected to mailbox: {MAILBOX_NAME}")

# ==========================================================
# OPEN INBOX
# ==========================================================

inbox = mailbox.Folders["Inbox"]
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

downloaded = set()

print("\nSearching emails...\n")

# ==========================================================
# DOWNLOAD REQUIRED FILES
# ==========================================================

for message in messages:

    try:
        if message.Class != 43:
            continue

        # Get sender
        sender = ""

        try:
            sender = message.SenderEmailAddress.lower()
        except:
            pass

        # Exchange mailbox support
        if not sender or "/o=" in sender:
            try:
                ex_user = message.Sender.GetExchangeUser()
                if ex_user:
                    sender = ex_user.PrimarySmtpAddress.lower()
            except:
                pass

        if sender != SENDER_EMAIL.lower():
            continue

        subject = message.Subject.strip()

        if subject not in SUBJECT_MAPPING:
            continue

        if subject in downloaded:
            continue

        print(f"Found Email: {subject}")

        for attachment in message.Attachments:

            original_name = attachment.FileName
            extension = os.path.splitext(original_name)[1]

            new_name = SUBJECT_MAPPING[subject] + extension
            save_path = os.path.join(SAVE_FOLDER, new_name)

            # Replace existing file
            if os.path.exists(save_path):
                os.remove(save_path)

            attachment.SaveAsFile(save_path)

            print(f"Downloaded: {new_name}")

            downloaded.add(subject)
            break

        if len(downloaded) == len(SUBJECT_MAPPING):
            break

    except Exception as e:
        print("Error:", e)

# ==========================================================
# CHECK DOWNLOADS
# ==========================================================

print("\n===============================")
print("DOWNLOAD SUMMARY")
print("===============================")

for subject in downloaded:
    print("✓", SUBJECT_MAPPING[subject])

missing = set(SUBJECT_MAPPING) - downloaded

if missing:
    print("\nMissing Emails:")
    for subject in missing:
        print("-", subject)

# ==========================================================
# CHANGE COLUMN NAMES IN DOWNLOADED FILES
# ==========================================================

print("\n===============================")
print("CHANGING COLUMN NAMES")
print("===============================")

processed_files = set()

for base_name, changes in FILES_TO_UPDATE.items():

    file_path = None

    # Find downloaded file regardless of extension
    for file in os.listdir(SAVE_FOLDER):

        name, ext = os.path.splitext(file)

        if name.lower() == base_name.lower():
            file_path = os.path.join(SAVE_FOLDER, file)
            break

    if file_path is None:
        print(f"File not found: {base_name}")
        continue

    print(f"\nProcessing: {os.path.basename(file_path)}")

    try:
        # Read existing downloaded file
        with open(
            file_path,
            "r",
            encoding="utf-8"
        ) as f:
            lines = f.readlines()

        if not lines:
            print("File is empty.")
            continue

        # Header
        header = lines[0].rstrip("\r\n").split("|")

        changed = False

        # Replace column names
        for old_name, new_name in changes.items():

            if old_name in header:

                index = header.index(old_name)
                header[index] = new_name

                print(
                    f"Changed column {index + 1}: "
                    f"'{old_name}' -> '{new_name}'"
                )

                changed = True

            else:
                print(f"Column '{old_name}' not found.")

        # Replace first line with modified header
        lines[0] = "|".join(header) + "\n"

        # ==================================================
        # SAVE BACK TO SAME FILE
        # ==================================================

        with open(
            file_path,
            "w",
            encoding="utf-8",
            newline=""
        ) as f:
            f.writelines(lines)

        print("File replaced successfully.")

        processed_files.add(base_name)

    except Exception as e:
        print(f"Error processing {base_name}: {e}")

# ==========================================================
# VERIFY ALL REQUIRED FILES
# ==========================================================

required_files = set(FILES_TO_UPDATE.keys())

if not required_files.issubset(processed_files):

    missing_processing = required_files - processed_files

    print("\nCannot run STLY pipeline.")
    print("Files not processed:")

    for file in missing_processing:
        print("-", file)

    sys.exit(1)

print("\nAll required files downloaded and column names replaced.")

# ==========================================================
# RUN STLY PIPELINE ONLY AFTER FILE REPLACEMENT
# ==========================================================

print("\n===============================")
print("RUNNING STLY PIPELINE")
print("===============================")

result = subprocess.run(
    [
        sys.executable,
        "stly_pipeline.py",
        "--force-rerun"
    ],
    cwd=WORKING_DIR,
    capture_output=True,
    text=True
)

# ==========================================================
# PIPELINE RESULT
# ==========================================================

print("\nReturn Code:", result.returncode)

print("\nSTDOUT:")
print(result.stdout)

if result.stderr:
    print("\nSTDERR:")
    print(result.stderr)

# ==========================================================
# FINAL STATUS
# ==========================================================

if result.returncode == 0:
    print("\n===================================")
    print("PROCESS COMPLETED SUCCESSFULLY")
    print("===================================")
else:
    print("\n===================================")
    print("STLY PIPELINE FAILED")
    print("===================================")