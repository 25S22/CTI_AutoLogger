import win32com.client
import os
import pandas as pd
import tempfile
import time
import uuid
from datetime import datetime, time as dt_time

# =============================================================================
#  CONFIGURATION — Edit this section to customise your scan targets
# =============================================================================

# List of Outlook folder names (inside Inbox) to scan.
# Leave empty [] to skip folder-based scanning.
TARGET_FOLDERS = [
    "cti",
    # "threat-intel",   # <-- Add more folder names here
]

# List of sender email addresses to scan across your ENTIRE inbox.
# Any email from these senders (regardless of folder) will be checked.
# Leave empty [] to skip sender-based scanning.
TARGET_SENDERS = [
    # "analyst@threatfeed.com",
    # "reports@isac.org",        # <-- Add sender emails here
]

# Output master file
MASTER_FILE = "Master_IOC_Sheet.xlsx"

# IOC column headers to look for inside Excel attachments
IOC_SEARCH_TERMS = {
    'md5':    ['md-5', 'md5'],
    'sha1':   ['sha-1', 'sha1'],
    'sha256': ['sha-2', 'sha256'],
    'ip':     ['ip address', 'ip_address', "ip's", "ips"],
    'domain': ['domain', 'domains', 'host'],
    'url':    ['url', 'urls', 'link', 'links', 'uri'],
    'email':  ['email', 'e-mail', 'emails', 'mail', 'sender', 'recipient'],
}

# =============================================================================
#  HELPERS
# =============================================================================

def get_valid_date(prompt_text):
    while True:
        try:
            return datetime.strptime(input(prompt_text).strip(), "%Y-%m-%d").date()
        except ValueError:
            print("❌  Invalid format. Use YYYY-MM-DD.")


def find_header_row(df):
    """Scans first 10 rows to locate the real header row."""
    all_keywords = [kw for kwlist in IOC_SEARCH_TERMS.values() for kw in kwlist]

    current_cols = [str(c).lower().strip() for c in df.columns]
    if any(k in c for k in all_keywords for c in current_cols):
        return df

    for i, row in df.head(10).iterrows():
        row_values = [str(v).lower().strip() for v in row.values]
        if any(k in v for k in all_keywords for v in row_values):
            df_new = df.iloc[i + 1:].copy()
            df_new.columns = row_values
            return df_new
    return None


def extract_iocs_from_df(df_clean, email_row):
    """Copies values straight from header-matched columns — no filtering or validation."""
    found_any = False
    df_clean.columns = [str(c).lower().strip() for c in df_clean.columns]

    for ioc_type, keywords in IOC_SEARCH_TERMS.items():
        found_col = None
        for kw in keywords:
            found_col = next((c for c in df_clean.columns if kw in c), None)
            if found_col:
                break
        if not found_col:
            continue

        vals = df_clean[found_col].dropna().astype(str).str.strip().tolist()
        vals = [v for v in vals if v and v.lower() not in ('nan', 'none', '')]
        if vals:
            email_row[ioc_type].extend(vals)
            found_any = True

    return found_any


def process_message(message, temp_dir, new_data, label_source):
    """
    Inspects a single Outlook message for Excel attachments and extracts IOCs.
    `label_source` is a string like 'Folder: cti' or 'Sender: x@y.com' shown in output.
    """
    try:
        msg_dt = message.ReceivedTime.replace(tzinfo=None)
    except Exception:
        return

    count = message.Attachments.Count
    if count == 0:
        return

    sender_name  = getattr(message, 'SenderName', 'Unknown')
    sender_email = getattr(message, 'SenderEmailAddress', 'Unknown')

    email_row = {
        'Subject':      message.Subject,
        'Date':         str(msg_dt.date()),
        'Sender Name':  sender_name,
        'Sender Email': sender_email,
        'Source':       label_source,
        'md5': [], 'sha1': [], 'sha256': [],
        'ip': [], 'domain': [], 'url': [], 'email': [],
    }

    email_has_data = False

    for i in range(1, count + 1):
        try:
            att      = message.Attachments.Item(i)
            fname    = att.FileName.lower()

            if not fname.endswith(('.xlsx', '.xls')):
                continue

            unique_name = f"{uuid.uuid4()}_{att.FileName}"
            save_path   = os.path.join(temp_dir, unique_name)
            att.SaveAsFile(save_path)
            time.sleep(0.4)

            file_has_ioc = False

            with pd.ExcelFile(save_path) as xls:
                    df = pd.read_excel(xls, sheet_name=0)   # first sheet only
                    if not df.empty:
                        df_clean = find_header_row(df)
                        if df_clean is not None:
                            if extract_iocs_from_df(df_clean, email_row):
                                file_has_ioc   = True
                                email_has_data = True

            status = "✅  Extracted" if file_has_ioc else "⚠️   No matching headers in"
            print(f"   {status}: {att.FileName}")

        except Exception as e:
            print(f"   ❌  Error reading attachment: {e}")

    if email_has_data:
        final_row = {
            'Subject':      email_row['Subject'],
            'Date':         email_row['Date'],
            'Sender Name':  email_row['Sender Name'],
            'Sender Email': email_row['Sender Email'],
            'Source':       email_row['Source'],
        }
        for key in ['md5', 'sha1', 'sha256', 'ip', 'domain', 'url', 'email']:
            unique_vals        = sorted(set(email_row[key]))
            final_row[key]     = "\n".join(unique_vals)
        new_data.append(final_row)


def write_master_file(new_data):
    """Appends new_data to the master Excel file with full wrap-text formatting."""

    col_order = ['Subject', 'Date', 'Sender Name', 'Sender Email', 'Source',
                 'md5', 'sha1', 'sha256', 'ip', 'domain', 'url', 'email']

    new_df = pd.DataFrame(new_data, columns=col_order)

    if os.path.exists(MASTER_FILE):
        try:
            existing_df = pd.read_excel(MASTER_FILE)
            # Drop fully blank rows and rows where every IOC column is empty
            existing_df.dropna(how='all', inplace=True)
            ioc_cols = ['md5', 'sha1', 'sha256', 'ip', 'domain', 'url', 'email']
            existing_df = existing_df[existing_df.apply(
                lambda r: any(str(r.get(c, '')).strip() not in ('', 'nan') for c in ioc_cols),
                axis=1
            )]
            for c in col_order:
                if c not in existing_df.columns:
                    existing_df[c] = ''
            final_df = pd.concat([existing_df[col_order], new_df], ignore_index=True)
        except Exception as e:
            print(f"❌  Error reading existing master file: {e}")
            return
    else:
        final_df = new_df

    # Sort oldest → newest before writing
    final_df['Date'] = pd.to_datetime(final_df['Date'], errors='coerce')
    final_df.sort_values('Date', ascending=True, inplace=True)
    final_df['Date'] = final_df['Date'].dt.strftime('%Y-%m-%d')
    final_df.reset_index(drop=True, inplace=True)

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, sheet_name='IOCs')

            wb  = writer.book
            ws  = writer.sheets['IOCs']

            # ── Formats ────────────────────────────────────────────────────
            header_fmt = wb.add_format({
                'bold':       True,
                'bg_color':   '#FFFF00',
                'font_color': '#000000',
                'font_name':  'Arial',
                'font_size':  10,
                'border':     1,
                'align':      'center',
                'valign':     'vcenter',
                'text_wrap':  True,
            })
            meta_fmt = wb.add_format({
                'font_name': 'Arial',
                'font_size': 9,
                'valign':    'top',
                'border':    1,
                'text_wrap': True,
            })
            ioc_fmt = wb.add_format({
                'font_name': 'Courier New',
                'font_size': 9,
                'valign':    'top',
                'border':    1,
                'text_wrap': True,
            })
            url_fmt = wb.add_format({
                'font_name':  'Courier New',
                'font_size':  9,
                'valign':     'top',
                'border':     1,
                'text_wrap':  True,
                'font_color': '#0563C1',  # hyperlink blue
            })

            # ── Column widths & formats ─────────────────────────────────────
            col_config = {
                # col_index : (width, fmt)
                0:  (30, meta_fmt),   # Subject
                1:  (12, meta_fmt),   # Date
                2:  (20, meta_fmt),   # Sender Name
                3:  (28, meta_fmt),   # Sender Email
                4:  (18, meta_fmt),   # Source
                5:  (36, ioc_fmt),    # md5
                6:  (42, ioc_fmt),    # sha1
                7:  (66, ioc_fmt),    # sha256
                8:  (18, ioc_fmt),    # ip
                9:  (30, ioc_fmt),    # domain
                10: (55, url_fmt),    # url
                11: (35, ioc_fmt),    # email
            }
            for idx, (width, fmt) in col_config.items():
                ws.set_column(idx, idx, width, fmt)

            # ── Write header row with custom format ─────────────────────────
            for col_idx, col_name in enumerate(col_order):
                ws.write(0, col_idx, col_name, header_fmt)

            # ── Row heights: calculated per row so every value is visible ─────
            # Count newlines in each cell to figure out how many lines it needs,
            # then set the row tall enough to show all of them without expanding.
            LINE_HEIGHT = 15   # pts per line of text
            for row_idx, (_, row) in enumerate(final_df.iterrows(), start=1):
                max_lines = max(
                    len(str(val).split('\n')) for val in row.values
                )
                ws.set_row(row_idx, max_lines * LINE_HEIGHT)

            # ── Freeze the header row ───────────────────────────────────────
            ws.freeze_panes(1, 0)

            # ── Auto-filter on header ───────────────────────────────────────
            ws.autofilter(0, 0, len(final_df), len(col_order) - 1)

        print(f"✅  Done! {len(new_data)} new row(s) saved → {MASTER_FILE}")

    except Exception as e:
        print(f"❌  Error saving file: {e}")
        print("    Try: pip install xlsxwriter")


# =============================================================================
#  MAIN
# =============================================================================

def process_cti():
    print("=" * 60)
    print("  Outlook CTI Processor  —  Multi-Source Edition")
    print("=" * 60)

    if not TARGET_FOLDERS and not TARGET_SENDERS:
        print("❌  No TARGET_FOLDERS or TARGET_SENDERS configured. Edit the script.")
        return

    # ── Connect to Outlook ────────────────────────────────────────────────
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox   = outlook.GetDefaultFolder(6)
    except Exception as e:
        print(f"❌  Could not connect to Outlook: {e}")
        return

    # ── Date range ────────────────────────────────────────────────────────
    start_date = get_valid_date("Enter Start Date (YYYY-MM-DD): ")
    end_date   = get_valid_date("Enter End Date   (YYYY-MM-DD): ")
    start_dt   = datetime.combine(start_date, dt_time.min)
    end_dt     = datetime.combine(end_date,   dt_time.max)

    new_data = []

    with tempfile.TemporaryDirectory() as temp_dir:

        # ── Scan configured folders ───────────────────────────────────────
        for folder_name in TARGET_FOLDERS:
            print(f"\n📁  Scanning folder: '{folder_name}'")
            try:
                folder   = inbox.Folders(folder_name)
                messages = folder.Items
                messages.Sort("[ReceivedTime]", True)
            except Exception as e:
                print(f"   ❌  Could not open folder '{folder_name}': {e}")
                continue

            for message in messages:
                if message.Class != 43:
                    continue
                try:
                    msg_dt = message.ReceivedTime.replace(tzinfo=None)
                except Exception:
                    continue
                if msg_dt < start_dt:
                    break
                if msg_dt > end_dt:
                    continue

                print(f"   📧  {msg_dt.date()} | {message.Subject[:60]}")
                process_message(message, temp_dir, new_data, f"Folder: {folder_name}")

        # ── Scan inbox for messages from specific senders ─────────────────
        for sender_email in TARGET_SENDERS:
            print(f"\n👤  Scanning for sender: '{sender_email}'")
            messages = inbox.Items
            # Use Outlook DASL filter for performance
            filter_str = (
                f"@SQL=\"urn:schemas:httpmail:fromemail\" = '{sender_email}'"
            )
            try:
                filtered = messages.Restrict(filter_str)
                filtered.Sort("[ReceivedTime]", True)
            except Exception:
                # Fallback: iterate manually
                filtered = messages

            for message in filtered:
                if message.Class != 43:
                    continue
                try:
                    msg_dt = message.ReceivedTime.replace(tzinfo=None)
                except Exception:
                    continue
                if msg_dt < start_dt:
                    break
                if msg_dt > end_dt:
                    continue

                # Confirm sender match (needed for fallback path)
                msg_sender = getattr(message, 'SenderEmailAddress', '').lower()
                if sender_email.lower() not in msg_sender:
                    continue

                print(f"   📧  {msg_dt.date()} | {message.Subject[:60]}")
                process_message(message, temp_dir, new_data, f"Sender: {sender_email}")

    # ── Write results ─────────────────────────────────────────────────────
    if new_data:
        print(f"\n💾  Writing {len(new_data)} row(s) to {MASTER_FILE}…")
        write_master_file(new_data)
    else:
        print("\n⚠️   No IOC data found in the specified date range.")

    input("\nPress Enter to exit…")


if __name__ == "__main__":
    process_cti()
