import win32com.client
import os
import pandas as pd
import tempfile
from datetime import datetime
import sys

# --- STATIC CONFIGURATION ---
TARGET_FOLDER_NAME = "Invoices"         # Folder in Outlook to search
MASTER_FILE = "Master_IOC_Sheet.xlsx"   # The file to append data to
IOC_COLUMNS = ['md5', 'sha1', 'sha256', 'ip'] # The headers we look for

def get_valid_date(prompt_text):
    """Asks the user for a date and validates the format."""
    while True:
        date_str = input(prompt_text).strip()
        try:
            # Try to parse the string into a date object
            valid_date = datetime.strptime(date_str, "%Y-%m-%d").date()
            return valid_date
        except ValueError:
            print("❌ Invalid format. Please use YYYY-MM-DD (e.g., 2026-01-31).")

def get_folder(base_folder, target_name):
    """Recursively find a folder by name within Outlook."""
    if base_folder.Name == target_name:
        return base_folder
    for folder in base_folder.Folders:
        found = get_folder(folder, target_name)
        if found: return found
    return None

def process_outlook_emails():
    # --- 1. USER INPUT ---
    print("--- Outlook IOC Extractor ---")
    start_date = get_valid_date("Enter Start Date (YYYY-MM-DD): ")
    end_date = get_valid_date("Enter End Date   (YYYY-MM-DD): ")

    if start_date > end_date:
        print("❌ Error: Start date cannot be after End date.")
        return

    # --- 2. CONNECT TO OUTLOOK ---
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6) # 6 = Inbox
    except Exception as e:
        print(f"❌ Could not connect to Outlook. Is it open? Error: {e}")
        return
    
    target_folder = get_folder(inbox, TARGET_FOLDER_NAME)
    if not target_folder:
        print(f"❌ Error: Folder '{TARGET_FOLDER_NAME}' not found in Inbox.")
        return

    print(f"\nScanning '{target_folder.Name}' for emails between {start_date} and {end_date}...")
    
    new_data = []

    # --- 3. PROCESS EMAILS ---
    with tempfile.TemporaryDirectory() as temp_dir:
        messages = target_folder.Items
        messages.Sort("[ReceivedTime]", True) # Newest first

        for message in messages:
            if message.Class != 43: continue # Skip non-emails

            # Check Date Range
            msg_date = message.ReceivedTime.date()
            if not (start_date <= msg_date <= end_date):
                continue

            # Check for Excel attachments
            excel_attachments = [
                att for att in message.Attachments 
                if att.FileName.lower().endswith(('.xlsx', '.xls'))
            ]

            if not excel_attachments:
                continue

            # Initialize Row
            email_row = {
                'Subject': message.Subject,
                'Date': str(msg_date),
                'md5': [], 'sha1': [], 'sha256': [], 'ip': []
            }

            has_data = False

            for attachment in excel_attachments:
                try:
                    save_path = os.path.join(temp_dir, attachment.FileName)
                    attachment.SaveAsFile(save_path)
                    
                    # Read Excel
                    df = pd.read_excel(save_path)
                    # Normalize headers to lowercase/stripped
                    df.columns = [str(c).lower().strip() for c in df.columns]

                    # Extract the fixed columns
                    for col in IOC_COLUMNS:
                        # Find column if it contains the keyword (e.g. "ip" matches "src_ip")
                        found_col = next((c for c in df.columns if col in c), None)
                        
                        if found_col:
                            values = df[found_col].dropna().astype(str).tolist()
                            if values:
                                email_row[col].extend(values)
                                has_data = True

                except Exception as e:
                    print(f"⚠️ Error reading attachment in '{message.Subject}': {e}")

            # Aggregate found data
            if has_data:
                final_row = {
                    'Subject': email_row['Subject'],
                    'Date': email_row['Date']
                }
                for col in IOC_COLUMNS:
                    unique_vals = sorted(list(set(email_row[col])))
                    final_row[col] = ", ".join(unique_vals)
                
                new_data.append(final_row)
                print(f"✅ Found data in: {message.Subject}")

    # --- 4. SAVE TO MASTER SHEET ---
    if new_data:
        new_df = pd.DataFrame(new_data)
        
        if os.path.exists(MASTER_FILE):
            print(f"\nAppending to existing file: {MASTER_FILE}")
            try:
                # Load existing to preserve history
                existing_df = pd.read_excel(MASTER_FILE)
                updated_df = pd.concat([existing_df, new_df], ignore_index=True)
                updated_df.to_excel(MASTER_FILE, index=False)
            except Exception as e:
                print(f"❌ Error updating master file. Please close Excel and try again! Error: {e}")
                return
        else:
            print(f"\nCreating new master file: {MASTER_FILE}")
            new_df.to_excel(MASTER_FILE, index=False)
            
        print(f"Success! {len(new_data)} emails processed.")
        # Keep window open if user double-clicked the script
        input("\nPress Enter to exit...")
    else:
        print("\nNo emails found with Excel attachments in that date range.")
        input("\nPress Enter to exit...")

if __name__ == "__main__":
    process_outlook_emails()
