import win32com.client
import os
import pandas as pd
import tempfile
import time
import uuid
from datetime import datetime, time as dt_time

# --- CONFIGURATION ---
TARGET_FOLDER_NAME = "cti"   
MASTER_FILE = "Master_IOC_Sheet.xlsx"

# SEARCH TERMS (Matches your specific headers)
IOC_SEARCH_TERMS = {
    'md5': ['md-5', 'md5'],
    'sha1': ['sha-1', 'sha1'],
    'sha256': ['sha-2', 'sha256'],
    'ip': ['ip address', 'ip_address']
}

def get_valid_date(prompt_text):
    while True:
        try:
            return datetime.strptime(input(prompt_text).strip(), "%Y-%m-%d").date()
        except ValueError:
            print("❌ Invalid format. Use YYYY-MM-DD.")

def find_header_row(df):
    """Scans for the real header row."""
    all_keywords = [item for sublist in IOC_SEARCH_TERMS.values() for item in sublist]
    
    # Check current headers
    current_cols = [str(c).lower().strip() for c in df.columns]
    if any(k in c for k in all_keywords for c in current_cols):
        return df

    # Scan first 10 rows
    for i, row in df.head(10).iterrows():
        row_values = [str(val).lower().strip() for val in row.values]
        if any(k in val for k in all_keywords for val in row_values):
            # Found headers
            df_new = df.iloc[i+1:].copy()
            df_new.columns = row_values
            return df_new
    return None

def process_cti_fixed_permissions():
    print(f"--- Outlook CTI Processor (Permission Fix) ---")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        target_folder = inbox.Folders(TARGET_FOLDER_NAME)
    except Exception as e:
        print(f"❌ Error connecting: {e}")
        return

    start_date = get_valid_date("Enter Start Date (YYYY-MM-DD): ")
    end_date = get_valid_date("Enter End Date   (YYYY-MM-DD): ")
    
    start_dt = datetime.combine(start_date, dt_time.min)
    end_dt = datetime.combine(end_date, dt_time.max)

    print(f"\nScanning '{TARGET_FOLDER_NAME}'...")
    
    new_data = []
    messages = target_folder.Items
    messages.Sort("[ReceivedTime]", True) 

    with tempfile.TemporaryDirectory() as temp_dir:
        for message in messages:
            if message.Class != 43: continue
            
            try:
                msg_dt = message.ReceivedTime.replace(tzinfo=None)
            except: continue

            if msg_dt < start_dt: break 
            if msg_dt > end_dt: continue

            count = message.Attachments.Count
            if count == 0: continue

            email_has_data = False
            email_row = {'Subject': message.Subject, 'Date': str(msg_dt.date()), 'md5': [], 'sha1': [], 'sha256': [], 'ip': []}

            for i in range(1, count + 1):
                try:
                    att = message.Attachments.Item(i)
                    fname = att.FileName.lower()

                    if not fname.endswith(('.xlsx', '.xls')):
                        continue
                    
                    print(f"   📎 Found Excel: {att.FileName}")

                    # --- FIX 1: UNIQUE FILENAME ---
                    # We add a random UUID to the filename. 
                    # This prevents overwriting "IOC.xlsx" with another "IOC.xlsx" (which causes the lock error)
                    unique_name = f"{uuid.uuid4()}_{att.FileName}"
                    save_path = os.path.join(temp_dir, unique_name)
                    
                    att.SaveAsFile(save_path)
                    
                    # --- FIX 2: TINY SLEEP ---
                    # Give Windows 0.5 seconds to finish writing/scanning the file before reading it
                    time.sleep(0.5) 
                    
                    # Use 'with' context to ensure file is closed immediately after reading
                    try:
                        with pd.ExcelFile(save_path) as xls:
                            for sheet_name in xls.sheet_names:
                                df = pd.read_excel(xls, sheet_name=sheet_name)
                                if df.empty: continue
                                
                                df_clean = find_header_row(df)
                                
                                if df_clean is not None:
                                    df_clean.columns = [str(c).lower().strip() for c in df_clean.columns]
                                    
                                    for ioc_type, keywords in IOC_SEARCH_TERMS.items():
                                        found_col = None
                                        for kw in keywords:
                                            found_col = next((c for c in df_clean.columns if kw in c), None)
                                            if found_col: break
                                        
                                        if found_col:
                                            vals = df_clean[found_col].dropna().astype(str).tolist()
                                            if vals:
                                                email_row[ioc_type].extend(vals)
                                                email_has_data = True
                    except PermissionError:
                        print(f"      ⚠️ Still locked. Skipping {att.FileName} (Antivirus might be holding it).")
                    except Exception as e:
                        print(f"      ⚠️ Read Error: {e}")

                except Exception as e:
                    print(f"      ⚠️ Attachment Error: {e}")

            if email_has_data:
                final_row = {'Subject': email_row['Subject'], 'Date': email_row['Date']}
                for key in ['md5', 'sha1', 'sha256', 'ip']:
                    unique_vals = sorted(list(set(email_row[key])))
                    final_row[key] = ", ".join(unique_vals)
                
                new_data.append(final_row)
                print(f"      ✅ Extracted Data.")

    if new_data:
        print(f"\nWriting {len(new_data)} rows to {MASTER_FILE}...")
        new_df = pd.DataFrame(new_data)
        
        if os.path.exists(MASTER_FILE):
            try:
                existing = pd.read_excel(MASTER_FILE)
                updated = pd.concat([existing, new_df], ignore_index=True)
                updated.to_excel(MASTER_FILE, index=False)
                print("✅ Done! Data appended.")
            except Exception as e:
                print(f"❌ Error: Close the Excel file! {e}")
        else:
            new_df.to_excel(MASTER_FILE, index=False)
            print("✅ Done! Created new file.")
    else:
        print("\n❌ No data found.")

    input("\nPress Enter to exit...")

if __name__ == "__main__":
    process_cti_fixed_permissions()
