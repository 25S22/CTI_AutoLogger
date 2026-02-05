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

# SEARCH TERMS (Updated)
IOC_SEARCH_TERMS = {
    'md5': ['md-5', 'md5'],
    'sha1': ['sha-1', 'sha1'],
    'sha256': ['sha-2', 'sha256'],
    # Added "ip's" to catch that specific header
    'ip': ['ip address', 'ip_address', "ip's", "ips"], 
    # New Category for Domains
    'domain': ['domain', 'domains', 'url', 'host'] 
}

def get_valid_date(prompt_text):
    while True:
        try:
            return datetime.strptime(input(prompt_text).strip(), "%Y-%m-%d").date()
        except ValueError:
            print("❌ Invalid format. Use YYYY-MM-DD.")

def find_header_row(df):
    """Scans first 10 rows to find the real header."""
    all_keywords = [item for sublist in IOC_SEARCH_TERMS.values() for item in sublist]
    
    current_cols = [str(c).lower().strip() for c in df.columns]
    if any(k in c for k in all_keywords for c in current_cols):
        return df

    for i, row in df.head(10).iterrows():
        row_values = [str(val).lower().strip() for val in row.values]
        if any(k in val for k in all_keywords for val in row_values):
            df_new = df.iloc[i+1:].copy()
            df_new.columns = row_values
            return df_new
    return None

def process_cti_final_v3():
    print(f"--- Outlook CTI Processor (Domains + Formatting) ---")

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
            # Initialize row with all columns (including domain)
            email_row = {
                'Subject': message.Subject, 
                'Date': str(msg_dt.date()), 
                'md5': [], 'sha1': [], 'sha256': [], 'ip': [], 'domain': []
            }
            
            for i in range(1, count + 1):
                try:
                    att = message.Attachments.Item(i)
                    fname = att.FileName.lower()

                    if not fname.endswith(('.xlsx', '.xls')):
                        continue
                    
                    unique_name = f"{uuid.uuid4()}_{att.FileName}"
                    save_path = os.path.join(temp_dir, unique_name)
                    att.SaveAsFile(save_path)
                    time.sleep(0.5) 
                    
                    file_has_ioc = False
                    
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
                                        # Strict matching for "ip's" to avoid issues
                                        found_col = next((c for c in df_clean.columns if kw in c), None)
                                        if found_col: break
                                    
                                    if found_col:
                                        vals = df_clean[found_col].dropna().astype(str).tolist()
                                        if vals:
                                            email_row[ioc_type].extend(vals)
                                            file_has_ioc = True
                                            email_has_data = True
                    
                    if file_has_ioc:
                        print(f"   ✅ Extracted data from: {att.FileName}")
                    else:
                        print(f"   ⚠️ Checked {att.FileName} but headers didn't match.")

                except Exception as e:
                    print(f"   ❌ Error reading attachment: {e}")

            if email_has_data:
                final_row = {'Subject': email_row['Subject'], 'Date': email_row['Date']}
                # Join with Newline for vertical stacking
                for key in ['md5', 'sha1', 'sha256', 'ip', 'domain']:
                    unique_vals = sorted(list(set(email_row[key])))
                    final_row[key] = "\n".join(unique_vals)
                
                new_data.append(final_row)

    if new_data:
        print(f"\nWriting {len(new_data)} rows to {MASTER_FILE}...")
        
        if os.path.exists(MASTER_FILE):
            try:
                existing_df = pd.read_excel(MASTER_FILE)
                final_df = pd.concat([existing_df, pd.DataFrame(new_data)], ignore_index=True)
            except Exception as e:
                print(f"❌ Error reading existing file: {e}")
                return
        else:
            final_df = pd.DataFrame(new_data)

        # Apply Formatting
        try:
            with pd.ExcelWriter(MASTER_FILE, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name='IOCs')
                
                workbook  = writer.book
                worksheet = writer.sheets['IOCs']
                
                wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
                
                # Format Subject/Date (Columns A-B)
                worksheet.set_column(0, 1, 20, workbook.add_format({'valign': 'top'}))
                
                # Format IOC Columns (C-G) -> MD5, SHA1, SHA256, IP, DOMAIN
                # We set them wide (45) and wrap text
                worksheet.set_column(2, 6, 45, wrap_format)

            print("✅ Done! Data saved with formatted columns.")
        except Exception as e:
            print(f"❌ Error saving file: {e}")
            print("   Try running: pip install xlsxwriter")
    else:
        print("\n❌ No data found.")

    input("\nPress Enter to exit...")

if __name__ == "__main__":
    process_cti_final_v3()
