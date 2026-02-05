import win32com.client
import os
import pandas as pd
import tempfile
from datetime import datetime, time

# --- CONFIGURATION ---
TARGET_FOLDER_NAME = "cti"   
MASTER_FILE = "Master_IOC_Sheet.xlsx"

# UPDATED MAPPING: Matches your specific column names
# Keys = Output Column Name
# Values = What to look for in the Excel Header (Lowercase)
IOC_SEARCH_TERMS = {
    'md5': ['md-5', 'md5'],          # Matches "MD-5"
    'sha1': ['sha-1', 'sha1'],        # Matches "SHA-1"
    'sha256': ['sha-2', 'sha256'],    # Matches "SHA-2"
    'ip': ['ip address', 'ip_address'] # Matches "IP Address"
}

def get_valid_date(prompt_text):
    while True:
        try:
            return datetime.strptime(input(prompt_text).strip(), "%Y-%m-%d").date()
        except ValueError:
            print("❌ Invalid format. Use YYYY-MM-DD.")

def find_header_row(df):
    """
    Scans the first 10 rows to find the row that contains your specific headers.
    """
    # Flatten keywords for easy searching
    all_keywords = [item for sublist in IOC_SEARCH_TERMS.values() for item in sublist]

    # 1. Check if the current columns already match
    current_cols = [str(c).lower().strip() for c in df.columns]
    if any(k in c for k in all_keywords for c in current_cols):
        return df

    # 2. Scan first 10 rows
    for i, row in df.head(10).iterrows():
        # Clean up the row data: lowercase, remove spaces
        row_values = [str(val).lower().strip() for val in row.values]
        
        # Check if this row contains 'md-5', 'sha-2', or 'ip address'
        if any(k in val for k in all_keywords for val in row_values):
            print(f"      🔎 Found headers on Row {i+2} (Pandas index {i})")
            
            # Reset the dataframe to start from this row
            df_new = df.iloc[i+1:].copy()
            df_new.columns = row_values
            return df_new
            
    return None

def process_cti_final_fixed():
    print(f"--- Outlook CTI Processor (Final Fix) ---")

    # 1. Connect
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        target_folder = inbox.Folders(TARGET_FOLDER_NAME)
    except Exception as e:
        print(f"❌ Error connecting to Outlook: {e}")
        return

    # 2. Dates
    start_date = get_valid_date("Enter Start Date (YYYY-MM-DD): ")
    end_date = get_valid_date("Enter End Date   (YYYY-MM-DD): ")
    
    start_dt = datetime.combine(start_date, time.min)
    end_dt = datetime.combine(end_date, time.max)

    print(f"\nScanning '{TARGET_FOLDER_NAME}' from {start_date} to {end_date}...")
    
    new_data = []
    messages = target_folder.Items
    messages.Sort("[ReceivedTime]", True) # Newest first

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

            # Loop through ALL attachments by index (Fixes the mixed image/excel issue)
            for i in range(1, count + 1):
                try:
                    att = message.Attachments.Item(i)
                    fname = att.FileName.lower()

                    if not fname.endswith(('.xlsx', '.xls')):
                        continue
                    
                    print(f"   📎 Found Excel: {att.FileName} in '{message.Subject}'")

                    save_path = os.path.join(temp_dir, att.FileName)
                    att.SaveAsFile(save_path)
                    
                    xls = pd.ExcelFile(save_path)
                    
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        if df.empty: continue
                        
                        # Find the row with "md-5" or "ip address"
                        df_clean = find_header_row(df)
                        
                        if df_clean is not None:
                            df_clean.columns = [str(c).lower().strip() for c in df_clean.columns]
                            
                            for ioc_type, keywords in IOC_SEARCH_TERMS.items():
                                found_col = None
                                # Try to match specific keywords (e.g. "md-5")
                                for kw in keywords:
                                    found_col = next((c for c in df_clean.columns if kw in c), None)
                                    if found_col: break
                                
                                if found_col:
                                    vals = df_clean[found_col].dropna().astype(str).tolist()
                                    if vals:
                                        email_row[ioc_type].extend(vals)
                                        email_has_data = True

                except Exception as e:
                    print(f"      ⚠️ Error reading attachment: {e}")

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
        print("Double check: Do the columns definitely say 'MD-5' and 'IP Address'?")

    input("\nPress Enter to exit...")

if __name__ == "__main__":
    process_cti_final_fixed()
