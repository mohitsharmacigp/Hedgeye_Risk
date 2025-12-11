import win32com.client
from bs4 import BeautifulSoup
import pandas as pd
import os
from datetime import datetime
import re
from datetime import datetime, date

def parse_email_body(html_body, email_date):
    soup = BeautifulSoup(html_body, features="html.parser")
    table_data = []
    table_rows = soup.find_all("tr")
    for row in table_rows:
        cells = row.find_all("td")
        if len(cells) == 4:
            index_full_text = cells[0].get_text(separator="\n", strip=True)
            index_parts = index_full_text.split("\n")
            index_name = index_parts[0].strip() if len(index_parts) >= 1 else index_full_text.strip()
            index_desc = index_parts[1].strip() if len(index_parts) >= 2 else ""

            buy_trade_str = cells[1].get_text(strip=True).replace(",", "")
            sell_trade_str = cells[2].get_text(strip=True).replace(",", "")
            prev_close_str = cells[3].get_text(strip=True).replace(",", "")

            buy_trade = pd.to_numeric(buy_trade_str, errors='coerce')
            sell_trade = pd.to_numeric(sell_trade_str, errors='coerce')
            prev_close = pd.to_numeric(prev_close_str, errors='coerce')

            table_data.append({
                "Date": email_date,  # Keep as full datetime
                "index": index_name,
                "index_desc": index_desc,
                "buy_trade": buy_trade,
                "sell_trade": sell_trade,
                "prev_close": prev_close,
            })

    if not table_data:
        print(f"No table data found in email dated {email_date.strftime('%B %d, %Y')}.")
        return None, None

    return table_data, email_date

def get_last_row_win32com(ws):
    """Get the last row with data in column A using win32com"""
    return ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp

def load_existing_dates(excel_file_path):
    """Load only dates from existing Raw worksheet using win32com"""
    existing_dates = set()
    if os.path.exists(excel_file_path):
        try:
            # Initialize Excel application
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # Run in background
            excel.DisplayAlerts = False  # Suppress alerts
            
            # Open workbook
            workbook = excel.Workbooks.Open(excel_file_path)
            
            # Check if Raw sheet exists
            sheet_exists = False
            for sheet in workbook.Sheets:
                if sheet.Name == "Raw":
                    sheet_exists = True
                    ws = workbook.Sheets("Raw")
                    break
            
            if sheet_exists:
                # Find the last row with data in column A
                last_row = get_last_row_win32com(ws)
                cell_value= ws.Cells(last_row, 1).Value
                if cell_value is not None:
                    if isinstance(cell_value, datetime):
                        date_obj = cell_value
                    else:
                        raw_value = ws.Cells(last_row, 1).Value2
                        # Check if it's a numeric Excel date
                        if isinstance(raw_value, (int, float)):
                            excel_zero = datetime(1899, 12, 30)
                            date_obj = excel_zero + pd.Timedelta(days=raw_value)
                        else:
                            # Try to parse as string
                            date_str = str(cell_value).strip()
                            try: 
                                # Try the format in your Excel: "11/24/2025 12:00:00 AM"
                                date_obj = datetime.strptime(date_str, "%m/%d/%Y %I:%M:%S %p")
                            except ValueError:
                                try:
                                    # Try without time
                                    date_obj = datetime.strptime(date_str, "%m/%d/%Y")
                                except ValueError:
                                    print(f"Could not parse date: {date_str}")             
            
            # Close workbook and quit Excel
            workbook.Close(SaveChanges=False)
            excel.Quit()
            
        except Exception as e:
            print(f"Error reading Excel file with win32com: {e}")
            # Ensure Excel is closed even on error
            try:
                excel.Quit()
            except:
                pass
    return date_obj

def append_new_data_to_excel(excel_file_path, new_data):
    """Append new data to the existing Raw worksheet using win32com"""
    try:
        # Initialize Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Run in background
        excel.DisplayAlerts = False  # Suppress alerts
        
        # Open workbook
        workbook = excel.Workbooks.Open(excel_file_path)
        
        # Check if Raw sheet exists, create if not
        sheet_exists = False
        for sheet in workbook.Sheets:
            if sheet.Name == "Raw":
                sheet_exists = True
                ws = workbook.Sheets("Raw")
                break
        
        if not sheet_exists:
            print("Raw sheet not found. Creating new one.")
            ws = workbook.Sheets.Add()
            ws.Name = "Raw"
            # Add headers
            headers = ["Date", "index", "index_desc", "buy_trade", "sell_trade", "prev_close"]
            for col_idx, header in enumerate(headers, 1):
                ws.Cells(1, col_idx).Value = header
            start_row = 2
        else:
            start_row = get_last_row_win32com(ws) + 1
        
        # Append new data
        for data_row in new_data:
            # Get the date from data_row 
            excel_date = data_row["Date"] + pd.Timedelta(hours=8)
            # Set the date value
            ws.Cells(start_row, 1).Value = excel_date

            # Apply date formatting to match existing format
            ws.Cells(start_row, 1).NumberFormat = "yyyy-mm-dd h:mm:ss"
            
            # Set other values
            ws.Cells(start_row, 2).Value = data_row["index"]
            ws.Cells(start_row, 3).Value = data_row["index_desc"]
            ws.Cells(start_row, 4).Value = data_row["buy_trade"]
            ws.Cells(start_row, 5).Value = data_row["sell_trade"]
            ws.Cells(start_row, 6).Value = data_row["prev_close"]
            
            start_row += 1
        
        # Save and close
        workbook.Save()
        workbook.Close(SaveChanges=True)
        excel.Quit()
        
        print(f"Appended {len(new_data)} rows to Raw worksheet in {excel_file_path}")
        return True
    except Exception as e:
        print(f"Error appending data to Excel with win32com: {e}")
        # Ensure Excel is closed even on error
        try:
            excel.Quit()
        except:
            pass
        return False
    

def to_date(obj):
    """Normalize obj to a datetime.date for safe day-level comparisons."""
    # Already a pure date
    if isinstance(obj, date) and not isinstance(obj, datetime):
        return obj

    # Standard datetime -> date
    if isinstance(obj, datetime):
        return obj.date()

    # pywintypes.datetime (Outlook/COM) -> date
    try:
        import pywintypes
        if isinstance(obj, pywintypes.TimeType):
            return obj.date()
    except Exception:
        # pywintypes may not be available outside Windows/win32com
        pass
    raise TypeError(f"Unsupported date type: {type(obj)}")

def main():
    # Ensure Outlook is running
    import subprocess
    import time
    try:
        # Check if Outlook is already running, if not, start it
        outlook_app = win32com.client.Dispatch("Outlook.Application")
    except:
        # If Outlook is not running, start it
        subprocess.Popen("outlook.exe")
        time.sleep(3)  # Wait for Outlook to start
        outlook_app = win32com.client.Dispatch("Outlook.Application")
    
    # Connect to Outlook
    try:
        outlook = outlook_app.GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
        # Navigate to the "Hedgeye Risk Ranges" subfolder
        hedgeye_folder = inbox.Folders("Hedgeye Risk Ranges")
    except Exception as e:
        print(f"Error connecting to Outlook or accessing folder: {e}")
        return
 
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Updated file name to "hedgeye risk ranges.xlsm"
    excel_file_path = os.path.join(script_dir, "Raw.xlsx")
 
    # Load only existing dates (not the full data)
    latest_date = load_existing_dates(excel_file_path)
    
    if latest_date:
        print(f"Latest Date: {latest_date}.strftime('%B %d, %Y')")
 
    # Filter emails using Outlook's Restrict method from the subfolder
    messages = hedgeye_folder.Items
    messages.Sort("[ReceivedTime]", True)
    filter_str = ('@SQL="urn:schemas:httpmail:fromemail" LIKE \'%info@hedgeye.com%\' '
                  'AND "urn:schemas:httpmail:subject" LIKE \'%risk range™ signals%\'')
    filtered_messages = messages.Restrict(filter_str)
 
    if not filtered_messages.Count:
        print("No matching emails found in 'Hedgeye Risk Ranges' folder.")
        return
 
    # Process emails in a single pass
    new_data = []
    update_dates = set()
    date_pattern = re.compile(r"risk range™ signals:\s*(\w+\s*\d+,\s*\d{4})", re.IGNORECASE)
 
    for message in filtered_messages:
        try:
            subject = message.Subject.lower()
            match = date_pattern.search(subject)
            if not match:
                continue
 
            date_text = match.group(1)
            try:
                email_date = datetime.strptime(date_text, "%B %d, %Y")
            except ValueError:
                print(f"Could not parse date from subject: {date_text}")
                continue
            
            email_day  = to_date(email_date)
            latest_day = to_date(latest_date)

            # Only process body if date is missing (compare date objects)
            if email_day > latest_day:
                html_body = message.HTMLBody
                table_data, _ = parse_email_body(html_body, email_date)
                if table_data:
                    new_data.extend(table_data)
                    update_dates.add(email_date.date())
                    print(f"Collected data for missing date: {email_date.strftime('%B %d, %Y')} with {len(table_data)} securities")
 
        except Exception as e:
            print(f"Error processing email: {e}")
            continue

    print(f"Dates to update {[d.strftime('%B %d, %Y') for d in update_dates]}")


    if not update_dates:
        print("No missing dates found. Raw sheet is up to date.")
        return
 
    if not new_data:
        print("No new data collected for missing dates.")
        return
 
    # Sort new_data by date and index_desc to maintain consistency
    new_data_sorted = sorted(new_data, key=lambda x: (x["Date"], x["index_desc"]))
    
    # Append new data to existing worksheet
    success = append_new_data_to_excel(excel_file_path, new_data_sorted)
    
    if success:
        print(f"Successfully appended {len(new_data_sorted)} rows of new data to the Raw worksheet.")
        print(f"Added data for dates: {sorted(list(update_dates))}")
    else:
        print("Failed to append new data to Excel file.")
 
if __name__ == "__main__":
    main()