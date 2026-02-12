# CIR Monday Reports Automation

Automated system for downloading, processing, and emailing Monday reports from Gmail attachments.

## 📋 Overview

This project automates the weekly Monday reports workflow by:
1. **Fetching** unread emails with specific subject lines from Gmail
2. **Downloading** Excel file attachments
3. **Processing** the Excel files (formatting, pivot tables, etc.)
4. **Sending** the processed files via email to designated recipients
5. **Uploading** reports to Google Drive with automatic date-based organization

---

## 🎯 Quick Start

### **Unified UI (Recommended)**
Run the unified interface to process any or all report types:
```bash
python MondayReportsUI.py
```
**Features:**
- ✅ Select multiple report types (runs sequentially)
- ✅ Choose recipients with checkboxes
- ✅ Email and/or Google Drive delivery
- ✅ Automatic date-based folder organization in Drive
- ✅ Modern dark-themed interface
- ✅ Live progress updates

### **Individual Scripts**
Each report can still be run independently:
```bash
python Monday_Andy&GregReports.py
python Monday_CameronFlatironsReports.py
python Monday_CamCrumpReports.py
python Monday_MalissaReports.py
```

---

## 📁 Project Structure

```
IND_Tools/
├── MondayReportsUI.py                  # 🆕 Unified UI for all reports (USE THIS)
├── Monday_Andy&GregReports.py          # Andy & Greg's reports (120 reports)
├── Monday_MalissaReports.py            # Malissa's reports (23 reports)
├── Monday_CamCrumpReports.py           # Cameron & Crump reports (3 reports)
├── Monday_CameronFlatironsReports.py   # Cameron Flatirons reports (56 reports)
├── DriveUploader.py                    # 🆕 Google Drive upload module
├── RepAutoGmail.py                     # Gmail API authentication module
├── Fix_defaultColWidthPt.py            # Excel column width fixer module
├── DRIVE_SETUP.md                      # 🆕 Google Drive setup guide
├── credentials.json                    # Google API OAuth credentials
├── token.pickle                        # Cached Gmail authentication token
├── token_drive.pickle                  # Cached Drive authentication token
├── requirements.txt                    # Python dependencies
└── __pycache__/                        # Python cache files
```

---

## 🚀 Setup Instructions

### 1. Prerequisites
- Python 3.13 or higher
- Windows OS (uses win32com for Excel automation)
- Gmail account with API access enabled

### 2. Install Dependencies

Install from requirements.txt:
```bash
pip install -r requirements.txt
```

**Key packages:**
- `google-auth-oauthlib` - Gmail & Drive OAuth authentication
- `google-auth-httplib2` - HTTP library for Google API
- `google-api-python-client` - Gmail & Drive API client
- `pandas` - Data manipulation
- `openpyxl` - Excel file handling
- `pywin32` - Windows COM automation for Excel
- `ttkbootstrap` - Modern themed UI components

### 3. Gmail API Setup

Already configured with `credentials.json`. On first run, the script will:
1. Open a browser window for Gmail authentication
2. Ask you to grant permissions (read, modify, send emails)
3. Save authentication token to `token.pickle` for future runs

### 4. Google Drive Setup (Optional)

For automatic Drive uploads with date-based organization:

1. **Enable Google Drive API** in your Google Cloud Console project
2. **First upload attempt** will open browser for Drive authentication
3. Grant Drive permissions to the app
4. Authentication saves to `token_drive.pickle`

**Drive Folder Structure:**
```
Monday Reports Test (Shared Drive)/
├── Andy & Greg/
│   ├── 2026-02-09/
│   │   ├── Report_1.xlsx
│   │   └── Report_2.xlsx
│   └── 2026-02-16/
├── Cameron Flatirons/
│   └── 2026-02-09/
├── Cameron & Crump/
│   └── 2026-02-09/
└── Malissa/
    └── 2026-02-09/
```

Reports are automatically organized into date subfolders (extracted from filenames).

See [DRIVE_SETUP.md](DRIVE_SETUP.md) for detailed configuration.

---

## 🎨 Unified UI Architecture

**MondayReportsUI.py** - Master interface for all report types

**Key Features:**
- Multi-select checkboxes for report types (can run multiple sequentially)
- Dynamic recipient list (shows union of emails from selected reports)
- Select All/Deselect All buttons for both reports and recipients
- Each report's `main()` function accepts `to_emails` and `status_callback` parameters
- Runs reports one after another when multiple selected
- Live status updates during processing

**How it works:**
1. Dynamically imports all 4 Monday report modules using `importlib`
2. User selects report type(s) via checkboxes
3. Recipient list updates based on selected report(s)
4. Calls each selected module's `main(to_emails, status_callback)` function
5. Processes reports sequentially with progress updates

**Individual report files retain their standalone UI** for backward compatibility.

---

## 📊 Report Scripts

### 1. Monday_Andy&GregReports.py

**Purpose:** Processes Andy & Greg's Monday reports

**Output Directory:** `C:\Users\Esteban\Desktop\Working\Python Outputs\Andy & Greg`

**Recipients:**
- aidan@tortintakeprofessionals.com
- martin@tortintakeprofessionals.com
- oroman@tortintakeprofessionals.com
- pjerome@tortintakeprofessionals.com
- esteban@tortintakeprofessionals.com
- ngaston@tortintakeprofessionals.com
- mclark@tortintakeprofessionals.com

**Email Subject:** "Andy & Greg's Monday Reports"

**Report Count:** 123+ different report types

**Processing Steps:**
1. Downloads reports with subjects matching: `"Report: A&G: [Report Name]"`
2. Fixes Excel column widths
3. Formats date columns (E-Sign Signed Date, Lead Created Date, Date of Birth)
4. Renames Sheet1 to "All Fields All Time"
5. Adds Excel table formatting
6. Creates 3 pivot tables:
   - Pivot Table Combined
   - Pivot Table Matches Dashboard
   - Pivot Table All Final
7. Auto-adjusts column widths and reorders sheets

---

### 2. Monday_MalissaReports.py

**Purpose:** Processes Malissa's Monday reports

**Output Directory:** `C:\Users\Esteban\Desktop\Working\Python Outputs\Malissa`

**Recipients:**
- aidan@tortintakeprofessionals.com
- ngaston@tortintakeprofessionals.com
- mclark@tortintakeprofessionals.com
- pjerome@tortintakeprofessionals.com
- esteban@tortintakeprofessionals.com

**Email Subject:** "Malissa Monday Reports"

**Processing:** Same as Andy & Greg reports

---

### 3. Monday_CamCrumpReports.py

**Purpose:** Processes Cameron & Crump reports

**Output Directory:** `C:\Users\Esteban\Desktop\Working\Python Outputs\Cameron\Other`

**Recipients:**
- aidan@tortintakeprofessionals.com
- martin@tortintakeprofessionals.com
- esteban@tortintakeprofessionals.com

**Email Subject:** "Cameron & Crump's Reports"

**Report Count:** 3 active reports (many commented out for testing)

**Processing Steps:**
1. Downloads reports with subjects matching: `"Report: BCL:..."` or `"Report: CAM:..."`
2. Same Excel processing as above
3. Creates 1 pivot table: "Pivot Table"

---

### 4. Monday_CameronFlatironsReports.py

**Purpose:** Processes Cameron Flatirons reports

**Output Directory:** `C:\Users\Esteban\Desktop\Working\Python Outputs\Cameron\Flatirons`

**Recipients:**
- aidan@tortintakeprofessionals.com
- ngaston@tortintakeprofessionals.com
- esteban@tortintakeprofessionals.com

**Email Subject:** "Cameron Flatirons Reports"

**Report Count:** 54 different MFI reports

**Processing Steps:**
1. Downloads reports with subjects matching: `"Report: MFI: [Report Name]"`
2. Same Excel processing
3. Creates 3 pivot tables:
   - Pivot Table Combined
   - Pivot Table Matches Benchmark
   - Pivot Table Matches Dashboard

---

## 🛠️ Helper Modules

### RepAutoGmail.py

**Purpose:** Handles Gmail API authentication

**Key Function:**
```python
confirm_auth() -> service
```
- Authenticates with Gmail using OAuth 2.0
- Loads existing token from `token.pickle` if available
- Creates new token if expired or missing
- Returns authenticated Gmail API service object

**Permissions Required:**
- `gmail.readonly` - Read emails
- `gmail.modify` - Mark emails as read
- `gmail.send` - Send emails

---

### DriveUploader.py

**Purpose:** Handles Google Drive uploads with automatic date-based organization

**Key Functions:**
```python
get_drive_service() -> service
# Authenticates with Google Drive API
# Returns authenticated Drive service object

upload_folder_to_drive(folder_path, folder_name, status_callback)
# Uploads all .xlsx files from local folder to Drive
# Automatically creates date subfolders (e.g., "2026-02-09")
# Returns count of successfully uploaded files

upload_file_to_drive(file_path, folder_name, status_callback, target_folder_id)
# Uploads single file to Drive folder
# Supports shared drives with supportsAllDrives=True
```

**Features:**
- Extracts date from filename using regex (`YYYY-MM-DD` format)
- Creates or reuses date subfolders automatically
- Supports Google Shared Drives
- Configurable folder mappings in `DRIVE_FOLDERS` dictionary
- Status callbacks for UI integration

**Configuration:**
```python
DRIVE_FOLDERS = {
    "Andy & Greg": "1qOnyoZl_lbWkUGk8r6iMroy8ZTKom91E",
    "Cameron Flatirons": "11V0Ity9HLncxsOkS-yedctUWRO6oiLFF",
    "Cameron & Crump": "1GHYmpl983zq264sZnYj-iR-M_1-w-5LJ",
    "Malissa": "1FroHjovKsopPtRTlr_LPwai7y4isX-DY"
}
```

---

### Fix_defaultColWidthPt.py

**Purpose:** Fixes Excel column width formatting issues

**Key Function:**
```python
XLSXFixer.fix_default_col_width(file_path)
```
- Opens Excel file with openpyxl
- Iterates through all worksheets
- Auto-adjusts column widths based on content
- Caps maximum width at 50 characters for readability
- Saves the modified file

---

## 🔄 Workflow Process

### For Each Script:

1. **Clear Output Directory**
   - Deletes all existing files in the output folder
   - Ensures fresh start for each run

2. **Download Reports**
   - Searches Gmail for unread emails matching subject filters
   - Downloads Excel attachments (.xlsx files)
   - Marks emails as read after download
   - Saves files to output directory

3. **Process Each Excel File**
   - Fixes column widths
   - Converts date columns to proper format (MM/DD/YYYY)
   - Renames Sheet1 to "All Fields All Time"
   - Adds formatted Excel table
   - Creates pivot tables with Status field counts
   - Auto-adjusts column widths
   - Reorders sheets (data sheet first, then pivot tables)

4. **Send Email (Optional)**
   - Collects all processed .xlsx files from output directory
   - Creates email with all files attached
   - Sends to designated recipients
   - Prints success/failure message

5. **Upload to Drive (Optional)**
   - Uploads processed files to Google Drive
   - Extracts date from filename (e.g., `2026-02-09`)
   - Creates date subfolder if it doesn't exist
   - Uploads all files to date subfolder
   - Supports both email and Drive or either one

---

## ▶️ Running the Scripts

### Using Virtual Environment:

```powershell
# Activate virtual environment
.venv\Scripts\Activate.ps1

# Run any script
python Monday_Andy&GregReports.py
python Monday_MalissaReports.py
python Monday_CamCrumpReports.py
python Monday_CameronFlatironsReports.py
```

### Direct Execution:

```powershell
C:/Users/Esteban/Documents/CIR_Monday-Reports/.venv/Scripts/python.exe Monday_Andy&GregReports.py
```

---

## ⚠️ Common Issues & Solutions

### Issue: AttributeError with win32com

**Error:**
```
AttributeError: module 'win32com.gen_py...' has no attribute 'CLSIDToClassMap'
```

**Solution:**
Delete the gen_py cache folder:
```powershell
Remove-Item -Recurse "C:\Users\Esteban\AppData\Local\Temp\gen_py"
```

### Issue: Gmail API Precondition Failed

**Error:**
```
HttpError 400: Precondition check failed
```

**Solution:**
Delete token and re-authenticate:
```powershell
Remove-Item token.pickle
# Run script again to re-authenticate
```

### Issue: No Emails Found

**Possible Causes:**
- No unread emails matching the subject filter
- Reports already processed (emails marked as read)
- Incorrect subject line in `subject_filters` list

### Issue: Excel Files Not Attaching

**Check:**
- Output directory exists and has .xlsx files
- Files aren't open in Excel (causes permission errors)
- File processing completed without errors

---

## 📝 Modifying Recipients

To add/remove email recipients, edit the `to_emails` list in each script:

```python
to_emails = [
    "aidan@tortintakeprofessionals.com",
    "esteban@tortintakeprofessionals.com",  # Uncommented = active
    # "brittany@tortintakeprofessionals.com",  # Commented = inactive
]
```

---

## 🔍 Adding New Reports

To process new report types, add subject filters to the `subject_filters` list:

```python
subject_filters = [
    "Report: A&G: New Report Name - Firm1 - Firm2 - Shield Legal",
    # Add more here...
]
```

**Subject Filter Format:**
- Must match Gmail subject line exactly
- Use quotes around the full subject
- Case-sensitive

---

## 📅 Scheduling (Future Enhancement)

To run automatically on Mondays:

### Option 1: Windows Task Scheduler
1. Open Task Scheduler
2. Create Basic Task
3. Trigger: Weekly on Mondays
4. Action: Start a Program
5. Program: `C:\Users\Esteban\Documents\CIR_Monday-Reports\.venv\Scripts\python.exe`
6. Arguments: `Monday_Andy&GregReports.py`
7. Start in: `C:\Users\Esteban\Documents\CIR_Monday-Reports`

### Option 2: Python Script with schedule library
```python
import schedule
import time

def run_reports():
    # Run all scripts
    pass

schedule.every().monday.at("09:00").do(run_reports)
while True:
    schedule.run_pending()
    time.sleep(60)
```

---

## 🔒 Security Notes

- `credentials.json` contains OAuth client secrets
- `token.pickle` contains Gmail authentication token
- `token_drive.pickle` contains Drive authentication token
- **DO NOT** commit these files to public repositories
- Keep credentials secure and rotate periodically
- Use service accounts for production environments
- Drive folders use shared drive for multi-user access

---

## 📊 Output Structure

Each script creates processed files in its output directory:

```
Working/Python Outputs/
├── Andy & Greg/
│   ├── Report_1.xlsx
│   ├── Report_2.xlsx
│   └── ...
├── Malissa/
│   └── ...
├── Cameron/
│   ├── Other/
│   │   └── ...
│   └── Flatirons/
│       └── ...
```

**Each Excel file contains:**
1. "All Fields All Time" sheet (formatted data)
2. Pivot tables (Status counts)

---

## 🤝 Support

For issues or questions:
- Check error messages in terminal output
- Review Gmail API quotas (quota limits may apply)
- Verify Excel isn't open when processing files
- Ensure all required packages are installed

---

## 📈 Statistics

**Total Scripts:** 4  
**Total Report Types:** 180+  
**Output Directories:** 4  
**Total Recipients:** 7 unique email addresses  
**Processing Time:** ~5-15 minutes per script (depends on report count)

---

## 🎯 Future Enhancements

- [x] ✅ Google Drive upload integration
- [x] ✅ Date-based folder organization
- [x] ✅ Optional email (Drive-only mode)
- [ ] Add error logging to file
- [ ] Create summary report of processed files
- [ ] Add retry logic for failed downloads
- [ ] Implement parallel processing for faster execution
- [ ] Add email notifications for failures
- [ ] Create dashboard for monitoring runs
- [ ] Add configuration file for settings
- [ ] Implement report-to-recipient mapping for selective sending

---

*Last Updated: February 12, 2026*  
*Maintained by: Esteban*
