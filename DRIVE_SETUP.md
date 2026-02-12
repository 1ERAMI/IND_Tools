# Google Drive Setup Instructions

## 🔧 Setup Steps

### 1. Enable Google Drive API

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Select your project (or create a new one - same one as Gmail API)
3. Go to "APIs & Services" > "Library"
4. Search for "Google Drive API"
5. Click "Enable"

### 2. Update OAuth Consent Screen (if needed)

1. Go to "APIs & Services" > "OAuth consent screen"
2. Click "Edit App"
3. Under "Scopes", click "Add or Remove Scopes"
4. Add the scope: `.../auth/drive.file` (allows creating/editing files created by this app)
5. Save and Continue

**Note:** You may need to re-authenticate after adding the Drive scope.

### 3. Create Google Drive Folders

Create 4 folders in your Google Drive:

```
📁 Monday Reports
   ├── 📁 Andy & Greg Reports
   ├── 📁 Cameron Flatirons Reports
   ├── 📁 Cameron & Crump Reports
   └── 📁 Malissa Reports
```

### 4. Get Folder IDs

For each folder you created:

1. Open the folder in Google Drive (web browser)
2. Look at the URL in your browser's address bar:
   ```
   https://drive.google.com/drive/folders/FOLDER_ID_HERE
                                         ^^^^^^^^^^^^^^^^^
                                         Copy this part
   ```
3. Copy the folder ID (the long string of letters/numbers after `/folders/`)

### 5. Update DriveUploader.py

Open `DriveUploader.py` and update the `DRIVE_FOLDERS` dictionary:

```python
DRIVE_FOLDERS = {
    "Andy & Greg": "1a2b3c4d5e6f7g8h9i0j",           # Replace with your folder ID
    "Cameron Flatirons": "9i8h7g6f5e4d3c2b1a",       # Replace with your folder ID
    "Cameron & Crump": "abcd1234efgh5678ijkl",       # Replace with your folder ID
    "Malissa": "wxyz9876stuv5432nopq"                # Replace with your folder ID
}
```

### 6. Test the Setup

Run this command to test Drive authentication:

```bash
python DriveUploader.py
```

This will:
- Show setup instructions
- Authenticate with Google Drive (first time only)
- Create `token_drive.pickle` for future use

### 7. Use in Reports

In the unified UI (`MondayReportsUI.py`):
- Check the "📤 Upload processed files to Google Drive" checkbox
- Process your reports as usual
- Files will be emailed AND uploaded to Drive

**Default behavior:** Without the checkbox, reports are only emailed (no Drive upload).

---

## 🔍 Troubleshooting

### "Folder not configured" Error
- Make sure you updated `DRIVE_FOLDERS` in `DriveUploader.py`
- Replace ALL "REPLACE_WITH_FOLDER_ID" placeholders with actual folder IDs

### Authentication Failed
- Delete `token_drive.pickle` and re-authenticate
- Make sure Google Drive API is enabled in your Cloud Console
- Check that the Drive scope is added to your OAuth consent screen

### Files Not Appearing in Drive
- Check folder IDs are correct (copy-paste from Drive URL)
- Make sure you're looking in the right folder
- Refresh your Google Drive page

---

## 📋 Quick Reference

**Files involved:**
- `DriveUploader.py` - Drive authentication and upload logic
- `token_drive.pickle` - Cached Drive authentication (auto-generated)
- Each Monday report file imports and uses DriveUploader

**To disable Drive uploads:**
- Simply uncheck the Drive checkbox in the UI
- Default is OFF, so existing workflow is unchanged
