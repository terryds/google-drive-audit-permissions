# Google Drive Audit - Open Source

An open-source Google Sheets tool that audits Google Drive files and their permissions with scheduled automatic updates.

## Features

- 📁 **Complete File Listing**: Audits all Google Drive files you have access to
- 🔒 **Permission Analysis**: Shows detailed permission information for each file
- 👥 **Access Control**: Identifies who has access and their roles (viewer, editor, owner, etc.)
- ⏰ **Scheduled Audits**: Automatically runs weekly audits using Apps Script triggers
- 📊 **Summary Dashboard**: Provides an overview of your audit results
- 🔍 **Filterable Results**: Easy-to-filter spreadsheet for finding specific sharing patterns
- ⏳ **Real-Time Status Tracking**: Monitor audit progress with live status updates
- 🔄 **Automatic Continuation**: Handles very large Drive accounts by processing in batches with 1-minute intervals - no timeouts!
- ⚡ **Fast Processing**: Uses 1-minute continuation intervals for quick completion

## What It Tracks

For each file, the tool captures:

- File name and ID
- Owner information
- File type and MIME type
- Creation and modification dates
- File size
- Direct URL to the file
- All permissions including:
  - Permission type (user, group, domain, anyone)
  - Role (owner, organizer, fileOrganizer, writer, commenter, reader)
  - Email address or domain
  - Display name

## Setup Instructions

### Prerequisites

- [Node.js](https://nodejs.org/) installed
- [clasp](https://github.com/google/clasp) CLI tool installed (`npm install -g @google/clasp`)
- A Google account

### Installation

1. **Clone or download this repository**

2. **Login to clasp**
   ```bash
   clasp login
   ```

3. **Create a new Google Sheets file** or open an existing one where you want the add-on

4. **Create a new Apps Script project bound to your sheet**
   ```bash
   clasp create --type sheets --title "Drive Audit"
   ```

5. **Push the code to Apps Script**
   ```bash
   clasp push
   ```

6. **Open the Apps Script editor to enable Drive API**
   ```bash
   clasp open
   ```

   In the Apps Script editor:
   - Go to **Services** (+ icon on the left sidebar)
   - Find and add **Google Drive API v3**
   - Click "Add"

7. **Go back to your Google Sheet** and refresh the page

8. **Authorize the script** when prompted

## Usage

### Running Manual Audit

1. In Google Sheets, click **Drive Audit** → **Run Audit Now**
2. You'll see a message that the audit is starting. Click OK to begin
3. The audit runs in the background with automatic continuation every minute:
   - ⏳ Check the **"Audit Status"** sheet for real-time progress
   - 📊 Or click **Drive Audit** → **Check Audit Status**
   - ⚡ Processes ~500 files every 5 minutes
4. When complete, review the results in the "Drive Audit" sheet

### Checking Audit Status

Click **Drive Audit** → **Check Audit Status** to see:
- Current status (Running, Completed, Cancelled, or Error)
- Progress percentage
- Last update time
- Estimated time remaining

### Cancelling a Running Audit

If you need to stop an audit that's currently running:

1. Click **Drive Audit** → **Cancel Running Audit**
2. Confirm that you want to cancel
3. The audit will be stopped immediately

**What happens when you cancel:**
- ✅ Audit process stops
- ✅ Continuation triggers are removed
- ✅ Results already written to the sheet remain
- ✅ You can start a new audit anytime

### Setting Up Scheduled Audits

1. Click **Drive Audit** → **Setup Weekly Schedule**
2. Click "Yes" to confirm
3. The audit will now run automatically every Monday at 6:00 AM

### Removing Scheduled Audits

1. Click **Drive Audit** → **Remove Schedule**
2. The weekly audit schedule will be removed
3. You can still run audits manually anytime from the Drive Audit menu

### Understanding the Results

**Audit Status Sheet:**
- 🟢 **COMPLETED** (green) - Audit finished successfully
- 🟡 **RUNNING** (yellow) - Audit is in progress
- 🛑 **CANCELLED** (gray) - Audit was cancelled by user
- 🔴 **ERROR** (red) - An error occurred
- Shows real-time progress percentage
- Last updated timestamp
- Detailed status message

**Audit Summary Sheet:**
- Shows total files and permissions audited
- Displays audit date and time
- Provides next steps and tips

**Drive Audit Sheet:**
- One row per permission (files with multiple permissions have multiple rows)
- Use filters to find:
  - Files shared with "anyone with the link"
  - Files shared with external domains
  - Files with specific permission roles
  - Files owned by specific users

## Common Use Cases

### Find Publicly Shared Files
Filter the "Permission Type" column for "anyone"

### Find Externally Shared Files
Filter the "Permission Email" column for domains outside your organization

### Find Files You Own
Filter the "Owner" column for your email address

### Find Files with Edit Access
Filter the "Permission Role" column for "writer" or "editor"

## OAuth Scopes

This script requires the following permissions:

- `https://www.googleapis.com/auth/spreadsheets.currentonly` - To read and write to the current spreadsheet
- `https://www.googleapis.com/auth/drive.readonly` - To read Drive files and permissions
- `https://www.googleapis.com/auth/script.scriptapp` - To create scheduled triggers
- `https://www.googleapis.com/auth/script.container.ui` - To display HTML dialogs and user interface

## Troubleshooting

### "Authorization Required" Error
- Make sure you've authorized the script when first opening the sheet
- Try running the audit manually first before setting up a schedule

### "Drive API has not been enabled" Error
- Open the Apps Script editor (`clasp open`)
- Add the Google Drive API service (see installation steps)

### Audit Takes Too Long / Timeout Errors
- **No longer an issue!** The tool now uses batch processing with 1-minute continuation intervals
- For very large Drive accounts (10,000+ files):
  - The audit automatically pauses before timeout (at 4.5 minutes)
  - It saves its progress and schedules a continuation in 1 minute
  - This repeats until all files are processed
  - You can monitor progress in the "Audit Status" sheet
  - The entire process is automatic - no manual intervention needed

### Missing Files
- The tool only shows files you have access to
- Files in shared drives require the `supportsAllDrives` parameter (already included)

## Customization

### Change Audit Schedule Time or Day
Edit the `setupWeeklyTrigger()` function in `Code.js`:

**Change the day of week:**
```javascript
ScriptApp.newTrigger('runDriveAudit')
  .timeBased()
  .onWeekDay(ScriptApp.WeekDay.MONDAY)  // Change to TUESDAY, WEDNESDAY, etc.
  .atHour(6)
  .create();
```

**Change the time:**
```javascript
ScriptApp.newTrigger('runDriveAudit')
  .timeBased()
  .onWeekDay(ScriptApp.WeekDay.MONDAY)
  .atHour(6)  // Change this to your preferred hour (0-23)
  .create();
```

**For daily audits instead of weekly:**
```javascript
ScriptApp.newTrigger('runDriveAudit')
  .timeBased()
  .atHour(6)
  .everyDays(1)  // Run every day
  .create();
```

### Filter Specific File Types
Modify the `getAllDriveFiles()` function to add a query filter:
```javascript
const response = Drive.Files.list({
  pageSize: 100,
  q: "mimeType='application/vnd.google-apps.spreadsheet'",  // Add this line
  fields: 'nextPageToken, files(...)',
  ...
});
```

### Audit Shared Drives Only
Add to the query in `getAllDriveFiles()`:
```javascript
q: "('sharedDriveId' in parents)",
```

## License

MIT License - feel free to modify and use as needed.

## Support

For issues or questions:
1. Check the Apps Script execution logs: `clasp logs`
2. Review the spreadsheet's execution transcript
3. Ensure all OAuth scopes are authorized
4. Open an issue on the GitHub repository

## How It Handles Large Drive Accounts

The tool uses an intelligent batch processing system to handle Drive accounts of any size:

1. **Time Monitoring**: Tracks execution time and stops at 4.5 minutes (safely before the 6-minute limit)
2. **State Persistence**: Saves progress using PropertiesService
3. **Auto-Continuation**: Automatically creates a trigger to resume after 1 minute
4. **Incremental Writing**: Writes audit data to the sheet as it processes, so no data is lost
5. **Smart Recovery**: Even if interrupted, the audit can resume from where it left off

### Fast Processing

This open-source version uses **1-minute continuation intervals** for fast processing:
- The audit processes ~500 files per batch (4.5 minutes)
- Continuation happens automatically every minute
- Large Drive accounts complete much faster than published add-ons
- You can audit Drive accounts with 50,000+ files without any manual intervention!

### Estimated Completion Times

Here's what to expect based on your Drive size:

| Files in Drive | Processing Time |
|---------------|-----------------|
| 0-500         | ~5 minutes      |
| 500-1,000     | ~5 minutes      |
| 1,000-5,000   | ~30 minutes     |
| 5,000-10,000  | ~60 minutes     |
| 10,000-20,000 | ~2 hours        |
| 50,000+       | ~5 hours        |

**Note:** These are approximate. Actual time depends on:
- Number of permissions per file
- Network latency
- API response times

## Version History

- **2.0.0** - Open Source Release
  - Changed to 1-minute continuation intervals (default)
  - Removed premium/paid features
  - Optimized for personal and organizational use
  - Updated documentation for open-source distribution
  
- **1.0.0** - Initial release
  - Basic file and permission auditing
  - Scheduled trigger support
  - Summary dashboard
  - Batch processing with automatic continuation
  - Real-time status tracking
  - Timeout protection

