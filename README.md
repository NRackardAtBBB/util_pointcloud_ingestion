# Folder Ingestion Script

Automated script to monitor a source directory for new subfolders, validate naming conventions, move folders to a destination, and log all file details to an Excel spreadsheet.

## Features

- ✅ **Automatic folder scanning** - Checks source directory for new subfolders
- ✅ **Naming validation** - Enforces `XXXX ProjectName` or `XXXX.XX ProjectName` convention
- ✅ **Smart renaming** - Attempts to fix common naming issues (underscores, missing spaces)
- ✅ **Duplicate prevention** - Skips folders already processed or in destination
- ✅ **Excel logging** - Records folder name, processing date, and file details
- ✅ **Retry logic** - Handles locked Excel files with automatic retries
- ✅ **Flag system** - Marks folders with 'X' if naming cannot be fixed

## Installation

1. **Install Python dependencies:**
   ```powershell
   pip install -r requirements.txt
   ```

2. **Configure the script:**
   Edit `config.yaml` and set your paths:
   ```yaml
   source_dir: "C:/path/to/source/folder"
   destination_dir: "C:/path/to/destination/folder"
   excel_log_path: "C:/path/to/log/processing_log.xlsx"
   ```

## Usage

### Manual Execution

Run the script manually from the command line:

```powershell
python ingest.py
```

The script will:
1. Scan the source directory for subfolders
2. Validate each folder name against the naming convention
3. Attempt to fix names that don't match
4. Move folders to the destination directory
5. Log all files to the Excel spreadsheet

### Automated Execution with Task Scheduler

For periodic automated execution:

1. **Open Task Scheduler:**
   - Press `Win + R`, type `taskschd.msc`, press Enter

2. **Create a New Task:**
   - Click "Create Task" (not "Create Basic Task")
   - Name: `Folder Ingestion Script`
   - Description: `Periodically checks for new folders and processes them`
   - Select "Run whether user is logged on or not"
   - Check "Run with highest privileges"

3. **Configure Triggers:**
   - Click "Triggers" tab → "New"
   - Choose frequency (e.g., Daily at 2:00 AM, or Every hour)
   - Set start date and time

4. **Configure Actions:**
   - Click "Actions" tab → "New"
   - Action: "Start a program"
   - Program/script: `python.exe` (or full path like `C:\Python311\python.exe`)
   - Add arguments: `ingest.py`
   - Start in: `C:\Users\nrackard\Code\util-pointcloud-ingestion`

5. **Configure Settings:**
   - Check "Allow task to be run on demand"
   - Check "Run task as soon as possible after a scheduled start is missed"
   - If the task fails, restart every: `10 minutes`

6. **Save and Test:**
   - Click OK and enter your Windows password
   - Right-click the task and select "Run" to test

## Naming Convention

The script expects folder names to follow this format:

```
<Project Number> <Project Name>
```

**Valid Examples:**
- `3019 Hart Island`
- `3055.12 NPS Phase 2`
- `1234 Building Survey`
- `5678.01 Bridge Inspection`

**Project Number Rules:**
- Must be 4 digits (e.g., `3019`)
- OR 4 digits, a period, and 2 digits (e.g., `3055.12`)
- Must be followed by a space and the project name

**Automatic Fixes:**
The script will attempt to fix common issues:
- `3019_Hart_Island` → `3019 Hart Island`
- `3055.12-NPS-Phase-2` → `3055.12 NPS Phase 2`
- `3019HartIsland` → `3019 HartIsland`

If the script cannot determine the correct format, it will:
- Copy the folder with its original name
- Flag it with 'X' in the Excel log

## Excel Log Format

The script creates/updates an Excel file with the following columns:

| Folder Name | Naming Flag | Processed Date | File Name | File Path | File Created Date |
|-------------|-------------|----------------|-----------|-----------|-------------------|
| 3019 Hart Island | OK | 2025-12-23 10:30:15 | survey.las | survey.las | 2025-12-20 14:22:01 |
| 3019 Hart Island | OK | 2025-12-23 10:30:15 | data.txt | subfolder/data.txt | 2025-12-21 09:15:33 |
| Bad_Folder_Name | X | 2025-12-23 10:30:20 | file.pdf | file.pdf | 2025-12-19 11:00:00 |

**Each file gets its own row** with the folder information repeated.

## Excel Locking

If the Excel file is open when the script runs, it will:
- Retry up to 10 times (configurable in `config.yaml`)
- Wait 2 seconds between retries (configurable)
- Display progress messages
- Fail gracefully if still locked after all retries

**Tip:** Close the Excel file before scheduled runs, or increase retry attempts.

## Configuration Reference

Edit `config.yaml` to customize behavior:

```yaml
# Directory to monitor for new subfolders
source_dir: "C:/path/to/source/folder"

# Directory where processed folders will be moved
destination_dir: "C:/path/to/destination/folder"

# Path to the Excel log file
excel_log_path: "C:/path/to/log/processing_log.xlsx"

# Retry settings for Excel file access
retry_attempts: 10
retry_delay_seconds: 2

# Naming convention pattern (regex)
naming_pattern: "^(\\d{4}|\\d{4}\\.\\d{2})\\s+(.+)$"
```

## Troubleshooting

### Script doesn't find any folders
- Check that `source_dir` path is correct in `config.yaml`
- Ensure the source directory contains subfolders (not just files)

### Excel file not updating
- Make sure the Excel file is closed
- Check that the path in `excel_log_path` is accessible
- Verify you have write permissions to the directory

### Task Scheduler not running
- Check that the Python path is correct in the Action
- Ensure "Start in" directory is set correctly
- Review Task Scheduler history for error messages
- Try running manually first to verify the script works

### Folders not being moved
- Verify `destination_dir` exists and is writable
- Check if folders with the same name already exist in destination
- Review script output for error messages

## License

MIT License - feel free to modify and use as needed.