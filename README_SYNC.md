# Excel Synchronization Setup

This website can automatically synchronize with the Excel file (`e2.xlsx`). 

## Available Scripts

### 1. `update_excel.py` - Update Excel File
Updates the Excel file itself by copying net total values from individual sheets to the estimate column in the TOTALLIST sheet.

**Usage:**
```bash
python3 update_excel.py
```

### 2. `sync_excel.py` - Sync Excel to Website
Reads the Excel file and generates `data.js` for the website.

**Usage:**
```bash
python3 sync_excel.py
```

### 3. `watch_excel.py` - Auto-Watch Mode (Recommended)
Automatically watches for Excel file changes and syncs to the website.

**Usage:**
```bash
python3 watch_excel.py
```

## Synchronization Methods

### Method 1: Manual Sync (One-time)

Run the sync script whenever you update the Excel file:

```bash
python3 sync_excel.py
```

This will read `e2.xlsx` and regenerate `data.js`.

### Method 2: Automatic Watch Mode (Recommended)

Keep the script running to automatically sync whenever the Excel file changes:

```bash
python3 watch_excel.py
```

This will:
- Watch for changes to `e2.xlsx`
- Automatically regenerate `data.js` when the file is saved
- Display a message when sync completes

**Note:** You'll need to refresh your browser to see the changes after sync.

## Complete Workflow

### Step 1: Update Excel Estimates (if needed)
If you've updated individual sheets and want to update the TOTALLIST estimates:

```bash
python3 update_excel.py
```

### Step 2: Sync Excel to Website

**Option A - Manual:**
```bash
python3 sync_excel.py
```

**Option B - Automatic (Recommended):**
```bash
python3 watch_excel.py
```

### Step 3: Start Web Server
In another terminal:

```bash
python3 -m http.server 8001
```

### Step 4: Open Browser
Navigate to `http://localhost:8001`

### Step 5: Edit and Refresh
- Edit the Excel file (`e2.xlsx`)
- If using watch mode, it will auto-sync
- Refresh your browser to see changes

## Requirements

Install required Python packages:

```bash
pip install pandas openpyxl watchdog
```

Or:

```bash
pip3 install pandas openpyxl watchdog
```

## Troubleshooting

- If `e2.xlsx` is not found, make sure it's in the same directory as the scripts
- If you get import errors, install the required packages (see Requirements above)
- The sync script reads from the 'TOTALLIST' sheet - make sure this sheet exists
- If port 8001 is in use, use a different port: `python3 -m http.server 8002`

