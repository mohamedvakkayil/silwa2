#!/usr/bin/env python3
"""
File watcher to automatically sync Excel file when it changes.
Requires watchdog package: pip install watchdog
"""

import time
import sys
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from sync_excel import sync_excel_to_js

class ExcelChangeHandler(FileSystemEventHandler):
    """Handle Excel file changes"""
    
    def __init__(self):
        self.last_modified = 0
        self.debounce_seconds = 2  # Wait 2 seconds after last change
    
    def on_modified(self, event):
        if event.src_path.endswith('e2.xlsx'):
            current_time = time.time()
            # Debounce: only process if enough time has passed since last modification
            if current_time - self.last_modified > self.debounce_seconds:
                print(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] Excel file changed, syncing...")
                sync_excel_to_js()
                self.last_modified = current_time
                print("✓ Sync complete. Refresh your browser to see changes.\n")

def watch_excel():
    """Watch for Excel file changes"""
    event_handler = ExcelChangeHandler()
    observer = Observer()
    
    # Watch current directory
    watch_path = '.'
    observer.schedule(event_handler, watch_path, recursive=False)
    observer.start()
    
    print("=" * 60)
    print("Excel File Watcher Started")
    print("=" * 60)
    print(f"Watching: {os.path.abspath(watch_path)}/e2.xlsx")
    print("Press Ctrl+C to stop watching")
    print("=" * 60)
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\n\nStopping file watcher...")
    
    observer.join()

if __name__ == '__main__':
    import os
    try:
        watch_excel()
    except ImportError:
        print("Error: watchdog package not installed.")
        print("Install it with: pip install watchdog")
        sys.exit(1)

