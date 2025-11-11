#!/bin/bash
# Quick sync script - updates Excel and syncs to website

echo "=== Silwa Tower Estimation - Sync Script ==="
echo ""

# Step 1: Update Excel estimates (optional)
read -p "Update Excel estimates from sheets? (y/n) " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    echo "Updating Excel file..."
    python3 update_excel.py
    echo ""
fi

# Step 2: Sync Excel to website
echo "Syncing Excel to website..."
python3 sync_excel.py
echo ""

echo "✓ Sync complete! Refresh your browser to see changes."

