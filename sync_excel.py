#!/usr/bin/env python3
"""
Script to synchronize Excel data with the website.
Reads e2.xlsx and generates data.js file.
"""

import json
import pandas as pd
import sys
import os
from pathlib import Path

def sync_excel_to_js():
    """Read Excel file and generate data.js"""
    
    excel_file = 'e2.xlsx'
    output_file = 'data.js'
    
    # Check if Excel file exists
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} not found!")
        return False
    
    try:
        # Read the TOTALLIST sheet
        print(f"Reading {excel_file}...")
        df = pd.read_excel(excel_file, sheet_name='TOTALLIST')
        
        # Convert DataFrame to list of dictionaries
        # Replace NaN values with empty strings or None
        data = df.where(pd.notna(df), None).to_dict('records')
        
        # Convert to JavaScript format
        js_content = "var allData = [\n"
        
        for i, row in enumerate(data):
            js_content += "  {\n"
            for key, value in row.items():
                if value is None or (isinstance(value, float) and pd.isna(value)):
                    js_value = '""'
                elif isinstance(value, (int, float)):
                    js_value = str(value)
                elif isinstance(value, str):
                    # Escape quotes and newlines
                    escaped_value = value.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n')
                    js_value = f'"{escaped_value}"'
                else:
                    js_value = json.dumps(value)
                
                js_content += f'    "{key}": {js_value},\n'
            
            # Remove trailing comma from last property
            js_content = js_content.rstrip(',\n') + '\n'
            js_content += "  }"
            
            if i < len(data) - 1:
                js_content += ",\n"
            else:
                js_content += "\n"
        
        js_content += "];\n\n"
        js_content += "// Ensure it's available on window object\n"
        js_content += "if (typeof window !== 'undefined') {\n"
        js_content += "    window.allData = allData;\n"
        js_content += "}\n"
        
        # Write to data.js
        print(f"Writing to {output_file}...")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(js_content)
        
        print(f"✓ Successfully synchronized {len(data)} rows from {excel_file} to {output_file}")
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    success = sync_excel_to_js()
    sys.exit(0 if success else 1)

