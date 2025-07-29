import os
import hashlib
from collections import defaultdict
from openpyxl import Workbook
import datetime

# Directory to scan (change this if needed)
root_dir = r"C:\Users\yourname\yourpath 

# Store files by SHA-256 hash
files_by_hash = defaultdict(list)

# Function to compute file hash
def compute_hash(filepath):
    hasher = hashlib.sha256()
    try:
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hasher.update(chunk)
        return hasher.hexdigest()
    except:
        return None

# Traverse all files
for dirpath, _, filenames in os.walk(root_dir):
    for filename in filenames:
        file_path = os.path.join(dirpath, filename)
        file_hash = compute_hash(file_path)
        if file_hash:
            files_by_hash[file_hash].append(file_path)

# Create Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "True Duplicates"

# Write headers
ws.append(["Hash", "Filename", "File Path", "Size (bytes)", "Last Modified"])

# Write only confirmed duplicates
for file_hash, paths in files_by_hash.items():
    if len(paths) > 1:
        for path in paths:
            try:
                name = os.path.basename(path)
                size = os.path.getsize(path)
                mod_time = datetime.datetime.fromtimestamp(os.path.getmtime(path))
                ws.append([file_hash, name, path, size, mod_time])
            except:
                continue

# Save Excel report to Desktop
output_path = r"C:\Users\your-name\your-path\filename.xlsx 
wb.save(output_path)
print(f"\n✅ Excel report saved to: {output_path}")
