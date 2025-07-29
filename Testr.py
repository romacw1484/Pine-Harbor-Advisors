import os
import datetime
from collections import defaultdict
from openpyxl import Workbook

# Set your directory path
# Root directory
root_dir = r"C:\Users\William Romac\My ShareSync\SharedDocs\IIS Operations"

# Store files by (filename, size)
files_seen = defaultdict(list)

# Walk through directory
for dirpath, dirnames, filenames in os.walk(root_dir):

    for filename in filenames:
        try:
            file_path = os.path.join(dirpath, filename)
            file_size = os.path.getsize(file_path)
            key = (filename.lower(), file_size)
            files_seen[key].append(file_path)
        except Exception:
            continue

# Prepare Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Duplicate Files"

# Headers
ws.append(["Filename", "Size (bytes)", "File Path", "Last Modified"])

# Write rows
for (name, size), paths in files_seen.items():
    if len(paths) > 1:
        for path in paths:
            try:
                mod_time = datetime.datetime.fromtimestamp(os.path.getmtime(path))
            except:
                mod_time = "Unknown"
            ws.append([name, size, path, mod_time])

# Save Excel file
output_path = r"C:\Users\William Romac\Desktop\duplicate_files_report.xlsx"
wb.save(output_path)
print(f"\n✅ Duplicate report saved to: {output_path}")
