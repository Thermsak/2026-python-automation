import os, re
from PyPDF2 import PdfMerger

folder_path = r'C:\Users\therm\OneDrive\Desktop\2025-03 ALLY\For uploaded'
output_folder = folder_path

# Match both: ... - Report - PO.pdf and ... - Report - PO-signed.pdf
merged_with_po_pattern = re.compile(r'^(IV\d{9}) - ([A-Z]+) - Report - PO(?:-signed)?\.pdf$', re.IGNORECASE)

# Pattern for receipt: RE-0253 - TC.pdf
receipt_pattern = re.compile(r'^RE-\d+ - ([A-Z]+)\.pdf$', re.IGNORECASE)

# Build receipt file lookup
receipt_files = {}
for f in os.listdir(folder_path):
    m = receipt_pattern.match(f)
    if m:
        branch = m.group(1).strip()
        receipt_files[branch] = os.path.join(folder_path, f)

# Combine PO-merged file with receipt
for f in os.listdir(folder_path):
    m = merged_with_po_pattern.match(f)
    if not m:
        continue

    full_inv_name, branch = m.groups()
    merged_with_po_path = os.path.join(folder_path, f)
    receipt_path = receipt_files.get(branch)

    if not receipt_path:
        print(f'⚠️ No receipt for branch "{branch}" — skipping {f}')
        continue

    merger = PdfMerger()
    merger.append(merged_with_po_path)
    merger.append(receipt_path)

    # Include "-signed" in output name if input file has it
    suffix = "-signed" if "-signed" in f else ""
    out_name = f"{full_inv_name} - {branch} - Report - PO{suffix} - Receipt.pdf"
    out_path = os.path.join(output_folder, out_name)
    merger.write(out_path)
    merger.close()

    print(f'✅ Final merged with receipt: {out_name}')
