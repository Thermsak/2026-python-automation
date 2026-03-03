import os
import re
from PyPDF2 import PdfMerger

folder_path = r'C:\Users\therm\OneDrive\Desktop\2025-07 ALLY'
output_folder = folder_path

# Pattern to detect invoice-related files (e.g. Report, Receipt, Receipt+Report, etc.)
invoice_pattern = re.compile(r'^(IV\d{9}) - ([A-Z]{2,}) - ([^\.]+)\.pdf$', re.IGNORECASE)

# Pattern to match PO file (e.g. ALLY-AMR-PO25001707 ∙ บริษัท เอ็นแพลน จำกัด.pdf)
po_pattern = re.compile(r'^ALLY-([A-Z]{2,})-PO\d+.*\.pdf$', re.IGNORECASE)

# Build PO file lookup by branch
po_files = {}
for file in os.listdir(folder_path):
    match = po_pattern.match(file)
    if match:
        branch = match.group(1).strip()
        po_files[branch] = os.path.join(folder_path, file)

# Process invoice/receipt/report files
for file in os.listdir(folder_path):
    match = invoice_pattern.match(file)
    if not match:
        continue

    invoice_id, branch, current_label = match.groups()
    inv_path = os.path.join(folder_path, file)
    po_path = po_files.get(branch)

    if not po_path:
        print(f'⚠️ No PO found for branch "{branch}" — skipping {file}')
        continue

    # Combine invoice and PO
    merger = PdfMerger()
    merger.append(inv_path)
    merger.append(po_path)

    # Create new name by adding +PO to the existing label
    new_label = current_label + '+PO'
    output_name = f"{invoice_id} - {branch} - {new_label}.pdf"
    output_path = os.path.join(output_folder, output_name)

    with open(output_path, 'wb') as f_out:
        merger.write(f_out)
    merger.close()

    print(f'✅ Created: {output_name}')
