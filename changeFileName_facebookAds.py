import os
from PyPDF2 import PdfReader

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file."""
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    return text

def rename_files_in_folder(folder_path):
    """Rename PDF files by keeping only the date and transaction part, then appending the reference number."""
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            file_path = os.path.join(folder_path, filename)
            
            # Extract text from the PDF file
            text = extract_text_from_pdf(file_path)
            
            # Find the reference number (assuming it's in the format 'Reference Number: 7N2W498NU2')
            start = text.find("Reference Number:")
            if start != -1:
                reference_number = text[start+len("Reference Number:"):].split()[0]
                
                # Extract only the '2024-08-09T09-42 Transaction' part
                date_transaction_part = filename.split("Transaction")[0] + "Transaction"
                
                # Rename the file if the reference number is found
                if reference_number:
                    new_file_name = f"{date_transaction_part}_{reference_number}.pdf"
                    new_file_path = os.path.join(folder_path, new_file_name)
                    os.rename(file_path, new_file_path)
                    print(f"Renamed '{filename}' to '{new_file_name}'")
                else:
                    print(f"No reference number found in '{filename}'")
            else:
                print(f"No 'Reference Number' found in '{filename}'")

# Usage
folder_path = r'H:\My Drive\# Billing - Account\2024-09\Payment VAT\2024-09 LOFT'  # Your target folder path
rename_files_in_folder(folder_path)
