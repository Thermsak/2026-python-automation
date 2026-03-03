import re
import csv
import pdfplumber

# Define the groups for categorization
groups = ['TC Ekamai', 'TCR', 'TCP', 'IMP', 'PLN', 'TS', 'SRM', 'SRP', 'APP', 'AMR', 'HPA','SRS','CDC']

# Initialize a dictionary for spending summary
spending_summary = {}

# Read the PDF file
pdf_file = r'C:\Users\therm\OneDrive\Desktop\Binder1.pdf'
with pdfplumber.open(pdf_file) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        if text:  # Ensure text is not empty
            for line in text.splitlines():
                line = line.strip()
                if '฿' in line:  # Check if the line contains a spending amount
                    # Check if the line matches any group
                    matched_group = None
                    for group in groups:
                        if group in line:
                            matched_group = group
                            break

                    # If no group matches, use the first 8 characters as the group
                    if not matched_group:
                        matched_group = line[:8].strip()

                    # Extract spending amounts
                    spending_matches = re.findall(r'฿([\d,]+(?:\.\d+)?)', line)
                    if spending_matches:
                        spending_amount = sum(float(amount.replace(',', '')) for amount in spending_matches)
                        if matched_group not in spending_summary:
                            spending_summary[matched_group] = 0.0
                        spending_summary[matched_group] += spending_amount

# Write the results to a CSV file
output_file = r'C:\Users\therm\OneDrive\Desktop\spending_summary.csv'
with open(output_file, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerow(['Group', 'Spending'])
    for group, spending in spending_summary.items():
        writer.writerow([group, spending])

print(f"Spending summary has been exported to {output_file}")
