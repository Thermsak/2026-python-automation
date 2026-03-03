import os
import csv

####------ Define the path to the master folder -------####
master_folder_path = r"H:\My Drive\# Billing - Account\2026-02\Payment tmp"

# Initialize a dictionary to store the count of different file types in each subfolder
folder_file_counts = {}

# Iterate over each subfolder in the master folder
for subfolder in os.listdir(master_folder_path):
    subfolder_path = os.path.join(master_folder_path, subfolder)
    if os.path.isdir(subfolder_path):
        # Initialize counters for file types
        jpg_count = 0
        pdf_count = 0
        png_count = 0
        other_count = 0

        # Count the number of files in the subfolder and separate them by type
        for f in os.listdir(subfolder_path):
            if os.path.isfile(os.path.join(subfolder_path, f)) and not f.startswith('.') and not f in ['desktop.ini', 'Thumbs.db']:
                # Check the file extension and count accordingly
                if f.lower().endswith('.jpg'):
                    jpg_count += 1
                elif f.lower().endswith('.pdf'):
                    pdf_count += 1
                elif f.lower().endswith('.png'):
                    png_count += 1
                else:
                    other_count += 1

        # Store the counts in the dictionary
        folder_file_counts[subfolder] = {'jpg': jpg_count, 'pdf': pdf_count, 'png': png_count, 'other': other_count}

# Sort the dictionary by subfolder names
sorted_folders = sorted(folder_file_counts.items())

# Print the output in table format
print("Folder Name - JPG Count - PDF Count - PNG Count - Other File Count")
for folder, counts in sorted_folders:
    print(f"{folder} - {counts['jpg']} - {counts['pdf']} - {counts['png']} - {counts['other']}")

# Save the output to a CSV file
output_file = os.path.join(master_folder_path, "file_type_counts.csv")
with open(output_file, "w", encoding="utf-8", newline='') as csvfile:
    csvwriter = csv.writer(csvfile)
    # Write the header
    csvwriter.writerow(["Folder Name", "JPG Count", "PDF Count", "PNG Count", "Other File Count"])
    # Write the data
    for folder, counts in sorted_folders:
        csvwriter.writerow([folder, counts['jpg'], counts['pdf'], counts['png'], counts['other']])

print(f"\nFile counts saved to {output_file}")
