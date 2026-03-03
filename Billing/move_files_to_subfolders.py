import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import os
import re
import shutil

# ----------------------------
# TESSERACT CONFIG
# ----------------------------
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'

eng_path = os.path.join(os.environ['TESSDATA_PREFIX'], 'eng.traineddata')
if not os.path.isfile(eng_path):
    raise FileNotFoundError(f"eng.traineddata not found at: {eng_path}")

# ----------------------------
# PATH CONFIG
# ----------------------------
main_folder = r'H:\My Drive\# Billing - Account\2026-02\Payment tmp'

# ----------------------------
# HELPERS
# ----------------------------
IMAGE_EXTS = ('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')

def extract_memo_text(image_path: str):
    """
    Extract 5-digit number followed by underscore from Slip_ OCR text.
    Example target: 12345_
    """
    img = Image.open(image_path)

    img = img.convert('L')
    img = ImageEnhance.Contrast(img).enhance(2)
    img = img.filter(ImageFilter.SHARPEN)

    custom_config = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(img, config=custom_config)

    print(f"\n--- Extracted Text from IMAGE: {image_path} ---")
    print(text)

    match = re.search(r'(\d{5})_', text)
    return match.group(1) if match else None

def find_subfolder_by_prefix(base_folder: str, prefix: str):
    """Find the first subfolder whose name starts with prefix."""
    for d in os.listdir(base_folder):
        full = os.path.join(base_folder, d)
        if os.path.isdir(full) and d.startswith(prefix):
            return full
    return None

def find_subfolder_by_suffix(base_folder: str, suffix: str):
    """Find the first subfolder whose name ends with suffix."""
    for d in os.listdir(base_folder):
        full = os.path.join(base_folder, d)
        if os.path.isdir(full) and d.endswith(suffix):
            return full
    return None

def extract_identifier_from_pdf(filename: str):
    """
    REQUIRED (your request):
    - For PDF names like: rec_B2601040_Cherryinmyjam.pdf
      return: Cherryinmyjam (the part AFTER the 7 digits + underscore)

    Also supports:
    - rec_A2601040_Cherryinmyjam.pdf
    - rec_B12345_Cherryinmyjam.pdf (5 digits, still ok)
    - 50ทวิ_A2601040_Cherryinmyjam.pdf
    - 50ทวิ_A12345_Cherryinmyjam.pdf

    If there is no suffix part (no underscore after digits), return None.
    """
    pattern = re.compile(
        r'^(?:50ทวิ_)?'          # optional Thai prefix
        r'(?:rec|50ทวิ)_'        # rec_ or 50ทวิ_
        r'[AB]'                  # A or B
        r'(\d{5}|\d{7})'         # 5 or 7 digits
        r'_(.+?)'                # underscore + identifier (required)
        r'\.pdf$',
        re.IGNORECASE
    )

    m = pattern.search(filename)
    if not m:
        return None

    identifier = m.group(2)

    # Clean common accidental trailing spaces
    identifier = identifier.strip()

    return identifier if identifier else None

def safe_move_file(src_path: str, dst_folder: str):
    filename = os.path.basename(src_path)
    dst_path = os.path.join(dst_folder, filename)

    if os.path.isdir(dst_path):
        print(f"Invalid destination (is a directory): {dst_path}")
        return False

    print(f"Moving: {src_path} -> {dst_path}")
    shutil.move(src_path, dst_path)
    print(f"Moved file to: {dst_path}")
    return True

# ----------------------------
# MAIN PROCESS
# ----------------------------
for filename in os.listdir(main_folder):
    file_path = os.path.join(main_folder, filename)

    # skip folders
    if os.path.isdir(file_path):
        continue

    lower = filename.lower()

    # ---------- IMAGES ----------
    if lower.endswith(IMAGE_EXTS):
        memo = extract_memo_text(file_path)

        if memo:
            subfolder = find_subfolder_by_prefix(main_folder, memo)
            if subfolder:
                print(f"Subfolder identified (prefix={memo}): {subfolder}")
                safe_move_file(file_path, subfolder)
            else:
                print(f"Subfolder starting with '{memo}' not found for image: {filename}")
        else:
            print(f"5-digit number followed by underscore not found in IMAGE: {filename}")

    # ---------- PDFS ----------
    elif lower.endswith('.pdf'):
        identifier = extract_identifier_from_pdf(filename)

        if identifier:
            print(f"\nExtracted identifier from PDF filename '{filename}': {identifier}")

            subfolder = find_subfolder_by_suffix(main_folder, identifier)
            if subfolder:
                print(f"Subfolder identified (endswith={identifier}): {subfolder}")
                safe_move_file(file_path, subfolder)
            else:
                print(f"Subfolder ending with '{identifier}' not found for PDF: {filename}")
        else:
            print(f"PDF filename pattern not matched / missing suffix identifier (skipped): {filename}")
