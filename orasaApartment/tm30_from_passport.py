import os
import glob
from datetime import datetime

import pandas as pd
from passporteye import read_mrz
import pytesseract


# === Tell Python where Tesseract is (same as your old script) ===
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

# ================== CONFIG ==================
INPUT_FOLDER = r"C:\Users\therm\OneDrive\Desktop\399"
TEMPLATE_PATH = r"C:\Users\therm\OneDrive\Desktop\TM_template.xlsx"
OUTPUT_PATH = r"C:\Users\therm\OneDrive\Desktop\TM_filled.xlsx"
TEMPLATE_SHEET_NAME = "แบบแจ้งที่พัก Inform Accom"
# ============================================

# Clean manual mapping for names (you can add more later)
FIXED_NAMES = {
    "CC8600803": "SOE HTET AUNG",
    "CC8600680": "HTET LIN AUNG",
    "CC8605941": "THANT ZIN HTWE",
    "MI984297":   "KYAR PHYU",
}


def yyMMdd_to_ddMMyyyy(yyMMdd: str) -> str:
    """Convert MRZ date format YYMMDD -> DD/MM/YYYY."""
    if not yyMMdd or len(yyMMdd) != 6 or not yyMMdd.isdigit():
        return ""

    yy = int(yyMMdd[0:2])
    mm = int(yyMMdd[2:4])
    dd = int(yyMMdd[4:6])
    year = 1900 + yy if yy >= 30 else 2000 + yy

    try:
        d = datetime(year, mm, dd)
        return d.strftime("%d/%m/%Y")
    except ValueError:
        return ""


def extract_passport_data(image_path: str):
    """Use passporteye to read MRZ from passport image."""
    print(f"Processing: {image_path}")
    mrz = read_mrz(image_path)

    if mrz is None:
        print(f"  !! MRZ not found or unreadable for: {image_path}")
        return None

    data = mrz.to_dict()

    passport_no = (data.get("number") or "").replace("<", "").strip()
    sex = (data.get("sex") or "").strip()
    dob_raw = (data.get("date_of_birth") or "").strip()
    exp_raw = (data.get("expiration_date") or "").strip()
    surname_raw = (data.get("surname") or "").replace("<", " ").strip()
    names_raw = (data.get("names") or "").replace("<", " ").strip()

    # Build full name from MRZ (we will override using FIXED_NAMES anyway)
    name_parts = " ".join(p for p in [surname_raw, names_raw] if p).strip()
    full_name = name_parts if name_parts else (names_raw or surname_raw)

    # For Myanmar style, treat whole thing as First Name
    first_name = full_name
    middle_name = ""
    last_name = ""

    dob = yyMMdd_to_ddMMyyyy(dob_raw)
    exp = yyMMdd_to_ddMMyyyy(exp_raw)

    # Override with clean mapping if available
    if passport_no in FIXED_NAMES:
        first_name = FIXED_NAMES[passport_no]
        middle_name = ""
        last_name = ""

    return {
        "first_name": first_name,
        "middle_name": middle_name,
        "last_name": last_name,
        "sex": sex,
        "passport_no": passport_no,
        "dob": dob,
        "expiry": exp,
    }


def find_column(columns, keyword):
    """Find template column whose header contains the English keyword."""
    for c in columns:
        if keyword in str(c):
            return c
    return None


def main():
    # ---------- 1) Load template ----------
    template_df = pd.read_excel(TEMPLATE_PATH, sheet_name=TEMPLATE_SHEET_NAME)
    template_columns = list(template_df.columns)

    col_first = find_column(template_columns, "First Name")
    col_middle = find_column(template_columns, "Middle Name")
    col_last = find_column(template_columns, "Last Name")
    col_gender = find_column(template_columns, "Gender")
    col_passport = find_column(template_columns, "Passport No.")
    col_nationality = find_column(template_columns, "Nationality")
    col_dob = find_column(template_columns, "Birth Date")

    print("Using columns:")
    print("  First Name:", col_first)
    print("  Middle Name:", col_middle)
    print("  Last Name:", col_last)
    print("  Gender:", col_gender)
    print("  Passport No.:", col_passport)
    print("  Nationality:", col_nationality)
    print("  Birth Date:", col_dob)
    print()

    # ---------- 2) Collect images ----------
    image_paths = []
    for ext in ("*.jpg", "*.jpeg", "*.png", "*.JPG", "*.JPEG", "*.PNG"):
        image_paths.extend(glob.glob(os.path.join(INPUT_FOLDER, ext)))

    if not image_paths:
        print("No images found in:", INPUT_FOLDER)
        return

    # ---------- 3) Extract data ----------
    rows = []

    for img_path in sorted(image_paths):
        p_data = extract_passport_data(img_path)
        if p_data is None:
            continue

        row = {col: "" for col in template_columns}

        if col_first:
            row[col_first] = p_data["first_name"]
        if col_middle:
            row[col_middle] = p_data["middle_name"]
        if col_last:
            row[col_last] = p_data["last_name"]
        if col_gender:
            row[col_gender] = p_data["sex"]
        if col_passport:
            row[col_passport] = p_data["passport_no"]
        if col_nationality:
            row[col_nationality] = "MMR"  # always MMR
        if col_dob:
            row[col_dob] = p_data["dob"]

        rows.append(row)

    if not rows:
        print("No valid passport rows extracted.")
        return

    filled_df = pd.DataFrame(rows, columns=template_columns)

    # ---------- 4) Remove duplicate passports ----------
    if col_passport:
        before = len(filled_df)
        filled_df = filled_df.drop_duplicates(subset=[col_passport])
        after = len(filled_df)
        print(f"Removed {before - after} duplicate rows based on Passport No.")

    # ---------- 5) Save ----------
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        filled_df.to_excel(writer, sheet_name=TEMPLATE_SHEET_NAME, index=False)

    print(f"\n✅ Done! Saved filled file to:\n  {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
