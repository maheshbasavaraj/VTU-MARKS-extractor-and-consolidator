#!/usr/bin/env python3
# vtu_extract.py  –  Extract VTU marks from PDF result sheets
import os
import re
import sys
import logging
from pathlib import Path
import pandas as pd

##############################################################################
# SUBJECT CODES IN THE ORDER YOUR EXCEL FILE SHOULD HAVE
##############################################################################
SUBJECT_CODES_ORDER = [
    "BAD601",
    "BDS602",
    "BCD685",
    "BCSL606",
    "BPEK658",
    "BNSK658",
    "BYOK658",
    "BIKS609",
    "BDS613B",
    "BCV654C",
    "BAIL657C",
]

# Build expanded subject columns for Excel
SUBJECT_COLUMNS = []
for code in SUBJECT_CODES_ORDER:
    SUBJECT_COLUMNS.extend([f"{code}_I", f"{code}_E", f"{code}_T"])

##############################################################################
# Regex to find subject lines
##############################################################################
"""SUBJECT_REGEX = re.compile(
    # Match the start of a subject line
    r"^(?P<code>[A-Z0-9]+)\s+(?P<name>.+?)\s+"
    # Match the three marks (internal, external, total)
    r"(?P<internal>\d+)\s+(?P<external>\d+)\s+(?P<total>\d+)\s+"
    # Match the result (P/F/A/W/X) and the date
    r"(?P<result>[PFAXW])\s+(?P<date>\d{4}-\d{2}-\d{2})$",
    re.MULTILINE,
)"""

##############################################################################
# NEW regex (overwrites the old SUBJECT_REGEX)
##############################################################################
SUBJECT_REGEX = re.compile(
    # 1. subject code
    r"(?P<code>BAD601|BDS602|BCD685|BCSL606|BPEK658|BNSK658|BYOK658|BIKS609|BDS613B|BCV654C|BAIL657C)\s+"
    # 2. subject name (allow letters, spaces, &/(), possibly wrapping line)
    r"(?P<name>[A-Z0-9\s/&()(-]+?)\s+"
    # 3. three space-separated numbers: internal, external, total
    r"(?P<internal>\d+)\s+(?P<external>\d+)\s+(?P<total>\d+)\s+"
    # 4. result letter + date
    r"(?P<result>[PFAXW])\s+(?P<date>\d{4}-\d{2}-\d{2})",
    re.IGNORECASE,
)
###############################################################################
# Setup logging
###############################################################################
log_formatter = logging.Formatter("[%(levelname)s] %(message)s")
log = logging.getLogger()
log.setLevel(logging.INFO)

# Log to file
file_handler = logging.FileHandler("marks_extraction.log", mode="w")
file_handler.setFormatter(log_formatter)
log.addHandler(file_handler)

# Log to console
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(log_formatter)
log.addHandler(console_handler)


###############################################################################
# Text extraction helper
###############################################################################
def pdf_to_text(pdf_path):
    # 1st try pdfplumber (better for tables)
    try:
        import pdfplumber

        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        if text.strip():
            return text
    except ImportError:
        log.warning("pdfplumber not found. Install with: pip install pdfplumber")
    except Exception as e:
        log.warning(f"pdfplumber failed on {pdf_path.name} ({e})")

    # Fallback to pypdf
    try:
        from PyPDF2 import PdfReader

        reader = PdfReader(pdf_path)
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
        return text
    except ImportError:
        log.warning("PyPDF2 not found. Install with: pip install PyPDF2")
    except Exception as e:
        log.warning(f"PyPDF2 failed on {pdf_path.name} ({e})")
    return ""


###############################################################################
# Parse one PDF and return a dict  {'subjectCode': total_marks}
###############################################################################
def parse_pdf(pdf_path):
    log.info(f"Parsing {pdf_path.name}...")
    text = pdf_to_text(pdf_path)
    if not text:
        log.warning(f"No text extracted from: {pdf_path.name}")
        return None, None, None

    # USN and name are always on separate lines
    usn_line = re.search(r"(?:University Seat Number|USN)[:\s]*([0-9A-Z]+)", text, re.I)
    name_line = re.search(r"(?:Student Name)[\s:]*([^\n\r]+)", text, re.I)

    usn = (usn_line.group(1) if usn_line else "").strip().upper()
    name = (name_line.group(1) if name_line else "").strip().title()

    if not usn:
        log.warning(f"Could not find USN in {pdf_path.name}")
    if not name:
        log.warning(f"Could not find Student Name in {pdf_path.name}")

    # Now get marks for each subject
    marks = {}
    # First, join lines where a subject name wraps
    clean_text = re.sub(r"\n(?!([A-Z0-9]+\s))", " ", text)

    for match in SUBJECT_REGEX.finditer(clean_text):
        code = match.group("code").upper()
        if code in SUBJECT_CODES_ORDER:
            try:
                marks[f"{code}_I"] = int(match.group("internal"))
                marks[f"{code}_E"] = int(match.group("external"))
                marks[f"{code}_T"] = int(match.group("total"))
            except (ValueError, IndexError):
                log.warning(f"Could not parse marks for {code} in {pdf_path.name}")
                marks[f"{code}_I"] = None
                marks[f"{code}_E"] = None
                marks[f"{code}_T"] = None

    log.info(f"  USN: {usn}, Name: {name}, Marks: {len(marks)//3} subjects found")
    return usn, name, marks


###############################################################################
# Walk folder & build dataframe
###############################################################################
def collect_marks(pdf_dir: Path):
    serial = 0
    rows = []
    pdf_files = sorted(pdf_dir.glob("*.pdf"))
    log.info(f"Found {len(pdf_files)} PDF files in {pdf_dir}")
    for pdf_file in pdf_files:
        serial += 1
        usn, name, mark_dict = parse_pdf(pdf_file)
        if not usn:
            log.warning(f"Could not read USN from {pdf_file.name}, using blank.")
            usn = ""
        if not name:
            log.warning(f"Could not read Name from {pdf_file.name}, using blank.")
            name = ""

        row = {"SerialNo": serial, "Name": name, "USN": usn}
        if not mark_dict:
            log.warning(f"No marks found for {pdf_file.name} (USN: {usn})")
            mark_dict = {}
        for subj_col in SUBJECT_COLUMNS:
            row[subj_col] = mark_dict.get(subj_col, None)
        rows.append(row)
    return rows


###############################################################################
# Excel output
###############################################################################
def write_excel(rows, output_path):
    if not rows:
        log.warning("No data was collected, skipping Excel export.")
        return
    df = pd.DataFrame(rows, columns=["SerialNo", "Name", "USN"] + SUBJECT_COLUMNS)
    try:
        df.to_excel(output_path, index=False, engine="openpyxl")
        log.info(f"Successfully saved Excel file to: {output_path}")
    except Exception as e:
        log.error(f"Failed to write Excel file: {e}")
        log.error("Please ensure you have 'openpyxl' installed: pip install openpyxl")


###############################################################################
# main
###############################################################################
if __name__ == "__main__":
    log.info("--- Starting VTU Marks Extraction ---")
    pdf_folder = (
        Path.cwd() / "6thsemresult"
    )  # run from inside PDF folder or change here
    output_excel = Path.cwd() / "extracted_marks.xlsx"

    if not pdf_folder.is_dir():
        log.error(f"PDF folder not found at: {pdf_folder}")
        log.error(
            "Please make sure the '6thsemresult' folder exists and contains your PDF files."
        )
        sys.exit(1)

    rows = collect_marks(pdf_folder)
    write_excel(rows, output_excel)
    log.info("--- Extraction Complete ---")
