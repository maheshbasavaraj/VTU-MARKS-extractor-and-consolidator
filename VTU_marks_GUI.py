#!/usr/bin/env python3
"""
VTU Marks Extractor – GUI file & folder selection.
No subject codes are hard-coded.
"""
import sys, re, logging
from pathlib import Path
import pandas as pd

# ----------------  GUI helpers  ----------------
try:
    from tkinter import Tk
    from tkinter.filedialog import askopenfilename, askdirectory
except ImportError:
    print("tkinter not found – please run on a desktop OS", file=sys.stderr)
    sys.exit(1)

# ----------------  logging  ----------------
logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("marks_extraction.log", "w"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger()


# ----------------  user picks files  ----------------
def prompt_subjects_file() -> Path:
    root = Tk()
    root.withdraw()
    f = askopenfilename(
        title="Choose your subject list file (one code per line)",
        filetypes=[("Text Files", "*.txt")],
    )
    root.destroy()
    if not f:
        sys.exit("No subject list chosen – exiting.")
    return Path(f)


def prompt_pdfs_folder() -> Path:
    root = Tk()
    root.withdraw()
    d = askdirectory(title="Choose folder that contains the VTU PDF files")
    root.destroy()
    if not d:
        sys.exit("No PDF folder chosen – exiting.")
    return Path(d)


# ----------------  load subject list  ----------------
def load_subject_list(txt_path: Path) -> list[str]:
    codes = [
        line.split("#")[0].strip().upper()
        for line in txt_path.read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.strip().startswith("#")
    ]
    if not codes:
        sys.exit("Subject list is empty.")
    log.info(f"Loaded {len(codes)} subject codes: {codes}")
    return codes


# ----------------  constant builder ----------------
SUBJECT_FILE = prompt_subjects_file()
SUBJECTS = load_subject_list(SUBJECT_FILE)

SUBJECT_COLS = []  # 'BAD601_I', 'BAD601_E', ...
for subj in SUBJECTS:
    SUBJECT_COLS.extend([f"{subj}_I", f"{subj}_E", f"{subj}_T"])


# ---------------- text extraction ----------------
def pdf_text_extract(pdf_path: Path) -> str:
    try:
        import pdfplumber

        with pdfplumber.open(pdf_path) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    except ImportError:
        import PyPDF2

        return "\n".join(
            PyPDF2.PdfReader(pdf_path).pages[i].extract_text() or ""
            for i in range(len(PyPDF2.PdfReader(pdf_path).pages))
        )


# ---------------- parsing ----------------
USER_SUBJECT_PATTERN = "|".join(map(re.escape, SUBJECTS))

SUBJECT_REGEX = re.compile(
    rf"(?P<code>{USER_SUBJECT_PATTERN})\s+"
    rf"(?P<name>[A-Z\s/&()\-]+?)\s+"
    rf"(?P<internal>\d+)\s+"
    rf"(?P<external>\d+)\s+"
    rf"(?P<total>\d+)\s+"
    rf"(?P<result>[PFAXW])\s+(?P<date>\d{{4}}-\d{{2}}-\d{{2}})",
    re.I,
)


def parse_pdf(pdf_path):
    text = pdf_text_extract(pdf_path)
    if not text:
        return None, None, None

    usn = re.search(r"(?:University Seat Number|USN)[:\s]*([0-9A-Z]+)", text, re.I)
    usn = usn.group(1).strip().upper() if usn else ""
    name = re.search(r"(?:Student Name)[:\s]*([^\r\n]+)", text, re.I)
    name = name.group(1).strip().title() if name else ""

    clean = re.sub(r"\n(?![A-Z0-9]+\s)", " ", text)
    marks = {}
    for m in SUBJECT_REGEX.finditer(clean):
        s = m.group("code").upper()
        if s in SUBJECTS:
            marks[f"{s}_I"] = int(m.group("internal"))
            marks[f"{s}_E"] = int(m.group("external"))
            marks[f"{s}_T"] = int(m.group("total"))
    return usn, name, marks


# ---------------- collector ----------------
PDF_FOLDER = prompt_pdfs_folder()
rows = []
serial = 0
for pdf in sorted(PDF_FOLDER.glob("*.pdf")):
    serial += 1
    usn, name, mm = parse_pdf(pdf)
    row = {"SerialNo": serial, "Name": name or "", "USN": usn or ""}
    for col in SUBJECT_COLS:
        row[col] = mm.get(col) if mm else None
    rows.append(row)

if not rows:
    sys.exit("No usable data extracted from the chosen folder.")

# ---------------- Excel writer with 2-row header ----------------
OUTPUT = PDF_FOLDER.with_suffix("").parent / "extracted_marks.xlsx"

with pd.ExcelWriter(OUTPUT, engine="openpyxl", mode="w") as writer:
    df = pd.DataFrame(rows, columns=["SerialNo", "Name", "USN"] + SUBJECT_COLS)
    df.to_excel(writer, index=False, header=False, startrow=2)
    ws = writer.sheets["Sheet1"]

    # row 1 : codes
    header1 = ["", "", ""] + SUBJECTS
    # row 2 : I E T under each
    header2 = ["SerialNo", "Name", "USN"]
    for _ in SUBJECTS:
        header2.extend(["I", "E", "T"])

    for c, val in enumerate(header1, 1):
        ws.cell(row=1, column=c, value=val)
    for c, val in enumerate(header2, 1):
        ws.cell(row=2, column=c, value=val)

log.info(f"All done – saved to: {OUTPUT}")
