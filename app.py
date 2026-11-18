import sys
import os

# ── PyInstaller path fix ──────────────────────────────────────────────────────
# When frozen as .exe, files are unpacked to a temp _MEIPASS folder.
# This function finds the correct path in both .py and .exe modes.
def resource_path(relative):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative)
    return os.path.join(os.path.abspath("."), relative)

# When running as exe, set working dir to folder where exe lives
# so uploads/ and outputs/ are created next to the exe (not in temp folder)
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))

from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
import pandas as pd
import re
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4

# Tell Flask where the templates folder is (important for exe)
app = Flask(__name__, template_folder=resource_path("templates"))

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ==============================================================================
# ▼▼▼  EDIT THIS SECTION ONLY  ▼▼▼
# ==============================================================================

YEAR_MAP = {
    "22": "4th Year",
    "23": "3rd Year",
    "24": "2nd Year",
    "25": "1st Year",
}

YEAR_SLICE = {
    "22": 3,
    "23": 4,
    "24": 3,
    "25": 3,
}

DEPT_MAP = {
    # ── Year 22 format  (3-digit numeric) ────────────────────────────────────
    "101": "AERO",
    "103": "CIVIL",
    "104": "CSE",
    "105": "EEE",
    "106": "ECE",
    "114": "MECH",
    "115": "MCO",
    "203": "CHEM",
    "205": "IT",
    "243": "AIDS",
    "244": "CSBS",

    # ── Year 23 format  (4-char with U prefix) ───────────────────────────────
    "U101": "AERO",
    "U103": "CIVIL",
    "U104": "CSE",
    "U105": "EEE",
    "U106": "ECE",
    "U114": "MECH",
    "U115": "MCO",
    "U203": "CHEM",
    "U205": "IT",
    "U243": "AIDS",
    "U244": "CSBS",

    # ── Year 24 / 25 format  (3-letter alphabetic) ───────────────────────────
    "BCS":  "CSE",
    "BEC":  "ECE",
    "BME":  "MECH",
    "BCE":  "CIVIL",
    "BEE":  "EEE",
    "BIT":  "IT",
    "BAE":  "AERO",
    "BCH":  "CHEM",
    "BMC":  "MCO",
    "BAD":  "AIDS",
    "BCB":  "CSBS",
    "BML":  "AIML",

    # ── PG programmes ─────────────────────────────────────────────────────────
    "MBA":  "MBA",
    "MCA":  "MCA",
    "MCS":  "MCS",
    "MEE":  "MEE",
}

# ==============================================================================
# ▲▲▲  STOP EDITING HERE  ▲▲▲
# ==============================================================================


def parse_reg_no(reg_no):
    reg_no = reg_no.strip()
    try:
        yr         = reg_no[4:6]
        year_label = YEAR_MAP.get(yr, f"Year({yr})")
        suffix     = reg_no[6:]
        n          = YEAR_SLICE.get(yr, 3)
        code       = suffix[0:n].upper()
        dept_label = DEPT_MAP.get(code, f"UNKNOWN({code})")
        return dept_label, year_label
    except Exception:
        return "UNKNOWN", "Unknown"

def get_department(reg_no): return parse_reg_no(reg_no)[0]
def get_year(reg_no):       return parse_reg_no(reg_no)[1]


ABSENT_MARKS = ("AB", "A", "ABSENT", "✖", "X")
# Cell texts that are attendance statuses, never student names
NON_NAME_WORDS = {"P", "PR", "PRESENT", "ABSENT", "AB", "A", "X"}

# Register number: 4-digit college code + 2-digit admission year + dept code + serial.
# e.g. 610823U104001. Pure-digit numbers must be 12+ digits (year-22 style) so
# that 10-digit answer-book numbers are not mistaken for register numbers.
REG_NO_RE = re.compile(r"^\d{6}[A-Z0-9]{4,10}$")
# Letters, spaces, dots (initials), hyphens, apostrophes; at least 2 letters.
NAME_RE = re.compile(r"^[A-Za-z][A-Za-z .'\-]*$")


def normalize_reg_no(text):
    return re.sub(r"\s+", "", text).upper()


def looks_like_reg_no(text):
    t = normalize_reg_no(text)
    if not REG_NO_RE.match(t):
        return False
    return not (t.isdigit() and len(t) < 12)


def looks_like_name(text):
    t = text.strip()
    return (NAME_RE.match(t) is not None
            and sum(ch.isalpha() for ch in t) >= 2
            and t.upper() not in NON_NAME_WORDS)


def extract_absentees(filepath):
    doc = Document(filepath)
    found       = {}          # reg no -> record (de-duplicates, keeps first-seen order)
    table_count = 0

    for table in doc.tables:
        table_count += 1
        for row in table.rows:
            texts = [c.text.strip() for c in row.cells]
            if len(texts) < 3:
                continue
            # A header row has no register-number-like cell, so it falls out
            # here; a table with no header row keeps its first data row.
            if not any(looks_like_reg_no(t) for t in texts):
                continue

            reg_no = name = status = ""
            for t in texts:
                if looks_like_reg_no(t):
                    if not reg_no:
                        reg_no = normalize_reg_no(t)
                elif t.upper() in ABSENT_MARKS:
                    status = "AB"
                elif looks_like_name(t):
                    if not name:
                        name = t

            if not (status and reg_no):
                continue
            rec = found.get(reg_no)
            if rec is None:
                found[reg_no] = {
                    "Reg No":     reg_no,
                    "Name":       name or "—",
                    "Department": get_department(reg_no),
                    "Year":       get_year(reg_no),
                }
            elif rec["Name"] == "—" and name:
                rec["Name"] = name

    return list(found.values()), table_count


@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if not file.filename.lower().endswith(".docx"):
        return jsonify({"error": "Please upload a .docx file"}), 400

    filepath = os.path.join(UPLOAD_FOLDER, "attendance.docx")
    file.save(filepath)

    try:
        absentees, table_count = extract_absentees(filepath)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if not absentees:
        return jsonify({"error": "No absentees found. Check AB column or file format."}), 404

    df = pd.DataFrame(absentees).sort_values(by=["Department", "Year", "Reg No"])
    df.to_excel(os.path.join(OUTPUT_FOLDER, "absentees.xlsx"), index=False)
    df.to_csv(os.path.join(OUTPUT_FOLDER,   "absentees.csv"),  index=False)

    summary = df.groupby("Department").size().reset_index(name="Count")
    return jsonify({
        "total":   len(absentees),
        "tables":  table_count,
        "summary": summary.to_dict(orient="records"),
        "records": df.to_dict(orient="records"),
    })

@app.route("/download/excel")
def download_excel():
    src = os.path.join(OUTPUT_FOLDER, "absentees.xlsx")
    if not os.path.exists(src):
        return jsonify({"error": "No data yet. Process a file first."}), 404
    # Save directly into the user's Downloads folder
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(downloads, exist_ok=True)
    dest = os.path.join(downloads, "absentees_dept_wise.xlsx")
    import shutil
    shutil.copy2(src, dest)
    return jsonify({"saved_to": dest})

@app.route("/download/pdf")
def download_pdf():
    csv_path = os.path.join(OUTPUT_FOLDER, "absentees.csv")
    if not os.path.exists(csv_path): return "No data yet", 404

    df       = pd.read_csv(csv_path)
    pdf_path = os.path.join(OUTPUT_FOLDER, "absentees.pdf")

    doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                            leftMargin=40, rightMargin=40,
                            topMargin=40, bottomMargin=40)
    styles     = getSampleStyleSheet()
    title_s    = ParagraphStyle("t", parent=styles["Heading1"], fontSize=16,
                                spaceAfter=12, textColor=colors.HexColor("#1a1a2e"))
    sub_s      = ParagraphStyle("s", parent=styles["Normal"],   fontSize=9,
                                textColor=colors.grey)
    dept_s     = ParagraphStyle("d", parent=styles["Heading2"], fontSize=11,
                                spaceBefore=10, spaceAfter=4,
                                textColor=colors.HexColor("#0f3460"))

    elements = [
        Paragraph("Exam Absentees Report", title_s),
        Paragraph(f"Total Absentees: {len(df)}  |  Departments: {df['Department'].nunique()}", sub_s),
        Spacer(1, 16),
    ]

    for dept in sorted(df["Department"].unique()):
        dept_df = df[df["Department"] == dept].reset_index(drop=True)
        elements.append(Paragraph(f"{dept}  ({len(dept_df)} absent)", dept_s))
        data = [["#", "Reg No", "Name", "Year"]]
        for i, row in dept_df.iterrows():
            data.append([str(i+1), row["Reg No"], row["Name"], row["Year"]])
        t = Table(data, colWidths=[25, 145, 200, 80])
        t.setStyle(TableStyle([
            ("BACKGROUND",     (0,0),(-1,0),  colors.HexColor("#0f3460")),
            ("TEXTCOLOR",      (0,0),(-1,0),  colors.white),
            ("FONTNAME",       (0,0),(-1,0),  "Helvetica-Bold"),
            ("FONTSIZE",       (0,0),(-1,-1), 8),
            ("ROWBACKGROUNDS", (0,1),(-1,-1), [colors.white, colors.HexColor("#f0f4ff")]),
            ("GRID",           (0,0),(-1,-1), 0.3, colors.HexColor("#ccccdd")),
            ("LEFTPADDING",    (0,0),(-1,-1), 6),
            ("RIGHTPADDING",   (0,0),(-1,-1), 6),
            ("TOPPADDING",     (0,0),(-1,-1), 4),
            ("BOTTOMPADDING",  (0,0),(-1,-1), 4),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 8))

    doc.build(elements)

    # Save directly into the user's Downloads folder
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(downloads, exist_ok=True)
    dest = os.path.join(downloads, "absentees_report.pdf")
    import shutil
    shutil.copy2(pdf_path, dest)
    return jsonify({"saved_to": dest})


if __name__ == "__main__":
    app.run(debug=True)
