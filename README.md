# ExamSync — Attendance Processor

Turns a hall-wise exam attendance Word file (`.docx`) into a department-wise
list of absentees, viewable on screen and exportable to Excel and PDF.

Stack: Flask + python-docx + pandas/openpyxl + reportlab, wrapped in a native
window with PyWebView, packaged to a single `.exe` with PyInstaller.

---

## Register number format

Department and year are read from the register number. Example: `610823U104001`

| Part            | Position (0-based) | Example  | Meaning                         |
|-----------------|--------------------|----------|---------------------------------|
| College code    | `[0:4]`            | `6108`   | Not used                        |
| Admission year  | `[4:6]`            | `23`     | Joined in June 2023             |
| Department code | starts at `[6]`    | `U104`   | Looked up in `DEPT_MAP`         |
| Serial          | the rest           | `001`    | Not used                        |

The length of the department code depends on the admission year
(`YEAR_SLICE` in `app.py`):

| Admission year | Length | Format          | Example         |
|----------------|--------|-----------------|-----------------|
| 22             | 3      | digits          | `610822104001`  |
| 23             | 4      | `U` + 3 digits  | `610823U104001` |
| 24, 25 (and any other) | 3 | 3 letters  | `610825BCS001`  |

### Department mappings (`DEPT_MAP`)

| Code (22 / 23)   | Code (24+) | Department |
|------------------|------------|------------|
| `101` / `U101`   | `BAE`      | AERO       |
| `103` / `U103`   | `BCE`      | CIVIL      |
| `104` / `U104`   | `BCS`      | CSE        |
| `105` / `U105`   | `BEE`      | EEE        |
| `106` / `U106`   | `BEC`      | ECE        |
| `114` / `U114`   | `BME`      | MECH       |
| `115` / `U115`   | `BMC`      | MCO        |
| `203` / `U203`   | `BCH`      | CHEM       |
| `205` / `U205`   | `BIT`      | IT         |
| `243` / `U243`   | `BAD`      | AIDS       |
| `244` / `U244`   | `BCB`      | CSBS       |
| —                | `BML`      | AIML       |

PG programmes: `MBA`, `MCA`, `MCS`, `MEE` map to themselves.

A code that is not in `DEPT_MAP` is reported as `UNKNOWN(<code>)`. The upload
response includes `unknown_count` and `unknown_reg_nos`, and the results page
shows an amber warning listing them. Add the code to `DEPT_MAP` to fix it.

### Year of study

Computed, not looked up. The academic year starts in **June**. With `Y` the
calendar year in which the current academic year began, and `yy` the admission
year code:

    year of study = Y - (2000 + yy) + 1

For example, in August 2025 (`Y = 2025`) code `25` is 1st Year and `22` is 4th
Year; from 1 June 2026 code `25` becomes 2nd Year and code `22` (now past year
4) is labelled `Batch 2022`.

**Assumptions:** students join in June of the admission year and progress one
year per academic year, with no gap years. If the result is not 1–4 (arrears
from older batches, or a code from the future) the label is `Batch 20yy`
instead of a guessed year. PG programmes use the same 1–4 scale, so a finished
2-year PG batch shows as 3rd/4th Year until it passes 4.

---

## Setup (run from source)

Python 3.9+ on Windows.

```bash
pip install -r requirements.txt
python desktop_app.py     # native window (needs a WebView2 runtime, present on Windows 10/11)
# or, in a browser:
python app.py             # then open http://127.0.0.1:5000
```

## How to use

1. Upload the attendance `.docx` (extension check is case-insensitive).
2. Click **Process Attendance File**.
3. Review absentees by department; check any amber warning about unknown codes.
4. **Download Excel** / **Download PDF** — files are copied to your
   `Downloads` folder (`absentees_dept_wise.xlsx`, `absentees_report.pdf`).

## What the parser expects in the Word file

It scans every table in the document, row by row:

- A row counts only if some cell looks like a register number (6 digits, then
  4–10 letters/digits; all-digit numbers must be 12+ digits so 10-digit answer
  book numbers are ignored). Rows without one — headers, blank rows — are
  skipped, so a table with no header row works.
- A cell equal to `AB`, `A`, `ABSENT`, `✖` or `X` (any case) marks the student
  absent.
- The name is the first remaining text cell made of letters, spaces, dots,
  hyphens and apostrophes (`R. Kumar`, `Mary-Anne`, `S.K. Devi`). If none is
  found the name is shown as `—`.
- Duplicate register numbers (across or within tables) are reported once.
- Rows need at least 3 cells.

Columns are detected by cell *content*, not position.

## Where data is stored

| Run mode      | Uploads / outputs                   |
|---------------|-------------------------------------|
| From source   | `uploads/`, `outputs/` next to `app.py` |
| Built `.exe`  | `%LOCALAPPDATA%\ExamSync\uploads` and `...\outputs` |

The uploaded `.docx` is deleted as soon as it has been parsed. The extracted
`absentees.xlsx`, `.csv` and `.pdf` in `outputs/` contain student names and
register numbers and are overwritten on the next upload; delete them yourself
if the machine is shared. Both folders are git-ignored.

## Build the .exe

```bat
BUILD.bat
```

It installs the dependencies and runs `pyinstaller ExamSync.spec --clean
--noconfirm`. The result is `dist\ExamSync.exe`, a windowed (no console)
single file that runs without Python and fully offline. `build/` and `dist/`
are git-ignored.

## Tests

```bash
pip install pytest
pytest
```

Tests build small synthetic `.docx` files (fake names) in a temp folder. They
cover register-number parsing, extraction (with/without header row, initials,
duplicates, unknown codes), the `/upload` response and the data folder choice.

## Limitations

- **Depends on the Word layout.** Absentees must be in real Word tables with an
  absent marker in a cell. Text in paragraphs, text boxes, images or scanned
  PDFs saved as Word is not read. Merged or nested tables may be misread.
- **Register-number scheme is hard-coded.** A new college code, department
  code or format change needs an edit to `DEPT_MAP` / `YEAR_SLICE` in `app.py`.
  Lateral-entry students and other schemes are not handled.
- **No ERP integration.** Output is a file for manual entry.
- **No exam date in reports.** The Excel/CSV/PDF contain register number, name,
  department and year only.
- Year labels depend on the computer's clock, so reprocessing an old file after
  June shifts every batch up a year.
- The Excel export is a single sheet; the PDF is built from the last processed
  file only.
- The `.exe` is unsigned, so Windows SmartScreen may warn on first run.
