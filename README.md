# ExamSync — Attendance Processor

Convert hall-wise exam attendance Word files into department-wise absentee lists instantly.

---

## Setup (One-time)

```bash
# 1. Install Python (3.9+) if not already installed

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the app
python app.py
```

Then open your browser: **http://127.0.0.1:5000**

---

## How to Use

1. Open `http://127.0.0.1:5000` in your browser
2. Upload the attendance `.docx` file
3. Click **Process Attendance File**
4. View department-wise absentees on screen
5. Download **Excel** or **PDF** for ERP entry

---

## Configuring for Your College

Open `app.py` and edit the top section:

```python
DEPT_MAP = {
    "104": "CSE",
    "105": "ECE",
    # Add your department codes here
}

YEAR_MAP = {
    "23": "2nd Year",
    "22": "3rd Year",
    # Adjust based on admission year
}

# Adjust index positions based on your register number format
# Example: 610823U104001
DEPT_START = 8   # dept code starts at position 8
DEPT_LEN   = 3   # dept code is 3 characters long
YEAR_START = 4   # year starts at position 4
YEAR_LEN   = 2   # year is 2 characters long
```

---

## Expected Word File Format

The tool looks for tables in your `.docx` file with columns like:

| S.No | Register No    | Name           | Answer Book | AB |
|------|----------------|----------------|-------------|-----|
| 1    | 610823U104001  | STUDENT NAME   | 12345       | AB  |

- **AB column**: Cells containing `AB`, `A`, `Absent`, or `✖` are treated as absent
- The tool auto-detects column positions

---

## Output

- **Excel file**: One sheet, all absentees sorted by department
- **PDF report**: Department-wise tables with count, formatted for printing

---

## Troubleshooting

**"No absentees found"** — Run this debug snippet to inspect your table structure:
```python
from docx import Document
doc = Document("attendance.docx")
for table in doc.tables:
    for row in table.rows[:3]:
        print([c.text for c in row.cells])
```
Check which column index has your AB marks and update `app.py` accordingly.
