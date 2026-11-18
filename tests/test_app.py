"""Tests for ExamSync. All data is synthetic: fake names, made-up register numbers."""
import datetime
import io
import os

import pytest
from docx import Document

import app as examsync

# Fixed "today" (August 2025 => academic year 2025-26) so year labels are stable.
TODAY = datetime.date(2025, 8, 1)

HEADER = ["S.No", "Register No", "Name", "Answer Book", "Attendance"]


def make_docx(path, tables):
    """tables: list of tables; each table is a list of rows (lists of strings)."""
    doc = Document()
    for rows in tables:
        t = doc.add_table(rows=len(rows), cols=len(rows[0]))
        for r, row in enumerate(rows):
            for c, text in enumerate(row):
                t.cell(r, c).text = text
        doc.add_paragraph("")
    doc.save(str(path))
    return str(path)


# ── parse_reg_no ─────────────────────────────────────────────────────────────

@pytest.mark.parametrize("reg, expected", [
    ("610822104001", ("CSE", "4th Year")),    # 22: 3-digit numeric code
    ("610823U106002", ("ECE", "3rd Year")),   # 23: U + 3 digits
    ("610824BME003", ("MECH", "2nd Year")),   # 24: 3-letter code
    ("610825BCS004", ("CSE", "1st Year")),    # 25: 3-letter code
    ("610825bad005", ("AIDS", "1st Year")),   # lower case is normalised
    ("  610825BIT006  ", ("IT", "1st Year")), # surrounding whitespace
])
def test_parse_reg_no_known(reg, expected):
    assert examsync.parse_reg_no(reg, today=TODAY) == expected


def test_parse_reg_no_unknown_dept_code():
    dept, year = examsync.parse_reg_no("610825BXX001", today=TODAY)
    assert dept == "UNKNOWN(BXX)"
    assert year == "1st Year"


def test_year_rolls_over_in_june():
    reg = "610825BCS004"
    assert examsync.parse_reg_no(reg, today=datetime.date(2026, 5, 31))[1] == "1st Year"
    assert examsync.parse_reg_no(reg, today=datetime.date(2026, 6, 1))[1] == "2nd Year"


def test_year_fallback_outside_one_to_four():
    # Older batch (arrears) and a batch from the future: no invented year.
    assert examsync.parse_reg_no("610821BCS001", today=TODAY)[1] == "Batch 2021"
    assert examsync.parse_reg_no("610827BCS001", today=TODAY)[1] == "Batch 2027"


def test_parse_reg_no_garbage_does_not_raise():
    dept, year = examsync.parse_reg_no("", today=TODAY)
    assert dept.startswith("UNKNOWN")


# ── extract_absentees ────────────────────────────────────────────────────────

def test_extracts_only_absentees_with_dept_and_year(tmp_path):
    path = make_docx(tmp_path / "a.docx", [[
        HEADER,
        ["1", "610825BCS001", "Asha Rao",   "1001", ""],
        ["2", "610825BCS002", "Bala Menon", "1002", "AB"],
        ["3", "610824BME003", "Chitra Iyer", "1003", "AB"],
    ]])
    rows, table_count = examsync.extract_absentees(path)
    assert table_count == 1
    assert [r["Reg No"] for r in rows] == ["610825BCS002", "610824BME003"]
    assert rows[0]["Name"] == "Bala Menon"
    assert rows[1]["Department"] == "MECH"


def test_table_without_header_row_keeps_first_row(tmp_path):
    path = make_docx(tmp_path / "a.docx", [[
        ["1", "610825BCS001", "Asha Rao", "1001", "AB"],   # row 0 is data
        ["2", "610825BCS002", "Bala Menon", "1002", ""],
        ["3", "610825BCS003", "Chitra Iyer", "1003", "AB"],
    ]])
    rows, _ = examsync.extract_absentees(path)
    assert [r["Reg No"] for r in rows] == ["610825BCS001", "610825BCS003"]


def test_header_row_is_not_treated_as_a_student(tmp_path):
    path = make_docx(tmp_path / "a.docx", [[
        ["S.No", "Register Number", "Student Name", "Answer Book No", "ABSENT"],
        ["1", "610825BCS001", "Asha Rao", "1001", "AB"],
    ]])
    rows, _ = examsync.extract_absentees(path)
    assert len(rows) == 1 and rows[0]["Reg No"] == "610825BCS001"


@pytest.mark.parametrize("name", [
    "R. Kumar", "S. KUMAR", "R.SANGAMI", "S.K. Devi", "A. B. C. Nair", "Mary-Anne Joseph", "K S Priya", "D'Souza Rita",
])
def test_names_with_initials_dots_and_hyphens(tmp_path, name):
    path = make_docx(tmp_path / "a.docx", [[
        HEADER,
        ["1", "610825BCS001", name, "1001", "AB"],
    ]])
    rows, _ = examsync.extract_absentees(path)
    assert rows[0]["Name"] == name


def test_missing_name_shows_dash(tmp_path):
    path = make_docx(tmp_path / "a.docx", [[
        HEADER,
        ["1", "610825BCS001", "", "1001", "AB"],
    ]])
    rows, _ = examsync.extract_absentees(path)
    assert rows[0]["Name"] == "—"


def test_duplicates_are_removed_across_tables(tmp_path):
    path = make_docx(tmp_path / "a.docx", [
        [HEADER, ["1", "610825BCS001", "Asha Rao", "1001", "AB"]],
        [HEADER, ["1", "610825BCS001", "Asha Rao", "1001", "AB"],
                 ["2", "610825bcs001 ", "Asha Rao", "1001", "AB"],   # case/space variant
                 ["3", "610825BCS002", "Bala Menon", "1002", "AB"]],
    ])
    rows, table_count = examsync.extract_absentees(path)
    assert table_count == 2
    assert [r["Reg No"] for r in rows] == ["610825BCS001", "610825BCS002"]


def test_all_digit_register_numbers_are_found(tmp_path):
    path = make_docx(tmp_path / "a.docx", [[
        HEADER,
        ["1", "610822104001", "Asha Rao", "1001", "AB"],
    ]])
    rows, _ = examsync.extract_absentees(path)
    assert rows[0]["Department"] == "CSE"


def test_answer_book_number_is_not_a_register_number(tmp_path):
    path = make_docx(tmp_path / "a.docx", [[
        HEADER,
        ["1", "610825BCS001", "Asha Rao", "1234567890", "AB"],
    ]])
    rows, _ = examsync.extract_absentees(path)
    assert rows[0]["Reg No"] == "610825BCS001"


def test_unknown_dept_codes_are_extracted_not_dropped(tmp_path):
    path = make_docx(tmp_path / "a.docx", [[
        HEADER,
        ["1", "610825BXX001", "Asha Rao", "1001", "AB"],
    ]])
    rows, _ = examsync.extract_absentees(path)
    assert rows[0]["Department"] == "UNKNOWN(BXX)"


# ── /upload ──────────────────────────────────────────────────────────────────

@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(examsync, "UPLOAD_FOLDER", str(tmp_path / "uploads"))
    monkeypatch.setattr(examsync, "OUTPUT_FOLDER", str(tmp_path / "outputs"))
    return examsync.app.test_client()


def _upload(client, tmp_path, tables, filename="Hall1.DOCX"):
    path = make_docx(tmp_path / "src.docx", tables)
    with open(path, "rb") as fh:
        return client.post("/upload", data={"file": (io.BytesIO(fh.read()), filename)},
                           content_type="multipart/form-data")


def test_upload_reports_unknown_and_deletes_upload(client, tmp_path):
    resp = _upload(client, tmp_path, [[
        HEADER,
        ["1", "610825BCS001", "Asha Rao", "1001", "AB"],
        ["2", "610825BXX002", "Bala Menon", "1002", "AB"],
        ["3", "610825BZZ003", "Chitra Iyer", "1003", "AB"],
    ]])
    assert resp.status_code == 200          # upper-case .DOCX accepted
    body = resp.get_json()
    assert body["total"] == 3
    assert body["unknown_count"] == 2
    assert sorted(body["unknown_reg_nos"]) == ["610825BXX002", "610825BZZ003"]
    assert os.listdir(examsync.UPLOAD_FOLDER) == []            # upload removed
    assert os.path.exists(os.path.join(examsync.OUTPUT_FOLDER, "absentees.xlsx"))


def test_upload_with_no_unknowns(client, tmp_path):
    body = _upload(client, tmp_path, [[
        HEADER, ["1", "610825BCS001", "Asha Rao", "1001", "AB"],
    ]]).get_json()
    assert body["unknown_count"] == 0 and body["unknown_reg_nos"] == []


def test_upload_rejects_non_docx(client):
    resp = client.post("/upload", data={"file": (io.BytesIO(b"x"), "notes.txt")},
                       content_type="multipart/form-data")
    assert resp.status_code == 400


def test_upload_deletes_file_even_when_parsing_fails(client):
    resp = client.post("/upload", data={"file": (io.BytesIO(b"not a docx"), "bad.docx")},
                       content_type="multipart/form-data")
    assert resp.status_code == 500
    assert os.listdir(examsync.UPLOAD_FOLDER) == []


# ── data directory ───────────────────────────────────────────────────────────

def test_data_dir_frozen_uses_localappdata(monkeypatch, tmp_path):
    monkeypatch.setattr(examsync.sys, "frozen", True, raising=False)
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    assert examsync.data_dir() == os.path.join(str(tmp_path), "ExamSync")


def test_data_dir_from_source_is_app_folder(monkeypatch):
    monkeypatch.delattr(examsync.sys, "frozen", raising=False)
    assert examsync.data_dir() == examsync.BASE_DIR
