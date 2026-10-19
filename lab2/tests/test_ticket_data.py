import random
import sys
from pathlib import Path

import pytest
from docx import Document
from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import ticket_data as td  # noqa: E402


@pytest.fixture
def students(tmp_path):
    wb = Workbook()
    wb.remove(wb.active)
    a = wb.create_sheet("TI-235")
    a.append(["Фамилия", "Имя"])
    a.append(["Anghis", "Dmitrii"])
    a.append([" Popescu ", "Ion"])
    b = wb.create_sheet("TI-237")
    b.append(["Фамилия", "Имя"])
    path = tmp_path / "students.xlsx"
    wb.save(path)
    return path


def make_docx(path, paragraphs):
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(path)
    return path


def rows(path):
    return [list(r) for r in load_workbook(path).active.iter_rows(values_only=True)]


# --- студенты по группе ---

def test_groups_are_sheets_and_students_by_group(students):
    groups = td.load_groups(students)
    assert list(groups) == ["TI-235", "TI-237"]
    assert groups["TI-235"] == [("Anghis", "Dmitrii"), ("Popescu", "Ion")]


def test_empty_group_has_no_students(students):
    assert td.load_groups(students)["TI-237"] == []


def test_missing_students_file(tmp_path):
    with pytest.raises(td.SourceError):
        td.load_groups(tmp_path / "nope.xlsx")


# --- парсинг билетов из Word ---

def test_parse_ticket_from_word(tmp_path):
    path = make_docx(tmp_path / "t.docx", [
        "Билет 1", "1. Вопрос А", "2. Вопрос Б", "3. Вопрос В", "",
        "Билет 2", "1. Q1", "2. Q2", "3. Q3",
    ])
    assert td.load_tickets(path) == {1: ["Вопрос А", "Вопрос Б", "Вопрос В"], 2: ["Q1", "Q2", "Q3"]}


def test_ticket_with_less_than_three_questions_is_skipped(tmp_path, caplog):
    path = make_docx(tmp_path / "t.docx", [
        "Билет 1", "1. A", "2. B",
        "Билет 2", "1. A", "2. B", "3. C",
    ])
    assert list(td.load_tickets(path)) == [2]
    assert "Билет 1 пропущен" in caplog.text


def test_unrecognized_header_is_ignored():
    lines = [("Bilet X", False), ("1. A", False), ("Билет 3", False),
             ("1. A", False), ("2. B", False), ("3. C", False)]
    assert td.parse_ticket_lines(lines) == {3: ["A", "B", "C"]}


def test_word_auto_numbered_list_is_parsed():
    lines = [("Билет 1", False), ("A", True), ("B", True), ("C", True)]
    assert td.parse_ticket_lines(lines) == {1: ["A", "B", "C"]}


def test_missing_tickets_file(tmp_path):
    with pytest.raises(td.SourceError):
        td.load_tickets(tmp_path / "nope.docx")


# --- журнал и идемпотентность ---

def test_new_student_gets_random_ticket_and_no_repeat(tmp_path):
    results = tmp_path / "results.xlsx"
    tickets = {n: ["a", "b", "c"] for n in range(1, 11)}
    number, repeat = td.assign_ticket(results, tickets, "TI-235", "Anghis", "Dmitrii", random.Random(0))
    assert number in tickets and repeat is False


def test_repeat_selection_returns_same_ticket(tmp_path):
    results = tmp_path / "results.xlsx"
    tickets = {n: ["a", "b", "c"] for n in range(1, 21)}
    first, _ = td.assign_ticket(results, tickets, "TI-235", "Anghis", "Dmitrii")
    td.append_result(results, "TI-235", "Anghis", "Dmitrii", first, False)
    for _ in range(5):
        number, repeat = td.assign_ticket(results, tickets, "TI-235", "Anghis", "Dmitrii")
        assert (number, repeat) == (first, True)
        td.append_result(results, "TI-235", "Anghis", "Dmitrii", number, repeat)


def test_same_name_in_other_group_is_different_student(tmp_path):
    results = tmp_path / "results.xlsx"
    td.append_result(results, "TI-235", "Anghis", "Dmitrii", 7, False)
    assert td.find_first_ticket(results, "TI-236", "Anghis", "Dmitrii") is None


def test_journal_appends_and_keeps_history(tmp_path):
    results = tmp_path / "results.xlsx"
    td.append_result(results, "TI-235", "Anghis", "Dmitrii", 7, False)
    td.append_result(results, "TI-235", "Popescu", "Ion", 3, False)
    before = rows(results)
    td.append_result(results, "TI-235", "Anghis", "Dmitrii", 7, True)
    after = rows(results)
    assert after[0] == td.RESULTS_HEADER
    assert after[:3] == before
    assert after[3][:4] == ["TI-235", "Anghis", "Dmitrii", 7] and after[3][5] == "да"
    assert [r[5] for r in after[1:]] == ["нет", "нет", "да"]
