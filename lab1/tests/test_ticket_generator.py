import random
import sys
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import ticket_generator as tg  # noqa: E402


def rows(path):
    return [list(r) for r in load_workbook(path).active.iter_rows(values_only=True)]


def test_clean_name_trims_spaces():
    assert tg.clean_name("  Anghis  ") == "Anghis"


@pytest.mark.parametrize("raw", ["", "   ", "\t"])
def test_clean_name_rejects_empty(raw):
    assert tg.clean_name(raw) is None


def test_ticket_in_range():
    rng = random.Random(1)
    assert all(1 <= tg.generate_ticket(rng) <= 20 for _ in range(1000))


def test_creates_file_with_header(tmp_path):
    path = tmp_path / "journal.xlsx"
    tg.append_record(path, "Anghis", "Dmitrii", 7, datetime(2026, 10, 19, 12, 0, 0))
    assert rows(path) == [tg.HEADER, ["Anghis", "Dmitrii", 7, "19.10.2026 12:00:00"]]


def test_appends_without_overwriting(tmp_path):
    path = tmp_path / "journal.xlsx"
    for i in range(3):
        tg.append_record(path, f"L{i}", f"F{i}", i + 1)
    before = rows(path)
    tg.append_record(path, "L3", "F3", 4)
    tg.append_record(path, "L4", "F4", 5)
    after = rows(path)
    assert len(after) == 6
    assert after[:4] == before


def test_locked_file_is_retried(tmp_path, monkeypatch, capsys):
    calls = {"n": 0}

    def flaky(*args, **kwargs):
        calls["n"] += 1
        if calls["n"] == 1:
            raise PermissionError
    monkeypatch.setattr(tg, "append_record", flaky)
    monkeypatch.setattr("builtins.input", lambda *_: "")
    tg.save_with_retry("A", "B", 1, path=str(tmp_path / "j.xlsx"))
    assert calls["n"] == 2
    assert "недоступен для записи" in capsys.readouterr().out
