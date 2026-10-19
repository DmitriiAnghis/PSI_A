"""Работа с файлами Лаб 2: students.xlsx, tickets.docx, results.xlsx."""
import logging
import os
import random
import re
from datetime import datetime

from docx import Document
from openpyxl import Workbook, load_workbook

log = logging.getLogger("tickets")

RESULTS_HEADER = ["Группа", "Фамилия", "Имя", "Номер билета", "Дата и время", "Повтор (да/нет)"]
QUESTIONS_PER_TICKET = 3

_TICKET_RE = re.compile(r"^\s*Билет\s*№?\s*(\d+)\s*$", re.IGNORECASE)
_QUESTION_RE = re.compile(r"^\s*\d+\s*[.)]\s*(.+?)\s*$")


class SourceError(Exception):
    """Входной файл отсутствует или не читается."""


def _cell_text(value):
    return "" if value is None else str(value).strip()


def load_groups(path):
    """Возвращает {группа: [(фамилия, имя), ...]} в порядке листов students.xlsx."""
    if not os.path.exists(path):
        raise SourceError(f"Не найден файл со студентами: {path}")
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise SourceError(f"Не удалось прочитать {path}: {exc}") from exc
    groups = {}
    for ws in wb.worksheets:
        students = []
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            last, first = (_cell_text(v) for v in (tuple(row) + (None, None))[:2])
            if last or first:
                students.append((last, first))
        groups[ws.title] = students
    wb.close()
    return groups


def _is_list_paragraph(paragraph):
    """Абзац оформлен автоматической нумерацией Word (номер не входит в текст)."""
    ppr = paragraph._p.pPr
    return ppr is not None and ppr.numPr is not None


def parse_ticket_lines(lines):
    """Разбирает строки документа в {номер: [вопросы]}.

    lines: список (текст, это_пункт_нумерованного_списка).
    Билеты, у которых меньше трёх вопросов, пропускаются с записью в лог.
    """
    tickets = {}
    current, questions = None, []

    def flush():
        if current is None:
            return
        if len(questions) < QUESTIONS_PER_TICKET:
            log.warning("Билет %s пропущен: найдено вопросов %d из %d",
                        current, len(questions), QUESTIONS_PER_TICKET)
        elif current in tickets:
            log.warning("Билет %s пропущен: номер повторяется", current)
        else:
            tickets[current] = questions[:QUESTIONS_PER_TICKET]

    for text, is_list_item in lines:
        text = text.strip()
        if not text:
            continue
        header = _TICKET_RE.match(text)
        if header:
            flush()
            current, questions = int(header.group(1)), []
            continue
        question = _QUESTION_RE.match(text)
        if current is None:
            log.warning("Строка вне билета пропущена: %r", text)
        elif question:
            questions.append(question.group(1))
        elif is_list_item:
            questions.append(text)
        else:
            log.warning("Нераспознанная строка в билете %s: %r", current, text)
    flush()
    return tickets


def load_tickets(path):
    """Читает tickets.docx и возвращает {номер билета: [3 вопроса]}."""
    if not os.path.exists(path):
        raise SourceError(f"Не найден файл с билетами: {path}")
    try:
        doc = Document(path)
    except Exception as exc:
        raise SourceError(f"Не удалось прочитать {path}: {exc}") from exc
    lines = [(p.text, _is_list_paragraph(p)) for p in doc.paragraphs]
    tickets = parse_ticket_lines(lines)
    if not tickets:
        raise SourceError(f"В {path} не найдено ни одного корректного билета")
    return tickets


def _key(group, last, first):
    return (group.strip().casefold(), last.strip().casefold(), first.strip().casefold())


def find_first_ticket(results_path, group, last, first):
    """Номер билета из самой первой записи студента в журнале или None."""
    if not os.path.exists(results_path):
        return None
    wb = load_workbook(results_path, read_only=True)
    try:
        target = _key(group, last, first)
        for row in wb.active.iter_rows(min_row=2, max_col=4, values_only=True):
            g, l, f, number = (tuple(row) + (None,) * 4)[:4]
            if number is not None and _key(_cell_text(g), _cell_text(l), _cell_text(f)) == target:
                return int(number)
    finally:
        wb.close()
    return None


def append_result(results_path, group, last, first, ticket, repeat, when=None):
    """Дописывает строку в конец results.xlsx и сразу сохраняет.

    PermissionError пробрасывается наверх (файл открыт в Excel).
    """
    when = when or datetime.now()
    if os.path.exists(results_path):
        wb = load_workbook(results_path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Results"
        ws.append(RESULTS_HEADER)
    ws.append([group, last, first, ticket, when.strftime("%d.%m.%Y %H:%M:%S"),
               "да" if repeat else "нет"])
    wb.save(results_path)


def assign_ticket(results_path, tickets, group, last, first, rng=random):
    """Возвращает (номер, повтор): первый номер студента или случайный из tickets."""
    previous = find_first_ticket(results_path, group, last, first)
    if previous is not None:
        return previous, True
    return rng.choice(sorted(tickets)), False
