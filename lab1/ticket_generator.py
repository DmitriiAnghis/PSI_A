"""Лабораторная работа №1: консольный генератор экзаменационных билетов.

Запрашивает фамилию и имя студента, выдаёт случайный номер билета 1..20
и сразу дописывает запись в journal.xlsx.
"""
import os
import random
import sys
from datetime import datetime

from openpyxl import Workbook, load_workbook

JOURNAL_FILE = "journal.xlsx"
HEADER = ["Last name", "First name", "Номер билета", "Дата и время"]
COLUMN_WIDTHS = {"A": 20, "B": 20, "C": 15, "D": 22}
TICKET_MIN, TICKET_MAX = 1, 20
ESC = "\x1b"


class ExitRequested(Exception):
    """Пользователь нажал ESC."""


# ---------- логика (покрыта тестами) ----------

def clean_name(raw):
    """Обрезает пробелы по краям. Возвращает None, если строка пустая."""
    value = raw.strip()
    return value or None


def generate_ticket(rng=random):
    return rng.randint(TICKET_MIN, TICKET_MAX)


def append_record(path, last_name, first_name, ticket, when=None):
    """Дописывает одну строку в конец журнала и сразу сохраняет файл.

    Если файла нет, создаёт его с шапкой. Существующие строки не меняются.
    PermissionError пробрасывается наверх (файл открыт в Excel).
    """
    when = when or datetime.now()
    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Journal"
        ws.append(HEADER)
        for column, width in COLUMN_WIDTHS.items():
            ws.column_dimensions[column].width = width
    ws.append([last_name, first_name, ticket, when.strftime("%d.%m.%Y %H:%M:%S")])
    wb.save(path)


# ---------- консольный ввод ----------

def _read_key():
    """Считывает одну клавишу без ожидания Enter."""
    if os.name == "nt":
        import msvcrt
        ch = msvcrt.getwch()
        if ch in ("\x00", "\xe0"):  # стрелки, F-клавиши: пропускаем второй код
            msvcrt.getwch()
            return ""
        return ch
    if not sys.stdin.isatty():
        ch = sys.stdin.read(1)
        if ch == "":
            raise EOFError
        return ch
    import termios
    import tty
    fd = sys.stdin.fileno()
    old = termios.tcgetattr(fd)
    try:
        tty.setraw(fd)
        return sys.stdin.read(1)
    finally:
        termios.tcsetattr(fd, termios.TCSADRAIN, old)


def ask(prompt):
    """Запрашивает непустую строку. ESC в начале ввода -> ExitRequested."""
    while True:
        print(prompt, end="", flush=True)
        first = _read_key()
        if first == ESC:
            print()
            raise ExitRequested
        if first in ("\r", "\n", "", "\x08", "\x7f"):  # Enter/Backspace на пустом поле
            print()
            print("Поле не может быть пустым, повторите ввод.")
            continue
        if first == "\x03":  # Ctrl+C
            raise KeyboardInterrupt
        print(first, end="", flush=True)
        value = clean_name(first + input())
        if value:
            return value
        print("Поле не может быть пустым, повторите ввод.")


def save_with_retry(last_name, first_name, ticket, path=JOURNAL_FILE):
    while True:
        try:
            append_record(path, last_name, first_name, ticket)
            return
        except PermissionError:
            print(f"Файл {path} открыт в другой программе (например, в Excel) "
                  "и недоступен для записи.")
            input("Закройте файл и нажмите Enter, чтобы повторить сохранение...")


def main():
    print("Для выхода нажмите ESC")
    try:
        while True:
            last_name = ask("Last name: ")
            first_name = ask("First name: ")
            ticket = generate_ticket()
            print(f"Билет № {ticket}")
            save_with_retry(last_name, first_name, ticket)
            print()
    except (ExitRequested, KeyboardInterrupt, EOFError):
        print("Работа завершена.")


if __name__ == "__main__":
    main()
