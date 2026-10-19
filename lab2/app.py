"""Лабораторная работа №2: генератор экзаменационных билетов с графическим интерфейсом."""
import logging
import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk

import ticket_data as td

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STUDENTS_FILE = os.path.join(DATA_DIR, "students.xlsx")
TICKETS_FILE = os.path.join(DATA_DIR, "tickets.docx")
RESULTS_FILE = os.path.join(DATA_DIR, "results.xlsx")
LOG_FILE = os.path.join(BASE_DIR, "tickets.log")


class TicketApp(tk.Tk):
    def __init__(self, groups, tickets):
        super().__init__()
        self.groups = groups
        self.tickets = tickets
        self.title("Генератор экзаменационных билетов")
        self.geometry("420x220")
        self.resizable(False, False)

        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Группа:").grid(row=0, column=0, sticky="w", pady=6)
        self.group_var = tk.StringVar()
        self.group_box = ttk.Combobox(frame, textvariable=self.group_var, state="readonly",
                                      values=list(groups), width=32)
        self.group_box.grid(row=0, column=1, pady=6)
        self.group_box.bind("<<ComboboxSelected>>", self.on_group_selected)

        ttk.Label(frame, text="Студент:").grid(row=1, column=0, sticky="w", pady=6)
        self.student_var = tk.StringVar()
        self.student_box = ttk.Combobox(frame, textvariable=self.student_var, state="readonly",
                                        values=[], width=32)
        self.student_box.grid(row=1, column=1, pady=6)
        self.student_box.bind("<<ComboboxSelected>>", lambda _e: self.update_button())

        self.generate_btn = ttk.Button(frame, text="Сгенерировать билет",
                                       command=self.on_generate, state="disabled")
        self.generate_btn.grid(row=2, column=0, columnspan=2, pady=18, ipadx=10, ipady=4)

        ttk.Label(frame, text=f"Билетов загружено: {len(tickets)}",
                  foreground="gray").grid(row=3, column=0, columnspan=2)

    # --- выбор ---

    def current_students(self):
        return self.groups.get(self.group_var.get(), [])

    def on_group_selected(self, _event=None):
        students = self.current_students()
        self.student_box["values"] = [f"{last} {first}".strip() for last, first in students]
        self.student_var.set("")
        self.update_button()

    def selected_student(self):
        index = self.student_box.current()
        students = self.current_students()
        if 0 <= index < len(students):
            return students[index]
        return None

    def update_button(self):
        ready = bool(self.group_var.get()) and self.selected_student() is not None
        self.generate_btn["state"] = "normal" if ready else "disabled"

    # --- генерация ---

    def on_generate(self):
        student = self.selected_student()
        if student is None:
            return
        group = self.group_var.get()
        last, first = student
        number, repeat = td.assign_ticket(RESULTS_FILE, self.tickets, group, last, first)
        if not self.save_result(group, last, first, number, repeat):
            return
        self.show_result(number, repeat)

    def save_result(self, group, last, first, number, repeat):
        while True:
            try:
                td.append_result(RESULTS_FILE, group, last, first, number, repeat)
                return True
            except PermissionError:
                retry = messagebox.askretrycancel(
                    "Файл занят",
                    f"Файл {os.path.basename(RESULTS_FILE)} открыт в другой программе "
                    "(например, в Excel).\nЗакройте его и нажмите «Повторить».",
                    parent=self)
                if not retry:
                    return False

    def show_result(self, number, repeat):
        win = tk.Toplevel(self)
        win.title(f"Билет № {number}")
        win.transient(self)
        win.grab_set()
        body = ttk.Frame(win, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=f"Ваш билет № {number}",
                  font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 12))
        questions = self.tickets.get(number)
        if questions:
            for i, question in enumerate(questions, 1):
                ttk.Label(body, text=f"{i}. {question}", wraplength=460,
                          justify="left").pack(anchor="w", pady=3)
        else:
            ttk.Label(body, text="Вопросы этого билета не найдены в tickets.docx.",
                      foreground="red").pack(anchor="w")
        if repeat:
            ttk.Label(body, text="Повторный запрос: выдан ранее назначенный билет.",
                      foreground="gray").pack(anchor="w", pady=(10, 0))
        ttk.Label(body, text="Esc — закрыть", foreground="gray").pack(anchor="e", pady=(14, 0))
        win.bind("<Escape>", lambda _e: win.destroy())
        win.focus_force()


def main():
    logging.basicConfig(filename=LOG_FILE, level=logging.INFO, encoding="utf-8",
                        format="%(asctime)s %(levelname)s %(message)s")
    try:
        groups = td.load_groups(STUDENTS_FILE)
        tickets = td.load_tickets(TICKETS_FILE)
    except td.SourceError as exc:
        logging.error("%s", exc)
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Ошибка исходных данных", str(exc))
        root.destroy()
        return 1
    TicketApp(groups, tickets).mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
