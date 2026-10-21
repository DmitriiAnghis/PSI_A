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


FONT = "Segoe UI"
COLORS = {
    "bg": "#F4F6FB",
    "card": "#FFFFFF",
    "border": "#E3E7F0",
    "header": "#1E2A4A",
    "header_text": "#FFFFFF",
    "header_muted": "#AEB8D6",
    "text": "#1F2430",
    "muted": "#6B7280",
    "accent": "#3B5BDB",
    "accent_hover": "#2F4BC0",
    "accent_disabled": "#B9C4EE",
    "badge_bg": "#EDF1FF",
    "warn_bg": "#FFF4E5",
    "warn_text": "#9A5B00",
    "error": "#C92A2A",
}


def setup_styles(root):
    style = ttk.Style(root)
    style.theme_use("clam")
    c = COLORS
    style.configure(".", font=(FONT, 10), background=c["card"], foreground=c["text"])
    style.configure("Card.TFrame", background=c["card"])
    style.configure("Field.TLabel", background=c["card"], foreground=c["muted"],
                    font=(FONT, 9, "bold"))
    style.configure("Hint.TLabel", background=c["card"], foreground=c["muted"], font=(FONT, 9))
    style.configure("TCombobox", padding=6, fieldbackground=c["card"],
                    bordercolor=c["border"], lightcolor=c["border"], darkcolor=c["border"],
                    arrowcolor=c["accent"])
    style.map("TCombobox", bordercolor=[("focus", c["accent"])],
              fieldbackground=[("readonly", c["card"])],
              selectbackground=[("readonly", c["card"])],
              selectforeground=[("readonly", c["text"])])
    style.configure("Accent.TButton", font=(FONT, 11, "bold"), padding=(18, 10),
                    background=c["accent"], foreground="#FFFFFF", borderwidth=0,
                    focuscolor=c["accent"])
    style.map("Accent.TButton",
              background=[("disabled", c["accent_disabled"]), ("active", c["accent_hover"])],
              foreground=[("disabled", "#FFFFFF")])
    root.option_add("*TCombobox*Listbox.font", (FONT, 10))
    root.option_add("*TCombobox*Listbox.selectBackground", c["accent"])


def header(parent, title, subtitle):
    c = COLORS
    bar = tk.Frame(parent, bg=c["header"], padx=24, pady=16)
    bar.pack(fill="x")
    tk.Label(bar, text=title, bg=c["header"], fg=c["header_text"],
             font=(FONT, 15, "bold")).pack(anchor="w")
    tk.Label(bar, text=subtitle, bg=c["header"], fg=c["header_muted"],
             font=(FONT, 9)).pack(anchor="w", pady=(2, 0))
    return bar


def card(parent):
    c = COLORS
    outer = tk.Frame(parent, bg=c["border"], padx=1, pady=1)
    outer.pack(fill="both", expand=True, padx=20, pady=20)
    inner = ttk.Frame(outer, style="Card.TFrame", padding=20)
    inner.pack(fill="both", expand=True)
    return inner


class TicketApp(tk.Tk):
    def __init__(self, groups, tickets):
        super().__init__()
        self.groups = groups
        self.tickets = tickets
        self.title("Генератор экзаменационных билетов")
        self.configure(bg=COLORS["bg"])
        self.resizable(False, False)
        setup_styles(self)

        header(self, "Генератор экзаменационных билетов",
               "Выберите группу и студента, чтобы выдать билет")
        frame = card(self)
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="ГРУППА", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        self.group_var = tk.StringVar()
        self.group_box = ttk.Combobox(frame, textvariable=self.group_var, state="readonly",
                                      values=list(groups), width=36, font=(FONT, 10))
        self.group_box.grid(row=1, column=0, sticky="ew", pady=(4, 14))
        self.group_box.bind("<<ComboboxSelected>>", self.on_group_selected)

        ttk.Label(frame, text="СТУДЕНТ", style="Field.TLabel").grid(row=2, column=0, sticky="w")
        self.student_var = tk.StringVar()
        self.student_box = ttk.Combobox(frame, textvariable=self.student_var, state="readonly",
                                        values=[], width=36, font=(FONT, 10))
        self.student_box.grid(row=3, column=0, sticky="ew", pady=(4, 18))
        self.student_box.bind("<<ComboboxSelected>>", lambda _e: self.update_button())

        self.generate_btn = ttk.Button(frame, text="Сгенерировать билет", style="Accent.TButton",
                                       command=self.on_generate, state="disabled",
                                       cursor="hand2")
        self.generate_btn.grid(row=4, column=0, sticky="ew")

        self.status_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.status_var, style="Hint.TLabel").grid(
            row=5, column=0, sticky="w", pady=(12, 0))
        self.update_status()

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
        self.update_status()

    def update_status(self):
        text = f"Билетов загружено: {len(self.tickets)}  ·  групп: {len(self.groups)}"
        if self.group_var.get():
            text += f"  ·  студентов в группе: {len(self.current_students())}"
        self.status_var.set(text)

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
        self.show_result(f"{last} {first}".strip(), group, number, repeat)

    def save_result(self, group, last, first, number, repeat):
        while True:
            try:
                td.append_result(RESULTS_FILE, group, last, first, number, repeat)
                return True
            except PermissionError:
                retry = messagebox.askretrycancel(
                    "Файл занят",
                    f"Файл {os.path.basename(RESULTS_FILE)} открыт в другой программе "
                    "(например, в Excel).\nЗакройте его и повторите попытку.",
                    parent=self)
                if not retry:
                    return False

    def show_result(self, student, group, number, repeat):
        c = COLORS
        win = tk.Toplevel(self)
        win.title(f"Билет № {number}")
        win.configure(bg=c["bg"])
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        header(win, f"Билет № {number}", f"{student}  ·  группа {group}")
        body = card(win)

        questions = self.tickets.get(number)
        if questions:
            for i, question in enumerate(questions, 1):
                row = tk.Frame(body, bg=c["card"])
                row.pack(fill="x", pady=5)
                tk.Label(row, text=str(i), bg=c["badge_bg"], fg=c["accent"], width=3,
                         font=(FONT, 11, "bold"), pady=4).pack(side="left", anchor="n")
                tk.Label(row, text=question, bg=c["card"], fg=c["text"], font=(FONT, 11),
                         wraplength=440, justify="left").pack(side="left", padx=(12, 0),
                                                              anchor="w")
        else:
            tk.Label(body, text="Вопросы этого билета не найдены в tickets.docx.",
                     bg=c["card"], fg=c["error"], font=(FONT, 10)).pack(anchor="w")
        if repeat:
            tk.Label(body, text="Повторный запрос: выдан ранее назначенный билет.",
                     bg=c["warn_bg"], fg=c["warn_text"], font=(FONT, 9, "bold"),
                     padx=10, pady=6, anchor="w").pack(fill="x", pady=(12, 0))
        ttk.Label(body, text="Esc — закрыть", style="Hint.TLabel").pack(anchor="e", pady=(14, 0))
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
