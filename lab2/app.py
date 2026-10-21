"""Лабораторная работа №2: генератор экзаменационных билетов с графическим интерфейсом."""
import logging
import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk

import ticket_data as td
from excel_closer import close_in_excel

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
    "secondary_bg": "#EDF1FF",
    "secondary_hover": "#DCE3FF",
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
    style.configure("Secondary.TButton", font=(FONT, 10, "bold"), padding=(14, 8),
                    background=c["secondary_bg"], foreground=c["accent"], borderwidth=0,
                    focuscolor=c["secondary_bg"])
    style.map("Secondary.TButton", background=[("active", c["secondary_hover"])])
    style.configure("TEntry", padding=6, fieldbackground=c["card"], bordercolor=c["border"],
                    lightcolor=c["border"], darkcolor=c["border"])
    style.map("TEntry", bordercolor=[("focus", c["accent"])])
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

        self.add_btn = ttk.Button(frame, text="+ Добавить студента", style="Secondary.TButton",
                                  command=self.open_add_student, cursor="hand2")
        self.add_btn.grid(row=5, column=0, sticky="ew", pady=(8, 0))

        self.status_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.status_var, style="Hint.TLabel").grid(
            row=6, column=0, sticky="w", pady=(12, 0))
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
        return self.write_with_retry(
            RESULTS_FILE,
            lambda: td.append_result(RESULTS_FILE, group, last, first, number, repeat))

    def write_with_retry(self, path, write, parent=None):
        """Выполняет запись в xlsx. Если файл открыт в Excel, сначала пробует
        закрыть его сама, иначе просит пользователя закрыть файл."""
        auto_closed = False
        while True:
            try:
                write()
                return True
            except PermissionError:
                if not auto_closed and close_in_excel(path):
                    auto_closed = True
                    continue
                retry = messagebox.askretrycancel(
                    "Файл занят",
                    f"Файл {os.path.basename(path)} открыт в другой программе "
                    "(например, в Excel).\nЗакройте его и повторите попытку.",
                    parent=parent or self)
                if not retry:
                    return False

    # --- добавление студента ---

    def open_add_student(self):
        AddStudentDialog(self)

    def student_added(self, group, last, first):
        self.groups.setdefault(group, []).append((last, first))
        self.group_var.set(group)
        self.on_group_selected()
        self.student_box.current(len(self.groups[group]) - 1)
        self.update_button()

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


class AddStudentDialog(tk.Toplevel):
    """Окно «Добавить студента»: группа, фамилия, имя с проверкой ввода."""

    def __init__(self, app):
        super().__init__(app)
        c = COLORS
        self.app = app
        self.title("Добавить студента")
        self.configure(bg=c["bg"])
        self.resizable(False, False)
        self.transient(app)
        self.grab_set()

        header(self, "Новый студент", "Фамилия и имя с заглавной буквы, без цифр")
        body = card(self)
        body.columnconfigure(0, weight=1)

        ttk.Label(body, text="ГРУППА", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        self.group_var = tk.StringVar(value=app.group_var.get() or next(iter(app.groups), ""))
        ttk.Combobox(body, textvariable=self.group_var, state="readonly", width=34,
                     values=list(app.groups), font=(FONT, 10)).grid(
            row=1, column=0, sticky="ew", pady=(4, 12))

        self.last_var = tk.StringVar()
        self.first_var = tk.StringVar()
        for row, (label, var) in enumerate((("ФАМИЛИЯ", self.last_var),
                                           ("ИМЯ", self.first_var))):
            ttk.Label(body, text=label, style="Field.TLabel").grid(
                row=2 + row * 2, column=0, sticky="w")
            entry = ttk.Entry(body, textvariable=var, width=36, font=(FONT, 10))
            entry.grid(row=3 + row * 2, column=0, sticky="ew", pady=(4, 12))
            if row == 0:
                entry.focus_set()

        self.error_var = tk.StringVar()
        tk.Label(body, textvariable=self.error_var, bg=c["card"], fg=c["error"],
                 font=(FONT, 9), wraplength=320, justify="left").grid(
            row=6, column=0, sticky="w")
        ttk.Button(body, text="Добавить", style="Accent.TButton", command=self.submit,
                   cursor="hand2").grid(row=7, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(body, text="Enter — добавить, Esc — отмена", style="Hint.TLabel").grid(
            row=8, column=0, sticky="e", pady=(10, 0))
        self.bind("<Return>", lambda _e: self.submit())
        self.bind("<Escape>", lambda _e: self.destroy())

    def submit(self):
        group = self.group_var.get()
        if not group:
            self.error_var.set("Выберите группу.")
            return
        result = {}

        def write():
            result["name"] = td.add_student(STUDENTS_FILE, group, self.last_var.get(),
                                            self.first_var.get())
        try:
            if not self.app.write_with_retry(STUDENTS_FILE, write, parent=self):
                return
        except (td.InvalidNameError, td.DuplicateStudentError, td.SourceError) as exc:
            self.error_var.set(str(exc))
            return
        last, first = result["name"]
        self.app.student_added(group, last, first)
        self.destroy()


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
