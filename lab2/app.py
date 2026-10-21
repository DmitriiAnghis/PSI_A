"""Лабораторная работа №2: генератор экзаменационных билетов с графическим интерфейсом.

Всё приложение работает в одном окне с вкладками:
«Билет» (выдача), «Студенты» (группы и студенты), «Журнал» (results.xlsx), «Билеты» (tickets.docx).
"""
import logging
import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk

import ticket_data as td
from excel_closer import close_in_office

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
    "error_bg": "#FFF0F0",
    "ok_bg": "#EBFBEE",
    "ok_text": "#2B8A3E",
    "secondary_bg": "#EDF1FF",
    "secondary_hover": "#DCE3FF",
    "row_alt": "#F8F9FD",
}
ALL_GROUPS = "Все группы"


def setup_styles(root):
    style = ttk.Style(root)
    style.theme_use("clam")
    c = COLORS
    style.configure(".", font=(FONT, 10), background=c["card"], foreground=c["text"])
    style.configure("Card.TFrame", background=c["card"])
    style.configure("Bg.TFrame", background=c["bg"])
    style.configure("Field.TLabel", background=c["card"], foreground=c["muted"],
                    font=(FONT, 9, "bold"))
    style.configure("Hint.TLabel", background=c["card"], foreground=c["muted"], font=(FONT, 9))
    style.configure("Title.TLabel", background=c["card"], foreground=c["text"],
                    font=(FONT, 12, "bold"))
    style.configure("TCombobox", padding=6, fieldbackground=c["card"],
                    bordercolor=c["border"], lightcolor=c["border"], darkcolor=c["border"],
                    arrowcolor=c["accent"])
    style.map("TCombobox", bordercolor=[("focus", c["accent"])],
              fieldbackground=[("readonly", c["card"])],
              selectbackground=[("readonly", c["card"])],
              selectforeground=[("readonly", c["text"])])
    style.configure("TEntry", padding=6, fieldbackground=c["card"], bordercolor=c["border"],
                    lightcolor=c["border"], darkcolor=c["border"])
    style.map("TEntry", bordercolor=[("focus", c["accent"])])
    style.configure("Accent.TButton", font=(FONT, 11, "bold"), padding=(18, 10),
                    background=c["accent"], foreground="#FFFFFF", borderwidth=0,
                    focuscolor=c["accent"])
    style.map("Accent.TButton",
              background=[("disabled", c["accent_disabled"]), ("active", c["accent_hover"])],
              foreground=[("disabled", "#FFFFFF")])
    style.configure("Secondary.TButton", font=(FONT, 10, "bold"), padding=(14, 8),
                    background=c["secondary_bg"], foreground=c["accent"], borderwidth=0,
                    focuscolor=c["secondary_bg"])
    style.map("Secondary.TButton",
              background=[("disabled", c["bg"]), ("active", c["secondary_hover"])],
              foreground=[("disabled", c["accent_disabled"])])
    style.configure("Danger.TButton", font=(FONT, 10, "bold"), padding=(14, 8),
                    background=c["error_bg"], foreground=c["error"], borderwidth=0,
                    focuscolor=c["error_bg"])
    style.map("Danger.TButton", background=[("disabled", c["bg"]), ("active", "#FFE3E3")],
              foreground=[("disabled", "#E8A0A0")])
    # вкладки
    style.configure("TNotebook", background=c["bg"], borderwidth=0, tabmargins=(20, 12, 20, 0),
                    bordercolor=c["bg"], lightcolor=c["bg"], darkcolor=c["bg"])
    style.configure("TNotebook.Tab", font=(FONT, 10, "bold"), padding=(18, 8),
                    background=c["bg"], foreground=c["muted"], borderwidth=0,
                    bordercolor=c["bg"], lightcolor=c["bg"], darkcolor=c["bg"],
                    focuscolor=c["bg"])
    style.map("TNotebook.Tab",
              background=[("selected", c["accent"]), ("active", c["secondary_bg"])],
              foreground=[("selected", "#FFFFFF"), ("active", c["accent"])],
              bordercolor=[("selected", c["accent"])], lightcolor=[("selected", c["accent"])],
              expand=[("selected", (0, 0, 0, 0))])
    style.layout("TNotebook.Tab", [("Notebook.tab", {"sticky": "nswe", "children": [
        ("Notebook.padding", {"side": "top", "sticky": "nswe", "children": [
            ("Notebook.label", {"side": "top", "sticky": ""})]})]})])
    # таблицы
    style.configure("Treeview", font=(FONT, 10), rowheight=28, background=c["card"],
                    fieldbackground=c["card"], bordercolor=c["border"], borderwidth=0)
    style.map("Treeview", background=[("selected", c["accent"])],
              foreground=[("selected", "#FFFFFF")])
    style.configure("Treeview.Heading", font=(FONT, 9, "bold"), background=c["bg"],
                    foreground=c["muted"], borderwidth=0, padding=(8, 6))
    style.map("Treeview.Heading", background=[("active", c["border"])])
    root.option_add("*TCombobox*Listbox.font", (FONT, 10))
    root.option_add("*TCombobox*Listbox.selectBackground", c["accent"])


def card(parent, **grid):
    """Белая карточка с тонкой рамкой. grid — параметры размещения внешней рамки."""
    outer = tk.Frame(parent, bg=COLORS["border"], padx=1, pady=1)
    outer.grid(**grid)
    inner = ttk.Frame(outer, style="Card.TFrame", padding=18)
    inner.pack(fill="both", expand=True)
    return inner


def label(parent, text="", fg="text", bg="card", size=10, bold=False, **kw):
    return tk.Label(parent, text=text, fg=COLORS[fg], bg=COLORS[bg],
                    font=(FONT, size, "bold" if bold else "normal"), **kw)


def make_table(parent, columns, widths, height=10):
    """Treeview со скроллбаром и полосатыми строками."""
    frame = ttk.Frame(parent, style="Card.TFrame")
    tree = ttk.Treeview(frame, columns=columns, show="headings", height=height,
                        selectmode="browse")
    for column, width in zip(columns, widths):
        tree.heading(column, text=column, anchor="w")
        tree.column(column, width=width, anchor="w", stretch=column == columns[-1])
    tree.tag_configure("alt", background=COLORS["row_alt"])
    scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    tree.pack(side="left", fill="both", expand=True)
    scroll.pack(side="right", fill="y")
    return frame, tree


def fill_table(tree, rows):
    tree.delete(*tree.get_children())
    for i, row in enumerate(rows):
        tree.insert("", "end", values=row, tags=("alt",) if i % 2 else ())


class TicketApp(tk.Tk):
    def __init__(self, groups, tickets):
        super().__init__()
        self.groups = groups
        self.tickets = tickets
        self.title("Генератор экзаменационных билетов")
        self.configure(bg=COLORS["bg"])
        self.geometry("980x700")
        self.minsize(900, 660)
        setup_styles(self)

        self._build_header()
        self._build_banner()
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)
        self.ticket_tab = TicketTab(self.notebook, self)
        self.students_tab = StudentsTab(self.notebook, self)
        self.journal_tab = JournalTab(self.notebook, self)
        self.tickets_tab = TicketsTab(self.notebook, self)
        for tab, title in ((self.ticket_tab, "  Билет  "), (self.students_tab, "  Студенты  "),
                           (self.journal_tab, "  Журнал  "), (self.tickets_tab, "  Билеты  ")):
            self.notebook.add(tab, text=title)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        self.bind("<Escape>", lambda _e: self.ticket_tab.clear_result())
        self.refresh_groups()

    # --- шапка и уведомления ---

    def _build_header(self):
        c = COLORS
        bar = tk.Frame(self, bg=c["header"], padx=24, pady=14)
        bar.pack(fill="x")
        label(bar, "Генератор экзаменационных билетов", "header_text", "header", 15,
              True).pack(anchor="w")
        self.header_stats = label(bar, "", "header_muted", "header", 9)
        self.header_stats.pack(anchor="w", pady=(2, 0))

    def _build_banner(self):
        self.banner = tk.Frame(self, bg=COLORS["bg"])
        self.banner_label = label(self.banner, "", bg="bg", size=10, anchor="w",
                                  justify="left", padx=12, pady=8)
        self.banner_label.pack(side="left", fill="x", expand=True)
        self.banner_action = ttk.Button(self.banner, style="Secondary.TButton")
        self.banner_close = tk.Label(self.banner, text="✕", cursor="hand2", padx=12,
                                     font=(FONT, 11))
        self.banner_close.bind("<Button-1>", lambda _e: self.hide_banner())
        self._banner_after = None

    def notify(self, text, kind="ok", action=None, action_text="Повторить"):
        """Полоса уведомления под шапкой вместо отдельных окон.

        kind: ok / warn / error. action — функция для кнопки (например, «Повторить»).
        """
        bg, fg = {"ok": ("ok_bg", "ok_text"), "warn": ("warn_bg", "warn_text"),
                  "error": ("error_bg", "error")}[kind]
        for widget in (self.banner, self.banner_label, self.banner_close):
            widget.configure(bg=COLORS[bg])
        self.banner_label.configure(text=text, fg=COLORS[fg])
        self.banner_close.configure(fg=COLORS[fg])
        self.banner_action.pack_forget()
        self.banner_close.pack_forget()
        if action:
            self.banner_action.configure(text=action_text,
                                         command=lambda: (self.hide_banner(), action()))
            self.banner_action.pack(side="left", padx=(0, 8), pady=6)
        self.banner_close.pack(side="right")
        self.banner.pack(fill="x", after=self.header_stats.master)
        if self._banner_after:
            self.after_cancel(self._banner_after)
            self._banner_after = None
        if kind == "ok":
            self._banner_after = self.after(4000, self.hide_banner)

    def hide_banner(self):
        self._banner_after = None
        self.banner.pack_forget()

    # --- запись в xlsx ---

    def write_file(self, path, write, retry):
        """Выполняет запись. Если файл открыт в Excel, сначала пробует закрыть его
        сама; если не вышло, показывает полосу с кнопкой «Повторить».
        Возвращает True при успехе."""
        try:
            write()
            return True
        except PermissionError:
            pass
        if close_in_office(path):
            try:
                write()
                return True
            except PermissionError:
                pass
        self.notify(f"Файл {os.path.basename(path)} открыт в другой программе (Excel или Word). "
                    "Закройте его и нажмите «Повторить».", "error", action=retry)
        return False

    # --- общее состояние ---

    def refresh_groups(self):
        names = list(self.groups)
        total = sum(len(s) for s in self.groups.values())
        self.header_stats.configure(
            text=f"Групп: {len(names)}   ·   студентов: {total}   ·   билетов: {len(self.tickets)}")
        self.ticket_tab.set_groups(names)
        self.students_tab.set_groups(names)
        self.journal_tab.set_groups(names)

    def reload_tickets(self):
        try:
            self.tickets = td.load_tickets(TICKETS_FILE)
        except td.SourceError:
            self.tickets = {}
        self.refresh_groups()
        self.tickets_tab.show_tickets()

    def on_tab_changed(self, _event=None):
        if self.notebook.select() == str(self.journal_tab):
            self.journal_tab.reload()


class TicketTab(ttk.Frame):
    """Вкладка «Билет»: выбор группы и студента, выдача билета на этой же странице."""

    def __init__(self, notebook, app):
        super().__init__(notebook, style="Bg.TFrame", padding=(20, 16, 20, 20))
        self.app = app
        self.columnconfigure(0, weight=0, minsize=320)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        form = card(self, row=0, column=0, sticky="nsew", padx=(0, 16))
        form.columnconfigure(0, weight=1)
        ttk.Label(form, text="Выдача билета", style="Title.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 14))
        ttk.Label(form, text="ГРУППА", style="Field.TLabel").grid(row=1, column=0, sticky="w")
        self.group_var = tk.StringVar()
        self.group_box = ttk.Combobox(form, textvariable=self.group_var, state="readonly",
                                      width=30, font=(FONT, 10))
        self.group_box.grid(row=2, column=0, sticky="ew", pady=(4, 14))
        self.group_box.bind("<<ComboboxSelected>>", self.on_group_selected)

        ttk.Label(form, text="СТУДЕНТ", style="Field.TLabel").grid(row=3, column=0, sticky="w")
        self.student_var = tk.StringVar()
        self.student_box = ttk.Combobox(form, textvariable=self.student_var, state="readonly",
                                        values=[], width=30, font=(FONT, 10))
        self.student_box.grid(row=4, column=0, sticky="ew", pady=(4, 18))
        self.student_box.bind("<<ComboboxSelected>>", lambda _e: self.update_button())

        self.generate_btn = ttk.Button(form, text="Сгенерировать билет", style="Accent.TButton",
                                       command=self.on_generate, state="disabled",
                                       cursor="hand2")
        self.generate_btn.grid(row=5, column=0, sticky="ew")
        self.count_var = tk.StringVar()
        ttk.Label(form, textvariable=self.count_var, style="Hint.TLabel").grid(
            row=6, column=0, sticky="w", pady=(12, 0))
        ttk.Label(form, text="Нет студента в списке? Добавьте его\nна вкладке «Студенты».",
                  style="Hint.TLabel", justify="left").grid(row=7, column=0, sticky="sw",
                                                            pady=(24, 0))
        form.rowconfigure(7, weight=1)

        self.result = card(self, row=0, column=1, sticky="nsew")
        self.clear_result()

    # --- выбор ---

    def set_groups(self, names):
        self.group_box["values"] = names
        if self.group_var.get() not in names:
            self.group_var.set("")
        self.on_group_selected(keep_student=True)

    def current_students(self):
        return self.app.groups.get(self.group_var.get(), [])

    def on_group_selected(self, _event=None, keep_student=False):
        students = self.current_students()
        names = [f"{last} {first}".strip() for last, first in students]
        self.student_box["values"] = names
        if not keep_student or self.student_var.get() not in names:
            self.student_var.set("")
        self.count_var.set(f"Студентов в группе: {len(students)}" if self.group_var.get() else
                           f"Билетов загружено: {len(self.app.tickets)}")
        self.update_button()

    def selected_student(self):
        index = self.student_box.current()
        students = self.current_students()
        if 0 <= index < len(students):
            return students[index]
        return None

    def select_student(self, group, last, first):
        self.group_var.set(group)
        self.on_group_selected()
        self.student_box.set(f"{last} {first}".strip())
        self.update_button()

    def update_button(self):
        ready = bool(self.group_var.get()) and self.selected_student() is not None
        self.generate_btn["state"] = "normal" if ready else "disabled"

    # --- генерация ---

    def on_generate(self):
        student = self.selected_student()
        if student is None:
            return
        if not self.app.tickets:
            self.app.notify("Нет ни одного билета: добавьте билеты на вкладке «Билеты».", "error")
            return
        group = self.group_var.get()
        last, first = student
        number, repeat = td.assign_ticket(RESULTS_FILE, self.app.tickets, group, last, first)
        saved = self.app.write_file(
            RESULTS_FILE,
            lambda: td.append_result(RESULTS_FILE, group, last, first, number, repeat),
            retry=self.on_generate)
        if saved:
            self.show_result(f"{last} {first}".strip(), group, number, repeat)

    def _reset_result(self):
        for child in self.result.winfo_children():
            child.destroy()

    def clear_result(self):
        """Пустое состояние карточки билета (Esc)."""
        self._reset_result()
        box = ttk.Frame(self.result, style="Card.TFrame")
        box.place(relx=0.5, rely=0.5, anchor="center")
        label(box, "№", "accent", "badge_bg", 26, True, width=3, pady=6).pack()
        label(box, "Здесь появится билет", size=12, bold=True).pack(pady=(6, 2))
        label(box, "Выберите группу и студента и нажмите\n«Сгенерировать билет»", "muted",
              size=9, justify="center").pack()

    def show_result(self, student, group, number, repeat):
        c = COLORS
        self._reset_result()
        top = tk.Frame(self.result, bg=c["card"])
        top.pack(fill="x")
        label(top, str(number), "accent", "badge_bg", 22, True, width=3, pady=6).pack(
            side="left")
        info = tk.Frame(top, bg=c["card"])
        info.pack(side="left", padx=14)
        label(info, f"Билет № {number}", size=15, bold=True).pack(anchor="w")
        label(info, f"{student}  ·  группа {group}", "muted", size=10).pack(anchor="w")
        tk.Frame(self.result, bg=c["border"], height=1).pack(fill="x", pady=14)

        questions = self.app.tickets.get(number)
        if questions:
            for i, question in enumerate(questions, 1):
                row = tk.Frame(self.result, bg=c["card"])
                row.pack(fill="x", pady=6)
                label(row, str(i), "accent", "badge_bg", 11, True, width=3, pady=4).pack(
                    side="left", anchor="n")
                label(row, question, size=11, wraplength=420, justify="left").pack(
                    side="left", padx=(12, 0), anchor="w")
        else:
            label(self.result, "Вопросы этого билета не найдены в tickets.docx.",
                  "error").pack(anchor="w")
        if repeat:
            label(self.result, "Повторный запрос: выдан ранее назначенный билет.", "warn_text",
                  "warn_bg", 9, True, padx=10, pady=6, anchor="w").pack(fill="x", pady=(14, 0))
        ttk.Label(self.result, text="Esc — очистить", style="Hint.TLabel").pack(
            side="bottom", anchor="e")


class StudentsTab(ttk.Frame):
    """Вкладка «Студенты»: список групп, студенты группы, добавление и удаление."""

    def __init__(self, notebook, app):
        super().__init__(notebook, style="Bg.TFrame", padding=(20, 16, 20, 20))
        self.app = app
        self.columnconfigure(0, minsize=250)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # группы
        left = card(self, row=0, column=0, sticky="nsew", padx=(0, 16))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        ttk.Label(left, text="Группы", style="Title.TLabel").grid(row=0, column=0, sticky="w",
                                                                  pady=(0, 10))
        frame, self.groups_tree = make_table(left, ("Группа", "Студентов"), (110, 100), 8)
        frame.grid(row=1, column=0, sticky="nsew")
        self.groups_tree.bind("<<TreeviewSelect>>", lambda _e: self.on_group_clicked())
        self._shown_group = None
        ttk.Label(left, text="НОВАЯ ГРУППА", style="Field.TLabel").grid(
            row=2, column=0, sticky="w", pady=(14, 0))
        self.new_group_var = tk.StringVar()
        entry = ttk.Entry(left, textvariable=self.new_group_var, font=(FONT, 10))
        entry.grid(row=3, column=0, sticky="ew", pady=(4, 8))
        entry.bind("<Return>", lambda _e: self.on_add_group())
        ttk.Button(left, text="+ Добавить группу", style="Secondary.TButton",
                   command=self.on_add_group, cursor="hand2").grid(row=4, column=0, sticky="ew")
        self.del_group_btn = ttk.Button(left, text="Удалить пустую группу", style="Danger.TButton",
                                        command=self.on_delete_group, cursor="hand2")
        self.del_group_btn.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        self.group_error = label(left, "", "error", size=9, wraplength=210, justify="left")
        self.group_error.grid(row=6, column=0, sticky="w", pady=(6, 0))

        # студенты
        right = card(self, row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)
        head = ttk.Frame(right, style="Card.TFrame")
        head.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.title_var = tk.StringVar(value="Студенты")
        ttk.Label(head, textvariable=self.title_var, style="Title.TLabel").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.show_group())
        search = ttk.Entry(head, textvariable=self.search_var, width=22, font=(FONT, 10))
        search.pack(side="right")
        ttk.Label(head, text="Поиск", style="Hint.TLabel").pack(side="right", padx=(0, 6))

        frame, self.students_tree = make_table(right, ("№", "Фамилия", "Имя"), (40, 170, 150), 9)
        frame.grid(row=2, column=0, sticky="nsew")
        self.students_tree.bind("<<TreeviewSelect>>", lambda _e: self.update_buttons())
        self.students_tree.bind("<Double-1>", lambda _e: self.open_in_ticket_tab())

        form = ttk.Frame(right, style="Card.TFrame")
        form.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)
        ttk.Label(form, text="ФАМИЛИЯ", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(form, text="ИМЯ", style="Field.TLabel").grid(row=0, column=1, sticky="w",
                                                               padx=(10, 0))
        self.last_var = tk.StringVar()
        self.first_var = tk.StringVar()
        last = ttk.Entry(form, textvariable=self.last_var, font=(FONT, 10))
        last.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        first = ttk.Entry(form, textvariable=self.first_var, font=(FONT, 10))
        first.grid(row=1, column=1, sticky="ew", pady=(4, 0), padx=(10, 0))
        for widget in (last, first):
            widget.bind("<Return>", lambda _e: self.on_add_student())
        self.add_btn = ttk.Button(form, text="+ Добавить студента", style="Accent.TButton",
                                  command=self.on_add_student, cursor="hand2")
        self.add_btn.grid(row=1, column=2, padx=(10, 0), pady=(4, 0), sticky="ns")
        self.student_error = label(form, "", "error", size=9)
        self.student_error.grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))

        actions = ttk.Frame(right, style="Card.TFrame")
        actions.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.ticket_btn = ttk.Button(actions, text="Выдать билет", style="Secondary.TButton",
                                     command=self.open_in_ticket_tab, cursor="hand2")
        self.ticket_btn.pack(side="left")
        self.delete_btn = ttk.Button(actions, text="Удалить", style="Danger.TButton",
                                     command=self.on_delete_student, cursor="hand2")
        self.delete_btn.pack(side="left", padx=(8, 0))
        ttk.Label(actions, text="Двойной щелчок — выдать билет",
                  style="Hint.TLabel").pack(side="right")

    # --- данные ---

    def set_groups(self, names, select=None):
        current = select or self.current_group()
        fill_table(self.groups_tree, [(n, len(self.app.groups[n])) for n in names])
        items = self.groups_tree.get_children()
        target = next((i for i in items if self.groups_tree.item(i, "values")[0] == current),
                      items[0] if items else None)
        if target:
            self.groups_tree.selection_set(target)
            self.groups_tree.see(target)
        self.show_group()

    def current_group(self):
        selection = self.groups_tree.selection()
        return self.groups_tree.item(selection[0], "values")[0] if selection else ""

    def visible_students(self):
        query = self.search_var.get().strip().casefold()
        students = self.app.groups.get(self.current_group(), [])
        return [s for s in students if query in f"{s[0]} {s[1]}".casefold()]

    def on_group_clicked(self):
        # событие приходит и после программного выбора: не перерисовываем ту же группу
        if self.current_group() != self._shown_group:
            self.show_group()

    def show_group(self):
        group = self.current_group()
        self._shown_group = group
        students = self.visible_students()
        self.title_var.set(f"Студенты группы {group}" if group else "Студенты")
        fill_table(self.students_tree, [(i, last, first)
                                        for i, (last, first) in enumerate(students, 1)])
        self.add_btn["state"] = "normal" if group else "disabled"
        self.update_buttons()

    def selected_student(self):
        selection = self.students_tree.selection()
        if not selection:
            return None
        _, last, first = self.students_tree.item(selection[0], "values")
        return last, first

    def update_buttons(self):
        state = "normal" if self.selected_student() else "disabled"
        self.delete_btn["state"] = state
        self.ticket_btn["state"] = state

    # --- действия ---

    def on_add_group(self):
        self.group_error.configure(text="")
        name = self.new_group_var.get()
        result = {}

        def write():
            result["name"] = td.add_group(STUDENTS_FILE, name)
        try:
            if not self.app.write_file(STUDENTS_FILE, write, retry=self.on_add_group):
                return
        except (td.InvalidNameError, td.DuplicateGroupError) as exc:
            self.group_error.configure(text=str(exc))
            return
        group = result["name"]
        self.app.groups[group] = []
        self.new_group_var.set("")
        self.app.refresh_groups()
        self.set_groups(list(self.app.groups), select=group)
        self.app.notify(f"Группа {group} добавлена.")

    def on_delete_group(self):
        self.group_error.configure(text="")
        group = self.current_group()
        if not group:
            return
        if self.app.groups.get(group):
            self.group_error.configure(text=f"В группе {group} есть студенты: сначала удалите их.")
            return
        self.app.notify(f"Удалить пустую группу {group}?", "warn",
                        action=lambda: self.delete_group(group), action_text="Удалить")

    def delete_group(self, group):
        try:
            if not self.app.write_file(STUDENTS_FILE, lambda: td.remove_group(STUDENTS_FILE, group),
                                       retry=lambda: self.delete_group(group)):
                return
        except (td.InvalidNameError, td.SourceError) as exc:
            self.group_error.configure(text=str(exc))
            return
        self.app.groups.pop(group, None)
        self.app.refresh_groups()
        self.app.notify(f"Группа {group} удалена.", "warn")

    def on_add_student(self):
        self.student_error.configure(text="")
        group = self.current_group()
        if not group:
            self.student_error.configure(text="Сначала выберите группу слева.")
            return
        result = {}

        def write():
            result["name"] = td.add_student(STUDENTS_FILE, group, self.last_var.get(),
                                            self.first_var.get())
        try:
            if not self.app.write_file(STUDENTS_FILE, write, retry=self.on_add_student):
                return
        except (td.InvalidNameError, td.DuplicateStudentError, td.SourceError) as exc:
            self.student_error.configure(text=str(exc))
            return
        last, first = result["name"]
        self.app.groups[group].append((last, first))
        self.last_var.set("")
        self.first_var.set("")
        self.search_var.set("")
        self.app.refresh_groups()
        self.select_in_table(last, first)
        self.app.notify(f"{last} {first} добавлен(а) в группу {group}.")

    def select_in_table(self, last, first):
        for item in self.students_tree.get_children():
            if tuple(self.students_tree.item(item, "values")[1:]) == (last, first):
                self.students_tree.selection_set(item)
                self.students_tree.see(item)

    def on_delete_student(self):
        student = self.selected_student()
        group = self.current_group()
        if not student or not group:
            return
        last, first = student
        self.app.notify(f"Удалить {last} {first} из группы {group}?", "warn",
                        action=lambda: self.delete_student(group, last, first),
                        action_text="Удалить")

    def delete_student(self, group, last, first):
        result = {}

        def write():
            result["ok"] = td.remove_student(STUDENTS_FILE, group, last, first)
        if not self.app.write_file(STUDENTS_FILE, write, retry=lambda: self.delete_student(group, last, first)):
            return
        students = self.app.groups[group]
        if (last, first) in students:
            students.remove((last, first))
        self.app.refresh_groups()
        self.app.notify(f"{last} {first} удалён(а) из группы {group}. "
                        "Записи в журнале остаются.", "warn")

    def open_in_ticket_tab(self):
        student = self.selected_student()
        if not student:
            return
        self.app.ticket_tab.select_student(self.current_group(), *student)
        self.app.notebook.select(self.app.ticket_tab)


class JournalTab(ttk.Frame):
    """Вкладка «Журнал»: содержимое results.xlsx, фильтр по группе и сводка."""

    def __init__(self, notebook, app):
        super().__init__(notebook, style="Bg.TFrame", padding=(20, 16, 20, 20))
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        tiles = ttk.Frame(self, style="Bg.TFrame")
        tiles.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        self.tile_vars = {}
        for i, (key, title) in enumerate((("issued", "Выдано билетов"),
                                          ("students", "Студентов"),
                                          ("repeats", "Повторных запросов"),
                                          ("top", "Чаще всего выпадал"))):
            tiles.columnconfigure(i, weight=1, uniform="tile")
            box = card(tiles, row=0, column=i, sticky="ew", padx=(0 if i == 0 else 12, 0))
            var = tk.StringVar(value="—")
            self.tile_vars[key] = var
            tk.Label(box, textvariable=var, bg=COLORS["card"], fg=COLORS["accent"],
                     font=(FONT, 18, "bold")).pack(anchor="w")
            ttk.Label(box, text=title, style="Hint.TLabel").pack(anchor="w")

        body = card(self, row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)
        bar = ttk.Frame(body, style="Card.TFrame")
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(bar, text="Журнал выдачи", style="Title.TLabel").pack(side="left")
        ttk.Button(bar, text="Открыть в Excel", style="Secondary.TButton",
                   command=self.open_file, cursor="hand2").pack(side="right")
        ttk.Button(bar, text="Экспорт CSV", style="Secondary.TButton",
                   command=self.export_csv, cursor="hand2").pack(side="right", padx=(0, 8))
        ttk.Button(bar, text="Обновить", style="Secondary.TButton", command=self.reload,
                   cursor="hand2").pack(side="right", padx=(0, 8))
        self.filter_var = tk.StringVar(value=ALL_GROUPS)
        self.filter_box = ttk.Combobox(bar, textvariable=self.filter_var, state="readonly",
                                       width=14, font=(FONT, 10))
        self.filter_box.pack(side="right", padx=(0, 12))
        self.filter_box.bind("<<ComboboxSelected>>", lambda _e: self.reload())

        columns = ("Группа", "Фамилия", "Имя", "Билет", "Дата и время", "Повтор")
        frame, self.tree = make_table(body, columns, (90, 140, 120, 60, 160, 70), 12)
        frame.grid(row=1, column=0, sticky="nsew")
        self.tree.tag_configure("repeat", foreground=COLORS["warn_text"])

    def set_groups(self, names):
        self.filter_box["values"] = [ALL_GROUPS] + names
        if self.filter_var.get() not in self.filter_box["values"]:
            self.filter_var.set(ALL_GROUPS)

    def reload(self):
        try:
            rows = td.load_results(RESULTS_FILE)
        except Exception as exc:  # файл повреждён или занят
            self.app.notify(f"Не удалось прочитать журнал: {exc}", "error", action=self.reload)
            return
        group = self.filter_var.get()
        if group != ALL_GROUPS:
            rows = [r for r in rows if r[0] == group]
        self.rows = rows
        rows = list(reversed(rows))  # новые сверху
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(rows):
            tags = ["alt"] if i % 2 else []
            if row[5].casefold() == "да":
                tags.append("repeat")
            self.tree.insert("", "end", values=row, tags=tags)
        stats = td.results_stats(rows)
        self.tile_vars["issued"].set(stats["issued"])
        self.tile_vars["students"].set(stats["students"])
        self.tile_vars["repeats"].set(stats["repeats"])
        self.tile_vars["top"].set(f"№ {stats['top_ticket']}  ×{stats['top_count']}"
                                  if stats["top_ticket"] else "—")

    def export_csv(self):
        """Сохраняет то, что сейчас показано в таблице (с учётом фильтра), в data/*.csv."""
        from datetime import datetime
        self.reload()
        rows = getattr(self, "rows", [])
        if not rows:
            self.app.notify("Нечего экспортировать: журнал пуст.", "warn")
            return
        group = self.filter_var.get()
        suffix = "" if group == ALL_GROUPS else f"_{group}"
        name = f"journal{suffix}_{datetime.now():%Y%m%d_%H%M}.csv"
        path = os.path.join(os.path.dirname(RESULTS_FILE), name)
        td.export_results_csv(rows, path)
        self.app.notify(f"Журнал сохранён: {path}")

    def open_file(self):
        if not os.path.exists(RESULTS_FILE):
            self.app.notify("Журнал ещё пуст: выдайте первый билет.", "warn")
            return
        if hasattr(os, "startfile"):
            os.startfile(RESULTS_FILE)  # Windows: откроется в Excel
        else:
            self.app.notify(f"Файл журнала: {RESULTS_FILE}", "warn")


class TicketsTab(ttk.Frame):
    """Вкладка «Билеты»: список билетов из tickets.docx, добавление, изменение, удаление."""

    def __init__(self, notebook, app):
        super().__init__(notebook, style="Bg.TFrame", padding=(20, 16, 20, 20))
        self.app = app
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, minsize=340)
        self.rowconfigure(0, weight=1)

        body = card(self, row=0, column=0, sticky="nsew", padx=(0, 16))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)
        head = ttk.Frame(body, style="Card.TFrame")
        head.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.title_var = tk.StringVar()
        ttk.Label(head, textvariable=self.title_var, style="Title.TLabel").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.show_tickets())
        ttk.Entry(head, textvariable=self.search_var, width=18, font=(FONT, 10)).pack(side="right")
        ttk.Label(head, text="Поиск", style="Hint.TLabel").pack(side="right", padx=(0, 6))

        frame = ttk.Frame(body, style="Card.TFrame")
        frame.grid(row=1, column=0, sticky="nsew")
        self.tree = ttk.Treeview(frame, show="tree", selectmode="browse")
        self.tree.column("#0", width=420, stretch=True)
        self.tree.tag_configure("ticket", font=(FONT, 10, "bold"), foreground=COLORS["accent"])
        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self.on_select())
        ttk.Label(body, text="Билеты, где меньше трёх вопросов, пропущены и записаны в "
                  "tickets.log", style="Hint.TLabel").grid(row=2, column=0, sticky="w",
                                                           pady=(10, 0))

        form = card(self, row=0, column=1, sticky="nsew")
        form.columnconfigure(0, weight=1)
        self.form_title = tk.StringVar(value="Новый билет")
        ttk.Label(form, textvariable=self.form_title, style="Title.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 12))
        ttk.Label(form, text="НОМЕР (пусто — следующий свободный)", style="Field.TLabel").grid(
            row=1, column=0, sticky="w")
        self.number_var = tk.StringVar()
        self.number_entry = ttk.Entry(form, textvariable=self.number_var, width=10,
                                      font=(FONT, 10))
        self.number_entry.grid(row=2, column=0, sticky="w", pady=(4, 12))
        self.question_vars = []
        for i in range(td.QUESTIONS_PER_TICKET):
            ttk.Label(form, text=f"ВОПРОС {i + 1}", style="Field.TLabel").grid(
                row=3 + i * 2, column=0, sticky="w")
            var = tk.StringVar()
            ttk.Entry(form, textvariable=var, font=(FONT, 10)).grid(
                row=4 + i * 2, column=0, sticky="ew", pady=(4, 8))
            self.question_vars.append(var)
        self.error = label(form, "", "error", size=9, wraplength=300, justify="left")
        self.error.grid(row=9, column=0, sticky="w")
        self.add_btn = ttk.Button(form, text="+ Добавить билет", style="Accent.TButton",
                                  command=self.on_add, cursor="hand2")
        self.add_btn.grid(row=10, column=0, sticky="ew", pady=(6, 0))
        self.save_btn = ttk.Button(form, text="Сохранить изменения", style="Secondary.TButton",
                                   command=self.on_save, cursor="hand2")
        self.save_btn.grid(row=11, column=0, sticky="ew", pady=(8, 0))
        row = ttk.Frame(form, style="Card.TFrame")
        row.grid(row=12, column=0, sticky="ew", pady=(8, 0))
        row.columnconfigure(0, weight=1)
        row.columnconfigure(1, weight=1)
        self.delete_btn = ttk.Button(row, text="Удалить билет", style="Danger.TButton",
                                     command=self.on_delete, cursor="hand2")
        self.delete_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(row, text="Очистить", style="Secondary.TButton", command=self.clear_form,
                   cursor="hand2").grid(row=0, column=1, sticky="ew", padx=(4, 0))
        self.selected = None
        self.show_tickets()
        self.clear_form()

    # --- список ---

    def show_tickets(self):
        query = self.search_var.get().strip().casefold()
        tickets = self.app.tickets
        self.title_var.set(f"Билеты из tickets.docx: {len(tickets)}")
        self.tree.delete(*self.tree.get_children())
        for number in sorted(tickets):
            questions = tickets[number]
            if query and query not in " ".join(questions).casefold() and query != str(number):
                continue
            parent = self.tree.insert("", "end", iid=f"t{number}", text=f"Билет № {number}",
                                      open=True, tags=("ticket",))
            for i, question in enumerate(questions, 1):
                self.tree.insert(parent, "end", text=f"{i}. {question}")
        if self.selected in tickets and self.tree.exists(f"t{self.selected}"):
            self.tree.selection_set(f"t{self.selected}")

    def selected_number(self):
        selection = self.tree.selection()
        if not selection:
            return None
        item = selection[0]
        parent = self.tree.parent(item) or item
        return int(parent[1:])

    def on_select(self):
        number = self.selected_number()
        if number is None or number not in self.app.tickets:
            return
        self.selected = number
        self.form_title.set(f"Билет № {number}")
        self.number_var.set(str(number))
        self.number_entry.state(["disabled"])
        for var, question in zip(self.question_vars, self.app.tickets[number]):
            var.set(question)
        self.error.configure(text="")
        self.save_btn.state(["!disabled"])
        self.delete_btn.state(["!disabled"])
        self.add_btn.state(["disabled"])

    def clear_form(self):
        self.selected = None
        self.tree.selection_remove(*self.tree.selection())
        self.form_title.set("Новый билет")
        self.number_entry.state(["!disabled"])
        self.number_var.set("")
        for var in self.question_vars:
            var.set("")
        self.error.configure(text="")
        self.save_btn.state(["disabled"])
        self.delete_btn.state(["disabled"])
        self.add_btn.state(["!disabled"])

    # --- действия ---

    def questions(self):
        return [var.get() for var in self.question_vars]

    def _run(self, write, retry):
        try:
            return self.app.write_file(TICKETS_FILE, write, retry=retry)
        except (td.TicketError, td.SourceError) as exc:
            self.error.configure(text=str(exc))
            return False

    def on_add(self):
        self.error.configure(text="")
        number = self.number_var.get().strip() or None
        result = {}

        def write():
            result["n"] = td.add_ticket(TICKETS_FILE, self.questions(), number=number)
        if self._run(write, self.on_add):
            self.app.reload_tickets()
            self.clear_form()
            self.app.notify(f"Билет № {result['n']} добавлен в tickets.docx.")

    def on_save(self):
        number = self.selected
        if number is None:
            return
        self.error.configure(text="")
        if self._run(lambda: td.update_ticket(TICKETS_FILE, number, self.questions()),
                     self.on_save):
            self.app.reload_tickets()
            self.selected = number
            self.show_tickets()
            self.app.notify(f"Билет № {number} сохранён.")

    def on_delete(self):
        number = self.selected
        if number is None:
            return
        self.app.notify(f"Удалить билет № {number} из tickets.docx? Записи в журнале останутся.",
                        "warn", action=lambda: self.delete(number), action_text="Удалить")

    def delete(self, number):
        if self._run(lambda: td.remove_ticket(TICKETS_FILE, number), lambda: self.delete(number)):
            self.app.reload_tickets()
            self.clear_form()
            self.app.notify(f"Билет № {number} удалён.", "warn")

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
