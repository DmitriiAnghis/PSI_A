# Prompt for GitHub Copilot Chat (Step 1 — Prompt Engineering)

Build a desktop Python 3 application with a Tkinter GUI: an exam ticket generator.

Inputs:
- `data/students.xlsx`: one sheet per group, sheet name = group name. Row 1 is the header, data from row 2: column A = last name, column B = first name.
- `data/tickets.docx`: tickets in the form "Билет N" followed by three numbered paragraphs "1. text", "2. text", "3. text".

Output:
- `data/results.xlsx` with columns: Группа | Фамилия | Имя | Номер билета | Дата и время | Повтор (да/нет). Create it with the header if missing; only append rows, never modify existing ones. Save immediately after each generation.

UI behaviour:
1. Combobox 1 is filled with the sheet names (groups) on start.
2. Selecting a group fills Combobox 2 with that group's students as "LastName FirstName".
3. The "Сгенерировать билет" button is disabled until both a group and a student are selected.
4. Clicking it opens a result window "Ваш билет № N" with the ticket's three questions.
5. ESC closes the result window only; the main window stays open for the next student.

Ticket assignment:
- If the student (group + last name + first name) already exists in results.xlsx, reuse the ticket number from their FIRST record and append a new row with "да" in the repeat column.
- Otherwise choose a random ticket from 1..N (N = number of tickets parsed from the docx) and append a row with "нет".

Robustness:
- Missing students.xlsx or tickets.docx: show a clear error message, do not crash.
- Empty group: empty student list, button stays disabled, no exception.
- results.xlsx locked by Excel (PermissionError): show a message with Retry/Cancel, do not crash.
- A ticket with fewer than 3 questions or an unrecognised header is skipped and logged; the app continues with the remaining tickets.

Structure: keep file logic in `ticket_data.py` (pure functions, no Tkinter) and UI in `app.py`.
Write pytest unit tests for: parsing a ticket from Word, listing students by group, idempotency (selecting the same student again returns the same ticket).
