# PSI — лабораторные работы

**Студент:** Dmitrii Anghis, гр. TI-235 · **Преподаватель:** Denis Ahramenco · **УТМ, 2026**

Репозиторий с лабораторными работами по дисциплине PSI. Лабы 1 и 2 лежат здесь, лабы 3 и 4 по условию сделаны в отдельных репозиториях.

| № | Тема | Где | Стек | Тесты |
|---|---|---|---|---|
| 1 | Консольный генератор экзаменационных билетов | [`lab1/`](lab1/) | Python, openpyxl | 8 |
| 2 | Генератор билетов с графическим интерфейсом | [`lab2/`](lab2/) | Python, Tkinter, openpyxl, python-docx | 12 |
| 3 | Конвертер валют BNM (Spec-Driven Development) | [currency-converter-sdd](https://github.com/DmitriiAnghis/currency-converter-sdd) | Python, Tkinter, httpx · Spec Kit, Copilot, Antigravity | 126 |
| 4 | Мобильный конвертер валют BNM | [currency-mobile](https://github.com/DmitriiAnghis/currency-mobile) | Kotlin, Jetpack Compose, Android | 33 |

## Быстрый старт

Нужен Python 3.12 или новее.

```bash
git clone https://github.com/DmitriiAnghis/PSI_A
cd PSI_A
python -m pip install -r lab1/requirements.txt -r lab2/requirements.txt
```

### Лаб 1: консольный генератор билетов

```bash
cd lab1
python ticket_generator.py   # запуск
python -m pytest             # тесты
```

- Запрашивает **Last name** и **First name**, пустой ввод не принимает.
- Выдаёт случайный билет **от 1 до 20** и сразу дописывает строку в `journal.xlsx` (фамилия, имя, номер, дата и время).
- **ESC** завершает программу.
- Если `journal.xlsx` открыт в Excel, программа не падает: просит закрыть файл и нажать Enter, потом повторяет сохранение.
- Записи сохраняются сразу, поэтому даже при принудительном закрытии (`taskkill`) уже введённые данные остаются в журнале.

Код: [`lab1/ticket_generator.py`](lab1/ticket_generator.py) · тесты: [`lab1/tests/`](lab1/tests/)

### Лаб 2: генератор билетов с GUI

```bash
cd lab2
python app.py                # запуск
python -m pytest             # тесты
python make_sample_data.py   # пересоздать тестовые students.xlsx и tickets.docx
```

- Окно Tkinter: выбор **группы** (листы `data/students.xlsx`) → выбор **студента** → «Сгенерировать билет».
- Показывает номер билета и три вопроса из `data/tickets.docx`, сразу записывает строку в `data/results.xlsx`.
- Повторный запрос для того же студента выдаёт **тот же номер** и пишет «Повтор = да».
- Нет исходного файла: понятное сообщение об ошибке. `results.xlsx` открыт в Excel: окно «Файл занят» с кнопкой «Повторить».
- **Esc** закрывает окно билета, приложение остаётся открытым.

Код: [`lab2/app.py`](lab2/app.py), [`lab2/ticket_data.py`](lab2/ticket_data.py) · тесты: [`lab2/tests/`](lab2/tests/)

### Лаб 3: конвертер валют BNM (отдельный репозиторий)

Desktop-конвертер по официальному курсу Национального банка Молдовы, сделан по методике Spec-Driven Development:
спецификации написал **GitHub Copilot** через **Spec Kit** (ветка `spec`), код по `tasks.md` реализовал агент **Google Antigravity** (ветка `implementation`), тесты дополнены в ветке `tests`.

```bash
git clone https://github.com/DmitriiAnghis/currency-converter-sdd
cd currency-converter-sdd
python -m pip install httpx pytest
set PYTHONPATH=src           # PowerShell: $env:PYTHONPATH="src"
python -m converter          # запуск
python -m pytest             # тесты
```

### Лаб 4: мобильный конвертер (отдельный репозиторий)

Android-приложение на Kotlin + Jetpack Compose: конвертер по курсу BNM, графики за неделю и месяц, работа без сети из кэша.
Открыть папку [currency-mobile](https://github.com/DmitriiAnghis/currency-mobile) в Android Studio и запустить на эмуляторе, тесты логики: `gradlew :core:test`.

## Структура репозитория

```
PSI_A/
├── lab1/                  Лаб 1: ticket_generator.py, tests/, requirements.txt
├── lab2/                  Лаб 2: app.py, ticket_data.py, data/, tests/, docs/
├── screenshots/
│   ├── lab1/              скриншоты запуска и проверок Лаб 1
│   ├── lab2/              скриншоты Лаб 2
│   └── lab3/              скриншоты этапов Лаб 3 (Spec Kit, Antigravity, GUI, pytest)
└── .github/workflows/     уведомление на почту о новых коммитах (см. ниже)
```

## Работа с Git

- Каждая лаба делается в своей ветке (`lab1-ticket-generator`, `lab2-ui-tickets`, …) и попадает в `main` только через Pull Request.
- Сообщения коммитов описывают, что сделано на шаге (логика, тесты, README, исправления).
- История каждой лабы видна во вкладке **Pull requests → Closed**.

## Уведомления о коммитах на почту

GitHub не присылает письма о твоих собственных push. Поэтому в репозитории есть workflow
[`.github/workflows/notify-push.yml`](.github/workflows/notify-push.yml): при каждом push он отправляет письмо со списком новых коммитов.
Пока секреты не заданы, он ничего не делает. Чтобы включить, в **Settings → Secrets and variables → Actions → New repository secret** добавь:

| Секрет | Значение |
|---|---|
| `MAIL_USERNAME` | адрес отправителя, например Gmail-адрес |
| `MAIL_PASSWORD` | пароль приложения (для Gmail: Google Account → Security → App passwords) |
| `MAIL_TO` | куда слать, например `dmitrii.anghis1@isa.utm.md` |

По умолчанию используется SMTP Gmail (`smtp.gmail.com:465`). Для другой почты добавь секреты `MAIL_SERVER` и `MAIL_PORT`.
