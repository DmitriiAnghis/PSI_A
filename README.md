<div align="center">

# 🎓 PSI: лабораторные работы

**Dmitrii Anghis** · группа TI-235 · преподаватель **Denis Ahramenco** · УТМ, 2026

![Python](https://img.shields.io/badge/Python-3.12+-3776AB?logo=python&logoColor=white)
![Kotlin](https://img.shields.io/badge/Kotlin-Android-7F52FF?logo=kotlin&logoColor=white)
![Tests](https://img.shields.io/badge/tests-227%20passed-2ea44f)
![Labs](https://img.shields.io/badge/labs-4%2F4-success)

</div>

| № | Тема | Где | Стек | Тесты |
|:-:|---|---|---|:-:|
| 1 | 🖥️ Консольный генератор экзаменационных билетов | [`lab1/`](lab1/) | Python, openpyxl | 23 ✅ |
| 2 | 🪟 Генератор билетов с графическим интерфейсом | [`lab2/`](lab2/) | Python, Tkinter, openpyxl, python-docx | 45 ✅ |
| 3 | 💱 Конвертер валют BNM (Spec-Driven Development) | [currency-converter-sdd](https://github.com/DmitriiAnghis/currency-converter-sdd) | Python, Tkinter, httpx · Spec Kit, Copilot, Antigravity | 126 ✅ |
| 4 | 📱 Мобильный конвертер валют BNM | [currency-mobile](https://github.com/DmitriiAnghis/currency-mobile) | Kotlin, Jetpack Compose | 33 ✅ |

**Содержание:** [Быстрый старт](#-быстрый-старт) · [Лаб 1](#-лаб-1-консольный-генератор-билетов) · [Лаб 2](#-лаб-2-генератор-билетов-с-gui) · [Лаб 3](#-лаб-3-конвертер-валют-bnm) · [Лаб 4](#-лаб-4-мобильный-конвертер) · [Структура](#-структура-репозитория) · [Git](#-работа-с-git) · [Уведомления](#-уведомления-о-коммитах-на-почту)

---

## 🚀 Быстрый старт

Нужен Python 3.12 или новее.

```bash
git clone https://github.com/DmitriiAnghis/PSI_A
cd PSI_A
python -m pip install -r lab1/requirements.txt -r lab2/requirements.txt
```

---

## 🖥️ Лаб 1: консольный генератор билетов

```bash
cd lab1
python ticket_generator.py   # запуск
python -m pytest -v          # тесты
```

- Запрашивает **Last name** и **First name**, пустой ввод не принимает.
- Выдаёт случайный билет **от 1 до 20** и сразу дописывает строку в `journal.xlsx`: фамилия, имя, номер, дата и время.
- **ESC** завершает программу.
- Если `journal.xlsx` открыт в Excel, программа не падает: просит закрыть файл и нажать Enter, потом повторяет сохранение.
- Записи сохраняются сразу, поэтому даже после `taskkill` введённые данные остаются в журнале.

<table>
<tr>
<td width="50%"><b>Работа программы и пустой ввод</b><br><img src="screenshots/lab1/lab1_empty_input_and_entries.png" alt="Ввод студентов"></td>
<td width="50%"><b>Выход по ESC</b><br><img src="screenshots/lab1/lab1_esc_exit.png" alt="Выход по ESC"></td>
</tr>
<tr>
<td><b>Файл открыт в Excel: программа не падает</b><br><img src="screenshots/lab1/lab1_locked_file_retry_saved.png" alt="Файл занят"></td>
<td><b>Процесс убит через taskkill</b><br><img src="screenshots/lab1/lab1_taskkill_console.png" alt="taskkill"></td>
</tr>
</table>

<details>
<summary>📊 Результат в <code>journal.xlsx</code></summary>

![journal.xlsx](screenshots/lab1/lab1_excel_journal.png)

</details>

**Тесты:**

![pytest lab1](screenshots/readme/lab1_pytest.svg)

Код: [`lab1/ticket_generator.py`](lab1/ticket_generator.py) · тесты: [`lab1/tests/`](lab1/tests/)

---

## 🪟 Лаб 2: генератор билетов с GUI

```bash
cd lab2
python app.py                # запуск
python -m pytest -v          # тесты
python make_sample_data.py   # пересоздать тестовые students.xlsx и tickets.docx
```

Всё приложение работает в **одном окне с вкладками**, без всплывающих окон:

| Вкладка | Что делает |
|---|---|
| **Билет** | Выбор группы и студента, «Сгенерировать билет». Номер и три вопроса из `data/tickets.docx` показываются прямо на странице, строка сразу пишется в `data/results.xlsx`. Повторный запрос выдаёт **тот же билет** и пишет «Повтор = да». **Esc** очищает карточку билета. |
| **Студенты** | Список групп и студентов, поиск, **добавление группы**, **добавление студента** (с заглавной буквы, без цифр, без повторов), удаление, «Выдать билет» по двойному щелчку. |
| **Журнал** | Таблица `results.xlsx` с фильтром по группе, сводка: выдано, студентов, повторов, самый частый билет. Кнопка «Открыть в Excel». |
| **Билеты** | Все билеты и вопросы из `tickets.docx`. |

- Если `results.xlsx` или `students.xlsx` открыт в Excel, программа сама сохраняет и закрывает эту книгу и записывает данные. Если не получилось, под шапкой появляется полоса «Файл занят» с кнопкой «Повторить».
- Сообщения об успехе и ошибках показываются полосой под шапкой, а не отдельными окнами.

<table>
<tr>
<td width="50%"><b>Билет</b><br><img src="screenshots/lab2/redesign/lab2_tab_ticket_repeat.png" alt="Билет"></td>
<td width="50%"><b>Студенты и группы</b><br><img src="screenshots/lab2/redesign/lab2_tab_students_added.png" alt="Студенты"></td>
</tr>
<tr>
<td><b>Проверка ввода</b><br><img src="screenshots/lab2/redesign/lab2_tab_students_errors.png" alt="Ошибки ввода"></td>
<td><b>Журнал</b><br><img src="screenshots/lab2/redesign/lab2_tab_journal.png" alt="Журнал"></td>
</tr>
</table>

<details>
<summary>📊 Результат в <code>results.xlsx</code></summary>

![results.xlsx](screenshots/lab2/lab2_results_excel.png)

</details>

**Тесты:**

![pytest lab2](screenshots/readme/lab2_pytest.svg)

Код: [`lab2/app.py`](lab2/app.py), [`lab2/ticket_data.py`](lab2/ticket_data.py) · тесты: [`lab2/tests/`](lab2/tests/)

---

## 💱 Лаб 3: конвертер валют BNM

> Отдельный репозиторий: **[currency-converter-sdd](https://github.com/DmitriiAnghis/currency-converter-sdd)**

Desktop-конвертер по официальному курсу Национального банка Молдовы, сделан по методике Spec-Driven Development:

1. **Spec Kit + GitHub Copilot** написали `constitution`, `spec`, `plan`, `tasks` (ветка `spec`).
2. **Google Antigravity** реализовал все 27 задач из `tasks.md` (ветка `implementation`).
3. Тесты на граничные случаи дописаны в ветке `tests`: **126 passed** на Python 3.12 и 3.14.

```bash
git clone https://github.com/DmitriiAnghis/currency-converter-sdd
cd currency-converter-sdd
python -m pip install -e ".[test]"
python -m converter          # запуск
python -m pytest             # тесты
```

<table>
<tr>
<td align="center"><b>EUR → MDL</b><br><img src="screenshots/lab3/lab3_11_gui_eur_mdl.png" width="230" alt="EUR MDL"></td>
<td align="center"><b>Кросс-курс USD → EUR</b><br><img src="screenshots/lab3/lab3_12_gui_cross_usd_eur.png" width="230" alt="USD EUR"></td>
<td align="center"><b>Без сети, из кэша</b><br><img src="screenshots/lab3/lab3_18_gui_offline_cache_result.png" width="230" alt="Кэш"></td>
</tr>
</table>

![pytest lab3](screenshots/lab3/lab3_25_pytest_tests.png)

<details>
<summary>🤖 Этапы Spec Kit и Antigravity</summary>

| `/speckit-specify` | `/speckit-analyze` |
|---|---|
| ![specify](screenshots/lab3/lab3_02_specify.png) | ![analyze](screenshots/lab3/lab3_05_analyze.png) |
| **Antigravity: старт** | **Antigravity: 27/27 задач** |
| ![start](screenshots/lab3/lab3_10_antigravity_start.png) | ![done](screenshots/lab3/lab3_20_antigravity_done.png) |

</details>

---

## 📱 Лаб 4: мобильный конвертер

> Отдельный репозиторий: **[currency-mobile](https://github.com/DmitriiAnghis/currency-mobile)**

Android-приложение на Kotlin + Jetpack Compose: конвертер по курсу BNM, графики курса за неделю и месяц (минимум, максимум, среднее, тренд), работа без сети из кэша.

Открыть проект в Android Studio и запустить на эмуляторе. Тесты логики: `gradlew :core:test`.

| Конвертер | Без сети |
|---|---|
| ![conv](screenshots/lab4/lab4_conv.png) | ![offline](screenshots/lab4/lab4_offline.png) |

<details>
<summary>📈 Графики и тесты</summary>

![stats](screenshots/lab4/lab4_stats.png)

![core test](screenshots/lab4/lab4_core_test.png)

</details>

---

## 📁 Структура репозитория

```text
PSI_A/
├── lab1/                  Лаб 1: ticket_generator.py, tests/, requirements.txt
├── lab2/                  Лаб 2: app.py, ticket_data.py, data/, tests/, docs/
├── screenshots/
│   ├── lab1/ … lab4/      скриншоты запуска и проверок
│   └── readme/            вывод pytest для этого README
└── .github/workflows/     уведомление на почту о новых коммитах
```

---

## 🌿 Работа с Git

- Каждая лаба делается в своей ветке (`lab1-ticket-generator`, `lab2-ui-tickets`, …) и попадает в `main` только через Pull Request.
- Сообщения коммитов описывают, что сделано на шаге: логика, тесты, README, исправления.
- История каждой лабы видна во вкладке **Pull requests → Closed**.

![git log](screenshots/lab1/lab1_git_log.png)

---

## 📬 Уведомления о коммитах на почту

GitHub не присылает письма о твоих собственных push. Поэтому в репозитории есть workflow
[`.github/workflows/notify-push.yml`](.github/workflows/notify-push.yml): при каждом push он отправляет письмо со списком новых коммитов.
Пока секреты не заданы, он ничего не делает. Чтобы включить, в **Settings → Secrets and variables → Actions → New repository secret** добавь:

| Секрет | Значение |
|---|---|
| `MAIL_USERNAME` | адрес отправителя, например Gmail-адрес |
| `MAIL_PASSWORD` | пароль приложения (для Gmail: Google Account → Security → App passwords) |
| `MAIL_TO` | куда слать, например `dmitrii.anghis1@isa.utm.md` |

По умолчанию используется SMTP Gmail (`smtp.gmail.com:465`). Для другой почты добавь секреты `MAIL_SERVER` и `MAIL_PORT`.
