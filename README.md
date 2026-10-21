# PSI — лабораторные работы

**Студент:** Dmitrii Anghis, гр. ti-235  
**Преподаватель:** Denis Ahramenco  
**Университет:** Технический университет Молдовы, FCIM

В этом репозитории Лаб 1 и Лаб 2. Лаб 3 и Лаб 4 по условию сделаны в отдельных репозиториях.

| № | Тема | Где | Стек |
|---|---|---|---|
| 1 | Консольный генератор экзаменационных билетов | [`lab1/`](lab1/) | Python, openpyxl, pytest |
| 2 | Генератор билетов с графическим интерфейсом | [`lab2/`](lab2/) | Python, Tkinter, openpyxl, python-docx, pytest |
| 3 | Конвертер валют через Spec Kit и Google Antigravity | [currency-converter-sdd](https://github.com/DmitriiAnghis/currency-converter-sdd) | Python, Tkinter, httpx, pytest |
| 4 | Мобильный конвертер валют с аналитикой курса | [currency-mobile](https://github.com/DmitriiAnghis/currency-mobile) | Kotlin, Jetpack Compose, JUnit |

## Быстрый старт

Нужен Python 3.12 или новее. Каждая лаба запускается из своей папки:

```bash
git clone https://github.com/DmitriiAnghis/PSI_A.git
cd PSI_A/lab1          # или lab2
pip install -r requirements.txt
python ticket_generator.py   # Лаб 1
python app.py                # Лаб 2
python -m pytest             # тесты
```

## Лаб 1 — консольный генератор билетов

Программа спрашивает `Last name:` и `First name:`, выдаёт случайный билет от 1 до 20 («Билет № 7») и сразу дописывает строку в `journal.xlsx`. Клавиша ESC завершает работу.

- Файл сохраняется после каждого студента, поэтому записи не теряются даже при аварийном завершении.
- Если журнала нет, он создаётся с шапкой. Если есть, строки дописываются в конец, старые не меняются.
- Пустой ввод не принимается, пробелы по краям обрезаются.
- Если `journal.xlsx` открыт в Excel, программа не падает: просит закрыть файл и повторяет сохранение.

Тесты: 8 unit-тестов в [`lab1/tests`](lab1/tests).

## Лаб 2 — генератор билетов с GUI

Окно на Tkinter: выбор группы (листы `students.xlsx`), выбор студента, кнопка «Сгенерировать билет». Номер и три вопроса берутся из `tickets.docx`, результат сразу записывается в `results.xlsx`.

- Повторный выбор того же студента выдаёт его первый билет и помечает запись «Повтор = да».
- Билеты, в которых меньше трёх вопросов, пропускаются и пишутся в `tickets.log`.
- Если файл занят или отсутствует, программа показывает понятное сообщение.
- Esc закрывает окно с билетом, приложение остаётся открытым.

Тесты: 12 unit-тестов в [`lab2/tests`](lab2/tests). Подробности в [`lab2/README.md`](lab2/README.md).

## Структура

```text
PSI_A/
├── lab1/
│   ├── ticket_generator.py   # программа
│   ├── requirements.txt
│   └── tests/
├── lab2/
│   ├── app.py                # окно Tkinter
│   ├── ticket_data.py        # чтение Excel и Word, логика выдачи
│   ├── make_sample_data.py   # пересоздание тестовых данных
│   ├── data/                 # students.xlsx, tickets.docx
│   ├── docs/copilot_prompt.md
│   └── tests/
└── screenshots/              # скриншоты для отчётов
```

## Работа с Git

Каждая лаба делалась в своей ветке и попадала в `main` только через Pull Request:

| Ветка | Pull Request |
|---|---|
| `lab1-ticket-generator` | [#1](https://github.com/DmitriiAnghis/PSI_A/pull/1) |
| `lab2-ui-tickets` | [#2](https://github.com/DmitriiAnghis/PSI_A/pull/2) |
| `report-screenshots` | [#3](https://github.com/DmitriiAnghis/PSI_A/pull/3) |
