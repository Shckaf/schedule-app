# -*- coding: utf-8 -*-
"""Тестовые данные расписания и общие константы.

Модуль без внешних зависимостей (кроме openpyxl в write_workbook), поэтому
его можно использовать и в Android-приложении (демо-данные), и в
generate_excel.py (pandas).
"""
import random

DAYS = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб"]
DAYS_FULL = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]

PAIR_TIMES = {
    1: ("8:30", "9:55"),
    2: ("10:05", "11:30"),
    3: ("11:40", "13:05"),
    4: ("13:15", "14:40"),
    5: ("15:05", "16:30"),
    6: ("16:40", "18:05"),
    7: ("18:15", "19:40"),
}

WEEK_OVER = "Над чертой"
WEEK_UNDER = "Под чертой"
WEEK_BOTH = "Каждая неделя"

COLUMNS = [
    "Курс", "Группа", "День недели", "Время / Номер пары", "Тип недели",
    "Предмет", "Преподаватель", "Аудитория", "Подгруппа",
]

FACULTIES = {
    "Факультет информатики": {
        "groups": {1: ["ИТ-11", "ИТ-12"], 2: ["ИТ-21", "ИТ-22"], 3: ["ИТ-31"]},
        "subjects": [
            ("Математический анализ", "Лекция"), ("Алгоритмы и структуры данных", "Лекция"),
            ("Программирование на Python", "ЛР"), ("Базы данных", "ЛР"),
            ("Дискретная математика", "ПЗ"), ("Операционные системы", "Лекция"),
            ("Компьютерные сети", "ЛР"), ("Английский язык", "ПЗ"),
            ("Теория вероятностей", "Семинар"), ("Физическая культура", "ПЗ"),
        ],
        "teachers": ["Иванов И. И.", "Петрова А. С.", "Сидоров К. В.", "Козлова М. Н.",
                     "Новиков Д. А.", "Морозова Е. П."],
        "buildings": [1, 2],
    },
    "Экономический факультет": {
        "groups": {1: ["ЭК-11", "ЭК-12"], 2: ["ЭК-21"], 4: ["ЭК-41", "ЭК-42"]},
        "subjects": [
            ("Микроэкономика", "Лекция"), ("Макроэкономика", "Лекция"),
            ("Бухгалтерский учёт", "ПЗ"), ("Статистика", "ЛР"),
            ("Финансовый менеджмент", "Семинар"), ("Маркетинг", "Лекция"),
            ("Эконометрика", "ЛР"), ("Иностранный язык", "ПЗ"),
            ("Правоведение", "Лекция"), ("Физическая культура", "ПЗ"),
        ],
        "teachers": ["Волков С. Р.", "Лебедева О. Г.", "Орлов П. Т.", "Гусева Н. Ю.",
                     "Фёдоров А. А."],
        "buildings": [3],
    },
    "Юридический факультет": {
        "groups": {1: ["ЮР-11"], 2: ["ЮР-21", "ЮР-22"], 3: ["ЮР-31"]},
        "subjects": [
            ("Теория государства и права", "Лекция"), ("Конституционное право", "Лекция"),
            ("Гражданское право", "Семинар"), ("Уголовное право", "Лекция"),
            ("Римское право", "ПЗ"), ("Латинский язык", "ПЗ"),
            ("История государства и права", "Семинар"), ("Философия", "Лекция"),
        ],
        "teachers": ["Белов Г. М.", "Тарасова И. К.", "Жуков В. Л.", "Соколова Т. Е."],
        "buildings": [4, 5],
    },
}


def build_data(seed=2026):
    """Вернуть {факультет: [строки-словари в формате Excel]}."""
    rng = random.Random(seed)
    result = {}
    for faculty, cfg in FACULTIES.items():
        rows = []
        for course, groups in cfg["groups"].items():
            for group in groups:
                for day in range(6):
                    count = rng.randint(2, 3) if day == 5 else rng.randint(2, 4)
                    slots = sorted(rng.sample(range(1, 6), count))
                    for pair in slots:
                        time_txt = "{} пара ({}–{})".format(pair, *PAIR_TIMES[pair])

                        def room():
                            return "{}-{}{:02d}".format(
                                rng.choice(cfg["buildings"]), rng.randint(1, 4),
                                rng.randint(1, 20))

                        def base(week, subj, sub=0):
                            name, kind = subj
                            return {
                                "Курс": course, "Группа": group,
                                "День недели": DAYS_FULL[day], "Время / Номер пары": time_txt,
                                "Тип недели": week, "Предмет": "{} ({})".format(name, kind),
                                "Преподаватель": rng.choice(cfg["teachers"]),
                                "Аудитория": room(), "Подгруппа": sub or "",
                            }

                        roll = rng.random()
                        subj = rng.choice(cfg["subjects"])
                        if roll < 0.28:      # разные предметы над / под чертой
                            other = rng.choice(cfg["subjects"])
                            rows.append(base(WEEK_OVER, subj))
                            rows.append(base(WEEK_UNDER, other))
                        elif subj[1] == "ЛР" and roll < 0.7:   # деление на подгруппы
                            week = rng.choice([WEEK_BOTH, WEEK_BOTH, WEEK_OVER, WEEK_UNDER])
                            rows.append(base(week, subj, 1))
                            rows.append(base(week, subj, 2))
                        else:
                            rows.append(base(WEEK_BOTH, subj))
        result[faculty] = rows
    return result


def write_workbook(path):
    """Записать тестовый xlsx только через openpyxl (используется в приложении)."""
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for faculty, rows in build_data().items():
        ws = wb.create_sheet(faculty[:31])
        ws.append(COLUMNS)
        for r in rows:
            ws.append([r[c] for c in COLUMNS])
    wb.save(path)
