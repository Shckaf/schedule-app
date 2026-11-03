# -*- coding: utf-8 -*-
"""Создаёт тестовый файл schedule_template.xlsx (pandas + openpyxl).

Запуск на компьютере:  python generate_excel.py
Каждый лист = факультет. Столбцы: Курс, Группа, День недели,
Время / Номер пары, Тип недели, Предмет, Преподаватель, Аудитория
(+ необязательный столбец «Подгруппа»: пусто = вся группа, 1 или 2).
Тип занятия указывается в скобках в конце названия предмета:
«Базы данных (ЛР)», «Математический анализ (Лекция)».
"""
import sys

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import sample_data

WIDTHS = [7, 11, 15, 24, 16, 46, 20, 12, 11]


def create_template(path="schedule_template.xlsx"):
    data = sample_data.build_data()
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for faculty, rows in data.items():
            df = pd.DataFrame(rows, columns=sample_data.COLUMNS)
            sheet = faculty[:31]  # лимит Excel на длину имени листа
            df.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            fill = PatternFill("solid", fgColor="1F4E9E")
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for i, w in enumerate(WIDTHS, start=1):
                ws.column_dimensions[get_column_letter(i)].width = w
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
    total = sum(len(r) for r in data.values())
    print("Готово: {} ({} листов, {} строк)".format(path, len(data), total))


if __name__ == "__main__":
    create_template(sys.argv[1] if len(sys.argv) > 1 else "schedule_template.xlsx")
