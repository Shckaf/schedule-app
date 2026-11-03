# -*- coding: utf-8 -*-
"""Расписание — Kivy-приложение (Android + десктоп) с импортом Excel и
виджетами рабочего стола.

Вкладки: «Расписание», «Экспорт» (импорт Excel), «Справка».
Данные хранятся в SQLite; для виджета Android дублируются в widget_data.json
(его читает Java-класс ScheduleWidgetProvider из папки java/).
"""
import datetime as dt
import json
import os
import re
import sqlite3
from contextlib import closing

import openpyxl
from kivy.app import App
from kivy.clock import Clock
from kivy.core.text import Label as CoreLabel
from kivy.core.window import Window
from kivy.graphics import Color, Ellipse, Line, Rectangle, Triangle
from kivy.lang import Builder
from kivy.metrics import dp, sp
from kivy.properties import (BooleanProperty, ColorProperty, NumericProperty,
                             StringProperty)
from kivy.uix.behaviors import ButtonBehavior
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.widget import Widget
from kivy.utils import escape_markup
from kivy.utils import get_color_from_hex as hx
from kivy.utils import platform

import sample_data
from sample_data import DAYS, DAYS_FULL, PAIR_TIMES

# ----------------------------------------------------------------- тема ----
BG = hx("#0d1117")
SURFACE = hx("#12161f")
CARD = hx("#1b2232")
ACCENT = hx("#3b82f6")
TEXT = hx("#e6edf3")
MUTED = hx("#8b949e")
LINE = hx("#2a3142")
CHIP_BG = hx("#1e2c47")
CHIP_FG = hx("#8ab4ff")

WEEK_NAMES = {"over": "над чертой", "under": "под чертой"}
REQ_PICK = 4711

# ------------------------------------------------- логика недель / дат ----


def monday_of(d):
    return d - dt.timedelta(days=d.weekday())


def week_type_for(date, anchor):
    """'over' (над чертой) или 'under'. anchor — понедельник недели «над чертой»."""
    idx = (monday_of(date) - anchor).days // 7
    return "over" if idx % 2 == 0 else "under"


def default_anchor():
    today = dt.date.today()
    year = today.year if today.month >= 8 else today.year - 1
    return monday_of(dt.date(year, 9, 1))


def default_view(anchor):
    """(день 0..5, тип недели) для показа «сейчас»; в воскресенье — понедельник."""
    today = dt.date.today()
    if today.weekday() == 6:
        ref = today + dt.timedelta(days=1)
        return 0, week_type_for(ref, anchor)
    return today.weekday(), week_type_for(today, anchor)


# ------------------------------------------------------ разбор Excel -------
REQUIRED = {
    "course": "Курс", "group": "Группа", "day": "День недели",
    "pair": "Время / Номер пары", "week": "Тип недели", "subject": "Предмет",
    "teacher": "Преподаватель", "room": "Аудитория",
}
HEADER_ALIASES = {
    "course": {"курс"},
    "group": {"группа"},
    "day": {"день недели", "день"},
    "pair": {"время/номер пары", "номер пары", "пара", "время"},
    "week": {"тип недели", "неделя"},
    "subject": {"предмет", "дисциплина"},
    "teacher": {"преподаватель"},
    "room": {"аудитория", "ауд.", "ауд"},
    "subgroup": {"подгруппа", "подгр.", "подгр"},
}
DAY_ALIASES = {
    0: {"пн", "пон", "понедельник", "1"}, 1: {"вт", "вто", "вторник", "2"},
    2: {"ср", "сре", "среда", "3"}, 3: {"чт", "чет", "четверг", "4"},
    4: {"пт", "пят", "пятница", "5"}, 5: {"сб", "суб", "суббота", "6"},
}
TYPE_MAP = {
    "лекция": "Лекция", "лек": "Лекция", "лр": "ЛР", "лаб": "ЛР",
    "лабораторная": "ЛР", "пз": "ПЗ", "практика": "ПЗ", "практическое": "ПЗ",
    "семинар": "Семинар",
}
TIME_RE = re.compile(r"(\d{1,2})[:.](\d{2})\s*[-–—]\s*(\d{1,2})[:.](\d{2})")
LEAD_NUM_RE = re.compile(r"\s*(\d{1,2})(?!\d)(?![:.]\d)")
SUBJECT_RE = re.compile(r"^(.*?)\s*[(\[]\s*([^()\[\]]+?)\s*[)\]]\s*$")


def _norm_header(h):
    s = re.sub(r"\s+", " ", str(h or "").strip().lower())
    return re.sub(r"\s*/\s*", "/", s)


def _map_headers(header):
    idx = {}
    for i, h in enumerate(header):
        n = _norm_header(h)
        for key, names in HEADER_ALIASES.items():
            if n in names and key not in idx:
                idx[key] = i
    return idx


def _cell(row, i):
    if i is None or i >= len(row) or row[i] is None:
        return ""
    return str(row[i]).strip()


def _parse_row(row, idx, faculty):
    course_s = _cell(row, idx["course"])
    m = re.match(r"\d+", course_s)
    if not m:
        raise ValueError("не указан курс")
    course = int(m.group())
    group = _cell(row, idx["group"])
    if not group:
        raise ValueError("не указана группа")

    day_s = _cell(row, idx["day"]).lower().rstrip(".")
    day = next((k for k, v in DAY_ALIASES.items() if day_s in v), None)
    if day is None:
        raise ValueError("неизвестный день недели «{}»".format(day_s))

    pair_s = _cell(row, idx["pair"])
    tm = TIME_RE.search(pair_s)
    start = end = None
    if tm:
        start = "{}:{}".format(int(tm.group(1)), tm.group(2))
        end = "{}:{}".format(int(tm.group(3)), tm.group(4))
    nm = LEAD_NUM_RE.match(pair_s)
    pair = int(nm.group(1)) if nm else None
    if pair is None and start:
        pair = next((p for p, (s, _e) in PAIR_TIMES.items() if s == start), 0)
    if pair is None:
        raise ValueError("не указан номер пары или время «{}»".format(pair_s))
    if not start:
        if pair not in PAIR_TIMES:
            raise ValueError("неизвестный номер пары {}".format(pair))
        start, end = PAIR_TIMES[pair]

    week_s = _cell(row, idx["week"]).lower()
    if "над" in week_s:
        week = "over"
    elif "под" in week_s:
        week = "under"
    elif week_s == "" or "кажд" in week_s or "обе" in week_s or "все" in week_s:
        week = "both"
    else:
        raise ValueError("неизвестный тип недели «{}»".format(week_s))

    subject = _cell(row, idx["subject"])
    if not subject:
        raise ValueError("не указан предмет")
    ltype = ""
    sm = SUBJECT_RE.match(subject)
    if sm and sm.group(2).lower().rstrip(".") in TYPE_MAP:
        subject = sm.group(1)
        ltype = TYPE_MAP[sm.group(2).lower().rstrip(".")]

    sub_s = _cell(row, idx.get("subgroup"))
    if sub_s.lower() in ("", "0", "все", "вся", "-"):
        subgroup = 0
    else:
        sg = re.match(r"\d+", sub_s)
        if not sg or int(sg.group()) not in (1, 2):
            raise ValueError("подгруппа должна быть 1 или 2 (или пусто)")
        subgroup = int(sg.group())

    return {
        "faculty": faculty, "course": course, "grp": group, "day": day,
        "pair": pair, "start": start, "end": end, "week": week,
        "subject": subject, "ltype": ltype,
        "teacher": _cell(row, idx["teacher"]), "room": _cell(row, idx["room"]),
        "subgroup": subgroup,
    }


def parse_workbook(path):
    """Прочитать xlsx > (строки, ошибки, предупреждения)."""
    rows, errors, warnings = [], [], []
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            it = ws.iter_rows(values_only=True)
            header = next(it, None)
            if not header or all(h is None for h in header):
                warnings.append("Лист «{}» пуст — пропущен".format(ws.title))
                continue
            idx = _map_headers(header)
            missing = [name for key, name in REQUIRED.items() if key not in idx]
            if missing:
                warnings.append("Лист «{}»: нет столбцов {} — пропущен".format(
                    ws.title, ", ".join(missing)))
                continue
            for n, row in enumerate(it, start=2):
                if row is None or all(c is None or str(c).strip() == "" for c in row):
                    continue
                try:
                    rows.append(_parse_row(row, idx, ws.title.strip()))
                except ValueError as exc:
                    errors.append("«{}», строка {}: {}".format(ws.title, n, exc))
    finally:
        wb.close()
    return rows, errors, warnings


# ------------------------------------------------------ база данных -------
class Database:
    def __init__(self, path):
        self.path = path
        with closing(self._conn()) as c, c:
            c.execute(
                "CREATE TABLE IF NOT EXISTS lessons (faculty TEXT, course INTEGER,"
                " grp TEXT, day INTEGER, pair INTEGER, start TEXT, end TEXT,"
                " week TEXT, subject TEXT, ltype TEXT, teacher TEXT, room TEXT,"
                " subgroup INTEGER)")

    def _conn(self):
        return sqlite3.connect(self.path)

    def replace_all(self, rows):
        with closing(self._conn()) as c, c:
            c.execute("DELETE FROM lessons")
            c.executemany(
                "INSERT INTO lessons VALUES (:faculty,:course,:grp,:day,:pair,:start,"
                ":end,:week,:subject,:ltype,:teacher,:room,:subgroup)", rows)

    def _col(self, sql, args=()):
        with closing(self._conn()) as c:
            return [r[0] for r in c.execute(sql, args)]

    def count(self):
        return self._col("SELECT COUNT(*) FROM lessons")[0]

    def faculties(self):
        return self._col("SELECT DISTINCT faculty FROM lessons ORDER BY faculty")

    def courses(self, faculty):
        return self._col("SELECT DISTINCT course FROM lessons WHERE faculty=?"
                         " ORDER BY course", (faculty,))

    def groups(self, faculty, course):
        return self._col("SELECT DISTINCT grp FROM lessons WHERE faculty=? AND"
                         " course=? ORDER BY grp", (faculty, course))

    def lessons(self, faculty, course, group, day, week, subgroup):
        sql = ("SELECT pair,start,end,week,subject,ltype,teacher,room,subgroup FROM"
               " lessons WHERE faculty=? AND course=? AND grp=? AND day=? AND"
               " (week='both' OR week=?) AND (?=0 OR subgroup=0 OR subgroup=?)"
               " ORDER BY pair, subgroup")
        keys = ("pair", "start", "end", "week", "subject", "ltype", "teacher",
                "room", "subgroup")
        with closing(self._conn()) as c:
            cur = c.execute(sql, (faculty, course, group, day, week, subgroup, subgroup))
            return [dict(zip(keys, r)) for r in cur]

    def widget_rows(self):
        keys = ("f", "c", "g", "d", "p", "s", "e", "w", "sj", "t", "te", "r", "sg")
        sql = ("SELECT faculty,course,grp,day,pair,start,end,week,subject,ltype,"
               "teacher,room,subgroup FROM lessons")
        with closing(self._conn()) as c:
            return [dict(zip(keys, r)) for r in c.execute(sql)]


# ------------------------------------------------------- виджеты UI -------
class Rounded(BoxLayout):
    bg = ColorProperty(CARD)
    radius = NumericProperty(dp(14))


class FlatButton(ButtonBehavior, Label):
    bg = ColorProperty(CARD)
    bg_down = ColorProperty(LINE)
    radius = NumericProperty(dp(12))


class SelButton(FlatButton):
    selected = BooleanProperty(False)


class WrapLabel(Label):
    pass


class Icon(Widget):
    kind = StringProperty("")
    color = ColorProperty(MUTED)

    def __init__(self, **kw):
        super().__init__(**kw)
        for p in ("pos", "size", "kind", "color"):
            self.bind(**{p: self.redraw})
        self.redraw()

    def redraw(self, *_):
        self.canvas.clear()
        if not self.kind or self.width <= 0:
            return
        s = min(self.width, self.height)
        cx, cy = self.center
        lw = max(dp(1.3), s * 0.07)
        with self.canvas:
            Color(*self.color)
            if self.kind == "room":
                Line(rounded_rectangle=(cx - s * .30, cy - s * .44, s * .60, s * .88, s * .08),
                     width=lw)
                Ellipse(pos=(cx + s * .08, cy - s * .08), size=(s * .14, s * .14))
            elif self.kind == "calendar":
                Line(rounded_rectangle=(cx - s * .44, cy - s * .42, s * .88, s * .76, s * .10),
                     width=lw)
                Line(points=[cx - s * .44, cy + s * .10, cx + s * .44, cy + s * .10], width=lw)
                Line(points=[cx - s * .20, cy + s * .26, cx - s * .20, cy + s * .50], width=lw)
                Line(points=[cx + s * .20, cy + s * .26, cx + s * .20, cy + s * .50], width=lw)
                for ox in (-.22, 0, .22):
                    Ellipse(pos=(cx + s * ox - s * .04, cy - s * .22), size=(s * .08, s * .08))
            elif self.kind == "refresh":
                r = s * .34
                Line(circle=(cx, cy, r, 30, 320), width=lw)
                import math
                th = math.radians(320)
                ex, ey = cx + r * math.sin(th), cy + r * math.cos(th)
                tx, ty = math.cos(th), -math.sin(th)      # касательная
                nx, ny = math.sin(th), math.cos(th)       # радиальная
                Triangle(points=[ex + tx * s * .22, ey + ty * s * .22,
                                 ex + nx * s * .16, ey + ny * s * .16,
                                 ex - nx * s * .16, ey - ny * s * .16])
            elif self.kind == "help":
                Line(circle=(cx, cy, s * .44), width=lw)
                lab = CoreLabel(text="?", font_size=s * .62, bold=True)
                lab.refresh()
                t = lab.texture
                Rectangle(texture=t, size=t.size, pos=(cx - t.width / 2, cy - t.height / 2))


class Chip(Rounded):
    text = StringProperty("")
    icon = StringProperty("")
    fg = ColorProperty(CHIP_FG)


class NavItem(ButtonBehavior, BoxLayout):
    kind = StringProperty("")
    text = StringProperty("")
    active = BooleanProperty(False)


class LessonCard(Rounded):
    pair_text = StringProperty("")
    start = StringProperty("")
    end = StringProperty("")

    def __init__(self, lesson, **kw):
        pair = lesson["pair"]
        super().__init__(pair_text=("{} пара".format(pair) if pair else "пара"),
                         start=lesson["start"], end=lesson["end"], **kw)
        info = self.ids.info
        info.add_widget(WrapLabel(text=lesson["subject"], bold=True, color=TEXT,
                                  font_size=sp(15)))
        chips = []
        if lesson["ltype"]:
            chips.append(Chip(text=lesson["ltype"]))
        if lesson["week"] != "both":
            chips.append(Chip(text=WEEK_NAMES[lesson["week"]]))
        if lesson["subgroup"]:
            chips.append(Chip(text="{} подгр.".format(lesson["subgroup"])))
        if chips:
            row = BoxLayout(size_hint_y=None, height=dp(24), spacing=dp(6))
            for c in chips:
                row.add_widget(c)
            row.add_widget(Widget())
            info.add_widget(row)
        if lesson["teacher"]:
            info.add_widget(WrapLabel(text=lesson["teacher"], color=MUTED, font_size=sp(13)))
        if lesson["room"]:
            row = BoxLayout(size_hint_y=None, height=dp(24))
            row.add_widget(Chip(text="ауд. " + lesson["room"], icon="room"))
            row.add_widget(Widget())
            info.add_widget(row)


class Root(BoxLayout):
    pass


KV = """
#:import dp kivy.metrics.dp
#:import sp kivy.metrics.sp
#:import hx kivy.utils.get_color_from_hex
#:import NoTransition kivy.uix.screenmanager.NoTransition
#:set BG hx('#0d1117')
#:set SURFACE hx('#12161f')
#:set CARD hx('#1b2232')
#:set ACCENT hx('#3b82f6')
#:set ACCENT_DOWN hx('#2563eb')
#:set TEXT hx('#e6edf3')
#:set MUTED hx('#8b949e')
#:set LINE hx('#2a3142')

<Rounded>:
    canvas.before:
        Color:
            rgba: self.bg
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [self.radius]

<FlatButton>:
    color: TEXT
    font_size: sp(14)
    halign: 'center'
    valign: 'middle'
    text_size: self.width - dp(10), self.height
    shorten: True
    canvas.before:
        Color:
            rgba: self.bg_down if self.state == 'down' else self.bg
        RoundedRectangle:
            pos: self.pos
            size: self.size
            radius: [self.radius]

<SelButton>:
    bg: ACCENT if self.selected else SURFACE
    bg_down: ACCENT_DOWN
    bold: self.selected
    color: (1, 1, 1, 1) if self.selected else MUTED

<WrapLabel>:
    size_hint_y: None
    text_size: self.width, None
    height: self.texture_size[1]
    halign: 'left'
    valign: 'top'
    color: TEXT
    font_size: sp(14)
    markup: True

<Chip>:
    size_hint: None, None
    height: dp(24)
    width: lbl.texture_size[0] + dp(16) + ic.width + (dp(4) if root.icon else 0)
    padding: dp(8), 0
    spacing: dp(4) if root.icon else 0
    radius: dp(12)
    bg: hx('#1e2c47')
    Icon:
        id: ic
        kind: root.icon
        color: root.fg
        size_hint: None, 1
        width: dp(13) if root.icon else 0
    Label:
        id: lbl
        text: root.text
        color: root.fg
        font_size: sp(12)
        size_hint_x: None
        width: self.texture_size[0]

<NavItem>:
    orientation: 'vertical'
    padding: 0, dp(8), 0, dp(6)
    spacing: dp(3)
    Icon:
        kind: root.kind
        color: ACCENT if root.active else MUTED
        size_hint: 1, None
        height: dp(24)
    Label:
        text: root.text
        font_size: sp(11)
        bold: root.active
        color: ACCENT if root.active else MUTED
        size_hint_y: None
        height: dp(16)

<LessonCard>:
    size_hint_y: None
    height: max(dp(96), info.height + dp(24))
    padding: 0
    spacing: 0
    BoxLayout:
        orientation: 'vertical'
        size_hint_x: None
        width: dp(84)
        padding: dp(4)
        canvas.before:
            Color:
                rgba: ACCENT
            RoundedRectangle:
                pos: self.pos
                size: self.size
                radius: [dp(14), 0, 0, dp(14)]
        Label:
            text: root.pair_text
            bold: True
            font_size: sp(15)
        Label:
            text: root.start
            font_size: sp(14)
        Label:
            text: root.end
            font_size: sp(14)
            color: 1, 1, 1, .8
    BoxLayout:
        id: info
        orientation: 'vertical'
        size_hint_y: None
        height: self.minimum_height
        pos_hint: {'center_y': .5}
        padding: dp(12), dp(8)
        spacing: dp(6)

<Root>:
    orientation: 'vertical'
    canvas.before:
        Color:
            rgba: BG
        Rectangle:
            pos: self.pos
            size: self.size
    ScreenManager:
        id: sm
        transition: NoTransition()
        Screen:
            name: 'schedule'
            BoxLayout:
                orientation: 'vertical'
                BoxLayout:
                    orientation: 'vertical'
                    size_hint_y: None
                    height: self.minimum_height
                    padding: dp(10)
                    spacing: dp(8)
                    canvas.before:
                        Color:
                            rgba: SURFACE
                        Rectangle:
                            pos: self.pos
                            size: self.size
                    BoxLayout:
                        size_hint_y: None
                        height: dp(38)
                        spacing: dp(6)
                        SelButton:
                            text: app.week_over_text
                            selected: app.view_week == 'over'
                            on_release: app.set_view_week('over')
                        SelButton:
                            text: app.week_under_text
                            selected: app.view_week == 'under'
                            on_release: app.set_view_week('under')
                        FlatButton:
                            text: 'Виджет'
                            bg: ACCENT
                            bg_down: ACCENT_DOWN
                            bold: True
                            size_hint_x: None
                            width: dp(84)
                            on_release: app.add_widget_to_desktop()
                    FlatButton:
                        text: app.fac_text
                        size_hint_y: None
                        height: dp(42)
                        on_release: app.pick('faculty')
                    BoxLayout:
                        size_hint_y: None
                        height: dp(42)
                        spacing: dp(8)
                        FlatButton:
                            text: app.course_text
                            size_hint_x: .26
                            on_release: app.pick('course')
                        FlatButton:
                            text: app.group_text
                            size_hint_x: .36
                            on_release: app.pick('group')
                        FlatButton:
                            text: app.sub_text
                            size_hint_x: .38
                            on_release: app.pick('subgroup')
                BoxLayout:
                    ScrollView:
                        bar_width: dp(3)
                        GridLayout:
                            id: cards
                            cols: 1
                            size_hint_y: None
                            height: self.minimum_height
                            padding: dp(12), dp(12), dp(6), dp(12)
                            spacing: dp(10)
                    BoxLayout:
                        id: days
                        orientation: 'vertical'
                        size_hint_x: None
                        width: dp(52)
                        padding: dp(4), dp(10), dp(8), dp(10)
                        spacing: dp(6)
        Screen:
            name: 'export'
            ScrollView:
                BoxLayout:
                    orientation: 'vertical'
                    size_hint_y: None
                    height: self.minimum_height
                    padding: dp(16)
                    spacing: dp(12)
                    WrapLabel:
                        text: '[b]Импорт расписания[/b]'
                        font_size: sp(22)
                    WrapLabel:
                        text: 'Загрузите Excel-файл (.xlsx): листы — факультеты, столбцы — Курс, Группа, День недели, Время / Номер пары, Тип недели, Предмет, Преподаватель, Аудитория (по желанию — Подгруппа). Текущее расписание будет заменено.'
                        color: MUTED
                    FlatButton:
                        text: 'Выбрать файл Excel (.xlsx)'
                        bg: ACCENT
                        bg_down: ACCENT_DOWN
                        bold: True
                        size_hint_y: None
                        height: dp(52)
                        on_release: app.choose_file()
                    FlatButton:
                        text: 'Загрузить демо-расписание'
                        size_hint_y: None
                        height: dp(46)
                        on_release: app.load_demo()
                    Rounded:
                        orientation: 'vertical'
                        size_hint_y: None
                        height: self.minimum_height
                        padding: dp(14)
                        spacing: dp(6)
                        WrapLabel:
                            text: app.stats_text
                            color: MUTED
                        WrapLabel:
                            text: app.status_text
                    Rounded:
                        orientation: 'vertical'
                        size_hint_y: None
                        height: self.minimum_height
                        padding: dp(14)
                        spacing: dp(10)
                        WrapLabel:
                            text: '[b]Какая неделя идёт сейчас?[/b]'
                        BoxLayout:
                            size_hint_y: None
                            height: dp(40)
                            spacing: dp(8)
                            SelButton:
                                text: 'Над чертой'
                                selected: app.cur_week == 'over'
                                on_release: app.set_current_week('over')
                            SelButton:
                                text: 'Под чертой'
                                selected: app.cur_week == 'under'
                                on_release: app.set_current_week('under')
        Screen:
            name: 'help'
            ScrollView:
                BoxLayout:
                    orientation: 'vertical'
                    size_hint_y: None
                    height: self.minimum_height
                    padding: dp(16)
                    spacing: dp(14)
                    WrapLabel:
                        text: app.help_text
                    FlatButton:
                        text: 'Добавить виджет на рабочий стол'
                        bg: ACCENT
                        bg_down: ACCENT_DOWN
                        bold: True
                        size_hint_y: None
                        height: dp(52)
                        on_release: app.add_widget_to_desktop()
    BoxLayout:
        size_hint_y: None
        height: dp(62)
        canvas.before:
            Color:
                rgba: SURFACE
            Rectangle:
                pos: self.pos
                size: self.size
            Color:
                rgba: LINE
            Rectangle:
                pos: self.x, self.top - dp(1)
                size: self.width, dp(1)
        NavItem:
            kind: 'calendar'
            text: 'Расписание'
            active: app.tab == 'schedule'
            on_release: app.set_tab('schedule')
        NavItem:
            kind: 'refresh'
            text: 'Экспорт'
            active: app.tab == 'export'
            on_release: app.set_tab('export')
        NavItem:
            kind: 'help'
            text: 'Справка'
            active: app.tab == 'help'
            on_release: app.set_tab('help')

Root:
"""

HELP_TEXT = """[b][size=22sp]Справка[/size][/b]

[b]Как это работает[/b]
Приложение показывает учебное расписание выбранной группы. Выберите факультет, курс, группу и (если нужно) подгруппу сверху, а день недели — справа. Кнопки «Над чертой» / «Под чертой» переключают тип недели; точка • отмечает текущую неделю.

[b]Импорт Excel[/b]
Откройте вкладку «Экспорт» > «Выбрать файл Excel». Каждый лист файла — факультет. Столбцы: Курс, Группа, День недели, Время / Номер пары, Тип недели («Над чертой», «Под чертой», «Каждая неделя»), Предмет, Преподаватель, Аудитория. Необязательный столбец «Подгруппа»: пусто — вся группа, 1 или 2. Тип занятия пишите в скобках в конце предмета: «Базы данных (ЛР)». Шаблон: schedule_template.xlsx (создаётся скриптом generate_excel.py) или кнопка «Загрузить демо-расписание».

[b]Какая неделя сейчас[/b]
Приложение считает чередование недель само, но если оно ошиблось — на вкладке «Экспорт» нажмите нужный тип в блоке «Какая неделя идёт сейчас?».

[b]Виджет на рабочий стол (Android)[/b]
1. Нажмите кнопку «Виджет» на главном экране или кнопку ниже — система предложит добавить виджет. Либо: долгое нажатие на пустом месте рабочего стола > «Виджеты» > «Расписание».
2. В окне настройки выберите группу (факультет · курс), подгруппу и день недели: «Автоматически» показывает сегодняшний день, а в воскресенье — понедельник.
3. Стрелки «<» и «>» на виджете листают дни, нажатие на заголовок открывает приложение.
4. Можно добавить несколько виджетов с разными группами. После нового импорта виджеты обновляются сами.

Если виджет пишет «Нет данных» — сначала импортируйте расписание в приложении."""


# ----------------------------------------------------------- приложение ---
class ScheduleApp(App):
    tab = StringProperty("schedule")
    fac_text = StringProperty("Факультет")
    course_text = StringProperty("Курс")
    group_text = StringProperty("Группа")
    sub_text = StringProperty("Все подгруппы")
    view_day = NumericProperty(0)
    view_week = StringProperty("over")
    cur_week = StringProperty("over")
    week_over_text = StringProperty("Над чертой")
    week_under_text = StringProperty("Под чертой")
    stats_text = StringProperty("")
    status_text = StringProperty("")
    help_text = StringProperty(HELP_TEXT)

    # --- служебное -----------------------------------------------------
    def files_dir(self):
        if platform == "android":
            from jnius import autoclass
            act = autoclass("org.kivy.android.PythonActivity").mActivity
            return act.getFilesDir().getAbsolutePath()
        return self.user_data_dir

    def bridge(self):
        """Java-класс ScheduleBridge и Activity (только Android)."""
        from jnius import autoclass
        return (autoclass("org.schedule.scheduleapp.ScheduleBridge"),
                autoclass("org.kivy.android.PythonActivity").mActivity)

    def build(self):
        self.title = "Расписание"
        if platform not in ("android", "ios"):
            Window.size = (410, 820)
        os.makedirs(self.files_dir(), exist_ok=True)
        self.db = Database(os.path.join(self.files_dir(), "schedule.db"))
        self.settings_path = os.path.join(self.files_dir(), "settings.json")
        self.settings = self._load_settings()
        try:
            self.anchor = dt.date.fromisoformat(self.settings["anchor"])
        except (KeyError, ValueError):
            self.anchor = default_anchor()
        self.sel = {
            "faculty": self.settings.get("faculty", ""),
            "course": self.settings.get("course", 0),
            "group": self.settings.get("group", ""),
            "subgroup": self.settings.get("subgroup", 0),
        }
        root = Builder.load_string(KV)
        self.root_widget = root
        self.day_buttons = []
        for i, name in enumerate(DAYS):
            b = SelButton(text=name, font_size=sp(15))
            b.bind(on_release=lambda _b, d=i: self.set_view_day(d))
            root.ids.days.add_widget(b)
            self.day_buttons.append(b)
        self.refresh_week_state()
        self.view_day, self.view_week = default_view(self.anchor)
        self.normalize()
        self.refresh_all()
        return root

    def on_start(self):
        if platform == "android":
            try:
                from android import activity
                activity.bind(on_activity_result=self._on_activity_result)
            except Exception as exc:  # noqa: BLE001
                self.status_text = "Ошибка инициализации: {}".format(escape_markup(str(exc)))
        self.export_widget_data()

    def on_pause(self):
        return True

    def on_resume(self):
        self.refresh_week_state()
        self.refresh_cards()

    def _load_settings(self):
        try:
            with open(self.settings_path, encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return {}

    def save_settings(self):
        data = dict(self.sel)
        data["anchor"] = self.anchor.isoformat()
        try:
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
        except OSError:
            pass

    # --- состояние -----------------------------------------------------
    def normalize(self):
        s, db = self.sel, self.db
        facs = db.faculties()
        if not facs:
            s.update(faculty="", course=0, group="")
            return
        if s["faculty"] not in facs:
            s["faculty"] = facs[0]
        courses = db.courses(s["faculty"])
        if s["course"] not in courses:
            s["course"] = courses[0]
        groups = db.groups(s["faculty"], s["course"])
        if s["group"] not in groups:
            s["group"] = groups[0]
        if s["subgroup"] not in (0, 1, 2):
            s["subgroup"] = 0

    def refresh_week_state(self):
        self.cur_week = week_type_for(dt.date.today(), self.anchor)
        self.week_over_text = "Над чертой" + (" •" if self.cur_week == "over" else "")
        self.week_under_text = "Под чертой" + (" •" if self.cur_week == "under" else "")

    def refresh_all(self):
        s = self.sel
        self.fac_text = s["faculty"] or "Нет данных"
        self.course_text = "{} курс".format(s["course"]) if s["course"] else "Курс"
        self.group_text = s["group"] or "Группа"
        self.sub_text = ("Все подгруппы" if not s["subgroup"]
                         else "{} подгруппа".format(s["subgroup"]))
        n = self.db.count()
        self.stats_text = ("В базе: {} занятий, факультетов: {}".format(
            n, len(self.db.faculties())) if n else "База пуста")
        for i, b in enumerate(self.day_buttons):
            b.selected = (i == self.view_day)
        self.refresh_cards()

    def refresh_cards(self):
        grid = self.root_widget.ids.cards
        grid.clear_widgets()
        s = self.sel
        if not s["faculty"]:
            grid.add_widget(self._note(
                "Расписание пусто.\nОткройте вкладку «Экспорт» и импортируйте "
                "Excel-файл или загрузите демо-расписание."))
            return
        lessons = self.db.lessons(s["faculty"], s["course"], s["group"], self.view_day,
                                  self.view_week, s["subgroup"])
        head = "[b]{}[/b] · {}".format(DAYS_FULL[self.view_day], WEEK_NAMES[self.view_week])
        grid.add_widget(self._note(head, color=MUTED))
        if not lessons:
            grid.add_widget(self._note("\n\nВ этот день занятий нет", halign="center",
                                       color=MUTED))
        for lesson in lessons:
            grid.add_widget(LessonCard(lesson))

    @staticmethod
    def _note(text, color=TEXT, halign="left"):
        return WrapLabel(text=text, color=color, halign=halign)

    # --- действия ------------------------------------------------------
    def set_tab(self, name):
        self.tab = name
        self.root_widget.ids.sm.current = name

    def set_view_day(self, day):
        self.view_day = day
        self.refresh_all()

    def set_view_week(self, week):
        self.view_week = week
        self.refresh_cards()

    def set_current_week(self, week):
        monday = monday_of(dt.date.today())
        self.anchor = monday if week == "over" else monday - dt.timedelta(days=7)
        self.save_settings()
        self.refresh_week_state()
        self.view_day, self.view_week = default_view(self.anchor)
        self.export_widget_data()
        self.refresh_all()

    def pick(self, kind):
        s, db = self.sel, self.db
        if not s["faculty"] and kind != "subgroup":
            return
        if kind == "faculty":
            title, opts = "Факультет", [(f, f) for f in db.faculties()]
        elif kind == "course":
            title = "Курс"
            opts = [("{} курс".format(c), c) for c in db.courses(s["faculty"])]
        elif kind == "group":
            title = "Группа"
            opts = [(g, g) for g in db.groups(s["faculty"], s["course"])]
        else:
            title = "Подгруппа"
            opts = [("Все подгруппы", 0), ("1 подгруппа", 1), ("2 подгруппа", 2)]

        def on_pick(value):
            s[kind] = value
            self.normalize()
            self.save_settings()
            self.export_widget_data()
            self.refresh_all()

        self.open_picker(title, opts, on_pick)

    def open_picker(self, title, options, on_pick):
        lst = GridLayout(cols=1, size_hint_y=None, spacing=dp(6), padding=dp(4))
        lst.bind(minimum_height=lst.setter("height"))
        pop = Popup(title=title, size_hint=(.88, None), auto_dismiss=True,
                    separator_color=ACCENT, background_color=(.45, .5, .65, 1))
        pop.height = min(Window.height * .75, dp(52) * len(options) + dp(96))
        for label, value in options:
            b = FlatButton(text=label, size_hint_y=None, height=dp(46))
            b.bind(on_release=lambda _b, v=value: (pop.dismiss(), on_pick(v)))
            lst.add_widget(b)
        sv = ScrollView()
        sv.add_widget(lst)
        pop.content = sv
        pop.open()

    def message(self, title, text):
        box = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(10))
        box.add_widget(WrapLabel(text=text))
        pop = Popup(title=title, content=box, size_hint=(.9, None), height=dp(300),
                    separator_color=ACCENT, background_color=(.45, .5, .65, 1))
        btn = FlatButton(text="Понятно", bg=ACCENT, size_hint_y=None, height=dp(44))
        btn.bind(on_release=pop.dismiss)
        box.add_widget(btn)
        pop.open()

    # --- импорт --------------------------------------------------------
    def choose_file(self):
        if platform == "android":
            try:
                from jnius import autoclass
                Intent = autoclass("android.content.Intent")
                act = autoclass("org.kivy.android.PythonActivity").mActivity
                intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
                intent.addCategory(Intent.CATEGORY_OPENABLE)
                intent.setType("*/*")
                act.startActivityForResult(intent, REQ_PICK)
            except Exception as exc:  # noqa: BLE001
                self.status_text = "[color=#f85149]Не удалось открыть выбор файла: {}[/color]".format(
                    escape_markup(str(exc)))
            return
        chooser = FileChooserListView(filters=["*.xlsx"], path=os.path.expanduser("~"))
        box = BoxLayout(orientation="vertical", spacing=dp(8))
        box.add_widget(chooser)
        row = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(8))
        pop = Popup(title="Выберите .xlsx", content=box, size_hint=(.95, .9),
                    background_color=(.45, .5, .65, 1))
        cancel = FlatButton(text="Отмена")
        ok = FlatButton(text="Открыть", bg=ACCENT)
        cancel.bind(on_release=pop.dismiss)

        def do_open(*_):
            if chooser.selection:
                pop.dismiss()
                self.import_file(chooser.selection[0])

        ok.bind(on_release=do_open)
        row.add_widget(cancel)
        row.add_widget(ok)
        box.add_widget(row)
        pop.open()

    def _on_activity_result(self, request, result, intent):
        if request != REQ_PICK or result != -1 or intent is None:   # -1 = RESULT_OK
            return
        try:
            bridge, act = self.bridge()
            dest = os.path.join(self.files_dir(), "import.xlsx")
            ok = bridge.copyUri(act, intent.getData(), dest)
        except Exception as exc:  # noqa: BLE001
            ok, dest = False, ""
            self.status_text = "[color=#f85149]{}[/color]".format(escape_markup(str(exc)))
        if ok:
            Clock.schedule_once(lambda _dt: self.import_file(dest))
        else:
            self.status_text = "[color=#f85149]Не удалось прочитать выбранный файл[/color]"

    def load_demo(self):
        path = os.path.join(self.files_dir(), "demo.xlsx")
        sample_data.write_workbook(path)
        self.import_file(path)

    def import_file(self, path):
        try:
            rows, errors, warnings = parse_workbook(path)
        except Exception as exc:  # noqa: BLE001 - битый файл, не xlsx и т.п.
            self.status_text = ("[color=#f85149]Не удалось прочитать файл как .xlsx:[/color] "
                                + escape_markup(str(exc)))
            return
        lines = []
        if not rows:
            lines.append("[color=#f85149]Импорт отменён: не найдено ни одной корректной "
                         "строки. Текущее расписание не изменено.[/color]")
        else:
            self.db.replace_all(rows)
            groups = {(r["faculty"], r["course"], r["grp"]) for r in rows}
            facs = {r["faculty"] for r in rows}
            lines.append("[color=#3fb950]Готово: загружено занятий — {}, факультетов — {}, "
                         "групп — {}.[/color]".format(len(rows), len(facs), len(groups)))
        if errors:
            lines.append("[color=#d29922]Пропущено строк с ошибками: {}[/color]".format(len(errors)))
            lines += [escape_markup(e) for e in errors[:10]]
            if len(errors) > 10:
                lines.append("… и ещё {}".format(len(errors) - 10))
        lines += ["[color=#d29922]{}[/color]".format(escape_markup(w)) for w in warnings]
        self.status_text = "\n".join(lines)
        if rows:
            self.normalize()
            self.save_settings()
            self.export_widget_data()
            self.view_day, self.view_week = default_view(self.anchor)
            self.refresh_all()

    # --- виджеты Android -----------------------------------------------
    def export_widget_data(self):
        """Записать widget_data.json для Java-виджета и обновить размещённые виджеты."""
        try:
            s = self.sel
            data = {
                "anchor": self.anchor.isoformat(),
                "defaults": {"faculty": s["faculty"], "course": s["course"],
                             "group": s["group"], "subgroup": s["subgroup"]},
                "lessons": self.db.widget_rows(),
            }
            path = os.path.join(self.files_dir(), "widget_data.json")
            with open(path + ".tmp", "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
            os.replace(path + ".tmp", path)
            if platform == "android":
                bridge, act = self.bridge()
                bridge.refreshAll(act)
        except Exception as exc:  # noqa: BLE001 - виджеты не критичны для приложения
            print("widget export failed:", exc)

    def add_widget_to_desktop(self):
        manual = ("Долгое нажатие на пустом месте рабочего стола > «Виджеты» > "
                  "«Расписание» > перетащите на экран и выберите группу, подгруппу "
                  "и день недели.")
        if platform != "android":
            self.message("Виджет", "Виджеты рабочего стола работают только в Android-версии "
                         "приложения (APK).\n\n" + manual)
            return
        self.export_widget_data()
        try:
            bridge, act = self.bridge()
            ok = bridge.requestPin(act)
        except Exception:  # noqa: BLE001
            ok = False
        if not ok:
            self.message("Добавьте виджет вручную", "Ваш лаунчер не поддерживает быстрое "
                         "добавление.\n\n" + manual)


if __name__ == "__main__":
    ScheduleApp().run()
