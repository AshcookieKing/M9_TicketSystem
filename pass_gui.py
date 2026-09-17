"""
Современный интерфейс автоматизации заявок ММТС-9.
Письмо клиента → Excel шаблон М9 → PDF с печатью выбранного гендиректора.
"""

from __future__ import annotations

import calendar as calmod
import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
import webbrowser


def _fix_tcl_paths():
    """Python 3.12 ломается, если в окружении TCL_LIBRARY от другой версии Python."""
    if getattr(sys, "frozen", False):
        return
    prefix = Path(sys.base_prefix)
    tcl = prefix / "tcl" / "tcl8.6"
    tk = prefix / "tcl" / "tk8.6"
    if tcl.exists():
        os.environ["TCL_LIBRARY"] = str(tcl)
    if tk.exists():
        os.environ["TK_LIBRARY"] = str(tk)


_fix_tcl_paths()

from tkinter import filedialog, messagebox
import tkinter as tk

import customtkinter as ctk

from pass_core import (
    ACTIONS,
    BUNDLE_DIR,
    OBJECT_ADDRESSES,
    OUTPUT_DIR,
    PURPOSES,
    EmailParser,
    Workspace,
    apply_company_defaults,
    build_work_description,
    collect_rooms,
    default_access_room,
    export_excel,
    export_pdf,
    import_client_excel,
    load_catalog,
    load_rooms,
    merge_unique_visitors,
    suggested_filenames,
    visitor_to_request,
)
from line_core import (
    CONNECTORS,
    CONNECTION_TYPES,
    TARIFFS,
    WORK_KINDS,
    LineParser,
    apply_line_company_defaults,
    assignment_from_json,
    empty_line_request,
    export_line_bundle,
    format_outgoing_number,
    line_stamp_path,
)

NAV_ITEMS = [
    ("mail", "Письмо", "Вставка письма и выпуск"),
    ("lines", "Протяжки", "Соединительные линии"),
    ("data", "Данные", "Реквизиты заявки"),
    ("people", "Люди", "Посетители"),
    ("gear", "Техника", "Оборудование"),
    ("car", "Авто", "Автотранспорт"),
    ("stamp", "Печати", "Гендиректора и печати"),
    ("help", "Инструкция", "Регламент работы"),
]

MONTHS_RU = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]
WEEKDAYS_RU = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]

from app_meta import APP_COMPANY, APP_DEVELOPER, APP_DEVELOPER_URL, APP_NAME, APP_VERSION
APP_TITLE = f"{APP_NAME} · {APP_VERSION}"


def _entry_get(widget):
    return widget.get().strip()


def _icon_path():
    ico = BUNDLE_DIR / "assets" / "app.ico"
    return ico if ico.exists() else None


def apply_window_icon(window, native=True):
    """Иконка окна, панели задач и заголовка. CTk на Windows сбрасывает её — ставим ещё и через WinAPI."""
    ico = _icon_path()
    if ico is None:
        return
    path = str(ico)
    try:
        window.iconbitmap(path)
    except Exception:
        pass
    try:
        window.wm_iconbitmap(path)
    except Exception:
        pass
    try:
        window.iconbitmap(default=path)
    except Exception:
        pass
    photos = getattr(window, "_icon_photos", None) or getattr(getattr(window, "master_app", None), "_icon_photos", None)
    if photos:
        try:
            window.iconphoto(False, *photos)
        except Exception:
            pass
    if not native:
        return
    try:
        import ctypes
        window.update_idletasks()
        user32 = ctypes.windll.user32
        IMAGE_ICON, LR_LOADFROMFILE = 1, 0x00000010
        WM_SETICON, ICON_SMALL, ICON_BIG = 0x0080, 0, 1
        GWL_STYLE, WS_CHILD = -16, 0x40000000
        hwnd = int(window.winfo_id())
        style = user32.GetWindowLongW(hwnd, GWL_STYLE)
        while style & WS_CHILD:
            parent = user32.GetParent(hwnd)
            if not parent:
                break
            hwnd = parent
            style = user32.GetWindowLongW(hwnd, GWL_STYLE)
        h_small = user32.LoadImageW(None, path, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        h_big = user32.LoadImageW(None, path, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
        if h_small:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, h_small)
        if h_big:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, h_big)
    except Exception:
        pass


def _parse_date(text):
    text = (text or "").strip()
    if not text:
        return datetime.now()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d.%m.%y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return datetime.now()


class FieldRow(ctk.CTkFrame):
    def __init__(self, master, label, placeholder="", width=420, **kwargs):
        super().__init__(master, fg_color="transparent")
        ctk.CTkLabel(self, text=label, text_color="#94A3B8", width=210, anchor="w").pack(side="left")
        self.entry = ctk.CTkEntry(self, placeholder_text=placeholder, width=width, height=34)
        self.entry.pack(side="left", fill="x", expand=True)

    def get(self):
        return self.entry.get().strip()

    def set(self, value):
        self.entry.delete(0, "end")
        if value:
            self.entry.insert(0, str(value))


class ComboField(ctk.CTkFrame):
    def __init__(self, master, label, values, width=420, **kwargs):
        super().__init__(master, fg_color="transparent")
        ctk.CTkLabel(self, text=label, text_color="#94A3B8", width=210, anchor="w").pack(side="left")
        vals = list(values) or ["-"]
        self.combo = ctk.CTkComboBox(self, values=vals, width=width, height=34)
        self.combo.pack(side="left", fill="x", expand=True)
        self.combo.set(vals[0])

    def get(self):
        value = (self.combo.get() or "").strip()
        if value in {"—", "-", "нет"}:
            return ""
        return value

    def set(self, value):
        value = "" if value is None else str(value).strip()
        values = list(self.combo.cget("values") or [])
        if not value:
            if "—" in values:
                self.combo.set("—")
                return
            if values:
                self.combo.set(values[0])
                return
        if value and value not in values:
            values.append(value)
            self.combo.configure(values=values)
        self.combo.set(value)

    def set_values(self, values):
        current = self.get()
        vals = list(values) or ["-"]
        self.combo.configure(values=vals)
        if current in vals:
            self.combo.set(current)
        else:
            self.combo.set(vals[0])


class DateField(ctk.CTkFrame):
    def __init__(self, master, label, width=280, **kwargs):
        super().__init__(master, fg_color="transparent")
        ctk.CTkLabel(self, text=label, text_color="#94A3B8", width=210, anchor="w").pack(side="left")
        wrap = ctk.CTkFrame(self, fg_color="transparent")
        wrap.pack(side="left", fill="x", expand=True)
        self.entry = ctk.CTkEntry(wrap, width=max(140, width - 44), height=34, placeholder_text="ДД.ММ.ГГГГ")
        self.entry.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(
            wrap, text="▾", width=36, height=34, fg_color="#1E293B", hover_color="#334155",
            command=self._open,
        ).pack(side="left", padx=(6, 0))

    def get(self):
        return self.entry.get().strip()

    def set(self, value):
        self.entry.delete(0, "end")
        if value:
            self.entry.insert(0, str(value))

    def _open(self):
        CalendarPopup(self.winfo_toplevel(), self.get(), self.set, anchor=self)


class CalendarPopup(ctk.CTkToplevel):
    def __init__(self, master, initial, on_pick, anchor=None):
        super().__init__(master)
        self.on_pick = on_pick
        self.current = _parse_date(initial)
        self.view = datetime(self.current.year, self.current.month, 1)
        self.overrideredirect(True)
        self.configure(fg_color="#0F172A")
        self.transient(master)
        self.attributes("-topmost", True)
        apply_window_icon(self, native=False)
        self.card = ctk.CTkFrame(self, fg_color="#111827", corner_radius=16, border_width=1, border_color="#2563EB")
        self.card.pack(fill="both", expand=True, padx=1, pady=1)
        self._build()
        self.bind("<Escape>", lambda _e: self.destroy())
        self.after(10, self._place, anchor)
        self.after(20, self._grab)

    def _grab(self):
        try:
            self.grab_set()
            self.focus_force()
        except Exception:
            pass

    def destroy(self):
        master = self.master
        try:
            self.grab_release()
        except Exception:
            pass
        super().destroy()
        try:
            master.grab_set()
        except Exception:
            pass

    def _place(self, anchor):
        self.update_idletasks()
        w, h = 292, 318
        if anchor is not None:
            ax = anchor.winfo_rootx()
            ay = anchor.winfo_rooty() + anchor.winfo_height() + 4
            self.geometry(f"{w}x{h}+{ax}+{ay}")
        else:
            self.geometry(f"{w}x{h}")
        self.lift()

    def _build(self):
        for child in self.card.winfo_children():
            child.destroy()
        head = ctk.CTkFrame(self.card, fg_color="transparent")
        head.pack(fill="x", padx=10, pady=(10, 4))
        ctk.CTkButton(head, text="‹", width=32, height=28, fg_color="#1E293B", command=self._prev).pack(side="left")
        ctk.CTkLabel(
            head, text=f"{MONTHS_RU[self.view.month - 1]} {self.view.year}",
            font=ctk.CTkFont(weight="bold"),
        ).pack(side="left", expand=True)
        ctk.CTkButton(head, text="›", width=32, height=28, fg_color="#1E293B", command=self._next).pack(side="right")

        grid = ctk.CTkFrame(self.card, fg_color="transparent")
        grid.pack(padx=10, pady=4)
        for i, name in enumerate(WEEKDAYS_RU):
            ctk.CTkLabel(grid, text=name, width=34, text_color="#64748B").grid(row=0, column=i, padx=1, pady=1)

        weeks = calmod.Calendar(0).monthdayscalendar(self.view.year, self.view.month)
        today = datetime.now().date()
        selected = self.current.date()
        for r, week in enumerate(weeks, start=1):
            for c, day in enumerate(week):
                if day == 0:
                    ctk.CTkLabel(grid, text="", width=34, height=28).grid(row=r, column=c, padx=1, pady=1)
                    continue
                dt = datetime(self.view.year, self.view.month, day).date()
                fg = "#1D4ED8" if dt == selected else ("#0F766E" if dt == today else "#1E293B")
                tk.Button(
                    grid, text=str(day), width=3, height=1, bd=0, relief="flat",
                    bg=fg, fg="#F8FAFC", activebackground="#2563EB",
                    font=("Segoe UI", 9),
                    command=lambda d=day: self._choose(d),
                ).grid(row=r, column=c, padx=1, pady=1)

        ctk.CTkButton(self.card, text="Сегодня", height=30, fg_color="#1E293B", command=self._today).pack(
            fill="x", padx=12, pady=(4, 10)
        )

    def _prev(self):
        month = self.view.month - 1
        year = self.view.year
        if month < 1:
            month, year = 12, year - 1
        self.view = datetime(year, month, 1)
        self._build()

    def _next(self):
        month = self.view.month + 1
        year = self.view.year
        if month > 12:
            month, year = 1, year + 1
        self.view = datetime(year, month, 1)
        self._build()

    def _today(self):
        self._apply(datetime.now())

    def _choose(self, day):
        self._apply(datetime(self.view.year, self.view.month, day))

    def _apply(self, dt):
        self.on_pick(dt.strftime("%d.%m.%Y"))
        self.destroy()


class RoomPicker(ctk.CTkFrame):
    def __init__(self, master, rooms, app):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self.rooms_catalog = list(rooms)
        self.selected = []
        ctk.CTkLabel(self, text="Помещения", text_color="#94A3B8", width=210, anchor="nw").pack(side="left", pady=6)
        right = ctk.CTkFrame(self, fg_color="transparent")
        right.pack(side="left", fill="x", expand=True)
        self.chips = ctk.CTkFrame(right, fg_color="#0F172A", corner_radius=10)
        self.chips.pack(fill="x")
        bar = ctk.CTkFrame(right, fg_color="transparent")
        bar.pack(fill="x", pady=(6, 0))
        self.search = ctk.CTkEntry(bar, placeholder_text="Найти помещение из списка шаблона…", height=34)
        self.search.pack(side="left", fill="x", expand=True)
        self.search.bind("<Return>", lambda _e: self._add_from_search())
        ctk.CTkButton(bar, text="Добавить", width=100, height=34, command=self._add_from_search).pack(side="left", padx=(6, 0))
        ctk.CTkButton(bar, text="Список", width=80, height=34, fg_color="#1E293B", command=self._open_list).pack(side="left", padx=(6, 0))
        self.hint = ctk.CTkLabel(right, text="", text_color="#64748B", anchor="w")
        self.hint.pack(fill="x", pady=(4, 0))

    def set_catalog(self, rooms):
        self.rooms_catalog = list(rooms)
        self._refresh_hint()

    def get(self):
        return list(self.selected)

    def set(self, values):
        if isinstance(values, str):
            values = [part.strip() for part in values.replace(";", ",").split(",") if part.strip()]
        seen = []
        for item in values or []:
            if item and item not in seen:
                seen.append(item)
        self.selected = seen[:6]
        self._render()

    def _render(self):
        for child in self.chips.winfo_children():
            child.destroy()
        if not self.selected:
            ctk.CTkLabel(self.chips, text="Помещение не выбрано", text_color="#64748B").pack(anchor="w", padx=10, pady=8)
        else:
            wrap = ctk.CTkFrame(self.chips, fg_color="transparent")
            wrap.pack(fill="x", padx=6, pady=6)
            for idx, room in enumerate(self.selected):
                chip = ctk.CTkFrame(wrap, fg_color="#1D4ED8", corner_radius=8)
                chip.pack(side="left", padx=4, pady=4)
                ctk.CTkLabel(chip, text=room, text_color="#F8FAFC").pack(side="left", padx=(8, 4), pady=4)
                ctk.CTkButton(
                    chip, text="×", width=22, height=22, fg_color="transparent", hover_color="#1E3A8A",
                    command=lambda i=idx: self._remove(i),
                ).pack(side="left", padx=(0, 4))
        self._refresh_hint()

    def _refresh_hint(self):
        left = 6 - len(self.selected)
        self.hint.configure(text=f"Из официального списка шаблона, до 6 помещений. Осталось слотов: {left}")

    def _remove(self, index):
        if 0 <= index < len(self.selected):
            del self.selected[index]
            self._render()

    def _match(self, query):
        query = (query or "").strip().lower()
        if not query:
            return list(self.rooms_catalog)
        compact = query.replace(" ", "").replace("№", "")
        ranked = []
        for room in self.rooms_catalog:
            hay = room.lower().replace(" ", "").replace("№", "")
            if query in room.lower() or compact in hay:
                ranked.append(room)
        return ranked

    def _add_room(self, room):
        room = (room or "").strip()
        if not room or room in self.selected:
            return
        if len(self.selected) >= 6:
            messagebox.showwarning("Помещения", "В шаблоне ММТС-9 не больше 6 помещений на заявку.")
            return
        if room not in self.rooms_catalog:
            messagebox.showwarning("Помещения", "Выберите помещение из списка шаблона, иначе Excel подсветит ошибку.")
            return
        self.selected.append(room)
        self.search.delete(0, "end")
        self._render()

    def _add_from_search(self):
        matches = self._match(self.search.get())
        available = [item for item in matches if item not in self.selected]
        if len(available) == 1:
            self._add_room(available[0])
            return
        self._open_list(prefill=self.search.get())

    def _open_list(self, prefill=""):
        self.app._room_picker_dialog(self, prefill)


HELP_SECTIONS = [
    (
        "Назначение",
        "M9 Gate готовит заявку на доступ на объект АО «ММТС-9» в том же составе, "
        "что и официальный Excel-шаблон: реквизиты заявителя, помещения из справочника, "
        "посетители, оборудование с вносом/выносом, автотранспорт и PDF с печатью выбранного "
        "генерального директора. Программа не заменяет визуальную проверку файла в Excel — "
        "шаблон содержит формулы контроля, и ММТС-9 принимает заявку только без ошибок заполнения.",
    ),
    (
        "Подготовка рабочего места",
        "1. В шапке выберите компанию-заявителя (Техноцентр или Тех РУ) и подписанта. "
        "Договор, ИНН и телефон подставятся сами.\n"
        "2. На вкладке «Печати» убедитесь, что у выбранного гендиректора загружен файл печати. "
        "Без него PDF сформируется, но штамп не встанет.\n"
        "3. Команда посетителей компании подтягивается автоматически. Состав можно править на вкладке «Люди».",
    ),
    (
        "Типовой маршрут заявки",
        "Откройте «Письмо». Вставьте текст клиента как есть и, если во вложении был Excel, "
        "прикрепите его. Нажмите «Разобрать письмо». Программа снимет даты, цель, стойки "
        "вида 12.35 / 12.28, серийники и сопоставит помещения с официальным списком шаблона. "
        "Договор клиента (например, 17033) в заявку не пишется — в Excel уходит договор заявителя с ММТС-9.\n\n"
        "Проверьте вкладки «Данные», «Люди», «Техника». Затем «Создать Excel + PDF» внизу слева. "
        "Файлы сохраняются в data/output. Перед отправкой откройте Excel и убедитесь, что нет красных ошибок шаблона.",
    ),
    (
        "Помещения",
        "Поле помещений — это не свободный текст, а справочник листа «Список» шаблона. "
        "Для Москвы (Бутлерова, 7) первым обычно идёт «Линия входных турникетов», дальше — залы, "
        "шахты, выделенные зоны. Для Нижнего Новгорода (Федосеенко, 35) список другой, он "
        "переключается вместе с объектом.\n\n"
        "Наберите номер зала (например, 12.28) и нажмите Enter либо «Список». В шаблоне не больше "
        "шести помещений на одну заявку — это ограничение ячеек C6–H6. Помещение оборудования "
        "тоже выбирается только из этого списка, иначе Excel пометит строку ошибкой.",
    ),
    (
        "Даты",
        "Все даты задаются календарём: договор, срок действия, период работ, дата заявки, "
        "рождение и выдача паспорта. Формат в файле — ДД.ММ.ГГГГ, как требует шаблон. "
        "Кнопки «Сегодня», «+3 дня», «+7 дней», «+14 дней» выставляют период работ. "
        "Дата начала не должна быть раньше текущей — шаблон это проверяет.",
    ),
    (
        "Посетители, техника, авто",
        "Посетители — до 20 человек: ФИО, дата рождения, паспорт, работодатель, при необходимости "
        "группа и удостоверение ЭБ, примечание. Гражданство заполняется только для иностранцев, "
        "как в Excel.\n\n"
        "Оборудование — до 40 позиций. Обязательны количество, единица измерения, «Внос» или «Вынос». "
        "Серийный номер и причина его отсутствия взаимоисключающие: либо номер, либо причина из списка "
        "(«Отсутствует на корпусе», «Не читаем / поврежден», «Не подлежит указанию…»). "
        "Ряд и место берутся из справочника шаблона и собираются в колонку «ряд , место».\n\n"
        "Автотранспорт заполняется, если нужен въезд под погрузку. В шаблоне заголовок говорит о пяти "
        "машинах; программа может выгрузить больше, но для приёмки лучше не превышать лимит бланка.",
    ),
    (
        "Выпуск и контроль",
        "«Создать Excel + PDF» заполняет актуальный шаблон ТехРу и собирает PDF на бланке заявки "
        "с печатью. «Только Excel» / «Только PDF» — если нужен один комплект. Номер заявки "
        "расходуется при выпуске и запоминается в профиле компании.\n\n"
        "После выпуска откройте xlsx в Excel. Красные пометки слева — ошибки шаблона: пустое помещение, "
        "просроченная дата, нет вноса/выноса, одновременно заполнены серийник и причина. "
        "Исправьте в программе и выпустите файл заново, не правьте вручную без нужды: формулы листа "
        "«Текст заявки на печать» зависят от тех же ячеек.",
    ),
    (
        "Клавиши и окна",
        "Карточки добавления и правки открываются поверх окна, без отдельной иконки в панели задач. "
        "Esc закрывает карточку без сохранения. Щелчок по затемнению — тоже закрытие. "
        "Сохранение только кнопкой «Сохранить».",
    ),
    (
        "Протяжки",
        "Вкладка «Протяжки» — это не пропуск, а письмо техническому директору АО «ММТС-9» "
        "(Ушмайкин К.Э.) на монтаж или демонтаж соединительной линии. Бланк как у 04-08.743 и C03036: "
        "шапка заявителя, таблица сторон A/B, приложение JSON. Подписывает инженер (Порозов, Пронякин), "
        "не гендиректор.\n\n"
        "Как пользоваться:\n"
        "1. В шапке выберите компанию-заявителя (обычно Тех РУ).\n"
        "2. Откройте «Протяжки». Вставьте письмо клиента или уведомление ДЦ как есть.\n"
        "3. Если к письму шло JSON-задание линии — «JSON-задание» и укажите txt/json.\n"
        "4. «Разобрать письмо». Проверьте вид работ, исх. номер, стороны A/B, разъёмы, волокна.\n"
        "5. Внизу слева «Создать PDF протяжки». Файлы: PDF письма и JSON в data/output.\n\n"
        "Два типовых письма:\n"
        "• Отказ от кроссировок («Прошу отказаться от кроссировок: 5276/25 и 5959/24») — демонтаж, "
        "номера исходных заданий в текст, печать обычно выключена.\n"
        "• Уведомление ДЦ о передаче протяжки (инв. C03036, стойка, комната ГО, разъёмы, клиент) — "
        "монтаж, срочность из JSON, на PDF ставится круглая печать ТЕХ.РУ.\n\n"
        "Переключатель «Печать ТЕХ.РУ на PDF» можно включить вручную. Исх. номер вроде C03036 "
        "не тратит счётчик 04-08; пустой номер берётся из префикса компании.",
    ),
    (
        "Обновления",
        "При доступе в сеть программа сверяет свою сборку с GitHub Releases "
        f"({APP_NAME}). Если вышла более новая версия, появится запрос на обновление: "
        "скачается архив релиза и приложение перезапустится уже на новой сборке. "
        "После перезапуска появится окно, что обновление установлено.",
    ),
    (
        "Разработка",
        f"{APP_NAME} {APP_VERSION}. Бета-версия.\n"
        f"Разработана для компании {APP_COMPANY}.\n"
        f"Разработчик: {APP_DEVELOPER}, {APP_DEVELOPER_URL}",
    ),
]


class OverlayShell(ctk.CTkToplevel):
    def __init__(self, master, width_ratio=0.56, height_ratio=0.82):
        super().__init__(master)
        self.master_app = master
        master._overlay = self
        self.overrideredirect(True)
        self.transient(master)
        self.configure(fg_color="#020617")
        self.attributes("-topmost", True)
        apply_window_icon(self, native=False)
        self.bind("<Escape>", lambda _e: self.close())
        self.bind("<Button-1>", self._click_dimmer)
        self._width_ratio = width_ratio
        self._height_ratio = height_ratio
        self.card = ctk.CTkFrame(self, fg_color="#111827", corner_radius=22, border_width=1, border_color="#3B82F6")
        self.card.place(relx=0.5, rely=0.5, anchor="center", relwidth=width_ratio, relheight=height_ratio)
        self.card.bind("<Button-1>", lambda e: "break")
        self._sync()
        self.after(20, self._grab)

    def _grab(self):
        try:
            self.grab_set()
            self.focus_force()
        except Exception:
            pass

    def _sync(self):
        self.update_idletasks()
        try:
            x = self.master_app.winfo_rootx()
            y = self.master_app.winfo_rooty()
            w = self.master_app.winfo_width()
            h = self.master_app.winfo_height()
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    def _click_dimmer(self, event):
        widget = event.widget
        try:
            if str(widget) == str(self):
                self.close()
        except Exception:
            pass

    def close(self):
        try:
            self.grab_release()
        except Exception:
            pass
        if getattr(self.master_app, "_overlay", None) is self:
            self.master_app._overlay = None
        self.destroy()


class OverlayForm(OverlayShell):
    def __init__(self, master, title, fields, data, on_save):
        super().__init__(master)
        self.on_save = on_save
        head = ctk.CTkFrame(self.card, fg_color="transparent")
        head.pack(fill="x", padx=22, pady=(18, 8))
        ctk.CTkLabel(head, text=title, font=ctk.CTkFont(size=22, weight="bold")).pack(side="left")
        ctk.CTkLabel(head, text="Esc — закрыть", text_color="#64748B").pack(side="right")
        box = ctk.CTkScrollableFrame(self.card, fg_color="transparent")
        box.pack(fill="both", expand=True, padx=14, pady=(0, 8))
        self.widgets = {}
        for spec in fields:
            key, label, *rest = spec
            kind = rest[0] if rest else "text"
            if kind == "date":
                row = DateField(box, label, width=280)
            elif kind == "combo":
                values = rest[1] if len(rest) > 1 else ["-"]
                row = ComboField(box, label, values, width=280)
            else:
                row = FieldRow(box, label, width=280)
            row.pack(fill="x", pady=4, padx=6)
            row.set(data.get(key, ""))
            self.widgets[key] = row
        actions = ctk.CTkFrame(self.card, fg_color="transparent")
        actions.pack(fill="x", padx=22, pady=(4, 18))
        ctk.CTkButton(actions, text="Отмена", fg_color="#1E293B", height=42, command=self.close).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(actions, text="Сохранить", height=42, command=self._save).pack(side="left", fill="x", expand=True)

    def _save(self):
        payload = {key: row.get() for key, row in self.widgets.items()}
        self.on_save(payload)
        self.close()


class OverlayRoomList(OverlayShell):
    def __init__(self, master, picker, prefill=""):
        super().__init__(master, width_ratio=0.5, height_ratio=0.78)
        self.picker = picker
        head = ctk.CTkFrame(self.card, fg_color="transparent")
        head.pack(fill="x", padx=22, pady=(18, 8))
        ctk.CTkLabel(head, text="Помещения шаблона ММТС-9", font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")
        ctk.CTkLabel(head, text="Esc — закрыть", text_color="#64748B").pack(side="right")
        self.search = ctk.CTkEntry(self.card, placeholder_text="Фильтр: 12.28, турникет, шахта…", height=36)
        self.search.pack(fill="x", padx=22, pady=(0, 10))
        self.search.insert(0, prefill or "")
        self.search.bind("<KeyRelease>", lambda _e: self._render())
        self.box = ctk.CTkScrollableFrame(self.card, fg_color="transparent")
        self.box.pack(fill="both", expand=True, padx=14, pady=(0, 18))
        self._render()
        self.after(50, self.search.focus_set)

    def _render(self):
        for child in self.box.winfo_children():
            child.destroy()
        query = self.search.get()
        items = [room for room in self.picker._match(query) if room not in self.picker.selected]
        if not items:
            ctk.CTkLabel(self.box, text="Нет совпадений в справочнике", text_color="#64748B").pack(pady=16)
            return
        for room in items:
            row = ctk.CTkFrame(self.box, fg_color="#0F172A", corner_radius=10)
            row.pack(fill="x", pady=3, padx=4)
            ctk.CTkLabel(row, text=room, anchor="w").pack(side="left", padx=12, pady=8)
            ctk.CTkButton(row, text="Выбрать", width=90, command=lambda r=room: self._choose(r)).pack(side="right", padx=8, pady=6)

    def _choose(self, room):
        self.picker._add_room(room)
        self.close()


class OverlayPool(OverlayShell):
    def __init__(self, master):
        super().__init__(master, width_ratio=0.5, height_ratio=0.78)
        head = ctk.CTkFrame(self.card, fg_color="transparent")
        head.pack(fill="x", padx=22, pady=(18, 8))
        ctk.CTkLabel(head, text="База людей", font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")
        ctk.CTkLabel(head, text="Esc — закрыть", text_color="#64748B").pack(side="right")
        box = ctk.CTkScrollableFrame(self.card, fg_color="transparent")
        box.pack(fill="both", expand=True, padx=14, pady=(0, 18))
        for visitor in master.workspace.visitors():
            row = ctk.CTkFrame(box, fg_color="#0F172A", corner_radius=10)
            row.pack(fill="x", pady=3, padx=4)
            title = f"{visitor.get('surname')} {visitor.get('name')} {visitor.get('patronymic')}"
            meta = f"паспорт {visitor.get('passport_series', '')} {visitor.get('passport_number', '')}"
            text = ctk.CTkFrame(row, fg_color="transparent")
            text.pack(side="left", fill="x", expand=True, padx=12, pady=6)
            ctk.CTkLabel(text, text=title, anchor="w").pack(anchor="w")
            ctk.CTkLabel(text, text=meta, text_color="#94A3B8", anchor="w").pack(anchor="w")
            ctk.CTkButton(row, text="В заявку", width=90, command=lambda v=visitor: self._add(v)).pack(side="right", padx=8, pady=8)

    def _add(self, visitor):
        company = self.master_app.workspace.get_company()
        self.master_app.visitors_data = merge_unique_visitors(
            self.master_app.visitors_data, [visitor_to_request(visitor, company.get("name", ""))]
        )
        self.master_app._render_people()


class PassApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1380x860")
        self.minsize(1180, 740)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.configure(fg_color="#0B1220")
        self._icon_photos = []
        self._set_app_icon()

        self.workspace = Workspace()
        self.catalog = load_catalog()
        self.rooms = load_rooms()
        self.parser = EmailParser(self.rooms)
        self.request_data = {}
        self.visitors_data = []
        self.equipment_data = []
        self.vehicles_data = []
        self.client_excel_path = None
        self.last_parsed = {}
        self.line_data = {}
        self.line_json_path = None
        self.current_page = "mail"
        self.pages = {}
        self.nav_buttons = {}
        self._menu_lock = False
        self._overlay = None
        self.bind("<Escape>", self._on_escape)

        self._build_layout()
        self._load_company_into_form(reset_people=True)
        self.show_page("mail")
        self.after(400, self._startup_update_flow)

    def _on_escape(self, _event=None):
        if self._overlay and self._overlay.winfo_exists():
            self._overlay.close()
            return "break"
        return None

    def _set_app_icon(self):
        png = BUNDLE_DIR / "assets" / "app.png"
        try:
            from PIL import Image, ImageTk
            if png.exists():
                src = Image.open(png).convert("RGB")
                img32 = src.resize((32, 32), Image.Resampling.LANCZOS)
                img16 = src.resize((16, 16), Image.Resampling.LANCZOS)
                self._icon_photo = ImageTk.PhotoImage(img32)
                self._icon_small = ImageTk.PhotoImage(img16)
                self._icon_photos = [self._icon_photo, self._icon_small]
        except Exception:
            self._icon_photos = []
        apply_window_icon(self)
        self.after(250, lambda w=self: apply_window_icon(w))

    # ── layout ──────────────────────────────────────────────
    def _build_layout(self):
        self.sidebar = ctk.CTkFrame(self, width=232, corner_radius=0, fg_color="#111827")
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        brand.pack(fill="x", padx=18, pady=(22, 8))
        ctk.CTkLabel(brand, text="M9 Gate", font=ctk.CTkFont(size=26, weight="bold"), text_color="#60A5FA").pack(anchor="w")
        ctk.CTkLabel(brand, text="система заявок", text_color="#94A3B8").pack(anchor="w")
        ctk.CTkLabel(brand, text=f"{APP_VERSION}  ·  {APP_COMPANY}", text_color="#475569").pack(anchor="w", pady=(4, 0))

        for key, title, _hint in NAV_ITEMS:
            btn = ctk.CTkButton(
                self.sidebar, text=title, anchor="w", height=40, corner_radius=10,
                fg_color="transparent", hover_color="#1E293B", text_color="#E5E7EB",
                command=lambda k=key: self.show_page(k),
            )
            btn.pack(fill="x", padx=12, pady=3)
            self.nav_buttons[key] = btn

        self.main_action = ctk.CTkButton(
            self.sidebar, text="Создать Excel + PDF", height=44, corner_radius=12,
            fg_color="#10B981", hover_color="#059669", text_color="#052E1C",
            font=ctk.CTkFont(weight="bold"), command=self._main_action,
        )
        self.main_action.pack(side="bottom", fill="x", padx=16, pady=18)

        right = ctk.CTkFrame(self, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True)

        self.topbar = ctk.CTkFrame(right, height=72, fg_color="#111827", corner_radius=0)
        self.topbar.pack(fill="x")
        self._build_topbar()

        self.content = ctk.CTkFrame(right, fg_color="transparent")
        self.content.pack(fill="both", expand=True, padx=18, pady=16)

        self._page_builders = {
            "mail": self._build_mail_page,
            "lines": self._build_lines_page,
            "data": self._build_data_page,
            "people": self._build_people_page,
            "gear": self._build_equipment_page,
            "car": self._build_vehicles_page,
            "stamp": self._build_stamps_page,
            "help": self._build_help_page,
        }
        self._build_mail_page()

    def _build_topbar(self):
        wrap = ctk.CTkFrame(self.topbar, fg_color="transparent")
        wrap.pack(fill="both", expand=True, padx=18, pady=12)

        ctk.CTkLabel(wrap, text="Компания-заявитель", text_color="#64748B").pack(side="left", padx=(0, 8))
        self.company_menu = ctk.CTkOptionMenu(
            wrap, width=220, height=34, command=self._on_company_changed, values=["-"]
        )
        self.company_menu.pack(side="left", padx=(0, 18))

        ctk.CTkLabel(wrap, text="Гендиректор / печать", text_color="#64748B").pack(side="left", padx=(0, 8))
        self.director_menu = ctk.CTkOptionMenu(
            wrap, width=240, height=34, command=self._on_director_changed, values=["-"]
        )
        self.director_menu.pack(side="left")

        self.stamp_hint = ctk.CTkLabel(wrap, text="", text_color="#94A3B8")
        self.stamp_hint.pack(side="left", padx=16)
        self._refresh_company_menus()

    def _refresh_company_menus(self):
        self._menu_lock = True
        try:
            companies = self.workspace.companies()
            names = [c.get("short_name") or c.get("name") for c in companies] or ["-"]
            self.company_menu.configure(values=names)
            active = self.workspace.get_company()
            if active:
                self.company_menu.set(active.get("short_name") or active.get("name"))
            self._refresh_director_menu()
        finally:
            self._menu_lock = False

    def _refresh_director_menu(self):
        company = self.workspace.get_company()
        directors = company.get("directors", []) if company else []
        names = [d.get("name") for d in directors] or ["-"]
        self.director_menu.configure(values=names)
        director = self.workspace.get_director()
        if director:
            self.director_menu.set(director.get("name"))
        stamp = self.workspace.stamp_path(director)
        if director and stamp:
            self.stamp_hint.configure(text=f"печать: {Path(stamp).name}", text_color="#34D399")
        elif director:
            self.stamp_hint.configure(text="печать не загружена", text_color="#F59E0B")
        else:
            self.stamp_hint.configure(text="")

    def show_page(self, key):
        self.current_page = key
        if key not in self.pages:
            builder = self._page_builders.get(key)
            if builder:
                builder()
        for name, frame in self.pages.items():
            frame.pack_forget()
        self.pages[key].pack(fill="both", expand=True)
        for name, btn in self.nav_buttons.items():
            btn.configure(fg_color="#1D4ED8" if name == key else "transparent")
        if hasattr(self, "main_action"):
            if key == "lines":
                self.main_action.configure(text="Создать PDF протяжки")
            else:
                self.main_action.configure(text="Создать Excel + PDF")
        if key == "stamp":
            self._render_directors()
        if key == "people":
            self._render_people()
        if key == "gear":
            self._render_equipment()
        if key == "car":
            self._render_vehicles()
        if key == "lines":
            self._fill_line_fields()

    # ── mail page ───────────────────────────────────────────
    def _build_mail_page(self):
        page = ctk.CTkFrame(self.content, fg_color="transparent")
        self.pages["mail"] = page

        left = ctk.CTkFrame(page, fg_color="#111827", corner_radius=18)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))
        ctk.CTkLabel(left, text="Письмо клиента", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=18, pady=(16, 4))
        ctk.CTkLabel(
            left,
            text="Вставьте письмо как есть. Программа сама достанет дату, договор, стойки и внос/вынос.",
            text_color="#94A3B8", wraplength=620, justify="left",
        ).pack(anchor="w", padx=18, pady=(0, 8))
        self.mail_box = ctk.CTkTextbox(left, font=ctk.CTkFont(family="Consolas", size=13), corner_radius=12)
        self.mail_box.pack(fill="both", expand=True, padx=18, pady=(0, 12))
        self.example_mail = (
            "Добрый день.\n"
            "Договор 17033 с ЭЙБИСИ-ЛАБС\n"
            "Планируем работы на 18 сентября\n"
            "Будет внос двух серверов из заявки с уставной в 12.35 ряд 3 место 4  и 12.28 ряд 9а место 12\n"
            "Будет вынос сервера 28062  из 12.35. Ряд 3 место 4\n"
            "Заявка во вложении\n\n"
            "--\n"
            "С уважением, Югов Антон\n"
            "Компания \"ABC-Labs\"\n"
            "Тел. +7 (499) 490-63-21\n"
        )
        self.mail_box.insert("1.0", self.example_mail)

        actions = ctk.CTkFrame(left, fg_color="transparent")
        actions.pack(fill="x", padx=18, pady=(0, 16))
        ctk.CTkButton(actions, text="Разобрать письмо", height=40, command=self.parse_mail).pack(side="left")
        ctk.CTkButton(actions, text="Прикрепить Excel клиента", fg_color="#1E293B", height=40, command=self.attach_excel).pack(side="left", padx=8)
        ctk.CTkButton(actions, text="Очистить", fg_color="#1E293B", height=40, command=self.clear_mail).pack(side="left")
        self.excel_label = ctk.CTkLabel(actions, text="вложение не выбрано", text_color="#64748B")
        self.excel_label.pack(side="left", padx=10)

        right = ctk.CTkFrame(page, width=420, fg_color="#111827", corner_radius=18)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)
        ctk.CTkLabel(right, text="Что получилось", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=18, pady=(16, 8))
        self.summary_box = ctk.CTkTextbox(right, state="disabled", corner_radius=12)
        self.summary_box.pack(fill="both", expand=True, padx=18, pady=(0, 12))
        ctk.CTkButton(right, text="Только Excel", fg_color="#1E293B", command=self.create_excel_only).pack(fill="x", padx=18, pady=(0, 8))
        ctk.CTkButton(right, text="Только PDF с печатью", fg_color="#1E293B", command=self.create_pdf_only).pack(fill="x", padx=18, pady=(0, 16))
        self._set_summary("Вставьте письмо и нажмите «Разобрать письмо».")

    def _set_summary(self, text):
        self.summary_box.configure(state="normal")
        self.summary_box.delete("1.0", "end")
        self.summary_box.insert("1.0", text)
        self.summary_box.configure(state="disabled")

    def clear_mail(self):
        self.mail_box.delete("1.0", "end")
        self.client_excel_path = None
        self.excel_label.configure(text="вложение не выбрано", text_color="#64748B")

    def attach_excel(self):
        path = filedialog.askopenfilename(title="Excel клиента", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")])
        if not path:
            return
        self.client_excel_path = path
        self.excel_label.configure(text=Path(path).name, text_color="#93C5FD")

    def parse_mail(self):
        text = self.mail_box.get("1.0", "end")
        is_example = text.strip() == self.example_mail.strip()
        if not text.strip() and not self.client_excel_path:
            messagebox.showwarning("Письмо", "Вставьте текст письма или прикрепите Excel клиента.")
            return
        if is_example and self.client_excel_path:
            parsed = self.parser.parse("")
        else:
            parsed = self.parser.parse(text)
        self.last_parsed = parsed
        imported = None
        if self.client_excel_path:
            try:
                imported = import_client_excel(self.client_excel_path, self.rooms)
            except Exception as exc:
                messagebox.showerror("Excel клиента", f"Не удалось прочитать вложение:\n{exc}")
                return

        suggested = self.workspace.suggest_company_from_text(text + " " + (parsed.get("client") or ""))
        if suggested and suggested != self.workspace.data.get("active_company_id"):
            self.workspace.set_active_company(suggested)
            self._refresh_company_menus()

        company = self.workspace.get_company()
        director = self.workspace.get_director()
        request = {}
        visitors = []
        equipment = []
        vehicles = []

        if imported:
            request = dict(imported.get("request_data") or {})
            visitors = imported.get("visitors") or []
            equipment = imported.get("equipment") or []
            vehicles = imported.get("vehicles") or []

        if parsed.get("work_start_date"):
            request["work_start_date"] = parsed["work_start_date"]
        if parsed.get("work_end_date"):
            request["work_end_date"] = parsed["work_end_date"]
        if parsed.get("purpose") and (not request.get("purpose") or request.get("purpose") in {"ТО", ""}):
            request["purpose"] = parsed["purpose"]
        if parsed.get("equipment"):
            if equipment:
                equipment = self._overlay_equipment(equipment, parsed["equipment"])
            else:
                equipment = parsed["equipment"]

        request["rooms"] = collect_rooms(equipment, request.get("rooms"), request.get("object_address"))
        request["work_description"] = build_work_description(equipment) or request.get("work_description", "")
        apply_company_defaults(request, company, director)
        request["request_number"] = self.workspace.peek_request_number(company)

        team = self.workspace.default_team(company)
        from_letter = self.workspace.find_visitors_in_text(text, company)
        for person in visitors:
            person["organization"] = company.get("name", person.get("organization", ""))
        visitors = merge_unique_visitors(from_letter, visitors, team)

        self.request_data = request
        self.visitors_data = visitors
        self.equipment_data = equipment
        self.vehicles_data = vehicles
        self._fill_data_fields(request)
        self._render_people()
        self._render_equipment()
        self._render_vehicles()
        self._set_summary(self._summary_text(parsed, imported, company, director))

    def _overlay_equipment(self, imported, from_mail):
        used = set()
        result = []
        for src in imported:
            result.append(dict(src))
        for mail_item in from_mail:
            matched = None
            for idx, item in enumerate(result):
                if idx in used:
                    continue
                same_room = (item.get("room") or "") == (mail_item.get("room") or "")
                same_action = (item.get("action") or "") == (mail_item.get("action") or "") or not item.get("action")
                serial_hit = mail_item.get("serial_number") and mail_item["serial_number"] in (item.get("serial_number") or "")
                if serial_hit or (same_room and same_action and not item.get("row")):
                    matched = idx
                    break
            if matched is None:
                result.append(dict(mail_item))
                continue
            used.add(matched)
            target = result[matched]
            for key in ("action", "room", "row", "place"):
                if mail_item.get(key) and not target.get(key):
                    target[key] = mail_item[key]
            if mail_item.get("serial_number") and not target.get("serial_number"):
                target["serial_number"] = mail_item["serial_number"]
        return result

    def _summary_text(self, parsed, imported, company, director):
        lines = [
            f"Заявитель: {company.get('name')}",
            f"Договор М9: {company.get('contract_number')} от {company.get('contract_date')}",
            f"Подпись: {director.get('title')} {director.get('name')}" if director else "Подпись: не выбран",
            f"Номер заявки: {self.request_data.get('request_number')}",
            f"Даты: {self.request_data.get('work_start_date')} — {self.request_data.get('work_end_date')}",
            f"Цель: {self.request_data.get('purpose')}",
        ]
        if parsed.get("client"):
            lines.append(f"Клиент из письма: {parsed['client']} (дог. {parsed.get('client_contract') or '—'})")
        lines.append(f"Помещения: {', '.join(self.request_data.get('rooms') or [])}")
        lines.append(f"Людей: {len(self.visitors_data)}")
        lines.append(f"Оборудования: {len(self.equipment_data)}")
        if imported:
            lines.append(f"Вложение: {imported.get('source_name')}")
        for note in parsed.get("notes") or []:
            lines.append("• " + note)
        stamp = self.workspace.stamp_path(director)
        lines.append("Печать: есть" if stamp else "Печать: загрузите на вкладке «Печати»")
        return "\n".join(lines)

    def _build_lines_page(self):
        page = ctk.CTkScrollableFrame(self.content, fg_color="#111827", corner_radius=18)
        self.pages["lines"] = page
        ctk.CTkLabel(page, text="Соединительные линии / протяжки", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=10, pady=(10, 4))
        ctk.CTkLabel(
            page,
            text="Письмо клиента, уведомление ДЦ или JSON-задание → заявка Ушмайкину К.Э. с таблицей сторон A/B.",
            text_color="#94A3B8", wraplength=900, justify="left",
        ).pack(anchor="w", padx=10, pady=(0, 8))
        self.line_mail = ctk.CTkTextbox(page, height=140, font=ctk.CTkFont(family="Consolas", size=13), corner_radius=12)
        self.line_mail.pack(fill="x", padx=10, pady=(0, 8))
        self.line_mail.insert("1.0", "Коллеги, добрый день!\n\nПрошу отказаться от кроссировок: 5276/25 и 5959/24\n")
        bar = ctk.CTkFrame(page, fg_color="transparent")
        bar.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(bar, text="Разобрать письмо", height=36, command=self.parse_line_mail).pack(side="left")
        ctk.CTkButton(bar, text="JSON-задание", height=36, fg_color="#1E293B", command=self.attach_line_json).pack(side="left", padx=8)
        self.line_json_label = ctk.CTkLabel(bar, text="задание не выбрано", text_color="#64748B")
        self.line_json_label.pack(side="left", padx=8)

        self.line_fields = {}
        self.line_fields["work_kind"] = ComboField(page, "Вид работ", WORK_KINDS)
        self.line_fields["tariff"] = ComboField(page, "Тариф", TARIFFS)
        self.line_fields["outgoing_number"] = FieldRow(page, "Исх. номер")
        self.line_fields["outgoing_date"] = FieldRow(page, "Дата исх. (ГГГГ-ММ-ДД)")
        self.line_fields["original_assignments"] = FieldRow(page, "Исходные задания")
        self.line_fields["object_address"] = ComboField(page, "Объект", OBJECT_ADDRESSES)
        self.line_fields["connection_type"] = ComboField(page, "Тип соединения", CONNECTION_TYPES)
        self.line_fields["quantity"] = FieldRow(page, "Количество волокон")
        self.line_fields["signer_title"] = FieldRow(page, "Должность подписанта")
        self.line_fields["signer_name"] = ComboField(page, "Подписант", ["-"])
        self.line_fields["contact_name"] = FieldRow(page, "Контактное лицо")
        self.line_fields["contact_phone"] = FieldRow(page, "Телефон")
        self.line_fields["contact_email"] = FieldRow(page, "E-mail")
        for key in (
            "work_kind", "tariff", "outgoing_number", "outgoing_date", "original_assignments",
            "object_address", "connection_type", "quantity", "signer_title", "signer_name",
            "contact_name", "contact_phone", "contact_email",
        ):
            self.line_fields[key].pack(fill="x", padx=10, pady=4)
        self.line_fields["signer_name"].combo.configure(command=self._on_line_signer_changed)

        ctk.CTkLabel(page, text="Сторона А", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(12, 4))
        self.line_side_a = self._build_side_fields(page, "a")
        ctk.CTkLabel(page, text="Сторона B", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(12, 4))
        self.line_side_b = self._build_side_fields(page, "b")
        self.line_fields["note"] = FieldRow(page, "Доп. информация")
        self.line_fields["note"].pack(fill="x", padx=10, pady=8)
        self.line_stamp_var = ctk.CTkSwitch(page, text="Печать ТЕХ.РУ на PDF", onvalue=1, offvalue=0)
        self.line_stamp_var.pack(anchor="w", padx=10, pady=(4, 16))
        self.line_stamp_var.select()
        self._refresh_line_signers()

    def _build_side_fields(self, page, prefix):
        fields = {}
        specs = [
            ("org", "Организация"),
            ("inn", "ИНН"),
            ("connector", "Разъём"),
            ("floor", "Этаж"),
            ("room", "Помещение"),
            ("row", "Ряд"),
            ("place", "Место"),
            ("tap", "Точка включения"),
        ]
        for key, label in specs:
            if key == "connector":
                row = ComboField(page, label, CONNECTORS)
            else:
                row = FieldRow(page, label)
            row.pack(fill="x", padx=10, pady=3)
            fields[key] = row
        return fields

    def _refresh_line_signers(self):
        if not getattr(self, "line_fields", None):
            return
        company = self.workspace.get_company()
        names = [s.get("name") for s in (company.get("line_signers") or [])] or ["-"]
        self.line_fields["signer_name"].combo.configure(values=names)
        current = (self.line_data or {}).get("signer_name")
        if current and current in names:
            self.line_fields["signer_name"].set(current)
        else:
            self.line_fields["signer_name"].set(names[0])

    def _on_line_signer_changed(self, name):
        company = self.workspace.get_company() or {}
        for signer in company.get("line_signers") or []:
            if signer.get("name") == name:
                self.line_fields["signer_title"].set(signer.get("title") or "Инженер")
                self.line_fields["contact_name"].set(signer.get("name") or "")
                if signer.get("email"):
                    self.line_fields["contact_email"].set(signer.get("email"))
                break

    def attach_line_json(self):
        path = filedialog.askopenfilename(title="JSON задания линии", filetypes=[("JSON / текст", "*.json *.txt"), ("Все файлы", "*.*")])
        if not path:
            return
        self.line_json_path = path
        self.line_json_label.configure(text=Path(path).name, text_color="#93C5FD")
        self.parse_line_mail()

    def parse_line_mail(self):
        text = self.line_mail.get("1.0", "end")
        assignment = None
        if self.line_json_path:
            try:
                assignment = json.loads(Path(self.line_json_path).read_text(encoding="utf-8-sig"))
            except Exception as exc:
                messagebox.showerror("JSON", f"Не удалось прочитать задание:\n{exc}")
                return
        parsed = LineParser().parse(text, assignment)
        company = self.workspace.get_company()
        apply_line_company_defaults(parsed, company)
        if not parsed.get("outgoing_number"):
            parsed["outgoing_number"] = format_outgoing_number(
                company.get("request_prefix") or "04-08",
                (company.get("last_request_seq") or 0) + 1,
            )
        if parsed.get("work_kind") == "монтаж":
            self.line_stamp_var.select()
        else:
            self.line_stamp_var.deselect()
        self.line_data = parsed
        self._fill_line_fields()
        orig = ", ".join(parsed.get("original_assignments") or [])
        messagebox.showinfo(
            "Протяжка",
            f"{parsed.get('work_kind')} · {parsed.get('outgoing_number')}\n"
            f"Задания: {orig or '—'}\n"
            f"A: {parsed.get('side_a', {}).get('room')} ряд {parsed.get('side_a', {}).get('row')} "
            f"B: {parsed.get('side_b', {}).get('org') or parsed.get('side_b', {}).get('room')}",
        )

    def _fill_line_fields(self):
        if not getattr(self, "line_fields", None):
            return
        company = self.workspace.get_company()
        data = dict(self.line_data or empty_line_request())
        apply_line_company_defaults(data, company)
        self.line_data = data
        self._refresh_line_signers()
        mapping = {
            "work_kind": data.get("work_kind"),
            "tariff": data.get("tariff"),
            "outgoing_number": data.get("outgoing_number"),
            "outgoing_date": data.get("outgoing_date"),
            "original_assignments": ", ".join(data.get("original_assignments") or []),
            "object_address": data.get("object_address"),
            "connection_type": data.get("connection_type"),
            "quantity": data.get("quantity"),
            "signer_title": data.get("signer_title"),
            "signer_name": data.get("signer_name"),
            "contact_name": data.get("contact_name"),
            "contact_phone": data.get("contact_phone"),
            "contact_email": data.get("contact_email"),
            "note": data.get("note"),
        }
        for key, value in mapping.items():
            if key in self.line_fields:
                self.line_fields[key].set(value or "")
        for store, fields in ((data.get("side_a") or {}, self.line_side_a), (data.get("side_b") or {}, self.line_side_b)):
            for key, row in fields.items():
                row.set(store.get(key) or "")
        if data.get("use_stamp"):
            self.line_stamp_var.select()
        else:
            self.line_stamp_var.deselect()

    def _collect_line_data(self):
        data = dict(self.line_data or empty_line_request())
        for key, row in self.line_fields.items():
            data[key] = row.get()
        data["original_assignments"] = [part.strip() for part in str(data.get("original_assignments") or "").replace(";", ",").split(",") if part.strip()]
        data["side_a"] = {key: row.get() for key, row in self.line_side_a.items()}
        data["side_b"] = {key: row.get() for key, row in self.line_side_b.items()}
        data["use_stamp"] = bool(self.line_stamp_var.get())
        company = self.workspace.get_company()
        apply_line_company_defaults(data, company)
        self.line_data = data
        return data

    def _main_action(self):
        if self.current_page == "lines":
            self.create_line_docs()
        else:
            self.create_all()

    def create_line_docs(self):
        try:
            if "lines" not in self.pages:
                self._build_lines_page()
            data = self._collect_line_data()
            company = self.workspace.get_company()
            number = (data.get("outgoing_number") or "").strip()
            peek = format_outgoing_number(company.get("request_prefix") or "04-08", (company.get("last_request_seq") or 0) + 1)
            official = number.replace("-", "") == peek.replace("-", "") or not number
            if official:
                self.workspace.consume_request_number(company)
                data["outgoing_number"] = format_outgoing_number(
                    company.get("request_prefix") or "04-08",
                    company.get("last_request_seq") or 0,
                )
                self.line_fields["outgoing_number"].set(data["outgoing_number"])
            stamp = line_stamp_path(self.workspace.get_director()) if data.get("use_stamp") else None
            paths = export_line_bundle(data, stamp)
            messagebox.showinfo("Протяжка", f"PDF: {paths['pdf']}\nJSON: {paths['json']}")
            self._open_folder(paths["pdf"])
        except Exception as exc:
            messagebox.showerror("Протяжка", str(exc))

    # ── data page ───────────────────────────────────────────
    def _build_data_page(self):
        page = ctk.CTkScrollableFrame(self.content, fg_color="#111827", corner_radius=18)
        self.pages["data"] = page
        ctk.CTkLabel(page, text="Реквизиты заявки", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=8, pady=(8, 4))
        ctk.CTkLabel(
            page,
            text="Помещения и даты берутся из тех же справочников, что и в официальном Excel ММТС-9.",
            text_color="#94A3B8",
        ).pack(anchor="w", padx=8, pady=(0, 12))
        self.fields = {}

        self.fields["object_address"] = ComboField(page, "Объект", self.catalog.get("addresses") or OBJECT_ADDRESSES)
        self.fields["object_address"].pack(fill="x", padx=8, pady=5)
        self.fields["object_address"].combo.configure(command=self._on_address_changed)

        self.fields["organization"] = FieldRow(page, "Организация")
        self.fields["inn"] = FieldRow(page, "ИНН")
        self.fields["contract_number"] = FieldRow(page, "Договор с ММТС-9")
        self.fields["contract_date"] = DateField(page, "Дата договора")
        self.fields["contract_valid_until"] = DateField(page, "Срок действия договора")
        for key in ("organization", "inn", "contract_number", "contract_date", "contract_valid_until"):
            self.fields[key].pack(fill="x", padx=8, pady=5)

        self.fields["rooms"] = RoomPicker(page, self.rooms, self)
        self.fields["rooms"].pack(fill="x", padx=8, pady=8)

        self.fields["work_description"] = FieldRow(page, "Стойки / описание работ")
        self.fields["work_description"].pack(fill="x", padx=8, pady=5)

        period = ctk.CTkFrame(page, fg_color="transparent")
        period.pack(fill="x", padx=8, pady=(8, 4))
        ctk.CTkLabel(period, text="Период работ", text_color="#94A3B8", width=210, anchor="w").pack(side="left")
        for title, days in (("Сегодня", 0), ("+3 дня", 3), ("+7 дней", 7), ("+14 дней", 14)):
            ctk.CTkButton(
                period, text=title, width=80, height=28, fg_color="#1E293B",
                command=lambda d=days: self._set_work_period(d),
            ).pack(side="left", padx=3)

        self.fields["work_start_date"] = DateField(page, "Дата начала")
        self.fields["work_end_date"] = DateField(page, "Дата окончания")
        self.fields["work_start_date"].pack(fill="x", padx=8, pady=5)
        self.fields["work_end_date"].pack(fill="x", padx=8, pady=5)

        self.fields["purpose"] = ComboField(page, "Цель визита", self.catalog.get("purposes") or PURPOSES)
        self.fields["purpose"].pack(fill="x", padx=8, pady=5)
        self.fields["responsible_person"] = FieldRow(page, "Должность подписанта")
        self.fields["responsible_name"] = FieldRow(page, "ФИО подписанта")
        self.fields["request_number"] = FieldRow(page, "Номер заявки")
        self.fields["request_date"] = DateField(page, "Дата заявки")
        self.fields["contact_phone"] = FieldRow(page, "Телефон")
        for key in ("responsible_person", "responsible_name", "request_number", "request_date", "contact_phone"):
            self.fields[key].pack(fill="x", padx=8, pady=5)

    def _on_address_changed(self, address=None):
        if self._menu_lock:
            return
        address = address or self.fields["object_address"].get()
        self.rooms = load_rooms(address)
        self.parser = EmailParser(self.rooms)
        self.fields["rooms"].set_catalog(self.rooms)
        current = [room for room in self.fields["rooms"].get() if room in self.rooms]
        if not current:
            current = [default_access_room(address)]
        self.fields["rooms"].set(current)

    def _set_work_period(self, extra_days):
        start = datetime.now()
        end = start + timedelta(days=extra_days)
        self.fields["work_start_date"].set(start.strftime("%d.%m.%Y"))
        self.fields["work_end_date"].set(end.strftime("%d.%m.%Y"))

    def _fill_data_fields(self, data):
        if not getattr(self, "fields", None):
            return
        self._menu_lock = True
        try:
            address = data.get("object_address") or OBJECT_ADDRESSES[0]
            if "object_address" in self.fields:
                self.fields["object_address"].set(address)
                self.rooms = load_rooms(address)
                self.fields["rooms"].set_catalog(self.rooms)
            for key, row in self.fields.items():
                value = data.get(key, "")
                if key == "object_address":
                    continue
                row.set(value)
        finally:
            self._menu_lock = False

    def _collect_request_data(self):
        if not getattr(self, "fields", None):
            self._build_data_page()
            self._fill_data_fields(self.request_data)
        data = {key: row.get() for key, row in self.fields.items()}
        rooms = data.get("rooms") or []
        if isinstance(rooms, str):
            rooms = [part.strip() for part in rooms.split(",") if part.strip()]
        default = default_access_room(data.get("object_address"))
        if default not in rooms:
            rooms.insert(0, default)
        data["rooms"] = rooms[:6]
        self.request_data = data
        return data

    def _load_company_into_form(self, reset_people=False):
        company = self.workspace.get_company()
        director = self.workspace.get_director()
        data = dict(self.request_data)
        apply_company_defaults(data, company, director, data.get("request_date"))
        if not data.get("request_number"):
            data["request_number"] = self.workspace.peek_request_number(company)
        if not data.get("work_start_date"):
            data["work_start_date"] = datetime.now().strftime("%d.%m.%Y")
        if not data.get("work_end_date"):
            data["work_end_date"] = data["work_start_date"]
        if not data.get("purpose"):
            data["purpose"] = company.get("purpose_default") or "Работа с оборудованием"
        if not data.get("object_address"):
            data["object_address"] = OBJECT_ADDRESSES[0]
        if not data.get("rooms"):
            data["rooms"] = [default_access_room(data.get("object_address"))]
        self.request_data = data
        self._fill_data_fields(data)
        if reset_people:
            self.visitors_data = self.workspace.default_team(company)
            self._render_people()
        self._refresh_director_menu()

    def _on_company_changed(self, name):
        if self._menu_lock:
            return
        for company in self.workspace.companies():
            if (company.get("short_name") or company.get("name")) == name:
                self.workspace.set_active_company(company["id"])
                break
        self._load_company_into_form(reset_people=True)
        self._refresh_line_signers()
        if getattr(self, "line_fields", None):
            self._fill_line_fields()

    def _on_director_changed(self, name):
        if self._menu_lock:
            return
        company = self.workspace.get_company()
        for director in company.get("directors", []):
            if director.get("name") == name:
                self.workspace.set_active_director(director["id"])
                if getattr(self, "fields", None):
                    self.fields["responsible_person"].set(director.get("title") or "Генеральный директор")
                    self.fields["responsible_name"].set(director.get("name") or "")
                break
        self._refresh_director_menu()

    # ── people / equipment / vehicles ───────────────────────
    def _build_people_page(self):
        page = ctk.CTkFrame(self.content, fg_color="#111827", corner_radius=18)
        self.pages["people"] = page
        bar = ctk.CTkFrame(page, fg_color="transparent")
        bar.pack(fill="x", padx=16, pady=12)
        ctk.CTkLabel(bar, text="Посетители", font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")
        ctk.CTkButton(bar, text="Добавить", width=110, command=lambda: self._edit_person()).pack(side="right", padx=4)
        ctk.CTkButton(bar, text="Из базы", width=110, fg_color="#1E293B", command=self._pick_from_pool).pack(side="right", padx=4)
        ctk.CTkButton(bar, text="Команда компании", width=170, fg_color="#1E293B", command=self._add_company_team).pack(side="right", padx=4)
        self.people_list = ctk.CTkScrollableFrame(page, fg_color="transparent")
        self.people_list.pack(fill="both", expand=True, padx=10, pady=(0, 12))

    def _render_people(self):
        if not hasattr(self, "people_list"):
            return
        for child in self.people_list.winfo_children():
            child.destroy()
        if not self.visitors_data:
            ctk.CTkLabel(self.people_list, text="Пока никого нет", text_color="#64748B").pack(pady=20)
            return
        for idx, person in enumerate(self.visitors_data):
            self._card(
                self.people_list,
                f"{person.get('surname', '')} {person.get('name', '')} {person.get('patronymic', '')}",
                f"{person.get('birth_date', '')} · паспорт {person.get('passport_series', '')} {person.get('passport_number', '')}",
                on_edit=lambda i=idx: self._edit_person(i),
                on_delete=lambda i=idx: self._delete_item("visitors_data", i, self._render_people),
            )

    def _build_equipment_page(self):
        page = ctk.CTkFrame(self.content, fg_color="#111827", corner_radius=18)
        self.pages["gear"] = page
        bar = ctk.CTkFrame(page, fg_color="transparent")
        bar.pack(fill="x", padx=16, pady=12)
        ctk.CTkLabel(bar, text="Оборудование", font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")
        ctk.CTkButton(bar, text="Добавить", width=110, command=lambda: self._edit_equipment()).pack(side="right")
        self.eq_list = ctk.CTkScrollableFrame(page, fg_color="transparent")
        self.eq_list.pack(fill="both", expand=True, padx=10, pady=(0, 12))

    def _render_equipment(self):
        if not hasattr(self, "eq_list"):
            return
        for child in self.eq_list.winfo_children():
            child.destroy()
        if not self.equipment_data:
            ctk.CTkLabel(self.eq_list, text="Оборудования нет", text_color="#64748B").pack(pady=20)
            return
        for idx, item in enumerate(self.equipment_data):
            title = " ".join(filter(None, [item.get("name"), item.get("brand"), item.get("model")]))
            meta = " · ".join(filter(None, [
                item.get("serial_number"), item.get("action"), item.get("room"),
                f"ряд {item.get('row')}" if item.get("row") else "",
                f"место {item.get('place')}" if item.get("place") else "",
            ]))
            self._card(
                self.eq_list, title or "Позиция", meta,
                on_edit=lambda i=idx: self._edit_equipment(i),
                on_delete=lambda i=idx: self._delete_item("equipment_data", i, self._render_equipment),
            )

    def _build_vehicles_page(self):
        page = ctk.CTkFrame(self.content, fg_color="#111827", corner_radius=18)
        self.pages["car"] = page
        bar = ctk.CTkFrame(page, fg_color="transparent")
        bar.pack(fill="x", padx=16, pady=12)
        ctk.CTkLabel(bar, text="Автотранспорт", font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")
        ctk.CTkButton(bar, text="Добавить", width=110, command=lambda: self._edit_vehicle()).pack(side="right")
        self.car_list = ctk.CTkScrollableFrame(page, fg_color="transparent")
        self.car_list.pack(fill="both", expand=True, padx=10, pady=(0, 12))

    def _render_vehicles(self):
        if not hasattr(self, "car_list"):
            return
        for child in self.car_list.winfo_children():
            child.destroy()
        if not self.vehicles_data:
            ctk.CTkLabel(self.car_list, text="Машин нет — это нормально", text_color="#64748B").pack(pady=20)
            return
        for idx, item in enumerate(self.vehicles_data):
            title = item.get("license_plate") or "Авто"
            meta = " ".join(filter(None, [item.get("driver_surname"), item.get("driver_name"), item.get("driver_patronymic")]))
            self._card(
                self.car_list, title, meta,
                on_edit=lambda i=idx: self._edit_vehicle(i),
                on_delete=lambda i=idx: self._delete_item("vehicles_data", i, self._render_vehicles),
            )

    def _card(self, parent, title, subtitle, on_edit, on_delete):
        card = ctk.CTkFrame(parent, fg_color="#0F172A", corner_radius=12)
        card.pack(fill="x", pady=5, padx=6)
        text = ctk.CTkFrame(card, fg_color="transparent")
        text.pack(side="left", fill="x", expand=True, padx=12, pady=10)
        ctk.CTkLabel(text, text=title, anchor="w", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        ctk.CTkLabel(text, text=subtitle, anchor="w", text_color="#94A3B8").pack(anchor="w")
        ctk.CTkButton(card, text="Изменить", width=90, fg_color="#1E293B", command=on_edit).pack(side="right", padx=(0, 8), pady=10)
        ctk.CTkButton(card, text="Удалить", width=80, fg_color="#7F1D1D", hover_color="#991B1B", command=on_delete).pack(side="right", padx=8, pady=10)

    def _delete_item(self, attr, index, refresh):
        items = getattr(self, attr)
        if 0 <= index < len(items):
            del items[index]
            refresh()

    def _edit_person(self, index=None):
        data = dict(self.visitors_data[index]) if index is not None else {"citizenship": ""}
        fields = [
            ("surname", "Фамилия"),
            ("name", "Имя"),
            ("patronymic", "Отчество"),
            ("birth_date", "Дата рождения", "date"),
            ("citizenship", "Гражданство (для иностранцев)"),
            ("passport_series", "Серия паспорта"),
            ("passport_number", "Номер паспорта"),
            ("passport_issue_date", "Дата выдачи", "date"),
            ("passport_issued_by", "Кем выдан"),
            ("organization", "Организация"),
            ("electrical_safety_group", "Группа ЭБ"),
            ("electrical_safety_certificate", "Удостоверение ЭБ"),
            ("note", "Примечание"),
        ]
        self._form_dialog("Посетитель", fields, data, lambda payload: self._save_list_item("visitors_data", index, payload, self._render_people))

    def _edit_equipment(self, index=None):
        data = dict(self.equipment_data[index]) if index is not None else {"quantity": "1", "unit": "шт.", "action": "Внос"}
        fields = [
            ("name", "Наименование"),
            ("brand", "Марка"),
            ("model", "Модель"),
            ("serial_number", "Серийный / заводской номер"),
            ("quantity", "Количество"),
            ("unit", "Единицы", "combo", self.catalog.get("units") or ["шт."]),
            ("action", "Внос / вынос ТМЦ", "combo", self.catalog.get("actions") or ACTIONS),
            ("note", "Примечание"),
            ("room", "Помещение", "combo", self.rooms),
            ("reason", "Причина отсутствия с/н", "combo", ["—"] + (self.catalog.get("reasons") or [])),
            ("row", "Ряд"),
            ("place", "Место"),
        ]
        self._form_dialog("Оборудование", fields, data, lambda payload: self._save_list_item("equipment_data", index, payload, self._after_equipment_change))

    def _after_equipment_change(self):
        self._render_equipment()
        if not getattr(self, "fields", None):
            return
        rooms = collect_rooms(
            self.equipment_data,
            self.fields["rooms"].get(),
            self.fields["object_address"].get() if "object_address" in self.fields else None,
        )
        self.fields["rooms"].set(rooms)
        self.fields["work_description"].set(build_work_description(self.equipment_data))

    def _edit_vehicle(self, index=None):
        data = dict(self.vehicles_data[index]) if index is not None else {"driver_citizenship": ""}
        fields = [
            ("license_plate", "Гос. номер"),
            ("driver_surname", "Фамилия"),
            ("driver_name", "Имя"),
            ("driver_patronymic", "Отчество"),
            ("driver_birth_date", "Дата рождения", "date"),
            ("driver_citizenship", "Гражданство (для иностранцев)"),
            ("passport_series", "Серия паспорта"),
            ("passport_number", "Номер паспорта"),
            ("passport_issue_date", "Дата выдачи", "date"),
            ("passport_issued_by", "Кем выдан"),
            ("driver_organization", "Организация водителя"),
        ]
        self._form_dialog("Автотранспорт", fields, data, lambda payload: self._save_list_item("vehicles_data", index, payload, self._render_vehicles))

    def _save_list_item(self, attr, index, payload, refresh):
        items = getattr(self, attr)
        if index is None:
            items.append(payload)
        else:
            items[index] = payload
        refresh()

    def _form_dialog(self, title, fields, data, on_save):
        OverlayForm(self, title, fields, data, on_save)

    def _room_picker_dialog(self, picker, prefill=""):
        OverlayRoomList(self, picker, prefill)

    def _add_company_team(self):
        team = self.workspace.default_team()
        self.visitors_data = merge_unique_visitors(self.visitors_data, team)
        self._render_people()

    def _pick_from_pool(self):
        OverlayPool(self)

    # ── stamps page ─────────────────────────────────────────
    def _build_stamps_page(self):
        page = ctk.CTkFrame(self.content, fg_color="#111827", corner_radius=18)
        self.pages["stamp"] = page
        bar = ctk.CTkFrame(page, fg_color="transparent")
        bar.pack(fill="x", padx=16, pady=12)
        ctk.CTkLabel(bar, text="Гендиректора и печати", font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")
        ctk.CTkButton(bar, text="Добавить подписанта", command=self._edit_director).pack(side="right")
        ctk.CTkButton(bar, text="Реквизиты компании", fg_color="#1E293B", command=self._edit_company).pack(side="right", padx=8)
        self.dir_list = ctk.CTkScrollableFrame(page, fg_color="transparent")
        self.dir_list.pack(fill="both", expand=True, padx=10, pady=(0, 12))

    def _render_directors(self):
        for child in self.dir_list.winfo_children():
            child.destroy()
        company = self.workspace.get_company()
        if not company:
            return
        for director in company.get("directors", []):
            card = ctk.CTkFrame(self.dir_list, fg_color="#0F172A", corner_radius=12)
            card.pack(fill="x", pady=6, padx=6)
            stamp = self.workspace.stamp_path(director)
            title = f"{director.get('title') or 'Генеральный директор'} · {director.get('name')}"
            meta = "печать загружена" if stamp else "печати нет — нажмите «Файл печати»"
            text = ctk.CTkFrame(card, fg_color="transparent")
            text.pack(side="left", fill="x", expand=True, padx=12, pady=10)
            ctk.CTkLabel(text, text=title, font=ctk.CTkFont(weight="bold")).pack(anchor="w")
            ctk.CTkLabel(text, text=meta, text_color="#34D399" if stamp else "#F59E0B").pack(anchor="w")
            ctk.CTkButton(card, text="Выбрать", width=90, command=lambda d=director: self._select_director(d)).pack(side="right", padx=6, pady=10)
            ctk.CTkButton(card, text="Файл печати", width=120, fg_color="#1E293B", command=lambda d=director: self._upload_stamp(d)).pack(side="right", padx=6, pady=10)
            ctk.CTkButton(card, text="Изменить", width=90, fg_color="#1E293B", command=lambda d=director: self._edit_director(d)).pack(side="right", padx=6, pady=10)

    def _select_director(self, director):
        self.workspace.set_active_director(director["id"])
        self._refresh_director_menu()
        if getattr(self, "fields", None):
            self.fields["responsible_person"].set(director.get("title") or "Генеральный директор")
            self.fields["responsible_name"].set(director.get("name") or "")
        self._render_directors()

    def _upload_stamp(self, director):
        path = filedialog.askopenfilename(title="Печать", filetypes=[("Изображения", "*.png *.jpg *.jpeg *.webp")])
        if not path:
            return
        self.workspace.save_stamp(path, director)
        self._refresh_director_menu()
        self._render_directors()
        messagebox.showinfo("Печать", f"Печать сохранена для {director.get('name')}")

    def _edit_director(self, director=None):
        data = dict(director or {"title": "Генеральный директор"})

        def save(payload):
            payload["id"] = data.get("id", "")
            payload["stamp_file"] = data.get("stamp_file", "")
            if not payload.get("title"):
                payload["title"] = "Генеральный директор"
            self.workspace.upsert_director(self.workspace.get_company()["id"], payload)
            self._refresh_director_menu()
            self._render_directors()

        OverlayForm(self, "Подписант", [("title", "Должность"), ("name", "ФИО")], data, save)

    def _edit_company(self):
        company = dict(self.workspace.get_company() or {})
        specs = [
            ("short_name", "Короткое имя"),
            ("name", "Наименование в заявке"),
            ("inn", "ИНН"),
            ("contract_number", "Договор с ММТС-9"),
            ("contract_date", "Дата договора", "date"),
            ("contract_valid_until", "Срок действия", "date"),
            ("phone", "Телефон"),
            ("phone_ext", "Добавочный"),
            ("request_prefix", "Префикс номера заявки"),
            ("purpose_default", "Цель по умолчанию", "combo", self.catalog.get("purposes") or PURPOSES),
        ]

        def save(payload):
            company.update(payload)
            self.workspace.upsert_company(company)
            self._refresh_company_menus()
            self._load_company_into_form()

        OverlayForm(self, "Реквизиты компании-заявителя", specs, company, save)

    def _build_help_page(self):
        page = ctk.CTkScrollableFrame(self.content, fg_color="#111827", corner_radius=18)
        self.pages["help"] = page
        for title, body in HELP_SECTIONS:
            ctk.CTkLabel(page, text=title, font=ctk.CTkFont(size=18, weight="bold"), text_color="#93C5FD").pack(anchor="w", padx=16, pady=(16, 6))
            ctk.CTkLabel(page, text=body, justify="left", anchor="w", wraplength=860, text_color="#CBD5E1").pack(anchor="w", padx=16, pady=(0, 8))
        ctk.CTkButton(
            page, text="Написать разработчику в Telegram", height=40,
            fg_color="#1D4ED8", hover_color="#1E40AF",
            command=lambda: webbrowser.open(APP_DEVELOPER_URL),
        ).pack(anchor="w", padx=16, pady=(8, 24))

    # ── export ──────────────────────────────────────────────
    def _prepare_export(self, consume_number=True):
        data = self._collect_request_data()
        if not data.get("organization") or not data.get("inn"):
            raise ValueError("Заполните организацию и ИНН")
        company = self.workspace.get_company()
        director = self.workspace.get_director()
        apply_company_defaults(data, company, director, data.get("request_date"))
        if consume_number:
            data["request_number"] = self.workspace.consume_request_number(company, data.get("request_number"))
            if getattr(self, "fields", None) and "request_number" in self.fields:
                self.fields["request_number"].set(data["request_number"])
        elif not data.get("request_number"):
            data["request_number"] = self.workspace.peek_request_number(company)
        self.request_data = data
        return data, director

    def create_excel_only(self):
        try:
            data, _director = self._prepare_export()
            names = suggested_filenames(data)
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=names["xlsx"].name, filetypes=[("Excel", "*.xlsx")])
            if not path:
                return
            export_excel(data, self.visitors_data, self.equipment_data, self.vehicles_data, path)
            messagebox.showinfo("Excel", f"Файл сохранён:\n{path}")
            self._open_folder(path)
        except Exception as exc:
            messagebox.showerror("Excel", str(exc))

    def create_pdf_only(self):
        try:
            data, director = self._prepare_export()
            names = suggested_filenames(data)
            path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=names["pdf"].name, filetypes=[("PDF", "*.pdf")])
            if not path:
                return
            stamp = self.workspace.stamp_path(director)
            if not stamp:
                if not messagebox.askyesno("Печать", "У выбранного гендиректора нет файла печати. Сохранить PDF без печати?"):
                    return
            export_pdf(data, self.visitors_data, self.equipment_data, self.vehicles_data, stamp, path)
            messagebox.showinfo("PDF", f"Файл сохранён:\n{path}")
            self._open_folder(path)
        except Exception as exc:
            messagebox.showerror("PDF", str(exc))

    def create_all(self):
        try:
            data, director = self._prepare_export()
            names = suggested_filenames(data)
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            xlsx_path = names["xlsx"]
            pdf_path = names["pdf"]
            export_excel(data, self.visitors_data, self.equipment_data, self.vehicles_data, xlsx_path)
            stamp = self.workspace.stamp_path(director)
            export_pdf(data, self.visitors_data, self.equipment_data, self.vehicles_data, stamp, pdf_path)
            extra = "" if stamp else "\n\nПечать не подставлена — загрузите её на вкладке «Печати»."
            messagebox.showinfo(
                "Готово",
                f"Excel: {xlsx_path}\nPDF: {pdf_path}{extra}",
            )
            self._open_folder(xlsx_path)
        except Exception as exc:
            messagebox.showerror("Выпуск", str(exc))

    def _startup_update_flow(self):
        self._show_update_notice()
        self._schedule_update_check()

    def _show_update_notice(self):
        from updater import consume_update_notice
        notice = consume_update_notice()
        if not notice:
            return
        target = notice.get("to") or APP_VERSION
        previous = notice.get("from") or ""
        extra = f"\nПредыдущая сборка: {previous}" if previous else ""
        messagebox.showinfo(
            "Обновление установлено",
            f"M9 Gate обновлён до {target}.{extra}",
        )

    def _schedule_update_check(self):
        from updater import start_background_check
        start_background_check(lambda release: self.after(0, lambda: self._prompt_update(release)))

    def _prompt_update(self, release):
        tag = release.get("tag") or ""
        if not messagebox.askyesno(
            "Обновление",
            f"Доступна новая сборка {tag} (сейчас {APP_VERSION}).\n\nСкачать и установить?",
        ):
            return
        try:
            from updater import apply_update
            apply_update(release)
            messagebox.showinfo("Обновление", "Загрузка завершена. Приложение перезапустится.")
            self.destroy()
        except Exception as exc:
            messagebox.showerror("Обновление", str(exc))

    def _open_folder(self, path):
        folder = str(Path(path).resolve().parent)
        try:
            os.startfile(folder)
        except Exception:
            pass


def main():
    app = PassApp()
    app.mainloop()


if __name__ == "__main__":
    main()
