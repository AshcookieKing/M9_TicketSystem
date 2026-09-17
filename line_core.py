"""Заявки на протяжки / соединительные линии ММТС-9."""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from pass_core import OBJECT_ADDRESSES, OUTPUT_DIR, STAMPS_DIR, draw_round_stamp, os_environ_windir

CONNECTION_TYPES = ["оптика SM", "оптика MM", "медь UTP", "медь FTP"]
TARIFFS = ["стандартный тариф", "срочный тариф"]
CONNECTORS = ["LC/UPC", "LC/APC", "SC/UPC", "SC/APC", "FC/UPC", "ST", "RJ45"]
WORK_KINDS = ["монтаж", "демонтаж"]


def empty_side():
    return {
        "org": "",
        "inn": "",
        "connector": "LC/UPC",
        "floor": "",
        "room": "",
        "row": "",
        "place": "",
        "tap": "",
        "notes": "",
    }


def empty_line_request():
    return {
        "work_kind": "монтаж",
        "tariff": "стандартный тариф",
        "object_address": OBJECT_ADDRESSES[0],
        "outgoing_number": "",
        "outgoing_date": datetime.now().strftime("%Y-%m-%d"),
        "original_assignments": [],
        "connection_type": "оптика SM",
        "quantity": "1",
        "note": "",
        "client": "",
        "client_contract": "",
        "contact_name": "",
        "contact_phone": "",
        "contact_email": "",
        "signer_title": "Инженер",
        "signer_name": "",
        "org_name": "",
        "org_inn": "",
        "side_a": empty_side(),
        "side_b": empty_side(),
        "use_stamp": True,
    }


def format_outgoing_number(prefix, seq):
    compact = re.sub(r"[^\w]", "", str(prefix or "0408"))
    return f"{compact}.{seq}"


def apply_line_company_defaults(data, company, signer=None):
    data["org_name"] = company.get("name") or data.get("org_name") or ""
    data["org_inn"] = company.get("inn") or data.get("org_inn") or ""
    phone = company.get("phone") or ""
    ext = company.get("phone_ext") or "1"
    if phone and not data.get("contact_phone"):
        pretty = phone if phone.startswith("+") else f"+7 ({phone[1:4]}) {phone[4:]}" if len(phone) >= 11 and phone.startswith("8") else phone
        if phone.startswith("8495") and len(phone) == 11:
            pretty = f"+7 (495) {phone[4:]}"
        elif phone.startswith("8") and len(phone) == 11:
            pretty = f"+7 ({phone[1:4]}) {phone[4:]}"
        data["contact_phone"] = f"{pretty} доб.{ext}" if ext else pretty
    if not data.get("object_address"):
        data["object_address"] = OBJECT_ADDRESSES[0]
    signers = company.get("line_signers") or []
    signer = signer or (signers[0] if signers else None)
    if signer:
        if not data.get("signer_name"):
            data["signer_name"] = signer.get("name") or ""
        if not data.get("signer_title"):
            data["signer_title"] = signer.get("title") or "Инженер"
        if not data.get("contact_name"):
            data["contact_name"] = signer.get("name") or ""
        if not data.get("contact_email"):
            data["contact_email"] = signer.get("email") or ""
    if not data.get("outgoing_date"):
        data["outgoing_date"] = datetime.now().strftime("%Y-%m-%d")
    if not data.get("side_a", {}).get("org"):
        data.setdefault("side_a", empty_side())
        data["side_a"]["org"] = data["org_name"]
        data["side_a"]["inn"] = data["org_inn"]
    return data


def assignment_from_json(raw: dict) -> dict:
    data = empty_line_request()
    company = raw.get("company") or {}
    line = raw.get("line") or {}
    side_a = raw.get("line_А") or raw.get("line_A") or {}
    side_b = raw.get("line_B") or raw.get("line_В") or {}
    kind = (company.get("Вид_работ") or "монтаж").strip().lower()
    data["work_kind"] = "демонтаж" if "демонт" in kind else "монтаж"
    tariff = (company.get("Срочность") or "").lower()
    data["tariff"] = "срочный тариф" if "сроч" in tariff else "стандартный тариф"
    data["org_name"] = company.get("Наименование") or ""
    data["org_inn"] = company.get("ИНН") or ""
    data["signer_title"] = company.get("Должность") or "Инженер"
    data["signer_name"] = company.get("ФИО") or ""
    data["outgoing_number"] = str(company.get("Номер_исходящего_письма") or "").strip()
    data["outgoing_date"] = str(company.get("Дата_исходящего_письма") or "").strip()
    data["contact_name"] = company.get("Контактное_лицо") or data["signer_name"]
    data["contact_phone"] = company.get("Контактный_телефон") or ""
    data["contact_email"] = company.get("Контактный_email") or ""
    data["object_address"] = company.get("Объект") or OBJECT_ADDRESSES[0]
    orig = (company.get("Номер_исходного_задания") or "").strip()
    data["original_assignments"] = [part.strip() for part in re.split(r"[,;]+", orig) if part.strip()]
    data["note"] = company.get("Примечание") or ""
    data["connection_type"] = line.get("Тип_соединения") or "оптика SM"
    data["quantity"] = str(line.get("Количество") or "1")
    data["side_a"] = _side_from_json(side_a)
    data["side_b"] = _side_from_json(side_b)
    data["use_stamp"] = data["work_kind"] == "монтаж"
    return data


def _side_from_json(block):
    side = empty_side()
    side["org"] = block.get("Наименование_организации") or ""
    side["inn"] = block.get("ИНН") or ""
    side["connector"] = block.get("Тип_разъема") or ""
    side["floor"] = str(block.get("Этаж") or "").strip()
    side["room"] = str(block.get("Помещение") or "").strip()
    side["row"] = str(block.get("Ряд") or "").strip()
    side["place"] = str(block.get("Место") or "").strip()
    side["notes"] = block.get("Примечания") or ""
    return side


def to_assignment_json(data: dict) -> dict:
    orig = data.get("original_assignments") or []
    if isinstance(orig, str):
        orig = [part.strip() for part in re.split(r"[,;]+", orig) if part.strip()]
    return {
        "company": {
            "Наименование": data.get("org_name") or "",
            "ИНН": data.get("org_inn") or "",
            "Номер_договора": "",
            "Дата_договора": "",
            "Должность": data.get("signer_title") or "",
            "ФИО": data.get("signer_name") or "",
            "Номер_исходящего_письма": data.get("outgoing_number") or "",
            "Дата_исходящего_письма": data.get("outgoing_date") or "",
            "Контактное_лицо": data.get("contact_name") or "",
            "Контактный_телефон": data.get("contact_phone") or "",
            "Контактный_email": data.get("contact_email") or "",
            "Объект": data.get("object_address") or "",
            "Сторона": "заказчик",
            "Вид_работ": data.get("work_kind") or "монтаж",
            "Срочность": data.get("tariff") or "стандартный тариф",
            "Номер_исходного_задания": ", ".join(orig),
            "Входящий_номер_заявки_заказчика": "",
            "Примечание": data.get("note") or "",
        },
        "line": {
            "Тип_соединения": data.get("connection_type") or "оптика SM",
            "Количество": str(data.get("quantity") or "1"),
        },
        "line_А": _side_to_json(data.get("side_a") or {}),
        "line_B": _side_to_json(data.get("side_b") or {}),
    }


def _side_to_json(side):
    return {
        "Наименование_организации": side.get("org") or "",
        "ИНН": side.get("inn") or "",
        "Тип_разъема": side.get("connector") or "",
        "Этаж": side.get("floor") or "",
        "Помещение": side.get("room") or "",
        "Ряд": side.get("row") or "",
        "Место": side.get("place") or "",
        "Примечания": side.get("notes") or "",
    }


def _fill(target, key, value):
    if value and not (target.get(key) or "").strip():
        target[key] = str(value).strip()


class LineParser:
    def parse(self, text: str, assignment: dict | None = None) -> dict:
        data = assignment_from_json(assignment) if assignment else empty_line_request()
        joined = (text or "").replace("\xa0", " ")
        has_json = assignment is not None

        refuse = re.search(r"отказаться|расформир|кроссиров", joined, re.IGNORECASE)
        if not has_json:
            if refuse or re.search(r"демонтаж", joined, re.IGNORECASE):
                data["work_kind"] = "демонтаж"
                data["use_stamp"] = False
            elif re.search(r"монтаж|протяж|соединительн|уведомление о передаче", joined, re.IGNORECASE):
                data["work_kind"] = "монтаж"
                data["use_stamp"] = True
        else:
            data["use_stamp"] = data.get("work_kind") == "монтаж"

        numbers = [re.sub(r"\s+", "", n) for n in re.findall(r"\b\d{3,5}\s*/\s*\d{2}\b", joined)]
        existing = list(data.get("original_assignments") or [])
        for item in numbers:
            if item not in existing:
                existing.append(item)
        data["original_assignments"] = existing

        inv = re.search(r"инв\.?\s*номер\s*[-–:]\s*([A-Za-zА-Яа-я]?\d{3,6})", joined, re.IGNORECASE)
        if inv:
            _fill(data, "outgoing_number", inv.group(1))
        out_letter = re.search(r"исх\.?\s*письм[ао][^\n]{0,24}[-–:]\s*([A-Za-zА-Яа-я]?\d{3,6})", joined, re.IGNORECASE)
        if out_letter:
            _fill(data, "outgoing_number", out_letter.group(1))

        if re.search(r"срочн", joined, re.IGNORECASE):
            data["tariff"] = "срочный тариф"

        fibers = re.search(r"волокон\s*/\s*линий\s*[-–:]?\s*(\d+)", joined, re.IGNORECASE)
        if fibers and (not has_json or str(data.get("quantity") or "") in ("", "1")):
            data["quantity"] = fibers.group(1)
        if re.search(r"одномодов|оптика\s*sm|\bsm\b", joined, re.IGNORECASE):
            _fill(data, "connection_type", "оптика SM")

        client = re.search(r"клиент\s*[-–:]\s*([^\n(]+)", joined, re.IGNORECASE)
        if client:
            _fill(data, "client", client.group(1).strip(" ."))
        contract = re.search(r"договор\s*[-–:]\s*([0-9A-Za-zА-Яа-я/._-]+)", joined, re.IGNORECASE)
        if contract:
            _fill(data, "client_contract", contract.group(1))

        side_a = data.get("side_a") or empty_side()
        loc_a = re.search(
            r"стойк\w*\s+(\d+)\s*этаж[^\d]{0,24}(?:блок\s*)?(\d+)?[^\d]{0,24}помещени[еюя]\s*(\d+(?:\.\d+)?)[^\n]{0,48}ряд\s*(\d+[А-ЯA-Za-z]?)[^\n]{0,24}место\s*(\d+[А-ЯA-Za-z]?)",
            joined,
            re.IGNORECASE,
        )
        if loc_a:
            floor = loc_a.group(1).lstrip("0") or loc_a.group(1)
            _fill(side_a, "floor", floor)
            _fill(side_a, "room", loc_a.group(3))
            _fill(side_a, "row", loc_a.group(4))
            _fill(side_a, "place", loc_a.group(5))
        conn = re.findall(r"разъем\s*[-–:]\s*([A-Z]{1,3}\s*/\s*[A-Z]{2,4})", joined, re.IGNORECASE)
        if conn:
            _fill(side_a, "connector", re.sub(r"\s+", "", conn[0]).upper())
            if len(conn) > 1:
                data.setdefault("side_b", empty_side())
                _fill(data["side_b"], "connector", re.sub(r"\s+", "", conn[1]).upper())

        go = re.search(r"комнат[аы]\s*(?:1[-\s])?(\d{3,4})\s*\(?\s*ГО", joined, re.IGNORECASE)
        if go:
            side_b = data.get("side_b") or empty_side()
            _fill(side_b, "org", "Комната ГО")
            _fill(side_b, "room", go.group(1))
            _fill(side_b, "floor", "8")
            data["side_b"] = side_b

        data["side_a"] = side_a
        if not (data.get("note") or "").strip():
            note = ""
            org_line = re.search(r"Организация соединительной линии[\s\S]{10,420}", joined)
            if org_line:
                note = re.sub(r"\s+", " ", org_line.group(0)).strip()
            if data.get("client"):
                extra = f'Клиент — {data["client"]}'
                if data.get("client_contract"):
                    extra += f' (договор {data["client_contract"]}).'
                note = f"{note} {extra}".strip() if note else extra
            data["note"] = note
        return data


def line_stamp_path(director=None):
    for name in ("texru.png", "techru.png"):
        path = STAMPS_DIR / name
        if path.exists():
            return path
    if director:
        filename = (director.get("stamp_file") or "").strip()
        if filename:
            path = STAMPS_DIR / Path(filename).name
            if path.exists():
                return path
    return None


def suggested_line_names(data):
    number = re.sub(r"[^\w.\-]+", "_", data.get("outgoing_number") or "protyazhka")
    kind = "demontazh" if data.get("work_kind") == "демонтаж" else "montazh"
    return {
        "pdf": OUTPUT_DIR / f"{number}_{kind}.pdf",
        "json": OUTPUT_DIR / f"{number}_{kind}.json",
    }


def _register_fonts():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    windir = Path(os_environ_windir())
    for regular, bold, name, bold_name in (
        (windir / "Fonts" / "arial.ttf", windir / "Fonts" / "arialbd.ttf", "Arial", "Arial-Bold"),
        (windir / "Fonts" / "calibri.ttf", windir / "Fonts" / "calibrib.ttf", "Calibri", "Calibri-Bold"),
    ):
        if regular.exists():
            pdfmetrics.registerFont(TTFont(name, str(regular)))
            pdfmetrics.registerFont(TTFont(bold_name, str(bold if bold.exists() else regular)))
            return name, bold_name
    return "Helvetica", "Helvetica-Bold"


def export_line_pdf(data, stamp_path, output_path):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.platypus import Paragraph, Table, TableStyle

    font, font_bold = _register_fonts()
    width, height = A4
    left, right = 18 * mm, width - 18 * mm
    pdf = pdf_canvas.Canvas(str(output_path), pagesize=A4)
    y = height - 18 * mm

    org = data.get("org_name") or ""
    inn = data.get("org_inn") or ""
    pdf.setFont(font, 11)
    pdf.drawString(left, y, org)
    pdf.drawRightString(right, y, "Техническому директору")
    y -= 14
    if inn:
        pdf.drawString(left, y, f"ИНН {inn}")
    pdf.drawRightString(right, y, 'АО "ММТС-9"')
    y -= 14
    pdf.drawRightString(right, y, "Ушмайкину К.Э.")
    y -= 28

    number = data.get("outgoing_number") or ""
    date = data.get("outgoing_date") or datetime.now().strftime("%Y-%m-%d")
    pdf.setFont(font, 11)
    pdf.drawString(left, y, f"исх. № {number} /дата {date}")
    y -= 18
    kind = data.get("work_kind") or "монтаж"
    title = "Заявка на демонтаж соединительной линии" if kind == "демонтаж" else "Заявка на монтаж"
    pdf.setFont(font_bold, 12)
    pdf.drawString(left, y, title)
    y -= 16
    pdf.setFont(font, 11)
    obj = data.get("object_address") or OBJECT_ADDRESSES[0]
    pdf.drawString(left, y, obj)
    y -= 24
    pdf.setFont(font_bold, 13)
    pdf.drawCentredString(width / 2, y, "Уважаемый Константин Эдуардович!")
    y -= 20

    tariff = data.get("tariff") or "стандартный тариф"
    tariff_phrase = "срочному тарифу" if "сроч" in tariff.lower() else "стандартному тарифу"
    orig = data.get("original_assignments") or []
    if isinstance(orig, str):
        orig = [part.strip() for part in re.split(r"[,;]+", orig) if part.strip()]
    orig_clause = ""
    if orig:
        label = "номер исходного задания" if len(orig) == 1 else "номера исходных заданий"
        orig_clause = f", {label} {', '.join(orig)}"
    action = "демонтаж" if kind == "демонтаж" else "монтаж"
    body = (
        f"Просим Вас организовать {action} соединительной линии на Объекте {obj} "
        f"в соответствии с таблицей 1 по {tariff_phrase}{orig_clause}, оплату гарантируем. "
        f"Контактное лицо, {data.get('contact_name') or data.get('signer_name') or ''} "
        f"телефон {data.get('contact_phone') or ''}, "
        f"E-mail {data.get('contact_email') or ''}"
    )
    style = ParagraphStyle("body", fontName=font, fontSize=11, leading=15, alignment=4)
    para = Paragraph(body.replace("&", "&amp;"), style)
    pw, ph = para.wrap(right - left, 80 * mm)
    para.drawOn(pdf, left, y - ph)
    y -= ph + 14

    cell = ParagraphStyle("c", fontName=font, fontSize=9, leading=12)
    head = ParagraphStyle("h", fontName=font_bold, fontSize=9, leading=12, alignment=1)

    def P(text, header=False):
        return Paragraph(str(text or "").replace("&", "&amp;").replace("<", "&lt;"), head if header else cell)

    a = data.get("side_a") or {}
    b = data.get("side_b") or {}
    rows = [
        [P("Таблица 1 - Параметры соединительной линии", True), "", "", ""],
        [P("Тип соединения", True), P(data.get("connection_type") or ""), "", ""],
        [P("Количество", True), P(data.get("quantity") or "1"), "", ""],
        [P(""), P("Сторона A", True), "", P("Сторона B", True)],
        [P("Наименование организации", True), P(a.get("org")), "", P(b.get("org"))],
        [P("ИНН организации", True), P(a.get("inn")), "", P(b.get("inn"))],
        [P("Тип разъема", True), P(a.get("connector")), "", P(b.get("connector"))],
        [P("Этаж", True), P(a.get("floor")), "", P(b.get("floor"))],
        [P("Помещение", True), P(a.get("room")), "", P(b.get("room"))],
        [P("Ряд", True), P(a.get("row")), "", P(b.get("row"))],
        [P("Место", True), P(a.get("place")), "", P(b.get("place"))],
        [P("Точка включения", True), P(a.get("tap")), "", P(b.get("tap"))],
    ]
    table = Table(rows, colWidths=[48 * mm, 52 * mm, 16 * mm, 54 * mm])
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("SPAN", (0, 0), (-1, 0)),
        ("SPAN", (1, 1), (-1, 1)),
        ("SPAN", (1, 2), (-1, 2)),
        ("ALIGN", (1, 3), (1, 3), "CENTER"),
        ("ALIGN", (3, 3), (3, 3), "CENTER"),
    ]))
    tw, th = table.wrap(right - left, y)
    table.drawOn(pdf, left, y - th)
    y -= th + 12

    note = (data.get("note") or "").strip()
    if note:
        pdf.setFont(font, 10)
        info = Paragraph(
            f"<b>Дополнительная информация:</b> {note.replace('&', '&amp;')}",
            ParagraphStyle("n", fontName=font, fontSize=10, leading=13),
        )
        iw, ih = info.wrap(right - left, 50 * mm)
        info.drawOn(pdf, left, y - ih)
        y -= ih + 10

    pdf.setFont(font, 10)
    pdf.drawString(left, y, "Приложение 1: Шаблон заявки в фомате json.")
    y -= 28
    pdf.setFont(font, 11)
    pdf.drawString(left, y, data.get("signer_title") or "Инженер")
    pdf.drawRightString(right, y, data.get("signer_name") or "")
    signature_y = y
    y -= 14
    pdf.drawString(left, y, org)

    if data.get("use_stamp") and stamp_path and Path(stamp_path).exists():
        draw_round_stamp(pdf, stamp_path, width, signature_y, size_mm=46)
    pdf.save()
    return output_path


def export_line_bundle(data, stamp_path=None, pdf_path=None, json_path=None):
    names = suggested_line_names(data)
    pdf_path = Path(pdf_path or names["pdf"])
    json_path = Path(json_path or names["json"])
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    json_path.write_text(json.dumps(to_assignment_json(data), ensure_ascii=False, indent=2), encoding="utf-8")
    export_line_pdf(data, stamp_path or line_stamp_path(), pdf_path)
    return {"pdf": pdf_path, "json": json_path}
