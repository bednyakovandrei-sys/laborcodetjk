#!/usr/bin/env python3
"""Автозаполнение трудовых договоров и приказов (.docx) из Excel.

Запуск:   python3 zapolnenie_dogovorov.py
Нужно:    pip install python-docx openpyxl

Как работает
  * Спрашивает: какой Excel-файл, какой ЛИСТ, какие шаблоны .docx (договор, приказ ...).
  * Каждая СТРОКА листа = один человек; ЗАГОЛОВОК колонки = код в шаблоне без скобок:
        {ФИО}  ->  колонка «ФИО»,   {ВАЗИФА_ХУРД}  ->  колонка «ВАЗИФА_ХУРД»  и т.д.
  * Пустая ячейка (или нет такой колонки) -> в документе остаются пробелы.
  * ФИО: везде, где код начинается с «ФИО», выводится «Кириллица (Pinyin)».
        - есть кириллица и латиница в колонках ФИО и ФИО_ЛАТ -> берутся обе;
        - есть только латиница (пиньин) -> кириллица создаётся автотранслитом (в отчёте помечено «проверить»);
        - есть только кириллица -> выводится она одна.
  * Даты: САНА / САНА_СОЛИ / СОЛ / МОҲ / РӮЗ можно брать из одной колонки с датой.
  * В конце — отчёт по каждому человеку и каждому шаблону: что заполнено, что пусто, чего нет в Excel.
    Отчёт печатается в консоль и сохраняется в report.txt и report.xlsx рядом с готовыми файлами.
"""
import argparse
import datetime as dt
import os
import re
import sys

import openpyxl
from docx import Document

SPACES = " " * 14                       # чем заменять пустые значения
FIO_BOTH_EVERYWHERE = True              # во всех кодах ФИО* писать «Кирилица (Pinyin)»
PH = re.compile(r"\{([^{}]+)\}")

MONTHS_TJ = ["январи", "феврали", "марти", "апрели", "майи", "июни",
             "июли", "августи", "сентябри", "октябри", "ноябри", "декабри"]

# Дополнительные названия колонок для одного и того же кода
ALIASES = {
    "岗位": ["岗位", "ВАЗИФА_КИТ", "ВАЗИФА_ЗХ"],
}

# ----------------------------------------------------------------- транслит
_LAT2TJ_DI = {"zh": "ж", "ch": "ч", "sh": "ш", "ng": "нг", "yu": "ю", "ya": "я",
              "yo": "ё", "ye": "е", "xi": "си", "qi": "чи", "ji": "ҷи", "ai": "ай",
              "ei": "эй", "ao": "ао", "ou": "оу", "ui": "уй", "iu": "ю"}
_LAT2TJ = dict(a="а", b="б", c="ц", d="д", e="е", f="ф", g="г", h="ҳ", i="и",
               j="ҷ", k="к", l="л", m="м", n="н", o="о", p="п", q="ч", r="р",
               s="с", t="т", u="у", v="в", w="в", x="кс", y="й", z="з")


def latin_to_tj(s: str) -> str:
    """Грубый транслит латиницы (пиньин) в таджикскую кириллицу. Всегда проверять вручную."""
    out = []
    for word in s.split():
        w, i, res = word.lower(), 0, ""
        while i < len(w):
            if w[i:i + 2] in _LAT2TJ_DI:
                res += _LAT2TJ_DI[w[i:i + 2]]
                i += 2
            else:
                res += _LAT2TJ.get(w[i], w[i])
                i += 1
        out.append(res[:1].upper() + res[1:])
    return " ".join(out)


def is_latin(s: str) -> bool:
    letters = [c for c in s if c.isalpha()]
    return bool(letters) and all(c.isascii() for c in letters)


# ------------------------------------------------------------- значения Excel
def norm(v):
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v
    return str(v).strip()


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        for f in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                return dt.datetime.strptime(v, f).date()
            except ValueError:
                pass
    return None


def date_tj(d: dt.date) -> str:
    return f"{d.day} {MONTHS_TJ[d.month - 1]} соли {d.year}"


class Person:
    """Значения одного человека + правила производных полей."""

    def __init__(self, row: dict):
        self.row = {k: norm(v) for k, v in row.items()}
        self.notes = []                 # пометки для отчёта

    def raw(self, code):
        for name in ALIASES.get(code, [code]):
            if name in self.row:
                return self.row[name]
        return None                     # None = колонки нет, "" = ячейка пуста

    def value(self, code):
        v = self.raw(code)
        if code.startswith("ФИО"):
            return self._fio(code)
        if code in ("САНА", "САНА_СОЛИ", "СОЛ", "МОҲ", "РӮЗ"):
            return self._date(code, v)
        if isinstance(v, (dt.date, dt.datetime)):
            return to_date(v).strftime("%d.%m.%Y")
        return v

    def _date(self, code, v):
        d = to_date(v) if v not in (None, "") else None
        if d is None:                   # взять из соседней колонки с датой
            for other in ("САНА", "САНА_СОЛИ", "СОЛ"):
                o = self.row.get(other)
                if o and to_date(o):
                    d = to_date(o)
                    note = f"{code}: взято из колонки {other}"
                    if note not in self.notes:
                        self.notes.append(note)
                    break
        if d is None:
            if code in ("СОЛ", "МОҲ", "РӮЗ"):
                s, m, r = (self.row.get(x) for x in ("СОЛ", "МОҲ", "РӮЗ"))
                return {"СОЛ": s, "МОҲ": m, "РӮЗ": r}[code] if self.raw(code) is not None else None
            if v is None and any(self.raw(c) is not None for c in ("САНА", "САНА_СОЛИ", "СОЛ")):
                return ""                # колонка с датой есть, ячейка пуста
            return v
        if code in ("САНА", "САНА_СОЛИ"):
            return v if (isinstance(v, str) and v and not to_date(v)) else date_tj(d)
        return str({"СОЛ": d.year, "МОҲ": d.month, "РӮЗ": d.day}[code])

    def _fio(self, code):
        cyr = lat = ""
        for name in ("ФИО", "ФИО_ТҶ", "ФИО_КИР"):
            if self.row.get(name):
                cyr = self.row[name]
                break
        lat = self.row.get("ФИО_ЛАТ", "") or ""
        if code not in ("ФИО", "ФИО_ЛАТ") and self.row.get(code):
            cyr, lat = self.row[code], ""
        if not cyr and not lat and all(self.raw(c) is None for c in ("ФИО", "ФИО_ЛАТ", code)):
            return None
        if cyr and is_latin(cyr) and not lat:          # в «ФИО» лежит пиньин
            cyr, lat = "", cyr
        if lat and not cyr:
            cyr = latin_to_tj(lat)
            note = f"{code}: кириллица получена автотранслитом из «{lat}» — ПРОВЕРИТЬ"
            if note not in self.notes:
                self.notes.append(note)
        if not cyr and not lat:
            return ""
        if not FIO_BOTH_EVERYWHERE or not lat or cyr == lat:
            return cyr or lat
        return f"{cyr} ({lat})"

    def label(self):
        return self.row.get("ФИО") or self.row.get("ФИО_ЛАТ") or "(без ФИО)"


# ------------------------------------------------------------- замена в docx
def iter_paragraphs(doc):
    def walk(container):
        for p in container.paragraphs:
            yield p
        for t in getattr(container, "tables", []):
            for row in t.rows:
                for cell in row.cells:
                    yield from walk(cell)
    yield from walk(doc)
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                     s.even_page_header, s.even_page_footer):
            if part is not None:
                yield from walk(part)


def replace_in_paragraph(p, resolve):
    """Заменяет коды {…}, даже если Word разбил их на несколько фрагментов (runs)."""
    pos = 0
    while True:
        runs = p.runs
        text = "".join(r.text for r in runs)
        m = PH.search(text, pos)
        if not m:
            return
        rep = resolve(m.group(1))
        start, end, acc = m.start(), m.end(), 0
        first = last = None
        for i, r in enumerate(runs):
            lo, hi = acc, acc + len(r.text)
            if first is None and lo <= start < hi:
                first, f_off = i, start - lo
            if lo < end <= hi:
                last, l_off = i, end - lo
                break
            acc = hi
        if first is None or last is None:
            return
        if first == last:
            t = runs[first].text
            runs[first].text = t[:f_off] + rep + t[l_off:]
        else:
            runs[first].text = runs[first].text[:f_off] + rep
            for k in range(first + 1, last):
                runs[k].text = ""
            runs[last].text = runs[last].text[l_off:]
        pos = start + len(rep)


def fill_template(path, person, out_path):
    doc = Document(path)
    stats = {"filled": [], "empty": [], "no_column": []}

    def resolve(code):
        v = person.value(code)
        if v is None:
            if code not in stats["no_column"]:
                stats["no_column"].append(code)
            return SPACES
        if v == "":
            if code not in stats["empty"]:
                stats["empty"].append(code)
            return SPACES
        if code not in stats["filled"]:
            stats["filled"].append(code)
        return str(v)

    for p in iter_paragraphs(doc):
        replace_in_paragraph(p, resolve)
    doc.save(out_path)
    return stats


# ----------------------------------------------------------------- диалоги
def ask_choice(title, options, multi=False):
    print(f"\n{title}")
    for i, o in enumerate(options, 1):
        print(f"  {i}. {o}")
    while True:
        s = input("Номер" + (" (можно несколько через запятую, «все» — все)" if multi else "") + ": ").strip().lower()
        try:
            if multi and s in ("все", "all", "*"):
                return list(range(len(options)))
            idx = [int(x) - 1 for x in s.replace(" ", "").split(",") if x]
            if idx and all(0 <= i < len(options) for i in idx) and (multi or len(idx) == 1):
                return idx if multi else idx[0]
        except ValueError:
            pass
        print("Не понял, повторите.")


def find_files(folder, exts):
    return sorted(f for f in os.listdir(folder)
                  if f.lower().endswith(exts) and not f.startswith("~$"))


def safe(s):
    return re.sub(r"[^\w\-. ()]+", "_", s).strip() or "bez_imeni"


def read_people(xlsx, sheet):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    head = [str(h).strip() if h is not None else "" for h in rows[0]]
    people = []
    for r in rows[1:]:
        if all(c in (None, "") for c in r):
            continue
        people.append(Person({h: c for h, c in zip(head, r) if h}))
    return people


def main():
    ap = argparse.ArgumentParser(description="Автозаполнение договоров и приказов из Excel")
    ap.add_argument("--xlsx"), ap.add_argument("--sheet")
    ap.add_argument("--templates", nargs="*")
    ap.add_argument("--out")
    a = ap.parse_args()
    here = os.getcwd()

    xlsx = a.xlsx
    if not xlsx:
        xs = find_files(here, (".xlsx", ".xlsm"))
        if not xs:
            sys.exit("В текущей папке нет Excel-файла (.xlsx). Положите его рядом со скриптом или укажите --xlsx.")
        xlsx = xs[ask_choice("Какой Excel-файл использовать?", xs)]
    wb = openpyxl.load_workbook(xlsx, read_only=True)
    sheet = a.sheet or wb.sheetnames[ask_choice("Какой ЛИСТ использовать?", wb.sheetnames)]
    wb.close()

    tpls = a.templates
    if not tpls:
        ds = [f for f in find_files(here, (".docx",)) if not f.startswith("zapolneno")]
        if not ds:
            sys.exit("В папке нет шаблонов .docx. Положите договор и приказ рядом или укажите --templates.")
        tpls = [ds[i] for i in ask_choice("Какие ШАБЛОНЫ заполнять?", ds, multi=True)]

    people = read_people(xlsx, sheet)
    if not people:
        sys.exit("На листе нет строк с данными (первая строка — заголовки, дальше — люди).")
    out = a.out or os.path.join(here, "zapolneno_" + dt.datetime.now().strftime("%Y%m%d_%H%M"))
    os.makedirs(out, exist_ok=True)

    report = []                                     # (человек, шаблон, файл, stats, notes)
    for n, person in enumerate(people, 1):
        for t in tpls:
            name = f"{n:02d}_{safe(str(person.label()))}_{os.path.splitext(os.path.basename(t))[0]}.docx"
            person.notes = []
            st = fill_template(t, person, os.path.join(out, name))
            report.append((person.label(), os.path.basename(t), name, st, list(person.notes)))

    # --- отчёт
    lines = [f"ОТЧЁТ  | Excel: {xlsx} | лист: {sheet} | людей: {len(people)} | шаблонов: {len(tpls)}", ""]
    for who, tpl, fname, st, notes in report:
        lines += [f"■ {who}  —  {tpl}  →  {fname}",
                  f"   заполнено ({len(st['filled'])}): {', '.join(st['filled']) or '—'}",
                  f"   ПУСТО в Excel, оставлены пробелы ({len(st['empty'])}): {', '.join(st['empty']) or '—'}",
                  f"   НЕТ колонки в Excel, оставлены пробелы ({len(st['no_column'])}): {', '.join(st['no_column']) or '—'}"]
        lines += [f"   ⚠ {x}" for x in notes]
        lines.append("")
    text = "\n".join(lines)
    print("\n" + text)
    with open(os.path.join(out, "report.txt"), "w", encoding="utf8") as f:
        f.write(text)
    wo = openpyxl.Workbook()
    ws = wo.active
    ws.title = "report"
    ws.append(["Человек", "Шаблон", "Файл", "Заполнено", "Пусто в Excel", "Нет колонки", "Пометки"])
    for who, tpl, fname, st, notes in report:
        ws.append([who, tpl, fname, ", ".join(st["filled"]), ", ".join(st["empty"]),
                   ", ".join(st["no_column"]), "; ".join(notes)])
    wo.save(os.path.join(out, "report.xlsx"))
    print(f"Готово. Папка: {out}")


if __name__ == "__main__":
    main()
