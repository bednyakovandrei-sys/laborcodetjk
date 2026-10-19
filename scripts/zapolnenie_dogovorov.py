#!/usr/bin/env python3
"""Автозаполнение трудовых договоров и приказов (.docx) из Excel.

Запуск:   python3 zapolnenie_dogovorov.py
Нужно:    pip install python-docx openpyxl

Как работает
  * Спрашивает: какой Excel-файл, какой ЛИСТ, какие шаблоны .docx (договор, приказ ...),
    в какой СТРОКЕ шапка (заголовки) и в КАКОЙ КОЛОНКЕ лежит каждый ключ ({ФИО}, {РМА} ...).
    Совпадающие по названию колонки подставляются сами — остаётся нажать Enter.
    Выбор сохраняется в mapping_<файл>_<лист>.json и предлагается при следующем запуске.
  * Каждая СТРОКА под шапкой = один человек. Для каждого человека создаётся своя папка.
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
import json
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
        return self.row.get(code)       # None = колонка не выбрана, "" = ячейка пуста

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


def col_letter(i):
    return openpyxl.utils.get_column_letter(i + 1)


def short(v, n=18):
    s = "" if v is None else str(v).replace("\n", " ")
    return s if len(s) <= n else s[:n - 1] + "…"


def load_rows(xlsx, sheet):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    return [list(r) for r in wb[sheet].iter_rows(values_only=True)]


def template_codes(paths):
    """Все коды {…} во всех выбранных шаблонах в порядке появления."""
    codes = []
    for t in paths:
        doc = Document(t)
        for p in iter_paragraphs(doc):
            for c in PH.findall(p.text):
                if c not in codes:
                    codes.append(c)
    return codes


def auto_column(code, header):
    for name in ALIASES.get(code, [code]):
        if name in header:
            return header.index(name)
    return None


def choose_header_row(rows):
    print("\nПервые строки листа:")
    for i, r in enumerate(rows[:10], 1):
        print(f"  строка {i}: " + " | ".join(short(c) for c in r[:10]))
    default = next((i for i, r in enumerate(rows, 1) if any(c not in (None, "") for c in r)), 1)
    while True:
        s = input(f"В какой строке ШАПКА (заголовки)? [Enter = {default}]: ").strip()
        try:
            n = int(s) if s else default
            if 1 <= n <= len(rows):
                return n
        except ValueError:
            pass
        print("Не понял, повторите.")


def choose_mapping(codes, header, saved):
    """codes -> индекс колонки или None. Показывает таблицу, даёт править."""
    print("\nКолонки выбранной шапки:")
    for i, h in enumerate(header):
        print(f"  {col_letter(i):>3} (№{i + 1}): {h or '(пусто)'}")
    mapping = {}
    for c in codes:
        if c in saved and saved[c] is not None and saved[c] < len(header):
            mapping[c] = saved[c]
        else:
            mapping[c] = auto_column(c, header)
    while True:
        print("\nКлюч в шаблоне  →  колонка Excel")
        for n, c in enumerate(codes, 1):
            m = mapping[c]
            print(f"  {n:>2}. {{{c}}}  →  " + (f"{col_letter(m)}: {header[m] or '(пусто)'}" if m is not None else "— не выбрана (будут пробелы)"))
        s = input("\nEnter = всё верно; или номер ключа, который изменить: ").strip()
        if not s:
            return mapping
        try:
            code = codes[int(s) - 1]
        except (ValueError, IndexError):
            print("Не понял.")
            continue
        v = input(f"Для {{{code}}} введите букву или номер колонки (A, B, 3 ...); 0 — не использовать: ").strip().upper()
        if v == "0":
            mapping[code] = None
        elif v.isdigit() and 1 <= int(v) <= len(header):
            mapping[code] = int(v) - 1
        elif v.isalpha() and 1 <= openpyxl.utils.column_index_from_string(v) <= len(header):
            mapping[code] = openpyxl.utils.column_index_from_string(v) - 1
        else:
            print("Такой колонки нет.")


def main():
    ap = argparse.ArgumentParser(description="Автозаполнение договоров и приказов из Excel")
    ap.add_argument("--xlsx"), ap.add_argument("--sheet")
    ap.add_argument("--templates", nargs="*")
    ap.add_argument("--out")
    ap.add_argument("--yes", action="store_true", help="не спрашивать про шапку и колонки (авто-подбор по названиям)")
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

    rows = load_rows(xlsx, sheet)
    if not rows:
        sys.exit("Лист пустой.")
    hrow = 1 if a.yes else choose_header_row(rows)
    if a.yes:
        hrow = next((i for i, r in enumerate(rows, 1) if any(c not in (None, "") for c in r)), 1)
    header = [str(h).strip() if h is not None else "" for h in rows[hrow - 1]]

    # ключи: из шаблонов + служебные, из которых считаются ФИО и даты
    codes = template_codes(tpls)
    for extra in ("ФИО", "ФИО_ЛАТ", "САНА"):
        if extra not in codes:
            codes.append(extra)

    map_file = f"mapping_{safe(os.path.splitext(os.path.basename(xlsx))[0])}_{safe(sheet)}.json"
    saved = {}
    if os.path.exists(map_file) and not a.yes:
        if input(f"\nНайден сохранённый выбор колонок ({map_file}). Использовать как основу? [Enter = да, n = нет]: ").strip().lower() != "n":
            with open(map_file, encoding="utf8") as f:
                saved = json.load(f)
    mapping = {c: auto_column(c, header) for c in codes} if a.yes else choose_mapping(codes, header, saved)
    with open(map_file, "w", encoding="utf8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=1)

    people = []
    for r in rows[hrow:]:
        if all(c in (None, "") for c in r):
            continue
        row = {c: (r[i] if i is not None and i < len(r) else "") for c, i in mapping.items() if i is not None}
        people.append(Person(row))
    if not people:
        sys.exit("Под шапкой нет строк с данными.")

    out = a.out or os.path.join(here, "zapolneno_" + dt.datetime.now().strftime("%Y%m%d_%H%M"))
    os.makedirs(out, exist_ok=True)

    report = []                                     # (человек, шаблон, файл, stats, notes)
    for n, person in enumerate(people, 1):
        folder = f"{n:02d}_{safe(str(person.label()))}"
        os.makedirs(os.path.join(out, folder), exist_ok=True)
        for t in tpls:
            name = f"{safe(os.path.splitext(os.path.basename(t))[0])}_{safe(str(person.label()))}.docx"
            person.notes = []
            st = fill_template(t, person, os.path.join(out, folder, name))
            report.append((person.label(), os.path.basename(t), f"{folder}/{name}", st, list(person.notes)))

    # --- отчёт
    lines = [f"ОТЧЁТ  | Excel: {xlsx} | лист: {sheet} | шапка: строка {hrow} | людей: {len(people)} | шаблонов: {len(tpls)}", ""]
    for who, tpl, fname, st, notes in report:
        lines += [f"■ {who}  —  {tpl}  →  {fname}",
                  f"   заполнено ({len(st['filled'])}): {', '.join(st['filled']) or '—'}",
                  f"   ПУСТО в Excel, оставлены пробелы ({len(st['empty'])}): {', '.join(st['empty']) or '—'}",
                  f"   колонка НЕ выбрана, оставлены пробелы ({len(st['no_column'])}): {', '.join(st['no_column']) or '—'}"]
        lines += [f"   ⚠ {x}" for x in notes]
        lines.append("")
    text = "\n".join(lines)
    print("\n" + text)
    with open(os.path.join(out, "report.txt"), "w", encoding="utf8") as f:
        f.write(text)
    wo = openpyxl.Workbook()
    ws = wo.active
    ws.title = "report"
    ws.append(["Человек", "Шаблон", "Файл", "Заполнено", "Пусто в Excel", "Колонка не выбрана", "Пометки"])
    for who, tpl, fname, st, notes in report:
        ws.append([who, tpl, fname, ", ".join(st["filled"]), ", ".join(st["empty"]),
                   ", ".join(st["no_column"]), "; ".join(notes)])
    wo.save(os.path.join(out, "report.xlsx"))
    print(f"Готово. Папка: {out}")


if __name__ == "__main__":
    main()
