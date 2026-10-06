"""Конвертер выгрузки «Адлия» (.docx, кнопка «Word») в Markdown для akty/tekst/.

Текст переносится дословно. Добавляются только заголовки Markdown:
- абзацы, выровненные по центру, и заголовки вида «Глава/Статья/Раздел» -> «## …»;
- абзацы, начинающиеся с «N. », -> перед ними «## Пункт N» (если номер идёт по порядку);
- моноширинные перечни (с колонками из пробелов, псевдографикой) -> блок ```.

Использование: python docx2md.py source.docx[+source2.docx] out.md "шапка (первые строки файла)"
"""
import re
import sys

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.table import Table
from docx.text.paragraph import Paragraph

BOX = '¦+'


def para_text(p):
    return p.text.replace('\xa0', ' ')


def is_pre(t):
    if re.search(r'^\+-{5,}', t.strip()):
        return True
    lines = t.split('\n')
    return len(lines) >= 2 and any(re.search(r'\S {4,}\S', l) or re.search(r'_{6,}', l) for l in lines)


def box_lines(t):
    # шапка псевдографической таблицы приходит одной строкой: режем по границам строк
    t = t.strip()
    t = re.sub(r'\+(?=¦)', '+\n', t)
    t = re.sub(r'¦(?=¦)', '¦\n', t)
    t = re.sub(r'¦(?=\+)', '¦\n', t)
    return t


def table_md(tb):
    rows = []
    for r in tb.rows:
        cells = []
        prev = None
        for c in r.cells:
            if prev is not None and c._tc is prev:
                continue  # объединённые ячейки
            prev = c._tc
            cells.append(re.sub(r'\s+', ' ', c.text.replace('\xa0', ' ')).strip().replace('|', '\\|'))
        rows.append(cells)
    if not rows:
        return ''
    n = max(len(r) for r in rows)
    rows = [r + [''] * (n - len(r)) for r in rows]
    out = ['| ' + ' | '.join(rows[0]) + ' |', '|' + ' --- |' * n]
    out += ['| ' + ' | '.join(r) + ' |' for r in rows[1:]]
    return '\n'.join(out)


def blocks(doc):
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.split('}')[1]
        if tag == 'p':
            yield Paragraph(child, doc)
        elif tag == 'tbl':
            yield Table(child, doc)


def convert(paths, header):
    out = [header.rstrip('\n'), '']
    for i, path in enumerate(paths.split('+')):
        if i:
            out += ['---', '']
        out += convert_one(path)
    return '\n'.join(out).rstrip() + '\n'


def convert_one(path):
    doc = docx.Document(path)
    out = []
    last = 0
    for b in blocks(doc):
        if isinstance(b, Table):
            out += [table_md(b), '']
            continue
        t = para_text(b)
        if not t.strip():
            continue
        if is_pre(t):
            body = box_lines(t) if t.strip().startswith('+') and '\n' not in t.strip() else t
            lines = [l.rstrip() for l in body.split('\n')]
            while lines and not lines[0].strip():
                lines.pop(0)
            out += ['```'] + lines + ['```', '']
            continue
        one = re.sub(r'[ \t]+', ' ', t).strip()
        lines = [re.sub(r'[ \t]+', ' ', l).strip() for l in t.split('\n') if l.strip()]
        al = b.alignment
        if re.match(r'^(Глава|ГЛАВА|Статья|СТАТЬЯ|Раздел|РАЗДЕЛ)\s', one) and len(one) < 300:
            out += ['## ' + one, '']
            continue
        if al == WD_ALIGN_PARAGRAPH.CENTER and len(one) < 400 and len(lines) <= 3 and not one.startswith('('):
            out += ['## ' + ' '.join(lines), '']
            continue
        m = re.match(r'^(\d+)\.(?!\d)\s*[^\s\d]', one)
        if m:
            n = int(m.group(1))
            if n == 1 or n == last + 1 or n == last:
                out += ['## Пункт %d' % n, '']
                last = n
        out += ['  \n'.join(lines), '']
    return out


if __name__ == '__main__':
    src, dst, header = sys.argv[1], sys.argv[2], sys.argv[3]
    md = convert(src, header)
    with open(dst, 'w', encoding='utf-8', newline='\n') as f:
        f.write(md)
    print(dst, len(md), 'chars')
