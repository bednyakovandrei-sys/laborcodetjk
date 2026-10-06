#!/usr/bin/env python3
"""Разбирает .docx Трудового кодекса РТ на Markdown-файлы по главам + оглавление.

Запуск (из корня репозитория):
    python3 -I scripts/build_kodeks.py source/<файл>.docx kodeks/

При выходе новой редакции кодекса: положить новый .docx в source/ и перезапустить.
"""
import re
import sys
from pathlib import Path

import docx

RAZDEL_RE = re.compile(r"^РАЗДЕЛ\s+([0-9IVXLC]+)\.\s*(.+)$")
GLAVA_RE = re.compile(r"^ГЛАВА\s+(\d+)\.\s*(.+)$")
STATYA_RE = re.compile(r"^Статья\s+(\d+(?:\(\d+\))?)\.\s*(.+)$")


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace(" ", " ")).strip()


def main(src: Path, out: Path) -> None:
    paragraphs = [clean(p.text) for p in docx.Document(str(src)).paragraphs]
    paragraphs = [p for p in paragraphs if p]

    header = []  # заголовок и преамбула до первого раздела
    chapters = []  # {num, title, razdel, articles: [{num, title, body: []}]}
    razdel = None
    chapter = None
    article = None

    for p in paragraphs:
        if m := RAZDEL_RE.match(p):
            razdel = f"Раздел {m[1]}. {m[2]}"
            continue
        if m := GLAVA_RE.match(p):
            chapter = {"num": int(m[1]), "title": m[2], "razdel": razdel, "articles": []}
            chapters.append(chapter)
            article = None
            continue
        if m := STATYA_RE.match(p):
            article = {"num": m[1], "title": m[2], "body": []}
            chapter["articles"].append(article)
            continue
        if article is None:
            header.append(p)
        else:
            article["body"].append(p)

    out.mkdir(parents=True, exist_ok=True)
    doc_title = header[0] if header else "Трудовой кодекс Республики Таджикистан"
    edition = header[1] if len(header) > 1 else ""

    index = [f"# {doc_title}", "", edition, "", "Оглавление: раздел → глава → статьи (ссылки на файлы глав).", ""]
    if len(header) > 2:
        index += ["> " + " ".join(header[2:]), ""]

    current_razdel = None
    total = 0
    for ch in chapters:
        fname = f"glava-{ch['num']:02d}.md"
        lines = [
            f"# Глава {ch['num']}. {ch['title']}",
            "",
            f"_{doc_title} {edition}_  ",
            f"_{ch['razdel']}_",
            "",
        ]
        for a in ch["articles"]:
            lines += [f"## Статья {a['num']}. {a['title']}", ""]
            for para in a["body"]:
                lines += [para, ""]
        (out / fname).write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")

        if ch["razdel"] != current_razdel:
            current_razdel = ch["razdel"]
            index += [f"## {current_razdel}", ""]
        index += [f"### [Глава {ch['num']}. {ch['title']}]({fname})", ""]
        for a in ch["articles"]:
            index.append(f"- ст. {a['num']} — {a['title']}")
            total += 1
        index.append("")

    (out / "INDEX.md").write_text("\n".join(index).rstrip() + "\n", encoding="utf-8")
    print(f"Глав: {len(chapters)}, статей: {total} → {out}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(Path(sys.argv[1]), Path(sys.argv[2]))
