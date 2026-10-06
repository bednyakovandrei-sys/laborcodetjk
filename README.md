# Трудовой кодекс РТ — база для HR-кейсов

Трудовой кодекс Республики Таджикистан от 23.07.2016 № 1329
(ред. по состоянию на 17.06.2026), разобранный в Markdown для работы с Claude.

- `kodeks/INDEX.md` — оглавление (41 глава, 367 статей)
- `kodeks/glava-NN.md` — текст по главам, статьи — заголовки `## Статья N. ...`
- `CLAUDE.md` — как Claude разбирает кейсы: поиск норм, формат ответа, навигатор «тема → статьи»
- `cases/` — сохранённые кейсы (шаблон `cases/_template.md`)
- `source/` — исходный .docx; `scripts/build_kodeks.py` — пересборка `kodeks/` из него

## Обновление редакции

```
python3 -m pip install python-docx
python3 -I scripts/build_kodeks.py source/<новый файл>.docx kodeks/
```
