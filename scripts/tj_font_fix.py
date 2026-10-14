"""Перекодирует таджикский текст из «шрифтовой» кодировки (і→ҳ, ї→ҷ, ѕ→қ, њ→ӣ, ў→ӯ, љ→ғ, ѐ→ё).
Запуск: pdftotext file.pdf file.txt && python3 -I scripts/tj_font_fix.py file.txt  → file.tj.txt"""
import sys
M=str.maketrans({'і':'ҳ','І':'Ҳ','ї':'ҷ','Ї':'Ҷ','њ':'ӣ','Њ':'Ӣ','ѕ':'қ','Ѕ':'Қ','ў':'ӯ','Ў':'Ӯ','ѐ':'ё','Ѐ':'Ё','љ':'ғ','Љ':'Ғ'})
for f in sys.argv[1:]:
    t=open(f,encoding='utf-8').read().translate(M)
    open(f.replace('.txt','.tj.txt'),'w',encoding='utf-8').write(t)
