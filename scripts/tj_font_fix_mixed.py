"""Перекодировка таджикского текста со смешанными шрифтовыми кодировками (построчно).
Кодировка A: і→ҳ ї→ҷ ѕ→қ њ→ӣ љ→ғ ў→ӯ; кодировка B: њ→ҳ ї→ӣ љ→ҷ ќ→қ ѓ→ғ ў→ӯ."""
import sys
A=str.maketrans({'і':'ҳ','І':'Ҳ','ї':'ҷ','Ї':'Ҷ','ѕ':'қ','Ѕ':'Қ','њ':'ӣ','Њ':'Ӣ','љ':'ғ','Љ':'Ғ','ў':'ӯ','Ў':'Ӯ','ѐ':'ё'})
B=str.maketrans({'њ':'ҳ','Њ':'Ҳ','ї':'ӣ','Ї':'Ӣ','љ':'ҷ','Љ':'Ҷ','ќ':'қ','Ќ':'Қ','ѓ':'ғ','Ѓ':'Ғ','ў':'ӯ','Ў':'Ӯ','ѐ':'ё'})
for f in sys.argv[1:]:
    out=[]; mode='A'
    for l in open(f,encoding='utf-8').read().split('\n'):
        if any(c in l for c in 'іІѕЅ'): mode='A'
        elif any(c in l for c in 'ќЌѓЃ'): mode='B'
        out.append(l.translate(A if mode=='A' else B))
    open(f.replace('.txt','.tj.txt'),'w',encoding='utf-8').write('\n'.join(out))
