"""tjbuild.py RAW_TJ_TXT OUT_MD HEADER_MD START_REGEX [HEADING_LINE ...]
Собирает .md из таджикского текста: до строки START_REGEX — преамбула (абзацы),
после — пункты '## Пункт N'; строки HEADING_LINE (точное совпадение начала) — '### ...'."""
import sys,re,collections
raw,out,headerf,start=sys.argv[1:5]; heads=sys.argv[5:]
lines=[l.strip() for l in open(raw,encoding='utf-8').read().replace('\f','\n').split('\n')]
lines=[l for l in lines if l]
started=False; paras=[]; cur=None; lasth=False
def flush():
    global cur
    if cur: paras.append(cur)
    cur=None
for l in lines:
    if lasth and (l[:1].islower() or l[:1]=='(') and not re.match(r'^(\d+\.|[IVX]+\.|- )',l) and not any(l.startswith(h) for h in heads):
        paras[-1]=(paras[-1][0],paras[-1][1]+' '+l); continue
    lasth=False
    if not started and re.search(start,l):
        flush(); started=True; paras.append(('h1',l)); lasth=True; continue
    h=next((h for h in heads if l.startswith(h)),None)
    if h is not None and started:
        flush(); paras.append(('h',l)); lasth=True; continue
    m=re.match(r'^(\d+)\.\s+\S',l)
    if m or l.startswith('- ') or re.match(r'^[абвгдеж]\)\s',l) or (not started and re.match(r'^(Раиси|аз \d|ш\.\s?Душанбе|Бо қарори|Замимаи)',l)):
        flush(); cur=['p' if not (m and started) else 'pt', l, m.group(1) if m else None]; continue
    if cur is None: cur=['p',l,None]
    else:
        prev=cur[1]
        cur[1]= prev[:-1]+l if prev.endswith('-') and not prev.endswith(' -') else prev+' '+l
flush()
md=[open(headerf,encoding='utf-8').read().rstrip(),'','---','','## Матни асл (таджикский оригинал)','']
for p in paras:
    if p[0]=='h1': md+=[f'### {p[1]}','']
    elif p[0]=='h': md+=[f'### {p[1]}','']
    elif p[0]=='pt': md+=[f'## Пункт {p[2]}','',p[1],'']
    else: md+=[p[1],'']
open(out,'w',encoding='utf-8').write('\n'.join(md).rstrip()+'\n')
src=collections.Counter(re.findall(r'\w+',' '.join(lines).lower()))
body='\n'.join(md[md.index('## Матни асл (таджикский оригинал)'):])
dst=collections.Counter(re.findall(r'\w+',body.lower()))
miss=src-dst
print('paras',len(paras),'missing words',sum(miss.values()),list(miss.items())[:12])
