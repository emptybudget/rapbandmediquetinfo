import openpyxl, glob, os, re, datetime, json
from collections import defaultdict
# 발주서 zip을 풀어둔 폴더를 SRC 로 지정
SRC = "./medyssey_orders"
TODAY = datetime.date.today(); HALF_LIFE = 180.0   # 반감기(일) — 조정 가능
def decode(n): return re.sub(r'#U([0-9a-fA-F]{4})', lambda m: chr(int(m.group(1),16)), n)  # 깨진 파일명 복원용(필요시)
def is_t(k): return ('Medyssey' in k) or ('양식' in k)
def canon(p):
    if p is None: return '미지정'
    s=p
    for w in ['교환건','교환','수정본','복사본','오발주건','오발주','추가분','추가','사이즈','로드분실건','소아병동','요청서','요청건','ILIAD']: s=s.replace(w,'')
    s=re.sub(r'\d+','',s).strip(' ()')
    al={'서순':'서울순천향','분당제생메듀사':'분당제생'}
    if ',' in s: return '복합('+s+')'
    return al.get(s, s if s else '미지정')
def nsz(d):
    if d in (None,''): return ''
    return re.sub(r'mm$','',re.sub(r'\s+','',str(d)).lower())
def ki(b): return re.sub(r'\s+',' ',str(b).strip()).casefold()
EX=['복사본','오발주']
rows=[]; disp=defaultdict(lambda:defaultdict(int))
for f in glob.glob(os.path.join(SRC,'*.xlsx')):
    k=os.path.basename(f)
    if '#U' in k: k=decode(k)
    if is_t(k) or any(w in k for w in EX): continue
    m=re.search(r'(20\d{6})',k); d=datetime.datetime.strptime(m.group(1),'%Y%m%d').date() if m else None
    pm=re.search(r'\(([^)]*)\)',k); h=canon(pm.group(1) if pm else None)
    try: wb=openpyxl.load_workbook(f,data_only=True,read_only=True)
    except: continue
    for ws in wb.worksheets:
        for r in ws.iter_rows(values_only=True):
            if len(r)<5: continue
            b,ds,q=r[1],r[3],r[4]
            if isinstance(b,str) and b.strip() and '합계' not in b and b.strip()!='품목' and isinstance(q,(int,float)) and not isinstance(q,bool):
                rows.append((d,h,ki(b),b.strip(),nsz(ds),q)); disp[ki(b)][b.strip()]+=1
dispn={ik:max(c.items(),key=lambda kv:kv[1])[0] for ik,c in disp.items()}
def wt(d):
    if d is None: return 0.3
    return 0.5**(((TODAY-d).days)/HALF_LIFE)
agg=defaultdict(lambda:defaultdict(lambda:defaultdict(lambda:[0.0,0,0.0,0.0,None])))
for d,h,ik,dp,sz,q in rows:
    c=agg[h][ik][sz]; c[0]+=q; c[1]+=1; c[2]+=q*wt(d); c[3]+=wt(d)
    if d and (c[4] is None or d>c[4]): c[4]=d
out={'meta':{'today':str(TODAY),'halfLifeDays':HALF_LIFE},'hospitals':{}}
for h in agg:
    out['hospitals'][h]={}
    for ik in agg[h]:
        sizes=[]
        for sz,c in agg[h][ik].items():
            sizes.append({'size':sz,'q':c[0],'n':c[1],'w':round(c[2],2),'last':str(c[4]) if c[4] else None})
        sizes.sort(key=lambda s:-s['w'])
        out['hospitals'][h][dispn[ik]]={'w':round(sum(s['w'] for s in sizes),2),'sizes':sizes}
open('medysseyData.js','w').write("export const MEDYSSEY_SEED = "+json.dumps(out,ensure_ascii=False,indent=1)+";\n")
print("medysseyData.js 생성 완료")