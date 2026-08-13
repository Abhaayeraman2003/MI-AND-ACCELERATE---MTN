import re
MONTHS=["January","February","March","April","May","June","July","August","September","October","November","December"]
KIND_LABEL={"MI":"MI Initiatives","Accelerate":"Accelerate Initiatives","Priorities":"MI & Accelerate Priorities (KPIs)","MI Tracker":"Monthly Actions Tracker"}
KIND_ORDER=["MI","Accelerate","Priorities","MI Tracker"]
READONLY_COLS={"MI":2,"Accelerate":2,"Priorities":3,"MI Tracker":2}
RAG_OPTS=["","Green — on track","Amber — delayed","Red — at risk","Blue — delivered / BAU"]
RAG_NUM_MAP={"1":"Green — on track","2":"Amber — delayed","3":"Red — at risk","4":""}
MATURITY=["Basic","Intermediate","Control","Advanced"]
MATURITY_LEVEL={"basic":1,"beginner":1,"initiation":1,"scaling":2,"intermediate":2,"control":3,"advanced":4}
RAG_COLOR={"Green":"#2E7D32","Amber":"#E8A317","Red":"#C62828","Blue":"#1565C0","Not set":"#CFCFCF"}
def ascii_slug(s):
    if not s: return ""
    t=str.maketrans("àáâãäåÀÁÂÃÄÅèéêëÈÉÊËìíîïÌÍÎÏòóôõöÒÓÔÕÖùúûüÙÚÛÜçÇñÑ","aaaaaaAAAAAAeeeeEEEEiiiiIIIIoooooOOOOOuuuuUUUUcCnN")
    return re.sub(r"[^A-Za-z0-9]+","",s.translate(t))
def is_section(row): return len([c for c in row if c and c.strip()])<=1
def field_type(h):
    t=(h or "").lower()
    if "rag" in t: return "rag"
    if "maturity" in t: return "maturity"
    if "due" in t or "date" in t or "timeline" in t: return "short"
    if any(k in t for k in ("impact","performance","ytd","bud","base")) or re.search(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)",t): return "short"
    return "long"
def rag_normalise(v):
    raw=str(v or "").strip()
    if not raw: return ""
    if raw in RAG_NUM_MAP: return RAG_NUM_MAP[raw]
    t=raw.lower()
    if t.startswith("green") or "on track" in t: return RAG_OPTS[1]
    if t.startswith("amber") or "delay" in t: return RAG_OPTS[2]
    if t.startswith("red") or "at risk" in t or "off track" in t: return RAG_OPTS[3]
    if t.startswith("blue") or "deliver" in t or "bau" in t: return RAG_OPTS[4]
    return ""
def rag_bucket(v):
    n=rag_normalise(v)
    for k in ("Green","Amber","Red","Blue"):
        if n.startswith(k): return k
    return "Not set"
def row_info(b,row,group):
    kind=b["kind"]; grouped=kind in ("Priorities","MI Tracker")
    if grouped:
        if row[0] and row[0].strip(): group=row[0].strip()
        name=row[1].strip() if len(row)>1 and row[1] and row[1].strip() else (row[0] or "").strip(); start=2
    else:
        if row[0] and row[0].strip(): name,start=row[0].strip(),1
        else: name=((row[1] if len(row)>1 else "") or "").strip(); start=2
    return name,group,start
def col_label(b,i,rm):
    h=b["headers"][i] if i<len(b["headers"]) else ("Column %d"%(i+1))
    if b["kind"]=="Priorities" and i==len(b["headers"])-1 and any(h.lower().startswith(m[:3].lower()) for m in MONTHS):
        return rm+" actual"
    return h
def parse_amount(s):
    if s is None: return None
    t=str(s).strip()
    if not t: return None
    tl=t.lower()
    if tl in ("na","n/a","nan","-","–","tbd","non-financial") or "work in progress" in tl: return None
    if "%" in t: return None
    mult=1.0
    if re.search(r"\bbn\b|billion|\bbil\b",tl): mult=1e9
    elif re.search(r"\bmn\b|\bmln\b|million",tl): mult=1e6
    elif re.search(r"\bm\b",tl): mult=1e6
    elif re.search(r"\bk\b",tl): mult=1e3
    m=re.search(r"[-+]?\d[\d ,.]*\d|\d",t)
    if not m: return None
    num=m.group(0).replace(" ","")
    if "," in num and "." in num: num=num.replace(",","")
    elif "," in num:
        parts=num.split(",")
        num=num.replace(",","") if len(parts[-1])==3 else num.replace(",",".")
    try: val=float(num)
    except ValueError: return None
    return val*mult
def accel_progress(fields):
    est_key=act_key=None
    for k in fields:
        kl=k.lower()
        if "rag" in kl: continue
        if ("actual" in kl or "ytd" in kl or "performance" in kl) and act_key is None: act_key=k
        if "estimat" in kl or "expected" in kl:
            if "annual" in kl: est_key=k
            elif est_key is None: est_key=k
    if not est_key or not act_key: return None,None,None,est_key,act_key
    est=parse_amount(fields.get(est_key)); act=parse_amount(fields.get(act_key))
    if est is None or act is None or est<=0: return None,act,est,est_key,act_key
    return max(0.0,min(100.0,act/est*100.0)),act,est,est_key,act_key
def maturity_progress(fields):
    cur=tgt=None
    for k,v in fields.items():
        kl=k.lower()
        if "current" in kl and "maturity" in kl: cur=v
        elif "target" in kl and "maturity" in kl: tgt=v
    c=MATURITY_LEVEL.get(str(cur or "").strip().lower()); t=MATURITY_LEVEL.get(str(tgt or "").strip().lower())
    if not c or not t or t<=0: return None,cur,tgt
    return max(0.0,min(100.0,c/t*100.0)),cur,tgt
def iter_baseline_records(DATA,country):
    out=[]
    for b in DATA[country]:
        kind=b["kind"]; group=""
        for row in b["rows"]:
            if is_section(row):
                group=next((c for c in row if c and c.strip()),""); continue
            name,group,_=row_info(b,row,group)
            if not name: continue
            fields={}
            for i,h in enumerate(b["headers"]):
                if h and h.strip() and i<len(row): fields[h]=row[i]
            out.append({"opco":country,"reportingMonth":"Baseline (deck)","type":KIND_LABEL[kind],"kind":kind,"section":"" if group==name else group,"initiative":name,"fields":fields,"changed":False})
    return out
def submissions_to_records(subs):
    kbl={v:k for k,v in KIND_LABEL.items()}; out=[]
    for p in subs:
        for s in p.get("sections",[]):
            t=s.get("type",""); kind=kbl.get(t,t)
            for it in s.get("items",[]):
                out.append({"opco":p.get("opco",""),"reportingMonth":p.get("reportingMonth",""),"type":t,"kind":kind,"section":it.get("section",""),"initiative":it.get("initiative",""),"fields":it.get("fields",{}) or {},"changed":bool(it.get("changed"))})
    return out
