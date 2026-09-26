import re,csv,json
from io import BytesIO,StringIO
from datetime import datetime,date,timedelta
from openpyxl import load_workbook

ALIASES={
    "ticker":["ticker","symbol","underlying","underlying symbol","security","security symbol"],
    "expiry":["expiry","expiration","expiration date","maturity","exp date"],
    "option_type":["option type","put/call","put call","call/put","cp","type"],
    "strike":["strike","strike price"],
    "quantity":["quantity","qty","position","contracts","contract quantity","net quantity"],
    "multiplier":["multiplier","contract multiplier"],
}

CALL_MONTH_CODES="ABCDEFGHIJKL"
PUT_MONTH_CODES="MNOPQRSTUVWX"
NIRVANA_OPTION_RE=re.compile(r"^O:(?P<ticker>[A-Z0-9.\-]+)\s+(?P<yy>\d{2})(?P<monthcode>[A-X])(?P<strike>\d+(?:\.\d+)?)D(?P<day>\d{1,2})$",re.I)

def norm(s):
    return re.sub(r"[^a-z0-9]+"," ",str(s).strip().lower()).strip()

def option_security(ticker,expiry,typ,strike):
    y,m,d=expiry.split("-")
    cp="C" if typ.upper().startswith("C") else "P"
    si=int(round(float(strike)*1000))
    return f"{ticker.upper()} {y[-2:]}{m}{d}{cp}{si:08d} Equity"

def detect_map(headers):
    nm={norm(h):h for h in headers}
    out={}
    for field,aliases in ALIASES.items():
        for a in aliases:
            if norm(a) in nm:
                out[field]=nm[norm(a)]
                break
    return out

def _row_dict(headers,row):
    return {headers[i]:row[i] if i<len(row) else None for i in range(len(headers))}

def parse_upload(data,filename):
    """Return (headers, rows, meta). Detects Nirvana Exposure Summary reports even when header is not row 1."""
    meta={"format":"generic"}
    if filename.lower().endswith(".csv"):
        txt=data.decode("utf-8-sig",errors="replace")
        reader=csv.DictReader(StringIO(txt))
        rows=list(reader)
        return list(reader.fieldnames or []),rows,meta

    wb=load_workbook(BytesIO(data),read_only=True,data_only=True)
    ws=wb.active
    vals=list(ws.iter_rows(values_only=True))
    if not vals:return [],[],meta

    # Nirvana report headers can begin after several title/report-date rows.
    header_idx=None
    for i,row in enumerate(vals[:25]):
        normalized={norm(x) for x in row if x is not None}
        if {"symbol","expiration date","position"}.issubset(normalized):
            header_idx=i
            break

    if header_idx is None:
        header_idx=0

    headers=[str(x).strip() if x is not None else "" for x in vals[header_idx]]
    rows=[]
    for row in vals[header_idx+1:]:
        if not any(x is not None and str(x).strip() for x in row):continue
        rows.append(_row_dict(headers,row))

    nheaders={norm(h) for h in headers}
    preamble="\n".join(" ".join(str(x) for x in r if x is not None) for r in vals[:header_idx])
    if {"symbol","expiration date","position","delta"}.issubset(nheaders) and any(str(r.get("Symbol","")).startswith("O:") for r in rows):
        meta["format"]="nirvana_exposure_by_underlying"
        m=re.search(r"Report Date:\s*([0-9/\-]+)",preamble,re.I)
        if m:meta["report_date"]=m.group(1)
        nav=None
        for r in rows:
            if str(r.get("Symbol","")).strip().lower().startswith("grand total"):
                try:nav=float(r.get("NAV"))
                except Exception:nav=None
                break
        meta["nav_usd"]=nav
    return headers,rows,meta

def parse_date(v):
    if v is None:return None
    if isinstance(v,datetime):return v.date().isoformat()
    if isinstance(v,date):return v.isoformat()
    if isinstance(v,(int,float)):
        # Excel 1900 date system used by Nirvana exports.
        try:return (datetime(1899,12,30)+timedelta(days=float(v))).date().isoformat()
        except Exception:return None
    s=str(v).strip()
    for fmt in ("%Y-%m-%d","%m/%d/%Y","%m/%d/%y","%d/%m/%Y","%d-%b-%Y","%b %d %Y"):
        try:return datetime.strptime(s,fmt).date().isoformat()
        except Exception:pass
    if re.match(r"^\d{4}-\d{2}-\d{2}",s):return s[:10]
    return None

def parse_type(v):
    s=str(v).strip().upper()
    if s.startswith("P"):return "P"
    if s.startswith("C"):return "C"
    return None

def parse_nirvana_option_symbol(symbol):
    s=str(symbol or "").strip().upper()
    m=NIRVANA_OPTION_RE.match(s)
    if not m:return None
    ticker=m.group("ticker")
    yy=int(m.group("yy")); code=m.group("monthcode").upper();strike=float(m.group("strike"));day=int(m.group("day"))
    if code in CALL_MONTH_CODES:
        typ="C";month=CALL_MONTH_CODES.index(code)+1
    elif code in PUT_MONTH_CODES:
        typ="P";month=PUT_MONTH_CODES.index(code)+1
    else:return None
    expiry=date(2000+yy,month,day).isoformat()
    return {"ticker":ticker,"expiry":expiry,"option_type":typ,"strike":strike}

def parse_nirvana_exposure_rows(rows):
    """Parse actual option rows only; summary/cash/grand-total rows are intentionally ignored."""
    out=[];errors=[];skipped=0
    for idx,r in enumerate(rows,start=1):
        sym=str(r.get("Symbol","") or "").strip()
        if not sym.startswith("O:"):
            skipped+=1
            continue
        parsed=parse_nirvana_option_symbol(sym)
        if not parsed:
            errors.append(f"{sym}: option symbol could not be parsed")
            continue
        try:q=float(str(r.get("Position")).replace(",",""))
        except Exception:
            errors.append(f"{sym}: invalid Position")
            continue
        file_expiry=parse_date(r.get("Expiration Date"))
        if file_expiry and file_expiry!=parsed["expiry"]:
            errors.append(f"{sym}: symbol expiry {parsed['expiry']} != file expiry {file_expiry}")
            continue
        out.append({**parsed,"quantity":q,"multiplier":100.0,"source_symbol":sym})
    return out,errors,{"skipped_non_option_rows":skipped}

def normalized_import(rows,mapping):
    out=[];errors=[]
    for idx,r in enumerate(rows,start=2):
        try:
            t=str(r.get(mapping["ticker"],"")).strip().upper()
            e=parse_date(r.get(mapping["expiry"]))
            typ=parse_type(r.get(mapping["type"]))
            k=float(r.get(mapping["strike"]))
            q=float(str(r.get(mapping["qty"])).replace(",",""))
            m=100.0
            if mapping.get("mult") and r.get(mapping["mult"]) not in (None,""):
                m=float(r.get(mapping["mult"]))
            if not t or not e or not typ: raise ValueError("missing ticker/expiry/type")
            out.append({"ticker":t,"expiry":e,"option_type":typ,"strike":k,"quantity":q,"multiplier":m})
        except Exception as ex:
            errors.append(f"row {idx}: {ex}")
    return out,errors

def keypos(x):
    return (x["ticker"],x["expiry"],x["option_type"],round(float(x["strike"]),6))
