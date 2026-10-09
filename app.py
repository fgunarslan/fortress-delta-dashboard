from __future__ import annotations
import os,json,hmac,html,uuid,re,math
from datetime import datetime,timezone,date
from zoneinfo import ZoneInfo
from urllib.parse import quote,unquote
from io import BytesIO
from openpyxl import load_workbook
from fastapi import FastAPI,Request,Form,UploadFile,File,HTTPException
from fastapi.responses import HTMLResponse,RedirectResponse

from db import SessionLocal,User,Setting,Position,Audit,Snapshot,ImportJob,DailyReturnBaseline,DailyReturnPosition,DailyReturnExposureBaseline,DailyReturnExposurePosition,get_setting,set_setting,audit,config_bump,utcnow
from security import ADMIN_USER,ensure_admin,hash_password,verify_password,sign_session,read_session
from helpers import option_security,detect_map,parse_upload,parse_nirvana_exposure_rows,normalized_import,keypos
from daily_return import parse_nirvana_pnl_pdf,calculate_mark_to_market_return,auto_refresh_allowed,ny_time_label

COLLECTOR_TOKEN=os.getenv("COLLECTOR_TOKEN","")
STALE_AFTER_SECONDS=max(int(os.getenv("STALE_AFTER_SECONDS","150")),150)
app=FastAPI(title="RPD Fund Management v5.7 — Opportunity Delta + Delay Badges",docs_url=None,redoc_url=None)

def init_data():
    ensure_admin()
    with SessionLocal() as db:
        defaults={"portfolio_name":"RPD Fortress Fund","nav_usd":"46865138.06795","limit_pct":"15.0","config_version":"1","config_updated_at":utcnow().isoformat(),"opportunity_config_version":"1","opportunity_limit_pct":""}
        for k,v in defaults.items():
            if db.get(Setting,k) is None: db.add(Setting(key=k,value=v))
        if db.query(Position).count()==0:
            seeds=[
                ("CG","2026-11-20","P",35,-530),("CG","2026-11-20","P",37.5,-550),
                ("COMP","2026-10-16","P",8,-4400),("GTM","2026-10-16","P",3,-11660),
                ("GTM","2026-10-02","P",3,-11667),("MNDY","2026-11-20","P",65,-1090),
                ("POOL","2026-10-16","P",155,-230),("QXO","2026-11-20","P",10,-2279),
                ("REAL","2026-10-16","P",7.5,-4700)]
            for t,e,typ,k,q in seeds:
                db.add(Position(ticker=t,expiry=e,option_type=typ,strike=k,quantity=q,multiplier=100,
                    bloomberg_security=option_security(t,e,typ,k),underlying_security=f"{t} US Equity",active=True,source="seed"))
        db.commit()
init_data()

def current_user(request):
    username=read_session(request.cookies.get("session",""))
    if not username:return None
    with SessionLocal() as db:
        u=db.query(User).filter_by(username=username,active=True).first()
        return None if not u else {"id":u.id,"username":u.username,"role":u.role}

def require_user(request):
    u=current_user(request)
    if not u:raise HTTPException(401,"Login required")
    return u

def require_admin(request):
    u=require_user(request)
    if u["username"]!=ADMIN_USER or u["role"]!="admin":raise HTTPException(403,"Admin only")
    return u

def page(title,body,user=None,refresh=None):
    adminlink='<a href="/admin">Admin</a>' if user and user["username"]==ADMIN_USER else ""
    if user:
        auth=(
            f'<a href="/overview">Overview</a>'
            f'<a href="/">Fortress Bloomberg</a>'
            f'<a href="/bloomberg-daily-return">Fortress Daily Return</a>'
            f'<a href="/daily-return">Fortress Yahoo</a>'
            f'<a href="/opportunity">Opportunity Bloomberg</a>'
            f'<a href="/opportunity/bloomberg-daily-return">Opportunity Daily Return</a>'
            f'<a href="/opportunity/yahoo-daily-return">Opportunity Yahoo</a>'
            f'{adminlink}<a href="/logout">Logout</a>'
            f'<span class="nav-user">{html.escape(user["username"])}</span>'
        )
    else:
        auth=""
    rf=f'<meta http-equiv="refresh" content="{refresh}">' if refresh else ""
    return HTMLResponse(f"""<!doctype html><html><head><meta charset="utf-8">{rf}<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)}</title><style>
*{{box-sizing:border-box}}body{{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f5f7fb;color:#172033}}.wrap{{max-width:1600px;margin:auto;padding:24px 28px 40px}}
.topbar{{background:#142b4d;color:#fff;box-shadow:0 1px 3px rgba(0,0,0,.12)}}.topinner{{max-width:1600px;margin:auto;padding:0 28px;display:flex;align-items:center;min-height:58px;gap:22px}}.brand{{color:#fff;text-decoration:none;font-size:20px;font-weight:800;white-space:nowrap;margin-right:auto}}nav{{display:flex;align-items:center;gap:4px;flex-wrap:wrap;justify-content:flex-end}}nav a{{color:#fff;text-decoration:none;padding:19px 10px 17px;font-size:13px;border-bottom:2px solid transparent;white-space:nowrap}}nav a:hover{{border-bottom-color:#fff}}.nav-user{{font-size:12px;margin-left:8px;opacity:.9}}
.card{{background:#fff;border:1px solid #dfe5ef;border-radius:12px;padding:16px;margin-bottom:14px}}.grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}}.metric .v{{font-size:25px;font-weight:700;margin-top:6px}}.muted{{color:#6d7b92;font-size:13px}}.ok{{color:#067647}}.bad{{color:#b42318}}.notice{{padding:10px 12px;background:#fff7e6;border:1px solid #ffd591;border-radius:8px}}
table{{width:100%;border-collapse:collapse}}th,td{{border-bottom:1px solid #e6ebf2;padding:9px;text-align:right;font-size:13px}}th:first-child,td:first-child{{text-align:left}}input,select,button{{padding:8px;border:1px solid #c9d3e1;border-radius:7px}}button{{background:#172b4d;color:white;cursor:pointer}}.danger{{background:#b42318}}.secondary{{background:#52657d}}.row{{display:flex;gap:8px;align-items:end;flex-wrap:wrap}}label{{display:flex;flex-direction:column;font-size:12px;color:#52657d;gap:4px}}h1{{font-size:32px;margin:4px 0 26px}}h2{{color:#13203a}}
@media(max-width:1250px){{.topinner{{align-items:flex-start;padding-top:8px;padding-bottom:8px}}nav a{{padding:8px 7px}}.grid{{grid-template-columns:repeat(2,1fr)}}}}@media(max-width:760px){{.topinner{{display:block}}.brand{{display:block;padding:12px 0}}nav{{justify-content:flex-start}}.wrap{{padding:18px 14px}}.grid{{grid-template-columns:1fr}}}}
</style></head><body><div class="topbar"><div class="topinner"><a class="brand" href="/overview">RPD Fund Management</a><nav>{auth}</nav></div></div><div class="wrap">{body}</div></body></html>""")

def admin_tabs():
    return '<p><a href="/overview">Overview</a> · <a href="/admin">Fortress Portfolio</a> · <a href="/admin/import">Fortress Nirvana Import</a> · <a href="/admin/daily-return">Fortress Daily Return</a> · <a href="/admin/opportunity">Opportunity Setup</a> · <a href="/admin/users">Viewer Users</a> · <a href="/admin/audit">Audit Log</a></p>'

@app.get("/health")
def health():return {"ok":True}

@app.get("/login")
def login_get():
    return page("Login","""<div class="card" style="max-width:420px;margin:80px auto"><h2>Sign in</h2>
<form method="post"><label>Username<input name="username"></label><br><label>Password<input type="password" name="password"></label><br><button>Sign in</button></form></div>""")

@app.post("/login")
def login_post(username:str=Form(...),password:str=Form(...)):
    name=username.strip().lower()
    with SessionLocal() as db:
        u=db.query(User).filter_by(username=name,active=True).first()
        if not u or not verify_password(password,u.password_hash):
            return page("Login failed",'<div class="card"><h2>Login failed</h2><a href="/login">Try again</a></div>')
        if name==ADMIN_USER:audit(db,name,"ADMIN_LOGIN","Admin signed in");db.commit()
    resp=RedirectResponse("/",303);resp.set_cookie("session",sign_session(name),httponly=True,secure=True,samesite="strict",max_age=604800)
    return resp

@app.get("/logout")
def logout():
    r=RedirectResponse("/login",303);r.delete_cookie("session");return r

def latest_snapshot(db):
    s=db.query(Snapshot).order_by(Snapshot.id.desc()).first()
    if not s:return None,None
    try:p=json.loads(s.payload)
    except Exception:return None,None
    try:
        ts=datetime.fromisoformat(p["timestamp_utc"].replace("Z","+00:00"))
        age=(datetime.now(timezone.utc)-ts.astimezone(timezone.utc)).total_seconds()
    except Exception:age=999999
    return p,age

def format_new_york_time(timestamp_utc):
    try:
        ts=datetime.fromisoformat(str(timestamp_utc).replace("Z","+00:00"))
        ny=ts.astimezone(ZoneInfo("America/New_York"))
        return ny.strftime("%d %b %Y, %H:%M:%S %Z")
    except Exception:
        return str(timestamp_utc or "")



def _json_setting(db,key,default=None):
    raw=get_setting(db,key,"")
    if not raw:
        return default
    try:
        x=json.loads(raw)
        return x
    except Exception:
        return default

def _set_json_setting(db,key,value):
    set_setting(db,key,json.dumps(value,separators=(",",":")))

def _opportunity_baseline(db):
    x=_json_setting(db,"opportunity_daily_baseline",None)
    return x if isinstance(x,dict) else None

def _opportunity_exposure(db):
    x=_json_setting(db,"opportunity_exposure_config",None)
    return x if isinstance(x,dict) else None

def _latest_opportunity_snapshot(db):
    p=_json_setting(db,"opportunity_bloomberg_snapshot",None)
    if not isinstance(p,dict):
        return None,None
    try:
        ts=datetime.fromisoformat(str(p.get("timestamp_utc") or "").replace("Z","+00:00"))
        age=(datetime.now(timezone.utc)-ts.astimezone(timezone.utc)).total_seconds()
    except Exception:
        age=999999
    return p,age

def _latest_yahoo_v2_fund_snapshot(db,fund_key):
    root=_json_setting(db,"yahoo_collector_snapshot_v2",None)
    if not isinstance(root,dict):
        return None
    funds=root.get("funds") or {}
    x=funds.get(fund_key)
    return x if isinstance(x,dict) else None

def _quote_data_from_snapshot(snapshot):
    out={}
    if not snapshot:
        return out
    for r in snapshot.get("positions",[]) or []:
        symbol=str(r.get("symbol") or "").upper().strip()
        if not symbol:
            continue
        out[symbol]={
            "mark":r.get("mark"),"bid":r.get("bid"),"ask":r.get("ask"),"last":r.get("last"),
            "source":r.get("source") or "","note":r.get("note") or "",
        }
    return out

def _dr_position_key(p):
    inst=str(p.get("instrument_type") or "OPTION").upper()
    if inst=="OPTION":
        return ("OPTION",str(p.get("ticker") or "").upper(),str(p.get("expiry") or ""),str(p.get("option_type") or "").upper(),round(float(p.get("strike") or 0),6))
    return ("EQUITY",str(p.get("ticker") or "").upper())

def _bloomberg_mark_map(snapshot):
    out={}
    if not snapshot:
        return out
    for r in snapshot.get("positions",[]) or []:
        try:
            inst=str(r.get("instrument_type") or "OPTION").upper()
            if inst=="OPTION":
                key=("OPTION",str(r.get("ticker") or "").upper(),str(r.get("expiry") or ""),str(r.get("option_type") or "").upper(),round(float(r.get("strike") or 0),6))
            else:
                key=("EQUITY",str(r.get("ticker") or "").upper())
            out[key]={
                "mark":r.get("market_mark",r.get("option_mark")),
                "source":r.get("market_source",r.get("option_mark_source")) or "",
                "bid":r.get("market_bid",r.get("option_bid")),
                "ask":r.get("market_ask",r.get("option_ask")),
                "last":r.get("market_last",r.get("option_last")),
            }
        except Exception:
            continue
    return out

def _calc_bloomberg_daily(positions,nav,snapshot,manual_field="bloomberg_manual_price"):
    marks=_bloomberg_mark_map(snapshot)
    total=0.0; valid=0; rows=[]
    for p in positions:
        q=marks.get(_dr_position_key(p),{})
        market=q.get("mark"); source=q.get("source") or ""; manual=p.get(manual_field)
        if source in ("PX_BID/ASK_MID","PX_ASK_HALF","EQUITY_BID/ASK_MID") and market is not None:
            current=float(market); eff="BLOOMBERG MID" if source!="PX_ASK_HALF" else "BLOOMBERG ASK/2"
        elif manual is not None:
            current=float(manual); eff="MANUAL OVERRIDE"
        elif market is not None:
            current=float(market); eff="BLOOMBERG LAST"
        else:
            current=None; eff="UNAVAILABLE"
        prev=p.get("baseline_price"); pnl=chg=contrib=None
        if current is not None and prev is not None:
            chg=current-float(prev); pnl=chg*float(p.get("quantity") or 0)*float(p.get("multiplier") or 1); total+=pnl; valid+=1
            contrib=pnl/float(nav)*100 if nav else None
        rows.append({**p,"market_mark":market,"market_source":source,"bid":q.get("bid"),"ask":q.get("ask"),"last":q.get("last"),"current_mark":current,"effective_source":eff,"change":chg,"estimated_pnl":pnl,"contribution_pct":contrib})
    n=len(positions)
    return {"rows":rows,"estimated_pnl":total,"estimated_return_pct":total/float(nav)*100 if nav else 0.0,"coverage_valid":valid,"coverage_total":n,"complete":bool(n and valid==n)}

def _opp_limit(db):
    raw=str(get_setting(db,"opportunity_limit_pct","") or "").strip()
    try:return float(raw) if raw else None
    except Exception:return None

def _bump_opportunity_config(db,actor,reason):
    v=int(get_setting(db,"opportunity_config_version","1") or 1)+1
    set_setting(db,"opportunity_config_version",v)
    audit(db,actor,"OPPORTUNITY_CONFIG_PUBLISHED",f"v{v}: {reason}")
    return v

def _parse_opportunity_exposure_xlsx(data:bytes,filename:str="Exposure.xlsx"):
    """
    Opportunity-only Exposure by Underlying parser.

    Important:
    - This function is used ONLY for Opportunity Delta / Exposure import.
    - It does not change Fortress parsing.
    - It does not change Opportunity/Fortress Daily Return baselines, Bid/Ask
      logic, Yahoo snapshots, Bloomberg Daily Return, or collector behavior.
    """
    wb=load_workbook(BytesIO(data),read_only=True,data_only=True)

    # Nirvana exports can change the workbook's active sheet depending on
    # report filters. Prefer the named Exposure Summary sheet, then scan all
    # sheets as a fallback instead of trusting wb.active.
    sheets=[]
    for s in wb.worksheets:
        if str(s.title or "").strip().lower()=="exposure summary":
            sheets.insert(0,s)
        else:
            sheets.append(s)

    ws=None; vals=None; header_i=None
    required_headers={"symbol","nav","position","delta"}
    for candidate in sheets:
        candidate_vals=list(candidate.iter_rows(values_only=True))
        # Search generously because Nirvana may add/remove report metadata rows.
        for i,row in enumerate(candidate_vals[:60]):
            norm=[str(x or "").strip() for x in (row or [])]
            norm_lower={x.lower() for x in norm if x}
            if required_headers.issubset(norm_lower):
                ws=candidate; vals=candidate_vals; header_i=i
                break
        if header_i is not None:
            break

    if ws is None or vals is None or header_i is None:
        raise ValueError("Could not locate Opportunity Exposure Summary header row (Symbol/NAV/Position/Delta).")

    headers=[str(x or "").strip() for x in vals[header_i]]
    hm={h.lower():i for i,h in enumerate(headers) if h}
    for req in ("symbol","nav","position","delta"):
        if req not in hm:
            raise ValueError(f"Opportunity Exposure is missing required column: {req.title()}")

    report_date=""
    for row in vals[:header_i]:
        line=" ".join(str(x or "") for x in row)
        m=re.search(r"Report Date:\s*(\d{1,2}/\d{1,2}/\d{4})",line,re.I)
        if m:
            report_date=datetime.strptime(m.group(1),"%m/%d/%Y").date().isoformat();break

    nav=None; positions=[]; seen=set()

    # Cash / money-market instruments must NOT contribute to Opportunity
    # Bloomberg Delta Monitor. NAV is still the report's Grand Total NAV.
    # Keeping the denominator unchanged is intentional.
    excluded_symbols={"FGTXX","MPFXX","USD","CASH"}

    exp_col=hm.get("expiration date")

    for row in vals[header_i+1:]:
        try:
            sym=str(row[hm["symbol"]] or "").strip()
        except Exception:
            continue
        if not sym:
            continue

        if sym.lower().startswith("grand total"):
            try: nav=float(row[hm["nav"]])
            except Exception: nav=None
            continue

        # Opportunity Delta-only filter. This does not affect Daily Return.
        clean_sym=sym.upper().strip()
        if clean_sym in excluded_symbols:
            continue

        try: qty=float(row[hm["position"]])
        except Exception: continue
        try: delta=float(row[hm["delta"]])
        except Exception: continue
        if abs(qty)<1e-12:
            continue

        if sym.startswith("O:"):
            m=re.match(r"^O:([^ ]+)\s+(\d{2})([A-X])([0-9.]+)D(\d{1,2})$",sym,re.I)
            if not m:
                continue
            ticker=m.group(1).upper(); code=m.group(3).upper(); strike=float(m.group(4))
            typ="C" if "A"<=code<="L" else "P"
            if exp_col is None:
                continue
            expv=row[exp_col]
            if hasattr(expv,"date"):
                expiry=expv.date().isoformat()
            else:
                try: expiry=datetime.strptime(str(expv)[:10],"%Y-%m-%d").date().isoformat()
                except Exception: continue
            key=("OPTION",ticker,expiry,typ,round(strike,6),round(qty,6))
            if key in seen:
                continue
            seen.add(key)
            positions.append({
                "id":len(positions)+1,
                "instrument_type":"OPTION",
                "ticker":ticker,
                "expiry":expiry,
                "option_type":typ,
                "strike":strike,
                "quantity":qty,
                "multiplier":100.0,
                "bloomberg_security":option_security(ticker,expiry,typ,strike),
                "underlying_security":f"{ticker} US Equity",
                "nirvana_delta":delta
            })

        elif abs(delta-1.0)<1e-9:
            ticker=clean_sym
            # Only real equities/ETFs remain here. Money-market/cash symbols
            # were excluded above.
            key=("EQUITY",ticker,round(qty,6))
            if key in seen:
                continue
            seen.add(key)
            positions.append({
                "id":len(positions)+1,
                "instrument_type":"EQUITY",
                "ticker":ticker,
                "expiry":"",
                "option_type":"",
                "strike":None,
                "quantity":qty,
                "multiplier":1.0,
                "bloomberg_security":f"{ticker} US Equity",
                "underlying_security":f"{ticker} US Equity",
                "nirvana_delta":delta
            })

    if nav is None or nav<=0:
        raise ValueError("Could not read Grand Total NAV from Opportunity exposure report.")
    if not positions:
        raise ValueError("No detailed Opportunity exposure positions were parsed.")

    return {
        "filename":filename,
        "report_date":report_date,
        "nav_usd":nav,
        "positions":positions
    }

def _opp_baseline_positions_for_calc(b):
    return list((b or {}).get("positions") or [])

def _market_config_position_from_daily(p):
    inst=str(p.get("instrument_type") or "OPTION").upper(); ticker=str(p.get("ticker") or "").upper()
    if inst=="OPTION":
        return {"instrument_type":"OPTION","ticker":ticker,"expiry":p.get("expiry") or "","option_type":p.get("option_type") or "","strike":p.get("strike"),"quantity":p.get("quantity"),"multiplier":p.get("multiplier",100),"bloomberg_security":option_security(ticker,p.get("expiry") or "",p.get("option_type") or "P",p.get("strike")),"underlying_security":f"{ticker} US Equity"}
    return {"instrument_type":"EQUITY","ticker":ticker,"expiry":"","option_type":"","strike":None,"quantity":p.get("quantity"),"multiplier":1.0,"bloomberg_security":f"{ticker} US Equity","underlying_security":f"{ticker} US Equity"}

def _merge_market_positions(delta_positions,daily_positions):
    merged={}
    def key(p):
        inst=str(p.get("instrument_type") or "OPTION").upper()
        if inst=="OPTION":return ("OPTION",str(p.get("ticker") or "").upper(),str(p.get("expiry") or ""),str(p.get("option_type") or "").upper(),round(float(p.get("strike") or 0),6))
        return (inst,str(p.get("ticker") or "").upper())
    for p in delta_positions or []:
        x=dict(p);x["use_for_delta"]=True;x["use_for_daily_return"]=False;merged[key(x)]=x
    for p in daily_positions or []:
        x=_market_config_position_from_daily(p);k=key(x)
        if k in merged: merged[k]["use_for_daily_return"]=True
        else:
            x["use_for_delta"]=False;x["use_for_daily_return"]=True;merged[k]=x
    out=[]
    for i,x in enumerate(merged.values(),1):
        x=dict(x);x["id"]=x.get("id") or i;out.append(x)
    return out
@app.get("/overview")
def overview(request:Request):
    u=require_user(request)
    with SessionLocal() as db:
        # Fortress delta
        fp,fage=latest_snapshot(db); fcfg=int(get_setting(db,"config_version","1")); flimit=float(get_setting(db,"limit_pct","15"))
        if fp:
            fm=fp.get("mode","LIVE"); fclosed=fm=="MARKET_CLOSED"; fstale=(fage>STALE_AFTER_SECONDS) and not fclosed
            fdelta_status="MARKET CLOSED" if fclosed else ("OFFLINE / STALE" if fstale else fm); fdelta_sub="Final closing snapshot" if fclosed else f"Age {fage:.1f}s"
            fdelta_pct=float(fp.get("delta_exposure_pct",0) or 0); fdelta_usd=float(fp.get("delta_exposure_usd",0) or 0); fcov=fp.get("coverage",{}) or {}; fdelta_cov=f'{fcov.get("valid",0)}/{fcov.get("total",0)}'; fcollector=fp.get("config_version","—")
            fbuffer=fp.get("buffer_pct"); fbuffer=float(fbuffer) if fbuffer is not None else flimit-fdelta_pct
        else:
            fdelta_status="WAITING";fdelta_sub="No collector snapshot";fdelta_pct=fdelta_usd=0.0;fdelta_cov="0/0";fcollector="—";fbuffer=flimit

        # Fortress daily summaries
        fb=_active_daily_baseline(db)
        if fb:
            fdb=db.query(DailyReturnPosition).filter_by(baseline_id=fb.id).order_by(DailyReturnPosition.id).all()
            fpos=[{"id":x.id,"security_name":x.security_name,"instrument_type":x.instrument_type,"ticker":x.ticker,"expiry":x.expiry,"option_type":x.option_type,"strike":x.strike,"quantity":x.quantity,"multiplier":x.multiplier,"baseline_price":x.baseline_price,"baseline_market_value":x.baseline_market_value,"yahoo_manual_price":x.yahoo_manual_price,"bloomberg_manual_price":x.bloomberg_manual_price} for x in fdb]
            fnav=float(fb.nav_usd or 0); freport=str(fb.report_date or "—")
            fys=_latest_yahoo_v2_fund_snapshot(db,"fortress") or _latest_yahoo_collector_snapshot(db)
            fycalc=calculate_mark_to_market_return(fpos,fnav,quote_data=_quote_data_from_snapshot(fys)); fy_pnl=float(fycalc.get("estimated_pnl") or 0);fy_ret=float(fycalc.get("estimated_return_pct") or 0);fy_cov=f'{fycalc.get("coverage_valid",0)}/{fycalc.get("coverage_total",0)}';fy_status="FULL" if fycalc.get("complete") else "PARTIAL"
            fbcalc=_calc_bloomberg_daily(fpos,fnav,fp);fb_pnl=float(fbcalc["estimated_pnl"]);fb_ret=float(fbcalc["estimated_return_pct"]);fb_cov=f'{fbcalc["coverage_valid"]}/{fbcalc["coverage_total"]}';fb_status="FULL" if fbcalc["complete"] else "PARTIAL"
        else:
            fnav=0;freport="—";fy_pnl=fy_ret=fb_pnl=fb_ret=0;fy_cov=fb_cov="0/0";fy_status=fb_status="NO BASELINE"

        # Opportunity delta
        op,oage=_latest_opportunity_snapshot(db); ocfg=int(get_setting(db,"opportunity_config_version","1") or 1); olimit=_opp_limit(db)
        if op:
            om=op.get("mode","LIVE"); oclosed=om=="MARKET_CLOSED"; ostale=(oage>STALE_AFTER_SECONDS) and not oclosed
            odelta_status="MARKET CLOSED" if oclosed else ("OFFLINE / STALE" if ostale else om); odelta_sub="Final closing snapshot" if oclosed else f"Age {oage:.1f}s"
            odelta_pct=float(op.get("delta_exposure_pct",0) or 0);odelta_usd=float(op.get("delta_exposure_usd",0) or 0);ocov=op.get("coverage",{}) or {};odelta_cov=f'{ocov.get("valid",0)}/{ocov.get("total",0)}';ocollector=op.get("config_version","—"); obuffer=op.get("buffer_pct")
            if obuffer is not None: obuffer=float(obuffer)
        else:
            odelta_status="WAITING";odelta_sub="No collector snapshot";odelta_pct=odelta_usd=0.0;odelta_cov="0/0";ocollector="—";obuffer=None

        ob=_opportunity_baseline(db)
        if ob:
            opos=_opp_baseline_positions_for_calc(ob);onav=float(ob.get("nav_usd") or 0);oreport=str(ob.get("report_date") or "—")
            oys=_latest_yahoo_v2_fund_snapshot(db,"opportunity"); oycalc=calculate_mark_to_market_return(opos,onav,quote_data=_quote_data_from_snapshot(oys));oy_pnl=float(oycalc.get("estimated_pnl") or 0);oy_ret=float(oycalc.get("estimated_return_pct") or 0);oy_cov=f'{oycalc.get("coverage_valid",0)}/{oycalc.get("coverage_total",0)}';oy_status="FULL" if oycalc.get("complete") else "PARTIAL"
            obcalc=_calc_bloomberg_daily(opos,onav,op);ob_pnl=float(obcalc["estimated_pnl"]);ob_ret=float(obcalc["estimated_return_pct"]);ob_cov=f'{obcalc["coverage_valid"]}/{obcalc["coverage_total"]}';ob_status="FULL" if obcalc["complete"] else "PARTIAL"
        else:
            onav=0;oreport="—";oy_pnl=oy_ret=ob_pnl=ob_ret=0;oy_cov=ob_cov="0/0";oy_status=ob_status="NO BASELINE"

    def cls(v): return "ok" if v>=0 else "bad"
    def limit_txt(x): return "—" if x is None else f"{x:.3f}%"
    def buffer_txt(x): return "—" if x is None else f"{x:+.3f}%"
    def fund_column(title,accent,delta_link,dstatus,dsub,dpct,dusd,limit,buffer,dcov,collector,portalv,bd_link,bd_status,bd_ret,bd_pnl,nav,report,bd_cov,y_link,y_status,y_ret,y_pnl,y_cov):
        return f'''<section class="fundcol"><div class="fundtitle {accent}"><h2>{title}</h2></div>
        <div class="summarypanel"><div class="panelhead"><div class="headgroup"><h3>Bloomberg Delta Monitor</h3><span class="delay-badge bloomberg">◷&nbsp; 2-min delayed</span></div><a href="{delta_link}">Open full monitor →</a></div><div class="mini-grid">
        <div class="mini"><span>Feed</span><b class="ok">{html.escape(dstatus)}</b><small>{html.escape(dsub)}</small></div><div class="mini"><span>Delta Exposure</span><b>{dpct:.3f}%</b><small>${dusd:,.0f}</small></div><div class="mini"><span>Limit</span><b>{limit_txt(limit)}</b></div><div class="mini"><span>Buffer</span><b class="{cls(buffer or 0) if buffer is not None else ''}">{buffer_txt(buffer)}</b></div><div class="mini"><span>Coverage</span><b>{dcov}</b><small>Collector v{collector} / Portal v{portalv}</small></div></div></div>
        <div class="summarypanel"><div class="panelhead"><div class="headgroup"><h3>Bloomberg Daily Return</h3><span class="delay-badge bloomberg">◷&nbsp; 2-min delayed</span></div><a href="{bd_link}">Open details →</a></div><div class="mini-grid">
        <div class="mini"><span>Status</span><b class="ok">{bd_status}</b><small>Bloomberg marks</small></div><div class="mini"><span>Estimated Daily Return</span><b class="{cls(bd_ret)}">{bd_ret:+.4f}%</b></div><div class="mini"><span>Estimated P&amp;L</span><b class="{cls(bd_pnl)}">${bd_pnl:+,.2f}</b></div><div class="mini"><span>Baseline NAV</span><b>${nav:,.2f}</b><small>Report {html.escape(report)}</small></div><div class="mini"><span>Coverage</span><b>{bd_cov}</b><small>Bloomberg marks</small></div></div></div>
        <div class="summarypanel"><div class="panelhead"><div class="headgroup"><h3>Yahoo Daily Return</h3><span class="delay-badge yahoo">◷&nbsp; 15-min delayed</span></div><a href="{y_link}">Open details →</a></div><div class="mini-grid">
        <div class="mini"><span>Status</span><b class="ok">{y_status}</b><small>Mark-to-market estimate</small></div><div class="mini"><span>Estimated Daily Return</span><b class="{cls(y_ret)}">{y_ret:+.4f}%</b></div><div class="mini"><span>Estimated P&amp;L</span><b class="{cls(y_pnl)}">${y_pnl:+,.2f}</b></div><div class="mini"><span>Baseline NAV</span><b>${nav:,.2f}</b><small>Report {html.escape(report)}</small></div><div class="mini"><span>Coverage</span><b>{y_cov}</b><small>Yahoo local collector marks</small></div></div></div></section>'''

    body=f'''<style>.funds{{display:grid;grid-template-columns:1fr 1fr;gap:18px;align-items:stretch}}.fundcol{{min-width:0;display:grid;grid-template-rows:auto repeat(3,minmax(0,1fr));row-gap:14px;align-self:stretch}}.fundtitle{{padding:12px 18px;border-radius:12px 12px 0 0;margin-bottom:0}}.fundtitle h2{{margin:0;font-size:24px}}.fundtitle.fortress{{background:#eef6ff}}.fundtitle.opportunity{{background:#f5efff}}.summarypanel{{background:#fff;border:1px solid #dfe5ef;border-radius:13px;padding:14px;margin-bottom:0;display:flex;flex-direction:column;height:100%;box-sizing:border-box}}.panelhead{{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:12px}}.headgroup{{display:flex;align-items:center;gap:9px;min-width:0;flex-wrap:wrap}}.panelhead h3{{margin:0;font-size:19px}}.panelhead a{{font-size:13px;white-space:nowrap}}.delay-badge{{display:inline-flex;align-items:center;white-space:nowrap;border-radius:999px;padding:4px 10px;font-size:11.5px;font-weight:600;line-height:1;border:1px solid transparent}}.delay-badge.bloomberg{{background:#eef3fb;color:#2d4771;border-color:#dde7f5}}.delay-badge.yahoo{{background:#fff1c9;color:#805b00;border-color:#f6dda0}}.mini-grid{{display:grid;grid-template-columns:.95fr 1fr 1.18fr 1.25fr .95fr;gap:9px;flex:1;align-items:stretch}}.mini{{border:1px solid #e4e9f1;border-radius:10px;padding:10px;min-height:98px;min-width:0;overflow:hidden;display:flex;flex-direction:column;height:100%;box-sizing:border-box}}.mini span{{display:block;color:#73819a;font-size:12px;line-height:1.25;min-height:30px}}.mini small{{display:block;color:#73819a;font-size:12px;line-height:1.25}}.mini b{{display:block;font-size:14.5px;line-height:1.15;margin:6px 0 4px;white-space:nowrap;letter-spacing:-.35px;font-variant-numeric:tabular-nums}}.mini:first-child b{{white-space:normal;font-size:14px}}.mini:nth-child(3) b,.mini:nth-child(4) b{{font-size:14px}}@media(min-width:1700px){{.mini b{{font-size:15.5px}}.mini:first-child b{{font-size:15px}}.mini:nth-child(3) b,.mini:nth-child(4) b{{font-size:15px}}}}@media(max-width:1250px){{.funds{{grid-template-columns:1fr}}.mini-grid{{grid-template-columns:.95fr 1fr 1.18fr 1.25fr .95fr}}.mini b{{font-size:18px}}.mini:first-child b{{font-size:17px}}.mini:nth-child(3) b,.mini:nth-child(4) b{{font-size:17px}}}}@media(max-width:800px){{.mini-grid{{grid-template-columns:1fr 1fr}}.mini b,.mini:first-child b,.mini:nth-child(3) b,.mini:nth-child(4) b{{font-size:17px}}.delay-badge{{font-size:11px;padding:4px 8px}}}}</style>
    <h1>RPD Fund Management — Overview</h1><div class="funds">
    {fund_column("RPD Opportunity Fund","opportunity","/opportunity",odelta_status,odelta_sub,odelta_pct,odelta_usd,olimit,obuffer,odelta_cov,ocollector,ocfg,"/opportunity/bloomberg-daily-return",ob_status,ob_ret,ob_pnl,onav,oreport,ob_cov,"/opportunity/yahoo-daily-return",oy_status,oy_ret,oy_pnl,oy_cov)}
    {fund_column("RPD Fortress Fund","fortress","/",fdelta_status,fdelta_sub,fdelta_pct,fdelta_usd,flimit,fbuffer,fdelta_cov,fcollector,fcfg,"/bloomberg-daily-return",fb_status,fb_ret,fb_pnl,fnav,freport,fb_cov,"/daily-return",fy_status,fy_ret,fy_pnl,fy_cov)}
    </div>'''
    return page("Overview",body,u,60)

@app.get("/")
def dashboard(request:Request):
    u=current_user(request)
    if not u:return RedirectResponse("/login",303)
    with SessionLocal() as db:
        p,age=latest_snapshot(db);cfg=int(get_setting(db,"config_version","1"));limit=float(get_setting(db,"limit_pct","15"))
        if not p:
            return page("Delta Monitor",f'<div class="card"><h1>RPD Fortress Fund — Bloomberg Delta Monitor</h1><div class="notice">WAITING — No collector snapshot yet.</div><p>Published config v{cfg}</p></div>',u,5)
        mode=p.get("mode","LIVE")
        closed_mode=(mode=="MARKET_CLOSED")
        stale=(age>STALE_AFTER_SECONDS) and not closed_mode
        status="MARKET CLOSED" if closed_mode else ("OFFLINE / STALE" if stale else mode)
        pct=p.get("delta_exposure_pct",0) or 0;buf=p.get("buffer_pct",0) or 0;exp=p.get("delta_exposure_usd",0) or 0;cov=p.get("coverage",{})
        rows=""
        for r in sorted(p.get("positions",[]),key=lambda x:abs(x.get("contribution_pct",0)),reverse=True):
            rows+=f'<tr><td>{html.escape(r.get("ticker",""))} {html.escape(r.get("expiry",""))} {html.escape(r.get("option_type",""))}{r.get("strike")}</td><td>{r.get("quantity",0):,.0f}</td><td>${r.get("delta_exposure_usd",0):,.0f}</td><td>{r.get("contribution_pct",0):.3f}%</td></tr>'
        mismatch=p.get("config_version")!=cfg
        body=f"""<h1>RPD Fortress Fund — Bloomberg Delta Monitor</h1><div class="grid">
<div class="card metric"><div class="muted">Feed</div><div class="v {'bad' if stale else 'ok'}">{status}</div><div class="muted">{'Final closing snapshot' if closed_mode else f'Age {age:.1f}s'}</div></div>
<div class="card metric"><div class="muted">Delta Exposure</div><div class="v">{pct:.3f}%</div><div class="muted">${exp:,.0f}</div></div>
<div class="card metric"><div class="muted">Limit</div><div class="v">{limit:.3f}%</div></div>
<div class="card metric"><div class="muted">Buffer</div><div class="v">{buf:.3f}%</div></div>
<div class="card metric"><div class="muted">Coverage</div><div class="v">{cov.get("valid",0)}/{cov.get("total",0)}</div><div class="muted">Collector v{p.get("config_version")} / Portal v{cfg}</div></div></div>
{('<div class="notice">A new portfolio config is published; collector has not reported it yet.</div>' if mismatch else '')}
<div class="card"><div class="muted">Last update (New York)</div>{html.escape(format_new_york_time(p.get("timestamp_utc","")))} · Collector: {html.escape(p.get("collector_id",""))}</div>
<div class="card"><table><thead><tr><th>Position</th><th>Qty</th><th>Delta Exposure $</th><th>Contribution % NAV</th></tr></thead><tbody>{rows}</tbody></table></div>"""
        return page("Delta Monitor",body,u,3)


@app.get("/opportunity")
def opportunity_dashboard(request:Request):
    u=require_user(request)
    with SessionLocal() as db:
        p,age=_latest_opportunity_snapshot(db);cfg=int(get_setting(db,"opportunity_config_version","1") or 1);limit=_opp_limit(db)
        if not p:
            return page("Opportunity Bloomberg Delta Monitor",f'<div class="card"><h1>RPD Opportunity Fund — Bloomberg Delta Monitor</h1><div class="notice">WAITING — No Opportunity collector snapshot yet.</div><p>Published config v{cfg}</p></div>',u,5)
        mode=p.get("mode","LIVE");closed=mode=="MARKET_CLOSED";stale=(age>STALE_AFTER_SECONDS) and not closed;status="MARKET CLOSED" if closed else ("OFFLINE / STALE" if stale else mode)
        pct=float(p.get("delta_exposure_pct",0) or 0);exp=float(p.get("delta_exposure_usd",0) or 0);buf=p.get("buffer_pct");cov=p.get("coverage",{}) or {};rows=""
        for r in sorted([x for x in (p.get("positions",[]) or []) if x.get("use_for_delta")],key=lambda x:abs(float(x.get("contribution_pct") or 0)),reverse=True):
            inst=str(r.get("instrument_type") or "OPTION").upper(); label=(f'{r.get("ticker","")} {r.get("expiry","")} {r.get("option_type","")}{r.get("strike")}' if inst=="OPTION" else str(r.get("ticker") or ""))
            rows+=f'<tr><td>{html.escape(label)}</td><td>{float(r.get("quantity") or 0):,.0f}</td><td>${float(r.get("delta_exposure_usd") or 0):,.0f}</td><td>{float(r.get("contribution_pct") or 0):.3f}%</td></tr>'
        limit_html='—' if limit is None else f'{limit:.3f}%';buf_html='—' if buf is None else f'{float(buf):+.3f}%'
        body=f'''<h1>RPD Opportunity Fund — Bloomberg Delta Monitor</h1><div class="grid"><div class="card metric"><div class="muted">Feed</div><div class="v {'bad' if stale else 'ok'}">{status}</div><div class="muted">{'Final closing snapshot' if closed else f'Age {age:.1f}s'}</div></div><div class="card metric"><div class="muted">Delta Exposure</div><div class="v">{pct:.3f}%</div><div class="muted">${exp:,.0f}</div></div><div class="card metric"><div class="muted">Limit</div><div class="v">{limit_html}</div></div><div class="card metric"><div class="muted">Buffer</div><div class="v">{buf_html}</div></div><div class="card metric"><div class="muted">Coverage</div><div class="v">{cov.get('valid',0)}/{cov.get('total',0)}</div><div class="muted">Collector v{p.get('config_version')} / Portal v{cfg}</div></div></div><div class="card"><div class="muted">Last update (New York)</div>{html.escape(format_new_york_time(p.get('timestamp_utc','')))} · Collector: {html.escape(str(p.get('collector_id','')))}</div><div class="card"><table><thead><tr><th>Position</th><th>Qty</th><th>Delta Exposure $</th><th>Contribution % NAV</th></tr></thead><tbody>{rows}</tbody></table></div>'''
        return page("Opportunity Bloomberg Delta Monitor",body,u,3)

def _render_opportunity_yahoo(request:Request):
    u=require_user(request)
    with SessionLocal() as db:
        b=_opportunity_baseline(db);snap=_latest_yahoo_v2_fund_snapshot(db,"opportunity")
    if not b:
        extra=' <a href="/admin/opportunity">Upload Opportunity P&amp;L Report</a>' if u["username"]==ADMIN_USER else ''
        return page("Opportunity Yahoo Daily Return",f'<h1>RPD Opportunity Fund — Yahoo Daily Return</h1><div class="card"><div class="notice">No Opportunity P&amp;L baseline loaded.{extra}</div></div>',u)
    positions=_opp_baseline_positions_for_calc(b);calc=calculate_mark_to_market_return(positions,float(b.get("nav_usd") or 0),quote_data=_quote_data_from_snapshot(snap));status="FULL" if calc["complete"] else "PARTIAL";ret=float(calc.get("estimated_return_pct") or 0);pnl=float(calc.get("estimated_pnl") or 0);rows=""
    for r in calc["rows"]:
        strike="" if r.get("strike") is None else f'{r["strike"]:g}';posname=(f'{r["ticker"]} {r["expiry"]} {r["option_type"]}{strike}' if r["instrument_type"]=="OPTION" else r["ticker"]);manual=r.get("manual_price");manual_val='' if manual is None else f'{float(manual):.4f}'
        if u["username"]==ADMIN_USER:
            mh=f'<form method="post" action="/admin/opportunity/manual-price" style="display:flex;gap:4px"><input type="hidden" name="position_id" value="{r["id"]}"><input type="hidden" name="source" value="YAHOO"><input type="number" step="0.0001" min="0" name="price" value="{manual_val}" style="width:82px"><button>Save</button><button name="clear" value="1">Clear</button></form>'
        else: mh='—' if not manual_val else f'${float(manual_val):.4f}'
        fmt=lambda v:'—' if v is None else f'${float(v):.2f}'
        rows+=f'<tr><td>{html.escape(posname)}</td><td>{float(r["quantity"]):,.0f}</td><td>{fmt(r.get("previous_mark"))}</td><td>{fmt(r.get("bid"))}</td><td>{fmt(r.get("ask"))}</td><td>{fmt(r.get("market_mark"))}</td><td>{mh}</td><td>{fmt(r.get("current_mark"))}</td><td>{"—" if r.get("change") is None else f"{r["change"]:+.2f}"}</td><td>{"—" if r.get("estimated_pnl") is None else f"${r["estimated_pnl"]:+,.2f}"}</td><td>{"—" if r.get("contribution_pct") is None else f"{r["contribution_pct"]:+.4f}%"}</td><td>{html.escape(r.get("effective_source") or "")}</td></tr>'
    warn='' if calc["complete"] else '<div class="notice"><strong>PARTIAL:</strong> Yahoo local collector has not supplied a usable current mark for every Opportunity position.</div><br>'
    body=f'''<h1>RPD Opportunity Fund — Yahoo Daily Return</h1><div class="grid"><div class="card metric"><div class="muted">Status</div><div class="v {'ok' if calc['complete'] else 'bad'}">{status}</div></div><div class="card metric"><div class="muted">Estimated Daily Return</div><div class="v">{ret:+.4f}%</div></div><div class="card metric"><div class="muted">Estimated P&amp;L</div><div class="v">${pnl:+,.2f}</div></div><div class="card metric"><div class="muted">Baseline NAV</div><div class="v">${float(b.get('nav_usd') or 0):,.2f}</div><div class="muted">Report {html.escape(str(b.get('report_date') or ''))}</div></div><div class="card metric"><div class="muted">Coverage</div><div class="v">{calc['coverage_valid']}/{calc['coverage_total']}</div><div class="muted">Yahoo local collector marks</div></div></div>{warn}<div class="card"><strong>{html.escape(str(b.get('filename') or 'Opportunity PNL'))}</strong> · Formula: (Current Mark − Previous P&amp;L Report Price) × Quantity × Multiplier.</div><div class="card"><table><thead><tr><th>Position</th><th>Qty</th><th>Previous Mark</th><th>Yahoo Bid</th><th>Yahoo Ask</th><th>Yahoo Market</th><th>Manual Override</th><th>Effective Price</th><th>Change</th><th>Estimated P&amp;L</th><th>Contribution</th><th>Effective Source</th></tr></thead><tbody>{rows}</tbody></table></div>'''
    return page("Opportunity Yahoo Daily Return",body,u,60 if auto_refresh_allowed() else None)

@app.get("/opportunity/yahoo-daily-return")
def opportunity_yahoo_daily_return(request:Request): return _render_opportunity_yahoo(request)

@app.get("/opportunity/bloomberg-daily-return")
def opportunity_bloomberg_daily_return(request:Request):
    u=require_user(request)
    with SessionLocal() as db:
        b=_opportunity_baseline(db);snap,age=_latest_opportunity_snapshot(db)
    if not b:
        extra=' <a href="/admin/opportunity">Upload Opportunity P&amp;L Report</a>' if u["username"]==ADMIN_USER else ''
        return page("Opportunity Bloomberg Daily Return",f'<h1>RPD Opportunity Fund — Bloomberg Daily Return</h1><div class="card"><div class="notice">No Opportunity P&amp;L baseline loaded.{extra}</div></div>',u)
    pos=_opp_baseline_positions_for_calc(b);calc=_calc_bloomberg_daily(pos,float(b.get("nav_usd") or 0),snap);ret=float(calc["estimated_return_pct"]);pnl=float(calc["estimated_pnl"]);rows=""
    for r in calc["rows"]:
        strike="" if r.get("strike") is None else f'{r["strike"]:g}';name=(f'{r["ticker"]} {r["expiry"]} {r["option_type"]}{strike}' if r["instrument_type"]=="OPTION" else r["ticker"]);manual=r.get("bloomberg_manual_price");mv='' if manual is None else f'{float(manual):.4f}'
        if u["username"]==ADMIN_USER:
            mh=f'<form method="post" action="/admin/opportunity/manual-price" style="display:flex;gap:4px"><input type="hidden" name="position_id" value="{r["id"]}"><input type="hidden" name="source" value="BLOOMBERG"><input type="number" step="0.0001" min="0" name="price" value="{mv}" style="width:82px"><button>Save</button><button name="clear" value="1">Clear</button></form>'
        else: mh='—' if not mv else f'${float(mv):.4f}'
        fmt=lambda v:'—' if v is None else f'${float(v):.2f}'
        rows+=f'<tr><td>{html.escape(name)}</td><td>{float(r["quantity"]):,.0f}</td><td>{fmt(r.get("baseline_price"))}</td><td>{fmt(r.get("bid"))}</td><td>{fmt(r.get("ask"))}</td><td>{fmt(r.get("market_mark"))}</td><td>{mh}</td><td>{fmt(r.get("current_mark"))}</td><td>{"—" if r.get("change") is None else f"{r["change"]:+.2f}"}</td><td>{"—" if r.get("estimated_pnl") is None else f"${r["estimated_pnl"]:+,.2f}"}</td><td>{"—" if r.get("contribution_pct") is None else f"{r["contribution_pct"]:+.4f}%"}</td><td>{html.escape(r.get("effective_source") or "")}</td></tr>'
    status="FULL" if calc["complete"] else "PARTIAL";warn='' if calc["complete"] else '<div class="notice"><strong>PARTIAL:</strong> Bloomberg has not supplied a usable current mark for every Opportunity position.</div><br>'
    body=f'''<h1>RPD Opportunity Fund — Bloomberg Daily Return</h1><div class="grid"><div class="card metric"><div class="muted">Status</div><div class="v {'ok' if calc['complete'] else 'bad'}">{status}</div></div><div class="card metric"><div class="muted">Estimated Daily Return</div><div class="v">{ret:+.4f}%</div></div><div class="card metric"><div class="muted">Estimated P&amp;L</div><div class="v">${pnl:+,.2f}</div></div><div class="card metric"><div class="muted">Baseline NAV</div><div class="v">${float(b.get('nav_usd') or 0):,.2f}</div><div class="muted">Report {html.escape(str(b.get('report_date') or ''))}</div></div><div class="card metric"><div class="muted">Coverage</div><div class="v">{calc['coverage_valid']}/{calc['coverage_total']}</div><div class="muted">Bloomberg marks</div></div></div>{warn}<div class="card"><div class="muted">Collector snapshot</div>{html.escape(format_new_york_time((snap or {}).get('timestamp_utc','')))}</div><div class="card"><table><thead><tr><th>Position</th><th>Qty</th><th>Previous Mark</th><th>Bloomberg Bid</th><th>Bloomberg Ask</th><th>Bloomberg Market</th><th>Manual Override</th><th>Effective Price</th><th>Change</th><th>Estimated P&amp;L</th><th>Contribution</th><th>Effective Source</th></tr></thead><tbody>{rows}</tbody></table></div>'''
    return page("Opportunity Bloomberg Daily Return",body,u,60 if auto_refresh_allowed() else None)

@app.get("/admin/opportunity")
def opportunity_admin(request:Request):
    u=require_admin(request)
    with SessionLocal() as db:
        exp=_opportunity_exposure(db);base=_opportunity_baseline(db);limit=get_setting(db,"opportunity_limit_pct","");ver=get_setting(db,"opportunity_config_version","1")
    exp_txt='<p>No Opportunity Exposure report loaded.</p>' if not exp else f'<div class="notice"><strong>Exposure:</strong> {html.escape(str(exp.get("filename") or ""))} · Report {html.escape(str(exp.get("report_date") or ""))} · NAV ${float(exp.get("nav_usd") or 0):,.2f} · Delta positions {len(exp.get("positions") or [])}</div>'
    base_txt='<p>No Opportunity P&amp;L baseline loaded.</p>' if not base else f'<div class="notice"><strong>P&amp;L baseline:</strong> {html.escape(str(base.get("filename") or ""))} · Report {html.escape(str(base.get("report_date") or ""))} · NAV ${float(base.get("nav_usd") or 0):,.2f} · Daily Return positions {len(base.get("positions") or [])}</div>'
    body=admin_tabs()+f'''<h1>Opportunity Setup</h1><div class="card"><h3>Opportunity Delta Monitor</h3><p>Upload the latest Nirvana Exposure by Underlying XLSX. This updates only Opportunity.</p><form method="post" action="/admin/opportunity/exposure-upload" enctype="multipart/form-data"><input type="file" name="file" accept=".xlsx" required><button>Upload Opportunity Exposure</button></form><br>{exp_txt}</div><div class="card"><h3>Opportunity Daily Return Baseline</h3><p>Upload the previous trading day's Opportunity PNL Report PDF. It supplies positions, previous marks and NAV for both Bloomberg and Yahoo Daily Return.</p><form method="post" action="/admin/opportunity/pnl-upload" enctype="multipart/form-data"><input type="file" name="file" accept=".pdf,application/pdf" required><button>Upload Opportunity P&amp;L</button></form><br>{base_txt}</div><div class="card"><h3>Opportunity Settings — Config v{html.escape(str(ver))}</h3><form method="post" action="/admin/opportunity/settings" class="row"><label>Delta Limit % (optional)<input name="limit_pct" value="{html.escape(str(limit or ''))}" placeholder="Leave blank if no limit"></label><button>Publish</button></form></div>'''
    return page("Opportunity Setup",body,u)

@app.post("/admin/opportunity/exposure-upload")
async def opportunity_exposure_upload(request:Request,file:UploadFile=File(...)):
    u=require_admin(request);filename=file.filename or "Opportunity Exposure.xlsx"
    if not filename.lower().endswith(".xlsx"): raise HTTPException(400,"Please upload the Opportunity Exposure by Underlying XLSX.")
    data=await file.read()
    try: parsed=_parse_opportunity_exposure_xlsx(data,filename)
    except Exception as e: raise HTTPException(400,f"Could not parse Opportunity Exposure: {e}")
    with SessionLocal() as db:
        _set_json_setting(db,"opportunity_exposure_config",parsed);_bump_opportunity_config(db,u["username"],f'Exposure {filename}; report={parsed.get("report_date")}; positions={len(parsed.get("positions") or [])}');db.commit()
    return RedirectResponse("/admin/opportunity",303)

@app.post("/admin/opportunity/pnl-upload")
async def opportunity_pnl_upload(request:Request,file:UploadFile=File(...)):
    u=require_admin(request);filename=file.filename or "Opportunity PNL.pdf"
    if not filename.lower().endswith(".pdf"): raise HTTPException(400,"Please upload the Opportunity PNL Report PDF.")
    data=await file.read()
    try: parsed=parse_nirvana_pnl_pdf(data)
    except Exception as e: raise HTTPException(400,f"Could not parse Opportunity PNL Report: {e}")
    positions=[]
    for i,x in enumerate(parsed["positions"],1):
        positions.append({**x,"id":i,"yahoo_manual_price":None,"bloomberg_manual_price":None})
    stored={"filename":filename,"report_date":parsed["report_date"],"run_date":parsed["run_date"],"nav_usd":parsed["nav_usd"],"positions":positions,"version":int(datetime.now(timezone.utc).timestamp())}
    with SessionLocal() as db:
        _set_json_setting(db,"opportunity_daily_baseline",stored);_bump_opportunity_config(db,u["username"],f'PNL baseline {filename}; report={parsed.get("report_date")}; positions={len(positions)}');db.commit()
    return RedirectResponse("/admin/opportunity",303)

@app.post("/admin/opportunity/settings")
async def opportunity_settings(request:Request,limit_pct:str=Form("")):
    u=require_admin(request);raw=str(limit_pct or "").strip()
    if raw:
        try:
            val=float(raw)
            if val<=0: raise ValueError
        except Exception: raise HTTPException(400,"Opportunity limit must be positive or blank.")
    with SessionLocal() as db:
        set_setting(db,"opportunity_limit_pct",raw);_bump_opportunity_config(db,u["username"],f"Limit={raw or 'N/A'}");db.commit()
    return RedirectResponse("/admin/opportunity",303)

@app.post("/admin/opportunity/manual-price")
async def opportunity_manual_price(request:Request):
    u=require_admin(request);form=await request.form();source=str(form.get("source") or "").upper();clear=str(form.get("clear") or "")=="1"
    try: pid=int(form.get("position_id"))
    except Exception: raise HTTPException(400,"Invalid position id.")
    val=None
    if not clear:
        try:
            val=float(str(form.get("price") or "").strip())
            if val<0: raise ValueError
        except Exception: raise HTTPException(400,"Manual price must be zero or greater.")
    with SessionLocal() as db:
        b=_opportunity_baseline(db)
        if not b: raise HTTPException(404,"Opportunity baseline not found.")
        found=False
        for p in b.get("positions",[]):
            if int(p.get("id") or 0)==pid:
                if source=="YAHOO": p["yahoo_manual_price"]=val;redirect="/opportunity/yahoo-daily-return"
                elif source=="BLOOMBERG": p["bloomberg_manual_price"]=val;redirect="/opportunity/bloomberg-daily-return"
                else: raise HTTPException(400,"Unknown source.")
                found=True;break
        if not found: raise HTTPException(404,"Position not found.")
        _set_json_setting(db,"opportunity_daily_baseline",b);audit(db,u["username"],"OPPORTUNITY_MANUAL_PRICE",f"{source} position_id={pid}; price={val}");db.commit()
    return RedirectResponse(redirect,303)

@app.get("/admin")
def admin_home(request:Request):
    u=require_admin(request)
    with SessionLocal() as db:
        nav=get_setting(db,"nav_usd","");limit=get_setting(db,"limit_pct","15");ver=get_setting(db,"config_version","1");nav_display=f"{float(nav):,.2f}"
        pos=db.query(Position).order_by(Position.active.desc(),Position.ticker,Position.expiry,Position.strike).all()
        today=date.today().isoformat();rows=""
        for p in pos:
            status="EXPIRED" if p.expiry<today else ("ACTIVE" if p.active else "ARCHIVED")
            action="" if not p.active else f"""<form method="post" action="/admin/positions/{p.id}/quantity" style="display:inline"><input name="quantity" value="{p.quantity:g}" size="7"><button>Qty</button></form>
<form method="post" action="/admin/positions/{p.id}/archive" style="display:inline"><button class="danger">Archive</button></form>"""
            rows+=f'<tr><td>{p.ticker}</td><td>{p.expiry}</td><td>{p.option_type}</td><td>{p.strike:g}</td><td>{p.quantity:,.0f}</td><td>{status}</td><td>{html.escape(p.bloomberg_security)}</td><td>{action}</td></tr>'
        body=admin_tabs()+f"""<div class="card"><h2>Admin — Fuat only</h2><div class="notice">Only Fuat can change portfolio, NAV, limit and users. Viewers are read-only.</div></div>
<div class="card"><h3>Fund Settings — Config v{ver}</h3><form method="post" action="/admin/settings" class="row"><label>NAV USD<input name="nav_usd" value="{nav_display}"></label><label>Delta Limit %<input name="limit_pct" value="{limit}"></label><button>Publish</button></form></div>
<div class="card"><h3>Add Position</h3><form method="post" action="/admin/positions/add" class="row"><label>Ticker<input name="ticker" required></label><label>Expiry YYYY-MM-DD<input name="expiry" required></label><label>Type<select name="option_type"><option>P</option><option>C</option></select></label><label>Strike<input name="strike" required></label><label>Qty<input name="quantity" required></label><label>Multiplier<input name="multiplier" value="100"></label><button>Add & Publish</button></form></div>
<div class="card"><h3>Portfolio</h3><table><thead><tr><th>Ticker</th><th>Expiry</th><th>Type</th><th>Strike</th><th>Qty</th><th>Status</th><th>Bloomberg Security</th><th>Action</th></tr></thead><tbody>{rows}</tbody></table></div>"""
        return page("Admin",body,u)

@app.post("/admin/settings")
def settings_update(request:Request,nav_usd:str=Form(...),limit_pct:str=Form(...)):
    u=require_admin(request);nav=float(nav_usd.replace(",",""));limit=float(limit_pct)
    if nav<=0 or not 0<limit<100:raise HTTPException(400,"Invalid NAV/limit")
    with SessionLocal() as db:
        set_setting(db,"nav_usd",nav);set_setting(db,"limit_pct",limit);config_bump(db,u["username"],f"NAV={nav}; limit={limit}%");db.commit()
    return RedirectResponse("/admin",303)

@app.post("/admin/positions/add")
def pos_add(request:Request,ticker:str=Form(...),expiry:str=Form(...),option_type:str=Form(...),strike:str=Form(...),quantity:str=Form(...),multiplier:str=Form("100")):
    u=require_admin(request);t=ticker.strip().upper();typ=option_type.strip().upper()[0];k=float(strike);q=float(quantity);m=float(multiplier);datetime.strptime(expiry,"%Y-%m-%d")
    if typ not in ("P","C"):raise HTTPException(400)
    with SessionLocal() as db:
        db.add(Position(ticker=t,expiry=expiry,option_type=typ,strike=k,quantity=q,multiplier=m,bloomberg_security=option_security(t,expiry,typ,k),underlying_security=f"{t} US Equity",active=True,source="manual"))
        audit(db,u["username"],"POSITION_ADDED",f"{t} {expiry} {typ}{k} qty={q}");config_bump(db,u["username"],"Position added");db.commit()
    return RedirectResponse("/admin",303)

@app.post("/admin/positions/{pid}/quantity")
def pos_qty(request:Request,pid:int,quantity:str=Form(...)):
    u=require_admin(request);q=float(quantity)
    with SessionLocal() as db:
        p=db.get(Position,pid)
        if not p:raise HTTPException(404)
        old=p.quantity;p.quantity=q;p.updated_at=utcnow();audit(db,u["username"],"QUANTITY_CHANGED",f"{p.ticker} {p.expiry} {p.option_type}{p.strike}: {old}->{q}");config_bump(db,u["username"],"Quantity changed");db.commit()
    return RedirectResponse("/admin",303)

@app.post("/admin/positions/{pid}/archive")
def pos_archive(request:Request,pid:int):
    u=require_admin(request)
    with SessionLocal() as db:
        p=db.get(Position,pid)
        if not p:raise HTTPException(404)
        if p.active:
            p.active=False;p.updated_at=utcnow();audit(db,u["username"],"POSITION_ARCHIVED",f"{p.ticker} {p.expiry} {p.option_type}{p.strike}");config_bump(db,u["username"],"Position archived");db.commit()
    return RedirectResponse("/admin",303)

@app.get("/admin/users")
def users_page(request:Request):
    u=require_admin(request)
    with SessionLocal() as db:
        users=db.query(User).order_by(User.username).all();rows=""
        for x in users:
            if x.username==ADMIN_USER:
                rows+=f"<tr><td>{x.username}</td><td>ADMIN (fixed)</td><td>Active</td><td>Protected</td></tr>"
            else:
                rows+=f"""<tr><td>{html.escape(x.username)}</td><td>Viewer</td><td>{'Active' if x.active else 'Disabled'}</td><td>
<form method="post" action="/admin/users/{x.id}/toggle" style="display:inline"><button class="secondary">{'Disable' if x.active else 'Enable'}</button></form>
<form method="post" action="/admin/users/{x.id}/reset" style="display:inline"><input type="password" name="password" placeholder="New password" required><button>Reset</button></form>
<form method="post" action="/admin/users/{x.id}/delete" style="display:inline"><button class="danger">Delete</button></form></td></tr>"""
        body=admin_tabs()+f"""<div class="card"><h2>Viewer Users</h2><div class="notice">Fuat is the only admin. New accounts are viewer-only.</div><br>
<form method="post" action="/admin/users/add" class="row"><label>Username<input name="username" required></label><label>Temporary password<input type="password" name="password" required></label><button>Add Viewer</button></form></div>
<div class="card"><table><thead><tr><th>User</th><th>Role</th><th>Status</th><th>Actions</th></tr></thead><tbody>{rows}</tbody></table></div>"""
        return page("Users",body,u)

@app.post("/admin/users/add")
def user_add(request:Request,username:str=Form(...),password:str=Form(...)):
    u=require_admin(request);name=username.strip().lower()
    if name==ADMIN_USER or not re.match(r"^[a-z0-9._-]{3,40}$",name):raise HTTPException(400,"Invalid/reserved username")
    if len(password)<8:raise HTTPException(400,"Password must be at least 8 characters")
    with SessionLocal() as db:
        if db.query(User).filter_by(username=name).first():raise HTTPException(400,"User exists")
        db.add(User(username=name,password_hash=hash_password(password),role="viewer",active=True));audit(db,u["username"],"VIEWER_USER_ADDED",name);db.commit()
    return RedirectResponse("/admin/users",303)

def viewer(db,uid):
    x=db.get(User,uid)
    if not x or x.username==ADMIN_USER:raise HTTPException(400,"Protected user")
    return x

@app.post("/admin/users/{uid}/toggle")
def user_toggle(request:Request,uid:int):
    u=require_admin(request)
    with SessionLocal() as db:
        x=viewer(db,uid);x.active=not x.active;x.updated_at=utcnow();audit(db,u["username"],"VIEWER_STATUS_CHANGED",f"{x.username}: active={x.active}");db.commit()
    return RedirectResponse("/admin/users",303)

@app.post("/admin/users/{uid}/reset")
def user_reset(request:Request,uid:int,password:str=Form(...)):
    u=require_admin(request)
    if len(password)<8:raise HTTPException(400)
    with SessionLocal() as db:
        x=viewer(db,uid);x.password_hash=hash_password(password);x.updated_at=utcnow();audit(db,u["username"],"VIEWER_PASSWORD_RESET",x.username);db.commit()
    return RedirectResponse("/admin/users",303)

@app.post("/admin/users/{uid}/delete")
def user_delete(request:Request,uid:int):
    u=require_admin(request)
    with SessionLocal() as db:
        x=viewer(db,uid);name=x.username;db.delete(x);audit(db,u["username"],"VIEWER_USER_DELETED",name);db.commit()
    return RedirectResponse("/admin/users",303)

@app.get("/admin/audit")
def audit_page(request:Request):
    u=require_admin(request)
    with SessionLocal() as db:
        items=db.query(Audit).order_by(Audit.id.desc()).limit(300).all()
        rows="".join(f"<tr><td>{x.ts}</td><td>{html.escape(x.actor)}</td><td>{html.escape(x.action)}</td><td>{html.escape(x.detail)}</td></tr>" for x in items)
    return page("Audit",admin_tabs()+f'<div class="card"><h2>Audit Log</h2><table><thead><tr><th>Time</th><th>Actor</th><th>Action</th><th>Detail</th></tr></thead><tbody>{rows}</tbody></table></div>',u)

@app.get("/admin/import")
def import_get(request:Request):
    u=require_admin(request)
    return page("Nirvana Import",admin_tabs()+"""<div class="card"><h2>Nirvana Position Import</h2>
<p>Upload the Nirvana <strong>Exposure Summary by Underlying</strong> XLSX directly. The portal recognizes this report automatically, extracts the real option rows, NAV and quantities, then shows a preview before anything changes.</p>
<p class="muted">Generic CSV/XLSX files are still supported with manual column mapping.</p>
<form method="post" enctype="multipart/form-data"><input type="file" name="file" accept=".xlsx,.csv" required><button>Upload & Preview</button></form></div>""",u)

def _compare_import(parsed,current):
    cur={keypos(x):x for x in current};new={keypos(x):x for x in parsed}
    added=[x for k,x in new.items() if k not in cur]
    changed=[(cur[k],x) for k,x in new.items() if k in cur and (float(cur[k]["quantity"])!=float(x["quantity"]) or float(cur[k]["multiplier"])!=float(x["multiplier"]))]
    missing=[x for k,x in cur.items() if k not in new]
    return added,changed,missing

def _preview_html(jid,parsed,errors,current,mapping,report_nav=None,current_nav=None,report_date=None,auto=False):
    added,changed,missing=_compare_import(parsed,current)
    def fmt(x):return f'{x["ticker"]} {x["expiry"]} {x["option_type"]}{x["strike"]:g} qty={x["quantity"]:,.0f}'
    detail="<h3>Added</h3><ul>"+"".join(f"<li>{html.escape(fmt(x))}</li>" for x in added[:50])+"</ul>"
    detail+="<h3>Changed</h3><ul>"+"".join(f'<li>{html.escape(fmt(a))} → qty={b["quantity"]:,.0f}</li>' for a,b in changed[:50])+"</ul>"
    detail+="<h3>Missing from uploaded file</h3><ul>"+"".join(f"<li>{html.escape(fmt(x))}</li>" for x in missing[:50])+"</ul>"
    if errors:detail+="<h3>Parse errors</h3><pre>"+html.escape("\n".join(errors[:30]))+"</pre>"
    navline=""
    if report_nav:
        navline=f'<div class="notice"><strong>NAV from Nirvana:</strong> ${report_nav:,.2f}'
        if current_nav is not None:navline+=f' &nbsp; | &nbsp; Current portal NAV: ${current_nav:,.2f}'
        navline+=' &nbsp; — NAV will be updated when you Confirm & Publish.</div><br>'
    source='Nirvana Exposure Summary auto-detected' if auto else 'Generic mapped import'
    if report_date:source+=f' · Report Date {html.escape(str(report_date))}'
    mj=quote(json.dumps(mapping))
    return admin_tabs()+f"""<div class="card"><h2>Import Preview</h2><p><strong>{source}</strong></p>{navline}
<p><strong>Parsed option positions:</strong> {len(parsed)} · <strong>Added:</strong> {len(added)} · <strong>Changed:</strong> {len(changed)} · <strong>Missing:</strong> {len(missing)} · <strong>Parse errors:</strong> {len(errors)}</p>{detail}
<form method="post" action="/admin/import/{jid}/apply"><input type="hidden" name="mapping_json" value="{html.escape(mj)}"><label>Mode<select name="mode"><option value="replace">Replace active portfolio — missing positions archived</option><option value="merge">Merge/update only — missing positions untouched</option></select></label><br><br><button>Confirm & Publish</button> <a href="/admin/import">Cancel</a></form></div>"""

@app.post("/admin/import")
async def import_upload(request:Request,file:UploadFile=File(...)):
    u=require_admin(request)
    data=await file.read()
    headers,rows,meta=parse_upload(data,file.filename or "upload")
    if not rows:raise HTTPException(400,"No rows found")
    jid=uuid.uuid4().hex
    envelope={"headers":headers,"meta":meta}
    with SessionLocal() as db:
        db.add(ImportJob(id=jid,filename=file.filename or "upload",headers_json=json.dumps(envelope),rows_json=json.dumps(rows,default=str)))
        audit(db,u["username"],"NIRVANA_FILE_UPLOADED",f"{file.filename}; rows={len(rows)}; format={meta.get('format')}")
        db.commit()

    if meta.get("format")=="nirvana_exposure_by_underlying":
        parsed,errors,_=parse_nirvana_exposure_rows(rows)
        if not parsed:raise HTTPException(400,"Nirvana report detected but no option rows could be parsed.")
        with SessionLocal() as db:
            current=[{"ticker":p.ticker,"expiry":p.expiry,"option_type":p.option_type,"strike":p.strike,"quantity":p.quantity,"multiplier":p.multiplier} for p in db.query(Position).filter_by(active=True).all()]
            current_nav=float(get_setting(db,"nav_usd","0") or 0)
        mapping={"format":"nirvana_exposure_by_underlying"}
        body=_preview_html(jid,parsed,errors,current,mapping,meta.get("nav_usd"),current_nav,meta.get("report_date"),True)
        return page("Nirvana Import Preview",body,u)

    mapping=detect_map(headers)
    def opts(sel=""):
        return "".join(f'<option value="{html.escape(h)}" {"selected" if h==sel else ""}>{html.escape(h)}</option>' for h in headers)
    body=admin_tabs()+f"""<div class="card"><h2>Map Columns</h2><p>{html.escape(file.filename or "")}: {len(rows)} rows</p><form method="post" action="/admin/import/{jid}/preview"><div class="row">
<label>Ticker<select name="ticker_col">{opts(mapping.get("ticker",""))}</select></label><label>Expiry<select name="expiry_col">{opts(mapping.get("expiry",""))}</select></label>
<label>Put/Call<select name="type_col">{opts(mapping.get("option_type",""))}</select></label><label>Strike<select name="strike_col">{opts(mapping.get("strike",""))}</select></label>
<label>Quantity<select name="qty_col">{opts(mapping.get("quantity",""))}</select></label><label>Multiplier<select name="mult_col"><option value="">Default 100</option>{opts(mapping.get("multiplier",""))}</select></label></div><br><button>Build Preview</button></form></div>"""
    return page("Map Import",body,u)

@app.post("/admin/import/{jid}/preview")
def import_preview(request:Request,jid:str,ticker_col:str=Form(...),expiry_col:str=Form(...),type_col:str=Form(...),strike_col:str=Form(...),qty_col:str=Form(...),mult_col:str=Form("")):
    u=require_admin(request);mapping={"ticker":ticker_col,"expiry":expiry_col,"type":type_col,"strike":strike_col,"qty":qty_col,"mult":mult_col}
    with SessionLocal() as db:
        job=db.get(ImportJob,jid)
        if not job:raise HTTPException(404)
        parsed,errors=normalized_import(json.loads(job.rows_json),mapping)
        current=[{"ticker":p.ticker,"expiry":p.expiry,"option_type":p.option_type,"strike":p.strike,"quantity":p.quantity,"multiplier":p.multiplier} for p in db.query(Position).filter_by(active=True).all()]
    body=_preview_html(jid,parsed,errors,current,mapping)
    return page("Import Preview",body,u)

@app.post("/admin/import/{jid}/apply")
def import_apply(request:Request,jid:str,mapping_json:str=Form(...),mode:str=Form(...)):
    u=require_admin(request);mapping=json.loads(unquote(mapping_json))
    with SessionLocal() as db:
        job=db.get(ImportJob,jid)
        if not job:raise HTTPException(404)
        rows=json.loads(job.rows_json)
        try:envelope=json.loads(job.headers_json)
        except Exception:envelope={"headers":[],"meta":{}}
        meta=envelope.get("meta",{}) if isinstance(envelope,dict) else {}
        if mapping.get("format")=="nirvana_exposure_by_underlying":
            parsed,errors,_=parse_nirvana_exposure_rows(rows)
        else:
            parsed,errors=normalized_import(rows,mapping)
        if not parsed:raise HTTPException(400,"No valid option positions parsed; nothing was changed.")
        if errors and len(errors)>max(3,len(parsed)//10):raise HTTPException(400,"Too many parse errors; nothing was changed.")
        active=db.query(Position).filter_by(active=True).all();cur={(p.ticker,p.expiry,p.option_type,round(p.strike,6)):p for p in active};incoming={keypos(x):x for x in parsed}
        added=changed=archived=0
        for k,x in incoming.items():
            p=cur.get(k)
            if p:
                if p.quantity!=x["quantity"] or p.multiplier!=x["multiplier"]:
                    p.quantity=x["quantity"];p.multiplier=x["multiplier"];p.updated_at=utcnow();changed+=1
            else:
                db.add(Position(ticker=x["ticker"],expiry=x["expiry"],option_type=x["option_type"],strike=x["strike"],quantity=x["quantity"],multiplier=x["multiplier"],bloomberg_security=option_security(x["ticker"],x["expiry"],x["option_type"],x["strike"]),underlying_security=f'{x["ticker"]} US Equity',active=True,source="nirvana"));added+=1
        if mode=="replace":
            for k,p in cur.items():
                if k not in incoming:p.active=False;p.updated_at=utcnow();archived+=1
        nav_note=""
        report_nav=meta.get("nav_usd") if mapping.get("format")=="nirvana_exposure_by_underlying" else None
        if report_nav:
            old_nav=float(get_setting(db,"nav_usd","0") or 0)
            set_setting(db,"nav_usd",float(report_nav))
            nav_note=f"; NAV {old_nav}->{float(report_nav)}"
        audit(db,u["username"],"NIRVANA_IMPORT_APPLIED",f"{job.filename}; mode={mode}; added={added}; changed={changed}; archived={archived}; parse_errors={len(errors)}{nav_note}")
        config_bump(db,u["username"],f"Nirvana import {job.filename}")
        db.delete(job);db.commit()
    return RedirectResponse("/admin",303)

def _active_daily_baseline(db):
    return db.query(DailyReturnBaseline).filter_by(active=True).order_by(DailyReturnBaseline.id.desc()).first()


def _latest_yahoo_collector_snapshot(db):
    """Return Fortress Yahoo snapshot; prefer the multi-fund local collector."""
    v2=_latest_yahoo_v2_fund_snapshot(db,"fortress")
    if v2:
        return v2
    raw=get_setting(db,"yahoo_collector_snapshot","")
    if not raw:
        return None
    try:
        data=json.loads(raw)
        return data if isinstance(data,dict) else None
    except Exception:
        return None

def _yahoo_quote_data_from_snapshot(snapshot):
    """Convert local collector rows into calculate_mark_to_market_return quote_data."""
    out={}
    if not snapshot:
        return out
    for r in snapshot.get("positions",[]) or []:
        symbol=str(r.get("symbol") or "").upper().strip()
        if not symbol:
            continue
        out[symbol]={
            "mark":r.get("mark"),
            "bid":r.get("bid"),
            "ask":r.get("ask"),
            "last":r.get("last"),
            "source":r.get("source") or "",
            "note":r.get("note") or "",
        }
    return out

@app.get("/daily-return")
def daily_return_page(request:Request):
    u=require_user(request)
    with SessionLocal() as db:
        b=_active_daily_baseline(db)
        if not b:
            extra=' <a href="/admin/daily-return">Upload P&amp;L Report</a>' if u["username"]==ADMIN_USER else ''
            return page("Estimated Daily Return",f'<h1>Estimated Daily Return</h1><div class="card"><div class="notice">No P&amp;L baseline has been uploaded yet.{extra}</div></div>',u)

        dbpos=db.query(DailyReturnPosition).filter_by(baseline_id=b.id).order_by(DailyReturnPosition.id).all()
        positions=[{
            "id":p.id,
            "security_name":p.security_name,
            "instrument_type":p.instrument_type,
            "ticker":p.ticker,
            "expiry":p.expiry,
            "option_type":p.option_type,
            "strike":p.strike,
            "quantity":p.quantity,
            "multiplier":p.multiplier,
            "baseline_price":p.baseline_price,
            "baseline_market_value":p.baseline_market_value,
            "yahoo_manual_price":p.yahoo_manual_price,
            "bloomberg_manual_price":p.bloomberg_manual_price,
        } for p in dbpos]
        baseline={
            "report_date":b.report_date,
            "run_date":b.run_date,
            "filename":b.filename,
            "nav_usd":b.nav_usd,
        }

    with SessionLocal() as db:
        yahoo_snapshot=_latest_yahoo_collector_snapshot(db)
    yahoo_quote_data=_yahoo_quote_data_from_snapshot(yahoo_snapshot)
    calc=calculate_mark_to_market_return(
        positions,
        float(baseline["nav_usd"]),
        quote_data=yahoo_quote_data,
    )
    calc["yahoo_error"]=""
    status="FULL" if calc["complete"] else "PARTIAL"
    ret=calc["estimated_return_pct"] or 0.0
    pnl=calc["estimated_pnl"] or 0.0

    rows=""
    for r in calc["rows"]:
        strike="" if r.get("strike") is None else f'{r["strike"]:g}'
        posname=(f'{r["ticker"]} {r["expiry"]} {r["option_type"]}{strike}' if r["instrument_type"]=="OPTION" else r["ticker"])
        previous='—' if r.get("previous_mark") is None else f'${r["previous_mark"]:.2f}'
        bid='—' if r.get("bid") is None else f'${r["bid"]:.2f}'
        ask='—' if r.get("ask") is None else f'${r["ask"]:.2f}'
        market='—' if r.get("market_mark") is None else f'${r["market_mark"]:.2f}'
        current='—' if r.get("current_mark") is None else f'${r["current_mark"]:.2f}'
        manual_val='' if r.get("manual_price") is None else f'{float(r["manual_price"]):.4f}'
        change='—' if r.get("change") is None else f'{r["change"]:+.2f}'
        epnl='—' if r.get("estimated_pnl") is None else f'${r["estimated_pnl"]:+,.2f}'
        contrib='—' if r.get("contribution_pct") is None else f'{r["contribution_pct"]:+.4f}%'
        source=(r.get("effective_source") or "")
        if u["username"]==ADMIN_USER:
            manual_html=(f'<form method="post" action="/admin/daily-return/manual-price" style="display:flex;gap:4px;align-items:center">'
                         f'<input type="hidden" name="position_id" value="{r["id"]}">'
                         f'<input type="hidden" name="source" value="YAHOO">'
                         f'<input type="number" step="0.0001" min="0" name="price" value="{manual_val}" style="width:82px">'
                         f'<button>Save</button><button name="clear" value="1">Clear</button></form>')
        else:
            manual_html='—' if not manual_val else f'${float(manual_val):.4f}'
        rows+=f'<tr><td>{html.escape(posname)}</td><td>{r["quantity"]:,.0f}</td><td>{previous}</td><td>{bid}</td><td>{ask}</td><td>{market}</td><td>{manual_html}</td><td>{current}</td><td>{change}</td><td>{epnl}</td><td>{contrib}</td><td>{html.escape(source)}</td></tr>'

    warning=""
    if not calc["complete"]:
        warning='<div class="notice"><strong>PARTIAL:</strong> Yahoo local collector has not supplied a usable current mark for every position. The return shown includes only covered positions.</div><br>'

    if yahoo_snapshot:
        yahoo_collector_name=str(yahoo_snapshot.get("collector_name") or "Yahoo-PC")
        yahoo_collected_at=str(yahoo_snapshot.get("collected_at_utc") or "")
        yahoo_collector_error=str(yahoo_snapshot.get("error") or "")
        try:
            ts=datetime.fromisoformat(yahoo_collected_at.replace("Z","+00:00"))
            age_seconds=max(0,(datetime.now(timezone.utc)-ts).total_seconds())
        except Exception:
            age_seconds=None
        if age_seconds is not None and age_seconds > 1800:
            warning+=f'<div class="notice"><strong>Yahoo Collector:</strong> Snapshot is stale ({age_seconds/60:.0f} minutes old) · {html.escape(yahoo_collector_name)}</div><br>'
        if yahoo_collector_error:
            warning+=f'<div class="notice"><strong>Yahoo Collector:</strong> {html.escape(yahoo_collector_error)}</div><br>'
    else:
        warning+='<div class="notice"><strong>Yahoo Collector:</strong> No local collector snapshot has been received yet.</div><br>'

    quality_class='ok' if calc['complete'] else 'bad'
    body=f'''<h1>RPD Fortress Fund — Yahoo Daily Return</h1>
<div class="grid">
<div class="card metric"><div class="muted">Status</div><div class="v {quality_class}">{status}</div><div class="muted">Mark-to-market estimate</div></div>
<div class="card metric"><div class="muted">Estimated Daily Return</div><div class="v">{ret:+.4f}%</div></div>
<div class="card metric"><div class="muted">Estimated P&amp;L</div><div class="v">${pnl:+,.2f}</div></div>
<div class="card metric"><div class="muted">Baseline NAV</div><div class="v">${baseline['nav_usd']:,.2f}</div><div class="muted">Report {html.escape(baseline['report_date'])}</div></div>
<div class="card metric"><div class="muted">Coverage</div><div class="v">{calc['coverage_valid']}/{calc['coverage_total']}</div><div class="muted">Yahoo local collector marks</div></div>
</div>
{warning}
<div class="card">
<div class="muted">Daily Return baseline — completely separate from Bloomberg</div>
<strong>{html.escape(baseline['filename'])}</strong> · Report Date {html.escape(baseline['report_date'])} · NAV ${baseline['nav_usd']:,.2f}
<br><div class="muted" style="margin-top:6px">Calculated {html.escape(ny_time_label(calc['calculated_at_utc']))}. Current Mark is supplied by the local Yahoo Collector: Bid/Ask midpoint when both are available; if Bid is zero/unavailable and Ask is available, Ask/2; otherwise Yahoo Last. Formula: (Current Mark − Previous P&amp;L Report Price) × Quantity × Multiplier. Daily Return = Total Estimated P&amp;L ÷ Baseline NAV.</div>
</div>
<div class="card"><table><thead><tr><th>Position</th><th>Qty</th><th>Previous Mark</th><th>Yahoo Bid</th><th>Yahoo Ask</th><th>Yahoo Market</th><th>Manual Override</th><th>Effective Price</th><th>Change</th><th>Estimated P&amp;L</th><th>Contribution</th><th>Effective Source</th></tr></thead><tbody>{rows}</tbody></table></div>'''
    return page("Yahoo Daily Return",body,u,60 if auto_refresh_allowed() else None)


@app.get("/bloomberg-daily-return")
def bloomberg_daily_return_page(request:Request):
    u=require_user(request)

    with SessionLocal() as db:
        b=_active_daily_baseline(db)
        if not b:
            extra=' <a href="/admin/daily-return">Upload P&amp;L Report</a>' if u["username"]==ADMIN_USER else ''
            return page(
                "Bloomberg Daily Return",
                f'<h1>RPD Fortress Fund — Bloomberg Daily Return</h1>'
                f'<div class="card"><div class="notice">No P&amp;L baseline has been uploaded yet.{extra}</div></div>',
                u,60
            )

        dbpos=db.query(DailyReturnPosition).filter_by(baseline_id=b.id).order_by(DailyReturnPosition.id).all()
        baseline={
            "report_date":b.report_date,
            "run_date":b.run_date,
            "filename":b.filename,
            "nav_usd":float(b.nav_usd),
        }
        snap,age=latest_snapshot(db)

    positions=[{
        "security_name":p.security_name,
        "instrument_type":p.instrument_type,
        "ticker":p.ticker,
        "expiry":p.expiry,
        "option_type":p.option_type,
        "strike":p.strike,
        "quantity":float(p.quantity),
        "multiplier":float(p.multiplier or 100),
        "baseline_price":float(p.baseline_price) if p.baseline_price is not None else None,
        "position_id":p.id,
        "bloomberg_manual_price":float(p.bloomberg_manual_price) if p.bloomberg_manual_price is not None else None,
    } for p in dbpos]

    bbg_marks={}
    snapshot_ts=""
    collector_id=""
    collector_mode=""
    if snap:
        snapshot_ts=snap.get("timestamp_utc","")
        collector_id=snap.get("collector_id","")
        collector_mode=snap.get("mode","")
        for r in snap.get("positions",[]) or []:
            try:
                key=(
                    str(r.get("ticker","")).upper(),
                    str(r.get("expiry","")),
                    str(r.get("option_type","")).upper(),
                    round(float(r.get("strike")),6),
                )
                mark=r.get("option_mark")
                bid=r.get("option_bid")
                ask=r.get("option_ask")
                last=r.get("option_last")
                bbg_marks[key]={
                    "mark":float(mark) if mark is not None else None,
                    "source":r.get("option_mark_source") or "",
                    "bid":float(bid) if bid is not None else None,
                    "ask":float(ask) if ask is not None else None,
                    "last":float(last) if last is not None else None,
                }
            except Exception:
                continue

    rows=""
    total_pnl=0.0
    valid=0
    for p in positions:
        strike=float(p.get("strike") or 0.0)
        key=(p["ticker"].upper(),p["expiry"],p["option_type"].upper(),round(strike,6))
        q=bbg_marks.get(key,{})
        market_mark=q.get("mark")
        market_source=q.get("source") or ""
        market_bid=q.get("bid")
        market_ask=q.get("ask")
        manual=p.get("bloomberg_manual_price")

        if market_source=="PX_BID/ASK_MID" and market_mark is not None:
            current=market_mark
            effective_source="BLOOMBERG MID"
        elif market_source=="PX_ASK_HALF" and market_mark is not None:
            current=market_mark
            effective_source="BLOOMBERG ASK/2"
        elif manual is not None:
            current=float(manual)
            effective_source="MANUAL OVERRIDE"
        else:
            current=market_mark
            effective_source="BLOOMBERG LAST" if current is not None else "UNAVAILABLE"

        previous=p.get("baseline_price")
        change=pnl=contrib=None
        if current is not None and previous is not None:
            change=float(current)-float(previous)
            pnl=change*float(p["quantity"])*float(p["multiplier"])
            contrib=pnl/baseline["nav_usd"]*100.0 if baseline["nav_usd"] else None
            total_pnl+=pnl
            valid+=1

        posname=(f'{p["ticker"]} {p["expiry"]} {p["option_type"]}{strike:g}' if p["instrument_type"]=="OPTION" else p["ticker"])
        prev_txt='—' if previous is None else f'${previous:.2f}'
        bid_txt='—' if market_bid is None else f'${market_bid:.2f}'
        ask_txt='—' if market_ask is None else f'${market_ask:.2f}'
        market_txt='—' if market_mark is None else f'${market_mark:.2f}'
        cur_txt='—' if current is None else f'${current:.2f}'
        manual_val='' if manual is None else f'{float(manual):.4f}'
        chg_txt='—' if change is None else f'{change:+.2f}'
        pnl_txt='—' if pnl is None else f'${pnl:+,.2f}'
        con_txt='—' if contrib is None else f'{contrib:+.4f}%'
        if u["username"]==ADMIN_USER:
            manual_html=(f'<form method="post" action="/admin/daily-return/manual-price" style="display:flex;gap:4px;align-items:center">'
                         f'<input type="hidden" name="position_id" value="{p["position_id"]}">'
                         f'<input type="hidden" name="source" value="BLOOMBERG">'
                         f'<input type="number" step="0.0001" min="0" name="price" value="{manual_val}" style="width:82px">'
                         f'<button>Save</button><button name="clear" value="1">Clear</button></form>')
        else:
            manual_html='—' if not manual_val else f'${float(manual_val):.4f}'
        rows+=(
            f'<tr><td>{html.escape(posname)}</td><td>{p["quantity"]:,.0f}</td>'
            f'<td>{prev_txt}</td><td>{bid_txt}</td><td>{ask_txt}</td><td>{market_txt}</td>'
            f'<td>{manual_html}</td><td>{cur_txt}</td><td>{chg_txt}</td>'
            f'<td>{pnl_txt}</td><td>{con_txt}</td><td>{html.escape(effective_source)}</td></tr>'
        )

    total=len(positions)
    ret=(total_pnl/baseline["nav_usd"]*100.0) if baseline["nav_usd"] else 0.0
    complete=(total>0 and valid==total)

    if not snap:
        status="WAITING"; status_class="bad"; feed_note="No Bloomberg Collector snapshot yet."
    elif valid==0:
        status="WAITING"; status_class="bad"; feed_note="The current snapshot does not yet contain Bloomberg option marks. The updated Collector must publish a new full snapshot."
    elif complete:
        status="FULL"; status_class="ok"; feed_note=f'Bloomberg snapshot age {age:.1f}s · {html.escape(collector_mode or "LIVE")}'
    else:
        status="PARTIAL"; status_class="bad"; feed_note=f'Bloomberg marks {valid}/{total} · snapshot age {age:.1f}s'

    warning=""
    if not complete:
        warning='<div class="notice"><strong>'+status+':</strong> '+feed_note+' The displayed P&amp;L includes covered positions only.</div><br>'

    body=f'''<h1>RPD Fortress Fund — Bloomberg Daily Return</h1>
<div class="grid">
<div class="card metric"><div class="muted">Status</div><div class="v {status_class}">{status}</div><div class="muted">Bloomberg marks</div></div>
<div class="card metric"><div class="muted">Estimated Daily Return</div><div class="v">{ret:+.4f}%</div></div>
<div class="card metric"><div class="muted">Estimated P&amp;L</div><div class="v">${total_pnl:+,.2f}</div></div>
<div class="card metric"><div class="muted">Baseline NAV</div><div class="v">${baseline['nav_usd']:,.2f}</div><div class="muted">Report {html.escape(baseline['report_date'])}</div></div>
<div class="card metric"><div class="muted">Coverage</div><div class="v">{valid}/{total}</div><div class="muted">Bloomberg option marks</div></div>
</div>
{warning}
<div class="card">
<div class="muted">Shared P&amp;L baseline</div>
<strong>{html.escape(baseline['filename'])}</strong> · Report Date {html.escape(baseline['report_date'])} · NAV ${baseline['nav_usd']:,.2f}
<br><div class="muted" style="margin-top:6px">
Bloomberg snapshot: {html.escape(format_new_york_time(snapshot_ts)) if snapshot_ts else '—'}
{(' · Collector: '+html.escape(collector_id)) if collector_id else ''}
<br>Current Mark = Bloomberg option Bid/Ask midpoint when both are available; otherwise Bloomberg Last.
Formula: (Current Mark − Previous P&amp;L Report Price) × Quantity × Multiplier.
Daily Return = Total Estimated P&amp;L ÷ Baseline NAV.
</div></div>
<div class="card"><table><thead><tr><th>Position</th><th>Qty</th><th>Previous Mark</th><th>Bloomberg Bid</th><th>Bloomberg Ask</th><th>Bloomberg Market</th><th>Manual Override</th><th>Effective Price</th><th>Change</th><th>Estimated P&amp;L</th><th>Contribution</th><th>Effective Source</th></tr></thead><tbody>{rows}</tbody></table></div>'''
    return page("Bloomberg Daily Return",body,u,60)

@app.get("/admin/daily-return")
def daily_return_admin(request:Request):
    u=require_admin(request)
    with SessionLocal() as db:
        b=_active_daily_baseline(db)
        current="<p>No P&amp;L baseline loaded.</p>"
        if b:
            pos=db.query(DailyReturnPosition).filter_by(baseline_id=b.id).order_by(DailyReturnPosition.id).all()
            prows=""
            for p in pos:
                strike="" if p.strike is None else f"{p.strike:g}"
                name=f"{p.ticker} {p.expiry} {p.option_type}{strike}" if p.instrument_type=="OPTION" else p.ticker
                prows+=f'<tr><td>{html.escape(name)}</td><td>{p.quantity:,.0f}</td><td>${p.baseline_price:.2f}</td></tr>'
            current=f'''<div class="notice"><strong>Active baseline:</strong> {html.escape(b.filename)} · Report {html.escape(b.report_date)} · NAV ${b.nav_usd:,.2f} · Positions {len(pos)}</div><br><table><thead><tr><th>Position</th><th>Qty</th><th>Previous Mark</th></tr></thead><tbody>{prows}</tbody></table>'''

    body=admin_tabs()+f'''
<div class="card"><h2>Daily Return — P&amp;L Baseline</h2>
<p>Upload the previous trading day's Nirvana <strong>PNL Report PDF</strong>. That single PDF supplies positions, quantities, previous marks and NAV and is shared by both Bloomberg Daily Return and Yahoo Daily Return.</p>
<form method="post" action="/admin/daily-return/upload" enctype="multipart/form-data"><input type="file" name="file" accept=".pdf,application/pdf" required><button>Upload P&amp;L Baseline</button></form><br>{current}</div>
<div class="card"><h3>Calculation</h3>
<p><strong>Current Mark = (Yahoo Bid + Yahoo Ask) ÷ 2 when both are available; otherwise Yahoo Last.</strong></p>
<p><strong>Position P&amp;L = (Current Mark − Previous P&amp;L Report Price) × Quantity × Multiplier</strong></p>
<p><strong>Estimated Daily Return = Total Estimated P&amp;L ÷ P&amp;L Report NAV</strong></p>
<p class="muted">No Exposure-by-Underlying upload is required. Yahoo option-chain expiration discovery is not used. This Daily Return module is separate from Bloomberg and does not change Bloomberg positions, NAV, Collector, snapshots, config version, users or limits.</p></div>'''
    return page("Daily Return Upload",body,u)


@app.post("/admin/daily-return/manual-price")
async def daily_return_manual_price(request:Request):
    u=require_admin(request)
    form=await request.form()
    try:
        position_id=int(form.get("position_id"))
    except Exception:
        raise HTTPException(400,"Invalid position id.")
    source=str(form.get("source") or "").upper()
    clear=str(form.get("clear") or "")=="1"
    raw=str(form.get("price") or "").strip()

    value=None
    if not clear:
        try:
            value=float(raw)
            if value < 0:
                raise ValueError
        except Exception:
            raise HTTPException(400,"Manual price must be zero or greater.")

    with SessionLocal() as db:
        p=db.query(DailyReturnPosition).filter_by(id=position_id).first()
        if not p:
            raise HTTPException(404,"Position not found.")
        if source=="YAHOO":
            p.yahoo_manual_price=value
            action="YAHOO_MANUAL_PRICE_CLEARED" if clear else "YAHOO_MANUAL_PRICE_SET"
            redirect="/daily-return"
        elif source=="BLOOMBERG":
            p.bloomberg_manual_price=value
            action="BLOOMBERG_MANUAL_PRICE_CLEARED" if clear else "BLOOMBERG_MANUAL_PRICE_SET"
            redirect="/bloomberg-daily-return"
        else:
            raise HTTPException(400,"Unknown source.")
        audit(db,u["username"],action,f"position_id={position_id}; price={value}")
        db.commit()

    return RedirectResponse(redirect,303)

@app.post("/admin/daily-return/upload")
async def daily_return_upload(request:Request,file:UploadFile=File(...)):
    u=require_admin(request)
    filename=file.filename or "PNL.pdf"
    if not filename.lower().endswith(".pdf"):
        raise HTTPException(400,"Please upload the Nirvana PNL Report as PDF.")
    data=await file.read()
    if len(data)>8*1024*1024:
        raise HTTPException(400,"PDF is too large.")
    try:
        parsed=parse_nirvana_pnl_pdf(data)
    except Exception as e:
        raise HTTPException(400,f"Could not parse PNL Report: {e}")

    with SessionLocal() as db:
        for old in db.query(DailyReturnBaseline).filter_by(active=True).all():
            old.active=False
        b=DailyReturnBaseline(report_date=parsed["report_date"],run_date=parsed["run_date"],filename=filename,
                              nav_usd=parsed["nav_usd"],active=True,uploaded_by=u["username"])
        db.add(b);db.flush()
        for x in parsed["positions"]:
            db.add(DailyReturnPosition(
                baseline_id=b.id,security_name=x["security_name"],instrument_type=x["instrument_type"],ticker=x["ticker"],
                expiry=x.get("expiry","") or "",option_type=x.get("option_type","") or "",strike=x.get("strike"),
                quantity=x["quantity"],multiplier=x["multiplier"],baseline_price=x.get("baseline_price"),
                baseline_market_value=x["baseline_market_value"],manual_price=None,yahoo_manual_price=None,bloomberg_manual_price=None))
        audit(db,u["username"],"DAILY_RETURN_PNL_BASELINE_UPLOADED",
              f'{filename}; report_date={parsed["report_date"]}; nav={parsed["nav_usd"]}; positions={len(parsed["positions"])}')
        db.commit()

    return RedirectResponse("/admin/daily-return",303)

def collector_auth(request):
    if not COLLECTOR_TOKEN or not hmac.compare_digest(request.headers.get("authorization",""),f"Bearer {COLLECTOR_TOKEN}"):
        raise HTTPException(401,"Bad collector token")


@app.get("/api/yahoo-collector/ping")
def yahoo_collector_ping(request:Request):
    collector_auth(request)
    return {"ok":True,"service":"yahoo-local-collector"}

@app.get("/api/yahoo-collector/config")
def yahoo_collector_config(request:Request):
    """Return the active Daily Return baseline positions to the local Yahoo collector."""
    collector_auth(request)
    with SessionLocal() as db:
        b=_active_daily_baseline(db)
        if not b:
            return {
                "version":0,
                "baseline_id":None,
                "report_date":None,
                "positions":[],
            }

        dbpos=db.query(DailyReturnPosition).filter_by(
            baseline_id=b.id
        ).order_by(DailyReturnPosition.id).all()

        positions=[]
        for p in dbpos:
            positions.append({
                "id":p.id,
                "instrument_type":str(p.instrument_type or "OPTION").upper(),
                "ticker":p.ticker,
                "expiry":p.expiry,
                "option_type":p.option_type,
                "strike":p.strike,
                "quantity":p.quantity,
                "multiplier":p.multiplier,
            })

        return {
            "version":int(b.id),
            "baseline_id":b.id,
            "report_date":b.report_date,
            "nav_usd":float(b.nav_usd or 0),
            "positions":positions,
        }

@app.post("/api/yahoo-collector/snapshot")
async def yahoo_collector_snapshot(request:Request):
    """Store the latest local Yahoo option-chain snapshot in Settings."""
    collector_auth(request)
    data=await request.json()

    if not isinstance(data,dict):
        raise HTTPException(400,"Invalid snapshot payload.")

    positions=data.get("positions")
    if not isinstance(positions,list):
        raise HTTPException(400,"positions must be a list.")

    # Keep payload intentionally small and derived only.
    cleaned=[]
    for r in positions[:200]:
        if not isinstance(r,dict):
            continue
        cleaned.append({
            "symbol":str(r.get("symbol") or "")[:64],
            "ticker":str(r.get("ticker") or "")[:30],
            "expiry":str(r.get("expiry") or "")[:10],
            "option_type":str(r.get("option_type") or "")[:4],
            "strike":r.get("strike"),
            "bid":r.get("bid"),
            "ask":r.get("ask"),
            "last":r.get("last"),
            "mark":r.get("mark"),
            "source":str(r.get("source") or "")[:80],
            "note":str(r.get("note") or "")[:160],
        })

    stored={
        "collector_name":str(data.get("collector_name") or "Yahoo-PC")[:80],
        "baseline_id":data.get("baseline_id"),
        "config_version":data.get("config_version"),
        "collected_at_utc":str(data.get("collected_at_utc") or utcnow().isoformat()),
        "coverage_valid":int(data.get("coverage_valid") or 0),
        "coverage_total":int(data.get("coverage_total") or len(cleaned)),
        "error":str(data.get("error") or "")[:1000],
        "positions":cleaned,
    }

    with SessionLocal() as db:
        set_setting(db,"yahoo_collector_snapshot",json.dumps(stored))
        set_setting(db,"yahoo_collector_updated_at",stored["collected_at_utc"])
        db.commit()

    return {"ok":True,"positions_received":len(cleaned)}


def _fortress_daily_positions(db):
    b=_active_daily_baseline(db)
    if not b:return None,[]
    dbpos=db.query(DailyReturnPosition).filter_by(baseline_id=b.id).order_by(DailyReturnPosition.id).all()
    return b,[{"id":p.id,"security_name":p.security_name,"instrument_type":p.instrument_type,"ticker":p.ticker,"expiry":p.expiry,"option_type":p.option_type,"strike":p.strike,"quantity":p.quantity,"multiplier":p.multiplier,"baseline_price":p.baseline_price,"baseline_market_value":p.baseline_market_value,"yahoo_manual_price":p.yahoo_manual_price,"bloomberg_manual_price":p.bloomberg_manual_price} for p in dbpos]

def _fortress_delta_positions(db):
    today=date.today().isoformat();out=[]
    for p in db.query(Position).filter_by(active=True).order_by(Position.ticker,Position.expiry,Position.strike).all():
        if p.expiry<today:continue
        out.append({"id":p.id,"instrument_type":"OPTION","ticker":p.ticker,"expiry":p.expiry,"option_type":p.option_type,"strike":p.strike,"quantity":p.quantity,"multiplier":p.multiplier,"bloomberg_security":p.bloomberg_security,"underlying_security":p.underlying_security})
    return out

@app.get("/api/collector/config-v2")
def collector_config_v2(request:Request):
    collector_auth(request)
    with SessionLocal() as db:
        fb,fdaily=_fortress_daily_positions(db);fdelta=_fortress_delta_positions(db);fcfg=int(get_setting(db,"config_version","1"));flimit=float(get_setting(db,"limit_pct","15"));fnav=float(get_setting(db,"nav_usd","0"))
        ob=_opportunity_baseline(db);oe=_opportunity_exposure(db);ocfg=int(get_setting(db,"opportunity_config_version","1") or 1);olimit=_opp_limit(db)
        opos_daily=list((ob or {}).get("positions") or []);opos_delta=list((oe or {}).get("positions") or [])
        fmerged=_merge_market_positions(fdelta,fdaily);omerged=_merge_market_positions(opos_delta,opos_daily)
        version=f'F{fcfg}-FB{getattr(fb,"id",0) if fb else 0}-O{ocfg}-OB{(ob or {}).get("version",0)}'
        return {"version":version,"poll_seconds":120,"funds":{"fortress":{"fund_key":"fortress","portfolio_name":"RPD Fortress Fund","version":fcfg,"nav_usd":fnav,"limit_pct":flimit,"positions":fmerged},"opportunity":{"fund_key":"opportunity","portfolio_name":"RPD Opportunity Fund","version":ocfg,"nav_usd":float((oe or {}).get("nav_usd") or 0),"limit_pct":olimit,"positions":omerged}}}

@app.post("/api/collector/snapshot-v2")
async def collector_snapshot_v2(request:Request):
    collector_auth(request);data=await request.json();funds=(data or {}).get("funds") or {}
    with SessionLocal() as db:
        fp=funds.get("fortress")
        if isinstance(fp,dict):
            db.add(Snapshot(payload=json.dumps(fp)));db.flush();count=db.query(Snapshot).count()
            if count>250:
                for x in db.query(Snapshot).order_by(Snapshot.id.asc()).limit(count-250).all():db.delete(x)
        op=funds.get("opportunity")
        if isinstance(op,dict): set_setting(db,"opportunity_bloomberg_snapshot",json.dumps(op))
        db.commit()
    return {"ok":True,"funds_received":[k for k,v in funds.items() if isinstance(v,dict)]}

@app.get("/api/yahoo-collector/config-v2")
def yahoo_collector_config_v2(request:Request):
    collector_auth(request)
    with SessionLocal() as db:
        fb,fdaily=_fortress_daily_positions(db);ob=_opportunity_baseline(db);odaily=list((ob or {}).get("positions") or [])
        def clean(rows):
            out=[]
            for p in rows:
                inst=str(p.get("instrument_type") or "OPTION").upper()
                if inst not in ("OPTION","EQUITY"):continue
                out.append({"id":p.get("id"),"instrument_type":inst,"ticker":p.get("ticker"),"expiry":p.get("expiry") or "","option_type":p.get("option_type") or "","strike":p.get("strike"),"quantity":p.get("quantity"),"multiplier":p.get("multiplier",100 if inst=="OPTION" else 1)})
            return out
        version=f'FB{getattr(fb,"id",0) if fb else 0}-OB{(ob or {}).get("version",0)}'
        return {"version":version,"funds":{"fortress":{"fund_key":"fortress","portfolio_name":"RPD Fortress Fund","baseline_id":getattr(fb,"id",None) if fb else None,"report_date":getattr(fb,"report_date",None) if fb else None,"positions":clean(fdaily)},"opportunity":{"fund_key":"opportunity","portfolio_name":"RPD Opportunity Fund","baseline_id":None,"report_date":(ob or {}).get("report_date"),"positions":clean(odaily)}}}

@app.post("/api/yahoo-collector/snapshot-v2")
async def yahoo_collector_snapshot_v2(request:Request):
    collector_auth(request);data=await request.json()
    if not isinstance(data,dict) or not isinstance(data.get("funds"),dict): raise HTTPException(400,"Invalid multi-fund Yahoo snapshot.")
    with SessionLocal() as db:
        set_setting(db,"yahoo_collector_snapshot_v2",json.dumps(data));set_setting(db,"yahoo_collector_v2_updated_at",str(data.get("collected_at_utc") or utcnow().isoformat()));db.commit()
    return {"ok":True,"funds_received":list(data.get("funds",{}).keys())}

@app.get("/api/collector/ping")
def collector_ping(request:Request):
    collector_auth(request);return {"ok":True}

@app.get("/api/collector/config")
def collector_config(request:Request):
    collector_auth(request)
    with SessionLocal() as db:
        today=date.today().isoformat();positions=[]
        for p in db.query(Position).filter_by(active=True).order_by(Position.ticker,Position.expiry,Position.strike).all():
            if p.expiry<today:continue
            positions.append({"id":p.id,"ticker":p.ticker,"expiry":p.expiry,"option_type":p.option_type,"strike":p.strike,"quantity":p.quantity,"multiplier":p.multiplier,"bloomberg_security":p.bloomberg_security,"underlying_security":p.underlying_security})
        return {"version":int(get_setting(db,"config_version","1")),"updated_at":get_setting(db,"config_updated_at",""),"portfolio_name":get_setting(db,"portfolio_name","RPD Fortress Fund"),"nav_usd":float(get_setting(db,"nav_usd","0")),"limit_pct":float(get_setting(db,"limit_pct","15")),"positions":positions}

@app.post("/api/collector/snapshot")
async def collector_snapshot(request:Request):
    collector_auth(request);data=await request.json()
    with SessionLocal() as db:
        db.add(Snapshot(payload=json.dumps(data)));db.flush()
        count=db.query(Snapshot).count()
        if count>250:
            for x in db.query(Snapshot).order_by(Snapshot.id.asc()).limit(count-250).all():db.delete(x)
        db.commit()
    return {"ok":True}
