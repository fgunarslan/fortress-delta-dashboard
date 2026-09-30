from __future__ import annotations
import os,json,hmac,html,uuid,re
from datetime import datetime,timezone,date
from zoneinfo import ZoneInfo
from urllib.parse import quote,unquote
from fastapi import FastAPI,Request,Form,UploadFile,File,HTTPException
from fastapi.responses import HTMLResponse,RedirectResponse

from db import SessionLocal,User,Setting,Position,Audit,Snapshot,ImportJob,get_setting,set_setting,audit,config_bump,utcnow
from security import ADMIN_USER,ensure_admin,hash_password,verify_password,sign_session,read_session
from helpers import option_security,detect_map,parse_upload,parse_nirvana_exposure_rows,normalized_import,keypos

COLLECTOR_TOKEN=os.getenv("COLLECTOR_TOKEN","")
STALE_AFTER_SECONDS=max(int(os.getenv("STALE_AFTER_SECONDS","150")),150)
app=FastAPI(title="Fortress Delta Dashboard v2",docs_url=None,redoc_url=None)

def init_data():
    ensure_admin()
    with SessionLocal() as db:
        defaults={"portfolio_name":"RPD Fortress Fund","nav_usd":"46865138.06795","limit_pct":"15.0","config_version":"1","config_updated_at":utcnow().isoformat()}
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
    auth=f'<span>{html.escape(user["username"])}</span> {adminlink} <a href="/logout">Logout</a>' if user else ""
    rf=f'<meta http-equiv="refresh" content="{refresh}">' if refresh else ""
    return HTMLResponse(f"""<!doctype html><html><head><meta charset="utf-8">{rf}<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)}</title><style>
body{{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f5f7fb;color:#172033}}.wrap{{max-width:1280px;margin:auto;padding:22px}}
nav{{display:flex;justify-content:space-between;margin-bottom:18px}}nav a{{margin-left:12px}}.brand{{color:#172033;text-decoration:none}}.card{{background:#fff;border:1px solid #dfe5ef;border-radius:12px;padding:16px;margin-bottom:14px}}
.grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}}.metric .v{{font-size:25px;font-weight:700;margin-top:6px}}.muted{{color:#6d7b92;font-size:13px}}
.ok{{color:#067647}}.bad{{color:#b42318}}.notice{{padding:10px 12px;background:#fff7e6;border:1px solid #ffd591;border-radius:8px}}
table{{width:100%;border-collapse:collapse}}th,td{{border-bottom:1px solid #e6ebf2;padding:9px;text-align:right;font-size:13px}}th:first-child,td:first-child{{text-align:left}}
input,select,button{{padding:8px;border:1px solid #c9d3e1;border-radius:7px}}button{{background:#172b4d;color:white;cursor:pointer}}.danger{{background:#b42318}}.secondary{{background:#52657d}}
.row{{display:flex;gap:8px;align-items:end;flex-wrap:wrap}}label{{display:flex;flex-direction:column;font-size:12px;color:#52657d;gap:4px}}
@media(max-width:900px){{.grid{{grid-template-columns:1fr 1fr}}}}
</style></head><body><div class="wrap"><nav><strong><a class="brand" href="/">Fortress Delta Monitor</a></strong><div>{auth}</div></nav>{body}</div></body></html>""")

def admin_tabs():
    return '<p><a href="/">Dashboard</a> · <a href="/admin">Portfolio</a> · <a href="/admin/import">Nirvana Import</a> · <a href="/admin/users">Viewer Users</a> · <a href="/admin/audit">Audit Log</a></p>'

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

def collector_auth(request):
    if not COLLECTOR_TOKEN or not hmac.compare_digest(request.headers.get("authorization",""),f"Bearer {COLLECTOR_TOKEN}"):
        raise HTTPException(401,"Bad collector token")

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
