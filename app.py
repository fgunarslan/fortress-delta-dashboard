from __future__ import annotations
import os, secrets, threading
from datetime import datetime, timezone
from typing import Any, Dict, Optional

from fastapi import FastAPI, Request, HTTPException, Depends
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from pydantic import BaseModel, Field

INGEST_TOKEN = os.getenv("INGEST_TOKEN", "")
VIEW_USER = os.getenv("VIEW_USER", "fuat")
VIEW_PASSWORD = os.getenv("VIEW_PASSWORD", "")
STALE_AFTER_SECONDS = int(os.getenv("STALE_AFTER_SECONDS", "20"))

if not INGEST_TOKEN:
    print("WARNING: INGEST_TOKEN is not set. /api/ingest will reject all requests.")
if not VIEW_PASSWORD:
    print("WARNING: VIEW_PASSWORD is not set. Dashboard authentication will reject access.")

app = FastAPI(title="Fortress Delta Dashboard", docs_url=None, redoc_url=None)
security = HTTPBasic()
_lock = threading.Lock()
_latest: Optional[Dict[str, Any]] = None

class Coverage(BaseModel):
    valid: int = 0
    total: int = 0

class PositionRow(BaseModel):
    ticker: str
    expiry: str
    option_type: str
    strike: float
    quantity: float
    delta_exposure_usd: float
    contribution_pct: float

class Snapshot(BaseModel):
    mode: str = "LIVE"
    portfolio: str
    collector_id: str
    timestamp_utc: str
    nav_usd: float
    limit_pct: float
    delta_exposure_usd: float
    delta_exposure_pct: float
    buffer_pct: float
    coverage: Coverage
    positions: list[PositionRow] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)

def require_view(credentials: HTTPBasicCredentials = Depends(security)):
    good_user = secrets.compare_digest(credentials.username or "", VIEW_USER or "")
    good_pass = secrets.compare_digest(credentials.password or "", VIEW_PASSWORD or "")
    if not (good_user and good_pass):
        raise HTTPException(status_code=401, detail="Unauthorized", headers={"WWW-Authenticate":"Basic"})
    return True

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/api/ingest")
async def ingest(request: Request):
    auth = request.headers.get("authorization","")
    expected = f"Bearer {INGEST_TOKEN}"
    if not INGEST_TOKEN or not secrets.compare_digest(auth, expected):
        raise HTTPException(status_code=401, detail="Bad ingest token")
    payload = Snapshot.model_validate(await request.json()).model_dump()
    payload["_server_received_utc"] = datetime.now(timezone.utc).isoformat()
    global _latest
    with _lock:
        _latest = payload
    return {"ok": True}

@app.get("/api/snapshot")
def snapshot(_: bool = Depends(require_view)):
    with _lock:
        data = dict(_latest) if _latest else None
    if not data:
        return JSONResponse({"available":False,"stale":True,"stale_seconds":None})
    try:
        ts = datetime.fromisoformat(data["timestamp_utc"].replace("Z","+00:00"))
        age = (datetime.now(timezone.utc)-ts.astimezone(timezone.utc)).total_seconds()
    except Exception:
        age = 999999
    data["available"] = True
    data["stale_seconds"] = round(age,1)
    data["stale"] = age > STALE_AFTER_SECONDS
    return data

DASHBOARD = r"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Fortress Delta Monitor</title>
<style>
:root{--bg:#0b1220;--card:#131d2f;--muted:#9db0ca;--text:#eef4ff;--green:#32d583;--yellow:#fdb022;--red:#f97066;--line:#263754;}
*{box-sizing:border-box} body{margin:0;background:var(--bg);color:var(--text);font-family:Inter,Segoe UI,Arial,sans-serif}
.wrap{max-width:1220px;margin:0 auto;padding:24px}
h1{font-size:26px;margin:0 0 4px}.sub{color:var(--muted);margin-bottom:18px}
.grid{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:14px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px}
.label{font-size:12px;text-transform:uppercase;color:var(--muted);letter-spacing:.06em}
.value{font-size:25px;font-weight:700;margin-top:6px}.small{font-size:13px;color:var(--muted);margin-top:6px}
.badge{display:inline-block;padding:5px 9px;border-radius:999px;font-weight:700;font-size:12px}
.ok{background:#113c2d;color:#70e0ad}.warn{background:#4a3512;color:#ffd27a}.bad{background:#4b1f24;color:#ffaaa3}
table{width:100%;border-collapse:collapse;background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden}
th,td{padding:10px 12px;border-bottom:1px solid var(--line);text-align:right;font-size:13px}
th{color:var(--muted);font-size:11px;text-transform:uppercase} th:first-child,td:first-child{text-align:left}
.footer{color:var(--muted);font-size:12px;margin-top:14px}
@media(max-width:850px){.grid{grid-template-columns:1fr 1fr}.wrap{padding:12px}}
</style>
</head>
<body><div class="wrap">
<h1 id="title">Fortress Delta Monitor</h1>
<div class="sub">Remote view of derived risk metrics from the Bloomberg Desktop collector. Raw Bloomberg quotes/Greeks are not displayed.</div>
<div class="grid">
 <div class="card"><div class="label">Feed Status</div><div class="value" id="status">WAITING</div><div class="small" id="age">No data</div></div>
 <div class="card"><div class="label">Delta Exposure</div><div class="value" id="pct">—</div><div class="small" id="usd">—</div></div>
 <div class="card"><div class="label">Limit</div><div class="value" id="limit">—</div><div class="small">Maximum delta exposure</div></div>
 <div class="card"><div class="label">Buffer</div><div class="value" id="buffer">—</div><div class="small">Distance to limit</div></div>
 <div class="card"><div class="label">Coverage</div><div class="value" id="coverage">—</div><div class="small" id="collector">—</div></div>
</div>
<div class="card" style="margin-bottom:12px">
 <div class="label">Last Bloomberg Collector Update</div>
 <div style="margin-top:7px" id="updated">—</div>
 <div class="small" id="warnings"></div>
</div>
<table>
<thead><tr><th>Position</th><th>Qty</th><th>Delta Exposure $</th><th>Contribution % NAV</th></tr></thead>
<tbody id="rows"><tr><td colspan="4" style="text-align:center;color:#9db0ca">Waiting for collector…</td></tr></tbody>
</table>
<div class="footer">If the collector stops, Bloomberg logs out, the PC sleeps, or the internet connection fails, this page changes to OFFLINE once the snapshot becomes stale.</div>
</div>
<script>
function money(x){return new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',maximumFractionDigits:0}).format(x)}
function num(x,n=3){return Number(x).toFixed(n)}
async function refresh(){
 try{
   const r=await fetch('/api/snapshot',{cache:'no-store'});
   if(!r.ok) throw new Error('HTTP '+r.status);
   const d=await r.json();
   if(!d.available){document.getElementById('status').textContent='WAITING';return}
   document.getElementById('title').textContent=d.portfolio+' — Bloomberg Delta Monitor';
   const status=document.getElementById('status');
   status.textContent=d.stale?'OFFLINE':(d.mode==='TEST'?'TEST':'LIVE');
   status.className='value '+(d.stale?'bad':(d.mode==='TEST'?'warn':''));
   document.getElementById('age').textContent='Age: '+d.stale_seconds+' sec';
   document.getElementById('pct').textContent=num(d.delta_exposure_pct)+'%';
   document.getElementById('usd').textContent=money(d.delta_exposure_usd);
   document.getElementById('limit').textContent=num(d.limit_pct)+'%';
   document.getElementById('buffer').textContent=num(d.buffer_pct)+'%';
   document.getElementById('coverage').textContent=d.coverage.valid+'/'+d.coverage.total;
   document.getElementById('collector').textContent=d.collector_id;
   document.getElementById('updated').textContent=d.timestamp_utc;
   document.getElementById('warnings').textContent=(d.warnings||[]).slice(0,3).join(' | ');
   const tb=document.getElementById('rows'); tb.innerHTML='';
   (d.positions||[]).sort((a,b)=>Math.abs(b.contribution_pct)-Math.abs(a.contribution_pct)).forEach(p=>{
     const tr=document.createElement('tr');
     tr.innerHTML=`<td>${p.ticker} ${p.expiry} ${p.option_type}${p.strike}</td><td>${Number(p.quantity).toLocaleString()}</td><td>${money(p.delta_exposure_usd)}</td><td>${num(p.contribution_pct)}%</td>`;
     tb.appendChild(tr);
   });
 }catch(e){
   document.getElementById('status').textContent='PORTAL ERROR';
 }
}
refresh(); setInterval(refresh,2000);
</script>
</body></html>"""

@app.get("/", response_class=HTMLResponse)
def dashboard(_: bool = Depends(require_view)):
    return HTMLResponse(DASHBOARD)
