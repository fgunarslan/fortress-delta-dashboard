from __future__ import annotations

import io
import math
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, time as dtime
from zoneinfo import ZoneInfo

from pypdf import PdfReader
import yfinance as yf

NY = ZoneInfo("America/New_York")
OPTION_RE = re.compile(
    r"^(?P<ticker>[A-Z0-9.\-]+)\s+(?P<expiry>\d{1,2}/\d{1,2}/\d{2})\s+(?P<strike>\d+(?:\.\d+)?)\s+(?P<type>PUT|CALL)$",
    re.I,
)

_quote_lock = threading.Lock()
_chain_cache = {}
_spot_cache = {}
CACHE_SECONDS = 55


def _num(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace("$", "").strip()
    if s in {"", "-", "—"}:
        return None
    try:
        x = float(s)
    except Exception:
        return None
    if not math.isfinite(x):
        return None
    return -x if neg else x


def _extract_layout_text(data: bytes) -> str:
    reader = PdfReader(io.BytesIO(data))
    pages = []
    for page in reader.pages:
        try:
            pages.append(page.extract_text(extraction_mode="layout") or "")
        except TypeError:
            pages.append(page.extract_text() or "")
    return "\n".join(pages)


def _column_slices(header: str):
    labels = [
        "Security Name", "Price", "Change", "%", "Quantity", "Unit Cost",
        "Market Value", "Day P&L", "MTD P&L", "YTD P&L", "ITD P&L",
    ]
    starts = []
    cursor = 0
    for lab in labels:
        idx = header.find(lab, cursor)
        if idx < 0:
            raise ValueError(f"PNL report column not found: {lab}")
        starts.append((lab, idx))
        cursor = idx + len(lab)
    slices = {}
    for i, (lab, start) in enumerate(starts):
        end = starts[i + 1][1] if i + 1 < len(starts) else None
        slices[lab] = (start, end)
    return slices


def _cell(line: str, col_slices, name: str) -> str:
    start, end = col_slices[name]
    if len(line) <= start:
        return ""
    return line[start:end].strip() if end is not None else line[start:].strip()


def parse_nirvana_pnl_pdf(data: bytes) -> dict:
    """Parse the Nirvana PNL PDF used as the previous-close return baseline."""
    text = _extract_layout_text(data)
    lines = text.splitlines()
    if not any("PNL Report" in x for x in lines[:20]):
        raise ValueError("This does not look like a Nirvana PNL Report PDF.")

    report_date = None
    run_date = None
    for line in lines[:30]:
        m = re.search(r"Report Date:\s*(\d{1,2}/\d{1,2}/\d{4})", line, re.I)
        if m:
            report_date = datetime.strptime(m.group(1), "%m/%d/%Y").date().isoformat()
        m2 = re.search(r"Run Date:\s*([^\n]+)$", line, re.I)
        if m2:
            run_date = m2.group(1).strip()

    header_idx = None
    for i, line in enumerate(lines):
        if "Security Name" in line and "Market Value" in line and "Day P&L" in line:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Could not locate the PNL table header.")
    cols = _column_slices(lines[header_idx])

    positions = []
    nav = None
    ignored = []
    for line in lines[header_idx + 1:]:
        if not line.strip():
            continue
        name = _cell(line, cols, "Security Name")
        if not name:
            continue
        if name.startswith("Page "):
            continue
        if name.startswith("Grand Total"):
            gm = re.search(r"Grand Total\s*:\s*([0-9,]+(?:\.[0-9]+)?)", line, re.I)
            nav = _num(gm.group(1)) if gm else _num(_cell(line, cols, "Market Value"))
            continue
        if name.upper() == "CLOSED":
            continue

        price = _num(_cell(line, cols, "Price"))
        qty = _num(_cell(line, cols, "Quantity"))
        market_value = _num(_cell(line, cols, "Market Value"))
        if qty is None or market_value is None:
            ignored.append(name)
            continue

        upper = name.upper().strip()
        if upper in {"FGTXX", "USD", "CASH"}:
            # Cash / money market stays in baseline NAV but is not live-repriced.
            continue

        om = OPTION_RE.match(upper)
        if om:
            expiry = datetime.strptime(om.group("expiry"), "%m/%d/%y").date().isoformat()
            typ = "P" if om.group("type").upper() == "PUT" else "C"
            strike = float(om.group("strike"))
            multiplier = 100.0
            baseline_mark = (market_value / (qty * multiplier)) if qty else price
            if baseline_mark is None or baseline_mark <= 0:
                baseline_mark = price
            positions.append({
                "security_name": name.strip(),
                "instrument_type": "OPTION",
                "ticker": om.group("ticker").upper(),
                "expiry": expiry,
                "option_type": typ,
                "strike": strike,
                "quantity": qty,
                "multiplier": multiplier,
                "baseline_price": baseline_mark,
                "display_price": price,
                "baseline_market_value": market_value,
            })
            continue

        # Basic stock support for future PNL reports.
        ticker = upper.split()[0]
        if re.fullmatch(r"[A-Z][A-Z0-9.\-]{0,14}", ticker):
            multiplier = 1.0
            baseline_mark = (market_value / qty) if qty else price
            positions.append({
                "security_name": name.strip(),
                "instrument_type": "EQUITY",
                "ticker": ticker,
                "expiry": "",
                "option_type": "",
                "strike": None,
                "quantity": qty,
                "multiplier": multiplier,
                "baseline_price": baseline_mark,
                "display_price": price,
                "baseline_market_value": market_value,
            })
        else:
            ignored.append(name)

    if nav is None or nav <= 0:
        raise ValueError("Could not read Grand Total / NAV from the PNL report.")
    if not positions:
        raise ValueError("No live option/equity positions were parsed from the PNL report.")

    return {
        "report_date": report_date or "",
        "run_date": run_date or "",
        "nav_usd": float(nav),
        "positions": positions,
        "ignored": ignored,
    }


def _positive(v):
    try:
        x = float(v)
        return x if math.isfinite(x) and x > 0 else None
    except Exception:
        return None


def _chain_key(ticker, expiry):
    return (str(ticker).upper(), str(expiry))


def _load_chain(ticker: str, expiry: str):
    key = _chain_key(ticker, expiry)
    now = time.time()
    with _quote_lock:
        hit = _chain_cache.get(key)
        if hit and now - hit[0] < CACHE_SECONDS:
            return hit[1]
    try:
        t = yf.Ticker(ticker)
        chain = t.option_chain(expiry)
        payload = (chain.calls.copy(), chain.puts.copy(), None)
    except Exception as e:
        payload = (None, None, str(e))
    with _quote_lock:
        _chain_cache[key] = (now, payload)
    return payload


def _load_spot(ticker: str):
    ticker = str(ticker).upper()
    now = time.time()
    with _quote_lock:
        hit = _spot_cache.get(ticker)
        if hit and now - hit[0] < CACHE_SECONDS:
            return hit[1]
    px = None
    source = ""
    err = ""
    try:
        t = yf.Ticker(ticker)
        fi = t.fast_info
        for key in ("last_price", "regular_market_price", "previous_close"):
            try:
                v = _positive(fi[key])
                if v is not None:
                    px = v
                    source = f"Yahoo {key}"
                    break
            except Exception:
                pass
        if px is None:
            hist = t.history(period="1d", interval="1m", auto_adjust=False, prepost=True)
            if hist is not None and not hist.empty:
                v = _positive(hist["Close"].dropna().iloc[-1])
                if v is not None:
                    px = v
                    source = "Yahoo 1m close"
    except Exception as e:
        err = str(e)
    result = (px, source, err or ("price unavailable" if px is None else ""))
    with _quote_lock:
        _spot_cache[ticker] = (now, result)
    return result


def _option_mark_from_chain(pos, chain_payload):
    calls, puts, err = chain_payload
    if err:
        return None, "UNAVAILABLE", err
    df = puts if pos.get("option_type") == "P" else calls
    if df is None or df.empty:
        return None, "UNAVAILABLE", "empty Yahoo option chain"
    try:
        strikes = df["strike"].astype(float)
        hits = df[(strikes - float(pos["strike"])).abs() < 1e-6]
        if hits.empty:
            return None, "UNAVAILABLE", f"strike {pos['strike']:g} not found"
        row = hits.iloc[0]
        bid = _positive(row.get("bid"))
        ask = _positive(row.get("ask"))
        last = _positive(row.get("lastPrice"))
        if bid is not None and ask is not None and ask >= bid:
            return (bid + ask) / 2.0, "Yahoo Bid/Ask Mid", f"bid {bid:.3f} / ask {ask:.3f}"
        if last is not None:
            return last, "Yahoo Last", "no usable two-sided quote"
        return None, "UNAVAILABLE", "no usable bid/ask or last"
    except Exception as e:
        return None, "UNAVAILABLE", str(e)


def live_marks(positions: list[dict]) -> dict[int, tuple]:
    """Return {position_id: (mark, source, note)} with 55-second caching."""
    result = {}
    option_groups = {}
    equities = []
    for p in positions:
        if p.get("instrument_type") == "OPTION":
            option_groups.setdefault(_chain_key(p["ticker"], p["expiry"]), []).append(p)
        elif p.get("instrument_type") == "EQUITY":
            equities.append(p)

    def task_chain(key):
        return key, _load_chain(*key)

    def task_spot(p):
        return p, _load_spot(p["ticker"])

    with ThreadPoolExecutor(max_workers=4) as ex:
        futures = [ex.submit(task_chain, key) for key in option_groups]
        futures += [ex.submit(task_spot, p) for p in equities]
        for fut in as_completed(futures):
            item = fut.result()
            if isinstance(item[0], tuple):
                key, payload = item
                for p in option_groups[key]:
                    result[int(p["id"])] = _option_mark_from_chain(p, payload)
            else:
                p, (px, src, err) = item
                result[int(p["id"])] = (px, src or "UNAVAILABLE", err)
    return result


def calculate_estimated_return(positions: list[dict], nav_usd: float) -> dict:
    marks = live_marks(positions)
    rows = []
    total_pnl = 0.0
    valid = 0
    for p in positions:
        override = p.get("manual_price")
        if override is not None and float(override) > 0:
            mark = float(override)
            source = "Manual Price Override"
            note = "admin override"
        else:
            mark, source, note = marks.get(int(p["id"]), (None, "UNAVAILABLE", "quote not returned"))
        pnl = None
        current_mv = None
        contrib = None
        if mark is not None:
            current_mv = float(p["quantity"]) * float(p.get("multiplier", 1)) * float(mark)
            pnl = current_mv - float(p["baseline_market_value"])
            contrib = (pnl / nav_usd * 100.0) if nav_usd else None
            total_pnl += pnl
            valid += 1
        rows.append({
            **p,
            "current_mark": mark,
            "price_source": source,
            "note": note,
            "current_market_value": current_mv,
            "estimated_pnl": pnl,
            "contribution_pct": contrib,
        })
    total = len(positions)
    return {
        "rows": rows,
        "estimated_pnl": total_pnl,
        "estimated_return_pct": (total_pnl / nav_usd * 100.0) if nav_usd else None,
        "coverage_valid": valid,
        "coverage_total": total,
        "complete": bool(total and valid == total),
        "calculated_at_utc": datetime.now(timezone.utc).isoformat(),
    }


def auto_refresh_allowed(now_utc=None):
    now_utc = now_utc or datetime.now(timezone.utc)
    ny = now_utc.astimezone(NY)
    t = ny.time().replace(tzinfo=None)
    return ny.weekday() < 5 and dtime(9, 30) <= t < dtime(16, 15)


def ny_time_label(timestamp_utc: str):
    try:
        ts = datetime.fromisoformat(str(timestamp_utc).replace("Z", "+00:00"))
        return ts.astimezone(NY).strftime("%d %b %Y, %H:%M:%S %Z")
    except Exception:
        return str(timestamp_utc or "")
