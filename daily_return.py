from __future__ import annotations

import io
import json
import math
import re
import threading
import time
from datetime import datetime, timezone, time as dtime
from urllib.parse import urlencode, quote
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError
from zoneinfo import ZoneInfo

from pypdf import PdfReader

NY = ZoneInfo("America/New_York")
OPTION_RE = re.compile(
    r"^(?P<ticker>[A-Z0-9.\-]+)\s+(?P<expiry>\d{1,2}/\d{1,2}/\d{2})\s+(?P<strike>\d+(?:\.\d+)?)\s+(?P<type>PUT|CALL)$",
    re.I,
)

_quote_lock = threading.Lock()
_quote_cache = {"ts": 0.0, "symbols": (), "data": {}, "error": ""}
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


def _positive(v):
    try:
        x = float(v)
        return x if math.isfinite(x) and x > 0 else None
    except Exception:
        return None


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
    """Parse one Nirvana PNL Report PDF.

    Baseline:
      Previous Mark = PDF Price column exactly.
      Quantity = PDF Quantity.
      NAV = PDF Grand Total.
    """
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
        if not name or name.startswith("Page "):
            continue

        if name.startswith("Grand Total"):
            gm = re.search(r"Grand Total\s*:\s*([0-9,]+(?:\.[0-9]+)?)", line, re.I)
            nav = _num(gm.group(1)) if gm else _num(_cell(line, cols, "Market Value"))
            continue

        upper = name.upper().strip()
        if upper in {"CLOSED", "FGTXX", "USD", "CASH"}:
            continue

        price = _num(_cell(line, cols, "Price"))
        qty = _num(_cell(line, cols, "Quantity"))
        market_value = _num(_cell(line, cols, "Market Value"))
        if qty is None or price is None:
            ignored.append(name)
            continue

        om = OPTION_RE.match(upper)
        if om:
            expiry = datetime.strptime(om.group("expiry"), "%m/%d/%y").date().isoformat()
            typ = "P" if om.group("type").upper() == "PUT" else "C"
            strike = float(om.group("strike"))
            positions.append({
                "security_name": name.strip(),
                "instrument_type": "OPTION",
                "ticker": om.group("ticker").upper(),
                "expiry": expiry,
                "option_type": typ,
                "strike": strike,
                "quantity": float(qty),
                "multiplier": 100.0,
                "baseline_price": float(price),
                "display_price": float(price),
                "baseline_market_value": float(market_value or (price * qty * 100.0)),
            })
            continue

        ticker = upper.split()[0]
        if re.fullmatch(r"[A-Z][A-Z0-9.\-]{0,14}", ticker):
            positions.append({
                "security_name": name.strip(),
                "instrument_type": "EQUITY",
                "ticker": ticker,
                "expiry": "",
                "option_type": "",
                "strike": None,
                "quantity": float(qty),
                "multiplier": 1.0,
                "baseline_price": float(price),
                "display_price": float(price),
                "baseline_market_value": float(market_value or (price * qty)),
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


def yahoo_option_symbol(ticker: str, expiry: str, option_type: str, strike: float) -> str:
    dt = datetime.strptime(expiry, "%Y-%m-%d")
    root = str(ticker).upper().replace(".", "-").strip()
    cp = str(option_type).upper()[:1]
    strike_code = int(round(float(strike) * 1000.0))
    return f"{root}{dt.strftime('%y%m%d')}{cp}{strike_code:08d}"


def _market_symbol(p: dict) -> str:
    if p.get("instrument_type") == "OPTION":
        return yahoo_option_symbol(p["ticker"], p["expiry"], p["option_type"], p["strike"])
    return str(p["ticker"]).upper().replace(".", "-")


def _http_json(url: str, timeout: int = 10):
    req = Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/142 Safari/537.36",
        "Accept": "application/json,text/plain,*/*",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "close",
    })
    try:
        with urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8")), ""
    except HTTPError as e:
        if e.code == 429:
            return None, "Yahoo rate limited (HTTP 429)"
        return None, f"Yahoo HTTP {e.code}"
    except URLError as e:
        return None, f"Yahoo network error: {getattr(e, 'reason', e)}"
    except Exception as e:
        return None, f"Yahoo error: {e}"


def _last_number(values):
    if not isinstance(values, list):
        return None
    for v in reversed(values):
        x = _positive(v)
        if x is not None:
            return x
    return None


def _parse_quote_batch(payload, wanted):
    """Parse Yahoo quote response.

    Preferred mark:
      Bid + Ask available -> midpoint
      otherwise regularMarketPrice -> last fallback
    """
    out = {}
    try:
        results = ((payload or {}).get("quoteResponse") or {}).get("result") or []
        for item in results:
            symbol = str(item.get("symbol") or "").upper()
            if symbol not in wanted:
                continue

            bid = _positive(item.get("bid"))
            ask = _positive(item.get("ask"))
            last = _positive(item.get("regularMarketPrice"))
            if last is None:
                last = _positive(item.get("postMarketPrice"))
            if last is None:
                last = _positive(item.get("preMarketPrice"))

            mark = None
            source = ""
            if bid is not None and ask is not None and ask >= bid:
                mark = (bid + ask) / 2.0
                source = "Yahoo Bid/Ask Mid"
            elif last is not None:
                mark = last
                source = "Yahoo Last / regularMarketPrice"

            out[symbol] = {
                "mark": mark,
                "bid": bid,
                "ask": ask,
                "last": last,
                "source": source if mark is not None else "UNAVAILABLE",
                "note": "",
            }
    except Exception:
        pass
    return out


def _parse_spark(payload, wanted):
    out = {}
    try:
        results = payload.get("spark", {}).get("result", [])
        for item in results:
            symbol = str(item.get("symbol") or "").upper()
            responses = item.get("response") or []
            if not responses:
                continue
            r = responses[0] or {}
            meta = r.get("meta") or {}
            px = _positive(meta.get("regularMarketPrice"))
            source = "Yahoo regularMarketPrice"
            if px is None:
                q = ((r.get("indicators") or {}).get("quote") or [{}])[0]
                px = _last_number(q.get("close"))
                source = "Yahoo latest 1m price"
            if symbol in wanted and px is not None:
                out[symbol] = (px, source, "")
    except Exception:
        pass
    return out


def _chart_quote(symbol):
    enc = quote(symbol, safe="")
    last_err = "Yahoo price unavailable"
    for host in ("query1.finance.yahoo.com", "query2.finance.yahoo.com"):
        url = f"https://{host}/v8/finance/chart/{enc}?interval=1m&range=1d&includePrePost=false"
        payload, err = _http_json(url)
        if err:
            last_err = err
            continue
        try:
            results = payload.get("chart", {}).get("result") or []
            if not results:
                last_err = "Yahoo chart returned no result"
                continue
            r = results[0]
            meta = r.get("meta") or {}
            px = _positive(meta.get("regularMarketPrice"))
            if px is not None:
                return px, "Yahoo regularMarketPrice", ""
            q = ((r.get("indicators") or {}).get("quote") or [{}])[0]
            px = _last_number(q.get("close"))
            if px is not None:
                return px, "Yahoo latest 1m price", ""
        except Exception as e:
            last_err = f"Yahoo chart parse error: {e}"
    return None, "UNAVAILABLE", last_err


def yahoo_current_marks(positions: list[dict]) -> tuple[dict, str]:
    """Fetch current marks with minimal Yahoo requests.

    Preferred current mark for every option/equity:
      1. Bid and Ask both usable -> (Bid + Ask) / 2
      2. Otherwise Yahoo regularMarketPrice / latest last price

    No option-expiration discovery is used.
    """
    symbols = tuple(sorted({_market_symbol(p) for p in positions}))
    now = time.time()

    with _quote_lock:
        if _quote_cache["symbols"] == symbols and now - _quote_cache["ts"] < CACHE_SECONDS:
            return dict(_quote_cache["data"]), _quote_cache["error"]

    data = {}
    global_err = ""
    wanted = set(symbols)

    if symbols:
        # First choice: one batch quote request for Bid / Ask / Last.
        quote_qs = urlencode({"symbols": ",".join(symbols)})
        quote_rate_limited = False

        for host in ("query1.finance.yahoo.com", "query2.finance.yahoo.com"):
            payload, err = _http_json(f"https://{host}/v7/finance/quote?{quote_qs}")
            if err:
                global_err = err
                if "429" in err or "rate limited" in err.lower():
                    quote_rate_limited = True
                continue
            parsed = _parse_quote_batch(payload or {}, wanted)
            if parsed:
                data.update(parsed)
                break

        # For missing symbols or missing mark, use ONE spark batch as fallback.
        missing = [s for s in symbols if s not in data or data[s].get("mark") is None]
        if missing and not quote_rate_limited:
            spark_qs = urlencode({
                "symbols": ",".join(missing),
                "range": "1d",
                "interval": "1m",
                "indicators": "close",
                "includePrePost": "false",
            })
            for host in ("query1.finance.yahoo.com", "query2.finance.yahoo.com"):
                payload, err = _http_json(f"https://{host}/v7/finance/spark?{spark_qs}")
                if err:
                    global_err = err
                    continue
                spark_data = _parse_spark(payload or {}, set(missing))
                for symbol, tup in spark_data.items():
                    px, src, note = tup
                    existing = data.get(symbol, {})
                    existing.update({
                        "mark": px,
                        "last": px,
                        "source": src,
                        "note": note,
                    })
                    data[symbol] = existing
                if spark_data:
                    break

        # Final individual chart fallback only for still-missing symbols and
        # only when Yahoo has not already rate-limited the server.
        if not quote_rate_limited:
            for symbol in symbols:
                if symbol in data and data[symbol].get("mark") is not None:
                    continue
                px, src, err = _chart_quote(symbol)
                existing = data.get(symbol, {})
                existing.update({
                    "mark": px,
                    "last": px if px is not None else existing.get("last"),
                    "source": src,
                    "note": err,
                })
                data[symbol] = existing
        else:
            for symbol in symbols:
                existing = data.get(symbol, {})
                if existing.get("mark") is None:
                    existing.update({
                        "mark": None,
                        "source": "UNAVAILABLE",
                        "note": global_err or "Yahoo rate limited",
                    })
                data[symbol] = existing

    with _quote_lock:
        _quote_cache.update({
            "ts": now,
            "symbols": symbols,
            "data": dict(data),
            "error": global_err,
        })
    return data, global_err

def calculate_mark_to_market_return(positions: list[dict], nav_usd: float, quote_data: dict | None = None) -> dict:
    """Simple mark-to-market.

    Current Mark:
      Bid+Ask midpoint when both are available; otherwise Yahoo last.

    Change = Current Mark - Previous P&L Report Price
    Position P&L = Change * Quantity * Multiplier
    Daily Return = Total P&L / Baseline NAV
    """
    if quote_data is None:
        quote_data, yahoo_error = yahoo_current_marks(positions)
    else:
        yahoo_error = ""

    rows = []
    total_pnl = 0.0
    valid = 0

    for p in positions:
        symbol = _market_symbol(p)
        q = quote_data.get(symbol, {})

        # Backward-compatible test/input tuple support.
        if isinstance(q, (tuple, list)):
            current = q[0] if len(q) > 0 else None
            source = q[1] if len(q) > 1 else ""
            note = q[2] if len(q) > 2 else ""
            bid = ask = None
            last = current
        else:
            current = q.get("mark")
            bid = q.get("bid")
            ask = q.get("ask")
            last = q.get("last")
            source = q.get("source") or ""
            note = q.get("note") or ""

        previous = _positive(p.get("baseline_price"))
        change = pnl = contrib = None

        if current is not None and previous is not None:
            change = float(current) - float(previous)
            pnl = change * float(p.get("quantity") or 0.0) * float(p.get("multiplier") or 1.0)
            contrib = pnl / nav_usd * 100.0 if nav_usd else None
            total_pnl += pnl
            valid += 1

        rows.append({
            **p,
            "yahoo_symbol": symbol,
            "previous_mark": previous,
            "bid": bid,
            "ask": ask,
            "last": last,
            "current_mark": current,
            "change": change,
            "estimated_pnl": pnl,
            "contribution_pct": contrib,
            "price_source": source,
            "note": note,
        })

    total = len(positions)
    return {
        "rows": rows,
        "estimated_pnl": total_pnl,
        "estimated_return_pct": (total_pnl / nav_usd * 100.0) if nav_usd else None,
        "coverage_valid": valid,
        "coverage_total": total,
        "complete": bool(total and valid == total),
        "yahoo_error": yahoo_error,
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
