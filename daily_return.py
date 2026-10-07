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

# Yahoo's public option-chain quotes are delayed, so there is no value in
# re-requesting the full chain every page refresh. Keep a short shared cache.
_chain_cache = {"ts": 0.0, "signature": (), "data": {}, "error": ""}
CHAIN_CACHE_SECONDS = 120

_yfdata_obj = None

def _yfinance_quote_payload(symbols):
    """Authenticated Yahoo quote request via yfinance.

    Yahoo's /v7/finance/quote endpoint now requires a matching cookie+crumb.
    yfinance's YfData manages that session state for us, so Bid/Ask can be
    requested in one batch without manually handling Yahoo authentication.
    """
    global _yfdata_obj
    if not symbols:
        return {}, ""
    try:
        from yfinance.data import YfData
        if _yfdata_obj is None:
            _yfdata_obj = YfData()
        response = _yfdata_obj.get(
            url="https://query1.finance.yahoo.com/v7/finance/quote",
            params={"symbols": ",".join(symbols)},
            timeout=12,
        )
        payload = response.json()
        err = ((payload.get("finance") or {}).get("error") or {})
        if err:
            return {}, f"Yahoo: {err.get('code','Error')} — {err.get('description','')}".strip()
        return payload, ""
    except Exception as e:
        return {}, f"Yahoo authenticated quote error: {e}"


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
      Bid missing/zero + Ask available -> Ask / 2
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
            elif bid is None and ask is not None:
                mark = ask / 2.0
                source = "Yahoo Ask/2"
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



def _options_chain_current_marks(positions: list[dict]) -> tuple[dict, str]:
    """Read Bid/Ask/Last directly from Yahoo's option chain.

    Positions are grouped by (ticker, expiry). For each chain, the exact OCC
    contract symbol generated by _market_symbol() is matched against Yahoo's
    contractSymbol column. This is the same contract identifier displayed on
    the Yahoo Finance Options page.

    Mark priority inside the chain:
      - Bid > 0 and Ask > 0 -> midpoint
      - Bid missing/zero and Ask > 0 -> Ask / 2
      - otherwise Last
    """
    if not positions:
        return {}, ""

    signature = tuple(sorted(
        (
            str(p.get("ticker") or "").upper(),
            str(p.get("expiry") or ""),
            str(p.get("option_type") or "").upper(),
            float(p.get("strike") or 0),
            _market_symbol(p),
        )
        for p in positions
    ))
    now = time.time()

    with _quote_lock:
        if (
            _chain_cache["signature"] == signature
            and now - _chain_cache["ts"] < CHAIN_CACHE_SECONDS
        ):
            return dict(_chain_cache["data"]), _chain_cache["error"]

    try:
        import yfinance as yf
    except Exception as e:
        return {}, f"Yahoo Options Chain import error: {e}"

    grouped = {}
    wanted_by_group = {}
    for p in positions:
        ticker = str(p.get("ticker") or "").upper().strip()
        expiry = str(p.get("expiry") or "").strip()
        if not ticker or not expiry:
            continue
        key = (ticker, expiry)
        grouped.setdefault(key, []).append(p)
        wanted_by_group.setdefault(key, set()).add(_market_symbol(p).upper())

    out = {}
    errors = []

    for (ticker, expiry), group_positions in grouped.items():
        try:
            chain = yf.Ticker(ticker).option_chain(expiry)
            frames = []
            # Every current Fortress position is a put today, but support calls
            # as well so this remains generic.
            if getattr(chain, "calls", None) is not None:
                frames.append(("CALL", chain.calls))
            if getattr(chain, "puts", None) is not None:
                frames.append(("PUT", chain.puts))

            wanted = wanted_by_group[(ticker, expiry)]
            found = set()

            for side, frame in frames:
                if frame is None or getattr(frame, "empty", True):
                    continue

                for _, row in frame.iterrows():
                    contract = str(row.get("contractSymbol") or "").upper().strip()
                    if contract not in wanted:
                        continue

                    bid_raw = row.get("bid")
                    ask_raw = row.get("ask")
                    last_raw = row.get("lastPrice")

                    # Preserve a displayed Yahoo 0.00 Bid as zero for UI, while
                    # treating it as "no usable bid" for mark construction.
                    try:
                        bid_display = float(bid_raw) if bid_raw is not None and math.isfinite(float(bid_raw)) else None
                    except Exception:
                        bid_display = None
                    try:
                        ask_display = float(ask_raw) if ask_raw is not None and math.isfinite(float(ask_raw)) else None
                    except Exception:
                        ask_display = None
                    try:
                        last_display = float(last_raw) if last_raw is not None and math.isfinite(float(last_raw)) else None
                    except Exception:
                        last_display = None

                    bid = bid_display if bid_display is not None and bid_display > 0 else None
                    ask = ask_display if ask_display is not None and ask_display > 0 else None
                    last = last_display if last_display is not None and last_display >= 0 else None

                    if bid is not None and ask is not None and ask >= bid:
                        mark = (bid + ask) / 2.0
                        source = "Yahoo Options Chain Mid"
                    elif ask is not None and (bid is None):
                        mark = ask / 2.0
                        source = "Yahoo Options Chain Ask/2"
                    elif last is not None:
                        mark = last
                        source = "Yahoo Options Chain Last"
                    else:
                        mark = None
                        source = "UNAVAILABLE"

                    out[contract] = {
                        "mark": mark,
                        "bid": bid_display,
                        "ask": ask_display,
                        "last": last_display,
                        "source": source,
                        "note": "Yahoo Options Chain",
                    }
                    found.add(contract)

            missing = sorted(wanted - found)
            if missing:
                errors.append(
                    f"{ticker} {expiry}: exact contract not found: {', '.join(missing)}"
                )

        except Exception as e:
            errors.append(f"{ticker} {expiry}: {e}")

    err = " | ".join(errors)

    with _quote_lock:
        _chain_cache.update({
            "ts": now,
            "signature": signature,
            "data": dict(out),
            "error": err,
        })

    return out, err



def yahoo_current_marks(positions: list[dict]) -> tuple[dict, str]:
    """Fetch Yahoo option marks with Options Chain as the primary source.

    Primary:
      Yahoo Finance Options Chain, exact contractSymbol match.

    Chain mark rules:
      1. Bid + Ask -> midpoint
      2. Bid missing/zero + Ask -> Ask / 2
      3. Otherwise chain Last

    The older quote/spark/chart path is retained only as a technical fallback
    for a contract the chain request could not return.
    """
    symbols = tuple(sorted({_market_symbol(p) for p in positions}))
    now = time.time()

    with _quote_lock:
        if _quote_cache["symbols"] == symbols and now - _quote_cache["ts"] < CACHE_SECONDS:
            return dict(_quote_cache["data"]), _quote_cache["error"]

    # 1) PRIMARY: exact contract rows from Yahoo Options Chain.
    chain_data, chain_error = _options_chain_current_marks(positions)
    data = dict(chain_data)
    wanted = set(symbols)

    # 2) Only unresolved contracts use the older Yahoo quote path.
    missing = [s for s in symbols if s not in data or data[s].get("mark") is None]
    auth_error = ""
    fallback_error = ""

    if missing:
        payload, auth_error = _yfinance_quote_payload(tuple(missing))
        if payload:
            quote_data = _parse_quote_batch(payload, set(missing))
            for symbol, q in quote_data.items():
                if symbol not in data or data[symbol].get("mark") is None:
                    data[symbol] = q

        missing = [s for s in symbols if s not in data or data[s].get("mark") is None]

    if missing:
        spark_qs = urlencode({
            "symbols": ",".join(missing),
            "range": "1d",
            "interval": "1m",
            "indicators": "close",
            "includePrePost": "false",
        })
        for host in ("query1.finance.yahoo.com", "query2.finance.yahoo.com"):
            spark_payload, err = _http_json(f"https://{host}/v7/finance/spark?{spark_qs}")
            if err:
                fallback_error = err
                continue
            spark_data = _parse_spark(spark_payload or {}, set(missing))
            for symbol, tup in spark_data.items():
                px, src, note = tup
                if symbol in data and data[symbol].get("mark") is not None:
                    continue
                data[symbol] = {
                    "mark": px,
                    "bid": None,
                    "ask": None,
                    "last": px,
                    "source": src,
                    "note": note,
                }
            if spark_data:
                break

    missing = [s for s in symbols if s not in data or data[s].get("mark") is None]
    for symbol in missing:
        px, src, err = _chart_quote(symbol)
        data[symbol] = {
            "mark": px,
            "bid": data.get(symbol, {}).get("bid"),
            "ask": data.get(symbol, {}).get("ask"),
            "last": px if px is not None else data.get(symbol, {}).get("last"),
            "source": src,
            "note": err,
        }

    incomplete = any(data.get(s, {}).get("mark") is None for s in symbols)
    global_err = ""
    if incomplete:
        global_err = chain_error or auth_error or fallback_error

    with _quote_lock:
        _quote_cache.update({
            "ts": now,
            "symbols": symbols,
            "data": dict(data),
            "error": global_err,
        })
    return data, global_err


def calculate_mark_to_market_return(positions: list[dict], nav_usd: float, quote_data: dict | None = None) -> dict:
    """Simple mark-to-market with Yahoo manual override.

    Rule:
      Yahoo Bid/Ask Mid available -> use market midpoint.
      Bid missing/zero + Ask available -> use Ask / 2.
      Otherwise, if manual override exists -> use manual price.
      Otherwise -> use Yahoo fallback mark/last.
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

        if isinstance(q, (tuple, list)):
            market_mark = q[0] if len(q) > 0 else None
            market_source = q[1] if len(q) > 1 else ""
            note = q[2] if len(q) > 2 else ""
            bid = ask = None
            last = market_mark
        else:
            market_mark = q.get("mark")
            bid = q.get("bid")
            ask = q.get("ask")
            last = q.get("last")
            market_source = q.get("source") or ""
            note = q.get("note") or ""

        manual = p.get("yahoo_manual_price")
        if market_source in ("Yahoo Options Chain Mid", "Yahoo Bid/Ask Mid") and market_mark is not None:
            current = market_mark
            effective_source = "YAHOO MID"
        elif market_source in ("Yahoo Options Chain Ask/2", "Yahoo Ask/2") and market_mark is not None:
            current = market_mark
            effective_source = "YAHOO ASK/2"
        elif market_source == "Yahoo Options Chain Last" and market_mark is not None:
            current = market_mark
            effective_source = "YAHOO LAST"
        elif manual is not None:
            current = _positive(manual)
            effective_source = "MANUAL OVERRIDE"
        else:
            current = market_mark
            effective_source = "YAHOO LAST" if current is not None else "UNAVAILABLE"

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
            "market_mark": market_mark,
            "market_source": market_source,
            "manual_price": manual,
            "current_mark": current,
            "effective_source": effective_source,
            "change": change,
            "estimated_pnl": pnl,
            "contribution_pct": contrib,
            "price_source": effective_source,
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
