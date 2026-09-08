#!/usr/bin/env python3
"""
isx_collector.py - collect Iraq Stock Exchange (ISX) data into a growing record.

Each run:
  1. Scrapes the ISX homepage market summary (index level + market breadth).
  2. Finds the newest DAILY trading report (Excel) on the ISX "Market Reports"
     page, downloads it, and archives the raw file under data/reports/.
  3. Parses the per-company table into tidy rows.
  4. Appends to CSVs (deduplicated) and writes data/latest.json for a front-end.

Designed to be run repeatedly by any scheduler (GitHub Actions, cron, a VPS
systemd timer, etc.). The CSVs are the 24/7 record - they only grow.

ISX only produces new data during sessions (Sun-Thu, ~10:00-13:00 Baghdad time),
so running overnight/weekends is harmless but adds nothing; the dedupe keeps
duplicates out.

FIRST RUN: run locally and read the console. It prints the exact (Arabic)
column headers found in the daily Excel. Compare them to COLUMN_MAP below and
adjust if any didn't map. Once it parses cleanly, enable the scheduler.

Dependencies:  pip install requests beautifulsoup4 pandas openpyxl xlrd
"""

from __future__ import annotations
import json
import re
import datetime as dt
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit, quote, urljoin
from html import unescape

import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE     = "http://www.isx-iq.net/isxportal/portal/"
HOME_EN  = BASE + "homePage.html?currLanguage=en"
REPORTS  = BASE + "uploadedFilesList.html"

DATA        = Path("data")
REPORTS_DIR = DATA / "reports"
PRICES_CSV  = DATA / "isx_daily_prices.csv"
SUMMARY_CSV = DATA / "isx_market_summary.csv"
LATEST_JSON = DATA / "latest.json"
HISTORY_JSON = DATA / "history.json"

# Be a polite bot: identify yourself and keep the request rate low.
HEADERS = {"User-Agent": "ISX-personal-recorder/1.0 (research use; contact you@example.com)"}

# ---- Column mapping for the ISX daily "Bulletin" sheet ----------------------
# The per-company price table is on the sheet named "Bullient " (the exchange's
# own misspelling of "Bulletin"). Headers are English but spaced inconsistently,
# so we match on a whitespace-normalised, lower-cased version. Columns we don't
# need (average price, prev average price) are simply left unmapped.
FIELD_SYNONYMS = {
    "ticker":     ["code", "symbol"],
    "name":       ["company name", "company names"],
    "open":       ["opening price"],
    "high":       ["highest price"],
    "low":        ["lowest price"],
    "close":      ["closing price"],
    "prev_close": ["prev closing price", "previous closing price"],
    "change_pct": ["change (%)", "change(%)", "change %", "change"],
    "trades":     ["no.of trades", "no. of trades", "number of trades"],
    "volume":     ["traded volume", "traded shares"],
    "value_iqd":  ["traded value"],
}
NUMERIC = ["open", "high", "low", "close", "prev_close",
           "change_pct", "volume", "value_iqd", "trades"]
TICKER_RE = re.compile(r"^[A-Z]{2,6}$")


def _norm(v) -> str:
    """Whitespace-collapsed, lower-cased text for header matching."""
    s = "" if v is None else str(v)
    s = s.replace("\u00a0", " ")
    return re.sub(r"\s+", " ", s).strip().lower()


def _resolve_columns(header_cells) -> dict:
    """field -> column index, from normalised header text (first free match wins)."""
    norm = [_norm(c) for c in header_cells]
    col_of = {}
    for field, syns in FIELD_SYNONYMS.items():
        for i, h in enumerate(norm):
            if h in syns and i not in col_of.values():
                col_of[field] = i
                break
    return col_of


def _find_header_row(df):
    """Return (row_index, col_of) for the first price-table header row, else (None, None)."""
    for i in range(len(df)):
        cells = df.iloc[i].tolist()
        norm = [_norm(c) for c in cells]
        if "code" in norm and ("company name" in norm or "company names" in norm):
            col_of = _resolve_columns(cells)
            if "ticker" in col_of and "close" in col_of:
                return i, col_of
    return None, None


def _extract_sheet(df) -> list:
    """Company rows from a price-table sheet. Skips title/sector/total rows and
    tolerates a second header partway down (Regular + Second Platform markets)."""
    hdr, col_of = _find_header_row(df)
    if hdr is None:
        return []
    tcol = col_of["ticker"]
    out = []
    for i in range(hdr + 1, len(df)):
        row = df.iloc[i].tolist()
        code = str(row[tcol]).strip() if tcol < len(row) else ""
        if not TICKER_RE.match(code):     # real tickers are 2-6 capitals; skips headers/sectors/totals
            continue
        rec = {"ticker": code}
        for field, ci in col_of.items():
            if field != "ticker" and ci < len(row):
                rec[field] = row[ci]
        out.append(rec)
    return out


def session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


def today() -> str:
    return dt.date.today().isoformat()


def _encode(url: str) -> str:
    """Percent-encode spaces / Arabic in the path so requests can fetch it."""
    p = urlsplit(url)
    return urlunsplit((p.scheme, p.netloc, quote(p.path, safe="/%"), p.query, p.fragment))


# ---------- 1. homepage market summary ----------
def fetch_market_summary(s: requests.Session) -> dict:
    r = s.get(HOME_EN, timeout=30)
    r.raise_for_status()
    text = BeautifulSoup(r.text, "html.parser").get_text("\n")

    def after(label: str):
        m = re.search(re.escape(label) + r"\s*\n\s*([-\d,\.]+%?)", text)
        return m.group(1).replace(",", "") if m else None

    return {
        "date": today(),
        "main_index":     after("Main Index"),
        "change":         after("Change"),
        "change_pct":     after("Change %"),
        "value_traded":   after("Value Traded"),
        "shares_traded":  after("Shares Traded"),
        "trades":         after("Trades"),
        "symbols_traded": after("Symbols Traded"),
        "symbols_up":     after("Symbols Up"),
        "symbols_down":   after("Symbols Down"),
        "flat":           after("Flat"),
    }


# ---------- 2. find + download newest daily Excel ----------
def find_latest_daily_excel(s: requests.Session):
    """Return (date_iso, absolute_url) for the newest DAILY Excel report, or None.

    Robust to the portal's quirks: forces UTF-8 (the site is UTF-8 but often
    omits the charset header, which otherwise mojibakes the Arabic), un-escapes
    any &#NNNN; entities, reads the raw <tr> HTML, strips the ';jsessionid=...'
    the Java portal appends, and resolves relative hrefs to absolute URLs.
    """
    r = s.get(REPORTS, timeout=30)
    r.raise_for_status()
    r.encoding = "utf-8"
    page = r.text

    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", page, re.S | re.I)
    excel_rows = daily_rows = 0
    sample = ""
    best = None  # (date_iso, url)
    for row in rows:
        m_file = re.search(r"""href=["']([^"']*\.(?:xlsx|xls))(?:;jsessionid=[^"']*)?["']""", row, re.I)
        text = unescape(row)               # turn &#1610;… into real Arabic if entity-encoded
        if m_file:
            excel_rows += 1
            if not sample:
                sample = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", text)).strip()[:200]
        is_daily = ("يومي" in text) or (re.search(r"\bdaily\b", text, re.I) is not None)
        if not is_daily:
            continue
        daily_rows += 1
        if not m_file:                     # e.g. the daily news PDF row
            continue
        href = urljoin(REPORTS, m_file.group(1))
        m_date = re.search(r"(\d{2})/(\d{2})/(\d{4})", text)
        d = f"{m_date.group(3)}-{m_date.group(2)}-{m_date.group(1)}" if m_date else ""
        if best is None or d > best[0]:
            best = (d, href)

    if best is None:
        print(f"(debug) <tr> rows={len(rows)}, rows with an Excel link={excel_rows}, "
              f"rows tagged daily={daily_rows}")
        if sample:
            print(f"(debug) first Excel row text: {sample!r}")
    return best


def download(s: requests.Session, url: str, dest: Path) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    r = s.get(_encode(url), timeout=60)
    r.raise_for_status()
    dest.write_bytes(r.content)
    return dest


# ---------- 3. parse the daily Excel ----------
def parse_daily_excel(path: Path, report_date: str) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    sheets = xls.sheet_names

    # Main equity price table lives on the "Bullient " sheet (misspelled Bulletin).
    main = next((s for s in sheets if "bull" in s.lower()), None)
    records = []
    if main is not None:
        records += _extract_sheet(pd.read_excel(path, sheet_name=main, header=None, dtype=str))

    # Companies that didn't trade: capture their carried-over closing price too,
    # so every listed company gets a row each day (volume/trades = 0).
    traded = {r["ticker"] for r in records}
    nt = next((s for s in sheets if "not trading" in s.lower() and "bond" not in s.lower()), None)
    if nt is not None:
        for r in _extract_sheet(pd.read_excel(path, sheet_name=nt, header=None, dtype=str)):
            if r["ticker"] not in traded:
                r.setdefault("volume", "0")
                r.setdefault("trades", "0")
                records.append(r)

    if not records:
        print("!! No company rows found. Sheets in file:", sheets)
        raise SystemExit("Could not parse the price table — paste the sheet names above.")

    df = pd.DataFrame.from_records(records)
    cols = [c for c in (["ticker", "name"] + NUMERIC) if c in df.columns]
    df = df[cols]
    for c in NUMERIC:
        if c in df.columns:
            df[c] = pd.to_numeric(
                df[c].astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False),
                errors="coerce",
            )
    df = df.drop_duplicates(subset=["ticker"], keep="first").reset_index(drop=True)
    df.insert(0, "date", report_date)
    return df


# ---------- 4. persist ----------
def append_csv(df: pd.DataFrame, path: Path, keys: list[str]) -> pd.DataFrame:
    path.parent.mkdir(parents=True, exist_ok=True)
    new = df.astype(str)
    if path.exists():
        old = pd.read_csv(path, dtype=str)
        both = pd.concat([old, new], ignore_index=True).drop_duplicates(subset=keys, keep="last")
    else:
        both = new
    both.to_csv(path, index=False)
    return both


def write_latest_json(prices: pd.DataFrame) -> None:
    latest = {}
    if {"ticker", "date"}.issubset(prices.columns):
        p = prices.copy()
        p["date"] = p["date"].astype(str)
        for tk, g in p.groupby("ticker"):
            row = g.sort_values("date").iloc[-1].to_dict()
            clean = {}
            for k, v in row.items():
                if v is None:
                    continue
                s = str(v).strip()
                if s == "" or s.lower() in ("nan", "none", "<na>"):
                    continue                       # skip blanks so the JSON stays valid
                clean[k] = s
            latest[str(tk)] = clean
    LATEST_JSON.write_text(json.dumps(latest, ensure_ascii=False, indent=2))


def _clean_num(v):
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "<na>"):
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def _merge_series(existing: dict, ticker: str, series: list) -> None:
    old = {d: v for d, v in existing.get(ticker, []) if isinstance(d, str)}
    for d, v in series:
        old[d] = v
    existing[ticker] = [[d, old[d]] for d in sorted(old)]


def write_history_json(prices: pd.DataFrame, summary_path: Path) -> None:
    """Merge per-ticker close history + the ISX60 index INTO the existing
    history.json (so a one-time backfill is preserved and daily runs just
    extend it). Shape: { "TICKER": [["YYYY-MM-DD", close], ...], "__ISX60": [...] }"""
    hist = {}
    if HISTORY_JSON.exists():
        try:
            hist = json.loads(HISTORY_JSON.read_text())
        except Exception:
            hist = {}
    if {"ticker", "date", "close"}.issubset(prices.columns):
        for tk, g in prices.groupby("ticker"):
            series = []
            for _, r in g.sort_values("date").iterrows():
                v = _clean_num(r.get("close"))
                if v is not None:
                    series.append([str(r["date"]), v])
            if series:
                _merge_series(hist, str(tk), series)
    # ISX60 index history from the market-summary log
    try:
        if summary_path.exists():
            s = pd.read_csv(summary_path, dtype=str)
            if {"date", "main_index"}.issubset(s.columns):
                idx = []
                for _, r in s.sort_values("date").iterrows():
                    v = _clean_num(r.get("main_index"))
                    if v is not None:
                        idx.append([str(r["date"]), v])
                if idx:
                    _merge_series(hist, "__ISX60", idx)
    except Exception as e:
        print("index history skipped:", e)
    HISTORY_JSON.write_text(json.dumps(hist, ensure_ascii=False))


def main() -> None:
    DATA.mkdir(exist_ok=True)
    s = session()

    # index-level summary (best effort - never fatal)
    try:
        summ = fetch_market_summary(s)
        append_csv(pd.DataFrame([summ]), SUMMARY_CSV, keys=["date"])
        print("summary:", summ.get("main_index"), summ.get("change_pct"))
    except Exception as e:
        print("summary failed:", e)

    # per-company daily prices
    found = find_latest_daily_excel(s)
    if not found:
        print("No daily Excel report found on the reports page.")
        return
    report_date, url = found
    report_date = report_date or today()
    dest = REPORTS_DIR / f"ISX_daily_{report_date}.xlsx"
    download(s, url, dest)
    df = parse_daily_excel(dest, report_date)
    print(f"parsed {len(df)} companies for {report_date}")

    prices = append_csv(df, PRICES_CSV, keys=["date", "ticker"])
    write_latest_json(prices)
    write_history_json(prices, SUMMARY_CSV)
    print("wrote", PRICES_CSV, "and", LATEST_JSON)


if __name__ == "__main__":
    main()
