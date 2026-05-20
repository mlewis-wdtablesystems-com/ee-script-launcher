"""
BOM unit costing from a CSV file: **Nexar Supply first**, then **web scraping** for gaps.

- Does **not** use Arena or any other PLM/ERP API (no Arena prices, BOM, or sourcing).

Pricing order per line:
  1. Nexar Supply (`supSearchMpn`) when `NEXAR_CLIENT_ID` / `NEXAR_CLIENT_SECRET` are set.
  2. If Nexar returns no usable unit price, optional **distributor web pages** (same VPN / stock
     number), controlled by `BOM_WEB_SCRAPE` (default 1).

Requires (recommended):
  NEXAR_CLIENT_ID, NEXAR_CLIENT_SECRET   (OAuth client_credentials, scope: supply.domain)

Optional:
  NEXAR_COUNTRY, NEXAR_CURRENCY, NEXAR_USD_TO_AUD, NEXAR_UNMATCHED_USE_BEST, NEXAR_REQUEST_DELAY_SEC
  BOM_WEB_SCRAPE          default 1  (set 0 to disable scraping)
  BOM_REQUEST_DELAY_SEC   default 0.18  (pause after each line’s network calls)
  BOM_INPUT_CSV / BOM_OUTPUT_CSV / BOM_OUTPUT_XLSX
  --no-scrape             disable web scraping for this run

Usage:
  python arena_bom_cost_tool.py
  python arena_bom_cost_tool.py C:\\path\\to\\bom_lines.csv
  python arena_bom_cost_tool.py bom.csv --no-scrape
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = SCRIPT_DIR / "263-70-001_BOM_FULL_cost_export.csv"
DEFAULT_OUT_CSV = SCRIPT_DIR / "263-70-001_BOM_NEXAR_cost_export.csv"
DEFAULT_OUT_XLSX = SCRIPT_DIR / "263-70-001_BOM_NEXAR_cost_export.xlsx"

NEXAR_TOKEN_URL = "https://identity.nexar.com/connect/token"
NEXAR_GRAPHQL_URL = "https://api.nexar.com/graphql"

CHROME_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)

GST = Decimal("1.1")


def d2(x: Decimal | float | int | None) -> Decimal:
    if x is None:
        return Decimal("0")
    if isinstance(x, Decimal):
        return x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def aud_ex_gst_from_inc(inc: Decimal) -> Decimal:
    return (inc / GST).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)


def fetch_url(url: str, timeout: int = 25) -> str:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": CHROME_UA,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-AU,en;q=0.9",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read().decode("utf-8", "replace")


def nexar_credentials_configured() -> bool:
    return bool(os.environ.get("NEXAR_CLIENT_ID", "").strip()) and bool(
        os.environ.get("NEXAR_CLIENT_SECRET", "").strip()
    )


_nexar_token: str | None = None
_nexar_token_expiry: float = 0.0


def _nexar_get_token() -> str:
    global _nexar_token, _nexar_token_expiry
    cid = os.environ.get("NEXAR_CLIENT_ID", "").strip()
    csec = os.environ.get("NEXAR_CLIENT_SECRET", "").strip()
    if not cid or not csec:
        raise RuntimeError("Set NEXAR_CLIENT_ID and NEXAR_CLIENT_SECRET (scope supply.domain).")

    now = time.time()
    if _nexar_token and now < _nexar_token_expiry - 60:
        return _nexar_token

    body = urllib.parse.urlencode(
        {
            "grant_type": "client_credentials",
            "client_id": cid,
            "client_secret": csec,
            "scope": "supply.domain",
        }
    ).encode()
    req = urllib.request.Request(
        NEXAR_TOKEN_URL,
        data=body,
        method="POST",
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
            "User-Agent": CHROME_UA,
        },
    )
    with urllib.request.urlopen(req, timeout=45) as resp:
        data = json.loads(resp.read().decode())
    tok = data.get("access_token")
    if not tok:
        raise RuntimeError(f"Nexar token response missing access_token: {data!r}")
    _nexar_token = tok
    _nexar_token_expiry = now + float(data.get("expires_in", 3600))
    return _nexar_token


def _norm_supplier(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def _bom_vendor_is_rs(bom_vendor: str) -> bool:
    s = (bom_vendor or "").lower().strip()
    if s == "rs" or s.startswith("rs ") or "rs components" in s or "rs online" in s:
        return True
    n = _norm_supplier(bom_vendor)
    return n in {"rs", "rscomponents", "rscomponentsplc"}


def _nexar_supplier_matches(bom_vendor: str, nexar_company: str) -> bool:
    a = _norm_supplier(bom_vendor)
    b = _norm_supplier(nexar_company)
    if not a or not b:
        return False
    if a in b or b in a:
        return True
    if "digi" in a and "digi" in b:
        return True
    if "mouser" in a and "mouser" in b:
        return True
    if ("farnell" in a or "element" in a or "newark" in a) and (
        "farnell" in b or "element" in b or "newark" in b or "premier" in b
    ):
        return True
    if _bom_vendor_is_rs(bom_vendor) and (
        "rscomponents" in b or "rsonline" in b or b == "rs" or "distrelec" in b
    ):
        return True
    return False


def _decimal_from(v: Any) -> Decimal | None:
    if v is None:
        return None
    try:
        return Decimal(str(v))
    except Exception:
        return None


def _pick_price_break(prices: list[dict[str, Any]], qty_q: Decimal) -> tuple[Decimal | None, str]:
    usd_to_aud = Decimal(os.environ.get("NEXAR_USD_TO_AUD", "1.55"))
    rows: list[tuple[Decimal, dict[str, Any]]] = []
    for p in prices:
        qn = p.get("quantity")
        try:
            qd = Decimal(str(qn)) if qn is not None else Decimal(0)
        except Exception:
            qd = Decimal(0)
        rows.append((qd, p))
    rows.sort(key=lambda x: x[0])
    chosen: dict[str, Any] | None = None
    for qd, prow in rows:
        if qd <= qty_q:
            chosen = prow
    if chosen is None and rows:
        chosen = rows[0][1]
    if not chosen:
        return None, "Nexar: no price breaks"

    conv = _decimal_from(chosen.get("convertedPrice"))
    conv_cur = (chosen.get("convertedCurrency") or "").upper()
    if conv is not None and conv_cur == "AUD":
        return conv.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP), "Nexar (converted AUD)"

    raw = _decimal_from(chosen.get("price"))
    cur = (chosen.get("currency") or "").upper()
    if raw is not None and cur == "AUD":
        return raw.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP), "Nexar (AUD list)"

    if raw is not None and cur == "USD":
        aud = (raw * usd_to_aud).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
        return aud, f"Nexar (USD→AUD @ {usd_to_aud})"

    if conv is not None and conv_cur == "USD":
        aud = (conv * usd_to_aud).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
        return aud, f"Nexar (converted USD→AUD @ {usd_to_aud})"

    return None, "Nexar: could not interpret currency on price break"


def _best_unit_from_seller_offers(
    seller_name: str, offers: list[dict[str, Any]], qty_q: Decimal
) -> tuple[Decimal | None, str]:
    best: tuple[Decimal, str] | None = None
    for off in offers or []:
        plist = off.get("prices") or []
        if not plist:
            continue
        unit, note = _pick_price_break(plist, qty_q)
        if unit is None or unit <= 0:
            continue
        cand = (unit, f"{note}; seller={seller_name}")
        if best is None or unit < best[0]:
            best = cand
    return (best[0], best[1]) if best else (None, "")


def nexar_sup_search_part(vpn: str) -> dict[str, Any] | None:
    tok = _nexar_get_token()
    qvpn = vpn.strip()
    if not qvpn:
        return None

    cur = os.environ.get("NEXAR_CURRENCY", "AUD").strip().upper() or "AUD"
    cc = os.environ.get("NEXAR_COUNTRY", "AU").strip().upper() or "AU"

    query = """
query BomPrice($q: String!, $cur: String!, $cc: String!) {
  supSearchMpn(q: $q, limit: 5, currency: $cur, country: $cc) {
    hits
    results {
      part {
        mpn
        sellers {
          company { name }
          offers {
            prices { quantity price currency convertedPrice convertedCurrency }
          }
        }
      }
    }
  }
}
"""
    payload = json.dumps(
        {"query": query, "variables": {"q": qvpn, "cur": cur, "cc": cc}}
    ).encode()
    req = urllib.request.Request(
        NEXAR_GRAPHQL_URL,
        data=payload,
        method="POST",
        headers={
            "Authorization": f"Bearer {tok}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": CHROME_UA,
        },
    )
    with urllib.request.urlopen(req, timeout=45) as resp:
        body = json.loads(resp.read().decode())

    if body.get("errors"):
        raise RuntimeError(f"Nexar GraphQL errors: {body['errors'][:2]}")

    data = body.get("data") or {}
    root = data.get("supSearchMpn") or {}
    results = root.get("results") or []
    if not results:
        return None
    part = results[0].get("part") or {}
    return part if part else None


def price_line_nexar_only(vendor_name: str, vendor_pn: str, qty_q: Decimal) -> tuple[Decimal | None, str]:
    part = nexar_sup_search_part(vendor_pn)
    if not part:
        return None, "Nexar: no search results for VPN"

    sellers = part.get("sellers") or []
    best_matched: tuple[Decimal, str] | None = None

    for sl in sellers:
        comp = (sl.get("company") or {}).get("name") or ""
        if not _nexar_supplier_matches(vendor_name, comp):
            continue
        unit, note = _best_unit_from_seller_offers(comp, sl.get("offers") or [], qty_q)
        if unit is not None and unit > 0:
            cand = (unit, note)
            if best_matched is None or unit < best_matched[0]:
                best_matched = cand

    if best_matched:
        return best_matched[0], best_matched[1]

    use_best = os.environ.get("NEXAR_UNMATCHED_USE_BEST", "1").strip() not in ("0", "false", "no")
    if not use_best:
        return None, "Nexar: no seller matched BOM vendor (NEXAR_UNMATCHED_USE_BEST=0)"

    best_any: tuple[Decimal, str] | None = None
    for sl in sellers:
        comp = (sl.get("company") or {}).get("name") or ""
        unit, note = _best_unit_from_seller_offers(comp, sl.get("offers") or [], qty_q)
        if unit is not None and unit > 0:
            cand = (unit, f"{note} [unmatched BOM vendor '{vendor_name}']")
            if best_any is None or unit < best_any[0]:
                best_any = cand

    if best_any:
        return best_any[0], best_any[1]
    return None, "Nexar: no priced offers for this part"


def _rs_pack_unit_ex_gst(html: str) -> tuple[Decimal | None, str]:
    m = re.search(
        r"Subtotal\s*\(\s*1\s*(?:pack|bag|reel)\s+of\s+(\d+)\s+units?\s*\)[^$]*\$([0-9,]+\.?[0-9]*)",
        html,
        re.I | re.S,
    )
    if not m:
        return None, "no pack subtotal"
    n = Decimal(m.group(1))
    pack_aud = Decimal(m.group(2).replace(",", ""))
    span_start = m.start()
    span = html[span_start : span_start + 500].lower()
    if "(exc." in span and "gst" in span:
        unit_ex = pack_aud / n
        return unit_ex.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP), "RS AU pack (exc-GST)"
    if "(inc." in span and "gst" in span:
        unit_ex = aud_ex_gst_from_inc(pack_aud / n)
        return unit_ex, "RS AU pack (inc-GST→ex-GST)"
    unit_ex = pack_aud / n
    return unit_ex.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP), "RS AU pack (tax unknown→treated exc-GST)"


def price_rs_au_stock(stock_no: str, qty_q: Decimal) -> tuple[Decimal | None, str]:
    stock = stock_no.strip().replace(" ", "")
    if not re.fullmatch(r"\d{3}-\d{2,5}", stock):
        return None, "RS: stock number format unexpected"
    search = f"https://au.rs-online.com/web/c/?searchTerm={urllib.parse.quote(stock)}"
    try:
        html = fetch_url(search, timeout=30)
    except Exception as e:
        return None, f"RS web: {e}"

    pack = _rs_pack_unit_ex_gst(html)
    if pack[0] is not None:
        return pack

    tiers = re.findall(
        r"(\d+)\s*-\s*(\d+)\s*\|\s*\$([0-9,]+\.?[0-9]*)\s*\|",
        html,
    )
    if tiers:
        qty_int = int(qty_q.to_integral_value(rounding=ROUND_HALF_UP))
        for lo_s, hi_s, aud_s in tiers:
            lo, hi = int(lo_s), int(hi_s)
            unit_aud = Decimal(aud_s.replace(",", ""))
            if lo <= qty_int <= hi:
                return unit_aud.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP), (
                    f"RS AU tier {lo}-{hi} ea (exc-GST)"
                )
        best_lo = -1
        picked: Decimal | None = None
        for lo_s, hi_s, aud_s in tiers:
            lo = int(lo_s)
            if lo <= qty_int and lo >= best_lo:
                best_lo = lo
                picked = Decimal(aud_s.replace(",", ""))
        if picked is not None:
            return picked.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP), "RS AU tier (MOQ≤Q)"
        first = Decimal(tiers[0][2].replace(",", ""))
        return first.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP), "RS AU tier (fallback)"
    return None, "RS: could not parse price"


def price_digikey_au(keyword: str) -> tuple[Decimal | None, str]:
    q = urllib.parse.quote(keyword.strip())
    url = f"https://www.digikey.com.au/en/products/result?keywords={q}"
    try:
        html = fetch_url(url, timeout=25)
    except Exception as e:
        return None, f"Digi-Key web: {e}"
    m3 = re.search(r"1\s*</td>\s*<td[^>]*>\s*AU\$\s*([0-9,]+\.?[0-9]*)", html, re.I)
    if m3:
        return Decimal(m3.group(1).replace(",", "")), "Digi-Key AU (list)"
    m4 = re.search(r"AU\$\s*([0-9,]+\.?[0-9]*)\s*</", html, re.I)
    if m4:
        return Decimal(m4.group(1).replace(",", "")), "Digi-Key AU (heuristic)"
    return None, "Digi-Key: no AUD price in HTML"


def price_element14_au(part: str) -> tuple[Decimal | None, str]:
    q = urllib.parse.quote(part.strip())
    url = f"https://au.element14.com/search?st={q}"
    try:
        html = fetch_url(url, timeout=30)
    except Exception as e:
        return None, f"element14 web: {e}"
    m = re.search(r"productDisplayPrice[^>]*>\s*AU\$\s*([0-9,]+\.?[0-9]*)", html, re.I)
    if m:
        inc = Decimal(m.group(1).replace(",", ""))
        return aud_ex_gst_from_inc(inc), "element14 AU (inc-GST→ex-GST est.)"
    m2 = re.search(r"\"prices\"\s*:\s*\{[^}]*\"AUD\"\s*:\s*([0-9.]+)", html)
    if m2:
        return Decimal(m2.group(1)), "element14 AUD JSON"
    m3 = re.search(r"AU\$\s*([0-9,]+\.?[0-9]*)", html)
    if m3:
        return aud_ex_gst_from_inc(Decimal(m3.group(1).replace(",", ""))), "element14 AU (generic AU$)"
    return None, "element14: no price parsed"


def price_mouser_au(keyword: str) -> tuple[Decimal | None, str]:
    q = urllib.parse.quote(keyword.strip())
    url = f"https://au.mouser.com/c/?q={q}"
    try:
        html = fetch_url(url, timeout=30)
    except Exception as e:
        return None, f"Mouser web: {e}"
    if "denied" in html.lower() and len(html) < 25000:
        return None, "Mouser: access denied (bot protection)"
    m = re.search(r"AU\$\s*([0-9,]+\.?[0-9]*)", html, re.I)
    if m:
        return Decimal(m.group(1).replace(",", "")), "Mouser AU (heuristic)"
    m2 = re.search(r"\"price\"\s*:\s*\"([0-9.]+)\"\s*,\s*\"currency\"\s*:\s*\"AUD\"", html, re.I)
    if m2:
        return Decimal(m2.group(1)), "Mouser AU (JSON)"
    return None, "Mouser: no AUD price parsed"


def price_wagner(sku: str) -> tuple[Decimal | None, str]:
    sku_u = sku.strip().upper()
    if sku_u == "SMP12V12A-25P":
        return Decimal("50.700"), "Wagner SMP12V12A-25P (static ex-GST proxy; verify)"
    q = urllib.parse.quote(sku.strip())
    url = f"https://www.wagneronline.com.au/search?q={q}"
    try:
        html = fetch_url(url, timeout=25)
    except Exception as e:
        return None, f"Wagner web: {e}"
    m = re.search(r"price[^$]*\$\s*([0-9,]+\.?[0-9]*)", html, re.I)
    if m:
        inc = Decimal(m.group(1).replace(",", ""))
        if inc <= 0:
            return None, "Wagner: non-positive price"
        return aud_ex_gst_from_inc(inc), "Wagner AU (inc-GST→ex-GST est.)"
    return None, "Wagner: no price parsed"


def scrape_vendor_unit(vendor_name: str, vendor_pn: str, qty_q: Decimal) -> tuple[Decimal | None, str]:
    """Distributor web pages only — no Arena."""
    sl = vendor_name.lower()
    vpn = vendor_pn.strip()
    if not vpn:
        return None, "empty VPN"

    if _bom_vendor_is_rs(vendor_name) or sl == "rs" or sl.startswith("rs "):
        if re.fullmatch(r"\d{3}-\d{2,5}", vpn.replace(" ", "")):
            return price_rs_au_stock(vpn, qty_q)
        return None, "RS: VPN not RS stock number"

    if "digi" in sl or "digi-key" in sl:
        return price_digikey_au(vpn)

    if "mouser" in sl:
        return price_mouser_au(vpn)

    if "element" in sl or "farnell" in sl or "newark" in sl:
        return price_element14_au(vpn)

    if "wagner" in sl:
        return price_wagner(vpn)

    return None, f"no web scraper for vendor '{vendor_name}'"


def price_line(
    vendor_name: str,
    vendor_pn: str,
    qty_q: Decimal,
    *,
    nexar_ready: bool,
    scrape: bool,
) -> tuple[Decimal | None, str]:
    notes: list[str] = []
    unit: Decimal | None = None

    if nexar_ready:
        try:
            u, n = price_line_nexar_only(vendor_name, vendor_pn, qty_q)
            if u is not None and u > 0:
                return u, n
            if n:
                notes.append(n)
        except Exception as e:
            notes.append(f"Nexar: {e}")

    if scrape and os.environ.get("BOM_WEB_SCRAPE", "1").strip().lower() not in ("0", "false", "no"):
        u2, n2 = scrape_vendor_unit(vendor_name, vendor_pn, qty_q)
        if u2 is not None and u2 > 0:
            suffix = f"Web scrape: {n2}"
            if notes:
                return u2, " | ".join(notes) + f" | {suffix}"
            return u2, suffix
        if n2:
            notes.append(f"Web: {n2}")

    return None, " | ".join(notes) if notes else "No price (Nexar + web scrape exhausted)"


def write_outputs(
    rows: list[dict[str, Any]],
    grand: Decimal,
    out_csv: Path,
    out_xlsx: Path,
) -> tuple[Path, Path]:
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter

    csv_grand = sum((r["qty_i"] * r["unit_r"]) for r in rows if r.get("unit_r") is not None)
    csv_grand = d2(csv_grand)

    COL_M, COL_N, COL_I, COL_Q, COL_R, COL_S = 13, 14, 9, 17, 18, 19

    wb = Workbook()
    ws = wb.active
    ws.title = "BOM Cost"
    ws.cell(1, COL_M, "BOM line-total sum (S1) AUD — Nexar + optional web scrape")
    ws.cell(2, COL_M, "Vendor (M)")
    ws.cell(2, COL_N, "Vendor item (N)")
    ws.cell(2, COL_I, "Qty (I)")
    ws.cell(2, COL_Q, "Price lookup qty (Q)")
    ws.cell(2, COL_R, "Unit cost (R)")
    ws.cell(2, COL_S, "Line total (S)")
    ws.cell(2, 20, "Internal PN")
    ws.cell(2, 21, "Item name")
    ws.cell(2, 22, "Pricing notes")

    start = 3
    last = start + len(rows) - 1
    for i, r in enumerate(rows, start=start):
        ws.cell(i, COL_M, r["vendor"])
        ws.cell(i, COL_N, r["vpn"])
        ws.cell(i, COL_I, float(r["qty_i"]))
        ws.cell(i, COL_Q, float(r["qty_q"]))
        if r["unit_r"] is not None:
            ws.cell(i, COL_R, float(r["unit_r"]))
        ci, cr, cs = get_column_letter(COL_I), get_column_letter(COL_R), get_column_letter(COL_S)
        ws.cell(i, COL_S, f"=IF(ISBLANK({cr}{i}),0,{ci}{i}*{cr}{i})")
        ws.cell(i, 20, r["internal_pn"])
        ws.cell(i, 21, r["name"])
        ws.cell(i, 22, r["note"])

    cs = get_column_letter(COL_S)
    ws.cell(1, COL_S, f"=SUM({cs}{start}:{cs}{last})")
    ws.freeze_panes = "A3"

    out_csv.parent.mkdir(parents=True, exist_ok=True)
    with out_csv.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(
            [
                "Vendor_M",
                "VendorItem_N",
                "Qty_I",
                "Qty_Q",
                "UnitCost_R_AUD_exGST",
                "LineTotal_S",
                "Internal_PN",
                "Name",
                "Notes",
            ]
        )
        w.writerow(
            [
                "",
                "",
                "",
                "",
                "",
                float(csv_grand),
                "",
                "",
                "S cumulative (rows with R filled; xlsx S1 is SUM formula)",
            ]
        )
        for r in rows:
            ur = r["unit_r"]
            lt = (r["qty_i"] * ur) if ur is not None else None
            w.writerow(
                [
                    r["vendor"],
                    r["vpn"],
                    float(r["qty_i"]),
                    float(r["qty_q"]),
                    float(ur) if ur is not None else "",
                    float(lt) if lt is not None else "",
                    r["internal_pn"],
                    r["name"],
                    r["note"],
                ]
            )

    try:
        wb.save(out_xlsx)
        xlsx_written = out_xlsx
    except PermissionError:
        xlsx_written = out_xlsx.with_name(out_xlsx.stem + "_alt.xlsx")
        wb.save(xlsx_written)

    return out_csv, xlsx_written


def _clean_arena_part_number(cell: str) -> str:
    s = (cell or "").strip()
    m = re.match(r"^(.+?)\s+rev\b", s, re.I)
    return m.group(1).strip() if m else s


def load_bom_rows_from_arena_export(lines: list[str], header_idx: int) -> list[dict[str, Any]]:
    """Arena / PLM CSV: preamble rows, then header with Item Number, Vendor, Vendor Item Number, Qty 100."""
    buf = io.StringIO("\n".join(lines[header_idx:]))
    rdr = csv.reader(buf)
    hdr = next(rdr)

    def col_index(name: str) -> int | None:
        for j, h in enumerate(hdr):
            if (h or "").strip() == name:
                return j
        return None

    j_v = col_index("Vendor")
    j_vpn = col_index("Vendor Item Number")
    j_q = col_index("Quantity")
    j_q100 = col_index("Qty 100")
    j_item = col_index("Item Number")
    j_name = col_index("Item Name")
    if j_v is None or j_vpn is None or j_q is None:
        raise RuntimeError(f"Arena-style export: missing columns in header {hdr!r}")

    rows_out: list[dict[str, Any]] = []
    for row in rdr:

        def cell(j: int | None) -> str:
            if j is None or j >= len(row):
                return ""
            return (row[j] or "").strip()

        vendor = cell(j_v)
        vpn = cell(j_vpn)
        if not vendor or not vpn:
            continue
        try:
            qty_i = Decimal(cell(j_q))
        except Exception:
            continue
        if qty_i <= 0:
            continue
        if j_q100 is not None and cell(j_q100):
            try:
                qty_q = Decimal(cell(j_q100))
            except Exception:
                qty_q = qty_i
        else:
            qty_q = qty_i

        internal_pn = _clean_arena_part_number(cell(j_item)) if j_item is not None else ""
        name = cell(j_name) if j_name is not None else ""
        rows_out.append(
            {
                "vendor": vendor,
                "vpn": vpn,
                "qty_i": qty_i,
                "qty_q": qty_q,
                "internal_pn": internal_pn,
                "name": name,
                "unit_r": None,
                "note": "",
            }
        )
    return rows_out


def load_bom_rows_from_csv(path: Path) -> list[dict[str, Any]]:
    raw = path.read_text(encoding="utf-8-sig")
    lines = raw.splitlines()
    for i, line in enumerate(lines):
        if (
            "Vendor Item Number" in line
            and "Item Number" in line
            and "Quantity" in line
            and "Vendor" in line
        ):
            return load_bom_rows_from_arena_export(lines, i)

    rows_out: list[dict[str, Any]] = []
    with path.open(newline="", encoding="utf-8-sig") as f:
        rdr = csv.DictReader(f)
        fieldnames = [x.strip() for x in (rdr.fieldnames or [])]
        if not fieldnames:
            raise RuntimeError(f"Empty or invalid CSV: {path}")

        def col(*names: str) -> str | None:
            for n in names:
                if n in (rdr.fieldnames or []):
                    return n
            for fn in rdr.fieldnames or []:
                for n in names:
                    if fn.strip().lower() == n.lower():
                        return fn
            return None

        c_vm = col("Vendor_M", "Vendor (M)", "M")
        c_vn = col("VendorItem_N", "Vendor item (N)", "N")
        c_qi = col("Qty_I", "Qty (I)", "I")
        c_qq = col("Qty_Q", "Qty (Q)", "Q")
        c_ip = col("Internal_PN", "Internal PN")
        c_nm = col("Name")
        if not c_vm or not c_vn or not c_qi:
            raise RuntimeError(
                f"CSV must include vendor, VPN, and Qty_I columns (got headers: {rdr.fieldnames})"
            )
        c_qq = c_qq or c_qi

        for row in rdr:
            vendor = (row.get(c_vm) or "").strip()
            if not vendor:
                continue
            vpn = (row.get(c_vn) or "").strip()
            try:
                qty_i = Decimal(str(row.get(c_qi) or "0"))
                qty_q = Decimal(str(row.get(c_qq) or row.get(c_qi) or "0"))
            except Exception:
                continue
            internal_pn = (row.get(c_ip) or "").strip() if c_ip else ""
            name = (row.get(c_nm) or "").strip() if c_nm else ""
            rows_out.append(
                {
                    "vendor": vendor,
                    "vpn": vpn,
                    "qty_i": qty_i,
                    "qty_q": qty_q,
                    "internal_pn": internal_pn,
                    "name": name,
                    "unit_r": None,
                    "note": "",
                }
            )
    return rows_out


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Price BOM CSV: Nexar Supply first, then web scraping. No Arena."
    )
    ap.add_argument(
        "input_csv",
        nargs="?",
        default=os.environ.get("BOM_INPUT_CSV", "").strip() or str(DEFAULT_INPUT),
        help="BOM CSV path",
    )
    ap.add_argument(
        "--out-csv",
        default=os.environ.get("BOM_OUTPUT_CSV", "").strip() or str(DEFAULT_OUT_CSV),
    )
    ap.add_argument(
        "--out-xlsx",
        default=os.environ.get("BOM_OUTPUT_XLSX", "").strip() or str(DEFAULT_OUT_XLSX),
    )
    ap.add_argument(
        "--no-scrape",
        action="store_true",
        help="Disable distributor web scraping (Nexar only)",
    )
    args = ap.parse_args()

    in_path = Path(args.input_csv).expanduser().resolve()
    if not in_path.is_file():
        print(f"Input CSV not found: {in_path}", file=sys.stderr)
        return 1

    out_csv = Path(args.out_csv).expanduser().resolve()
    out_xlsx = Path(args.out_xlsx).expanduser().resolve()

    nexar_ready = False
    if nexar_credentials_configured():
        try:
            _nexar_get_token()
            nexar_ready = True
        except Exception as e:
            print(f"Nexar token failed: {e}. Continuing without Nexar.", file=sys.stderr)
    else:
        print("Nexar not configured (optional). Web scraping will be used where supported.", file=sys.stderr)

    if not nexar_ready and args.no_scrape:
        print("Nothing to do: Nexar unavailable and --no-scrape set.", file=sys.stderr)
        return 1

    rows = load_bom_rows_from_csv(in_path)
    delay = float(os.environ.get("BOM_REQUEST_DELAY_SEC", os.environ.get("NEXAR_REQUEST_DELAY_SEC", "0.18")))
    scrape = not args.no_scrape

    grand = Decimal(0)
    for r in rows:
        unit, note = price_line(r["vendor"], r["vpn"], r["qty_q"], nexar_ready=nexar_ready, scrape=scrape)
        r["unit_r"] = unit
        r["note"] = note
        if unit is not None:
            grand += d2(r["qty_i"] * unit)
        time.sleep(delay)

    csv_p, xlsx_p = write_outputs(rows, d2(grand), out_csv, out_xlsx)
    priced = sum(1 for x in rows if x.get("unit_r") is not None)
    print(
        f"Read {in_path}\nWrote {csv_p} and {xlsx_p}\n"
        f"{len(rows)} lines, {priced} with unit price, subtotal AUD {d2(grand)} "
        f"(nexar={'on' if nexar_ready else 'off'}, scrape={'on' if scrape else 'off'})"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
