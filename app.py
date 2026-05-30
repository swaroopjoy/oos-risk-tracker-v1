import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date, timedelta, datetime
from pathlib import Path
import base64
import re
import xlsxwriter

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="OOS Sentinel",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM CSS — OOS Sentinel brand: navy + cyan, dark/light adaptive
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Minimal clean overrides only */
.main .block-container { padding: 1.5rem 2rem; max-width: 100%; }
#MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

TODAY = date.today()

# ── Embedded input template (base64) ─────────────────────────────────────────
_TEMPLATE_B64 = "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"
_TEMPLATE_BYTES = base64.b64decode(_TEMPLATE_B64)

# ─────────────────────────────────────────────────────────────────────────────
# SKU PARSING
# ─────────────────────────────────────────────────────────────────────────────
def parse_sku(sku: str) -> list[dict]:
    """Parse 'A+Bx6+Cx2' → [{'base':'A','mult':1},{'base':'B','mult':6},{'base':'C','mult':2}]
    Handles both lowercase x and uppercase X as multiplier separator.
    """
    components = []
    for part in str(sku).split("+"):
        m = re.match(r"^(\d+?)(?:[xX](\d+))?$", part.strip())
        if m:
            components.append({"base": m.group(1), "mult": int(m.group(2)) if m.group(2) else 1})
    return components

def scope_key(seller, country, base_sku):
    return f"{seller}|{country}|{base_sku}"

# ─────────────────────────────────────────────────────────────────────────────
# XLSX UPLOAD PARSER
# ─────────────────────────────────────────────────────────────────────────────
def _detect_header_row(file) -> int:
    """
    Detect which row contains the actual data column headers.
    The template has 4 decorative rows before headers.
    We find the row that contains at least 3 of the known key column names.
    """
    KEY_COLS = {"SKU", "Brand", "Channel", "Country", "Seller Code",
                "Stock Locked / Reserved", "Total Reserved - All Campaigns",
                "Stock for Base SKU", "Campaign / Promotion Name", "Base SKU"}
    for header_row in range(0, 8):
        try:
            df = pd.read_excel(file, sheet_name=0, dtype=str, header=header_row, nrows=0)
            cols = {c.strip() for c in df.columns}
            if len(cols & KEY_COLS) >= 3:
                return header_row
        except Exception:
            continue
    return 0  # fallback


def validate_stock(promos: list[dict], stock_map: dict) -> list[str]:
    """
    Check every locked promo's scope_key exists in stock_map with a value > 0.
    Returns a list of missing scope descriptions (empty = all good).
    """
    missing = []
    seen = set()
    for p in promos:
        if not p.get("stock_lock"):
            continue
        sk = scope_key(p["seller"], p["country"], p["base_sku"])
        if sk in seen:
            continue
        seen.add(sk)
        if stock_map.get(sk, 0) == 0:
            missing.append(
                f"**{p['seller']} · {p['country']} · {p['base_sku']}** "
                f"(Promo SKU: {p['sku']})"
            )
    return missing


# ─────────────────────────────────────────────────────────────────────────────
def parse_upload(file) -> tuple[list[dict], dict]:
    """
    Parse the reservation tracker sheet.

    Key rules:
    - Auto-detects header row (handles OOS Sentinel template with 4 decorative rows).
    - Column lookup is always by TITLE, never by position.
    - Base SKU column is read directly when present; otherwise derived from SKU string.
    - Stock Lock is read from the lock column; inherited by continuation rows.
    - Total reservation = 'Total Reserved - All Campaigns' column.
    - Stock is stored per seller+country+base_sku and SHARED across all promos
      in that country for the same seller.
    - Continuation rows (blank SKU, blank seller) supply stock/reservations
      for the next component of the combo SKU above them.

    Returns:
        promos    — list of promo component dicts ready for compute_rows
        stock_map — {scope_key: stock} keyed "seller|country|base_sku"
    """
    header_row = _detect_header_row(file)
    df = pd.read_excel(file, sheet_name=0, dtype=str, header=header_row)
    df.columns = [c.strip() for c in df.columns]
    # Drop completely empty rows
    df = df.dropna(how="all")

    # ── Column name aliases (title-based, never positional) ────────────────
    COL = {
        "seller":    ["Seller Code", "Seller code", "Seller"],
        "country":   ["Country", "Country (SG/MY/TH)"],
        "brand":     ["Brand"],
        "channel":   ["Channel"],
        "sku":       ["SKU", "Promo SKU", "sku"],
        "base_sku":  ["Base SKU", "Base Sku", "base_sku"],          # ← NEW
        "campaign":  ["Campaign / Promotion Name", "Campaign Name"],
        "type":      ["Campaign Type", "Campaign type"],
        "lock":      ["Stock Locked / Reserved",
                      "Stock Lock / Reserved",
                      "Stock Lock", "Stock lock"],
        "start":     ["Promo Start Date", "Promo Start Date ",
                      "Start Date"],
        "end":       ["Promo End Date", "End Date"],
        "stock":     ["Stock for Base SKU",
                      "Today's Stock for Base SKU - (24 may)",
                      "Today's Stock for Base SKU",
                      "Today's Stock", "Stock"],
        "total_res": ["Total Reserved - All Campaigns",
                      "Total Reserved Across All Campaigns",
                      "Total Reserved", "total_reserved"],
        "nominated": ["Nominated stock (Non Reservation)",
                      "Nominated Stock", "Nominated stock"],
        "res_mp":    ["Reserved by Marketplace",                     # ← NEW alias
                      "Reserved by MP"],
        "res_seller": ["Reserved by Seller", "Reserved by DKSH"],
        "res_graas": ["Reserved by Graas"],
    }

    def find_col(key):
        for alias in COL.get(key, []):
            if alias in df.columns:
                return alias
        return None

    def gv(row, key):
        col = find_col(key)
        if not col:
            return None
        val = row.get(col)
        if val is None:
            return None
        s = str(val).strip()
        return None if s in ("", "nan", "None", "NaT", "NaN") else s

    def gn(row, key):
        v = gv(row, key)
        try:    return float(v) if v is not None else 0.0
        except: return 0.0

    def gdate(row, key, fallback=""):
        v = gv(row, key)
        if not v:
            return fallback
        try:    return pd.to_datetime(v).strftime("%Y-%m-%d")
        except: return fallback

    # ── Parse rows ────────────────────────────────────────────────────────
    records   = []
    stock_map = {}
    last      = {}
    pending_comps = []

    for _, row in df.iterrows():
        raw_sku      = gv(row, "sku")
        raw_seller   = gv(row, "seller")
        raw_lock     = gv(row, "lock")
        raw_stock    = gv(row, "stock")

        is_main = bool(raw_sku and raw_seller)
        is_cont = bool(not raw_sku and not raw_seller and raw_stock and last)

        if is_main:
            seller   = raw_seller
            country  = gv(row, "country")  or last.get("country", "")
            brand    = gv(row, "brand")    or last.get("brand", "")
            channel  = gv(row, "channel")  or last.get("channel", "")
            campaign = gv(row, "campaign") or last.get("campaign", "")
            typ      = gv(row, "type")     or last.get("type", "")
            start    = gdate(row, "start",  last.get("start", ""))
            end      = gdate(row, "end",    last.get("end", ""))

            lock_val   = str(raw_lock or "").strip().lower()
            stock_lock = lock_val in ("yes", "true", "1", "y")

            stock     = gn(row, "stock")
            total_res = gn(row, "total_res")
            nominated = gn(row, "nominated")

            # ── Base SKU resolution ──────────────────────────────────────
            # Priority: explicit Base SKU column → parse from SKU string
            raw_base_sku_col = gv(row, "base_sku")
            if raw_base_sku_col:
                # Column present — derive multiplier by comparing to SKU string
                components_from_sku = parse_sku(raw_sku)
                # Map base SKUs from parsed components, override with column value
                # For single-component SKUs the column IS the answer
                components = []
                for comp in components_from_sku:
                    if comp["base"] == raw_base_sku_col.strip():
                        components.append(comp)
                # If no match found (e.g. combo), fall back to full parse
                if not components:
                    components = components_from_sku
            else:
                components = parse_sku(raw_sku)

            # Store stock for ALL components — shared per seller+country
            for comp in components:
                sk = scope_key(seller, country, comp["base"])
                if stock > 0:   # only store if a real value was entered
                    stock_map[sk] = int(stock)

            pending_comps = []

            last = dict(seller=seller, country=country, brand=brand,
                        channel=channel, campaign=campaign, type=typ,
                        start=start, end=end, stock_lock=stock_lock, sku=raw_sku)

            for comp in components:
                records.append(dict(
                    seller=seller, country=country, brand=brand, channel=channel,
                    sku=raw_sku, base_sku=comp["base"], mult=comp["mult"],
                    campaign=campaign, type=typ, start=start, end=end,
                    stock_lock=stock_lock,
                    total_res=total_res,
                    nominated=nominated,
                    _stock_raw=int(stock),
                ))

        elif is_cont and pending_comps:
            comp       = pending_comps.pop(0)
            cont_stock = gn(row, "stock")
            cont_res   = gn(row, "total_res")
            cont_nom   = gn(row, "nominated")

            sk = scope_key(last["seller"], last["country"], comp["base"])
            if cont_stock > 0:
                stock_map[sk] = int(cont_stock)

            for rec in reversed(records):
                if (rec["base_sku"] == comp["base"]
                        and rec["campaign"] == last["campaign"]
                        and rec["sku"] == last["sku"]):
                    rec["total_res"]  = cont_res
                    rec["nominated"]  = cont_nom
                    rec["_stock_raw"] = int(cont_stock)
                    break

    return records, stock_map


# ─────────────────────────────────────────────────────────────────────────────
# CORE CALCULATIONS — scoped to country + base_sku
# (stock is shared per country regardless of seller/channel)
# ─────────────────────────────────────────────────────────────────────────────
def compute_rows(promos: list[dict], stock_map: dict) -> pd.DataFrame:
    """
    For each promo-component record:
      - demand = total_res × mult  (reservations already reflect orders × units)
      - stock  = shared stock for that country + base_sku
      - gap    = stock − total demand across ALL promos for that scope
                 (NOT per-row gap — aggregated below in compute_scope_summary)
    """
    rows = []
    for p in promos:
        sk    = scope_key(p["seller"], p["country"], p["base_sku"])
        stock = stock_map.get(sk, p.get("_stock_raw", 0))
        # demand per base SKU unit = total_res × multiplier
        demand = int(p["total_res"] * p["mult"]) if p["stock_lock"] else 0

        rows.append({
            "seller":     p["seller"],
            "country":    p["country"],
            "brand":      p["brand"],
            "channel":    p["channel"],
            "scope_key":  sk,
            "scope_label": f"{p['seller']} · {p['country']} · {p['base_sku']}",
            "base_sku":   p["base_sku"],
            "promo_sku":  p["sku"],
            "mult":       p["mult"],
            "campaign":   p["campaign"],
            "type":       p["type"],
            "start":      p["start"],
            "end":        p["end"],
            "stock_lock": p["stock_lock"],
            "stock":      int(stock),
            "total_res":  int(p["total_res"]),
            "nominated":  int(p.get("nominated", 0)),
            "demand":     demand,          # base units consumed by this promo
        })

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)

    # ── Aggregate total demand per scope across all locked promos ──────────
    locked = df[df["stock_lock"]]
    scope_demand = locked.groupby("scope_key")["demand"].sum().rename("total_demand")
    df = df.merge(scope_demand, on="scope_key", how="left")
    df["total_demand"] = df["total_demand"].fillna(0).astype(int)

    # ── Gap and OOS based on aggregated demand vs shared stock ─────────────
    # Use pd.NA for unlocked rows so the column stays integer-compatible
    df["gap"]     = df.apply(
        lambda r: int(r["stock"] - r["total_demand"]) if r["stock_lock"] else pd.NA, axis=1
    )
    df["oos"]     = df.apply(
        lambda r: bool(r["stock_lock"] and pd.notna(r["gap"]) and r["gap"] < 0), axis=1
    )
    df["restock"] = df.apply(
        lambda r: int(abs(r["gap"])) if r["oos"] else 0, axis=1
    )
    df["status"]  = df.apply(lambda r:
        "OOS"     if r["oos"] else
        "Watch"   if (r["stock_lock"] and pd.notna(r["gap"]) and r["gap"] < 20) else
        "Safe"    if r["stock_lock"] else
        "No lock", axis=1)

    return df

def compute_conflicts(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    locked = df[df["stock_lock"]].copy()
    conflicts = []
    # Group by scope — each scope has its total_demand vs stock already computed
    # Conflicts exist when 2+ different campaigns overlap in date within same scope
    for scope, grp in locked.groupby("scope_key"):
        rows = grp.drop_duplicates(subset=["campaign","start","end"]).to_dict("records")
        for i in range(len(rows)):
            for j in range(i+1, len(rows)):
                a, b = rows[i], rows[j]
                if a["campaign"] == b["campaign"]:
                    continue
                if a["end"] >= b["start"] and b["end"] >= a["start"]:
                    os_ = max(a["start"], b["start"])
                    oe_ = min(a["end"],   b["end"])
                    combined = a["demand"] + b["demand"]
                    conflicts.append({
                        "scope_key":       scope,
                        "scope_label":     a["scope_label"],
                        "seller":          a["seller"],
                        "country":         a["country"],
                        "brand":           a["brand"],
                        "channel":         a["channel"],
                        "stock":           a["stock"],
                        "campaign_a":      a["campaign"],
                        "campaign_b":      b["campaign"],
                        "overlap_start":   os_,
                        "overlap_end":     oe_,
                        "combined_demand": combined,
                        "verdict":         "OOS risk" if combined > a["stock"] else "Overlap",
                    })
    return pd.DataFrame(conflicts) if conflicts else pd.DataFrame()

def compute_heatmap(df: pd.DataFrame, ref_date: date) -> pd.DataFrame:
    rows = []
    locked = df[df["stock_lock"]] if not df.empty else pd.DataFrame()
    for d in range(-3, 8):
        dt   = ref_date + timedelta(days=d)
        dstr = dt.strftime("%Y-%m-%d")
        label = "D" if d == 0 else (f"D+{d}" if d > 0 else f"D{d}")

        scope_map = {}
        if not locked.empty:
            active = locked[(locked["start"] <= dstr) & (locked["end"] >= dstr)]
            for _, r in active.iterrows():
                k = r["scope_key"]
                if k not in scope_map:
                    scope_map[k] = {"demand": 0, "stock": r["stock"]}
                scope_map[k]["demand"] += r["demand"]

        total_demand = sum(s["demand"] for s in scope_map.values())
        total_stock  = sum(s["stock"]  for s in scope_map.values())
        units_risk   = sum(max(0, s["demand"] - s["stock"]) for s in scope_map.values())
        ratio        = round(total_demand / total_stock * 100, 1) if total_stock > 0 else 0
        risk_level   = ("None" if ratio == 0 else
                        "Low"    if ratio < 33 else
                        "Medium" if ratio < 66 else "High")

        rows.append({
            "day_label":    label,
            "date":         dt,
            "date_str":     dt.strftime("%d %b"),
            "active_scopes": len(scope_map),
            "total_demand": total_demand,
            "total_stock":  total_stock,
            "ratio":        ratio,
            "units_at_risk": units_risk,
            "risk_level":   risk_level,
            "scope_keys":   ", ".join(scope_map.keys()) or "—",
            "is_today":     d == 0,
        })
    return pd.DataFrame(rows)

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────
STATUS_COLORS = {
    "OOS":     {"bg": "#FDF0EE", "font": "#C0392B"},
    "Watch":   {"bg": "#FEF9EC", "font": "#B45309"},
    "Safe":    {"bg": "#EDF7F2", "font": "#1A6B3C"},
    "No lock": {"bg": "#F5F4F0", "font": "#9E9B91"},
}
RISK_COLORS = {
    "High":   "#FDF0EE",
    "Medium": "#FEF9EC",
    "Low":    "#EDF7F2",
    "None":   "#F5F4F0",
}

def safe_val(v):
    """Convert pandas/numpy scalars to plain Python types xlsxwriter can handle."""
    if v is None or v is pd.NA or v is pd.NaT:
        return "—"
    try:
        import numpy as np
        if isinstance(v, (np.integer,)):   return int(v)
        if isinstance(v, (np.floating,)):  return float(v)
        if isinstance(v, (np.bool_,)):     return bool(v)
    except ImportError:
        pass
    if isinstance(v, float) and (v != v):  # NaN check
        return "—"
    return v

def safe_num(v, fallback="—"):
    """Return plain int/float or fallback string."""
    sv = safe_val(v)
    return sv if isinstance(sv, (int, float)) else fallback

def build_excel(df_rows: pd.DataFrame, df_conf: pd.DataFrame, hm_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})

    # ── Formats ──────────────────────────────────────────────────────────────
    hdr_fmt  = wb.add_format({"bold":True,"bg_color":"#F0EFE9","border":1,"font_name":"Calibri","font_size":10,"align":"center","valign":"vcenter"})
    base_fmt = wb.add_format({"border":1,"font_name":"Calibri","font_size":10,"valign":"vcenter"})
    num_fmt  = wb.add_format({"border":1,"font_name":"Calibri","font_size":10,"valign":"vcenter","align":"right"})
    mono_fmt = wb.add_format({"border":1,"font_name":"Courier New","font_size":9,"valign":"vcenter"})

    def row_fmt(status, align="left"):
        c = STATUS_COLORS.get(status, STATUS_COLORS["No lock"])
        return wb.add_format({"border":1,"font_name":"Calibri","font_size":10,"valign":"vcenter",
                               "align":align,"bg_color":c["bg"],"font_color":c["font"]})
    def risk_fmt(level):
        return wb.add_format({"border":1,"font_name":"Calibri","font_size":10,"valign":"vcenter",
                               "bg_color":RISK_COLORS.get(level,"#F5F4F0")})

    # ── Sheet 1: SKU breakdown ───────────────────────────────────────────────
    ws1 = wb.add_worksheet("SKU breakdown")
    ws1.freeze_panes(1, 0)
    ws1.set_row(0, 20)
    headers = ["Seller","Country","Brand","Channel","Scope key","Base SKU","Promo SKU",
               "Campaign","Type","Start","End","Stock lock","Stock","Reserved","Demand","Gap","Status"]
    widths  = [10,8,12,22,30,14,28,26,16,11,11,10,7,9,7,6,8]
    for c, (h, w) in enumerate(zip(headers, widths)):
        ws1.write(0, c, h, hdr_fmt)
        ws1.set_column(c, c, w)

    if not df_rows.empty:
        sorted_rows = df_rows.sort_values(["seller","country","base_sku","stock_lock"],
                                          ascending=[True,True,True,False])
        for ri, (_, r) in enumerate(sorted_rows.iterrows(), start=1):
            st_ = r["status"]
            rf  = row_fmt(st_)
            rf_num = wb.add_format({"border":1,"font_name":"Calibri","font_size":10,"valign":"vcenter",
                                    "align":"right","bg_color":STATUS_COLORS.get(st_,STATUS_COLORS["No lock"])["bg"],
                                    "font_color":STATUS_COLORS.get(st_,STATUS_COLORS["No lock"])["font"]})
            rf_mono = wb.add_format({"border":1,"font_name":"Courier New","font_size":9,"valign":"vcenter",
                                     "bg_color":STATUS_COLORS.get(st_,STATUS_COLORS["No lock"])["bg"],
                                     "font_color":STATUS_COLORS.get(st_,STATUS_COLORS["No lock"])["font"]})
            ws1.write(ri, 0,  safe_val(r["seller"]),    rf)
            ws1.write(ri, 1,  safe_val(r["country"]),   rf)
            ws1.write(ri, 2,  safe_val(r["brand"]),     rf)
            ws1.write(ri, 3,  safe_val(r["channel"]),   rf)
            ws1.write(ri, 4,  safe_val(r["scope_key"]), rf_mono)
            ws1.write(ri, 5,  safe_val(r["base_sku"]),  rf_mono)
            ws1.write(ri, 6,  safe_val(r["promo_sku"]), rf_mono)
            ws1.write(ri, 7,  safe_val(r["campaign"]),  rf)
            ws1.write(ri, 8,  safe_val(r["type"]),      rf)
            ws1.write(ri, 9,  safe_val(r["start"]),     rf)
            ws1.write(ri, 10, safe_val(r["end"]),        rf)
            ws1.write(ri, 11, "Yes" if r["stock_lock"] else "No", rf)
            ws1.write(ri, 12, safe_num(r["stock"]),      rf_num)
            ws1.write(ri, 13, safe_num(r["total_res"]) if r["stock_lock"] else "—", rf_num)
            ws1.write(ri, 14, safe_num(r["demand"])   if r["stock_lock"] else "—", rf_num)
            ws1.write(ri, 15, safe_num(r["gap"])      if r["gap"] is not None else "—", rf_num)
            ws1.write(ri, 16, safe_val(st_),           rf)

    # ── Sheet 2: Conflicts ───────────────────────────────────────────────────
    ws2 = wb.add_worksheet("Conflicts")
    ws2.freeze_panes(1, 0)
    ws2.set_row(0, 20)
    conf_headers = ["Scope key","Seller","Country","Brand","Channel","Stock",
                    "Campaign A","Campaign B","Overlap start","Overlap end","Combined demand","Verdict"]
    conf_widths  = [30,10,8,12,22,7,26,26,12,10,16,10]
    for c, (h, w) in enumerate(zip(conf_headers, conf_widths)):
        ws2.write(0, c, h, hdr_fmt)
        ws2.set_column(c, c, w)

    if not df_conf.empty:
        for ri, (_, r) in enumerate(df_conf.sort_values("scope_key").iterrows(), start=1):
            is_oos = r["verdict"] == "OOS risk"
            bg = "#FDF0EE" if is_oos else "#FEF9EC"
            fc = "#C0392B" if is_oos else "#B45309"
            cf = wb.add_format({"border":1,"font_name":"Calibri","font_size":10,
                                 "valign":"vcenter","bg_color":bg,"font_color":fc})
            cf_num = wb.add_format({"border":1,"font_name":"Calibri","font_size":10,"valign":"vcenter",
                                    "align":"right","bg_color":bg,"font_color":fc})
            cf_mono = wb.add_format({"border":1,"font_name":"Courier New","font_size":9,
                                     "valign":"vcenter","bg_color":bg,"font_color":fc})
            ws2.write(ri, 0,  safe_val(r["scope_key"]),       cf_mono)
            ws2.write(ri, 1,  safe_val(r["seller"]),           cf)
            ws2.write(ri, 2,  safe_val(r["country"]),          cf)
            ws2.write(ri, 3,  safe_val(r["brand"]),            cf)
            ws2.write(ri, 4,  safe_val(r["channel"]),          cf)
            ws2.write(ri, 5,  safe_num(r["stock"]),            cf_num)
            ws2.write(ri, 6,  safe_val(r["campaign_a"]),       cf)
            ws2.write(ri, 7,  safe_val(r["campaign_b"]),       cf)
            ws2.write(ri, 8,  safe_val(r["overlap_start"]),    cf)
            ws2.write(ri, 9,  safe_val(r["overlap_end"]),      cf)
            ws2.write(ri, 10, safe_num(r["combined_demand"]),  cf_num)
            ws2.write(ri, 11, safe_val(r["verdict"]),          cf)

    # ── Sheet 3: Restock ─────────────────────────────────────────────────────
    ws3 = wb.add_worksheet("Restock")
    ws3.freeze_panes(1, 0)
    ws3.set_row(0, 20)
    rs_headers = ["Type","Scope key","Seller","Country","Brand","Base SKU",
                  "Stock","Locked demand","Nominated qty","Combined demand","Gap","Restock needed"]
    rs_widths  = [10,30,10,8,12,14,7,14,13,16,8,14]
    for c, (h, w) in enumerate(zip(rs_headers, rs_widths)):
        ws3.write(0, c, h, hdr_fmt)
        ws3.set_column(c, c, w)

    rs_rows = []
    if not df_rows.empty:
        # Hard OOS
        hard_by_scope = {}
        for _, r in df_rows[df_rows["oos"]].iterrows():
            k = r["scope_key"]
            if k not in hard_by_scope or hard_by_scope[k]["restock"] < r["restock"]:
                hard_by_scope[k] = r
        for r in hard_by_scope.values():
            rs_rows.append(("Hard OOS", r["scope_key"], r["seller"], r["country"], r["brand"],
                            r["base_sku"], r["stock"], r["demand"], 0, r["demand"], r["gap"], r["restock"]))

        # Soft risk
        scope_agg = {}
        for _, r in df_rows.iterrows():
            k = r["scope_key"]
            if k not in scope_agg:
                scope_agg[k] = {"locked":0,"nominated":0,"stock":r["stock"],"base_sku":r["base_sku"],
                                 "seller":r["seller"],"country":r["country"],"brand":r["brand"]}
            if r["stock_lock"]: scope_agg[k]["locked"]   += r["demand"]
            else:               scope_agg[k]["nominated"] += r["nominated"]

        for k, s in scope_agg.items():
            if s["nominated"] > 0:
                combined = s["locked"] + s["nominated"]
                gap_soft = s["stock"] - combined
                rs_rows.append(("Soft risk", k, s["seller"], s["country"], s["brand"],
                                s["base_sku"], s["stock"], s["locked"], s["nominated"],
                                combined, gap_soft, max(0, -gap_soft)))

    for ri, row in enumerate(rs_rows, start=1):
        is_hard = row[0] == "Hard OOS"
        bg = "#FDF0EE" if is_hard else "#FEF9EC"
        fc = "#C0392B" if is_hard else "#B45309"
        rf  = wb.add_format({"border":1,"font_name":"Calibri","font_size":10,"valign":"vcenter","bg_color":bg,"font_color":fc})
        rfn = wb.add_format({"border":1,"font_name":"Calibri","font_size":10,"valign":"vcenter","align":"right","bg_color":bg,"font_color":fc})
        rfm = wb.add_format({"border":1,"font_name":"Courier New","font_size":9,"valign":"vcenter","bg_color":bg,"font_color":fc})
        ws3.write(ri, 0,  safe_val(row[0]),  rf)
        ws3.write(ri, 1,  safe_val(row[1]),  rfm)
        ws3.write(ri, 2,  safe_val(row[2]),  rf)
        ws3.write(ri, 3,  safe_val(row[3]),  rf)
        ws3.write(ri, 4,  safe_val(row[4]),  rf)
        ws3.write(ri, 5,  safe_val(row[5]),  rfm)
        ws3.write(ri, 6,  safe_num(row[6]),  rfn)
        ws3.write(ri, 7,  safe_num(row[7]),  rfn)
        ws3.write(ri, 8,  safe_num(row[8]),  rfn)
        ws3.write(ri, 9,  safe_num(row[9]),  rfn)
        ws3.write(ri, 10, safe_num(row[10]), rfn)
        ws3.write(ri, 11, safe_num(row[11]), rfn)

    # ── Sheet 4: Heatmap summary ─────────────────────────────────────────────
    ws4 = wb.add_worksheet("Heatmap summary")
    ws4.freeze_panes(1, 0)
    ws4.set_row(0, 20)
    hm_headers = ["Day label","Date","Active scopes","Total demand","Total stock",
                  "Demand ratio %","Units at risk","Risk level","Active scope keys"]
    hm_widths  = [9,12,14,13,11,14,13,9,50]
    for c, (h, w) in enumerate(zip(hm_headers, hm_widths)):
        ws4.write(0, c, h, hdr_fmt)
        ws4.set_column(c, c, w)

    for ri, (_, r) in enumerate(hm_df.iterrows(), start=1):
        bg   = RISK_COLORS.get(r["risk_level"], "#F5F4F0")
        bold = r["is_today"]
        rf   = wb.add_format({"border":1,"font_name":"Calibri","font_size":10,"valign":"vcenter",
                               "bg_color":bg,"bold":bold})
        rfn  = wb.add_format({"border":1,"font_name":"Calibri","font_size":10,"valign":"vcenter",
                               "align":"right","bg_color":bg,"bold":bold})
        ws4.write(ri, 0, safe_val(r["day_label"]),      rf)
        ws4.write(ri, 1, safe_val(r["date_str"]),        rf)
        ws4.write(ri, 2, safe_num(r["active_scopes"]),   rfn)
        ws4.write(ri, 3, safe_num(r["total_demand"]),    rfn)
        ws4.write(ri, 4, safe_num(r["total_stock"]),     rfn)
        ws4.write(ri, 5, str(safe_val(r["ratio"]))+"%",  rfn)
        ws4.write(ri, 6, safe_num(r["units_at_risk"]),   rfn)
        ws4.write(ri, 7, safe_val(r["risk_level"]),      rf)
        ws4.write(ri, 8, safe_val(r["scope_keys"]),      rf)

    wb.close()
    output.seek(0)
    return output.read()

# ─────────────────────────────────────────────────────────────────────────────
# DISPLAY HELPERS
# ─────────────────────────────────────────────────────────────────────────────
STATUS_EMOJI = {"OOS": "🔴 OOS", "Watch": "🟡 Watch", "Safe": "🟢 Safe", "No lock": "⚪ No lock"}

def style_status(val):
    colors = {"OOS":"background-color:#FDF0EE;color:#C0392B;font-weight:500",
              "Watch":"background-color:#FEF9EC;color:#B45309;font-weight:500",
              "Safe":"background-color:#EDF7F2;color:#1A6B3C",
              "No lock":"background-color:#F5F4F0;color:#9E9B91"}
    return colors.get(val, "")

def style_gap(val):
    try:
        v = float(val)
        if v < 0:  return "color:#C0392B;font-weight:500"
        if v < 20: return "color:#B45309;font-weight:500"
        return "color:#1A6B3C"
    except: return ""

def style_verdict(val):
    if val == "OOS risk": return "background-color:#FDF0EE;color:#C0392B;font-weight:500"
    if val == "Overlap":  return "background-color:#FEF9EC;color:#B45309;font-weight:500"
    return ""

def style_risk(val):
    m = {"High":"background-color:#FDF0EE;color:#C0392B;font-weight:500",
         "Medium":"background-color:#FEF9EC;color:#B45309;font-weight:500",
         "Low":"background-color:#EDF7F2;color:#1A6B3C",
         "None":"background-color:#F5F4F0;color:#9E9B91"}
    return m.get(val,"")

def render_heatmap(hm_df: pd.DataFrame, mode: str):
    """Render heatmap as coloured HTML cards."""
    RISK_BG   = {"High":"#2A0A0C","Medium":"#2A1A08","Low":"#092014","None":"rgba(0,0,0,0)"}
    RISK_TC   = {"High":"#FF6B7A","Medium":"#F4A261","Low":"#52D9A0","None":"#4A7090"}
    RISK_BC   = {"High":"rgba(230,57,70,0.4)","Medium":"rgba(244,162,97,0.4)","Low":"rgba(0,180,216,0.3)","None":"rgba(0,180,216,0.15)"}
    # Light mode overrides via CSS var — handled by inline styles with contrasting pairs
    RISK_BG_L = {"High":"#FEF0F1","Medium":"#FEF6EE","Low":"#E8F9F3","None":"#F0F4FA"}
    RISK_TC_L = {"High":"#C0392B","Medium":"#B45309","Low":"#1A7A50","None":"#7A95B0"}
    RISK_BC_L = {"High":"#F5C6C2","Medium":"#FAC77A","Low":"#A0DFC0","None":"#C8D6E5"}

    if mode == "Demand ratio %": vals = hm_df["ratio"]
    elif mode == "Promo count":  vals = hm_df["active_scopes"]
    else:                        vals = hm_df["units_at_risk"]

    cards_html = '<div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px">'
    for (_, row), val in zip(hm_df.iterrows(), vals):
        rl   = row["risk_level"]
        bg   = RISK_BG_L.get(rl, "#F0F4FA")
        tc   = RISK_TC_L.get(rl, "#7A95B0")
        bc   = RISK_BC_L.get(rl, "#C8D6E5")
        disp = f"{val:.0f}%" if mode == "Demand ratio %" else str(int(val))
        today_border = "border:2px solid #00B4D8!important;" if row["is_today"] else f"border:1px solid {bc};"
        cards_html += f"""
        <div title="{row['scope_keys']}" style="width:68px;height:76px;border-radius:8px;
            background:{bg};{today_border}display:flex;flex-direction:column;
            align-items:center;justify-content:center;cursor:default;gap:1px;">
          <span style="font-size:10px;font-weight:600;color:{tc};">{row['day_label']}</span>
          <span style="font-size:9px;color:{tc};opacity:.75">{row['date_str']}</span>
          <span style="font-size:15px;font-weight:600;color:{tc};margin-top:2px;">{disp}</span>
          <span style="font-size:9px;color:{tc};opacity:.65">{int(row['active_scopes'])} scope{'s' if row['active_scopes']!=1 else ''}</span>
        </div>"""
    cards_html += "</div>"

    legend_html = """<div style="display:flex;gap:14px;flex-wrap:wrap;font-size:11px;color:#7A95B0">
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#F0F4FA;border:1px solid #C8D6E5;margin-right:4px"></span>No promos</span>
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#E8F9F3;border:1px solid #A0DFC0;margin-right:4px"></span>Low</span>
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#FEF6EE;border:1px solid #FAC77A;margin-right:4px"></span>Medium</span>
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#FEF0F1;border:1px solid #F5C6C2;margin-right:4px"></span>High</span>
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:transparent;border:2px solid #00B4D8;margin-right:4px"></span>Today (D)</span>
    </div>"""
    st.markdown(cards_html + legend_html, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# RESTOCK PANEL
# ─────────────────────────────────────────────────────────────────────────────
def render_restock(df_rows: pd.DataFrame):
    if df_rows.empty:
        st.info("No data to analyse.")
        return

    # ── Hard OOS ─────────────────────────────────────────────────────────────
    st.markdown("**🔒 Hard OOS — locked stock only**")
    # One entry per scope (de-duplicated — gap/restock are the same for all rows in a scope)
    hard_by_scope = {}
    for _, r in df_rows[df_rows["oos"]].iterrows():
        k = r["scope_key"]
        if k not in hard_by_scope:
            hard_by_scope[k] = r

    if not hard_by_scope:
        st.caption("No hard OOS for current selection.")
    else:
        for r in hard_by_scope.values():
            st.markdown(f"""
            <div style="display:flex;align-items:center;justify-content:space-between;
                padding:11px 14px;background:#FDF0EE;border:1px solid #F5C6C2;
                border-radius:8px;margin-bottom:6px">
              <div>
                <div style="font-family:'DM Mono',monospace;font-size:12px;font-weight:500;color:#C0392B">{r['base_sku']}</div>
                <div style="font-size:10px;color:#C0392B;opacity:.7;font-family:'DM Mono',monospace">{r['seller']} · {r['country']}</div>
                <div style="font-size:11px;color:#C0392B;opacity:.8;margin-top:2px">
                  {r['brand']} · stock {r['stock']} · total demand across all promos {r['total_demand']} · deficit {abs(int(r['gap']))}
                </div>
              </div>
              <div style="font-size:16px;font-weight:600;color:#C0392B;margin-left:16px;white-space:nowrap">+{abs(int(r['gap']))} units</div>
            </div>""", unsafe_allow_html=True)

    # ── Soft Risk ─────────────────────────────────────────────────────────────
    st.markdown("**⚠️ Soft risk — locked + nominated combined**")
    scope_agg = {}
    for _, r in df_rows.iterrows():
        k = r["scope_key"]
        if k not in scope_agg:
            scope_agg[k] = {"locked": int(r["total_demand"]) if r["stock_lock"] else 0,
                             "nominated": 0, "stock": r["stock"],
                             "base_sku": r["base_sku"], "seller": r["seller"],
                             "country": r["country"], "brand": r["brand"], "channel": r["channel"]}
        if not r["stock_lock"]:
            scope_agg[k]["nominated"] += int(r["nominated"])

    soft_items = [(k, s) for k, s in scope_agg.items()
                  if s["nominated"] > 0 and s["stock"] - s["locked"] - s["nominated"] < 0]

    if not soft_items:
        st.caption("No soft risk when nominated stock is included.")
    else:
        for k, s in soft_items:
            combined = s["locked"] + s["nominated"]
            gap_soft = s["stock"] - combined
            extra    = abs(gap_soft)
            st.markdown(f"""
            <div style="display:flex;align-items:center;justify-content:space-between;
                padding:11px 14px;background:#FEF9EC;border:1px solid #F5DFA0;
                border-radius:8px;margin-bottom:6px">
              <div>
                <div style="font-family:'DM Mono',monospace;font-size:12px;font-weight:500;color:#B45309">{s['base_sku']}</div>
                <div style="font-size:10px;color:#B45309;opacity:.7;font-family:'DM Mono',monospace">{s['seller']} · {s['country']}</div>
                <div style="font-size:11px;color:#B45309;opacity:.8;margin-top:2px">{s['brand']} · stock {s['stock']} · locked {s['locked']} + nominated {s['nominated']} = combined {combined}</div>
              </div>
              <div style="font-size:16px;font-weight:600;color:#B45309;margin-left:16px;white-space:nowrap">+{extra} units</div>
            </div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🛡️ OOS Sentinel")
    st.caption("Stock risk intelligence platform")
    st.divider()

    # ── Data source ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-hdr">Data source</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload sheet", type=["xlsx","csv"], label_visibility="collapsed")

    # Template download — embedded directly in app, always available
    st.download_button(
        label="⬇ Download input template",
        data=_TEMPLATE_BYTES,
        file_name="OOS_Sentinel_Input_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    if uploaded:
        try:
            promos, stock_map = parse_upload(uploaded)
            st.success(f"✓ {uploaded.name}  ·  {len(promos)} rows loaded")
            # On first upload (or new file), reset date range to today
            if st.session_state.get("last_uploaded") != uploaded.name:
                st.session_state["last_uploaded"] = uploaded.name
                st.session_state["d_from"] = TODAY
                st.session_state["d_to"]   = TODAY
            # Run stock validation — store missing scopes in session state
            st.session_state["missing_stock"] = validate_stock(promos, stock_map)
        except Exception as e:
            st.error(f"Parse error: {e}")
            promos, stock_map = [], {}
    else:
        promos, stock_map = [], {}
        st.markdown("""
        <div style="background:rgba(0,180,216,0.08);border:1px solid rgba(0,180,216,0.25);
             border-radius:8px;padding:10px 12px;font-size:12px;color:#00B4D8;margin-top:6px">
            📂 Upload your tracker sheet or download the template above.
        </div>""", unsafe_allow_html=True)

    st.markdown("---")

    # ── Date range — defaults to today on upload, user-adjustable after ───────
    st.markdown('<div class="section-hdr">Date range</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        d_from = st.date_input("From", value=st.session_state.get("d_from", TODAY),
                               key="d_from", label_visibility="collapsed")
    with col2:
        d_to   = st.date_input("To",   value=st.session_state.get("d_to",   TODAY),
                               key="d_to",   label_visibility="collapsed")
    st.caption(f"From {d_from.strftime('%d %b')} → {d_to.strftime('%d %b %Y')}")

    st.markdown("---")

    # ── Dimension filters ─────────────────────────────────────────────────────
    all_sellers   = sorted(set(p["seller"]  for p in promos))
    all_countries = sorted(set(p["country"] for p in promos))
    all_brands    = sorted(set(p["brand"]   for p in promos))
    all_channels  = sorted(set(p["channel"] for p in promos))
    all_types     = sorted(set(p["type"]    for p in promos))

    st.markdown('<div class="section-hdr">Seller code</div>', unsafe_allow_html=True)
    sel_sellers   = st.multiselect("Seller",   all_sellers,   default=all_sellers,   label_visibility="collapsed")
    st.markdown('<div class="section-hdr">Country</div>', unsafe_allow_html=True)
    sel_countries = st.multiselect("Country",  all_countries, default=all_countries, label_visibility="collapsed")
    st.markdown('<div class="section-hdr">Brand</div>', unsafe_allow_html=True)
    sel_brands    = st.multiselect("Brand",    all_brands,    default=all_brands,    label_visibility="collapsed")
    st.markdown('<div class="section-hdr">Channel</div>', unsafe_allow_html=True)
    sel_channels  = st.multiselect("Channel",  all_channels,  default=all_channels,  label_visibility="collapsed")

    st.markdown("---")

    st.markdown('<div class="section-hdr">SKU lookup</div>', unsafe_allow_html=True)
    sku_input = st.text_input("Add SKU", placeholder="e.g. 101336151x8", label_visibility="collapsed")
    if "sku_list" not in st.session_state:
        st.session_state.sku_list = []
    if sku_input and sku_input not in st.session_state.sku_list:
        st.session_state.sku_list.append(sku_input)
    if st.session_state.sku_list:
        to_remove = st.multiselect("Active SKUs", st.session_state.sku_list,
                                   default=st.session_state.sku_list, label_visibility="collapsed")
        st.session_state.sku_list = to_remove

    st.markdown("---")

    st.markdown('<div class="section-hdr">More filters</div>', unsafe_allow_html=True)
    sel_type = st.selectbox("Campaign type", ["All types"] + all_types, label_visibility="collapsed")
    sel_lock = st.selectbox("Stock lock",    ["All","Locked only","Unlocked only"], label_visibility="collapsed")
    sel_show = st.selectbox("Show",          ["All SKUs","OOS risk only","Safe stock only"], label_visibility="collapsed")

    st.markdown("---")
    run = st.button("▶ Analyse", use_container_width=True, type="primary")

# ─────────────────────────────────────────────────────────────────────────────
# FILTER + COMPUTE  (runs on load and whenever Analyse is pressed)
# ─────────────────────────────────────────────────────────────────────────────
df_str = d_from.strftime("%Y-%m-%d")
dt_str = d_to.strftime("%Y-%m-%d")

filtered_promos = [
    p for p in promos
    if p["seller"]  in (sel_sellers  or all_sellers)
    and p["country"] in (sel_countries or all_countries)
    and p["brand"]   in (sel_brands   or all_brands)
    and p["channel"] in (sel_channels or all_channels)
    and p["end"]     >= df_str
    and p["start"]   <= dt_str
    and (sel_type == "All types" or p["type"] == sel_type)
    and (sel_lock == "All" or (sel_lock == "Locked only" and p["stock_lock"]) or (sel_lock == "Unlocked only" and not p["stock_lock"]))
    and (not st.session_state.sku_list or p["sku"] in st.session_state.sku_list)
]

df_rows = compute_rows(filtered_promos, stock_map)
if not df_rows.empty and sel_show == "OOS risk only":
    df_rows = df_rows[df_rows["oos"]]
elif not df_rows.empty and sel_show == "Safe stock only":
    df_rows = df_rows[~df_rows["oos"] & df_rows["stock_lock"]]

df_conf = compute_conflicts(df_rows) if not df_rows.empty else pd.DataFrame()
hm_df   = compute_heatmap(df_rows, TODAY)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN CONTENT
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("## 🛡️ OOS Sentinel")
st.caption("Stock scoped per **Seller + Country** — same SKU in different countries is treated as separate stock. "
           "Demand & OOS calculated only for **Stock lock = Yes** rows.")

# ── No file uploaded yet — show welcome screen ───────────────────────────────
if not promos:
    st.markdown("---")
    st.markdown("### 📂 No data loaded yet")
    st.info("Upload your Stock Reservation Tracker sheet using the sidebar to begin analysis.")
    st.markdown("""
    **Required columns:**
    `Seller Code` · `Country` · `Brand` · `Channel` · `SKU` · `Base SKU` ·
    `Campaign / Promotion Name` · `Campaign Type` · `Stock Locked / Reserved` ·
    `Promo Start Date` · `Promo End Date` · `Stock for Base SKU` ·
    `Total Reserved - All Campaigns`
    """)
    st.stop()

# ── Stock validation — block analysis if any locked scope is missing stock ────
missing_stock = st.session_state.get("missing_stock", [])
if missing_stock:
    st.error(
        "⛔ **Analysis blocked — stock missing for the following scopes.**\n\n"
        "Every locked promotion must have a `Stock for Base SKU` value before "
        "OOS calculations can run. Fill in the missing values in your sheet and re-upload.\n"
    )
    st.markdown("**Missing stock — Seller · Country · Base SKU:**")
    for scope in missing_stock:
        st.markdown(f"- {scope}")
    st.markdown("---")
    st.caption(
        "💡 Stock is shared per Seller + Country + Base SKU. "
        "You only need to enter the value once per combination — "
        "the tool applies it to all promo rows with the same scope."
    )
    st.stop()

# ── Metrics ──────────────────────────────────────────────────────────────────
locked_rows  = df_rows[df_rows["stock_lock"]] if not df_rows.empty else pd.DataFrame()
# Deduplicate by scope_key for metric counts — one scope = one country+seller+base_sku bucket
unique_scopes     = locked_rows.drop_duplicates("scope_key") if not locked_rows.empty else pd.DataFrame()
n_scopes    = unique_scopes["scope_key"].nunique()                              if not unique_scopes.empty else 0
n_oos       = unique_scopes[unique_scopes["oos"]]["scope_key"].nunique()        if not unique_scopes.empty else 0
n_conf      = len(df_conf)
n_restock   = int(unique_scopes[unique_scopes["oos"]]["restock"].sum())         if not unique_scopes.empty else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Stock scopes",      n_scopes)
c2.metric("OOS risk scopes",   n_oos,    delta=f"{n_oos} at risk" if n_oos else None, delta_color="inverse")
c3.metric("Promo conflicts",   n_conf,   delta=f"{n_conf} conflicts" if n_conf else None, delta_color="inverse")
c4.metric("Total restock needed", n_restock)

# ── Info bars ─────────────────────────────────────────────────────────────────
st.info("ℹ️ Demand and OOS calculated **only** for Stock lock = Yes rows. Unlocked rows shown for reference only.")
st.warning("📍 Stock partitioned by **Seller + Country + Base SKU**. Same SKU in TH and MY is calculated independently.")

# ── SKU breakdown table ───────────────────────────────────────────────────────
st.markdown('### Base SKU breakdown')
st.caption("Sorted by Seller · Country · Base SKU · locked rows highlighted")

if df_rows.empty:
    st.info("No rows match current filters.")
else:
    display_df = df_rows.sort_values(["seller","country","base_sku","stock_lock"],
                                      ascending=[True,True,True,False]).copy()
    display_df["status_label"]   = display_df["status"].map(STATUS_EMOJI)
    display_df["mult_label"]     = display_df["mult"].apply(lambda m: f"×{m}")
    display_df["demand_disp"]    = display_df.apply(lambda r: r["demand"]       if r["stock_lock"] else "—", axis=1)
    display_df["total_res_disp"] = display_df.apply(lambda r: r["total_res"]    if r["stock_lock"] else "—", axis=1)
    display_df["total_dem_disp"] = display_df.apply(lambda r: r["total_demand"] if r["stock_lock"] else "—", axis=1)
    display_df["gap_disp"]       = display_df.apply(
        lambda r: int(r["gap"]) if pd.notna(r["gap"]) else "—", axis=1
    )

    show_cols = {
        "scope_label":    "Scope (Seller · Country · Base SKU)",
        "base_sku":       "Base SKU",
        "mult_label":     "Multiplier",
        "promo_sku":      "Promo SKU",
        "brand":          "Brand",
        "channel":        "Channel",
        "campaign":       "Campaign",
        "type":           "Type",
        "start":          "Start",
        "end":            "End",
        "stock_lock":     "Lock",
        "stock":          "Stock (shared)",
        "total_res_disp": "Orders reserved",
        "demand_disp":    "Base units demand",
        "total_dem_disp": "Total demand (all promos)",
        "gap_disp":       "Gap",
        "status_label":   "Status",
    }
    tbl = display_df[list(show_cols.keys())].rename(columns=show_cols)
    tbl["Lock"] = tbl["Lock"].map({True:"🔒 Yes", False:"🔓 No"})

    styled = (tbl.style
        .map(style_status, subset=["Status"])
        .map(style_gap,    subset=["Gap"])
        .set_properties(**{"font-family":"DM Mono, monospace","font-size":"11px"},
                         subset=["Scope (Seller · Country · Base SKU)","Base SKU","Promo SKU"])
    )
    st.dataframe(styled, use_container_width=True, hide_index=True, height=min(40 + 35*len(tbl), 500))

# ── Conflicts table ───────────────────────────────────────────────────────────
st.markdown('### Conflicting promotions — by Seller · Country · Base SKU')
st.caption("Only locked promotions within same seller+country scope flagged as conflicts")

if df_conf.empty:
    st.success("✅ No conflicts found in selected date range.")
else:
    conf_show = df_conf.sort_values("scope_key")[
        ["scope_label","brand","channel","stock","campaign_a","campaign_b",
         "overlap_start","overlap_end","combined_demand","verdict"]
    ].rename(columns={
        "scope_label":     "Scope (Seller · Country · Base SKU)",
        "brand":           "Brand",
        "channel":         "Channel",
        "stock":           "Stock (shared)",
        "campaign_a":      "Campaign A",
        "campaign_b":      "Campaign B",
        "overlap_start":   "Overlap start",
        "overlap_end":     "Overlap end",
        "combined_demand": "Combined demand",
        "verdict":         "Verdict",
    })
    styled_conf = conf_show.style.map(style_verdict, subset=["Verdict"])
    st.dataframe(styled_conf, use_container_width=True, hide_index=True)

# ── Heatmap ───────────────────────────────────────────────────────────────────
st.markdown('### OOS risk calendar — D-3 to D+7')
st.caption("Locked promotions only · today = D · hover for scope details")

hm_mode = st.radio("Intensity", ["Demand ratio %","Promo count","Units at risk"],
                   horizontal=True, label_visibility="collapsed")
render_heatmap(hm_df, hm_mode)

with st.expander("📋 View heatmap as table"):
    hm_show = hm_df[["day_label","date_str","active_scopes","total_demand",
                      "total_stock","ratio","units_at_risk","risk_level","scope_keys"]].copy()
    hm_show.columns = ["Day","Date","Active scopes","Total demand","Total stock",
                        "Demand ratio %","Units at risk","Risk level","Scope keys"]
    styled_hm = hm_show.style.map(style_risk, subset=["Risk level"])
    st.dataframe(styled_hm, use_container_width=True, hide_index=True)

# ── Restock ────────────────────────────────────────────────────────────────────
st.markdown('### Restock recommendations')
st.caption("Per Seller · Country · Base SKU scope · locked and combined (locked + nominated) views")
render_restock(df_rows if not df_rows.empty else pd.DataFrame())

# ── Export ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('### Export report')
export_bytes = build_excel(df_rows if not df_rows.empty else pd.DataFrame(),
                            df_conf if not df_conf.empty else pd.DataFrame(),
                            hm_df)
filename = f"OOS_Sentinel_Export_{TODAY.strftime('%Y-%m-%d')}.xlsx"
st.download_button(
    label="⬇ Download Excel report",
    data=export_bytes,
    file_name=filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary",
)
st.caption("4 sheets: SKU breakdown · Conflicts · Restock · Heatmap summary")

