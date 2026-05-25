import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date, timedelta, datetime
import re
import xlsxwriter

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Stock OOS Risk Tracker",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

/* Sidebar */
section[data-testid="stSidebar"] { background: #ffffff; border-right: 1px solid #e2e0d8; }
section[data-testid="stSidebar"] .block-container { padding: 1.5rem 1rem; }

/* Metric cards */
[data-testid="metric-container"] {
    background: #ffffff; border: 1px solid #e2e0d8;
    border-radius: 12px; padding: 1rem;
}

/* Info boxes */
.info-blue  { background:#eff4fe; border:1px solid #bfcff8; color:#1d4ed8; border-radius:8px; padding:10px 14px; font-size:13px; margin-bottom:8px; }
.info-purple{ background:#f5f3ff; border:1px solid #ddd6fe; color:#6d28d9; border-radius:8px; padding:10px 14px; font-size:13px; margin-bottom:8px; }

/* Section headers */
.section-hdr { font-size:11px; font-weight:600; text-transform:uppercase; letter-spacing:.08em; color:#9e9b91; margin:12px 0 6px; }

/* Scope pill */
.scope-pill { display:inline-block; font-family:'DM Mono',monospace; font-size:10px; padding:2px 7px; border-radius:4px; background:#f5f3ff; color:#6d28d9; border:1px solid #ddd6fe; }

/* Restock section labels */
.rs-label-hard { font-size:11px; font-weight:600; text-transform:uppercase; letter-spacing:.07em; color:#c0392b; margin:14px 0 6px; border-bottom:1px solid #f5c6c2; padding-bottom:4px; }
.rs-label-soft { font-size:11px; font-weight:600; text-transform:uppercase; letter-spacing:.07em; color:#b45309; margin:14px 0 6px; border-bottom:1px solid #f5dfa0; padding-bottom:4px; }

/* Hide streamlit branding */
#MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

TODAY = date.today()

# ─────────────────────────────────────────────────────────────────────────────
# SEED DATA  (used until a file is uploaded)
# ─────────────────────────────────────────────────────────────────────────────
SEED_PROMOS = [
    {"seller":"AAFHU","country":"TH","brand":"Oatside", "channel":"Lazada - Oatside",
     "sku":"101336151x8","campaign":"LazFlash Mid Month","type":"Flash Sale",
     "stock_lock":True,"start":"2026-05-15","end":"2026-05-20","reserved":24,"nominated":0,"stock":100},
    {"seller":"AAFHU","country":"TH","brand":"Oatside", "channel":"Lazada - Oatside",
     "sku":"101336151x8","campaign":"LazFlash Mid Month","type":"Flash Sale",
     "stock_lock":True,"start":"2026-05-23","end":"2026-05-25","reserved":40,"nominated":0,"stock":100},
    {"seller":"AAFHU","country":"TH","brand":"Hiruscar","channel":"TikTok-Hiruscar",
     "sku":"101336151+101336152x6","campaign":"TikTok Live Promotion","type":"Flash Sale",
     "stock_lock":True,"start":"2026-05-20","end":"2026-05-25","reserved":40,"nominated":0,"stock":100},
    {"seller":"AAFHU","country":"TH","brand":"Hiruscar","channel":"TikTok-Hiruscar",
     "sku":"101336151+101336152x6","campaign":"TikTok Live Promotion","type":"Voucher Promo",
     "stock_lock":False,"start":"2026-05-20","end":"2026-05-25","reserved":0,"nominated":10,"stock":100},
]

# ─────────────────────────────────────────────────────────────────────────────
# SKU PARSING
# ─────────────────────────────────────────────────────────────────────────────
def parse_sku(sku: str) -> list[dict]:
    """Parse 'A+Bx6+Cx2' → [{'base':'A','mult':1},{'base':'B','mult':6},{'base':'C','mult':2}]"""
    components = []
    for part in str(sku).split("+"):
        m = re.match(r"^(\w+)(?:x(\d+))?$", part.strip())
        if m:
            components.append({"base": m.group(1), "mult": int(m.group(2)) if m.group(2) else 1})
    return components

def scope_key(seller, country, base_sku):
    return f"{seller}|{country}|{base_sku}"

# ─────────────────────────────────────────────────────────────────────────────
# XLSX UPLOAD PARSER
# ─────────────────────────────────────────────────────────────────────────────
def parse_upload(file) -> list[dict]:
    df = pd.read_excel(file, sheet_name=0, dtype=str)
    df.columns = [c.strip() for c in df.columns]

    col_map = {
        "seller":     ["Seller code","Seller Code","seller"],
        "country":    ["Country","country"],
        "brand":      ["Brand","brand"],
        "channel":    ["Channel","channel"],
        "sku":        ["SKU","sku","Promo SKU"],
        "campaign":   ["Campaign / Promotion Name","Campaign Name","campaign"],
        "type":       ["Campaign Type","Campaign type","type"],
        "stock_lock": ["Stock Lock / Reserved","Stock lock","stock_lock"],
        "start":      ["Promo Start Date","Start Date","start"],
        "end":        ["Promo End Date","End Date","end"],
        "stock":      ["Today's Stock for Base SKU - (24 may)","Today's Stock","Stock","stock"],
        "reserved_dksh":  ["Reserved by DKSH","reserved_dksh"],
        "reserved_graas": ["Reserved by Graas","reserved_graas"],
        "reserved_mp":    ["Reserved by MP","reserved_mp"],
        "nominated":  ["Nominated stock (Non Reservation)","Nominated Stock","nominated"],
    }

    def find_col(key):
        for name in col_map[key]:
            if name in df.columns:
                return name
        return None

    promos = []
    last = {}
    for _, row in df.iterrows():
        def g(key, default=None):
            col = find_col(key)
            val = row.get(col, default) if col else default
            return val if pd.notna(val) and str(val).strip() not in ("","nan","None") else default

        seller   = g("seller")   or last.get("seller")
        country  = g("country")  or last.get("country")
        brand    = g("brand")    or last.get("brand")
        channel  = g("channel")  or last.get("channel")
        sku      = g("sku")      or last.get("sku")
        campaign = g("campaign") or last.get("campaign")
        typ      = g("type")     or last.get("type")

        start_raw = g("start")
        end_raw   = g("end")
        try:
            start = pd.to_datetime(start_raw).strftime("%Y-%m-%d") if start_raw else last.get("start","")
        except Exception:
            start = last.get("start","")
        try:
            end = pd.to_datetime(end_raw).strftime("%Y-%m-%d") if end_raw else last.get("end","")
        except Exception:
            end = last.get("end","")

        lock_raw = str(g("stock_lock") or "").strip().lower()
        stock_lock = lock_raw in ("yes","true","1","y")

        try: stock = float(g("stock") or 0)
        except: stock = 0

        try:
            reserved = sum(float(g(k) or 0) for k in ["reserved_dksh","reserved_graas","reserved_mp"])
        except: reserved = 0

        try: nominated = float(g("nominated") or 0)
        except: nominated = 0

        if sku and start and seller and country:
            last = dict(seller=seller,country=country,brand=brand,channel=channel,
                        sku=sku,campaign=campaign,type=typ,start=start,end=end)
            promos.append(dict(seller=seller,country=country,brand=brand,channel=channel,
                               sku=sku,campaign=campaign,type=typ,stock_lock=stock_lock,
                               start=start,end=end,reserved=reserved,nominated=nominated,stock=stock))
    return promos

# ─────────────────────────────────────────────────────────────────────────────
# CORE CALCULATIONS  — all scoped to seller + country + base_sku
# ─────────────────────────────────────────────────────────────────────────────
def compute_rows(promos: list[dict]) -> pd.DataFrame:
    rows = []
    for p in promos:
        for comp in parse_sku(p["sku"]):
            base  = comp["base"]
            mult  = comp["mult"]
            stock = p["stock"]
            demand    = int(p["reserved"] * mult) if p["stock_lock"] else 0
            gap       = (stock - demand) if p["stock_lock"] else None
            oos       = bool(p["stock_lock"] and gap is not None and gap < 0)
            restock   = abs(gap) if oos else 0
            rows.append({
                "seller":     p["seller"],
                "country":    p["country"],
                "brand":      p["brand"],
                "channel":    p["channel"],
                "scope_key":  scope_key(p["seller"], p["country"], base),
                "base_sku":   base,
                "promo_sku":  p["sku"],
                "mult":       mult,
                "campaign":   p["campaign"],
                "type":       p["type"],
                "start":      p["start"],
                "end":        p["end"],
                "stock_lock": p["stock_lock"],
                "stock":      int(stock),
                "reserved":   int(p["reserved"]),
                "nominated":  int(p["nominated"]),
                "demand":     demand,
                "gap":        gap,
                "oos":        oos,
                "restock":    restock,
                "status":     ("OOS" if oos else
                               "Watch" if (p["stock_lock"] and gap is not None and gap < 20) else
                               "Safe" if p["stock_lock"] else "No lock"),
            })
    return pd.DataFrame(rows) if rows else pd.DataFrame()

def compute_conflicts(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    locked = df[df["stock_lock"]].copy()
    conflicts = []
    for scope, grp in locked.groupby("scope_key"):
        rows = grp.to_dict("records")
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
            ws1.write(ri, 0,  r["seller"],    rf)
            ws1.write(ri, 1,  r["country"],   rf)
            ws1.write(ri, 2,  r["brand"],     rf)
            ws1.write(ri, 3,  r["channel"],   rf)
            ws1.write(ri, 4,  r["scope_key"], rf_mono)
            ws1.write(ri, 5,  r["base_sku"],  rf_mono)
            ws1.write(ri, 6,  r["promo_sku"], rf_mono)
            ws1.write(ri, 7,  r["campaign"],  rf)
            ws1.write(ri, 8,  r["type"],      rf)
            ws1.write(ri, 9,  r["start"],     rf)
            ws1.write(ri, 10, r["end"],       rf)
            ws1.write(ri, 11, "Yes" if r["stock_lock"] else "No", rf)
            ws1.write(ri, 12, r["stock"],     rf_num)
            ws1.write(ri, 13, r["reserved"] if r["stock_lock"] else "—", rf_num)
            ws1.write(ri, 14, r["demand"]   if r["stock_lock"] else "—", rf_num)
            ws1.write(ri, 15, r["gap"]      if r["gap"] is not None else "—", rf_num)
            ws1.write(ri, 16, st_,           rf)

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
            ws2.write(ri, 0,  r["scope_key"],       cf_mono)
            ws2.write(ri, 1,  r["seller"],           cf)
            ws2.write(ri, 2,  r["country"],          cf)
            ws2.write(ri, 3,  r["brand"],            cf)
            ws2.write(ri, 4,  r["channel"],          cf)
            ws2.write(ri, 5,  r["stock"],            cf_num)
            ws2.write(ri, 6,  r["campaign_a"],       cf)
            ws2.write(ri, 7,  r["campaign_b"],       cf)
            ws2.write(ri, 8,  r["overlap_start"],    cf)
            ws2.write(ri, 9,  r["overlap_end"],      cf)
            ws2.write(ri, 10, r["combined_demand"],  cf_num)
            ws2.write(ri, 11, r["verdict"],          cf)

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
        ws3.write(ri, 0,  row[0],  rf)
        ws3.write(ri, 1,  row[1],  rfm)
        ws3.write(ri, 2,  row[2],  rf)
        ws3.write(ri, 3,  row[3],  rf)
        ws3.write(ri, 4,  row[4],  rf)
        ws3.write(ri, 5,  row[5],  rfm)
        ws3.write(ri, 6,  row[6],  rfn)
        ws3.write(ri, 7,  row[7],  rfn)
        ws3.write(ri, 8,  row[8],  rfn)
        ws3.write(ri, 9,  row[9],  rfn)
        ws3.write(ri, 10, row[10], rfn)
        ws3.write(ri, 11, row[11], rfn)

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
        ws4.write(ri, 0, r["day_label"],     rf)
        ws4.write(ri, 1, r["date_str"],      rf)
        ws4.write(ri, 2, r["active_scopes"], rfn)
        ws4.write(ri, 3, r["total_demand"],  rfn)
        ws4.write(ri, 4, r["total_stock"],   rfn)
        ws4.write(ri, 5, str(r["ratio"])+"%",rfn)
        ws4.write(ri, 6, r["units_at_risk"], rfn)
        ws4.write(ri, 7, r["risk_level"],    rf)
        ws4.write(ri, 8, r["scope_keys"],    rf)

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
    RISK_BG   = {"High":"#FDF0EE","Medium":"#FEF9EC","Low":"#EDF7F2","None":"#F5F4F0"}
    RISK_TC   = {"High":"#C0392B","Medium":"#B45309","Low":"#1A6B3C","None":"#9E9B91"}
    RISK_BC   = {"High":"#F5C6C2","Medium":"#F5DFA0","Low":"#B5DFC8","None":"#E2E0D8"}

    if mode == "Demand ratio %": vals = hm_df["ratio"]
    elif mode == "Promo count":  vals = hm_df["active_scopes"]
    else:                        vals = hm_df["units_at_risk"]

    cards_html = '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px">'
    for (_, row), val in zip(hm_df.iterrows(), vals):
        rl   = row["risk_level"]
        bg   = RISK_BG.get(rl, "#F5F4F0")
        tc   = RISK_TC.get(rl, "#9E9B91")
        bc   = RISK_BC.get(rl, "#E2E0D8")
        disp = f"{val:.0f}%" if mode == "Demand ratio %" else str(int(val))
        today_border = f"border:2.5px solid #1a1917!important;" if row["is_today"] else f"border:1px solid {bc};"
        cards_html += f"""
        <div title="{row['scope_keys']}" style="width:66px;height:76px;border-radius:8px;
            background:{bg};{today_border}display:flex;flex-direction:column;
            align-items:center;justify-content:center;gap:2px;cursor:default;">
          <span style="font-size:10px;font-weight:600;color:{tc}">{row['day_label']}</span>
          <span style="font-size:9px;color:{tc};opacity:.75">{row['date_str']}</span>
          <span style="font-size:15px;font-weight:600;color:{tc};margin-top:2px">{disp}</span>
          <span style="font-size:9px;color:{tc};opacity:.65">{int(row['active_scopes'])} scope{'s' if row['active_scopes']!=1 else ''}</span>
        </div>"""
    cards_html += "</div>"

    legend_html = """<div style="display:flex;gap:14px;flex-wrap:wrap;font-size:11px;color:#6b6860">
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#F5F4F0;border:1px solid #e2e0d8;margin-right:4px"></span>No promos</span>
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#EDF7F2;border:1px solid #b5dfc8;margin-right:4px"></span>Low</span>
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#FEF9EC;border:1px solid #f5dfa0;margin-right:4px"></span>Medium</span>
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#FDF0EE;border:1px solid #f5c6c2;margin-right:4px"></span>High</span>
      <span><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:white;border:2px solid #1a1917;margin-right:4px"></span>Today (D)</span>
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
    st.markdown('<div class="rs-label-hard">🔒 Hard OOS — locked stock only</div>', unsafe_allow_html=True)
    hard_by_scope = {}
    for _, r in df_rows[df_rows["oos"]].iterrows():
        k = r["scope_key"]
        if k not in hard_by_scope or hard_by_scope[k]["restock"] < r["restock"]:
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
                <div style="font-size:11px;color:#C0392B;opacity:.8;margin-top:2px">{r['brand']} · {r['channel']} · {r['campaign']} · stock {r['stock']}, locked demand {r['demand']}</div>
              </div>
              <div style="font-size:16px;font-weight:600;color:#C0392B;margin-left:16px;white-space:nowrap">+{r['restock']} units</div>
            </div>""", unsafe_allow_html=True)

    # ── Soft Risk ─────────────────────────────────────────────────────────────
    st.markdown('<div class="rs-label-soft">⚠️ Soft risk — locked + nominated combined</div>', unsafe_allow_html=True)
    scope_agg = {}
    for _, r in df_rows.iterrows():
        k = r["scope_key"]
        if k not in scope_agg:
            scope_agg[k] = {"locked":0,"nominated":0,"stock":r["stock"],"base_sku":r["base_sku"],
                             "seller":r["seller"],"country":r["country"],"brand":r["brand"],"channel":r["channel"]}
        if r["stock_lock"]: scope_agg[k]["locked"]   += r["demand"]
        else:               scope_agg[k]["nominated"] += r["nominated"]

    soft_items = [(k, s) for k, s in scope_agg.items()
                  if s["nominated"] > 0 and s["stock"] - s["locked"] - s["nominated"] < 0]

    if not soft_items:
        st.caption("No soft risk when nominated stock is included.")
    else:
        for k, s in soft_items:
            combined  = s["locked"] + s["nominated"]
            gap_soft  = s["stock"] - combined
            extra     = abs(gap_soft)
            st.markdown(f"""
            <div style="display:flex;align-items:center;justify-content:space-between;
                padding:11px 14px;background:#FEF9EC;border:1px solid #F5DFA0;
                border-radius:8px;margin-bottom:6px">
              <div>
                <div style="font-family:'DM Mono',monospace;font-size:12px;font-weight:500;color:#B45309">{s['base_sku']}</div>
                <div style="font-size:10px;color:#B45309;opacity:.7;font-family:'DM Mono',monospace">{s['seller']} · {s['country']}</div>
                <div style="font-size:11px;color:#B45309;opacity:.8;margin-top:2px">{s['brand']} · {s['channel']} · stock {s['stock']} · locked {s['locked']} + nominated {s['nominated']} = combined {combined}</div>
              </div>
              <div style="font-size:16px;font-weight:600;color:#B45309;margin-left:16px;white-space:nowrap">+{extra} units</div>
            </div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📦 OOS Risk Tracker")
    st.caption("Stock scoped per Seller · Country · Base SKU")
    st.divider()

    # File upload
    st.markdown('<div class="section-hdr">Data source</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload sheet", type=["xlsx","csv"], label_visibility="collapsed")

    if uploaded:
        try:
            promos = parse_upload(uploaded)
            st.success(f"✓ {uploaded.name}  ·  {len(promos)} rows loaded")
        except Exception as e:
            st.error(f"Parse error: {e}")
            promos = SEED_PROMOS
    else:
        promos = SEED_PROMOS
        st.info("Using sample data — upload your sheet above.")

    st.divider()

    # Date range
    st.markdown('<div class="section-hdr">Date range</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        d_from = st.date_input("From", value=TODAY - timedelta(days=9), label_visibility="collapsed")
    with col2:
        d_to   = st.date_input("To",   value=TODAY + timedelta(days=16), label_visibility="collapsed")
    st.caption(f"From {d_from.strftime('%d %b')} → {d_to.strftime('%d %b %Y')}")

    st.divider()

    # Dimension chip-style filters using multiselect
    all_sellers  = sorted(set(p["seller"]  for p in promos))
    all_countries= sorted(set(p["country"] for p in promos))
    all_brands   = sorted(set(p["brand"]   for p in promos))
    all_channels = sorted(set(p["channel"] for p in promos))
    all_types    = sorted(set(p["type"]    for p in promos))

    st.markdown('<div class="section-hdr">Seller code</div>', unsafe_allow_html=True)
    sel_sellers  = st.multiselect("Seller",  all_sellers,  default=all_sellers,  label_visibility="collapsed")

    st.markdown('<div class="section-hdr">Country</div>', unsafe_allow_html=True)
    sel_countries= st.multiselect("Country", all_countries,default=all_countries,label_visibility="collapsed")

    st.markdown('<div class="section-hdr">Brand</div>', unsafe_allow_html=True)
    sel_brands   = st.multiselect("Brand",   all_brands,   default=all_brands,   label_visibility="collapsed")

    st.markdown('<div class="section-hdr">Channel</div>', unsafe_allow_html=True)
    sel_channels = st.multiselect("Channel", all_channels, default=all_channels, label_visibility="collapsed")

    st.divider()

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

    st.divider()

    st.markdown('<div class="section-hdr">More filters</div>', unsafe_allow_html=True)
    sel_type = st.selectbox("Campaign type", ["All types"] + all_types, label_visibility="collapsed")
    sel_lock = st.selectbox("Stock lock",    ["All","Locked only","Unlocked only"], label_visibility="collapsed")
    sel_show = st.selectbox("Show",          ["All SKUs","OOS risk only","Safe stock only"], label_visibility="collapsed")

    st.divider()
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

df_rows = compute_rows(filtered_promos)
if not df_rows.empty and sel_show == "OOS risk only":
    df_rows = df_rows[df_rows["oos"]]
elif not df_rows.empty and sel_show == "Safe stock only":
    df_rows = df_rows[~df_rows["oos"] & df_rows["stock_lock"]]

df_conf = compute_conflicts(df_rows) if not df_rows.empty else pd.DataFrame()
hm_df   = compute_heatmap(df_rows, TODAY)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN CONTENT
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("## 📦 Stock OOS Risk Tracker")
st.caption("Stock scoped per **Seller + Country** — same SKU in different countries is treated as separate stock. "
           "Demand & OOS calculated only for **Stock lock = Yes** rows.")

# ── Metrics ──────────────────────────────────────────────────────────────────
locked_rows = df_rows[df_rows["stock_lock"]] if not df_rows.empty else pd.DataFrame()
n_scopes    = locked_rows["scope_key"].nunique()     if not locked_rows.empty else 0
n_oos       = locked_rows[locked_rows["oos"]]["scope_key"].nunique() if not locked_rows.empty else 0
n_conf      = len(df_conf)
n_restock   = int(locked_rows[locked_rows["oos"]]["restock"].sum()) if not locked_rows.empty else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Stock scopes",      n_scopes)
c2.metric("OOS risk scopes",   n_oos,    delta=f"{n_oos} at risk" if n_oos else None, delta_color="inverse")
c3.metric("Promo conflicts",   n_conf,   delta=f"{n_conf} conflicts" if n_conf else None, delta_color="inverse")
c4.metric("Total restock needed", n_restock)

# ── Info bars ─────────────────────────────────────────────────────────────────
st.markdown('<div class="info-blue">ℹ️ Demand and OOS are calculated <b>only</b> for Stock lock = Yes rows. Unlocked rows shown for reference only.</div>', unsafe_allow_html=True)
st.markdown('<div class="info-purple">📍 Stock is partitioned by <b>Seller + Country + Base SKU</b>. The same SKU for the same seller in TH and MY is calculated independently.</div>', unsafe_allow_html=True)

# ── SKU breakdown table ───────────────────────────────────────────────────────
st.markdown("### Base SKU breakdown")
st.caption("Sorted by Seller · Country · Base SKU · locked rows highlighted")

if df_rows.empty:
    st.info("No rows match current filters.")
else:
    display_df = df_rows.sort_values(["seller","country","base_sku","stock_lock"],
                                      ascending=[True,True,True,False]).copy()
    display_df["status_label"] = display_df["status"].map(STATUS_EMOJI)
    display_df["demand_disp"]  = display_df.apply(lambda r: r["demand"] if r["stock_lock"] else "—", axis=1)
    display_df["gap_disp"]     = display_df.apply(lambda r: r["gap"] if r["gap"] is not None else "—", axis=1)
    display_df["reserved_disp"]= display_df.apply(lambda r: r["reserved"] if r["stock_lock"] else "—", axis=1)

    show_cols = {
        "scope_key":      "Scope (Seller·Country·SKU)",
        "brand":          "Brand",
        "channel":        "Channel",
        "promo_sku":      "Promo SKU",
        "campaign":       "Campaign",
        "type":           "Type",
        "start":          "Start",
        "end":            "End",
        "stock_lock":     "Lock",
        "stock":          "Stock",
        "demand_disp":    "Demand",
        "reserved_disp":  "Reserved",
        "gap_disp":       "Gap",
        "status_label":   "Status",
    }
    tbl = display_df[list(show_cols.keys())].rename(columns=show_cols)
    tbl["Lock"] = tbl["Lock"].map({True:"🔒 Yes", False:"🔓 No"})

    styled = (tbl.style
        .applymap(style_status, subset=["Status"])
        .applymap(style_gap,    subset=["Gap"])
        .set_properties(**{"font-family":"DM Mono, monospace","font-size":"11px"},
                         subset=["Scope (Seller·Country·SKU)","Promo SKU"])
    )
    st.dataframe(styled, use_container_width=True, hide_index=True, height=min(40 + 35*len(tbl), 500))

# ── Conflicts table ───────────────────────────────────────────────────────────
st.markdown("### Conflicting promotions — by Seller · Country · Base SKU")
st.caption("Only locked promotions within same seller+country scope flagged as conflicts")

if df_conf.empty:
    st.success("✅ No conflicts found in selected date range.")
else:
    conf_show = df_conf.sort_values("scope_key")[
        ["scope_key","brand","channel","stock","campaign_a","campaign_b",
         "overlap_start","overlap_end","combined_demand","verdict"]
    ].rename(columns={
        "scope_key":       "Scope key",
        "brand":           "Brand",
        "channel":         "Channel",
        "stock":           "Stock",
        "campaign_a":      "Campaign A",
        "campaign_b":      "Campaign B",
        "overlap_start":   "Overlap start",
        "overlap_end":     "Overlap end",
        "combined_demand": "Combined demand",
        "verdict":         "Verdict",
    })
    styled_conf = conf_show.style.applymap(style_verdict, subset=["Verdict"])
    st.dataframe(styled_conf, use_container_width=True, hide_index=True)

# ── Heatmap ───────────────────────────────────────────────────────────────────
st.markdown("### OOS risk calendar — D-3 to D+7")
st.caption("Locked promotions only · today = D · hover for scope details")

hm_mode = st.radio("Intensity", ["Demand ratio %","Promo count","Units at risk"],
                   horizontal=True, label_visibility="collapsed")
render_heatmap(hm_df, hm_mode)

with st.expander("📋 View heatmap as table"):
    hm_show = hm_df[["day_label","date_str","active_scopes","total_demand",
                      "total_stock","ratio","units_at_risk","risk_level","scope_keys"]].copy()
    hm_show.columns = ["Day","Date","Active scopes","Total demand","Total stock",
                        "Demand ratio %","Units at risk","Risk level","Scope keys"]
    styled_hm = hm_show.style.applymap(style_risk, subset=["Risk level"])
    st.dataframe(styled_hm, use_container_width=True, hide_index=True)

# ── Restock ────────────────────────────────────────────────────────────────────
st.markdown("### Restock recommendations")
st.caption("Per Seller · Country · Base SKU scope · locked and combined (locked + nominated) views")
render_restock(df_rows if not df_rows.empty else pd.DataFrame())

# ── Export ────────────────────────────────────────────────────────────────────
st.divider()
st.markdown("### Export")
export_bytes = build_excel(df_rows if not df_rows.empty else pd.DataFrame(),
                            df_conf if not df_conf.empty else pd.DataFrame(),
                            hm_df)
filename = f"OOS_Risk_Export_{TODAY.strftime('%Y-%m-%d')}.xlsx"
st.download_button(
    label="⬇ Download Excel report",
    data=export_bytes,
    file_name=filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary",
)
st.caption(f"Exports 4 sheets: SKU breakdown · Conflicts · Restock · Heatmap summary")

