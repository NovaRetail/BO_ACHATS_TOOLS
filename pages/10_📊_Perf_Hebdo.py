"""
10_📊_Perf_Hebdo.py — SmartBuyer Hub [v3.0]
Performance commerciale hebdomadaire · Charte SmartBuyer v2
🔧 Fix v3.0 : Suppression des doublons article×magasin
   - Filtre arts sur Site nom long == "Total" (une seule ligne par article = total réseau)
   - Colonnes PBI : "CA Promo" (pas "CA HT Promo")
   - Colonnes optionnelles sécurisées
"""

import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Perf Hebdo · SmartBuyer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CHARTE SMARTBUYER ────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display",
                 "SF Pro Text", "Helvetica Neue", Arial, sans-serif !important;
    background-color: #F2F2F7;
}
.stApp { background: #F2F2F7; }
.main .block-container { padding-top: 1.8rem; max-width: 1200px; }

[data-testid="stSidebar"] {
    background: #F2F2F7 !important;
    border-right: 0.5px solid #D1D1D6 !important;
}
[data-testid="stMetric"] {
    background: #FFFFFF !important;
    border: 0.5px solid #E5E5EA !important;
    border-radius: 12px !important;
    padding: 16px 18px !important;
}
[data-testid="stMetricLabel"]  { font-size: 11px !important; font-weight: 500 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stMetricValue"]  { font-size: 24px !important; font-weight: 600 !important; color: #1C1C1E !important; letter-spacing: -0.02em !important; }
[data-testid="stMetricDelta"]  { font-size: 12px !important; }

[data-testid="stTabs"] button[role="tab"] { font-size: 13px !important; font-weight: 500 !important; padding: 8px 16px !important; color: #8E8E93 !important; border-radius: 0 !important; border-bottom: 2px solid transparent !important; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color: #007AFF !important; border-bottom: 2px solid #007AFF !important; background: transparent !important; }
[data-testid="stTabs"] [role="tablist"] { border-bottom: 0.5px solid #E5E5EA !important; }

[data-testid="stDataFrame"] { border: 0.5px solid #E5E5EA !important; border-radius: 10px !important; }
[data-testid="stDataFrame"] th { background: #F2F2F7 !important; font-size: 11px !important; font-weight: 600 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }

[data-testid="stFileUploader"] { border: 1.5px dashed #D1D1D6 !important; border-radius: 10px !important; background: #F9F9FB !important; }
[data-testid="baseButton-primary"] { background: #007AFF !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; }

.stDownloadButton > button {
    background: #007AFF !important; color: white !important; border: none !important;
    border-radius: 8px !important; font-weight: 500 !important; font-size: 13px !important;
    padding: 10px 24px !important; width: 100% !important;
}
hr { border-color: #E5E5EA !important; margin: 1rem 0 !important; }

.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 1.5rem; }
.section-label { font-size: 11px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 10px; }

.alert-card { padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; font-size: 13px; line-height: 1.5; border-left: 3px solid; }
.alert-red   { background: #FFF2F2; border-color: #FF3B30; color: #3A0000; }
.alert-amber { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }
.alert-green { background: #F0FFF4; border-color: #34C759; color: #003A10; }
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }

.col-required { background: #F0F8FF; border: 0.5px solid #B3D9FF; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; display: flex; align-items: flex-start; gap: 10px; }
.col-name    { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc    { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
.col-example { font-size: 11px; color: #8E8E93; font-family: monospace; margin-top: 2px; }

.kpi-rayon      { background: #FFFFFF; border: 0.5px solid #E5E5EA; border-radius: 12px; padding: 14px 16px; }
.kpi-rayon-name { font-size: 12px; font-weight: 700; margin-bottom: 6px; }
.kpi-rayon-ca   { font-size: 20px; font-weight: 700; margin-bottom: 2px; }
.kpi-rayon-sub  { font-size: 11px; color: #8E8E93; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTES ───────────────────────────────────────────────────────────────
RAYON_MAP = {
    "00014 - EPICERIE":           "Épicerie",
    "00010 - BOISSONS":           "Boissons",
    "00012 - PARFUMERIE HYGIENE": "DPH",
    "00011 - DROGUERIE":          "DPH",
}
COLORS = {
    "Épicerie": "#FF9500",
    "Boissons": "#007AFF",
    "DPH":      "#AF52DE",
}
RED   = "#FF3B30"
GREEN = "#34C759"
AMBER = "#FF9500"
PL_FONT = dict(family="-apple-system, Helvetica Neue, Arial", color="#3A3A3C", size=12)

REQUIRED_COLS = {
    "Rayon":           ("Rayon de l'article",               "ex: 00014 - EPICERIE"),
    "Famille":         ("Famille de l'article",             "ex: 00147 - CONDIMENT, ASSAISONNEMENT"),
    "Article":         ("Code et libellé article",          "ex: 14006584 - 5KG RIZ RIZIERE"),
    "Site nom long":   ("Niveau d'agrégation",              "ex: Total (ligne réseau) ou 10301 - Hyper Marcory"),
    "CA":              ("Chiffre d'affaires HT (FCFA)",     "ex: 17 405 450"),
    "Marge":           ("Marge brute HT (FCFA)",            "ex: 1 098 943"),
    "%Marge":          ("Taux de marge (%)",                "ex: 0.0631"),
    "CA Promo":        ("CA réalisé sous promotion (FCFA)", "ex: 9 756 175"),
    "Marge Promo":     ("Marge sur ventes en promotion",    "ex: 80 769"),
    "Qté Vente":       ("Quantité vendue sur la période",   "ex: 6 263"),
    "Casse (Valeur)":  ("Valeur de la casse (FCFA)",        "ex: -183 976"),
}

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def fmt_fcfa(n):
    if pd.isna(n): return "—"
    a = abs(n)
    if a >= 1_000_000: return f"{n/1_000_000:.1f} M"
    if a >= 1_000:     return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def alert_html(level, title, action):
    cls = {"🔴": "alert-red", "🟡": "alert-amber", "🟢": "alert-green"}.get(level, "alert-blue")
    ico = {"🔴": "⚠️", "🟡": "⚠️", "🟢": "✅"}.get(level, "ℹ️")
    return (f'<div class="alert-card {cls}"><strong>{ico} {title}</strong><br>'
            f'<span style="font-size:12px;opacity:.85">→ {action}</span></div>')

def clean_label(s):
    if pd.isna(s): return ""
    m = re.match(r"^\d+ - (.+)$", str(s))
    return m.group(1) if m else str(s)

# ─── PARSING ──────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_file(file_bytes, filename):
    if filename.endswith(".csv"):
        df = pd.read_csv(BytesIO(file_bytes), encoding="latin-1")
    else:
        df = pd.read_excel(BytesIO(file_bytes), engine="openpyxl")
    df.columns = df.columns.str.strip()

    # ── Fix v3.0 : garder uniquement les lignes article-niveau réseau total ──
    # Le fichier PBI contient une ligne par article×magasin + une ligne Total par article.
    # On filtre sur Site nom long == "Total" pour ne conserver que le total réseau
    # et éviter les doublons dans les classements.
    if "Site nom long" in df.columns:
        arts = df[
            df["Article"].notna() &
            (df["Article"] != "Total") &
            (df["Site nom long"] == "Total")
        ].copy()
    else:
        # Fallback si colonne absente (format alternatif)
        arts = df[df["Article"].notna() & (df["Article"] != "Total")].copy()

    arts["art_label"]   = arts["Article"].apply(clean_label)
    arts["rayon_label"] = arts["Rayon"].apply(
        lambda x: RAYON_MAP.get(str(x).strip(), clean_label(x))
    )

    # Normaliser "CA Promo" → "CA HT Promo" pour compatibilité interne
    if "CA Promo" in arts.columns and "CA HT Promo" not in arts.columns:
        arts = arts.rename(columns={"CA Promo": "CA HT Promo"})

    num_cols = ["CA", "Marge", "%Marge", "CA HT Promo", "Marge Promo",
                "%CA Poids Promo", "Qté Vente", "Casse (Valeur)", "Casse (Qté)"]
    for col in num_cols:
        if col in arts.columns:
            arts[col] = pd.to_numeric(arts[col], errors="coerce").fillna(0)

    # Assurer l'existence des colonnes optionnelles
    for col in ["CA HT Promo", "Marge Promo", "%CA Poids Promo", "Casse (Valeur)", "Casse (Qté)"]:
        if col not in arts.columns:
            arts[col] = 0

    # ── Totaux rayon : lignes Famille == "Total", Site == NaN (vrais totaux rayon) ──
    rayon_tots = df[df["Famille"] == "Total"].copy()
    rayon_tots["rayon_label"] = rayon_tots["Rayon"].apply(
        lambda x: RAYON_MAP.get(str(x).strip(), clean_label(x))
    )
    for col in ["CA", "Marge", "Casse (Valeur)"]:
        if col in rayon_tots.columns:
            rayon_tots[col] = pd.to_numeric(rayon_tots[col], errors="coerce").fillna(0)
    if "Casse (Valeur)" not in rayon_tots.columns:
        rayon_tots["Casse (Valeur)"] = 0

    rayon_tots = rayon_tots.groupby("rayon_label", as_index=False).agg(
        CA=("CA", "sum"),
        Marge=("Marge", "sum"),
        Casse=("Casse (Valeur)", "sum"),
    )
    rayon_tots["%Marge"] = rayon_tots["Marge"] / rayon_tots["CA"].replace(0, np.nan)

    return arts, rayon_tots

# ─── CALCULS ──────────────────────────────────────────────────────────────────
def compute_kpis(arts):
    ca     = arts["CA"].sum()
    marge  = arts["Marge"].sum()
    casse  = arts["Casse (Valeur)"].sum() if "Casse (Valeur)" in arts.columns else 0
    nb_neg  = int((arts["Marge"] < 0).sum())
    nb_cs   = int((arts["Casse (Valeur)"] < 0).sum()) if "Casse (Valeur)" in arts.columns else 0
    return ca, marge, marge / ca if ca else 0, casse, nb_neg, nb_cs

def top_ca(arts, n=10):
    return (arts.nlargest(n, "CA")
            [["art_label", "rayon_label", "CA", "Marge", "%Marge", "Qté Vente"]]
            .reset_index(drop=True))

def top_marge(arts, n=10):
    return (arts.nlargest(n, "Marge")
            [["art_label", "rayon_label", "CA", "Marge", "%Marge"]]
            .reset_index(drop=True))

def top_promo(arts, n=10):
    if "CA HT Promo" not in arts.columns or arts["CA HT Promo"].max() <= 0:
        return pd.DataFrame()
    return (arts[arts["CA HT Promo"] > 0]
            .nlargest(n, "CA HT Promo")
            [["art_label", "rayon_label", "CA HT Promo", "Marge Promo", "%CA Poids Promo"]]
            .reset_index(drop=True))

def flop_marge(arts, n=15):
    return (arts[arts["Marge"] < 0]
            .nsmallest(n, "Marge")
            [["art_label", "rayon_label", "CA", "Marge", "%Marge"]]
            .reset_index(drop=True))

def top_casse(arts, n=10):
    if "Casse (Valeur)" not in arts.columns:
        return pd.DataFrame()
    return (arts[arts["Casse (Valeur)"] < 0]
            .nsmallest(n, "Casse (Valeur)")
            [["art_label", "rayon_label", "Casse (Valeur)", "Casse (Qté)"]]
            .reset_index(drop=True))

# ─── AFFICHAGE DATAFRAME ──────────────────────────────────────────────────────
def show_df(df, rename_map, num_cols=(), pct_cols=(), neg_cols=(), green_cols=()):
    df = df.rename(columns=rename_map).copy()
    df.index = range(1, len(df) + 1)

    # Cast int64 pour supprimer les décimales
    for col in [rename_map.get(c, c) for c in num_cols] + ["Qté vendue", "Casse qté"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).round(0).astype("int64")

    fmt = {}
    for c in num_cols:
        c2 = rename_map.get(c, c)
        if c2 in df.columns:
            fmt[c2] = lambda v, _c=c2: f"{int(v):,}".replace(",", "\u202f")
    for c in pct_cols:
        c2 = rename_map.get(c, c)
        if c2 in df.columns:
            fmt[c2] = "{:.1%}"

    style = df.style
    if fmt:
        style = style.format(fmt, na_rep="—")
    for c in neg_cols:
        c2 = rename_map.get(c, c)
        if c2 in df.columns:
            style = style.map(
                lambda v: "color:#FF3B30;font-weight:600"
                if isinstance(v, (int, float)) and v < 0 else "",
                subset=[c2])
    for c in green_cols:
        c2 = rename_map.get(c, c)
        if c2 in df.columns:
            style = style.map(
                lambda v: "color:#34C759;font-weight:600"
                if isinstance(v, (int, float)) and v > 0 else "",
                subset=[c2])
    st.dataframe(style, use_container_width=True, height=400, hide_index=False)

# ─── HELPERS EXCEL ────────────────────────────────────────────────────────────
def hdr_fill(h):
    return PatternFill("solid", fgColor=h.replace("#", ""))

def mk_border(s="thin", c="D1D1D6"):
    x = Side(style=s, color=c)
    return Border(left=x, right=x, top=x, bottom=x)

def mk_bottom(c="D1D1D6"):
    return Border(bottom=Side(style="thin", color=c))

def xfnt(bold=False, color="1C1C1E", size=10):
    return Font(bold=bold, color=color.replace("#", ""), size=size, name="Arial")

def xaln(h="left", v="center"):
    return Alignment(horizontal=h, vertical=v)

def xl_hdr(ws, row, col, title, color, ncols):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + ncols - 1)
    c = ws.cell(row=row, column=col, value=f"  {title}")
    c.font = Font(bold=True, color="FFFFFF", size=11, name="Arial")
    c.fill = hdr_fill(color)
    c.alignment = xaln("left")
    ws.row_dimensions[row].height = 24

def xl_table(ws, sr, sc, headers, rows, widths, hc,
             num_cols=None, pct_cols=None, neg_cols=None, green_cols=None,
             rank=True, rayon_col=None):
    num_cols   = num_cols   or []
    pct_cols   = pct_cols   or []
    neg_cols   = neg_cols   or []
    green_cols = green_cols or []

    ah = (["#"] + headers) if rank else headers
    aw = ([4]   + widths)  if rank else widths
    for j, (h, w) in enumerate(zip(ah, aw)):
        col = sc + j
        c = ws.cell(row=sr, column=col, value=h)
        c.font      = Font(bold=True, color="FFFFFF", size=10, name="Arial")
        c.fill      = hdr_fill(hc)
        c.alignment = xaln("center")
        c.border    = mk_border("thin", "555555")
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[sr].height = 22

    off = 1 if rank else 0
    for i, rd in enumerate(rows):
        r = sr + 1 + i
        ws.row_dimensions[r].height = 19
        bg = hdr_fill("F9F9FB" if i % 2 else "FFFFFF")
        if rank:
            c = ws.cell(row=r, column=sc, value=i + 1)
            c.font      = Font(color="AAAAAA", size=9, name="Arial")
            c.fill      = bg
            c.alignment = xaln("center")
            c.border    = mk_bottom()
        for j, val in enumerate(rd):
            col = sc + j + off
            if rayon_col is not None and j == rayon_col:
                label, rcolor = val
                c = ws.cell(row=r, column=col, value=label)
                c.fill      = bg
                c.border    = mk_bottom()
                c.font      = Font(bold=True, color=rcolor.replace("#", ""), size=10, name="Arial")
                c.alignment = xaln("center")
                continue
            c = ws.cell(row=r, column=col, value=val)
            c.fill   = bg
            c.border = mk_bottom()
            if j in num_cols:
                c.number_format = '#,##0'
                c.alignment     = xaln("right")
                c.font          = xfnt()
            elif j in pct_cols:
                c.number_format = '0.0%'
                c.alignment     = xaln("center")
                c.font          = xfnt()
            else:
                c.alignment = xaln("left")
                c.font      = xfnt()
            if j in neg_cols and isinstance(val, (int, float)) and val < 0:
                c.font = Font(bold=True, color="FF3B30", size=10, name="Arial")
            if j in green_cols and isinstance(val, (int, float)) and val > 0:
                c.font = Font(bold=True, color="34C759", size=10, name="Arial")
    return sr + 1 + len(rows)

def sp(ws, row, h=12):
    ws.row_dimensions[row].height = h
    return row + 1

# ─── EXPORT EXCEL ─────────────────────────────────────────────────────────────
def generate_excel(arts, rayon_tots):
    wb  = Workbook()
    RCOL_W = 16

    def rc_rows(df, art_col, rayon_col, *extra):
        return [
            tuple([r[art_col], (r[rayon_col], COLORS.get(r[rayon_col], "555555"))]
                  + [r[c] for c in extra])
            for _, r in df.iterrows()
        ]

    tca = top_ca(arts)
    tmg = top_marge(arts)
    tpr = top_promo(arts)
    tfl = flop_marge(arts)
    tcs = top_casse(arts)
    ca_t, mg_t, pct_t, cs_t, nb_neg, nb_cs = compute_kpis(arts)
    nb_art = len(arts)

    # ── ONGLET SYNTHÈSE RÉSEAU ────────────────────────────────────────────────
    ws0 = wb.active
    ws0.title = "📊 Synthèse Réseau"
    ws0.sheet_view.showGridLines = False
    ws0.row_dimensions[1].height = 8

    ws0.merge_cells("A2:J2")
    ws0["A2"] = "PERFORMANCE COMMERCIALE HEBDOMADAIRE — RÉSEAU CARREFOUR CÔTE D'IVOIRE"
    ws0["A2"].font      = Font(bold=True, color="FFFFFF", size=15, name="Arial")
    ws0["A2"].fill      = hdr_fill("#1C1C1E")
    ws0["A2"].alignment = xaln("center")
    ws0.row_dimensions[2].height = 36

    ws0.merge_cells("A3:J3")
    ws0["A3"] = f"Extraction PBI · {nb_art:,} articles actifs · Semaine en cours"
    ws0["A3"].font      = Font(color="8E8E93", size=10, name="Arial")
    ws0["A3"].fill      = hdr_fill("#F2F2F7")
    ws0["A3"].alignment = xaln("center")
    ws0.row_dimensions[3].height = 20
    ws0.row_dimensions[4].height = 10

    # Bloc KPIs colorés
    kpis_net = [
        ("CA TOTAL RÉSEAU",  f"{ca_t/1e6:.1f} M FCFA",  "",                   "#007AFF", 1),
        ("MARGE BRUTE",      f"{mg_t/1e6:.1f} M FCFA",  f"{pct_t:.1%} du CA", "#34C759", 3),
        ("CASSE RÉSEAU",     f"{cs_t/1e6:.2f} M FCFA",  f"{nb_cs} articles",  "#FF3B30", 5),
        ("MARGES NÉGATIVES", f"{nb_neg} articles",       "à corriger",         "#FF9500", 7),
    ]
    for label, val, sub, color, col in kpis_net:
        ec = col + 1
        for r in range(5, 10):
            ws0.merge_cells(start_row=r, start_column=col, end_row=r, end_column=ec)
        ws0.cell(row=5, column=col).fill = hdr_fill(color)
        ws0.row_dimensions[5].height = 5
        for row_num, txt, sz, bold in [(6, label, 9, True), (7, val, 13, True), (8, sub, 9, False)]:
            c = ws0.cell(row=row_num, column=col, value=txt)
            c.font      = Font(bold=bold, color="FFFFFF", size=sz, name="Arial")
            c.fill      = hdr_fill(color)
            c.alignment = xaln("center")
            ws0.row_dimensions[row_num].height = 16 if row_num != 7 else 26
        ws0.cell(row=9, column=col).fill = hdr_fill(color)
        ws0.row_dimensions[9].height = 5

    ws0.row_dimensions[10].height = 14

    # Récapitulatif par rayon
    xl_hdr(ws0, 11, 1, "RÉCAPITULATIF PAR RAYON", "#3A3A3C", 7)
    rh = ["RAYON", "CA (FCFA)", "MARGE (FCFA)", "% MARGE", "CASSE (FCFA)", "ART. ACTIFS"]
    rw = [26, 18, 18, 10, 16, 12]
    for j, (h, w) in enumerate(zip(rh, rw)):
        c = ws0.cell(row=12, column=1 + j, value=h)
        c.font      = Font(bold=True, color="FFFFFF", size=10, name="Arial")
        c.fill      = hdr_fill("#3A3A3C")
        c.alignment = xaln("center")
        c.border    = mk_border("thin", "555555")
        ws0.column_dimensions[get_column_letter(1 + j)].width = w
    ws0.row_dimensions[12].height = 22

    nb_arts_r = arts.groupby("rayon_label").size().to_dict()
    for i, (_, row) in enumerate(rayon_tots.iterrows()):
        r  = 13 + i
        ws0.row_dimensions[r].height = 20
        bg = hdr_fill("F9F9FB" if i % 2 else "FFFFFF")
        vals = [row["rayon_label"], row["CA"], row["Marge"], row["%Marge"],
                row["Casse"], nb_arts_r.get(row["rayon_label"], 0)]
        fmts = [None, '#,##0', '#,##0', '0.0%', '#,##0', '#,##0']
        for j, (v, fmt) in enumerate(zip(vals, fmts)):
            c = ws0.cell(row=r, column=1 + j, value=v)
            c.fill      = bg
            c.border    = mk_bottom()
            if fmt:
                c.number_format = fmt
            c.alignment = xaln("right" if isinstance(v, (int, float)) else "left")
            c.font      = xfnt(bold=(j == 0),
                               color=COLORS.get(str(v), "#1C1C1E") if j == 0 else "#1C1C1E")
            if j == 4 and isinstance(v, (int, float)) and v < 0:
                c.font = Font(bold=True, color="FF3B30", size=10, name="Arial")

    r_tot = 13 + len(rayon_tots)
    ws0.row_dimensions[r_tot].height = 22
    totals = [
        ("TOTAL",  None,    "#1C1C1E"),
        (ca_t,     '#,##0', "#007AFF"),
        (mg_t,     '#,##0', "#34C759"),
        (pct_t,    '0.0%',  "#1C1C1E"),
        (cs_t,     '#,##0', "#FF3B30"),
        (nb_art,   '#,##0', "#1C1C1E"),
    ]
    for j, (v, fmt, clr) in enumerate(totals):
        c = ws0.cell(row=r_tot, column=1 + j, value=v)
        c.fill      = hdr_fill("#F2F2F7")
        c.border    = mk_border("medium", "AAAAAA")
        if fmt:
            c.number_format = fmt
        c.font      = Font(bold=True, color=clr.replace("#", ""), size=10, name="Arial")
        c.alignment = xaln("right" if j > 0 else "left")

    cur = r_tot + 1
    cur = sp(ws0, cur, 18)

    # Classements réseau
    xl_hdr(ws0, cur, 1, "TOP 10 CA — TOUS RAYONS CONFONDUS", "#1C1C1E", 8)
    cur += 1
    cur = xl_table(ws0, cur, 1,
        ["ARTICLE", "RAYON", "CA (FCFA)", "MARGE (FCFA)", "% MARGE", "QTÉ VENDUE"],
        rc_rows(tca, "art_label", "rayon_label", "CA", "Marge", "%Marge", "Qté Vente"),
        [40, RCOL_W, 16, 16, 10, 12], "#1C1C1E",
        num_cols=[2, 3, 5], pct_cols=[4], green_cols=[3], neg_cols=[3], rayon_col=1)
    cur = sp(ws0, cur, 14)

    xl_hdr(ws0, cur, 1, "TOP 10 MARGE BRUTE — TOUS RAYONS CONFONDUS", "#34C759", 7)
    cur += 1
    cur = xl_table(ws0, cur, 1,
        ["ARTICLE", "RAYON", "CA (FCFA)", "MARGE (FCFA)", "% MARGE"],
        rc_rows(tmg, "art_label", "rayon_label", "CA", "Marge", "%Marge"),
        [40, RCOL_W, 16, 16, 10], "#34C759",
        num_cols=[2, 3], pct_cols=[4], green_cols=[3], rayon_col=1)
    cur = sp(ws0, cur, 14)

    xl_hdr(ws0, cur, 1, "TOP 10 VENTES PROMO — TOUS RAYONS CONFONDUS", "#FF9500", 7)
    cur += 1
    if not tpr.empty:
        cur = xl_table(ws0, cur, 1,
            ["ARTICLE", "RAYON", "CA PROMO (FCFA)", "MARGE PROMO (FCFA)", "POIDS PROMO"],
            rc_rows(tpr, "art_label", "rayon_label", "CA HT Promo", "Marge Promo", "%CA Poids Promo"),
            [40, RCOL_W, 18, 18, 12], "#FF9500",
            num_cols=[2, 3], pct_cols=[4], green_cols=[3], rayon_col=1)
    cur = sp(ws0, cur, 14)

    xl_hdr(ws0, cur, 1, "FLOP 15 MARGES NÉGATIVES — TOUS RAYONS CONFONDUS", "#FF3B30", 7)
    cur += 1
    cur = xl_table(ws0, cur, 1,
        ["ARTICLE", "RAYON", "CA (FCFA)", "MARGE (FCFA)", "% MARGE"],
        rc_rows(tfl, "art_label", "rayon_label", "CA", "Marge", "%Marge"),
        [40, RCOL_W, 14, 14, 12], "#FF3B30",
        num_cols=[2, 3], pct_cols=[4], neg_cols=[3, 4], rayon_col=1)
    cur = sp(ws0, cur, 14)

    xl_hdr(ws0, cur, 1, "TOP 10 CASSE — TOUS RAYONS CONFONDUS", "#555555", 6)
    cur += 1
    if not tcs.empty:
        xl_table(ws0, cur, 1,
            ["ARTICLE", "RAYON", "CASSE VALEUR (FCFA)", "CASSE QTÉ"],
            rc_rows(tcs, "art_label", "rayon_label", "Casse (Valeur)", "Casse (Qté)"),
            [40, RCOL_W, 20, 12], "#555555",
            num_cols=[2, 3], neg_cols=[2, 3], rayon_col=1)

    ws0.freeze_panes = "A4"

    # ── ONGLETS PAR RAYON ─────────────────────────────────────────────────────
    for rayon, color in COLORS.items():
        arts_r = arts[arts["rayon_label"] == rayon]
        if arts_r.empty:
            continue

        ca_r  = arts_r["CA"].sum()
        mg_r  = arts_r["Marge"].sum()
        pct_r = mg_r / ca_r if ca_r else 0
        cs_r  = arts_r["Casse (Valeur)"].sum() if "Casse (Valeur)" in arts_r.columns else 0
        neg_r = int((arts_r["Marge"] < 0).sum())
        cas_r = int((arts_r["Casse (Valeur)"] < 0).sum()) if "Casse (Valeur)" in arts_r.columns else 0

        tca_r = top_ca(arts_r)
        tmg_r = top_marge(arts_r)
        tpr_r = top_promo(arts_r)
        tfl_r = flop_marge(arts_r)
        tcs_r = top_casse(arts_r)

        ws = wb.create_sheet(f"📋 {rayon}")
        ws.sheet_view.showGridLines = False
        ws.row_dimensions[1].height = 8

        ws.merge_cells("A2:G2")
        ws["A2"] = f"CLASSEMENTS SEMAINE — {rayon.upper()}"
        ws["A2"].font      = Font(bold=True, color="FFFFFF", size=14, name="Arial")
        ws["A2"].fill      = hdr_fill(color)
        ws["A2"].alignment = xaln("center")
        ws.row_dimensions[2].height = 32

        ws.merge_cells("A3:G3")
        ws["A3"] = (f"CA : {ca_r:,.0f} FCFA  |  Marge : {mg_r:,.0f} FCFA ({pct_r:.1%})"
                    f"  |  Casse : {cs_r:,.0f} FCFA  |  Marges nég. : {neg_r} art.")
        ws["A3"].font      = Font(color=color.replace("#", ""), size=10, name="Arial", bold=True)
        ws["A3"].fill      = hdr_fill("#F9F9FB")
        ws["A3"].alignment = xaln("center")
        ws.row_dimensions[3].height = 22

        def sr(df, cols):
            return [tuple(r[c] for c in cols) for _, r in df.iterrows()]

        cur2 = 5
        xl_hdr(ws, cur2, 1, f"TOP 10 CA — {rayon.upper()}", color, 7)
        cur2 += 1
        cur2 = xl_table(ws, cur2, 1,
            ["ARTICLE", "CA (FCFA)", "MARGE (FCFA)", "% MARGE", "QTÉ VENDUE"],
            sr(tca_r, ["art_label", "CA", "Marge", "%Marge", "Qté Vente"]),
            [42, 16, 16, 10, 12], color,
            num_cols=[1, 2, 4], pct_cols=[3], green_cols=[2], neg_cols=[2, 3])
        cur2 = sp(ws, cur2, 14)

        xl_hdr(ws, cur2, 1, f"TOP 10 MARGE — {rayon.upper()}", "#34C759", 6)
        cur2 += 1
        cur2 = xl_table(ws, cur2, 1,
            ["ARTICLE", "CA (FCFA)", "MARGE (FCFA)", "% MARGE"],
            sr(tmg_r, ["art_label", "CA", "Marge", "%Marge"]),
            [44, 16, 16, 10], "#34C759",
            num_cols=[1, 2], pct_cols=[3], green_cols=[2])
        cur2 = sp(ws, cur2, 14)

        xl_hdr(ws, cur2, 1, f"TOP PROMO — {rayon.upper()}", "#FF9500", 6)
        cur2 += 1
        if not tpr_r.empty:
            cur2 = xl_table(ws, cur2, 1,
                ["ARTICLE", "CA PROMO (FCFA)", "MARGE PROMO (FCFA)", "POIDS PROMO"],
                sr(tpr_r, ["art_label", "CA HT Promo", "Marge Promo", "%CA Poids Promo"]),
                [44, 18, 18, 12], "#FF9500",
                num_cols=[1, 2], pct_cols=[3], green_cols=[2])
        cur2 = sp(ws, cur2, 14)

        xl_hdr(ws, cur2, 1, f"FLOP MARGES NÉGATIVES — {rayon.upper()}", "#FF3B30", 6)
        cur2 += 1
        ws.merge_cells(start_row=cur2, start_column=1, end_row=cur2, end_column=6)
        note = ws.cell(row=cur2, column=1,
                       value=f"⚠  {neg_r} articles en perte — vérifier PA, promos non compensées, conditions fournisseurs.")
        note.font      = Font(italic=True, color="FF3B30", size=9, name="Arial")
        note.fill      = hdr_fill("#FFF5F5")
        note.alignment = xaln("left")
        ws.row_dimensions[cur2].height = 17
        cur2 += 1
        cur2 = xl_table(ws, cur2, 1,
            ["ARTICLE", "CA (FCFA)", "MARGE (FCFA)", "% MARGE"],
            sr(tfl_r, ["art_label", "CA", "Marge", "%Marge"]),
            [44, 14, 14, 12], "#FF3B30",
            num_cols=[1, 2], pct_cols=[3], neg_cols=[2, 3])
        cur2 = sp(ws, cur2, 14)

        xl_hdr(ws, cur2, 1, f"TOP CASSE — {rayon.upper()}", "#555555", 5)
        cur2 += 1
        ws.merge_cells(start_row=cur2, start_column=1, end_row=cur2, end_column=5)
        note2 = ws.cell(row=cur2, column=1,
                        value=f"Casse totale : {cs_r:,.0f} FCFA  ·  {cas_r} articles impactés")
        note2.font      = Font(italic=True, color="555555", size=9, name="Arial")
        note2.fill      = hdr_fill("#F2F2F7")
        note2.alignment = xaln("left")
        ws.row_dimensions[cur2].height = 17
        cur2 += 1
        if not tcs_r.empty:
            xl_table(ws, cur2, 1,
                ["ARTICLE", "CASSE VALEUR (FCFA)", "CASSE QTÉ"],
                sr(tcs_r, ["art_label", "Casse (Valeur)", "Casse (Qté)"]),
                [46, 20, 12], "#555555",
                num_cols=[1, 2], neg_cols=[1, 2])

        ws.freeze_panes = "A4"

    wb.active = wb["📊 Synthèse Réseau"]
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:18px'>
  <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>🛍️ SmartBuyer</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · Équipe Achats</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Navigation</div>", unsafe_allow_html=True)
    nav_pages = [
        ("app.py",                             "🏠  Accueil"),
        ("pages/01_📊_Analyse_Scoring_ABC.py", "📊  Scoring ABC"),
        ("pages/02_📈_Ventes_PBI.py",          "📈  Ventes PBI"),
        ("pages/09_📦_Tasks_Trackers.py",      "📋  Task Tracker"),
        ("pages/10_📊_Perf_Hebdo.py",          "📊  Perf Hebdo"),
    ]
    for path, label in nav_pages:
        try:
            st.page_link(path, label=label)
        except Exception:
            pass
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Import fichier</div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Extraction PBI (semaine en cours)", type=["xlsx", "xls", "csv"], key="pbi")

# ─── PAGE PRINCIPALE ──────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>📊 Performance commerciale hebdo</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Classements hebdomadaires · Épicerie · Boissons · DPH</div>", unsafe_allow_html=True)

# ─── ÉCRAN D'ACCUEIL (pas de fichier) ────────────────────────────────────────
if not uploaded:
    st.markdown("---")
    st.markdown("""
<div class="alert-card alert-blue">
  <strong>ℹ️ À quoi sert ce module ?</strong><br>
  Le module <strong>Performance Hebdo</strong> transforme votre extraction PBI en rapport de synthèse
  actionnable pour toute l'équipe Achats. Il calcule automatiquement, pour l'ensemble du réseau
  et par rayon, les cinq classements clés de la semaine :
  <strong>Top CA</strong>, <strong>Top masse de marge</strong>,
  <strong>Meilleures ventes en promotion</strong>,
  <strong>Flop marges négatives</strong> et <strong>Top casse</strong>.<br><br>
  Le bouton <strong>Exporter Excel</strong> produit un fichier structuré en 4 onglets —
  Synthèse réseau (avec tous les classements tous rayons) + un onglet par rayon
  (Épicerie, Boissons, DPH) — directement partageable avec les acheteurs.
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Colonnes attendues dans le fichier PBI</div>", unsafe_allow_html=True)
    icons = {
        "Rayon": "🏷️", "Famille": "🏷️", "Article": "🏷️", "Site nom long": "🏷️",
        "CA": "🔢", "Marge": "🔢", "%Marge": "🔢", "CA Promo": "🔢",
        "Marge Promo": "🔢", "Qté Vente": "🔢", "Casse (Valeur)": "🔢",
    }
    c1, c2 = st.columns(2)
    for i, (col_name, (desc, example)) in enumerate(REQUIRED_COLS.items()):
        with (c1 if i < 6 else c2):
            st.markdown(f"""
<div class="col-required">
  <div style="font-size:16px;margin-top:1px">{icons.get(col_name,"📌")}</div>
  <div>
    <div class="col-name">{col_name}</div>
    <div class="col-desc">{desc}</div>
    <div class="col-example">{example}</div>
  </div>
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
<div class="alert-card alert-blue">
  <strong>📌 Format du fichier</strong><br>
  Export PBI standard avec hiérarchie <strong>Rayon → Famille → Sous-Famille → Article × Site</strong>.<br>
  Le module conserve automatiquement uniquement les lignes <em>Total réseau</em> par article
  (colonne <code>Site nom long = "Total"</code>) pour éviter tout doublon dans les classements.
  Les 4 rayons reconnus : <strong>Épicerie, Boissons, DPH (Droguerie + Parfumerie Hygiène)</strong>.
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.info("⬆️ Charge ton extraction PBI dans la sidebar pour lancer l'analyse.")
    st.stop()

# ─── CHARGEMENT ───────────────────────────────────────────────────────────────
with st.spinner("Lecture du fichier…"):
    arts, rayon_tots = parse_file(uploaded.getvalue(), uploaded.name)

if arts.empty:
    st.error("Impossible de lire les données. Vérifier le format du fichier.")
    st.stop()

ca_tot, marge_tot, pct_tot, casse_tot, nb_neg, nb_cs = compute_kpis(arts)

# ─── FILTRES SIDEBAR ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Filtres</div>", unsafe_allow_html=True)
    sel_rayon = st.multiselect(
        "Rayon",
        sorted(arts["rayon_label"].unique()),
        default=sorted(arts["rayon_label"].unique()),
    )
    st.caption(f"Fichier : **{uploaded.name}**")
    st.caption(f"{len(arts):,} articles actifs")

# ─── KPIs RÉSEAU ──────────────────────────────────────────────────────────────
st.markdown(f"<div class='section-label'>Vue réseau · {uploaded.name}</div>", unsafe_allow_html=True)
k1, k2, k3, k4 = st.columns(4)
k1.metric("CA Total Réseau",  fmt_fcfa(ca_tot))
k2.metric("Marge Brute",      fmt_fcfa(marge_tot), f"{pct_tot:.1%} du CA")
k3.metric("Casse",            fmt_fcfa(casse_tot), f"{nb_cs} articles")
k4.metric("Marges négatives", f"{nb_neg} articles")

# ─── CARDS PAR RAYON ──────────────────────────────────────────────────────────
st.markdown("---")
r1, r2, r3 = st.columns(3)
ca_max = rayon_tots["CA"].max() if not rayon_tots.empty else 1
for col_ui, rayon in zip([r1, r2, r3], ["Épicerie", "Boissons", "DPH"]):
    row = rayon_tots[rayon_tots["rayon_label"] == rayon]
    if row.empty:
        continue
    rv     = row.iloc[0]
    color  = COLORS[rayon]
    pct_bar = int(rv["CA"] / ca_max * 100) if ca_max else 0
    col_ui.markdown(f"""<div class="kpi-rayon">
      <div class="kpi-rayon-name" style="color:{color}">{rayon}</div>
      <div class="kpi-rayon-ca"   style="color:{color}">{fmt_fcfa(rv['CA'])} FCFA</div>
      <div class="kpi-rayon-sub">Marge {rv['%Marge']:.1%} &nbsp;·&nbsp; Casse {fmt_fcfa(rv['Casse'])} FCFA</div>
      <div style="height:4px;background:#E5E5EA;border-radius:2px;margin-top:10px;overflow:hidden">
        <div style="height:4px;width:{pct_bar}%;background:{color};border-radius:2px"></div>
      </div>
    </div>""", unsafe_allow_html=True)

# ─── ALERTES ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Alertes &amp; actions prioritaires</div>", unsafe_allow_html=True)

alerts = []
critiques = int((arts["Marge"] < -100_000).sum())
if critiques > 0:
    alerts.append(("🔴",
                   f"{critiques} article{'s' if critiques > 1 else ''} avec marge < -100 000 FCFA",
                   "Vérifier PA et conditions fournisseurs — action corrective avant S+1."))
for rayon in ["Épicerie", "Boissons", "DPH"]:
    row = rayon_tots[rayon_tots["rayon_label"] == rayon]
    if row.empty:
        continue
    neg_r = int((arts[arts["rayon_label"] == rayon]["Marge"] < 0).sum())
    if neg_r > 20:
        alerts.append(("🟡",
                       f"{rayon} : {neg_r} articles en marge négative",
                       "Analyser les causes rayon par rayon — conditions ou PA à corriger."))
if not top_casse(arts).empty and abs(casse_tot) > 1_000_000:
    alerts.append(("🟡",
                   f"Casse totale : {fmt_fcfa(casse_tot)} FCFA",
                   "Revoir les DLC et les politiques de commande sur les articles les plus exposés."))

if not alerts:
    st.success("✅ Aucune alerte critique cette semaine.")
else:
    st.markdown("".join(alert_html(*a) for a in alerts), unsafe_allow_html=True)

# ─── CLASSEMENTS ──────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Classements</div>", unsafe_allow_html=True)

rayon_options = ["Tous rayons"] + [
    r for r in ["Épicerie", "Boissons", "DPH"] if r in arts["rayon_label"].unique()
]
rayon_filtre = st.segmented_control(
    "Rayon", rayon_options, default="Tous rayons", label_visibility="collapsed"
)
arts_f = arts if rayon_filtre == "Tous rayons" else arts[arts["rayon_label"] == rayon_filtre]

RENAME = {
    "art_label":        "Article",
    "rayon_label":      "Rayon",
    "CA":               "CA (FCFA)",
    "Marge":            "Marge (FCFA)",
    "%Marge":           "% Marge",
    "Qté Vente":        "Qté vendue",
    "CA HT Promo":      "CA Promo (FCFA)",
    "Marge Promo":      "Marge Promo (FCFA)",
    "%CA Poids Promo":  "Poids Promo",
    "Casse (Valeur)":   "Casse valeur (FCFA)",
    "Casse (Qté)":      "Casse qté",
}

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🏆 Top CA", "💚 Top Marge", "🎯 Top Promo", "🔴 Flop Marges", "🗑️ Casse"
])
with tab1:
    show_df(top_ca(arts_f), RENAME,
            num_cols=["CA", "Marge", "Qté Vente"], pct_cols=["%Marge"],
            neg_cols=["Marge"], green_cols=["Marge"])
with tab2:
    show_df(top_marge(arts_f), RENAME,
            num_cols=["CA", "Marge"], pct_cols=["%Marge"], green_cols=["Marge"])
with tab3:
    df_pr = top_promo(arts_f)
    if df_pr.empty:
        st.info("Aucun article promotionnel sur ce périmètre.")
    else:
        show_df(df_pr, RENAME,
                num_cols=["CA HT Promo", "Marge Promo"], pct_cols=["%CA Poids Promo"],
                green_cols=["Marge Promo"])
with tab4:
    df_fl = flop_marge(arts_f)
    st.warning(f"⚠️ {len(df_fl)} articles en marge négative sur ce périmètre.")
    show_df(df_fl, RENAME,
            num_cols=["CA", "Marge"], pct_cols=["%Marge"], neg_cols=["Marge", "%Marge"])
with tab5:
    df_cs = top_casse(arts_f)
    if df_cs.empty:
        st.info("Aucune casse enregistrée sur ce périmètre.")
    else:
        show_df(df_cs, RENAME,
                num_cols=["Casse (Valeur)", "Casse (Qté)"],
                neg_cols=["Casse (Valeur)", "Casse (Qté)"])

# ─── EXPORT ───────────────────────────────────────────────────────────────────
st.markdown("---")
with st.expander("📥 Export Excel — Synthèse réseau · Classements · 1 onglet par rayon"):
    st.caption("1 fichier · 4 onglets · Synthèse réseau + classements tous rayons + Épicerie, Boissons, DPH")
    if st.button("Générer le fichier Excel", type="primary"):
        with st.spinner("Génération en cours…"):
            buf = generate_excel(arts, rayon_tots)
        fname = (f"Perf_Hebdo_SmartBuyer_"
                 f"{uploaded.name.replace('.xlsx','').replace('.xls','').replace('.csv','')}.xlsx")
        st.download_button(
            "⬇️ Télécharger", data=buf, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
