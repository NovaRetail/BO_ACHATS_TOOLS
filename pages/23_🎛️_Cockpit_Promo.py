"""
23_🎛️_Cockpit_Promo.py — Module Cockpit Promo · SmartBuyer Hub
Balaie un export PROMO et sort les alertes Disponibilité (stock ≤0, avec/sans
réappro RAL) et Marge (vente à perte, marge faible, PMP manquant), avec un
flag combiné "Disponibilité + Marge" pour les cas cumulés prioritaires.
Colonne "Action" : message actionnable par ligne (Rupture -> Passer commande,
RAL -> Accélérer livraison, etc.) — pas juste un statut descriptif.
Export Excel "Cockpit Promo - AAAAMMJJ.xlsx" — valeurs calculées en Python,
aucune formule dans le classeur.
"""

import datetime as _dt

import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# CONFIG & CHARTE (Apple clair — identique aux autres modules SmartBuyer)
# ============================================================
st.set_page_config(page_title="Cockpit Promo · SmartBuyer", page_icon="🎛️", layout="wide", initial_sidebar_state="expanded")

BLUE = "#007AFF"
GREEN = "#34C759"
RED = "#FF3B30"
AMBER = "#FF9500"
DARK = "#1D1D1F"
GREY = "#86868B"
BG = "#F2F2F7"

MARGE_FAIBLE_DEFAUT = 10  # % — modifiable dans la sidebar
VALID_SITES = {"0010301", "0010202", "0010203"}  # à ajuster si le périmètre magasins change

REQUIRED_COLS = ["Code site", "Rayon", "Libellé rayon", "Libellé article", "Code article",
                  "DPR", "PV Promo", "Taux TVA", "PMP", "Stock", "RAL", "Marge en cours", "Four."]
OPTIONAL_COLS = ["Quantité vendue", "Montant vente HT", "Montant achat"]

st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display",
                 "SF Pro Text", "Helvetica Neue", Arial, sans-serif !important;
    background-color: #F2F2F7;
}
.stApp { background: #F2F2F7; }
.main .block-container { padding-top: 1.8rem; max-width: 1350px; }
[data-testid="stSidebar"] { background: #F2F2F7 !important; border-right: 0.5px solid #D1D1D6 !important; }
[data-testid="stMetric"] { background: #FFFFFF !important; border: 0.5px solid #E5E5EA !important; border-radius: 12px !important; padding: 16px 18px !important; }
[data-testid="stMetricLabel"] { font-size: 11px !important; font-weight: 500 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 600 !important; color: #1C1C1E !important; letter-spacing: -0.02em !important; }
[data-testid="stDataFrame"] { border: 0.5px solid #E5E5EA !important; border-radius: 12px !important; overflow: hidden !important; }
[data-testid="stDataFrame"] th { background: #FAFAFC !important; font-size: 11px !important; font-weight: 600 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; border-bottom: 0.5px solid #E5E5EA !important; }
[data-testid="stDataFrame"] td { font-size: 13px !important; border-bottom: 0.5px solid #F2F2F7 !important; }
[data-testid="stFileUploader"] { border: 1.5px dashed #D1D1D6 !important; border-radius: 10px !important; background: #F9F9FB !important; }
.stDownloadButton > button { background: #007AFF !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; font-size: 13px !important; padding: 10px 24px !important; width: 100% !important; }
hr { border-color: #E5E5EA !important; margin: 1rem 0 !important; }

.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 1.5rem; }
.section-label { font-size: 11px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 10px; }
.alert-card  { padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; font-size: 13px; line-height: 1.5; border-left: 3px solid; background: #FFFFFF; }
.alert-red   { background: #FFF2F2; border-color: #FF3B30; color: #3A0000; }
.alert-amber { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }
.alert-green { background: #F0FFF4; border-color: #34C759; color: #003A10; }
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }

.format-card { border-radius: 12px; padding: 14px 16px; margin-bottom: 6px; border: 0.5px solid; }
.format-hyper  { background: #EFF6FF; border-color: #B3D9FF; }
.format-market { background: #FFFBF0; border-color: #FFD9A0; }
.format-supeco { background: #FFF2F2; border-color: #FFB3B3; }

.badge { display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: 600; }
.badge-hyper  { background: #154360; color: #FFFFFF; }
.badge-red    { background: #FF3B30; color: #FFFFFF; }
.badge-amber  { background: #FF9500; color: #FFFFFF; }
.badge-green  { background: #34C759; color: #FFFFFF; }

.col-required { background: #F0F8FF; border: 0.5px solid #B3D9FF; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; display: flex; align-items: flex-start; gap: 10px; }
.col-optional { background: #F9F9FB; border-color: #D1D1D6; }
.col-name { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
.card { background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:16px;margin-bottom:10px; }
.small-muted { font-size:12px;color:#8E8E93; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# HELPERS DE FORMAT (convention SmartBuyer)
# ============================================================
def fmt(n):
    if n is None or (isinstance(n, float) and (pd.isna(n) or not np.isfinite(n))):
        return "—"
    a = abs(n)
    if a >= 1_000_000: return f"{n/1_000_000:.1f} M"
    if a >= 1_000:     return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def fmt_pct(v, dec=1):
    if v is None or pd.isna(v): return "—"
    return f"{v:.{dec}f}%"

def read_csv_robust(file) -> pd.DataFrame:
    """Lecture CSV avec repli d'encodage UTF-8 -> CP1252 -> Latin-1 (convention SmartBuyer)."""
    raw = file.read()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(raw), sep=";", encoding=enc, dtype=str)
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError("Impossible de décoder le fichier (UTF-8 / CP1252 / Latin-1 ont échoué).")

def to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.strip(), errors="coerce")

# ============================================================
# CALCUL DES ALERTES
# ============================================================
@st.cache_data(show_spinner=False)
def compute_alerts(file_bytes, seuil_marge_faible):
    df = read_csv_robust(io.BytesIO(file_bytes))
    df.columns = [c.strip() for c in df.columns]
    df = df[df["Code site"].isin(VALID_SITES)].copy()

    for c in ["DPR", "PV Promo", "Taux TVA", "PMP", "Stock", "RAL", "Marge en cours",
              "Quantité vendue", "Montant vente HT", "Montant achat", "Four."]:
        df[c] = to_num(df[c]) if c in df.columns else np.nan

    df["Site"] = df["Code site"].str.lstrip("0").astype(int)
    df["Rayon_aff"] = df["Libellé rayon"].str.strip()
    df["Article_aff"] = df["Libellé article"].str.strip()
    df["Code_article"] = df["Code article"].astype(str).str.lstrip("0")
    df["RAL_int"] = df["RAL"].fillna(0).astype(int)

    df["stat_stock"] = np.where(
        (df["Stock"] <= 0) & (df["RAL_int"] <= 0), "Rupture sans réappro",
        np.where((df["Stock"] <= 0) & (df["RAL_int"] > 0), "Réappro en cours", ""))
    df["stat_marge"] = np.where(
        df["Marge en cours"] < 0, "Vente à perte",
        np.where(df["Marge en cours"] < seuil_marge_faible, "Marge faible", ""))
    df["pmp_manquant"] = df["PMP"].isna()
    df["flag_dispo"] = df["stat_stock"] != ""
    df["flag_marge"] = (df["stat_marge"] != "") | df["pmp_manquant"]
    df["type_alerte"] = np.where(
        df["flag_dispo"] & df["flag_marge"], "Disponibilité + Marge",
        np.where(df["flag_dispo"], "Disponibilité",
                 np.where(df["flag_marge"], "Marge", "")))
    df["marge_realisee"] = df["Montant vente HT"] - df["Montant achat"]
    df["pmp_effectif"] = df["PMP"].fillna(df["DPR"])

    # ---- Colonne Action : message actionnable, pas un statut descriptif ----
    df["action_stock"] = np.where(
        df["stat_stock"] == "Rupture sans réappro", "Rupture - Passer commande",
        np.where(df["stat_stock"] == "Réappro en cours",
                 "RAL - Accélérer livraison (" + df["RAL_int"].astype(str) + ")", ""))
    df["action_marge"] = np.where(
        df["stat_marge"] != "", df["stat_marge"],
        np.where(df["pmp_manquant"], "PMP manquant", ""))
    df["action"] = np.where(
        df["type_alerte"] == "Disponibilité + Marge",
        df["action_stock"] + " + " + df["action_marge"],
        np.where(df["type_alerte"] == "Disponibilité", df["action_stock"], df["action_marge"]))

    return df[df["type_alerte"] != ""].copy()

def synthese_par_site(alerts):
    rows = []
    for site in sorted(alerts["Site"].unique()):
        sub = alerts[alerts["Site"] == site]
        rows.append({
            "Site": site,
            "Rupture sans réappro": int((sub["stat_stock"] == "Rupture sans réappro").sum()),
            "Réappro en cours": int((sub["stat_stock"] == "Réappro en cours").sum()),
            "Vente à perte": int((sub["stat_marge"] == "Vente à perte").sum()),
            "Marge faible": int((sub["stat_marge"] == "Marge faible").sum()),
            "PMP manquant": int((sub["action_marge"] == "PMP manquant").sum()),
            "Dispo + Marge": int((sub["type_alerte"] == "Disponibilité + Marge").sum()),
        })
    return pd.DataFrame(rows)

def build_headline(alerts, synth):
    n = len(alerts)
    n_cumul = int((alerts["type_alerte"] == "Disponibilité + Marge").sum())
    n_perte = int((alerts["stat_marge"] == "Vente à perte").sum())
    pire_site = synth.loc[synth["Dispo + Marge"].idxmax()] if len(synth) else None
    line1 = (f"{n} lignes en alerte &nbsp;·&nbsp; "
             f"{n_cumul} en cumul Disponibilité + Marge (prioritaires) &nbsp;·&nbsp; "
             f"{n_perte} ventes à perte")
    line2 = ""
    if pire_site is not None and pire_site["Dispo + Marge"] > 0:
        line2 = f"📌 Site à prioriser : <b>{int(pire_site['Site'])}</b> ({int(pire_site['Dispo + Marge'])} alertes cumulées)"
    return line1, line2

# ============================================================
# STYLE DES TABLEAUX STREAMLIT — coloration cellule par cellule
# ============================================================
DISPLAY_COLS = {"Site": "Site", "Rayon_aff": "Rayon", "Article_aff": "Article",
                 "Stock": "Stock", "RAL": "RAL", "Marge en cours": "Marge %", "action": "Action"}

def _color_marge(v, seuil):
    if pd.isna(v): return ""
    if v < 0: return "background-color:#FFF2F2; color:#C0392B; font-weight:600;"
    if v < seuil: return "background-color:#FFFBF0; color:#B36B00; font-weight:600;"
    return "background-color:#F0FFF4; color:#1E7B3C;"

def _color_stock(v):
    if pd.isna(v): return ""
    return "color:#C0392B; font-weight:600;" if v <= 0 else "color:#1C1C1E;"

def _color_action(v):
    if not isinstance(v, str):
        return ""
    if "Rupture" in v:
        return "background-color:#FFF2F2; color:#C0392B; font-weight:600;"
    if "RAL" in v:
        return "background-color:#FFFBF0; color:#B36B00; font-weight:600;"
    if "Vente à perte" in v:
        return "background-color:#FFF2F2; color:#C0392B; font-weight:600;"
    if "Marge faible" in v:
        return "background-color:#FFFBF0; color:#B36B00; font-weight:600;"
    return "background-color:#F5F5F7; color:#5A5A5E;"

def render_table(d, seuil, height=None):
    if d.empty:
        st.caption("Aucune ligne pour ce filtre.")
        return
    disp = d[list(DISPLAY_COLS.keys())].rename(columns=DISPLAY_COLS).reset_index(drop=True)
    sty = (disp.style
           .format({"Marge %": lambda v: fmt_pct(v)})
           .map(lambda v: _color_marge(v, seuil), subset=["Marge %"])
           .map(_color_stock, subset=["Stock"])
           .map(_color_action, subset=["Action"]))
    st.dataframe(sty, use_container_width=True, hide_index=True, height=height)

# ============================================================
# EXPORT EXCEL — valeurs calculées en Python, aucune formule
# Couleur réservée à ce qui a un vrai impact financier (Cumul, Marge) ;
# Disponibilité reste neutre (c'est un sujet opérationnel, pas une perte).
# ============================================================
def build_excel(alerts, synth):
    BLUE_H = "FF007AFF"; NEUTRAL_H = "FF48484A"; RED_H = "FFFF3B30"; AMBER_H = "FFFF9500"
    WHITE_H = "FFFFFFFF"; LGREY_H = "FFF7F7F9"; ARIAL = "Arial"
    thin = Side(style="thin", color="FFE0E0E2")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    QTY = "#,##0"; PCT = "0.0%"; ACC = "#,##0"

    def section_bar(ws, row, ncols, text, color=NEUTRAL_H):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        c = ws.cell(row=row, column=1, value=text)
        c.font = Font(name=ARIAL, bold=True, size=11, color=WHITE_H)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 20

    def header_row(ws, row, labels):
        for i, lbl in enumerate(labels, start=1):
            c = ws.cell(row=row, column=i, value=lbl)
            c.font = Font(name=ARIAL, bold=True, size=10, color=WHITE_H)
            c.fill = PatternFill("solid", fgColor=BLUE_H)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def data_row(ws, row, values, zebra=False, left_cols=()):
        fillc = LGREY_H if zebra else "FFFFFFFF"
        for i, v in enumerate(values, start=1):
            c = ws.cell(row=row, column=i, value=v)
            c.font = Font(name=ARIAL, size=10)
            c.border = box
            c.fill = PatternFill("solid", fgColor=fillc)
            c.alignment = Alignment(horizontal="left" if i in left_cols else "center")

    def autosize(ws, widths):
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    wb = Workbook()
    ws = wb.active
    ws.title = "Cockpit Promo"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:J1")
    ws["A1"] = "COCKPIT PROMO — ALERTES DISPONIBILITÉ & MARGE"
    ws["A1"].font = Font(name=ARIAL, bold=True, size=14, color=WHITE_H)
    ws["A1"].fill = PatternFill("solid", fgColor=BLUE_H)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26
    ws["A2"] = "Généré le :"; ws["A2"].font = Font(name=ARIAL, bold=True, size=10)
    ws["B2"] = _dt.date.today().strftime("%d/%m/%Y"); ws["B2"].font = Font(name=ARIAL, size=10, color="FF0000FF", bold=True)

    r = 4
    section_bar(ws, r, 10, "1.  SYNTHÈSE RÉSEAU"); r += 1
    header_row(ws, r, ["Indicateur", "Valeur"]); r += 1
    n_cumul = int((alerts["type_alerte"] == "Disponibilité + Marge").sum())
    kpi_rows = [
        ("Total lignes en alerte", len(alerts)),
        ("Disponibilité", int((alerts["type_alerte"] == "Disponibilité").sum())),
        ("Marge", int((alerts["type_alerte"] == "Marge").sum())),
        ("Disponibilité + Marge (cumul)", n_cumul),
    ]
    for i, (label, v) in enumerate(kpi_rows):
        data_row(ws, r, [label, v], zebra=(i % 2 == 1), left_cols=(1,))
        ws.cell(row=r, column=2).number_format = QTY
        r += 1
    r += 1

    section_bar(ws, r, 10, "2.  SYNTHÈSE PAR MAGASIN"); r += 1
    header_row(ws, r, list(synth.columns)); r += 1
    for i, (_, row_) in enumerate(synth.iterrows()):
        data_row(ws, r, list(row_.values), zebra=(i % 2 == 1), left_cols=(1,))
        for col in range(2, len(synth.columns) + 1):
            ws.cell(row=r, column=col).number_format = QTY
        r += 1
    r += 1

    def alert_section(title, sub_df, color):
        nonlocal r
        section_bar(ws, r, 10, f"{title} — {len(sub_df)} lignes", color=color); r += 1
        cols = ["Site", "Rayon_aff", "Article_aff", "Code_article", "Stock", "RAL",
                "pmp_effectif", "PV Promo", "Marge en cours", "action", "Four."]
        labels = ["Site", "Rayon", "Article", "Code article", "Stock", "RAL",
                   "PMP", "PV Promo", "Marge %", "Action", "Fournisseur"]
        header_row(ws, r, labels); r += 1
        sub_sorted = sub_df.sort_values("Marge en cours")
        for i, (_, row_) in enumerate(sub_sorted.iterrows()):
            vals = [row_["Site"], row_["Rayon_aff"], row_["Article_aff"], row_["Code_article"],
                    row_["Stock"], (row_["RAL"] if pd.notna(row_["RAL"]) else 0),
                    row_["pmp_effectif"], row_["PV Promo"],
                    (row_["Marge en cours"] / 100 if pd.notna(row_["Marge en cours"]) else None),
                    row_["action"],
                    (int(row_["Four."]) if pd.notna(row_["Four."]) else None)]
            data_row(ws, r, vals, zebra=(i % 2 == 1), left_cols=(2, 3, 10))
            ws.cell(row=r, column=9).number_format = PCT
            for c in (5, 6, 7, 8):
                ws.cell(row=r, column=c).number_format = ACC
            r += 1
        r += 1

    # Couleur réservée aux enjeux financiers : Cumul = rouge, Marge = ambre.
    # Disponibilité = neutre (sujet opérationnel, pas une perte en soi).
    alert_section("3.  DISPONIBILITÉ + MARGE (CUMUL, PRIORITAIRE)",
                   alerts[alerts["type_alerte"] == "Disponibilité + Marge"], RED_H)
    alert_section("4.  DISPONIBILITÉ", alerts[alerts["type_alerte"] == "Disponibilité"], NEUTRAL_H)
    alert_section("5.  MARGE", alerts[alerts["type_alerte"] == "Marge"], AMBER_H)

    autosize(ws, {'A': 9, 'B': 22, 'C': 32, 'D': 13, 'E': 9, 'F': 8, 'G': 10, 'H': 11, 'I': 10, 'J': 30})
    ws.freeze_panes = "A5"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ============================================================
# INTERFACE
# ============================================================
st.markdown("<div class='page-title'>🎛️ Cockpit Promo</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Alertes Disponibilité & Marge sur l'export PROMO · flag combiné pour les cas cumulés · synthèse par magasin</div>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("""
<div style='margin-bottom:18px'>
  <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>🛍️ SmartBuyer</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · Équipe Achats</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Import fichier</div>", unsafe_allow_html=True)
    up = st.file_uploader("Export PROMO (.csv)", type=["csv"], key="up_promo")
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Paramètres</div>", unsafe_allow_html=True)
    seuil = st.slider("Seuil marge faible (%)", 0, 30, MARGE_FAIBLE_DEFAUT, step=1)
    st.markdown("---")
    st.caption("SmartBuyer Hub · Module Cockpit Promo")

if up is None:
    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>ℹ️ À quoi sert ce module ?</strong><br>
  Ce module balaie un export PROMO et sort les lignes à risque de rupture et/ou de marge
  négative, avec un flag d'action par ligne et une synthèse par magasin.
  Un fichier à charger dans la barre latérale.
</div>
""", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("<div class='section-label'>Contenu du module</div>", unsafe_allow_html=True)
        st.markdown("""
<div class='card'>
  <div style='font-size:14px;font-weight:700;color:#1C1C1E;margin-bottom:8px'>📦 Disponibilité</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:1.5'>
    Stock ≤ 0 : "Rupture - Passer commande" (RAL = 0) ou "RAL - Accélérer livraison" (RAL &gt; 0).
  </div>
</div>
<div class='card'>
  <div style='font-size:14px;font-weight:700;color:#1C1C1E;margin-bottom:8px'>📉 Marge</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:1.5'>
    Marge % calculée depuis PMP, PV Promo et Taux TVA : vente à perte, marge faible
    (seuil réglable), ou PMP manquant.
  </div>
</div>
<div class='card'>
  <div style='font-size:14px;font-weight:700;color:#1C1C1E;margin-bottom:8px'>🚨 Cumul prioritaire</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:1.5'>
    Colonne Action combinée pour les articles qui cumulent les deux — à traiter en premier.
  </div>
</div>
""", unsafe_allow_html=True)
    with c2:
        st.markdown("<div class='section-label'>Types d'alerte</div>", unsafe_allow_html=True)
        st.markdown("""
<div class='format-card format-hyper'>
  <span class='badge badge-hyper'>Disponibilité</span>
  <div class='small-muted' style='margin-top:6px'>Rupture sans réappro ou réappro en cours</div>
</div>
<div class='format-card format-market'>
  <span class='badge badge-amber'>Marge faible</span>
  <div class='small-muted' style='margin-top:6px'>Sous le seuil réglable (10% par défaut)</div>
</div>
<div class='format-card format-supeco'>
  <span class='badge badge-red'>Vente à perte / Cumul</span>
  <div class='small-muted' style='margin-top:6px'>Marge négative, seule ou combinée à une rupture</div>
</div>
""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div class='section-label'>Fonctionnement</div>", unsafe_allow_html=True)
        st.markdown("""
<div class='alert-card alert-green'>
  <strong>1.</strong> Charge l'export PROMO dans la sidebar.<br>
  <strong>2.</strong> Ajuste si besoin le seuil de marge faible.<br>
  <strong>3.</strong> Consulte le cockpit et la synthèse par magasin.<br>
  <strong>4.</strong> Télécharge l'Excel (valeurs figées, prêt à diffuser).
</div>
""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Colonnes attendues</div>", unsafe_allow_html=True)
    col_left, col_right = st.columns(2)
    with col_left:
        for c in REQUIRED_COLS:
            st.markdown(f"""
<div class='col-required'>
  <div style='font-size:16px'>▪️</div>
  <div class='col-name'>{c}</div>
</div>""", unsafe_allow_html=True)
    with col_right:
        st.caption("Optionnelles (marge réalisée)")
        for c in OPTIONAL_COLS:
            st.markdown(f"""
<div class='col-required col-optional'>
  <div style='font-size:16px'>▪️</div>
  <div class='col-name'>{c}</div>
</div>""", unsafe_allow_html=True)

    st.info("⬅️ Charge ton export PROMO dans la barre latérale pour démarrer.")
    st.stop()

# ============================================================
# TRAITEMENT
# ============================================================
try:
    alerts = compute_alerts(up.getvalue(), seuil)
except Exception as e:
    st.error(f"Lecture du fichier impossible : {e}")
    st.stop()

if alerts.empty:
    st.success("Aucune alerte détectée sur ce périmètre — rien à traiter aujourd'hui.")
    st.stop()

synth = synthese_par_site(alerts)
line1, line2 = build_headline(alerts, synth)
st.markdown(
    f"<div class='alert-card alert-blue'><strong>{line1}</strong>"
    + (f"<br>{line2}" if line2 else "")
    + "</div>", unsafe_allow_html=True)

st.markdown("<div class='section-label'>Vue d'ensemble</div>", unsafe_allow_html=True)
n_dispo = int((alerts["type_alerte"] == "Disponibilité").sum())
n_marge = int((alerts["type_alerte"] == "Marge").sum())
n_cumul = int((alerts["type_alerte"] == "Disponibilité + Marge").sum())
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total alertes", len(alerts))
c2.metric("Disponibilité", n_dispo)
c3.metric("Marge", n_marge)
c4.metric("Cumul (critique)", n_cumul, delta_color="off")

st.markdown("<div class='section-label'>Synthèse par magasin</div>", unsafe_allow_html=True)
sty_synth = (synth.style
             .background_gradient(subset=["Vente à perte"], cmap="Reds")
             .background_gradient(subset=["Dispo + Marge"], cmap="Oranges"))
st.dataframe(sty_synth, use_container_width=True, hide_index=True)

st.markdown("<div class='section-label'>Détail des alertes</div>", unsafe_allow_html=True)

st.markdown(f"<span class='badge badge-red' style='margin:14px 0 6px;'>🚨 Disponibilité + Marge — prioritaire ({n_cumul})</span>", unsafe_allow_html=True)
render_table(alerts[alerts["type_alerte"] == "Disponibilité + Marge"].sort_values("Marge en cours"), seuil)

st.markdown(f"<span class='badge badge-hyper' style='margin:14px 0 6px;'>📦 Disponibilité ({n_dispo})</span>", unsafe_allow_html=True)
render_table(alerts[alerts["type_alerte"] == "Disponibilité"].sort_values("Stock"), seuil, height=320)

st.markdown(f"<span class='badge badge-amber' style='margin:14px 0 6px;'>📉 Marge ({n_marge})</span>", unsafe_allow_html=True)
render_table(alerts[alerts["type_alerte"] == "Marge"].sort_values("Marge en cours"), seuil, height=320)

st.markdown("<div class='section-label'>Export</div>", unsafe_allow_html=True)
xls = build_excel(alerts, synth)
today_str = _dt.date.today().strftime("%Y%m%d")
export_filename = f"Cockpit Promo - {today_str}.xlsx"
st.download_button(f"📥 Télécharger {export_filename}", xls, file_name=export_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption("Fichier à une feuille : synthèse réseau, synthèse par magasin, puis le détail des 3 sections d'alerte — valeurs figées, aucune formule. Couleur réservée aux sections à enjeu financier (Marge, Cumul) ; Disponibilité reste neutre.")
