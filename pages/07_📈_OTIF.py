"""
SmartBuyer · On Time In Full — v4
──────────────────────────────────
Page Streamlit de pilotage OTIF fournisseur.
Nouveauté v4 : onglet Fiche Fournisseur avec export Excel envoyable au fournisseur.
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date as _date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go

# ══════════════════════════════════════════════════════════════════════════════
# CONFIG PAGE
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="On Time In Full · SmartBuyer",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
# CHARTE GRAPHIQUE SMARTBUYER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display",
                 "SF Pro Text", "Helvetica Neue", Arial, sans-serif !important;
    background-color: #F2F2F7;
}
.stApp { background: #F2F2F7; }
.main .block-container { padding-top: 1.8rem; max-width: 1280px; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: #F2F2F7 !important;
    border-right: 0.5px solid #D1D1D6 !important;
}

/* Metric cards */
[data-testid="stMetric"] {
    background: #FFFFFF !important;
    border: 0.5px solid #E5E5EA !important;
    border-radius: 12px !important;
    padding: 16px 18px !important;
}
[data-testid="stMetricLabel"] {
    font-size: 11px !important; font-weight: 500 !important;
    color: #8E8E93 !important; text-transform: uppercase !important;
    letter-spacing: 0.04em !important;
}
[data-testid="stMetricValue"] {
    font-size: 24px !important; font-weight: 600 !important;
    color: #1C1C1E !important; letter-spacing: -0.02em !important;
}

/* Tabs */
[data-testid="stTabs"] button[role="tab"] {
    font-size: 13px !important; font-weight: 500 !important;
    padding: 8px 16px !important; color: #8E8E93 !important;
    border-bottom: 2px solid transparent !important;
}
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
    color: #007AFF !important; border-bottom: 2px solid #007AFF !important;
    background: transparent !important;
}
[data-testid="stTabs"] [role="tablist"] { border-bottom: 0.5px solid #E5E5EA !important; }

/* DataFrames */
[data-testid="stDataFrame"] {
    border: 0.5px solid #E5E5EA !important; border-radius: 10px !important;
}
[data-testid="stDataFrame"] th {
    background: #F2F2F7 !important; font-size: 11px !important;
    font-weight: 600 !important; color: #8E8E93 !important;
    text-transform: uppercase !important; letter-spacing: 0.04em !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    border: 1.5px dashed #D1D1D6 !important;
    border-radius: 10px !important; background: #F9F9FB !important;
}

/* Download buttons */
.stDownloadButton > button {
    background: #007AFF !important; color: white !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 500 !important; font-size: 13px !important;
    padding: 10px 24px !important; width: 100% !important;
}

hr { border-color: #E5E5EA !important; margin: 1rem 0 !important; }

/* Custom classes */
.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 1.5rem; }
.section-label {
    font-size: 11px; font-weight: 600; color: #8E8E93;
    text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 10px;
}
.alert-card  { padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; font-size: 13px; line-height: 1.6; border-left: 3px solid; }
.alert-red   { background: #FFF2F2; border-color: #FF3B30; color: #3A0000; }
.alert-amber { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }
.alert-green { background: #F0FFF4; border-color: #34C759; color: #003A10; }
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }

.kpi-focus {
    background: #EFF6FF; border: 1px solid #B3D9FF;
    border-radius: 12px; padding: 16px 18px;
}
.kpi-focus-label { font-size: 11px; font-weight: 500; color: #007AFF; text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 3px; }
.kpi-focus-value { font-size: 24px; font-weight: 700; color: #007AFF; letter-spacing: -0.02em; }
.kpi-focus-sub   { font-size: 12px; color: #0066CC; margin-top: 3px; font-weight: 500; }

.doc-card {
    background: #FFFFFF; border: 0.5px solid #E5E5EA;
    border-radius: 12px; padding: 20px 22px; margin-bottom: 12px;
}
.doc-card-title { font-size: 13px; font-weight: 700; color: #1C1C1E; margin-bottom: 8px; }
.doc-card-body  { font-size: 13px; color: #3A3A3C; line-height: 1.7; }
code { background: #F2F2F7; padding: 2px 6px; border-radius: 4px; font-size: 12px; }

.fiche-header {
    background: #FFFFFF; border: 0.5px solid #E5E5EA;
    border-radius: 14px; padding: 22px 26px; margin-bottom: 16px;
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES MÉTIER
# ══════════════════════════════════════════════════════════════════════════════
TECHNICAL_SUPPLIERS = {
    "FOURNISSEUR STOCK",
    "FOURNISSEUR PLATEFORME LOCAL",
    "FOURNISSEUR PLATEFORME IMPORT",
}

# Colonnes date prévue — testées dans l'ordre, première non-vide retenue
DATE_EXPECTED_CANDIDATES = ["H Date", "Date livraison", "Date prévue", "Date"]

# Seuils score global
SEUIL_EXCELLENT  = 97
SEUIL_SURVEILLER = 90


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS AFFICHAGE
# ══════════════════════════════════════════════════════════════════════════════
def fmt(n) -> str:
    """Formate un nombre en K / M lisible."""
    if pd.isna(n) or n is None:
        return "—"
    a = abs(float(n))
    if a >= 1_000_000:
        return f"{n / 1_000_000:.1f} M"
    if a >= 1_000:
        return f"{int(n / 1_000)} K"
    return f"{int(n):,}"


def fmt_pct(v, decimals: int = 1) -> str:
    if pd.isna(v) or v is None:
        return "—"
    return f"{v:.{decimals}f}%"


def score_band(v) -> str:
    """Niveau qualitatif basé sur le score global (0–100)."""
    if pd.isna(v):
        return "Inconnu"
    if v >= SEUIL_EXCELLENT:
        return "🟢 Excellent"
    if v >= SEUIL_SURVEILLER:
        return "🟠 À surveiller"
    return "🔴 Critique"


def score_color(v) -> str:
    if pd.isna(v):
        return "#8E8E93"
    if v >= SEUIL_EXCELLENT:
        return "#34C759"
    if v >= SEUIL_SURVEILLER:
        return "#FF9500"
    return "#FF3B30"


def safe_div(a, b) -> float:
    return a / b if b not in (0, None) and not pd.isna(b) else 0.0


# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT & NETTOYAGE ERP
# ══════════════════════════════════════════════════════════════════════════════
def detect_expected_date_column(df: pd.DataFrame):
    """Retourne la première colonne date prévue non entièrement nulle, ou None."""
    for col in DATE_EXPECTED_CANDIDATES:
        if col in df.columns:
            if pd.to_datetime(df[col], errors="coerce").notna().sum() > 0:
                return col
    return None


@st.cache_data(show_spinner=False)
def load_erp(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Charge le CSV depuis les bytes bruts.
    Clé de cache = (bytes, filename) → invalidation automatique à chaque nouvel upload.
    """
    df = pd.read_csv(BytesIO(file_bytes), sep=";", low_memory=False)
    df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]

    # Nettoyage colonnes parasites
    df = df.dropna(axis=1, how="all")
    df = df.drop(columns=[c for c in df.columns if c.startswith("Unnamed:")], errors="ignore")

    # Nettoyage chaînes de caractères
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip().replace({"nan": None, "": None})

    # Parsing dates
    date_cols = [
        "Dt Rec", "Date de commande", "Date", "Date facture",
        "Date comptable du rapprochement", "H Date", "Date livraison", "Date prévue",
    ]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")

    # Parsing numériques
    num_cols = [
        "Site", "Département", "N° Cde", "Sit", "Fou", "Famille", "Sous-famille",
        "Code", "Article", "Qté cde", "Poids cde", "Qté rec", "Poids rec",
        "PV", "Px revient", "Prix de vente HT", "Marge ligne", "Taux TVA",
        "EAN", "Poids unitaire", "PCB", "Colis",
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def prepare_dataset(df: pd.DataFrame, exclude_technical: bool = True, cap_sur_receipt: bool = True):
    """
    Applique toutes les règles métier OTIF.
    Retourne (df_clean, quality_dict).
    """
    if df.empty:
        return df.copy(), {}

    work = df.copy()
    raw_len = len(work)

    # ── Labels normalisés
    def _col(name):
        return work[name] if name in work.columns else pd.Series("Inconnu", index=work.index)

    work["site_label"]    = _col("Libellé site").fillna("Inconnu").astype(str).str.strip()
    work["supplier_name"] = _col("Nom fourn.").fillna("Inconnu").astype(str).str.strip()
    work["article_label"] = _col("Libellé article").fillna("Inconnu").astype(str).str.strip()
    work["dept_label"]    = _col("Libellé département").fillna("Inconnu").astype(str).str.strip()
    work["famille_label"] = _col("Libellé famille").fillna("Inconnu").astype(str).str.strip()

    # ── Quantités & prix
    work["qte_cde"] = pd.to_numeric(work.get("Qté cde"), errors="coerce").fillna(0)
    work["qte_rec"] = pd.to_numeric(work.get("Qté rec"), errors="coerce").fillna(0)
    work["pv_ht"]   = pd.to_numeric(work.get("Prix de vente HT"), errors="coerce").fillna(0)
    pv_zero_rows    = int((work["pv_ht"] == 0).sum())

    # ── Compteurs avant exclusion
    is_tech     = work["supplier_name"].str.upper().isin(TECHNICAL_SUPPLIERS)
    is_zero_qty = work["qte_cde"] <= 0
    sur_receipt = int((work["qte_rec"] > work["qte_cde"]).sum())
    excl_tech   = int(is_tech.sum())
    excl_zero   = int(is_zero_qty.sum())

    # ── Application exclusions
    if exclude_technical:
        work = work[~is_tech].copy()
    work = work[work["qte_cde"] > 0].copy()

    # ── Cap sur-réception
    work["qte_rec_retained"] = (
        work[["qte_rec", "qte_cde"]].min(axis=1) if cap_sur_receipt else work["qte_rec"]
    )

    # ── Détection colonne date prévue
    expected_col = detect_expected_date_column(work)
    work["date_expected"] = work[expected_col] if expected_col else pd.NaT
    work["date_received"] = work["Dt Rec"] if "Dt Rec" in work.columns else pd.NaT

    missing_exp = int(work["date_expected"].isna().sum())

    # ── Métriques ligne
    work["qty_missing"]       = (work["qte_cde"] - work["qte_rec_retained"]).clip(lower=0)
    work["service_gap_value"] = work["qty_missing"] * work["pv_ht"]
    work["line_fill_rate"]    = np.where(
        work["qte_cde"] > 0, work["qte_rec_retained"] / work["qte_cde"], 0.0
    )
    work["delay_days"] = (work["date_received"] - work["date_expected"]).dt.days

    # Règle temporaire : date manquante → On Time par défaut
    work["on_time"] = work["date_expected"].isna() | (work["date_received"] <= work["date_expected"])
    work["otif"]    = ((work["qte_rec_retained"] >= work["qte_cde"]) & work["on_time"]).astype(int)

    # Criticité ligne = impact CA × taux de défaillance
    work["criticality_score"] = work["service_gap_value"] * (1 - work["line_fill_rate"])

    quality = {
        "raw_rows":              raw_len,
        "clean_rows":            len(work),
        "excluded_zero_qty":     excl_zero,
        "excluded_technical":    excl_tech,
        "sur_receipt_rows":      sur_receipt,
        "missing_expected_date": missing_exp,
        "all_dates_missing":     (missing_exp == len(work)),
        "pv_zero_rows":          pv_zero_rows,
        "expected_col":          expected_col or "Aucune",
        "usable_rate":           round(safe_div(len(work), raw_len) * 100, 1),
    }
    return work, quality


# ══════════════════════════════════════════════════════════════════════════════
# KPI GLOBAUX
# ══════════════════════════════════════════════════════════════════════════════
def compute_global_kpis(df: pd.DataFrame) -> dict:
    if df.empty:
        return dict(fill_rate=0, on_time=0, otif=0, score=0,
                    ordered_qty=0, received_qty=0, missing_qty=0,
                    impact_value=0, suppliers=0, articles=0, orders=0, sites=0)

    ordered   = df["qte_cde"].sum()
    received  = df["qte_rec_retained"].sum()
    fill_rate = safe_div(received, ordered) * 100
    on_time   = df["on_time"].mean() * 100
    otif      = df["otif"].mean() * 100
    score     = 0.5 * fill_rate + 0.3 * on_time + 0.2 * otif

    return dict(
        fill_rate=fill_rate, on_time=on_time, otif=otif, score=score,
        ordered_qty=ordered, received_qty=received,
        missing_qty=df["qty_missing"].sum(),
        impact_value=df["service_gap_value"].sum(),
        suppliers=df["supplier_name"].nunique(),
        articles=df["Code"].nunique() if "Code" in df.columns else 0,
        orders=df["N° Cde"].nunique() if "N° Cde" in df.columns else 0,
        sites=df["site_label"].nunique(),
    )


# ══════════════════════════════════════════════════════════════════════════════
# AGRÉGATIONS
# ══════════════════════════════════════════════════════════════════════════════
def _enrich(g: pd.DataFrame) -> pd.DataFrame:
    """Ajoute fill_rate%, on_time%, otif%, score et Niveau à un groupe agrégé."""
    g["fill_rate"] = np.where(g["qte_cde"] > 0, g["qte_rec"] / g["qte_cde"] * 100, 0.0)
    g["on_time"]  *= 100
    g["otif"]     *= 100
    g["score"]     = 0.5 * g["fill_rate"] + 0.3 * g["on_time"] + 0.2 * g["otif"]
    g["Niveau"]    = g["score"].apply(score_band)
    # Criticité agrégée = impact CA proxy × taux de défaillance
    g["criticality_score"] = g["impact_value"] * (1 - g["fill_rate"] / 100)
    return g


def agg_supplier(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["Fou", "supplier_name"], as_index=False).agg(
        qte_cde      =("qte_cde",          "sum"),
        qte_rec      =("qte_rec_retained",  "sum"),
        qty_missing  =("qty_missing",       "sum"),
        impact_value =("service_gap_value", "sum"),
        on_time      =("on_time",           "mean"),
        otif         =("otif",              "mean"),
        orders       =("N° Cde",            "nunique"),
        articles     =("Code",              "nunique"),
        sites        =("site_label",        "nunique"),
    )
    return _enrich(g).sort_values("criticality_score", ascending=False).reset_index(drop=True)


def agg_site(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["Site", "site_label"], as_index=False).agg(
        qte_cde      =("qte_cde",          "sum"),
        qte_rec      =("qte_rec_retained",  "sum"),
        qty_missing  =("qty_missing",       "sum"),
        impact_value =("service_gap_value", "sum"),
        on_time      =("on_time",           "mean"),
        otif         =("otif",              "mean"),
        suppliers    =("supplier_name",     "nunique"),
        articles     =("Code",             "nunique"),
    )
    return _enrich(g).sort_values("criticality_score", ascending=False).reset_index(drop=True)


def agg_article(df: pd.DataFrame) -> pd.DataFrame:
    gcols = [c for c in ["Code", "article_label", "supplier_name"] if c in df.columns]
    g = df.groupby(gcols, as_index=False).agg(
        qte_cde      =("qte_cde",          "sum"),
        qte_rec      =("qte_rec_retained",  "sum"),
        qty_missing  =("qty_missing",       "sum"),
        impact_value =("service_gap_value", "sum"),
        on_time      =("on_time",           "mean"),
        otif         =("otif",              "mean"),
        sites        =("site_label",        "nunique"),
        orders       =("N° Cde",            "nunique"),
    )
    return _enrich(g).sort_values("criticality_score", ascending=False).reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
# GRAPHIQUES PLOTLY
# ══════════════════════════════════════════════════════════════════════════════
_PLOTLY_BASE = dict(
    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    font=dict(family="-apple-system, Helvetica Neue", color="#3A3A3C", size=11),
    margin=dict(t=10, b=10, l=10, r=80),
    xaxis=dict(showgrid=True, gridcolor="#F2F2F7"),
    yaxis=dict(showgrid=False, title=""),
)


def bar_h(data: pd.DataFrame, x_col: str, y_col: str, color: str,
          x_title: str, height: int = 500, fmt_fn=None) -> go.Figure:
    top = data.head(15).sort_values(x_col)
    texts = [fmt_fn(v) if fmt_fn else f"{v:,.0f}" for v in top[x_col]]
    fig = go.Figure(go.Bar(
        x=top[x_col], y=top[y_col].astype(str),
        orientation="h", marker_color=color, marker_line_width=0,
        text=texts, textposition="outside",
    ))
    fig.update_layout(**{**_PLOTLY_BASE, "height": height,
                         "xaxis": {**_PLOTLY_BASE["xaxis"], "title": x_title}})
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# COMPOSANT KPI ROW
# ══════════════════════════════════════════════════════════════════════════════
def render_kpi_row(kpi: dict):
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Fill Rate",    fmt_pct(kpi["fill_rate"]))
    c2.metric("On Time",      fmt_pct(kpi["on_time"]))
    c3.metric("OTIF",         fmt_pct(kpi["otif"]))
    c4.metric("Score global", fmt_pct(kpi["score"]))
    with c5:
        st.markdown(f"""
<div class='kpi-focus'>
  <div class='kpi-focus-label'>Volume manquant</div>
  <div class='kpi-focus-value'>{fmt(kpi['missing_qty'])}</div>
  <div class='kpi-focus-sub'>Impact CA proxy : {fmt(kpi['impact_value'])} FCFA</div>
</div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL — SYNTHÈSE GLOBALE
# ══════════════════════════════════════════════════════════════════════════════
def _xl_write_sheet(ws, title: str, df: pd.DataFrame,
                    hdr_fill=None, hdr_font=None, ctr=None):
    """Écrit titre + entêtes + données dans un onglet openpyxl."""
    ws.append([title])
    ws.cell(1, 1).font = Font(bold=True, size=13, color="1C3557")
    ws.append([])
    headers = list(df.columns)
    ws.append(headers)
    for i, _ in enumerate(headers, 1):
        c = ws.cell(3, i)
        if hdr_fill: c.fill = hdr_fill
        if hdr_font: c.font = hdr_font
        if ctr:      c.alignment = ctr
    for row in df.itertuples(index=False):
        ws.append(list(row))
    # Ajustement largeur robuste
    for col_cells in ws.iter_cols(min_row=3, max_row=3):
        letter = get_column_letter(col_cells[0].column)
        val = str(col_cells[0].value or "")
        ws.column_dimensions[letter].width = max(12, min(32, len(val) + 4))


def build_export_excel(df, by_supplier, by_site, by_article, quality) -> BytesIO:
    wb = Workbook()
    H_FILL = PatternFill("solid", fgColor="1C3557")
    H_FONT = Font(bold=True, color="FFFFFF", size=11)
    CTR    = Alignment(horizontal="center", vertical="center")

    # Synthèse qualité
    ws1 = wb.active
    ws1.title = "Synthese"
    synthese = pd.DataFrame([
        ["Lignes brutes",                      quality.get("raw_rows", 0)],
        ["Lignes exploitables",                quality.get("clean_rows", 0)],
        ["Taux exploitable %",                 quality.get("usable_rate", 0)],
        ["Date prévue utilisée",               quality.get("expected_col", "Aucune")],
        ["Toutes dates prévues manquantes",     "OUI" if quality.get("all_dates_missing") else "NON"],
        ["Qté cde ≤ 0 exclues",                quality.get("excluded_zero_qty", 0)],
        ["Fournisseurs techniques exclus",      quality.get("excluded_technical", 0)],
        ["Sur-réceptions capées",              quality.get("sur_receipt_rows", 0)],
        ["Dates prévues manquantes",            quality.get("missing_expected_date", 0)],
        ["Lignes PV HT = 0 (impact nul)",      quality.get("pv_zero_rows", 0)],
    ], columns=["Indicateur", "Valeur"])
    _xl_write_sheet(ws1, "Synthèse qualité de données", synthese, H_FILL, H_FONT, CTR)

    _xl_write_sheet(wb.create_sheet("Par fournisseur"), "Analyse fournisseur", by_supplier, H_FILL, H_FONT, CTR)
    _xl_write_sheet(wb.create_sheet("Par magasin"),     "Analyse magasin",     by_site,     H_FILL, H_FONT, CTR)
    _xl_write_sheet(wb.create_sheet("Par article"),     "Analyse article",     by_article,  H_FILL, H_FONT, CTR)

    # Lignes critiques (top 500 non OTIF)
    detail_cols = [c for c in [
        "date_received", "date_expected", "site_label", "supplier_name",
        "Code", "article_label", "N° Cde", "qte_cde", "qte_rec_retained",
        "qty_missing", "service_gap_value", "on_time", "otif", "delay_days",
    ] if c in df.columns]
    crit = (df[df["otif"] == 0][detail_cols]
            .sort_values(["qty_missing", "service_gap_value"], ascending=[False, False])
            .head(500))
    _xl_write_sheet(wb.create_sheet("Lignes critiques"), "Lignes non OTIF", crit, H_FILL, H_FONT, CTR)

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL — FICHE FOURNISSEUR
# ══════════════════════════════════════════════════════════════════════════════
def build_fiche_excel(fournisseur: str, df_all: pd.DataFrame, df_bad: pd.DataFrame,
                      site_recap: pd.DataFrame, art_recap: pd.DataFrame,
                      kpis: dict, seuil: int) -> BytesIO:
    """
    Génère un Excel fiche fournisseur propre et envoyable.
    4 onglets : Synthèse KPI · Par magasin · Par article · Détail lignes
    """
    wb = Workbook()
    today_str = _date.today().strftime("%d/%m/%Y")

    # Styles
    H_FILL   = PatternFill("solid", fgColor="1C3557")
    H_FONT   = Font(bold=True, color="FFFFFF", size=10)
    T_FONT   = Font(bold=True, size=13, color="1C3557")
    S_FONT   = Font(bold=True, size=10)
    CTR      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LFT      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    RED_FILL = PatternFill("solid", fgColor="FFF2F2")
    AMB_FILL = PatternFill("solid", fgColor="FFFBF0")
    GRN_FILL = PatternFill("solid", fgColor="F0FFF4")
    THIN     = Border(
        left=Side(style="thin", color="E5E5EA"),
        right=Side(style="thin", color="E5E5EA"),
        top=Side(style="thin", color="E5E5EA"),
        bottom=Side(style="thin", color="E5E5EA"),
    )

    def auto_width(ws, min_w=12, max_w=34):
        for col in ws.iter_cols():
            best = min_w
            for cell in col:
                if cell.value:
                    best = max(best, min(max_w, len(str(cell.value)) + 3))
            ws.column_dimensions[get_column_letter(col[0].column)].width = best

    def write_header_row(ws, row_n, headers):
        for i, h in enumerate(headers, 1):
            c = ws.cell(row_n, i, h)
            c.fill = H_FILL; c.font = H_FONT
            c.alignment = CTR; c.border = THIN

    def kpi_fill(val_str: str):
        """Retourne le fill approprié pour une valeur en %."""
        try:
            v = float(val_str.replace("%", ""))
            return GRN_FILL if v >= SEUIL_EXCELLENT else (AMB_FILL if v >= SEUIL_SURVEILLER else RED_FILL)
        except Exception:
            return None

    # ─────────────────────────────────────────────
    # Onglet 1 : Synthèse KPI
    # ─────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Synthèse KPI"
    ws1.row_dimensions[1].height = 30
    ws1.row_dimensions[2].height = 18

    ws1.cell(1, 1, f"BILAN DE PERFORMANCE — {fournisseur.upper()}").font = Font(bold=True, size=14, color="1C1C1E")
    ws1.cell(2, 1, f"Généré le {today_str}  ·  Période : export ERP  ·  Seuil de détection : Fill Rate < {seuil}%").font = Font(size=10, italic=True, color="8E8E93")
    ws1.merge_cells("A1:D1")
    ws1.merge_cells("A2:D2")
    ws1.append([])

    kpi_data = [
        ("Fill Rate",       fmt_pct(kpis["fill_rate"]),    "Part de la qté commandée effectivement livrée"),
        ("On Time",         fmt_pct(kpis["on_time"]),      "Part des livraisons respectant la date prévue"),
        ("OTIF",            fmt_pct(kpis["otif"]),         "Livraisons complètes ET à l'heure"),
        ("Score global",    fmt_pct(kpis["score"]),        "50% × Fill Rate + 30% × On Time + 20% × OTIF"),
        ("Niveau",          score_band(kpis["score"]),     "Évaluation synthétique"),
        ("Vol. manquant",   f"{int(kpis['missing_qty']):,} unités",  "Qté commandée − Qté reçue retenue"),
        ("Impact CA proxy", f"{kpis['impact_value']:,.0f} FCFA",     "Vol. manquant × Prix de vente HT"),
        ("Commandes",       str(kpis["orders"]),           "Nombre de N° de commandes distincts"),
        ("Articles",        str(kpis["articles"]),         "Références concernées"),
        ("Magasins",        str(kpis["sites"]),            "Sites livrés sur la période"),
    ]
    write_header_row(ws1, 4, ["Indicateur", "Valeur", "Commentaire"])
    for i, (label, val, comment) in enumerate(kpi_data, 5):
        ws1.cell(i, 1, label).font = S_FONT
        ws1.cell(i, 1).border = THIN
        c_val = ws1.cell(i, 2, val)
        c_val.alignment = CTR; c_val.border = THIN
        if label in ("Fill Rate", "On Time", "OTIF", "Score global"):
            f = kpi_fill(val)
            if f: c_val.fill = f
        c_com = ws1.cell(i, 3, comment)
        c_com.alignment = LFT; c_com.border = THIN
    auto_width(ws1)

    # ─────────────────────────────────────────────
    # Onglet 2 : Par magasin
    # ─────────────────────────────────────────────
    ws2 = wb.create_sheet("Par magasin")
    ws2.cell(1, 1, f"Livraisons incomplètes par magasin — seuil Fill Rate < {seuil}%").font = T_FONT
    ws2.append([])
    headers2 = ["Magasin", "Fill Rate", "Vol. manquant", "Impact CA proxy (FCFA)", "Cmdes", "Articles"]
    write_header_row(ws2, 3, headers2)
    for _, r in site_recap.sort_values("qty_missing", ascending=False).iterrows():
        row_data = [
            r.get("site_label", ""),
            r.get("Fill Rate", ""),
            r.get("Vol. manquant", ""),
            r.get("Impact CA proxy", ""),
            r.get("nb_cdes", ""),
            r.get("nb_articles", ""),
        ]
        n = ws2.max_row + 1
        ws2.append(row_data)
        for col_i in range(1, 7):
            ws2.cell(n, col_i).border = THIN
            ws2.cell(n, col_i).alignment = CTR if col_i > 1 else LFT
        fill_val = r.get("Fill Rate", "")
        f = kpi_fill(str(fill_val)) if fill_val else None
        if f: ws2.cell(n, 2).fill = f
    auto_width(ws2)

    # ─────────────────────────────────────────────
    # Onglet 3 : Par article
    # ─────────────────────────────────────────────
    ws3 = wb.create_sheet("Par article")
    ws3.cell(1, 1, f"Livraisons incomplètes par article — seuil Fill Rate < {seuil}%").font = T_FONT
    ws3.append([])
    headers3 = ["Code article", "Désignation", "Fill Rate", "Vol. manquant",
                "Impact CA proxy (FCFA)", "Magasins", "Cmdes"]
    write_header_row(ws3, 3, headers3)
    for _, r in art_recap.iterrows():
        row_data = [
            r.get("Code", ""),
            r.get("article_label", ""),
            r.get("Fill Rate", ""),
            r.get("Vol. manquant", ""),
            r.get("Impact CA proxy", ""),
            r.get("nb_sites", ""),
            r.get("nb_cdes", ""),
        ]
        n = ws3.max_row + 1
        ws3.append(row_data)
        for col_i in range(1, 8):
            ws3.cell(n, col_i).border = THIN
            ws3.cell(n, col_i).alignment = LFT if col_i <= 2 else CTR
        fill_val = r.get("Fill Rate", "")
        f = kpi_fill(str(fill_val)) if fill_val else None
        if f: ws3.cell(n, 3).fill = f
    auto_width(ws3)

    # ─────────────────────────────────────────────
    # Onglet 4 : Détail lignes
    # ─────────────────────────────────────────────
    ws4 = wb.create_sheet("Détail lignes")
    ws4.cell(1, 1, f"Détail des livraisons incomplètes — {fournisseur}  (Fill Rate < {seuil}%)").font = T_FONT
    ws4.append([])
    headers4 = [
        "N° Commande", "Date réception", "Date prévue", "Magasin",
        "Code article", "Désignation article",
        "Qté commandée", "Qté reçue", "Qté manquante",
        "Fill Rate ligne", "Impact CA proxy (FCFA)",
        "Livré à l'heure", "Retard (jours)",
    ]
    write_header_row(ws4, 3, headers4)

    src_cols = [c for c in [
        "N° Cde", "date_received", "date_expected", "site_label",
        "Code", "article_label", "qte_cde", "qte_rec_retained",
        "qty_missing", "line_fill_rate", "service_gap_value", "on_time", "delay_days",
    ] if c in df_bad.columns]

    for _, row in df_bad[src_cols].iterrows():
        fr   = row.get("line_fill_rate", 0)
        ot   = row.get("on_time", True)
        n    = ws4.max_row + 1
        ws4.append([
            row.get("N° Cde", ""),
            row.get("date_received", pd.NaT),
            row.get("date_expected", pd.NaT),
            row.get("site_label", ""),
            row.get("Code", ""),
            row.get("article_label", ""),
            int(row.get("qte_cde", 0)),
            int(row.get("qte_rec_retained", 0)),
            int(row.get("qty_missing", 0)),
            f"{fr * 100:.1f}%" if pd.notna(fr) else "—",
            round(row.get("service_gap_value", 0), 0),
            "OUI" if ot else "NON",
            int(row.get("delay_days", 0)) if pd.notna(row.get("delay_days")) else "—",
        ])
        row_fill = RED_FILL if pd.notna(fr) and fr < 0.90 else None
        for col_i in range(1, len(headers4) + 1):
            cell = ws4.cell(n, col_i)
            cell.border = THIN
            cell.alignment = LFT if col_i in (4, 6) else CTR
            if row_fill and col_i in (9, 10, 11):  # colonnes manque & impact
                cell.fill = row_fill
        # Cellule Fill Rate colorée individuellement
        fr_fill = kpi_fill(f"{fr*100:.1f}%") if pd.notna(fr) else None
        if fr_fill:
            ws4.cell(n, 10).fill = fr_fill
    auto_width(ws4)

    # Freeze panes ligne 3 (en-têtes fixes)
    for ws in [ws2, ws3, ws4]:
        ws.freeze_panes = "A4"

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:18px'>
  <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>📦 SmartBuyer</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · On Time In Full</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    # ── Import
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Import fichier</div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Extraction ERP OTIF (CSV ; )", type=["csv"], key="otif")
    st.markdown("---")

    # ── Règles métier
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Règles métier</div>", unsafe_allow_html=True)
    exclude_technical = st.checkbox("Exclure fournisseurs techniques", value=True)
    cap_sur_receipt   = st.checkbox("Caper Qté reçue ≤ Qté commandée", value=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<div class='page-title'>📦 On Time In Full</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Pilotage OTIF fournisseur · Magasin · Article · Fiche fournisseur</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# ÉCRAN D'ACCUEIL — MÉTHODOLOGIE
# ══════════════════════════════════════════════════════════════════════════════
if uploaded is None:
    st.markdown("---")
    st.markdown("<div class='section-label'>Méthodologie de calcul</div>", unsafe_allow_html=True)

    docs = [
        ("① Fill Rate — Taux de service quantitatif",
         """Mesure la part de la quantité commandée effectivement reçue.<br>
<code>Fill Rate = Qté reçue retenue / Qté commandée × 100</code><br><br>
<strong>Règle cap :</strong> si la quantité reçue dépasse la quantité commandée,
elle est ramenée à la quantité commandée pour éviter un taux &gt; 100%.<br>
<code>Qté reçue retenue = min(Qté reçue, Qté commandée)</code><br>
<span style='color:#34C759;font-weight:600'>✓ Objectif cible : ≥ 97%</span>"""),

        ("② On Time — Taux de respect du délai",
         """Mesure la proportion de lignes livrées à la date prévue ou avant.<br>
<code>On Time = 1 si Date réception ≤ Date prévue, sinon 0</code><br><br>
<strong>Colonne date prévue (ordre de priorité) :</strong>
<code>H Date</code> → <code>Date livraison</code> → <code>Date prévue</code> → <code>Date</code><br><br>
<strong>⚠️ Règle temporaire ERP :</strong> si la date prévue est absente,
la ligne est considérée <em>On Time</em> par défaut.
Si aucune date n'est dispo dans l'export, le On Time sera <strong>100% artificiel</strong>.<br>
<span style='color:#34C759;font-weight:600'>✓ Objectif cible : ≥ 95%</span>"""),

        ("③ OTIF — On Time In Full",
         """Indicateur synthétique : une livraison est OTIF uniquement si elle est
<strong>à la fois complète ET à l'heure</strong>.<br>
<code>OTIF = 1 si (Qté reçue retenue ≥ Qté commandée) ET (On Time = 1), sinon 0</code><br><br>
Le taux OTIF global est la moyenne des OTIF ligne par ligne.<br>
<span style='color:#34C759;font-weight:600'>✓ Objectif cible : ≥ 95%</span>"""),

        ("④ Score global — Synthèse pondérée",
         """Agrège les trois indicateurs avec des pondérations reflétant leur priorité métier.<br>
<code>Score = 50% × Fill Rate + 30% × On Time + 20% × OTIF</code><br><br>
🟢 <strong>Excellent</strong> ≥ 97% &nbsp;|&nbsp; 🟠 <strong>À surveiller</strong> 90–97% &nbsp;|&nbsp; 🔴 <strong>Critique</strong> &lt; 90%<br><br>
<strong>Exemple :</strong> fournisseur avec 850/1 000 livrées, 7/10 lignes à l'heure, 5/10 OTIF<br>
→ Fill Rate = <strong>85%</strong> · On Time = <strong>70%</strong> · OTIF = <strong>50%</strong><br>
→ Score = 50%×85 + 30%×70 + 20%×50 = 42.5 + 21 + 10 = <strong>73.5% 🔴 Critique</strong>"""),

        ("⑤ Score de criticité — Priorisation opérationnelle",
         """Classe les fournisseurs par ordre de priorité d'action en combinant
<strong>l'impact financier du manque</strong> et <strong>le taux de défaillance</strong>.<br>
<code>Criticité = Impact CA proxy × (1 − Fill Rate)</code><br>
<code>Impact CA proxy = Qté manquante × Prix de vente HT</code><br><br>
<strong>Pourquoi cette formule ?</strong> Un fournisseur avec fort volume manquant
sur articles bon marché sera moins prioritaire qu'un fournisseur avec moins de manque
sur des références à forte valeur.<br><br>
<strong>Exemple :</strong><br>
Fournisseur A — Impact 50 M FCFA, Fill Rate 80% → Criticité = 50M × 20% = <strong>10 M</strong><br>
Fournisseur B — Impact 30 M FCFA, Fill Rate 20% → Criticité = 30M × 80% = <strong>24 M</strong><br>
→ B est prioritaire malgré un impact brut inférieur car il est <strong>beaucoup moins fiable</strong>."""),

        ("⑥ Impact CA proxy — Valorisation du manque",
         """Traduit le volume non livré en valeur commerciale potentiellement perdue.<br>
<code>Impact CA proxy = Qté manquante × Prix de vente HT</code><br><br>
Estimation haute : on suppose que toutes les unités manquantes auraient été vendues
au prix catalogue. Ne tient pas compte des stocks disponibles en magasin.<br>
<span style='color:#FF9500;font-weight:600'>⚠️ Si PV HT = 0 pour certaines lignes, l'impact est sous-estimé.</span>"""),
    ]

    for title, body in docs:
        st.markdown(f"""
<div class='doc-card'>
  <div class='doc-card-title'>{title}</div>
  <div class='doc-card-body'>{body}</div>
</div>""", unsafe_allow_html=True)

    st.markdown("""
<div class='alert-card alert-amber'>
  <strong>⚠️ Règle temporaire ERP active sur toutes les analyses</strong><br>
  Si la date prévue est absente ou nulle, la ligne est considérée <strong>On Time</strong> par défaut.
  Demandez à l'IT un export incluant la colonne <code>H Date</code> ou <code>Date livraison</code>.
</div>""", unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT DONNÉES
# ══════════════════════════════════════════════════════════════════════════════
with st.spinner("Lecture et préparation des données…"):
    file_bytes = uploaded.read()
    raw = load_erp(file_bytes, uploaded.name)
    df, quality = prepare_dataset(raw, exclude_technical=exclude_technical, cap_sur_receipt=cap_sur_receipt)

if df.empty:
    st.error("Aucune ligne exploitable après nettoyage. Vérifiez le format du fichier.")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# FILTRES SIDEBAR (après chargement)
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Filtres globaux</div>", unsafe_allow_html=True)

    all_sites   = sorted([x for x in df["site_label"].dropna().unique()   if x not in ("", "Inconnu")])
    all_sup     = sorted([x for x in df["supplier_name"].dropna().unique() if x not in ("", "Inconnu")])
    all_depts   = sorted([x for x in df["dept_label"].dropna().unique()    if x not in ("", "Inconnu")])

    sel_sites   = st.multiselect("Magasin",     all_sites, default=all_sites)
    sel_sup     = st.multiselect("Fournisseur", all_sup,   default=[])
    sel_depts   = st.multiselect("Département", all_depts, default=all_depts)
    only_crit   = st.checkbox("Uniquement OTIF = 0", value=False)

    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>📋 Fiche Fournisseur</div>", unsafe_allow_html=True)

    fiche_supplier = st.selectbox(
        "Fournisseur à analyser",
        options=["— Choisir —"] + all_sup,
        key="fiche_sel",
    )
    seuil_fill = st.slider(
        "Seuil Fill Rate (lignes 'mauvaises')",
        min_value=50, max_value=100, value=100, step=5, format="%d%%",
        help="Sont incluses dans la fiche toutes les lignes dont le Fill Rate ligne est STRICTEMENT inférieur à ce seuil.",
    )
    st.caption(f"Lignes retenues dans la fiche : Fill Rate ligne < {seuil_fill}%")

# ── Garde-fous filtres vides
if not sel_sites: sel_sites = all_sites
if not sel_depts: sel_depts = all_depts

view = df[df["site_label"].isin(sel_sites) & df["dept_label"].isin(sel_depts)].copy()
if sel_sup:
    view = view[view["supplier_name"].isin(sel_sup)].copy()
if only_crit:
    view = view[view["otif"] == 0].copy()

if view.empty:
    st.warning("Aucune donnée après filtrage — ajustez les filtres dans la sidebar.")
    st.stop()

# ── Calculs agrégés
kpi         = compute_global_kpis(view)
by_supplier = agg_supplier(view)
by_site     = agg_site(view)
by_article  = agg_article(view)


# ══════════════════════════════════════════════════════════════════════════════
# KPI GLOBAUX
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(
    f"<div class='section-label'>{kpi['sites']} magasin(s) · "
    f"{kpi['suppliers']} fournisseur(s) · {kpi['orders']} commande(s)</div>",
    unsafe_allow_html=True,
)
render_kpi_row(kpi)


# ══════════════════════════════════════════════════════════════════════════════
# ALERTES
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("<div class='section-label'>Alertes &amp; points d'attention</div>", unsafe_allow_html=True)

if quality.get("all_dates_missing"):
    st.markdown("""
<div class='alert-card alert-red'>
  <strong>🔴 On Time non significatif — aucune date prévue dans ce fichier</strong><br>
  On Time = 100% artificiel. L'OTIF ne reflète que le Fill Rate.
  Demandez un export avec la colonne <code>H Date</code> ou <code>Date livraison</code>.
</div>""", unsafe_allow_html=True)
elif quality.get("missing_expected_date", 0) > 0:
    pct = round(quality["missing_expected_date"] / quality["clean_rows"] * 100, 1)
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ Dates prévues partiellement manquantes</strong><br>
  {quality['missing_expected_date']:,} lignes sans date prévue ({pct}%) → considérées On Time par défaut.
</div>""", unsafe_allow_html=True)

if not by_supplier.empty:
    top3_sup = by_supplier.head(3)
    st.markdown(f"""
<div class='alert-card alert-red'>
  <strong>🔴 Fournisseurs les plus critiques</strong><br>
  {" · ".join(f"<strong>{r.supplier_name}</strong> (score {r.score:.1f}%)" for r in top3_sup.itertuples())}
</div>""", unsafe_allow_html=True)

if not by_site.empty:
    top3_site = by_site.head(3)
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ Magasins les plus impactés</strong><br>
  {" · ".join(f"<strong>{r.site_label}</strong> ({int(r.qty_missing):,} unités manquantes)" for r in top3_site.itertuples())}
</div>""", unsafe_allow_html=True)

if quality.get("pv_zero_rows", 0) > 0:
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ Impact CA proxy sous-estimé</strong><br>
  {quality['pv_zero_rows']:,} ligne(s) avec Prix de vente HT = 0 → impact CA proxy = 0 FCFA pour ces lignes.
</div>""", unsafe_allow_html=True)

st.markdown(f"""
<div class='alert-card alert-blue'>
  <strong>ℹ️ Qualité de données</strong> &nbsp;·&nbsp;
  Date prévue : <strong>{quality['expected_col']}</strong> &nbsp;·&nbsp;
  Dates manquantes : <strong>{quality['missing_expected_date']:,}</strong> &nbsp;·&nbsp;
  Sur-réceptions capées : <strong>{quality['sur_receipt_rows']:,}</strong> &nbsp;·&nbsp;
  Taux exploitable : <strong>{quality['usable_rate']:.1f}%</strong>
</div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    f"🚛 Fournisseurs ({len(by_supplier)})",
    f"🏪 Magasins ({len(by_site)})",
    f"📦 Articles ({len(by_article)})",
    "🚨 Lignes critiques",
    "🧪 Qualité des données",
    "📋 Fiche Fournisseur",
])


# ── Tab 1 : Fournisseurs ──────────────────────────────────────────────────────
with tab1:
    st.caption("Trié par criticité décroissante = Impact CA proxy × (1 − Fill Rate)")
    st.plotly_chart(
        bar_h(by_supplier, "criticality_score", "supplier_name", "#FF3B30",
              "Score de criticité (FCFA-équivalent)"),
        use_container_width=True,
    )
    d = by_supplier.copy()
    d["Fill Rate"]       = d["fill_rate"].apply(fmt_pct)
    d["On Time"]         = d["on_time"].apply(fmt_pct)
    d["OTIF"]            = d["otif"].apply(fmt_pct)
    d["Score"]           = d["score"].apply(fmt_pct)
    d["Vol. manquant"]   = d["qty_missing"].apply(fmt)
    d["Impact CA proxy"] = d["impact_value"].apply(fmt)
    st.dataframe(
        d[["supplier_name", "Fill Rate", "On Time", "OTIF", "Score",
           "Vol. manquant", "Impact CA proxy", "orders", "articles", "sites", "Niveau"]]
        .rename(columns={"supplier_name": "Fournisseur", "orders": "Cmdes",
                         "articles": "Articles", "sites": "Magasins"}),
        use_container_width=True, hide_index=True,
    )


# ── Tab 2 : Magasins ─────────────────────────────────────────────────────────
with tab2:
    st.plotly_chart(
        bar_h(by_site, "qty_missing", "site_label", "#FF9500",
              "Volume manquant (unités)", fmt_fn=lambda v: f"{int(v):,}"),
        use_container_width=True,
    )
    d = by_site.copy()
    for src, dst in [("fill_rate","Fill Rate"),("on_time","On Time"),("otif","OTIF"),("score","Score")]:
        d[dst] = d[src].apply(fmt_pct)
    d["Vol. manquant"]   = d["qty_missing"].apply(fmt)
    d["Impact CA proxy"] = d["impact_value"].apply(fmt)
    st.dataframe(
        d[["site_label", "Fill Rate", "On Time", "OTIF", "Score",
           "Vol. manquant", "Impact CA proxy", "suppliers", "articles", "Niveau"]]
        .rename(columns={"site_label": "Magasin", "suppliers": "Fournisseurs", "articles": "Articles"}),
        use_container_width=True, hide_index=True,
    )


# ── Tab 3 : Articles ─────────────────────────────────────────────────────────
with tab3:
    st.markdown("<div class='section-label'>Top 20 articles critiques</div>", unsafe_allow_html=True)
    st.plotly_chart(
        bar_h(by_article.head(20), "qty_missing", "article_label", "#007AFF",
              "Volume manquant (unités)", height=620, fmt_fn=lambda v: f"{int(v):,}"),
        use_container_width=True,
    )
    d = by_article.copy()
    for src, dst in [("fill_rate","Fill Rate"),("on_time","On Time"),("otif","OTIF"),("score","Score")]:
        d[dst] = d[src].apply(fmt_pct)
    d["Vol. manquant"]   = d["qty_missing"].apply(fmt)
    d["Impact CA proxy"] = d["impact_value"].apply(fmt)
    show_cols = [c for c in ["Code","article_label","supplier_name","Fill Rate","On Time",
                              "OTIF","Score","Vol. manquant","Impact CA proxy","sites","orders"]
                 if c in d.columns]
    st.dataframe(
        d[show_cols].rename(columns={"article_label":"Article","supplier_name":"Fournisseur",
                                      "sites":"Magasins","orders":"Cmdes"}),
        use_container_width=True, hide_index=True,
    )


# ── Tab 4 : Lignes critiques ──────────────────────────────────────────────────
with tab4:
    crit_lines = view[view["otif"] == 0].sort_values(
        ["qty_missing", "service_gap_value"], ascending=[False, False]
    )
    pct_non_otif = len(crit_lines) / len(view) * 100
    st.caption(f"{len(crit_lines):,} lignes non OTIF sur {len(view):,} ({pct_non_otif:.1f}%)")

    dcols = [c for c in [
        "date_received", "date_expected", "site_label", "supplier_name",
        "Code", "article_label", "N° Cde", "qte_cde", "qte_rec_retained",
        "qty_missing", "service_gap_value", "delay_days",
    ] if c in crit_lines.columns]
    dl = crit_lines[dcols].copy()
    if "service_gap_value" in dl.columns:
        dl["service_gap_value"] = dl["service_gap_value"].apply(fmt)
    st.dataframe(
        dl.rename(columns={
            "date_received":    "Date réception",
            "date_expected":    "Date prévue",
            "site_label":       "Magasin",
            "supplier_name":    "Fournisseur",
            "article_label":    "Article",
            "N° Cde":           "N° Commande",
            "qte_cde":          "Qté cde",
            "qte_rec_retained": "Qté reçue",
            "qty_missing":      "Qté manquante",
            "service_gap_value":"Impact CA proxy",
            "delay_days":       "Retard (j)",
        }),
        use_container_width=True, hide_index=True,
    )


# ── Tab 5 : Qualité des données ───────────────────────────────────────────────
with tab5:
    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Lignes brutes",       f"{quality['raw_rows']:,}")
    q2.metric("Lignes exploitables", f"{quality['clean_rows']:,}")
    q3.metric("Taux exploitable",    fmt_pct(quality['usable_rate']))
    q4.metric("Date prévue utilisée",quality["expected_col"])

    q5, q6, q7, q8 = st.columns(4)
    q5.metric("Qté cde ≤ 0 exclues",       f"{quality['excluded_zero_qty']:,}")
    q6.metric("Fournisseurs tech. exclus",  f"{quality['excluded_technical']:,}")
    q7.metric("Dates prévues manquantes",   f"{quality['missing_expected_date']:,}")
    q8.metric("Sur-réceptions capées",      f"{quality['sur_receipt_rows']:,}")

    if quality.get("all_dates_missing"):
        st.error("🔴 Aucune date prévue dans ce fichier — On Time = 100% artificiel.")
    else:
        st.warning("Règle temporaire : date prévue absente → ligne considérée On Time.")
    if quality.get("pv_zero_rows", 0) > 0:
        st.warning(f"⚠️ {quality['pv_zero_rows']:,} ligne(s) avec PV HT = 0 → impact CA proxy nul.")


# ── Tab 6 : Fiche Fournisseur ─────────────────────────────────────────────────
with tab6:

    if fiche_supplier == "— Choisir —":
        st.markdown("""
<div class='alert-card alert-blue'>
  <strong>ℹ️ Comment utiliser la Fiche Fournisseur ?</strong><br><br>
  1. Sélectionnez un fournisseur dans le menu <strong>📋 Fiche Fournisseur</strong> de la sidebar<br>
  2. Ajustez le <strong>seuil Fill Rate</strong> pour cibler les livraisons incomplètes à remonter<br>
  3. Consultez la synthèse KPI, le détail par magasin, par article, et ligne par ligne<br>
  4. Cliquez sur <strong>Générer la fiche Excel</strong> pour télécharger un fichier propre à envoyer au fournisseur
</div>""", unsafe_allow_html=True)

    else:
        # ── Données du fournisseur sélectionné (sur le périmètre filtré)
        fiche_df  = view[view["supplier_name"] == fiche_supplier].copy()
        fiche_bad = fiche_df[fiche_df["line_fill_rate"] < seuil_fill / 100].sort_values(
            ["qty_missing", "service_gap_value"], ascending=[False, False]
        )
        fkpi = compute_global_kpis(fiche_df)
        col_txt = score_color(fkpi["score"])

        # ── Header fiche
        fou_code = "—"
        if "Fou" in fiche_df.columns and not fiche_df["Fou"].isna().all():
            fou_code = str(int(fiche_df["Fou"].dropna().iloc[0]))

        st.markdown(f"""
<div class='fiche-header'>
  <div style='display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:12px'>
    <div>
      <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>{fiche_supplier}</div>
      <div style='font-size:12px;color:#8E8E93;margin-top:4px'>
        Code fournisseur : <strong>{fou_code}</strong> &nbsp;·&nbsp;
        {fkpi['orders']} commande(s) &nbsp;·&nbsp;
        {fkpi['articles']} article(s) &nbsp;·&nbsp;
        {fkpi['sites']} magasin(s)
      </div>
    </div>
    <div style='background:{col_txt}1A;border:1.5px solid {col_txt};border-radius:8px;
                padding:8px 16px;font-size:13px;font-weight:700;color:{col_txt};white-space:nowrap'>
      {score_band(fkpi["score"])} &nbsp;·&nbsp; Score {fkpi["score"]:.1f}%
    </div>
  </div>
</div>""", unsafe_allow_html=True)

        # ── KPIs fournisseur
        render_kpi_row(fkpi)
        st.markdown("---")

        # ── Récap par magasin
        st.markdown("<div class='section-label'>Livraisons incomplètes par magasin</div>", unsafe_allow_html=True)
        site_recap = (
            fiche_bad.groupby("site_label", as_index=False).agg(
                nb_cdes     =("N° Cde",           "nunique"),
                nb_articles =("Code",             "nunique"),
                qte_cde     =("qte_cde",          "sum"),
                qte_rec     =("qte_rec_retained",  "sum"),
                qty_missing =("qty_missing",       "sum"),
                impact_value=("service_gap_value", "sum"),
            )
            .sort_values("qty_missing", ascending=False)
        )
        site_recap["Fill Rate"]       = (site_recap["qte_rec"] / site_recap["qte_cde"] * 100).apply(fmt_pct)
        site_recap["Vol. manquant"]   = site_recap["qty_missing"].apply(fmt)
        site_recap["Impact CA proxy"] = site_recap["impact_value"].apply(fmt)
        st.dataframe(
            site_recap[["site_label","Fill Rate","Vol. manquant","Impact CA proxy","nb_cdes","nb_articles"]]
            .rename(columns={"site_label":"Magasin","nb_cdes":"Cmdes","nb_articles":"Articles"}),
            use_container_width=True, hide_index=True,
        )

        # ── Récap par article
        st.markdown("<div class='section-label' style='margin-top:18px'>Livraisons incomplètes par article</div>", unsafe_allow_html=True)
        art_recap = (
            fiche_bad.groupby(["Code","article_label"], as_index=False).agg(
                nb_sites    =("site_label",        "nunique"),
                nb_cdes     =("N° Cde",            "nunique"),
                qte_cde     =("qte_cde",           "sum"),
                qte_rec     =("qte_rec_retained",   "sum"),
                qty_missing =("qty_missing",        "sum"),
                impact_value=("service_gap_value",  "sum"),
            )
            .sort_values("qty_missing", ascending=False)
        )
        art_recap["Fill Rate"]       = (art_recap["qte_rec"] / art_recap["qte_cde"] * 100).apply(fmt_pct)
        art_recap["Vol. manquant"]   = art_recap["qty_missing"].apply(fmt)
        art_recap["Impact CA proxy"] = art_recap["impact_value"].apply(fmt)
        st.dataframe(
            art_recap[["Code","article_label","Fill Rate","Vol. manquant","Impact CA proxy","nb_sites","nb_cdes"]]
            .rename(columns={"article_label":"Article","nb_sites":"Magasins","nb_cdes":"Cmdes"}),
            use_container_width=True, hide_index=True,
        )

        # ── Détail ligne par ligne
        st.markdown("<div class='section-label' style='margin-top:18px'>Détail des livraisons incomplètes</div>", unsafe_allow_html=True)
        st.caption(
            f"{len(fiche_bad):,} ligne(s) avec Fill Rate < {seuil_fill}% "
            f"sur {len(fiche_df):,} total · triées par volume manquant décroissant"
        )
        dcols_f = [c for c in [
            "N° Cde", "date_received", "date_expected", "site_label",
            "Code", "article_label", "qte_cde", "qte_rec_retained",
            "qty_missing", "line_fill_rate", "service_gap_value", "on_time", "delay_days",
        ] if c in fiche_bad.columns]
        df_disp = fiche_bad[dcols_f].copy()
        if "line_fill_rate" in df_disp.columns:
            df_disp["line_fill_rate"] = df_disp["line_fill_rate"].apply(lambda v: fmt_pct(v * 100))
        if "service_gap_value" in df_disp.columns:
            df_disp["service_gap_value"] = df_disp["service_gap_value"].apply(fmt)
        if "on_time" in df_disp.columns:
            df_disp["on_time"] = df_disp["on_time"].map({True: "✅", False: "❌", 1: "✅", 0: "❌"}).fillna("—")
        st.dataframe(
            df_disp.rename(columns={
                "N° Cde":           "N° Commande",
                "date_received":    "Date réception",
                "date_expected":    "Date prévue",
                "site_label":       "Magasin",
                "article_label":    "Article",
                "qte_cde":          "Qté cde",
                "qte_rec_retained": "Qté reçue",
                "qty_missing":      "Qté manquante",
                "line_fill_rate":   "Fill Rate ligne",
                "service_gap_value":"Impact CA proxy",
                "on_time":          "À l'heure",
                "delay_days":       "Retard (j)",
            }),
            use_container_width=True, hide_index=True,
        )

        # ── Export fiche Excel
        st.markdown("---")
        if st.button("📥 Générer la fiche Excel fournisseur", type="primary", key="btn_fiche"):
            with st.spinner("Génération de la fiche…"):
                buf_fiche = build_fiche_excel(
                    fournisseur=fiche_supplier,
                    df_all=fiche_df,
                    df_bad=fiche_bad,
                    site_recap=site_recap,
                    art_recap=art_recap,
                    kpis=fkpi,
                    seuil=seuil_fill,
                )
            safe_name = fiche_supplier.strip().replace(" ", "_").replace("/", "-")[:40]
            st.download_button(
                label=f"⬇️ Télécharger la fiche — {fiche_supplier}",
                data=buf_fiche,
                file_name=f"SmartBuyer_OTIF_{safe_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_fiche",
            )


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL GLOBAL
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
with st.expander("📥 Export global — Synthèse · Fournisseur · Magasin · Article · Lignes critiques"):
    st.caption(f"{len(by_supplier)} fournisseurs · {len(by_site)} magasin(s) · {len(by_article)} articles")
    if st.button("Générer l'export global Excel", type="primary", key="btn_global"):
        with st.spinner("Génération en cours…"):
            buf_global = build_export_excel(view, by_supplier, by_site, by_article, quality)
        st.download_button(
            "⬇️ Télécharger l'export OTIF global",
            data=buf_global,
            file_name="SmartBuyer_OTIF_Global.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_global",
        )
