import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go

st.set_page_config(
    page_title="On Time In Full · SmartBuyer",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────────────────────
# CHARTE GRAPHIQUE SMARTBUYER
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display",
                 "SF Pro Text", "Helvetica Neue", Arial, sans-serif !important;
    background-color: #F2F2F7;
}
.stApp { background: #F2F2F7; }
.main .block-container { padding-top: 1.8rem; max-width: 1250px; }
[data-testid="stSidebar"] { background: #F2F2F7 !important; border-right: 0.5px solid #D1D1D6 !important; }
[data-testid="stMetric"] { background: #FFFFFF !important; border: 0.5px solid #E5E5EA !important; border-radius: 12px !important; padding: 16px 18px !important; }
[data-testid="stMetricLabel"] { font-size: 11px !important; font-weight: 500 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 600 !important; color: #1C1C1E !important; letter-spacing: -0.02em !important; }
[data-testid="stTabs"] button[role="tab"] { font-size: 13px !important; font-weight: 500 !important; padding: 8px 16px !important; color: #8E8E93 !important; border-radius: 0 !important; border-bottom: 2px solid transparent !important; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color: #007AFF !important; border-bottom: 2px solid #007AFF !important; background: transparent !important; }
[data-testid="stTabs"] [role="tablist"] { border-bottom: 0.5px solid #E5E5EA !important; }
[data-testid="stDataFrame"] { border: 0.5px solid #E5E5EA !important; border-radius: 10px !important; }
[data-testid="stDataFrame"] th { background: #F2F2F7 !important; font-size: 11px !important; font-weight: 600 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stFileUploader"] { border: 1.5px dashed #D1D1D6 !important; border-radius: 10px !important; background: #F9F9FB !important; }
.stDownloadButton > button { background: #007AFF !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; font-size: 13px !important; padding: 10px 24px !important; width: 100% !important; }
hr { border-color: #E5E5EA !important; margin: 1rem 0 !important; }
.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 1.5rem; }
.section-label { font-size: 11px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 10px; }
.alert-card  { padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; font-size: 13px; line-height: 1.5; border-left: 3px solid; }
.alert-red   { background: #FFF2F2; border-color: #FF3B30; color: #3A0000; }
.alert-amber { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }
.alert-green { background: #F0FFF4; border-color: #34C759; color: #003A10; }
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }
.kpi-focus { background: #EFF6FF; border: 1px solid #B3D9FF; border-radius: 12px; padding: 16px 18px; }
.kpi-focus-label { font-size: 11px; font-weight: 500; color: #007AFF; text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 3px; }
.kpi-focus-value { font-size: 24px; font-weight: 700; color: #007AFF; letter-spacing: -0.02em; }
.kpi-focus-sub   { font-size: 12px; color: #0066CC; margin-top: 3px; font-weight: 500; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# CONSTANTES MÉTIER
# ──────────────────────────────────────────────────────────────────────────────
TECHNICAL_SUPPLIERS = {
    "FOURNISSEUR STOCK",
    "FOURNISSEUR PLATEFORME LOCAL",
    "FOURNISSEUR PLATEFORME IMPORT",
}

# Colonnes date prévue à tester dans l'ordre de priorité
DATE_EXPECTED_CANDIDATES = ["H Date", "Date livraison", "Date prévue", "Date"]

# Seuils score OTIF
SEUIL_EXCELLENT = 97
SEUIL_SURVEILLER = 90

# ──────────────────────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────────────────────
def fmt(n):
    if pd.isna(n) or n is None:
        return "—"
    a = abs(float(n))
    if a >= 1_000_000:
        return f"{n/1_000_000:.1f} M"
    if a >= 1_000:
        return f"{int(n/1_000)} K"
    return f"{int(n):,}"


def fmt_pct(v, decimals=1):
    if pd.isna(v) or v is None:
        return "—"
    return f"{v:.{decimals}f}%"


def score_band(v):
    """Niveau de performance basé sur le score global (0–100)."""
    if pd.isna(v):
        return "Inconnu"
    if v >= SEUIL_EXCELLENT:
        return "🟢 Excellent"
    if v >= SEUIL_SURVEILLER:
        return "🟠 À surveiller"
    return "🔴 Critique"


def safe_div(a, b):
    return a / b if b not in (0, None) and not pd.isna(b) else 0


def detect_expected_date_column(df: pd.DataFrame):
    """
    Retourne la première colonne de date prévue disponible ET non entièrement vide.
    Retourne None si aucune colonne utilisable n'est trouvée.
    """
    for col in DATE_EXPECTED_CANDIDATES:
        if col in df.columns:
            parsed = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")
            non_null = parsed.notna().sum()
            if non_null > 0:
                return col
    return None  # Aucune date prévue disponible dans ce fichier


# ──────────────────────────────────────────────────────────────────────────────
# CHARGEMENT & NETTOYAGE
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_erp(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Charge le CSV ERP depuis les bytes bruts (pas l'objet file).
    On passe file_bytes + filename pour que le cache de Streamlit
    invalide correctement dès qu'un nouveau fichier est uploadé.
    """
    from io import BytesIO

    df = pd.read_csv(BytesIO(file_bytes), sep=";", low_memory=False)
    df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]

    # Supprimer les colonnes entièrement vides (ex: Unnamed: 34)
    df = df.dropna(axis=1, how="all")
    # Supprimer les colonnes "Unnamed: X"
    unnamed_cols = [c for c in df.columns if c.startswith("Unnamed:")]
    df = df.drop(columns=unnamed_cols, errors="ignore")

    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"nan": None, "": None})

    date_cols = [
        "Dt Rec", "Date de commande", "Date", "Date facture",
        "Date comptable du rapprochement", "H Date",
        "Date livraison", "Date prévue",
    ]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")

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
    Applique toutes les règles métier OTIF et retourne (df_clean, quality_dict).
    """
    if df.empty:
        return df.copy(), {}

    work = df.copy()
    raw_len = len(work)

    # ── Labels
    work["site_label"]     = work.get("Libellé site",        pd.Series(index=work.index, dtype="object")).fillna("Inconnu").astype(str).str.strip()
    work["supplier_name"]  = work.get("Nom fourn.",           pd.Series(index=work.index, dtype="object")).fillna("Inconnu").astype(str).str.strip()
    work["article_label"]  = work.get("Libellé article",      pd.Series(index=work.index, dtype="object")).fillna("Inconnu").astype(str).str.strip()
    work["dept_label"]     = work.get("Libellé département",  pd.Series(index=work.index, dtype="object")).fillna("Inconnu").astype(str).str.strip()
    work["famille_label"]  = work.get("Libellé famille",      pd.Series(index=work.index, dtype="object")).fillna("Inconnu").astype(str).str.strip()

    # ── Quantités
    work["qte_cde"] = pd.to_numeric(work.get("Qté cde"), errors="coerce").fillna(0)
    work["qte_rec"] = pd.to_numeric(work.get("Qté rec"), errors="coerce").fillna(0)

    # ── Prix de vente HT — utilisé pour l'impact CA proxy
    work["pv_ht"] = pd.to_numeric(work.get("Prix de vente HT"), errors="coerce").fillna(0)
    pv_zero_rows = int((work["pv_ht"] == 0).sum())  # lignes sans prix → impact sous-estimé

    # ── Flags exclusion
    work["is_technical_supplier"] = work["supplier_name"].str.upper().isin(TECHNICAL_SUPPLIERS)
    work["line_valid_qty"]        = work["qte_cde"] > 0

    excluded_technical = int(work["is_technical_supplier"].sum())
    excluded_zero_qty  = int((~work["line_valid_qty"]).sum())

    # ── Sur-réceptions avant cap
    sur_receipt_rows = int((work["qte_rec"] > work["qte_cde"]).sum())

    # ── Application des exclusions
    if exclude_technical:
        work = work[~work["is_technical_supplier"]].copy()
    work = work[work["line_valid_qty"]].copy()

    # ── Cap quantité reçue
    if cap_sur_receipt:
        work["qte_rec_retained"] = work[["qte_rec", "qte_cde"]].min(axis=1)
    else:
        work["qte_rec_retained"] = work["qte_rec"]

    # ── Détection colonne date prévue (sur données filtrées)
    expected_col = detect_expected_date_column(work)
    work["date_expected"] = work[expected_col] if expected_col else pd.NaT
    work["date_received"] = work.get("Dt Rec", pd.Series(index=work.index, dtype="datetime64[ns]"))

    missing_expected_date = int(work["date_expected"].isna().sum())
    all_dates_missing = (missing_expected_date == len(work))

    # ── Calculs OTIF
    work["qty_missing"]       = (work["qte_cde"] - work["qte_rec_retained"]).clip(lower=0)
    work["service_gap_value"] = work["qty_missing"] * work["pv_ht"]
    work["line_fill_rate"]    = np.where(work["qte_cde"] > 0, work["qte_rec_retained"] / work["qte_cde"], 0)
    work["delay_days"]        = (work["date_received"] - work["date_expected"]).dt.days

    # Règle temporaire : date prévue absente → On Time par défaut
    work["on_time"] = work["date_expected"].isna() | (work["date_received"] <= work["date_expected"])
    work["otif"]    = ((work["qte_rec_retained"] >= work["qte_cde"]) & work["on_time"]).astype(int)
    work["criticality_score"] = work["service_gap_value"] * (1 - work["line_fill_rate"])

    quality = {
        "raw_rows":              raw_len,
        "clean_rows":            len(work),
        "excluded_zero_qty":     excluded_zero_qty,
        "excluded_technical":    excluded_technical,
        "sur_receipt_rows":      sur_receipt_rows,
        "missing_expected_date": missing_expected_date,
        "all_dates_missing":     all_dates_missing,
        "pv_zero_rows":          pv_zero_rows,
        "expected_col":          expected_col if expected_col else "Aucune",
        "usable_rate":           round(safe_div(len(work), raw_len) * 100, 1) if raw_len else 0,
    }
    return work, quality


# ──────────────────────────────────────────────────────────────────────────────
# KPI GLOBAUX
# ──────────────────────────────────────────────────────────────────────────────
def compute_global_kpis(df: pd.DataFrame) -> dict:
    if df.empty:
        return {k: 0 for k in [
            "fill_rate", "on_time", "otif", "score",
            "ordered_qty", "received_qty", "missing_qty",
            "impact_value", "suppliers", "articles", "orders", "sites",
        ]}

    ordered  = df["qte_cde"].sum()
    received = df["qte_rec_retained"].sum()
    fill_rate = safe_div(received, ordered) * 100
    on_time   = df["on_time"].mean() * 100
    otif      = df["otif"].mean() * 100
    score     = 0.5 * fill_rate + 0.3 * on_time + 0.2 * otif

    return {
        "fill_rate":   fill_rate,
        "on_time":     on_time,
        "otif":        otif,
        "score":       score,
        "ordered_qty": ordered,
        "received_qty":received,
        "missing_qty": df["qty_missing"].sum(),
        "impact_value":df["service_gap_value"].sum(),
        "suppliers":   df["supplier_name"].nunique(),
        "articles":    df["Code"].nunique()  if "Code"   in df.columns else 0,
        "orders":      df["N° Cde"].nunique() if "N° Cde" in df.columns else 0,
        "sites":       df["site_label"].nunique(),
    }


# ──────────────────────────────────────────────────────────────────────────────
# AGRÉGATIONS
# ──────────────────────────────────────────────────────────────────────────────
def _add_perf_cols(g: pd.DataFrame) -> pd.DataFrame:
    """Calcule fill_rate, on_time%, otif%, score et score_band."""
    g["fill_rate"] = np.where(g["qte_cde"] > 0, g["qte_rec"] / g["qte_cde"] * 100, 0)
    g["on_time"]  *= 100
    g["otif"]     *= 100
    g["score"]     = 0.5 * g["fill_rate"] + 0.3 * g["on_time"] + 0.2 * g["otif"]
    g["Niveau"]    = g["score"].apply(score_band)
    return g


def agg_supplier(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["Fou", "supplier_name"], as_index=False).agg(
        qte_cde      =("qte_cde",           "sum"),
        qte_rec      =("qte_rec_retained",   "sum"),
        qty_missing  =("qty_missing",        "sum"),
        impact_value =("service_gap_value",  "sum"),
        on_time      =("on_time",            "mean"),
        otif         =("otif",               "mean"),
        orders       =("N° Cde",             "nunique"),
        articles     =("Code",               "nunique"),
        sites        =("site_label",         "nunique"),
    )
    g = _add_perf_cols(g)
    # Criticité = Impact CA proxy × taux de défaillance (1 - Fill Rate)
    # Logique : un fournisseur est critique s'il manque BEAUCOUP en valeur ET que son taux de service est faible
    # Exemple : impact 100 M FCFA à 40% Fill Rate → criticité = 100M × 60% = 60 M
    g["criticality_score"] = g["impact_value"] * (1 - g["fill_rate"] / 100)
    return g.sort_values(["criticality_score", "impact_value"], ascending=[False, False]).reset_index(drop=True)


def agg_site(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["Site", "site_label"], as_index=False).agg(
        qte_cde      =("qte_cde",           "sum"),
        qte_rec      =("qte_rec_retained",   "sum"),
        qty_missing  =("qty_missing",        "sum"),
        impact_value =("service_gap_value",  "sum"),
        on_time      =("on_time",            "mean"),
        otif         =("otif",               "mean"),
        suppliers    =("supplier_name",      "nunique"),
        articles     =("Code",              "nunique"),
    )
    g = _add_perf_cols(g)
    g["criticality_score"] = g["impact_value"] * (1 - g["fill_rate"] / 100)
    return g.sort_values(["criticality_score", "impact_value"], ascending=[False, False]).reset_index(drop=True)


def agg_article(df: pd.DataFrame) -> pd.DataFrame:
    group_cols = ["Code", "article_label", "supplier_name"]
    # On s'assure que les colonnes existent
    group_cols = [c for c in group_cols if c in df.columns]
    g = df.groupby(group_cols, as_index=False).agg(
        qte_cde      =("qte_cde",          "sum"),
        qte_rec      =("qte_rec_retained",  "sum"),
        qty_missing  =("qty_missing",       "sum"),
        impact_value =("service_gap_value", "sum"),
        on_time      =("on_time",           "mean"),
        otif         =("otif",              "mean"),
        sites        =("site_label",        "nunique"),
        orders       =("N° Cde",            "nunique"),
    )
    g = _add_perf_cols(g)
    g["criticality_score"] = g["impact_value"] * (1 - g["fill_rate"] / 100)
    return g.sort_values(["criticality_score", "impact_value"], ascending=[False, False]).reset_index(drop=True)


# ──────────────────────────────────────────────────────────────────────────────
# GRAPHIQUES PLOTLY
# ──────────────────────────────────────────────────────────────────────────────
PLOTLY_LAYOUT = dict(
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    font=dict(family="-apple-system, Helvetica Neue", color="#3A3A3C", size=11),
    margin=dict(t=10, b=10, l=10, r=70),
    xaxis=dict(showgrid=True, gridcolor="#F2F2F7"),
    yaxis=dict(showgrid=False, title=""),
)


def bar_h(data, x_col, y_col, color, x_title, height=500, fmt_fn=None):
    top = data.head(15).sort_values(x_col)
    texts = [fmt_fn(v) if fmt_fn else f"{v:.0f}" for v in top[x_col]]
    fig = go.Figure(go.Bar(
        x=top[x_col], y=top[y_col],
        orientation="h",
        marker_color=color, marker_line_width=0,
        text=texts, textposition="outside",
    ))
    layout = {**PLOTLY_LAYOUT, "height": height, "xaxis": {**PLOTLY_LAYOUT["xaxis"], "title": x_title}}
    fig.update_layout(**layout)
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# EXPORT EXCEL
# ──────────────────────────────────────────────────────────────────────────────
def build_export_excel(df, by_supplier, by_site, by_article, quality) -> BytesIO:
    wb = Workbook()
    hdr_fill = PatternFill("solid", fgColor="1C3557")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    ctr = Alignment(horizontal="center", vertical="center")

    def write_sheet(ws, title: str, dataframe: pd.DataFrame):
        ws.append([title])
        ws.cell(1, 1).font = Font(bold=True, size=13)
        ws.append([])
        headers = list(dataframe.columns)
        ws.append(headers)
        for i, _ in enumerate(headers, start=1):
            c = ws.cell(3, i)
            c.fill = hdr_fill
            c.font = hdr_font
            c.alignment = ctr
        for row in dataframe.itertuples(index=False):
            ws.append(list(row))
        # Ajustement largeur — robuste quelle que soit la longueur de colonne
        for col_cells in ws.iter_cols(min_row=3, max_row=3):
            col_letter = get_column_letter(col_cells[0].column)
            header_val = str(col_cells[0].value or "")
            ws.column_dimensions[col_letter].width = max(12, min(30, len(header_val) + 4))

    # Feuille synthèse qualité
    ws1 = wb.active
    ws1.title = "Synthese"
    synthese = pd.DataFrame([
        ["Lignes brutes",                       quality.get("raw_rows", 0)],
        ["Lignes exploitables",                 quality.get("clean_rows", 0)],
        ["Taux exploitable %",                  quality.get("usable_rate", 0)],
        ["Date prévue utilisée",                quality.get("expected_col", "Aucune")],
        ["Toutes dates prévues manquantes",      "OUI" if quality.get("all_dates_missing") else "NON"],
        ["Qté cde ≤ 0 exclues",                 quality.get("excluded_zero_qty", 0)],
        ["Fournisseurs techniques exclus",       quality.get("excluded_technical", 0)],
        ["Sur-réceptions détectées",             quality.get("sur_receipt_rows", 0)],
        ["Dates prévues manquantes",             quality.get("missing_expected_date", 0)],
        ["Lignes avec PV HT = 0 (impact nul)",  quality.get("pv_zero_rows", 0)],
    ], columns=["Indicateur", "Valeur"])
    write_sheet(ws1, "Synthèse qualité de données", synthese)

    # Feuilles analytiques
    write_sheet(wb.create_sheet("Par fournisseur"), "Analyse fournisseur", by_supplier)
    write_sheet(wb.create_sheet("Par magasin"),      "Analyse magasin",     by_site)
    write_sheet(wb.create_sheet("Par article"),      "Analyse article",     by_article)

    # Lignes critiques — top 500 non OTIF
    detail_cols = [c for c in [
        "date_received", "date_expected", "site_label", "supplier_name",
        "Code", "article_label", "N° Cde", "qte_cde", "qte_rec_retained",
        "qty_missing", "service_gap_value", "on_time", "otif", "delay_days",
    ] if c in df.columns]
    critical_df = (
        df[df["otif"] == 0][detail_cols]
        .sort_values(["qty_missing", "service_gap_value"], ascending=[False, False])
        .head(500)
    )
    write_sheet(wb.create_sheet("Lignes critiques"), "Détail lignes non OTIF", critical_df)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:18px'>
  <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>📦 SmartBuyer</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · On Time In Full</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Import fichier</div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Extraction ERP OTIF (CSV)", type=["csv"], key="otif")
    st.markdown("---")
    exclude_technical = st.checkbox("Exclure fournisseurs techniques", value=True)
    cap_sur_receipt   = st.checkbox("Caper Qté reçue à Qté commandée", value=True)


# ──────────────────────────────────────────────────────────────────────────────
# PAGE PRINCIPALE
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>📦 On Time In Full</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Pilotage OTIF fournisseur · Magasin · Article · Alertes opérationnelles · Impact volume</div>", unsafe_allow_html=True)

if uploaded is None:
    st.markdown("---")
    st.markdown("<div class='section-label'>Méthodologie de calcul</div>", unsafe_allow_html=True)

    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>ℹ️ Objectif du module</strong><br>
  Mesurer la performance de service fournisseur à partir d'un export ERP de réceptions.
  Quatre indicateurs complémentaires permettent d'identifier les fournisseurs critiques,
  les magasins impactés et les articles à risque.
</div>
""", unsafe_allow_html=True)

    # KPI 1 — Fill Rate
    st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:20px 22px;margin-bottom:12px'>
  <div style='font-size:13px;font-weight:700;color:#1C1C1E;margin-bottom:6px'>① Fill Rate — Taux de service quantitatif</div>
  <div style='font-size:13px;color:#3A3A3C;line-height:1.7'>
    Mesure la part de la quantité commandée effectivement reçue.<br>
    <code style='background:#F2F2F7;padding:2px 6px;border-radius:4px;font-size:12px'>
      Fill Rate = Qté reçue retenue / Qté commandée × 100
    </code><br><br>
    <strong>Règle cap :</strong> si la quantité reçue dépasse la quantité commandée (sur-réception),
    elle est ramenée à la quantité commandée pour éviter un taux supérieur à 100%.<br>
    <code style='background:#F2F2F7;padding:2px 6px;border-radius:4px;font-size:12px'>
      Qté reçue retenue = min(Qté reçue, Qté commandée)
    </code><br><br>
    <span style='color:#34C759;font-weight:600'>✓ Objectif cible : ≥ 97%</span>
  </div>
</div>
""", unsafe_allow_html=True)

    # KPI 2 — On Time
    st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:20px 22px;margin-bottom:12px'>
  <div style='font-size:13px;font-weight:700;color:#1C1C1E;margin-bottom:6px'>② On Time — Taux de respect du délai</div>
  <div style='font-size:13px;color:#3A3A3C;line-height:1.7'>
    Mesure la proportion de lignes livrées à la date prévue ou avant.<br>
    <code style='background:#F2F2F7;padding:2px 6px;border-radius:4px;font-size:12px'>
      On Time = 1 si Date réception ≤ Date prévue, sinon 0
    </code><br><br>
    <strong>Colonne date prévue utilisée (ordre de priorité) :</strong>
    <code>H Date</code> → <code>Date livraison</code> → <code>Date prévue</code> → <code>Date</code><br><br>
    <strong>⚠️ Règle temporaire ERP :</strong> si la date prévue est absente,
    la ligne est considérée <em>On Time</em> par défaut.
    Si aucune date prévue n'est disponible dans l'export, le On Time sera <strong>100% artificiel</strong>
    et l'OTIF ne reflétera que le Fill Rate.<br><br>
    <span style='color:#34C759;font-weight:600'>✓ Objectif cible : ≥ 95%</span>
  </div>
</div>
""", unsafe_allow_html=True)

    # KPI 3 — OTIF
    st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:20px 22px;margin-bottom:12px'>
  <div style='font-size:13px;font-weight:700;color:#1C1C1E;margin-bottom:6px'>③ OTIF — On Time In Full</div>
  <div style='font-size:13px;color:#3A3A3C;line-height:1.7'>
    Indicateur synthétique : une livraison est OTIF uniquement si elle est <strong>à la fois complète ET à l'heure</strong>.<br>
    <code style='background:#F2F2F7;padding:2px 6px;border-radius:4px;font-size:12px'>
      OTIF = 1 si (Qté reçue retenue ≥ Qté commandée) ET (On Time = 1), sinon 0
    </code><br><br>
    Le taux OTIF global est la moyenne des OTIF ligne par ligne.<br><br>
    <span style='color:#34C759;font-weight:600'>✓ Objectif cible : ≥ 95%</span>
  </div>
</div>
""", unsafe_allow_html=True)

    # KPI 4 — Score global
    st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:20px 22px;margin-bottom:12px'>
  <div style='font-size:13px;font-weight:700;color:#1C1C1E;margin-bottom:6px'>④ Score global — Synthèse pondérée</div>
  <div style='font-size:13px;color:#3A3A3C;line-height:1.7'>
    Agrège les trois indicateurs avec des pondérations reflétant leur priorité métier.<br>
    <code style='background:#F2F2F7;padding:2px 6px;border-radius:4px;font-size:12px'>
      Score = 50% × Fill Rate + 30% × On Time + 20% × OTIF
    </code><br><br>
    <strong>Niveaux de performance :</strong><br>
    🟢 <strong>Excellent</strong> → Score ≥ 97% &nbsp;|&nbsp;
    🟠 <strong>À surveiller</strong> → Score entre 90% et 97% &nbsp;|&nbsp;
    🔴 <strong>Critique</strong> → Score &lt; 90%<br><br>
    <strong>Exemple concret :</strong><br>
    Un fournisseur livre <strong>850 unités sur 1 000 commandées</strong>, avec <strong>7 lignes sur 10 à l'heure</strong>.<br>
    → Fill Rate = 850/1 000 = <strong>85%</strong><br>
    → On Time = 7/10 = <strong>70%</strong><br>
    → OTIF : parmi les 7 lignes à l'heure, seules celles avec réception complète comptent.
    Supposons 5 lignes complètes et à l'heure → OTIF = 5/10 = <strong>50%</strong><br>
    → Score = 50%×85 + 30%×70 + 20%×50 = 42.5 + 21 + 10 = <strong>73.5% 🔴 Critique</strong>
  </div>
</div>
""", unsafe_allow_html=True)

    # KPI 5 — Criticité
    st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:20px 22px;margin-bottom:12px'>
  <div style='font-size:13px;font-weight:700;color:#1C1C1E;margin-bottom:6px'>⑤ Score de criticité — Priorisation opérationnelle</div>
  <div style='font-size:13px;color:#3A3A3C;line-height:1.7'>
    Permet de classer les fournisseurs par ordre de priorité d'action en combinant
    <strong>l'impact financier du manque</strong> et <strong>le taux de défaillance</strong>.<br><br>
    <code style='background:#F2F2F7;padding:2px 6px;border-radius:4px;font-size:12px'>
      Criticité = Impact CA proxy × (1 − Fill Rate)
    </code><br>
    <code style='background:#F2F2F7;padding:2px 6px;border-radius:4px;font-size:12px'>
      Impact CA proxy = Qté manquante × Prix de vente HT
    </code><br><br>
    <strong>Pourquoi cette formule ?</strong><br>
    Un fournisseur avec un fort volume manquant mais sur des articles peu chers sera moins
    prioritaire qu'un fournisseur avec moins de manque mais sur des références à forte valeur.<br><br>
    <strong>Exemple concret :</strong><br>
    <em>Fournisseur A</em> — Impact CA proxy = 50 M FCFA, Fill Rate = 80% → Criticité = 50M × 20% = <strong>10 M</strong><br>
    <em>Fournisseur B</em> — Impact CA proxy = 30 M FCFA, Fill Rate = 20% → Criticité = 30M × 80% = <strong>24 M</strong><br>
    → Le fournisseur B est prioritaire malgré un impact brut inférieur, car il est <strong>beaucoup moins fiable</strong>.
  </div>
</div>
""", unsafe_allow_html=True)

    # Impact CA proxy
    st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:20px 22px;margin-bottom:12px'>
  <div style='font-size:13px;font-weight:700;color:#1C1C1E;margin-bottom:6px'>⑥ Impact CA proxy — Valorisation du manque</div>
  <div style='font-size:13px;color:#3A3A3C;line-height:1.7'>
    Traduit le volume non livré en valeur commerciale potentiellement perdue.<br>
    <code style='background:#F2F2F7;padding:2px 6px;border-radius:4px;font-size:12px'>
      Impact CA proxy = Qté manquante × Prix de vente HT
    </code><br><br>
    Il s'agit d'une <strong>estimation haute</strong> (on suppose que les unités manquantes auraient
    toutes été vendues au prix catalogue). Il ne tient pas compte des stocks disponibles en magasin.<br><br>
    <span style='color:#FF9500;font-weight:600'>⚠️ Si Prix de vente HT = 0 pour certaines lignes, l'impact est sous-estimé — une alerte s'affiche.</span>
  </div>
</div>
""", unsafe_allow_html=True)

    st.markdown("""
<div class='alert-card alert-amber'>
  <strong>⚠️ Règle temporaire ERP active</strong><br>
  Si la date prévue est absente ou nulle dans l'export, la ligne est considérée <strong>On Time</strong> par défaut.
  Le On Time sera artificiellement à 100% si aucune date prévue n'est renseignée.
  Demandez à l'IT un export incluant la colonne <code>H Date</code> ou <code>Date livraison</code>.
</div>
""", unsafe_allow_html=True)

    st.stop()

# ── Chargement (cache sur bytes + filename pour invalidation correcte)
with st.spinner("Lecture et préparation des données…"):
    file_bytes = uploaded.read()
    raw = load_erp(file_bytes, uploaded.name)
    df, quality = prepare_dataset(raw, exclude_technical=exclude_technical, cap_sur_receipt=cap_sur_receipt)

if df.empty:
    st.error("Aucune ligne exploitable après nettoyage. Vérifiez le format du fichier.")
    st.stop()

# ── Filtres sidebar (après chargement)
with st.sidebar:
    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Filtres</div>", unsafe_allow_html=True)

    sites       = sorted([x for x in df["site_label"].dropna().unique() if x and x != "Inconnu"])
    suppliers   = sorted([x for x in df["supplier_name"].dropna().unique() if x and x != "Inconnu"])
    departments = sorted([x for x in df["dept_label"].dropna().unique() if x and x != "Inconnu"])

    sel_sites      = st.multiselect("Magasin",      sites,       default=sites)
    sel_suppliers  = st.multiselect("Fournisseur",  suppliers,   default=[])
    sel_depts      = st.multiselect("Département",  departments, default=departments)
    only_critical  = st.checkbox("Afficher uniquement OTIF = 0", value=False)

# ── Garde-fou : filtres vides → on garde tout
if not sel_sites:
    sel_sites = sites
if not sel_depts:
    sel_depts = departments

view = df[df["site_label"].isin(sel_sites) & df["dept_label"].isin(sel_depts)].copy()
if sel_suppliers:
    view = view[view["supplier_name"].isin(sel_suppliers)].copy()
if only_critical:
    view = view[view["otif"] == 0].copy()

if view.empty:
    st.warning("Aucune donnée après filtrage — ajustez les filtres dans la sidebar.")
    st.stop()

# ── Calculs
kpi         = compute_global_kpis(view)
by_supplier = agg_supplier(view)
by_site     = agg_site(view)
by_article  = agg_article(view)

# ──────────────────────────────────────────────────────────────────────────────
# KPI CARDS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown(f"<div class='section-label'>{kpi['sites']} magasin(s) · {kpi['suppliers']} fournisseur(s) · {kpi['orders']} commande(s)</div>", unsafe_allow_html=True)

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
</div>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# ALERTES
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Alertes &amp; points d'attention</div>", unsafe_allow_html=True)

# Alerte OTIF biaisé si toutes les dates manquent
if quality.get("all_dates_missing"):
    st.markdown("""
<div class='alert-card alert-red'>
  <strong>🔴 On Time non significatif</strong><br>
  Aucune date prévue disponible dans ce fichier. Le On Time est à <strong>100%</strong> par application de la règle temporaire ERP.
  L'OTIF ne reflète que le Fill Rate dans ce cas. Demandez un export avec la colonne <code>H Date</code> ou <code>Date livraison</code>.
</div>
""", unsafe_allow_html=True)
elif quality.get("missing_expected_date", 0) > 0:
    pct_missing = round(quality["missing_expected_date"] / quality["clean_rows"] * 100, 1)
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ Dates prévues partiellement manquantes</strong><br>
  {quality['missing_expected_date']:,} lignes sans date prévue ({pct_missing}%) → considérées On Time par défaut.
</div>
""", unsafe_allow_html=True)

crit_sup  = by_supplier.head(3)
crit_site = by_site.head(3)

if not crit_sup.empty:
    st.markdown(f"""
<div class='alert-card alert-red'>
  <strong>🔴 Fournisseurs les plus critiques</strong><br>
  {" · ".join([f"{r.supplier_name} (score&nbsp;{r.score:.1f}%)" for r in crit_sup.itertuples()])}
</div>
""", unsafe_allow_html=True)

if not crit_site.empty:
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ Magasins les plus impactés</strong><br>
  {" · ".join([f"{r.site_label} ({int(r.qty_missing):,}&nbsp;unités manquantes)" for r in crit_site.itertuples()])}
</div>
""", unsafe_allow_html=True)

if quality.get("pv_zero_rows", 0) > 0:
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ Impact CA proxy sous-estimé</strong><br>
  {quality['pv_zero_rows']:,} lignes avec Prix de vente HT = 0 → impact CA proxy = 0 pour ces lignes.
</div>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class='alert-card alert-blue'>
  <strong>ℹ️ Qualité de données</strong><br>
  Colonne date prévue utilisée : <strong>{quality['expected_col']}</strong> ·
  Dates prévues manquantes : <strong>{quality['missing_expected_date']:,}</strong> ·
  Sur-réceptions capées : <strong>{quality['sur_receipt_rows']:,}</strong> ·
  Taux exploitable : <strong>{quality['usable_rate']:.1f}%</strong>
</div>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# TABS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("---")
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    f"🚛 Fournisseurs ({len(by_supplier)})",
    f"🏪 Magasins ({len(by_site)})",
    f"📦 Articles ({len(by_article)})",
    "🚨 Lignes critiques",
    "🧪 Qualité des données",
])

# ── Tab 1 : Fournisseurs
with tab1:
    st.caption("Trié par criticité décroissante = volume manquant × dégradation du score")
    st.plotly_chart(
        bar_h(by_supplier, "criticality_score", "supplier_name", "#FF3B30", "Score de criticité"),
        use_container_width=True,
    )
    disp = by_supplier.copy()
    disp["Fill Rate"]       = disp["fill_rate"].apply(fmt_pct)
    disp["On Time"]         = disp["on_time"].apply(fmt_pct)
    disp["OTIF"]            = disp["otif"].apply(fmt_pct)
    disp["Score"]           = disp["score"].apply(fmt_pct)
    disp["Vol. manquant"]   = disp["qty_missing"].apply(fmt)
    disp["Impact CA proxy"] = disp["impact_value"].apply(fmt)
    st.dataframe(
        disp[["supplier_name", "Fill Rate", "On Time", "OTIF", "Score",
              "Vol. manquant", "Impact CA proxy", "orders", "articles", "sites", "Niveau"]]
        .rename(columns={"supplier_name": "Fournisseur", "orders": "Cmdes",
                         "articles": "Articles", "sites": "Magasins"}),
        use_container_width=True, hide_index=True,
    )

# ── Tab 2 : Magasins
with tab2:
    st.plotly_chart(
        bar_h(by_site, "qty_missing", "site_label", "#FF9500", "Volume manquant",
              fmt_fn=lambda v: f"{int(v):,}"),
        use_container_width=True,
    )
    disp = by_site.copy()
    for src, dst in [("fill_rate","Fill Rate"),("on_time","On Time"),("otif","OTIF"),("score","Score")]:
        disp[dst] = disp[src].apply(fmt_pct)
    disp["Vol. manquant"]   = disp["qty_missing"].apply(fmt)
    disp["Impact CA proxy"] = disp["impact_value"].apply(fmt)
    st.dataframe(
        disp[["site_label", "Fill Rate", "On Time", "OTIF", "Score",
              "Vol. manquant", "Impact CA proxy", "suppliers", "articles", "Niveau"]]
        .rename(columns={"site_label": "Magasin", "suppliers": "Fournisseurs", "articles": "Articles"}),
        use_container_width=True, hide_index=True,
    )

# ── Tab 3 : Articles
with tab3:
    st.markdown("<div class='section-label'>Top 20 articles critiques</div>", unsafe_allow_html=True)
    art_top = by_article.head(20)
    st.plotly_chart(
        bar_h(art_top, "qty_missing", "article_label", "#007AFF", "Volume manquant", height=620,
              fmt_fn=lambda v: f"{int(v):,}"),
        use_container_width=True,
    )
    disp = by_article.copy()
    for src, dst in [("fill_rate","Fill Rate"),("on_time","On Time"),("otif","OTIF"),("score","Score")]:
        disp[dst] = disp[src].apply(fmt_pct)
    disp["Vol. manquant"]   = disp["qty_missing"].apply(fmt)
    disp["Impact CA proxy"] = disp["impact_value"].apply(fmt)
    cols = [c for c in ["Code","article_label","supplier_name","Fill Rate","On Time","OTIF","Score",
                         "Vol. manquant","Impact CA proxy","sites","orders"] if c in disp.columns]
    st.dataframe(
        disp[cols].rename(columns={"article_label":"Article","supplier_name":"Fournisseur",
                                    "sites":"Magasins","orders":"Cmdes"}),
        use_container_width=True, hide_index=True,
    )

# ── Tab 4 : Lignes critiques
with tab4:
    critical_lines = view[view["otif"] == 0].sort_values(
        ["qty_missing", "service_gap_value"], ascending=[False, False]
    )
    st.caption(f"{len(critical_lines):,} lignes non OTIF sur {len(view):,} total ({len(critical_lines)/len(view)*100:.1f}%)")
    detail_cols = [c for c in [
        "date_received", "date_expected", "site_label", "supplier_name",
        "Code", "article_label", "N° Cde", "qte_cde", "qte_rec_retained",
        "qty_missing", "service_gap_value", "delay_days",
    ] if c in critical_lines.columns]
    disp_lines = critical_lines[detail_cols].copy()
    if "service_gap_value" in disp_lines.columns:
        disp_lines["service_gap_value"] = disp_lines["service_gap_value"].apply(fmt)
    st.dataframe(
        disp_lines.rename(columns={
            "date_received":    "Date réception",
            "date_expected":    "Date prévue",
            "site_label":       "Magasin",
            "supplier_name":    "Fournisseur",
            "article_label":    "Article",
            "N° Cde":           "Commande",
            "qte_cde":          "Qté cde",
            "qte_rec_retained": "Qté reçue retenue",
            "qty_missing":      "Qté manquante",
            "service_gap_value":"Impact CA proxy",
            "delay_days":       "Retard (j)",
        }),
        use_container_width=True, hide_index=True,
    )

# ── Tab 5 : Qualité des données
with tab5:
    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Lignes brutes",       f"{quality['raw_rows']:,}")
    q2.metric("Lignes exploitables", f"{quality['clean_rows']:,}")
    q3.metric("Taux exploitable",    fmt_pct(quality['usable_rate']))
    q4.metric("Date prévue utilisée",quality["expected_col"])

    q5, q6, q7, q8 = st.columns(4)
    q5.metric("Qté cde ≤ 0 exclues",      f"{quality['excluded_zero_qty']:,}")
    q6.metric("Fournisseurs tech. exclus", f"{quality['excluded_technical']:,}")
    q7.metric("Dates prévues manquantes",  f"{quality['missing_expected_date']:,}")
    q8.metric("Sur-réceptions capées",     f"{quality['sur_receipt_rows']:,}")

    if quality.get("all_dates_missing"):
        st.error("🔴 Aucune date prévue dans ce fichier. On Time = 100% artificiel — KPI non représentatif.")
    else:
        st.warning("Règle temporaire active : si la date prévue est absente, la ligne est considérée On Time.")

    if quality.get("pv_zero_rows", 0) > 0:
        st.warning(f"⚠️ {quality['pv_zero_rows']:,} ligne(s) avec PV HT = 0 → impact CA proxy nul pour ces lignes.")

# ──────────────────────────────────────────────────────────────────────────────
# EXPORT EXCEL
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("---")
with st.expander("📥 Export Excel — Synthèse · Fournisseur · Magasin · Article · Lignes critiques"):
    st.caption(f"{len(by_supplier)} fournisseurs · {len(by_site)} magasin(s) · {len(by_article)} articles")
    if st.button("Générer le fichier Excel", type="primary"):
        with st.spinner("Génération en cours…"):
            buf = build_export_excel(view, by_supplier, by_site, by_article, quality)
        st.download_button(
            "⬇️ Télécharger l'export OTIF",
            data=buf,
            file_name="SmartBuyer_OTIF.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
