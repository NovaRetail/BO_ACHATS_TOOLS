"""
05_🏪_Suivi_Implantation.py — SmartBuyer Hub
Suivi Implantation Nouvelles Références — v7.1
Charte visuelle SmartBuyer v2.3 (alignée app.py)
Source stock : fichiers CSV ERP par magasin (1 à 13 fichiers)
"""

import io
import re
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from datetime import date

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────────────────────
TODAY      = pd.Timestamp(date.today())
TODAY_STR  = date.today().strftime("%d %b %Y")
TODAY_FILE = date.today().strftime("%Y%m%d")

STOCK_COLS_REQUIRED = [
    "Site", "Libellé site", "Code article", "Libellé article",
    "Code etat", "Nouveau stock", "Ral", "Pcb",
    "Code marketing", "Nom fourn.", "Libellé rayon", "Libellé famille",
    "Date dernière entrée", "Type saisonnalité",
]

ETAT_ACTIF    = "2"
ETAT_PURGE    = "P"
ETAT_ANOMALIE = {"B", "S", "F", "6", "5", "1"}
ETAT_LABEL    = {
    "B": "Rayon générique", "S": "Suspendu",
    "F": "Fin de vie",      "6": "Déréférencé", "5": "Autre",
}

ALERTES = {
    "✅ Implanté":        "#34C759",
    "🔵 Appro en cours":  "#007AFF",
    "🛒 Passer commande": "#FF9500",
    "🚩 Anomalie référ.": "#FF9500",
}
ACTION_LABEL = {
    "✅ Implanté":        "—",
    "🔵 Appro en cours":  "Accélérer livraison",
    "🛒 Passer commande": "Passer commande fournisseur",
    "🚩 Anomalie référ.": "Vérifier référencement magasin",
}

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Implantation · SmartBuyer",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CHARTE SMARTBUYER v2.3 — alignée app.py
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display",
                 "SF Pro Text", "Helvetica Neue", Arial, sans-serif !important;
    background-color: #F2F2F7;
}
.stApp { background: #F2F2F7; }
.main .block-container { padding-top: 1.8rem; max-width: 1300px; }

[data-testid="stSidebar"] { background: #FFFFFF !important; border-right: 0.5px solid #E5E5EA !important; }

[data-testid="stMetric"] {
    background: #FFFFFF !important; border: 0.5px solid #E5E5EA !important;
    border-radius: 12px !important; padding: 16px 18px !important;
}
[data-testid="stMetricLabel"] { font-size: 11px !important; font-weight: 500 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 600 !important; color: #1C1C1E !important; letter-spacing: -0.02em !important; }

[data-testid="stTabs"] button[role="tab"] { font-size: 13px !important; font-weight: 500 !important; padding: 8px 16px !important; color: #8E8E93 !important; border-bottom: 2px solid transparent !important; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color: #007AFF !important; border-bottom: 2px solid #007AFF !important; background: transparent !important; }
[data-testid="stTabs"] [role="tablist"] { border-bottom: 0.5px solid #E5E5EA !important; }

[data-testid="stDataFrame"] { border: 0.5px solid #E5E5EA !important; border-radius: 10px !important; }
[data-testid="stDataFrame"] th { background: #F2F2F7 !important; font-size: 11px !important; font-weight: 600 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }

[data-testid="stFileUploader"] { border: 1.5px dashed #D1D1D6 !important; border-radius: 10px !important; background: #F9F9FB !important; }
.stDownloadButton > button { background: #007AFF !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; font-size: 13px !important; padding: 10px 24px !important; width: 100% !important; }
hr { border-color: #E5E5EA !important; margin: 1rem 0 !important; }

/* Titres — style app.py */
.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 1.5rem; }
.section-label { font-size: 11px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 10px; }

/* KPI barre haute — style app.py */
.kpi-bar { background: #FFFFFF; border-radius: 14px; border: 0.5px solid #E5E5EA; padding: 0.85rem 1.25rem; text-align: center; }
.kpi-bar-val   { font-size: 20px; font-weight: 700; color: #1C1C1E; }
.kpi-bar-label { font-size: 11px; color: #8E8E93; margin-top: 2px; }

/* Alertes — style Marges Négatives */
.alert-card  { padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; font-size: 13px; line-height: 1.5; border-left: 3px solid; }
.alert-red   { background: #FFF2F2; border-color: #FF3B30; color: #3A0000; }
.alert-amber { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }
.alert-green { background: #F0FFF4; border-color: #34C759; color: #003A10; }
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }

/* Scorecard magasins */
.scorecard-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 10px; margin-bottom: 16px; }
.scorecard-card { background: #FFFFFF; border: 0.5px solid #E5E5EA; border-radius: 12px; padding: 14px 16px; position: relative; }
.scorecard-card.ok   { border-color: #6EE7B7; background: #F0FFF4; }
.scorecard-card.warn { border-color: #FCD34D; background: #FFFBF0; }
.scorecard-card.ko   { border-color: #FECACA; background: #FFF2F2; }
.scorecard-dot  { width: 8px; height: 8px; border-radius: 50%; position: absolute; top: 14px; right: 14px; }
.scorecard-name { font-size: 11px; font-weight: 600; color: #1C1C1E; margin-bottom: 6px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 88%; }
.scorecard-pct  { font-size: 28px; font-weight: 700; line-height: 1; }
.scorecard-sub  { font-size: 10px; color: #8E8E93; margin-top: 3px; }

/* Bannière actions */
.action-banner {
    background: #FFFFFF; border: 0.5px solid #E5E5EA; border-left: 3px solid #FF3B30;
    border-radius: 10px; padding: 12px 16px; margin-bottom: 12px;
    display: flex; align-items: center; gap: 16px; flex-wrap: wrap;
}
.action-item { display: flex; flex-direction: column; align-items: center; padding: 0 12px; border-right: 0.5px solid #E5E5EA; }
.action-item:last-child { border-right: none; }
.action-num  { font-size: 22px; font-weight: 700; line-height: 1; }
.action-lbl  { font-size: 10px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: .05em; margin-top: 2px; }

/* Val-box résumé données */
.val-box { background: #FFFFFF; border: 0.5px solid #E5E5EA; border-radius: 10px; padding: 12px 18px; margin-bottom: 14px; display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
.val-item { display: flex; flex-direction: column; align-items: center; padding: 0 12px; border-right: 0.5px solid #E5E5EA; }
.val-item:last-child { border-right: none; }
.val-num { font-size: 18px; font-weight: 700; line-height: 1; }
.val-lbl { font-size: 10px; color: #8E8E93; text-transform: uppercase; letter-spacing: .05em; margin-top: 2px; }

/* Info box */
.info-box { border-radius: 10px; padding: 12px 16px; margin-bottom: 10px; border: 0.5px solid; font-size: 13px; line-height: 1.6; border-left: 3px solid; }
.info-box.blue   { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }
.info-box.green  { background: #F0FFF4; border-color: #34C759; color: #003A10; }
.info-box.amber  { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }

/* Section header (tiret) */
.sh { font-size: 11px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.07em; margin: 20px 0 10px; padding-bottom: 6px; border-bottom: 0.5px solid #E5E5EA; }

/* Col requise — style Marges Négatives */
.col-required { background: #F0F8FF; border: 0.5px solid #B3D9FF; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; display: flex; align-items: flex-start; gap: 10px; }
.col-name { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def fmt_n(n) -> str:
    try:
        return f"{int(n):,}".replace(",", "\u202f")
    except Exception:
        return str(n)

def color_taux(t: float) -> str:
    if t >= 80: return "#34C759"
    if t >= 50: return "#FF9500"
    return "#FF3B30"

def scorecard_cls(t: float) -> str:
    if t >= 80: return "ok"
    if t >= 50: return "warn"
    return "ko"

def extract_site_from_filename(filename: str) -> str | None:
    m = re.search(r'_(\d{5})_', filename)
    return m.group(1) if m else None

def extract_date_from_filename(filename: str) -> pd.Timestamp:
    m8 = re.search(r'_(\d{8})_', filename)
    if m8:
        for fmt in ('%Y%m%d', '%d%m%Y'):
            try:
                return pd.Timestamp(pd.to_datetime(m8.group(1), format=fmt))
            except Exception:
                continue
    return TODAY

def _sem_sort(s) -> int:
    cleaned = re.sub(r"[Ss]", "", str(s).strip())
    return int(cleaned) if cleaned.isdigit() else 99

def get_alerte(row) -> str:
    etat  = str(row.get("Code etat", "")).strip()
    stock = row.get("Nouveau stock")
    ral   = float(row.get("Ral", 0) or 0)
    if not etat or etat in ("nan", ETAT_PURGE) or etat in ETAT_ANOMALIE:
        return "🚩 Anomalie référ."
    if pd.isna(stock):
        return "🚩 Anomalie référ."
    if float(stock) != 0:
        return "✅ Implanté"
    return "🔵 Appro en cours" if ral > 0 else "🛒 Passer commande"


# ─────────────────────────────────────────────────────────────────────────────
# PARSERS
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_t1(file_bytes: bytes, filename: str):
    buf = io.BytesIO(file_bytes)
    try:
        df = pd.read_excel(buf, header=None, dtype=str) \
             if filename.lower().endswith((".xlsx", ".xls")) \
             else pd.read_csv(buf, header=None, sep=None, engine="python",
                              encoding="latin1", dtype=str, on_bad_lines="skip")
    except Exception as e:
        return None, f"Lecture T1 : {e}"

    first = str(df.iloc[0, 0]).strip().replace(".0", "")
    has_header = not first.isdigit()
    if has_header:
        df.columns = df.iloc[0].astype(str).str.strip().str.upper()
        df = df.iloc[1:].reset_index(drop=True)
    else:
        df.columns = ["ARTICLE"] + [f"_COL{i}" for i in range(1, len(df.columns))]

    df.columns = (df.columns.astype(str).str.strip().str.upper()
                  .str.replace("\ufeff", "", regex=False)
                  .str.replace("\xa0", " ", regex=False))

    if "ARTICLE" not in df.columns:
        return None, "Colonne ARTICLE introuvable dans le fichier T1"

    df["SKU"] = (df["ARTICLE"].astype(str).str.strip()
                 .str.replace(r"\.0$", "", regex=True).str.zfill(8).str[:8])
    df = df[df["SKU"].str.match(r"^\d{8}$", na=False)].drop_duplicates("SKU").copy()

    for col, val in {"LIBELLÉ ARTICLE": "", "LIBELLÉ FOURNISSEUR ORIGINE": "",
                     "MODE APPRO": "", "SEMAINE RECEPTION": "", "DATE LIV.": ""}.items():
        if col not in df.columns:
            df[col] = val

    df["SEMAINE RECEPTION"] = df["SEMAINE RECEPTION"].astype(str).str.strip().replace("nan", "")
    df["SEM_NUM"] = df["SEMAINE RECEPTION"].apply(
        lambda s: int(re.sub(r"[Ss]", "", s)) if re.sub(r"[Ss]", "", s).isdigit() else 99
    )
    df["ORIGINE"] = df["MODE APPRO"].apply(
        lambda m: "IM" if "IMPORT" in str(m).upper() else "LO"
    )
    return df, None


@st.cache_data(show_spinner=False)
def parse_stock_file(file_bytes: bytes, filename: str, sku_scope: tuple):
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), sep=";", encoding="latin1",
                         low_memory=False, on_bad_lines="skip", dtype=str)
    except Exception as e:
        return None, f"Lecture {filename} : {e}"

    missing = [c for c in STOCK_COLS_REQUIRED if c not in df.columns]
    if missing:
        return None, f"{filename} — colonnes manquantes : {', '.join(missing)}"

    for col in ["Libellé site", "Libellé article", "Nom fourn.", "Libellé rayon",
                "Code etat", "Code marketing", "Type saisonnalité"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    df["SKU"]       = (df["Code article"].astype(str).str.strip()
                       .str.replace(r"\.0$", "", regex=True).str.zfill(8).str[:8])
    df["Code site"] = df["Site"].astype(str).str.strip()

    site_from_name = extract_site_from_filename(filename)
    if site_from_name and df["Code site"].isin(["nan", ""]).all():
        df["Code site"] = site_from_name

    df = df[df["Code etat"] != ETAT_PURGE].copy()
    if sku_scope:
        df = df[df["SKU"].isin(sku_scope)].copy()
    df = df[df["SKU"].str.match(r"^\d{8}$", na=False)].copy()

    df["Nouveau stock"]  = pd.to_numeric(df["Nouveau stock"], errors="coerce")
    df["Ral"]            = pd.to_numeric(df["Ral"], errors="coerce").fillna(0).astype(int)
    df["Pcb"]            = pd.to_numeric(df["Pcb"], errors="coerce").fillna(1)
    df["Origine"]        = df["Code marketing"].apply(lambda m: "IM" if str(m).strip().upper() == "IM" else "LO")
    df["_date_fichier"]  = extract_date_from_filename(filename)

    return df, None


def consolidate_stock(uploaded_files, sku_scope: tuple):
    frames, errors = [], []
    for f in uploaded_files:
        df_site, err = parse_stock_file(f.read(), f.name, sku_scope)
        if err:
            errors.append(err)
        elif df_site is not None and not df_site.empty:
            frames.append(df_site)
    return (pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()), errors


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:18px'>
  <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>🛍️ SmartBuyer</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · Équipe Achats</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Import fichiers</div>", unsafe_allow_html=True)
    st.markdown("**① T1 — Nouvelles Références**")
    t1_file = st.file_uploader("T1", type=["csv", "xlsx", "xls"],
                               key="t1", label_visibility="collapsed")
    st.markdown("**② Stock par magasin** *(multi-upload)*")
    st.caption("1 à 13 fichiers CSV · nommage : *_10206_20260507_*.csv")
    stk_files = st.file_uploader("Stock", type=["csv"], accept_multiple_files=True,
                                  key="stk", label_visibility="collapsed")


# ─────────────────────────────────────────────────────────────────────────────
# PAGE PRINCIPALE
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>🏪 Suivi Implantation</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Nouvelles références T1 · Stock ERP par magasin · Alertes · Cessions inter-magasins</div>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# ÉCRAN D'ACCUEIL
# ─────────────────────────────────────────────────────────────────────────────
if not t1_file:
    st.markdown("---")

    st.markdown("""
<div class='info-box blue'>
  <strong>ℹ️ À quoi sert ce module ?</strong><br>
  Suivi en temps réel de l'implantation des nouvelles références T1 dans le réseau.
  À partir de la liste T1 et des extractions ERP par magasin, le module calcule le
  <strong>taux d'implantation</strong> site par site et propose des actions concrètes :
  accélérer une livraison, passer une commande, planifier une cession.
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("<div class='section-label'>Les 4 statuts d'alerte</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    alertes_doc = [
        ("✅", "Implanté",        "#34C759", "Stock ≠ 0 (positif ou négatif)",
         "stock != 0", "Article présent en rayon — aucune action requise."),
        ("🔵", "Appro en cours",  "#007AFF", "Stock = 0 · RAL > 0",
         "stock == 0 & RAL > 0", "Livraison en cours — accélérer si délai dépassé."),
        ("🛒", "Passer commande", "#FF9500", "Stock = 0 · RAL = 0",
         "stock == 0 & RAL == 0", "Aucune commande en cours — passer commande fournisseur."),
        ("🚩", "Anomalie référ.", "#FF9500", "Code état ≠ 2 ou absent du stock",
         "Code etat not in ['2']", "Vérifier le référencement magasin ou la liste T1."),
    ]
    for i, (ico, titre, color, cond, formule, action) in enumerate(alertes_doc):
        with (c1 if i % 2 == 0 else c2):
            st.markdown(f"""
<div class='module-card' style='background:#FFFFFF;border:0.5px solid #E5E5EA;
     border-radius:12px;padding:16px;border-left:3px solid {color};margin-bottom:10px'>
  <div style='display:flex;align-items:center;gap:8px;margin-bottom:8px'>
    <span style='font-size:18px'>{ico}</span>
    <span style='font-size:14px;font-weight:600;color:#1C1C1E'>{titre}</span>
  </div>
  <div style='font-size:12px;color:#3A3A3C;margin-bottom:4px'>{cond}</div>
  <div style='font-size:11px;color:{color};font-family:monospace;background:#F9F9FB;
              padding:4px 8px;border-radius:6px;margin-bottom:6px'>{formule}</div>
  <div style='font-size:11px;color:#8E8E93;font-style:italic'>→ {action}</div>
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Fichiers attendus</div>", unsafe_allow_html=True)
    cf1, cf2 = st.columns(2)
    with cf1:
        st.markdown("""
<div class='col-required'>
  <div style='font-size:16px'>📋</div>
  <div>
    <div class='col-name'>T1 — Nouvelles Références</div>
    <div class='col-desc'>CSV ou Excel · colonne ARTICLE (code 8 chiffres)</div>
    <div class='col-desc' style='color:#8E8E93;font-size:11px;margin-top:2px'>
      Colonnes optionnelles : MODE APPRO · SEMAINE RECEPTION · LIBELLÉ FOURNISSEUR ORIGINE</div>
  </div>
</div>""", unsafe_allow_html=True)
    with cf2:
        st.markdown("""
<div class='col-required'>
  <div style='font-size:16px'>🏪</div>
  <div>
    <div class='col-name'>Stock ERP par magasin</div>
    <div class='col-desc'>CSV ; latin1 · 1 fichier par site · nommage *_10206_20260507_*.csv</div>
    <div class='col-desc' style='color:#8E8E93;font-size:11px;margin-top:2px'>
      Colonnes : Site · Code article · Code etat · Nouveau stock · Ral · Pcb · Code marketing</div>
  </div>
</div>""", unsafe_allow_html=True)

    st.info("⬆️ Charge les fichiers dans la sidebar pour lancer l'analyse.")
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# CHARGEMENT T1
# ─────────────────────────────────────────────────────────────────────────────
with st.spinner("Lecture T1…"):
    t1_df, t1_err = parse_t1(t1_file.read(), t1_file.name)

if t1_err or t1_df is None:
    st.error(f"❌ T1 : {t1_err}")
    st.stop()

if not stk_files:
    st.markdown(
        f'<div class="info-box blue">✅ T1 chargé — <strong>{len(t1_df):,}</strong> références. '
        f'⬆️ Charge maintenant les fichiers stock (1 par magasin).</div>',
        unsafe_allow_html=True,
    )
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR FILTRES (après chargement T1)
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Filtres</div>", unsafe_allow_html=True)

    orig_sel = st.multiselect("Flux", ["IM", "LO"], default=["IM", "LO"])
    sem_dispo = sorted(
        [s for s in t1_df["SEMAINE RECEPTION"].unique()
         if str(s).strip() not in ("nan", "", "99")],
        key=_sem_sort,
    )
    sem_sel = st.multiselect("Semaine réception", sem_dispo, default=sem_dispo)

    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Cessions</div>", unsafe_allow_html=True)
    mag_labels_sidebar = []  # peuplé après consolidation
    min_1pcb = st.toggle("Qté min = 1 PCB", value=True)


# ─────────────────────────────────────────────────────────────────────────────
# CONSOLIDATION STOCK
# ─────────────────────────────────────────────────────────────────────────────
SKU_SCOPE = tuple(sorted(t1_df["SKU"].unique()))

with st.spinner(f"Consolidation de {len(stk_files)} fichier(s) stock…"):
    df_stock, stk_errors = consolidate_stock(stk_files, SKU_SCOPE)

for err in stk_errors:
    st.warning(f"⚠️ {err}")

if df_stock.empty:
    st.error("Aucune donnée stock valide après consolidation.")
    st.stop()

dates_stock    = df_stock["_date_fichier"].unique()
date_stock_str = pd.Timestamp(max(dates_stock)).strftime("%d %b %Y")
age_stock      = (TODAY - pd.Timestamp(max(dates_stock))).days

if age_stock > 7:
    st.markdown(f"<div class='info-box amber'>⚠️ Données stock du <strong>{date_stock_str}</strong> — {age_stock} jours. Recharge des fichiers plus récents.</div>", unsafe_allow_html=True)

site_ref  = (df_stock[["Code site", "Libellé site"]]
             .drop_duplicates("Code site")
             .set_index("Code site")["Libellé site"].to_dict())
all_codes = sorted(site_ref.keys())

# Filtres T1
mask_t1 = t1_df["ORIGINE"].isin(orig_sel)
if sem_sel:
    mask_t1 = mask_t1 & t1_df["SEMAINE RECEPTION"].isin(sem_sel)
t1_scope  = t1_df[mask_t1].copy()
sku_scope = t1_scope["SKU"].unique()

if len(sku_scope) == 0:
    st.warning("Aucun SKU pour les filtres sélectionnés.")
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# VALIDATION CROISÉE
# ─────────────────────────────────────────────────────────────────────────────
sku_dans_stock = set(df_stock[df_stock["Code site"].isin(all_codes)]["SKU"].unique())
n_sku         = len(sku_scope)
n_sku_trouves = len([s for s in sku_scope if s in sku_dans_stock])
n_sku_absents = n_sku - n_sku_trouves
n_mag         = len(all_codes)

st.markdown(f"""
<div class="val-box">
  <div style="font-size:13px;font-weight:600;color:#1C1C1E;margin-right:6px;">📋 T1 × Stock</div>
  <div class="val-item"><div class="val-num" style="color:#007AFF">{fmt_n(n_sku)}</div><div class="val-lbl">SKUs T1</div></div>
  <div class="val-item"><div class="val-num" style="color:#34C759">{fmt_n(n_sku_trouves)}</div><div class="val-lbl">Trouvés</div></div>
  <div class="val-item"><div class="val-num" style="color:{'#FF3B30' if n_sku_absents > 0 else '#34C759'}">{fmt_n(n_sku_absents)}</div><div class="val-lbl">Absents</div></div>
  <div class="val-item"><div class="val-num" style="color:#8E8E93">{n_mag}</div><div class="val-lbl">Magasins</div></div>
  <div class="val-item"><div class="val-num" style="color:#8E8E93">{len(stk_files)}</div><div class="val-lbl">Fichiers</div></div>
  <div style="margin-left:auto;font-size:11px;color:#8E8E93;">
    Stock du <strong style="color:#007AFF">{date_stock_str}</strong>
    {"&nbsp;·&nbsp;<span style='color:#FF3B30'>⚠️ " + str(age_stock) + "j</span>" if age_stock > 7 else ""}
  </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# CONSTRUCTION DATASET
# ─────────────────────────────────────────────────────────────────────────────
stk_filt = df_stock[df_stock["Code site"].isin(all_codes) & df_stock["SKU"].isin(sku_scope)].copy()

grid = pd.DataFrame(
    pd.MultiIndex.from_product([all_codes, sku_scope], names=["Code site", "SKU"]).tolist(),
    columns=["Code site", "SKU"]
)

KEEP = ["Code site", "SKU", "Nouveau stock", "Ral", "Code etat",
        "Origine", "Libellé article", "Nom fourn.", "Libellé rayon", "Libellé famille", "Pcb"]
merged = grid.merge(stk_filt[[c for c in KEEP if c in stk_filt.columns]], on=["Code site", "SKU"], how="left")

t1_ref = t1_scope.set_index("SKU")[[
    "LIBELLÉ ARTICLE", "LIBELLÉ FOURNISSEUR ORIGINE",
    "MODE APPRO", "SEMAINE RECEPTION", "DATE LIV.", "ORIGINE", "SEM_NUM"
]].rename(columns={
    "LIBELLÉ ARTICLE": "T1_lib", "LIBELLÉ FOURNISSEUR ORIGINE": "Fournisseur T1",
    "MODE APPRO": "Mode Appro", "SEMAINE RECEPTION": "Sem. Réception",
    "DATE LIV.": "Date Livraison", "ORIGINE": "Origine_T1", "SEM_NUM": "SEM_NUM",
})
merged = merged.merge(t1_ref.reset_index(), on="SKU", how="left")

merged["Libellé article"] = merged["Libellé article"].fillna("").astype(str)
merged["Libellé article"] = merged.apply(
    lambda r: r["Libellé article"] if r["Libellé article"] else r.get("T1_lib", ""), axis=1)
merged["Origine"] = merged.apply(
    lambda r: r.get("Origine") if pd.notna(r.get("Origine")) and str(r.get("Origine")) not in ("nan", "")
    else r.get("Origine_T1", "LO"), axis=1)
merged.drop(columns=["T1_lib", "Origine_T1"], errors="ignore", inplace=True)
merged["Magasin"]   = merged["Code site"].map(site_ref).fillna(merged["Code site"])
merged["Code etat"] = merged["Code etat"].fillna("").astype(str)
merged["Ral"]       = pd.to_numeric(merged["Ral"], errors="coerce").fillna(0)
merged["Pcb"]       = pd.to_numeric(merged["Pcb"], errors="coerce").fillna(1)
merged["Alerte"]    = merged.apply(get_alerte, axis=1)
merged["Action"]    = merged["Alerte"].map(ACTION_LABEL)


# ─────────────────────────────────────────────────────────────────────────────
# MÉTRIQUES
# ─────────────────────────────────────────────────────────────────────────────
merged_actif = merged[merged["Code etat"].str.strip() == ETAT_ACTIF]
n_base_taux  = len(merged_actif)
n_impl       = int((merged["Alerte"] == "✅ Implanté").sum())
n_appro      = int((merged["Alerte"] == "🔵 Appro en cours").sum())
n_cmd        = int((merged["Alerte"] == "🛒 Passer commande").sum())
n_anomalie   = int((merged["Alerte"] == "🚩 Anomalie référ.").sum())
taux_reseau  = int(n_impl / n_base_taux * 100) if n_base_taux else 0
n_sku_im     = int((t1_scope["ORIGINE"] == "IM").sum())
n_sku_lo     = int((t1_scope["ORIGINE"] == "LO").sum())
total_cells  = len(merged)

def taux_mag(mag):
    dm = merged_actif[merged_actif["Magasin"] == mag]
    return int((dm["Alerte"] == "✅ Implanté").sum() / len(dm) * 100) if len(dm) else 0

pivot_mag = (
    merged.groupby(["Magasin", "Alerte"]).size()
    .unstack(fill_value=0)
    .reindex(columns=list(ALERTES.keys()), fill_value=0)
    .reset_index()
)
pivot_mag.columns.name = None
pivot_mag["Taux (%)"] = pivot_mag["Magasin"].apply(taux_mag)

rayon_pivot = pd.DataFrame()
if "Libellé rayon" in merged.columns:
    try:
        rayon_pivot = (
            merged_actif.groupby(["Libellé rayon", "Magasin"])
            .apply(lambda x: int((x["Alerte"] == "✅ Implanté").sum() / len(x) * 100) if len(x) else 0)
            .reset_index(name="Taux (%)")
            .pivot(index="Libellé rayon", columns="Magasin", values="Taux (%)")
            .fillna(0).astype(int)
        )
    except Exception:
        rayon_pivot = pd.DataFrame()


# ─────────────────────────────────────────────────────────────────────────────
# KPIs BARRE — style app.py
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-label'>Indicateurs réseau</div>", unsafe_allow_html=True)
k1, k2, k3, k4, k5 = st.columns(5)
for col, val, label in [
    (k1, f"{taux_reseau}%", "Taux implanté"),
    (k2, fmt_n(n_impl),     "✅ Implantés"),
    (k3, fmt_n(n_appro),    "🔵 Appro en cours"),
    (k4, fmt_n(n_cmd),      "🛒 À commander"),
    (k5, fmt_n(n_anomalie), "🚩 Anomalies"),
]:
    col.markdown(f"""
    <div class="kpi-bar">
        <div class="kpi-bar-val">{val}</div>
        <div class="kpi-bar-label">{label}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

k6, k7, k8 = st.columns(3)
for col, val, label in [
    (k6, f"IM {n_sku_im} · LO {n_sku_lo}", "Flux"),
    (k7, fmt_n(n_sku),                      "SKUs analysés"),
    (k8, date_stock_str,                    "Données stock"),
]:
    col.markdown(f"""
    <div class="kpi-bar">
        <div class="kpi-bar-val">{val}</div>
        <div class="kpi-bar-label">{label}</div>
    </div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# BANNIÈRE ACTIONS
# ─────────────────────────────────────────────────────────────────────────────
n_actions = n_appro + n_cmd + n_anomalie
if n_actions > 0:
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.markdown(f"""
<div class="action-banner">
  <span style="font-size:12px;font-weight:700;color:#FF3B30;">⚡ ACTIONS</span>
  <div class="action-item"><div class="action-num" style="color:#007AFF">{fmt_n(n_appro)}</div><div class="action-lbl">Appro en cours</div></div>
  <div class="action-item"><div class="action-num" style="color:#FF9500">{fmt_n(n_cmd)}</div><div class="action-lbl">À commander</div></div>
  <div class="action-item"><div class="action-num" style="color:#B8860B">{fmt_n(n_anomalie)}</div><div class="action-lbl">Anomalies</div></div>
  <div style="margin-left:auto;font-size:11px;color:#8E8E93;">{n_mag} magasin(s) · {fmt_n(n_sku)} SKU</div>
</div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# SCORECARD MAGASINS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Scorecard magasins</div>", unsafe_allow_html=True)
sc_html = '<div class="scorecard-grid">'
for _, row in pivot_mag.sort_values("Taux (%)", ascending=False).iterrows():
    t_   = row["Taux (%)"]
    col  = color_taux(t_)
    impl = int(row.get("✅ Implanté", 0))
    app  = int(row.get("🔵 Appro en cours", 0))
    cmd  = int(row.get("🛒 Passer commande", 0))
    an   = int(row.get("🚩 Anomalie référ.", 0))
    sc_html += f"""
<div class="scorecard-card {scorecard_cls(t_)}">
  <div class="scorecard-dot" style="background:{col}"></div>
  <div class="scorecard-name">{row['Magasin']}</div>
  <div class="scorecard-pct" style="color:{col}">{t_}%</div>
  <div class="scorecard-sub">{impl}✅ {app}🔵 {cmd}🛒 {an}🚩</div>
</div>"""
sc_html += "</div>"
st.markdown(sc_html, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# ONGLETS
# ─────────────────────────────────────────────────────────────────────────────
tab_copil, tab_reseau, tab_alertes, tab_cessions, tab_export = st.tabs([
    "📋 COPIL", "📊 Vue Réseau", "🚨 Alertes", "🔄 Cessions", "📥 Export",
])


# ══ COPIL ═════════════════════════════════════════════════════════════════════
with tab_copil:
    st.markdown("<div class='sh'>SYNTHÈSE PAR MAGASIN</div>", unsafe_allow_html=True)
    cols_aff = ["Magasin"] + [c for c in ALERTES if c in pivot_mag.columns] + ["Taux (%)"]
    st.dataframe(
        pivot_mag[cols_aff].sort_values("Taux (%)", ascending=False).reset_index(drop=True)
        .style.format({"Taux (%)": "{}%"}),
        use_container_width=True, hide_index=True,
    )

    if not rayon_pivot.empty:
        st.markdown("<div class='sh'>TAUX PAR RAYON × MAGASIN</div>", unsafe_allow_html=True)
        def color_cell(val):
            if val >= 80: return "background-color:#D1FAE5;color:#065F46;font-weight:600"
            if val >= 50: return "background-color:#FEF3C7;color:#92400E;font-weight:600"
            return "background-color:#FEE2E2;color:#991B1B;font-weight:600"
        st.dataframe(rayon_pivot.style.map(color_cell).format("{}%"), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("<div class='sh'>🛒 TOP ARTICLES À COMMANDER</div>", unsafe_allow_html=True)
        df_cmd = merged[merged["Alerte"] == "🛒 Passer commande"].groupby("SKU").agg(
            Libellé=("Libellé article", "first"),
            Fournisseur=("Fournisseur T1", "first"),
            Origine=("Origine", "first"),
            Nb_magasins=("Magasin", "count"),
        ).reset_index().sort_values("Nb_magasins", ascending=False).head(10)
        if df_cmd.empty:
            st.markdown("<div class='info-box green'>✅ Aucun article sans commande.</div>", unsafe_allow_html=True)
        else:
            st.dataframe(df_cmd.rename(columns={"Nb_magasins": "Mag. sans stock"}),
                         use_container_width=True, hide_index=True)

    with c2:
        st.markdown("<div class='sh'>🔵 TOP APPROS À ACCÉLÉRER</div>", unsafe_allow_html=True)
        df_acc = merged[merged["Alerte"] == "🔵 Appro en cours"].groupby("SKU").agg(
            Libellé=("Libellé article", "first"),
            Fournisseur=("Fournisseur T1", "first"),
            Origine=("Origine", "first"),
            Nb_magasins=("Magasin", "count"),
            RAL_total=("Ral", "sum"),
        ).reset_index().sort_values("Nb_magasins", ascending=False).head(10)
        if df_acc.empty:
            st.markdown("<div class='info-box green'>✅ Aucune appro en attente.</div>", unsafe_allow_html=True)
        else:
            st.dataframe(df_acc.rename(columns={"Nb_magasins": "Mag. en attente", "RAL_total": "RAL total"}),
                         use_container_width=True, hide_index=True)


# ══ VUE RÉSEAU ════════════════════════════════════════════════════════════════
with tab_reseau:
    c1, c2 = st.columns([3, 2])
    with c1:
        alertes_aff = [a for a in ALERTES if a in pivot_mag.columns and a != "✅ Implanté"]
        mel = pivot_mag.melt(id_vars="Magasin",
                              value_vars=["✅ Implanté"] + alertes_aff,
                              var_name="Alerte", value_name="N")
        fig = px.bar(mel, x="Magasin", y="N", color="Alerte",
                     color_discrete_map=ALERTES, barmode="stack",
                     title="Situation par magasin")
        fig.update_layout(paper_bgcolor="#FFFFFF", plot_bgcolor="#F9F9FB",
                          height=380, font=dict(family="-apple-system, Helvetica Neue", size=12),
                          margin=dict(l=10, r=10, t=44, b=20),
                          legend=dict(orientation="h", y=-0.3))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        labels_d = [a for a in ALERTES if a in pivot_mag.columns]
        vals_d   = [int(pivot_mag[a].sum()) for a in labels_d]
        fig_d = go.Figure(go.Pie(
            labels=labels_d, values=vals_d, hole=0.65,
            marker=dict(colors=[ALERTES[a] for a in labels_d], line=dict(color="#fff", width=3)),
            textinfo="percent", textfont=dict(size=12),
        ))
        fig_d.add_annotation(
            text=f"<b>{taux_reseau}%</b><br>implanté",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=20, color=color_taux(taux_reseau))
        )
        fig_d.update_layout(
            paper_bgcolor="#FFFFFF", height=380,
            margin=dict(l=10, r=10, t=44, b=20),
            legend=dict(orientation="v", x=1.0, y=0.5, font=dict(size=11)),
            title=dict(text="Répartition réseau", font=dict(size=13, color="#8E8E93"))
        )
        st.plotly_chart(fig_d, use_container_width=True)

    if not rayon_pivot.empty:
        st.markdown("<div class='sh'>TAUX PAR RAYON × MAGASIN</div>", unsafe_allow_html=True)
        st.dataframe(rayon_pivot.style.map(color_cell).format("{}%"), use_container_width=True)


# ══ ALERTES ═══════════════════════════════════════════════════════════════════
with tab_alertes:
    alertes_dispo  = [a for a in ALERTES if a != "✅ Implanté" and (merged["Alerte"] == a).any()]
    alerte_sel     = st.multiselect("Filtrer par alerte", alertes_dispo, default=alertes_dispo,
                                     format_func=lambda a: f"{a} — {ACTION_LABEL[a]}")
    mag_alerte_sel = st.multiselect("Filtrer par magasin", sorted(merged["Magasin"].unique()),
                                     default=sorted(merged["Magasin"].unique()))

    # Ruptures communes
    df_non_impl = merged[merged["Alerte"].isin(["🛒 Passer commande", "🔵 Appro en cours"])]
    sku_rupt    = df_non_impl.groupby("SKU")["Magasin"].count()
    sku_rupt    = sku_rupt[sku_rupt == n_mag].index.tolist()
    if sku_rupt:
        st.markdown(f"""
<div class='alert-card alert-red'>
  <strong>🚨 {len(sku_rupt)} rupture(s) communes réseau</strong> — aucun stock sur tous les magasins<br>
  <span style='font-size:12px;opacity:.85'>→ Escalade critique — commander en urgence.</span>
</div>""", unsafe_allow_html=True)
        with st.expander(f"Voir les {len(sku_rupt)} ruptures communes"):
            df_rupt = (merged[merged["SKU"].isin(sku_rupt)]
                       [["SKU", "Libellé article", "Origine", "Fournisseur T1", "Alerte"]]
                       .drop_duplicates("SKU").sort_values("Alerte"))
            st.dataframe(df_rupt, use_container_width=True, hide_index=True)

    df_al = merged[merged["Alerte"].isin(alerte_sel) & merged["Magasin"].isin(mag_alerte_sel)]
    if df_al.empty:
        st.markdown("<div class='info-box green'>✅ Aucune alerte pour les filtres sélectionnés.</div>", unsafe_allow_html=True)
    else:
        for alerte in alerte_sel:
            df_a = df_al[df_al["Alerte"] == alerte]
            if df_a.empty: continue
            color = ALERTES.get(alerte, "#8E8E93")
            st.markdown(f"""
<div class='alert-card' style='background:{color}18;border-color:{color};color:#1C1C1E'>
  <strong>{alerte}</strong> — {len(df_a)} ligne(s)<br>
  <span style='font-size:12px;opacity:.85'>→ {ACTION_LABEL.get(alerte, '')}</span>
</div>""", unsafe_allow_html=True)

        COLS_AL = ["Magasin", "SKU", "Libellé article", "Origine", "Code etat",
                   "Nouveau stock", "Ral", "Mode Appro", "Sem. Réception",
                   "Fournisseur T1", "Alerte", "Action"]
        st.dataframe(
            df_al[[c for c in COLS_AL if c in df_al.columns]]
            .sort_values(["Alerte", "Magasin"]).reset_index(drop=True),
            use_container_width=True, hide_index=True,
        )


# ══ CESSIONS ══════════════════════════════════════════════════════════════════
with tab_cessions:
    st.markdown("<div class='sh'>MOTEUR CESSIONS INTER-MAGASINS</div>", unsafe_allow_html=True)

    # Suggestions automatiques (ruptures totales)
    df_non_impl_all = merged[merged["Alerte"].isin(["🛒 Passer commande", "🔵 Appro en cours"])]
    sku_rupt_tot    = (df_non_impl_all.groupby("SKU")["Magasin"].count()
                       .pipe(lambda s: s[s == n_mag]).index.tolist())
    if sku_rupt_tot:
        st.markdown(f'<div class="info-box blue">🤖 <strong>{len(sku_rupt_tot)} article(s)</strong> sans stock sur tous les magasins — suggestions automatiques.</div>', unsafe_allow_html=True)
        auto_sugg = []
        for sku in sku_rupt_tot:
            sku_df = df_stock[df_stock["SKU"] == sku].copy()
            sku_df["Nouveau stock"] = pd.to_numeric(sku_df["Nouveau stock"], errors="coerce").fillna(0)
            sku_df["Pcb"]           = pd.to_numeric(sku_df.get("Pcb", pd.Series(1, index=sku_df.index)), errors="coerce").fillna(1).clip(lower=1)
            sku_df["Reserve_2pcb"]  = (sku_df["Pcb"] * 2).astype(int)
            best = sku_df[sku_df["Nouveau stock"] > sku_df["Reserve_2pcb"]].sort_values("Nouveau stock", ascending=False)
            lib  = sku_df["Libellé article"].iloc[0] if len(sku_df) else sku
            if not best.empty:
                b   = best.iloc[0]
                qty = int(b["Nouveau stock"]) - int(b["Reserve_2pcb"])
                auto_sugg.append({"SKU": sku, "Libellé": lib,
                                   "Cédant": site_ref.get(b["Code site"], b["Code site"]),
                                   "Stock cédant": int(b["Nouveau stock"]), "Qté cessible": qty})
        if auto_sugg:
            st.dataframe(pd.DataFrame(auto_sugg), use_container_width=True, hide_index=True)
        else:
            st.info("Aucun magasin cédant disponible.")

    # Cessions manuelles
    mag_labels_all = sorted([site_ref.get(c, c) for c in all_codes])
    mag_detresse   = st.multiselect("Magasins en détresse", mag_labels_all, default=[])
    seuil_det      = st.number_input("Seuil stock (≤)", 0, 50, 0, 1)

    if mag_detresse:
        mag_det_codes = [c for c in all_codes if site_ref.get(c, c) in mag_detresse]
        mag_ced_codes = [c for c in all_codes if c not in mag_det_codes]
        suggestions   = []
        for sku in sku_scope:
            sku_df = df_stock[df_stock["SKU"] == sku].copy()
            if sku_df.empty: continue
            lib = sku_df["Libellé article"].iloc[0] if "Libellé article" in sku_df.columns else sku
            sku_df["Nouveau stock"] = pd.to_numeric(sku_df["Nouveau stock"], errors="coerce").fillna(0)
            sku_df["Pcb"]           = pd.to_numeric(sku_df.get("Pcb", pd.Series(1, index=sku_df.index)), errors="coerce").fillna(1).clip(lower=1)
            det_rows  = sku_df[sku_df["Code site"].isin(mag_det_codes) & (sku_df["Nouveau stock"] <= seuil_det)]
            if det_rows.empty: continue
            sku_df["Reserve_2pcb"] = (sku_df["Pcb"] * 2).astype(int)
            ced_rows = sku_df[sku_df["Code site"].isin(mag_ced_codes) &
                               (sku_df["Nouveau stock"] > sku_df["Reserve_2pcb"])].sort_values("Nouveau stock", ascending=False)
            for _, dr in det_rows.iterrows():
                if ced_rows.empty:
                    suggestions.append({"SKU": sku, "Libellé": lib,
                                         "Magasin détresse": site_ref.get(dr["Code site"], dr["Code site"]),
                                         "Stock détresse": int(dr["Nouveau stock"]),
                                         "Cédant": "⚠️ Aucun", "Stock cédant": 0, "Qté cessible": 0,
                                         "Faisabilité": "🔴 Impossible"})
                else:
                    best = ced_rows.iloc[0]
                    qty  = int(best["Nouveau stock"]) - int(best["Reserve_2pcb"])
                    suggestions.append({"SKU": sku, "Libellé": lib,
                                         "Magasin détresse": site_ref.get(dr["Code site"], dr["Code site"]),
                                         "Stock détresse": int(dr["Nouveau stock"]),
                                         "Cédant": site_ref.get(best["Code site"], best["Code site"]),
                                         "Stock cédant": int(best["Nouveau stock"]),
                                         "Réserve (2 PCB)": int(best["Reserve_2pcb"]),
                                         "Qté cessible": qty,
                                         "Faisabilité": "🟢 Possible" if qty >= 1 else "🟠 Partielle"})

        if not suggestions:
            st.markdown("<div class='info-box green'>✅ Aucune cession nécessaire.</div>", unsafe_allow_html=True)
        else:
            df_all  = pd.DataFrame(suggestions)
            df_cess = df_all[df_all["Faisabilité"] == "🟢 Possible"].copy()
            if min_1pcb and "Réserve (2 PCB)" in df_cess.columns:
                df_cess = df_cess[df_cess["Qté cessible"] >= (df_cess["Réserve (2 PCB)"] / 2).clip(lower=1)]
            df_cess = df_cess.sort_values("Qté cessible", ascending=False).reset_index(drop=True)
            st.dataframe(df_cess, use_container_width=True, hide_index=True)
            buf_c = io.BytesIO()
            with pd.ExcelWriter(buf_c, engine="openpyxl") as w:
                df_cess.to_excel(w, sheet_name="Plan Cessions", index=False)
            buf_c.seek(0)
            st.download_button(f"📥 Plan_Cessions_{TODAY_FILE}.xlsx", data=buf_c,
                               file_name=f"Plan_Cessions_{TODAY_FILE}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ══ EXPORT ════════════════════════════════════════════════════════════════════
with tab_export:
    from openpyxl.styles import Font as XFont, PatternFill as XFill, Alignment as XAlignment
    from openpyxl.utils import get_column_letter as gcl

    st.markdown("""
<div class='info-box blue'>
  <strong>📋 Contenu de l'export (3 feuilles)</strong><br>
  <strong>Synthèse Réseau</strong> — taux par magasin · <strong>Détail Complet</strong> — toutes lignes ·
  <strong>Alertes & Actions</strong> — lignes non implantées uniquement
</div>""", unsafe_allow_html=True)

    st.caption(f"Données stock : {date_stock_str} · {fmt_n(len(merged))} lignes · {n_mag} magasin(s)")

    if st.button("Générer l'export Excel", type="primary"):
        ALERTE_FILLS = {
            "✅ Implanté":        ("D1FAE5", "065F46"),
            "🔵 Appro en cours":  ("DBEAFE", "1D4ED8"),
            "🛒 Passer commande": ("FEF3C7", "92400E"),
            "🚩 Anomalie référ.": ("FFFDE7", "795548"),
        }
        COLS_DET = ["Magasin", "SKU", "Libellé article", "Origine", "Code etat",
                    "Nouveau stock", "Ral", "Mode Appro", "Sem. Réception",
                    "Fournisseur T1", "Alerte", "Action"]
        buf_x = io.BytesIO()
        with pd.ExcelWriter(buf_x, engine="openpyxl") as writer:
            cols_s = ["Magasin"] + [c for c in ALERTES if c in pivot_mag.columns] + ["Taux (%)"]
            pivot_mag[cols_s].sort_values("Taux (%)", ascending=False).to_excel(
                writer, sheet_name="Synthèse Réseau", index=False)
            merged[[c for c in COLS_DET if c in merged.columns]].to_excel(
                writer, sheet_name="Détail Complet", index=False)
            merged[merged["Alerte"] != "✅ Implanté"][[c for c in COLS_DET if c in merged.columns]]\
                .sort_values(["Alerte", "Magasin"]).to_excel(writer, sheet_name="Alertes & Actions", index=False)
            wb = writer.book
            FH = XFill("solid", fgColor="1C3557")
            FT = XFont(bold=True, color="FFFFFF", name="Arial", size=11)
            for sn in wb.sheetnames:
                ws = wb[sn]
                for cell in ws[1]:
                    cell.fill = FH; cell.font = FT
                    cell.alignment = XAlignment(horizontal="center")
                for col in ws.columns:
                    ws.column_dimensions[gcl(col[0].column)].width = min(
                        max((len(str(c.value)) for c in col if c.value), default=10) + 4, 50)
                ws.freeze_panes = "A2"
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        v = str(cell.value)
                        if v in ALERTE_FILLS:
                            bg, fg = ALERTE_FILLS[v]
                            cell.fill = XFill("solid", fgColor=bg)
                            cell.font = XFont(color=fg, name="Arial", size=10)
        buf_x.seek(0)
        st.download_button(
            f"📥 Implantation_T1_{TODAY_FILE}.xlsx", data=buf_x,
            file_name=f"Implantation_T1_{TODAY_FILE}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success(f"✅ Export — {fmt_n(len(merged))} lignes · 3 feuilles")


# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(f"""
<div style='text-align:center;color:#C7C7CC;font-size:11px;padding:8px 0'>
    NovaRetail Solutions · SmartBuyer v2.3 · Implantation · {TODAY_STR} · Données : {date_stock_str}
</div>
""", unsafe_allow_html=True)
