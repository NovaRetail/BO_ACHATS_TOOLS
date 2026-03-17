"""
03_📦_Detention_Top_CA.py — SmartBuyer Hub
Taux de détention Top CA · Flux IM/LO · Code état par article × magasin
Charte SmartBuyer v2
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Détention Top CA · SmartBuyer",
    page_icon="📦",
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
[data-testid="stSidebar"] { background: #F2F2F7 !important; border-right: 0.5px solid #D1D1D6 !important; }
[data-testid="stMetric"] { background: #FFFFFF !important; border: 0.5px solid #E5E5EA !important; border-radius: 12px !important; padding: 16px 18px !important; }
[data-testid="stMetricLabel"] { font-size: 11px !important; font-weight: 500 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 600 !important; color: #1C1C1E !important; letter-spacing: -0.02em !important; }
[data-testid="stMetricDelta"] { font-size: 12px !important; }
[data-testid="stTabs"] button[role="tab"] { font-size: 13px !important; font-weight: 500 !important; padding: 8px 16px !important; color: #8E8E93 !important; border-radius: 0 !important; border-bottom: 2px solid transparent !important; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color: #007AFF !important; border-bottom: 2px solid #007AFF !important; background: transparent !important; }
[data-testid="stTabs"] [role="tablist"] { border-bottom: 0.5px solid #E5E5EA !important; }
[data-testid="stDataFrame"] { border: 0.5px solid #E5E5EA !important; border-radius: 10px !important; }
[data-testid="stDataFrame"] th { background: #F2F2F7 !important; font-size: 11px !important; font-weight: 600 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stFileUploader"] { border: 1.5px dashed #D1D1D6 !important; border-radius: 10px !important; background: #F9F9FB !important; }
[data-testid="baseButton-primary"] { background: #007AFF !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; }
.stDownloadButton > button { background: #007AFF !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; font-size: 13px !important; padding: 10px 24px !important; width: 100% !important; }
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
.col-name     { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc     { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
.col-example { font-size: 11px; color: #8E8E93; font-family: monospace; margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTES ───────────────────────────────────────────────────────────────
ETAT_LABELS = {
    "2": ("Actif", "#34C759",  "#F0FFF4", "Inclus dans le taux"),
    "P": ("Permanent", "#007AFF", "#E6F1FB", "Exclu — signalé"),
    "S": ("Saisonnier", "#FF9500", "#FFFBF0", "Exclu — signalé"),
    "B": ("Bloqué", "#FF3B30",  "#FFF2F2", "Exclu — signalé"),
    "F": ("Fin de vie", "#8E8E93", "#F2F2F7", "Exclu — à retirer"),
    "1": ("Autre", "#8E8E93", "#F2F2F7", "Exclu"),
    "5": ("Autre", "#8E8E93", "#F2F2F7", "Exclu"),
    "6": ("Autre", "#8E8E93", "#F2F2F7", "Exclu"),
}
FLUX_LABELS = {"IM": ("Import", "#7C3AED", "#F0EEFF"), "LO": ("Local", "#007AFF", "#E6F1FB")}

def fmt(n):
    if pd.isna(n): return "—"
    a = abs(n)
    if a >= 1_000_000: return f"{n/1_000_000:.1f} M"
    if a >= 1_000:     return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def norm_code(s):
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(8)

# ─── PARSING ──────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_stock(files_data):
    dfs = []
    for byt, name in files_data:
        try:
            df = pd.read_csv(BytesIO(byt), sep=";", encoding="utf-8-sig",
                             dtype=str, low_memory=False)
            dfs.append(df)
        except Exception as e:
            st.warning(f"Erreur lecture {name} : {e}")
    if not dfs:
        return pd.DataFrame()
    raw = pd.concat(dfs, ignore_index=True)

    for col in ["Nouveau stock", "Ral", "Nb colis", "Prix de vente", "Prix d'achat"]:
        if col in raw.columns:
            raw[col] = pd.to_numeric(raw[col], errors="coerce").fillna(0)

    raw["Code article"] = norm_code(raw["Code article"])
    raw["Code etat"]    = raw["Code etat"].astype(str).str.strip().str.upper()
    raw["Code marketing"] = raw["Code marketing"].astype(str).str.strip().str.upper() if "Code marketing" in raw.columns else "?"
    raw["Libellé marketing"] = raw.get("Libellé marketing", pd.Series("?", index=raw.index)).fillna("?")

    PGC = {"BOISSONS","DROGUERIE","PARFUMERIE HYGIENE","EPICERIE"}
    if "Libellé rayon" in raw.columns:
        raw = raw[raw["Libellé rayon"].str.upper().isin(PGC)]

    return raw

@st.cache_data(show_spinner=False)
def load_topca(file_bytes):
    try:
        # On tente le CSV d'abord
        df = pd.read_csv(BytesIO(file_bytes), header=None, names=["Code article"], dtype=str)
        # Si on a trop de colonnes, c'est probablement un mauvais format (Excel mal lu en CSV)
        if df.shape[1] > 10: raise ValueError()
    except Exception:
        # En cas d'erreur ou mauvais format, on tente l'Excel
        df = pd.read_excel(BytesIO(file_bytes), header=None, names=["Code article"], dtype=str)
    df["Code article"] = norm_code(df["Code article"])
    return set(df["Code article"].dropna().unique())

# ─── CALCUL DÉTENTION ────────────────────────────────────────────────────────
def compute_detention(df_stock, top_codes):
    df = df_stock[df_stock["Code article"].isin(top_codes)].copy()

    grp = df.groupby(["Code article","Libellé site","Code marketing","Libellé marketing"]).agg(
        code_etat       = ("Code etat",       lambda x: x.mode().iloc[0] if len(x) else "?"),
        stock           = ("Nouveau stock",   "sum"),
        ral             = ("Ral",             "sum"),
        nb_colis        = ("Nb colis",        "first"),
        lib_article     = ("Libellé article", "first"),
        lib_rayon       = ("Libellé rayon",   "first") if "Libellé rayon" in df.columns else ("Code article","first"),
        lib_fournisseur = ("Nom fourn.",      "first") if "Nom fourn." in df.columns else ("Code article","first"),
    ).reset_index()

    found = set(df["Code article"].unique())
    absents = sorted(top_codes - found)

    return grp, absents

def compute_taux(grp, top_codes):
    sites = sorted(grp["Libellé site"].unique())
    rows = []
    for site in sites:
        s = grp[grp["Libellé site"] == site]
        for flux in ["IM","LO","ALL"]:
            if flux == "ALL":
                sf = s
            else:
                sf = s[s["Code marketing"] == flux]

            actifs      = sf[sf["code_etat"] == "2"]
            n_actifs    = len(actifs)
            n_stock     = (actifs["stock"] > 0).sum()
            taux        = n_stock / n_actifs * 100 if n_actifs > 0 else None
            n_bloques   = (sf["code_etat"] == "B").sum()
            n_autres    = (sf["code_etat"].isin(["P","S","F","1","5","6"])).sum()
            n_rupture   =
