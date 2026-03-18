"""
06_💸_Marges_Negatives.py — SmartBuyer Hub
Diagnostic Rentabilité Réseau · Flop 100 · Analyse par format et rayon
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
    page_title="Marges Négatives · SmartBuyer",
    page_icon="💸",
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
.main .block-container { padding-top: 1.8rem; max-width: 1300px; }
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
.alert-purple{ background: #F5F0FF; border-color: #AF52DE; color: #1A0033; }

/* Bloc format */
.format-card { border-radius: 12px; padding: 14px 16px; margin-bottom: 6px; border: 0.5px solid; }
.format-hyper  { background: #EFF6FF; border-color: #B3D9FF; }
.format-market { background: #F0FFF4; border-color: #A8E6BF; }
.format-supeco { background: #F5F0FF; border-color: #D9B3FF; }

/* Badge format */
.badge { display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: 600; }
.badge-hyper  { background: #154360; color: #FFFFFF; }
.badge-market { background: #145A32; color: #FFFFFF; }
.badge-supeco { background: #6E2F8A; color: #FFFFFF; }

.col-required { background: #F0F8FF; border: 0.5px solid #B3D9FF; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; display: flex; align-items: flex-start; gap: 10px; }
.col-name { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def fmt(n):
    if pd.isna(n) or n is None: return "—"
    a = abs(n)
    if a >= 1_000_000: return f"{n/1_000_000:.1f} M"
    if a >= 1_000:     return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def fmt_pct(v, dec=1):
    if pd.isna(v) or v is None: return "—"
    return f"{v:.{dec}f}%"

def fmt_delta(v):
    if pd.isna(v) or v is None: return "—"
    return f"{v:+.1f} pts"

def get_format(site_name):
    s = str(site_name)
    if "Supeco" in s: return "Supeco"
    if "Hyper"  in s: return "Hyper"
    return "Market"

def short_name(s):
    """Retourne le libellé après le ' - '"""
    s = str(s)
    return s.split(" - ", 1)[-1].strip() if " - " in s else s

def extract_periode(df_raw):
    try:
        for val in df_raw.iloc[:, 0].astype(str):
            m = re.search(r"après le (\d{2}/\d{2}/\d{4}) et est avant le (\d{2}/\d{2}/\d{4})", val)
            if m:
                from datetime import datetime
                d1 = datetime.strptime(m.group(1), "%d/%m/%Y")
                d2 = datetime.strptime(m.group(2), "%d/%m/%Y")
                nb = (d2 - d1).days
                return f"{m.group(1)} → {m.group(2)}", nb
    except: pass
    return "Période inconnue", 1

# ─── CHARGEMENT ───────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(byt, fname):
    ext = fname.lower().rsplit(".", 1)[-1]
    if ext in ("xlsx", "xls"):
        df = pd.read_excel(BytesIO(byt), dtype=str)
    else:
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                df = pd.read_csv(BytesIO(byt), sep=";", encoding=enc, dtype=str)
                break
            except: continue

    periode, nb_jours = extract_periode(df)

    num_cols = ["CA", "Marge", "CA Hors Promo", "Marge Hors Promo",
                "CA HT Promo", "Marge Promo", "Qté Vente",
                "Casse (Valeur)", "Casse (Qté)", "%Marge", "%CA Poids Promo"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Libellés courts
    df["lib_art"]    = df["Article"].apply(       lambda s: short_name(s) if pd.notna(s) else None)
    df["code_art"]   = df["Article"].apply(       lambda s: str(s).split(" - ", 1)[0].strip() if pd.notna(s) and " - " in str(s) else None)
    df["lib_site"]   = df["Site nom long"].apply( lambda s: short_name(s) if pd.notna(s) else None)
    df["lib_rayon"]  = df["Rayon"].apply(         lambda s: short_name(s) if pd.notna(s) else None)
    df["lib_fam"]    = df["Famille"].apply(       lambda s: short_name(s) if pd.notna(s) else None)
    df["format"]     = df["Site nom long"].apply( lambda s: get_format(s) if pd.notna(s) else None)

    # Nettoyage : garder uniquement lignes article × site réelles
    df_clean = df[
        df["lib_art"].notna()  & (df["lib_art"]  != "Total") &
        df["lib_site"].notna() & (df["lib_site"] != "Total") &
        df["lib_rayon"].notna()&
        ~df["Rayon"].astype(str).str.startswith("Filtres") &
        (df["lib_rayon"] != "Total") &
        df["lib_fam"].notna()  & (df["lib_fam"]  != "Total")
    ].copy()

    return df_clean, periode, nb_jours


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:18px'>
  <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>🛍️ SmartBuyer</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · Équipe Achats</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Navigation</div>", unsafe_allow_html=True)
    st.page_link("app.py",                                       label="🏠  Accueil")
    st.page_link("pages/01_📊_Analyse_Scoring_ABC.py",           label="📊  Scoring ABC")
    st.page_link("pages/02_📈_Ventes_PBI.py",                    label="📈  Ventes PBI")
    st.page_link("pages/03_📦_Detention_Top_CA.py",              label="📦  Détention Top CA")
    st.page_link("pages/04_💸_Performance_Promo.py",             label="💸  Performance Promo")
    st.page_link("pages/05_🏪_Suivi_Implantation.py",            label="🏪  Suivi Implantation")
    st.page_link("pages/06_💸_Marges_Negatives.py",              label="💸  Marges Négatives")
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Import fichier</div>", unsafe_allow_html=True)
    f_pbi = st.file_uploader("Export PBI ventes (Excel)", type=["xlsx", "xls", "csv"], key="pbi_marge")

# ─── HEADER PAGE ──────────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>💸 Diagnostic Rentabilité Réseau</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Analyse des marges · Flop 100 destructeurs · Décomposition par format (Hyper / Market / Supeco) · Fuites de valeur</div>", unsafe_allow_html=True)

# ─── ÉCRAN D'ACCUEIL ──────────────────────────────────────────────────────────
if not f_pbi:
    st.markdown("---")
    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>ℹ️ À quoi sert ce module ?</strong><br>
  Ce module réalise un <strong>diagnostic complet de la rentabilité réseau</strong> à partir d'un export PBI ventes.
  Il identifie précisément <strong>où se perdent les marges</strong> : par rayon, par format de magasin, par article.
  <br><br>
  <strong>1. Vue réseau globale</strong> — KPIs, synthèse par format et par rayon, palmarès magasins<br>
  <strong>2. Matrice rayon × magasin</strong> — Taux de marge croisé pour repérer les combinaisons critiques<br>
  <strong>3. Flop 100</strong> — Articles destructeurs de marge avec impact par site (Hyper / Market / Supeco)<br>
  <strong>4. Analyse des fuites</strong> — Effet promo, casse, familles sous seuil de rentabilité
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Les 5 indicateurs de fuite de valeur analysés</div>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    indics = [
        ("📉", "Marge négative", "#FF3B30",
         "Articles dont la marge brute est inférieure à 0",
         "Marge = CA − (PA × Qté)",
         "Signal immédiat : le magasin perd de l'argent sur chaque unité vendue."),
        ("🔀", "Effet promo sur la marge", "#FF9500",
         "Écart entre marge hors promo et marge sous promotion",
         "Δ = Tx marge HP − Tx marge Promo",
         "Mesure l'érosion de marge causée par la mécanique promotionnelle."),
        ("🏪", "Écart de rentabilité par format", "#AF52DE",
         "Différence de taux de marge entre Hypers, Markets et Supecos",
         "Tx marge format vs moyenne réseau",
         "Identifie le format qui dégrade la rentabilité globale."),
        ("🗑️", "Taux de casse", "#8E8E93",
         "Valeur perdue en démarque sur le CA total",
         "Casse (valeur) ÷ CA × 100",
         "Un taux > 1% sur un site est un signal d'alerte opérationnel."),
        ("📦", "Familles sous seuil", "#007AFF",
         "Familles de produits à marge structurellement basse (< 8%)",
         "Tx marge famille sur la période",
         "Ces familles absorbent du CA sans générer de marge suffisante pour couvrir les charges."),
    ]
    for i, (ico, titre, color, desc, formule, interp) in enumerate(indics):
        with (c1 if i % 2 == 0 else c2):
            st.markdown(f"""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;
            padding:16px;border-left:3px solid {color};margin-bottom:10px'>
  <div style='display:flex;align-items:center;gap:8px;margin-bottom:8px'>
    <span style='font-size:18px'>{ico}</span>
    <span style='font-size:14px;font-weight:600;color:#1C1C1E'>{titre}</span>
  </div>
  <div style='font-size:12px;color:#3A3A3C;margin-bottom:4px'>{desc}</div>
  <div style='font-size:11px;color:{color};font-family:monospace;background:#F9F9FB;
              padding:4px 8px;border-radius:6px;margin-bottom:6px'>{formule}</div>
  <div style='font-size:11px;color:#8E8E93;font-style:italic'>{interp}</div>
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Fichier attendu</div>", unsafe_allow_html=True)
    st.markdown("""
<div class='col-required'><div style='font-size:16px'>📊</div>
<div><div class='col-name'>Export PBI ventes réseau</div>
<div class='col-desc'>Excel · Axes : Rayon / Famille / Article / Site nom long · Colonnes : CA, Marge, CA HT Promo, Marge Promo, CA Hors Promo, Marge Hors Promo, Qté Vente, Casse (Valeur)</div>
<div class='col-desc' style='margin-top:4px'>Le fichier doit inclure tous les formats de magasins (Hyper, Market, Supeco) pour une analyse réseau complète.</div>
</div></div>""", unsafe_allow_html=True)

    st.info("⬆️ Charge le fichier export PBI dans la sidebar pour lancer le diagnostic.")
    st.stop()

# ─── CHARGEMENT & CALCULS ─────────────────────────────────────────────────────
with st.spinner("Lecture et analyse des données…"):
    df, periode, nb_jours = load_data(f_pbi.read(), f_pbi.name)

if df.empty:
    st.error("Fichier vide ou colonnes non reconnues. Vérifier le format de l'export PBI.")
    st.stop()

# ── Filtres sidebar ────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Filtres</div>", unsafe_allow_html=True)

    formats_dispo = sorted(df["format"].dropna().unique())
    sel_format = st.multiselect("Format magasin", formats_dispo, default=formats_dispo)

    rayons_dispo = sorted(df["lib_rayon"].dropna().unique())
    sel_rayon = st.multiselect("Rayon", rayons_dispo, default=rayons_dispo)

    sites_dispo = sorted(df[df["format"].isin(sel_format)]["lib_site"].dropna().unique())
    sel_site = st.multiselect("Magasin", sites_dispo, default=sites_dispo)

    st.markdown("---")
    st.caption(f"**Période :** {periode}")
    st.caption(f"**Durée :** {nb_jours} jour(s)")

df_f = df[
    df["format"].isin(sel_format) &
    df["lib_rayon"].isin(sel_rayon) &
    df["lib_site"].isin(sel_site)
].copy()

if df_f.empty:
    st.warning("Aucune donnée pour la sélection en cours.")
    st.stop()

# ── Agrégations principales ────────────────────────────────────────────────────
sites_actifs = df_f[df_f["CA"] > 0]["lib_site"].unique()
nb_sites_actifs = len(sites_actifs)

# Par format
agg_fmt = df_f.groupby("format").agg(
    CA=("CA","sum"), Marge=("Marge","sum"),
    CA_Promo=("CA HT Promo","sum"), CA_HP=("CA Hors Promo","sum"),
    Marge_HP=("Marge Hors Promo","sum"), Marge_Promo=("Marge Promo","sum"),
    Casse=("Casse (Valeur)","sum")
).reset_index()
agg_fmt["TxMarge"]       = (agg_fmt["Marge"] / agg_fmt["CA"] * 100).where(agg_fmt["CA"] > 0)
agg_fmt["PdsPromo"]      = (agg_fmt["CA_Promo"] / agg_fmt["CA"] * 100).where(agg_fmt["CA"] > 0)
agg_fmt["TxMarge_HP"]    = (agg_fmt["Marge_HP"] / agg_fmt["CA_HP"] * 100).where(agg_fmt["CA_HP"] > 0)
agg_fmt["TxMarge_Promo"] = (agg_fmt["Marge_Promo"] / agg_fmt["CA_Promo"] * 100).where(agg_fmt["CA_Promo"] > 0)
agg_fmt["TxCasse"]       = (agg_fmt["Casse"].abs() / agg_fmt["CA"] * 100).where(agg_fmt["CA"] > 0)
agg_fmt = agg_fmt.sort_values("TxMarge", ascending=False)

# Par rayon
agg_rax = df_f.groupby("lib_rayon").agg(
    CA=("CA","sum"), Marge=("Marge","sum"),
    CA_Promo=("CA HT Promo","sum"), CA_HP=("CA Hors Promo","sum"),
    Marge_HP=("Marge Hors Promo","sum"), Marge_Promo=("Marge Promo","sum"),
    Casse=("Casse (Valeur)","sum")
).reset_index()
agg_rax["TxMarge"]       = (agg_rax["Marge"] / agg_rax["CA"] * 100).where(agg_rax["CA"] > 0)
agg_rax["PdsPromo"]      = (agg_rax["CA_Promo"] / agg_rax["CA"] * 100).where(agg_rax["CA"] > 0)
agg_rax["TxMarge_HP"]    = (agg_rax["Marge_HP"] / agg_rax["CA_HP"] * 100).where(agg_rax["CA_HP"] > 0)
agg_rax["TxMarge_Promo"] = (agg_rax["Marge_Promo"] / agg_rax["CA_Promo"] * 100).where(agg_rax["CA_Promo"] > 0)
agg_rax["TxCasse"]       = (agg_rax["Casse"].abs() / agg_rax["CA"] * 100).where(agg_rax["CA"] > 0)
agg_rax["PoidsCA"]       = (agg_rax["CA"] / agg_rax["CA"].sum() * 100)
agg_rax = agg_rax.sort_values("TxMarge")

# Par magasin
agg_site = df_f.groupby(["lib_site","format"]).agg(
    CA=("CA","sum"), Marge=("Marge","sum"),
    CA_Promo=("CA HT Promo","sum"), Casse=("Casse (Valeur)","sum")
).reset_index()
agg_site["TxMarge"]  = (agg_site["Marge"] / agg_site["CA"] * 100).where(agg_site["CA"] > 0)
agg_site["PdsPromo"] = (agg_site["CA_Promo"] / agg_site["CA"] * 100).where(agg_site["CA"] > 0)
agg_site["TxCasse"]  = (agg_site["Casse"].abs() / agg_site["CA"] * 100).where(agg_site["CA"] > 0)
agg_site = agg_site[agg_site["CA"] > 0].sort_values("TxMarge", ascending=False).reset_index(drop=True)
moy_marge_site = agg_site["TxMarge"].mean()

# Matrice rayon × site
mat = df_f.groupby(["lib_site", "lib_rayon"]).agg(
    CA=("CA","sum"), Marge=("Marge","sum")
).reset_index()
mat["TxMarge"] = (mat["Marge"] / mat["CA"] * 100).where(mat["CA"] > 0)
mat_pivot = mat.pivot_table(index="lib_rayon", columns="lib_site", values="TxMarge").round(1)

# Article × site (pour blocs format)
art_site = df_f[df_f["CA"] > 0].groupby(["Article", "lib_site", "format"]).agg(
    CA=("CA","sum"), Marge=("Marge","sum")
).reset_index()
art_site["TxMarge_site"] = (art_site["Marge"] / art_site["CA"] * 100).where(art_site["CA"] > 0)
art_site["lib_court"]    = art_site["lib_site"]

# Agrégation article globale
agg_art = df_f.groupby(["Article", "lib_art", "lib_rayon", "lib_fam"]).agg(
    CA=("CA","sum"), Marge=("Marge","sum"),
    CA_Promo=("CA HT Promo","sum"), Qte=("Qté Vente","sum")
).reset_index()
agg_art["TxMarge"]  = (agg_art["Marge"] / agg_art["CA"] * 100).where(agg_art["CA"] > 0)
agg_art["PdsPromo"] = (agg_art["CA_Promo"] / agg_art["CA"] * 100).where(agg_art["CA"] > 0)

# Flop 100 : pire taux de marge, seuil CA > 5 000 FCFA
flop100 = agg_art[agg_art["CA"] > 5000].nsmallest(100, "TxMarge").copy()
flop100 = flop100.reset_index(drop=True)
flop100["Rang"] = range(1, len(flop100) + 1)

# Construire les blocs format pour chaque article du flop
def build_bloc(article_full, fmt_name):
    rows = art_site[
        (art_site["Article"] == article_full) &
        (art_site["format"]  == fmt_name)
    ].sort_values("TxMarge_site")
    if rows.empty:
        return "—"
    parts = []
    for _, r in rows.iterrows():
        tm = r["TxMarge_site"]
        if pd.notna(tm):
            parts.append(f"{r['lib_court']}: {tm:.1f}%")
    return "  |  ".join(parts) if parts else "—"

flop100["Bloc_Hyper"]  = flop100["Article"].apply(lambda a: build_bloc(a, "Hyper"))
flop100["Bloc_Market"] = flop100["Article"].apply(lambda a: build_bloc(a, "Market"))
flop100["Bloc_Supeco"] = flop100["Article"].apply(lambda a: build_bloc(a, "Supeco"))

# KPIs globaux
ca_total     = df_f["CA"].sum()
marge_total  = df_f["Marge"].sum()
ca_promo     = df_f["CA HT Promo"].sum()
ca_hp        = df_f["CA Hors Promo"].sum()
m_promo      = df_f["Marge Promo"].sum()
m_hp         = df_f["Marge Hors Promo"].sum()
casse_total  = df_f["Casse (Valeur)"].sum()
tx_marge     = marge_total / ca_total * 100  if ca_total > 0 else 0
tx_m_promo   = m_promo / ca_promo * 100      if ca_promo > 0 else 0
tx_m_hp      = m_hp / ca_hp * 100            if ca_hp > 0    else 0
poids_promo  = ca_promo / ca_total * 100     if ca_total > 0 else 0
delta_hp_p   = tx_m_hp - tx_m_promo
tx_casse     = abs(casse_total) / ca_total * 100 if ca_total > 0 else 0
nb_art_neg   = int((agg_art["TxMarge"] < 0).sum())
nb_flop_neg  = int((flop100["TxMarge"] < 0).sum())

# ─── KPIs GLOBAUX ─────────────────────────────────────────────────────────────
st.markdown(f"<div class='section-label'>{nb_sites_actifs} magasin(s) actifs · {len(sel_rayon)} rayon(s) · {periode}</div>", unsafe_allow_html=True)

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("CA Réseau",         fmt(ca_total),              "FCFA")
k2.metric("Marge Brute",       fmt(marge_total),            "FCFA")
k3.metric("Taux de Marge",     fmt_pct(tx_marge),          f"HP {fmt_pct(tx_m_hp)}")
k4.metric("Effet Promo",       f"−{delta_hp_p:.1f} pts",   f"promo {fmt_pct(tx_m_promo)} vs HP {fmt_pct(tx_m_hp)}")
k5.metric("Poids Promo",       fmt_pct(poids_promo),        fmt(ca_promo) + " FCFA")
k6.metric("Casse Réseau",      fmt_pct(tx_casse, dec=2),   fmt(abs(casse_total)) + " FCFA")

# ─── ALERTES ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Signaux critiques réseau</div>", unsafe_allow_html=True)

# Supeco
supeco_row = agg_fmt[agg_fmt["format"] == "Supeco"]
hyper_row  = agg_fmt[agg_fmt["format"] == "Hyper"]
if not supeco_row.empty and not hyper_row.empty:
    tm_sup = supeco_row["TxMarge"].values[0]
    tm_hyp = hyper_row["TxMarge"].values[0]
    ecart  = tm_hyp - tm_sup
    ca_sup = supeco_row["CA"].values[0]
    if tm_sup < 10:
        st.markdown(f"""
<div class='alert-card alert-purple'>
  <strong>🏪 Format Supeco : taux de marge {tm_sup:.1f}%</strong>
  — écart de {ecart:.1f} pts vs Hypers ({tm_hyp:.1f}%)<br>
  CA concerné : <strong>{fmt(ca_sup)} FCFA</strong>
  · Poids promo Supeco : {supeco_row['PdsPromo'].values[0]:.1f}%<br>
  <span style='font-size:12px;opacity:.85'>→ Mix produit défavorable et sur-pression promotionnelle dans les Supecos à retravailler.</span>
</div>""", unsafe_allow_html=True)

# Articles à marge négative
if nb_art_neg > 0:
    st.markdown(f"""
<div class='alert-card alert-red'>
  <strong>🔴 {nb_art_neg} article(s) à marge négative sur le réseau</strong>
  · {nb_flop_neg} dans le Flop 100<br>
  <span style='font-size:12px;opacity:.85'>→ Chaque vente de ces articles génère une perte nette. Vérification PA / PV / mécanique promo urgente.</span>
</div>""", unsafe_allow_html=True)

# Effet promo
if delta_hp_p > 5:
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ La promotion dégrade la marge de {delta_hp_p:.1f} pts</strong>
  — Hors promo : {fmt_pct(tx_m_hp)} · Sous promo : {fmt_pct(tx_m_promo)}<br>
  <span style='font-size:12px;opacity:.85'>→ Revoir les conditions d'achat promo ou réviser les prix de vente promotionnels.</span>
</div>""", unsafe_allow_html=True)

# Casse élevée
sites_casse = agg_site[agg_site["TxCasse"] > 1].sort_values("TxCasse", ascending=False)
if not sites_casse.empty:
    noms = ", ".join([f"{r['lib_site']} ({r['TxCasse']:.1f}%)" for _, r in sites_casse.iterrows()])
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>🗑️ Taux de casse anormal sur {len(sites_casse)} site(s)</strong> : {noms}<br>
  <span style='font-size:12px;opacity:.85'>→ Vérifier les procédures de démarque. Un taux > 1% du CA indique un problème opérationnel.</span>
</div>""", unsafe_allow_html=True)

# Familles sous seuil
agg_fam = df_f.groupby("lib_fam").agg(CA=("CA","sum"), Marge=("Marge","sum")).reset_index()
agg_fam["TxMarge"] = (agg_fam["Marge"] / agg_fam["CA"] * 100).where(agg_fam["CA"] > 0)
fam_sous_seuil = agg_fam[(agg_fam["TxMarge"] < 8) & (agg_fam["CA"] > 500_000)]
if not fam_sous_seuil.empty:
    noms_fam = ", ".join([f"{r['lib_fam']} ({r['TxMarge']:.1f}%)" for _, r in fam_sous_seuil.iterrows()])
    st.markdown(f"""
<div class='alert-card alert-blue'>
  <strong>📦 {len(fam_sous_seuil)} famille(s) sous 8% de marge</strong> (CA > 500K) : {noms_fam}<br>
  <span style='font-size:12px;opacity:.85'>→ Ces familles absorbent du volume sans générer de marge suffisante pour couvrir les charges.</span>
</div>""", unsafe_allow_html=True)

# ─── TABS PRINCIPAUX ──────────────────────────────────────────────────────────
st.markdown("---")
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Synthèse Réseau",
    "🔢 Matrice Rayon × Magasin",
    f"💣 Flop {min(100, len(flop100))}",
    "📥 Export Excel",
])

# ═══ TAB 1 — SYNTHÈSE RÉSEAU ══════════════════════════════════════════════════
with tab1:

    # Section A : Par format
    st.markdown("<div class='section-label'>Performance par format de magasin</div>", unsafe_allow_html=True)

    fmt_cols = st.columns(len(agg_fmt))
    fmt_colors = {"Hyper": ("#154360","#EFF6FF","#B3D9FF"),
                  "Market":("#145A32","#F0FFF4","#A8E6BF"),
                  "Supeco":("#6E2F8A","#F5F0FF","#D9B3FF")}
    for i, (_, row) in enumerate(agg_fmt.iterrows()):
        fc, bg, border = fmt_colors.get(row["format"], ("#3A3A3C","#F9F9FB","#CCCCCC"))
        with fmt_cols[i]:
            ecart_p = row["TxMarge_HP"] - row["TxMarge_Promo"] if pd.notna(row.get("TxMarge_Promo")) else None
            st.markdown(f"""
<div style='background:{bg};border:1px solid {border};border-radius:12px;padding:16px;margin-bottom:8px'>
  <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:10px'>
    <span style='font-size:15px;font-weight:700;color:{fc}'>{row["format"]}</span>
    <span style='font-size:11px;color:#8E8E93'>{fmt(row["CA"])} FCFA</span>
  </div>
  <div style='font-size:26px;font-weight:700;color:{fc};letter-spacing:-0.02em'>{fmt_pct(row["TxMarge"])}</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:2px'>Taux de marge</div>
  <hr style='margin:10px 0;border-color:{border}'>
  <div style='display:grid;grid-template-columns:1fr 1fr;gap:6px;font-size:12px'>
    <div><span style='color:#8E8E93'>Promo</span><br><strong>{fmt_pct(row.get("TxMarge_Promo"))}</strong></div>
    <div><span style='color:#8E8E93'>Hors promo</span><br><strong>{fmt_pct(row.get("TxMarge_HP"))}</strong></div>
    <div><span style='color:#8E8E93'>Pds promo</span><br><strong>{fmt_pct(row["PdsPromo"])}</strong></div>
    <div><span style='color:#8E8E93'>Casse</span><br><strong>{fmt_pct(row["TxCasse"], dec=2)}</strong></div>
  </div>
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Section B : Récap par rayon
    st.markdown("<div class='section-label'>Récapitulatif par rayon — Fond de rayon vs Promotion</div>", unsafe_allow_html=True)
    disp_rax = agg_rax.copy()
    disp_rax["Rayon"]           = disp_rax["lib_rayon"]
    disp_rax["CA (FCFA)"]       = disp_rax["CA"].apply(fmt)
    disp_rax["Poids CA"]        = disp_rax["PoidsCA"].apply(lambda x: fmt_pct(x))
    disp_rax["Marge (FCFA)"]    = disp_rax["Marge"].apply(fmt)
    disp_rax["Tx Marge"]        = disp_rax["TxMarge"].apply(fmt_pct)
    disp_rax["Tx Marge HP"]     = disp_rax["TxMarge_HP"].apply(fmt_pct)
    disp_rax["Tx Marge Promo"]  = disp_rax["TxMarge_Promo"].apply(fmt_pct)
    disp_rax["Pds Promo"]       = disp_rax["PdsPromo"].apply(fmt_pct)
    disp_rax["Tx Casse"]        = disp_rax["TxCasse"].apply(lambda x: fmt_pct(x, dec=2))
    disp_rax["Écart HP−Promo"]  = (disp_rax["TxMarge_HP"] - disp_rax["TxMarge_Promo"]).apply(
        lambda x: fmt_delta(x) if pd.notna(x) else "—")

    st.dataframe(
        disp_rax[["Rayon","CA (FCFA)","Poids CA","Marge (FCFA)","Tx Marge",
                  "Tx Marge HP","Tx Marge Promo","Écart HP−Promo","Pds Promo","Tx Casse"]],
        use_container_width=True, hide_index=True,
        column_config={
            "Rayon": st.column_config.TextColumn("Rayon", width="medium"),
        }
    )
    st.caption("Écart HP−Promo : différence entre le taux de marge hors promo et sous promotion — mesure l'érosion causée par la mécanique promotionnelle.")

    st.markdown("<br>", unsafe_allow_html=True)

    # Section C : Palmarès magasins
    st.markdown("<div class='section-label'>Palmarès magasins — Classé par taux de marge décroissant</div>", unsafe_allow_html=True)

    disp_site = agg_site.copy()
    disp_site["Rang"]       = range(1, len(disp_site) + 1)
    disp_site["Magasin"]    = disp_site["lib_site"]
    disp_site["Format"]     = disp_site["format"]
    disp_site["CA (FCFA)"]  = disp_site["CA"].apply(fmt)
    disp_site["Marge"]      = disp_site["Marge"].apply(fmt)
    disp_site["Tx Marge"]   = disp_site["TxMarge"].apply(fmt_pct)
    disp_site["Pds Promo"]  = disp_site["PdsPromo"].apply(fmt_pct)
    disp_site["Tx Casse"]   = disp_site["TxCasse"].apply(lambda x: fmt_pct(x, dec=2))
    disp_site["Δ vs Moy."]  = (disp_site["TxMarge"] - moy_marge_site).apply(fmt_delta)

    st.dataframe(
        disp_site[["Rang","Magasin","Format","CA (FCFA)","Marge","Tx Marge","Pds Promo","Tx Casse","Δ vs Moy."]],
        use_container_width=True, hide_index=True,
        column_config={
            "Magasin": st.column_config.TextColumn("Magasin", width="medium"),
            "Format":  st.column_config.TextColumn("Format",  width="small"),
        }
    )

    # Graphique barres horizontales taux de marge par magasin
    try:
        import plotly.graph_objects as go
        s = agg_site.sort_values("TxMarge")
        colors_bar = []
        for _, r in s.iterrows():
            if r["format"] == "Hyper":   colors_bar.append("#154360")
            elif r["format"] == "Market": colors_bar.append("#145A32")
            else:                         colors_bar.append("#6E2F8A")

        fig = go.Figure(go.Bar(
            x=s["TxMarge"].tolist(),
            y=s["lib_site"].tolist(),
            orientation="h",
            marker_color=colors_bar,
            marker_line_width=0,
            text=[f"{v:.1f}%" for v in s["TxMarge"]],
            textposition="outside",
        ))
        fig.add_vline(x=moy_marge_site, line_width=1.5, line_dash="dot", line_color="#FF9500",
                      annotation_text=f" Moy. {moy_marge_site:.1f}%",
                      annotation_font=dict(color="#FF9500", size=10))
        fig.update_layout(
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="-apple-system, Helvetica Neue", color="#3A3A3C", size=11),
            height=max(280, len(agg_site) * 40 + 60),
            margin=dict(t=10, b=10, l=10, r=70),
            xaxis=dict(title="Taux de marge (%)", ticksuffix="%",
                       showgrid=True, gridcolor="#F2F2F7", range=[0, max(s["TxMarge"]) * 1.25]),
            yaxis=dict(showgrid=False, title=""),
        )
        st.plotly_chart(fig, use_container_width=True)
        st.caption("🔵 Hyper  ·  🟢 Market  ·  🟣 Supeco  ·  Ligne pointillée = moyenne réseau")
    except ImportError:
        pass

# ═══ TAB 2 — MATRICE RAYON × MAGASIN ══════════════════════════════════════════
with tab2:
    st.markdown("<div class='section-label'>Taux de marge (%) par combinaison Rayon × Magasin</div>", unsafe_allow_html=True)
    st.caption("Lecture : chaque cellule = taux de marge brute pour ce rayon dans ce magasin · Une cellule vide = aucune vente ce jour-là")

    if not mat_pivot.empty:
        # Afficher avec formatage %
        mat_display = mat_pivot.copy()
        for col in mat_display.columns:
            mat_display[col] = mat_display[col].apply(
                lambda x: f"{x:.1f}%" if pd.notna(x) else "—"
            )
        st.dataframe(mat_display, use_container_width=True)

        # Heatmap plotly
        try:
            import plotly.graph_objects as go
            import plotly.express as px

            sites_ordered = mat_pivot.columns.tolist()
            rayons_ordered = mat_pivot.index.tolist()
            z = mat_pivot.values.tolist()
            text_z = [[f"{v:.1f}%" if pd.notna(v) else "—" for v in row] for row in z]

            fig_h = go.Figure(go.Heatmap(
                z=z,
                x=sites_ordered,
                y=rayons_ordered,
                text=text_z,
                texttemplate="%{text}",
                textfont=dict(size=12, family="-apple-system, Helvetica Neue"),
                colorscale=[
                    [0.0,  "#C0392B"],
                    [0.25, "#E74C3C"],
                    [0.45, "#F39C12"],
                    [0.60, "#F0E68C"],
                    [0.75, "#A8D5A2"],
                    [1.0,  "#27AE60"],
                ],
                showscale=True,
                colorbar=dict(title="Tx Marge %", ticksuffix="%", len=0.8),
                zmin=0, zmax=30,
            ))
            fig_h.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                font=dict(family="-apple-system, Helvetica Neue", color="#3A3A3C", size=11),
                height=max(250, len(rayons_ordered) * 80 + 80),
                margin=dict(t=20, b=60, l=20, r=20),
                xaxis=dict(tickangle=-35, tickfont=dict(size=10)),
                yaxis=dict(tickfont=dict(size=11)),
            )
            st.plotly_chart(fig_h, use_container_width=True)
            st.caption("🔴 < 10%  ·  🟠 10–15%  ·  🟡 15–20%  ·  🟢 20–25%  ·  ✅ > 25%")
        except ImportError:
            pass
    else:
        st.info("Pas de données suffisantes pour construire la matrice.")

# ═══ TAB 3 — FLOP 100 ═════════════════════════════════════════════════════════
with tab3:
    st.markdown(f"""
<div class='alert-card alert-red'>
  <strong>💣 {nb_flop_neg} article(s) à marge négative dans le Flop {len(flop100)}</strong>
  · Pertes cumulées : <strong>{fmt(flop100[flop100['Marge']<0]['Marge'].sum())} FCFA</strong>
  · CA concerné : <strong>{fmt(flop100['CA'].sum())} FCFA</strong><br>
  <span style='font-size:12px;opacity:.85'>
    Chaque bloc de magasins est trié du pire au meilleur taux de marge.
    Un site absent d'un bloc = article non vendu dans ce format sur la période.
  </span>
</div>""", unsafe_allow_html=True)

    # Filtres inline
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        filtre_rayon_f = st.selectbox("Rayon", ["Tous"] + sorted(flop100["lib_rayon"].dropna().unique().tolist()), key="f100_rayon")
    with fc2:
        filtre_marge = st.selectbox("Statut marge", ["Tous", "Négatif uniquement", "< 3%", "< 8%"], key="f100_marge")
    with fc3:
        filtre_promo = st.selectbox("Promo", ["Tous", "100% sous promo", "Hors promo uniquement"], key="f100_promo")

    df_flop = flop100.copy()
    if filtre_rayon_f != "Tous":
        df_flop = df_flop[df_flop["lib_rayon"] == filtre_rayon_f]
    if filtre_marge == "Négatif uniquement":
        df_flop = df_flop[df_flop["TxMarge"] < 0]
    elif filtre_marge == "< 3%":
        df_flop = df_flop[df_flop["TxMarge"] < 3]
    elif filtre_marge == "< 8%":
        df_flop = df_flop[df_flop["TxMarge"] < 8]
    if filtre_promo == "100% sous promo":
        df_flop = df_flop[df_flop["PdsPromo"] >= 99.9]
    elif filtre_promo == "Hors promo uniquement":
        df_flop = df_flop[df_flop["PdsPromo"].fillna(0) < 1]

    st.markdown(f"<div style='font-size:12px;color:#8E8E93;margin-bottom:8px'>{len(df_flop)} article(s) affichés</div>", unsafe_allow_html=True)

    # Préparer affichage
    disp_flop = df_flop.copy()
    disp_flop["#"]              = disp_flop["Rang"]
    disp_flop["Article"]        = disp_flop["lib_art"]
    disp_flop["Rayon"]          = disp_flop["lib_rayon"]
    disp_flop["Famille"]        = disp_flop["lib_fam"]
    disp_flop["CA (FCFA)"]      = disp_flop["CA"].apply(fmt)
    disp_flop["Marge (FCFA)"]   = disp_flop["Marge"].apply(fmt)
    disp_flop["Tx Marge"]       = disp_flop["TxMarge"].apply(fmt_pct)
    disp_flop["Pds Promo"]      = disp_flop["PdsPromo"].apply(lambda x: fmt_pct(x) if pd.notna(x) else "—")
    disp_flop["Qté"]            = disp_flop["Qte"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "—")
    disp_flop["🔵 HYPER"]       = disp_flop["Bloc_Hyper"]
    disp_flop["🟢 MARKET"]      = disp_flop["Bloc_Market"]
    disp_flop["🟣 SUPECO"]      = disp_flop["Bloc_Supeco"]

    st.dataframe(
        disp_flop[["#","Article","Rayon","Famille","CA (FCFA)","Marge (FCFA)",
                   "Tx Marge","Pds Promo","Qté","🔵 HYPER","🟢 MARKET","🟣 SUPECO"]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "#":          st.column_config.NumberColumn("#",        width=40),
            "Article":    st.column_config.TextColumn("Article",    width="large"),
            "Rayon":      st.column_config.TextColumn("Rayon",      width="medium"),
            "Famille":    st.column_config.TextColumn("Famille",    width="medium"),
            "🔵 HYPER":   st.column_config.TextColumn("🔵 HYPER",   width="large"),
            "🟢 MARKET":  st.column_config.TextColumn("🟢 MARKET",  width="large"),
            "🟣 SUPECO":  st.column_config.TextColumn("🟣 SUPECO",  width="large"),
        }
    )
    st.caption("Blocs magasins : triés du taux de marge le plus bas au plus élevé · '—' = article non vendu dans ce format · Format : Magasin: Tx%")

# ═══ TAB 4 — EXPORT EXCEL ═════════════════════════════════════════════════════
with tab4:
    st.markdown("<div class='section-label'>Export Excel — Rapport complet</div>", unsafe_allow_html=True)
    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>📋 Contenu du fichier exporté :</strong><br>
  <strong>Onglet 1 — Synthèse Réseau</strong> : KPIs globaux, résumé par format, palmarès magasins<br>
  <strong>Onglet 2 — Récap par Rayon</strong> : Taux de marge, décomposition HP vs Promo, casse<br>
  <strong>Onglet 3 — Matrice Marge</strong> : Taux de marge croisé Rayon × Magasin<br>
  <strong>Onglet 4 — Flop 100</strong> : Articles destructeurs avec impact par bloc (Hyper / Market / Supeco)
</div>""", unsafe_allow_html=True)

    st.caption(f"Périmètre : {len(sel_site)} magasin(s) · {len(sel_rayon)} rayon(s) · {periode}")

    if st.button("Générer le fichier Excel", type="primary", key="gen_excel"):
        with st.spinner("Génération du rapport…"):

            wb_exp = Workbook()

            # Styles communs
            C_HDR = "1B2A4A"; C_SUB = "2E4B7A"; C_WH = "FFFFFF"; C_DK = "1A1A2E"
            C_HYP = "154360"; C_MKT = "145A32"; C_SUP = "6E2F8A"

            def xfill(h): return PatternFill("solid", fgColor=h)
            def xbdr():
                s = Side(style="thin", color="CCCCCC")
                return Border(left=s, right=s, top=s, bottom=s)
            def xctr(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
            def xrgt(): return Alignment(horizontal="right",  vertical="center")
            def xlft(w=False): return Alignment(horizontal="left", vertical="center", wrap_text=w)

            def write_header_row(ws, row_num, headers, widths, bg=C_SUB):
                for i, (h, w) in enumerate(zip(headers, widths)):
                    c = ws.cell(row=row_num, column=i+1, value=h)
                    c.font      = Font("Calibri", size=10, bold=True, color=C_WH)
                    c.fill      = xfill(bg)
                    c.alignment = xctr()
                    c.border    = xbdr()
                    ws.column_dimensions[get_column_letter(i+1)].width = w
                ws.row_dimensions[row_num].height = 24

            def title_block(ws, txt, span=10):
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
                c = ws.cell(row=1, column=1, value=txt)
                c.font = Font("Calibri", size=13, bold=True, color=C_WH)
                c.fill = xfill(C_HDR); c.alignment = xctr()
                ws.row_dimensions[1].height = 30
                ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
                c2 = ws.cell(row=2, column=1, value=f"  Période : {periode}  ·  {nb_jours} jour(s)")
                c2.font = Font("Calibri", size=9, italic=True, color="AABBCC")
                c2.fill = xfill(C_HDR); c2.alignment = xlft()
                ws.row_dimensions[2].height = 16
                ws.row_dimensions[3].height = 6

            # ── Onglet 1 : Synthèse réseau ────────────────────────────────────
            ws1 = wb_exp.active; ws1.title = "Synthèse Réseau"
            title_block(ws1, "DIAGNOSTIC RENTABILITÉ RÉSEAU — SYNTHÈSE", span=8)

            # KPIs
            r = 4
            kpi_data = [
                ("CA Réseau (FCFA)",    f"{ca_total:,.0f}"),
                ("Marge Brute (FCFA)",  f"{marge_total:,.0f}"),
                ("Taux de Marge",       f"{tx_marge:.1f}%"),
                ("Taux Marge HP",       f"{tx_m_hp:.1f}%"),
                ("Taux Marge Promo",    f"{tx_m_promo:.1f}%"),
                ("Effet Promo (pts)",   f"−{delta_hp_p:.1f}"),
                ("Poids Promo",         f"{poids_promo:.1f}%"),
                ("Taux Casse",          f"{tx_casse:.2f}%"),
            ]
            for ci, (lbl, val) in enumerate(kpi_data):
                c1e = ws1.cell(row=r, column=ci+1, value=lbl)
                c1e.font = Font("Calibri", size=9, bold=True, color=C_WH)
                c1e.fill = xfill(C_SUB); c1e.alignment = xctr(); c1e.border = xbdr()
                c2e = ws1.cell(row=r+1, column=ci+1, value=val)
                c2e.font = Font("Calibri", size=12, bold=True, color=C_DK)
                c2e.fill = xfill("FFFFFF"); c2e.alignment = xctr(); c2e.border = xbdr()
                ws1.column_dimensions[get_column_letter(ci+1)].width = 18
            ws1.row_dimensions[r].height = 20; ws1.row_dimensions[r+1].height = 28
            r += 3

            # Par format
            ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
            c = ws1.cell(row=r, column=1, value="  PERFORMANCE PAR FORMAT")
            c.font = Font("Calibri", size=10, bold=True, color=C_WH)
            c.fill = xfill(C_SUB); c.alignment = xlft(); ws1.row_dimensions[r].height = 22; r += 1

            write_header_row(ws1, r,
                ["Format","CA (FCFA)","Marge (FCFA)","Tx Marge","Tx Marge HP","Tx Marge Promo","Pds Promo","Tx Casse"],
                [14,18,18,12,14,16,12,12])
            r += 1
            for ri2, (_, fd) in enumerate(agg_fmt.iterrows()):
                bg_f = {"Hyper":"D6EAF8","Market":"D5F5E3","Supeco":"E8DAEF"}.get(fd["format"],"FFFFFF")
                for ci2, (v, fmt3) in enumerate([
                    (fd["format"], None), (fd["CA"], "#,##0"), (fd["Marge"], "#,##0"),
                    (fd["TxMarge"]/100, "0.0%"), (fd.get("TxMarge_HP",None), "0.0%"),
                    (fd.get("TxMarge_Promo",None), "0.0%"),
                    (fd["PdsPromo"]/100 if pd.notna(fd["PdsPromo"]) else None, "0.0%"),
                    (fd["TxCasse"]/100 if pd.notna(fd["TxCasse"]) else None, "0.0%"),
                ]):
                    if pd.notna(v) if not isinstance(v, str) else True:
                        pv = v/100 if fmt3 == "0.0%" and isinstance(v, float) and not (0 <= abs(v) <= 1.5) else v
                    else:
                        pv = None
                    c = ws1.cell(row=r, column=ci2+1, value=pv)
                    c.font = Font("Calibri", size=10, color=C_DK)
                    c.fill = xfill(bg_f); c.border = xbdr()
                    if fmt3: c.number_format = fmt3
                    c.alignment = xrgt() if ci2 in [1,2] else xctr()
                ws1.row_dimensions[r].height = 20; r += 1

            # Palmarès magasins
            r += 1
            ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
            c = ws1.cell(row=r, column=1, value="  PALMARÈS MAGASINS — classé par taux de marge décroissant")
            c.font = Font("Calibri", size=10, bold=True, color=C_WH)
            c.fill = xfill(C_SUB); c.alignment = xlft(); ws1.row_dimensions[r].height = 22; r += 1

            write_header_row(ws1, r, ["Rang","Magasin","Format","CA (FCFA)","Marge (FCFA)","Tx Marge","Pds Promo","Tx Casse"],
                             [5,24,10,18,18,12,12,12])
            r += 1
            for ri3, (_, sd) in enumerate(agg_site.iterrows()):
                bg_s = "F7F7F7" if ri3 % 2 == 0 else "FFFFFF"
                vals3 = [ri3+1, sd["lib_site"], sd["format"], sd["CA"], sd["Marge"],
                         sd["TxMarge"]/100 if pd.notna(sd["TxMarge"]) else None,
                         sd["PdsPromo"]/100 if pd.notna(sd["PdsPromo"]) else None,
                         sd["TxCasse"]/100 if pd.notna(sd["TxCasse"]) else None]
                fmts3 = [None,None,None,"#,##0","#,##0","0.0%","0.0%","0.0%"]
                for ci3,(v,f3) in enumerate(zip(vals3,fmts3)):
                    c = ws1.cell(row=r, column=ci3+1, value=v)
                    c.font = Font("Calibri", size=10, color=C_DK)
                    c.fill = xfill(bg_s); c.border = xbdr()
                    if f3: c.number_format = f3
                    c.alignment = xctr() if ci3==0 else xrgt() if ci3 in [3,4] else xctr()
                ws1.row_dimensions[r].height = 20; r += 1

            ws1.freeze_panes = "A4"

            # ── Onglet 2 : Récap rayon ────────────────────────────────────────
            ws2 = wb_exp.create_sheet("Récap Rayon")
            title_block(ws2, "RÉCAPITULATIF PAR RAYON — Fond de Rayon vs Promotion", span=10)
            write_header_row(ws2, 4,
                ["Rayon","CA (FCFA)","Poids CA","Marge (FCFA)","Tx Marge","Tx Marge HP","Tx Marge Promo","Écart HP−Promo","Pds Promo","Tx Casse"],
                [22,16,10,16,12,14,16,16,12,12])
            for ri4,(_, rd4) in enumerate(agg_rax.iterrows()):
                r4 = ri4 + 5
                bg4 = "F7F7F7" if ri4%2==0 else "FFFFFF"
                ecart4 = rd4["TxMarge_HP"] - rd4["TxMarge_Promo"] if pd.notna(rd4.get("TxMarge_Promo")) else None
                vals4 = [rd4["lib_rayon"], rd4["CA"], rd4["PoidsCA"]/100,
                         rd4["Marge"], rd4["TxMarge"]/100,
                         rd4["TxMarge_HP"]/100 if pd.notna(rd4.get("TxMarge_HP")) else None,
                         rd4["TxMarge_Promo"]/100 if pd.notna(rd4.get("TxMarge_Promo")) else None,
                         ecart4/100 if ecart4 is not None else None,
                         rd4["PdsPromo"]/100, rd4["TxCasse"]/100]
                fmts4 = [None,"#,##0","0.0%","#,##0","0.0%","0.0%","0.0%","0.0%","0.0%","0.0%"]
                for ci4,(v,f4) in enumerate(zip(vals4,fmts4)):
                    c = ws2.cell(row=r4, column=ci4+1, value=v)
                    c.font = Font("Calibri",size=10,color=C_DK); c.fill=xfill(bg4); c.border=xbdr()
                    if f4: c.number_format=f4
                    c.alignment = xrgt() if ci4 in [1,3] else xctr()
                ws2.row_dimensions[r4].height = 20
            ws2.freeze_panes = "A5"

            # ── Onglet 3 : Matrice ────────────────────────────────────────────
            ws3 = wb_exp.create_sheet("Matrice Marge")
            title_block(ws3, "MATRICE TAUX DE MARGE — RAYON × MAGASIN", span=len(mat_pivot.columns)+2)
            ws3.cell(row=4, column=1, value="Rayon").font = Font("Calibri", size=10, bold=True, color=C_WH)
            ws3.cell(row=4, column=1).fill   = xfill(C_SUB)
            ws3.cell(row=4, column=1).alignment = xctr()
            ws3.cell(row=4, column=1).border    = xbdr()
            ws3.column_dimensions["A"].width = 22
            for ci5, col_name in enumerate(mat_pivot.columns):
                c = ws3.cell(row=4, column=ci5+2, value=short_name(col_name))
                c.font = Font("Calibri", size=9, bold=True, color=C_WH)
                c.fill = xfill(C_SUB); c.alignment = xctr(); c.border = xbdr()
                ws3.column_dimensions[get_column_letter(ci5+2)].width = 16
            ws3.row_dimensions[4].height = 36
            for ri5, rayon_idx in enumerate(mat_pivot.index):
                r5 = ri5 + 5
                c0 = ws3.cell(row=r5, column=1, value=rayon_idx)
                c0.font = Font("Calibri", size=10, bold=True, color=C_DK)
                c0.fill = xfill("F7F7F7"); c0.alignment = xlft(); c0.border = xbdr()
                for ci5, col_name in enumerate(mat_pivot.columns):
                    v5 = mat_pivot.loc[rayon_idx, col_name]
                    c = ws3.cell(row=r5, column=ci5+2)
                    if pd.notna(v5):
                        c.value = v5/100; c.number_format = "0.0%"
                    c.font = Font("Calibri", size=11, bold=True, color=C_DK)
                    c.alignment = xctr(); c.border = xbdr()
                    c.fill = xfill("FFFFFF")
                ws3.row_dimensions[r5].height = 28
            ws3.freeze_panes = "B5"

            # ── Onglet 4 : Flop 100 ──────────────────────────────────────────
            ws4 = wb_exp.create_sheet("Flop 100")
            title_block(ws4, f"FLOP {len(flop100)} — DESTRUCTEURS DE MARGE · Taux de marge par bloc magasin", span=12)
            ws4.merge_cells("A3:L3")
            c = ws4.cell(row=3, column=1,
                value=f"  {nb_flop_neg} articles à marge négative · Pertes : {flop100[flop100['Marge']<0]['Marge'].sum():,.0f} FCFA · Blocs triés du pire au meilleur taux")
            c.font = Font("Calibri", size=9, italic=True, color="AABBCC")
            c.fill = xfill(C_HDR); c.alignment = xlft()
            ws4.row_dimensions[3].height = 14

            hdrs4  = ["#","Article","Rayon","Famille","CA (FCFA)","Marge (FCFA)","Tx Marge","Pds Promo","Qté",
                      "🔵 HYPER — Tx marge % par site",
                      "🟢 MARKET — Tx marge % par site",
                      "🟣 SUPECO — Tx marge % par site"]
            wdths4 = [5, 44, 16, 24, 13, 13, 10, 10, 8, 42, 46, 50]

            # En-têtes avec couleurs blocs
            bloc_bg = {9: C_HYP, 10: C_MKT, 11: C_SUP}
            for ci6, (h, w) in enumerate(zip(hdrs4, wdths4)):
                bg6 = bloc_bg.get(ci6, C_SUB)
                c = ws4.cell(row=4, column=ci6+1, value=h)
                c.font = Font("Calibri", size=9, bold=True, color=C_WH)
                c.fill = xfill(bg6); c.alignment = xctr(); c.border = xbdr()
                ws4.column_dimensions[get_column_letter(ci6+1)].width = w
            ws4.row_dimensions[4].height = 28

            bloc_fill = {9: "D6EAF8", 10: "D5F5E3", 11: "E8DAEF"}
            for ri6, (_, rd6) in enumerate(flop100.iterrows()):
                r6 = ri6 + 5
                bg6 = "F7F7F7" if ri6 % 2 == 0 else "FFFFFF"
                tm6 = rd6["TxMarge"]; pp6 = rd6["PdsPromo"]
                vals6 = [
                    rd6["Rang"], rd6["lib_art"], rd6["lib_rayon"], rd6["lib_fam"],
                    rd6["CA"], rd6["Marge"],
                    tm6/100 if pd.notna(tm6) else None,
                    pp6/100 if pd.notna(pp6) else None,
                    int(rd6["Qte"]) if pd.notna(rd6["Qte"]) else None,
                    rd6["Bloc_Hyper"], rd6["Bloc_Market"], rd6["Bloc_Supeco"],
                ]
                fmts6 = [None,None,None,None,"#,##0","#,##0","0.0%","0.0%","#,##0",None,None,None]
                for ci6,(v,f6) in enumerate(zip(vals6,fmts6)):
                    col6 = ci6
                    c = ws4.cell(row=r6, column=ci6+1, value=v)
                    cell_bg = bloc_fill.get(col6, bg6) if v and str(v) != "—" else bg6
                    c.font  = Font("Calibri", size=10 if ci6<9 else 9, color=C_DK)
                    c.fill  = xfill(cell_bg); c.border = xbdr()
                    if f6: c.number_format = f6
                    if ci6 == 0:   c.font = Font("Calibri", size=10, bold=True, color=C_DK); c.alignment = xctr()
                    elif ci6 in [4,5]:  c.alignment = xrgt()
                    elif ci6 in [6,7,8]: c.alignment = xctr()
                    elif ci6 >= 9: c.alignment = xlft(w=True)
                    else: c.alignment = xlft(w=(ci6 in [1,3]))
                ws4.row_dimensions[r6].height = 30

            ws4.freeze_panes = "A5"

            buf = BytesIO()
            wb_exp.save(buf)
            buf.seek(0)

        st.download_button(
            label="⬇️ Télécharger le rapport Excel",
            data=buf,
            file_name=f"SmartBuyer_Diagnostic_Reseau_{periode.replace('/','').replace(' ','_').replace('→','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
