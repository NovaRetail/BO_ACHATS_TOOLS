"""
07_💸_Fidelite_Cagnotte.py — SmartBuyer Hub
Suivi Hebdomadaire & Performance Terrain · Investissement via Cagnottage
"""

import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Fidélité Cagnotte · SmartBuyer",
    page_icon="💸",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CHARTE SMARTBUYER (INJECTION CSS EXACTE MARGES NEGATIVES) ────────────────
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

/* Style des KPI Cards comme dans Marges Négatives */
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
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }

.col-required { background: #F0F8FF; border: 0.5px solid #B3D9FF; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; display: flex; align-items: flex-start; gap: 10px; }
.col-name { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
</style>
""", unsafe_allow_html=True)

# ─── HELPERS FORMATTAGE ───────────────────────────────────────────────────────
def fmt(n):
    if pd.isna(n) or n is None: return "—"
    a = abs(n)
    if a >= 1_000_000: return f"{n/1_000_000:.2f} M"
    if a >= 1_000:     return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def fmt_pct(v, dec=1):
    if pd.isna(v) or v is None: return "—"
    return f"{v:.{dec}f}%"

def short_name(s):
    s = str(s)
    return s.split(" - ", 1)[-1].strip() if " - " in s else s

def normaliser_mois(s):
    if not isinstance(s, str): return ""
    dico = {"janvier":"Janvier","fevrier":"Fevrier","mars":"Mars","avril":"Avril","mai":"Mai","juin":"Juin","juillet":"Juillet","aout":"Aout","septembre":"Septembre","octobre":"Octobre","novembre":"Novembre","decembre":"Decembre"}
    s_clean = (s.strip().lower().replace("é","e").replace("è","e").replace("ê","e").replace("û","u").replace("à","a").replace("ç","c"))
    return dico.get(s_clean, s.strip().capitalize())

def normaliser_rayon(s):
    if not isinstance(s, str): return ""
    return (s.strip().replace("é","e").replace("è","e").replace("ê","e").replace("û","u").replace("à","a").replace("ç","c").upper())

# ─── PARSING STRUCTURE FLAT SÉCURISÉE ─────────────────────────────────────────
def extract_periode(df):
    if "Site nom long" not in df.columns: return "Période inconnue"
    pattern = re.compile(r"apr[eè]s\s+le\s+(\d{2}/\d{2}/\d{4})\s+et\s+est\s+avant\s+le\s+(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
    for val in df["Site nom long"].dropna():
        if "Filtres" in str(val):
            m = pattern.search(str(val))
            if m:
                d_fin = (pd.Timestamp(datetime.strptime(m.group(2), "%d/%m/%Y").date()) - pd.Timedelta(days=1)).date()
                mois_fr = {1:"Janvier", 2:"Fevrier", 3:"Mars", 4:"Avril", 5:"Mai", 6:"Juin", 7:"Juillet", 8:"Aout", 9:"Septembre", 10:"Octobre", 11:"Novembre", 12:"Decembre"}
                return {"semaine": f"S{d_fin.isocalendar().week:02d}", "mois_court": mois_fr[d_fin.month], "mois_long": f"{mois_fr[d_fin.month]} {d_fin.year}"}
    return "Période inconnue"

def parser_pbi(fichier) -> dict:
    try:
        df = pd.read_excel(fichier, header=0, dtype=str)
    except: return None

    per = extract_periode(df)
    if isinstance(per, str): return None

    to_f = lambda v: pd.to_numeric(str(v).replace(",", ".").strip(), errors='coerce') if pd.notnull(v) else 0.0

    c_ca = next((c for c in df.columns if "CA" in str(c)), df.columns[4])
    c_mg = next((c for c in df.columns if "Marge" in str(c)), df.columns[7])
    c_qt = next((c for c in df.columns if "Qté" in str(c) or "Qte" in str(c)), df.columns[27])

    df["_site"] = df["Site nom long"].fillna("").astype(str).str.strip()
    df["_rayon"] = df["Rayon"].fillna("").astype(str).str.strip()
    df["_article"] = df["Article"].fillna("").astype(str).str.strip()

    pat_site = re.compile(r"^\d{4,6}\s*-\s*.+")
    mask_art = (df["_article"].str.match(re.compile(r"^\d{7,9}\s*-\s*.+")) & df["_site"].str.match(pat_site))
    df_art = df[mask_art].copy()
    
    if df_art.empty: return None

    df_art["Code Article"] = df_art["_article"].apply(lambda s: s.split("-")[0].strip())
    df_art["Article_Label"] = df_art["_article"].apply(lambda s: s.split("-", 1)[1].strip() if "-" in s else s)
    df_art["Magasin"] = df_art["_site"].apply(lambda s: short_name(s))
    df_art["Rayon"] = df_art["_rayon"].apply(lambda s: short_name(s))
    df_art["Rayon_Norm"] = df_art["_rayon"].apply(normaliser_rayon)
    df_art["Semaine"] = per["semaine"]
    df_art["Mois"] = per["mois_long"]
    df_art["CA"] = df_art[c_ca].apply(to_f)
    df_art["Marge"] = df_art[c_mg].apply(to_f)
    df_art["Qte"] = df_art[c_qt].apply(to_f)

    mask_ray = (df["_site"].str.match(pat_site) & (df["_rayon"] != "") & (df["Famille"].fillna("").str.strip() == "Total"))
    totaux_rayon = {}
    for _, r in df[mask_ray].iterrows():
        rn = normaliser_rayon(r["_rayon"])
        totaux_rayon[rn] = totaux_rayon.get(rn, 0.0) + to_f(r[c_ca])

    return {"periode": per, "lignes": df_art, "totaux_rayon": totaux_rayon}

# ─── SIDEBAR NAVIGATION & IMPORT ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:18px'>
  <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>🛍️ SmartBuyer</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · Équipe Achats</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("<div class='section-label'>Navigation</div>", unsafe_allow_html=True)
    st.page_link("app.py",                                       label="🏠  Accueil")
    st.page_link("pages/06_💸_Marges_Negatives.py",              label="💸  Marges Négatives")
    st.page_link("pages/07_💸_Fidelite_Cagnotte.py",              label="💸  Fidélité Cagnotte")
    st.markdown("---")

    st.markdown("<div class='section-label'>Import fichiers</div>", unsafe_allow_html=True)
    fichiers_pbi = st.file_uploader("Extractions PBI (xlsx, plusieurs possibles)", type=["xlsx"], accept_multiple_files=True, key="cagnotte_pbi")
    fichier_liste = st.file_uploader("Liste articles fidélité (csv)", type=["csv"], key="cagnotte_liste")

# ─── HEADER PAGE ──────────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>🎯 Ciblage & Investissement Fidélité</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Suivi de la performance magasin par semaine · Confrontation avec les adhésions · Arbitrage budgétaire et contrôle de l'enveloppe de cagnottage</div>", unsafe_allow_html=True)

# ─── ÉCRAN D'ACCUEIL ──────────────────────────────────────────────────────────
if not fichiers_pbi or not fichier_liste:
    st.markdown("---")
    st.markdown("""
<div class='alert-card alert-blue'>
    <strong>ℹ️ À quoi sert ce module ?</strong><br>
    Ce module permet de piloter la rentabilité et le déploiement du programme de fidélité réseau sur deux axes stratégiques :
    <br><br>
    <strong>1. Suivi Terrain (Magasin & Semaine)</strong> — Analyse croisée par point de vente et par semaine pour cibler les baisses de régime d'adhésion ou de chiffre d'affaires fidélité.<br>
    <strong>2. Pilotage Financier (Investissement)</strong> — Isolation de l'enveloppe financière distribuée via le cagnottage et calcul du ROI face à la marge commerciale générée.
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Structure des fichiers attendus</div>", unsafe_allow_html=True)
    st.markdown("""
<div class='col-required'><div style='font-size:16px'>📊</div>
<div><div class='col-name'>Extractions PBI Ventes Hebdomadaires (.xlsx)</div>
<div class='col-desc'>Fichiers plats Power BI par semaine contenant les ventes magasins avec les colonnes CA, Marge, et Qté Vente.</div>
</div></div>
<div class='col-required'><div style='font-size:16px'>📄</div>
<div><div class='col-name'>Référentiel des Offres Nationales (.csv)</div>
<div class='col-desc'>Séparateur point-virgule ou virgule · Colonnes obligatoires : <span style='font-family:monospace'>Article</span> (Code) ; <span style='font-family:monospace'>Cagnotte</span> (Montant unitaire) ; <span style='font-family:monospace'>Mois</span>.</div>
</div></div>""", unsafe_allow_html=True)
    st.info("⬆️ Chargez vos extractions PBI et la liste des offres dans la sidebar pour lancer les analyses.")
    st.stop()

# ─── CHARGEMENT & TRAITEMENT DES DONNÉES ──────────────────────────────────────
with st.spinner("Analyse et consolidation des fichiers en cours…"):
    try:
        ref_df = pd.read_csv(fichier_liste, sep=";", engine='python', dtype={"Article": str})
    except:
        try:
            fichier_liste.seek(0)
            ref_df = pd.read_csv(fichier_liste, sep=",", engine='python', dtype={"Article": str})
        except:
            st.error("Impossible de lire le fichier référentiel CSV. Vérifiez sa structure.")
            st.stop()

    ref_df["Article"] = ref_df["Article"].astype(str).str.strip()
    ref_df["Mois_norm"] = ref_df["Mois"].apply(normaliser_mois)

    all_rows = []
    global_rayons_ca = {}
    semaines_traitees = set()
    label_periode = ""

    for f in fichiers_pbi:
        data = parser_pbi(f)
        if not data: continue
        
        sem = data["periode"]["semaine"]
        semaines_traitees.add(sem)
        label_periode = data["periode"]["mois_long"]

        for k, v in data["totaux_rayon"].items():
            global_rayons_ca[k] = global_rayons_ca.get(k, 0.0) + v

        m_court = data["periode"]["mois_court"]
        valid_rows = ref_df[ref_df["Mois_norm"] == m_court]
        cagnotte_map = dict(zip(valid_rows["Article"], valid_rows["Cagnotte"]))
        
        df_f = data["lignes"][data["lignes"]["Code Article"].isin(cagnotte_map.keys())].copy()
        if not df_f.empty:
            df_f["Cagnotte Unitaire"] = df_f["Code Article"].map(cagnotte_map)
            all_rows.append(df_f)

if not all_rows:
    st.warning("⚠️ Aucun article de la liste d'offres n'
