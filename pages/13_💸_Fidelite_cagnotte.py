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
[data-testid="stMetricLabel"] { font-size: 11px !important; font-weight: 500 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; margin-bottom: 4px !important; }
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
    st.warning("⚠️ Aucun article de la liste d'offres n'a été identifié dans vos extractions Power BI.")
    st.stop()

df_base = pd.concat(all_rows, ignore_index=True)

# ─── CONSOLIDATION DES ENSEMBLES DÉCISIONNELS ─────────────────────────────────
df_magasin = df_base.groupby(["Semaine", "Mois", "Magasin", "Rayon", "Code Article", "Article_Label"]).agg(
    CA_Fid=("CA", "sum"), Marge_Fid=("Marge", "sum"), Qte_Fid=("Qte", "sum")
).reset_index().rename(columns={"Article_Label": "Article", "CA_Fid": "CA Fidélité", "Marge_Fid": "Marge Fidélité", "Qte_Fid": "Qté Vendue"})
df_magasin = df_magasin.sort_values(by=["Semaine", "Magasin", "CA Fidélité"], ascending=[True, True, False])

df_finance = df_base.groupby(["Semaine", "Rayon", "Rayon_Norm", "Code Article", "Article_Label"]).agg(
    Cagnotte_U=("Cagnotte Unitaire", "first"), Qte_Tot=("Qte", "sum"), CA_Tot=("CA", "sum"), Marge_Tot=("Marge", "sum")
).reset_index().rename(columns={"Article_Label": "Article", "Cagnotte_U": "Cagnotte Unitaire", "Qte_Tot": "Qté Totale", "CA_Tot": "CA Généré", "Marge_Tot": "Marge Générée"})

df_finance["Investissement Cagnottage"] = df_finance["Cagnotte Unitaire"] * df_finance["Qté Totale"]
df_finance["ROI Écart"] = df_finance.apply(lambda r: r["CA Généré"] / r["Investissement Cagnottage"] if r["Investissement Cagnottage"] > 0 else 0, axis=1)

df_poids_rayon = df_finance.groupby(["Rayon", "Rayon_Norm"]).agg({"CA Généré": "sum", "Marge Générée": "sum"}).reset_index()
df_poids_rayon["CA Rayon Réseau"] = df_poids_rayon["Rayon_Norm"].map(global_rayons_ca).fillna(0.0)
df_poids_rayon["Poids CA %"] = df_poids_rayon.apply(lambda r: (r["CA Généré"] / r["CA Rayon Réseau"] * 100) if r["CA Rayon Réseau"] > 0 else 0, axis=1)
df_poids_rayon = df_poids_rayon.sort_values("Poids CA %", ascending=True)

invest_total = df_finance["Investissement Cagnottage"].sum()
ca_total = df_finance["CA Généré"].sum()
marge_totale = df_finance["Marge Générée"].sum()
nb_magasins_actifs = len(df_magasin["Magasin"].unique())

# ─── AFFICHAGE DES KPI CARDS (ALIGNEMENT EXACT CHARTE MARGES NEGATIVES) ───────
st.markdown(f"<div class='section-label'>{nb_magasins_actifs} magasin(s) suivis · {len(semaines_traitees)} semaine(s) · {label_periode}</div>", unsafe_allow_html=True)

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Investissement Total", fmt(invest_total), "Cagnottage FCFA")
k2.metric("CA Fidélité Généré",   fmt(ca_total),     "FCFA Réseau")
k3.metric("Marge Commerciale",   fmt(marge_totale), "FCFA Brute")
k4.metric("Multiplicateur CA",    f"{ca_total / invest_total:.1f}x" if invest_total > 0 else "—", "Effet de Levier")
k5.metric("Poids Invest / CA",    fmt_pct((invest_total / ca_total * 100) if ca_total > 0 else 0), "Taux d'Effort")
k6.metric("Articles Cibles",       f"{len(df_finance['Code Article'].unique())}", "Offres Actives")

# ─── CRITIQUES ET ALERTES RÉSEAU ──────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Alertes et dérives budgétaires</div>", unsafe_allow_html=True)

flop_roi = df_finance[df_finance["ROI Écart"] < 3].sort_values("Investissement Cagnottage", ascending=False)
if not flop_roi.empty:
    st.markdown(f"""
<div class='alert-card alert-red'>
    <strong>🔴 Alerte Investissement : {len(flop_roi)} offre(s) génèrent un levier inférieur à 3x</strong><br>
    L'offre la plus critique : <strong>{flop_roi.iloc[0]['Article']}</strong> (Levier: {flop_roi.iloc[0]['ROI Écart']:.1f}x | Investissement: {fmt(flop_roi.iloc[0]['Investissement Cagnottage'])} FCFA). 
    Le coût du cagnottage absorbe une part disproportionnée du chiffre d'affaires. Conditions à revoir d'urgence avec le fournisseur.
</div>""", unsafe_allow_html=True)

mag_perf = df_magasin.groupby("Magasin").agg({"CA Fidélité": "sum"}).reset_index()
mag_perf = mag_perf.sort_values("CA Fidélité", ascending=True)
if not mag_perf.empty and mag_perf.iloc[0]["CA Fidélité"] < 100_000:
    st.markdown(f"""
<div class='alert-card alert-amber'>
    <strong>⚠️ Point de vigilance terrain : Volume fidélité critique</strong><br>
    Le site <strong>{mag_perf.iloc[0]['Magasin']}</strong> enregistre une performance très basse avec seulement {fmt(mag_perf.iloc[0]['CA Fidélité'])} FCFA générés sur la période. 
    Vérifier la bonne exécution des supports de communication en magasin et relancer le taux de passage en caisse.
</div>""", unsafe_allow_html=True)

# ─── ACCORDÉONS COMPOSANTS ET TABS ────────────────────────────────────────────
st.markdown("---")
tab1, tab2, tab3 = st.tabs([
    "📍 Suivi Terrain & Magasins",
    "📉 Pilotage Financier Budget",
    "📥 Export Fichiers Excel",
])

# ═══ TAB 1 : VUE TERRAIN ══════════════════════════════════════════════════════
with tab1:
    st.markdown("<div class='section-label'>Analyse de pénétration par Rayon à l'échelle réseau</div>", unsafe_allow_html=True)
    try:
        import plotly.graph_objects as go
        fig = go.Figure(go.Bar(
            x=df_poids_rayon["Poids CA %"].tolist(),
            y=df_poids_rayon["Rayon"].tolist(),
            orientation="h",
            marker_color="#007AFF",
            text=[f"{v:.2f}%" for v in df_poids_rayon["Poids CA %"]],
            textposition="outside",
        ))
        fig.update_layout(
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="-apple-system, Helvetica Neue", color="#3A3A3C", size=11),
            height=max(250, len(df_poids_rayon) * 35 + 40),
            margin=dict(t=10, b=10, l=10, r=70),
            xaxis=dict(title="Poids de la Fidélité dans le CA du Rayon (%)", ticksuffix="%", showgrid=True, gridcolor="#F2F2F7"),
            yaxis=dict(showgrid=False, title=""),
        )
        st.plotly_chart(fig, use_container_width=True)
    except: pass

    st.markdown("<br><div class='section-label'>Détail des performances par Point de Vente et Semaine</div>", unsafe_allow_html=True)
    col_f1, col_f2 = st.columns(2)
    with col_f1: magasins_cibles = st.multiselect("Filtrer par site (Magasin)", options=sorted(df_magasin["Magasin"].unique()))
    with col_f2: rayons_cibles = st.multiselect("Filtrer par Rayon", options=sorted(df_magasin["Rayon"].unique()))

    df_m_aff = df_magasin.copy()
    if magasins_cibles: df_m_aff = df_m_aff[df_m_aff["Magasin"].isin(magasins_cibles)]
    if rayons_cibles:   df_m_aff = df_m_aff[df_m_aff["Rayon"].isin(rayons_cibles)]

    st.dataframe(
        df_m_aff.style.format({"CA Fidélité": "{:,.0f}", "Marge Fidélité": "{:,.0f}", "Qté Vendue": "{:,.1f}"}),
        use_container_width=True, hide_index=True
    )

# ═══ TAB 2 : VUE BUDGETAIRE ═══════════════════════════════════════════════════
with tab2:
    st.markdown("<div class='section-label'>Contrôle de rentabilité des offres nationales et coût du cagnottage</div>", unsafe_allow_html=True)
    sem_cibles = st.multiselect("Isoler une semaine d'analyse", options=sorted(df_finance["Semaine"].unique()))
    df_f_aff = df_finance.copy()
    if sem_cibles: df_f_aff = df_f_aff[df_f_aff["Semaine"].isin(sem_cibles)]

    df_f_display = df_f_aff.drop(columns=["Rayon_Norm"]).rename(columns={
        "Cagnotte_U": "Cagnotte/u", "Qte_Tot": "Qté Totale", "CA_Tot": "CA Généré", "Marge_Tot": "Marge Générée", "ROI Écart": "Levier CA/Invest"
    })
    st.dataframe(
        df_f_display.style.format({
            "Cagnotte/u": "{:,.0f}", "Qté Totale": "{:,.1f}", "Investissement Cagnottage": "{:,.0f}", 
            "CA Généré": "{:,.0f}", "Marge Générée": "{:,.0f}", "Levier CA/Invest": "{:.1f}x"
        }), use_container_width=True, hide_index=True
    )

# ═══ TAB 3 : EXPORT EXCEL BIONGLAT ════════════════════════════════════════════
with tab3:
    st.markdown("<div class='section-label'>Export Excel Professionnel — SmartBuyer Hub</div>", unsafe_allow_html=True)
    st.markdown("""
<div class='alert-card alert-blue'>
    <strong>📋 Structure réglementaire du classeur exporté :</strong><br>
    <strong>Onglet 1 — Suivi Terrain</strong> : Pilotage opérationnel destiné aux équipes magasins (CA, Marges, Volumes par site et par semaine).<br>
    <strong>Onglet 2 — Budget Cagnottage</strong> : Données financières pour la direction commerciale (Montants investis, ROI de l'enveloppe, Marges nettes).
</div>""", unsafe_allow_html=True)

    if st.button("Générer le rapport Excel bionglat", type="primary"):
        with st.spinner("Mise en forme des données et application des styles…"):
            wb_exp = Workbook()
            C_HDR = "1B2A4A"; C_SUB = "2E4B7A"; C_WH = "FFFFFF"; C_DK = "1C1C1E"
            def xfill(h): return PatternFill("solid", fgColor=h)
            def xbdr():
                s = Side(style="thin", color="E5E5EA")
                return Border(left=s, right=s, top=s, bottom=s)
            
            def apply_title_block(ws, title_text, span=9):
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
                c = ws.cell(row=1, column=1, value=title_text)
                c.font = Font("Calibri", size=13, bold=True, color=C_WH)
                c.fill = xfill(C_HDR); c.alignment = Alignment(horizontal="center", vertical="center")
                ws.row_dimensions[1].height = 30
                ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
                c2 = ws.cell(row=2, column=1, value=f"   Périmètre Réseau · {label_periode}")
                c2.font = Font("Calibri", size=9, italic=True, color="AABBCC")
                c2.fill = xfill(C_HDR); c2.alignment = Alignment(horizontal="left", vertical="center")
                ws.row_dimensions[2].height = 16
                ws.row_dimensions[3].height = 6

            def write_headers(ws, row, columns_headers, bg=C_SUB):
                for col_idx, h in enumerate(columns_headers, 1):
                    c = ws.cell(row=row, column=col_idx, value=h)
                    c.font = Font("Calibri", size=10, bold=True, color=C_WH)
                    c.fill = xfill(bg)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.border = xbdr()
                ws.row_dimensions[row].height = 24

            # --- ONGLET 1 : TERRAIN ---
            ws1 = wb_exp.active; ws1.title = "Suivi Terrain"
            apply_title_block(ws1, "PILOTAGE OPERATIONNEL - PERFORMANCE MAGASINS", span=9)
            headers_t = ["Semaine", "Mois", "Magasin", "Rayon", "Code Article", "Article", "CA Fidélité", "Marge Fidélité", "Qté Vendue"]
            write_headers(ws1, 4, headers_t)
            row_idx = 5
            for _, r in df_magasin.iterrows():
                bg_row = "F7F7F7" if row_idx % 2 == 0 else "FFFFFF"
                vals = [r["Semaine"], r["Mois"], r["Magasin"], r["Rayon"], r["Code Article"], r["Article"], r["CA Fidélité"], r["Marge Fidélité"], r["Qté Vendue"]]
                for c_idx, val in enumerate(vals, 1):
                    c = ws1.cell(row=row_idx, column=c_idx, value=val)
                    c.font = Font("Calibri", size=10, color=C_DK); c.fill = xfill("E6F2FF" if c_idx == 3 else bg_row); c.border = xbdr()
                    if c_idx in [7, 8]: c.number_format = "#,##0"; c.alignment = Alignment(horizontal="right")
                    elif c_idx == 9: c.number_format = "#,##0.0"; c.alignment = Alignment(horizontal="right")
                    elif c_idx in [1, 5]: c.alignment = Alignment(horizontal="center")
                    else: c.alignment = Alignment(horizontal="left")
                ws1.row_dimensions[row_idx].height = 20
                row_idx += 1
            ws1.freeze_panes = "A5"

            # --- ONGLET 2 : FINANCIER ---
            ws2 = wb_exp.create_sheet("Budget Cagnottage")
            apply_title_block(ws2, "SUIVI BUDGÉTAIRE ET ENVELOPPE DE CAGNOTTAGE", span=9)
            headers_f = ["Semaine", "Rayon", "Code Article", "Article", "Cagnotte Unitaire", "Qté Totale", "Investissement Cagnottage", "CA Généré", "Marge Générée"]
            write_headers(ws2, 4, headers_f)
            row_idx = 5
            for _, r in df_finance.iterrows():
                bg_row = "F7F7F7" if row_idx % 2 == 0 else "FFFFFF"
                vals = [r["Semaine"], r["Rayon"], r["Code Article"], r["Article"], r["Cagnotte Unitaire"], r["Qté Totale"], r["Investissement Cagnottage"], r["CA Généré"], r["Marge Générée"]]
                for c_idx, val in enumerate(vals, 1):
                    c = ws2.cell(row=row_idx, column=c_idx, value=val)
                    c.font = Font("Calibri", size=10, color=C_DK); c.fill = xfill("FFF9E6" if c_idx == 7 else bg_row); c.border = xbdr()
                    if c_idx in [5, 7, 8, 9]: c.number_format = "#,##0"; c.alignment = Alignment(horizontal="right")
                    elif c_idx == 6: c.number_format = "#,##0.0"; c.alignment = Alignment(horizontal="right")
                    elif c_idx in [1, 3]: c.alignment = Alignment(horizontal="center")
                    else: c.alignment = Alignment(horizontal="left")
                    if c_idx == 7: c.font = Font("Calibri", size=10, bold=True, color=C_DK)
                ws2.row_dimensions[row_idx].height = 20
                row_idx += 1
                
            ws2.cell(row=row_idx, column=1, value="TOTAL").font = Font("Calibri", size=10, bold=True)
            ws2.cell(row=row_idx, column=1).fill = xfill("FFF9E6"); ws2.cell(row=row_idx, column=1).border = xbdr()
            for c_idx, col_let in [(6, "F"), (7, "G"), (8, "H"), (9, "I")]:
                c = ws2.cell(row=row_idx, column=c_idx, value=f"=SUM({col_let}5:{col_let}{row_idx-1})")
                c.font = Font("Calibri", size=10, bold=True); c.fill = xfill("FFF9E6"); c.border = xbdr()
                c.number_format = "#,##0.0" if c_idx == 6 else "#,##0"; c.alignment = Alignment(horizontal="right")
            ws2.freeze_panes = "A5"

            for ws in [ws1, ws2]:
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 12)

            buf = BytesIO()
            wb_exp.save(buf)
            buf.seek(0)
            
        st.download_button(
            label="⬇️ Télécharger le rapport Double-Onglet Excel", data=buf,
            file_name=f"SmartBuyer_Performance_Fidelite_Export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
