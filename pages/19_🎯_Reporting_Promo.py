"""
19_🎯_Reporting_Promo.py — SmartBuyer Hub
=========================================
Reporting rapide de performance des articles en promotion (esprit Tesco).

Colonnes obligatoires — fichier VENTES (extraction) :
  Article, Site nom long, CA, Marge, CA Promo, Marge Promo, Qté Vente, %CA Poids Promo
  (Departement, Rayon, Famille, Sous Famille : optionnelles / contextuelles —
   tous les niveaux de sous-total, quel que soit leur nombre, sont filtrés
   automatiquement via le grain Article × Site="Total".)

Colonnes obligatoires — fichier PRÉVISION :
  Code Article, Libellé, Prévision Qté, Prévision CA, Prévision Marge
  + la période promo en dernière cellule de la colonne A (texte libre avec
    dates jj/mm/aaaa), lue automatiquement et retirée des données.

Rien n'est exploité tant que les contrôles ne sont pas au vert.
"""

import streamlit as st
import pandas as pd
import utils_promo as up

st.set_page_config(page_title="Reporting Promo · SmartBuyer", page_icon="🎯", layout="wide")

BLUE = "#007AFF"
GREEN = "#34C759"
RED = "#FF3B30"
AMBER = "#FF9500"
DARK = "#1D1D1F"
GREY = "#86868B"
BG = "#F2F2F7"

st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display",
                 "SF Pro Text", "Helvetica Neue", Arial, sans-serif !important;
    background-color: #F2F2F7;
}
.stApp { background: #F2F2F7; }
.main .block-container { padding-top: 1.8rem; max-width: 1350px; }
[data-testid="stMetric"] { background: #FFFFFF !important; border: 0.5px solid #E5E5EA !important; border-radius: 12px !important; padding: 16px 18px !important; }
[data-testid="stMetricLabel"] { font-size: 11px !important; font-weight: 500 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 600 !important; color: #1C1C1E !important; letter-spacing: -0.02em !important; }
[data-testid="stDataFrame"] { border: 0.5px solid #E5E5EA !important; border-radius: 10px !important; }
[data-testid="stDataFrame"] th { background: #F2F2F7 !important; font-size: 11px !important; font-weight: 600 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
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

.badge { display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: 600; }
.badge-red    { background: #FF3B30; color: #FFFFFF; }
.badge-green  { background: #34C759; color: #FFFFFF; }
.badge-blue   { background: #007AFF; color: #FFFFFF; }
.badge-amber  { background: #FF9500; color: #FFFFFF; }

.col-required { background: #F0F8FF; border: 0.5px solid #B3D9FF; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; display: flex; align-items: flex-start; gap: 10px; }
.col-name { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
.card { background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:16px;margin-bottom:10px; }
.card-kpi { background:#fff; border-radius:14px; padding:14px 18px; box-shadow:0 1px 3px rgba(0,0,0,.06); }
.kpi  { font-size:22px; font-weight:700; color:#1C1C1E; }
.kpi-l{ font-size:12px; color:#8E8E93; text-transform:uppercase; letter-spacing:.4px; }
.small-muted { font-size:12px;color:#8E8E93; }
.chip-ok   { background:#D6F5DD; color:#1E7A34; }
.chip-warn { background:#FFECD1; color:#C96A00; }
.chip-err  { background:#FFD5D2; color:#FF3B30; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='page-title'>🎯 Reporting Promo — Performance commerciale</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Scorecard article vs prévision + classement magasin, esprit Tesco. "
           "Périmètre piloté par ta liste prévision (la promo hors plan est ignorée). "
           "Rien ne s'affiche tant que les contrôles ne sont pas au vert.</div>", unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    st.markdown("<div class='section-label'>1 · Extraction des ventes</div>", unsafe_allow_html=True)
    f_ventes = st.file_uploader("Article, Site, CA, Marge, CA/Marge Promo, Qté Vente…",
                                type=["xlsx", "xls", "csv"], key="ventes")
with c2:
    st.markdown("<div class='section-label'>2 · Liste prévisions</div>", unsafe_allow_html=True)
    f_prev = st.file_uploader("Code · Libellé · Prév. Qté · Prév. CA · Prév. Marge",
                              type=["xlsx", "xls", "csv"], key="prev")

# =========================================================================== #
# LANDING PAGE — tant que les 2 fichiers ne sont pas chargés
# =========================================================================== #
if not (f_ventes and f_prev):
    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>ℹ️ À quoi sert ce module ?</strong><br>
  Ce module croise ton extraction de ventes avec ta liste de prévisions promo pour produire,
  en un clic, la <strong>scorecard de performance article</strong> (atteinte vs prévision, marge,
  poids CA) et le <strong>classement des magasins</strong> sur le périmètre suivi.
  Deux fichiers à charger ci-dessus, un rapport de contrôle qualité avant toute exploitation.
</div>
""", unsafe_allow_html=True)

    cc1, cc2 = st.columns(2)

    with cc1:
        st.markdown("<div class='section-label'>Contenu du module</div>", unsafe_allow_html=True)
        st.markdown("""
<div class='card'>
  <div style='font-size:14px;font-weight:700;color:#1C1C1E;margin-bottom:8px'>📋 Scorecard articles</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:1.5'>
    Rang (CA réel décroissant) · Sous-famille · Qté/CA/Marge prévus vs réalisés ·
    Poids CA de l'article dans le CA total · % Atteinte avec icône de statut
    (🔴 sous-perf · 🟠 sous tendance · 🟢 cible · 🔵 sur-perf) · alerte marge négative prioritaire.
  </div>
</div>
<div class='card'>
  <div style='font-size:14px;font-weight:700;color:#1C1C1E;margin-bottom:8px'>🏬 Classement magasin</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:1.5'>
    CA, marge et quantité réels par magasin sur tout le périmètre promo suivi, triés du plus gros
    contributeur au plus petit — avec la part de chaque site dans le CA réseau du périmètre.
  </div>
</div>
<div class='card'>
  <div style='font-size:14px;font-weight:700;color:#1C1C1E;margin-bottom:8px'>🔍 Contrôles qualité</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:1.5'>
    Structure des fichiers, cohérence de la jointure prévision ↔ ventes, valeurs manquantes ou
    négatives, marges promo destructrices — tout est vérifié avant que les tableaux ne s'affichent.
  </div>
</div>
""", unsafe_allow_html=True)

    with cc2:
        st.markdown("<div class='section-label'>Fonctionnement</div>", unsafe_allow_html=True)
        st.markdown("""
<div class='alert-card alert-green'>
  <strong>1.</strong> Charge l'extraction ventes de la période.<br>
  <strong>2.</strong> Charge ta liste prévision (Code, Libellé, Prév. Qté/CA/Marge).<br>
  <strong>3.</strong> Vérifie le rapport de contrôle qualité — corrige si un point est bloquant.<br>
  <strong>4.</strong> Consulte la scorecard article et le classement magasin.
</div>
""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div class='section-label'>Règles de calcul</div>", unsafe_allow_html=True)
        st.markdown(f"""
<div class='card'>
  <span class='badge badge-blue'>Périmètre</span>
  <div class='small-muted' style='margin-top:6px'>Piloté par la liste prévision — un article vendu en promo mais absent du plan n'apparaît pas.</div>
</div>
<div class='card'>
  <span class='badge badge-amber'>% Atteinte</span>
  <div class='small-muted' style='margin-top:6px'>CA réel ÷ CA prévu, comparé au CA <u>total</u> (pas seulement promo) pour le Poids CA.</div>
</div>
<div class='card'>
  <span class='badge badge-red'>Alerte marge</span>
  <div class='small-muted' style='margin-top:6px'>Une marge promo négative prime sur tout autre statut, quel que soit le volume vendu.</div>
</div>
""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Colonnes attendues dans les fichiers</div>", unsafe_allow_html=True)

    cols_ventes = [
        ("Article", "Code + libellé — obligatoire, sert de clé de jointure"),
        ("Site nom long", "\"Total\" = grain réseau exploité — obligatoire"),
        ("CA / Marge", "Valeurs totales période — obligatoires (Poids CA, en-tête)"),
        ("CA Promo / Marge Promo", "Univers promo + alerte marge négative — obligatoires"),
        ("Qté Vente", "Axe quantité réalisée — obligatoire"),
        ("%CA Poids Promo", "Signal de dépendance à la promo — obligatoire"),
        ("Departement / Rayon / Famille / Sous Famille", "Contextuel — tous les sous-totaux sont filtrés automatiquement"),
        ("Casse (Qté)", "Optionnelle — NaN traité comme absence de casse"),
    ]
    cols_prev = [
        ("Code Article", "Clé de jointure — obligatoire"),
        ("Libellé", "Affiché dans la scorecard"),
        ("Prévision Qté / CA / Marge", "3 axes — obligatoires"),
        ("Dernière cellule colonne A", "Période promo (texte libre avec dates jj/mm/aaaa) — lue automatiquement"),
    ]

    t1, t2 = st.columns(2)
    with t1:
        st.markdown("<div class='small-muted' style='margin-bottom:6px'><b>Fichier ventes</b></div>", unsafe_allow_html=True)
        for name, desc in cols_ventes:
            st.markdown(f"""
<div class='col-required'>
  <div style='font-size:16px'>▪️</div>
  <div><div class='col-name'>{name}</div><div class='col-desc'>{desc}</div></div>
</div>
""", unsafe_allow_html=True)
    with t2:
        st.markdown("<div class='small-muted' style='margin-bottom:6px'><b>Fichier prévision</b></div>", unsafe_allow_html=True)
        for name, desc in cols_prev:
            st.markdown(f"""
<div class='col-required'>
  <div style='font-size:16px'>▪️</div>
  <div><div class='col-name'>{name}</div><div class='col-desc'>{desc}</div></div>
</div>
""", unsafe_allow_html=True)

    st.info("⬅️ Charge les deux fichiers ci-dessus pour démarrer.")
    st.stop()

# =========================================================================== #
# CHARGEMENT + CONTRÔLES
# =========================================================================== #
try:
    df_raw = up.load_ventes(f_ventes)
    meta = up.extract_meta(df_raw)
    df_art = up.to_article_reseau(df_raw)
    df_prev, found, prev_meta = up.load_previsions(f_prev)
except Exception as e:
    st.error(f"Lecture impossible : {e}")
    st.stop()

report1, bloq1 = up.run_controls(df_raw, df_art, df_prev, found, meta, prev_meta)
perim = up.build_perimetre(df_art, df_prev)
report2, bloq2 = up.controls_jointure(perim, df_prev)
kpi = up.compute_kpi(perim)
report3, bloq3 = up.controls_kpi(kpi)

all_reports = report1 + report2 + report3
bloquant = bloq1 or bloq2 or bloq3

def _row(r):
    b = {"ok": "chip-ok", "warn": "chip-warn", "err": "chip-err"}[r["statut"]]
    lbl = {"ok": "OK", "warn": "WARN", "err": "STOP"}[r["statut"]]
    return (f'<div class="card" style="margin-bottom:6px;"><span class="badge {b}">{lbl}</span> '
            f'<b>{r["contrôle"]}</b> &nbsp;·&nbsp; {r["valeur"]}'
            f'<div style="color:#8E8E93;font-size:12px;margin-top:2px">{r["message"]}</div></div>')

with st.expander(f"🔍 Rapport de contrôle qualité "
                 f"({sum(1 for r in all_reports if r['statut']=='ok')} OK · "
                 f"{sum(1 for r in all_reports if r['statut']=='warn')} à surveiller · "
                 f"{sum(1 for r in all_reports if r['statut']=='err')} bloquant)",
                 expanded=bloquant):
    st.markdown("".join(_row(r) for r in all_reports), unsafe_allow_html=True)

if bloquant:
    st.error("⛔ Contrôles bloquants détectés — corrige les fichiers avant d'exploiter les données. "
             "Voir le détail ci-dessus.")
    st.stop()

st.success("✅ Contrôles au vert. Données exploitées.")

# =========================================================================== #
# EN-TÊTE : KPI atteinte, CA total, Marge totale, Période promo
# =========================================================================== #
V = up.VENTES_MAP
sc = up.build_scorecard(kpi, df_art)
ca_total_ext = df_art[V["ca"]].sum()
marge_total_ext = df_art[V["marge"]].sum()
kpi_atteinte = sc["CA réal."].sum() / sc["CA prév."].sum() if sc["CA prév."].sum() else 0
periode_txt = (f"{prev_meta['periode_debut']} → {prev_meta['periode_fin']}"
              if prev_meta.get("periode_debut") else "à préciser")

h1, h2, h3, h4 = st.columns(4)
for col, label, val in [
    (h1, "% Atteinte CA", f"{kpi_atteinte:.0%}"),
    (h2, "CA total (extraction)", f"{ca_total_ext:,.0f} F".replace(",", " ")),
    (h3, "Marge totale (extraction)", f"{marge_total_ext:,.0f} F".replace(",", " ")),
    (h4, "Période promo", periode_txt),
]:
    col.markdown(f'<div class="card-kpi"><div class="kpi-l">{label}</div>'
                 f'<div class="kpi">{val}</div></div>', unsafe_allow_html=True)

st.caption("Période lue automatiquement dans la dernière cellule de la colonne A de la liste prévision.")

# =========================================================================== #
# SCORECARD ARTICLE
# =========================================================================== #
st.markdown("<div class='section-label' style='margin-top:1.5rem'>Scorecard articles</div>", unsafe_allow_html=True)

disp = sc.copy()
disp["% Atteinte"] = (disp["% Atteinte"] * 100).round(0).astype("Int64")
disp["Poids CA"] = disp["Poids CA"].round(2)
disp["Icône"] = disp["Statut"].map(up.RAG_ICON).fillna("⚪")
disp["% Atteinte"] = disp["Icône"] + " " + disp["% Atteinte"].astype(str) + "%"

show_cols = ["Rang", "Sous Famille", "Code Article", "Libellé",
            "Qté prév.", "CA prév.", "Marge prév.",
            "Qté réal.", "CA réal.", "Marge réal.",
            "Poids CA", "% Atteinte", "Alerte"]

st.dataframe(
    disp[show_cols],
    use_container_width=True,
    hide_index=True,
    column_config={
        "CA prév.": st.column_config.NumberColumn(format="%d F"),
        "Marge prév.": st.column_config.NumberColumn(format="%d F"),
        "CA réal.": st.column_config.NumberColumn(format="%d F"),
        "Marge réal.": st.column_config.NumberColumn(format="%d F"),
        "Poids CA": st.column_config.NumberColumn(format="%.2f%%"),
    },
)
st.caption("Rang = CA réel décroissant · % Atteinte = CA réal. ÷ CA prév. "
          "· Poids CA = CA réel de l'article ÷ CA total de l'extraction (tous articles). "
          "🔴 <70% · 🟠 70–90% · 🟢 90–115% · 🔵 >115%")

# =========================================================================== #
# CLASSEMENT MAGASIN
# =========================================================================== #
st.markdown("<div class='section-label' style='margin-top:1.5rem'>Classement par magasin</div>", unsafe_allow_html=True)
st.caption("Sur tout le périmètre promo suivi. Valeurs réelles (pas de % réalisation : "
          "pas de prévision éclatée par magasin).")

mag = up.store_ranking(df_raw, perim["code"].tolist())
mag_disp = mag.copy()
mag_disp["Part réseau"] = mag_disp["Part réseau"].round(1)

st.dataframe(
    mag_disp,
    use_container_width=True,
    hide_index=True,
    column_config={
        "CA": st.column_config.NumberColumn(format="%d F"),
        "Marge": st.column_config.NumberColumn(format="%d F"),
        "Part réseau": st.column_config.ProgressColumn(
            format="%.1f%%", min_value=0, max_value=float(mag_disp["Part réseau"].max() or 1)),
    },
)
