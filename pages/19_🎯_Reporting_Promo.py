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

st.set_page_config(page_title="Reporting Promo", page_icon="🎯", layout="wide")

st.markdown("""
<style>
  .stApp { background:#F2F2F7; }
  h1,h2,h3 { color:#1C1C1E; font-family:-apple-system,'Helvetica Neue',Arial,sans-serif; }
  .card { background:#fff; border-radius:14px; padding:14px 18px;
          box-shadow:0 1px 3px rgba(0,0,0,.06); }
  .kpi  { font-size:22px; font-weight:700; color:#1C1C1E; }
  .kpi-l{ font-size:12px; color:#8E8E93; text-transform:uppercase; letter-spacing:.4px; }
  .badge{ display:inline-block; padding:2px 10px; border-radius:999px; font-size:12px; font-weight:600; }
  .ok   { background:#D6F5DD; color:#1E7A34; }
  .warn { background:#FFECD1; color:#C96A00; }
  .err  { background:#FFD5D2; color:#FF3B30; }
</style>
""", unsafe_allow_html=True)

st.title("🎯 Reporting Promo — Performance commerciale")
st.caption("Scorecard article + classement magasin. Périmètre piloté par la liste prévision "
           "(la promo hors plan est ignorée). Rien ne s'affiche tant que les contrôles ne sont pas au vert.")

c1, c2 = st.columns(2)
with c1:
    st.markdown("##### 1 · Extraction des ventes")
    f_ventes = st.file_uploader("Article, Site, CA, Marge, CA/Marge Promo, Qté Vente…",
                                type=["xlsx", "xls", "csv"], key="ventes")
with c2:
    st.markdown("##### 2 · Liste prévisions")
    f_prev = st.file_uploader("Code · Libellé · Prév. Qté · Prév. CA · Prév. Marge",
                              type=["xlsx", "xls", "csv"], key="prev")

if not (f_ventes and f_prev):
    st.info("Charge les deux fichiers pour lancer les contrôles.")
    st.stop()

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
    b = r["statut"]
    lbl = {"ok": "OK", "warn": "WARN", "err": "STOP"}[b]
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
    col.markdown(f'<div class="card"><div class="kpi-l">{label}</div>'
                 f'<div class="kpi">{val}</div></div>', unsafe_allow_html=True)

st.caption("Période lue automatiquement dans la dernière cellule de la colonne A de la liste prévision.")

st.markdown("### Scorecard articles")

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

st.markdown("### Classement par magasin")
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
