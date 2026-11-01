"""SmartBuyer Hub · Suivi implantation des nouvelles références (T1).

Page de pilotage (usage personnel) : on charge la liste T1 et les extractions de stock ERP
(un fichier par magasin) ; le module calcule le taux d'implantation par magasin et par rayon,
donne l'action à mener par article × magasin et propose des cessions inter-magasins.

Les calculs et les exports sont dans modules/implantation/engine.py (logique de la v7.1).
"""
from __future__ import annotations

import pathlib
import sys

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core import theme as T  # noqa: E402
from modules.implantation import engine as E  # noqa: E402

st.set_page_config(page_title="Suivi implantation · SmartBuyer", page_icon="🏪", layout="wide", initial_sidebar_state="expanded")
T.apply()

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
IMPL, APPRO, CMD, ANOM = "✅ Implanté", "🔵 Appro en cours", "🛒 Passer commande", "🚩 Anomalie référ."
STATUT_COLOR = {IMPL: T.TEAL, APPRO: T.BLEU, CMD: T.ORANGE, ANOM: T.RED}
STATUT_KIND = {IMPL: "ok", APPRO: "info", CMD: "att", ANOM: "crit"}


def n(v) -> str:
    return E.fmt_n(v)


def band(t) -> str:
    return "ok" if t >= 80 else ("att" if t >= 50 else "crit")


def table(df, **kw):
    st.dataframe(df, use_container_width=True, hide_index=True, **kw)


# ══════════════════════════════════════════════════════════════════════════════
# BARRE LATÉRALE
# ══════════════════════════════════════════════════════════════════════════════
T.sidebar_brand("🏪", "Suivi implantation")
T.sidebar_label("1 · Liste T1")
with st.sidebar:
    t1_file = st.file_uploader("Nouvelles références T1", type=["csv", "xlsx", "xls"], key="t1",
                               help="Colonne ARTICLE (code 8 chiffres) ; MODE APPRO, SEMAINE RECEPTION et fournisseur en option.")
T.sidebar_label("2 · Stocks ERP par magasin")
with st.sidebar:
    stk_files = st.file_uploader("Un fichier CSV par magasin", type=["csv"], accept_multiple_files=True, key="stk",
                                 help="CSV « ; », nommage *_10206_20260507_*.csv (code site, date).")


# ══════════════════════════════════════════════════════════════════════════════
# ACCUEIL (tant que les deux fichiers ne sont pas déposés)
# ══════════════════════════════════════════════════════════════════════════════
t1_df = None
if t1_file:
    with st.spinner("Lecture T1…"):
        t1_df, t1_err = E.parse_t1(t1_file.read(), t1_file.name)
    if t1_err or t1_df is None:
        T.page_header("🏪", "Suivi implantation")
        T.alert("Liste T1 illisible", str(t1_err), "crit")
        st.stop()

if t1_df is None or not stk_files:
    T.sidebar_footer("Taux = articles implantés ÷ articles actifs (code état 2)")
    rows = [("Liste T1 des nouvelles références", "ok" if t1_df is not None else "wait",
             f"{n(len(t1_df))} réf." if t1_df is not None else "à déposer"),
            ("Stocks ERP par magasin", "ok" if stk_files else "wait",
             f"{len(stk_files)} fichier(s)" if stk_files else "1 à 13 fichiers")]
    T.landing_hero("LES NOUVEAUTÉS SONT-ELLES <b>EN RAYON</b> ?",
                   "Taux d'implantation des références T1 magasin par magasin, et l'action à mener pour chaque "
                   "article manquant : accélérer, commander, corriger le référencement ou céder entre magasins.",
                   eyebrow="Suivi implantation · SmartBuyer Hub",
                   status=T.landing_status("Préparation de l'analyse", rows, "fichiers reçus"))
    T.steps([("Déposer", "la liste T1 de la période"),
             ("Ajouter", "l'extraction stock ERP de chaque magasin"),
             ("Lire", "le taux d'implantation par magasin et par rayon"),
             ("Agir", "commander, accélérer, corriger le référencement"),
             ("Céder", "le stock en trop d'un magasin à l'autre")])
    T.section("Avant de déposer")
    c1, c2, c3 = st.columns(3)
    with c1:
        T.columns_card("Colonnes attendues", [
            ("Liste T1", ["ARTICLE", "LIBELLÉ ARTICLE", "MODE APPRO", "SEMAINE RECEPTION", "LIBELLÉ FOURNISSEUR ORIGINE"]),
            ("Stock ERP (CSV ;)", ["Site", "Libellé site", "Code article", "Code etat", "Nouveau stock", "Ral", "Pcb",
                                   "Code marketing", "Nom fourn.", "Libellé rayon", "Libellé famille"]),
        ], "Seule ARTICLE est obligatoire dans la T1. Le code site et la date viennent du nom de fichier si besoin.")
    with c2:
        T.rules_card("🚦 Les 4 statuts", [
            (T.badge("Implanté", "ok"), "stock ≠ 0"),
            (T.badge("Appro en cours", "info"), "stock = 0 et RAL > 0 : accélérer"),
            (T.badge("À commander", "att"), "stock = 0 et RAL = 0 : commander"),
            (T.badge("Anomalie", "crit"), "code état ≠ 2 ou absent : vérifier"),
        ])
    with c3:
        T.rules_card("🔄 Règles de calcul", [
            ("Taux d'implantation", "implantés ÷ articles actifs"),
            ("Rupture réseau", "article non implanté partout"),
            ("Cession", "le cédant garde 2 PCB"),
            ("Stock ancien", "alerte au-delà de 7 jours"),
        ])
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# FILTRES, CONSOLIDATION, CALCULS
# ══════════════════════════════════════════════════════════════════════════════
T.sidebar_label("3 · Filtres")
with st.sidebar:
    orig_sel = st.multiselect("Flux", ["IM", "LO"], default=["IM", "LO"], format_func=lambda x: {"IM": "Import", "LO": "Local"}[x])
    sem_dispo = sorted([s for s in t1_df["SEMAINE RECEPTION"].unique() if str(s).strip() not in ("nan", "", "99")], key=E._sem_sort)
    sem_sel = st.multiselect("Semaine de réception", sem_dispo, default=sem_dispo)

SKU_SCOPE = tuple(sorted(t1_df["SKU"].unique()))
with st.spinner(f"Consolidation de {len(stk_files)} fichier(s) stock…"):
    df_stock, stk_errors = E.consolidate_stock(stk_files, SKU_SCOPE)
if df_stock.empty:
    T.page_header("🏪", "Suivi implantation")
    T.alert("Aucune donnée stock valide", " · ".join(stk_errors) or "Les fichiers ne contiennent aucun article de la T1.", "crit")
    st.stop()

R = E.build(t1_df, df_stock, orig_sel, sem_sel)
if R is None:
    T.page_header("🏪", "Suivi implantation")
    T.alert("Aucun article pour ces filtres", "Élargissez le flux ou les semaines de réception.", "att")
    st.stop()

merged, pm = R["merged"], R["pivot_mag"]
date_stock = pd.Timestamp(max(df_stock["_date_fichier"].unique()))
age_stock = (E.TODAY - date_stock).days
T.sidebar_footer(f"Stock du {date_stock:%d/%m/%Y} · {len(stk_files)} fichier(s)<br>Taux = implantés ÷ articles actifs")


# ══════════════════════════════════════════════════════════════════════════════
# BANDEAU, KPI, ALERTES
# ══════════════════════════════════════════════════════════════════════════════
n_actions = R["n_appro"] + R["n_cmd"] + R["n_anomalie"]
T.hero(f"IMPLANTATION T1 · <b>{R['taux_reseau']} %</b>", f"Stock du {date_stock:%d/%m/%Y}", [
    f"{R['n_mag']} magasins", f"{n(R['n_sku'])} réf. T1 · import {R['n_sku_im']} · local {R['n_sku_lo']}",
    (f"{n(n_actions)} actions à mener", "alert") if n_actions else ("Tout est implanté", "ok"),
    (f"{len(R['sku_rupt'])} ruptures réseau", "alert") if R["sku_rupt"] else ("Aucune rupture réseau", "ok"),
], eyebrow="Suivi implantation · nouvelles références")

T.kpi_row([
    T.kpi("Taux d'implantation", str(R["taux_reseau"]), "objectif 80 %", {"ok": "pos", "att": "mut", "crit": "neg"}[band(R["taux_reseau"])],
          "articles actifs implantés", "blue", "🏪", "%"),
    T.kpi("Implantés", n(R["n_impl"]), "", "pos", "article × magasin", "teal", "✅"),
    T.kpi("Appro en cours", n(R["n_appro"]), "accélérer", "mut", "RAL > 0, stock nul", "violet", "🚚"),
    T.kpi("À commander", n(R["n_cmd"]), "commander", "neg" if R["n_cmd"] else "pos", "ni stock ni RAL", "violet", "🛒"),
    T.kpi("Anomalies", n(R["n_anomalie"]), "référencement", "neg" if R["n_anomalie"] else "pos", "code état ≠ 2", "", "🚩"),
])

msgs = []
if stk_errors:
    msgs.append((f"{len(stk_errors)} fichier(s) stock ignoré(s)", " · ".join(stk_errors), "crit"))
if age_stock > 7:
    msgs.append((f"Stock vieux de {age_stock} jours", "Rechargez des extractions plus récentes avant de décider.", "att"))
if R["n_sku_absents"]:
    msgs.append((f"{n(R['n_sku_absents'])} réf. T1 absentes de tous les stocks",
                 f"{n(R['n_sku_trouves'])} sur {n(R['n_sku'])} trouvées : comptées en anomalie de référencement.", "att"))
if R["sku_rupt"]:
    msgs.append((f"{len(R['sku_rupt'])} rupture(s) sur tout le réseau", "Aucun magasin n'a de stock : commander en urgence ou céder "
                 "depuis un magasin hors périmètre (onglet Cessions).", "crit"))
if msgs:
    cols = st.columns(min(len(msgs), 2))
    for i, (t, l, k) in enumerate(msgs):
        with cols[i % len(cols)]:
            T.alert(t, l, k)


# ══════════════════════════════════════════════════════════════════════════════
# ONGLETS
# ══════════════════════════════════════════════════════════════════════════════
tab0, tab1, tab2, tab3, tab4 = st.tabs(["🎯 Synthèse", f"🚨 Actions · {n(n_actions)}", "🔢 Rayon × magasin", "🔄 Cessions", "⬇️ Export"])

# ── Synthèse ──────────────────────────────────────────────────────────────────
with tab0:
    left, right = st.columns([1.3, 1])
    with left:
        T.section("Taux d'implantation par magasin")
        rows = ""
        for _, vals in pm.sort_values("Taux (%)").iterrows():
            tot = sum(int(vals.get(s, 0)) for s in STATUT_COLOR) or 1
            seg = "".join(f"<span title='{s}' style='width:{int(vals.get(s, 0)) / tot * 100:.1f}%;background:{c}'></span>"
                          for s, c in STATUT_COLOR.items())
            rows += (f"<div style='display:grid;grid-template-columns:minmax(0,1.1fr) minmax(0,1.4fr) auto;gap:14px;align-items:center;"
                     f"padding:9px 0;border-bottom:1px solid {T.LIGNE}'>"
                     f"<div style='font-weight:600;font-size:13.5px'>{vals['Magasin']}<small style='display:block;color:{T.GRIS_TEXTE};"
                     f"font-weight:500'>{int(vals.get(CMD, 0))} à commander · {int(vals.get(APPRO, 0))} en appro · "
                     f"{int(vals.get(ANOM, 0))} anomalie(s)</small></div>"
                     f"<div style='display:flex;height:10px;border-radius:6px;overflow:hidden;background:{T.GRIS_CLAIR}'>{seg}</div>"
                     f"{T.badge(str(int(vals['Taux (%)'])) + ' %', band(vals['Taux (%)']))}</div>")
        legend = " ".join(f"<span style='display:inline-flex;align-items:center;gap:5px;margin-right:12px'>"
                          f"<span style='width:10px;height:10px;border-radius:3px;background:{c}'></span>{s[2:]}</span>"
                          for s, c in STATUT_COLOR.items())
        T.card(rows + f"<div class='note'>{legend}</div>")
    with right:
        T.section("Répartition réseau")
        labels = [s for s in STATUT_COLOR if s in pm.columns]
        fig = go.Figure(go.Pie(labels=[s[2:] for s in labels], values=[int(pm[s].sum()) for s in labels], hole=0.66, sort=False,
                               marker=dict(colors=[STATUT_COLOR[s] for s in labels], line=dict(color="#fff", width=3)),
                               textinfo="percent", textfont=dict(size=12, color="#fff")))
        fig.add_annotation(text=f"<b>{R['taux_reseau']} %</b><br>implanté", x=0.5, y=0.5, showarrow=False,
                           font=dict(size=20, color=T.NAVY))
        T.plotly_layout(fig, 300)
        st.plotly_chart(fig, use_container_width=True)
        T.section("Articles à commander en priorité")
        dc = merged[merged["Alerte"] == CMD].groupby("SKU").agg(Article=("Libellé article", "first"),
                                                               Fournisseur=("Fournisseur T1", "first"),
                                                               Magasins=("Magasin", "count")).reset_index()
        if dc.empty:
            T.alert("Rien à commander", "Chaque article a du stock ou une commande en cours.", "ok")
        else:
            table(dc.sort_values("Magasins", ascending=False).head(10))
    T.section("Appros à accélérer")
    da = merged[merged["Alerte"] == APPRO].groupby("SKU").agg(Article=("Libellé article", "first"), Fournisseur=("Fournisseur T1", "first"),
                                                             Magasins=("Magasin", "count"), RAL=("Ral", "sum")).reset_index()
    if da.empty:
        T.alert("Aucune appro en attente", kind="ok")
    else:
        table(da.sort_values("Magasins", ascending=False).head(15))

# ── Actions ───────────────────────────────────────────────────────────────────
with tab1:
    c1, c2 = st.columns(2)
    dispo = [a for a in (CMD, APPRO, ANOM) if (merged["Alerte"] == a).any()]
    sel_a = c1.multiselect("Statut", dispo, default=dispo, format_func=lambda a: f"{a[2:]} · {E.ACTION_LABEL[a]}")
    mags = sorted(merged["Magasin"].unique())
    sel_m = c2.multiselect("Magasin", mags, default=[], placeholder="Tous")
    d = merged[merged["Alerte"].isin(sel_a)]
    if sel_m:
        d = d[d["Magasin"].isin(sel_m)]
    if d.empty:
        T.alert("Aucune action pour ces filtres", kind="ok")
    else:
        T.html("<div style='display:flex;gap:8px;flex-wrap:wrap;margin:4px 0 10px'>" + "".join(
            T.badge(f"{a[2:]} · {int((d['Alerte'] == a).sum())} · {E.ACTION_LABEL[a]}", STATUT_KIND[a]) for a in sel_a) + "</div>")
        cols = [c for c in E.COLS_DET if c in d.columns]
        table(d[cols].sort_values(["Alerte", "Magasin"]))
    if R["sku_rupt"]:
        with st.expander(f"🚨 {len(R['sku_rupt'])} rupture(s) sur tout le réseau"):
            table(merged[merged["SKU"].isin(R["sku_rupt"])][["SKU", "Libellé article", "Origine", "Fournisseur T1", "Alerte"]]
                  .drop_duplicates("SKU").sort_values("Alerte"))

# ── Rayon × magasin ───────────────────────────────────────────────────────────
with tab2:
    rp = R["rayon_pivot"]
    if rp.empty:
        T.alert("Pas de rayon dans les stocks", "La colonne « Libellé rayon » est vide ou absente.", "att")
    else:
        T.section("Taux d'implantation · rayon × magasin")
        head = "".join(f"<th style='text-align:center'>{c}</th>" for c in rp.columns)
        body = ""
        for ray, line in rp.iterrows():
            cells = ""
            for v in line:
                bg, fg = T.heat(v - 65, 35)
                cells += f"<td style='text-align:center;background:{bg};color:{fg};font-weight:700'>{int(v)} %</td>"
            body += f"<tr><td><b>{ray}</b></td>{cells}</tr>"
        T.card(f"<div style='overflow-x:auto'><table class='tbl'><thead><tr><th>Rayon</th>{head}</tr></thead><tbody>{body}</tbody></table></div>"
               "<div class='note'>Vert au-dessus de 65 %, rouge en dessous ; plus la couleur est forte, plus on s'en éloigne.</div>")

# ── Cessions ──────────────────────────────────────────────────────────────────
with tab3:
    if R["sku_rupt"]:
        T.section("Ruptures réseau : magasin qui peut céder")
        auto = E.auto_cessions(df_stock, R["site_ref"], R["sku_rupt"])
        if auto.empty:
            T.alert("Aucun magasin ne peut céder", "Aucun stock au-delà de 2 PCB pour ces articles.", "att")
        else:
            table(auto)
    T.section("Plan de cessions vers des magasins en difficulté")
    c1, c2, c3 = st.columns([2, 1, 1])
    mag_labels = sorted(R["site_ref"].get(c, c) for c in R["all_codes"])
    mag_det = c1.multiselect("Magasins à réapprovisionner", mag_labels, default=[])
    seuil = c2.number_input("Stock inférieur ou égal à", 0, 50, 0, 1)
    min_1pcb = c3.toggle("Au moins 1 PCB cédé", value=True)
    if not mag_det:
        T.alert("Choisissez un ou plusieurs magasins", "Le module cherche, pour chaque article sous le seuil, le magasin qui a le plus "
                "de stock au-delà de 2 PCB.", "info")
    else:
        dc = E.manual_cessions(df_stock, R["site_ref"], R["all_codes"], R["sku_scope"], mag_det, seuil, min_1pcb)
        if dc.empty:
            T.alert("Aucune cession possible", kind="ok")
        else:
            table(dc)
            st.download_button(f"⬇️ Plan de cessions ({len(dc)} lignes)", data=E.build_cessions_export(dc),
                               file_name=f"Plan_Cessions_{E.TODAY_FILE}.xlsx", mime=XLSX)

# ── Export ────────────────────────────────────────────────────────────────────
with tab4:
    T.section("Export Excel · valeurs uniquement")
    T.card("<div style='font-size:13px;line-height:1.6'>Trois onglets : synthèse par magasin, détail complet article × magasin, "
           "et alertes à traiter avec l'action attendue.</div>"
           f"<div class='note'>Stock du {date_stock:%d/%m/%Y} · {n(len(merged))} lignes · {R['n_mag']} magasin(s)</div>", "Classeur implantation T1")
    if st.button("Générer le fichier Excel", type="primary", key="gen_impl"):
        with st.spinner("Génération…"):
            st.session_state["impl_xlsx"] = E.build_export(merged, pm)
    if st.session_state.get("impl_xlsx"):
        st.download_button("⬇️ Télécharger le classeur", data=st.session_state["impl_xlsx"],
                           file_name=f"Implantation_T1_{E.TODAY_FILE}.xlsx", mime=XLSX, key="dl_impl")
