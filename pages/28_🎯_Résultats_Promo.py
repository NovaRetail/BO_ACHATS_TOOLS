"""SmartBuyer Hub · Résultats Promo — challenge des prévisions acheteurs.

Page de pilotage (usage personnel) : on charge l'extraction ventes PBI de la période promo et la liste
promo des acheteurs (Code, Prévision Qté, Prix promo, PA, TVA). On confronte ce qui a été promis
(quantité, CA, marge) à ce qui est sorti, puis on compare le CA et la marge réalisés à N-1.
Export Excel : valeurs uniquement, aucune formule.

Les calculs et l'export sont dans modules/promo_resultats/engine.py.
"""
from __future__ import annotations

import pathlib
import re
import sys
from datetime import datetime

import pandas as pd
import streamlit as st

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core import theme as T  # noqa: E402
from modules.promo_resultats import engine as E  # noqa: E402

st.set_page_config(page_title="Résultats Promo · SmartBuyer", page_icon="🎯", layout="wide",
                   initial_sidebar_state="expanded")
T.apply()

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def rbadge(v: str) -> str:
    return T.badge(v, E.RELIABILITY[v]["kind"])


def nbadge(v: str) -> str:
    return T.badge(v, E.VS_N1[v]["kind"])


def pct_badge(x, low, high) -> str:
    if pd.isna(x):
        return "—"
    kind = "crit" if x < low else ("att" if x > high else "ok")
    return T.badge(E.fmt_pct(x, 0), kind)


def evo_badge(x, band) -> str:
    if pd.isna(x):
        return "—"
    kind = "ok" if x > band else ("crit" if x < -band else "info")
    return T.badge(E.fmt_pct(x, 0, True), kind)


def sm(text) -> str:
    return f"<small style='display:block;color:{T.GRIS_TEXTE};font-weight:500'>{text}</small>"


# ══════════════════════════════════════════════════════════════════════════════
# BARRE LATÉRALE
# ══════════════════════════════════════════════════════════════════════════════
T.sidebar_brand("🎯", "Résultats Promo")
T.sidebar_label("1 · Extraction ventes")
with st.sidebar:
    f_sales = st.file_uploader("Export PBI ventes par article (xlsx / csv)", type=["xlsx", "xlsm", "csv"], key="rp_sales",
                               help="Export hiérarchique Departement › Rayon › Famille › Article › Site, filtré sur la période de la promo.")
T.sidebar_label("2 · Liste promo des acheteurs")
with st.sidebar:
    f_promo = st.file_uploader("Code, Prévision Qté, Prix promo, PA, TVA", type=["xlsx", "xlsm", "csv"], key="rp_promo",
                               help="Une ligne par article. Le CA prévu et la marge prévue sont calculés par le module.")
    price_ttc = st.toggle("Prix promo saisi TTC", value=True, help="Si activé, le CA prévu est ramené en HT avec la colonne TVA.")
T.sidebar_label("3 · Règles")
with st.sidebar:
    low = st.slider("Prévision surestimée en dessous de", 50, 100, 80, 5, format="%d %%") / 100
    high = st.slider("Prévision sous-estimée au-delà de", 100, 200, 120, 5, format="%d %%") / 100
    band = st.slider("Stable vs N-1 (±)", 0, 20, 5, 1, format="%d %%") / 100
    converted = st.multiselect("Sites au N-1 non comparable", E.DEFAULT_CONVERTED_SITES, default=E.DEFAULT_CONVERTED_SITES)
params = {"low": low, "high": max(high, low), "stable_band": band, "converted_sites": converted}


# ══════════════════════════════════════════════════════════════════════════════
# ACCUEIL
# ══════════════════════════════════════════════════════════════════════════════
if f_sales is None or f_promo is None:
    T.sidebar_footer("CA prévu HT = Qté × Prix promo ÷ (1 + TVA)<br>Marge prévue = Qté × (Prix HT − PA)")
    rows = [("Extraction ventes PBI", "ok" if f_sales else "wait", "reçue" if f_sales else "à déposer"),
            ("Liste promo des acheteurs", "ok" if f_promo else "wait", "reçue" if f_promo else "à déposer")]
    T.landing_hero("PRÉVU <b>VS RÉALISÉ</b>",
                   "Ce que l'acheteur a promis en quantité, CA et marge, ce qui est réellement sorti, "
                   "et si la promo a fait mieux que l'an dernier.",
                   eyebrow="Résultats Promo · SmartBuyer Hub",
                   status=T.landing_status("Préparation de l'analyse", rows, "fichiers obligatoires reçus", count=2))
    T.steps([("Collecter", "la liste promo de chaque acheteur avant l'opération : quantité, prix promo, PA, TVA"),
             ("Extraire", "les ventes par article × site dans Power BI, sur la seule période de la promo"),
             ("Déposer", "les deux fichiers dans la barre latérale"),
             ("Challenger", "les écarts prévu vs réalisé, article par article"),
             ("Exporter", "le classeur pour la revue avec les acheteurs")])
    T.section("Format des fichiers attendus")
    c1, c2, c3 = st.columns(3)
    with c1:
        T.columns_card("Colonnes attendues", [
            ("Liste promo des acheteurs", ["Code", "Prévision Qté (quantité totale, tous sites)", "Prix promo (TTC par défaut)",
                                           "PA (prix d'achat HT)", "TVA (18, 0,18 ou 18 %)", "Libellé (facultatif)"]),
            ("Extraction ventes (export PBI)", ["Article", "Site nom long", "CA", "Marge", "Qté Vente",
                                                "CA N-1 · Qté Vente N-1 · Marge Hors Promo N-1 · Marge Promo N-1"]),
        ], "Article au format « 12003043 - LIBELLÉ » : le code avant le tiret sert de clé. Les lignes « Total » sont ignorées ; "
           "le bloc « Filtres appliqués » est lu pour contrôler période, fournisseur et sites.")
    with c2:
        T.rules_card("📐 Méthode de calcul", [
            ("CA prévu HT", "Prévision Qté × Prix promo ÷ (1 + TVA)"),
            ("Marge prévue", "Prévision Qté × (Prix promo HT − PA)"),
            ("Réalisation", "Réalisé ÷ Prévu, en quantité, CA et marge"),
            ("Fiabilité (sur le CA)", f"{T.badge('Fiable', 'ok')} {T.badge('Surestimée', 'crit')} {T.badge('Sous-estimée', 'att')}"),
            ("Vs N-1 (sur le CA)", f"{T.badge('Mieux', 'ok')} {T.badge('Stable', 'info')} {T.badge('Moins bien', 'crit')}"),
            ("Marge réalisée", "valeur et taux, comparés à la prévision et à N-1"),
        ])
    with c3:
        T.rules_card("⚠️ Points d'attention", [
            ("Période", "l'extraction doit couvrir uniquement la durée de la promo"),
            ("Montants", "les exports PBI sont en HT : le prix promo TTC est ramené en HT"),
            ("Quantité réalisée", "quantité totale vendue (la quantité promo n'est pas dans l'export)"),
            ("Marge N-1", "reconstituée : Marge Hors Promo N-1 + Marge Promo N-1"),
            ("Sites convertis", "Cité verte, Aboboté, Toit rouge : N-1 marqué non comparable"),
            ("Exports Excel", "valeurs uniquement, aucune formule"),
        ])
    T.section("Exemple de liste promo")
    st.dataframe(pd.DataFrame({"Code": [12003029, 12003043, 12503300],
                               "Libellé": ["X42 COUCH STD PACK MINI JOYA", "X40 COUCH STD PACK MAXI SMILEY", "X55 MOUCHOIRS FACIALE 4P STYLE"],
                               "Prévision Qté": [200, 1500, 900], "Prix promo": [3900, 2950, 450], "PA": [2700, 2250, 260], "TVA": [18, 18, 18]}),
                 use_container_width=True, hide_index=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT
# ══════════════════════════════════════════════════════════════════════════════
with st.spinner("Lecture et préparation des données…"):
    try:
        sales, filters = E.parse_sales(f_sales)
    except Exception as exc:  # noqa: BLE001
        T.page_header("🎯", "Résultats Promo")
        T.alert("Extraction ventes illisible", str(exc), "crit")
        st.stop()
    try:
        promo, meta = E.parse_promo(f_promo, price_ttc=price_ttc)
    except Exception as exc:  # noqa: BLE001
        T.page_header("🎯", "Résultats Promo")
        T.alert("Liste promo illisible", str(exc) + " Colonnes attendues : Code, Prévision Qté, Prix promo, PA, TVA.", "crit")
        st.stop()
if sales.empty:
    T.page_header("🎯", "Résultats Promo")
    T.alert("Aucune ligne exploitable", "L'export ne contient aucune ligne article × site (5 axes attendus).", "crit")
    st.stop()

art, site, controls, tot = E.compute(sales, promo, params)
sentence = E.synthesis_sentence(art, tot)
rc, nc = tot["rel_counts"], tot["n1_counts"]
has_m = meta["has_pa"]
per = (f"du {filters['period_from']:%d/%m/%Y} au {filters['period_to']:%d/%m/%Y}" if filters["period_from"] else "période non lue")
T.sidebar_footer(f"{tot['n']} articles · {tot['sites']} sites<br>{per}<br>CA prévu : {meta['ca_mode']}")


# ══════════════════════════════════════════════════════════════════════════════
# BANDEAU ET KPI
# ══════════════════════════════════════════════════════════════════════════════
chips = [f"{tot['n']} articles",
         (f"{int(rc['Surestimée'])} surestimée(s)", "alert") if rc["Surestimée"] else ("Aucune surestimation", "ok"),
         f"{int(rc['Fiable'])} fiable(s)", f"{int(rc['Sous-estimée'])} sous-estimée(s)"]
if rc["Pas de vente"]:
    chips.append((f"{int(rc['Pas de vente'])} sans vente", "alert"))
chips.append((f"CA vs N-1 {E.fmt_pct(tot['ca_evo'], 0, True)}", "ok" if (tot["ca_evo"] or 0) > 0 else "alert"))
T.hero(f"CA RÉALISÉ · <b>{E.fmt_pct(tot['real_ca'], 0)}</b> DU PRÉVU",
       f"{filters['supplier'] or 'Fournisseur non lu'} · extraction {per}", chips, eyebrow="Résultats Promo")

sgn = lambda x: "pos" if (x or 0) >= params["low"] else "neg"  # noqa: E731
cards = [
    T.kpi("Quantité", E.fmt_int(tot["qty"]), E.fmt_pct(tot["real_qty"], 0), sgn(tot["real_qty"]),
          f"prévu {E.fmt_int(tot['prev_qty'])}", "blue", "📦"),
    T.kpi("CA HT", E.fmt_fcfa(tot["ca"]), E.fmt_pct(tot["real_ca"], 0), sgn(tot["real_ca"]),
          f"prévu {E.fmt_fcfa(tot['prev_ca'])}", "blue", "🧾", "FCFA"),
]
if has_m:
    cards.append(T.kpi("Marge", E.fmt_fcfa(tot["marge"]), E.fmt_pct(tot["real_marge"], 0), sgn(tot["real_marge"]),
                       f"prévu {E.fmt_fcfa(tot['prev_marge'])} · taux {E.fmt_pct(tot['rate'])} vs {E.fmt_pct(tot['prev_rate'])} prévu",
                       "violet", "💰", "FCFA"))
cards.append(T.kpi("CA promis non réalisé", E.fmt_fcfa(tot["ca_missed"]), "", "neg" if tot["ca_missed"] > 0 else "pos",
                   (f"marge promise non réalisée {E.fmt_fcfa(tot['marge_missed'])}" if has_m else "somme des écarts négatifs"),
                   "violet", "📉", "FCFA"))
T.kpi_row(cards)
T.card(f"<div style='font-size:15px;font-weight:700;line-height:1.5'>{sentence}</div>")
if not has_m:
    T.alert("Marge prévue non calculée", "Ajoutez les colonnes PA et TVA dans la liste promo pour challenger aussi la marge.", "att")


# ══════════════════════════════════════════════════════════════════════════════
# ONGLETS
# ══════════════════════════════════════════════════════════════════════════════
n_ctrl = len(controls["not_found"]) + len(controls["duplicates"]) + len(controls["no_price"]) + len(controls["neg_prev_marge"]) \
    + (len(controls["no_pa"]) if has_m else 0)
tab1, tab2, tab3, tab4 = st.tabs(["🎯 Prévu vs réalisé", "📅 Vs N-1", f"🧪 Contrôles · {n_ctrl}", "⬇️ Export"])

# ── Prévu vs réalisé ──────────────────────────────────────────────────────────
with tab1:
    c1, c2 = st.columns([2, 1])
    sel = c1.multiselect("Fiabilité", E.RELIABILITY_ORDER, default=[], placeholder="Toutes", key="rp_f_rel")
    sort = c2.selectbox("Tri", ["Écart CA (pire d'abord)", "Réalisation CA ↑", "Écart marge (pire d'abord)", "CA prévu ↓"], key="rp_sort")
    d = art[art["reliability"].isin(sel)] if sel else art
    key = {"Écart CA (pire d'abord)": ("gap_ca", True), "Réalisation CA ↑": ("real_ca", True),
           "Écart marge (pire d'abord)": ("gap_marge", True), "CA prévu ↓": ("prev_ca", False)}[sort]
    d = d.sort_values(key[0], ascending=key[1], na_position="last")

    T.section("Ce qui a été promis, ce qui est sorti")
    head = "<th>Article</th><th>Qté prévue</th><th>Qté réalisée</th><th>Réal.</th><th>CA prévu HT</th><th>CA réalisé</th><th>Réal.</th><th>Écart CA</th>"
    if has_m:
        head += "<th>Marge prévue</th><th>Marge réalisée</th><th>Réal.</th>"
    head += "<th style='text-align:right'>Fiabilité</th>"
    body = ""
    for _, r in d.iterrows():
        row = (f"<td><b>{r['label']}</b>{sm(r['code'])}</td>"
               f"<td>{E.fmt_int(r['prev_qty'])}</td><td>{E.fmt_int(r['Qté Vente'])}</td><td>{pct_badge(r['real_qty'], low, high)}</td>"
               f"<td>{E.fmt_fcfa(r['prev_ca'])}</td><td>{E.fmt_fcfa(r['CA'])}</td><td>{pct_badge(r['real_ca'], low, high)}</td>"
               f"<td>{E.fmt_fcfa(r['gap_ca'])}</td>")
        if has_m:
            row += (f"<td>{E.fmt_fcfa(r['prev_marge'])}{sm(E.fmt_pct(r['prev_rate']))}</td>"
                    f"<td>{E.fmt_fcfa(r['Marge'])}{sm(E.fmt_pct(r['rate']))}</td><td>{pct_badge(r['real_marge'], low, high)}</td>")
        row += f"<td style='text-align:right'>{rbadge(r['reliability'])}</td>"
        body += f"<tr>{row}</tr>"
    foot = (f"<tr><td><b>Total</b></td><td><b>{E.fmt_int(tot['prev_qty'])}</b></td><td><b>{E.fmt_int(tot['qty'])}</b></td>"
            f"<td>{pct_badge(tot['real_qty'], low, high)}</td><td><b>{E.fmt_fcfa(tot['prev_ca'])}</b></td><td><b>{E.fmt_fcfa(tot['ca'])}</b></td>"
            f"<td>{pct_badge(tot['real_ca'], low, high)}</td><td><b>{E.fmt_fcfa(tot['ca'] - tot['prev_ca'])}</b></td>")
    if has_m:
        foot += (f"<td><b>{E.fmt_fcfa(tot['prev_marge'])}</b>{sm(E.fmt_pct(tot['prev_rate']))}</td>"
                 f"<td><b>{E.fmt_fcfa(tot['marge'])}</b>{sm(E.fmt_pct(tot['rate']))}</td><td>{pct_badge(tot['real_marge'], low, high)}</td>")
    foot += "<td></td></tr>"
    T.card(f"<div style='overflow-x:auto'><table class='tbl'><thead><tr>{head}</tr></thead><tbody>{body}{foot}</tbody></table></div>"
           f"<div class='note'>Fiabilité jugée sur le CA : surestimée sous {E.fmt_pct(low, 0)}, sous-estimée au-delà de {E.fmt_pct(high, 0)}. "
           f"CA prévu : {meta['ca_mode']}. Montants en FCFA HT.</div>")

# ── Vs N-1 ────────────────────────────────────────────────────────────────────
with tab2:
    T.kpi_row([
        T.kpi("CA HT", E.fmt_fcfa(tot["ca"]), E.fmt_pct(tot["ca_evo"], 0, True), "pos" if (tot["ca_evo"] or 0) >= 0 else "neg",
              f"N-1 {E.fmt_fcfa(tot['ca_n1'])}", "blue", "🧾", "FCFA"),
        T.kpi("Marge", E.fmt_fcfa(tot["marge"]), E.fmt_pct(tot["marge_evo"], 0, True), "pos" if (tot["marge_evo"] or 0) >= 0 else "neg",
              f"N-1 {E.fmt_fcfa(tot['marge_n1'])}", "violet", "💰", "FCFA"),
        T.kpi("Taux de marge", E.fmt_pct(tot["rate"]).replace(" %", ""), E.fmt_pts(tot["rate"] - tot["rate_n1"]),
              "pos" if (tot["rate"] or 0) >= (tot["rate_n1"] or 0) else "neg", f"N-1 {E.fmt_pct(tot['rate_n1'])}", "teal", "📊", "%"),
        T.kpi("Articles mieux que N-1", str(int(nc["Mieux que N-1"])), f"{int(nc['Moins bien que N-1'])} moins bien", "mut",
              f"{int(nc['Stable'])} stable(s) · {int(nc['Pas de N-1'])} sans N-1", "", "📅"),
    ])
    if controls["converted_present"]:
        T.alert("N-1 partiellement non comparable",
                f"Sites convertis en 2026 présents dans l'extraction : {', '.join(controls['converted_present'])}. "
                "Les articles concernés sont marqués ; à lire avec prudence.", "att")
    T.section("CA et marge réalisés vs N-1")
    head = ("<th>Article</th><th>CA réalisé</th><th>CA N-1</th><th>Évo CA</th><th>Marge réalisée</th><th>Marge N-1</th>"
            "<th>Évo marge</th><th>Taux</th><th>Taux N-1</th><th style='text-align:right'>Lecture</th>")
    body = ""
    for _, r in art.sort_values("ca_evo", na_position="last").iterrows():
        flag = " " + T.badge("N-1 non comp.", "att") if r["n1_partial"] else ""
        body += (f"<tr><td><b>{r['label']}</b>{flag}{sm(r['code'])}</td>"
                 f"<td>{E.fmt_fcfa(r['CA'])}</td><td class='mut'>{E.fmt_fcfa(r['CA N-1'])}</td><td>{evo_badge(r['ca_evo'], band)}</td>"
                 f"<td>{E.fmt_fcfa(r['Marge'])}</td><td class='mut'>{E.fmt_fcfa(r['Marge N-1'])}</td><td>{evo_badge(r['marge_evo'], band)}</td>"
                 f"<td>{E.fmt_pct(r['rate'])}</td><td class='mut'>{E.fmt_pct(r['rate_n1'])}</td>"
                 f"<td style='text-align:right'>{nbadge(r['vs_n1'])}</td></tr>")
    body += (f"<tr><td><b>Total</b></td><td><b>{E.fmt_fcfa(tot['ca'])}</b></td><td><b>{E.fmt_fcfa(tot['ca_n1'])}</b></td>"
             f"<td>{evo_badge(tot['ca_evo'], band)}</td><td><b>{E.fmt_fcfa(tot['marge'])}</b></td><td><b>{E.fmt_fcfa(tot['marge_n1'])}</b></td>"
             f"<td>{evo_badge(tot['marge_evo'], band)}</td><td><b>{E.fmt_pct(tot['rate'])}</b></td><td><b>{E.fmt_pct(tot['rate_n1'])}</b></td><td></td></tr>")
    T.card(f"<div style='overflow-x:auto'><table class='tbl'><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>"
           f"<div class='note'>Lecture sur le CA : stable dans ±{E.fmt_pct(band, 0)}. Marge N-1 = Marge Hors Promo N-1 + Marge Promo N-1.</div>")

# ── Contrôles ─────────────────────────────────────────────────────────────────
with tab3:
    def rr(ok, lab, det):
        return (f"<div class='rule-row'><span><b>{lab}</b>{sm(det)}</span>"
                f"<span style='align-self:center'>{T.badge('OK', 'ok') if ok else T.badge('À vérifier', 'att')}</span></div>")
    days = (filters["period_to"] - filters["period_from"]).days + 1 if filters["period_from"] else None
    T.section("Périmètre de l'extraction")
    T.card(rr(days is not None and days <= 31, "Période",
              f"{per}" + (f", soit {days} jours. Au-delà de la durée de la promo, les résultats sont dilués." if days else ""))
           + rr(bool(filters["supplier"]), "Fournisseur", filters["supplier"] or "Non lu.")
           + rr(True, "Sites", f"{tot['sites']} sites" + (f" : {', '.join(filters['sites'])}" if filters["sites"] else "")))
    T.section("Liste promo")
    T.card(rr(not controls["not_found"], "Codes trouvés dans l'extraction",
              f"{int(art['found'].sum())} / {len(art)}" + (f" · introuvables : {', '.join(controls['not_found'])}" if controls["not_found"] else ""))
           + rr(not controls["duplicates"], "Doublons", ", ".join(controls["duplicates"]) or "Aucun.")
           + rr(not controls["no_price"], "Prix promo manquant", ", ".join(controls["no_price"]) or "Aucun.")
           + (rr(not controls["no_pa"], "PA manquant", ", ".join(controls["no_pa"]) or "Aucun.") if has_m else
              rr(False, "PA et TVA", "Colonnes absentes : la marge prévue n'est pas calculée."))
           + rr(not controls["neg_prev_marge"], "Marge prévue négative (prix promo HT < PA)", ", ".join(controls["neg_prev_marge"]) or "Aucune.")
           + rr(meta["has_vat"] or not price_ttc, "TVA", "Prix ramené en HT avec la TVA de chaque article." if meta["has_vat"] and price_ttc
                else "Prix promo déclaré HT." if not price_ttc else "Colonne TVA absente : prix promo pris tel quel, CA prévu surévalué si le prix est TTC."))

# ── Export ────────────────────────────────────────────────────────────────────
with tab4:
    T.section("Classeur Excel · valeurs uniquement")
    T.card("<div style='font-size:13px;line-height:1.55'>Onglets : <b>Synthèse</b>, <b>Fiabilité prévisions</b> (prévu vs réalisé en "
           "quantité, CA et marge), <b>Vs N-1</b> (CA et marge en valeur et taux), <b>Détail sites</b>, <b>Contrôles</b>.</div>"
           "<div class='note'>Généré avec les règles de la barre latérale · aucune formule.</div>")
    if st.button("Générer le classeur", type="primary", key="rp_btn_xlsx", use_container_width=True):
        with st.spinner("Génération du classeur…"):
            st.session_state["rp_xlsx"] = E.build_excel(art, site, controls, tot, filters, meta, params, sentence)
            sup = re.sub(r"[^A-Za-z0-9]+", "_", (filters["supplier"] or "promo").split(" - ")[-1]).strip("_")
            st.session_state["rp_xlsx_name"] = f"SmartBuyer_Resultats_Promo_{sup}_{datetime.now():%Y%m%d}.xlsx"
    if st.session_state.get("rp_xlsx"):
        st.download_button("⬇️ Télécharger le classeur", data=st.session_state["rp_xlsx"],
                           file_name=st.session_state["rp_xlsx_name"], mime=XLSX, key="rp_dl", use_container_width=True)
