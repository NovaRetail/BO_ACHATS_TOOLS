"""SmartBuyer Hub · Résultats Promo — réalisé vs prévision, marge dégagée, verdict par article.

Page de pilotage (usage personnel) : on charge l'extraction ventes PBI de la période promo et la
liste promo (Code + prévision), on lit la synthèse, les verdicts et les actions par article, puis
on télécharge le classeur Excel (valeurs uniquement, aucune formule) avec un onglet « Bilan
fournisseur » sans marge.

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


def vbadge(verdict: str) -> str:
    return T.badge(verdict, E.VERDICTS[verdict]["kind"])


def table(df):
    st.dataframe(df, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# BARRE LATÉRALE
# ══════════════════════════════════════════════════════════════════════════════
T.sidebar_brand("🎯", "Résultats Promo")
T.sidebar_label("1 · Extraction ventes")
with st.sidebar:
    f_sales = st.file_uploader("Export PBI ventes par article (xlsx / csv)", type=["xlsx", "xlsm", "csv"], key="rp_sales",
                               help="Export hiérarchique Departement › Rayon › Famille › Article › Site, filtré sur la période de la promo.")
T.sidebar_label("2 · Liste promo")
with st.sidebar:
    f_promo = st.file_uploader("Articles et prévisions (xlsx / csv)", type=["xlsx", "xlsm", "csv"], key="rp_promo",
                               help="Colonnes : Code, Prévision (qté) et/ou Prévision CA ; Libellé, Site, Financement facultatifs.")
T.sidebar_label("3 · Règles de verdict")
with st.sidebar:
    low = st.slider("Sous-performance en dessous de", 50, 100, 80, 5, format="%d %%") / 100
    high = st.slider("Succès à partir de", 100, 200, 120, 5, format="%d %%") / 100
    over = st.slider("Prévision sous-estimée au-delà de", 120, 500, 200, 10, format="%d %%") / 100
    tol = st.number_input("Tolérance marge (points)", 0.0, 10.0, 1.0, 0.5,
                          help="Écart admis entre le taux de marge promo et le taux hors promo des mêmes sites.") / 100
    converted = st.multiselect("Sites au N-1 non comparable", E.DEFAULT_CONVERTED_SITES, default=E.DEFAULT_CONVERTED_SITES)
params = {"low": low, "high": max(high, low), "over": max(over, high), "margin_tol": tol, "converted_sites": converted}


# ══════════════════════════════════════════════════════════════════════════════
# ACCUEIL (avant dépôt des fichiers)
# ══════════════════════════════════════════════════════════════════════════════
if f_sales is None or f_promo is None:
    T.sidebar_footer("Réalisé = quantité totale vendue<br>Marge de référence = hors promo des mêmes sites")
    rows = [("Extraction ventes PBI", "ok" if f_sales else "wait", "reçue" if f_sales else "à déposer"),
            ("Liste promo avec prévisions", "ok" if f_promo else "wait", "reçue" if f_promo else "à déposer")]
    T.landing_hero("RÉSULTATS <b>PROMO</b>",
                   "La promo a-t-elle fait le volume prévu, combien de marge a-t-elle dégagée et coûté, "
                   "et que doit faire l'acheteur article par article.",
                   eyebrow="Résultats Promo · SmartBuyer Hub",
                   status=T.landing_status("Préparation de l'analyse", rows, "fichiers obligatoires reçus", count=2))
    T.steps([("Extraire", "les ventes par article × site dans Power BI, sur la seule période de la promo"),
             ("Lister", "les articles de la promo avec leur prévision (quantité et/ou CA)"),
             ("Déposer", "les deux fichiers dans la barre latérale"),
             ("Lire", "la synthèse, les verdicts et les actions recommandées"),
             ("Exporter", "le classeur Excel et le bilan à partager avec le fournisseur")])
    T.section("Format des fichiers attendus")
    c1, c2, c3 = st.columns(3)
    with c1:
        T.columns_card("Colonnes attendues", [
            ("Extraction ventes (export PBI)", ["Departement", "Rayon", "Famille", "Article", "Site nom long", "CA", "Marge",
                                                "CA Promo", "Marge Promo", "CA Hors Promo", "Marge Hors Promo", "Qté Vente",
                                                "CA N-1 · Qté Vente N-1 · Marge Hors Promo N-1 · Marge Promo N-1 (facultatives)"]),
            ("Liste promo", ["Code", "Prévision (quantité totale, tous sites)", "Prévision CA (facultative)",
                             "Libellé · Site · Financement (facultatives)"]),
        ], "Article au format « 12003043 - LIBELLÉ » : le code avant le tiret sert de clé de jointure. "
           "Les lignes « Total » de l'export sont ignorées ; le bloc « Filtres appliqués » en pied de fichier est lu pour "
           "contrôler période, fournisseur et sites.")
    with c2:
        T.rules_card("📐 Méthode de calcul", [
            ("Réalisation", "Qté vendue ÷ Prévision (ou CA ÷ Prévision CA si seule la prévision CA est fournie)"),
            ("Marge dégagée", "Marge totale de l'article sur la période, dont marge promo"),
            ("Taux de référence", "Taux de marge hors promo des sites où la promo a tourné"),
            ("Marge cédée", "CA promo × (taux de référence − taux promo) : le coût de la remise"),
            ("Marge nette", "Marge promo + Financement fournisseur, si la colonne est fournie"),
            ("Verdicts", f"{T.badge('Succès', 'ok')} {T.badge('Conforme', 'ok')} {T.badge('Sous-estimée', 'info')} "
                         f"{T.badge('Volume sans marge', 'att')} {T.badge('Non activée', 'att')} {T.badge('Sous-perf.', 'crit')}"),
        ])
    with c3:
        T.rules_card("⚠️ Points d'attention", [
            ("Période", "l'extraction doit couvrir uniquement la durée de la promo, sinon les résultats sont dilués"),
            ("Quantité promo", "absente de l'export PBI : le réalisé est la quantité totale vendue"),
            ("Marge N-1", "reconstituée : Marge Hors Promo N-1 + Marge Promo N-1"),
            ("Sites convertis", "Cité verte, Aboboté, Toit rouge : N-1 marqué non comparable"),
            ("Casse", "non utilisée (devise erronée dans l'export)"),
            ("Exports Excel", "valeurs uniquement, aucune formule ; bilan fournisseur sans marge"),
        ])
    T.section("Exemple de liste promo")
    table(pd.DataFrame({"Code": [12003029, 12003043, 12503300],
                        "Libellé": ["X42 COUCH STD PACK MINI JOYA", "X40 COUCH STD PACK MAXI SMILEY", "X55 MOUCHOIRS FACIALE 4P STYLE"],
                        "Prévision": [10, 30, 80], "Prévision CA": [35000, 80000, 30000], "Financement": [None, 50000, None]}))
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
        promo, meta = E.parse_promo(f_promo)
    except Exception as exc:  # noqa: BLE001
        T.page_header("🎯", "Résultats Promo")
        T.alert("Liste promo illisible", str(exc), "crit")
        st.stop()

if sales.empty:
    T.page_header("🎯", "Résultats Promo")
    T.alert("Aucune ligne exploitable", "L'export ne contient aucune ligne article × site. Vérifiez qu'il est bien "
            "hiérarchique (5 axes) avec les lignes de détail par site.", "crit")
    st.stop()

art, site, controls, tot = E.compute(sales, promo, meta, params)
sentence = E.synthesis_sentence(art, tot, meta)
vc = tot["verdict_counts"]
per = (f"du {filters['period_from']:%d/%m/%Y} au {filters['period_to']:%d/%m/%Y}" if filters["period_from"] else "période non lue")
T.sidebar_footer(f"{len(art)} articles · {tot['sites_total']} sites<br>{per}")

n_alerts = (len(controls["not_found"]) + len(controls["duplicates"]) + (1 if len(controls["partial"]) else 0)
            + (1 if controls["converted_present"] else 0) + (1 if len(controls["promo_not_listed"]) else 0))


# ══════════════════════════════════════════════════════════════════════════════
# BANDEAU ET KPI
# ══════════════════════════════════════════════════════════════════════════════
chips = [f"{len(art)} articles", f"{tot['sites_promo']} / {tot['sites_total']} sites en promo"]
for v in E.VERDICT_ORDER:
    if vc[v]:
        kind = {"ok": "ok", "crit": "alert", "att": "alert"}.get(E.VERDICTS[v]["kind"], "")
        chips.append((f"{int(vc[v])} {v.lower()}", kind))
chips.append((f"{n_alerts} alerte(s) de périmètre", "alert") if n_alerts else ("Périmètre OK", "ok"))
T.hero(f"RÉALISATION · <b>{E.fmt_pct(tot['real'], 0)}</b>",
       f"{filters['supplier'] or 'Fournisseur non lu'} · extraction {per}", chips, eyebrow="Résultats Promo")

real_sign = "pos" if (tot["real"] or 0) >= params["high"] else ("neg" if (tot["real"] or 0) < params["low"] else "mut")
cards = [
    T.kpi("Qté vendue vs prévue" if meta["has_qty"] else "CA vs prévu",
          E.fmt_int(tot["qty"]) if meta["has_qty"] else E.fmt_fcfa(tot["ca"]),
          E.fmt_pct(tot["real"], 0), real_sign,
          f"prévu {E.fmt_int(tot['prev_qty']) if meta['has_qty'] else E.fmt_fcfa(tot['prev_ca'])}", "blue", "📦",
          "" if meta["has_qty"] else "FCFA"),
    T.kpi("Marge dégagée", E.fmt_fcfa(tot["marge"]), f"dont {E.fmt_fcfa(tot['marge_promo'])} en promo", "mut",
          f"{E.fmt_pct(E.safe_div(tot['marge_promo'], tot['marge']), 0)} de la marge en promo", "violet", "💰", "FCFA"),
    T.kpi("Coût de la promo", E.fmt_fcfa(tot["ceded"]), f"{E.fmt_pct(tot['rate_promo'])} vs {E.fmt_pct(tot['rate_hp'])}",
          "neg" if tot["ceded"] > 0 else "pos", "marge cédée par la remise · taux promo vs hors promo", "violet", "🏷️", "FCFA"),
    T.kpi("CA promo", E.fmt_fcfa(tot["ca_promo"]), f"{E.fmt_pct(tot['promo_share'], 0)} du CA", "mut",
          f"CA total {E.fmt_fcfa(tot['ca'])}", "teal", "🧾", "FCFA"),
]
if meta["has_funding"] and not pd.isna(tot["funding"]):
    cards.insert(3, T.kpi("Financement fournisseur", E.fmt_fcfa(tot["funding"]),
                          f"marge nette {E.fmt_fcfa(tot['marge_promo'] + tot['funding'])}", "pos",
                          "marge promo + financement", "teal", "🤝", "FCFA"))
T.kpi_row(cards)
T.card(f"<div style='font-size:15px;font-weight:700;line-height:1.5'>{sentence}</div>")


# ══════════════════════════════════════════════════════════════════════════════
# ONGLETS
# ══════════════════════════════════════════════════════════════════════════════
tab0, tab1, tab2, tab3, tab4 = st.tabs([
    "🎯 Synthèse", f"📦 Articles · {len(art)}", "🏪 Drill-down site", f"🧪 Contrôles · {n_alerts}", "⬇️ Exports",
])

# ── Synthèse ──────────────────────────────────────────────────────────────────
with tab0:
    left, right = st.columns([1, 1.35])
    with left:
        T.section("Répartition des verdicts")
        n = max(1, len(art))
        body = ""
        for v in E.VERDICT_ORDER:
            k = int(vc[v])
            body += (f"<div style='display:grid;grid-template-columns:minmax(0,1.3fr) minmax(0,1fr) auto;gap:14px;align-items:center;"
                     f"padding:8px 0;border-bottom:1px solid {T.LIGNE}'><div style='font-weight:600;font-size:13.5px'>{v}</div>"
                     f"<div style='background:{T.GRIS_CLAIR};border-radius:6px;height:8px'><div style='width:{max(k / n * 100, 2 if k else 0):.0f}%;"
                     f"height:8px;border-radius:6px;background:#{E.VERDICTS[v]['hex']}'></div></div>"
                     f"<div style='font-weight:700;min-width:22px;text-align:right'>{k}</div></div>")
        T.card(body)

        T.section("Top 5 · marge dégagée")
        top = art.sort_values("Marge", ascending=False).head(5)
        T.card("".join(f"<div class='rule-row'><span>{r['label']}<small style='display:block;color:{T.GRIS_TEXTE}'>"
                       f"{E.fmt_fcfa(r['Marge'])} FCFA · {E.fmt_int(r['Qté Vente'])} vendus</small></span>"
                       f"<span style='align-self:center'>{vbadge(r['verdict'])}</span></div>" for _, r in top.iterrows()))
        T.section("Flop 5 · marge dégagée")
        flop = art[art["found"]].sort_values("Marge").head(5)
        T.card("".join(f"<div class='rule-row'><span>{r.label}<small style='display:block;color:{T.GRIS_TEXTE}'>"
                       f"{E.fmt_fcfa(r.Marge)} FCFA</small></span><span style='align-self:center'>{vbadge(r.verdict)}</span></div>"
                       for r in flop.itertuples()))
    with right:
        T.section("À traiter en priorité")
        prio = art[~art["verdict"].isin(["Conforme", "Succès"])].sort_values(["verdict_rank", "CA"], ascending=[True, False]).head(8)
        if prio.empty:
            T.alert("Aucun point d'attention", "Tous les articles sont conformes ou en succès.", "ok")
        for _, r in prio.iterrows():
            q, q1, v = r["Qté Vente"], r["Qté Vente N-1"], r["verdict"]
            if v == "Prévision sous-estimée":
                txt = f"{E.fmt_int(q)} vendus pour {E.fmt_int(r['prev_qty'])} prévus ({E.fmt_pct(r['real'], 0)}) : risque de rupture."
            elif v == "Sous-performance":
                txt = f"{E.fmt_int(q)} vendus pour {E.fmt_int(r['prev_qty'])} prévus ({E.fmt_pct(r['real'], 0)})" + (f", contre {E.fmt_int(q1)} en N-1." if q1 else ".")
            elif v == "Promo non activée":
                txt = f"{E.fmt_int(q)} vendus mais aucun CA promo sur {int(r['n_sites_total'])} sites."
            elif v == "Pas de vente":
                txt = "Aucune vente sur la période" + (f" ({E.fmt_int(q1)} en N-1)." if q1 else ".")
            elif v == "Volume sans marge":
                txt = (f"Volume atteint mais {E.fmt_fcfa(r['margin_ceded'])} FCFA de marge cédée "
                       f"({E.fmt_pts(r['margin_delta_pts'])} vs hors promo des mêmes sites).")
            else:
                txt = f"Marge promo {E.fmt_fcfa(r['Marge Promo'])} FCFA ({E.fmt_pct(r['rate_promo'])})."
            kind = E.VERDICTS[v]["kind"]
            T.alert(f"{r['label']} · {v}", f"{txt}<br><b>Action :</b> {r['action']}", kind)

# ── Articles ──────────────────────────────────────────────────────────────────
with tab1:
    c1, c2 = st.columns([2, 1])
    sel = c1.multiselect("Verdict", E.VERDICT_ORDER, default=[], placeholder="Tous les verdicts", key="rp_f_verdict")
    sort = c2.selectbox("Tri", ["Marge dégagée ↓", "Réalisation ↓", "Réalisation ↑", "CA promo ↓", "Marge cédée ↓", "Verdict"], key="rp_sort")
    d = art[art["verdict"].isin(sel)] if sel else art
    key = {"Marge dégagée ↓": ("Marge", False), "Réalisation ↓": ("real", False), "Réalisation ↑": ("real", True),
           "CA promo ↓": ("CA Promo", False), "Marge cédée ↓": ("margin_ceded", False), "Verdict": ("verdict_rank", True)}[sort]
    d = d.sort_values(key[0], ascending=key[1], na_position="last")

    T.section("Réalisé, marge et action par article")
    head = "<th>Article</th>"
    if meta["has_qty"]:
        head += "<th>Prévu</th><th>Vendu</th><th>Réal.</th>"
    if meta["has_ca"]:
        head += "<th>CA prévu</th><th>CA</th><th>Réal. CA</th>"
    head += "<th>Marge dégagée</th><th>Marge promo</th><th>Marge cédée</th>"
    if meta["has_funding"]:
        head += "<th>Financ.</th><th>Marge nette</th>"
    head += "<th>Marge vs N-1</th><th>Sites promo</th><th>Verdict</th><th>Action recommandée</th>"
    body = ""
    for _, r in d.iterrows():
        row = f"<td><b>{r['label']}</b><small style='display:block;color:{T.GRIS_TEXTE}'>{r['code']}</small></td>"
        if meta["has_qty"]:
            row += f"<td>{E.fmt_int(r['prev_qty'])}</td><td>{E.fmt_int(r['Qté Vente'])}</td><td>{E.fmt_pct(r['real_qty'], 0)}</td>"
        if meta["has_ca"]:
            row += f"<td>{E.fmt_fcfa(r['prev_ca'])}</td><td>{E.fmt_fcfa(r['CA'])}</td><td>{E.fmt_pct(r['real_ca'], 0)}</td>"
        mp = (f"{E.fmt_fcfa(r['Marge Promo'])}<small style='display:block;color:{T.GRIS_TEXTE}'>{E.fmt_pct(r['rate_promo'])} "
              f"({E.fmt_pts(r['margin_delta_pts'])})</small>") if r["CA Promo"] > 0 else "—"
        row += f"<td>{E.fmt_fcfa(r['Marge'])}</td><td>{mp}</td><td>{E.fmt_fcfa(r['margin_ceded']) if r['margin_ceded'] else '—'}</td>"
        if meta["has_funding"]:
            row += f"<td>{E.fmt_fcfa(r['funding'])}</td><td>{E.fmt_fcfa(r['margin_net'])}</td>"
        sp = f"{int(r['n_sites_promo'])} / {int(r['n_sites_total'])}" + (" · partielle" if r["promo_partial"] else "")
        row += (f"<td>{E.fmt_fcfa(r['marge_evo_n1']) if r['marge_evo_n1'] else '—'}</td><td>{sp}</td><td>{vbadge(r['verdict'])}</td>"
                f"<td class='mut' style='white-space:normal;min-width:220px'>{r['action']}</td>")
        body += f"<tr>{row}</tr>"
    T.card(f"<div style='overflow-x:auto'><table class='tbl'><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>"
           "<div class='note'>Marge promo : valeur, puis taux et écart en points vs le taux hors promo des mêmes sites. "
           "Marge cédée = CA promo × (taux hors promo − taux promo).</div>")

# ── Drill-down site ───────────────────────────────────────────────────────────
with tab2:
    opts = {f"{r.label} ({r.code})": r.code for r in art[art["found"]].itertuples()}
    if not opts:
        T.alert("Aucun article trouvé dans l'extraction", kind="att")
    else:
        choice = st.selectbox("Article", list(opts), key="rp_drill")
        a = art[art["code"] == opts[choice]].iloc[0]
        ds = site[site["code"] == opts[choice]].sort_values(["promo_active", "CA Promo", "Qté Vente"], ascending=[False, False, False])
        fchips = [f"{int(a['n_sites_promo'])} / {int(a['n_sites_total'])} sites en promo", f"{E.fmt_int(a['Qté Vente'])} vendus",
                  f"{E.fmt_fcfa(a['Marge'])} FCFA de marge"]
        if (~ds["n1_comparable"]).any():
            fchips.append(("N-1 partiellement non comparable", "alert"))
        T.hero(f"<b>{a['label']}</b>", f"{vbadge(a['verdict'])} · {a['action']}", fchips, eyebrow="Drill-down site")
        body = ""
        for _, r in ds.iterrows():
            mp = (E.fmt_fcfa(r["Marge Promo"]) + " (" + E.fmt_pct(r["rate_promo"]) + ")") if r["CA Promo"] > 0 else "—"
            body += (f"<tr><td>{r['site']}{' ' + T.badge('N-1 non comp.', 'att') if not r['n1_comparable'] else ''}</td>"
                     f"<td>{E.fmt_int(r['Qté Vente'])}</td><td class='mut'>{E.fmt_int(r['Qté Vente N-1'])}</td>"
                     f"<td>{E.fmt_pct(r['qty_evo'], 0, True)}</td><td>{E.fmt_fcfa(r['CA'])}</td>"
                     f"<td>{E.fmt_fcfa(r['CA Promo']) if r['CA Promo'] else '—'}</td><td>{E.fmt_fcfa(r['Marge'])}</td>"
                     f"<td>{mp}</td><td class='mut'>{E.fmt_pct(r['rate_hp'])}</td>"
                     f"<td style='text-align:right'>{T.badge('Active', 'ok') if r['promo_active'] else T.badge('Absente', 'info')}</td></tr>")
        T.card("<table class='tbl'><thead><tr><th>Site</th><th>Qté</th><th>Qté N-1</th><th>Évo</th><th>CA</th><th>CA promo</th>"
               "<th>Marge</th><th>Marge promo</th><th>Taux hors promo</th><th style='text-align:right'>Promo</th></tr></thead>"
               f"<tbody>{body}</tbody></table>")

# ── Contrôles ─────────────────────────────────────────────────────────────────
with tab3:
    def rule_rows(items):
        return "".join(f"<div class='rule-row'><span><b>{lab}</b><small style='display:block;color:{T.GRIS_TEXTE}'>{det}</small></span>"
                       f"<span style='align-self:center'>{T.badge(txt, kind)}</span></div>" for kind, txt, lab, det in items)

    if filters["period_from"]:
        days = (filters["period_to"] - filters["period_from"]).days + 1
        per_item = ("att" if days > 31 else "ok", "À vérifier" if days > 31 else "OK", "Période",
                    f"{per}, soit {days} jours." + (" Si la promo a duré moins longtemps, filtrez l'extraction sur sa fenêtre." if days > 31 else ""))
    else:
        per_item = ("att", "À vérifier", "Période", "Bloc « Filtres appliqués » absent : période non lue.")
    T.section("Périmètre de l'extraction")
    T.card(rule_rows([per_item,
                      ("ok" if filters["supplier"] else "att", "OK" if filters["supplier"] else "À vérifier", "Fournisseur", filters["supplier"] or "Non lu."),
                      ("ok", "OK", "Sites", f"{tot['sites_total']} sites" + (f" : {', '.join(filters['sites'])}" if filters["sites"] else ""))]))
    pnl = controls["promo_not_listed"]
    T.section("Qualité de la liste promo")
    T.card(rule_rows([
        ("ok" if not controls["not_found"] else "att", "OK" if not controls["not_found"] else "À vérifier", "Jointure des codes",
         f"{int(art['found'].sum())} codes sur {len(art)} trouvés (codes normalisés)." + (f" Introuvables : {', '.join(controls['not_found'])}." if controls["not_found"] else "")),
        ("ok" if not controls["duplicates"] else "att", "OK" if not controls["duplicates"] else "À vérifier", "Doublons",
         "Aucun." if not controls["duplicates"] else f"{', '.join(controls['duplicates'])} (prévisions sommées)."),
        ("ok" if pnl.empty else "att", "OK" if pnl.empty else "À vérifier", "Articles en promo non listés",
         "Aucun autre article n'a de CA promo." if pnl.empty else f"{len(pnl)} articles ont du CA promo sans figurer dans la liste (détail ci-dessous)."),
        ("ok", "OK", "Prévisions fournies", (("Quantité" if meta["has_qty"] else "") + (" et CA" if meta["has_qty"] and meta["has_ca"] else "CA" if meta["has_ca"] else "")
                                            + (" · financement" if meta["has_funding"] else "") + (" · par site" if meta["has_site"] else ""))),
    ]))
    part = controls["partial"]
    T.section("Fiabilité des comparaisons")
    T.card(rule_rows([
        ("ok" if part.empty else "att", "OK" if part.empty else "À vérifier", "Promo partielle",
         "Toutes les promos actives le sont sur tous les sites." if part.empty else
         ", ".join(f"{r.label} ({int(r.n_sites_promo)}/{int(r.n_sites_total)})" for r in part.itertuples()) + ". Taux de référence calculé sur les mêmes sites."),
        ("ok" if not controls["converted_present"] else "att", "OK" if not controls["converted_present"] else "À vérifier", "N-1 non comparable",
         "Aucun site converti." if not controls["converted_present"] else f"Sites convertis en 2026 : {', '.join(controls['converted_present'])} ; évolutions N-1 marquées."),
    ]))
    if not pnl.empty:
        with st.expander(f"🔍 {len(pnl)} article(s) en promo absents de la liste"):
            table(pnl.rename(columns={"code": "Code", "label": "Article", "CA Promo": "CA promo"}))

# ── Exports ───────────────────────────────────────────────────────────────────
with tab4:
    T.section("Classeur Excel · valeurs uniquement")
    T.card("<div style='font-size:13px;line-height:1.55'>Cinq onglets : <b>Synthèse</b> (KPI, verdicts, paramètres), "
           "<b>Détail articles</b>, <b>Détail sites</b>, <b>Contrôles</b> et <b>Bilan fournisseur</b> (volumes, CA et sites, "
           "sans aucune marge, prêt à partager pour la négociation).</div>"
           "<div class='note'>Généré sur les seuils de la barre latérale · aucune formule.</div>")
    if st.button("Générer le classeur", type="primary", key="rp_btn_xlsx", use_container_width=True):
        with st.spinner("Génération du classeur…"):
            st.session_state["rp_xlsx"] = E.build_excel(art, site, controls, tot, filters, meta, params, sentence)
            sup = re.sub(r"[^A-Za-z0-9]+", "_", (filters["supplier"] or "promo").split(" - ")[-1]).strip("_")
            st.session_state["rp_xlsx_name"] = f"SmartBuyer_Resultats_Promo_{sup}_{datetime.now():%Y%m%d}.xlsx"
    if st.session_state.get("rp_xlsx"):
        st.download_button("⬇️ Télécharger le classeur", data=st.session_state["rp_xlsx"],
                           file_name=st.session_state["rp_xlsx_name"], mime=XLSX, key="rp_dl", use_container_width=True)
