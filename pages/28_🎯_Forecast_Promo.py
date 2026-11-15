"""SmartBuyer Hub · Forecast Promo — challenger les prévisions des acheteurs, avant et après la promo.

Page de pilotage (usage personnel). Deux temps, mêmes fichiers :
  ① Avant la promo : la prévision de chaque article est confrontée au CA N-1 de la même période.
  ② Après la promo : prévu vs réalisé sans compensation entre articles, et comparaison à N-1.
Lecture en matrice dépliable Département › Rayon › Famille › Sous-famille › Article.
Export Excel : valeurs uniquement, aucune formule.

Les calculs et l'export sont dans modules/forecast_promo/engine.py.
"""
from __future__ import annotations

import html as _html
import pathlib
import re
import sys
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core import theme as T  # noqa: E402
from modules.forecast_promo import engine as E  # noqa: E402

st.set_page_config(page_title="Forecast Promo · SmartBuyer", page_icon="🎯", layout="wide", initial_sidebar_state="expanded")
T.apply()

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ══════════════════════════════════════════════════════════════════════════════
# CACHE (les fichiers ne sont relus que s'ils changent)
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner="Lecture de l'extraction…")
def load_sales(data: bytes, name: str):
    return E.parse_sales(data, name)


@st.cache_data(show_spinner=False)
def load_promo(data: bytes, name: str, price_ttc: bool):
    return E.parse_promo(data, name, price_ttc)


@st.cache_data(show_spinner=False)
def run(sales, promo, params):
    return E.compute(sales, promo, dict(params))


@st.cache_data(show_spinner=False)
def export(a, tot, controls, filters, meta, params):
    return E.build_excel(a, tot, controls, filters, meta, dict(params))


# ══════════════════════════════════════════════════════════════════════════════
# MATRICE DÉPLIABLE (même rendu que les tableaux du thème)
# ══════════════════════════════════════════════════════════════════════════════
KIND_CSS = {"ok": (f"{T.TEAL}17", T.VERT_TEXTE), "att": (f"{T.ORANGE}1F", T.ORANGE_TEXTE),
            "crit": (f"{T.RED}14", T.ROUGE_TEXTE), "info": ("#24235C12", T.NAVY)}


def b(text, kind):
    bg, fg = KIND_CSS[kind]
    return f"<span class='b' style='background:{bg};color:{fg}'>{text}</span>"


def b_real(x, low, high):
    return b("—", "info") if E._nan(x) else b(E.fmt_pct(x), "crit" if x < low else "att" if x > high else "ok")


def b_ratio(x, watch, justify):
    return b("pas de N-1", "att") if E._nan(x) or np.isinf(x) else b(E.fmt_ratio(x), "crit" if x > justify else "att" if x > watch else "ok")


def b_evo(x, band):
    return b("—", "info") if E._nan(x) else b(E.fmt_pct(x, 0, True), "ok" if x > band else "crit" if x < -band else "info")


def b_bias(x):
    return b("—", "info") if E._nan(x) else b(E.fmt_pct(x), "crit" if x >= .5 else "att" if x >= .25 else "ok")


def rate_line(t, r, p, n1):
    f = lambda v: "—" if E._nan(v) else f"{v * 100:.1f} %".replace(".", ",")  # noqa: E731
    return f"<small>taux {f(r)} · prévu {f(p)} · N-1 {f(n1)}</small>"


@st.cache_data(show_spinner=False)
def matrix_html(a: pd.DataFrame, kind: str, params) -> str:
    p = dict(params)
    levels = [k for k, _ in E.LEVELS]
    rows, n = [], [0]

    def cells(t, art=None):
        if kind == "avant":
            last = (b(art["control"], E.CONTROL[art["control"]]) if art is not None else
                    b(f"{t['justify']} à justifier", "crit") if t["justify"] else b(f"{t['watch']} à surveiller", "att") if t["watch"] else b("RAS", "ok"))
            return (f"<td>{E.fmt_fcfa(t['prev_ca'])}</td><td class='mut'>{E.fmt_fcfa(t['ca_n1'])}</td>"
                    f"<td>{b_ratio(t['ratio'], p['watch'], p['justify'])}</td><td>{E.fmt_fcfa(t['prev_ca'] - t['ca_n1'])}</td><td>{last}</td>")
        bias = (b(art["result"], E.RESULT[art["result"]]) if art is not None else b_bias(t["bias"]))
        return (f"<td>{E.fmt_fcfa(t['prev_ca'])}</td><td><b>{E.fmt_fcfa(t['ca'])}</b></td><td>{b_real(t['real_ca'], p['low'], p['high'])}</td>"
                f"<td class='neg'>{E.fmt_fcfa(t['short']) if t['short'] else '—'}</td><td>{bias}</td>"
                f"<td class='mut'>{E.fmt_fcfa(t['ca_n1'])}</td><td>{b_evo(t['ca_evo'], p['stable'])}</td>"
                f"<td>{E.fmt_fcfa(t['marge'])}{rate_line(t, t['rate'], t['prev_rate'], t['rate_n1'])}</td>")

    def sub(t, art=None):
        if art is not None:
            extra = " · N-1 non comparable" if art["n1_partial"] else ""
            return f"<small>{art['code']}{extra}</small>"
        if kind == "avant":
            return f"<small>{t['n']} art. · {t['justify']} à justifier · {t['nohist']} sans N-1</small>"
        return f"<small>{t['n']} art. · {t['over']} surestimé(s) ou sans vente</small>"

    def add(level, label, t, parent, kids, art=None, total=False):
        n[0] += 1
        rid = f"r{n[0]}"
        cls = f"l{level}" + (" tot" if total else "") + (" hide" if level > 0 else "")
        tog = "<span class='tg'>▸</span>" if kids else "<span class='tg'></span>"
        on = " onclick='tgl(this)'" if kids else ""
        rows.append(f"<tr id='{rid}' class='{cls}' data-p='{parent}' data-lvl='{level}'{on}>"
                    f"<td class='lab' style='padding-left:{12 + level * 20}px'>{tog}<span class='nm'>{_html.escape(str(label))}</span>{sub(t, art)}</td>"
                    f"{cells(t, art)}</tr>")
        return rid

    def rec(d, depth, parent):
        if depth == len(levels):
            key, asc = ("ratio", False) if kind == "avant" else ("short", False)
            for r in d.sort_values(key, ascending=asc, na_position="first").to_dict("records"):
                t = {"prev_ca": r["prev_ca"], "ca_n1": r["CA N-1"], "ratio": r["ratio"], "ca": r["CA"], "real_ca": r["real_ca"],
                     "short": r["short"], "ca_evo": r["ca_evo"], "marge": r["Marge"], "rate": r["rate"], "prev_rate": r["prev_rate"],
                     "rate_n1": r["rate_n1"]}
                add(depth, r["label"], t, parent, False, art=r)
            return
        sort_key = (lambda g: -(g["control"] == "À justifier").sum()) if kind == "avant" else (lambda g: -g["short"].sum())
        for lab, g in sorted(d.groupby(levels[depth], sort=False), key=lambda kv: sort_key(kv[1])):
            rec(g, depth + 1, add(depth, lab, E.summarize(g), parent, True))

    add(0, "Total opération", E.summarize(a), "", False, total=True)
    rec(a, 0, "")
    head = ("<th>Niveau</th><th>CA prévu HT</th><th>CA N-1</th><th>Prévu / N-1</th><th>Écart vs N-1</th><th>Contrôle</th>" if kind == "avant" else
            "<th>Niveau</th><th>CA prévu HT</th><th>CA réalisé</th><th>Réal.</th><th>CA promis non réalisé</th><th>Biais</th><th>CA N-1</th><th>Vs N-1</th><th>Marge</th>")
    return f"""<html><head><meta charset='utf-8'>
<link href='https://fonts.googleapis.com/css2?family=Inter:wght@500;600;700;800&display=swap' rel='stylesheet'>
<style>
body{{margin:0;font-family:{T.PLOTLY_FONT};color:{T.NAVY};background:#fff}}
.bar{{display:flex;justify-content:flex-end;gap:6px;padding:4px 2px 8px}}
.btn{{font:inherit;font-size:12.5px;font-weight:700;border:1px solid {T.LIGNE};background:#fff;color:{T.NAVY};border-radius:9px;padding:5px 12px;cursor:pointer}}
table{{width:100%;border-collapse:separate;border-spacing:0;font-size:13px;min-width:900px}}
th{{font-size:10.5px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:{T.GRIS_TEXTE};text-align:right;padding:9px 8px;border-bottom:1px solid {T.LIGNE};position:sticky;top:0;background:#fff;z-index:1}}
th:first-child{{text-align:left}}
td{{text-align:right;padding:8px;font-weight:600;color:{T.NAVY};border-bottom:1px solid #F0F1F5;white-space:nowrap}}
td small{{display:block;color:{T.GRIS_TEXTE};font-weight:500;font-size:11px}}
td.lab{{text-align:left;white-space:normal}} td.mut{{color:{T.GRIS_TEXTE}}} td.neg{{color:{T.ROUGE_TEXTE}}}
.nm{{font-weight:700}} tr[onclick]{{cursor:pointer}} tr[onclick]:hover td{{background:{T.FOND}}}
.l0 td{{background:{T.GRIS_CLAIR};font-weight:800}} .l0 .nm{{font-weight:800}} .l1 td{{background:#FAFAFD}}
.tot td{{background:{T.NAVY};color:#fff}} .tot small{{color:#C9CAE8}}
.tg{{display:inline-block;width:16px;color:{T.GRIS_TEXTE};transition:transform .15s}} .open>td .tg{{transform:rotate(90deg)}}
.hide{{display:none}} .b{{display:inline-block;border-radius:999px;padding:2px 9px;font-size:12px;font-weight:700}}
</style></head><body>
<div class='bar'><button class='btn' onclick='all_(1)'>Tout déplier</button><button class='btn' onclick='all_(0)'>Tout replier</button></div>
<table><thead><tr>{head}</tr></thead><tbody>{''.join(rows)}</tbody></table>
<script>
function kids(id){{return document.querySelectorAll('tr[data-p="'+id+'"]')}}
function hideDeep(id){{kids(id).forEach(function(r){{r.classList.add('hide');r.classList.remove('open');hideDeep(r.id)}})}}
function tgl(tr){{if(tr.classList.toggle('open')){{kids(tr.id).forEach(function(r){{r.classList.remove('hide')}})}}else{{hideDeep(tr.id)}}}}
function all_(o){{document.querySelectorAll('tbody tr').forEach(function(r){{if(r.dataset.lvl!=='0')r.classList.toggle('hide',!o);if(r.hasAttribute('onclick'))r.classList.toggle('open',!!o)}})}}
</script></body></html>"""


def matrix(a, kind, params):
    page = matrix_html(a, kind, params)  # libellés échappés : aucun HTML venant des fichiers n'est injecté
    if hasattr(st, "iframe"):
        st.iframe(page, height=560)
    else:  # versions de Streamlit antérieures à st.iframe
        components.html(page, height=560, scrolling=True)


# ══════════════════════════════════════════════════════════════════════════════
# BARRE LATÉRALE
# ══════════════════════════════════════════════════════════════════════════════
T.sidebar_brand("🎯", "Forecast Promo")
T.sidebar_label("1 · Extraction ventes")
with st.sidebar:
    f_sales = st.file_uploader("Export PBI ventes par article (xlsx / csv)", type=["xlsx", "xlsm", "csv"], key="fp_sales",
                               help="Export hiérarchique Département › Rayon › Famille › Sous-famille › Article › Site, sur les dates de la promo.")
T.sidebar_label("2 · Liste promo des acheteurs")
with st.sidebar:
    f_promo = st.file_uploader("Code, Prévision Qté, Prix promo, PA, TVA", type=["xlsx", "xlsm", "csv"], key="fp_promo",
                               help="Une ligne par article. CA prévu et marge prévue sont calculés par le module.")
    price_ttc = st.toggle("Prix promo saisi TTC", value=True, help="Activé : le prix promo est ramené en HT avec la colonne TVA.")
T.sidebar_label("3 · Règles")
with st.sidebar:
    watch = st.slider("① À surveiller au-delà de (× N-1)", 1.0, 2.0, 1.3, 0.1)
    justify = st.slider("① À justifier au-delà de (× N-1)", 1.2, 5.0, 2.0, 0.1)
    low = st.slider("② Surestimée sous", 50, 100, 80, 5, format="%d %%") / 100
    high = st.slider("② Sous-estimée au-delà de", 100, 200, 120, 5, format="%d %%") / 100
    stable = st.slider("② Stable vs N-1 (±)", 0, 20, 5, 1, format="%d %%") / 100
    converted = st.multiselect("Sites au N-1 non comparable", E.DEFAULT_CONVERTED_SITES, default=E.DEFAULT_CONVERTED_SITES)
params = (("watch", watch), ("justify", max(justify, watch)), ("low", low), ("high", max(high, low)), ("stable", stable),
          ("converted_sites", tuple(converted)))
P = dict(params)


# ══════════════════════════════════════════════════════════════════════════════
# ACCUEIL
# ══════════════════════════════════════════════════════════════════════════════
if f_sales is None or f_promo is None:
    T.sidebar_footer("CA prévu HT = Qté × Prix promo ÷ (1 + TVA)<br>Marge prévue = Qté × (Prix HT − PA)")
    rows = [("Extraction ventes PBI", "ok" if f_sales else "wait", "reçue" if f_sales else "à déposer"),
            ("Liste promo des acheteurs", "ok" if f_promo else "wait", "reçue" if f_promo else "à déposer")]
    T.landing_hero("FORECAST <b>PROMO</b>",
                   "La prévision de l'acheteur tient-elle face à l'an dernier, et a-t-elle été tenue ? "
                   "On la contrôle avant la promo, on la mesure après, du département jusqu'à l'article.",
                   eyebrow="Forecast Promo · SmartBuyer Hub",
                   status=T.landing_status("Préparation de l'analyse", rows, "fichiers obligatoires reçus", count=2))
    T.steps([("Collecter", "la liste promo de chaque acheteur : Code, Prévision Qté, Prix promo, PA, TVA"),
             ("Extraire", "les ventes PBI sur les dates de la promo (avant : seules les colonnes N-1 sont remplies)"),
             ("① Contrôler", "avant validation : les prévisions à × 2 le N-1 sont à justifier"),
             ("② Mesurer", "après la promo : CA promis non réalisé et biais, niveau par niveau"),
             ("Exporter", "le classeur pour la revue avec les acheteurs")])
    T.section("Format des fichiers attendus")
    c1, c2, c3 = st.columns(3)
    with c1:
        T.columns_card("Colonnes attendues", [
            ("Liste promo des acheteurs", ["Code", "Prévision Qté", "Prix promo (TTC par défaut)", "PA (HT)", "TVA (18, 0,18 ou 18 %)",
                                           "Libellé (facultatif)"]),
            ("Extraction ventes (export PBI)", ["Departement · Rayon · Famille · Sous-Famille", "Article", "Site nom long", "CA · Marge · Qté Vente",
                                                "CA N-1 · Qté Vente N-1 · Marge Hors Promo N-1 · Marge Promo N-1"]),
        ], "Article au format « 12003043 - LIBELLÉ » : le code avant le tiret sert de clé. Les lignes « Total » sont ignorées. "
           "Sans l'axe Sous-Famille, ce niveau reste vide dans la matrice.")
    with c2:
        T.rules_card("📐 Méthode de calcul", [
            ("CA prévu HT", "Qté × Prix promo ÷ (1 + TVA)"),
            ("Marge prévue", "Qté × (Prix promo HT − PA)"),
            ("① Contrôle", f"{T.badge('Cohérente', 'ok')} {T.badge('× 1,3', 'att')} {T.badge('× 2', 'crit')}"),
            ("② Prévision", f"{T.badge('Fiable', 'ok')} {T.badge('Surestimée', 'crit')} {T.badge('Sous-estimée', 'att')}"),
            ("CA promis non réalisé", "somme des manques article par article"),
            ("Biais", "part d'articles surestimés ou sans vente"),
        ])
    with c3:
        T.rules_card("⚠️ Points d'attention", [
            ("Période", "l'extraction couvre uniquement les dates de la promo"),
            ("Montants", "exports PBI en HT : le prix promo TTC est ramené en HT"),
            ("Quantité réalisée", "quantité totale vendue (pas de quantité promo dans l'export)"),
            ("Marge N-1", "Marge Hors Promo N-1 + Marge Promo N-1"),
            ("Sites convertis", "Cité verte, Aboboté, Toit rouge : N-1 non comparable"),
            ("Exports Excel", "valeurs uniquement, aucune formule"),
        ])
    T.section("📖 Lexique")
    half = (len(E.LEXIQUE) + 1) // 2
    c1, c2 = st.columns(2)
    with c1:
        T.rules_card("Avant la promo et calculs", [(f"<b>{t}</b>", d) for t, d in E.LEXIQUE[:half]], tall=False)
    with c2:
        T.rules_card("Après la promo et comparaison N-1", [(f"<b>{t}</b>", d) for t, d in E.LEXIQUE[half:]], tall=False)
    T.section("Exemple de liste promo")
    st.dataframe(pd.DataFrame({"Code": [12003029, 12003043, 12503300],
                               "Libellé": ["X42 COUCH STD PACK MINI JOYA", "X40 COUCH STD PACK MAXI SMILEY", "X55 MOUCHOIRS FACIALE 4P STYLE"],
                               "Prévision Qté": [200, 1500, 900], "Prix promo": [3900, 2950, 450], "PA": [2700, 2250, 260], "TVA": [18, 18, 18]}),
                 width="stretch", hide_index=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT (en cache)
# ══════════════════════════════════════════════════════════════════════════════
try:
    sales, filters, smeta = load_sales(f_sales.getvalue(), f_sales.name)
except Exception as exc:  # noqa: BLE001
    T.page_header("🎯", "Forecast Promo")
    T.alert("Extraction ventes illisible", str(exc), "crit")
    st.stop()
try:
    promo, meta = load_promo(f_promo.getvalue(), f_promo.name, price_ttc)
except Exception as exc:  # noqa: BLE001
    T.page_header("🎯", "Forecast Promo")
    T.alert("Liste promo illisible", f"{exc} Colonnes attendues : Code, Prévision Qté, Prix promo, PA, TVA.", "crit")
    st.stop()
if sales.empty:
    T.page_header("🎯", "Forecast Promo")
    T.alert("Aucune ligne exploitable", "L'export ne contient aucune ligne article × site.", "crit")
    st.stop()

a, tot, controls = run(sales, promo, params)
per = f"du {filters['period_from']:%d/%m/%Y} au {filters['period_to']:%d/%m/%Y}" if filters["period_from"] else "période non lue"
before = tot["phase"] == "avant"
T.sidebar_footer(f"{tot['n']} articles · {tot['sites']} sites<br>{per}<br>CA prévu : {meta['ca_mode']}")


# ══════════════════════════════════════════════════════════════════════════════
# BANDEAU
# ══════════════════════════════════════════════════════════════════════════════
if before:
    title = f"FORECAST · <b>{E.fmt_ratio(tot['ratio'])}</b> LE N-1"
    chips = [f"{tot['n']} articles", (f"{tot['justify']} à justifier", "alert") if tot["justify"] else ("Aucune prévision à justifier", "ok"),
             f"{tot['watch']} à surveiller", f"{tot['nohist']} sans historique", ("Promo pas encore vendue", "")]
else:
    title = f"CA RÉALISÉ · <b>{E.fmt_pct(tot['real_ca'])}</b> DU FORECAST"
    chips = [f"{tot['n']} articles", (f"CA promis non réalisé {E.fmt_fcfa(tot['short'])}", "alert") if tot["short"] else ("Forecast tenu", "ok"),
             (f"Biais {E.fmt_pct(tot['bias'])}", "alert" if (tot["bias"] or 0) >= .25 else "ok"),
             f"{tot['justify']} prévision(s) étaient à justifier", f"CA vs N-1 {E.fmt_pct(tot['ca_evo'], 0, True)}"]
T.hero(title, f"{filters['supplier'] or 'Fournisseur non lu'} · extraction {per}", chips, eyebrow="Forecast Promo")
if not smeta["has_sf"]:
    T.alert("Sous-famille absente de l'extraction", "Ajoutez l'axe Sous-Famille entre Famille et Article dans l'export PBI pour ce niveau de la matrice.", "info")

n_ctrl = sum(len(controls[k]) for k in ["not_found", "duplicates", "no_price", "neg_marge"]) + (len(controls["no_pa"]) if meta["has_pa"] else 1)
tab1, tab2, tab3, tab4 = st.tabs(["① Avant la promo", "② Après la promo", f"🧪 Contrôles · {n_ctrl}", "⬇️ Export"])

# ── ① Avant la promo ──────────────────────────────────────────────────────────
with tab1:
    T.kpi_row([
        T.kpi("CA prévu HT", E.fmt_fcfa(tot["prev_ca"]), E.fmt_ratio(tot["ratio"]) + " le N-1",
              "neg" if (tot["ratio"] or 0) > P["justify"] else "mut", f"CA N-1 même période {E.fmt_fcfa(tot['ca_n1'])}", "blue", "🎯", "FCFA"),
        T.kpi("Prévisions à justifier", str(tot["justify"]), f"{tot['watch']} à surveiller", "neg" if tot["justify"] else "pos",
              f"prévision > × {P['justify']:.1f} le N-1".replace(".", ","), "violet", "🚩"),
        T.kpi("CA prévu à justifier", E.fmt_fcfa(tot["justify_ca"]), E.fmt_pct(E.div(tot["justify_ca"], tot["prev_ca"])) + " du forecast",
              "neg" if tot["justify_ca"] else "pos", "part du forecast portée par ces prévisions", "violet", "⚖️", "FCFA"),
        T.kpi("Sans historique", str(tot["nohist"]), "", "mut", "aucune vente N-1 sur la période", "teal", "🆕"),
    ])
    T.card(f"<div style='font-size:14px;font-weight:600;line-height:1.55'>{E.sentence_before(tot)}</div>")
    T.section("Forecast vs CA N-1 de la même période")
    matrix(a, "avant", params)
    T.card("<div class='note' style='margin-top:0'>Cliquer une ligne pour descendre. Niveaux triés par nombre de prévisions à justifier, "
           "articles triés du plus fort ratio au plus faible.</div>")

# ── ② Après la promo ──────────────────────────────────────────────────────────
with tab2:
    if before:
        T.alert("La promo n'a pas encore de ventes", "L'extraction ne contient que le N-1 : utilisez l'onglet ① pour contrôler le forecast. "
                "Rechargez la même extraction après l'opération pour mesurer le résultat.", "info")
    else:
        sg = lambda x: "pos" if (x or 0) >= P["low"] else "neg"  # noqa: E731
        cards = [
            T.kpi("CA réalisé HT", E.fmt_fcfa(tot["ca"]), E.fmt_pct(tot["real_ca"]) + " du forecast", sg(tot["real_ca"]),
                  f"forecast {E.fmt_fcfa(tot['prev_ca'])}", "blue", "🧾", "FCFA"),
            T.kpi("CA promis non réalisé", E.fmt_fcfa(tot["short"]), "à challenger" if tot["short"] else "", "neg" if tot["short"] else "pos",
                  "somme des manques, sans compensation", "violet", "📉", "FCFA"),
            T.kpi("Biais", E.fmt_pct(tot["bias"]).replace(" %", ""), f"{tot['over']} / {tot['n']} articles",
                  "neg" if (tot["bias"] or 0) >= .25 else "pos", "surestimés ou sans vente", "violet", "🎯", "%"),
            T.kpi("CA vs N-1", E.fmt_pct(tot["ca_evo"], 0, True).replace(" %", ""), f"marge {E.fmt_pct(tot['marge_evo'], 0, True)}",
                  "pos" if (tot["ca_evo"] or 0) >= 0 else "neg",
                  f"taux {E.fmt_pct(tot['rate'], 1)} · prévu {E.fmt_pct(tot['prev_rate'], 1)} · N-1 {E.fmt_pct(tot['rate_n1'], 1)}", "teal", "📅", "%"),
        ]
        T.kpi_row(cards)
        T.card(f"<div style='font-size:14px;font-weight:600;line-height:1.55'>{E.sentence_after(tot)}</div>")
        if not meta["has_pa"]:
            T.alert("Marge prévue non calculée", "Ajoutez PA et TVA dans la liste promo pour comparer la marge au forecast.", "att")
        if controls["converted"]:
            T.alert("N-1 partiellement non comparable", f"Sites convertis présents : {', '.join(controls['converted'])}. "
                    "Les articles concernés sont marqués dans la matrice.", "att")
        T.section("Forecast vs réalisé vs N-1")
        matrix(a, "après", params)
        T.card("<div class='note' style='margin-top:0'>Cliquer une ligne pour descendre. Niveaux triés par CA promis non réalisé. "
               "Un article qui dépasse ne compense pas un article gonflé.</div>")

# ── Contrôles ─────────────────────────────────────────────────────────────────
with tab3:
    def rr(ok, lab, det):
        return (f"<div class='rule-row'><span><b>{lab}</b><small style='display:block;color:{T.GRIS_TEXTE};font-weight:500'>{det}</small></span>"
                f"<span style='align-self:center'>{T.badge('OK', 'ok') if ok else T.badge('À vérifier', 'att')}</span></div>")
    days = (filters["period_to"] - filters["period_from"]).days + 1 if filters["period_from"] else None
    T.section("Extraction")
    T.card(rr(days is not None and days <= 31, "Période", per + (f", soit {days} jours : au-delà de la promo, les résultats sont dilués." if days else ""))
           + rr(bool(filters["supplier"]), "Fournisseur", filters["supplier"] or "Non lu dans les filtres.")
           + rr(True, "Sites", f"{tot['sites']} sites" + (f" : {', '.join(filters['sites'])}" if filters["sites"] else ""))
           + rr(smeta["has_sf"], "Niveaux de la matrice", " › ".join(smeta["levels_found"]) or "aucun"))
    T.section("Liste promo")
    T.card(rr(not controls["not_found"], "Codes trouvés", f"{int(a['found'].sum())} / {len(a)}"
              + (f" · introuvables : {', '.join(controls['not_found'])}" if controls["not_found"] else ""))
           + rr(not controls["duplicates"], "Doublons", ", ".join(controls["duplicates"]) or "Aucun.")
           + rr(not controls["no_price"], "Prix promo manquant", ", ".join(controls["no_price"]) or "Aucun.")
           + (rr(not controls["no_pa"], "PA manquant", ", ".join(controls["no_pa"]) or "Aucun.") if meta["has_pa"]
              else rr(False, "PA et TVA", "Colonnes absentes : la marge prévue n'est pas calculée."))
           + rr(not controls["neg_marge"], "Marge prévue négative (prix promo HT < PA)", ", ".join(controls["neg_marge"]) or "Aucune.")
           + rr(meta["has_vat"] or not price_ttc, "TVA", "Prix ramené en HT." if meta["has_vat"] and price_ttc else
                "Prix promo déclaré HT." if not price_ttc else "Colonne TVA absente : CA prévu surévalué si le prix est TTC."))

# ── Export ────────────────────────────────────────────────────────────────────
with tab4:
    T.section("Classeur Excel · valeurs uniquement")
    T.card("<div style='font-size:13px;line-height:1.55'>Onglets : <b>Synthèse</b>, <b>① Avant la promo</b> (forecast vs N-1 par article), "
           "<b>② Après la promo</b> (forecast vs réalisé vs N-1, marge en valeur et taux), <b>Contrôles</b>, <b>Lexique</b>. "
           "Chaque ligne porte Département, Rayon, Famille et Sous-famille pour filtrer ou faire un TCD.</div>"
           "<div class='note'>Généré avec les règles de la barre latérale · aucune formule.</div>")
    if st.button("Générer le classeur", type="primary", key="fp_btn", width="stretch"):
        st.session_state["fp_xlsx"] = export(a, tot, controls, filters, meta, params)
        sup = re.sub(r"[^A-Za-z0-9]+", "_", (filters["supplier"] or "promo").split(" - ")[-1]).strip("_")
        st.session_state["fp_name"] = f"SmartBuyer_Forecast_Promo_{sup}_{datetime.now():%Y%m%d}.xlsx"
    if st.session_state.get("fp_xlsx"):
        st.download_button("⬇️ Télécharger le classeur", data=st.session_state["fp_xlsx"], file_name=st.session_state["fp_name"],
                           mime=XLSX, key="fp_dl", width="stretch")
