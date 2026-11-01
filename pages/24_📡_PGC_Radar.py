"""SmartBuyer Hub · PGC Radar — pilotage hebdomadaire Réel vs N-1 (budget en référence).

Page de pilotage (usage personnel) : un seul export PBI Département / Rayon / Site ;
lecture en entonnoir (vue d'ensemble → format → rayon → alertes rayon × site),
causes à saisir et export Excel figé.

Les calculs et l'export sont dans modules/pgc_radar/engine.py (inchangés).
"""
from __future__ import annotations

import pathlib
import sys

import pandas as pd
import streamlit as st

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core import theme as T  # noqa: E402
from modules.pgc_radar import engine as E  # noqa: E402

st.set_page_config(page_title="PGC Radar · SmartBuyer", page_icon="📡", layout="wide", initial_sidebar_state="expanded")
T.apply()

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
KIND = {"Critique": "crit", "Attention": "att", "OK": "ok"}
SIGN = {"neg": "neg", "pos": "pos", "mut": "mut"}


def short(lbl) -> str:
    s = str(lbl)
    return s.split(" - ", 1)[1].title() if " - " in s else s


def contrib_bar(df, label_col) -> str:
    """Barre segmentée : part de chaque élément dans le recul total vs N-1."""
    negs = df[df["ecart_n1"] < 0].sort_values("ecart_n1")
    if negs.empty:
        return ""
    tot = negs["ecart_n1"].sum()
    segs = "".join(f"<span title='{getattr(r, label_col)} · {r.ecart_n1 / tot * 100:.0f} %' style='width:{r.ecart_n1 / tot * 100:.1f}%;"
                   f"background:rgba(230,34,68,{max(0.25, 1 - i * 0.18):.2f})'></span>" for i, r in enumerate(negs.itertuples()))
    return f"<div style='display:flex;height:12px;border-radius:7px;overflow:hidden;gap:2px;margin:8px 0 4px'>{segs}</div>"


def level_card(title, sub, r, extra="") -> str:
    a = E.alerte(r["Vs_N1"], r["Marge_pt"])
    return (f"<div class='card'><div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:6px'>"
            f"<div style='font-weight:800;font-size:15px'>{title} <span style='color:{T.GRIS_TEXTE};font-weight:600;font-size:12.5px'>"
            f"{sub}</span></div>{T.badge(a, KIND[a])}</div>"
            f"<div style='font-size:13px'>Vs N-1 <b class='{E.cls(r['Vs_N1'])}'>{E.fmt_pct(r['Vs_N1'])}</b> · "
            f"écart <b class='{E.cls(r['ecart_n1'])}'>{E.fmt_m(r['ecart_n1'])}</b> · marge "
            f"<b class='{E.cls(r['Marge_pt'])}'>{E.fmt_pt(r['Marge_pt'])}</b> · "
            f"<span class='mut'>budget {E.fmt_pct(r['Vs_Bgt'])}</span></div>{extra}</div>")


# ══════════════════════════════════════════════════════════════════════════════
# BARRE LATÉRALE ET ACCUEIL
# ══════════════════════════════════════════════════════════════════════════════
T.sidebar_brand("📡", "PGC Radar")
T.sidebar_label("1 · Export Power BI")
with st.sidebar:
    up = st.file_uploader("Département / Rayon / Site (avec Budget)", type=["xlsx"], key="radar",
                          help="Onglet « Export » ; un seul export hebdomadaire suffit.")

if not up:
    T.sidebar_footer("Alertes sur Vs N-1 et marge<br>Budget en référence, jamais déclencheur")
    T.landing_hero("OÙ EN EST LE <b>PGC</b>, ET QUI DOIT CREUSER QUOI ?",
                   "Vue direction, détail par format et par rayon, puis la liste des couples rayon × magasin "
                   "en alerte à transmettre aux acheteurs, avec l'Excel prêt à partager.",
                   eyebrow="PGC Radar · SmartBuyer Hub",
                   status=T.landing_status("Préparation de l'analyse", [("Export PBI Dépt / Rayon / Site", "wait", "à déposer")], "fichier reçu"))
    T.steps([("Exporter", "PBI, semaine en cours, Dépt / Rayon / Site"), ("Déposer", "le fichier dans la barre latérale"),
             ("Lire", "vue d'ensemble, puis format et rayon"), ("Qualifier", "les causes des alertes"),
             ("Transmettre", "l'Excel aux acheteurs")])
    T.section("Avant de déposer")
    c1, c2 = st.columns(2)
    with c1:
        T.rules_card("🚨 Règles d'alerte", [
            ("Déclencheur", "Vs N-1 et marge, jamais le budget"),
            ("CA", f"{T.badge('Critique < −5 %', 'crit')} {T.badge('Attention < −2 %', 'att')}"),
            ("Marge", f"{T.badge('Critique < −1 pt', 'crit')} {T.badge('Attention < −0,5 pt', 'att')}"),
            ("Matérialité", "écart ≥ 1 M FCFA (rayon × magasin)"),
            ("Poids dans le recul", "part des écarts négatifs vs N-1"),
        ])
    with c2:
        T.columns_card("Colonnes attendues (onglet « Export »)", [
            ("Axes", ["Département", "Rayon", "Site"]),
            ("Mesures", ["CA", "CA N-1", "Budget", "Vs N-1 (%)", "Vs Bgt (%)", "Marge", "Marge N-1",
                         "Taux de Marge N Vs N-1", "Débit", "Panier"]),
        ], "Budget absent pour Supeco : affiché « — ». Débit et panier : Vs N-1, fiables au niveau rayon × magasin.")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# CALCULS
# ══════════════════════════════════════════════════════════════════════════════
try:
    R = E.compute(up)
except Exception as ex:  # export non conforme
    T.page_header("📡", "PGC Radar")
    T.alert("Export non reconnu", f"Vérifiez l'onglet « Export » et les colonnes attendues ({ex}).", "crit")
    st.stop()

pgc, al = R["pgc"], R["al_actives"]
n_crit = int((al["Alerte"] == "Critique").sum())
T.sidebar_footer("Alertes sur Vs N-1 et marge<br>Budget en référence, jamais déclencheur")

T.hero(f"PGC RADAR · <b>{E.fmt_m(pgc['CA'])}</b>", "Semaine en cours · Vs N-1", [
    f"Vs N-1 {E.fmt_pct(pgc['Vs N-1 (%)'])}", f"marge {E.fmt_pt(pgc['Taux de Marge N Vs N-1'])}",
    (f"{n_crit} alertes critiques", "alert") if n_crit else ("Aucune alerte critique", "ok"),
    f"{len(al)} alertes rayon × magasin",
], eyebrow="Pilotage hebdomadaire PGC")

T.kpi_row([
    T.kpi("CA PGC", E.fmt_m(pgc["CA"]).replace(" M", ""), E.fmt_pct(pgc["Vs N-1 (%)"]) + " vs N-1", SIGN[E.cls(pgc["Vs N-1 (%)"])],
          f"N-1 {E.fmt_m(pgc['CA N-1'])}", "blue", "💰", "M FCFA"),
    T.kpi("Taux de marge", f"{pgc['Taux de Marge'] * 100:.1f}".replace(".", ","), E.fmt_pt(pgc["Taux de Marge N Vs N-1"]),
          SIGN[E.cls(pgc["Taux de Marge N Vs N-1"])], "écart de taux vs N-1", "violet", "📈", "%"),
    T.kpi("Budget", E.fmt_m(pgc["Budget"]).replace(" M", ""), E.fmt_pct(pgc["Vs Bgt (%)"]), "mut", "référence, hors alertes", "", "🎯", "M FCFA"),
    T.kpi("Alertes", str(len(al)), f"{n_crit} critiques", "neg" if n_crit else "pos", "rayon × magasin ≥ 1 M", "teal", "🚨"),
])

tab1, tab2, tab3, tab4, tab5 = st.tabs(["🧭 Vue d'ensemble", "🏬 Format", "🗂️ Rayon", f"🚨 Alertes · {len(al)}", "⬇️ Export"])

# ── Vue d'ensemble ────────────────────────────────────────────────────────────
with tab1:
    left, right = st.columns([1.2, 1])
    with left:
        T.section("Benchmark départements · Vs N-1")
        autres = R["dept"][R["dept"]["Département"] != "01 - PGC"]
        allv = pd.concat([R["dept"][R["dept"]["Département"] == "01 - PGC"], autres])
        vmax = allv["Vs N-1 (%)"].abs().max() * 100 or 1
        T.card("".join(T.dv_row(short(r["Département"]), E.fmt_m(r["CA"]), r["Vs N-1 (%)"] * 100, vmax, E.fmt_pct(r["Vs N-1 (%)"]))
                       for _, r in allv.iterrows())
               + f"<div class='note'>Total magasin {E.fmt_m(R['ca_reseau'])}</div>")
    with right:
        T.section("Formats")
        T.html("".join(level_card(r["format"], f"{int(r['nb_sites'])} sites · {E.fmt_m(r['CA'])}", r)
                       for _, r in R["fmt_lvl"].dropna(subset=["CA"]).iterrows()))

# ── Format ────────────────────────────────────────────────────────────────────
with tab2:
    for _, f in R["fmt_lvl"].dropna(subset=["CA"]).iterrows():
        sf = R["site_lvl"][R["site_lvl"]["format"] == f["format"]].sort_values("ecart_n1")
        gros, petits = sf[sf["poids"] >= E.SEUIL_POIDS_SITE], sf[sf["poids"] < E.SEUIL_POIDS_SITE]
        rows = "".join(f"<div class='rule-row'><span>{s.site_court}</span><span><b class='{E.cls(s.Vs_N1)}'>{E.fmt_pct(s.Vs_N1)}</b>"
                       f" · {E.fmt_m(s.ecart_n1)} · <b>{s.poids * 100:.0f} % du recul</b></span></div>" for s in gros.itertuples())
        if len(petits) and petits["poids"].sum() > 0:
            rows += f"<div class='note'>+ {len(petits)} autres magasins, {petits['poids'].sum() * 100:.0f} % du recul cumulé</div>"
        elif gros.empty:
            rows += "<div class='note'>Aucun magasin significatif dans le recul.</div>"
        T.html(level_card(f["format"], f"{int(f['nb_sites'])} sites · {E.fmt_m(f['CA'])}", f, contrib_bar(sf, "site_court") + rows))

# ── Rayon ─────────────────────────────────────────────────────────────────────
with tab3:
    rl = R["rayon_lvl"]
    bar = contrib_bar(rl, "Rayon")
    if bar:
        T.card(bar + "<div class='note'>Chaque segment = un rayon en recul ; largeur = sa part du recul réseau vs N-1.</div>",
               "Poids des rayons dans le recul")
    T.section("Rayons, tous formats confondus")
    vmax = rl["ecart_n1"].abs().max() or 1
    T.card("".join(T.dv_row(short(r["Rayon"]), f"Vs N-1 {E.fmt_pct(r['Vs_N1'])} · marge {E.fmt_pt(r['Marge_pt'])}"
                            + (f" · {r['poids'] * 100:.0f} % du recul" if r["poids"] > 0 else ""),
                            r["ecart_n1"], vmax, E.fmt_m(r["ecart_n1"])) for _, r in rl.iterrows()))

# ── Alertes ───────────────────────────────────────────────────────────────────
with tab4:
    if al.empty:
        T.alert("Aucune alerte cette semaine", kind="ok")
    else:
        cards = ""
        for _, r in al.iterrows():
            site_c = r["Site"].split(" - ")[1] if " - " in r["Site"] else r["Site"]
            nom = site_c.split(" ", 1)[1] if " " in site_c else site_c
            kind = KIND[r["Alerte"]]
            cards += (f"<div class='alert-card alert-{kind}'><div style='display:flex;justify-content:space-between;gap:10px'>"
                      f"<div class='alert-title'>{nom} · {short(r['Rayon'])}</div>{T.badge(r['Alerte'], kind)}</div>"
                      f"<div class='alert-line'>CA {E.fmt_m(r['CA'])} · <b class='{E.cls(r['Vs_N1'])}'>{E.fmt_pct(r['Vs_N1'])} N-1</b> · "
                      f"marge <b class='{E.cls(r['Marge_pt'])}'>{E.fmt_pt(r['Marge_pt'])}</b> · débit "
                      f"<b class='{E.cls(r['Debit_vs_n1'])}'>{E.fmt_pct(r['Debit_vs_n1'])}</b> · panier "
                      f"<b class='{E.cls(r['Panier_vs_n1'])}'>{E.fmt_pct(r['Panier_vs_n1'])}</b> · "
                      f"<b>{r['poids'] * 100:.0f} % du recul</b></div></div>")
        T.html(f"<div style='display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px'>{cards}</div>")
        T.section("Causes à saisir pour les acheteurs")
        saisie = al[["Site", "Rayon", "Alerte"]].copy()
        saisie["Écart (M)"] = (al["ecart_n1"] / 1e6).round(1)
        saisie["Cause"] = E.CAUSES[0]
        edited = st.data_editor(saisie, hide_index=True, use_container_width=True, key="causes_editor",
                                column_config={"Cause": st.column_config.SelectboxColumn(options=E.CAUSES)})
        st.download_button("⬇️ Liste des alertes avec causes (CSV)", data=edited.to_csv(index=False).encode("utf-8-sig"),
                           file_name="alertes_pgc_radar.csv", mime="text/csv")

# ── Export ────────────────────────────────────────────────────────────────────
with tab5:
    T.card("<div style='font-size:13px;line-height:1.6'>Cinq onglets figés (aucune formule) : synthèse, formats, rayons, "
           "magasins et alertes avec la cause à choisir en liste déroulante.</div>", "Classeur PGC Radar")
    st.download_button("⬇️ Télécharger le classeur", type="primary", file_name="PGC_Radar.xlsx", mime=XLSX,
                       data=E.build_excel(pgc, R["dept"], R["ca_reseau"], R["fmt_lvl"], R["rayon_lvl"], R["site_lvl"], al))
