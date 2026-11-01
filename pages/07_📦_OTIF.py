"""SmartBuyer Hub · OTIF — taux de service fournisseurs.

Page de pilotage (usage personnel) : on charge l'extraction ERP des commandes / réceptions
et, si besoin, la liste de surveillance GOLD/SILVER ; on lit le taux de service par format,
fournisseur, magasin et article, puis on télécharge le plan d'action et les exports Excel
(valeurs uniquement, aucune formule).

Les calculs et les exports sont dans modules/otif/engine.py (inchangés depuis la v6.4).
"""
from __future__ import annotations

import pathlib
import sys
from datetime import date as _date

import numpy as np
import pandas as pd
import streamlit as st

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core import theme as T  # noqa: E402
from modules.otif import engine as E  # noqa: E402

st.set_page_config(page_title="OTIF · SmartBuyer", page_icon="📦", layout="wide", initial_sidebar_state="expanded")
T.apply()

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
FORMAT_ORDER = {"Hyper": 0, "Market": 1, "Supeco": 2}


def band(v) -> str:
    """Classe de couleur d'un taux de service : ok ≥ 97 %, att 90–97 %, crit < 90 %."""
    return "ok" if v >= E.SEUIL_EXCELLENT else ("att" if v >= E.SEUIL_SURVEILLER else "crit")


def sign(v) -> str:
    return {"ok": "pos", "att": "mut", "crit": "neg"}[band(v)]


def bar_chart(data, x_col, y_col, color, x_title, height=460, classe_col=None):
    """Barres horizontales aux couleurs de la charte (GOLD/SILVER mis en orange)."""
    fig = E.bar_h(data, x_col, y_col, color, x_title, height=height,
                  fmt_fn=lambda v: f"{int(v):,}".replace(",", " "), classe_col=classe_col)
    if classe_col and classe_col in data.columns:
        top = data.head(15).sort_values(x_col)
        fig.update_traces(marker_color=[T.ORANGE if str(c).strip() else color for c in top[classe_col]])
    T.plotly_layout(fig, height)
    fig.update_traces(textfont=dict(color=T.INK, size=12), cliponaxis=False)
    fig.update_layout(margin=dict(l=14, r=70, t=10, b=14),
                      xaxis=dict(title=x_title, gridcolor=T.LIGNE, zeroline=False),
                      yaxis=dict(title="", automargin=True))
    return fig


def supplier_rows(df, vmax) -> str:
    """Liste classée : nom, barre de volume manquant, badge du taux de service."""
    out = ""
    for r in df.itertuples():
        w = (r.qty_missing / vmax * 100) if vmax else 0
        col = T.RED if band(r.fill_rate) == "crit" else T.ORANGE
        out += (f"<div style='display:grid;grid-template-columns:minmax(0,1.6fr) minmax(0,1fr) auto;gap:14px;"
                f"align-items:center;padding:9px 0;border-bottom:1px solid {T.LIGNE}'>"
                f"<div style='font-weight:600;font-size:13.5px'>{r.supplier_name}"
                f"<small style='display:block;color:{T.GRIS_TEXTE};font-weight:500'>{E.fmt(r.qty_missing)} unités manquantes · "
                f"{r.orders} cde · {r.sites} magasin(s)</small></div>"
                f"<div style='background:{T.GRIS_CLAIR};border-radius:6px;height:8px'><div style='width:{max(w, 2):.0f}%;"
                f"height:8px;border-radius:6px;background:{col}'></div></div>"
                f"{T.badge(E.fmt_pct(r.fill_rate), band(r.fill_rate))}</div>")
    return out


def table(df):
    st.dataframe(df, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# BARRE LATÉRALE
# ══════════════════════════════════════════════════════════════════════════════
T.sidebar_brand("📦", "OTIF")
T.sidebar_label("1 · Extraction ERP")
with st.sidebar:
    uploaded = st.file_uploader("Commandes et réceptions (CSV ;)", type=["csv"], key="otif",
                                help="Export ERP des lignes de commande avec Qté cde et Qté rec, séparateur « ; ».")
T.sidebar_label("2 · Liste de surveillance (facultatif)")
with st.sidebar:
    watch_file = st.file_uploader("Articles GOLD / SILVER", type=["csv", "xlsx", "xls"], key="watchlist",
                                  help="Deux colonnes : Code article + Classe (GOLD, SILVER, A, B…).")
    watch_label = st.text_input("Nom de la liste", value="20/80", placeholder="ex : 20/80, Stratégiques…")
    watchdict: dict = {}
    if watch_file is not None:
        watchdict = E.load_watchlist(watch_file.read(), watch_file.name)
        if watchdict:
            st.success(f"{len(watchdict)} codes · {', '.join(sorted(set(watchdict.values())))}")
        else:
            st.warning("Aucun code détecté : vérifiez les deux colonnes du fichier.")
    watchlist = E.watchdict_codes(watchdict)
T.sidebar_label("3 · Règles de calcul")
with st.sidebar:
    exclude_technical = st.checkbox("Exclure les fournisseurs techniques", value=True,
                                    help="Fournisseur stock, plateformes local et import.")
    cap_sur_receipt = st.checkbox("Plafonner Qté reçue à Qté commandée", value=True)


# ══════════════════════════════════════════════════════════════════════════════
# ACCUEIL (avant dépôt du fichier)
# ══════════════════════════════════════════════════════════════════════════════
if uploaded is None:
    T.sidebar_footer("Score = Taux de Service · cible ≥ 97 %<br>Ponctualité non mesurée (dates prévues absentes)")
    rows = [("Extraction ERP commandes / réceptions", "wait", "à déposer"),
            (f"Liste de surveillance {watch_label}", "ok" if watchdict else "wait",
             f"{len(watchdict)} codes" if watchdict else "facultatif")]
    T.landing_hero("TAUX DE <b>SERVICE</b> FOURNISSEURS",
                   "Qui livre mal, sur quels magasins et quels articles, et à qui l'acheteur doit-il "
                   "demander une explication cette semaine.",
                   eyebrow="OTIF · SmartBuyer Hub",
                   status=T.landing_status("Préparation de l'analyse", rows, "fichier obligatoire reçu", count=1))
    T.steps([("Extraire", "les lignes de commande de la période dans l'ERP"),
             ("Déposer", "le CSV et, si besoin, la liste GOLD / SILVER"),
             ("Lire", "le taux de service par format, fournisseur et magasin"),
             ("Distribuer", "le plan d'action aux acheteurs"),
             ("Consolider", "les retours fournisseurs pour les magasins")])
    T.section("Avant de déposer")
    c1, c2, c3 = st.columns(3)
    with c1:
        T.columns_card("Colonnes attendues", [
            ("Extraction ERP (CSV ;)", ["Libellé site", "Nom fournisseur", "Fou", "N° Cde", "Code", "Libellé article",
                                         "Libellé rayon", "Qté cde", "Qté rec", "Prix de vente HT", "Date de commande", "Dt Rec"]),
            ("Liste de surveillance", ["Code article", "Classe"]),
        ], "Le rayon vient de « Libellé rayon » : Droguerie et Parfumerie Hygiène sont regroupés en DPH.")
    with c2:
        T.rules_card("📐 Méthode de calcul", [
            ("Taux de Service", "Qté reçue retenue ÷ Qté commandée"),
            ("Qté reçue retenue", "min(Qté reçue, Qté commandée)"),
            ("Score retenu", "Taux de Service seul"),
            ("Niveaux", f"{T.badge('≥ 97 %', 'ok')} {T.badge('90–97 %', 'att')} {T.badge('< 90 %', 'crit')}"),
            ("Criticité", "volume manquant, puis valeur"),
        ])
    with c3:
        T.rules_card("⚠️ Points d'attention", [
            ("Ponctualité et OTIF", "affichés pour information : dates prévues quasi absentes de l'ERP"),
            ("« Libellé département »", "inutilisable (toujours « PGC »)"),
            ("Classes composites", "« GOLD ET PRIX PLAFONNE » compte comme GOLD"),
            ("Exports Excel", "valeurs uniquement, aucune formule"),
        ])
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT
# ══════════════════════════════════════════════════════════════════════════════
with st.spinner("Lecture et préparation des données…"):
    raw = E.load_erp(uploaded.read(), uploaded.name)
    df, quality = E.prepare_dataset(raw, exclude_technical=exclude_technical,
                                    cap_sur_receipt=cap_sur_receipt, watchdict=watchdict)

if df.empty:
    T.page_header("📦", "OTIF")
    T.alert("Aucune ligne exploitable", "Après nettoyage, le fichier ne contient aucune ligne de commande. "
            "Vérifiez le séparateur « ; » et les colonnes Qté cde / Qté rec.", "crit")
    st.stop()

# ── Filtres (après chargement) ────────────────────────────────────────────────
def _vals(col):
    return sorted(x for x in df[col].dropna().unique() if x not in ("", "Inconnu"))


all_sites, all_sup, all_depts, all_rayons = _vals("site_label"), _vals("supplier_name"), _vals("dept_label"), _vals("rayon_label")
T.sidebar_label("4 · Filtres")
with st.sidebar:
    sel_rayons = st.multiselect("Rayon", all_rayons, default=all_rayons)
    sel_sites = st.multiselect("Magasin", all_sites, default=all_sites)
    sel_sup = st.multiselect("Fournisseur", all_sup, default=[], placeholder="Tous")
    sel_depts = all_depts
    only_crit = st.checkbox("Lignes non OTIF uniquement", value=False)
    only_watch = st.checkbox(f"Articles ⭐ {watch_label} uniquement", value=False, disabled=not watchdict)
T.sidebar_footer(f"{quality['clean_rows']:,} lignes exploitables".replace(",", " ")
                 + f" sur {quality['raw_rows']:,}".replace(",", " ") + "<br>Score = Taux de Service · cible ≥ 97 %")

sel_sites, sel_rayons = sel_sites or all_sites, sel_rayons or all_rayons
view = df[df["site_label"].isin(sel_sites) & df["dept_label"].isin(sel_depts) & df["rayon_label"].isin(sel_rayons)].copy()
if sel_sup:
    view = view[view["supplier_name"].isin(sel_sup)]
if only_crit:
    view = view[view["otif"] == 0]
if only_watch and watchdict:
    view = view[view["is_watched"]]
if view.empty:
    T.page_header("📦", "OTIF")
    T.alert("Aucune donnée après filtrage", "Élargissez les filtres de la barre latérale.", "att")
    st.stop()

kpi = E.compute_global_kpis(view)
by_supplier = E.agg_supplier(view)
by_site = E.agg_site(view)
by_article = E.agg_article(view, watchdict=watchdict)
watch_kpi = E.compute_global_kpis(view[view["is_watched"]]) if watchdict else None
crit_only = by_supplier[by_supplier["Niveau"] == "🔴 Critique"]


# ══════════════════════════════════════════════════════════════════════════════
# BANDEAU ET KPI
# ══════════════════════════════════════════════════════════════════════════════
periode = ""
if "Date de commande" in view.columns and view["Date de commande"].notna().any():
    d0, d1 = view["Date de commande"].min(), view["Date de commande"].max()
    periode = f"Commandes du {d0:%d/%m/%Y} au {d1:%d/%m/%Y}"
chips = [f"{kpi['sites']} magasins", f"{kpi['suppliers']} fournisseurs", f"{kpi['orders']:,} commandes".replace(",", " "),
         (f"{len(crit_only)} fournisseurs critiques", "alert") if len(crit_only) else ("Aucun fournisseur critique", "ok")]
if watchdict:
    chips.append((f"⭐ {watch_label} · {view.loc[view['is_watched'], 'Code'].nunique()} réf. trouvées", ""))
T.hero(f"TAUX DE SERVICE · <b>{E.fmt_pct(kpi['fill_rate'])}</b>", periode or "Extraction ERP", chips, eyebrow="OTIF fournisseurs")

cards = [
    T.kpi("Taux de Service", E.fmt_pct(kpi["fill_rate"]).replace("%", "").strip(), E.score_band(kpi["fill_rate"]).split(" ", 1)[-1],
          sign(kpi["fill_rate"]), "cible ≥ 97 %", "blue", "📦", "%"),
    T.kpi("Volume manquant", E.fmt(kpi["missing_qty"]), "", "mut",
          f"sur {E.fmt(kpi['ordered_qty'])} commandées", "violet", "📉", "unités"),
    T.kpi("Fournisseurs critiques", str(len(crit_only)), f"sur {kpi['suppliers']}", "neg" if len(crit_only) else "pos",
          "taux de service < 90 %", "violet", "🚛", ""),
    T.kpi("OTIF", E.fmt_pct(kpi["otif"]).replace("%", "").strip(), "information", "mut",
          "dates prévues absentes", "", "⏱️", "%"),
]
if watch_kpi:
    cards.insert(2, T.kpi(f"⭐ {watch_label}", E.fmt_pct(watch_kpi["fill_rate"]).replace("%", "").strip(),
                          f"{E.fmt(watch_kpi['missing_qty'])} manquants", sign(watch_kpi["fill_rate"]),
                          "taux de service des articles suivis", "teal", "⭐", "%"))
T.kpi_row(cards)

if watch_kpi and watch_kpi["fill_rate"] < E.SEUIL_SURVEILLER:
    T.alert(f"Articles {watch_label} en sous-service",
            f"Taux de service {E.fmt_pct(watch_kpi['fill_rate'])} · {E.fmt(watch_kpi['missing_qty'])} unités manquantes. "
            "Escalade au Responsable Achat dans le plan d'action.", "crit")


# ══════════════════════════════════════════════════════════════════════════════
# ONGLETS
# ══════════════════════════════════════════════════════════════════════════════
tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🎯 Synthèse", f"🚛 Fournisseurs · {len(by_supplier)}", f"🏪 Magasins · {len(by_site)}",
    f"📦 Articles · {len(by_article)}", "📋 Fiche fournisseur", "⬇️ Exports",
])

# ── Synthèse (vue Responsable Achat) ──────────────────────────────────────────
with tab0:
    fv = view.assign(format_magasin=view["site_label"].apply(E.format_magasin))
    fa = fv[fv["format_magasin"] != "Autre"].groupby("format_magasin", as_index=False).agg(
        qte_cde=("qte_cde", "sum"), qte_rec=("qte_rec_retained", "sum"), qty_missing=("qty_missing", "sum"),
        sites=("site_label", "nunique"))
    fa["fill_rate"] = np.where(fa["qte_cde"] > 0, fa["qte_rec"] / fa["qte_cde"] * 100, 0.0)
    fa = fa.sort_values("format_magasin", key=lambda s: s.map(FORMAT_ORDER))
    if len(fa):
        T.section("Taux de service par format")
        T.kpi_row([T.kpi(r.format_magasin, E.fmt_pct(r.fill_rate).replace("%", "").strip(), E.score_band(r.fill_rate).split(" ", 1)[-1],
                         sign(r.fill_rate), f"{r.sites} magasin(s) · {E.fmt(r.qty_missing)} manquants",
                         "teal", {"Hyper": "🏬", "Market": "🏪", "Supeco": "🛒"}.get(r.format_magasin, ""), "%")
                   for r in fa.itertuples()])

    left, right = st.columns([1.35, 1])
    with left:
        T.section("Fournisseurs critiques à appeler en priorité")
        top = crit_only.sort_values("qty_missing", ascending=False).head(10)
        if top.empty:
            T.alert("Aucun fournisseur sous 90 %", "Aucun appel prioritaire cette semaine sur ce périmètre.", "ok")
        else:
            T.card(supplier_rows(top, top["qty_missing"].max())
                   + f"<div class='note'>{len(crit_only)} fournisseur(s) critique(s) · classés par volume manquant.</div>")
    with right:
        T.section("Taux de service par rayon")
        ra = view.groupby("rayon_label", as_index=False).agg(qte_cde=("qte_cde", "sum"), qte_rec=("qte_rec_retained", "sum"),
                                                             qty_missing=("qty_missing", "sum"))
        ra["fill_rate"] = np.where(ra["qte_cde"] > 0, ra["qte_rec"] / ra["qte_cde"] * 100, 0.0)
        body = "".join(f"<div class='rule-row'><span>{r.rayon_label}<small style='display:block;color:{T.GRIS_TEXTE}'>"
                       f"{E.fmt(r.qty_missing)} unités manquantes</small></span><span style='align-self:center'>{T.badge(E.fmt_pct(r.fill_rate), band(r.fill_rate))}</span></div>"
                       for r in ra.sort_values("fill_rate").itertuples())
        T.card(body)
        T.section("Plan d'action acheteurs")
        T.card(f"<div style='font-size:13px;line-height:1.55'>Fichier Excel à 4 onglets : résumé, suivi fournisseur × rayon "
               f"avec cause racine et statut à saisir, matrice par magasin, détail des commandes non servies.</div>"
               f"<div class='note'>Tous les fournisseurs et articles critiques · fichier vierge chaque semaine.</div>")
        if st.button("Générer le plan d'action", type="primary", key="btn_plan_action", use_container_width=True):
            with st.spinner("Génération du plan d'action (10 à 30 s)…"):
                st.session_state["plan_action_bytes"] = E.build_plan_action_excel(view, by_supplier, quality, watchdict=watchdict).getvalue()
                st.session_state["plan_action_date"] = _date.today().strftime("%d-%m-%Y")
        if st.session_state.get("plan_action_bytes"):
            st.download_button("⬇️ Télécharger le plan d'action", data=st.session_state["plan_action_bytes"],
                               file_name=f"Taux de Service - {st.session_state['plan_action_date']}.xlsx",
                               mime=XLSX, key="dl_plan_action", use_container_width=True)

    with st.expander("🧪 Qualité des données"):
        n = quality["clean_rows"] or 1
        T.kpi_row([
            T.kpi("Lignes exploitables", f"{quality['clean_rows']:,}".replace(",", " "), E.fmt_pct(quality["usable_rate"]), "mut",
                  f"sur {quality['raw_rows']:,} brutes".replace(",", " ")),
            T.kpi("Sur-réceptions plafonnées", f"{quality['sur_receipt_rows']:,}".replace(",", " ")),
            T.kpi("Sans date prévue", f"{quality['missing_expected_date'] / n * 100:.1f}".replace(".", ","), "", "mut",
                  f"colonne : {quality['expected_col']}", unit="%"),
            T.kpi("Sans prix de vente HT", f"{quality['pv_zero_rows'] / n * 100:.1f}".replace(".", ","), "", "mut",
                  "impact CA non affiché", unit="%"),
        ])

# ── Fournisseurs ──────────────────────────────────────────────────────────────
with tab1:
    sup_sorted = by_supplier.sort_values("qty_missing", ascending=False)
    T.section("Top 15 · volume manquant (unités)")
    st.plotly_chart(bar_chart(sup_sorted, "qty_missing", "supplier_name", T.RED, "Volume manquant (unités)"), use_container_width=True)
    d = sup_sorted.assign(**{"Taux de Service": sup_sorted["fill_rate"].apply(E.fmt_pct), "OTIF": sup_sorted["otif"].apply(E.fmt_pct),
                             "Vol. manquant": sup_sorted["qty_missing"].apply(E.fmt)})
    table(d[["supplier_name", "Niveau", "Taux de Service", "Vol. manquant", "OTIF", "orders", "articles", "sites"]].rename(
        columns={"supplier_name": "Fournisseur", "orders": "Cdes", "articles": "Articles", "sites": "Magasins"}))

# ── Magasins ──────────────────────────────────────────────────────────────────
with tab2:
    s = by_site.copy()
    s["format"] = s["site_label"].apply(E.format_magasin)
    s = s.sort_values(["format", "fill_rate"], key=lambda c: c.map(FORMAT_ORDER).fillna(9) if c.name == "format" else c)
    T.section("Taux de service par magasin")
    body = "".join(f"<tr><td>{r.site_label}</td><td class='mut'>{r.format}</td><td>{E.fmt(r.qty_missing)}</td>"
                   f"<td>{r.suppliers}</td><td style='text-align:right'>{T.badge(E.fmt_pct(r.fill_rate), band(r.fill_rate))}</td></tr>"
                   for r in s.itertuples())
    T.card("<table class='tbl'><thead><tr><th>Magasin</th><th>Format</th><th>Vol. manquant</th><th>Fournisseurs</th>"
           f"<th style='text-align:right'>Taux de Service</th></tr></thead><tbody>{body}</tbody></table>")
    T.section("Volume manquant par magasin (unités)")
    st.plotly_chart(bar_chart(by_site, "qty_missing", "site_label", T.ORANGE, "Volume manquant (unités)", height=420),
                    use_container_width=True)

# ── Articles ──────────────────────────────────────────────────────────────────
with tab3:
    d_art = by_article.copy()
    if watchdict and st.toggle(f"⭐ Liste {watch_label} uniquement ({int((d_art['Classe'] != '').sum())} réf.)", key="tgl_watch"):
        d_art = d_art[d_art["Classe"] != ""]
    if d_art.empty:
        T.alert(f"Aucun article de la liste {watch_label} dans le périmètre filtré", kind="att")
    else:
        T.section("Top 15 articles · volume manquant (unités)" + (" · ⭐ en orange" if watchdict else ""))
        st.plotly_chart(bar_chart(d_art, "qty_missing", "article_label", T.NAVY, "Volume manquant (unités)", height=560,
                                  classe_col="Classe" if watchdict else None), use_container_width=True)
        d = d_art.assign(**{"Taux de Service": d_art["fill_rate"].apply(E.fmt_pct), "OTIF": d_art["otif"].apply(E.fmt_pct),
                            "Vol. manquant": d_art["qty_missing"].apply(E.fmt)})
        cols = [c for c in ["Classe", "Code", "article_label", "supplier_name", "Taux de Service", "Vol. manquant", "OTIF", "sites", "orders"]
                if c in d.columns]
        table(d[cols].rename(columns={"article_label": "Article", "supplier_name": "Fournisseur", "sites": "Magasins", "orders": "Cdes"}))
    if watchdict:
        matched = set(by_article.loc[by_article["Classe"] != "", "code_str"]) if "code_str" in by_article.columns else set()
        missing = sorted(watchlist - matched)
        if missing:
            with st.expander(f"🔍 {len(missing)} code(s) de la liste absents des commandes"):
                table(pd.DataFrame(missing, columns=["Code non trouvé"]))

# ── Fiche fournisseur ─────────────────────────────────────────────────────────
with tab4:
    c1, c2 = st.columns([2, 1])
    order = list(by_supplier.sort_values("qty_missing", ascending=False)["supplier_name"])
    fiche_supplier = c1.selectbox("Fournisseur", ["— Choisir —"] + order, key="fiche_sel",
                                  help="Classés du plus gros volume manquant au plus petit.")
    seuil_fill = c2.slider("Lignes retenues : taux de service sous", 50, 100, 100, 5, format="%d%%")
    if fiche_supplier == "— Choisir —":
        T.alert("Choisissez un fournisseur", "La fiche détaille ses livraisons incomplètes par magasin, par article et par "
                "commande, avec un export Excel prêt à envoyer.", "info")
    else:
        fiche_df = view[view["supplier_name"] == fiche_supplier]
        fiche_bad = E.safe_sort(fiche_df[fiche_df["line_fill_rate"] < seuil_fill / 100],
                                keys=["is_watched", "qty_missing", "service_gap_value"], ascending=[False, False, False])
        fk = E.compute_global_kpis(fiche_df)
        fou = str(int(fiche_df["Fou"].dropna().iloc[0])) if "Fou" in fiche_df.columns and fiche_df["Fou"].notna().any() else "—"
        n_w = int(fiche_df["is_watched"].sum()) if "is_watched" in fiche_df.columns else 0
        fchips = [f"{fk['orders']} commandes", f"{fk['articles']} articles", f"{fk['sites']} magasins"]
        if watchdict and n_w:
            fchips.append((f"⭐ {n_w} lignes {watch_label}", "alert"))
        T.hero(f"<b>{fiche_supplier}</b>", f"Code {fou}", fchips, eyebrow="Fiche fournisseur")
        T.kpi_row([
            T.kpi("Taux de Service", E.fmt_pct(fk["fill_rate"]).replace("%", "").strip(), E.score_band(fk["fill_rate"]).split(" ", 1)[-1],
                  sign(fk["fill_rate"]), "cible ≥ 97 %", "blue", "📦", "%"),
            T.kpi("Volume manquant", E.fmt(fk["missing_qty"]), "", "mut", f"sur {E.fmt(fk['ordered_qty'])} commandées", "violet", "📉", "unités"),
            T.kpi("Lignes incomplètes", f"{len(fiche_bad):,}".replace(",", " "), "", "mut",
                  f"sur {len(fiche_df):,} lignes".replace(",", " "), "violet", "🧾"),
        ])
        agg = dict(nb_cdes=("N° Cde", "nunique"), qte_cde=("qte_cde", "sum"), qte_rec=("qte_rec_retained", "sum"),
                   qty_missing=("qty_missing", "sum"), impact_value=("service_gap_value", "sum"))
        site_recap = fiche_bad.groupby("site_label", as_index=False).agg(nb_articles=("Code", "nunique"), **agg).sort_values("qty_missing", ascending=False)
        art_recap = fiche_bad.groupby(["Code", "article_label"], as_index=False).agg(nb_sites=("site_label", "nunique"), **agg).sort_values("qty_missing", ascending=False)
        for r in (site_recap, art_recap):
            r["Taux de Service"] = (r["qte_rec"] / r["qte_cde"] * 100).apply(E.fmt_pct)
            r["Vol. manquant"] = r["qty_missing"].apply(E.fmt)
            r["Impact CA proxy"] = r["impact_value"].apply(E.fmt)
        a, b = st.columns(2)
        with a:
            T.section("Par magasin")
            table(site_recap[["site_label", "Taux de Service", "Vol. manquant", "nb_cdes", "nb_articles"]].rename(
                columns={"site_label": "Magasin", "nb_cdes": "Cdes", "nb_articles": "Articles"}))
        with b:
            T.section("Par article")
            cols_art = ["Code", "article_label", "Taux de Service", "Vol. manquant", "nb_sites", "nb_cdes"]
            if watchdict:
                art_recap["Classe"] = art_recap["Code"].apply(lambda c: E.get_classe(c, watchdict))
                cols_art = ["Classe"] + cols_art
            table(art_recap[cols_art].rename(columns={"article_label": "Article", "nb_sites": "Magasins", "nb_cdes": "Cdes"}))
        T.section("Détail des livraisons incomplètes" + (f" · ⭐ {watch_label} en tête" if watchdict else ""))
        dcols = [c for c in ["watch_classe", "N° Cde", "date_received", "site_label", "Code", "article_label", "qte_cde",
                             "qte_rec_retained", "qty_missing", "line_fill_rate"] if c in fiche_bad.columns]
        disp = fiche_bad[dcols].copy()
        if "line_fill_rate" in disp.columns:
            disp["line_fill_rate"] = disp["line_fill_rate"].apply(lambda v: E.fmt_pct(v * 100))
        table(disp.rename(columns={"watch_classe": "Classe", "N° Cde": "N° Commande", "date_received": "Date réception",
                                   "site_label": "Magasin", "article_label": "Article", "qte_cde": "Qté cde",
                                   "qte_rec_retained": "Qté reçue", "qty_missing": "Qté manquante",
                                   "line_fill_rate": "Taux de Service ligne"}))
        if st.button("Générer la fiche Excel", type="primary", key="btn_fiche"):
            with st.spinner("Génération de la fiche…"):
                buf = E.build_fiche_excel(fournisseur=fiche_supplier, df_all=fiche_df, df_bad=fiche_bad, site_recap=site_recap,
                                          art_recap=art_recap, kpis=fk, seuil=seuil_fill, watchdict=watchdict)
            safe = fiche_supplier.strip().replace(" ", "_").replace("/", "-")[:40]
            st.download_button(f"⬇️ Télécharger la fiche {fiche_supplier}", data=buf,
                               file_name=f"SmartBuyer_OTIF_{safe}.xlsx", mime=XLSX, key="dl_fiche")

# ── Exports ───────────────────────────────────────────────────────────────────
with tab5:
    seuil_exp = 100  # export « toutes les fiches » : toutes les lignes incomplètes
    w = "<br>⭐ Classe GOLD / SILVER incluse" if watchdict else ""
    exports = [
        ("🎯 Récap priorisation fournisseurs", f"Une ligne par fournisseur : réf. commandées et livrées, taux de service, "
         f"magasins touchés. Trié par criticité.{w}", "btn_recap", "dl_recap", "SmartBuyer_OTIF_Recap_Fournisseurs.xlsx",
         lambda: E.build_recap_fournisseur_excel(view, watchdict=watchdict)),
        ("📋 Toutes les fiches fournisseurs", f"Un onglet : toutes les livraisons incomplètes de tous les fournisseurs, "
         f"trié par criticité.{w}", "btn_all", "dl_all", "SmartBuyer_OTIF_Fiches_Fournisseurs.xlsx",
         lambda: E.build_export_all_fiches(view, by_supplier, seuil_exp, watchdict=watchdict)),
        ("📊 Synthèse analytique", f"Onglets par fournisseur, magasin, article, lignes critiques et qualité des données.{w}",
         "btn_global", "dl_global", "SmartBuyer_OTIF_Synthese.xlsx",
         lambda: E.build_export_excel(view, by_supplier, by_site, by_article, quality, watchdict=watchdict)),
    ]
    T.section("Exports Excel · périmètre filtré, valeurs uniquement")
    for col, (title, text, bk, dk, fname, build) in zip(st.columns(3), exports):
        with col:
            T.card(f"<div style='font-size:13px;line-height:1.55;min-height:62px'>{text}</div>", title)
            if st.button("Générer", type="primary", key=bk, use_container_width=True):
                with st.spinner("Génération…"):
                    st.session_state[dk + "_bytes"] = build().getvalue()
            if st.session_state.get(dk + "_bytes"):
                st.download_button("⬇️ Télécharger", data=st.session_state[dk + "_bytes"], file_name=fname, mime=XLSX,
                                   key=dk, use_container_width=True)
