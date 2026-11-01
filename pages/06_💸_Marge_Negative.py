"""SmartBuyer Hub · Marge négative (ex « Diagnostic Rentabilité Réseau »).

Page de pilotage (usage personnel) : on charge l'export PBI ventes article × magasin,
on repère où la marge part en négatif (article, magasin, rayon, format, effet promo),
puis on télécharge le classeur Excel (valeurs uniquement).

Les calculs et l'export sont dans modules/marge_negative/engine.py (inchangés depuis la v2.2).
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
from modules.marge_negative import engine as E  # noqa: E402

st.set_page_config(page_title="Marge négative · SmartBuyer", page_icon="💸", layout="wide", initial_sidebar_state="expanded")
T.apply()

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
FORMATS = ["Hyper", "Market", "Supeco"]
FMT_ICON = {"Hyper": "🏬", "Market": "🏪", "Supeco": "🛒"}


def num(v, dec=1) -> str:
    """Nombre à la française sans le signe % (pour les cartes KPI)."""
    return "—" if v is None or pd.isna(v) else f"{v:.{dec}f}".replace(".", ",")


def pct(v, dec=1) -> str:
    return "—" if v is None or pd.isna(v) else f"{num(v, dec)} %"


def money(v) -> str:
    return E.fmt(v).replace(",", " ")


def tx_badge(v) -> str:
    """Pastille du taux de marge : rouge < 0, orange < 8 %, vert au-delà."""
    if v is None or pd.isna(v):
        return T.badge("—", "info")
    return T.badge(pct(v), "crit" if v < 0 else ("att" if v < 8 else "ok"))


def table(df, **kw):
    st.dataframe(df, use_container_width=True, hide_index=True, **kw)


# ══════════════════════════════════════════════════════════════════════════════
# BARRE LATÉRALE
# ══════════════════════════════════════════════════════════════════════════════
T.sidebar_brand("💸", "Marge négative")
T.sidebar_label("1 · Export Power BI")
with st.sidebar:
    f_pbi = st.file_uploader("Ventes article × magasin", type=["xlsx", "xls", "csv"], key="pbi_marge",
                             help="Axes Rayon / Famille / Article / Site nom long ; mesures CA, Marge, promo et hors promo, Qté Vente.")


# ══════════════════════════════════════════════════════════════════════════════
# ACCUEIL
# ══════════════════════════════════════════════════════════════════════════════
if not f_pbi:
    T.sidebar_footer("Montants en FCFA HT<br>Seuil d'attention : taux de marge < 8 %")
    T.landing_hero("OÙ LA MARGE <b>PASSE SOUS ZÉRO</b>",
                   "Quels articles vendent à perte, dans quels magasins, et si c'est la promo, le format ou le rayon "
                   "qui tire la marge vers le bas.",
                   eyebrow="Marge négative · SmartBuyer Hub",
                   status=T.landing_status("Préparation de l'analyse",
                                           [("Export PBI ventes article × magasin", "wait", "à déposer")],
                                           "fichier reçu"))
    T.steps([("Extraire", "l'export PBI ventes de la période"),
             ("Déposer", "le fichier dans la barre latérale"),
             ("Repérer", "les articles à marge négative par magasin"),
             ("Comprendre", "promo, format, rayon, famille"),
             ("Corriger", "PA, PV ou mécanique promo avec l'acheteur")])
    T.section("Avant de déposer")
    c1, c2, c3 = st.columns(3)
    with c1:
        T.columns_card("Colonnes attendues", [
            ("Axes", ["Rayon", "Famille", "Article", "Site nom long"]),
            ("Mesures", ["CA", "Marge", "CA Promo", "Marge Promo", "CA Hors Promo", "Marge Hors Promo", "Qté Vente"]),
        ], "Article et magasin au format « code - libellé ». La période est lue dans la ligne de filtres de l'export.")
    with c2:
        T.rules_card("📐 Ce que calcule le module", [
            ("Taux de marge", "Marge ÷ CA"),
            ("Effet promo", "taux hors promo − taux promo"),
            ("Flop 100", "articles au taux le plus bas (CA > 5 000)"),
            ("Marge négative", "article × magasin sous zéro"),
            ("Niveaux", f"{T.badge('< 0 %', 'crit')} {T.badge('0–8 %', 'att')} {T.badge('≥ 8 %', 'ok')}"),
        ])
    with c3:
        T.rules_card("⚠️ Points d'attention", [
            ("Marges négatives", "souvent dégagement DLC ou ventes B2B, pas forcément une erreur de prix"),
            ("Casse", "valeur de casse non fiable dans PBI : en information seulement"),
            ("Supeco", "magasins convertis en Market en cours d'année"),
            ("Export Excel", "valeurs uniquement, aucune formule"),
        ])
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT, FILTRES, CALCULS
# ══════════════════════════════════════════════════════════════════════════════
with st.spinner("Lecture et analyse des données…"):
    df, periode, nb_jours = E.load_data(f_pbi.read(), f_pbi.name)

if df.empty:
    T.page_header("💸", "Marge négative")
    T.alert("Fichier vide ou colonnes non reconnues", "Vérifiez les axes Rayon, Famille, Article et Site nom long.", "crit")
    st.stop()

T.sidebar_label("2 · Filtres")
with st.sidebar:
    formats_dispo = [f for f in FORMATS if f in set(df["format"].dropna())]
    sel_format = st.multiselect("Format", formats_dispo, default=formats_dispo)
    rayons_dispo = sorted(df["lib_rayon"].dropna().unique())
    sel_rayon = st.multiselect("Rayon", rayons_dispo, default=rayons_dispo)
    sites_dispo = sorted(df[df["format"].isin(sel_format)]["lib_site"].dropna().unique())
    sel_site = st.multiselect("Magasin", sites_dispo, default=sites_dispo)
T.sidebar_footer(f"Période : {periode}<br>{nb_jours} jour(s) · montants en FCFA HT")

df_f = df[df["format"].isin(sel_format) & df["lib_rayon"].isin(sel_rayon) & df["lib_site"].isin(sel_site)].copy()
if df_f.empty:
    T.page_header("💸", "Marge négative")
    T.alert("Aucune donnée pour cette sélection", "Élargissez les filtres de la barre latérale.", "att")
    st.stop()

R = E.compute(df_f)
neg = R["art_neg_site"]
perte = neg["Marge"].sum() if not neg.empty else 0
nb_art_neg_site = neg["Article"].nunique() if not neg.empty else 0


# ══════════════════════════════════════════════════════════════════════════════
# BANDEAU, KPI, SIGNAUX
# ══════════════════════════════════════════════════════════════════════════════
T.hero(f"MARGE NÉGATIVE · <b>{money(abs(perte))} FCFA</b> perdus", f"{periode} · {nb_jours} jour(s)", [
    f"{R['nb_sites_actifs']} magasins", f"{len(sel_rayon)} rayons",
    (f"{nb_art_neg_site} articles sous zéro dans au moins un magasin", "alert") if nb_art_neg_site else ("Aucune vente à perte", "ok"),
    (f"{len(neg)} lignes article × magasin", "alert" if len(neg) else "ok"),
], eyebrow="Marge négative · réseau")

T.kpi_row([
    T.kpi("Chiffre d'affaires", money(R["ca_total"]), "", "mut", "périmètre filtré", "blue", "💰", "FCFA"),
    T.kpi("Taux de marge", num(R["tx_marge"]), f"hors promo {pct(R['tx_m_hp'])}", "mut", f"marge {money(R['marge_total'])}", "violet", "📈", "%"),
    T.kpi("Perte sur ventes à perte", money(abs(perte)), f"{len(neg)} lignes", "neg" if len(neg) else "pos",
          "article × magasin < 0", "violet", "🔻", "FCFA"),
    T.kpi("Effet promo", f"−{num(R['delta_hp_p'])}", f"promo {pct(R['tx_m_promo'])}", "neg" if R["delta_hp_p"] > 5 else "mut",
          "écart de taux HP − promo", "teal", "🏷️", "pts"),
    T.kpi("Poids promo", num(R["poids_promo"]), "", "mut", "part du CA en promo", "teal", "📣", "%"),
])

signals = []
if R["nb_art_neg"]:
    signals.append((f"{R['nb_art_neg']} article(s) à marge négative sur le réseau",
                    f"{R['nb_flop_neg']} dans le Flop 100 · à vérifier avec l'acheteur : prix d'achat, prix de vente, mécanique promo, "
                    "dégagement DLC ou vente B2B.", "crit"))
fr = R["agg_fmt"].set_index("format")
if "Supeco" in fr.index and "Hyper" in fr.index and fr.loc["Supeco", "TxMarge"] < 10:
    signals.append((f"Supeco à {pct(fr.loc['Supeco', 'TxMarge'])} de marge",
                    f"{num(fr.loc['Hyper', 'TxMarge'] - fr.loc['Supeco', 'TxMarge'])} pts sous les Hyper · poids promo "
                    f"{pct(fr.loc['Supeco', 'PdsPromo'])}.", "att"))
if R["delta_hp_p"] > 5:
    signals.append((f"La promo retire {num(R['delta_hp_p'])} pts de taux de marge",
                    f"Hors promo {pct(R['tx_m_hp'])} · promo {pct(R['tx_m_promo'])}.", "att"))
fs = R["fam_sous_seuil"]
if not fs.empty:
    signals.append((f"{len(fs)} famille(s) sous 8 % de marge (CA > 500 K)",
                    ", ".join(f"{r.lib_fam} ({pct(r.TxMarge)})" for r in fs.sort_values("TxMarge").head(6).itertuples()), "info"))
if signals:
    cols = st.columns(min(len(signals), 2))
    for i, (t, l, k) in enumerate(signals):
        with cols[i % len(cols)]:
            T.alert(t, l, k)


# ══════════════════════════════════════════════════════════════════════════════
# ONGLETS
# ══════════════════════════════════════════════════════════════════════════════
tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
    f"🔻 Ventes à perte · {len(neg)}", "📊 Synthèse réseau", "🔢 Rayon × magasin",
    f"💣 Flop {len(R['flop100'])}", "🗑️ Casse (info)", "⬇️ Export",
])

# ── Ventes à perte par magasin ────────────────────────────────────────────────
with tab0:
    if neg.empty:
        T.alert("Aucune vente à perte", "Aucun couple article × magasin sous zéro sur ce périmètre.", "ok")
    else:
        left, right = st.columns([1, 1.4])
        with left:
            T.section("Perte par magasin")
            ps = neg.groupby(["lib_site", "format"], as_index=False).agg(Marge=("Marge", "sum"), n=("Article", "nunique")).sort_values("Marge")
            vmax = ps["Marge"].abs().max()
            T.card("".join(T.dv_row(r.lib_site, f"{r.format} · {r.n} article(s)", r.Marge, vmax, money(r.Marge)) for r in ps.itertuples()))
        with right:
            T.section("Perte par rayon et famille")
            pr = neg.groupby(["lib_rayon", "lib_fam"], as_index=False).agg(Marge=("Marge", "sum"), n=("Article", "nunique")).sort_values("Marge").head(12)
            vmax = pr["Marge"].abs().max()
            T.card("".join(T.dv_row(r.lib_fam, f"{r.lib_rayon} · {r.n} article(s)", r.Marge, vmax, money(r.Marge)) for r in pr.itertuples()))
        T.section("Détail article × magasin")
        c1, c2 = st.columns(2)
        f_site = c1.selectbox("Magasin", ["Tous"] + sorted(neg["lib_site"].unique()), key="neg_site")
        f_ray = c2.selectbox("Rayon", ["Tous"] + sorted(neg["lib_rayon"].dropna().unique()), key="neg_ray")
        d = neg.copy()
        if f_site != "Tous":
            d = d[d["lib_site"] == f_site]
        if f_ray != "Tous":
            d = d[d["lib_rayon"] == f_ray]
        d = d.sort_values("Marge")
        table(pd.DataFrame({
            "Magasin": d["lib_site"], "Format": d["format"], "Code": d["code_art"], "Article": d["lib_art"],
            "Rayon": d["lib_rayon"], "Famille": d["lib_fam"], "CA (FCFA)": d["CA"].round(0), "Marge (FCFA)": d["Marge"].round(0),
            "Taux de marge": d["TxMarge_site"].round(1), "Qté": d["Qte"],
        }), column_config={"Taux de marge": st.column_config.NumberColumn(format="%.1f %%"),
                           "CA (FCFA)": st.column_config.NumberColumn(format="%d"),
                           "Marge (FCFA)": st.column_config.NumberColumn(format="%d")})
        st.caption(f"{len(d)} ligne(s) · triées de la plus grosse perte à la plus petite.")

# ── Synthèse réseau ───────────────────────────────────────────────────────────
with tab1:
    T.section("Taux de marge par format")
    af = R["agg_fmt"].set_index("format")
    T.kpi_row([T.kpi(f, num(af.loc[f, "TxMarge"]), f"promo {pct(af.loc[f, 'TxMarge_Promo'])}",
                     "neg" if af.loc[f, "TxMarge"] < 8 else "mut",
                     f"hors promo {pct(af.loc[f, 'TxMarge_HP'])} · CA {money(af.loc[f, 'CA'])}", "teal", FMT_ICON[f], "%")
               for f in FORMATS if f in af.index])
    left, right = st.columns([1.2, 1])
    with left:
        T.section("Taux de marge par magasin")
        s = R["agg_site"].sort_values("TxMarge")
        colors = [{"Hyper": T.NAVY, "Market": T.TEAL}.get(f, T.ORANGE) for f in s["format"]]
        fig = go.Figure(go.Bar(x=s["TxMarge"], y=s["lib_site"], orientation="h", marker_color=colors, marker_line_width=0,
                               text=[pct(v) for v in s["TxMarge"]], textposition="outside", cliponaxis=False))
        fig.add_vline(x=R["moy_marge_site"], line_width=1.5, line_dash="dot", line_color=T.ORANGE,
                      annotation_text=f" moyenne {pct(R['moy_marge_site'])}", annotation_font=dict(color=T.ORANGE, size=11))
        T.plotly_layout(fig, max(300, len(s) * 34 + 60))
        fig.update_layout(margin=dict(l=14, r=60, t=24, b=14), xaxis=dict(ticksuffix=" %", gridcolor=T.LIGNE),
                          yaxis=dict(title="", automargin=True))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Navy = Hyper · teal = Market · orange = Supeco · pointillé = moyenne des magasins")
    with right:
        T.section("Par rayon")
        rx = R["agg_rax"].sort_values("TxMarge")
        body = "".join(f"<tr><td>{r.lib_rayon}</td><td>{pct(r.PoidsCA)}</td><td>{pct(r.TxMarge_HP)}</td>"
                       f"<td>{pct(r.TxMarge_Promo)}</td><td style='text-align:right'>{tx_badge(r.TxMarge)}</td></tr>"
                       for r in rx.itertuples())
        T.card("<table class='tbl'><thead><tr><th>Rayon</th><th>Poids CA</th><th>Hors promo</th><th>Promo</th>"
               f"<th style='text-align:right'>Taux</th></tr></thead><tbody>{body}</tbody></table>")

# ── Matrice ───────────────────────────────────────────────────────────────────
with tab2:
    mp = R["mat_pivot"]
    if mp.empty:
        T.alert("Pas assez de données pour la matrice", kind="att")
    else:
        T.section("Taux de marge · rayon × magasin")
        sites = list(mp.columns)
        head = "".join(f"<th style='text-align:center'>{c}</th>" for c in sites)
        rows = ""
        for ray, line in mp.iterrows():
            cells = ""
            for c in sites:
                v = line[c]
                if pd.isna(v):
                    cells += "<td style='text-align:center;color:#9A9DB5'>—</td>"
                    continue
                bg, fg = T.heat(v - 8, 12)
                cells += f"<td style='text-align:center;background:{bg};color:{fg};font-weight:700'>{num(v)}</td>"
            rows += f"<tr><td><b>{ray}</b></td>{cells}</tr>"
        T.card(f"<div style='overflow-x:auto'><table class='tbl'><thead><tr><th>Rayon</th>{head}</tr></thead><tbody>{rows}</tbody></table></div>"
               "<div class='note'>Rouge sous 8 % de marge, vert au-dessus ; plus la couleur est forte, plus l'écart à 8 % est grand.</div>")

# ── Flop 100 ──────────────────────────────────────────────────────────────────
with tab3:
    fl = R["flop100"]
    T.alert(f"{R['nb_flop_neg']} article(s) à marge négative dans le Flop {len(fl)}",
            f"Pertes {money(fl.loc[fl['Marge'] < 0, 'Marge'].sum())} FCFA · CA du Flop {money(fl['CA'].sum())} FCFA", "crit")
    c1, c2, c3 = st.columns(3)
    fr_ = c1.selectbox("Rayon", ["Tous"] + sorted(fl["lib_rayon"].dropna().unique()), key="f100_rayon")
    fm_ = c2.selectbox("Taux de marge", ["Tous", "Négatif uniquement", "< 3 %", "< 8 %"], key="f100_marge")
    fp_ = c3.selectbox("Promo", ["Tous", "100 % sous promo", "Hors promo uniquement"], key="f100_promo")
    d = fl.copy()
    if fr_ != "Tous":
        d = d[d["lib_rayon"] == fr_]
    d = {"Négatif uniquement": d[d["TxMarge"] < 0], "< 3 %": d[d["TxMarge"] < 3], "< 8 %": d[d["TxMarge"] < 8]}.get(fm_, d)
    if fp_ == "100 % sous promo":
        d = d[d["PdsPromo"] >= 99.9]
    elif fp_ == "Hors promo uniquement":
        d = d[d["PdsPromo"].fillna(0) < 1]
    table(pd.DataFrame({
        "#": d["Rang"], "Code": d["code_art"], "Article": d["lib_art"], "Rayon": d["lib_rayon"], "Famille": d["lib_fam"],
        "CA (FCFA)": d["CA"].round(0), "Marge (FCFA)": d["Marge"].round(0), "Taux de marge": d["TxMarge"].round(1),
        "Poids promo": d["PdsPromo"].round(1), "Qté": d["Qte"],
        "🏬 Hyper": d["Bloc_Hyper"], "🏪 Market": d["Bloc_Market"], "🛒 Supeco": d["Bloc_Supeco"],
    }), column_config={"#": st.column_config.NumberColumn(width=40),
                       "Article": st.column_config.TextColumn(width="large"),
                       "Taux de marge": st.column_config.NumberColumn(format="%.1f %%"),
                       "Poids promo": st.column_config.NumberColumn(format="%.1f %%"),
                       "CA (FCFA)": st.column_config.NumberColumn(format="%d"),
                       "Marge (FCFA)": st.column_config.NumberColumn(format="%d"),
                       **{k: st.column_config.TextColumn(width="large") for k in ("🏬 Hyper", "🏪 Market", "🛒 Supeco")}})
    st.caption(f"{len(d)} article(s) · colonnes magasins triées du taux de marge le plus bas.")

# ── Casse (information) ───────────────────────────────────────────────────────
with tab4:
    T.alert("Valeur de casse non fiable", "La casse en valeur de l'export PBI a une monnaie erronée : "
            "ces chiffres servent à repérer des magasins ou articles, pas à chiffrer une perte.", "att")
    cs = R["agg_casse_site"]
    a, b = st.columns(2)
    with a:
        T.section("Par magasin")
        table(pd.DataFrame({"Magasin": cs["lib_site"], "Format": cs["format"], "Taux de casse": cs["TxCasse"].round(2)}),
              column_config={"Taux de casse": st.column_config.NumberColumn(format="%.2f %%")})
    with b:
        T.section("Top 30 articles")
        t = R["top30_casse"]
        table(pd.DataFrame({"Code": t["code_art"], "Article": t["lib_art"], "Rayon": t["lib_rayon"],
                            "Taux de casse": t["TxCasse"].round(2)}),
              column_config={"Taux de casse": st.column_config.NumberColumn(format="%.2f %%")})

# ── Export ────────────────────────────────────────────────────────────────────
with tab5:
    T.section("Export Excel · périmètre filtré, valeurs uniquement")
    T.card("<div style='font-size:13px;line-height:1.6'>Six onglets : synthèse réseau, récap rayon, matrice de marge, "
           "Flop 100 avec code article, casse, et marges négatives par magasin.</div>"
           f"<div class='note'>{len(sel_site)} magasin(s) · {len(sel_rayon)} rayon(s) · {periode}</div>", "Classeur Marge négative")
    if st.button("Générer le fichier Excel", type="primary", key="gen_excel"):
        with st.spinner("Génération du classeur…"):
            st.session_state["mn_xlsx"] = E.build_excel(R, periode, nb_jours).getvalue()
    if st.session_state.get("mn_xlsx"):
        tag = periode.replace("/", "").replace(" ", "_").replace("→", "_")
        st.download_button("⬇️ Télécharger le classeur", data=st.session_state["mn_xlsx"],
                           file_name=f"SmartBuyer_Marge_Negative_{tag}.xlsx", mime=XLSX, key="dl_excel")
