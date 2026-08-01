# -*- coding: utf-8 -*-
"""
18_💸_Reporting_Vente CA.py
============================================================
SmartBuyer Hub — Module Reporting Commercial (Synthèse Exécutive & Flops)

Traite un export Excel (feuille "Export") contenant les métriques de ventes
(CA, CA N-1, Budget, Marges, Débits, Panier moyen, Volumes) à 3 niveaux :
    - Global (Société)
    - Rayon (toutes enseignes confondues)
    - Couple Magasin x Rayon (niveau de détection des Flops)

Règles de flop (validées) :
    C1 - Décrochage CA vs N-1   : Vs N-1 (%)  <= -10%
    C2 - Écart vs Budget        : Vs Bgt (%)  <= -10%  (ignoré si Budget NaN)
    C3 - Dégradation marge      : Delta Marge (pts) <= -0.8 pt
    C4 - Rupture / fermeture    : CA NaN/0 alors que CA N-1 > 0 (prioritaire)

Sévérité = cumul simple des critères applicables déclenchés :
    C4 déclenché         -> CRITIQUE
    >= 2 critères KO      -> FLOP MAJEUR
    1 critère KO          -> FLOP MODÉRÉ
    0 critère KO          -> OK

Charte : Dark Fintech (fond #0B0B0F, cartes #151519, bordures #2A2A30)

Architecture du fichier :
    0-4  : config, palette, chargement données, moteurs de calcul, helpers
           -> pur Python/pandas, importable et testable sans Streamlit actif
    5    : main() -> tout le rendu Streamlit (sidebar, KPI, tabs)
    6    : tests unitaires (fonctions test_*)
    7    : point d'entrée -> main() en prod Streamlit, tests si
           RUN_DASHBOARD_TESTS=1
============================================================
"""

import io
import os
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# ============================================================
# 0. CONFIGURATION PAGE & THEME
# ============================================================

st.set_page_config(
    page_title="Reporting Vente CA",
    page_icon="💸",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- Palette Dark Fintech ---
COL_BG_PAGE = "#0B0B0F"
COL_BG_CARD = "#151519"
COL_BORDER = "#2A2A30"
COL_TEXT_PRIMARY = "#F5F5F7"
COL_TEXT_SECONDARY = "#9A9AA0"

COL_RED = "#F0997B"
COL_RED_BG = "#4A1B0C"
COL_ORANGE = "#EF9F27"
COL_ORANGE_BG = "#633806"
COL_AMBER = "#FAC775"
COL_AMBER_BG = "#854F0B"
COL_GREEN = "#97C459"
COL_GREEN_BG = "#173404"
COL_BLUE = "#85B7EB"
COL_PURPLE = "#AFA9EC"
COL_GRAY = "#5F5E5A"

SEVERITY_STYLE = {
    "Critique": {"text": COL_RED, "bg": COL_RED_BG, "emoji": "🔴"},
    "Flop majeur": {"text": COL_ORANGE, "bg": COL_ORANGE_BG, "emoji": "🟠"},
    "Flop modéré": {"text": COL_AMBER, "bg": COL_AMBER_BG, "emoji": "🟡"},
    "OK": {"text": COL_GREEN, "bg": COL_GREEN_BG, "emoji": "🟢"},
}

CUSTOM_CSS = f"""
<style>
    .stApp {{
        background-color: {COL_BG_PAGE};
    }}
    section[data-testid="stSidebar"] {{
        background-color: {COL_BG_CARD};
        border-right: 0.5px solid {COL_BORDER};
    }}
    h1, h2, h3, h4, p, span, div, label {{
        color: {COL_TEXT_PRIMARY};
    }}
    .kpi-card {{
        background-color: {COL_BG_CARD};
        border: 0.5px solid {COL_BORDER};
        border-radius: 10px;
        padding: 0.85rem 1rem;
        margin-bottom: 8px;
    }}
    .kpi-label {{
        font-size: 12px;
        color: {COL_TEXT_SECONDARY};
        margin: 0 0 6px 0;
    }}
    .kpi-value {{
        font-size: 22px;
        font-weight: 600;
        color: {COL_TEXT_PRIMARY};
        margin: 0;
    }}
    .badge-pill {{
        font-size: 11px;
        padding: 3px 10px;
        border-radius: 20px;
        display: inline-block;
        font-weight: 600;
    }}
    .rayon-card {{
        background-color: {COL_BG_CARD};
        border: 0.5px solid {COL_BORDER};
        border-radius: 12px;
        padding: 1rem;
        margin-bottom: 8px;
        height: 100%;
    }}
    .flop-row {{
        background-color: {COL_BG_CARD};
        border: 0.5px solid {COL_BORDER};
        border-radius: 10px;
        padding: 0.75rem 1rem;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }}
    div[data-testid="stMetric"] {{
        background-color: {COL_BG_CARD};
        border: 0.5px solid {COL_BORDER};
        border-radius: 10px;
        padding: 0.75rem 1rem;
    }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

PLOTLY_TEMPLATE = go.layout.Template(
    layout=go.Layout(
        paper_bgcolor=COL_BG_PAGE,
        plot_bgcolor=COL_BG_PAGE,
        font=dict(color=COL_TEXT_PRIMARY, size=12),
        xaxis=dict(gridcolor=COL_BORDER, zerolinecolor=COL_BORDER),
        yaxis=dict(gridcolor=COL_BORDER, zerolinecolor=COL_BORDER),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
)

# ============================================================
# 1. CHARGEMENT & NETTOYAGE DES DONNÉES
# ============================================================

EXPECTED_COLS = [
    "Société", "Rayon", "Site", "CA N-1", "Budget", "CA", "Poids",
    "Vs N-1 (%)", "Vs Bgt (%)", "Marge N-1", "Marge",
    "Taux de Marge N-1", "Taux de Marge", "Taux de Marge N Vs N-1",
    "Débit N-1", "Débit", "Vs N-1 (%).1", "Panier N-1", "Panier",
    "Panier N Vs N-1", "Panier Qté N-1", "Panier Qté",
    "Panier Qté N Vs N-1", "Volume N-1", "Volume", "Volume N Vs N-1",
]

FORMAT_KEYWORDS = {
    "Hyper": "Hyper",
    "Market": "Market",
    "Supeco": "Supeco",
}


def split_code_libelle(serie: pd.Series) -> pd.DataFrame:
    """Sépare une colonne 'CODE - Libellé' en 2 colonnes Code / Libellé.
    Gère les valeurs NaN et les valeurs sans séparateur ' - '."""
    serie = serie.astype("string")
    split = serie.str.split(" - ", n=1, expand=True)
    if split.shape[1] == 1:
        split[1] = np.nan
    code = split[0].str.strip()
    libelle = split[1].str.strip()
    libelle = libelle.fillna(serie)
    return pd.DataFrame({"Code": code, "Libelle": libelle})


def detect_format(libelle) -> str:
    """Déduit le format de magasin (Hyper/Market/Supeco) à partir du libellé Site."""
    if not isinstance(libelle, str):
        return "Autre"
    for kw, fmt in FORMAT_KEYWORDS.items():
        if kw.lower() in libelle.lower():
            return fmt
    return "Autre"


def _load_data_impl(file) -> dict:
    """Implémentation pure (sans décorateur cache) — utilisée par main() et
    par les tests, pour rester importable/testable hors contexte Streamlit."""
    raw = pd.read_excel(file, sheet_name="Export")

    missing_cols = [c for c in EXPECTED_COLS if c not in raw.columns]

    df = raw.dropna(how="all").copy()
    df = df[df["Société"].astype("string") != "Total"]
    df = df[~df["Société"].astype("string").str.startswith("Filtres appliqués", na=False)]
    df = df[~(df["Rayon"].isna() & df["Site"].isna())]

    is_global = (df["Rayon"].astype("string") == "Total") & (df["Site"].isna())
    is_rayon = (df["Rayon"].astype("string") != "Total") & (df["Site"].astype("string") == "Total")
    is_couple = (df["Rayon"].astype("string") != "Total") & (
        ~df["Site"].isna() & (df["Site"].astype("string") != "Total")
    )

    df_global = df[is_global].copy()
    df_rayon = df[is_rayon].copy()
    df_couple = df[is_couple].copy()

    for d in (df_rayon, df_couple):
        rayon_split = split_code_libelle(d["Rayon"])
        d["Rayon_Code"] = rayon_split["Code"].values
        d["Rayon_Libelle"] = rayon_split["Libelle"].values

    site_split = split_code_libelle(df_couple["Site"])
    df_couple["Site_Code"] = site_split["Code"].values
    df_couple["Site_Libelle"] = site_split["Libelle"].values
    df_couple["Format"] = df_couple["Site_Libelle"].apply(detect_format)

    return {
        "global": df_global.reset_index(drop=True),
        "rayon": df_rayon.reset_index(drop=True),
        "couple": df_couple.reset_index(drop=True),
        "missing_cols": missing_cols,
    }


@st.cache_data(show_spinner="Chargement des données...")
def load_data(file) -> dict:
    return _load_data_impl(file)


# ============================================================
# 2. MOTEUR DE CALCUL — FLOPS & SÉVÉRITÉ (niveau Couple)
# ============================================================

def compute_delta_marge_pts(taux_marge: pd.Series, taux_marge_n1: pd.Series) -> pd.Series:
    """Delta de taux de marge en points (ex: 21.7% - 20.9% = +0.8 pt).
    Robuste aux NaN (pas de division impliquée -> pas de risque ZeroDivisionError)."""
    return (taux_marge - taux_marge_n1) * 100


def compute_flops(df: pd.DataFrame, seuil_ca: float, seuil_bgt: float, seuil_marge: float) -> pd.DataFrame:
    """Calcule les 4 critères de flop + la sévérité pour chaque couple Magasin x Rayon.

    seuil_ca, seuil_bgt : négatifs, ex. -0.10 pour -10%
    seuil_marge         : négatif, en points, ex. -0.8
    """
    out = df.copy()

    out["Delta_Marge_pt"] = compute_delta_marge_pts(out["Taux de Marge"], out["Taux de Marge N-1"])

    # --- C4 : rupture / fermeture (prioritaire) ---
    ca_is_missing_or_zero = out["CA"].isna() | (out["CA"] == 0)
    ca_n1_positif = out["CA N-1"].fillna(0) > 0
    out["C4_Rupture"] = ca_is_missing_or_zero & ca_n1_positif

    # --- C1 : décrochage CA vs N-1 ---
    out["C1_Decrochage_CA"] = (out["Vs N-1 (%)"] <= seuil_ca).fillna(False) & ~out["C4_Rupture"]

    # --- C2 : écart vs budget (uniquement si Budget renseigné) ---
    budget_applicable = out["Budget"].notna()
    out["C2_Applicable"] = budget_applicable
    out["C2_Ecart_Budget"] = np.where(budget_applicable, out["Vs Bgt (%)"] <= seuil_bgt, False)
    out["C2_Ecart_Budget"] = out["C2_Ecart_Budget"].astype(bool) & ~out["C4_Rupture"]

    # --- C3 : dégradation de marge (nécessite Taux de Marge des 2 périodes) ---
    marge_applicable = out["Taux de Marge"].notna() & out["Taux de Marge N-1"].notna()
    out["C3_Applicable"] = marge_applicable
    out["C3_Degradation_Marge"] = np.where(marge_applicable, out["Delta_Marge_pt"] <= seuil_marge, False)
    out["C3_Degradation_Marge"] = out["C3_Degradation_Marge"].astype(bool) & ~out["C4_Rupture"]

    out["Nb_Criteres_Applicables"] = 1 + out["C2_Applicable"].astype(int) + out["C3_Applicable"].astype(int)
    out["Nb_Criteres_KO"] = (
        out["C1_Decrochage_CA"].astype(int)
        + out["C2_Ecart_Budget"].astype(int)
        + out["C3_Degradation_Marge"].astype(int)
    )

    def severity(row):
        if row["C4_Rupture"]:
            return "Critique"
        if row["Nb_Criteres_KO"] >= 2:
            return "Flop majeur"
        if row["Nb_Criteres_KO"] == 1:
            return "Flop modéré"
        return "OK"

    out["Severite"] = out.apply(severity, axis=1)
    out["Score_Label"] = out.apply(
        lambda r: "4/4" if r["C4_Rupture"] else f"{r['Nb_Criteres_KO']}/{r['Nb_Criteres_Applicables']}",
        axis=1,
    )

    return out


# ============================================================
# 3. MOTEUR DE COMMENTAIRE — RENTABILITÉ (niveau Rayon)
# ============================================================

def classe_ca(vs_n1, seuil_ca: float) -> str:
    if pd.isna(vs_n1):
        return "inconnu"
    if vs_n1 >= 0:
        return "croissance"
    if vs_n1 <= seuil_ca:
        return "decrochage"
    return "leger_recul"


def classe_marge(delta_pt, seuil_marge: float) -> str:
    if pd.isna(delta_pt):
        return "inconnu"
    if delta_pt >= abs(seuil_marge):
        return "amelioration"
    if delta_pt <= seuil_marge:
        return "degradation"
    return "stable"


COMMENT_MATRIX = {
    ("croissance", "amelioration"): "Rayon performant : croissance rentable, le CA progresse et la marge s'améliore.",
    ("croissance", "stable"): "Croissance saine, marge maîtrisée malgré la hausse d'activité.",
    ("croissance", "degradation"): "Croissance en trompe-l'œil : le CA progresse mais au prix d'une érosion de la marge — vérifier pression prix/promo.",
    ("leger_recul", "amelioration"): "Repli limité mais pilotage marge efficace : le rayon protège sa rentabilité malgré le tassement du CA.",
    ("leger_recul", "stable"): "Rayon stable, sans signal d'alerte notable sur la période.",
    ("leger_recul", "degradation"): "Vigilance : CA en léger repli et marge qui s'effrite simultanément — surveiller l'évolution.",
    ("decrochage", "amelioration"): "Décrochage de CA compensé par un pilotage marge défensif (moins de volume, marge préservée voire renforcée).",
    ("decrochage", "stable"): "Décrochage de CA préoccupant, marge stable — le problème est un problème de volume, pas de rentabilité unitaire.",
    ("decrochage", "degradation"): "Rayon en difficulté structurelle : perte de CA doublée d'une érosion de marge — cumul des deux signaux, priorité d'action.",
}


def budget_suffix(vs_bgt, seuil_bgt: float) -> str:
    if pd.isna(vs_bgt):
        return " — pas de budget alloué sur ce rayon."
    if vs_bgt >= 0:
        return " — objectif budgétaire atteint."
    if vs_bgt <= seuil_bgt:
        return " — loin de l'objectif budgétaire, à traiter en priorité."
    return " — légèrement en retard sur l'objectif budgétaire."


def build_rentabilite_comment(row: pd.Series, seuil_ca: float, seuil_bgt: float, seuil_marge: float) -> str:
    delta_marge = compute_delta_marge_pts(
        pd.Series([row.get("Taux de Marge")]), pd.Series([row.get("Taux de Marge N-1")])
    ).iloc[0]

    c_ca = classe_ca(row.get("Vs N-1 (%)"), seuil_ca)
    c_marge = classe_marge(delta_marge, seuil_marge)

    if c_ca == "inconnu" or c_marge == "inconnu":
        return "Données insuffisantes pour établir un diagnostic de rentabilité sur ce rayon."

    base = COMMENT_MATRIX.get((c_ca, c_marge), "Diagnostic non déterminé.")
    return base + budget_suffix(row.get("Vs Bgt (%)"), seuil_bgt)


# ============================================================
# 4. HELPERS D'AFFICHAGE
# ============================================================

def fmt_fcfa(x) -> str:
    if pd.isna(x):
        return "n/a"
    if abs(x) >= 1_000_000:
        return f"{x / 1_000_000:,.1f}M".replace(",", " ")
    if abs(x) >= 1_000:
        return f"{x / 1_000:,.0f}K".replace(",", " ")
    return f"{x:,.0f}".replace(",", " ")


def fmt_pct(x) -> str:
    if pd.isna(x):
        return "n/a"
    return f"{x * 100:+.1f}%"


def fmt_pt(x) -> str:
    if pd.isna(x):
        return "n/a"
    return f"{x:+.2f} pt"


def kpi_card(label: str, value: str, color: str = COL_TEXT_PRIMARY) -> str:
    return f"""
    <div class="kpi-card">
        <p class="kpi-label">{label}</p>
        <p class="kpi-value" style="color:{color}">{value}</p>
    </div>
    """


def severity_badge(severite: str) -> str:
    style = SEVERITY_STYLE.get(severite, SEVERITY_STYLE["OK"])
    return (
        f'<span class="badge-pill" style="background:{style["bg"]};color:{style["text"]}">'
        f'{style["emoji"]} {severite}</span>'
    )


def variation_color(x) -> str:
    if pd.isna(x):
        return COL_TEXT_SECONDARY
    return COL_GREEN if x >= 0 else COL_RED


# ============================================================
# 5. MAIN — RENDU STREAMLIT (sidebar, KPI, tabs)
# ============================================================

def main():
    with st.sidebar:
        st.markdown("### 💸 Reporting Vente CA")
        st.markdown("---")

        uploaded_file = st.file_uploader("Export Excel (feuille 'Export')", type=["xlsx"])

        if uploaded_file is None:
            st.info("Charge un fichier `data.xlsx` pour démarrer.")
            st.stop()

        data = load_data(uploaded_file)
        df_global_raw = data["global"]
        df_rayon_raw = data["rayon"]
        df_couple_raw = data["couple"]

        if data["missing_cols"]:
            st.warning(f"Colonnes manquantes dans l'export : {', '.join(data['missing_cols'])}")

        if df_couple_raw.empty:
            st.error("Aucune ligne de niveau Couple Magasin x Rayon détectée. Vérifie la structure du fichier.")
            st.stop()

        st.markdown("#### Filtres")
        societes = sorted(df_couple_raw["Société"].dropna().unique().tolist())
        societe_sel = st.selectbox("Société", options=societes) if societes else None

        rayons_dispo = sorted(df_couple_raw["Rayon_Libelle"].dropna().unique().tolist())
        rayons_sel = st.multiselect("Rayon", options=rayons_dispo, default=rayons_dispo)

        formats_dispo = sorted(df_couple_raw["Format"].dropna().unique().tolist())
        formats_sel = st.multiselect("Format", options=formats_dispo, default=formats_dispo)

        magasins_dispo = sorted(
            df_couple_raw.loc[df_couple_raw["Format"].isin(formats_sel), "Site_Libelle"].dropna().unique().tolist()
        )
        magasins_sel = st.multiselect("Magasin", options=magasins_dispo, default=magasins_dispo)

        st.markdown("---")
        st.markdown("#### Seuils d'alerte")
        seuil_ca_pct = st.slider("Décrochage CA vs N-1 (%)", min_value=-50, max_value=0, value=-10, step=1)
        seuil_bgt_pct = st.slider("Écart vs Budget (%)", min_value=-50, max_value=0, value=-10, step=1)
        seuil_marge_pt = st.slider("Dégradation marge (points)", min_value=-5.0, max_value=0.0, value=-0.8, step=0.1)

        seuil_ca = seuil_ca_pct / 100
        seuil_bgt = seuil_bgt_pct / 100
        seuil_marge = seuil_marge_pt

    # --- Application des filtres ---
    df_couple_f = df_couple_raw[
        (df_couple_raw["Société"] == societe_sel)
        & (df_couple_raw["Rayon_Libelle"].isin(rayons_sel))
        & (df_couple_raw["Format"].isin(formats_sel))
        & (df_couple_raw["Site_Libelle"].isin(magasins_sel))
    ].copy()

    df_rayon_f = df_rayon_raw[
        (df_rayon_raw["Société"] == societe_sel) & (df_rayon_raw["Rayon_Libelle"].isin(rayons_sel))
    ].copy()

    df_global_f = df_global_raw[df_global_raw["Société"] == societe_sel].copy()

    if df_couple_f.empty:
        st.warning("Aucune donnée ne correspond aux filtres sélectionnés.")
        st.stop()

    df_flops = compute_flops(df_couple_f, seuil_ca, seuil_bgt, seuil_marge)

    # --- Header + KPI globaux ---
    st.markdown(f"## Synthèse Commerciale — {societe_sel}")
    st.caption("Niveaux : Global · Rayon · Couple Magasin x Rayon — Détection automatique des Flops")

    if not df_global_f.empty:
        g = df_global_f.iloc[0]
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.markdown(kpi_card("CA total", fmt_fcfa(g["CA"]), COL_BLUE), unsafe_allow_html=True)
        with c2:
            st.markdown(kpi_card("Vs N-1", fmt_pct(g["Vs N-1 (%)"]), variation_color(g["Vs N-1 (%)"])), unsafe_allow_html=True)
        with c3:
            st.markdown(kpi_card("Vs Budget", fmt_pct(g["Vs Bgt (%)"]), variation_color(g["Vs Bgt (%)"])), unsafe_allow_html=True)
        with c4:
            st.markdown(kpi_card("Marge totale", fmt_fcfa(g["Marge"]), COL_PURPLE), unsafe_allow_html=True)
        with c5:
            st.markdown(kpi_card("Taux de marge", fmt_pct(g["Taux de Marge"]).replace("+", ""), COL_TEXT_PRIMARY), unsafe_allow_html=True)
    else:
        st.info("Pas de ligne Global disponible pour cette société (vérifier les filtres).")

    nb_critique = int((df_flops["Severite"] == "Critique").sum())
    if nb_critique > 0:
        st.markdown(
            f"""<div style="background:{COL_RED_BG};border:0.5px solid {COL_RED};border-radius:10px;
            padding:0.75rem 1rem;margin:12px 0;color:{COL_RED};">
            🔴 {nb_critique} rupture(s) totale de CA détectée(s) sur le périmètre — voir onglet Flops</div>""",
            unsafe_allow_html=True,
        )

    st.markdown("---")

    tab_exec, tab_flops, tab_rayons, tab_magasins, tab_detail = st.tabs(
        ["🎯 Exécutive", "🚩 Flops", "🏷️ Rayons", "🏬 Magasins", "📋 Détail"]
    )

    # ---------------- TAB EXÉCUTIVE ----------------
    with tab_exec:
        col_a, col_b = st.columns([1, 2])

        with col_a:
            st.markdown("#### Répartition sévérité")
            counts = df_flops["Severite"].value_counts()
            order = ["Critique", "Flop majeur", "Flop modéré", "OK"]
            counts = counts.reindex(order).fillna(0)
            colors_map = [SEVERITY_STYLE[s]["text"] for s in order]

            fig_donut = go.Figure(
                data=[go.Pie(labels=order, values=counts.values, hole=0.55, marker=dict(colors=colors_map))]
            )
            fig_donut.update_layout(template=PLOTLY_TEMPLATE, height=300, showlegend=True, margin=dict(t=10, b=10))
            st.plotly_chart(fig_donut, use_container_width=True)

        with col_b:
            st.markdown("#### Top 5 Flops les plus sévères")
            top5 = (
                df_flops[df_flops["Severite"] != "OK"]
                .sort_values(["Nb_Criteres_KO", "Vs N-1 (%)"], ascending=[False, True])
                .head(5)
            )
            if top5.empty:
                st.markdown('<div class="flop-row"><span>Aucun flop détecté sur ce périmètre 🎉</span></div>', unsafe_allow_html=True)
            else:
                for _, r in top5.iterrows():
                    st.markdown(
                        f"""<div class="flop-row">
                            <div style="display:flex;align-items:center;gap:10px">
                                {severity_badge(r['Severite'])}
                                <span>{r['Site_Libelle']} — {r['Rayon_Libelle']}</span>
                            </div>
                            <span style="color:{COL_TEXT_SECONDARY}">Vs N-1 {fmt_pct(r['Vs N-1 (%)'])} · Score {r['Score_Label']}</span>
                        </div>""",
                        unsafe_allow_html=True,
                    )

        st.markdown("#### Commentaire rentabilité par rayon")
        rayon_cols = st.columns(min(4, max(1, len(df_rayon_f))))
        for i, (_, r) in enumerate(df_rayon_f.iterrows()):
            comment = build_rentabilite_comment(r, seuil_ca, seuil_bgt, seuil_marge)
            with rayon_cols[i % len(rayon_cols)]:
                st.markdown(
                    f"""<div class="rayon-card">
                        <p style="font-weight:600;margin:0 0 6px 0">{r['Rayon_Libelle']}</p>
                        <p style="font-size:13px;color:{COL_TEXT_SECONDARY};margin:0">{comment}</p>
                    </div>""",
                    unsafe_allow_html=True,
                )

    # ---------------- TAB FLOPS ----------------
    with tab_flops:
        filtre_severite = st.radio(
            "Filtrer par sévérité", options=["Tous", "Critique", "Flop majeur", "Flop modéré"], horizontal=True
        )
        df_flops_view = df_flops if filtre_severite == "Tous" else df_flops[df_flops["Severite"] == filtre_severite]
        df_flops_view = df_flops_view.sort_values(["Nb_Criteres_KO", "Vs N-1 (%)"], ascending=[False, True])

        st.markdown(f"**{len(df_flops_view)}** couple(s) Magasin x Rayon affiché(s)")

        for _, r in df_flops_view.iterrows():
            with st.expander(
                f"{SEVERITY_STYLE[r['Severite']]['emoji']} {r['Site_Libelle']} — {r['Rayon_Libelle']}  "
                f"(Score {r['Score_Label']})"
            ):
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("CA", fmt_fcfa(r["CA"]), fmt_pct(r["Vs N-1 (%)"]))
                c2.metric("Vs Budget", fmt_pct(r["Vs Bgt (%)"]) if r["C2_Applicable"] else "n/a")
                c3.metric("Delta Marge", fmt_pt(r["Delta_Marge_pt"]) if r["C3_Applicable"] else "n/a")
                c4.metric("Sévérité", r["Severite"])

                st.markdown("**Détail des critères :**")
                st.markdown(f"- {'❌' if r['C4_Rupture'] else '✅'} C4 — Rupture / fermeture totale de CA")
                st.markdown(f"- {'❌' if r['C1_Decrochage_CA'] else '✅'} C1 — Décrochage CA vs N-1 (seuil {seuil_ca_pct}%)")
                c2_txt = "n/a (pas de budget alloué)" if not r["C2_Applicable"] else ("❌" if r["C2_Ecart_Budget"] else "✅")
                st.markdown(f"- {c2_txt} C2 — Écart vs Budget (seuil {seuil_bgt_pct}%)")
                c3_txt = "n/a (marge N-1 ou N manquante)" if not r["C3_Applicable"] else ("❌" if r["C3_Degradation_Marge"] else "✅")
                st.markdown(f"- {c3_txt} C3 — Dégradation marge (seuil {seuil_marge_pt} pt)")

        st.markdown("---")
        st.markdown("#### Vue globale — Vs N-1 x Δ Marge")
        scatter_df = df_flops.dropna(subset=["Vs N-1 (%)", "Delta_Marge_pt", "CA"])
        if not scatter_df.empty:
            fig_scatter = px.scatter(
                scatter_df,
                x="Vs N-1 (%)",
                y="Delta_Marge_pt",
                size="CA",
                color="Rayon_Libelle",
                hover_name="Site_Libelle",
                labels={"Vs N-1 (%)": "Vs N-1 (%)", "Delta_Marge_pt": "Δ Marge (pt)"},
            )
            fig_scatter.update_layout(template=PLOTLY_TEMPLATE, height=420)
            fig_scatter.add_vline(x=seuil_ca, line_dash="dash", line_color=COL_RED)
            fig_scatter.add_hline(y=seuil_marge, line_dash="dash", line_color=COL_RED)
            st.plotly_chart(fig_scatter, use_container_width=True)
        else:
            st.info("Pas assez de données valides pour tracer le scatter plot.")

        st.markdown("---")
        export_cols = [
            "Site_Libelle", "Rayon_Libelle", "Format", "CA", "CA N-1", "Vs N-1 (%)",
            "Budget", "Vs Bgt (%)", "Delta_Marge_pt", "Severite", "Score_Label",
        ]
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_flops[export_cols].to_excel(writer, index=False, sheet_name="Flops")
        st.download_button(
            "📥 Export Excel Flops",
            data=buffer.getvalue(),
            file_name=f"flops_{societe_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ---------------- TAB RAYONS ----------------
    with tab_rayons:
        if df_rayon_f.empty:
            st.info("Aucune donnée rayon pour ce périmètre.")
        else:
            display_rayon = df_rayon_f[
                ["Rayon_Libelle", "CA", "CA N-1", "Vs N-1 (%)", "Budget", "Vs Bgt (%)", "Marge", "Taux de Marge"]
            ].copy()
            display_rayon["Vs N-1 (%)"] = display_rayon["Vs N-1 (%)"].apply(fmt_pct)
            display_rayon["Vs Bgt (%)"] = display_rayon["Vs Bgt (%)"].apply(fmt_pct)
            display_rayon["Taux de Marge"] = display_rayon["Taux de Marge"].apply(lambda x: fmt_pct(x).replace("+", ""))
            display_rayon["CA"] = display_rayon["CA"].apply(fmt_fcfa)
            display_rayon["CA N-1"] = display_rayon["CA N-1"].apply(fmt_fcfa)
            display_rayon["Budget"] = display_rayon["Budget"].apply(fmt_fcfa)
            display_rayon["Marge"] = display_rayon["Marge"].apply(fmt_fcfa)
            st.dataframe(display_rayon, use_container_width=True, hide_index=True)

            fig_bar = go.Figure()
            fig_bar.add_bar(name="Vs N-1 (%)", x=df_rayon_f["Rayon_Libelle"], y=df_rayon_f["Vs N-1 (%)"] * 100, marker_color=COL_BLUE)
            fig_bar.add_bar(name="Vs Budget (%)", x=df_rayon_f["Rayon_Libelle"], y=df_rayon_f["Vs Bgt (%)"] * 100, marker_color=COL_PURPLE)
            fig_bar.update_layout(template=PLOTLY_TEMPLATE, barmode="group", height=380, yaxis_title="%")
            st.plotly_chart(fig_bar, use_container_width=True)

    # ---------------- TAB MAGASINS ----------------
    with tab_magasins:
        agg_magasin = (
            df_flops.groupby(["Site_Libelle", "Format"], as_index=False)
            .agg(
                CA=("CA", "sum"),
                CA_N1=("CA N-1", "sum"),
                Marge=("Marge", "sum"),
                Nb_Flops=("Severite", lambda s: (s != "OK").sum()),
                Nb_Critiques=("Severite", lambda s: (s == "Critique").sum()),
            )
            .sort_values("Nb_Flops", ascending=False)
        )
        agg_magasin["Vs N-1 (%)"] = np.where(
            agg_magasin["CA_N1"] > 0, (agg_magasin["CA"] - agg_magasin["CA_N1"]) / agg_magasin["CA_N1"], np.nan
        )

        display_mag = agg_magasin.copy()
        display_mag["CA"] = display_mag["CA"].apply(fmt_fcfa)
        display_mag["Marge"] = display_mag["Marge"].apply(fmt_fcfa)
        display_mag["Vs N-1 (%)"] = display_mag["Vs N-1 (%)"].apply(fmt_pct)
        st.dataframe(
            display_mag[["Site_Libelle", "Format", "CA", "Vs N-1 (%)", "Marge", "Nb_Flops", "Nb_Critiques"]],
            use_container_width=True,
            hide_index=True,
        )

    # ---------------- TAB DÉTAIL ----------------
    with tab_detail:
        st.markdown("Détail brut (toutes colonnes) — pour audit et export libre")
        st.dataframe(df_flops, use_container_width=True, hide_index=True)
        csv_buffer = df_flops.to_csv(index=False).encode("utf-8-sig")
        st.download_button("📥 Export CSV complet", data=csv_buffer, file_name=f"detail_{societe_sel}.csv", mime="text/csv")


# ============================================================
# 6. TESTS UNITAIRES / ASSERTIONS
# ============================================================
# Isolés du flux Streamlit (aucun appel top-level) : ne s'exécutent que si
# le fichier est lancé directement avec RUN_DASHBOARD_TESTS=1, ou collectés
# par pytest (fonctions préfixées test_).

def _make_couple_row(**overrides) -> pd.DataFrame:
    base = dict(
        **{
            "Société": "TEST", "Rayon": "010 - BOISSON", "Site": "999 - Magasin Test",
            "CA N-1": 1000.0, "Budget": 1000.0, "CA": 1000.0,
            "Vs N-1 (%)": 0.0, "Vs Bgt (%)": 0.0,
            "Taux de Marge N-1": 0.20, "Taux de Marge": 0.20,
        }
    )
    base.update(overrides)
    return pd.DataFrame([base])


def test_c4_rupture_prioritaire_sur_les_autres():
    """Une rupture totale doit toujours être classée CRITIQUE, même si les
    autres critères ne seraient pas déclenchés isolément."""
    df = _make_couple_row(CA=np.nan, **{"Vs N-1 (%)": -1.0})
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert res.loc[0, "Severite"] == "Critique"
    assert res.loc[0, "C4_Rupture"] == True  # noqa: E712


def test_flop_majeur_deux_criteres():
    """CA en décrochage ET marge en dégradation -> 2 critères KO -> Flop majeur."""
    df = _make_couple_row(**{
        "Vs N-1 (%)": -0.15, "Vs Bgt (%)": 0.05,
        "Taux de Marge N-1": 0.20, "Taux de Marge": 0.18,  # -2 pt
    })
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert res.loc[0, "Severite"] == "Flop majeur"
    assert res.loc[0, "Nb_Criteres_KO"] == 2


def test_flop_modere_un_critere():
    df = _make_couple_row(**{"Vs N-1 (%)": -0.15, "Vs Bgt (%)": 0.05})
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert res.loc[0, "Severite"] == "Flop modéré"
    assert res.loc[0, "Nb_Criteres_KO"] == 1


def test_ok_aucun_critere():
    df = _make_couple_row()
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert res.loc[0, "Severite"] == "OK"
    assert res.loc[0, "Nb_Criteres_KO"] == 0


def test_budget_nan_exclu_du_scoring():
    """Un couple sans Budget (ex: Supeco) ne doit pas être pénalisé sur C2 :
    C2 doit être non-applicable et ne pas compter dans Nb_Criteres_KO."""
    df = _make_couple_row(Budget=np.nan, **{"Vs Bgt (%)": np.nan})
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert res.loc[0, "C2_Applicable"] == False  # noqa: E712
    assert res.loc[0, "Nb_Criteres_Applicables"] == 2  # C1 + C3 uniquement
    assert res.loc[0, "Severite"] == "OK"


def test_marge_nan_exclu_du_scoring():
    df = _make_couple_row(**{"Taux de Marge": np.nan})
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert res.loc[0, "C3_Applicable"] == False  # noqa: E712


def test_ca_nul_ne_plante_pas_le_calcul():
    """Vérifie qu'un CA nul (donc taux de marge nul) ne provoque aucune
    exception lors du calcul du delta de marge (pas de division impliquée)."""
    df = _make_couple_row(CA=0.0, Marge=0.0, **{"Taux de Marge": 0.0})
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert len(res) == 1
    assert res.loc[0, "Severite"] == "Critique"  # CA=0 avec CA N-1>0 -> rupture


def test_commentaire_rentabilite_decrochage_marge_ok():
    row = pd.Series({
        "Vs N-1 (%)": -0.20, "Vs Bgt (%)": 0.03,
        "Taux de Marge N-1": 0.18, "Taux de Marge": 0.19,  # +1pt -> amélioration
    })
    comment = build_rentabilite_comment(row, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert "pilotage marge défensif" in comment
    assert "objectif budgétaire atteint" in comment


def test_commentaire_rentabilite_donnees_manquantes():
    row = pd.Series({"Vs N-1 (%)": np.nan, "Vs Bgt (%)": np.nan,
                      "Taux de Marge N-1": np.nan, "Taux de Marge": np.nan})
    comment = build_rentabilite_comment(row, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert "insuffisantes" in comment


def test_split_code_libelle_gere_nan():
    s = pd.Series(["010 - BOISSON", np.nan, "SansSeparateur"])
    res = split_code_libelle(s)
    assert pd.isna(res.loc[1, "Code"])
    assert res.loc[2, "Libelle"] == "SansSeparateur"


def test_detect_format_variantes():
    assert detect_format("10301 - Hyper Marcory") == "Hyper"
    assert detect_format("10705 - Market 7 Décembre") == "Market"
    assert detect_format("10601 - Supeco Niangon") == "Supeco"
    assert detect_format(None) == "Autre"


def test_load_data_exclut_lignes_parasites():
    """Vérifie que le grand total société et les lignes de footer sont bien
    exclus, et que les 3 niveaux (global/rayon/couple) sont correctement
    séparés, sur un jeu de données synthétique reproduisant la structure
    réelle de l'export."""
    rows = [
        {"Société": "SOC", "Rayon": "Total", "Site": np.nan, "CA": 100, "CA N-1": 90},
        {"Société": "SOC", "Rayon": "010 - BOISSON", "Site": "Total", "CA": 60, "CA N-1": 55},
        {"Société": "SOC", "Rayon": "010 - BOISSON", "Site": "111 - Hyper Test", "CA": 60, "CA N-1": 55},
        {"Société": "Total", "Rayon": np.nan, "Site": np.nan, "CA": 100, "CA N-1": 90},
        {"Société": np.nan, "Rayon": np.nan, "Site": np.nan, "CA": np.nan, "CA N-1": np.nan},
        {"Société": "Filtres appliqués : \naxes...", "Rayon": np.nan, "Site": np.nan, "CA": np.nan, "CA N-1": np.nan},
    ]
    df = pd.DataFrame(rows)
    for col in EXPECTED_COLS:
        if col not in df.columns:
            df[col] = np.nan

    buf = io.BytesIO()
    df.to_excel(buf, sheet_name="Export", index=False)
    buf.seek(0)

    result = _load_data_impl(buf)
    assert len(result["global"]) == 1
    assert len(result["rayon"]) == 1
    assert len(result["couple"]) == 1
    assert result["couple"].loc[0, "Format"] == "Hyper"


def run_all_tests():
    """Exécute tous les tests locaux et affiche un résumé (utilisable sans
    dépendance pytest obligatoire)."""
    tests = [obj for name, obj in list(globals().items()) if name.startswith("test_") and callable(obj)]
    passed, failed = 0, []
    for t in tests:
        try:
            t()
            passed += 1
            print(f"  OK   - {t.__name__}")
        except AssertionError as e:
            failed.append((t.__name__, str(e)))
            print(f"  FAIL - {t.__name__} : {e}")
    print(f"\n{passed}/{len(tests)} tests passés.")
    if failed:
        raise SystemExit(1)


# ============================================================
# 7. POINT D'ENTRÉE
# ============================================================
if __name__ == "__main__":
    if os.environ.get("RUN_DASHBOARD_TESTS") == "1":
        run_all_tests()
    else:
        main()
