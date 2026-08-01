# -*- coding: utf-8 -*-
"""
18_💸_Reporting_Vente CA.py — V2
============================================================
SmartBuyer Hub — Module Reporting Commercial (Synthèse Exécutive & Flops)

V2 — améliorations ergonomie par rapport à la V1 :
    - Table maître-détail dans l'onglet Flops (au lieu de 48 expanders)
    - Recherche rapide magasin/rayon
    - Presets de seuils (Strict / Standard / Souple) + persistance session
    - Lecteur I/O robuste (xlsx et csv, encodage auto) façon utils_io.py
    - Mapping Format par code magasin (fiable) + fallback mot-clé
    - Export "Synthèse COPIL" en 1 clic (classeur 3 onglets)
    - Scatter plot avec quadrants annotés

Règles de flop (validées) :
    C1 - Décrochage CA vs N-1   : Vs N-1 (%)  <= seuil CA
    C2 - Écart vs Budget        : Vs Bgt (%)  <= seuil Budget (ignoré si Budget NaN)
    C3 - Dégradation marge      : Delta Marge (pts) <= seuil Marge
    C4 - Rupture / fermeture    : CA NaN/0 alors que CA N-1 > 0 (prioritaire)

Architecture :
    0-4  : config, palette, I/O, moteurs de calcul, helpers
           -> pur Python/pandas, importable et testable sans Streamlit actif
    5    : main() -> rendu Streamlit
    6    : tests unitaires (fonctions test_*)
    7    : point d'entrée
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

COL_BG_PAGE = "#0B0B0F"
COL_BG_CARD = "#151519"
COL_BG_CARD_HOVER = "#1C1C22"
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

SEVERITY_STYLE = {
    "Critique": {"text": COL_RED, "bg": COL_RED_BG, "emoji": "🔴"},
    "Flop majeur": {"text": COL_ORANGE, "bg": COL_ORANGE_BG, "emoji": "🟠"},
    "Flop modéré": {"text": COL_AMBER, "bg": COL_AMBER_BG, "emoji": "🟡"},
    "OK": {"text": COL_GREEN, "bg": COL_GREEN_BG, "emoji": "🟢"},
}
SEVERITY_ORDER = ["Critique", "Flop majeur", "Flop modéré", "OK"]

# Presets de seuils (CA %, Budget %, Marge pt) — cumul simple validé
THRESHOLD_PRESETS = {
    "Strict":   {"ca": -5,  "bgt": -5,  "marge": -0.5},
    "Standard": {"ca": -10, "bgt": -10, "marge": -0.8},
    "Souple":   {"ca": -15, "bgt": -15, "marge": -1.2},
}

CUSTOM_CSS = f"""
<style>
    .stApp {{ background-color: {COL_BG_PAGE}; }}
    section[data-testid="stSidebar"] {{
        background-color: {COL_BG_CARD};
        border-right: 0.5px solid {COL_BORDER};
    }}
    h1, h2, h3, h4, p, span, div, label {{ color: {COL_TEXT_PRIMARY}; }}
    .kpi-card {{
        background-color: {COL_BG_CARD};
        border: 0.5px solid {COL_BORDER};
        border-radius: 10px;
        padding: 0.85rem 1rem;
        margin-bottom: 8px;
    }}
    .kpi-label {{ font-size: 12px; color: {COL_TEXT_SECONDARY}; margin: 0 0 6px 0; }}
    .kpi-value {{ font-size: 22px; font-weight: 600; color: {COL_TEXT_PRIMARY}; margin: 0; }}
    .badge-pill {{
        font-size: 11px; padding: 3px 10px; border-radius: 20px;
        display: inline-block; font-weight: 600;
    }}
    .rayon-card {{
        background-color: {COL_BG_CARD}; border: 0.5px solid {COL_BORDER};
        border-radius: 12px; padding: 1rem; margin-bottom: 8px; height: 100%;
    }}
    .detail-panel {{
        background-color: {COL_BG_CARD}; border: 0.5px solid {COL_BORDER};
        border-radius: 12px; padding: 1.1rem 1.25rem;
    }}
    .criterion-line {{
        display: flex; justify-content: space-between; padding: 6px 0;
        border-bottom: 0.5px solid {COL_BORDER}; font-size: 13.5px;
    }}
    .criterion-line:last-child {{ border-bottom: none; }}
    div[data-testid="stMetric"] {{
        background-color: {COL_BG_CARD}; border: 0.5px solid {COL_BORDER};
        border-radius: 10px; padding: 0.75rem 1rem;
    }}
    .preset-caption {{ font-size: 11px; color: {COL_TEXT_SECONDARY}; margin-top: -6px; }}
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
# 1. LECTEUR I/O ROBUSTE (xlsx / csv, encodage auto) — façon utils_io.py
# ============================================================

EXPECTED_COLS = [
    "Société", "Rayon", "Site", "CA N-1", "Budget", "CA", "Poids",
    "Vs N-1 (%)", "Vs Bgt (%)", "Marge N-1", "Marge",
    "Taux de Marge N-1", "Taux de Marge", "Taux de Marge N Vs N-1",
    "Débit N-1", "Débit", "Vs N-1 (%).1", "Panier N-1", "Panier",
    "Panier N Vs N-1", "Panier Qté N-1", "Panier Qté",
    "Panier Qté N Vs N-1", "Volume N-1", "Volume", "Volume N Vs N-1",
]

# Mapping Format par code magasin — priorité sur la détection par mot-clé.
# À compléter/ajuster avec la liste officielle des 12 sites du réseau.
FORMAT_BY_CODE = {}

FORMAT_KEYWORDS = {"hyper": "Hyper", "market": "Market", "supeco": "Supeco"}


class DataLoadError(Exception):
    """Erreur métier levée par le lecteur I/O — message affichable tel quel à l'utilisateur."""
    pass


def read_any_export(file, sheet_name: str = "Export") -> pd.DataFrame:
    """Lecteur universel encoding-safe (xlsx ou csv), aligné sur le pattern
    utils_io.py utilisé dans les autres modules SmartBuyer Hub.

    - .xlsx / .xls -> pandas.read_excel sur la feuille demandée
    - .csv         -> tentative UTF-8, puis UTF-8-SIG, puis Latin-1,
                       avec détection auto du séparateur (`;` ou `,`)
    Lève DataLoadError avec un message clair plutôt qu'une traceback brute.
    """
    name = getattr(file, "name", "") or ""
    ext = name.lower().rsplit(".", 1)[-1] if "." in name else "xlsx"

    try:
        if ext in ("xlsx", "xls"):
            return pd.read_excel(file, sheet_name=sheet_name)

        if ext == "csv":
            raw_bytes = file.read()
            for encoding in ("utf-8", "utf-8-sig", "latin-1"):
                try:
                    text = raw_bytes.decode(encoding)
                    sep = ";" if text.count(";") > text.count(",") else ","
                    return pd.read_csv(io.StringIO(text), sep=sep)
                except (UnicodeDecodeError, pd.errors.ParserError):
                    continue
            raise DataLoadError("Impossible de décoder le CSV (encodages testés : UTF-8, UTF-8-SIG, Latin-1).")

        raise DataLoadError(f"Extension '.{ext}' non supportée — utilise .xlsx ou .csv.")

    except ValueError as e:
        if "Worksheet" in str(e) or "sheet" in str(e).lower():
            raise DataLoadError(f"Feuille '{sheet_name}' introuvable dans le fichier. Vérifie le nom de l'onglet Excel.")
        raise DataLoadError(f"Erreur de lecture du fichier : {e}")


def split_code_libelle(serie: pd.Series) -> pd.DataFrame:
    """Sépare 'CODE - Libellé' en 2 colonnes Code / Libellé. Gère NaN et absence de séparateur."""
    serie = serie.astype("string")
    split = serie.str.split(" - ", n=1, expand=True)
    if split.shape[1] == 1:
        split[1] = np.nan
    code = split[0].str.strip()
    libelle = split[1].str.strip().fillna(serie)
    return pd.DataFrame({"Code": code, "Libelle": libelle})


def detect_format(code, libelle) -> str:
    """Format magasin : priorité au mapping explicite par code, fallback mot-clé sur le libellé."""
    if code in FORMAT_BY_CODE:
        return FORMAT_BY_CODE[code]
    if not isinstance(libelle, str):
        return "Autre"
    lib_low = libelle.lower()
    for kw, fmt in FORMAT_KEYWORDS.items():
        if kw in lib_low:
            return fmt
    return "Autre"


def _load_data_impl(file, sheet_name: str = "Export") -> dict:
    """Implémentation pure (sans cache Streamlit) — testable hors contexte Streamlit."""
    raw = read_any_export(file, sheet_name=sheet_name)

    missing_cols = [c for c in EXPECTED_COLS if c not in raw.columns]
    required_minimum = {"Société", "Rayon", "Site", "CA", "CA N-1"}
    if not required_minimum.issubset(set(raw.columns)):
        raise DataLoadError(
            f"Colonnes indispensables manquantes : {sorted(required_minimum - set(raw.columns))}. "
            "Vérifie que l'export correspond au format attendu."
        )

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
        rs = split_code_libelle(d["Rayon"])
        d["Rayon_Code"] = rs["Code"].values
        d["Rayon_Libelle"] = rs["Libelle"].values

    ss = split_code_libelle(df_couple["Site"])
    df_couple["Site_Code"] = ss["Code"].values
    df_couple["Site_Libelle"] = ss["Libelle"].values
    df_couple["Format"] = [
        detect_format(c, l) for c, l in zip(df_couple["Site_Code"], df_couple["Site_Libelle"])
    ]

    if df_couple.empty:
        raise DataLoadError(
            "Aucune ligne de niveau Couple Magasin x Rayon détectée. "
            "Vérifie que l'export contient bien 3 niveaux (Global / Rayon / Magasin x Rayon)."
        )

    return {
        "global": df_global.reset_index(drop=True),
        "rayon": df_rayon.reset_index(drop=True),
        "couple": df_couple.reset_index(drop=True),
        "missing_cols": missing_cols,
    }


@st.cache_data(show_spinner="Chargement des données...")
def load_data(file, sheet_name: str = "Export") -> dict:
    return _load_data_impl(file, sheet_name=sheet_name)


# ============================================================
# 2. MOTEUR DE CALCUL — FLOPS & SÉVÉRITÉ (niveau Couple)
# ============================================================

def compute_delta_marge_pts(taux_marge: pd.Series, taux_marge_n1: pd.Series) -> pd.Series:
    return (taux_marge - taux_marge_n1) * 100


def compute_flops(df: pd.DataFrame, seuil_ca: float, seuil_bgt: float, seuil_marge: float) -> pd.DataFrame:
    """seuil_ca, seuil_bgt : négatifs (ex -0.10) ; seuil_marge : négatif, en points (ex -0.8)."""
    out = df.copy()
    out["Delta_Marge_pt"] = compute_delta_marge_pts(out["Taux de Marge"], out["Taux de Marge N-1"])

    ca_is_missing_or_zero = out["CA"].isna() | (out["CA"] == 0)
    ca_n1_positif = out["CA N-1"].fillna(0) > 0
    out["C4_Rupture"] = ca_is_missing_or_zero & ca_n1_positif

    out["C1_Decrochage_CA"] = (out["Vs N-1 (%)"] <= seuil_ca).fillna(False) & ~out["C4_Rupture"]

    budget_applicable = out["Budget"].notna()
    out["C2_Applicable"] = budget_applicable
    out["C2_Ecart_Budget"] = np.where(budget_applicable, out["Vs Bgt (%)"] <= seuil_bgt, False)
    out["C2_Ecart_Budget"] = out["C2_Ecart_Budget"].astype(bool) & ~out["C4_Rupture"]

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
    return f"""<div class="kpi-card"><p class="kpi-label">{label}</p>
    <p class="kpi-value" style="color:{color}">{value}</p></div>"""


def severity_badge(severite: str) -> str:
    s = SEVERITY_STYLE.get(severite, SEVERITY_STYLE["OK"])
    return f'<span class="badge-pill" style="background:{s["bg"]};color:{s["text"]}">{s["emoji"]} {severite}</span>'


def variation_color(x) -> str:
    if pd.isna(x):
        return COL_TEXT_SECONDARY
    return COL_GREEN if x >= 0 else COL_RED


def build_copil_export(df_global: pd.DataFrame, df_rayon: pd.DataFrame, df_flops: pd.DataFrame,
                        seuil_ca: float, seuil_bgt: float, seuil_marge: float) -> bytes:
    """Classeur Excel 3 onglets pour diffusion COPIL : KPI Global, Synthèse Rayons
    (avec commentaire rentabilité), Flops complets triés par sévérité."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if not df_global.empty:
            g = df_global.iloc[0]
            kpi_df = pd.DataFrame({
                "Indicateur": ["CA total", "Vs N-1 (%)", "Vs Budget (%)", "Marge totale", "Taux de marge"],
                "Valeur": [
                    fmt_fcfa(g.get("CA")), fmt_pct(g.get("Vs N-1 (%)")), fmt_pct(g.get("Vs Bgt (%)")),
                    fmt_fcfa(g.get("Marge")), fmt_pct(g.get("Taux de Marge")).replace("+", ""),
                ],
            })
            kpi_df.to_excel(writer, index=False, sheet_name="KPI Global")

        rayon_export = df_rayon.copy()
        if not rayon_export.empty:
            rayon_export["Commentaire rentabilité"] = rayon_export.apply(
                lambda r: build_rentabilite_comment(r, seuil_ca, seuil_bgt, seuil_marge), axis=1
            )
            cols = ["Rayon_Libelle", "CA", "CA N-1", "Vs N-1 (%)", "Budget", "Vs Bgt (%)",
                    "Marge", "Taux de Marge", "Commentaire rentabilité"]
            rayon_export[[c for c in cols if c in rayon_export.columns]].to_excel(
                writer, index=False, sheet_name="Synthese Rayons"
            )

        flops_export = df_flops.sort_values(["Nb_Criteres_KO", "Vs N-1 (%)"], ascending=[False, True])
        cols = ["Site_Libelle", "Rayon_Libelle", "Format", "CA", "CA N-1", "Vs N-1 (%)",
                "Budget", "Vs Bgt (%)", "Delta_Marge_pt", "Severite", "Score_Label"]
        flops_export[[c for c in cols if c in flops_export.columns]].to_excel(
            writer, index=False, sheet_name="Flops"
        )
    return buffer.getvalue()


# ============================================================
# 5. MAIN — RENDU STREAMLIT
# ============================================================

def _init_threshold_state():
    if "seuil_ca_pct" not in st.session_state:
        preset = THRESHOLD_PRESETS["Standard"]
        st.session_state.seuil_ca_pct = preset["ca"]
        st.session_state.seuil_bgt_pct = preset["bgt"]
        st.session_state.seuil_marge_pt = preset["marge"]


def _apply_preset(name: str):
    preset = THRESHOLD_PRESETS[name]
    st.session_state.seuil_ca_pct = preset["ca"]
    st.session_state.seuil_bgt_pct = preset["bgt"]
    st.session_state.seuil_marge_pt = preset["marge"]


def main():
    _init_threshold_state()

    with st.sidebar:
        st.markdown("### 💸 Reporting Vente CA")
        st.markdown("---")

        uploaded_file = st.file_uploader("Export ventes (.xlsx ou .csv)", type=["xlsx", "xls", "csv"])
        if uploaded_file is None:
            st.info("Charge un export pour démarrer.")
            st.stop()

        try:
            data = load_data(uploaded_file)
        except DataLoadError as e:
            st.error(str(e))
            st.stop()

        df_global_raw, df_rayon_raw, df_couple_raw = data["global"], data["rayon"], data["couple"]
        if data["missing_cols"]:
            st.caption(f"⚠ Colonnes absentes (ignorées) : {', '.join(data['missing_cols'])}")

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
        p1, p2, p3 = st.columns(3)
        p1.button("Strict", use_container_width=True, on_click=_apply_preset, args=("Strict",))
        p2.button("Standard", use_container_width=True, on_click=_apply_preset, args=("Standard",))
        p3.button("Souple", use_container_width=True, on_click=_apply_preset, args=("Souple",))
        st.markdown('<p class="preset-caption">Presets — ajustables ensuite via les curseurs</p>', unsafe_allow_html=True)

        seuil_ca_pct = st.slider("Décrochage CA vs N-1 (%)", min_value=-50, max_value=0, step=1, key="seuil_ca_pct")
        seuil_bgt_pct = st.slider("Écart vs Budget (%)", min_value=-50, max_value=0, step=1, key="seuil_bgt_pct")
        seuil_marge_pt = st.slider("Dégradation marge (points)", min_value=-5.0, max_value=0.0, step=0.1, key="seuil_marge_pt")

        seuil_ca, seuil_bgt, seuil_marge = seuil_ca_pct / 100, seuil_bgt_pct / 100, seuil_marge_pt

    # --- Filtrage ---
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

    # --- Header + KPI ---
    top_l, top_r = st.columns([4, 1])
    with top_l:
        st.markdown(f"## Synthèse Commerciale — {societe_sel}")
        st.caption("Global · Rayon · Couple Magasin x Rayon — Détection automatique des Flops")
    with top_r:
        st.markdown("<div style='padding-top:14px'></div>", unsafe_allow_html=True)
        copil_bytes = build_copil_export(df_global_f, df_rayon_f, df_flops, seuil_ca, seuil_bgt, seuil_marge)
        st.download_button(
            "📊 Export COPIL", data=copil_bytes, file_name=f"synthese_copil_{societe_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
        )

    if not df_global_f.empty:
        g = df_global_f.iloc[0]
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.markdown(kpi_card("CA total", fmt_fcfa(g["CA"]), COL_BLUE), unsafe_allow_html=True)
        c2.markdown(kpi_card("Vs N-1", fmt_pct(g["Vs N-1 (%)"]), variation_color(g["Vs N-1 (%)"])), unsafe_allow_html=True)
        c3.markdown(kpi_card("Vs Budget", fmt_pct(g["Vs Bgt (%)"]), variation_color(g["Vs Bgt (%)"])), unsafe_allow_html=True)
        c4.markdown(kpi_card("Marge totale", fmt_fcfa(g["Marge"]), COL_PURPLE), unsafe_allow_html=True)
        c5.markdown(kpi_card("Taux de marge", fmt_pct(g["Taux de Marge"]).replace("+", ""), COL_TEXT_PRIMARY), unsafe_allow_html=True)
    else:
        st.info("Pas de ligne Global disponible pour cette société.")

    nb_critique = int((df_flops["Severite"] == "Critique").sum())
    if nb_critique > 0:
        st.markdown(
            f"""<div style="background:{COL_RED_BG};border:0.5px solid {COL_RED};border-radius:10px;
            padding:0.75rem 1rem;margin:12px 0;color:{COL_RED};">
            🔴 {nb_critique} rupture(s) totale de CA détectée(s) — voir onglet Flops</div>""",
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
            counts = df_flops["Severite"].value_counts().reindex(SEVERITY_ORDER).fillna(0)
            colors_map = [SEVERITY_STYLE[s]["text"] for s in SEVERITY_ORDER]
            fig_donut = go.Figure(data=[go.Pie(labels=SEVERITY_ORDER, values=counts.values, hole=0.55, marker=dict(colors=colors_map))])
            fig_donut.update_layout(template=PLOTLY_TEMPLATE, height=300, showlegend=True, margin=dict(t=10, b=10))
            st.plotly_chart(fig_donut, use_container_width=True)

        with col_b:
            st.markdown("#### Top 5 Flops les plus sévères")
            top5 = df_flops[df_flops["Severite"] != "OK"].sort_values(["Nb_Criteres_KO", "Vs N-1 (%)"], ascending=[False, True]).head(5)
            if top5.empty:
                st.success("Aucun flop détecté sur ce périmètre.")
            else:
                for _, r in top5.iterrows():
                    st.markdown(
                        f"""<div class="detail-panel" style="margin-bottom:6px;padding:0.6rem 1rem;display:flex;align-items:center;justify-content:space-between">
                            <div style="display:flex;align-items:center;gap:10px">{severity_badge(r['Severite'])}
                            <span>{r['Site_Libelle']} — {r['Rayon_Libelle']}</span></div>
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
                    f"""<div class="rayon-card"><p style="font-weight:600;margin:0 0 6px 0">{r['Rayon_Libelle']}</p>
                    <p style="font-size:13px;color:{COL_TEXT_SECONDARY};margin:0">{comment}</p></div>""",
                    unsafe_allow_html=True,
                )

    # ---------------- TAB FLOPS (maître-détail) ----------------
    with tab_flops:
        f_col1, f_col2 = st.columns([2, 1])
        with f_col1:
            recherche = st.text_input("🔎 Recherche magasin ou rayon", placeholder="ex: Yopougon, Boisson...")
        with f_col2:
            filtre_severite = st.selectbox("Sévérité", options=["Tous"] + SEVERITY_ORDER)

        df_view = df_flops.copy()
        if filtre_severite != "Tous":
            df_view = df_view[df_view["Severite"] == filtre_severite]
        if recherche:
            mask = (
                df_view["Site_Libelle"].str.contains(recherche, case=False, na=False)
                | df_view["Rayon_Libelle"].str.contains(recherche, case=False, na=False)
            )
            df_view = df_view[mask]
        df_view = df_view.sort_values(["Nb_Criteres_KO", "Vs N-1 (%)"], ascending=[False, True]).reset_index(drop=True)

        st.caption(f"{len(df_view)} couple(s) affiché(s) — clique une ligne pour voir le détail")

        table_col, detail_col = st.columns([2.2, 1])

        with table_col:
            display_df = pd.DataFrame({
                "Sévérité": df_view["Severite"].apply(lambda s: f"{SEVERITY_STYLE[s]['emoji']} {s}"),
                "Magasin": df_view["Site_Libelle"],
                "Rayon": df_view["Rayon_Libelle"],
                "Format": df_view["Format"],
                "CA": df_view["CA"].apply(fmt_fcfa),
                "Vs N-1": df_view["Vs N-1 (%)"].apply(fmt_pct),
                "Vs Budget": df_view["Vs Bgt (%)"].apply(fmt_pct),
                "Δ Marge": df_view["Delta_Marge_pt"].apply(fmt_pt),
                "Score": df_view["Score_Label"],
            })

            event = st.dataframe(
                display_df, use_container_width=True, hide_index=True, height=460,
                on_select="rerun", selection_mode="single-row",
            )

        with detail_col:
            selected_rows = event.selection.rows if hasattr(event, "selection") else []
            if not selected_rows:
                st.markdown(
                    f"""<div class="detail-panel"><p style="color:{COL_TEXT_SECONDARY};margin:0">
                    Sélectionne une ligne dans le tableau pour afficher le détail des critères.</p></div>""",
                    unsafe_allow_html=True,
                )
            else:
                r = df_view.iloc[selected_rows[0]]
                c4_txt = "❌" if r["C4_Rupture"] else "✅"
                c1_txt = "❌" if r["C1_Decrochage_CA"] else "✅"
                c2_txt = "n/a" if not r["C2_Applicable"] else ("❌" if r["C2_Ecart_Budget"] else "✅")
                c3_txt = "n/a" if not r["C3_Applicable"] else ("❌" if r["C3_Degradation_Marge"] else "✅")

                st.markdown(
                    f"""<div class="detail-panel">
                        <p style="font-weight:600;font-size:15px;margin:0 0 2px 0">{r['Site_Libelle']}</p>
                        <p style="font-size:13px;color:{COL_TEXT_SECONDARY};margin:0 0 12px 0">{r['Rayon_Libelle']} · {r['Format']}</p>
                        {severity_badge(r['Severite'])}
                        <div style="margin-top:14px">
                            <div class="criterion-line"><span>{c4_txt} C4 — Rupture / fermeture</span></div>
                            <div class="criterion-line"><span>{c1_txt} C1 — Décrochage CA vs N-1</span><span>{fmt_pct(r['Vs N-1 (%)'])}</span></div>
                            <div class="criterion-line"><span>{c2_txt} C2 — Écart vs Budget</span><span>{fmt_pct(r['Vs Bgt (%)']) if r['C2_Applicable'] else '—'}</span></div>
                            <div class="criterion-line"><span>{c3_txt} C3 — Dégradation marge</span><span>{fmt_pt(r['Delta_Marge_pt']) if r['C3_Applicable'] else '—'}</span></div>
                        </div>
                    </div>""",
                    unsafe_allow_html=True,
                )

        st.markdown("---")
        st.markdown("#### Vue globale — Vs N-1 x Δ Marge")
        scatter_df = df_flops.dropna(subset=["Vs N-1 (%)", "Delta_Marge_pt", "CA"])
        if not scatter_df.empty:
            fig_scatter = px.scatter(
                scatter_df, x="Vs N-1 (%)", y="Delta_Marge_pt", size="CA", color="Rayon_Libelle",
                hover_name="Site_Libelle", labels={"Vs N-1 (%)": "Vs N-1 (%)", "Delta_Marge_pt": "Δ Marge (pt)"},
            )
            fig_scatter.update_layout(template=PLOTLY_TEMPLATE, height=420)
            fig_scatter.add_vline(x=seuil_ca, line_dash="dash", line_color=COL_RED)
            fig_scatter.add_hline(y=seuil_marge, line_dash="dash", line_color=COL_RED)
            fig_scatter.add_annotation(
                x=scatter_df["Vs N-1 (%)"].min(), y=scatter_df["Delta_Marge_pt"].min(),
                text="Zone à risque", showarrow=False, font=dict(color=COL_RED, size=11),
                xanchor="left", yanchor="bottom",
            )
            st.plotly_chart(fig_scatter, use_container_width=True)
        else:
            st.info("Pas assez de données valides pour tracer le scatter plot.")

        st.markdown("---")
        export_cols = ["Site_Libelle", "Rayon_Libelle", "Format", "CA", "CA N-1", "Vs N-1 (%)",
                       "Budget", "Vs Bgt (%)", "Delta_Marge_pt", "Severite", "Score_Label"]
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_flops[export_cols].to_excel(writer, index=False, sheet_name="Flops")
        st.download_button(
            "📥 Export Excel Flops (périmètre filtré)", data=buffer.getvalue(),
            file_name=f"flops_{societe_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ---------------- TAB RAYONS ----------------
    with tab_rayons:
        if df_rayon_f.empty:
            st.info("Aucune donnée rayon pour ce périmètre.")
        else:
            display_rayon = df_rayon_f[["Rayon_Libelle", "CA", "CA N-1", "Vs N-1 (%)", "Budget", "Vs Bgt (%)", "Marge", "Taux de Marge"]].copy()
            display_rayon["Vs N-1 (%)"] = display_rayon["Vs N-1 (%)"].apply(fmt_pct)
            display_rayon["Vs Bgt (%)"] = display_rayon["Vs Bgt (%)"].apply(fmt_pct)
            display_rayon["Taux de Marge"] = display_rayon["Taux de Marge"].apply(lambda x: fmt_pct(x).replace("+", ""))
            for col in ("CA", "CA N-1", "Budget", "Marge"):
                display_rayon[col] = display_rayon[col].apply(fmt_fcfa)
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
            .agg(CA=("CA", "sum"), CA_N1=("CA N-1", "sum"), Marge=("Marge", "sum"),
                 Nb_Flops=("Severite", lambda s: (s != "OK").sum()),
                 Nb_Critiques=("Severite", lambda s: (s == "Critique").sum()))
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
            use_container_width=True, hide_index=True,
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

def _make_couple_row(**overrides) -> pd.DataFrame:
    base = dict(**{
        "Société": "TEST", "Rayon": "010 - BOISSON", "Site": "999 - Magasin Test",
        "CA N-1": 1000.0, "Budget": 1000.0, "CA": 1000.0,
        "Vs N-1 (%)": 0.0, "Vs Bgt (%)": 0.0,
        "Taux de Marge N-1": 0.20, "Taux de Marge": 0.20,
    })
    base.update(overrides)
    return pd.DataFrame([base])


def test_c4_rupture_prioritaire_sur_les_autres():
    df = _make_couple_row(CA=np.nan, **{"Vs N-1 (%)": -1.0})
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert res.loc[0, "Severite"] == "Critique"
    assert res.loc[0, "C4_Rupture"] == True  # noqa: E712


def test_flop_majeur_deux_criteres():
    df = _make_couple_row(**{"Vs N-1 (%)": -0.15, "Vs Bgt (%)": 0.05, "Taux de Marge N-1": 0.20, "Taux de Marge": 0.18})
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


def test_budget_nan_exclu_du_scoring():
    df = _make_couple_row(Budget=np.nan, **{"Vs Bgt (%)": np.nan})
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert res.loc[0, "C2_Applicable"] == False  # noqa: E712
    assert res.loc[0, "Nb_Criteres_Applicables"] == 2
    assert res.loc[0, "Severite"] == "OK"


def test_marge_nan_exclu_du_scoring():
    df = _make_couple_row(**{"Taux de Marge": np.nan})
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert res.loc[0, "C3_Applicable"] == False  # noqa: E712


def test_ca_nul_ne_plante_pas_le_calcul():
    df = _make_couple_row(CA=0.0, Marge=0.0, **{"Taux de Marge": 0.0})
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert res.loc[0, "Severite"] == "Critique"


def test_commentaire_rentabilite_decrochage_marge_ok():
    row = pd.Series({"Vs N-1 (%)": -0.20, "Vs Bgt (%)": 0.03, "Taux de Marge N-1": 0.18, "Taux de Marge": 0.19})
    comment = build_rentabilite_comment(row, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert "pilotage marge défensif" in comment
    assert "objectif budgétaire atteint" in comment


def test_commentaire_rentabilite_donnees_manquantes():
    row = pd.Series({"Vs N-1 (%)": np.nan, "Vs Bgt (%)": np.nan, "Taux de Marge N-1": np.nan, "Taux de Marge": np.nan})
    comment = build_rentabilite_comment(row, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert "insuffisantes" in comment


def test_split_code_libelle_gere_nan():
    s = pd.Series(["010 - BOISSON", np.nan, "SansSeparateur"])
    res = split_code_libelle(s)
    assert pd.isna(res.loc[1, "Code"])
    assert res.loc[2, "Libelle"] == "SansSeparateur"


def test_detect_format_priorite_code_sur_motcle():
    global FORMAT_BY_CODE
    original = FORMAT_BY_CODE.copy()
    try:
        FORMAT_BY_CODE["999"] = "Hyper"
        assert detect_format("999", "Un libellé quelconque sans mot-clé") == "Hyper"
        assert detect_format("111", "10301 - Hyper Marcory") == "Hyper"
        assert detect_format("222", "10601 - Supeco Niangon") == "Supeco"
        assert detect_format("333", None) == "Autre"
    finally:
        FORMAT_BY_CODE.clear()
        FORMAT_BY_CODE.update(original)


def test_load_data_exclut_lignes_parasites():
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
    buf.name = "test.xlsx"
    result = _load_data_impl(buf)
    assert len(result["global"]) == 1
    assert len(result["rayon"]) == 1
    assert len(result["couple"]) == 1
    assert result["couple"].loc[0, "Format"] == "Hyper"


def test_read_any_export_extension_invalide():
    class FakeFile:
        name = "export.pdf"
    try:
        read_any_export(FakeFile())
        raise AssertionError("DataLoadError attendue pour une extension non supportée")
    except DataLoadError:
        pass


def test_missing_required_columns_leve_dataloaderror():
    df = pd.DataFrame({"Société": ["SOC"], "Rayon": ["Total"], "Site": [np.nan]})
    buf = io.BytesIO()
    df.to_excel(buf, sheet_name="Export", index=False)
    buf.seek(0)
    buf.name = "test.xlsx"
    try:
        _load_data_impl(buf)
        raise AssertionError("DataLoadError attendue si CA/CA N-1 absents")
    except DataLoadError:
        pass


def run_all_tests():
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
