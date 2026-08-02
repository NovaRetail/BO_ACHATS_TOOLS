# -*- coding: utf-8 -*-
"""
18_💸_Reporting_Vente CA.py — V4.1
============================================================
SmartBuyer Hub — Module Reporting Commercial (Point de situation & Pilotage)

V4.1 — refonte de l'écran d'accueil :
    - Accueil aligné sur la charte du module Reporting Ventes
      (titre pleine largeur + callout, colonne "Contenu du module",
       les 4 critères C1–C4 en boîtes colorées, bloc "Fonctionnement")
    - Reste identique à V4 :
        * CA à risque (FCFA) : priorisation Pareto des flops
        * Benchmark par pairs de format
        * Décomposition Trafic / Panier
        * Cartes KPI effet 3D + halo conditionnel
        * Onglets Vue d'ensemble / Par Rayon / Flops (maître-détail) / Méthodologie
    Export : un seul classeur Excel COPIL multi-onglets

Règles de flop (validées) :
    C1 - Décrochage CA vs N-1   : Vs N-1 (%)  <= seuil CA
    C2 - Écart vs Budget        : Vs Bgt (%)  <= seuil Budget (ignoré si Budget NaN)
    C3 - Dégradation marge      : Delta Marge (pts) <= seuil Marge
    C4 - Rupture / fermeture    : CA NaN/0 alors que CA N-1 > 0 (prioritaire)

Architecture :
    0-4  : config, palette, I/O, calculs, analytique, helpers (pur, testable)
    5    : main() -> rendu Streamlit
    6    : tests unitaires
    7    : point d'entrée
============================================================
"""

import io
import os
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# 0. CONFIGURATION PAGE & THEME — Apple charter
# ============================================================

st.set_page_config(
    page_title="Reporting Vente CA",
    page_icon="💸",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Bascule rendu cartes KPI : True = effet 3D + halo coloré, False = flat simple
KPI_STYLE_3D = True

COL_BG_PAGE = "#F2F2F7"
COL_BG_CARD = "#FFFFFF"
COL_BG_QUADRANT = "#F7F7F9"
COL_BORDER = "#E5E5EA"
COL_TEXT_PRIMARY = "#1C1C1E"
COL_TEXT_SECONDARY = "#8E8E93"

COL_BLUE = "#007AFF"
COL_RED = "#FF3B30"
COL_RED_BG = "#FFEBEA"
COL_ORANGE = "#FF9500"
COL_ORANGE_BG = "#FFF4E5"
COL_AMBER = "#C9A400"
COL_AMBER_BG = "#FFFBE5"
COL_GREEN = "#34C759"
COL_GREEN_BG = "#E9F9EE"
COL_PURPLE = "#AF52DE"
COL_PURPLE_BG = "#F6E9FC"

SEVERITY_STYLE = {
    "Critique": {"text": COL_RED, "bg": COL_RED_BG, "emoji": "🔴"},
    "Flop majeur": {"text": COL_ORANGE, "bg": COL_ORANGE_BG, "emoji": "🟠"},
    "Flop modéré": {"text": COL_AMBER, "bg": COL_AMBER_BG, "emoji": "🟡"},
    "OK": {"text": COL_GREEN, "bg": COL_GREEN_BG, "emoji": "🟢"},
}
SEVERITY_ORDER = ["Critique", "Flop majeur", "Flop modéré", "OK"]

THRESHOLD_PRESETS = {
    "Strict":   {"ca": -5,  "bgt": -5,  "marge": -0.5},
    "Standard": {"ca": -10, "bgt": -10, "marge": -0.8},
    "Souple":   {"ca": -15, "bgt": -15, "marge": -1.2},
}

RAYON_TO_BUYER = {
    "BOISSON": "CK",
    "EPICERIE": "GB",
    "DROGUERIE": "AC",
    "PARFUMERIE HYGIENE": "AC",
}


def get_buyer_code(rayon_libelle) -> str:
    if not isinstance(rayon_libelle, str):
        return "N/A"
    lib_up = rayon_libelle.upper()
    for key, code in RAYON_TO_BUYER.items():
        if key in lib_up:
            return code
    return "N/A"


def rayon_trigram(rayon_libelle) -> str:
    """Trigramme d'un rayon pour l'avatar (BOISSON -> BOI)."""
    if not isinstance(rayon_libelle, str) or not rayon_libelle.strip():
        return "?"
    return rayon_libelle.strip()[:3].upper()


CUSTOM_CSS = f"""
<style>
    .stApp {{ background-color: {COL_BG_PAGE}; }}
    section[data-testid="stSidebar"] {{
        background-color: {COL_BG_CARD};
        border-right: 0.5px solid {COL_BORDER};
    }}
    h1, h2, h3, h4 {{ color: {COL_TEXT_PRIMARY}; font-weight: 600; }}
    p, span, div, label {{ color: {COL_TEXT_PRIMARY}; }}
    .kpi-flat {{
        background-color: {COL_BG_CARD}; border-radius: 14px;
        padding: 0.9rem 1rem; margin-bottom: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    .kpi-label {{ font-size: 12px; color: {COL_TEXT_SECONDARY}; margin: 0 0 6px 0; font-weight: 500; }}
    .kpi-value {{ font-size: 24px; font-weight: 700; margin: 0; }}
    .badge-pill {{
        font-size: 11px; padding: 4px 11px; border-radius: 20px;
        display: inline-block; font-weight: 600;
    }}
    .card {{
        background-color: {COL_BG_CARD}; border-radius: 16px;
        padding: 1.1rem 1.25rem; margin-bottom: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    .flop-row {{
        background-color: {COL_BG_CARD}; border-radius: 14px;
        padding: 0.7rem 1rem; margin-bottom: 6px;
        display: flex; align-items: center; justify-content: space-between;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }}
    .criterion-line {{
        display: flex; justify-content: space-between; padding: 7px 0;
        border-bottom: 0.5px solid {COL_BORDER}; font-size: 13.5px;
    }}
    .criterion-line:last-child {{ border-bottom: none; }}
    .buyer-avatar {{
        width: 44px; height: 44px; border-radius: 50%;
        background: {COL_BLUE}1A; color: {COL_BLUE};
        display: flex; align-items: center; justify-content: center;
        font-weight: 700; font-size: 13px;
    }}
    div[data-testid="stMetric"] {{
        background-color: {COL_BG_CARD}; border-radius: 14px;
        padding: 0.75rem 1rem; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    .preset-caption {{ font-size: 11px; color: {COL_TEXT_SECONDARY}; margin-top: -6px; }}
    .method-block {{
        background-color: {COL_BG_CARD}; border-radius: 14px;
        padding: 1.1rem 1.4rem; margin-bottom: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    .method-block code {{
        background: {COL_BG_QUADRANT}; padding: 2px 6px; border-radius: 6px;
        font-size: 13px;
    }}
    .welcome-section-title {{
        font-size: 12px; font-weight: 700; color: {COL_TEXT_SECONDARY};
        letter-spacing: 0.06em; text-transform: uppercase; margin: 0 0 12px 0;
    }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

PLOTLY_TEMPLATE = go.layout.Template(
    layout=go.Layout(
        paper_bgcolor=COL_BG_CARD,
        plot_bgcolor=COL_BG_CARD,
        font=dict(color=COL_TEXT_PRIMARY, size=12),
        xaxis=dict(gridcolor=COL_BORDER, zerolinecolor=COL_BORDER),
        yaxis=dict(gridcolor=COL_BORDER, zerolinecolor=COL_BORDER),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
)

# ============================================================
# 1. LECTEUR I/O ROBUSTE (xlsx / csv, encodage auto)
# ============================================================

EXPECTED_COLS = [
    "Société", "Rayon", "Site", "CA N-1", "Budget", "CA", "Poids",
    "Vs N-1 (%)", "Vs Bgt (%)", "Marge N-1", "Marge",
    "Taux de Marge N-1", "Taux de Marge", "Taux de Marge N Vs N-1",
    "Débit N-1", "Débit", "Vs N-1 (%).1", "Panier N-1", "Panier",
    "Panier N Vs N-1", "Panier Qté N-1", "Panier Qté",
    "Panier Qté N Vs N-1", "Volume N-1", "Volume", "Volume N Vs N-1",
]

FORMAT_BY_CODE = {}
FORMAT_KEYWORDS = {"hyper": "Hyper", "market": "Market", "supeco": "Supeco"}


class DataLoadError(Exception):
    """Erreur métier levée par le lecteur I/O — message affichable tel quel."""
    pass


def read_any_export(file, sheet_name: str = "Export") -> pd.DataFrame:
    """Lecteur universel encoding-safe (xlsx ou csv)."""
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
            raise DataLoadError("Impossible de décoder le CSV (UTF-8, UTF-8-SIG, Latin-1 testés).")
        raise DataLoadError(f"Extension '.{ext}' non supportée — utilise .xlsx ou .csv.")
    except ValueError as e:
        if "Worksheet" in str(e) or "sheet" in str(e).lower():
            raise DataLoadError(f"Feuille '{sheet_name}' introuvable. Vérifie le nom de l'onglet Excel.")
        raise DataLoadError(f"Erreur de lecture du fichier : {e}")


def split_code_libelle(serie: pd.Series) -> pd.DataFrame:
    serie = serie.astype("string")
    split = serie.str.split(" - ", n=1, expand=True)
    if split.shape[1] == 1:
        split[1] = np.nan
    code = split[0].str.strip()
    libelle = split[1].str.strip().fillna(serie)
    return pd.DataFrame({"Code": code, "Libelle": libelle})


def detect_format(code, libelle) -> str:
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
    raw = read_any_export(file, sheet_name=sheet_name)

    missing_cols = [c for c in EXPECTED_COLS if c not in raw.columns]
    required_minimum = {"Société", "Rayon", "Site", "CA", "CA N-1"}
    if not required_minimum.issubset(set(raw.columns)):
        raise DataLoadError(
            f"Colonnes indispensables manquantes : {sorted(required_minimum - set(raw.columns))}."
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
        d["Acheteur"] = d["Rayon_Libelle"].apply(get_buyer_code)

    ss = split_code_libelle(df_couple["Site"])
    df_couple["Site_Code"] = ss["Code"].values
    df_couple["Site_Libelle"] = ss["Libelle"].values
    df_couple["Format"] = [detect_format(c, l) for c, l in zip(df_couple["Site_Code"], df_couple["Site_Libelle"])]

    if df_couple.empty:
        raise DataLoadError(
            "Aucune ligne de niveau Couple Magasin x Rayon détectée. Vérifie la structure de l'export."
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
# 2. MOTEUR DE CALCUL — FLOPS, SÉVÉRITÉ & ANALYTIQUE
# ============================================================

def compute_delta_marge_pts(taux_marge: pd.Series, taux_marge_n1: pd.Series) -> pd.Series:
    return (taux_marge - taux_marge_n1) * 100


def compute_ca_a_risque(ca: pd.Series, ca_n1: pd.Series) -> pd.Series:
    """CA perdu en valeur absolue (FCFA). Positif = perte. Lecture top-line secondaire."""
    return ca_n1.fillna(0) - ca.fillna(0)


def compute_marge_a_risque(marge: pd.Series, marge_n1: pd.Series) -> pd.Series:
    """Marge perdue en valeur absolue (FCFA). Positif = valeur détruite.
    Clé de priorisation retail : l'acheteur est jugé sur sa marge, pas sur le top-line."""
    return marge_n1.fillna(0) - marge.fillna(0)


def compute_benchmark_pairs(df: pd.DataFrame) -> pd.Series:
    """Écart du Vs N-1 (%) de chaque couple à la moyenne de ses pairs
    (même Format + même Rayon). Négatif = sous-performe ses pairs."""
    if "Format" not in df.columns or "Rayon_Libelle" not in df.columns:
        return pd.Series(np.nan, index=df.index)
    moyenne_pairs = df.groupby(["Format", "Rayon_Libelle"])["Vs N-1 (%)"].transform("mean")
    return df["Vs N-1 (%)"] - moyenne_pairs


def decompose_trafic_panier(row) -> dict:
    """Décomposition symétrique de la variation de CA en Effet Trafic + Effet Panier.
        CA = Débit x Panier moyen
        Effet Trafic = ΔDébit x (Panier_N + Panier_N-1)/2
        Effet Panier = ΔPanier x (Débit_N + Débit_N-1)/2
    Retourne None si données nécessaires absentes."""
    debit_n = row.get("Débit")
    debit_n1 = row.get("Débit N-1")
    panier_n = row.get("Panier")
    panier_n1 = row.get("Panier N-1")

    if any(pd.isna(v) for v in (debit_n, debit_n1, panier_n, panier_n1)):
        return {"effet_trafic": None, "effet_panier": None, "delta_ca": None}

    effet_trafic = (debit_n - debit_n1) * (panier_n + panier_n1) / 2
    effet_panier = (panier_n - panier_n1) * (debit_n + debit_n1) / 2
    return {
        "effet_trafic": effet_trafic,
        "effet_panier": effet_panier,
        "delta_ca": effet_trafic + effet_panier,
    }


def compute_flops(df: pd.DataFrame, seuil_ca: float, seuil_bgt: float, seuil_marge: float) -> pd.DataFrame:
    out = df.copy()
    out["Delta_Marge_pt"] = compute_delta_marge_pts(out["Taux de Marge"], out["Taux de Marge N-1"])
    out["CA_a_risque"] = compute_ca_a_risque(out["CA"], out["CA N-1"])
    if "Marge" in out.columns and "Marge N-1" in out.columns:
        out["Marge_a_risque"] = compute_marge_a_risque(out["Marge"], out["Marge N-1"])
    else:
        out["Marge_a_risque"] = np.nan
    out["Ecart_vs_pairs"] = compute_benchmark_pairs(out)

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
# 4. HELPERS D'AFFICHAGE + ANALYTIQUE
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


def variation_color(x) -> str:
    if pd.isna(x):
        return COL_TEXT_SECONDARY
    return COL_GREEN if x >= 0 else COL_RED


def _hex_to_rgb(hex_color: str) -> tuple:
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def kpi_card(label: str, value: str, accent: str = COL_BLUE, value_color: str = None) -> str:
    """Carte KPI. accent = couleur barre/halo. value_color = couleur du chiffre.
    Rendu 3D (halo + relief) ou flat selon le flag KPI_STYLE_3D."""
    value_color = value_color or COL_TEXT_PRIMARY
    if not KPI_STYLE_3D:
        return f"""<div class="kpi-flat"><p class="kpi-label">{label}</p>
        <p class="kpi-value" style="color:{value_color}">{value}</p></div>"""

    r, g, b = _hex_to_rgb(accent)
    tint = f"rgba({r},{g},{b},0.045)"
    halo = f"rgba({r},{g},{b},0.15)"
    return f"""<div style="background:linear-gradient(160deg,#FFFFFF 0%,{tint} 100%);
        border-radius:16px;padding:0.9rem 1rem;margin-bottom:8px;
        box-shadow:0 1px 2px rgba(0,0,0,0.04),0 6px 16px {halo};
        border-top:2.5px solid {accent}">
        <p class="kpi-label">{label}</p>
        <p class="kpi-value" style="color:{value_color}">{value}</p></div>"""


def kpi_card_conditional(label: str, value: str, metric_value) -> str:
    """Carte KPI dont le halo/couleur dépend du signe de la métrique (Vs N-1, Vs Budget)."""
    color = variation_color(metric_value)
    return kpi_card(label, value, accent=color, value_color=color)


def severity_badge(severite: str) -> str:
    s = SEVERITY_STYLE.get(severite, SEVERITY_STYLE["OK"])
    return f'<span class="badge-pill" style="background:{s["bg"]};color:{s["text"]}">{s["emoji"]} {severite}</span>'


def dot_html(color: str) -> str:
    return f'<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:{color};margin-right:6px"></span>'


def dot_color_for(x, seuil: float) -> str:
    if pd.isna(x):
        return COL_TEXT_SECONDARY
    if x >= 0:
        return COL_GREEN
    if x <= seuil:
        return COL_RED
    return COL_ORANGE


def quadrant_html(label: str, rows_html: str, muted: bool = False) -> str:
    opacity = "0.55" if muted else "1"
    return f"""<div style="background:{COL_BG_QUADRANT};border-radius:10px;padding:0.7rem 0.85rem;opacity:{opacity}">
        <p style="font-size:11px;color:{COL_TEXT_SECONDARY};margin:0 0 8px 0;font-weight:600;letter-spacing:0.02em;text-transform:uppercase">{label}</p>
        {rows_html}
    </div>"""


def quadrant_row(label: str, value: str, dot: str = "") -> str:
    return f"""<div style="display:flex;justify-content:space-between;font-size:13px;padding:2px 0">
        <span>{label}</span><span>{dot}{value}</span></div>"""


def build_steering_wheel_card(rayon_row: pd.Series, nb_flops_actifs: int, nb_magasins: int,
                              seuil_ca: float, seuil_marge: float) -> str:
    """Carte 4 cadrans (Client / Finance / Activité / Opérations placeholder) par rayon."""
    debit_var = rayon_row.get("Vs N-1 (%).1")
    panier_var = rayon_row.get("Panier N Vs N-1")
    ca_var = rayon_row.get("Vs N-1 (%)")
    delta_marge = compute_delta_marge_pts(
        pd.Series([rayon_row.get("Taux de Marge")]), pd.Series([rayon_row.get("Taux de Marge N-1")])
    ).iloc[0]

    client_q = quadrant_html("Client", (
        quadrant_row("Débit (trafic)", fmt_pct(debit_var), dot_html(dot_color_for(debit_var, seuil_ca)))
        + quadrant_row("Panier moyen", fmt_pct(panier_var), dot_html(dot_color_for(panier_var, seuil_ca)))
    ))
    finance_q = quadrant_html("Finance", (
        quadrant_row("CA vs N-1", fmt_pct(ca_var), dot_html(dot_color_for(ca_var, seuil_ca)))
        + quadrant_row("Δ marge", fmt_pt(delta_marge), dot_html(dot_color_for(delta_marge, seuil_marge)))
    ))
    activite_q = quadrant_html("Activité", (
        quadrant_row("Points actifs", str(nb_flops_actifs))
        + quadrant_row("Magasins", str(nb_magasins))
    ))
    operations_q = quadrant_html("Opérations", (
        quadrant_row("Disponibilité", "à connecter")
    ), muted=True)

    return f"""<div class="card">
        <p style="font-weight:600;font-size:15px;margin:0 0 10px 0">{rayon_row.get('Rayon_Libelle', '')}</p>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
            {client_q}{finance_q}{activite_q}{operations_q}
        </div>
    </div>"""


def get_selected_rows(event) -> list:
    """Extrait les index sélectionnés depuis st.dataframe(..., on_select="rerun").
    Robuste aux deux formats renvoyés selon la version de Streamlit."""
    if event is None:
        return []
    selection = getattr(event, "selection", None)
    if selection is None and isinstance(event, dict):
        selection = event.get("selection")
    if selection is None:
        return []
    if isinstance(selection, dict):
        return selection.get("rows", []) or []
    return getattr(selection, "rows", []) or []


def top_magasins(flops_rayon: pd.DataFrame, by: str, n: int = 3) -> pd.DataFrame:
    """Top n magasins d'un rayon. by='ca' -> plus gros CA ; by='progression' -> meilleur Vs N-1."""
    if flops_rayon.empty:
        return flops_rayon
    if by == "ca":
        return flops_rayon.sort_values("CA", ascending=False).head(n)
    return flops_rayon.dropna(subset=["Vs N-1 (%)"]).sort_values("Vs N-1 (%)", ascending=False).head(n)


def build_rayon_brief(rayon_row: pd.Series, flops_rayon: pd.DataFrame,
                      seuil_ca: float, seuil_bgt: float, seuil_marge: float) -> str:
    """Brief texte prêt à partager, liste verticale, incluant les tops magasins."""
    nom = rayon_row.get("Rayon_Libelle", "")
    date = pd.Timestamp.today().strftime("%d/%m/%Y")
    delta_marge = compute_delta_marge_pts(
        pd.Series([rayon_row.get("Taux de Marge")]), pd.Series([rayon_row.get("Taux de Marge N-1")])
    ).iloc[0]
    comment = build_rentabilite_comment(rayon_row, seuil_ca, seuil_bgt, seuil_marge)

    lines = [f"POINT RAYON {nom} — {date}", "", "SYNTHÈSE"]
    lines.append(f"  CA      : {fmt_fcfa(rayon_row.get('CA'))} ({fmt_pct(rayon_row.get('Vs N-1 (%)'))} vs N-1)")
    taux = fmt_pct(rayon_row.get("Taux de Marge")).replace("+", "")
    lines.append(f"  Marge   : {fmt_fcfa(rayon_row.get('Marge'))} ({taux} · {fmt_pt(delta_marge)})")
    lines.append(f"  Qté     : {fmt_fcfa(rayon_row.get('Panier Qté'))} ({fmt_pct(rayon_row.get('Panier Qté N Vs N-1'))})")
    lines.append(f"  Débit   : {fmt_pct(rayon_row.get('Vs N-1 (%).1'))}")
    lines.append(f"  Panier  : {fmt_pct(rayon_row.get('Panier N Vs N-1'))}")
    lines.append(f"  {comment}")
    lines.append("")

    top_ca = top_magasins(flops_rayon, "ca", 3)
    top_prog = top_magasins(flops_rayon, "progression", 3)
    lines.append("TOP MAGASINS")
    if not top_ca.empty:
        ca_str = ", ".join(f"{r['Site_Libelle']} ({fmt_fcfa(r['CA'])})" for _, r in top_ca.iterrows())
        lines.append(f"  CA          : {ca_str}")
    if not top_prog.empty:
        prog_str = ", ".join(f"{r['Site_Libelle']} ({fmt_pct(r['Vs N-1 (%)'])})" for _, r in top_prog.iterrows())
        lines.append(f"  Progression : {prog_str}")
    lines.append("")

    a_traiter = flops_rayon[flops_rayon["Severite"] != "OK"].sort_values("Marge_a_risque", ascending=False)
    lines.append(f"POINTS D'ATTENTION ({len(a_traiter)})")
    if a_traiter.empty:
        lines.append("  Aucun point d'alerte à date.")
    else:
        for _, r in a_traiter.head(10).iterrows():
            if r["C4_Rupture"]:
                detail = "rupture totale de CA"
            else:
                detail = f"CA {fmt_pct(r['Vs N-1 (%)'])}, Δ marge {fmt_pt(r['Delta_Marge_pt'])}"
            lines.append(f"  - [{r['Severite']}] {r['Site_Libelle']} : {detail} (marge à risque {fmt_fcfa(r['Marge_a_risque'])})")

    return "\n".join(lines)


def build_copil_export(df_global: pd.DataFrame, df_rayon: pd.DataFrame, df_flops: pd.DataFrame,
                        seuil_ca: float, seuil_bgt: float, seuil_marge: float) -> bytes:
    """Un seul classeur Excel COPIL : KPI Global, Par Rayon, Flops, Données brutes."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if not df_global.empty:
            g = df_global.iloc[0]
            pd.DataFrame({
                "Indicateur": ["CA total", "Vs N-1 (%)", "Vs Budget (%)", "Marge totale", "Taux de marge"],
                "Valeur": [
                    fmt_fcfa(g.get("CA")), fmt_pct(g.get("Vs N-1 (%)")), fmt_pct(g.get("Vs Bgt (%)")),
                    fmt_fcfa(g.get("Marge")), fmt_pct(g.get("Taux de Marge")).replace("+", ""),
                ],
            }).to_excel(writer, index=False, sheet_name="KPI Global")

        rayon_export = df_rayon.copy()
        if not rayon_export.empty:
            rayon_export["Commentaire rentabilité"] = rayon_export.apply(
                lambda r: build_rentabilite_comment(r, seuil_ca, seuil_bgt, seuil_marge), axis=1
            )
            cols = ["Rayon_Libelle", "CA", "CA N-1", "Vs N-1 (%)", "Budget", "Vs Bgt (%)",
                    "Marge", "Taux de Marge", "Poids", "Commentaire rentabilité"]
            rayon_export[[c for c in cols if c in rayon_export.columns]].to_excel(
                writer, index=False, sheet_name="Par Rayon"
            )

        flops_export = df_flops.sort_values("Marge_a_risque", ascending=False)
        cols = ["Site_Libelle", "Rayon_Libelle", "Format", "CA", "CA N-1", "Vs N-1 (%)",
                "Budget", "Vs Bgt (%)", "Marge", "Marge N-1", "Delta_Marge_pt",
                "Marge_a_risque", "CA_a_risque", "Ecart_vs_pairs",
                "Severite", "Score_Label"]
        flops_export[[c for c in cols if c in flops_export.columns]].to_excel(
            writer, index=False, sheet_name="Flops"
        )

        df_flops.to_excel(writer, index=False, sheet_name="Données brutes")
    return buffer.getvalue()


def _welcome_criterion(tag: str, titre: str, formule: str, color: str) -> str:
    """Boîte critère de flop pour l'accueil (style objectifs colorés du module Reporting Ventes)."""
    r, g, b = _hex_to_rgb(color)
    tint = f"rgba({r},{g},{b},0.07)"
    border = f"rgba({r},{g},{b},0.20)"
    return f"""<div style="background:{tint};border:1px solid {border};border-radius:14px;
        padding:0.85rem 1rem;margin-bottom:9px">
        <span style="background:{color};color:#FFFFFF;font-size:12px;font-weight:700;
            padding:4px 10px;border-radius:8px;display:inline-block;margin-bottom:7px">{tag}</span>
        <p style="font-size:13.5px;font-weight:600;color:{COL_TEXT_PRIMARY};margin:0 0 3px 0">{titre}</p>
        <p style="font-size:12px;color:{COL_TEXT_SECONDARY};margin:0;
            font-family:ui-monospace,Menlo,monospace">{formule}</p>
    </div>"""


def render_welcome():
    """Écran d'accueil (charte type module Reporting Ventes) : présentation + 4 critères de flop."""
    blue_r, blue_g, blue_b = _hex_to_rgb(COL_BLUE)
    callout_bg = f"rgba({blue_r},{blue_g},{blue_b},0.07)"

    # --- Titre + sous-titre (niveau page) ---
    st.markdown(
        f"""<div style="display:flex;align-items:center;gap:12px;margin-bottom:2px">
            <span style="font-size:34px">💸</span>
            <span style="font-size:30px;font-weight:700;color:{COL_TEXT_PRIMARY}">Reporting Vente CA</span>
        </div>
        <p style="font-size:15px;color:{COL_TEXT_SECONDARY};margin:0 0 18px 0">
        Point de situation commercial et orientation des acheteurs, du global au couple Magasin × Rayon.</p>""",
        unsafe_allow_html=True,
    )

    # --- Callout "À quoi sert ce module ?" ---
    st.markdown(
        f"""<div style="background:{callout_bg};border-left:4px solid {COL_BLUE};
            border-radius:12px;padding:1rem 1.25rem;margin-bottom:22px">
            <p style="font-weight:600;font-size:15px;margin:0 0 6px 0">ℹ️ À quoi sert ce module ?</p>
            <p style="font-size:14px;line-height:1.6;margin:0;color:{COL_TEXT_PRIMARY}">
            Il transforme un export ventes (Global → Rayon → couple Magasin × Rayon) en point de situation
            commercial : température du réseau, détection des <b>Flops</b> selon 4 critères, diagnostic par rayon
            (trafic, panier, marge) et synthèse COPIL. Un seul fichier à charger dans la barre latérale.</p>
        </div>""",
        unsafe_allow_html=True,
    )

    left, right = st.columns([1.15, 1], gap="large")

    # --- Colonne gauche : contenu du module ---
    with left:
        st.markdown('<p class="welcome-section-title">Contenu du module</p>', unsafe_allow_html=True)
        st.markdown(
            f"""<div class="card">
                <p style="font-weight:600;font-size:15.5px;margin:0 0 8px 0">🎯 Vue d'ensemble & pilotage</p>
                <p style="font-size:13.5px;line-height:1.6;color:{COL_TEXT_PRIMARY};margin:0">
                Température du réseau en un coup d'œil : KPI clés, Top 5 des points d'attention priorisés par
                <b>marge à risque</b> (valeur détruite), point de situation par rayon et magasins les plus en difficulté.</p>
            </div>""",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"""<div class="card">
                <p style="font-weight:600;font-size:15.5px;margin:0 0 8px 0">🚩 Détection & diagnostic des Flops</p>
                <p style="font-size:13.5px;line-height:1.6;color:{COL_TEXT_PRIMARY};margin:0">
                Couples Magasin × Rayon en sous-performance selon 4 critères, décomposition
                <b>Trafic / Panier</b>, écart vs pairs de format, brief prêt à partager et export COPIL multi-onglets.</p>
            </div>""",
            unsafe_allow_html=True,
        )

    # --- Colonne droite : les 4 critères de flop ---
    with right:
        st.markdown('<p class="welcome-section-title">Les 4 critères de flop</p>', unsafe_allow_html=True)
        st.markdown(_welcome_criterion(
            "C1 · Décrochage CA", "CA en repli vs N-1",
            "Vs N-1 (%) ≤ seuil CA", COL_BLUE), unsafe_allow_html=True)
        st.markdown(_welcome_criterion(
            "C2 · Écart Budget", "Objectif budgétaire manqué",
            "Vs Bgt (%) ≤ seuil Bgt · ignoré si pas de budget", COL_ORANGE), unsafe_allow_html=True)
        st.markdown(_welcome_criterion(
            "C3 · Dégradation marge", "Taux de marge en recul",
            "(Taux N − Taux N-1) × 100 ≤ seuil marge (pts)", COL_AMBER), unsafe_allow_html=True)
        st.markdown(_welcome_criterion(
            "C4 · Rupture / fermeture", "CA nul alors que N-1 > 0 · prioritaire",
            "CA absent/0 & CA N-1 > 0 → Critique", COL_RED), unsafe_allow_html=True)

    # --- Fonctionnement ---
    st.markdown(
        f"""<div style="background:{COL_GREEN_BG};border-left:4px solid {COL_GREEN};
            border-radius:12px;padding:1rem 1.25rem;margin-top:8px">
            <p style="font-weight:700;font-size:12px;color:{COL_TEXT_SECONDARY};
            letter-spacing:0.06em;text-transform:uppercase;margin:0 0 10px 0">Fonctionnement</p>
            <p style="font-size:14px;line-height:1.95;margin:0;color:{COL_TEXT_PRIMARY}">
            <b>1.</b> Charge ton export ventes (.xlsx ou .csv) dans la barre latérale.<br>
            <b>2.</b> Choisis un preset de seuils (Strict / Standard / Souple) puis ajuste les curseurs.<br>
            <b>3.</b> Explore les onglets : Vue d'ensemble → Par Rayon → Flops → Méthodologie.<br>
            <b>4.</b> Télécharge la synthèse COPIL prête à diffuser.</p>
        </div>""",
        unsafe_allow_html=True,
    )


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
        render_welcome()
        st.stop()

    try:
        data = load_data(uploaded_file)
    except DataLoadError as e:
        st.error(str(e))
        st.stop()

    df_global_raw, df_rayon_raw, df_couple_raw = data["global"], data["rayon"], data["couple"]

    with st.sidebar:
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

    # --- Header + export unique ---
    top_l, top_r = st.columns([4, 1])
    with top_l:
        st.markdown(f"## Point de situation — {societe_sel}")
        st.caption("Vue d'ensemble · Par rayon · Détail des flops · Méthodologie")
    with top_r:
        st.markdown("<div style='padding-top:14px'></div>", unsafe_allow_html=True)
        copil_bytes = build_copil_export(df_global_f, df_rayon_f, df_flops, seuil_ca, seuil_bgt, seuil_marge)
        st.download_button(
            "📊 Export COPIL", data=copil_bytes, file_name=f"synthese_copil_{societe_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
        )

    # --- KPI globaux (3D + halo conditionnel) ---
    if not df_global_f.empty:
        g = df_global_f.iloc[0]
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.markdown(kpi_card("CA total", fmt_fcfa(g["CA"]), accent=COL_BLUE, value_color=COL_BLUE), unsafe_allow_html=True)
        c2.markdown(kpi_card_conditional("Vs N-1", fmt_pct(g["Vs N-1 (%)"]), g["Vs N-1 (%)"]), unsafe_allow_html=True)
        c3.markdown(kpi_card_conditional("Vs Budget", fmt_pct(g["Vs Bgt (%)"]), g["Vs Bgt (%)"]), unsafe_allow_html=True)
        c4.markdown(kpi_card("Marge totale", fmt_fcfa(g["Marge"]), accent=COL_PURPLE, value_color=COL_PURPLE), unsafe_allow_html=True)
        c5.markdown(kpi_card("Taux de marge", fmt_pct(g["Taux de Marge"]).replace("+", ""), accent=COL_BLUE), unsafe_allow_html=True)
    else:
        st.info("Pas de ligne Global disponible pour cette société.")

    nb_critique = int((df_flops["Severite"] == "Critique").sum())
    if nb_critique > 0:
        st.markdown(
            f"""<div style="background:{COL_RED_BG};border-radius:14px;
            padding:0.75rem 1rem;margin:12px 0;color:{COL_RED};font-weight:500">
            🔴 {nb_critique} rupture(s) totale de CA détectée(s) — voir onglet Flops</div>""",
            unsafe_allow_html=True,
        )

    st.markdown("---")
    tab_exec, tab_rayon, tab_flops, tab_method = st.tabs(
        ["🎯 Vue d'ensemble", "🏷️ Par Rayon", "🚩 Flops", "📖 Méthodologie"]
    )

    # ---------------- TAB VUE D'ENSEMBLE ----------------
    with tab_exec:
        st.markdown("#### Top 5 points d'attention")
        top5 = df_flops[df_flops["Severite"] != "OK"].sort_values("Marge_a_risque", ascending=False).head(5)
        if top5.empty:
            st.success("Aucun flop détecté sur ce périmètre.")
        else:
            for _, r in top5.iterrows():
                st.markdown(
                    f"""<div class="flop-row">
                        <div style="display:flex;align-items:center;gap:10px">{severity_badge(r['Severite'])}
                        <span>{r['Site_Libelle']} — {r['Rayon_Libelle']}</span></div>
                        <span style="color:{COL_TEXT_SECONDARY}">Vs N-1 {fmt_pct(r['Vs N-1 (%)'])} · marge à risque {fmt_fcfa(r['Marge_a_risque'])}</span>
                    </div>""",
                    unsafe_allow_html=True,
                )
            total_risque = df_flops[df_flops["Severite"] != "OK"]["Marge_a_risque"].clip(lower=0).sum()
            top5_risque = top5["Marge_a_risque"].clip(lower=0).sum()
            part = (top5_risque / total_risque * 100) if total_risque > 0 else 0
            st.caption(f"Top 5 = {fmt_fcfa(top5_risque)} FCFA de marge à risque, soit {part:.0f}% de la valeur détruite sur le périmètre.")

        st.markdown("#### Point de situation par rayon")
        rayon_cols = st.columns(min(2, max(1, len(df_rayon_f))))
        for i, (_, r) in enumerate(df_rayon_f.iterrows()):
            sub = df_flops[df_flops["Rayon_Libelle"] == r["Rayon_Libelle"]]
            nb_flops_rayon = int((sub["Severite"] != "OK").sum())
            nb_magasins = int(sub["Site_Libelle"].nunique())
            with rayon_cols[i % len(rayon_cols)]:
                st.markdown(build_steering_wheel_card(r, nb_flops_rayon, nb_magasins, seuil_ca, seuil_marge), unsafe_allow_html=True)

        st.markdown("#### Magasins les plus en difficulté")
        agg_magasin = (
            df_flops.groupby(["Site_Libelle", "Format"], as_index=False)
            .agg(CA=("CA", "sum"), CA_N1=("CA N-1", "sum"),
                 CA_a_risque=("CA_a_risque", "sum"), Marge_a_risque=("Marge_a_risque", "sum"),
                 Nb_Flops=("Severite", lambda s: (s != "OK").sum()),
                 Nb_Critiques=("Severite", lambda s: (s == "Critique").sum()))
            .sort_values(["Nb_Critiques", "Marge_a_risque"], ascending=False)
            .head(8)
        )
        agg_magasin["Vs N-1 (%)"] = np.where(
            agg_magasin["CA_N1"] > 0, (agg_magasin["CA"] - agg_magasin["CA_N1"]) / agg_magasin["CA_N1"], np.nan
        )
        display_mag = agg_magasin.copy()
        display_mag["CA"] = display_mag["CA"].apply(fmt_fcfa)
        display_mag["Marge à risque"] = display_mag["Marge_a_risque"].apply(fmt_fcfa)
        display_mag["CA à risque"] = display_mag["CA_a_risque"].apply(fmt_fcfa)
        display_mag["Vs N-1 (%)"] = display_mag["Vs N-1 (%)"].apply(fmt_pct)
        st.dataframe(
            display_mag[["Site_Libelle", "Format", "CA", "Vs N-1 (%)", "Marge à risque", "CA à risque", "Nb_Flops", "Nb_Critiques"]],
            use_container_width=True, hide_index=True,
        )

    # ---------------- TAB PAR RAYON ----------------
    with tab_rayon:
        rayons_present = sorted(df_rayon_f["Rayon_Libelle"].dropna().unique().tolist())
        if not rayons_present:
            st.info("Aucun rayon sur ce périmètre.")
        else:
            rayon_sel = st.radio("Rayon", options=rayons_present, horizontal=True)
            rayon_row = df_rayon_f[df_rayon_f["Rayon_Libelle"] == rayon_sel].iloc[0]
            flops_rayon = df_flops[df_flops["Rayon_Libelle"] == rayon_sel]

            nb_crit = int((flops_rayon["Severite"] == "Critique").sum())
            nb_maj = int((flops_rayon["Severite"] == "Flop majeur").sum())
            nb_mod = int((flops_rayon["Severite"] == "Flop modéré").sum())
            nb_magasins = int(flops_rayon["Site_Libelle"].nunique())

            st.markdown(
                f"""<div style="display:flex;align-items:center;gap:12px;margin-bottom:14px">
                    <div class="buyer-avatar">{rayon_trigram(rayon_sel)}</div>
                    <div><p style="font-weight:600;font-size:16px;margin:0">{rayon_sel}</p>
                    <p style="font-size:13px;color:{COL_TEXT_SECONDARY};margin:0">{nb_magasins} magasins · CA {fmt_fcfa(rayon_row.get('CA'))}</p></div>
                </div>""",
                unsafe_allow_html=True,
            )

            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.markdown(kpi_card("CA rayon", fmt_fcfa(rayon_row.get("CA")), accent=COL_BLUE, value_color=COL_BLUE), unsafe_allow_html=True)
            m2.markdown(kpi_card("Marge", fmt_fcfa(rayon_row.get("Marge")), accent=COL_PURPLE, value_color=COL_PURPLE), unsafe_allow_html=True)
            poids = rayon_row.get("Poids")
            m3.markdown(kpi_card("Poids CA", fmt_pct(poids).replace("+", "") if pd.notna(poids) else "n/a", accent=COL_BLUE), unsafe_allow_html=True)
            m4.markdown(kpi_card("Critiques", str(nb_crit), accent=COL_RED if nb_crit else COL_GREEN, value_color=COL_RED if nb_crit else COL_GREEN), unsafe_allow_html=True)
            m5.markdown(kpi_card("Flops majeurs", str(nb_maj), accent=COL_ORANGE if nb_maj else COL_GREEN, value_color=COL_ORANGE if nb_maj else COL_GREEN), unsafe_allow_html=True)
            m6.markdown(kpi_card("Flops modérés", str(nb_mod), accent=COL_AMBER if nb_mod else COL_GREEN, value_color=COL_AMBER if nb_mod else COL_GREEN), unsafe_allow_html=True)

            tcol1, tcol2 = st.columns(2)
            with tcol1:
                st.markdown("##### 🏆 Top CA")
                for _, r in top_magasins(flops_rayon, "ca", 3).iterrows():
                    st.markdown(
                        f"""<div class="flop-row"><span>{r['Site_Libelle']}</span>
                        <span style="color:{COL_BLUE};font-weight:600">{fmt_fcfa(r['CA'])}</span></div>""",
                        unsafe_allow_html=True,
                    )
            with tcol2:
                st.markdown("##### 📈 Top progression")
                for _, r in top_magasins(flops_rayon, "progression", 3).iterrows():
                    st.markdown(
                        f"""<div class="flop-row"><span>{r['Site_Libelle']}</span>
                        <span style="color:{variation_color(r['Vs N-1 (%)'])};font-weight:600">{fmt_pct(r['Vs N-1 (%)'])}</span></div>""",
                        unsafe_allow_html=True,
                    )

            st.markdown("##### Magasins à traiter en priorité")
            a_traiter = flops_rayon[flops_rayon["Severite"] != "OK"].sort_values("CA_a_risque", ascending=False)
            if a_traiter.empty:
                st.success("Aucun point d'alerte sur ce rayon.")
            else:
                for _, r in a_traiter.iterrows():
                    st.markdown(
                        f"""<div class="flop-row">
                            <div style="display:flex;align-items:center;gap:10px">{severity_badge(r['Severite'])}
                            <span>{r['Site_Libelle']}</span></div>
                            <span style="color:{COL_TEXT_SECONDARY}">Vs N-1 {fmt_pct(r['Vs N-1 (%)'])} · risque {fmt_fcfa(r['CA_a_risque'])}</span>
                        </div>""",
                        unsafe_allow_html=True,
                    )

            st.markdown("##### Brief prêt à partager")
            brief = build_rayon_brief(rayon_row, flops_rayon, seuil_ca, seuil_bgt, seuil_marge)
            st.text_area("Copier ce texte pour le point rayon", value=brief, height=320)

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
        df_view = df_view.sort_values("Marge_a_risque", ascending=False).reset_index(drop=True)

        st.caption(f"{len(df_view)} couple(s) affiché(s) — trié par marge à risque décroissante. Clique une ligne pour le détail.")

        table_col, detail_col = st.columns([2.2, 1])
        with table_col:
            display_df = pd.DataFrame({
                "Sévérité": df_view["Severite"].apply(lambda s: f"{SEVERITY_STYLE[s]['emoji']} {s}"),
                "Magasin": df_view["Site_Libelle"],
                "Rayon": df_view["Rayon_Libelle"],
                "CA": df_view["CA"].apply(fmt_fcfa),
                "Vs N-1": df_view["Vs N-1 (%)"].apply(fmt_pct),
                "Vs Budget": df_view["Vs Bgt (%)"].apply(fmt_pct),
                "Δ Marge": df_view["Delta_Marge_pt"].apply(fmt_pt),
                "Marge à risque": df_view["Marge_a_risque"].apply(fmt_fcfa),
                "CA à risque": df_view["CA_a_risque"].apply(fmt_fcfa),
                "Score": df_view["Score_Label"],
            })
            event = st.dataframe(
                display_df, use_container_width=True, hide_index=True, height=460,
                on_select="rerun", selection_mode="single-row",
            )

        with detail_col:
            selected_rows = get_selected_rows(event)
            if not selected_rows:
                st.markdown(
                    f"""<div class="card"><p style="color:{COL_TEXT_SECONDARY};margin:0">
                    Sélectionne une ligne pour afficher le détail des critères, la décomposition Trafic/Panier et l'écart vs pairs.</p></div>""",
                    unsafe_allow_html=True,
                )
            else:
                r = df_view.iloc[selected_rows[0]]
                c4_txt = "❌" if r["C4_Rupture"] else "✅"
                c1_txt = "❌" if r["C1_Decrochage_CA"] else "✅"
                c2_txt = "n/a" if not r["C2_Applicable"] else ("❌" if r["C2_Ecart_Budget"] else "✅")
                c3_txt = "n/a" if not r["C3_Applicable"] else ("❌" if r["C3_Degradation_Marge"] else "✅")

                decomp = decompose_trafic_panier(r)
                if decomp["effet_trafic"] is None:
                    decomp_html = f'<p style="font-size:13px;color:{COL_TEXT_SECONDARY};margin:8px 0 0 0">Décomposition Trafic/Panier : données insuffisantes.</p>'
                else:
                    decomp_html = f"""<div style="margin-top:10px">
                        <p style="font-size:12px;color:{COL_TEXT_SECONDARY};margin:0 0 4px 0;font-weight:600;text-transform:uppercase">Décomposition ΔCA</p>
                        <div class="criterion-line"><span>Effet Trafic</span><span>{fmt_fcfa(decomp['effet_trafic'])}</span></div>
                        <div class="criterion-line"><span>Effet Panier</span><span>{fmt_fcfa(decomp['effet_panier'])}</span></div>
                    </div>"""

                ecart = r.get("Ecart_vs_pairs")
                if pd.isna(ecart):
                    pairs_html = ""
                else:
                    interpret = "spécifique au magasin" if ecart < -0.03 else ("aligné sur ses pairs" if abs(ecart) <= 0.03 else "meilleur que ses pairs")
                    pairs_html = f"""<div class="criterion-line"><span>Écart vs pairs {r['Format']}</span>
                        <span style="color:{variation_color(ecart)}">{fmt_pct(ecart)} ({interpret})</span></div>"""

                st.markdown(
                    f"""<div class="card">
                        <p style="font-weight:600;font-size:15px;margin:0 0 2px 0">{r['Site_Libelle']}</p>
                        <p style="font-size:13px;color:{COL_TEXT_SECONDARY};margin:0 0 12px 0">{r['Rayon_Libelle']} · {r['Format']}</p>
                        {severity_badge(r['Severite'])}
                        <div style="margin-top:14px">
                            <div class="criterion-line"><span>{c4_txt} C4 — Rupture / fermeture</span></div>
                            <div class="criterion-line"><span>{c1_txt} C1 — Décrochage CA vs N-1</span><span>{fmt_pct(r['Vs N-1 (%)'])}</span></div>
                            <div class="criterion-line"><span>{c2_txt} C2 — Écart vs Budget</span><span>{fmt_pct(r['Vs Bgt (%)']) if r['C2_Applicable'] else '—'}</span></div>
                            <div class="criterion-line"><span>{c3_txt} C3 — Dégradation marge</span><span>{fmt_pt(r['Delta_Marge_pt']) if r['C3_Applicable'] else '—'}</span></div>
                            {pairs_html}
                            <div class="criterion-line"><span>Marge à risque</span><span style="font-weight:600;color:{COL_RED if pd.notna(r['Marge_a_risque']) and r['Marge_a_risque'] > 0 else COL_TEXT_PRIMARY}">{fmt_fcfa(r['Marge_a_risque'])}</span></div>
                            <div class="criterion-line"><span>CA à risque</span><span>{fmt_fcfa(r['CA_a_risque'])}</span></div>
                        </div>
                        {decomp_html}
                    </div>""",
                    unsafe_allow_html=True,
                )

    # ---------------- TAB MÉTHODOLOGIE ----------------
    with tab_method:
        render_methodology()


def render_methodology():
    """Onglet Méthodologie — documente la logique, les formules et les concepts."""
    blocks = [
        ("1. Périmètre & niveaux de données",
         "L'export est lu à trois niveaux imbriqués. <b>Global</b> (Société) : ligne où le Rayon vaut "
         "« Total » et le Site est vide. <b>Rayon</b> (toutes enseignes) : Rayon renseigné, Site = « Total ». "
         "<b>Couple Magasin × Rayon</b> : Rayon et Site renseignés — c'est le niveau où sont détectés les flops. "
         "Sont exclues au chargement : la ligne grand total Société, les lignes vides et la ligne de bas de page "
         "« Filtres appliqués »."),
        ("2. Les 4 critères de Flop",
         "<b>C1 — Décrochage CA</b> : <code>Vs N-1 (%) ≤ seuil CA</code>. "
         "<b>C2 — Écart vs Budget</b> : <code>Vs Bgt (%) ≤ seuil Budget</code>, <b>ignoré si aucun budget n'est alloué</b> "
         "(fréquent sur les Supeco) — on ne pénalise jamais un couple sur un objectif qui n'existe pas. "
         "<b>C3 — Dégradation de marge</b> : <code>(Taux marge N − Taux marge N-1) × 100 ≤ seuil marge</code>, en points ; "
         "ignoré si un des deux taux est absent. <b>C4 — Rupture / fermeture</b> : CA nul ou absent alors que le CA N-1 "
         "était positif — c'est le signal le plus grave, prioritaire sur tous les autres."),
        ("3. Sévérité (cumul simple)",
         "On compte le nombre de critères déclenchés parmi ceux qui sont applicables. "
         "<b>Critique</b> : C4 déclenché (rupture). <b>Flop majeur</b> : au moins 2 critères KO. "
         "<b>Flop modéré</b> : exactement 1 critère KO. <b>OK</b> : aucun. Le score affiché (ex <code>2/3</code>) se lit "
         "« critères KO / critères applicables » : un couple sans budget n'a que 2 critères applicables au lieu de 3."),
        ("4. Marge à risque (clé de priorisation)",
         "<code>Marge à risque = Marge N-1 − Marge</code>, en FCFA — la <b>valeur détruite</b>. C'est la métrique de "
         "priorisation retenue : en retail FMCG, un acheteur est jugé sur sa marge (sa ligne P&L), pas sur le seul "
         "top-line. Deux flops de même sévérité n'ont pas le même enjeu : un Hyper à −8 % peut détruire bien plus de "
         "valeur qu'un petit Supeco à −40 %. On trie donc par ce montant (logique Pareto : quelques couples concentrent "
         "l'essentiel de la valeur perdue). Avantage clé : le signe est toujours juste — un magasin qui gagne de la marge "
         "affiche une marge à risque négative et sort naturellement du haut de liste, contrairement au CA à risque qui "
         "pouvait afficher des valeurs négatives trompeuses sur des couples flaggés uniquement sur la marge. "
         "Le <code>CA à risque = CA N-1 − CA</code> reste affiché en <b>lecture top-line secondaire</b>."),
        ("5. Benchmark par pairs de format",
         "<code>Écart vs pairs = Vs N-1 (couple) − moyenne Vs N-1 (mêmes Format et Rayon)</code>. "
         "Comparer un magasin à ses pairs (les autres Supeco sur le même rayon, par exemple) permet de distinguer un "
         "problème <b>spécifique au magasin</b> (il décroche plus que ses pairs) d'un <b>sujet réseau</b> (tout le format "
         "baisse ensemble — ce n'est alors pas la faute du site mais un enjeu d'assortiment, d'appro ou de prix central)."),
        ("6. Décomposition Trafic / Panier",
         "Le CA se lit comme <code>CA = Débit × Panier moyen</code>. On décompose sa variation en deux effets "
         "(méthode symétrique, moyenne des périodes) : <code>Effet Trafic = ΔDébit × (Panier N + Panier N-1)/2</code> et "
         "<code>Effet Panier = ΔPanier × (Débit N + Débit N-1)/2</code>. Leur somme reconstitue ΔCA. Lecture : une perte "
         "portée par le <b>trafic</b> (moins de clients) appelle une action affluence/commerciale ; une perte portée par "
         "le <b>panier</b> (chaque client dépense moins) oriente vers l'assortiment, le prix ou la disponibilité rayon."),
        ("7. Commentaire rentabilité (matrice CA × Marge)",
         "Chaque rayon reçoit un commentaire automatique en croisant sa tendance CA (croissance / léger recul / "
         "décrochage) avec l'évolution de sa marge (amélioration / stable / dégradation). Neuf cas, neuf lectures : par "
         "exemple, un décrochage de CA accompagné d'une marge qui s'améliore signale un pilotage défensif (on vend moins "
         "mais on protège la marge), tandis qu'un décrochage doublé d'une dégradation de marge est le signal le plus "
         "préoccupant. Un suffixe précise la position vs budget."),
        ("8. Presets de seuils",
         "Trois réglages rapides : <b>Strict</b> (−5 % CA, −5 % Budget, −0,5 pt marge), <b>Standard</b> "
         "(−10 % / −10 % / −0,8), <b>Souple</b> (−15 % / −15 % / −1,2). Ils pré-remplissent les curseurs, qui restent ajustables "
         "manuellement. Le choix persiste durant la session."),
    ]
    for title, body in blocks:
        st.markdown(
            f"""<div class="method-block">
                <p style="font-weight:600;font-size:15px;margin:0 0 6px 0">{title}</p>
                <p style="font-size:14px;line-height:1.65;margin:0;color:{COL_TEXT_PRIMARY}">{body}</p>
            </div>""",
            unsafe_allow_html=True,
        )


# ============================================================
# 6. TESTS UNITAIRES / ASSERTIONS
# ============================================================

def _make_couple_row(**overrides) -> pd.DataFrame:
    base = dict(**{
        "Société": "TEST", "Rayon": "010 - BOISSON", "Site": "999 - Magasin Test",
        "Rayon_Libelle": "BOISSON", "Format": "Hyper", "Site_Libelle": "Magasin Test",
        "CA N-1": 1000.0, "Budget": 1000.0, "CA": 1000.0,
        "Vs N-1 (%)": 0.0, "Vs Bgt (%)": 0.0,
        "Marge N-1": 200.0, "Marge": 200.0,
        "Taux de Marge N-1": 0.20, "Taux de Marge": 0.20,
        "Débit N-1": 100.0, "Débit": 100.0, "Panier N-1": 10.0, "Panier": 10.0,
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


def test_ca_a_risque():
    ca = pd.Series([800.0, np.nan, 1200.0])
    ca_n1 = pd.Series([1000.0, 500.0, 1000.0])
    res = compute_ca_a_risque(ca, ca_n1)
    assert res.iloc[0] == 200.0   # perte de 200
    assert res.iloc[1] == 500.0   # rupture : perte de tout le CA N-1
    assert res.iloc[2] == -200.0  # progression : "risque" négatif


def test_marge_a_risque():
    marge = pd.Series([150.0, np.nan, 250.0])
    marge_n1 = pd.Series([200.0, 100.0, 200.0])
    res = compute_marge_a_risque(marge, marge_n1)
    assert res.iloc[0] == 50.0    # valeur détruite : 50
    assert res.iloc[1] == 100.0   # marge tombée : perte de toute la marge N-1
    assert res.iloc[2] == -50.0   # marge en hausse : "risque" négatif = gain


def test_marge_a_risque_dans_compute_flops():
    df = _make_couple_row(**{"Marge N-1": 200.0, "Marge": 160.0})
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert res.loc[0, "Marge_a_risque"] == 40.0


def test_marge_a_risque_colonnes_absentes():
    df = _make_couple_row()
    df = df.drop(columns=["Marge", "Marge N-1"])
    res = compute_flops(df, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8)
    assert pd.isna(res.loc[0, "Marge_a_risque"])


def test_benchmark_pairs():
    df = pd.DataFrame({
        "Format": ["Supeco", "Supeco", "Hyper"],
        "Rayon_Libelle": ["BOISSON", "BOISSON", "BOISSON"],
        "Vs N-1 (%)": [-0.20, -0.10, -0.05],
    })
    res = compute_benchmark_pairs(df)
    # moyenne des 2 Supeco = -0.15 ; le premier est à -0.20 -> écart -0.05
    assert abs(res.iloc[0] - (-0.05)) < 1e-9
    assert abs(res.iloc[1] - (0.05)) < 1e-9
    # le Hyper est seul dans son groupe -> écart nul
    assert abs(res.iloc[2] - 0.0) < 1e-9


def test_decomposition_trafic_panier_somme_egale_delta_ca():
    row = pd.Series({"Débit N-1": 100.0, "Débit": 90.0, "Panier N-1": 10.0, "Panier": 11.0})
    d = decompose_trafic_panier(row)
    delta_ca_reel = 90 * 11 - 100 * 10  # CA_N - CA_N-1 = 990 - 1000 = -10
    assert abs(d["delta_ca"] - delta_ca_reel) < 1e-6
    assert d["effet_trafic"] is not None and d["effet_panier"] is not None


def test_decomposition_trafic_panier_donnees_manquantes():
    row = pd.Series({"Débit N-1": np.nan, "Débit": 90.0, "Panier N-1": 10.0, "Panier": 11.0})
    d = decompose_trafic_panier(row)
    assert d["effet_trafic"] is None


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


def test_rayon_trigram():
    assert rayon_trigram("BOISSON") == "BOI"
    assert rayon_trigram("EPICERIE") == "EPI"
    assert rayon_trigram("") == "?"
    assert rayon_trigram(None) == "?"


def test_get_selected_rows_gere_dict_et_objet():
    class FakeSelectionObj:
        rows = [2]

    class FakeEventObj:
        selection = FakeSelectionObj()

    assert get_selected_rows(FakeEventObj()) == [2]
    assert get_selected_rows({"selection": {"rows": [3]}}) == [3]
    assert get_selected_rows({"selection": {}}) == []
    assert get_selected_rows(None) == []


def test_top_magasins_ca_et_progression():
    df = pd.DataFrame({
        "Site_Libelle": ["A", "B", "C"],
        "CA": [100.0, 300.0, 200.0],
        "Vs N-1 (%)": [0.05, -0.10, 0.20],
    })
    top_ca = top_magasins(df, "ca", 2)
    assert list(top_ca["Site_Libelle"]) == ["B", "C"]
    top_prog = top_magasins(df, "progression", 2)
    assert list(top_prog["Site_Libelle"]) == ["C", "A"]


def test_kpi_card_flat_et_3d_ne_plantent_pas():
    html_flat = kpi_card("CA", "10M", accent=COL_BLUE)
    html_cond = kpi_card_conditional("Vs N-1", "-8%", -0.08)
    assert "CA" in html_flat and "10M" in html_flat
    assert "Vs N-1" in html_cond


def test_welcome_criterion_contient_tag_et_formule():
    html = _welcome_criterion("C1 · Décrochage CA", "CA en repli vs N-1", "Vs N-1 (%) ≤ seuil CA", COL_BLUE)
    assert "C1" in html
    assert "seuil CA" in html
    assert COL_BLUE in html


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
