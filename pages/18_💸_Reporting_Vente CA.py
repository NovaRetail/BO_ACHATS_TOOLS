# -*- coding: utf-8 -*-
"""
18_💸_Reporting_Vente CA.py — V3
============================================================
SmartBuyer Hub — Module Reporting Commercial (Point de situation & Pilotage Acheteurs)

V3 — refonte design + recentrage usage :
    - Charte Apple (cohérente avec les autres modules SmartBuyer Hub) :
      fond #F2F2F7, cartes blanches, bleu #007AFF, rouge #FF3B30, vert #34C759
    - 4 onglets au lieu de 5 (Magasins fondu dans Vue d'ensemble) :
      1) Vue d'ensemble — le point de situation en 10 secondes
      2) Par Acheteur   — brief prêt à l'emploi par acheteur (GB/CK/AC)
      3) Flops          — table maître-détail complète, drill-down
      4) Export         — données brutes + exports
    - Correctifs techniques :
      * lecture robuste de st.dataframe(...).selection (dict ou objet)
      * scatter plot Vs N-1 x Δ Marge avec range X fixe (lisible même à -100%)

Règles de flop (validées) :
    C1 - Décrochage CA vs N-1   : Vs N-1 (%)  <= seuil CA
    C2 - Écart vs Budget        : Vs Bgt (%)  <= seuil Budget (ignoré si Budget NaN)
    C3 - Dégradation marge      : Delta Marge (pts) <= seuil Marge
    C4 - Rupture / fermeture    : CA NaN/0 alors que CA N-1 > 0 (prioritaire)
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
# 0. CONFIGURATION PAGE & THEME — Apple charter
# ============================================================

st.set_page_config(
    page_title="Reporting Vente CA",
    page_icon="💸",
    layout="wide",
    initial_sidebar_state="expanded",
)

COL_BG_PAGE = "#F2F2F7"
COL_BG_CARD = "#FFFFFF"
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

# Mapping Rayon -> Acheteur, aligné sur l'organisation PGC
# (GB Épicerie, CK Boissons, AC Droguerie-Parfumerie-Hygiène). À ajuster si
# l'organisation évolue.
RAYON_TO_BUYER = {
    "BOISSON": {"code": "CK", "perimetre": "Boissons"},
    "EPICERIE": {"code": "GB", "perimetre": "Épicerie"},
    "DROGUERIE": {"code": "AC", "perimetre": "Droguerie / Parfumerie-Hygiène"},
    "PARFUMERIE HYGIENE": {"code": "AC", "perimetre": "Droguerie / Parfumerie-Hygiène"},
}


def get_buyer_code(rayon_libelle) -> str:
    if not isinstance(rayon_libelle, str):
        return "N/A"
    lib_up = rayon_libelle.upper()
    for key, info in RAYON_TO_BUYER.items():
        if key in lib_up:
            return info["code"]
    return "N/A"


CUSTOM_CSS = f"""
<style>
    .stApp {{ background-color: {COL_BG_PAGE}; }}
    section[data-testid="stSidebar"] {{
        background-color: {COL_BG_CARD};
        border-right: 0.5px solid {COL_BORDER};
    }}
    h1, h2, h3, h4 {{ color: {COL_TEXT_PRIMARY}; font-weight: 600; }}
    p, span, div, label {{ color: {COL_TEXT_PRIMARY}; }}
    .kpi-card {{
        background-color: {COL_BG_CARD};
        border-radius: 16px;
        padding: 1rem 1.1rem;
        margin-bottom: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    .kpi-label {{ font-size: 12px; color: {COL_TEXT_SECONDARY}; margin: 0 0 6px 0; font-weight: 500; }}
    .kpi-value {{ font-size: 26px; font-weight: 700; color: {COL_TEXT_PRIMARY}; margin: 0; }}
    .badge-pill {{
        font-size: 11px; padding: 4px 11px; border-radius: 20px;
        display: inline-block; font-weight: 600;
    }}
    .card {{
        background-color: {COL_BG_CARD};
        border-radius: 16px;
        padding: 1.1rem 1.25rem;
        margin-bottom: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    .flop-row {{
        background-color: {COL_BG_CARD};
        border-radius: 14px;
        padding: 0.7rem 1rem;
        margin-bottom: 6px;
        display: flex; align-items: center; justify-content: space-between;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }}
    .criterion-line {{
        display: flex; justify-content: space-between; padding: 7px 0;
        border-bottom: 0.5px solid {COL_BORDER}; font-size: 13.5px;
    }}
    .criterion-line:last-child {{ border-bottom: none; }}
    .buyer-header {{
        display: flex; align-items: center; gap: 12px; margin-bottom: 4px;
    }}
    .buyer-avatar {{
        width: 44px; height: 44px; border-radius: 50%;
        background: {COL_BLUE}1A; color: {COL_BLUE};
        display: flex; align-items: center; justify-content: center;
        font-weight: 700; font-size: 15px;
    }}
    div[data-testid="stMetric"] {{
        background-color: {COL_BG_CARD}; border-radius: 14px;
        padding: 0.75rem 1rem; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    .preset-caption {{ font-size: 11px; color: {COL_TEXT_SECONDARY}; margin-top: -6px; }}
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
    """Lecteur universel encoding-safe (xlsx ou csv), façon utils_io.py."""
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
# 2. MOTEUR DE CALCUL — FLOPS & SÉVÉRITÉ (niveau Couple)
# ============================================================

def compute_delta_marge_pts(taux_marge: pd.Series, taux_marge_n1: pd.Series) -> pd.Series:
    return (taux_marge - taux_marge_n1) * 100


def compute_flops(df: pd.DataFrame, seuil_ca: float, seuil_bgt: float, seuil_marge: float) -> pd.DataFrame:
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


def get_selected_rows(event) -> list:
    """Extrait la liste des index sélectionnés depuis st.dataframe(..., on_select="rerun").
    Robuste aux deux formats renvoyés selon la version de Streamlit :
      - objet avec .selection.rows (API récente)
      - dict brut {"selection": {"rows": [...]}}
    """
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


def build_buyer_brief(buyer_code: str, rayons_info: pd.DataFrame, flops_buyer: pd.DataFrame,
                       seuil_ca: float, seuil_bgt: float, seuil_marge: float) -> str:
    """Génère un texte de brief prêt à copier-coller pour discuter avec l'acheteur."""
    lines = [f"POINT ACHETEUR {buyer_code} — {pd.Timestamp.today().strftime('%d/%m/%Y')}", ""]
    for _, r in rayons_info.iterrows():
        comment = build_rentabilite_comment(r, seuil_ca, seuil_bgt, seuil_marge)
        lines.append(f"• {r['Rayon_Libelle']} : CA {fmt_fcfa(r['CA'])} ({fmt_pct(r['Vs N-1 (%)'])} vs N-1)")
        lines.append(f"  {comment}")
        lines.append("")

    flops_sorted = flops_buyer[flops_buyer["Severite"] != "OK"].sort_values(
        ["Nb_Criteres_KO", "Vs N-1 (%)"], ascending=[False, True]
    )
    if flops_sorted.empty:
        lines.append("Aucun point d'alerte magasin à date.")
    else:
        lines.append(f"Points d'attention magasins ({len(flops_sorted)}) :")
        for _, r in flops_sorted.head(10).iterrows():
            lines.append(
                f"  - [{r['Severite']}] {r['Site_Libelle']} : CA {fmt_pct(r['Vs N-1 (%)'])} vs N-1, "
                f"Δ marge {fmt_pt(r['Delta_Marge_pt'])}"
            )
    return "\n".join(lines)


def build_copil_export(df_global: pd.DataFrame, df_rayon: pd.DataFrame, df_flops: pd.DataFrame,
                        seuil_ca: float, seuil_bgt: float, seuil_marge: float) -> bytes:
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
            cols = ["Rayon_Libelle", "Acheteur", "CA", "CA N-1", "Vs N-1 (%)", "Budget", "Vs Bgt (%)",
                    "Marge", "Taux de Marge", "Commentaire rentabilité"]
            rayon_export[[c for c in cols if c in rayon_export.columns]].to_excel(
                writer, index=False, sheet_name="Synthese Rayons"
            )

        flops_export = df_flops.sort_values(["Nb_Criteres_KO", "Vs N-1 (%)"], ascending=[False, True])
        cols = ["Site_Libelle", "Rayon_Libelle", "Acheteur", "Format", "CA", "CA N-1", "Vs N-1 (%)",
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
        st.markdown(f"## Point de situation — {societe_sel}")
        st.caption("Vue d'ensemble · Par acheteur · Détail des flops")
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
            f"""<div style="background:{COL_RED_BG};border-radius:14px;
            padding:0.75rem 1rem;margin:12px 0;color:{COL_RED};font-weight:500">
            🔴 {nb_critique} rupture(s) totale de CA détectée(s) — voir onglet Flops</div>""",
            unsafe_allow_html=True,
        )

    st.markdown("---")
    tab_exec, tab_buyers, tab_flops, tab_export = st.tabs(
        ["🎯 Vue d'ensemble", "👤 Par Acheteur", "🚩 Flops", "📤 Export"]
    )

    # ---------------- TAB VUE D'ENSEMBLE ----------------
    with tab_exec:
        col_a, col_b = st.columns([1, 2])
        with col_a:
            st.markdown("#### Répartition sévérité")
            counts = df_flops["Severite"].value_counts().reindex(SEVERITY_ORDER).fillna(0)
            colors_map = [SEVERITY_STYLE[s]["text"] for s in SEVERITY_ORDER]
            fig_donut = go.Figure(data=[go.Pie(labels=SEVERITY_ORDER, values=counts.values, hole=0.55, marker=dict(colors=colors_map))])
            fig_donut.update_layout(template=PLOTLY_TEMPLATE, height=280, showlegend=True, margin=dict(t=10, b=10))
            st.plotly_chart(fig_donut, use_container_width=True)

        with col_b:
            st.markdown("#### Top 5 points d'attention")
            top5 = df_flops[df_flops["Severite"] != "OK"].sort_values(["Nb_Criteres_KO", "Vs N-1 (%)"], ascending=[False, True]).head(5)
            if top5.empty:
                st.success("Aucun flop détecté sur ce périmètre.")
            else:
                for _, r in top5.iterrows():
                    st.markdown(
                        f"""<div class="flop-row">
                            <div style="display:flex;align-items:center;gap:10px">{severity_badge(r['Severite'])}
                            <span>{r['Site_Libelle']} — {r['Rayon_Libelle']}</span></div>
                            <span style="color:{COL_TEXT_SECONDARY}">Vs N-1 {fmt_pct(r['Vs N-1 (%)'])} · Score {r['Score_Label']}</span>
                        </div>""",
                        unsafe_allow_html=True,
                    )

        st.markdown("#### Rentabilité par rayon")
        rayon_cols = st.columns(min(4, max(1, len(df_rayon_f))))
        for i, (_, r) in enumerate(df_rayon_f.iterrows()):
            comment = build_rentabilite_comment(r, seuil_ca, seuil_bgt, seuil_marge)
            with rayon_cols[i % len(rayon_cols)]:
                st.markdown(
                    f"""<div class="card"><p style="font-weight:600;margin:0 0 6px 0">{r['Rayon_Libelle']}</p>
                    <p style="font-size:13px;color:{COL_TEXT_SECONDARY};margin:0">{comment}</p></div>""",
                    unsafe_allow_html=True,
                )

        st.markdown("#### Magasins les plus en difficulté")
        agg_magasin = (
            df_flops.groupby(["Site_Libelle", "Format"], as_index=False)
            .agg(CA=("CA", "sum"), CA_N1=("CA N-1", "sum"),
                 Nb_Flops=("Severite", lambda s: (s != "OK").sum()),
                 Nb_Critiques=("Severite", lambda s: (s == "Critique").sum()))
            .sort_values(["Nb_Critiques", "Nb_Flops"], ascending=False)
            .head(8)
        )
        agg_magasin["Vs N-1 (%)"] = np.where(
            agg_magasin["CA_N1"] > 0, (agg_magasin["CA"] - agg_magasin["CA_N1"]) / agg_magasin["CA_N1"], np.nan
        )
        display_mag = agg_magasin.copy()
        display_mag["CA"] = display_mag["CA"].apply(fmt_fcfa)
        display_mag["Vs N-1 (%)"] = display_mag["Vs N-1 (%)"].apply(fmt_pct)
        st.dataframe(
            display_mag[["Site_Libelle", "Format", "CA", "Vs N-1 (%)", "Nb_Flops", "Nb_Critiques"]],
            use_container_width=True, hide_index=True,
        )

    # ---------------- TAB PAR ACHETEUR ----------------
    with tab_buyers:
        buyers_present = sorted(
            set(df_rayon_f["Acheteur"].dropna().unique()) | set(df_flops["Acheteur"].dropna().unique())
        )
        buyers_present = [b for b in buyers_present if b != "N/A"]

        if not buyers_present:
            st.info("Aucun acheteur identifié sur ce périmètre (vérifie le mapping Rayon -> Acheteur).")
        else:
            buyer_sel = st.radio("Acheteur", options=buyers_present, horizontal=True)

            rayons_buyer = df_rayon_f[df_rayon_f["Acheteur"] == buyer_sel]
            flops_buyer = df_flops[df_flops["Acheteur"] == buyer_sel]
            perimetre = ", ".join(sorted(rayons_buyer["Rayon_Libelle"].unique()))

            nb_crit = int((flops_buyer["Severite"] == "Critique").sum())
            nb_maj = int((flops_buyer["Severite"] == "Flop majeur").sum())
            nb_mod = int((flops_buyer["Severite"] == "Flop modéré").sum())
            ca_total = rayons_buyer["CA"].sum()

            st.markdown(
                f"""<div class="buyer-header">
                    <div class="buyer-avatar">{buyer_sel}</div>
                    <div><p style="font-weight:600;font-size:16px;margin:0">Acheteur {buyer_sel}</p>
                    <p style="font-size:13px;color:{COL_TEXT_SECONDARY};margin:0">{perimetre}</p></div>
                </div>""",
                unsafe_allow_html=True,
            )
            st.markdown("<br>", unsafe_allow_html=True)

            m1, m2, m3, m4 = st.columns(4)
            m1.markdown(kpi_card("CA périmètre", fmt_fcfa(ca_total), COL_BLUE), unsafe_allow_html=True)
            m2.markdown(kpi_card("Critiques", str(nb_crit), COL_RED if nb_crit else COL_GREEN), unsafe_allow_html=True)
            m3.markdown(kpi_card("Flops majeurs", str(nb_maj), COL_ORANGE if nb_maj else COL_GREEN), unsafe_allow_html=True)
            m4.markdown(kpi_card("Flops modérés", str(nb_mod), COL_AMBER if nb_mod else COL_GREEN), unsafe_allow_html=True)

            st.markdown("#### Diagnostic par rayon")
            for _, r in rayons_buyer.iterrows():
                comment = build_rentabilite_comment(r, seuil_ca, seuil_bgt, seuil_marge)
                st.markdown(
                    f"""<div class="card">
                        <p style="font-weight:600;margin:0 0 4px 0">{r['Rayon_Libelle']}</p>
                        <p style="font-size:13px;color:{COL_TEXT_SECONDARY};margin:0 0 8px 0">
                        CA {fmt_fcfa(r['CA'])} · {fmt_pct(r['Vs N-1 (%)'])} vs N-1 · {fmt_pct(r['Vs Bgt (%)'])} vs Budget</p>
                        <p style="font-size:14px;margin:0">{comment}</p>
                    </div>""",
                    unsafe_allow_html=True,
                )

            st.markdown("#### Magasins à traiter en priorité")
            flops_a_traiter = flops_buyer[flops_buyer["Severite"] != "OK"].sort_values(
                ["Nb_Criteres_KO", "Vs N-1 (%)"], ascending=[False, True]
            )
            if flops_a_traiter.empty:
                st.success("Aucun point d'alerte sur le périmètre de cet acheteur.")
            else:
                for _, r in flops_a_traiter.iterrows():
                    st.markdown(
                        f"""<div class="flop-row">
                            <div style="display:flex;align-items:center;gap:10px">{severity_badge(r['Severite'])}
                            <span>{r['Site_Libelle']} — {r['Rayon_Libelle']}</span></div>
                            <span style="color:{COL_TEXT_SECONDARY}">Vs N-1 {fmt_pct(r['Vs N-1 (%)'])} · Δ marge {fmt_pt(r['Delta_Marge_pt'])}</span>
                        </div>""",
                        unsafe_allow_html=True,
                    )

            st.markdown("#### Brief prêt à partager")
            brief_text = build_buyer_brief(buyer_sel, rayons_buyer, flops_buyer, seuil_ca, seuil_bgt, seuil_marge)
            st.text_area("Copier ce texte pour le point avec l'acheteur", value=brief_text, height=260)

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
                "Acheteur": df_view["Acheteur"],
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
            selected_rows = get_selected_rows(event)
            if not selected_rows:
                st.markdown(
                    f"""<div class="card"><p style="color:{COL_TEXT_SECONDARY};margin:0">
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
                    f"""<div class="card">
                        <p style="font-weight:600;font-size:15px;margin:0 0 2px 0">{r['Site_Libelle']}</p>
                        <p style="font-size:13px;color:{COL_TEXT_SECONDARY};margin:0 0 12px 0">{r['Rayon_Libelle']} · {r['Format']} · Acheteur {r['Acheteur']}</p>
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

            # Range X fixe : les ruptures totales (-100%) restent lisibles sans écraser le nuage de points
            max_val = float(scatter_df["Vs N-1 (%)"].max())
            x_range = [-1.05, max(0.2, max_val + 0.05)]
            fig_scatter.update_xaxes(range=x_range)
            fig_scatter.add_annotation(
                x=x_range[0] + 0.03, y=scatter_df["Delta_Marge_pt"].min(),
                text="Zone à risque", showarrow=False, font=dict(color=COL_RED, size=11),
                xanchor="left", yanchor="bottom",
            )
            st.plotly_chart(fig_scatter, use_container_width=True)
        else:
            st.info("Pas assez de données valides pour tracer le scatter plot.")

    # ---------------- TAB EXPORT ----------------
    with tab_export:
        st.markdown("#### Exports disponibles")
        exp1, exp2 = st.columns(2)
        with exp1:
            export_cols = ["Site_Libelle", "Rayon_Libelle", "Acheteur", "Format", "CA", "CA N-1", "Vs N-1 (%)",
                           "Budget", "Vs Bgt (%)", "Delta_Marge_pt", "Severite", "Score_Label"]
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_flops[export_cols].to_excel(writer, index=False, sheet_name="Flops")
            st.download_button(
                "📥 Excel Flops (périmètre filtré)", data=buffer.getvalue(),
                file_name=f"flops_{societe_sel}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with exp2:
            csv_buffer = df_flops.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "📥 CSV complet (toutes colonnes)", data=csv_buffer,
                file_name=f"detail_{societe_sel}.csv", mime="text/csv", use_container_width=True,
            )

        st.markdown("#### Données brutes")
        st.dataframe(df_flops, use_container_width=True, hide_index=True)


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


def test_get_buyer_code_mapping():
    assert get_buyer_code("BOISSON") == "CK"
    assert get_buyer_code("EPICERIE") == "GB"
    assert get_buyer_code("DROGUERIE") == "AC"
    assert get_buyer_code("PARFUMERIE HYGIENE") == "AC"
    assert get_buyer_code("RAYON INCONNU") == "N/A"
    assert get_buyer_code(None) == "N/A"


def test_get_selected_rows_gere_dict_et_objet():
    class FakeSelectionObj:
        rows = [2]

    class FakeEventObj:
        selection = FakeSelectionObj()

    assert get_selected_rows(FakeEventObj()) == [2]
    assert get_selected_rows({"selection": {"rows": [3]}}) == [3]
    assert get_selected_rows({"selection": {}}) == []
    assert get_selected_rows(None) == []


def test_scatter_x_range_couvre_rupture_totale():
    """Vérifie que le calcul du range X du scatter plot couvre bien -100% (rupture)."""
    scatter_df = pd.DataFrame({"Vs N-1 (%)": [-1.0, -0.3, 0.05], "Delta_Marge_pt": [0, -1, 2]})
    max_val = float(scatter_df["Vs N-1 (%)"].max())
    x_range = [-1.05, max(0.2, max_val + 0.05)]
    assert x_range[0] <= -1.0
    assert x_range[1] >= scatter_df["Vs N-1 (%)"].max()


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
    assert result["rayon"].loc[0, "Acheteur"] == "CK"


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
