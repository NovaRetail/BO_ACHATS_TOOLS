
"""
12_🏪_Bascule_XD.py — SmartBuyer Hub
Commando XD · Analyse DL vers Cross-Docking
Version corrigée :
- upload uniquement dans la sidebar
- fichier à un seul onglet : lecture automatique du premier onglet
- alias corrigé pour "Qté rec" / "Qte rec"
- landing page explicative au style SmartBuyer
- charte visuelle inspirée des apps SmartBuyer existantes
"""

from __future__ import annotations

import io
import re
import sys
import importlib.util
import unicodedata
from datetime import timedelta

import numpy as np
import pandas as pd
import streamlit as st


# ═══════════════════════════════════════════════════════════════════════════════
# CONFIG PAGE
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Bascule XD · SmartBuyer",
    page_icon="🏪",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ═══════════════════════════════════════════════════════════════════════════════
# CHARTE SMARTBUYER
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display",
                 "SF Pro Text", "Helvetica Neue", Arial, sans-serif !important;
    background-color: #F2F2F7;
}
.stApp { background: #F2F2F7; }
.main .block-container { padding-top: 1.8rem; max-width: 1350px; }
[data-testid="stSidebar"] { background: #F2F2F7 !important; border-right: 0.5px solid #D1D1D6 !important; }
[data-testid="stMetric"] { background: #FFFFFF !important; border: 0.5px solid #E5E5EA !important; border-radius: 12px !important; padding: 16px 18px !important; }
[data-testid="stMetricLabel"] { font-size: 11px !important; font-weight: 500 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 600 !important; color: #1C1C1E !important; letter-spacing: -0.02em !important; }
[data-testid="stTabs"] button[role="tab"] { font-size: 13px !important; font-weight: 500 !important; padding: 8px 16px !important; color: #8E8E93 !important; border-radius: 0 !important; border-bottom: 2px solid transparent !important; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color: #007AFF !important; border-bottom: 2px solid #007AFF !important; background: transparent !important; }
[data-testid="stTabs"] [role="tablist"] { border-bottom: 0.5px solid #E5E5EA !important; }
[data-testid="stDataFrame"] { border: 0.5px solid #E5E5EA !important; border-radius: 10px !important; }
[data-testid="stDataFrame"] th { background: #F2F2F7 !important; font-size: 11px !important; font-weight: 600 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stFileUploader"] { border: 1.5px dashed #D1D1D6 !important; border-radius: 10px !important; background: #F9F9FB !important; }
.stDownloadButton > button { background: #007AFF !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; font-size: 13px !important; padding: 10px 24px !important; width: 100% !important; }
.stButton > button[kind="primary"] { background: #007AFF !important; border: none !important; border-radius: 8px !important; font-weight: 600 !important; }
hr { border-color: #E5E5EA !important; margin: 1rem 0 !important; }

.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 1.5rem; }
.section-label { font-size: 11px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 10px; }
.alert-card  { padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; font-size: 13px; line-height: 1.5; border-left: 3px solid; background: #FFFFFF; }
.alert-red   { background: #FFF2F2; border-color: #FF3B30; color: #3A0000; }
.alert-amber { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }
.alert-green { background: #F0FFF4; border-color: #34C759; color: #003A10; }
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }
.alert-purple{ background: #F5F0FF; border-color: #AF52DE; color: #1A0033; }

.format-card { border-radius: 12px; padding: 14px 16px; margin-bottom: 6px; border: 0.5px solid; }
.format-hyper  { background: #EFF6FF; border-color: #B3D9FF; }
.format-market { background: #F0FFF4; border-color: #A8E6BF; }
.format-supeco { background: #F5F0FF; border-color: #D9B3FF; }

.badge { display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: 600; }
.badge-hyper  { background: #154360; color: #FFFFFF; }
.badge-market { background: #145A32; color: #FFFFFF; }
.badge-supeco { background: #6E2F8A; color: #FFFFFF; }

.col-required { background: #F0F8FF; border: 0.5px solid #B3D9FF; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; display: flex; align-items: flex-start; gap: 10px; }
.col-name { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
.card { background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:16px;margin-bottom:10px; }
.small-muted { font-size:12px;color:#8E8E93; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PARAMÈTRES MÉTIER
# ═══════════════════════════════════════════════════════════════════════════════

DEFAULT_START_DATE = pd.Timestamp("2026-01-01")
DEFAULT_XD_THRESHOLD = 100_000
DEFAULT_PLATFORM_COST_PER_PACKAGE = 90
DEFAULT_MIN_ORDERS = 5

HYPERS = {"10202", "10203", "10301"}
MARKETS = {"10604", "10206", "10208", "10209", "10705"}
SUPECO = {"10601", "10602", "10603", "10605"}

DAYS_FR = {
    0: "Lundi",
    1: "Mardi",
    2: "Mercredi",
    3: "Jeudi",
    4: "Vendredi",
    5: "Samedi",
    6: "Dimanche",
}
DAY_TO_NUM = {v: k for k, v in DAYS_FR.items()}


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS GÉNÉRAUX
# ═══════════════════════════════════════════════════════════════════════════════

def package_installed(package_name: str) -> bool:
    return importlib.util.find_spec(package_name) is not None


def fmt(n):
    if pd.isna(n) or n is None:
        return "—"
    try:
        n = float(n)
    except Exception:
        return "—"
    a = abs(n)
    if a >= 1_000_000:
        return f"{n/1_000_000:.1f} M"
    if a >= 1_000:
        return f"{int(n/1_000)} K"
    return f"{int(n):,}".replace(",", " ")


def fmt_xof(n):
    if pd.isna(n) or n is None:
        return "—"
    return f"{float(n):,.0f} FCFA".replace(",", " ")


def fmt_pct(v, dec=1):
    if pd.isna(v) or v is None:
        return "—"
    return f"{float(v):.{dec}f}%"


def normalize_text(value) -> str:
    if value is None:
        return ""
    value = str(value).strip().lower()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(c for c in value if not unicodedata.combining(c))
    value = re.sub(r"[^a-z0-9]+", "", value)
    return value


def find_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    normalized_cols = {normalize_text(c): c for c in df.columns}
    for alias in aliases:
        key = normalize_text(alias)
        if key in normalized_cols:
            return normalized_cols[key]
    return None


def detect_columns(df: pd.DataFrame) -> dict:
    """
    Détection robuste des colonnes.
    Correction clé : ajout de "Qté rec" / "Qte rec" pour éviter l'erreur qte_rec.
    """
    aliases = {
        "fou": [
            "Fou", "FOU", "Code fournisseur", "Fournisseur",
            "Code fourn", "Code fourn.", "Code Four", "Four."
        ],
        "nom_fourn": [
            "Nom fourn,", "Nom fourn", "Nom fournisseur",
            "Libellé fournisseur", "Libelle fournisseur", "Nom fourn.",
            "Nom four", "Fournisseur libellé", "Nom Four."
        ],
        "site": [
            "Site", "Code site", "Magasin", "Code magasin",
            "Etablissement", "Établissement", "Code établissement"
        ],
        "code": [
            "Code", "Code article", "Article", "Code produit",
            "SKU", "Code EAN", "Référence", "Reference"
        ],
        "n_cde": [
            "N° Cde", "N Cde", "No Cde", "Num Cde",
            "Numero commande", "Numéro commande", "N° commande",
            "Commande", "No commande", "N commande"
        ],
        "date_cde": [
            "Date de commande", "Date commande", "Dt Cde",
            "Date Cde", "Date cmd", "Date Cmd", "Dt commande"
        ],
        "dt_rec": [
            "Dt Rec", "Date réception", "Date reception",
            "Date de réception", "Date de reception",
            "Date rec", "Date Rec", "Dt réception", "Dt reception"
        ],
        "qte_cde": [
            "Qté cde", "Qte cde", "Quantité commandée",
            "Quantite commandee", "Qte commande", "Qté commandée",
            "Qte Cde", "Qté commande", "Qte cdee"
        ],
        "qte_rec": [
            "Qté rec", "Qte rec", "Qté reçue", "Qte recue",
            "Quantité reçue", "Quantite recue", "Qte reception",
            "Qté réception", "Qte Rec", "Qté Rec", "Qté réceptionnée",
            "Qte receptionnee"
        ],
        "px_revient": [
            "Px revient", "Prix revient", "Prix de revient",
            "PR", "Px Revient", "Prix achat", "Prix d'achat", "PA"
        ],
        "colis": [
            "Colis", "Nb colis", "Nombre colis", "PCB",
            "Nb Colis", "Nombre de colis", "Nb. colis"
        ],
        "sit": [
            "Sit", "Situation", "Statut", "Statut commande",
            "Code situation", "Statut Cde", "Code Sit"
        ],
    }
    return {key: find_column(df, names) for key, names in aliases.items()}


def clean_numeric(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    s = series.astype(str).str.strip()
    s = s.str.replace("\u00a0", "", regex=False)
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    s = s.replace({
        "": np.nan,
        "nan": np.nan,
        "NaN": np.nan,
        "None": np.nan,
        "NULL": np.nan,
        "-": np.nan,
    })
    return pd.to_numeric(s, errors="coerce")


def parse_date_series(series: pd.Series) -> pd.Series:
    dt = pd.to_datetime(series, errors="coerce", dayfirst=True)

    numeric = pd.to_numeric(series, errors="coerce")
    mask_excel_serial = dt.isna() & numeric.between(20_000, 80_000)

    if mask_excel_serial.any():
        dt.loc[mask_excel_serial] = pd.to_datetime(
            numeric.loc[mask_excel_serial],
            unit="D",
            origin="1899-12-30",
            errors="coerce",
        )
    return dt


def safe_div(num, den):
    try:
        if den is None or pd.isna(den) or den == 0:
            return np.nan
        return num / den
    except Exception:
        return np.nan


def mode_day(series: pd.Series) -> str:
    s = pd.to_datetime(series, errors="coerce").dropna()
    if s.empty:
        return "N/A"
    m = s.dt.dayofweek.mode()
    if m.empty:
        return "N/A"
    return DAYS_FR.get(int(m.iloc[0]), "N/A")


def classify_group(site) -> str:
    if pd.isna(site):
        return "Site hors groupe"
    site = str(site).strip().split(".")[0]
    if site in HYPERS:
        return "Hypers"
    if site in MARKETS:
        return "Markets"
    if site in SUPECO:
        return "Supeco"
    return "Site hors groupe"


def join_unique(values) -> str:
    vals = pd.Series(values).dropna().astype(str).unique().tolist()
    vals = [v for v in vals if v and v.lower() != "nan"]
    return " / ".join(sorted(vals)) if vals else "N/A"


def yes_no(value: bool) -> str:
    return "Oui" if bool(value) else "Non"


def compute_current_cadence(cdes_mois: float) -> str:
    if pd.isna(cdes_mois):
        return "N/A"
    if cdes_mois >= 20:
        return "Hebdo"
    if cdes_mois >= 8:
        return "Bi-mensuel"
    return "Mensuel"


def get_cycles_from_cadence(cadence: str) -> int:
    if cadence == "Hebdo":
        return 4
    if cadence == "Bi-mensuel":
        return 2
    if cadence == "Mensuel":
        return 1
    return 0


def compute_xd_cadence(colis_xd_mois: float) -> tuple[str, int, float, str]:
    if pd.isna(colis_xd_mois) or colis_xd_mois <= 0:
        return "N/A", 0, 0, "N/A"

    if colis_xd_mois > 3000:
        cadence = "Hebdo"
    elif colis_xd_mois >= 1000:
        cadence = "Bi-mensuel"
    else:
        cadence = "Mensuel"

    cycles = get_cycles_from_cadence(cadence)
    colis_livraison = colis_xd_mois / cycles

    if cadence == "Mensuel" and colis_livraison > 500:
        cadence = "Bi-mensuel"
        cycles = 2
        colis_livraison = colis_xd_mois / cycles

    if cadence == "Bi-mensuel" and colis_livraison > 600:
        cadence = "Hebdo"
        cycles = 4
        colis_livraison = colis_xd_mois / cycles

    if colis_livraison < 500:
        alerte = "🟢 < 500"
    elif colis_livraison < 1000:
        alerte = "🟠 500-999"
    else:
        alerte = "🔴 >= 1000"

    return cadence, cycles, colis_livraison, alerte


def compute_cutoff_day(livraison_day: str, lead_time_days) -> str:
    if livraison_day in ["N/A", "", None] or pd.isna(lead_time_days):
        return "N/A"
    base = DAY_TO_NUM.get(livraison_day)
    if base is None:
        return "N/A"
    try:
        lt = int(round(float(lead_time_days)))
    except Exception:
        return "N/A"

    cutoff = base - lt
    if cutoff >= 0:
        return DAYS_FR[cutoff]

    while cutoff < 0:
        cutoff += 7
    return f"S-1 {DAYS_FR[cutoff]}"


# ═══════════════════════════════════════════════════════════════════════════════
# LECTURE FICHIER — UN SEUL ONGLET
# ═══════════════════════════════════════════════════════════════════════════════

def get_file_extension(filename: str) -> str:
    return filename.lower().rsplit(".", 1)[-1].strip()


def check_engine_available(filename: str):
    ext = get_file_extension(filename)

    if ext == "xlsb" and not package_installed("pyxlsb"):
        raise ImportError(
            "Le fichier est au format .xlsb mais la librairie pyxlsb n'est pas installée. "
            "Ajoute pyxlsb dans requirements.txt ou convertis le fichier en .xlsx."
        )

    if ext == "xls" and not package_installed("xlrd"):
        raise ImportError(
            "Le fichier est au format .xls mais la librairie xlrd n'est pas installée. "
            "Ajoute xlrd dans requirements.txt ou convertis le fichier en .xlsx."
        )

    if ext == "xlsx" and not package_installed("openpyxl"):
        raise ImportError(
            "Le fichier est au format .xlsx mais la librairie openpyxl n'est pas installée. "
            "Ajoute openpyxl dans requirements.txt."
        )


def get_excel_engine(filename: str) -> str:
    ext = get_file_extension(filename)
    if ext == "xlsb":
        return "pyxlsb"
    if ext == "xls":
        return "xlrd"
    return "openpyxl"


@st.cache_data(show_spinner=False)
def read_input_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Le fichier ne contient qu'un onglet : on lit automatiquement le premier onglet.
    Pas de selectbox d'onglet.
    """
    ext = get_file_extension(filename)
    buffer = io.BytesIO(file_bytes)

    if ext == "csv":
        for encoding in ["utf-8-sig", "utf-8", "latin1"]:
            for sep in [None, ";", ",", "\t"]:
                try:
                    buffer.seek(0)
                    return pd.read_csv(buffer, encoding=encoding, sep=sep, engine="python", dtype=str)
                except Exception:
                    continue
        raise ValueError("Impossible de lire le CSV. Vérifie l'encodage ou le séparateur.")

    check_engine_available(filename)
    engine = get_excel_engine(filename)
    return pd.read_excel(buffer, sheet_name=0, engine=engine, dtype=str)


# ═══════════════════════════════════════════════════════════════════════════════
# PRÉPARATION DONNÉES
# ═══════════════════════════════════════════════════════════════════════════════

def prepare_data(raw_df: pd.DataFrame, start_date: pd.Timestamp, mapping: dict) -> tuple[pd.DataFrame, dict]:
    df = raw_df.copy()
    initial_rows = len(df)

    required = [
        "fou", "nom_fourn", "site", "code", "n_cde",
        "date_cde", "dt_rec", "qte_cde", "qte_rec",
        "px_revient", "colis", "sit"
    ]

    missing = [field for field in required if mapping.get(field) is None]
    if missing:
        readable_missing = {
            "fou": "Fou / Code fournisseur",
            "nom_fourn": "Nom fourn, / Nom fournisseur",
            "site": "Site",
            "code": "Code / Code article",
            "n_cde": "N° Cde",
            "date_cde": "Date de commande",
            "dt_rec": "Dt Rec / Date réception",
            "qte_cde": "Qté cde",
            "qte_rec": "Qté rec / Qté reçue",
            "px_revient": "Px revient",
            "colis": "Colis",
            "sit": "Sit",
        }
        msg = "\n".join([f"- {readable_missing.get(m, m)}" for m in missing])
        available = "\n".join([f"- {c}" for c in df.columns])

        raise ValueError(
            "Colonnes obligatoires non détectées :\n"
            f"{msg}\n\n"
            "Colonnes disponibles dans ton fichier :\n"
            f"{available}\n\n"
            "Solution : corrige l'alias dans detect_columns() ou renomme la colonne dans le fichier source."
        )

    df = df.rename(columns={
        mapping["fou"]: "Fou",
        mapping["nom_fourn"]: "Nom fournisseur",
        mapping["site"]: "Site",
        mapping["code"]: "Code article",
        mapping["n_cde"]: "N° Cde",
        mapping["date_cde"]: "Date de commande",
        mapping["dt_rec"]: "Dt Rec",
        mapping["qte_cde"]: "Qté cde",
        mapping["qte_rec"]: "Qté rec",
        mapping["px_revient"]: "Px revient",
        mapping["colis"]: "Colis",
        mapping["sit"]: "Sit",
    })

    df["Fou"] = df["Fou"].astype(str).str.strip()
    df["Nom fournisseur"] = df["Nom fournisseur"].astype(str).str.strip()
    df["Site"] = df["Site"].astype(str).str.strip().str.split(".").str[0]
    df["Code article"] = df["Code article"].astype(str).str.strip()
    df["N° Cde"] = df["N° Cde"].astype(str).str.strip()

    df["Date de commande"] = parse_date_series(df["Date de commande"])
    df["Dt Rec"] = parse_date_series(df["Dt Rec"])

    df["Qté cde"] = clean_numeric(df["Qté cde"])
    df["Qté rec"] = clean_numeric(df["Qté rec"])
    df["Px revient"] = clean_numeric(df["Px revient"])
    df["Colis"] = clean_numeric(df["Colis"]).fillna(0)

    df["Sit brut"] = df["Sit"].astype(str).str.strip()
    df["Sit clean"] = (
        df["Sit brut"]
        .str.replace(".0", "", regex=False)
        .str.extract(r"(\d+)", expand=False)
        .fillna(df["Sit brut"])
    )

    qte_missing_or_zero = int(df["Qté cde"].isna().sum() + df["Qté cde"].fillna(0).eq(0).sum())
    px_missing = int(df["Px revient"].isna().sum())
    date_cde_missing = int(df["Date de commande"].isna().sum())
    dt_rec_missing = int(df["Dt Rec"].isna().sum())

    df = df[df["Date de commande"].notna()].copy()
    df = df[df["Date de commande"] >= start_date].copy()

    if df.empty:
        raise ValueError(
            "Aucune ligne disponible après filtre de date. "
            "Vérifie la date de début d’analyse ou le format de la colonne Date de commande."
        )

    last_date = df["Date de commande"].max()
    nb_days = max((last_date - start_date).days + 1, 1)
    nb_months = max(nb_days / 30.44, 1 / 30.44)

    df["Groupe magasin"] = df["Site"].apply(classify_group)
    df["Valeur commande"] = df["Qté cde"].fillna(0) * df["Px revient"].fillna(0)
    df["BC unique"] = (
        df["N° Cde"].astype(str)
        + "|"
        + df["Site"].astype(str)
        + "|"
        + df["Fou"].astype(str)
    )
    df["Fournisseur key"] = df["Fou"].astype(str) + "|" + df["Nom fournisseur"].astype(str)
    df["Sit95 flag"] = df["Sit clean"].astype(str).eq("95")
    df["Lead time brut"] = (df["Dt Rec"] - df["Date de commande"]).dt.days
    df["Lead time valide"] = df["Lead time brut"].where(
        (df["Lead time brut"] >= 0) & (df["Lead time brut"] <= 30)
    )

    site_hors_groupe = sorted(
        df.loc[df["Groupe magasin"].eq("Site hors groupe"), "Site"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    quality = {
        "lignes_initiales": initial_rows,
        "lignes_apres_filtre_date": len(df),
        "date_debut_analyse": start_date,
        "date_fin_analyse": last_date,
        "nb_mois_analyse": nb_months,
        "methode_nb_mois": "Nombre de jours entre date début et date fin / 30,44",
        "qte_cde_manquante_ou_nulle": qte_missing_or_zero,
        "px_revient_manquant": px_missing,
        "date_commande_manquante": date_cde_missing,
        "dt_rec_manquante": dt_rec_missing,
        "sites_hors_groupe": ", ".join(site_hors_groupe) if site_hors_groupe else "Aucun",
    }

    return df, quality


# ═══════════════════════════════════════════════════════════════════════════════
# ANALYSE FOURNISSEURS
# ═══════════════════════════════════════════════════════════════════════════════

def aggregate_group_metrics(df: pd.DataFrame, nb_months: float) -> pd.DataFrame:
    rows = []
    for (fkey, group), g in df.groupby(["Fournisseur key", "Groupe magasin"], dropna=False):
        bc_count = g["BC unique"].nunique()
        qte_cde = g["Qté cde"].fillna(0).sum()
        qte_rec = g["Qté rec"].fillna(0).sum()
        sit95_bc = g.loc[g["Sit95 flag"], "BC unique"].nunique()
        colis_total = g["Colis"].fillna(0).sum()

        rows.append({
            "Fournisseur key": fkey,
            "Groupe magasin": group,
            "BC": bc_count,
            "Colis": colis_total,
            "Colis/mois": colis_total / nb_months,
            "TS%": safe_div(qte_rec, qte_cde) * 100,
            "%Sit95": safe_div(sit95_bc, bc_count) * 100,
            "Actif": bc_count > 0,
        })
    return pd.DataFrame(rows)


def get_metric_for_group(group_metrics: pd.DataFrame, fkey: str, group: str, metric: str):
    sub = group_metrics[
        (group_metrics["Fournisseur key"].eq(fkey))
        & (group_metrics["Groupe magasin"].eq(group))
    ]
    return np.nan if sub.empty else sub.iloc[0][metric]


def is_group_active(group_metrics: pd.DataFrame, fkey: str, group: str) -> bool:
    sub = group_metrics[
        (group_metrics["Fournisseur key"].eq(fkey))
        & (group_metrics["Groupe magasin"].eq(group))
    ]
    return False if sub.empty else bool(sub.iloc[0]["Actif"])


def is_group_defective(group_metrics: pd.DataFrame, fkey: str, group: str) -> bool:
    if not is_group_active(group_metrics, fkey, group):
        return False
    ts = get_metric_for_group(group_metrics, fkey, group, "TS%")
    sit = get_metric_for_group(group_metrics, fkey, group, "%Sit95")
    ts_bad = False if pd.isna(ts) else ts < 60
    sit_bad = False if pd.isna(sit) else sit > 30
    return ts_bad or sit_bad


def is_group_correct(group_metrics: pd.DataFrame, fkey: str, group: str) -> bool:
    if not is_group_active(group_metrics, fkey, group):
        return False
    ts = get_metric_for_group(group_metrics, fkey, group, "TS%")
    sit = get_metric_for_group(group_metrics, fkey, group, "%Sit95")
    if pd.isna(ts) or pd.isna(sit):
        return False
    return ts >= 60 and sit <= 30


def build_supplier_state(
    df: pd.DataFrame,
    quality: dict,
    xd_threshold: float,
    min_orders: int,
    platform_cost_per_package: float,
) -> pd.DataFrame:

    nb_months = quality["nb_mois_analyse"]
    last_date = quality["date_fin_analyse"]
    last_60_start = last_date - timedelta(days=60)
    group_metrics = aggregate_group_metrics(df, nb_months)

    couple = (
        df.groupby(["Fournisseur key", "Site"], dropna=False)
        .agg(
            valeur_cde=("Valeur commande", "sum"),
            nb_bc=("BC unique", "nunique"),
        )
        .reset_index()
    )
    couple["valeur_moyenne_livraison"] = couple["valeur_cde"] / couple["nb_bc"].replace(0, np.nan)

    below_threshold = (
        couple[couple["valeur_moyenne_livraison"] < xd_threshold]
        .groupby("Fournisseur key")
        .agg(nb_couples_sous_seuil=("Site", "nunique"))
        .reset_index()
    )
    total_couples = (
        couple.groupby("Fournisseur key")
        .agg(nb_couples_total=("Site", "nunique"))
        .reset_index()
    )

    rows = []
    for fkey, g in df.groupby("Fournisseur key", dropna=False):
        fou = g["Fou"].iloc[0]
        nom = g["Nom fournisseur"].iloc[0]
        nb_bc = g["BC unique"].nunique()
        nb_refs = g["Code article"].nunique()
        nb_sites = g["Site"].nunique()
        groupes_presents = join_unique(g["Groupe magasin"])
        cdes_mois = nb_bc / nb_months

        qte_cde = g["Qté cde"].fillna(0).sum()
        qte_rec = g["Qté rec"].fillna(0).sum()
        sit95_bc = g.loc[g["Sit95 flag"], "BC unique"].nunique()

        ts_global = safe_div(qte_rec, qte_cde) * 100
        sit95_global = safe_div(sit95_bc, nb_bc) * 100

        colis_total = g["Colis"].fillna(0).sum()
        colis_mois = colis_total / nb_months
        colis_cde_moyen = safe_div(colis_total, nb_bc)

        valeur_totale = g["Valeur commande"].sum()
        valeur_moyenne_bc = safe_div(valeur_totale, nb_bc)

        last_order = g["Date de commande"].max()
        recent_order = bool(last_order >= last_60_start)

        total_couple_match = total_couples[total_couples["Fournisseur key"].eq(fkey)]
        nb_couples_total = int(total_couple_match["nb_couples_total"].iloc[0]) if not total_couple_match.empty else 0

        below_match = below_threshold[below_threshold["Fournisseur key"].eq(fkey)]
        nb_couples_sous_seuil = int(below_match["nb_couples_sous_seuil"].iloc[0]) if not below_match.empty else 0
        pct_couples_sous_seuil = safe_div(nb_couples_sous_seuil, nb_couples_total) * 100

        if nb_bc < min_orders:
            categorie = "Sans données suffisantes"
        elif nb_couples_sous_seuil >= 1:
            categorie = "Candidat XD"
        else:
            categorie = "Hors périmètre XD"

        hypers_active = is_group_active(group_metrics, fkey, "Hypers")
        hypers_def = is_group_defective(group_metrics, fkey, "Hypers")
        hypers_correct = is_group_correct(group_metrics, fkey, "Hypers")
        markets_def = is_group_defective(group_metrics, fkey, "Markets")
        supeco_def = is_group_defective(group_metrics, fkey, "Supeco")
        ms_def = markets_def or supeco_def

        if categorie != "Candidat XD":
            decision = "Non applicable"
            flag = "Non applicable"
            raison = "Fournisseur non candidat XD selon les règles de périmètre."
        else:
            ts_is_zero = not pd.isna(ts_global) and np.isclose(ts_global, 0)
            sit_is_100 = not pd.isna(sit95_global) and np.isclose(sit95_global, 100)

            if ts_is_zero and sit_is_100 and not recent_order:
                decision = "Inactif probable"
                flag = "Inactif"
                raison = "TS global = 0%, Sit95 global = 100%, aucune commande dans les 60 derniers jours."
            elif ts_is_zero and sit_is_100 and recent_order:
                decision = "Litige probable"
                flag = "Litige"
                raison = "TS global = 0%, Sit95 global = 100%, commandes encore actives sur les 60 derniers jours."
            elif hypers_active and hypers_def and ms_def:
                decision = "XD Total"
                flag = "Actif"
                raison = "Hypers défaillants et Markets/Supeco défaillants : bascule globale recommandée."
            elif (not hypers_active or hypers_correct) and ms_def:
                decision = "XD Markets+Supeco"
                flag = "Actif"
                raison = "Hypers absents ou corrects, mais Markets/Supeco défaillants : bascule petits formats."
            else:
                decision = "DL — Surveiller"
                flag = "À surveiller"
                raison = "Candidat XD au seuil valeur, mais performance service acceptable ou bascule non prioritaire."

        ts_hypers = get_metric_for_group(group_metrics, fkey, "Hypers", "TS%")
        ts_markets = get_metric_for_group(group_metrics, fkey, "Markets", "TS%")
        ts_supeco = get_metric_for_group(group_metrics, fkey, "Supeco", "TS%")

        sit_hypers = get_metric_for_group(group_metrics, fkey, "Hypers", "%Sit95")
        sit_markets = get_metric_for_group(group_metrics, fkey, "Markets", "%Sit95")
        sit_supeco = get_metric_for_group(group_metrics, fkey, "Supeco", "%Sit95")

        colis_hypers_mois = get_metric_for_group(group_metrics, fkey, "Hypers", "Colis/mois")
        colis_markets_mois = get_metric_for_group(group_metrics, fkey, "Markets", "Colis/mois")
        colis_supeco_mois = get_metric_for_group(group_metrics, fkey, "Supeco", "Colis/mois")
        colis_hors_groupe_mois = get_metric_for_group(group_metrics, fkey, "Site hors groupe", "Colis/mois")

        colis_hypers_mois = 0 if pd.isna(colis_hypers_mois) else colis_hypers_mois
        colis_markets_mois = 0 if pd.isna(colis_markets_mois) else colis_markets_mois
        colis_supeco_mois = 0 if pd.isna(colis_supeco_mois) else colis_supeco_mois
        colis_hors_groupe_mois = 0 if pd.isna(colis_hors_groupe_mois) else colis_hors_groupe_mois

        if decision == "XD Total":
            colis_xd_mois = colis_hypers_mois + colis_markets_mois + colis_supeco_mois + colis_hors_groupe_mois
        elif decision == "XD Markets+Supeco":
            colis_xd_mois = colis_markets_mois + colis_supeco_mois
        else:
            colis_xd_mois = 0

        cost_xd_month = colis_xd_mois * platform_cost_per_package
        cost_xd_year = cost_xd_month * 12

        rows.append({
            "Fournisseur key": fkey,
            "Code fournisseur": fou,
            "Nom fournisseur": nom,
            "Catégorie périmètre": categorie,
            "Décision XD": decision,
            "Flag statut": flag,
            "Nb références": nb_refs,
            "Nb magasins actifs": nb_sites,
            "Groupes présents": groupes_presents,
            "Nb BC uniques": nb_bc,
            "Cdes/mois": cdes_mois,
            "Jour de commande dominant": mode_day(g["Date de commande"]),
            "Jour de réception dominant": mode_day(g["Dt Rec"]),
            "Lead time médian (j)": g["Lead time valide"].median(),
            "Colis/cde moyen": colis_cde_moyen,
            "Colis/mois": colis_mois,
            "Colis Hypers/mois": colis_hypers_mois,
            "Colis Markets/mois": colis_markets_mois,
            "Colis Supeco/mois": colis_supeco_mois,
            "Colis hors groupe/mois": colis_hors_groupe_mois,
            "Colis XD/mois": colis_xd_mois,
            "Valeur commande totale": valeur_totale,
            "Valeur moyenne BC": valeur_moyenne_bc,
            "Nb couples fournisseur/magasin < seuil": nb_couples_sous_seuil,
            "% couples fournisseur/magasin < seuil": pct_couples_sous_seuil,
            "TS% global": ts_global,
            "TS% Hypers": ts_hypers,
            "TS% Markets": ts_markets,
            "TS% Supeco": ts_supeco,
            "%Sit95 global": sit95_global,
            "%Sit95 Hypers": sit_hypers,
            "%Sit95 Markets": sit_markets,
            "%Sit95 Supeco": sit_supeco,
            "Dernière date de commande": last_order,
            "Commande dans les 60 derniers jours": yes_no(recent_order),
            "Coût traitement XD/mois": cost_xd_month,
            "Coût traitement XD/an": cost_xd_year,
            "Raison de décision": raison,
        })

    suppliers = pd.DataFrame(rows)
    if not suppliers.empty:
        suppliers = suppliers.sort_values("Cdes/mois", ascending=False).reset_index(drop=True)
    return suppliers


# ═══════════════════════════════════════════════════════════════════════════════
# PLAN XD + CHARGE QUAI
# ═══════════════════════════════════════════════════════════════════════════════

def assign_delivery_days(plan: pd.DataFrame) -> pd.DataFrame:
    plan = plan.copy()
    plan["Jour livraison XD"] = "N/A"

    hebdo_idx = plan[plan["Cadence XD"].eq("Hebdo")].sort_values("Colis XD/mois", ascending=False).index
    charges_hebdo = {"Lundi": 0.0, "Mercredi": 0.0}
    for idx in hebdo_idx:
        day = min(charges_hebdo, key=charges_hebdo.get)
        plan.loc[idx, "Jour livraison XD"] = day
        charges_hebdo[day] += plan.loc[idx, "Colis XD/mois"] / 4

    bim_idx = plan[plan["Cadence XD"].eq("Bi-mensuel")].sort_values("Colis XD/mois", ascending=False).index
    charges_bim = {"Jeudi": 0.0, "Vendredi": 0.0}
    for idx in bim_idx:
        day = min(charges_bim, key=charges_bim.get)
        plan.loc[idx, "Jour livraison XD"] = day
        charges_bim[day] += plan.loc[idx, "Colis XD/mois"] / 2

    mens_idx = plan[plan["Cadence XD"].eq("Mensuel")].sort_values("Colis XD/mois", ascending=False).index
    charges_mens = {"Lundi": 0.0, "Mercredi": 0.0, "Jeudi": 0.0, "Vendredi": 0.0}
    for idx in mens_idx:
        day = min(charges_mens, key=charges_mens.get)
        plan.loc[idx, "Jour livraison XD"] = day
        charges_mens[day] += plan.loc[idx, "Colis XD/mois"]

    return plan


def build_smoothing_plan(suppliers: pd.DataFrame, platform_cost_per_package: float) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    plan = suppliers[suppliers["Décision XD"].isin(["XD Total", "XD Markets+Supeco"])].copy()

    if plan.empty:
        charge = pd.DataFrame(columns=[
            "Jour", "Colis/semaine simulés", "Nombre de fournisseurs",
            "Nombre de réceptions", "Charge moyenne par réception",
            "Coût traitement XD/semaine", "Dépassement seuil 800 colis/jour", "Alerte"
        ])
        stats = {"charge_max": 0, "charge_min": 0, "ratio_pic_creux": 0, "flag_ratio": "N/A", "total_cost_month": 0, "total_cost_year": 0}
        return plan, charge, stats

    plan["Groupes basculés XD"] = np.where(plan["Décision XD"].eq("XD Total"), "Tous groupes actifs", "Markets / Supeco")
    plan["Groupes maintenus DL"] = np.where(plan["Décision XD"].eq("XD Markets+Supeco"), "Hypers", "Aucun")
    plan["Cadence actuelle"] = plan["Cdes/mois"].apply(compute_current_cadence)
    plan["Jour de commande actuel"] = plan["Jour de commande dominant"]
    plan["Jour de livraison actuel"] = plan["Jour de réception dominant"]
    plan["BC/mois actuel"] = plan["Cdes/mois"]
    plan["Colis/mois actuel"] = plan["Colis/mois"]
    plan["Colis/cde actuel"] = plan["Colis/cde moyen"]

    cadence_data = plan["Colis XD/mois"].apply(compute_xd_cadence)
    plan["Cadence XD"] = cadence_data.apply(lambda x: x[0])
    plan["Cycles XD/mois"] = cadence_data.apply(lambda x: x[1])
    plan["Colis/livraison XD"] = cadence_data.apply(lambda x: x[2])
    plan["Alerte colis"] = cadence_data.apply(lambda x: x[3])

    plan = assign_delivery_days(plan)

    plan["Jour cut-off"] = plan.apply(
        lambda r: compute_cutoff_day(r["Jour livraison XD"], r["Lead time médian (j)"]),
        axis=1,
    )
    plan["BC XD/mois"] = plan["Cycles XD/mois"]
    plan["Réduction BC/mois"] = plan["BC/mois actuel"] - plan["BC XD/mois"]
    plan["% réduction BC/mois"] = plan["Réduction BC/mois"] / plan["BC/mois actuel"].replace(0, np.nan) * 100
    plan["Coût traitement plateforme / colis"] = platform_cost_per_package
    plan["Coût traitement XD/mois"] = plan["Colis XD/mois"] * platform_cost_per_package
    plan["Coût traitement XD/an"] = plan["Coût traitement XD/mois"] * 12
    plan["Coût traitement XD/livraison"] = plan["Colis/livraison XD"] * platform_cost_per_package
    plan["Coût traitement XD par cycle"] = plan["Coût traitement XD/livraison"]
    plan["Colis Hypers maintenus DL"] = np.where(plan["Décision XD"].eq("XD Markets+Supeco"), plan["Colis Hypers/mois"], 0)
    plan["Coût théorique Hypers exclu XD"] = plan["Colis Hypers maintenus DL"] * platform_cost_per_package

    charge_rows = []
    for day in ["Lundi", "Mercredi", "Jeudi", "Vendredi"]:
        sub = plan[plan["Jour livraison XD"].eq(day)].copy()
        weekly_packages = 0.0
        receptions = 0.0

        for _, r in sub.iterrows():
            cycles = r["Cycles XD/mois"]
            if cycles > 0:
                weekly_packages += r["Colis XD/mois"] / 4
                receptions += max(cycles / 4, 0.25)

        nb_suppliers = sub["Code fournisseur"].nunique()
        avg_per_reception = safe_div(weekly_packages, receptions)
        cost_week = weekly_packages * platform_cost_per_package
        over_800 = weekly_packages > 800
        alert = "🟢" if weekly_packages < 500 else ("🟠" if weekly_packages < 800 else "🔴")

        charge_rows.append({
            "Jour": day,
            "Colis/semaine simulés": weekly_packages,
            "Nombre de fournisseurs": nb_suppliers,
            "Nombre de réceptions": receptions,
            "Charge moyenne par réception": avg_per_reception,
            "Coût traitement XD/semaine": cost_week,
            "Dépassement seuil 800 colis/jour": yes_no(over_800),
            "Alerte": alert,
        })

    charge = pd.DataFrame(charge_rows)
    non_zero = charge.loc[charge["Colis/semaine simulés"] > 0, "Colis/semaine simulés"]

    if non_zero.empty:
        charge_max, charge_min, ratio = 0, 0, 0
    else:
        charge_max = non_zero.max()
        charge_min = non_zero.min()
        ratio = safe_div(charge_max, charge_min)

    stats = {
        "charge_max": charge_max,
        "charge_min": charge_min,
        "ratio_pic_creux": ratio,
        "flag_ratio": "OK" if ratio <= 3 else "À lisser",
        "total_cost_month": plan["Coût traitement XD/mois"].sum(),
        "total_cost_year": plan["Coût traitement XD/an"].sum(),
    }

    export_cols = [
        "Code fournisseur", "Nom fournisseur", "Décision XD", "Groupes basculés XD", "Groupes maintenus DL",
        "Cadence actuelle", "Jour de commande actuel", "Jour de livraison actuel", "Lead time médian (j)",
        "Colis/cde actuel", "BC/mois actuel", "Colis/mois actuel",
        "Colis XD/mois", "Cadence XD", "Cycles XD/mois", "Colis/livraison XD", "Alerte colis",
        "Jour livraison XD", "Jour cut-off", "BC XD/mois", "Réduction BC/mois", "% réduction BC/mois",
        "Coût traitement plateforme / colis", "Coût traitement XD/mois", "Coût traitement XD/an",
        "Coût traitement XD/livraison", "Coût traitement XD par cycle",
        "Colis Hypers maintenus DL", "Coût théorique Hypers exclu XD",
    ]

    return plan[export_cols].reset_index(drop=True), charge, stats


# ═══════════════════════════════════════════════════════════════════════════════
# À STATUER + BDD ARTICLES
# ═══════════════════════════════════════════════════════════════════════════════

def build_to_decide(suppliers: pd.DataFrame) -> pd.DataFrame:
    subset = suppliers[
        suppliers["Décision XD"].isin(["DL — Surveiller", "Litige probable", "Inactif probable"])
        | suppliers["Catégorie périmètre"].eq("Hors périmètre XD")
        | suppliers["Catégorie périmètre"].eq("Sans données suffisantes")
    ].copy()

    rows = []
    for _, r in subset.iterrows():
        if r["Catégorie périmètre"] == "Hors périmètre XD":
            decision = "Hors périmètre XD"
            reason = "Tous les couples fournisseur/magasin sont au-dessus ou égaux au seuil XD."
            action = "Maintien DL ; revoir uniquement si baisse de valeur commande ou dégradation TS."
        elif r["Catégorie périmètre"] == "Sans données suffisantes":
            decision = "Sans données suffisantes"
            reason = "Moins de 5 BC uniques sur la période."
            action = "Compléter l’historique avant décision ; surveiller les prochaines commandes."
        elif r["Décision XD"] == "DL — Surveiller":
            decision = "DL — Surveiller"
            reason = r["Raison de décision"]
            action = "Revoir dans 3 mois avec suivi TS%, Sit95 et valeur moyenne commande."
        elif r["Décision XD"] == "Litige probable":
            decision = "Litige probable"
            reason = r["Raison de décision"]
            action = "Escalader aux Achats / clarifier litige fournisseur / bloquer commandes si nécessaire."
        elif r["Décision XD"] == "Inactif probable":
            decision = "Inactif probable"
            reason = r["Raison de décision"]
            action = "Vérifier référencement / suspendre ou nettoyer base fournisseur."
        else:
            decision = r["Décision XD"]
            reason = r["Raison de décision"]
            action = "À analyser."

        rows.append({
            "Code fournisseur": r["Code fournisseur"],
            "Nom fournisseur": r["Nom fournisseur"],
            "Catégorie périmètre": r["Catégorie périmètre"],
            "Décision XD": decision,
            "Raison principale": reason,
            "TS% global": r["TS% global"],
            "%Sit95 global": r["%Sit95 global"],
            "Cdes/mois": r["Cdes/mois"],
            "Colis/mois": r["Colis/mois"],
            "Coût traitement XD/mois": 0,
            "Coût traitement XD/an": 0,
            "Dernière date de commande": r["Dernière date de commande"],
            "Action recommandée": action,
        })

    return pd.DataFrame(rows)


def build_article_db(df: pd.DataFrame, suppliers: pd.DataFrame, platform_cost_per_package: float) -> pd.DataFrame:
    agg = (
        df.groupby(["Fournisseur key", "Fou", "Nom fournisseur", "Code article"], dropna=False)
        .agg(
            nb_magasins=("Site", "nunique"),
            groupes=("Groupe magasin", join_unique),
            qte_commandee=("Qté cde", "sum"),
            qte_recue=("Qté rec", "sum"),
            valeur_commande=("Valeur commande", "sum"),
            colis_total=("Colis", "sum"),
            nb_bc=("BC unique", "nunique"),
            derniere_commande=("Date de commande", "max"),
        )
        .reset_index()
    )
    agg["TS% article"] = agg["qte_recue"] / agg["qte_commandee"].replace(0, np.nan) * 100

    article_group = (
        df.groupby(["Fournisseur key", "Code article", "Groupe magasin"], dropna=False)
        .agg(colis=("Colis", "sum"))
        .reset_index()
    )

    ms_colis = (
        article_group[article_group["Groupe magasin"].isin(["Markets", "Supeco"])]
        .groupby(["Fournisseur key", "Code article"])
        .agg(colis_ms=("colis", "sum"))
        .reset_index()
    )
    all_xd_colis = (
        article_group
        .groupby(["Fournisseur key", "Code article"])
        .agg(colis_all=("colis", "sum"))
        .reset_index()
    )

    agg = agg.merge(ms_colis, on=["Fournisseur key", "Code article"], how="left")
    agg = agg.merge(all_xd_colis, on=["Fournisseur key", "Code article"], how="left")
    agg["colis_ms"] = agg["colis_ms"].fillna(0)
    agg["colis_all"] = agg["colis_all"].fillna(0)

    sup_info = suppliers[["Fournisseur key", "Catégorie périmètre", "Décision XD"]].drop_duplicates()
    agg = agg.merge(sup_info, on="Fournisseur key", how="left")

    def groups_switched(decision):
        if decision == "XD Total":
            return "Tous groupes actifs"
        if decision == "XD Markets+Supeco":
            return "Markets / Supeco"
        return "Aucun"

    def article_cost(row):
        if row["Décision XD"] == "XD Total":
            return row["colis_all"] * platform_cost_per_package
        if row["Décision XD"] == "XD Markets+Supeco":
            return row["colis_ms"] * platform_cost_per_package
        return 0

    agg["Groupes basculés XD"] = agg["Décision XD"].apply(groups_switched)
    agg["Coût traitement XD article théorique"] = agg.apply(article_cost, axis=1)
    agg["Commentaire"] = np.where(
        agg["Décision XD"].isin(["XD Total", "XD Markets+Supeco"]),
        "Article rattaché à un fournisseur basculé XD.",
        "Article rattaché à un fournisseur non basculé XD.",
    )

    final = agg.rename(columns={
        "Fou": "Code fournisseur",
        "nb_magasins": "Nb magasins où l’article est commandé",
        "groupes": "Groupes magasins présents",
        "qte_commandee": "Qté commandée totale",
        "qte_recue": "Qté reçue totale",
        "valeur_commande": "Valeur commande totale article",
        "colis_total": "Colis total article",
        "nb_bc": "Nb BC article",
        "derniere_commande": "Dernière date de commande article",
        "Catégorie périmètre": "Catégorie périmètre fournisseur",
        "Décision XD": "Décision XD fournisseur",
    })

    cols = [
        "Code fournisseur", "Nom fournisseur", "Code article",
        "Nb magasins où l’article est commandé", "Groupes magasins présents",
        "Qté commandée totale", "Qté reçue totale", "TS% article",
        "Valeur commande totale article", "Colis total article", "Nb BC article",
        "Dernière date de commande article", "Catégorie périmètre fournisseur",
        "Décision XD fournisseur", "Groupes basculés XD",
        "Coût traitement XD article théorique", "Commentaire",
    ]

    final = final[cols].copy()
    if not final.empty:
        final = final.sort_values(["Code fournisseur", "Valeur commande totale article"], ascending=[True, False])
    return final


# ═══════════════════════════════════════════════════════════════════════════════
# CONTRÔLES + EXPORT EXCEL
# ═══════════════════════════════════════════════════════════════════════════════

def build_control_sheet(
    suppliers: pd.DataFrame,
    plan: pd.DataFrame,
    charge_stats: dict,
    quality: dict,
    platform_cost_per_package: float,
) -> pd.DataFrame:

    total_suppliers = suppliers["Fournisseur key"].nunique()
    sans_data = int((suppliers["Catégorie périmètre"] == "Sans données suffisantes").sum())
    candidats = int((suppliers["Catégorie périmètre"] == "Candidat XD").sum())
    hors = int((suppliers["Catégorie périmètre"] == "Hors périmètre XD").sum())

    sum_categories = sans_data + candidats + hors
    diff_categories = total_suppliers - sum_categories

    candidate_decisions = suppliers.loc[suppliers["Catégorie périmètre"].eq("Candidat XD"), "Décision XD"]
    xd_total = int((candidate_decisions == "XD Total").sum())
    xd_ms = int((candidate_decisions == "XD Markets+Supeco").sum())
    dl_surv = int((candidate_decisions == "DL — Surveiller").sum())
    litige = int((candidate_decisions == "Litige probable").sum())
    inactif = int((candidate_decisions == "Inactif probable").sum())

    sum_decisions = xd_total + xd_ms + dl_surv + litige + inactif
    diff_decisions = candidats - sum_decisions

    plan_count = len(plan)
    expected_plan = xd_total + xd_ms
    total_colis_xd_mois = plan["Colis XD/mois"].sum() if not plan.empty else 0
    total_cost_month = total_colis_xd_mois * platform_cost_per_package
    total_cost_year = total_cost_month * 12
    financial_check = np.isclose(total_cost_month, charge_stats.get("total_cost_month", 0))

    rows = [
        ["Synthèse fournisseurs", "Nombre total de fournisseurs uniques", total_suppliers],
        ["Synthèse fournisseurs", "Sans données suffisantes", sans_data],
        ["Synthèse fournisseurs", "Candidats XD", candidats],
        ["Synthèse fournisseurs", "Hors périmètre XD", hors],
        ["Synthèse fournisseurs", "Somme catégories", sum_categories],
        ["Synthèse fournisseurs", "Écart catégories", diff_categories],
        ["Synthèse fournisseurs", "Flag contrôle fournisseurs", "OK" if diff_categories == 0 else "ÉCART À CORRIGER"],

        ["Décisions candidats XD", "XD Total", xd_total],
        ["Décisions candidats XD", "XD Markets+Supeco", xd_ms],
        ["Décisions candidats XD", "DL — Surveiller", dl_surv],
        ["Décisions candidats XD", "Litige probable", litige],
        ["Décisions candidats XD", "Inactif probable", inactif],
        ["Décisions candidats XD", "Total décisions candidats", sum_decisions],
        ["Décisions candidats XD", "Écart décisions", diff_decisions],
        ["Décisions candidats XD", "Flag contrôle décisions", "OK" if diff_decisions == 0 else "ÉCART À CORRIGER"],

        ["Plan de lissage", "Fournisseurs dans plan de lissage", plan_count],
        ["Plan de lissage", "XD Total + XD Markets+Supeco attendus", expected_plan],
        ["Plan de lissage", "Flag contrôle plan", "OK" if plan_count == expected_plan else "ÉCART À CORRIGER"],
        ["Plan de lissage", "Ratio pic/creux charge quai", charge_stats.get("ratio_pic_creux", 0)],
        ["Plan de lissage", "Flag ratio charge quai", charge_stats.get("flag_ratio", "N/A")],

        ["Contrôle financier XD", "Total colis XD/mois", total_colis_xd_mois],
        ["Contrôle financier XD", "Coût unitaire traitement plateforme", platform_cost_per_package],
        ["Contrôle financier XD", "Coût total traitement XD/mois", total_cost_month],
        ["Contrôle financier XD", "Coût total traitement XD/an", total_cost_year],
        ["Contrôle financier XD", "Vérification coût = colis × coût unitaire", "OK" if financial_check else "ÉCART À CORRIGER"],

        ["Contrôle qualité données", "Lignes initiales", quality["lignes_initiales"]],
        ["Contrôle qualité données", "Lignes après filtre date", quality["lignes_apres_filtre_date"]],
        ["Contrôle qualité données", "Date début analyse", quality["date_debut_analyse"]],
        ["Contrôle qualité données", "Date fin analyse", quality["date_fin_analyse"]],
        ["Contrôle qualité données", "Nombre de mois analyse", quality["nb_mois_analyse"]],
        ["Contrôle qualité données", "Méthode nombre de mois", quality["methode_nb_mois"]],
        ["Contrôle qualité données", "Qté cde manquante ou nulle", quality["qte_cde_manquante_ou_nulle"]],
        ["Contrôle qualité données", "Px revient manquant", quality["px_revient_manquant"]],
        ["Contrôle qualité données", "Date commande manquante", quality["date_commande_manquante"]],
        ["Contrôle qualité données", "Dt Rec manquante", quality["dt_rec_manquante"]],
        ["Contrôle qualité données", "Sites hors groupe", quality["sites_hors_groupe"]],
    ]

    return pd.DataFrame(rows, columns=["Section", "Indicateur", "Valeur"])


def write_excel(
    control: pd.DataFrame,
    suppliers: pd.DataFrame,
    plan: pd.DataFrame,
    charge: pd.DataFrame,
    to_decide: pd.DataFrame,
    article_db: pd.DataFrame,
) -> bytes:

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        control.to_excel(writer, sheet_name="1_Controle_exhaustivite", index=False)

        suppliers_export = suppliers.drop(columns=["Fournisseur key"], errors="ignore")
        suppliers_export.to_excel(writer, sheet_name="2_Etat_DL_complet", index=False)

        plan.to_excel(writer, sheet_name="3_Plan_lissage_XD", index=False, startrow=0)
        start_charge = len(plan) + 4
        charge.to_excel(writer, sheet_name="3_Plan_lissage_XD", index=False, startrow=start_charge)

        to_decide.to_excel(writer, sheet_name="4_A_statuer", index=False)
        article_db.to_excel(writer, sheet_name="5_BDD_articles", index=False)

        workbook = writer.book
        fmt_header = workbook.add_format({
            "bold": True,
            "bg_color": "#1F4E78",
            "font_color": "white",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        fmt_header_orange = workbook.add_format({
            "bold": True,
            "bg_color": "#F4B183",
            "font_color": "black",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        fmt_money = workbook.add_format({"num_format": '#,##0 "FCFA"'})
        fmt_num = workbook.add_format({"num_format": "#,##0.0"})
        fmt_date = workbook.add_format({"num_format": "dd/mm/yyyy"})
        fmt_red = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
        fmt_orange = workbook.add_format({"bg_color": "#FCE4D6", "font_color": "#9C6500"})
        fmt_green = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})

        all_sheets = {
            "1_Controle_exhaustivite": control,
            "2_Etat_DL_complet": suppliers_export,
            "4_A_statuer": to_decide,
            "5_BDD_articles": article_db,
        }

        for sheet_name, df_sheet in all_sheets.items():
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)

            if not df_sheet.empty:
                ws.autofilter(0, 0, len(df_sheet), max(len(df_sheet.columns) - 1, 0))

            for col_num, col_name in enumerate(df_sheet.columns):
                ws.write(0, col_num, col_name, fmt_header)
                width = min(max(len(str(col_name)) + 2, 12), 42)
                lower = str(col_name).lower()

                if "coût" in lower or "valeur" in lower or "prix" in lower:
                    ws.set_column(col_num, col_num, width, fmt_money)
                elif "date" in lower:
                    ws.set_column(col_num, col_num, width, fmt_date)
                elif "ts%" in lower or "%sit" in lower or "%" in lower or "colis" in lower or "bc" in lower:
                    ws.set_column(col_num, col_num, width, fmt_num)
                else:
                    ws.set_column(col_num, col_num, width)

        ws_plan = writer.sheets["3_Plan_lissage_XD"]
        ws_plan.freeze_panes(1, 0)

        if not plan.empty:
            ws_plan.autofilter(0, 0, len(plan), max(len(plan.columns) - 1, 0))

        for col_num, col_name in enumerate(plan.columns):
            header_fmt = fmt_header_orange if "Coût" in str(col_name) else fmt_header
            ws_plan.write(0, col_num, col_name, header_fmt)
            width = min(max(len(str(col_name)) + 2, 12), 42)
            lower = str(col_name).lower()

            if "coût" in lower:
                ws_plan.set_column(col_num, col_num, width, fmt_money)
            elif "date" in lower:
                ws_plan.set_column(col_num, col_num, width, fmt_date)
            elif "colis" in lower or "bc" in lower or "%" in lower or "réduction" in lower:
                ws_plan.set_column(col_num, col_num, width, fmt_num)
            else:
                ws_plan.set_column(col_num, col_num, width)

        for col_num, col_name in enumerate(charge.columns):
            ws_plan.write(start_charge, col_num, col_name, fmt_header)
            width = min(max(len(str(col_name)) + 2, 12), 42)
            if "coût" in str(col_name).lower():
                ws_plan.set_column(col_num, col_num, width, fmt_money)
            else:
                ws_plan.set_column(col_num, col_num, width)

        if not plan.empty and "Alerte colis" in plan.columns:
            alert_col = list(plan.columns).index("Alerte colis")
            ws_plan.conditional_format(1, alert_col, len(plan), alert_col, {
                "type": "text", "criteria": "containing", "value": "🔴", "format": fmt_red
            })
            ws_plan.conditional_format(1, alert_col, len(plan), alert_col, {
                "type": "text", "criteria": "containing", "value": "🟠", "format": fmt_orange
            })
            ws_plan.conditional_format(1, alert_col, len(plan), alert_col, {
                "type": "text", "criteria": "containing", "value": "🟢", "format": fmt_green
            })

    output.seek(0)
    return output.read()


# ═══════════════════════════════════════════════════════════════════════════════
# PIPELINE
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def run_analysis_cached(
    file_bytes: bytes,
    filename: str,
    start_date,
    xd_threshold: float,
    min_orders: int,
    platform_cost_per_package: float,
) -> dict:

    raw_df = read_input_file(file_bytes, filename)
    mapping = detect_columns(raw_df)
    df, quality = prepare_data(raw_df, pd.Timestamp(start_date), mapping)

    suppliers = build_supplier_state(
        df=df,
        quality=quality,
        xd_threshold=xd_threshold,
        min_orders=min_orders,
        platform_cost_per_package=platform_cost_per_package,
    )

    plan, charge, charge_stats = build_smoothing_plan(
        suppliers=suppliers,
        platform_cost_per_package=platform_cost_per_package,
    )

    to_decide = build_to_decide(suppliers)

    article_db = build_article_db(
        df=df,
        suppliers=suppliers,
        platform_cost_per_package=platform_cost_per_package,
    )

    control = build_control_sheet(
        suppliers=suppliers,
        plan=plan,
        charge_stats=charge_stats,
        quality=quality,
        platform_cost_per_package=platform_cost_per_package,
    )

    excel_bytes = write_excel(
        control=control,
        suppliers=suppliers,
        plan=plan,
        charge=charge,
        to_decide=to_decide,
        article_db=article_db,
    )

    return {
        "raw_df": raw_df,
        "mapping": mapping,
        "df": df,
        "quality": quality,
        "suppliers": suppliers,
        "plan": plan,
        "charge": charge,
        "charge_stats": charge_stats,
        "to_decide": to_decide,
        "article_db": article_db,
        "control": control,
        "excel_bytes": excel_bytes,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════

def safe_page_link(page: str, label: str):
    try:
        st.page_link(page, label=label)
    except Exception:
        pass


with st.sidebar:
    st.markdown("""
<div style='margin-bottom:18px'>
  <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>🛍️ SmartBuyer</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · Équipe Achats</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Navigation</div>", unsafe_allow_html=True)
    safe_page_link("app.py", "🏠  Accueil")
    safe_page_link("pages/01_📊_Analyse_Scoring_ABC.py", "📊  Scoring ABC")
    safe_page_link("pages/02_📈_Ventes_PBI.py", "📈  Ventes PBI")
    safe_page_link("pages/03_📦_Detention_Top_CA.py", "📦  Détention Top CA")
    safe_page_link("pages/04_💸_Performance_Promo.py", "💸  Performance Promo")
    safe_page_link("pages/05_🏪_Suivi_Implantation.py", "🏪  Suivi Implantation")
    safe_page_link("pages/06_💸_Marges_Negatives.py", "💸  Marges Négatives")
    st.markdown("<div style='font-size:13px;font-weight:600;color:#007AFF;margin-top:6px'>🏪  Bascule XD</div>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Import fichier</div>", unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Cahier d’entrée commandes",
        type=["xlsx", "xlsb", "xls", "csv"],
        key="xd_file",
        help="Le fichier doit contenir un seul onglet. Le premier onglet sera lu automatiquement.",
    )

    st.caption("Format attendu : fichier unique avec un seul onglet de données.")
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Paramètres</div>", unsafe_allow_html=True)

    start_date = st.date_input("Date début analyse", value=DEFAULT_START_DATE.date())
    xd_threshold = st.number_input(
        "Seuil XD valeur moyenne",
        min_value=0,
        value=DEFAULT_XD_THRESHOLD,
        step=10_000,
        help="Règle stricte : candidat si valeur moyenne fournisseur/magasin < seuil.",
    )
    min_orders = st.number_input("Minimum BC données suffisantes", min_value=1, value=DEFAULT_MIN_ORDERS, step=1)
    platform_cost = st.number_input("Coût plateforme / colis", min_value=0, value=DEFAULT_PLATFORM_COST_PER_PACKAGE, step=10)

    launch = False
    if uploaded_file is not None:
        st.markdown("---")
        launch = st.button("🚀 Lancer l’analyse", type="primary", use_container_width=True)

    st.markdown("---")
    st.caption(f"Python : {sys.version.split()[0]}")
    if package_installed("pyxlsb"):
        st.caption("✅ pyxlsb installé : lecture .xlsb OK")
    else:
        st.caption("⚠️ pyxlsb absent : convertir .xlsb en .xlsx ou ajouter pyxlsb")


# ═══════════════════════════════════════════════════════════════════════════════
# HEADER PAGE
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("<div class='page-title'>🏪 Commando XD — Bascule DL vers Cross-Docking</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Analyse fournisseurs · seuil petites commandes · taux de service · plan de lissage plateforme · coût XD à 90 FCFA / colis</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# LANDING PAGE
# ═══════════════════════════════════════════════════════════════════════════════

if uploaded_file is None:
    st.markdown("---")
    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>ℹ️ À quoi sert ce module ?</strong><br>
  Ce module analyse le flux <strong>Direct Livraison fournisseur → magasin</strong> pour identifier les fournisseurs à basculer vers une
  <strong>plateforme de Cross-Docking XD</strong>. L’objectif est de réduire les petites commandes non rentables,
  massifier les flux, améliorer le taux de service et lisser les réceptions plateforme.
</div>
""", unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("<div class='section-label'>Logique métier appliquée</div>", unsafe_allow_html=True)
        st.markdown(f"""
<div class='card'>
  <div style='font-size:14px;font-weight:700;color:#1C1C1E;margin-bottom:8px'>🎯 Seuil XD</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:1.5'>
    Un couple <strong>Fournisseur / Magasin</strong> devient candidat si sa valeur moyenne de livraison est
    <strong>strictement inférieure à {DEFAULT_XD_THRESHOLD:,.0f} FCFA</strong>.<br>
    La règle est donc : <code>valeur moyenne &lt; seuil</code>. Une valeur exactement égale au seuil n’est pas candidate.
  </div>
</div>
<div class='card'>
  <div style='font-size:14px;font-weight:700;color:#1C1C1E;margin-bottom:8px'>📦 Coût plateforme</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:1.5'>
    Le budget XD est calculé uniquement sur les colis réellement basculés :<br>
    <code>Coût XD/mois = Colis XD/mois × {DEFAULT_PLATFORM_COST_PER_PACKAGE} FCFA</code>
  </div>
</div>
<div class='card'>
  <div style='font-size:14px;font-weight:700;color:#1C1C1E;margin-bottom:8px'>🧾 Livrables Excel</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:1.5'>
    1. Contrôle exhaustivité<br>
    2. État DL complet<br>
    3. Plan de lissage XD<br>
    4. À statuer<br>
    5. BDD articles
  </div>
</div>
""".replace(",", " "), unsafe_allow_html=True)

    with c2:
        st.markdown("<div class='section-label'>Groupes magasins</div>", unsafe_allow_html=True)
        st.markdown("""
<div class='format-card format-hyper'>
  <span class='badge badge-hyper'>Hypers</span>
  <div class='small-muted' style='margin-top:6px'>Sites : 10202 · 10203 · 10301</div>
</div>
<div class='format-card format-market'>
  <span class='badge badge-market'>Markets</span>
  <div class='small-muted' style='margin-top:6px'>Sites : 10604 · 10206 · 10208 · 10209 · 10705</div>
</div>
<div class='format-card format-supeco'>
  <span class='badge badge-supeco'>Supeco</span>
  <div class='small-muted' style='margin-top:6px'>Sites : 10601 · 10602 · 10603 · 10605</div>
</div>
""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div class='section-label'>Fonctionnement</div>", unsafe_allow_html=True)
        st.markdown("""
<div class='alert-card alert-green'>
  <strong>1.</strong> Charge le cahier d’entrée dans la sidebar.<br>
  <strong>2.</strong> Vérifie les paramètres : seuil XD, coût colis, date de début.<br>
  <strong>3.</strong> Clique sur <strong>Lancer l’analyse</strong>.<br>
  <strong>4.</strong> Télécharge le fichier Excel final.
</div>
""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Colonnes attendues dans le fichier</div>", unsafe_allow_html=True)

    cols_expected = [
        ("Fou", "Code fournisseur"),
        ("Nom fourn,", "Nom fournisseur"),
        ("Site", "Code magasin"),
        ("Code", "Code article"),
        ("N° Cde", "Numéro de commande"),
        ("Date de commande", "Date de création de commande"),
        ("Dt Rec", "Date de réception"),
        ("Qté cde", "Quantité commandée"),
        ("Qté rec", "Quantité reçue — alias accepté : Qté reçue / Qte rec"),
        ("Px revient", "Prix de revient"),
        ("Colis", "Nombre de colis / PCB réel"),
        ("Sit", "Statut commande, dont 95 pour totalement non livrée"),
    ]

    col_left, col_right = st.columns(2)
    for i, (name, desc) in enumerate(cols_expected):
        target = col_left if i % 2 == 0 else col_right
        with target:
            st.markdown(f"""
<div class='col-required'>
  <div style='font-size:16px'>▪️</div>
  <div>
    <div class='col-name'>{name}</div>
    <div class='col-desc'>{desc}</div>
  </div>
</div>
""", unsafe_allow_html=True)

    st.info("⬅️ Charge le fichier dans la sidebar pour lancer l’analyse.")
    st.stop()


if not launch:
    st.markdown("""
<div class='alert-card alert-blue'>
  Fichier chargé. Clique maintenant sur <strong>🚀 Lancer l’analyse</strong> dans la sidebar pour démarrer le traitement.
</div>
""", unsafe_allow_html=True)
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════════
# EXÉCUTION
# ═══════════════════════════════════════════════════════════════════════════════

try:
    file_bytes = uploaded_file.getvalue()
    filename = uploaded_file.name

    with st.spinner("Lecture du fichier et calcul Commando XD…"):
        result = run_analysis_cached(
            file_bytes=file_bytes,
            filename=filename,
            start_date=start_date,
            xd_threshold=float(xd_threshold),
            min_orders=int(min_orders),
            platform_cost_per_package=float(platform_cost),
        )

    raw_df = result["raw_df"]
    mapping = result["mapping"]
    suppliers = result["suppliers"]
    plan = result["plan"]
    charge = result["charge"]
    control = result["control"]
    stats = result["charge_stats"]
    quality = result["quality"]
    excel_bytes = result["excel_bytes"]

    total_suppliers = suppliers["Fournisseur key"].nunique()
    candidats = int((suppliers["Catégorie périmètre"] == "Candidat XD").sum())
    xd_total = int((suppliers["Décision XD"] == "XD Total").sum())
    xd_ms = int((suppliers["Décision XD"] == "XD Markets+Supeco").sum())
    dl_surv = int((suppliers["Décision XD"] == "DL — Surveiller").sum())
    litige = int((suppliers["Décision XD"] == "Litige probable").sum())
    inactif = int((suppliers["Décision XD"] == "Inactif probable").sum())
    hors = int((suppliers["Catégorie périmètre"] == "Hors périmètre XD").sum())
    sans_data = int((suppliers["Catégorie périmètre"] == "Sans données suffisantes").sum())

    total_colis_xd = plan["Colis XD/mois"].sum() if not plan.empty else 0
    total_cost_month = total_colis_xd * float(platform_cost)
    total_cost_year = total_cost_month * 12

    st.markdown(f"<div class='section-label'>{quality['lignes_apres_filtre_date']:,} ligne(s) analysée(s) · période : {quality['date_debut_analyse'].strftime('%d/%m/%Y')} → {quality['date_fin_analyse'].strftime('%d/%m/%Y')}</div>".replace(",", " "), unsafe_allow_html=True)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Fournisseurs", f"{total_suppliers:,.0f}".replace(",", " "))
    k2.metric("Candidats XD", f"{candidats:,.0f}".replace(",", " "))
    k3.metric("XD Total", f"{xd_total:,.0f}".replace(",", " "))
    k4.metric("XD M+S", f"{xd_ms:,.0f}".replace(",", " "))

    k5, k6, k7, k8 = st.columns(4)
    k5.metric("DL Surveiller", f"{dl_surv:,.0f}".replace(",", " "))
    k6.metric("Litiges", f"{litige:,.0f}".replace(",", " "))
    k7.metric("Inactifs", f"{inactif:,.0f}".replace(",", " "))
    k8.metric("Hors périmètre", f"{hors:,.0f}".replace(",", " "))

    f1, f2, f3, f4 = st.columns(4)
    f1.metric("Sans données", f"{sans_data:,.0f}".replace(",", " "))
    f2.metric("Colis XD/mois", f"{total_colis_xd:,.0f}".replace(",", " "))
    f3.metric("Coût XD/mois", fmt_xof(total_cost_month))
    f4.metric("Coût XD/an", fmt_xof(total_cost_year))

    st.markdown("---")

    # Alertes principales
    if quality["sites_hors_groupe"] != "Aucun":
        st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ Sites hors groupe détectés</strong><br>
  {quality["sites_hors_groupe"]}<br>
  <span style='font-size:12px;opacity:.85'>Ces sites restent inclus dans l’analyse fournisseur, mais ne sont pas rattachés à Hypers / Markets / Supeco.</span>
</div>
""", unsafe_allow_html=True)

    if stats.get("flag_ratio") == "À lisser":
        st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ Charge quai à lisser</strong><br>
  Ratio pic/creux : <strong>{stats.get("ratio_pic_creux", 0):.2f}x</strong> — objectif ≤ 3x.
</div>
""", unsafe_allow_html=True)
    else:
        st.markdown(f"""
<div class='alert-card alert-green'>
  <strong>✅ Contrôle charge quai</strong><br>
  Ratio pic/creux : <strong>{stats.get("ratio_pic_creux", 0):.2f}x</strong>.
</div>
""", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "✅ Contrôles",
        "📦 État DL complet",
        "🚚 Plan XD",
        "🏭 Charge quai",
        "📥 Export",
    ])

    with tab1:
        st.markdown("<div class='section-label'>Contrôle d’exhaustivité et qualité données</div>", unsafe_allow_html=True)
        st.dataframe(control, use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div class='section-label'>Détection automatique des colonnes</div>", unsafe_allow_html=True)
        st.dataframe(
            pd.DataFrame([
                {"Champ attendu": k, "Colonne détectée": v if v else "NON DÉTECTÉE"}
                for k, v in mapping.items()
            ]),
            use_container_width=True,
            hide_index=True,
        )

        with st.expander("Aperçu des 20 premières lignes du fichier"):
            st.dataframe(raw_df.head(20), use_container_width=True)

    with tab2:
        st.markdown("<div class='section-label'>État actuel Direct Livraison — un fournisseur par ligne</div>", unsafe_allow_html=True)
        st.dataframe(
            suppliers.drop(columns=["Fournisseur key"], errors="ignore"),
            use_container_width=True,
            hide_index=True,
        )

    with tab3:
        st.markdown("<div class='section-label'>Plan de lissage XD — fournisseurs basculés</div>", unsafe_allow_html=True)
        if plan.empty:
            st.info("Aucun fournisseur avec décision XD Total ou XD Markets+Supeco.")
        else:
            st.dataframe(plan, use_container_width=True, hide_index=True)

    with tab4:
        st.markdown("<div class='section-label'>Simulation de charge quai</div>", unsafe_allow_html=True)
        st.dataframe(charge, use_container_width=True, hide_index=True)
        st.caption(f"Ratio pic/creux : {stats.get('ratio_pic_creux', 0):.2f} — {stats.get('flag_ratio', 'N/A')}")

    with tab5:
        st.markdown("""
<div class='alert-card alert-blue'>
  <strong>📋 Contenu du fichier exporté :</strong><br>
  <strong>Onglet 1 — Contrôle exhaustivité</strong> : contrôles fournisseurs, décisions, finance et qualité données<br>
  <strong>Onglet 2 — État DL complet</strong> : un fournisseur par ligne avec TS, Sit95, colis, coût XD<br>
  <strong>Onglet 3 — Plan de lissage XD</strong> : cadence cible, jours de livraison, charge quai et coûts<br>
  <strong>Onglet 4 — À statuer</strong> : DL surveiller, litiges, inactifs, hors périmètre, sans données<br>
  <strong>Onglet 5 — BDD articles</strong> : liste articles par fournisseur avec décision XD associée
</div>
""", unsafe_allow_html=True)

        st.download_button(
            label="📥 Télécharger Analyse_Commando_XD.xlsx",
            data=excel_bytes,
            file_name="Analyse_Commando_XD.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

except ImportError as e:
    st.error("Dépendance Python manquante.")
    st.code(str(e))
    st.markdown("""
<div class='alert-card alert-amber'>
  <strong>Correction recommandée</strong><br>
  Ajoute les dépendances suivantes dans <code>requirements.txt</code>, puis redéploie l’application.
</div>
""", unsafe_allow_html=True)
    st.code("""streamlit
pandas
numpy
openpyxl
xlsxwriter
pyxlsb
xlrd""")

except ValueError as e:
    st.error("Erreur de structure ou de données fichier.")
    st.code(str(e))

except Exception as e:
    st.error("Erreur inattendue pendant le traitement.")
    st.exception(e)
