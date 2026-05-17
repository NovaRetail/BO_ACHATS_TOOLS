import io
import re
import unicodedata
from datetime import timedelta

import numpy as np
import pandas as pd
import streamlit as st


# ============================================================
# PARAMÈTRES MÉTIER PAR DÉFAUT
# ============================================================

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


# ============================================================
# OUTILS GÉNÉRAUX
# ============================================================

def normalize_text(value: str) -> str:
    """Normalise un nom de colonne pour rendre la détection robuste."""
    if value is None:
        return ""
    value = str(value).strip().lower()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(c for c in value if not unicodedata.combining(c))
    value = re.sub(r"[^a-z0-9]+", "", value)
    return value


def find_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    """Trouve une colonne à partir d'une liste d'alias possibles."""
    normalized_cols = {normalize_text(c): c for c in df.columns}
    for alias in aliases:
        key = normalize_text(alias)
        if key in normalized_cols:
            return normalized_cols[key]
    return None


def detect_columns(df: pd.DataFrame) -> dict:
    """Mappe les colonnes du fichier vers les noms internes attendus."""
    aliases = {
        "fou": ["Fou", "Code fournisseur", "Fournisseur", "FOU"],
        "nom_fourn": ["Nom fourn,", "Nom fourn", "Nom fournisseur", "Libellé fournisseur", "Nom fourn."],
        "site": ["Site", "Code site", "Magasin", "Code magasin"],
        "code": ["Code", "Code article", "Article", "Code produit"],
        "n_cde": ["N° Cde", "N Cde", "No Cde", "Num Cde", "Numero commande", "Numéro commande", "N° commande"],
        "date_cde": ["Date de commande", "Date commande", "Dt Cde", "Date Cde"],
        "dt_rec": ["Dt Rec", "Date réception", "Date reception", "Date de réception", "Date de reception"],
        "qte_cde": ["Qté cde", "Qte cde", "Quantité commandée", "Quantite commandee", "Qte commande"],
        "qte_rec": ["Qté reçue", "Qte recue", "Quantité reçue", "Quantite recue", "Qte reception"],
        "px_revient": ["Px revient", "Prix revient", "Prix de revient", "PR"],
        "colis": ["Colis", "Nb colis", "Nombre colis", "PCB"],
        "sit": ["Sit", "Situation", "Statut", "Statut commande"],
    }

    mapping = {}
    for key, names in aliases.items():
        mapping[key] = find_column(df, names)

    return mapping


def clean_numeric(series: pd.Series) -> pd.Series:
    """Convertit proprement des nombres au format français ou texte."""
    if series is None:
        return pd.Series(dtype=float)

    s = series.astype(str).str.strip()
    s = s.str.replace("\u00a0", "", regex=False)
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    s = s.replace({"": np.nan, "nan": np.nan, "None": np.nan})
    return pd.to_numeric(s, errors="coerce")


def safe_div(num, den):
    if den is None or pd.isna(den) or den == 0:
        return np.nan
    return num / den


def mode_day(series: pd.Series) -> str:
    s = pd.to_datetime(series, errors="coerce").dropna()
    if s.empty:
        return "N/A"
    mode_value = s.dt.dayofweek.mode()
    if mode_value.empty:
        return "N/A"
    return DAYS_FR.get(int(mode_value.iloc[0]), "N/A")


def mode_value(series: pd.Series):
    s = series.dropna()
    if s.empty:
        return np.nan
    m = s.mode()
    if m.empty:
        return np.nan
    return m.iloc[0]


def classify_group(site) -> str:
    site = str(site).strip().split(".")[0]
    if site in HYPERS:
        return "Hypers"
    if site in MARKETS:
        return "Markets"
    if site in SUPECO:
        return "Supeco"
    return "Site hors groupe"


def join_unique(values) -> str:
    vals = [str(v) for v in pd.Series(values).dropna().unique()]
    vals = [v for v in vals if v and v != "nan"]
    if not vals:
        return "N/A"
    return " / ".join(sorted(vals))


def yes_no(value: bool) -> str:
    return "Oui" if value else "Non"


def format_currency_xof(value):
    if pd.isna(value):
        return "N/A"
    return f"{value:,.0f} XOF".replace(",", " ")


def get_cycles_from_cadence(cadence: str) -> int:
    if cadence == "Hebdo":
        return 4
    if cadence == "Bi-mensuel":
        return 2
    return 1


def compute_current_cadence(cdes_mois: float) -> str:
    if pd.isna(cdes_mois):
        return "N/A"
    if cdes_mois >= 20:
        return "Hebdo"
    if cdes_mois >= 8:
        return "Bi-mensuel"
    return "Mensuel"


def compute_xd_cadence(colis_xd_mois: float) -> tuple[str, int, float, str]:
    """
    Retourne cadence, cycles/mois, colis/livraison, alerte.
    Applique les règles de plafond :
    Mensuel > 500 => Bi-mensuel
    Bi-mensuel > 600 => Hebdo
    Hebdo > 800 => alerte capacité
    """
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

    cutoff_num = base - lt

    if cutoff_num >= 0:
        return DAYS_FR[cutoff_num]

    while cutoff_num < 0:
        cutoff_num += 7

    return f"S-1 {DAYS_FR[cutoff_num]}"


# ============================================================
# LECTURE FICHIER
# ============================================================

def get_excel_sheets(uploaded_file) -> list[str]:
    name = uploaded_file.name.lower()

    if name.endswith(".xlsb"):
        xls = pd.ExcelFile(uploaded_file, engine="pyxlsb")
    elif name.endswith(".xls"):
        xls = pd.ExcelFile(uploaded_file, engine="xlrd")
    else:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")

    return xls.sheet_names


def read_uploaded_file(uploaded_file, sheet_name=None) -> pd.DataFrame:
    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file, sep=None, engine="python")
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, sep=";")

    if name.endswith(".xlsb"):
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="pyxlsb")

    if name.endswith(".xls"):
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="xlrd")

    return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl")


# ============================================================
# PRÉPARATION DES DONNÉES
# ============================================================

def prepare_data(
    raw_df: pd.DataFrame,
    start_date: pd.Timestamp,
    mapping: dict,
) -> tuple[pd.DataFrame, dict]:
    """
    Nettoie et prépare la base de commandes.
    Retourne df préparé + diagnostic qualité.
    """
    df = raw_df.copy()
    initial_rows = len(df)

    required = [
        "fou", "nom_fourn", "site", "code", "n_cde", "date_cde",
        "dt_rec", "qte_cde", "qte_rec", "px_revient", "colis", "sit"
    ]

    missing = [k for k in required if mapping.get(k) is None]

    if missing:
        missing_labels = ", ".join(missing)
        raise ValueError(f"Colonnes obligatoires introuvables : {missing_labels}")

    df = df.rename(columns={
        mapping["fou"]: "Fou",
        mapping["nom_fourn"]: "Nom fournisseur",
        mapping["site"]: "Site",
        mapping["code"]: "Code article",
        mapping["n_cde"]: "N° Cde",
        mapping["date_cde"]: "Date de commande",
        mapping["dt_rec"]: "Dt Rec",
        mapping["qte_cde"]: "Qté cde",
        mapping["qte_rec"]: "Qté reçue",
        mapping["px_revient"]: "Px revient",
        mapping["colis"]: "Colis",
        mapping["sit"]: "Sit",
    })

    df["Fou"] = df["Fou"].astype(str).str.strip()
    df["Nom fournisseur"] = df["Nom fournisseur"].astype(str).str.strip()
    df["Site"] = df["Site"].astype(str).str.strip().str.split(".").str[0]
    df["Code article"] = df["Code article"].astype(str).str.strip()
    df["N° Cde"] = df["N° Cde"].astype(str).str.strip()
    df["Sit"] = df["Sit"].astype(str).str.strip()

    df["Date de commande"] = pd.to_datetime(df["Date de commande"], errors="coerce")
    df["Dt Rec"] = pd.to_datetime(df["Dt Rec"], errors="coerce")

    df["Qté cde"] = clean_numeric(df["Qté cde"])
    df["Qté reçue"] = clean_numeric(df["Qté reçue"])
    df["Px revient"] = clean_numeric(df["Px revient"])
    df["Colis"] = clean_numeric(df["Colis"]).fillna(0)

    qte_missing = int(df["Qté cde"].isna().sum())
    px_missing = int(df["Px revient"].isna().sum())
    date_cde_missing = int(df["Date de commande"].isna().sum())
    dt_rec_missing = int(df["Dt Rec"].isna().sum())

    df = df[df["Date de commande"].notna()].copy()
    df = df[df["Date de commande"] >= start_date].copy()

    if df.empty:
        raise ValueError("Aucune ligne disponible après filtre de date.")

    last_date = df["Date de commande"].max()
    nb_days = max((last_date - start_date).days + 1, 1)
    nb_months = nb_days / 30.44

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
    df["Sit95 flag"] = df["Sit"].astype(str).str.replace(".0", "", regex=False).eq("95")

    df["Lead time brut"] = (df["Dt Rec"] - df["Date de commande"]).dt.days
    df["Lead time valide"] = df["Lead time brut"].where(
        (df["Lead time brut"] >= 0) & (df["Lead time brut"] <= 30)
    )

    site_hors_groupe = sorted(df.loc[df["Groupe magasin"].eq("Site hors groupe"), "Site"].dropna().unique())

    quality = {
        "lignes_initiales": initial_rows,
        "lignes_apres_filtre_date": len(df),
        "date_debut_analyse": start_date,
        "date_fin_analyse": last_date,
        "nb_mois_analyse": nb_months,
        "methode_nb_mois": "Nombre de jours entre date début et date fin / 30,44",
        "qte_cde_manquante_ou_nulle": int(qte_missing + (df["Qté cde"].fillna(0).eq(0)).sum()),
        "px_revient_manquant": px_missing,
        "date_commande_manquante": date_cde_missing,
        "dt_rec_manquante": dt_rec_missing,
        "sites_hors_groupe": ", ".join(site_hors_groupe) if site_hors_groupe else "Aucun",
    }

    return df, quality


# ============================================================
# CALCULS FOURNISSEURS
# ============================================================

def aggregate_group_metrics(df: pd.DataFrame, nb_months: float) -> pd.DataFrame:
    """Calcule les métriques par fournisseur et groupe magasin."""
    rows = []

    for (fkey, group), g in df.groupby(["Fournisseur key", "Groupe magasin"], dropna=False):
        bc_count = g["BC unique"].nunique()
        qte_cde = g["Qté cde"].fillna(0).sum()
        qte_rec = g["Qté reçue"].fillna(0).sum()
        sit95_bc = g.loc[g["Sit95 flag"], "BC unique"].nunique()

        rows.append({
            "Fournisseur key": fkey,
            "Groupe magasin": group,
            "BC": bc_count,
            "Colis": g["Colis"].fillna(0).sum(),
            "Colis/mois": g["Colis"].fillna(0).sum() / nb_months,
            "TS%": safe_div(qte_rec, qte_cde) * 100,
            "%Sit95": safe_div(sit95_bc, bc_count) * 100,
            "Actif": bc_count > 0,
        })

    return pd.DataFrame(rows)


def get_metric_for_group(group_metrics, fkey, group, metric):
    sub = group_metrics[
        (group_metrics["Fournisseur key"] == fkey)
        & (group_metrics["Groupe magasin"] == group)
    ]

    if sub.empty:
        return np.nan

    return sub.iloc[0][metric]


def is_group_active(group_metrics, fkey, group) -> bool:
    sub = group_metrics[
        (group_metrics["Fournisseur key"] == fkey)
        & (group_metrics["Groupe magasin"] == group)
    ]

    if sub.empty:
        return False

    return bool(sub.iloc[0]["Actif"])


def is_group_defective(group_metrics, fkey, group) -> bool:
    if not is_group_active(group_metrics, fkey, group):
        return False

    ts = get_metric_for_group(group_metrics, fkey, group, "TS%")
    sit = get_metric_for_group(group_metrics, fkey, group, "%Sit95")

    ts_bad = False if pd.isna(ts) else ts < 60
    sit_bad = False if pd.isna(sit) else sit > 30

    return ts_bad or sit_bad


def is_group_correct(group_metrics, fkey, group) -> bool:
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

    # Valeur moyenne fournisseur / magasin
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
        .agg(
            nb_couples_sous_seuil=("Site", "nunique"),
        )
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
        qte_rec = g["Qté reçue"].fillna(0).sum()
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

        nb_couples_total = int(
            total_couples.loc[
                total_couples["Fournisseur key"].eq(fkey),
                "nb_couples_total"
            ].iloc[0]
        )

        match_below = below_threshold[below_threshold["Fournisseur key"].eq(fkey)]
        nb_couples_sous_seuil = int(match_below["nb_couples_sous_seuil"].iloc[0]) if not match_below.empty else 0
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
            if np.isclose(ts_global if not pd.isna(ts_global) else -1, 0) and np.isclose(sit95_global if not pd.isna(sit95_global) else -1, 100) and not recent_order:
                decision = "Inactif probable"
                flag = "Inactif"
                raison = "TS global = 0%, Sit95 global = 100%, aucune commande dans les 60 derniers jours."
            elif np.isclose(ts_global if not pd.isna(ts_global) else -1, 0) and np.isclose(sit95_global if not pd.isna(sit95_global) else -1, 100) and recent_order:
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

        # Metrics par groupe
        ts_hypers = get_metric_for_group(group_metrics, fkey, "Hypers", "TS%")
        ts_markets = get_metric_for_group(group_metrics, fkey, "Markets", "TS%")
        ts_supeco = get_metric_for_group(group_metrics, fkey, "Supeco", "TS%")

        sit_hypers = get_metric_for_group(group_metrics, fkey, "Hypers", "%Sit95")
        sit_markets = get_metric_for_group(group_metrics, fkey, "Markets", "%Sit95")
        sit_supeco = get_metric_for_group(group_metrics, fkey, "Supeco", "%Sit95")

        colis_hypers_mois = get_metric_for_group(group_metrics, fkey, "Hypers", "Colis/mois")
        colis_markets_mois = get_metric_for_group(group_metrics, fkey, "Markets", "Colis/mois")
        colis_supeco_mois = get_metric_for_group(group_metrics, fkey, "Supeco", "Colis/mois")

        colis_hypers_mois = 0 if pd.isna(colis_hypers_mois) else colis_hypers_mois
        colis_markets_mois = 0 if pd.isna(colis_markets_mois) else colis_markets_mois
        colis_supeco_mois = 0 if pd.isna(colis_supeco_mois) else colis_supeco_mois

        if decision == "XD Total":
            colis_xd_mois = colis_hypers_mois + colis_markets_mois + colis_supeco_mois
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
    suppliers = suppliers.sort_values("Cdes/mois", ascending=False).reset_index(drop=True)

    return suppliers


# ============================================================
# PLAN DE LISSAGE XD
# ============================================================

def assign_delivery_days(plan: pd.DataFrame) -> pd.DataFrame:
    plan = plan.copy()
    plan["Jour livraison XD"] = "N/A"

    # Hebdo : Lundi / Mercredi par équilibrage
    hebdo_idx = plan[plan["Cadence XD"].eq("Hebdo")].sort_values("Colis XD/mois", ascending=False).index
    charges_hebdo = {"Lundi": 0.0, "Mercredi": 0.0}

    for idx in hebdo_idx:
        day = min(charges_hebdo, key=charges_hebdo.get)
        plan.loc[idx, "Jour livraison XD"] = day
        charges_hebdo[day] += plan.loc[idx, "Colis XD/mois"] / 4

    # Bi-mensuel : Jeudi / Vendredi par équilibrage
    bim_idx = plan[plan["Cadence XD"].eq("Bi-mensuel")].sort_values("Colis XD/mois", ascending=False).index
    charges_bim = {"Jeudi": 0.0, "Vendredi": 0.0}

    for idx in bim_idx:
        day = min(charges_bim, key=charges_bim.get)
        plan.loc[idx, "Jour livraison XD"] = day
        charges_bim[day] += plan.loc[idx, "Colis XD/mois"] / 2

    # Mensuel : Lundi / Mercredi / Jeudi / Vendredi par équilibrage
    mens_idx = plan[plan["Cadence XD"].eq("Mensuel")].sort_values("Colis XD/mois", ascending=False).index
    charges_mens = {"Lundi": 0.0, "Mercredi": 0.0, "Jeudi": 0.0, "Vendredi": 0.0}

    for idx in mens_idx:
        day = min(charges_mens, key=charges_mens.get)
        plan.loc[idx, "Jour livraison XD"] = day
        charges_mens[day] += plan.loc[idx, "Colis XD/mois"]

    return plan


def build_smoothing_plan(
    suppliers: pd.DataFrame,
    platform_cost_per_package: float,
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    plan = suppliers[
        suppliers["Décision XD"].isin(["XD Total", "XD Markets+Supeco"])
    ].copy()

    if plan.empty:
        empty_charge = pd.DataFrame(columns=[
            "Jour", "Colis/semaine simulés", "Nombre de fournisseurs",
            "Nombre de réceptions", "Charge moyenne par réception",
            "Coût traitement XD/semaine",
            "Dépassement seuil 800 colis/jour", "Alerte"
        ])

        stats = {
            "charge_max": 0,
            "charge_min": 0,
            "ratio_pic_creux": 0,
            "flag_ratio": "N/A",
            "total_cost_month": 0,
            "total_cost_year": 0,
        }

        return plan, empty_charge, stats

    plan["Groupes basculés XD"] = np.where(
        plan["Décision XD"].eq("XD Total"),
        "Hypers / Markets / Supeco",
        "Markets / Supeco"
    )
    plan["Groupes maintenus DL"] = np.where(
        plan["Décision XD"].eq("XD Markets+Supeco"),
        "Hypers",
        "Aucun"
    )

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
        axis=1
    )

    plan["BC XD/mois"] = plan["Cycles XD/mois"]
    plan["Réduction BC/mois"] = plan["BC/mois actuel"] - plan["BC XD/mois"]
    plan["% réduction BC/mois"] = (
        plan["Réduction BC/mois"] / plan["BC/mois actuel"].replace(0, np.nan) * 100
    )

    plan["Coût traitement plateforme / colis"] = platform_cost_per_package
    plan["Coût traitement XD/mois"] = plan["Colis XD/mois"] * platform_cost_per_package
    plan["Coût traitement XD/an"] = plan["Coût traitement XD/mois"] * 12
    plan["Coût traitement XD/livraison"] = plan["Colis/livraison XD"] * platform_cost_per_package
    plan["Coût traitement XD par cycle"] = plan["Coût traitement XD/livraison"]

    plan["Colis Hypers maintenus DL"] = np.where(
        plan["Décision XD"].eq("XD Markets+Supeco"),
        plan["Colis Hypers/mois"],
        0
    )

    plan["Coût théorique Hypers exclu XD"] = (
        plan["Colis Hypers maintenus DL"] * platform_cost_per_package
    )

    # Charge quai hebdomadaire
    charge_rows = []
    for day in ["Lundi", "Mercredi", "Jeudi", "Vendredi"]:
        sub = plan[plan["Jour livraison XD"].eq(day)].copy()

        # Simulation colis/semaine :
        # Hebdo = colis/mois / 4
        # Bi-mensuel = colis/mois / 2
        # Mensuel = colis/mois / 4 pour lisser en équivalent hebdo
        weekly_packages = 0
        receptions = 0

        for _, r in sub.iterrows():
            cycles = r["Cycles XD/mois"]
            if cycles > 0:
                weekly_packages += r["Colis XD/mois"] / 4
                receptions += max(cycles / 4, 0.25)

        nb_suppliers = sub["Fournisseur key"].nunique()
        avg_per_reception = safe_div(weekly_packages, receptions)
        cost_week = weekly_packages * platform_cost_per_package
        over = weekly_packages > 800

        if weekly_packages < 500:
            alert = "🟢"
        elif weekly_packages < 800:
            alert = "🟠"
        else:
            alert = "🔴"

        charge_rows.append({
            "Jour": day,
            "Colis/semaine simulés": weekly_packages,
            "Nombre de fournisseurs": nb_suppliers,
            "Nombre de réceptions": receptions,
            "Charge moyenne par réception": avg_per_reception,
            "Coût traitement XD/semaine": cost_week,
            "Dépassement seuil 800 colis/jour": yes_no(over),
            "Alerte": alert,
        })

    charge = pd.DataFrame(charge_rows)

    non_zero = charge.loc[charge["Colis/semaine simulés"] > 0, "Colis/semaine simulés"]
    charge_max = non_zero.max() if not non_zero.empty else 0
    charge_min = non_zero.min() if not non_zero.empty else 0
    ratio = safe_div(charge_max, charge_min) if charge_min else 0

    stats = {
        "charge_max": charge_max,
        "charge_min": charge_min,
        "ratio_pic_creux": ratio,
        "flag_ratio": "OK" if ratio <= 3 else "À lisser",
        "total_cost_month": plan["Coût traitement XD/mois"].sum(),
        "total_cost_year": plan["Coût traitement XD/an"].sum(),
    }

    export_cols = [
        "Code fournisseur",
        "Nom fournisseur",
        "Décision XD",
        "Groupes basculés XD",
        "Groupes maintenus DL",
        "Cadence actuelle",
        "Jour de commande actuel",
        "Jour de livraison actuel",
        "Lead time médian (j)",
        "Colis/cde actuel",
        "BC/mois actuel",
        "Colis/mois actuel",
        "Colis XD/mois",
        "Cadence XD",
        "Cycles XD/mois",
        "Colis/livraison XD",
        "Alerte colis",
        "Jour livraison XD",
        "Jour cut-off",
        "BC XD/mois",
        "Réduction BC/mois",
        "% réduction BC/mois",
        "Coût traitement plateforme / colis",
        "Coût traitement XD/mois",
        "Coût traitement XD/an",
        "Coût traitement XD/livraison",
        "Coût traitement XD par cycle",
        "Colis Hypers maintenus DL",
        "Coût théorique Hypers exclu XD",
    ]

    return plan[export_cols].reset_index(drop=True), charge, stats


# ============================================================
# À STATUER
# ============================================================

def build_to_decide(suppliers: pd.DataFrame) -> pd.DataFrame:
    rows = []

    action_map = {
        "DL — Surveiller": "Revoir dans 3 mois avec suivi TS%, Sit95 et valeur moyenne commande.",
        "Litige probable": "Escalader aux Achats / clarifier litige fournisseur / bloquer commandes si nécessaire.",
        "Inactif probable": "Vérifier référencement / suspendre ou nettoyer base fournisseur.",
        "Non applicable": "Maintien DL ; revoir uniquement si baisse de valeur commande ou dégradation TS.",
    }

    subset = suppliers[
        suppliers["Décision XD"].isin(["DL — Surveiller", "Litige probable", "Inactif probable"])
        | suppliers["Catégorie périmètre"].eq("Hors périmètre XD")
    ].copy()

    for _, r in subset.iterrows():
        if r["Catégorie périmètre"] == "Hors périmètre XD":
            decision = "Hors périmètre XD"
            action = action_map["Non applicable"]
            reason = "Tous les couples fournisseur/magasin sont au-dessus ou égaux au seuil XD."
        else:
            decision = r["Décision XD"]
            action = action_map.get(decision, "À analyser.")
            reason = r["Raison de décision"]

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


# ============================================================
# BDD ARTICLES
# ============================================================

def build_article_db(
    df: pd.DataFrame,
    suppliers: pd.DataFrame,
    platform_cost_per_package: float,
) -> pd.DataFrame:
    agg = (
        df.groupby(["Fournisseur key", "Fou", "Nom fournisseur", "Code article"], dropna=False)
        .agg(
            nb_magasins=("Site", "nunique"),
            groupes=("Groupe magasin", join_unique),
            qte_commandee=("Qté cde", "sum"),
            qte_recue=("Qté reçue", "sum"),
            valeur_commande=("Valeur commande", "sum"),
            colis_total=("Colis", "sum"),
            nb_bc=("BC unique", "nunique"),
            derniere_commande=("Date de commande", "max"),
        )
        .reset_index()
    )

    agg["TS% article"] = agg["qte_recue"] / agg["qte_commandee"].replace(0, np.nan) * 100

    # Colis article par groupe pour calcul XD M+S
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

    agg = agg.merge(ms_colis, on=["Fournisseur key", "Code article"], how="left")
    agg["colis_ms"] = agg["colis_ms"].fillna(0)

    sup_info = suppliers[[
        "Fournisseur key",
        "Catégorie périmètre",
        "Décision XD",
    ]].drop_duplicates()

    agg = agg.merge(sup_info, on="Fournisseur key", how="left")

    def groups_switched(decision):
        if decision == "XD Total":
            return "Hypers / Markets / Supeco"
        if decision == "XD Markets+Supeco":
            return "Markets / Supeco"
        return "Aucun"

    def article_cost(row):
        if row["Décision XD"] == "XD Total":
            return row["colis_total"] * platform_cost_per_package
        if row["Décision XD"] == "XD Markets+Supeco":
            return row["colis_ms"] * platform_cost_per_package
        return 0

    agg["Groupes basculés XD"] = agg["Décision XD"].apply(groups_switched)
    agg["Coût traitement XD article théorique"] = agg.apply(article_cost, axis=1)
    agg["Commentaire"] = np.where(
        agg["Décision XD"].isin(["XD Total", "XD Markets+Supeco"]),
        "Article rattaché à un fournisseur basculé XD.",
        "Article rattaché à un fournisseur non basculé XD."
    )

    final = agg.rename(columns={
        "Fou": "Code fournisseur",
        "Code article": "Code article",
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
        "Code fournisseur",
        "Nom fournisseur",
        "Code article",
        "Nb magasins où l’article est commandé",
        "Groupes magasins présents",
        "Qté commandée totale",
        "Qté reçue totale",
        "TS% article",
        "Valeur commande totale article",
        "Colis total article",
        "Nb BC article",
        "Dernière date de commande article",
        "Catégorie périmètre fournisseur",
        "Décision XD fournisseur",
        "Groupes basculés XD",
        "Coût traitement XD article théorique",
        "Commentaire",
    ]

    return final[cols].sort_values(["Code fournisseur", "Valeur commande totale article"], ascending=[True, False])


# ============================================================
# CONTRÔLE EXHAUSTIVITÉ
# ============================================================

def build_control_sheet(
    df: pd.DataFrame,
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

    decisions = suppliers.loc[suppliers["Catégorie périmètre"].eq("Candidat XD"), "Décision XD"]

    xd_total = int((decisions == "XD Total").sum())
    xd_ms = int((decisions == "XD Markets+Supeco").sum())
    dl_surv = int((decisions == "DL — Surveiller").sum())
    litige = int((decisions == "Litige probable").sum())
    inactif = int((decisions == "Inactif probable").sum())
    sum_decisions = xd_total + xd_ms + dl_surv + litige + inactif
    diff_decisions = candidats - sum_decisions

    plan_count = len(plan)
    expected_plan = xd_total + xd_ms

    total_colis_xd_mois = plan["Colis XD/mois"].sum() if not plan.empty else 0
    total_cost_month = total_colis_xd_mois * platform_cost_per_package
    total_cost_year = total_cost_month * 12
    financial_check = np.isclose(total_cost_month, charge_stats.get("total_cost_month", 0))

    rows = [
        {"Section": "Synthèse fournisseurs", "Indicateur": "Nombre total de fournisseurs uniques", "Valeur": total_suppliers},
        {"Section": "Synthèse fournisseurs", "Indicateur": "Sans données suffisantes", "Valeur": sans_data},
        {"Section": "Synthèse fournisseurs", "Indicateur": "Candidats XD", "Valeur": candidats},
        {"Section": "Synthèse fournisseurs", "Indicateur": "Hors périmètre XD", "Valeur": hors},
        {"Section": "Synthèse fournisseurs", "Indicateur": "Somme catégories", "Valeur": sum_categories},
        {"Section": "Synthèse fournisseurs", "Indicateur": "Écart catégories", "Valeur": diff_categories},
        {"Section": "Synthèse fournisseurs", "Indicateur": "Flag contrôle fournisseurs", "Valeur": "OK" if diff_categories == 0 else "ÉCART À CORRIGER"},

        {"Section": "Décisions candidats XD", "Indicateur": "XD Total", "Valeur": xd_total},
        {"Section": "Décisions candidats XD", "Indicateur": "XD Markets+Supeco", "Valeur": xd_ms},
        {"Section": "Décisions candidats XD", "Indicateur": "DL — Surveiller", "Valeur": dl_surv},
        {"Section": "Décisions candidats XD", "Indicateur": "Litige probable", "Valeur": litige},
        {"Section": "Décisions candidats XD", "Indicateur": "Inactif probable", "Valeur": inactif},
        {"Section": "Décisions candidats XD", "Indicateur": "Total décisions candidats", "Valeur": sum_decisions},
        {"Section": "Décisions candidats XD", "Indicateur": "Écart décisions", "Valeur": diff_decisions},
        {"Section": "Décisions candidats XD", "Indicateur": "Flag contrôle décisions", "Valeur": "OK" if diff_decisions == 0 else "ÉCART À CORRIGER"},

        {"Section": "Plan de lissage", "Indicateur": "Fournisseurs dans plan de lissage", "Valeur": plan_count},
        {"Section": "Plan de lissage", "Indicateur": "XD Total + XD Markets+Supeco attendus", "Valeur": expected_plan},
        {"Section": "Plan de lissage", "Indicateur": "Flag contrôle plan", "Valeur": "OK" if plan_count == expected_plan else "ÉCART À CORRIGER"},
        {"Section": "Plan de lissage", "Indicateur": "Ratio pic/creux charge quai", "Valeur": charge_stats.get("ratio_pic_creux", 0)},
        {"Section": "Plan de lissage", "Indicateur": "Flag ratio charge quai", "Valeur": charge_stats.get("flag_ratio", "N/A")},

        {"Section": "Contrôle financier XD", "Indicateur": "Total colis XD/mois", "Valeur": total_colis_xd_mois},
        {"Section": "Contrôle financier XD", "Indicateur": "Coût unitaire traitement plateforme", "Valeur": platform_cost_per_package},
        {"Section": "Contrôle financier XD", "Indicateur": "Coût total traitement XD/mois", "Valeur": total_cost_month},
        {"Section": "Contrôle financier XD", "Indicateur": "Coût total traitement XD/an", "Valeur": total_cost_year},
        {"Section": "Contrôle financier XD", "Indicateur": "Vérification coût = colis × 90", "Valeur": "OK" if financial_check else "ÉCART À CORRIGER"},

        {"Section": "Contrôle qualité données", "Indicateur": "Lignes initiales", "Valeur": quality["lignes_initiales"]},
        {"Section": "Contrôle qualité données", "Indicateur": "Lignes après filtre date", "Valeur": quality["lignes_apres_filtre_date"]},
        {"Section": "Contrôle qualité données", "Indicateur": "Date début analyse", "Valeur": quality["date_debut_analyse"]},
        {"Section": "Contrôle qualité données", "Indicateur": "Date fin analyse", "Valeur": quality["date_fin_analyse"]},
        {"Section": "Contrôle qualité données", "Indicateur": "Nombre de mois analyse", "Valeur": quality["nb_mois_analyse"]},
        {"Section": "Contrôle qualité données", "Indicateur": "Méthode nombre de mois", "Valeur": quality["methode_nb_mois"]},
        {"Section": "Contrôle qualité données", "Indicateur": "Qté cde manquante ou nulle", "Valeur": quality["qte_cde_manquante_ou_nulle"]},
        {"Section": "Contrôle qualité données", "Indicateur": "Px revient manquant", "Valeur": quality["px_revient_manquant"]},
        {"Section": "Contrôle qualité données", "Indicateur": "Date commande manquante", "Valeur": quality["date_commande_manquante"]},
        {"Section": "Contrôle qualité données", "Indicateur": "Dt Rec manquante", "Valeur": quality["dt_rec_manquante"]},
        {"Section": "Contrôle qualité données", "Indicateur": "Sites hors groupe", "Valeur": quality["sites_hors_groupe"]},
    ]

    return pd.DataFrame(rows)


# ============================================================
# EXPORT EXCEL
# ============================================================

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
        suppliers.drop(columns=["Fournisseur key"], errors="ignore").to_excel(
            writer,
            sheet_name="2_Etat_DL_complet",
            index=False
        )

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
        })

        fmt_money = workbook.add_format({"num_format": '#,##0 "XOF"'})
        fmt_num = workbook.add_format({"num_format": "#,##0.0"})
        fmt_int = workbook.add_format({"num_format": "#,##0"})
        fmt_pct = workbook.add_format({"num_format": "0.0%"})
        fmt_date = workbook.add_format({"num_format": "dd/mm/yyyy"})
        fmt_alert_red = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
        fmt_alert_orange = workbook.add_format({"bg_color": "#FCE4D6", "font_color": "#9C6500"})
        fmt_alert_green = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})

        decision_colors = {
            "XD Total": "#F4B183",
            "XD Markets+Supeco": "#FFD966",
            "DL — Surveiller": "#FFF2CC",
            "Litige probable": "#FFC7CE",
            "Inactif probable": "#D9EAD3",
            "Hors périmètre XD": "#DDEBF7",
            "Sans données suffisantes": "#E7E6E6",
        }

        for sheet_name, df_sheet in [
            ("1_Controle_exhaustivite", control),
            ("2_Etat_DL_complet", suppliers.drop(columns=["Fournisseur key"], errors="ignore")),
            ("4_A_statuer", to_decide),
            ("5_BDD_articles", article_db),
        ]:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(len(df_sheet), 1), max(len(df_sheet.columns) - 1, 0))

            for col_num, col_name in enumerate(df_sheet.columns):
                ws.write(0, col_num, col_name, fmt_header)
                width = min(max(len(str(col_name)) + 2, 12), 35)
                ws.set_column(col_num, col_num, width)

        ws_plan = writer.sheets["3_Plan_lissage_XD"]
        ws_plan.freeze_panes(1, 0)
        ws_plan.autofilter(0, 0, max(len(plan), 1), max(len(plan.columns) - 1, 0))

        for col_num, col_name in enumerate(plan.columns):
            header_fmt = fmt_header_orange if "Coût" in col_name else fmt_header
            ws_plan.write(0, col_num, col_name, header_fmt)
            width = min(max(len(str(col_name)) + 2, 12), 35)
            ws_plan.set_column(col_num, col_num, width)

        # Header de la charge quai
        for col_num, col_name in enumerate(charge.columns):
            ws_plan.write(start_charge, col_num, col_name, fmt_header)

        # Formats par type de colonne
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            # largeur par défaut
            ws.set_default_row(18)

        # Mise en forme conditionnelle décisions sur Etat DL
        ws = writer.sheets["2_Etat_DL_complet"]
        if not suppliers.empty:
            cols = list(suppliers.drop(columns=["Fournisseur key"], errors="ignore").columns)
            if "Décision XD" in cols:
                col_idx = cols.index("Décision XD")
                col_letter = chr(ord("A") + col_idx) if col_idx < 26 else None
                if col_letter:
                    for decision, color in decision_colors.items():
                        fmt = workbook.add_format({"bg_color": color})
                        ws.conditional_format(
                            1, col_idx, len(suppliers), col_idx,
                            {
                                "type": "text",
                                "criteria": "containing",
                                "value": decision,
                                "format": fmt,
                            }
                        )

        # Mise en évidence alertes colis dans plan
        if not plan.empty and "Alerte colis" in plan.columns:
            alert_col = list(plan.columns).index("Alerte colis")
            ws_plan.conditional_format(1, alert_col, len(plan), alert_col, {
                "type": "text", "criteria": "containing", "value": "🔴", "format": fmt_alert_red
            })
            ws_plan.conditional_format(1, alert_col, len(plan), alert_col, {
                "type": "text", "criteria": "containing", "value": "🟠", "format": fmt_alert_orange
            })
            ws_plan.conditional_format(1, alert_col, len(plan), alert_col, {
                "type": "text", "criteria": "containing", "value": "🟢", "format": fmt_alert_green
            })

        # Formats numériques simples par nom de colonne
        all_sheets_data = {
            "1_Controle_exhaustivite": control,
            "2_Etat_DL_complet": suppliers.drop(columns=["Fournisseur key"], errors="ignore"),
            "3_Plan_lissage_XD": plan,
            "4_A_statuer": to_decide,
            "5_BDD_articles": article_db,
        }

        for sheet_name, df_sheet in all_sheets_data.items():
            ws = writer.sheets[sheet_name]
            for idx, col in enumerate(df_sheet.columns):
                col_lower = str(col).lower()

                if "coût" in col_lower or "valeur" in col_lower or "prix" in col_lower:
                    ws.set_column(idx, idx, 18, fmt_money)
                elif "date" in col_lower:
                    ws.set_column(idx, idx, 16, fmt_date)
                elif "%" in col_lower or "ts%" in col_lower:
                    ws.set_column(idx, idx, 14, fmt_num)
                elif "colis" in col_lower or "bc" in col_lower or "nb " in col_lower:
                    ws.set_column(idx, idx, 14, fmt_num)

    output.seek(0)
    return output.read()


# ============================================================
# PIPELINE COMPLET
# ============================================================

def run_analysis(
    raw_df: pd.DataFrame,
    start_date,
    xd_threshold: float,
    min_orders: int,
    platform_cost_per_package: float,
) -> dict:
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
    article_db = build_article_db(df, suppliers, platform_cost_per_package)

    control = build_control_sheet(
        df=df,
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


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Commando XD",
    page_icon="🚚",
    layout="wide",
)

st.title("🚚 Commando XD — Analyse DL vers Cross-Docking")
st.caption("Analyse fournisseurs, candidats XD, plan de lissage, coût plateforme et export Excel.")

with st.sidebar:
    st.header("Paramètres")

    start_date = st.date_input(
        "Date début analyse",
        value=DEFAULT_START_DATE.date(),
    )

    xd_threshold = st.number_input(
        "Seuil valeur moyenne livraison XD",
        min_value=0,
        value=DEFAULT_XD_THRESHOLD,
        step=10_000,
        help="Règle stricte : candidat si valeur moyenne < seuil.",
    )

    min_orders = st.number_input(
        "Minimum BC pour données suffisantes",
        min_value=1,
        value=DEFAULT_MIN_ORDERS,
        step=1,
    )

    platform_cost = st.number_input(
        "Coût traitement plateforme par colis",
        min_value=0,
        value=DEFAULT_PLATFORM_COST_PER_PACKAGE,
        step=10,
    )

uploaded_file = st.file_uploader(
    "Charge ton fichier de commandes fournisseurs",
    type=["xlsx", "xlsb", "xls", "csv"],
)

if uploaded_file is None:
    st.info("Charge un fichier Excel, XLSB ou CSV pour lancer l’analyse.")
    st.stop()

try:
    if not uploaded_file.name.lower().endswith(".csv"):
        sheets = get_excel_sheets(uploaded_file)
        uploaded_file.seek(0)
        sheet_name = st.selectbox("Sélectionne l’onglet à analyser", sheets)
    else:
        sheet_name = None

    raw_df = read_uploaded_file(uploaded_file, sheet_name=sheet_name)

    st.subheader("Aperçu fichier")
    st.dataframe(raw_df.head(20), use_container_width=True)

    mapping_preview = detect_columns(raw_df)
    with st.expander("Voir détection des colonnes"):
        st.dataframe(
            pd.DataFrame([
                {"Champ attendu": k, "Colonne détectée": v}
                for k, v in mapping_preview.items()
            ]),
            use_container_width=True,
        )

    if st.button("🚀 Lancer l’analyse Commando XD", type="primary"):
        with st.spinner("Analyse en cours..."):
            result = run_analysis(
                raw_df=raw_df,
                start_date=start_date,
                xd_threshold=xd_threshold,
                min_orders=int(min_orders),
                platform_cost_per_package=platform_cost,
            )

        suppliers = result["suppliers"]
        plan = result["plan"]
        charge = result["charge"]
        control = result["control"]
        stats = result["charge_stats"]

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
        total_cost_month = total_colis_xd * platform_cost
        total_cost_year = total_cost_month * 12

        st.success("Analyse terminée.")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Fournisseurs", f"{total_suppliers:,.0f}".replace(",", " "))
        c2.metric("Candidats XD", f"{candidats:,.0f}".replace(",", " "))
        c3.metric("XD Total", f"{xd_total:,.0f}".replace(",", " "))
        c4.metric("XD M+S", f"{xd_ms:,.0f}".replace(",", " "))

        c5, c6, c7, c8 = st.columns(4)
        c5.metric("DL Surveiller", f"{dl_surv:,.0f}".replace(",", " "))
        c6.metric("Litiges", f"{litige:,.0f}".replace(",", " "))
        c7.metric("Inactifs", f"{inactif:,.0f}".replace(",", " "))
        c8.metric("Hors périmètre", f"{hors:,.0f}".replace(",", " "))

        f1, f2, f3 = st.columns(3)
        f1.metric("Colis XD/mois", f"{total_colis_xd:,.0f}".replace(",", " "))
        f2.metric("Coût XD/mois", format_currency_xof(total_cost_month))
        f3.metric("Coût XD/an", format_currency_xof(total_cost_year))

        st.subheader("Contrôle d’exhaustivité")
        st.dataframe(control, use_container_width=True)

        st.subheader("État DL complet")
        st.dataframe(
            suppliers.drop(columns=["Fournisseur key"], errors="ignore"),
            use_container_width=True,
        )

        st.subheader("Plan de lissage XD")
        st.dataframe(plan, use_container_width=True)

        st.subheader("Simulation charge quai")
        st.dataframe(charge, use_container_width=True)

        st.info(
            f"Ratio pic/creux charge quai : {stats.get('ratio_pic_creux', 0):.2f} "
            f"— {stats.get('flag_ratio', 'N/A')}"
        )

        st.download_button(
            label="📥 Télécharger Analyse_Commando_XD.xlsx",
            data=result["excel_bytes"],
            file_name="Analyse_Commando_XD.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

except Exception as e:
    st.error("Erreur pendant le traitement.")
    st.exception(e)
