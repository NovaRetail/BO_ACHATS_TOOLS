"""Moteur du module OTIF : lecture de l’extraction ERP, calculs et exports Excel (repris tels quels de la page v6.4)."""
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO          # ← UNIQUEMENT pour les exports Excel en sortie
from datetime import date as _date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import plotly.graph_objects as go

# ── Import module I/O robuste SmartBuyer Hub ──────────────────────────────
# utils_io.py doit être à la RACINE du repo (même niveau que ce fichier).
# Il fournit read_csv_robust, read_excel_robust, clean_dataframe.
try:
    from utils_io import read_csv_robust, read_excel_robust, clean_dataframe
except ImportError:
    # Fallback minimal si utils_io absent (ne devrait pas arriver en prod)
    import io as _io

    def read_csv_robust(file_bytes, sep=";", decimal=",", **kw):
        # Retirer les params gérés explicitement pour éviter les doublons dans **kw
        kw.pop("low_memory", None)
        kw.pop("thousands", None)
        kw.pop("dtype", None)
        kw.pop("parse_dates", None)
        engine = kw.pop("engine", None)
        # v6.3 : low_memory incompatible avec engine='python' (ou sep=None)
        use_python_engine = (engine == "python") or (sep is None)
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                if use_python_engine:
                    return pd.read_csv(_io.BytesIO(file_bytes), sep=sep, decimal=decimal,
                                       encoding=enc, engine="python", **kw)
                return pd.read_csv(_io.BytesIO(file_bytes), sep=sep, decimal=decimal,
                                   encoding=enc, low_memory=False, **kw)
            except (UnicodeDecodeError, UnicodeError):
                continue
        if use_python_engine:
            return pd.read_csv(_io.BytesIO(file_bytes), sep=sep, decimal=decimal,
                               encoding="latin-1", encoding_errors="replace",
                               engine="python", **kw)
        return pd.read_csv(_io.BytesIO(file_bytes), sep=sep, decimal=decimal,
                           encoding="latin-1", encoding_errors="replace",
                           low_memory=False, **kw)

    def read_excel_robust(file_bytes, **kw):
        return pd.read_excel(_io.BytesIO(file_bytes), engine="openpyxl", **kw)

    def clean_dataframe(df, strip_strings=True):
        df = df.loc[:, ~df.columns.str.contains(r"^Unnamed", na=False)]
        df.columns = df.columns.str.strip()
        if strip_strings:
            sc = df.select_dtypes(include="object").columns
            df[sc] = df[sc].apply(lambda c: c.str.strip() if c.dtype == "object" else c)
        return df


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES MÉTIER
# ══════════════════════════════════════════════════════════════════════════════
TECHNICAL_SUPPLIERS = {
    "FOURNISSEUR STOCK",
    "FOURNISSEUR PLATEFORME LOCAL",
    "FOURNISSEUR PLATEFORME IMPORT",
}
DATE_EXPECTED_CANDIDATES = ["H Date", "Date livraison", "Date prévue", "Date"]
SEUIL_EXCELLENT  = 97
SEUIL_SURVEILLER = 90

# ── Rayon (Libellé rayon) — Droguerie + Parfumerie Hygiène fusionnés en DPH
# pour retrouver les 3 rayons réels de l'équipe (Épicerie, Boissons, DPH).
# "Libellé département" n'est PAS utilisable pour cette segmentation : il ne
# contient qu'une seule valeur ("PGC") sur l'export ERP, quel que soit le rayon.
RAYON_MAP = {
    "EPICERIE": "Épicerie",
    "BOISSONS": "Boissons",
    "DROGUERIE": "DPH",
    "PARFUMERIE HYGIENE": "DPH",
}


def rayon_bucket(v) -> str:
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return "Inconnu"
    return RAYON_MAP.get(s.upper(), s)

WATCH_GOLD_HEX  = "FFD60A"
WATCH_LIGHT_HEX = "FFFDE7"

# ── PLAN D'ACTION OTIF — Responsable Achat ─────────────────────────────────
CFAO_NAVY   = "24235C"
CFAO_ORANGE = "F09F39"

CAUSES_RACINE = [
    "Rupture usine / capacité fournisseur",
    "Rupture matière première",
    "Logistique-transport",
    "Erreur de commande / donnée",
    "Non-conformité qualité",
    "Blocage financier",
    "Fournisseur non joignable / sans réponse",
    "Autre",
]
STATUTS_SUIVI = ["À contacter", "Contacté", "Cause identifiée", "Plan en cours", "Résolu"]


def classe_bucket(classe: str) -> str:
    """GOLD/SILVER par sous-chaîne — couvre les libellés composites (cf. _classe_fills)."""
    c = str(classe).strip().upper()
    if "GOLD" in c:
        return "GOLD"
    if "SILVER" in c:
        return "SILVER"
    return ""


def format_magasin(site_label: str) -> str:
    u = str(site_label).upper()
    if "HYPER" in u:
        return "Hyper"
    if "SUPECO" in u:
        return "Supeco"
    if "MARKET" in u:
        return "Market"
    return "Autre"


# ══════════════════════════════════════════════════════════════════════════════
# HELPER GÉNÉRIQUE — Tri défensif
# ══════════════════════════════════════════════════════════════════════════════
def safe_sort(df: pd.DataFrame, keys, ascending) -> pd.DataFrame:
    """Trie uniquement sur les colonnes présentes — évite les KeyError."""
    if df is None or df.empty:
        return df
    present = [(k, a) for k, a in zip(keys, ascending) if k in df.columns]
    if not present:
        return df
    cols = [k for k, _ in present]
    asc  = [a for _, a in present]
    return df.sort_values(cols, ascending=asc)


# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT LISTE SURVEILLANCE
# v6.3 : lecture CSV réécrite pour éviter le conflit low_memory/engine='python'
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load_watchlist(file_bytes: bytes, filename: str) -> dict:
    """
    Charge la liste de surveillance depuis un CSV ou Excel.
    Cherche :
      - une colonne 'code/article/ref/ean' → codes articles
      - une colonne 'classe/class/segment/categ' → libellé de classe
    Retourne un dict {code_normalisé: classe_str}.

    v6.3 — FIX lecture CSV :
      L'ancienne implémentation appelait
      `read_csv_robust(file_bytes, sep=None, engine="python")`, ce qui
      provoquait l'erreur pandas :
        "The 'low_memory' option is not supported with the 'python' engine"
      car `read_csv_robust` (utils_io.py) injecte `low_memory=False` en
      interne sans vérifier le moteur demandé.

      Nouvelle stratégie, indépendante du comportement interne de
      utils_io.py :
        1) essayer les séparateurs usuels (";", ",", "\\t") un par un
           via read_csv_robust AVEC un sep explicite (→ moteur C,
           low_memory est alors compatible, aucun risque de conflit) ;
        2) si aucun de ces séparateurs ne donne ≥ 2 colonnes exploitables,
           repli ultime sur une auto-détection pandas native (sans passer
           par read_csv_robust) avec engine='python' et SANS low_memory.
    """
    try:
        fname_lower = filename.lower()
        if fname_lower.endswith((".xlsx", ".xls", ".xlsm")):
            # ── Excel : read_excel_robust (engine openpyxl avec fallback xlrd)
            wdf = read_excel_robust(file_bytes)
        else:
            # ── CSV : tentative avec séparateurs usuels (moteur C, sûr)
            wdf = None
            last_err = None
            for candidate_sep in (";", ",", "\t"):
                try:
                    test_df = read_csv_robust(file_bytes, sep=candidate_sep)
                    if test_df is not None and test_df.shape[1] >= 2:
                        wdf = test_df
                        break
                except Exception as e:
                    last_err = e
                    continue

            if wdf is None:
                # ── Repli ultime : auto-détection python, SANS low_memory,
                #    en bypassant volontairement read_csv_robust pour ne
                #    pas hériter d'un low_memory injecté en interne.
                import io as _io_wl
                for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
                    try:
                        wdf = pd.read_csv(
                            _io_wl.BytesIO(file_bytes), sep=None,
                            engine="python", encoding=enc,
                        )
                        break
                    except (UnicodeDecodeError, UnicodeError) as e:
                        last_err = e
                        continue
                if wdf is None:
                    wdf = pd.read_csv(
                        _io_wl.BytesIO(file_bytes), sep=None, engine="python",
                        encoding="latin-1", encoding_errors="replace",
                    )

        wdf = clean_dataframe(wdf)

        # Détection colonne code
        code_col = None
        for col in wdf.columns:
            if any(kw in col.lower() for kw in ["code", "article", "ref", "ean"]):
                code_col = col
                break
        if code_col is None:
            code_col = wdf.columns[0]

        # Détection colonne classe
        classe_col = None
        for col in wdf.columns:
            if col == code_col:
                continue
            if any(kw in col.lower() for kw in ["class", "type", "segment", "categ", "tier", "niveau", "groupe"]):
                classe_col = col
                break
        if classe_col is None and len(wdf.columns) >= 2:
            classe_col = [c for c in wdf.columns if c != code_col][0]

        def norm(v):
            return str(v).strip().replace(".0", "")

        result = {}
        for _, row in wdf.dropna(subset=[code_col]).iterrows():
            code = norm(row[code_col])
            if not code or code == "nan":
                continue
            classe = str(row[classe_col]).strip() if classe_col and pd.notna(row[classe_col]) else "⭐"
            result[code] = classe

        return result

    except Exception as e:
        st.sidebar.error(f"Erreur lecture liste surveillance : {e}")
        return {}


def normalise_code(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip().replace(".0", "")


def watchdict_codes(watchdict: dict) -> set:
    return set(watchdict.keys())


def get_classe(code, watchdict: dict) -> str:
    return watchdict.get(normalise_code(code), "")


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS AFFICHAGE
# ══════════════════════════════════════════════════════════════════════════════
def fmt(n) -> str:
    if pd.isna(n) or n is None:
        return "—"
    a = abs(float(n))
    if a >= 1_000_000:
        return f"{n / 1_000_000:.1f} M"
    if a >= 1_000:
        return f"{int(n / 1_000)} K"
    return f"{int(n):,}"


def fmt_pct(v, decimals: int = 1) -> str:
    if pd.isna(v) or v is None:
        return "—"
    return f"{v:.{decimals}f}%"


def score_band(v) -> str:
    if pd.isna(v):
        return "Inconnu"
    if v >= SEUIL_EXCELLENT:
        return "🟢 Excellent"
    if v >= SEUIL_SURVEILLER:
        return "🟠 À surveiller"
    return "🔴 Critique"


def score_color(v) -> str:
    if pd.isna(v):
        return "#8E8E93"
    if v >= SEUIL_EXCELLENT:
        return "#34C759"
    if v >= SEUIL_SURVEILLER:
        return "#FF9500"
    return "#FF3B30"


def safe_div(a, b) -> float:
    return a / b if b not in (0, None) and not pd.isna(b) else 0.0


# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT & NETTOYAGE ERP
# ── v6.2 : read_csv_robust remplace pd.read_csv(BytesIO(...))
#            → gère CP1252, Latin-1, UTF-8, UTF-8-BOM automatiquement
# ══════════════════════════════════════════════════════════════════════════════
def detect_expected_date_column(df: pd.DataFrame):
    for col in DATE_EXPECTED_CANDIDATES:
        if col in df.columns:
            if pd.to_datetime(df[col], errors="coerce").notna().sum() > 0:
                return col
    return None


@st.cache_data(show_spinner="⏳ Chargement du fichier ERP…")
def load_erp(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Lecture robuste du CSV ERP.
    Détecte automatiquement l'encodage (UTF-8, CP1252, Latin-1, UTF-8-BOM…).
    Ne lève JAMAIS UnicodeDecodeError.
    """
    try:
        df = read_csv_robust(
            file_bytes,
            sep=";",
            decimal=",",
            thousands=" ",
            low_memory=False,
        )
    except Exception as e:
        st.error(
            f"❌ Impossible de lire le fichier **{filename}**.\n\n"
            f"Vérifiez que c'est bien un CSV avec séparateur `;`.\n\n"
            f"Détail : `{e}`"
        )
        st.stop()

    # Nettoyage post-import standard
    df = clean_dataframe(df)

    # Nettoyage spécifique ERP : BOM résiduel sur première colonne, ponctuation finale
    df.columns = [str(c).replace("\ufeff", "").strip().rstrip(",.") for c in df.columns]

    # Supprimer colonnes vides
    df = df.dropna(axis=1, how="all")

    # Valeurs objet → None propre
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip().replace({"nan": None, "": None})

    # Parsing dates
    date_cols = [
        "Dt Rec", "Date de commande", "Date", "Date facture",
        "Date comptable du rapprochement", "H Date", "Date livraison", "Date prévue",
    ]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")

    # Parsing numériques
    num_cols = [
        "Site", "Département", "N° Cde", "Sit", "Fou", "Famille", "Sous-famille",
        "Code", "Article", "Qté cde", "Poids cde", "Qté rec", "Poids rec",
        "PV", "Px revient", "Prix de vente HT", "Marge ligne", "Taux TVA",
        "EAN", "Poids unitaire", "PCB", "Colis",
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def prepare_dataset(df: pd.DataFrame, exclude_technical: bool = True,
                    cap_sur_receipt: bool = True, watchdict: dict = None):
    if df.empty:
        return df.copy(), {}

    work = df.copy()
    raw_len = len(work)

    def _col(name):
        return work[name] if name in work.columns else pd.Series("Inconnu", index=work.index)

    work["site_label"]    = _col("Libellé site").fillna("Inconnu").astype(str).str.strip()
    supplier_col = next((c for c in work.columns if c.startswith("Nom fourn")), None)
    work["supplier_name"] = (work[supplier_col] if supplier_col else pd.Series("Inconnu", index=work.index)).fillna("Inconnu").astype(str).str.strip()
    work["article_label"] = _col("Libellé article").fillna("Inconnu").astype(str).str.strip()
    work["dept_label"]    = _col("Libellé département").fillna("Inconnu").astype(str).str.strip()
    work["rayon_label"]   = _col("Libellé rayon").fillna("Inconnu").astype(str).str.strip().apply(rayon_bucket)
    work["famille_label"] = _col("Libellé famille").fillna("Inconnu").astype(str).str.strip()

    work["qte_cde"] = pd.to_numeric(work.get("Qté cde"), errors="coerce").fillna(0)
    work["qte_rec"] = pd.to_numeric(work.get("Qté rec"), errors="coerce").fillna(0)
    work["pv_ht"]   = pd.to_numeric(work.get("Prix de vente HT"), errors="coerce").fillna(0)
    pv_zero_rows    = int((work["pv_ht"] == 0).sum())

    is_tech     = work["supplier_name"].str.upper().isin(TECHNICAL_SUPPLIERS)
    is_zero_qty = work["qte_cde"] <= 0
    sur_receipt = int((work["qte_rec"] > work["qte_cde"]).sum())
    excl_tech   = int(is_tech.sum())
    excl_zero   = int(is_zero_qty.sum())

    if exclude_technical:
        work = work[~is_tech].copy()
    work = work[work["qte_cde"] > 0].copy()

    work["qte_rec_retained"] = (
        work[["qte_rec", "qte_cde"]].min(axis=1) if cap_sur_receipt else work["qte_rec"]
    )

    expected_col = detect_expected_date_column(work)
    work["date_expected"] = work[expected_col] if expected_col else pd.NaT
    work["date_received"] = work["Dt Rec"] if "Dt Rec" in work.columns else pd.NaT

    missing_exp = int(work["date_expected"].isna().sum())

    work["qty_missing"]       = (work["qte_cde"] - work["qte_rec_retained"]).clip(lower=0)
    work["service_gap_value"] = work["qty_missing"] * work["pv_ht"]
    work["line_fill_rate"]    = np.where(
        work["qte_cde"] > 0, work["qte_rec_retained"] / work["qte_cde"], 0.0
    )
    work["delay_days"] = (work["date_received"] - work["date_expected"]).dt.days
    work["on_time"] = work["date_expected"].isna() | (work["date_received"] <= work["date_expected"])
    work["otif"]    = ((work["qte_rec_retained"] >= work["qte_cde"]) & work["on_time"]).astype(int)
    work["criticality_score"] = work["service_gap_value"] * (1 - work["line_fill_rate"])

    # ── FLAG SURVEILLANCE : toujours créé (valeurs neutres si pas de watchdict)
    if "Code" in work.columns:
        work["code_str"] = work["Code"].apply(normalise_code)
    else:
        work["code_str"] = ""

    if watchdict:
        work["watch_classe"] = work["code_str"].map(watchdict).fillna("")
        work["is_watched"]   = work["watch_classe"] != ""
    else:
        work["watch_classe"] = ""
        work["is_watched"]   = False

    quality = {
        "raw_rows":              raw_len,
        "clean_rows":            len(work),
        "excluded_zero_qty":     excl_zero,
        "excluded_technical":    excl_tech,
        "sur_receipt_rows":      sur_receipt,
        "missing_expected_date": missing_exp,
        "all_dates_missing":     (missing_exp == len(work)),
        "pv_zero_rows":          pv_zero_rows,
        "expected_col":          expected_col or "Aucune",
        "usable_rate":           round(safe_div(len(work), raw_len) * 100, 1),
        "watched_in_data":       int(work["is_watched"].sum()),
    }
    return work, quality


# ══════════════════════════════════════════════════════════════════════════════
# KPI GLOBAUX
# ══════════════════════════════════════════════════════════════════════════════
def compute_global_kpis(df: pd.DataFrame) -> dict:
    if df.empty:
        return dict(fill_rate=0, on_time=0, otif=0, score=0,
                    ordered_qty=0, received_qty=0, missing_qty=0,
                    impact_value=0, suppliers=0, articles=0, orders=0, sites=0)
    ordered   = df["qte_cde"].sum()
    received  = df["qte_rec_retained"].sum()
    fill_rate = safe_div(received, ordered) * 100
    on_time   = df["on_time"].mean() * 100
    otif      = df["otif"].mean() * 100
    score     = fill_rate  # Score = Taux de Service seul (Ponctualité et OTIF non retenus, cf. alerte qualité)
    return dict(
        fill_rate=fill_rate, on_time=on_time, otif=otif, score=score,
        ordered_qty=ordered, received_qty=received,
        missing_qty=df["qty_missing"].sum(),
        impact_value=df["service_gap_value"].sum(),
        suppliers=df["supplier_name"].nunique(),
        articles=df["Code"].nunique() if "Code" in df.columns else 0,
        orders=df["N° Cde"].nunique() if "N° Cde" in df.columns else 0,
        sites=df["site_label"].nunique(),
    )


# ══════════════════════════════════════════════════════════════════════════════
# AGRÉGATIONS
# ══════════════════════════════════════════════════════════════════════════════
def _enrich(g: pd.DataFrame) -> pd.DataFrame:
    g["fill_rate"] = np.where(g["qte_cde"] > 0, g["qte_rec"] / g["qte_cde"] * 100, 0.0)
    g["on_time"]  *= 100
    g["otif"]     *= 100
    g["score"]     = g["fill_rate"]  # Score = Taux de Service seul (Ponctualité et OTIF non retenus, cf. alerte qualité)
    g["Niveau"]    = g["score"].apply(score_band)
    g["criticality_score"] = g["impact_value"] * (1 - g["fill_rate"] / 100)
    return g


def agg_supplier(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["Fou", "supplier_name"], as_index=False).agg(
        qte_cde      =("qte_cde",          "sum"),
        qte_rec      =("qte_rec_retained",  "sum"),
        qty_missing  =("qty_missing",       "sum"),
        impact_value =("service_gap_value", "sum"),
        on_time      =("on_time",           "mean"),
        otif         =("otif",              "mean"),
        orders       =("N° Cde",            "nunique"),
        articles     =("Code",              "nunique"),
        sites        =("site_label",        "nunique"),
    )
    return _enrich(g).sort_values("criticality_score", ascending=False).reset_index(drop=True)


def agg_site(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["Site", "site_label"], as_index=False).agg(
        qte_cde      =("qte_cde",          "sum"),
        qte_rec      =("qte_rec_retained",  "sum"),
        qty_missing  =("qty_missing",       "sum"),
        impact_value =("service_gap_value", "sum"),
        on_time      =("on_time",           "mean"),
        otif         =("otif",              "mean"),
        suppliers    =("supplier_name",     "nunique"),
        articles     =("Code",             "nunique"),
    )
    return _enrich(g).sort_values("criticality_score", ascending=False).reset_index(drop=True)


def agg_article(df: pd.DataFrame, watchdict: dict = None) -> pd.DataFrame:
    gcols = [c for c in ["Code", "article_label", "supplier_name"] if c in df.columns]
    g = df.groupby(gcols, as_index=False).agg(
        qte_cde      =("qte_cde",          "sum"),
        qte_rec      =("qte_rec_retained",  "sum"),
        qty_missing  =("qty_missing",       "sum"),
        impact_value =("service_gap_value", "sum"),
        on_time      =("on_time",           "mean"),
        otif         =("otif",              "mean"),
        sites        =("site_label",        "nunique"),
        orders       =("N° Cde",            "nunique"),
    )
    g = _enrich(g).sort_values("criticality_score", ascending=False).reset_index(drop=True)

    if watchdict and "Code" in g.columns:
        g["code_str"]  = g["Code"].apply(normalise_code)
        g["Classe"]    = g["code_str"].map(watchdict).fillna("")
    else:
        g["code_str"]  = g["Code"].apply(normalise_code) if "Code" in g.columns else ""
        g["Classe"]    = ""
    return g


# ══════════════════════════════════════════════════════════════════════════════
# GRAPHIQUES PLOTLY
# ══════════════════════════════════════════════════════════════════════════════
_PLOTLY_BASE = dict(
    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    font=dict(family="-apple-system, Helvetica Neue", color="#3A3A3C", size=11),
    margin=dict(t=10, b=10, l=10, r=80),
    xaxis=dict(showgrid=True, gridcolor="#F2F2F7"),
    yaxis=dict(showgrid=False, title=""),
)


def bar_h(data: pd.DataFrame, x_col: str, y_col: str, color: str,
          x_title: str, height: int = 500, fmt_fn=None,
          classe_col: str = None) -> go.Figure:
    top = data.head(15).sort_values(x_col)
    texts = [fmt_fn(v) if fmt_fn else f"{v:,.0f}" for v in top[x_col]]

    if classe_col and classe_col in top.columns:
        bar_colors = [
            "#FFD60A" if str(c).strip() != "" else color
            for c in top[classe_col]
        ]
    else:
        bar_colors = color

    y_labels = top[y_col].astype(str).tolist()
    if classe_col and classe_col in top.columns:
        y_labels = [
            f"[{c}] {lbl}" if str(c).strip() != "" else lbl
            for lbl, c in zip(y_labels, top[classe_col])
        ]

    fig = go.Figure(go.Bar(
        x=top[x_col], y=y_labels,
        orientation="h", marker_color=bar_colors, marker_line_width=0,
        text=texts, textposition="outside",
    ))
    fig.update_layout(**{**_PLOTLY_BASE, "height": height,
                         "xaxis": {**_PLOTLY_BASE["xaxis"], "title": x_title}})
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# COMPOSANT KPI ROW
# ══════════════════════════════════════════════════════════════════════════════
def render_kpi_row(kpi: dict, watch_kpi: dict = None):
    if watch_kpi:
        c1, c2, c3, c4 = st.columns(4)
    else:
        c1, c2, c3 = st.columns(3)

    c1.metric("Taux de Service", fmt_pct(kpi["fill_rate"]))
    c2.metric("OTIF",            fmt_pct(kpi["otif"]))
    c3.metric("Vol. manquant",   fmt(kpi["missing_qty"]))

    if watch_kpi:
        with c4:
            st.markdown(f"""
<div class='kpi-watch'>
  <div class='kpi-watch-label'>⭐ Articles surveillés</div>
  <div class='kpi-watch-value'>{fmt(watch_kpi['missing_qty'])}</div>
  <div class='kpi-watch-sub'>Vol. manquant · TS {fmt_pct(watch_kpi['fill_rate'])}</div>
</div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL — helpers couleur classe
# Note : BytesIO ici = PRODUCTION de bytes en sortie. Usage correct et voulu.
# v6.3 : _classe_fills détecte désormais par sous-chaîne ("GOLD" in c /
#        "SILVER" in c) au lieu d'une égalité stricte, pour couvrir des
#        libellés composites comme "GOLD ET PRIX PLAFONNE".
# ══════════════════════════════════════════════════════════════════════════════
GOLD_FILL    = PatternFill("solid", fgColor=WATCH_GOLD_HEX)
GOLDL_FILL   = PatternFill("solid", fgColor=WATCH_LIGHT_HEX)
SILVER_FILL  = PatternFill("solid", fgColor="E8E8E8")
SILVERL_FILL = PatternFill("solid", fgColor="F5F5F5")


def _classe_fills(classe: str):
    c = str(classe).strip().upper()
    if "GOLD" in c:
        return GOLDL_FILL, GOLD_FILL
    elif "SILVER" in c:
        return SILVERL_FILL, SILVER_FILL
    elif c in ("A", ""):
        return GOLDL_FILL, GOLD_FILL
    else:
        return SILVERL_FILL, SILVER_FILL


def _xl_write_sheet(ws, title: str, df: pd.DataFrame,
                    hdr_fill=None, hdr_font=None, ctr=None,
                    classe_col_idx: int = None):
    ws.append([title])
    ws.cell(1, 1).font = Font(bold=True, size=13, color="1C3557")
    ws.append([])
    headers = list(df.columns)
    ws.append(headers)
    for i, _ in enumerate(headers, 1):
        c = ws.cell(3, i)
        if hdr_fill: c.fill = hdr_fill
        if hdr_font: c.font = hdr_font
        if ctr:      c.alignment = ctr
    for row in df.itertuples(index=False):
        ws.append(list(row))
        if classe_col_idx is not None:
            val = str(list(row)[classe_col_idx - 1]).strip()
            if val != "":
                fill_l, _ = _classe_fills(val)
                n = ws.max_row
                for col_i in range(1, len(headers) + 1):
                    ws.cell(n, col_i).fill = fill_l
    for col_cells in ws.iter_cols(min_row=3, max_row=3):
        letter = get_column_letter(col_cells[0].column)
        val = str(col_cells[0].value or "")
        ws.column_dimensions[letter].width = max(12, min(32, len(val) + 4))


def build_export_excel(df, by_supplier, by_site, by_article, quality,
                       watchdict: dict = None) -> BytesIO:
    wb = Workbook()
    H_FILL = PatternFill("solid", fgColor="1C3557")
    H_FONT = Font(bold=True, color="FFFFFF", size=11)
    CTR    = Alignment(horizontal="center", vertical="center")

    ws1 = wb.active
    ws1.title = "Synthese"
    synthese = pd.DataFrame([
        ["Lignes brutes",                      quality.get("raw_rows", 0)],
        ["Lignes exploitables",                quality.get("clean_rows", 0)],
        ["Taux exploitable %",                 quality.get("usable_rate", 0)],
        ["Date prévue utilisée",               quality.get("expected_col", "Aucune")],
        ["Toutes dates prévues manquantes",     "OUI" if quality.get("all_dates_missing") else "NON"],
        ["Qté cde ≤ 0 exclues",                quality.get("excluded_zero_qty", 0)],
        ["Fournisseurs techniques exclus",      quality.get("excluded_technical", 0)],
        ["Sur-réceptions capées",              quality.get("sur_receipt_rows", 0)],
        ["Dates prévues manquantes",            quality.get("missing_expected_date", 0)],
        ["Lignes PV HT = 0 (impact nul)",      quality.get("pv_zero_rows", 0)],
        ["Articles surveillance matchés",       quality.get("watched_in_data", 0)],
    ], columns=["Indicateur", "Valeur"])
    _xl_write_sheet(ws1, "Synthèse qualité de données", synthese, H_FILL, H_FONT, CTR)

    by_article_export = by_article.drop(columns=["code_str"], errors="ignore")
    art_cols_export = list(by_article_export.columns)
    classe_idx_export = (art_cols_export.index("Classe") + 1) if "Classe" in art_cols_export else None

    _xl_write_sheet(wb.create_sheet("Par fournisseur"), "Analyse fournisseur", by_supplier, H_FILL, H_FONT, CTR)
    _xl_write_sheet(wb.create_sheet("Par magasin"),     "Analyse magasin",     by_site,     H_FILL, H_FONT, CTR)
    _xl_write_sheet(wb.create_sheet("Par article"),     "Analyse article",     by_article_export,  H_FILL, H_FONT, CTR,
                    classe_col_idx=classe_idx_export)

    if watchdict and "Classe" in by_article_export.columns:
        watched_art = by_article_export[by_article_export["Classe"] != ""].copy()
        if not watched_art.empty:
            ws_w = wb.create_sheet("⭐ Articles surveillés")
            _xl_write_sheet(ws_w, "Articles de la liste surveillance", watched_art, H_FILL, H_FONT, CTR,
                            classe_col_idx=classe_idx_export)
            for i in range(1, len(watched_art.columns) + 1):
                c = ws_w.cell(3, i)
                c.fill = PatternFill("solid", fgColor=WATCH_GOLD_HEX)
                c.font = Font(bold=True, color="1C1C1E", size=11)

    # Lignes critiques — is_watched inclus pour tri, retiré de l'export
    detail_cols = [c for c in [
        "date_received", "date_expected", "site_label", "supplier_name",
        "Code", "article_label", "N° Cde", "qte_cde", "qte_rec_retained",
        "qty_missing", "service_gap_value", "on_time", "otif", "delay_days",
        "watch_classe", "is_watched",
    ] if c in df.columns]

    crit = df[df["otif"] == 0][detail_cols].copy()
    crit = safe_sort(crit,
                     keys=["is_watched", "qty_missing", "service_gap_value"],
                     ascending=[False, False, False])
    crit = crit.head(500)

    if "is_watched" in crit.columns:
        crit = crit.drop(columns=["is_watched"])
    if "watch_classe" in crit.columns:
        crit = crit.rename(columns={"watch_classe": "Classe"})

    crit_classe_idx = (list(crit.columns).index("Classe") + 1) if "Classe" in crit.columns else None
    _xl_write_sheet(wb.create_sheet("Lignes critiques"), "Lignes non OTIF", crit, H_FILL, H_FONT, CTR,
                    classe_col_idx=crit_classe_idx)

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL — FICHE FOURNISSEUR
# ══════════════════════════════════════════════════════════════════════════════
def build_fiche_excel(fournisseur: str, df_all: pd.DataFrame, df_bad: pd.DataFrame,
                      site_recap: pd.DataFrame, art_recap: pd.DataFrame,
                      kpis: dict, seuil: int, watchdict: dict = None) -> BytesIO:
    wb = Workbook()
    today_str = _date.today().strftime("%d/%m/%Y")

    H_FILL   = PatternFill("solid", fgColor="1C3557")
    H_FONT   = Font(bold=True, color="FFFFFF", size=10)
    T_FONT   = Font(bold=True, size=13, color="1C3557")
    S_FONT   = Font(bold=True, size=10)
    CTR      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LFT      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    RED_FILL = PatternFill("solid", fgColor="FFF2F2")
    AMB_FILL = PatternFill("solid", fgColor="FFFBF0")
    GRN_FILL = PatternFill("solid", fgColor="F0FFF4")
    THIN     = Border(
        left=Side(style="thin", color="E5E5EA"),
        right=Side(style="thin", color="E5E5EA"),
        top=Side(style="thin", color="E5E5EA"),
        bottom=Side(style="thin", color="E5E5EA"),
    )

    def auto_width(ws, min_w=12, max_w=34):
        for col in ws.iter_cols():
            best = min_w
            for cell in col:
                if cell.value:
                    best = max(best, min(max_w, len(str(cell.value)) + 3))
            ws.column_dimensions[get_column_letter(col[0].column)].width = best

    def write_header_row(ws, row_n, headers, gold=False):
        for i, h in enumerate(headers, 1):
            c = ws.cell(row_n, i, h)
            c.fill = PatternFill("solid", fgColor=WATCH_GOLD_HEX) if gold else H_FILL
            c.font = Font(bold=True, color="1C1C1E" if gold else "FFFFFF", size=10)
            c.alignment = CTR; c.border = THIN

    def kpi_fill(val_str: str):
        try:
            v = float(str(val_str).replace("%", ""))
            return GRN_FILL if v >= SEUIL_EXCELLENT else (AMB_FILL if v >= SEUIL_SURVEILLER else RED_FILL)
        except Exception:
            return None

    def get_code_classe(code) -> str:
        if not watchdict:
            return ""
        return watchdict.get(normalise_code(code), "")

    # Onglet 1 : Synthèse KPI
    ws1 = wb.active
    ws1.title = "Synthèse KPI"
    ws1.row_dimensions[1].height = 30
    ws1.row_dimensions[2].height = 18
    ws1.cell(1, 1, f"BILAN DE PERFORMANCE — {fournisseur.upper()}").font = Font(bold=True, size=14, color="1C1C1E")
    ws1.cell(2, 1, f"Généré le {today_str}  ·  Période : export ERP  ·  Seuil : Taux de Service < {seuil}%").font = Font(size=10, italic=True, color="8E8E93")
    ws1.merge_cells("A1:D1"); ws1.merge_cells("A2:D2")
    ws1.append([])
    kpi_data = [
        ("Taux de Service",       fmt_pct(kpis["fill_rate"]),    "Part de la qté commandée effectivement livrée"),
        ("Ponctualité",         fmt_pct(kpis["on_time"]),      "Part des livraisons respectant la date prévue"),
        ("OTIF",            fmt_pct(kpis["otif"]),         "Livraisons complètes ET à l'heure"),
        ("Score global",    fmt_pct(kpis["score"]),        "50% × Taux de Service + 30% × Ponctualité + 20% × OTIF"),
        ("Niveau",          score_band(kpis["score"]),     "Évaluation synthétique"),
        ("Vol. manquant",   f"{int(kpis['missing_qty']):,} unités",  "Qté commandée − Qté reçue retenue"),
        ("Impact CA proxy", f"{kpis['impact_value']:,.0f} FCFA",     "Vol. manquant × Prix de vente HT"),
        ("Commandes",       str(kpis["orders"]),           "Nombre de N° de commandes distincts"),
        ("Articles",        str(kpis["articles"]),         "Références concernées"),
        ("Magasins",        str(kpis["sites"]),            "Sites livrés sur la période"),
    ]
    write_header_row(ws1, 4, ["Indicateur", "Valeur", "Commentaire"])
    for i, (label, val, comment) in enumerate(kpi_data, 5):
        ws1.cell(i, 1, label).font = S_FONT; ws1.cell(i, 1).border = THIN
        c_val = ws1.cell(i, 2, val)
        c_val.alignment = CTR; c_val.border = THIN
        if label in ("Taux de Service", "Ponctualité", "OTIF", "Score global"):
            f = kpi_fill(val)
            if f: c_val.fill = f
        ws1.cell(i, 3, comment).alignment = LFT; ws1.cell(i, 3).border = THIN
    auto_width(ws1)

    # Onglet 2 : Par magasin
    ws2 = wb.create_sheet("Par magasin")
    ws2.cell(1, 1, f"Livraisons incomplètes par magasin — seuil Taux de Service < {seuil}%").font = T_FONT
    ws2.append([])
    headers2 = ["Magasin", "Taux de Service", "Vol. manquant", "Impact CA proxy (FCFA)", "Cmdes", "Articles"]
    write_header_row(ws2, 3, headers2)
    site_recap_sorted = safe_sort(site_recap, keys=["qty_missing"], ascending=[False])
    for _, r in site_recap_sorted.iterrows():
        ws2.append([r.get("site_label",""), r.get("Taux de Service",""), r.get("Vol. manquant",""),
                    r.get("Impact CA proxy",""), r.get("nb_cdes",""), r.get("nb_articles","")])
        n = ws2.max_row
        for col_i in range(1, 7):
            ws2.cell(n, col_i).border = THIN
            ws2.cell(n, col_i).alignment = CTR if col_i > 1 else LFT
        f = kpi_fill(str(r.get("Taux de Service", ""))) if r.get("Taux de Service") else None
        if f: ws2.cell(n, 2).fill = f
    auto_width(ws2)

    # Onglet 3 : Par article
    ws3 = wb.create_sheet("Par article")
    ws3.cell(1, 1, f"Livraisons incomplètes par article — seuil Taux de Service < {seuil}%").font = T_FONT
    ws3.append([])
    headers3 = ["Classe", "Code article", "Désignation", "Taux de Service", "Vol. manquant",
                "Impact CA proxy (FCFA)", "Magasins", "Cmdes"]
    write_header_row(ws3, 3, headers3)
    for _, r in art_recap.iterrows():
        code   = r.get("Code", "")
        classe = get_code_classe(code)
        row_data = [
            classe, code, r.get("article_label",""),
            r.get("Taux de Service",""), r.get("Vol. manquant",""),
            r.get("Impact CA proxy",""), r.get("nb_sites",""), r.get("nb_cdes",""),
        ]
        n = ws3.max_row + 1
        ws3.append(row_data)
        for col_i in range(1, 9):
            ws3.cell(n, col_i).border = THIN
            ws3.cell(n, col_i).alignment = LFT if col_i in (2, 3) else CTR
        if classe:
            fill_l, fill_b = _classe_fills(classe)
            for col_i in range(1, 9):
                ws3.cell(n, col_i).fill = fill_l
            ws3.cell(n, 1).fill = fill_b
        else:
            f = kpi_fill(str(r.get("Taux de Service", ""))) if r.get("Taux de Service") else None
            if f: ws3.cell(n, 4).fill = f
    auto_width(ws3)

    # Onglet 4 : Détail lignes
    ws4 = wb.create_sheet("Détail lignes")
    ws4.cell(1, 1, f"Détail des livraisons incomplètes — {fournisseur}  (Taux de Service < {seuil}%)").font = T_FONT
    ws4.append([])
    headers4 = [
        "Classe", "N° Commande", "Date réception", "Date prévue", "Magasin",
        "Code article", "Désignation article",
        "Qté commandée", "Qté reçue", "Qté manquante",
        "Taux de Service ligne", "Impact CA proxy (FCFA)",
        "Livré à l'heure", "Retard (jours)",
    ]
    write_header_row(ws4, 3, headers4)

    src_cols = [c for c in [
        "N° Cde", "date_received", "date_expected", "site_label",
        "Code", "article_label", "qte_cde", "qte_rec_retained",
        "qty_missing", "line_fill_rate", "service_gap_value", "on_time", "delay_days",
    ] if c in df_bad.columns]

    for _, row in df_bad[src_cols].iterrows():
        fr     = row.get("line_fill_rate", 0)
        ot     = row.get("on_time", True)
        code   = row.get("Code", "")
        classe = get_code_classe(code)
        n      = ws4.max_row + 1
        ws4.append([
            classe,
            row.get("N° Cde", ""),
            row.get("date_received", pd.NaT),
            row.get("date_expected", pd.NaT),
            row.get("site_label", ""),
            code,
            row.get("article_label", ""),
            int(row.get("qte_cde", 0)),
            int(row.get("qte_rec_retained", 0)),
            int(row.get("qty_missing", 0)),
            f"{fr * 100:.1f}%" if pd.notna(fr) else "—",
            round(row.get("service_gap_value", 0), 0),
            "OUI" if ot else "NON",
            int(row.get("delay_days", 0)) if pd.notna(row.get("delay_days")) else "—",
        ])
        for col_i in range(1, len(headers4) + 1):
            cell = ws4.cell(n, col_i)
            cell.border = THIN
            cell.alignment = LFT if col_i in (5, 7) else CTR
        if classe:
            fill_l, fill_b = _classe_fills(classe)
            for col_i in range(1, len(headers4) + 1):
                ws4.cell(n, col_i).fill = fill_l
            ws4.cell(n, 1).fill = fill_b
        else:
            row_fill = RED_FILL if pd.notna(fr) and fr < 0.90 else None
            if row_fill:
                for col_i in (10, 11, 12):
                    ws4.cell(n, col_i).fill = row_fill
            fr_fill = kpi_fill(f"{fr*100:.1f}%") if pd.notna(fr) else None
            if fr_fill:
                ws4.cell(n, 11).fill = fr_fill

    # Onglet 5 : Articles surveillés uniquement
    if watchdict:
        df_bad_watched = df_bad[df_bad["is_watched"]] if "is_watched" in df_bad.columns else df_bad.iloc[0:0]
        if not df_bad_watched.empty:
            ws5 = wb.create_sheet("⭐ Articles surveillés")
            ws5.cell(1, 1, f"Articles de surveillance — {fournisseur}").font = T_FONT
            ws5.cell(2, 1, f"Filtre : Taux de Service < {seuil}% · liste surveillance active").font = Font(size=9, italic=True, color="8E8E93")
            ws5.append([])
            write_header_row(ws5, 4, headers4, gold=True)
            for _, row in df_bad_watched[src_cols].iterrows():
                fr     = row.get("line_fill_rate", 0)
                ot     = row.get("on_time", True)
                code   = row.get("Code", "")
                classe = get_code_classe(code)
                n      = ws5.max_row + 1
                ws5.append([
                    classe,
                    row.get("N° Cde", ""),
                    row.get("date_received", pd.NaT),
                    row.get("date_expected", pd.NaT),
                    row.get("site_label", ""),
                    code,
                    row.get("article_label", ""),
                    int(row.get("qte_cde", 0)),
                    int(row.get("qte_rec_retained", 0)),
                    int(row.get("qty_missing", 0)),
                    f"{fr * 100:.1f}%" if pd.notna(fr) else "—",
                    round(row.get("service_gap_value", 0), 0),
                    "OUI" if ot else "NON",
                    int(row.get("delay_days", 0)) if pd.notna(row.get("delay_days")) else "—",
                ])
                for col_i in range(1, len(headers4) + 1):
                    cell = ws5.cell(n, col_i)
                    cell.border = THIN
                    cell.alignment = LFT if col_i in (5, 7) else CTR
                fill_l, fill_b = _classe_fills(classe)
                for col_i in range(1, len(headers4) + 1):
                    ws5.cell(n, col_i).fill = fill_l
                ws5.cell(n, 1).fill = fill_b
                fr_fill = kpi_fill(f"{fr*100:.1f}%") if pd.notna(fr) else None
                if fr_fill:
                    ws5.cell(n, 11).fill = fr_fill
            auto_width(ws5)
            ws5.freeze_panes = "A5"

    for ws in [ws2, ws3, ws4]:
        ws.freeze_panes = "A4"

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL — TOUTES LES FICHES
# ══════════════════════════════════════════════════════════════════════════════
def build_export_all_fiches(df: pd.DataFrame, by_supplier: pd.DataFrame,
                             seuil: int, watchdict: dict = None) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Toutes livraisons incomplètes"
    today_str = _date.today().strftime("%d/%m/%Y")

    H_FILL    = PatternFill("solid", fgColor="1C3557")
    H_FONT    = Font(bold=True, color="FFFFFF", size=10)
    T_FONT    = Font(bold=True, size=13, color="1C3557")
    GRN_FILL  = PatternFill("solid", fgColor="E8F8EE")
    AMB_FILL  = PatternFill("solid", fgColor="FFF8E8")
    RED_FILL  = PatternFill("solid", fgColor="FFF0F0")
    CTR       = Alignment(horizontal="center", vertical="center")
    LFT       = Alignment(horizontal="left",   vertical="center")
    THIN_SIDE = Side(style="thin", color="E5E5EA")
    THIN      = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

    ws.cell(1, 1, "LIVRAISONS INCOMPLÈTES — TOUS FOURNISSEURS").font = T_FONT
    ws.cell(2, 1, (
        f"Généré le {today_str}  ·  "
        f"Seuil : Taux de Service < {seuil}% ou OTIF = 0  ·  "
        f"Trié par criticité fournisseur"
        + ("  ·  Classe = segment 20/80" if watchdict else "")
    )).font = Font(size=9, italic=True, color="8E8E93")
    ws.merge_cells("A1:R1"); ws.merge_cells("A2:R2")
    ws.append([])

    HEADERS = [
        "Classe", "Fournisseur", "Code Fou.", "Score four.", "Niveau",
        "N° Commande", "Date réception", "Date prévue",
        "Magasin", "Code article", "Désignation article",
        "Qté commandée", "Qté reçue", "Qté manquante",
        "Taux de Service ligne", "Impact CA proxy (FCFA)",
        "OTIF", "Retard (j)",
    ]
    ws.append(HEADERS)
    hdr_row = 4
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(hdr_row, i)
        c.value = h; c.fill = H_FILL; c.font = H_FONT
        c.alignment = CTR; c.border = THIN

    sup_meta = (
        by_supplier[["supplier_name", "Fou", "score", "Niveau", "criticality_score"]]
        .copy()
        .rename(columns={"Fou": "fou_code", "criticality_score": "crit_sup"})
    )
    mask = (df["line_fill_rate"] < seuil / 100) | (df["otif"] == 0)
    export_df = df[mask].copy()
    export_df = export_df.merge(sup_meta, on="supplier_name", how="left")

    export_df = safe_sort(
        export_df,
        keys=["is_watched", "crit_sup", "qty_missing", "service_gap_value"],
        ascending=[False, False, False, False],
    )

    for _, row in export_df.iterrows():
        fr      = row.get("line_fill_rate", np.nan)
        ot      = row.get("otif", 0)
        dly     = row.get("delay_days", np.nan)
        classe  = str(row.get("watch_classe", "")).strip()

        row_data = [
            classe,
            row.get("supplier_name", ""),
            int(row.get("fou_code", 0)) if pd.notna(row.get("fou_code")) else "",
            f"{row.get('score', 0):.1f}%",
            row.get("Niveau", ""),
            row.get("N° Cde", ""),
            row.get("date_received", pd.NaT),
            row.get("date_expected", pd.NaT),
            row.get("site_label", ""),
            row.get("Code", ""),
            row.get("article_label", ""),
            int(row.get("qte_cde", 0)),
            int(row.get("qte_rec_retained", 0)),
            int(row.get("qty_missing", 0)),
            f"{fr * 100:.1f}%" if pd.notna(fr) else "—",
            round(row.get("service_gap_value", 0), 0),
            "OUI" if ot == 1 else "NON",
            int(dly) if pd.notna(dly) else "—",
        ]
        n = ws.max_row + 1
        ws.append(row_data)

        for col_i in range(1, len(HEADERS) + 1):
            cell = ws.cell(n, col_i)
            cell.border = THIN
            cell.alignment = LFT if col_i in (2, 5, 9, 11) else CTR

        if classe:
            fill_l, fill_b = _classe_fills(classe)
            for col_i in range(1, len(HEADERS) + 1):
                ws.cell(n, col_i).fill = fill_l
            ws.cell(n, 1).fill = fill_b
        else:
            if pd.notna(fr):
                fr_pct = fr * 100
                fill_fr = GRN_FILL if fr_pct >= 97 else (AMB_FILL if fr_pct >= 90 else RED_FILL)
                ws.cell(n, 15).fill = fill_fr
            niv = row.get("Niveau", "")
            if "Excellent"    in str(niv): ws.cell(n, 5).fill = GRN_FILL
            elif "surveiller" in str(niv): ws.cell(n, 5).fill = AMB_FILL
            elif "Critique"   in str(niv): ws.cell(n, 5).fill = RED_FILL
            ws.cell(n, 17).fill = GRN_FILL if ot == 1 else RED_FILL

    if watchdict:
        ws_w = wb.create_sheet("⭐ Articles surveillés")
        ws_w.cell(1, 1, "ARTICLES SURVEILLANCE — TOUS FOURNISSEURS").font = T_FONT
        ws_w.merge_cells("A1:R1")
        ws_w.append([])
        ws_w.append(HEADERS)
        watch_hdr_row = 3
        for i, h in enumerate(HEADERS, 1):
            c = ws_w.cell(watch_hdr_row, i)
            c.value = h
            c.fill = PatternFill("solid", fgColor=WATCH_GOLD_HEX)
            c.font = Font(bold=True, color="1C1C1E", size=10)
            c.alignment = CTR; c.border = THIN

        watched_rows = export_df[export_df["is_watched"] == True] if "is_watched" in export_df.columns else export_df.iloc[0:0]
        for _, row in watched_rows.iterrows():
            fr     = row.get("line_fill_rate", np.nan)
            ot     = row.get("otif", 0)
            dly    = row.get("delay_days", np.nan)
            classe = str(row.get("watch_classe", "")).strip()
            ws_w.append([
                classe,
                row.get("supplier_name", ""),
                int(row.get("fou_code", 0)) if pd.notna(row.get("fou_code")) else "",
                f"{row.get('score', 0):.1f}%",
                row.get("Niveau", ""),
                row.get("N° Cde", ""),
                row.get("date_received", pd.NaT),
                row.get("date_expected", pd.NaT),
                row.get("site_label", ""),
                row.get("Code", ""),
                row.get("article_label", ""),
                int(row.get("qte_cde", 0)),
                int(row.get("qte_rec_retained", 0)),
                int(row.get("qty_missing", 0)),
                f"{fr * 100:.1f}%" if pd.notna(fr) else "—",
                round(row.get("service_gap_value", 0), 0),
                "OUI" if ot == 1 else "NON",
                int(dly) if pd.notna(dly) else "—",
            ])
            n = ws_w.max_row
            for col_i in range(1, len(HEADERS) + 1):
                cell = ws_w.cell(n, col_i)
                cell.border = THIN
                cell.alignment = LFT if col_i in (2, 5, 9, 11) else CTR
            fill_l, fill_b = _classe_fills(classe)
            for col_i in range(1, len(HEADERS) + 1):
                ws_w.cell(n, col_i).fill = fill_l
            ws_w.cell(n, 1).fill = fill_b
            if pd.notna(fr):
                fr_pct = fr * 100
                fill_fr = GRN_FILL if fr_pct >= 97 else (AMB_FILL if fr_pct >= 90 else RED_FILL)
                ws_w.cell(n, 15).fill = fill_fr

        ws_w.freeze_panes = "A4"
        col_widths = [10, 28, 10, 11, 14, 14, 16, 14, 22, 13, 32, 13, 11, 14, 14, 22, 8, 11]
        for i, w in enumerate(col_widths, 1):
            ws_w.column_dimensions[get_column_letter(i)].width = w

    col_widths = [10, 28, 10, 11, 14, 14, 16, 14, 22, 13, 32, 13, 11, 14, 14, 22, 8, 11]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(len(HEADERS))}{ws.max_row}"
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[hdr_row].height = 28

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL — RÉCAP PRIORISATION FOURNISSEURS
# ══════════════════════════════════════════════════════════════════════════════
def build_recap_fournisseur_excel(df: pd.DataFrame, watchdict: dict = None) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Récap priorisation fournisseurs"
    today_str = _date.today().strftime("%d/%m/%Y")

    H_FILL    = PatternFill("solid", fgColor="1C3557")
    H_FONT    = Font(bold=True, color="FFFFFF", size=10)
    T_FONT    = Font(bold=True, size=13, color="1C3557")
    GRN_FILL  = PatternFill("solid", fgColor="E8F8EE")
    AMB_FILL  = PatternFill("solid", fgColor="FFF8E8")
    RED_FILL  = PatternFill("solid", fgColor="FFF0F0")
    CTR       = Alignment(horizontal="center", vertical="center")
    LFT       = Alignment(horizontal="left",   vertical="center")
    THIN_SIDE = Side(style="thin", color="E5E5EA")
    THIN      = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

    total_cols = 19 if watchdict else 17
    ws.cell(1, 1, "RÉCAP PRIORISATION FOURNISSEURS — OTIF").font = T_FONT
    ws.cell(2, 1,
        f"Généré le {today_str}  ·  "
        "Trié par criticité décroissante  ·  "
        "Réfs livrées = au moins 1 unité reçue sur la période"
        + ("  ·  Cols GOLD/SILVER = nb réfs en sous-service (TS < 97%)" if watchdict else "")
    ).font = Font(size=9, italic=True, color="8E8E93")
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    ws.merge_cells(f"A2:{get_column_letter(total_cols)}2")
    ws.append([])

    HEADERS = [
        "Rang", "Fournisseur", "Code Fou.",
        "Réfs commandées", "Réfs livrées", "Taux couverture réfs",
        "Taux de Service", "Ponctualité", "OTIF", "Score", "Niveau",
        "Sites impactés", "Sites total",
        "Vol. manquant", "Impact CA proxy (FCFA)",
        "Nb commandes", "Nb lignes",
    ]
    if watchdict:
        HEADERS += ["Réfs GOLD sous-service", "Réfs SILVER sous-service"]

    ws.append(HEADERS)
    hdr_row = 4
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(hdr_row, i)
        c.value = h; c.fill = H_FILL; c.font = H_FONT
        c.alignment = CTR; c.border = THIN
    ws.row_dimensions[hdr_row].height = 30

    grp = df.groupby(["Fou", "supplier_name"], as_index=False)
    agg = grp.agg(
        refs_total    =("Code",              "nunique"),
        qte_cde_sum   =("qte_cde",           "sum"),
        qte_rec_sum   =("qte_rec_retained",  "sum"),
        qty_missing   =("qty_missing",        "sum"),
        impact_value  =("service_gap_value",  "sum"),
        on_time_mean  =("on_time",            "mean"),
        otif_mean     =("otif",               "mean"),
        nb_orders     =("N° Cde",             "nunique"),
        nb_lines      =("Code",               "count"),
        sites_total   =("site_label",         "nunique"),
    )

    refs_livrees = (
        df[df["qte_rec_retained"] > 0]
        .groupby(["Fou", "supplier_name"], as_index=False)["Code"]
        .nunique()
        .rename(columns={"Code": "refs_livrees"})
    )
    agg = agg.merge(refs_livrees, on=["Fou", "supplier_name"], how="left")
    agg["refs_livrees"] = agg["refs_livrees"].fillna(0).astype(int)

    sites_impactes = (
        df[df["line_fill_rate"] < 1.0]
        .groupby(["Fou", "supplier_name"], as_index=False)["site_label"]
        .nunique()
        .rename(columns={"site_label": "sites_impactes"})
    )
    agg = agg.merge(sites_impactes, on=["Fou", "supplier_name"], how="left")
    agg["sites_impactes"] = agg["sites_impactes"].fillna(0).astype(int)

    agg["fill_rate"]       = np.where(agg["qte_cde_sum"] > 0, agg["qte_rec_sum"] / agg["qte_cde_sum"] * 100, 0.0)
    agg["on_time_pct"]     = agg["on_time_mean"] * 100
    agg["otif_pct"]        = agg["otif_mean"]    * 100
    agg["score"]           = agg["fill_rate"]  # Score = Taux de Service seul (Ponctualité et OTIF non retenus, cf. alerte qualité)
    agg["taux_couverture"] = np.where(agg["refs_total"] > 0, agg["refs_livrees"] / agg["refs_total"] * 100, 0.0)
    agg["criticality"]     = agg["impact_value"] * (1 - agg["fill_rate"] / 100)
    agg                    = agg.sort_values("criticality", ascending=False).reset_index(drop=True)
    agg["rang"]            = agg.index + 1

    if watchdict:
        df_w = df.copy()
        df_w["watch_c"] = df_w["Code"].apply(normalise_code).map(watchdict).fillna("")
        art_fr = (
            df_w.groupby(["Fou", "supplier_name", "Code", "watch_c"], as_index=False)
            .agg(qte_c=("qte_cde","sum"), qte_r=("qte_rec_retained","sum"))
        )
        art_fr["fr"] = np.where(art_fr["qte_c"] > 0, art_fr["qte_r"] / art_fr["qte_c"] * 100, 0.0)
        art_under = art_fr[art_fr["fr"] < 97]

        # v6.3 : détection GOLD/SILVER par sous-chaîne (couvre "GOLD ET PRIX PLAFONNE")
        gold_under = (
            art_under[art_under["watch_c"].str.upper().str.contains("GOLD", na=False)]
            .groupby(["Fou", "supplier_name"])["Code"].nunique()
            .reset_index().rename(columns={"Code": "refs_gold_under"})
        )
        silver_under = (
            art_under[art_under["watch_c"].str.upper().str.contains("SILVER", na=False)]
            .groupby(["Fou", "supplier_name"])["Code"].nunique()
            .reset_index().rename(columns={"Code": "refs_silver_under"})
        )
        agg = agg.merge(gold_under,   on=["Fou", "supplier_name"], how="left")
        agg = agg.merge(silver_under, on=["Fou", "supplier_name"], how="left")
        agg["refs_gold_under"]   = agg["refs_gold_under"].fillna(0).astype(int)
        agg["refs_silver_under"] = agg["refs_silver_under"].fillna(0).astype(int)

    def _niveau(score):
        if score >= SEUIL_EXCELLENT:  return "Excellent"
        if score >= SEUIL_SURVEILLER: return "À surveiller"
        return "Critique"

    def _score_fill(score):
        if score >= SEUIL_EXCELLENT:  return GRN_FILL
        if score >= SEUIL_SURVEILLER: return AMB_FILL
        return RED_FILL

    for _, row in agg.iterrows():
        score    = row["score"]
        fr       = row["fill_rate"]
        ot       = row["on_time_pct"]
        otif     = row["otif_pct"]
        taux_cov = row["taux_couverture"]

        data = [
            int(row["rang"]),
            row["supplier_name"],
            int(row["Fou"]) if pd.notna(row["Fou"]) else "",
            int(row["refs_total"]),
            int(row["refs_livrees"]),
            f"{taux_cov:.1f}%",
            f"{fr:.1f}%",
            f"{ot:.1f}%",
            f"{otif:.1f}%",
            f"{score:.1f}%",
            _niveau(score),
            int(row["sites_impactes"]),
            int(row["sites_total"]),
            int(row["qty_missing"]),
            round(row["impact_value"], 0),
            int(row["nb_orders"]),
            int(row["nb_lines"]),
        ]
        if watchdict:
            data += [int(row["refs_gold_under"]), int(row["refs_silver_under"])]

        n = ws.max_row + 1
        ws.append(data)

        for col_i in range(1, len(HEADERS) + 1):
            cell = ws.cell(n, col_i)
            cell.border = THIN
            cell.alignment = LFT if col_i == 2 else CTR

        ws.cell(n, 7).fill  = _score_fill(fr)
        ws.cell(n, 9).fill  = _score_fill(otif)
        ws.cell(n, 10).fill = _score_fill(score)
        ws.cell(n, 11).fill = _score_fill(score)
        ws.cell(n, 6).fill  = _score_fill(taux_cov)

        if watchdict:
            if int(row["refs_gold_under"]) > 0:
                ws.cell(n, 18).fill = PatternFill("solid", fgColor=WATCH_GOLD_HEX)
                ws.cell(n, 18).font = Font(bold=True, color="1C1C1E", size=10)
            if int(row["refs_silver_under"]) > 0:
                ws.cell(n, 19).fill = PatternFill("solid", fgColor="E8E8E8")

    col_widths = [6, 30, 10, 16, 14, 20, 12, 12, 10, 10, 16, 14, 12, 16, 22, 14, 12]
    if watchdict:
        col_widths += [18, 20]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(len(HEADERS))}{ws.max_row}"
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 14

    ws_leg = wb.create_sheet("Légende")
    ws_leg.cell(1, 1, "LÉGENDE — Récap priorisation fournisseurs").font = Font(bold=True, size=12, color="1C3557")
    leg_data = [
        ("Réfs commandées",         "Nombre de codes articles distincts présents dans au moins 1 commande sur la période"),
        ("Réfs livrées",            "Nombre de codes articles avec au moins 1 unité reçue (qte_rec > 0)"),
        ("Taux couverture réfs",    "Réfs livrées / Réfs commandées × 100 — mesure la largeur de l'assortiment effectivement approvisionné"),
        ("Taux de Service",               "Qté reçue retenue / Qté commandée × 100 — mesure la profondeur du service"),
        ("Ponctualité",                 "% de lignes livrées à la date prévue ou avant"),
        ("OTIF",                    "% de lignes complètes ET à l'heure"),
        ("Score",                   "50% × Taux de Service + 30% × Ponctualité + 20% × OTIF"),
        ("Niveau",                  "Excellent ≥ 97% · À surveiller 90–97% · Critique < 90%"),
        ("Sites impactés",          "Nb de magasins avec au moins une ligne Taux de Service < 100%"),
        ("Vol. manquant",           "Somme (Qté commandée − Qté reçue retenue) pour toutes les lignes"),
        ("Impact CA proxy",         "Vol. manquant × Prix de vente HT — estimation du CA non réalisé"),
        ("Réfs GOLD sous-service",  "Nb de réfs classées GOLD (watchlist, y compris variantes comme 'GOLD ET PRIX PLAFONNE') avec Taux de Service < 97% chez ce fournisseur"),
        ("Réfs SILVER sous-service","Nb de réfs SILVER (watchlist) avec Taux de Service < 97% chez ce fournisseur"),
    ]
    ws_leg.append([])
    ws_leg.append(["Indicateur", "Définition"])
    for i in range(1, 3):
        ws_leg.cell(3, i).fill = H_FILL
        ws_leg.cell(3, i).font = H_FONT
        ws_leg.cell(3, i).alignment = CTR
        ws_leg.cell(3, i).border = THIN
    for label, defn in leg_data:
        ws_leg.append([label, defn])
        n = ws_leg.max_row
        ws_leg.cell(n, 1).font = Font(bold=True, size=10)
        ws_leg.cell(n, 1).border = THIN
        ws_leg.cell(n, 2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws_leg.cell(n, 2).border = THIN
        ws_leg.row_dimensions[n].height = 22
    ws_leg.column_dimensions["A"].width = 26
    ws_leg.column_dimensions["B"].width = 72
    ws_leg.freeze_panes = "A4"

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL — PLAN D'ACTION OTIF (vue Responsable Achat)
# 4 onglets, zéro formule (règle SmartBuyer Hub), couleurs CFAO Consumer,
# police Segoe UI appliquée à toutes les cellules en fin de génération.
# ══════════════════════════════════════════════════════════════════════════════
PA_H_FILL      = PatternFill("solid", fgColor=CFAO_NAVY)
PA_H_FONT      = Font(name="Segoe UI", bold=True, color="FFFFFF", size=10)
PA_T_FONT      = Font(name="Segoe UI", bold=True, size=13, color=CFAO_NAVY)
PA_SUB_FONT    = Font(name="Segoe UI", size=9, italic=True, color="8E8E93")
PA_BASE_FONT   = Font(name="Segoe UI", size=10)
PA_BOLD_FONT   = Font(name="Segoe UI", bold=True, size=10)
PA_DQ_FONT     = Font(name="Segoe UI", size=9, italic=True, color="8E8E93")
PA_DQ_TITLE    = Font(name="Segoe UI", bold=True, size=10, color="A32D2D")
PA_INPUT_FILL  = PatternFill("solid", fgColor="FFFDE7")
PA_GRN_FILL    = PatternFill("solid", fgColor="E8F8EE")
PA_AMB_FILL    = PatternFill("solid", fgColor="FFF8E8")
PA_RED_FILL    = PatternFill("solid", fgColor="FFF0F0")
PA_GOLD_FILL   = PatternFill("solid", fgColor=WATCH_GOLD_HEX)
PA_SILVER_FILL = PatternFill("solid", fgColor="E8E8E8")
PA_GOLD_ROW    = PatternFill("solid", fgColor="FFF9C4")
PA_SILVER_ROW  = PatternFill("solid", fgColor="EFEFEF")
PA_ALT_FILL    = PatternFill("solid", fgColor="F7F7F5")
PA_NODATA_FONT = Font(name="Segoe UI", color="C7C7CC", size=10)
PA_CTR   = Alignment(horizontal="center", vertical="center")
PA_LFT   = Alignment(horizontal="left", vertical="center", wrap_text=True)
PA_THIN_SIDE = Side(style="thin", color="E5E5EA")
PA_THIN  = Border(left=PA_THIN_SIDE, right=PA_THIN_SIDE, top=PA_THIN_SIDE, bottom=PA_THIN_SIDE)
PA_DATE_FMT = "dd/mm/yyyy"


def _pa_score_fill(v):
    if v >= SEUIL_EXCELLENT:
        return PA_GRN_FILL
    if v >= SEUIL_SURVEILLER:
        return PA_AMB_FILL
    return PA_RED_FILL


def _pa_auto_width(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _pa_write_title(ws, title, subtitle, n_cols):
    ws.cell(1, 1, title).font = PA_T_FONT
    ws.cell(2, 1, subtitle).font = PA_SUB_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 14


def _pa_write_header(ws, row_n, headers):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row_n, i, h)
        c.fill = PA_H_FILL
        c.font = PA_H_FONT
        c.alignment = PA_CTR
        c.border = PA_THIN
    ws.row_dimensions[row_n].height = 28


def _pa_format_key(site):
    u = str(site).upper()
    if "HYPER" in u:
        fmt = 0
    elif "MARKET" in u:
        fmt = 1
    elif "SUPECO" in u:
        fmt = 2
    else:
        fmt = 3
    return (fmt, site)


def build_plan_action_excel(df: pd.DataFrame, by_supplier: pd.DataFrame, quality: dict,
                             watchdict: dict = None) -> BytesIO:
    """
    Plan d'action OTIF pour les situations Niveau Critique — 4 onglets :
    Résumé · Suivi Fournisseurs · Taux de Service par Site · Détail Commandes.
    Zéro formule (règle SmartBuyer Hub) : tout est calculé en Python, le
    classeur produit est le résultat figé de la semaine.

    Unité d'analyse : Fournisseur × Rayon, pas Fournisseur seul. Un fournisseur
    livrant plusieurs rayons (Épicerie/Boissons/DPH) a un Taux de Service
    distinct par rayon — le blender masquerait un rayon en difficulté derrière
    la moyenne d'un autre (ex. TRANSMED SA CI : 26.8% Épicerie vs 42.7% DPH,
    contre 26.7% affiché si on blend). "Libellé département" n'est pas utilisable
    pour cette segmentation (une seule valeur "PGC" sur tout l'export ERP) ;
    "Libellé rayon" est le champ fiable, utilisé ici via rayon_label.
    """
    from datetime import datetime as _datetime
    now = _datetime.now()
    today_str = now.strftime("%d/%m/%Y %H:%M")
    watchdict = watchdict or {}

    # ── Agrégation Fournisseur × Rayon (remplace l'agrégation Fournisseur seul
    #    de by_supplier pour toute la logique de criticité de ce classeur) ──
    by_sr = df.groupby(["Fou", "supplier_name", "rayon_label"], as_index=False).agg(
        qte_cde     =("qte_cde",            "sum"),
        qte_rec     =("qte_rec_retained",    "sum"),
        qty_missing =("qty_missing",         "sum"),
        otif        =("otif",                "mean"),
        orders      =("N° Cde",              "nunique"),
        articles    =("Code",                "nunique"),
        sites       =("site_label",          "nunique"),
    )
    by_sr["fill_rate"] = np.where(by_sr["qte_cde"] > 0, by_sr["qte_rec"] / by_sr["qte_cde"] * 100, 0.0)
    by_sr["otif"]      = by_sr["otif"] * 100
    by_sr["score"]     = by_sr["fill_rate"]  # Score = Taux de Service seul (cf. alerte qualité)
    by_sr["Niveau"]    = by_sr["score"].apply(score_band)

    crit_sup = by_sr[by_sr["Niveau"] == "🔴 Critique"].sort_values(
        "qty_missing", ascending=False
    ).reset_index(drop=True)
    n_crit = len(crit_sup)
    n_crit_fou = crit_sup["Fou"].nunique()
    crit_pairs = set(zip(crit_sup["Fou"], crit_sup["rayon_label"]))

    df["_pair"] = list(zip(df["Fou"], df["rayon_label"]))
    crit_lines = df[(df["otif"] == 0) & (df["_pair"].isin(crit_pairs))].copy()
    all_crit_supplier_lines = df[df["_pair"].isin(crit_pairs)].copy()

    pair_rank = {(row["Fou"], row["rayon_label"]): idx for idx, row in crit_sup.iterrows()}
    crit_lines["_rank"] = crit_lines["_pair"].map(pair_rank)
    det = crit_lines.sort_values(
        ["_rank", "qty_missing", "service_gap_value"], ascending=[True, False, False]
    ).reset_index(drop=True)

    # Réfs GOLD/SILVER en sous-service (< 97%) par couple (Fournisseur, Rayon)
    gold_counts, silver_counts = {}, {}
    if watchdict:
        w = all_crit_supplier_lines.copy()
        w["bucket"] = w["code_str"].map(watchdict).fillna("").apply(classe_bucket)
        art_fr = w.groupby(["Fou", "rayon_label", "Code", "bucket"], as_index=False).agg(
            qte_c=("qte_cde", "sum"), qte_r=("qte_rec_retained", "sum")
        )
        art_fr["fr"] = (art_fr["qte_r"] / art_fr["qte_c"] * 100).fillna(0)
        under = art_fr[art_fr["fr"] < 97]
        gold_counts = under[under["bucket"] == "GOLD"].groupby(["Fou", "rayon_label"])["Code"].nunique().to_dict()
        silver_counts = under[under["bucket"] == "SILVER"].groupby(["Fou", "rayon_label"])["Code"].nunique().to_dict()

    # Matrice Taux de Service par Site — colonnes triées par format
    all_sites = sorted(all_crit_supplier_lines["site_label"].dropna().unique(), key=_pa_format_key)
    pivot = all_crit_supplier_lines.groupby(["Fou", "rayon_label", "site_label"], as_index=False).agg(
        qte_cde=("qte_cde", "sum"), qte_rec=("qte_rec_retained", "sum")
    )
    pivot["fr"] = (pivot["qte_rec"] / pivot["qte_cde"] * 100).round(1)
    pivot_map = {(r["Fou"], r["rayon_label"], r["site_label"]): r["fr"] for _, r in pivot.iterrows()}

    wb = Workbook()

    # ── Onglet 1 : Résumé ────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Résumé"
    ws1.sheet_properties.tabColor = CFAO_ORANGE
    _pa_write_title(ws1, "PLAN D'ACTION OTIF — RÉSUMÉ",
                     f"Généré le {today_str} · Périmètre : situations Fournisseur × Rayon Niveau Critique (Score < {SEUIL_SURVEILLER}%)", 2)
    ws1.append([])
    _pa_write_header(ws1, 4, ["Indicateur", "Valeur"])
    n_gold_matched = sum(1 for c in watchdict.values() if classe_bucket(c) == "GOLD") if watchdict else 0
    n_silver_matched = sum(1 for c in watchdict.values() if classe_bucket(c) == "SILVER") if watchdict else 0
    resume_rows = [
        ("Situations critiques (Fournisseur × Rayon)", n_crit),
        ("Fournisseurs distincts concernés", n_crit_fou),
        ("Situations à surveiller (Fournisseur × Rayon)", int((by_sr["Niveau"] == "🟠 À surveiller").sum())),
        ("Situations Excellent (Fournisseur × Rayon)", int((by_sr["Niveau"] == "🟢 Excellent").sum())),
        ("Vol. manquant — situations critiques", int(crit_lines["qty_missing"].sum())),
        ("Magasins concernés", int(crit_lines["site_label"].nunique())),
        ("Commandes concernées", int(crit_lines["N° Cde"].nunique())),
        ("Liste de surveillance — codes GOLD", n_gold_matched),
        ("Liste de surveillance — codes SILVER", n_silver_matched),
    ]
    for label, val in resume_rows:
        n = ws1.max_row + 1
        ws1.append([label, val])
        ws1.cell(n, 1).font = PA_BOLD_FONT
        ws1.cell(n, 1).border = PA_THIN
        ws1.cell(n, 2).font = PA_BASE_FONT
        ws1.cell(n, 2).alignment = PA_CTR
        ws1.cell(n, 2).border = PA_THIN

    pct_no_pv = round(quality["pv_zero_rows"] / quality["clean_rows"] * 100, 1) if quality["clean_rows"] else 0
    pct_no_date = round(quality["missing_expected_date"] / quality["clean_rows"] * 100, 1) if quality["clean_rows"] else 0
    ws1.append([])
    n = ws1.max_row + 1
    ws1.cell(n, 1, "Qualité de la donnée — à lire avant d'interpréter les chiffres").font = PA_DQ_TITLE
    ws1.merge_cells(start_row=n, start_column=1, end_row=n, end_column=2)
    dq_rows = [
        f"{pct_no_pv}% des lignes n'ont pas de Prix de Vente HT — l'Impact CA proxy est sous-estimé sur cette part et n'est pas affiché dans ce fichier.",
        f"{pct_no_date}% des lignes n'ont pas de date prévue — la Ponctualité n'est pas mesurable. L'OTIF en dépendant aussi, le Score retenu est le Taux de Service seul.",
        f"Lignes brutes export ERP : {quality['raw_rows']:,} · Lignes exploitables : {quality['clean_rows']:,} ({quality['usable_rate']}%).",
        "Un même fournisseur peut apparaître plusieurs fois dans ce fichier (une ligne par rayon livré) — le Taux de Service affiché est propre à chaque rayon, jamais mélangé.",
    ]
    for line in dq_rows:
        n = ws1.max_row + 1
        ws1.append([line])
        ws1.merge_cells(start_row=n, start_column=1, end_row=n, end_column=2)
        ws1.cell(n, 1).font = PA_DQ_FONT
        ws1.cell(n, 1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws1.row_dimensions[n].height = 26

    _pa_auto_width(ws1, [55, 20])
    ws1.freeze_panes = "A5"

    # ── Onglet 2 : Suivi Fournisseurs ───────────────────────────────────
    ws2 = wb.create_sheet("Suivi Fournisseurs")
    ws2.sheet_properties.tabColor = CFAO_ORANGE
    headers2 = [
        "Code Fournisseur", "Fournisseur", "Rayon", "Qté commandée", "Qté reçue", "Qté manquante",
        "Nb commandes", "Nb articles", "Nb sites", "Taux de Service", "Liste de surveillance",
        "Escalade suggérée", "Acheteur", "Cause racine", "Plan d'action",
        "Date de retour à la normale", "Statut",
    ]
    _pa_write_title(ws2, "SUIVI FOURNISSEURS — PLAN D'ACTION OTIF",
                     f"Généré le {today_str} · {n_crit} situations Niveau Critique ({n_crit_fou} fournisseurs distincts) · "
                     f"Trié par volume manquant décroissant · 1 ligne = 1 fournisseur sur 1 rayon", len(headers2))
    ws2.append([])
    _pa_write_header(ws2, 4, headers2)

    row2_start = 5
    for idx, row in crit_sup.iterrows():
        fou, rayon = row["Fou"], row["rayon_label"]
        watch_n = int(gold_counts.get((fou, rayon), 0)) + int(silver_counts.get((fou, rayon), 0))
        escalade = "Responsable Achat" if (watch_n > 0 and row["fill_rate"] < 20) else "Acheteur"
        n = ws2.max_row + 1
        ws2.append([
            int(fou) if pd.notna(fou) else "",
            row["supplier_name"],
            rayon,
            int(row["qte_cde"]),
            int(row["qte_rec"]),
            int(row["qty_missing"]),
            int(row["orders"]),
            int(row["articles"]),
            int(row["sites"]),
            round(row["fill_rate"], 1),
            watch_n,
            escalade,
            "", "", "", "", "",
        ])
        for col_i in range(1, len(headers2) + 1):
            cell = ws2.cell(n, col_i)
            cell.font = PA_BASE_FONT
            cell.border = PA_THIN
            cell.alignment = PA_LFT if col_i in (2, 15) else PA_CTR
        ws2.cell(n, 10).fill = _pa_score_fill(row["fill_rate"])
        ws2.cell(n, 10).number_format = '0.0"%"'
        if watch_n > 0:
            ws2.cell(n, 11).fill = PA_GOLD_FILL
        if escalade == "Responsable Achat":
            ws2.cell(n, 12).fill = PA_RED_FILL
        for col_i in range(13, 18):
            ws2.cell(n, col_i).fill = PA_INPUT_FILL
        ws2.cell(n, 16).number_format = PA_DATE_FMT

    row2_end = ws2.max_row

    cause_dv = DataValidation(type="list", formula1='"' + ",".join(CAUSES_RACINE) + '"', allow_blank=True)
    cause_dv.error = "Choisissez une cause dans la liste."
    cause_dv.errorTitle = "Cause non valide"
    ws2.add_data_validation(cause_dv)
    cause_dv.add(f"N{row2_start}:N{row2_end}")

    statut_dv = DataValidation(type="list", formula1='"' + ",".join(STATUTS_SUIVI) + '"', allow_blank=True)
    statut_dv.error = "Choisissez un statut dans la liste."
    statut_dv.errorTitle = "Statut non valide"
    ws2.add_data_validation(statut_dv)
    statut_dv.add(f"Q{row2_start}:Q{row2_end}")

    _pa_auto_width(ws2, [10, 26, 11, 10, 10, 12, 9, 9, 8, 11, 14, 16, 14, 22, 30, 18, 16])
    ws2.freeze_panes = "A5"
    ws2.auto_filter.ref = f"A4:{get_column_letter(len(headers2))}{row2_end}"

    # ── Onglet 3 : Taux de Service par Site ─────────────────────────────
    ws3 = wb.create_sheet("Taux de Service par Site")
    ws3.sheet_properties.tabColor = CFAO_ORANGE
    headers3 = ["Fournisseur", "Rayon"] + all_sites
    _pa_write_title(ws3, "TAUX DE SERVICE PAR SITE",
                     f"Généré le {today_str} · {n_crit} situations critiques × {len(all_sites)} magasins · "
                     f"« - » = aucune commande sur ce couple · En-tête coloré par format (navy = Hyper, bleu = Market, orange = Supeco)",
                     len(headers3))
    ws3.append([])
    _pa_write_header(ws3, 4, headers3)

    PA_FORMAT_HEADER_FILL = {
        0: (PatternFill("solid", fgColor=CFAO_NAVY), Font(bold=True, color="FFFFFF", size=10)),
        1: (PatternFill("solid", fgColor="2A78D6"), Font(bold=True, color="FFFFFF", size=10)),
        2: (PatternFill("solid", fgColor=CFAO_ORANGE), Font(bold=True, color=CFAO_NAVY, size=10)),
        3: (PatternFill("solid", fgColor="8E8E93"), Font(bold=True, color="FFFFFF", size=10)),
    }
    for j, site in enumerate(all_sites, start=3):
        fmt_idx = _pa_format_key(site)[0]
        fill, font = PA_FORMAT_HEADER_FILL[fmt_idx]
        ws3.cell(4, j).fill = fill
        ws3.cell(4, j).font = font

    for idx, row in crit_sup.iterrows():
        fou, rayon = row["Fou"], row["rayon_label"]
        n = ws3.max_row + 1
        line = [row["supplier_name"], rayon]
        for site in all_sites:
            v = pivot_map.get((fou, rayon, site))
            line.append(v if v is not None else "-")
        ws3.append(line)
        ws3.cell(n, 1).font = PA_BASE_FONT
        ws3.cell(n, 1).border = PA_THIN
        ws3.cell(n, 1).alignment = PA_LFT
        ws3.cell(n, 2).font = PA_BASE_FONT
        ws3.cell(n, 2).border = PA_THIN
        ws3.cell(n, 2).alignment = PA_CTR
        for j, site in enumerate(all_sites, start=3):
            cell = ws3.cell(n, j)
            cell.border = PA_THIN
            cell.alignment = PA_CTR
            v = pivot_map.get((fou, rayon, site))
            if v is None:
                cell.font = PA_NODATA_FONT
            else:
                cell.font = PA_BASE_FONT
                cell.number_format = "0.0"
                cell.fill = _pa_score_fill(v)

    _pa_auto_width(ws3, [26, 11] + [13] * len(all_sites))
    ws3.freeze_panes = "C5"
    ws3.auto_filter.ref = f"A4:{get_column_letter(len(headers3))}{ws3.max_row}"

    # ── Onglet 4 : Détail Commandes ─────────────────────────────────────
    ws4 = wb.create_sheet("Détail Commandes")
    ws4.sheet_properties.tabColor = CFAO_ORANGE
    headers4 = [
        "Fournisseur", "Code Fou.", "Rayon", "Magasin", "Code article", "Désignation",
        "N° Commande", "Date de commande", "Jours depuis commande",
        "Qté cde", "Qté reçue", "Qté manquante",
    ]
    _pa_write_title(ws4, "DÉTAIL COMMANDES — SITUATIONS CRITIQUES",
                     f"Généré le {today_str} · Toutes les lignes non-OTIF des {n_crit} situations Fournisseur × Rayon critiques "
                     f"({len(det)} lignes) · Groupées par situation (ordre de criticité), triées par volume manquant · "
                     f"Fond doré = article GOLD · Fond gris = article SILVER",
                     len(headers4))
    ws4.append([])
    _pa_write_header(ws4, 4, headers4)

    # Écriture allégée : pas de style cellule par cellule sur des dizaines de
    # milliers de lignes (le quadrillage Excel natif suffit) — seuls les
    # fonds GOLD/SILVER, porteurs de sens métier, sont appliqués.
    n_cols4 = len(headers4)
    for _, row in det.iterrows():
        date_cde = row.get("Date de commande", None)
        jours = (now.date() - date_cde.date()).days if pd.notna(date_cde) else ""
        date_txt = date_cde.strftime("%d/%m/%Y") if pd.notna(date_cde) else ""
        ws4.append([
            row.get("supplier_name", ""),
            int(row["Fou"]) if pd.notna(row["Fou"]) else "",
            row.get("rayon_label", ""),
            row.get("site_label", ""),
            row.get("Code", ""),
            row.get("article_label", ""),
            row.get("N° Cde", ""),
            date_txt,
            jours,
            int(row.get("qte_cde", 0)),
            int(row.get("qte_rec_retained", 0)),
            int(row.get("qty_missing", 0)),
        ])
        bucket = classe_bucket(row.get("watch_classe", ""))
        if bucket:
            fill = PA_GOLD_ROW if bucket == "GOLD" else PA_SILVER_ROW
            n = ws4.max_row
            for col_i in range(1, n_cols4 + 1):
                ws4.cell(n, col_i).fill = fill

    _pa_auto_width(ws4, [26, 10, 11, 20, 12, 30, 12, 14, 12, 10, 10, 12])
    ws4.freeze_panes = "A5"
    ws4.auto_filter.ref = f"A4:{get_column_letter(len(headers4))}{ws4.max_row}"

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf


