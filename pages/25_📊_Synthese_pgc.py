"""
SmartBuyer Hub — Module 25 · Synthèse PGC (version action)

Input  : one or more "Hier" exports + optional "Cette Semaine" export (Power BI),
         optional history CSV, optional recipients CSV.
Output : action-oriented synthesis (Actions / Analyse semaine / État du jour),
         standalone HTML export, updated history CSV (no formulas).

Python internals in English, user-facing labels in French.
"""

from __future__ import annotations

import html
import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

try:
    import plotly.graph_objects as go
except ImportError:  # pragma: no cover
    go = None


# ════════════════════════════════════════════════════════════════════════════
# CONSTANTS & SETTINGS
# ════════════════════════════════════════════════════════════════════════════

PGC_DEPT = "01 - PGC"
RAYON_LABELS = {
    "010 - BOISSON": "Boisson",
    "011 - DROGUERIE": "Droguerie",
    "012 - PARFUMERIE HYGIENE": "Parfumerie-Hygiène",
    "014 - EPICERIE": "Épicerie",
}
RAYON_ORDER = ["Boisson", "Épicerie", "Parfumerie-Hygiène", "Droguerie"]
SITE_ORDER = [
    "Hyper Marcory", "Hyper Palmeraie", "Hyper Yopougon",
    "Market 7 Décembre", "Market Riviera", "Market 2 Plateaux", "Market Kokoh Mall",
    "Market Aboboté", "Market Cité verte",
    "Supeco Niangon", "Supeco Terminus 47", "Supeco Toit rouge",
]
UNBUDGETED_FORMATS = {"Supeco"}
WEEKDAYS_FR = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]

NUM_COLS = {
    "CA": "ca", "CA N-1": "ca_n1", "Budget": "budget",
    "Marge": "marge", "Marge N-1": "marge_n1",
    "Débit": "debit", "Débit N-1": "debit_n1",
    "Volume": "volume", "Volume N-1": "volume_n1",
}
HISTORY_COLS = ["date", "level", "rayon", "site", "fmt", "ca", "ca_n1", "budget",
                "marge", "marge_n1", "debit", "debit_n1", "volume", "volume_n1",
                "effet_volume", "effet_taux", "export_ts"]

DEFAULT_SETTINGS = {
    "bulk_basket_mult": 2.0,     # basket N >= x * basket N-1
    "bulk_max_rate": 0.03,       # and margin rate < 3 %
    "materiality_k": 50.0,       # min stake (k FCFA) to create an action
    "max_week_actions": 6,       # cap of "Cette semaine" actions (others listed as secondary gaps)
    "today_k": 200.0,            # taux effect threshold for "Aujourd'hui"
    "persist_k": 10.0,           # daily margin loss (k) counted for persistence
    "traffic_drop": 0.25,        # tickets < -25 % => traffic action
    "ticket_stable": 0.10,       # |tickets| <= 10 % => "tickets stables"
    "bulk_floor_rate": 0.05,     # floor quoted in B2B message
    # article export
    "price_anomaly_rate": -0.50,  # margin rate below => price anomaly
    "article_min_loss_k": 10.0,   # min loss (k) for a price anomaly
    "rupture_min_ca_k": 20.0,     # usual daily CA (k) to watch a stock-out
    "rupture_other_sites": 3,     # still sold in at least N other stores
    "bulk_qty_mult": 3.0,         # article qty >= x * reference qty
    "bulk_min_qty": 20.0,
}

URG_TODAY, URG_WEEK, URG_FOLLOW = 0, 1, 2
URG_LABELS = {URG_TODAY: "À traiter aujourd'hui", URG_WEEK: "Cette semaine", URG_FOLLOW: "À suivre"}


# ════════════════════════════════════════════════════════════════════════════
# FORMATTING (French)
# ════════════════════════════════════════════════════════════════════════════

MINUS = "−"
NBSP = " "


def _num(v: float, dec: int = 0) -> str:
    s = f"{abs(v):,.{dec}f}".replace(",", NBSP).replace(".", ",")
    return s


def fmt_m(v: float, dec: int = 1) -> str:
    sign = MINUS if v < 0 else ""
    return f"{sign}{_num(v / 1e6, dec)} M"


def fmt_k(v_k: float, signed: bool = True) -> str:
    """v_k already in thousands."""
    if abs(v_k) >= 1000:
        s = f"{_num(v_k / 1000, 2)} M"
    else:
        s = f"{_num(v_k, 0)} k"
    if v_k < 0:
        return MINUS + s
    return ("+" + s) if signed else s


def fmt_pct(v: float, dec: int = 1, signed: bool = True) -> str:
    if v is None or pd.isna(v):
        return "n.d."
    s = _num(v * 100, dec) + " %"
    if v < 0:
        return MINUS + s
    return ("+" + s) if signed else s


def fmt_rate(t: float, dec: int = 1) -> str:
    if t is None or pd.isna(t):
        return "n.d."
    s = _num(t * 100, dec) + " %"
    return (MINUS + s) if t < 0 else s


def fmt_bps(v: float) -> str:
    b = round(v * 10000)
    return (MINUS if b < 0 else "+") + f"{abs(b)} bps"


def short_site(site: str) -> str:
    return re.sub(r"^\d+\s*-\s*", "", str(site)).strip()


def site_fmt(site: str) -> str:
    s = short_site(site)
    return s.split(" ")[0] if s else ""


def site_name_only(site: str) -> str:
    s = short_site(site)
    parts = s.split(" ", 1)
    return parts[1] if len(parts) > 1 else s


def esc(s) -> str:
    return html.escape(str(s))


def safe_div(a, b):
    try:
        return a / b if b else float("nan")
    except ZeroDivisionError:
        return float("nan")


def weekday_fr(d: date) -> str:
    return WEEKDAYS_FR[d.weekday()]


def dfr(d: date) -> str:
    return d.strftime("%d/%m")


# ════════════════════════════════════════════════════════════════════════════
# PARSING
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class ExportFile:
    name: str
    kind: str                     # "day" | "week" | "unknown"
    export_ts: datetime | None
    display_type: str
    detail: pd.DataFrame          # PGC lines rayon × site (normalized)
    hors_pgc: dict                # {"ca", "ca_n1", "marge", "marge_n1"} network, non-PGC depts
    store_total: dict             # {"debit", "debit_n1", "ca", "ca_n1"}
    sales_date: date | None = None
    week_start: date | None = None
    week_end: date | None = None
    monday: bool = False
    error: str | None = None

    @property
    def empty(self) -> bool:
        return self.detail.empty or float(self.detail["ca"].sum()) == 0.0


FILENAME_TS = re.compile(r"(\d{4}-\d{2}-\d{2})T(\d{2})(\d{2})(\d{2})")


def parse_export_ts(filename: str) -> datetime | None:
    m = FILENAME_TS.search(filename)
    if not m:
        return None
    return datetime.strptime(f"{m.group(1)} {m.group(2)}{m.group(3)}{m.group(4)}", "%Y-%m-%d %H%M%S")


def read_export(file_bytes: bytes, filename: str, raw: pd.DataFrame | None = None) -> ExportFile:
    try:
        raw = raw if raw is not None else pd.read_excel(io.BytesIO(file_bytes))
    except Exception as exc:  # noqa: BLE001
        return ExportFile(filename, "unknown", None, "", pd.DataFrame(), {}, {}, error=f"Lecture impossible : {exc}")

    first_col = raw.columns[0]
    # Filter footer ("Filtres appliqués : ...")
    footer = ""
    for v in raw[first_col].astype(str).tolist()[::-1]:
        if v.startswith("Filtres appliqués"):
            footer = v
            break
    m = re.search(r"Type d'affichage est ([^\n]+)", footer)
    display_type = m.group(1).strip() if m else ""
    low = display_type.lower()
    kind = "day" if low == "hier" else "week" if "semaine" in low else "unknown"

    missing = [c for c in ["Département", "Rayon", "Site", "CA", "CA N-1", "Marge", "Marge N-1"] if c not in raw.columns]
    if missing:
        return ExportFile(filename, kind, parse_export_ts(filename), display_type, pd.DataFrame(), {}, {},
                          error=f"Colonnes manquantes : {', '.join(missing)}")

    df = raw.copy()
    for src in NUM_COLS:
        if src not in df.columns:
            df[src] = 0.0
        df[src] = pd.to_numeric(df[src], errors="coerce")

    is_pgc = df["Département"].astype(str).str.strip() == PGC_DEPT
    is_line = df["Site"].notna() & (df["Site"].astype(str).str.strip() != "Total")
    det = df[is_pgc & is_line].copy()
    det = det[det["Rayon"].isin(RAYON_LABELS)]
    out = pd.DataFrame({
        "rayon": det["Rayon"].map(RAYON_LABELS),
        "site": det["Site"].map(short_site),
    })
    for src, dst in NUM_COLS.items():
        out[dst] = det[src].fillna(0.0).astype(float).values
    out["fmt"] = out["site"].map(lambda s: s.split(" ")[0])
    out = out.reset_index(drop=True)

    # Network totals (grand total row = Département "Total")
    grand = df[df["Département"].astype(str).str.strip() == "Total"]
    pgc_tot = df[is_pgc & (df["Rayon"].astype(str).str.strip() == "Total") & df["Site"].isna()]
    hors_pgc, store_total = {}, {}
    if len(grand) and len(pgc_tot):
        g, p = grand.iloc[0], pgc_tot.iloc[0]
        hors_pgc = {
            "ca": float(g["CA"] or 0) - float(p["CA"] or 0),
            "ca_n1": float(g["CA N-1"] or 0) - float(p["CA N-1"] or 0),
            "marge": float(g["Marge"] or 0) - float(p["Marge"] or 0),
            "marge_n1": float(g["Marge N-1"] or 0) - float(p["Marge N-1"] or 0),
        }
        store_total = {
            "ca": float(g["CA"] or 0), "ca_n1": float(g["CA N-1"] or 0),
            "debit": float(0 if pd.isna(g["Débit"]) else g["Débit"]),
            "debit_n1": float(0 if pd.isna(g["Débit N-1"]) else g["Débit N-1"]),
        }

    ef = ExportFile(filename, kind, parse_export_ts(filename), display_type, out, hors_pgc, store_total)
    if kind == "unknown":
        ef.error = f"Type d'export non reconnu (« {display_type or 'absent'} »). Attendu : « Hier » ou « Cette Semaine »."
    return ef


def assign_dates(ef: ExportFile, manual_export_date: date | None = None) -> None:
    """Sales date (day export) or week period (week export) from the export date."""
    exp_d = ef.export_ts.date() if ef.export_ts else manual_export_date
    if exp_d is None:
        return
    if ef.kind == "day":
        ef.sales_date = exp_d - timedelta(days=1)
    elif ef.kind == "week":
        if exp_d.weekday() == 0:
            # Monday: current week has no past day => content is empty or previous full week
            ef.monday = True
            ef.week_start = exp_d - timedelta(days=7)
            ef.week_end = exp_d - timedelta(days=1)
        else:
            ef.week_start = exp_d - timedelta(days=exp_d.weekday())   # Monday of export week
            ef.week_end = exp_d - timedelta(days=1)                   # "today" excluded


# ─── Article export (article × site, one day, N-1 NOT iso-day) ────────────────

ART_RAYON_KEYS = [("BOISSON", "Boisson"), ("EPICERIE", "Épicerie"), ("ÉPICERIE", "Épicerie"),
                  ("PARFUMERIE", "Parfumerie-Hygiène"), ("DROGUERIE", "Droguerie")]
ART_COLS = {
    "CA": "ca", "CA N-1": "ca_n1", "Marge": "marge",
    "CA Promo": "ca_promo", "CA Promo N-1": "ca_promo_n1", "Marge Promo": "marge_promo",
    "Marge Promo N-1": "marge_promo_n1", "CA Hors Promo": "ca_hp", "CA Hors Promo N-1": "ca_hp_n1",
    "Marge Hors Promo": "marge_hp", "Marge Hors Promo N-1": "marge_hp_n1",
    "Qté Vente": "qte", "Qté Vente N-1": "qte_n1", "Casse (Valeur)": "casse", "Casse (Qté)": "casse_qte",
}


@dataclass
class ArticleExport:
    name: str
    day: date | None
    lines: pd.DataFrame
    excluded_sites: list = field(default_factory=list)
    error: str | None = None


def _map_rayon(r) -> str | None:
    up = str(r).upper()
    for k, v in ART_RAYON_KEYS:
        if k in up:
            return v
    return None


def _footer(raw: pd.DataFrame) -> str:
    first_col = raw.columns[0]
    for v in raw[first_col].astype(str).tolist()[::-1]:
        if v.startswith("Filtres appliqués"):
            return v
    return ""


def read_article_export(raw: pd.DataFrame, filename: str) -> ArticleExport:
    footer = _footer(raw)
    m = re.search(r"Date est le ou après le (\d{2}/\d{2}/\d{4})", footer) or re.search(r"(\d{2}/\d{2}/\d{4})", footer)
    day = datetime.strptime(m.group(1), "%d/%m/%Y").date() if m else None
    if day is None:
        ts = parse_export_ts(filename)
        day = (ts.date() - timedelta(days=1)) if ts else None
    site_col = "Site nom long" if "Site nom long" in raw.columns else "Site"
    missing = [c for c in ["Rayon", "Article", site_col, "CA", "Marge"] if c not in raw.columns]
    if missing:
        return ArticleExport(filename, day, pd.DataFrame(), error=f"Colonnes manquantes : {', '.join(missing)}")
    L = raw[raw["Article"].notna() & raw[site_col].notna()
            & (raw[site_col].astype(str).str.strip() != "Total")].copy()
    out = pd.DataFrame({
        "rayon": L["Rayon"].map(_map_rayon).values,
        "site": L[site_col].map(short_site).values,
        "article": L["Article"].astype(str).str.strip().values,
    })
    parts = out["article"].str.split(" - ", n=1, expand=True)
    out["code"] = parts[0].str.strip()
    out["lib"] = (parts[1] if parts.shape[1] > 1 else parts[0]).fillna(out["article"]).str.strip()
    for src, dst in ART_COLS.items():
        out[dst] = pd.to_numeric(L[src], errors="coerce").fillna(0.0).values if src in L.columns else 0.0
    # Marge N-1 is not exported as such: rebuilt from promo + non-promo N-1 margins
    out["marge_n1"] = out["marge_hp_n1"] + out["marge_promo_n1"]
    out = out[out["rayon"].notna()]
    excluded = sorted(set(out["site"]) - set(SITE_ORDER))
    out = out[out["site"].isin(SITE_ORDER)].reset_index(drop=True)
    out["fmt"] = out["site"].map(lambda x: x.split(" ")[0])
    return ArticleExport(filename, day, out, excluded)


@st.cache_data(show_spinner=False)
def classify_file(file_bytes: bytes, filename: str):
    """Detect the file type from its content. Returns (kind, payload)."""
    low = filename.lower()
    if low.endswith(".csv"):
        head = None
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                head = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine="python", encoding=enc, nrows=5)
                break
            except Exception:  # noqa: BLE001
                continue
        if head is None:
            return "unknown", "CSV illisible."
        cols = {c.strip().lower() for c in head.columns}
        if {"type", "cle", "nom"} <= cols:
            return "recipients", pd.read_csv(io.BytesIO(file_bytes), sep=None, engine="python", encoding="utf-8-sig")
        if {"date", "level"} <= cols:
            try:
                return "history", load_history(file_bytes)
            except ValueError as exc:
                return "unknown", str(exc)
        return "unknown", "CSV non reconnu (ni historique, ni destinataires)."
    try:
        raw = pd.read_excel(io.BytesIO(file_bytes))
    except Exception as exc:  # noqa: BLE001
        return "unknown", f"Lecture impossible : {exc}"
    if "Article" in raw.columns:
        ae = read_article_export(raw, filename)
        return ("unknown", ae.error) if ae.error else ("article", ae)
    ef = read_export(file_bytes, filename, raw=raw)
    if ef.error:
        return "unknown", ef.error
    return ef.kind, ef


# ════════════════════════════════════════════════════════════════════════════
# HISTORY
# ════════════════════════════════════════════════════════════════════════════

def day_rows(ef: ExportFile) -> pd.DataFrame:
    d = ef.detail.copy()
    d["level"] = "detail"
    rows = [d]
    if ef.hors_pgc:
        rows.append(pd.DataFrame([{"level": "hors_pgc", "rayon": "Hors PGC", "site": "Réseau", "fmt": "",
                                   **ef.hors_pgc}]))
    if ef.store_total:
        rows.append(pd.DataFrame([{"level": "magasin", "rayon": "Total", "site": "Réseau", "fmt": "",
                                   **ef.store_total}]))
    out = pd.concat(rows, ignore_index=True)
    out["date"] = ef.sales_date.isoformat()
    out["export_ts"] = ef.export_ts.isoformat() if ef.export_ts else ""
    return out


def load_history(file_bytes: bytes) -> pd.DataFrame:
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            h = pd.read_csv(io.BytesIO(file_bytes), sep=";", encoding=enc, decimal=",")
            break
        except Exception:  # noqa: BLE001
            h = None
    if h is None or "date" not in h.columns:
        raise ValueError("Historique illisible : colonne « date » absente.")
    text_cols = ("date", "level", "rayon", "site", "fmt", "export_ts")
    for c in HISTORY_COLS:
        if c not in h.columns:
            h[c] = "" if c in text_cols else 0.0
    for c in text_cols:
        h[c] = h[c].fillna("").astype(str)
    for c in set(HISTORY_COLS) - set(text_cols):
        h[c] = pd.to_numeric(h[c], errors="coerce").fillna(0.0)
    h["date"] = pd.to_datetime(h["date"], dayfirst=False, errors="coerce").dt.date.astype(str)
    return h[HISTORY_COLS].drop(columns=["effet_volume", "effet_taux"])


@dataclass
class Issue:
    level: str        # "r" | "o" | "g" | "i"
    tag: str
    title: str
    detail: str = ""
    action: bool = False
    site: str | None = None
    day: str | None = None


def merge_history(history: pd.DataFrame | None, days: list[ExportFile]) -> tuple[pd.DataFrame, list[Issue]]:
    issues: list[Issue] = []
    frames = []
    if history is not None and len(history):
        h = history.copy()
        h["_src"] = "history"
        frames.append(h)
    for ef in days:
        r = day_rows(ef)
        r["_src"] = ef.name
        frames.append(r)
    if not frames:
        return pd.DataFrame(columns=HISTORY_COLS), issues
    allr = pd.concat(frames, ignore_index=True)
    allr["export_ts"] = allr["export_ts"].fillna("").astype(str)

    # One version per date: latest export wins
    keep = []
    for d, g in allr.groupby("date"):
        srcs = g.groupby("_src")["export_ts"].max().sort_values()
        if len(srcs) > 1:
            det = g[g["level"] == "detail"]
            n1 = det.groupby("_src")["ca_n1"].sum()
            if n1.max() > 0 and (n1.max() - n1.min()) / n1.max() > 0.005:
                issues.append(Issue("o", "Même jour, N-1 différent",
                                    f"{dfr(date.fromisoformat(d))} : chargé {len(srcs)} fois avec des N-1 différents",
                                    f"Écart de N-1 : {fmt_m(n1.max() - n1.min())}. La version la plus récente est gardée.",
                                    action=True, day=d))
            else:
                issues.append(Issue("i", "Doublon", f"{dfr(date.fromisoformat(d))} chargé {len(srcs)} fois",
                                    "La version la plus récente est gardée."))
        keep.append(g[g["_src"] == srcs.index[-1]])
    merged = pd.concat(keep, ignore_index=True).drop(columns=["_src"])
    merged = merged.sort_values(["date", "level", "rayon", "site"]).reset_index(drop=True)
    return merged, issues


# ════════════════════════════════════════════════════════════════════════════
# ENGINE — lines, Bennet, bulk sales
# ════════════════════════════════════════════════════════════════════════════

def analyze_lines(df: pd.DataFrame, s: dict) -> pd.DataFrame:
    """Line metrics + symmetric Bennet decomposition (zero residual)."""
    d = df.copy()
    d["tm"] = [safe_div(m, c) if c > 0 else 0.0 for m, c in zip(d.marge, d.ca)]
    d["tm_n1"] = [safe_div(m, c) if c > 0 else 0.0 for m, c in zip(d.marge_n1, d.ca_n1)]
    d["dM"] = d.marge - d.marge_n1
    d["effet_volume"] = (d.ca - d.ca_n1) * (d.tm + d.tm_n1) / 2
    d["effet_taux"] = (d.tm - d.tm_n1) * (d.ca + d.ca_n1) / 2
    d["panier"] = [safe_div(c, n) for c, n in zip(d.ca, d.debit)]
    d["panier_n1"] = [safe_div(c, n) for c, n in zip(d.ca_n1, d.debit_n1)]
    d["ca_var"] = [safe_div(a, b) - 1 if b else float("nan") for a, b in zip(d.ca, d.ca_n1)]
    d["deb_var"] = [safe_div(a, b) - 1 if b else float("nan") for a, b in zip(d.debit, d.debit_n1)]
    d["pan_var"] = [safe_div(a, b) - 1 if b and b == b else float("nan") for a, b in zip(d.panier, d.panier_n1)]
    k, r = s["bulk_basket_mult"], s["bulk_max_rate"]
    d["bulk"] = (d.panier >= k * d.panier_n1) & (d.tm < r) & (d.ca > 0)
    d["base_bulk"] = (d.panier_n1 >= k * d.panier) & (d.tm_n1 < 0.08) & (d.ca_n1 > 0)
    d["zero"] = (d.ca == 0) & (d.debit == 0)
    return d


def totals(d: pd.DataFrame) -> dict:
    budgeted = d[~d.fmt.isin(UNBUDGETED_FORMATS)]
    t = {
        "ca": d.ca.sum(), "ca_n1": d.ca_n1.sum(), "marge": d.marge.sum(), "marge_n1": d.marge_n1.sum(),
        "budget": d.budget.sum(), "ca_budgeted": budgeted.ca.sum(), "budget_real": budgeted.budget.sum(),
        "debit": d.debit.sum(), "debit_n1": d.debit_n1.sum(),
    }
    t["ca_var"] = safe_div(t["ca"], t["ca_n1"]) - 1
    t["bgt_shown"] = safe_div(t["ca"], t["budget"]) - 1 if t["budget"] else float("nan")
    t["bgt_real"] = safe_div(t["ca_budgeted"], t["budget_real"]) - 1 if t["budget_real"] else float("nan")
    t["dM"] = t["marge"] - t["marge_n1"]
    t["tm"] = safe_div(t["marge"], t["ca"])
    t["tm_n1"] = safe_div(t["marge_n1"], t["ca_n1"])
    return t


def bennet_group(d: pd.DataFrame, by: str) -> pd.DataFrame:
    """Aggregate Bennet: volume (aggregate), pure rate (Σ line rates), mix (residual of aggregation)."""
    rows = []
    for key, g in d.groupby(by):
        t = totals(g)
        vol_agg = (t["ca"] - t["ca_n1"]) * ((t["tm"] or 0) + (t["tm_n1"] or 0)) / 2
        taux_agg = ((t["tm"] or 0) - (t["tm_n1"] or 0)) * (t["ca"] + t["ca_n1"]) / 2
        taux_pure = g.effet_taux.sum()
        rows.append({by: key, **t, "effet_volume": vol_agg, "effet_taux": taux_pure,
                     "effet_mix": taux_agg - taux_pure})
    return pd.DataFrame(rows)


# ════════════════════════════════════════════════════════════════════════════
# DATA QUALITY
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class WeekRef:
    label: str
    source: str                   # "export" | "history" | "closure"
    start: date
    end: date
    lines: pd.DataFrame           # raw lines (not analyzed)
    hors_pgc: dict
    days_in_history: list = field(default_factory=list)


def quality_checks(hist: pd.DataFrame, week: ExportFile | None, s: dict) -> tuple[list[Issue], set]:
    issues: list[Issue] = []
    incomplete: set = set()   # (date, site)
    det = hist[hist.level == "detail"] if len(hist) else hist
    if not len(det):
        return issues, incomplete
    all_sites = sorted(set(det.site), key=lambda x: SITE_ORDER.index(x) if x in SITE_ORDER else 99)
    dates = sorted(set(det.date))

    for d in dates:
        g = det[det.date == d]
        dd = date.fromisoformat(d)
        for site in all_sites:
            sg = g[g.site == site]
            if sg.empty:
                issues.append(Issue("r", "Magasin absent", f"{site} · {dfr(dd)}",
                                    "Aucune ligne dans l'export « Hier ».", action=True, site=site, day=d))
                incomplete.add((d, site))
                continue
            if sg.ca.sum() == 0 and sg.debit.sum() == 0:
                bud = sg.budget.sum()
                issues.append(Issue("r", "Magasin à zéro", f"{site} · {dfr(dd)}",
                                    f"0 CA, 0 ticket{'' if bud else ' et pas de budget'} dans l'export « Hier ». "
                                    f"Fermeture ou défaut de remontée caisse ? "
                                    f"Pèse {fmt_m(-sg.ca_n1.sum())} de CA vs N-1.",
                                    action=True, site=site, day=d))
                incomplete.add((d, site))
                continue
            missing_r = set(RAYON_ORDER) - set(sg.rayon)
            if missing_r:
                issues.append(Issue("o", "Rayon manquant", f"{site} · {dfr(dd)}",
                                    f"Rayon(s) absent(s) : {', '.join(sorted(missing_r))}.", site=site, day=d))
            if site_fmt(site) not in UNBUDGETED_FORMATS and sg.budget.sum() == 0:
                issues.append(Issue("o", "Budget à zéro", f"{site} · {dfr(dd)}",
                                    "Magasin normalement budgété, budget nul ce jour.", site=site, day=d))

    if week is not None and week.week_start and week.week_end and not week.empty:
        span = [week.week_start + timedelta(days=i) for i in range((week.week_end - week.week_start).days + 1)]
        missing_days = [x for x in span if x.isoformat() not in dates]
        if missing_days:
            issues.append(Issue("o", "Jours absents",
                                ", ".join(f"{weekday_fr(x).capitalize()} {x.day}" for x in missing_days)
                                + " : exports « Hier » non chargés",
                                "La courbe jour par jour les regroupe en une valeur déduite de l'export semaine."))
        # Week vs loaded days, per site
        cols = ["ca", "ca_n1", "budget", "debit", "debit_n1"]
        wk = week.detail.groupby("site")[cols].sum()
        in_week = det[det.date.isin([x.isoformat() for x in span])]
        ld = in_week.groupby("site")[cols].sum().reindex(wk.index).fillna(0)
        implied = wk - ld
        for site in wk.index:
            if (implied.loc[site, ["ca", "ca_n1"]] < -1000).any():
                issues.append(Issue("r", "Semaine ≠ jours", site,
                                    "Les jours chargés dépassent l'export semaine : base N-1 ou périmètre différent.",
                                    action=True, site=site))
                continue
            if missing_days:
                # tickets ratio: robust to bulk sales (which inflate CA, not tickets)
                r_imp = safe_div(implied.loc[site, "debit"], implied.loc[site, "debit_n1"])
                r_ld = safe_div(ld.loc[site, "debit"], ld.loc[site, "debit_n1"])
                if r_imp == r_imp and r_ld == r_ld and abs(r_imp - r_ld) > 0.5:
                    n = len(missing_days)
                    issues.append(Issue(
                        "o", "Semaine ≠ jours", site,
                        f"La semaine moins les jours chargés laisse {fmt_m(implied.loc[site, 'ca'])} pour "
                        f"{n} jour{'s' if n > 1 else ''}, soit {fmt_m(implied.loc[site, 'ca'] / n)} par jour "
                        f"(N-1 : {fmt_m(implied.loc[site, 'ca_n1'] / n)}), avec des tickets "
                        f"{fmt_pct(r_imp - 1, 0)} vs N-1 contre {fmt_pct(r_ld - 1, 0)} sur les jours chargés. "
                        "Relancer les exports « Hier » concernés.",
                        action=True, site=site))
            else:
                diff = safe_div(abs(wk.loc[site, "ca"] - ld.loc[site, "ca"]), wk.loc[site, "ca"])
                if diff == diff and diff > 0.01:
                    issues.append(Issue("o", "Semaine ≠ jours", site,
                                        f"Écart de {fmt_pct(diff, signed=False)} entre l'export semaine et la somme des jours.",
                                        action=True, site=site))
    return issues, incomplete


def build_week_ref(hist: pd.DataFrame, week: ExportFile | None, latest: date | None,
                   issues: list[Issue]) -> WeekRef | None:
    det = hist[hist.level == "detail"] if len(hist) else hist
    hp = hist[hist.level == "hors_pgc"] if len(hist) else hist

    def from_history(start: date, end: date, label: str, source: str) -> WeekRef | None:
        span = [(start + timedelta(days=i)).isoformat() for i in range((end - start).days + 1)]
        g = det[det.date.isin(span)]
        if g.empty:
            return None
        lines = g.groupby(["rayon", "site", "fmt"], as_index=False)[list(NUM_COLS.values())].sum()
        hpg = hp[hp.date.isin(span)]
        hors = {k: hpg[k].sum() for k in ("ca", "ca_n1", "marge", "marge_n1")} if len(hpg) else {}
        return WeekRef(label, source, start, end, lines, hors, sorted(set(g.date)))

    if week is not None and week.week_start:
        if week.empty:
            # Empty export (Monday morning) => close previous week from history
            if week.monday:
                prev_start, prev_end = week.week_start, week.week_end
            else:
                prev_end = week.week_start - timedelta(days=1)
                prev_start = prev_end - timedelta(days=6)
            issues.append(Issue("g", "Nouvelle semaine", "Export semaine vide : début de semaine",
                                "Le cumul affiché est la semaine précédente clôturée, reconstituée à partir des exports « Hier »."))
            wr = from_history(prev_start, prev_end, f"Semaine clôturée · {dfr(prev_start)} → {dfr(prev_end)}", "closure")
            if wr is not None:
                return wr
        elif week.monday:
            # Monday export with content => previous full week (checked against history in quality_checks)
            issues.append(Issue("i", "Export du lundi", "L'export semaine du lundi contient la semaine précédente",
                                "Il est utilisé comme clôture officielle de la semaine précédente."))
            wr = WeekRef(f"Semaine clôturée · {dfr(week.week_start)} → {dfr(week.week_end)}", "export",
                         week.week_start, week.week_end, week.detail.copy(), week.hors_pgc)
            return wr
        else:
            wr = WeekRef(f"Semaine à date · {weekday_fr(week.week_start)[:3]} {week.week_start.day} → "
                         f"{weekday_fr(week.week_end)[:3]} {week.week_end.day}/{week.week_end.month:02d}",
                         "export", week.week_start, week.week_end, week.detail.copy(), week.hors_pgc)
            span = [(week.week_start + timedelta(days=i)).isoformat()
                    for i in range((week.week_end - week.week_start).days + 1)]
            wr.days_in_history = sorted(set(det[det.date.isin(span)].date)) if len(det) else []
            return wr
    if latest is None:
        return None
    start = latest - timedelta(days=latest.weekday())
    return from_history(start, latest, f"Semaine reconstituée · {dfr(start)} → {dfr(latest)}", "history")


# ════════════════════════════════════════════════════════════════════════════
# ENGINE — persistence & actions
# ════════════════════════════════════════════════════════════════════════════

def streaks(hist: pd.DataFrame, incomplete: set, s: dict) -> dict:
    """(rayon, site) -> (consecutive loss days up to latest, last-day loss k, previous-day loss k)."""
    det = hist[hist.level == "detail"]
    if det.empty:
        return {}
    dates = sorted(set(det.date))
    piv = det.assign(dM=(det.marge - det.marge_n1) / 1000).pivot_table(
        index=["rayon", "site"], columns="date", values="dM", aggfunc="sum")
    out = {}
    for idx, row in piv.iterrows():
        n = 0
        for d in reversed(dates):
            if (d, idx[1]) in incomplete or pd.isna(row.get(d)):
                break
            if row[d] < -s["persist_k"]:
                n += 1
            else:
                break
        last = row.get(dates[-1], float("nan"))
        prev = row.get(dates[-2], float("nan")) if len(dates) > 1 else float("nan")
        out[idx] = (n, last, prev)
    return out


def traffic_streaks(hist: pd.DataFrame, incomplete: set, s: dict) -> dict:
    """site -> consecutive days (up to latest) with tickets below -traffic_drop vs N-1."""
    det = hist[hist.level == "detail"]
    if det.empty:
        return {}
    dates = sorted(set(det.date))
    g = det.groupby(["site", "date"])[["debit", "debit_n1"]].sum()
    out = {}
    for site in set(det.site):
        n = 0
        for d in reversed(dates):
            if (d, site) in incomplete or (site, d) not in g.index:
                break
            x = g.loc[(site, d)]
            if x.debit_n1 and x.debit / x.debit_n1 - 1 < -s["traffic_drop"]:
                n += 1
            else:
                break
        out[site] = n
    return out


def secondary_html(sec: list) -> str:
    if not sec:
        return ""
    items = "".join(f"<li><b>{esc(a.title)}</b> · {fmt_k(a.stake_k) if a.stake_k is not None else '—'} "
                    f"<span style='color:#6B7280'>({esc(a.stake_label)})</span></li>" for a in sec)
    return (f'<div class="mini"><h3>Autres écarts, sous le seuil de priorité ({len(sec)})</h3>'
            f"<ul>{items}</ul></div>")


@dataclass
class Action:
    kind: str
    urgency: int
    title: str
    stake_k: float | None
    stake_label: str
    trail_label: str
    trail: list
    fact: str
    cause: str
    todo: str
    owners: list                  # [(role, key)]
    due: str
    message: str = ""
    tags: list = field(default_factory=list)   # [(cls, label)]
    lines: list = field(default_factory=list)  # [(rayon, site)]
    articles: list = field(default_factory=list)
    articles_label: str = ""
    scope: str = "week"                        # "week" (rayon export) | "day" (article export)

    @property
    def filters(self) -> set:
        return {r for r, _ in self.owners}


ROLE_LABEL = {"achat": "Achat", "supply": "Supply", "magasin": "Magasin", "cdg": "Contrôle de gestion"}


class Recipients:
    def __init__(self, df: pd.DataFrame | None):
        self.map = {}
        if df is not None and len(df):
            for _, r in df.iterrows():
                self.map[(str(r["type"]).strip().lower(), str(r["cle"]).strip())] = str(r["nom"]).strip()

    def name(self, role: str, key: str) -> str:
        if role == "achat":
            return self.map.get(("rayon", key), f"Acheteur {key}")
        if role == "magasin":
            return self.map.get(("site", key), f"Directeur {key}")
        if role == "supply":
            return self.map.get(("fonction", "Supply"), "Supply")
        return self.map.get(("fonction", "Contrôle de gestion"), "Contrôle de gestion")

    def owner_text(self, owners: list) -> str:
        roles = []
        for role, key in owners:
            lab = f"{ROLE_LABEL[role]} {key}" if role == "achat" else ROLE_LABEL[role]
            if role == "magasin":
                lab = f"Magasin {site_name_only(key)}" if key else "Magasin"
            if lab not in roles:
                roles.append(lab)
        return " + ".join(roles)

    def mentions(self, owners: list) -> str:
        names = []
        for role, key in owners:
            n = self.name(role, key)
            if n not in names:
                names.append(n)
        return " ".join("@" + n for n in names)


def _day_line(day_lines: pd.DataFrame | None, rayon: str, site: str):
    if day_lines is None or day_lines.empty:
        return None
    g = day_lines[(day_lines.rayon == rayon) & (day_lines.site == site)]
    return g.iloc[0] if len(g) else None


def build_actions(ref: pd.DataFrame, day: pd.DataFrame | None, day_date: date | None,
                  stk: dict, issues: list[Issue], s: dict, rec: Recipients,
                  traffic_streak: dict | None = None) -> list[Action]:
    """ref = analyzed week lines; day = analyzed latest-day lines."""
    acts: list[Action] = []
    used: set = set()
    traffic_streak = traffic_streak or {}
    mat = s["materiality_k"] * 1000
    today = s["today_k"] * 1000
    hier = f"{weekday_fr(day_date).capitalize()}" if day_date else "Hier"

    def tag_for(keys: list) -> list:
        best = max((stk.get(k, (0, 0, 0)) for k in keys), key=lambda x: x[0], default=(0, 0, 0))
        n, last, prev = best
        if n >= 2:
            if prev == prev and last < prev * 1.5 and last < -s["persist_k"] * 5:
                return [("agg", f"Aggravé · {n} jours")]
            return [("pers", f"Persistant · {n} jours")]
        if n == 1:
            return [("new", f"Nouveau · {hier.lower()}")]
        return []

    # ── A. Sold at a loss (week or latest day)
    loss_keys = set(tuple(x) for x in ref[(ref.marge < 0) & (ref.ca > 0)][["rayon", "site"]].values)
    if day is not None:
        loss_keys |= set(tuple(x) for x in day[(day.marge < 0) & (day.ca > 0)][["rayon", "site"]].values)
    by_site: dict = {}
    for r, st_ in loss_keys:
        by_site.setdefault(st_, []).append(r)
    for site, rayons in sorted(by_site.items()):
        rayons = sorted(rayons, key=lambda r: RAYON_ORDER.index(r) if r in RAYON_ORDER else 9)
        g = ref[(ref.site == site) & (ref.rayon.isin(rayons))]
        stake = g.dM.sum()
        worst = g.sort_values("tm").iloc[0]
        dl = _day_line(day, worst.rayon, site)
        trail = [f"N-1 {fmt_rate(worst.tm_n1)}", f"Semaine {fmt_rate(worst.tm)}"]
        if dl is not None:
            trail.append(f"{hier} {fmt_rate(dl.tm)}")
        ray_txt = " et ".join(rayons)
        owners = [("achat", r) for r in rayons] + [("magasin", site)]
        a = Action("perte", URG_TODAY, f"{site} : {ray_txt} vendue{'s' if len(rayons) > 1 else ''} à perte",
                   stake / 1000, "marge semaine", f"Taux {worst.rayon} {site_name_only(site)}", trail,
                   f"{fmt_m(g.ca.sum())} de CA pour {fmt_k(g.marge.sum() / 1000)} de marge sur la semaine.",
                   "Prix de vente sous le coût d'achat, souvent sur des ventes en gros.",
                   "Identifier les références vendues sous le PMP et corriger les prix dès aujourd'hui.",
                   owners, "Aujourd'hui", lines=[(r, site) for r in rayons])
        a.message = (f"{rec.mentions(owners)}, {site} : {ray_txt} vendue{'s' if len(rayons) > 1 else ''} à perte "
                     f"({fmt_rate(worst.tm)} de marge sur la semaine"
                     + (f", {fmt_rate(dl.tm)} {hier.lower()}" if dl is not None else "") +
                     "). Sors aujourd'hui la liste des références vendues sous le PMP et corrige les prix. "
                     f"Enjeu : {fmt_k(stake / 1000)} de marge.")
        a.tags = tag_for(a.lines)
        acts.append(a)
        used |= set(a.lines)

    # ── B. Bulk sales under the floor
    bulk_keys = set(tuple(x) for x in ref[ref.bulk][["rayon", "site"]].values)
    if day is not None:
        bulk_keys |= set(tuple(x) for x in day[day.bulk][["rayon", "site"]].values)
    bulk_keys -= used
    if bulk_keys:
        g = ref[[(r, st_) in bulk_keys for r, st_ in zip(ref.rayon, ref.site)]]
        stake = g.dM.sum()
        if True:  # bulk sales under the floor are always actioned (policy breach)
            parts = []
            for _, x in g.sort_values("dM").iterrows():
                mult = safe_div(x.panier, x.panier_n1)
                parts.append(f"{x.rayon} {x.site} : {fmt_m(x.ca)} à {fmt_rate(x.tm)} de marge"
                             + (f", panier ×{_num(mult, 1)}" if mult == mult and mult > 1.2 else ""))
            w = g.sort_values("dM").iloc[0]
            dl = _day_line(day, w.rayon, w.site)
            trail = [f"N-1 {fmt_rate(w.tm_n1)}", f"Semaine {fmt_rate(w.tm)}"]
            if dl is not None:
                trail.append(f"{hier} {fmt_rate(dl.tm)}")
            sites = sorted(set(g.site))
            owners = [("achat", r) for r in sorted(set(g.rayon))] + [("magasin", x) for x in sites]
            floor = _num(s["bulk_floor_rate"] * 100, 0)
            a = Action("gros", URG_TODAY if abs(stake) >= today else URG_WEEK,
                       f"Ventes en gros sous le plancher de marge ({len(g)} ligne{'s' if len(g) > 1 else ''})",
                       stake / 1000, "marge semaine", f"Taux {w.rayon} {site_name_only(w.site)}", trail,
                       ". ".join(parts) + ".", "Ventes en gros sans plancher de marge.",
                       f"Suspendre toute vente en gros sous {floor} % de marge et faire valider les prix de cession "
                       "par l'Achat avant chaque vente.",
                       owners, "Aujourd'hui" if abs(stake) >= today else "Cette semaine",
                       lines=list(zip(g.rayon, g.site)))
            a.message = (f"{rec.mentions(owners)}, ventes en gros à faible marge : "
                         + "; ".join(f"{x.rayon} {x.site} à {fmt_rate(x.tm)}" for _, x in g.iterrows())
                         + f". À partir d'aujourd'hui, aucune vente en gros sous {floor} % de marge sans "
                           f"validation Achat du prix de cession. Enjeu : {fmt_k(stake / 1000)} de marge.")
            a.tags = tag_for(a.lines)
            acts.append(a)
            used |= set(a.lines)

    # base-effect lines (bulk in N-1) are explained, not actioned
    base_keys = set(tuple(x) for x in ref[ref.base_bulk][["rayon", "site"]].values) - used
    used |= base_keys

    free = ref[[(r, st_) not in used for r, st_ in zip(ref.rayon, ref.site)]]

    # ── E. Traffic collapse per site (computed first, to route volume lines to "Magasin")
    traffic_sites = []
    for site, g in ref.groupby("site"):
        dv = safe_div(g.debit.sum(), g.debit_n1.sum()) - 1
        if dv == dv and dv < -s["traffic_drop"]:
            traffic_sites.append((site, dv, g))

    # ── C. Rate effect
    rate = free[free.effet_taux <= -mat].sort_values("effet_taux")
    big = rate[rate.effet_taux <= -today]
    for _, x in big.iterrows():
        dl = _day_line(day, x.rayon, x.site)
        trail = [f"N-1 {fmt_rate(x.tm_n1)}", f"Semaine {fmt_rate(x.tm)}"]
        if dl is not None:
            trail.append(f"{hier} {fmt_rate(dl.tm)}")
        owners = [("achat", x.rayon)]
        a = Action("taux", URG_TODAY, f"{x.rayon} {x.site} : taux de marge en chute",
                   x.effet_taux / 1000, "effet taux semaine", f"Taux {x.rayon} {site_name_only(x.site)}", trail,
                   f"CA {fmt_pct(x.ca_var)} vs N-1 (effet volume {fmt_k(x.effet_volume / 1000)}), "
                   f"mais effet taux {fmt_k(x.effet_taux / 1000)}.",
                   "Le taux décroche plus que le volume : erreur de prix, promo à perte ou coût mal chargé.",
                   f"Contrôler le prix caisse contre la fiche article, les promos actives et le PMP chargé "
                   f"sur les 20 premières références {x.rayon} de {site_name_only(x.site)}.",
                   owners, "Aujourd'hui 12h", lines=[(x.rayon, x.site)])
        a.message = (f"{rec.mentions(owners)}, {x.rayon} {x.site} : taux de marge "
                     + (f"{fmt_rate(dl.tm)} {hier.lower()}, " if dl is not None else "")
                     + f"{fmt_rate(x.tm)} sur la semaine ({fmt_rate(x.tm_n1)} en N-1). Vérifie avant 12h le prix "
                       f"caisse, les promos actives et le PMP sur le top 20. Enjeu : {fmt_k(x.effet_taux / 1000)} "
                       "d'effet taux.")
        a.tags = tag_for(a.lines)
        acts.append(a)
        used.add((x.rayon, x.site))
    rest = rate[rate.effet_taux > -today]
    for rayon, g in rest.groupby("rayon"):
        stake = g.effet_taux.sum()
        tot = ref[ref.rayon == rayon]
        tr_n1, tr = safe_div(tot.marge_n1.sum(), tot.ca_n1.sum()), safe_div(tot.marge.sum(), tot.ca.sum())
        trail = [f"N-1 {fmt_rate(tr_n1)}", f"Semaine {fmt_rate(tr)}"]
        if day is not None:
            dt = day[day.rayon == rayon]
            trail.append(f"{hier} {fmt_rate(safe_div(dt.marge.sum(), dt.ca.sum()))}")
        owners = [("achat", rayon)]
        detail = ", ".join(f"{site_name_only(x.site)} {fmt_rate(x.tm_n1)} → {fmt_rate(x.tm)}"
                           for _, x in g.iterrows())
        if len(g) >= 2:
            title = f"Marge {rayon} : érosion sur {len(g)} sites"
        else:
            title = f"{rayon} {g.iloc[0].site} : taux de marge en baisse"
        a = Action("taux_groupe", URG_WEEK, title, stake / 1000, "effet taux semaine", f"Taux {rayon} PGC", trail,
                   detail + ".", "Taux en baisse à volume comparable : promos ou hausses fournisseurs non répercutées.",
                   f"Lister les promos {rayon} en cours et les hausses tarifaires reçues, et chiffrer ce qui n'a "
                   "pas été répercuté.", owners, "Revue hebdo lundi", lines=list(zip(g.rayon, g.site)))
        a.message = (f"{rec.mentions(owners)}, marge {rayon} en recul sur "
                     f"{', '.join(site_name_only(x) for x in g.site)} ({fmt_k(stake / 1000)} d'effet taux sur la "
                     "semaine). Pour la revue de lundi : liste des promos en cours et des hausses fournisseurs "
                     "non répercutées.")
        a.tags = tag_for(a.lines)
        if abs(stake) >= mat:
            acts.append(a)
            used |= set(a.lines)

    # ── D. Volume effect, tickets not collapsing => availability / offer
    traffic_site_names = {x[0] for x in traffic_sites}
    free2 = ref[[(r, st_) not in used for r, st_ in zip(ref.rayon, ref.site)]]
    vol = free2[(free2.effet_volume <= -mat) & (~free2.site.isin(traffic_site_names))].sort_values("effet_volume")
    for _, x in vol.iterrows():
        dl = _day_line(day, x.rayon, x.site)
        rupture = x.pan_var == x.pan_var and x.pan_var < -0.25
        trail = [f"Semaine {fmt_pct(x.ca_var)}"]
        if dl is not None:
            trail.append(f"{hier} {fmt_pct(dl.ca_var)}")
        trail.append(f"Panier {fmt_pct(x.pan_var, 0)}")
        owners = [("supply", "")] + ([] if rupture else [("achat", x.rayon)])
        stable = x.deb_var == x.deb_var and abs(x.deb_var) <= s["ticket_stable"]
        a = Action("volume", URG_WEEK, f"{x.rayon} {x.site} : " + ("rupture probable" if rupture else "volume en recul"),
                   x.effet_volume / 1000, "effet volume semaine", "CA vs N-1", trail,
                   f"Tickets {fmt_pct(x.deb_var, 0)}{' (stables)' if stable else ''}, panier {fmt_pct(x.pan_var, 0)}. "
                   + (f"Écart budget {fmt_k((x.ca - x.budget) / 1000)}." if x.budget else ""),
                   "Panier qui s'effondre : articles absents du rayon." if rupture
                   else "Trafic présent, panier en baisse : ruptures ou offre incomplète.",
                   f"Contrôler les ruptures sur le top 30 {x.rayon} de {site_name_only(x.site)} et relancer les "
                   "commandes bloquées." if rupture else
                   "Sortir les 50 références qui perdent le plus de volume et vérifier leur disponibilité en "
                   "rayon et en réserve.",
                   owners, "Lundi" if rupture else "Cette semaine", lines=[(x.rayon, x.site)])
        a.message = (f"{rec.mentions(owners)}, {x.rayon} {x.site} : CA {fmt_pct(x.ca_var)} vs N-1, tickets "
                     f"{fmt_pct(x.deb_var, 0)} et panier {fmt_pct(x.pan_var, 0)} sur la semaine. "
                     + ("Vérifie les ruptures sur le top 30 et relance les commandes bloquées."
                        if rupture else "Vérifie la disponibilité des 50 références qui perdent le plus de volume."))
        a.tags = tag_for(a.lines)
        acts.append(a)
        used.add((x.rayon, x.site))

    # ── E. Traffic actions
    for site, dv, g in traffic_sites:
        stake = g[[(r, x) not in used for r, x in zip(g.rayon, g.site)]].effet_volume.sum()
        dl_deb = None
        if day is not None:
            dg = day[day.site == site]
            dl_deb = safe_div(dg.debit.sum(), dg.debit_n1.sum()) - 1 if dg.debit_n1.sum() else None
        trail = [f"Semaine {fmt_pct(dv, 0)}"] + ([f"{hier} {fmt_pct(dl_deb, 0)}"] if dl_deb is not None else [])
        owners = [("magasin", site)]
        a = Action("trafic", URG_FOLLOW, f"{site} : trafic en chute", stake / 1000 if stake < 0 else None,
                   "effet volume semaine", "Tickets vs N-1", trail,
                   f"CA {fmt_pct(safe_div(g.ca.sum(), g.ca_n1.sum()) - 1)} vs N-1 sur la semaine.",
                   "À établir : travaux, concurrence, horaires ou incident.",
                   f"Inscrire {site_name_only(site)} à la revue cause racine hebdo avec le directeur de magasin.",
                   owners, "Revue cause racine", lines=list(zip(g.rayon, g.site)))
        a.message = (f"{rec.mentions(owners)}, {site} : tickets {fmt_pct(dv, 0)} vs N-1 sur la semaine. "
                     "Remonte-nous les causes connues (travaux, concurrence, horaires, incident) pour la revue de lundi.")
        n_t = traffic_streak.get(site, 0)
        a.tags = [("pers", f"Persistant · {n_t} jours")] if n_t >= 2 else [("new", f"Nouveau · {hier.lower()}")]
        acts.append(a)

    # ── F. Data issues needing an owner
    grouped: dict = {}
    for iss in [i for i in issues if i.action]:
        grouped.setdefault(iss.site or iss.title, []).append(iss)
    for key, its in list(grouped.items())[:3]:
        owners = [("cdg", "")]
        title = f"Données : {key}" if its[0].site else f"Données : {its[0].title}"
        fact = " ".join(f"{i.tag} ({i.title}) : {i.detail}" for i in its)
        a = Action("data", URG_FOLLOW, title, None, "à qualifier", " · ".join(i.tag for i in its), [],
                   fact, "Anomalie dans les exports Power BI.",
                   "Vérifier dans Power BI et relancer l'export concerné, puis le recharger dans le module.",
                   owners, "Lundi")
        a.message = f"{rec.mentions(owners)}, contrôle données sur {key} : " + " ".join(i.detail for i in its) \
            + " Peux-tu vérifier dans Power BI ?"
        a.tags = [("data", "Données")]
        acts.append(a)

    acts.sort(key=lambda a: (a.urgency, a.stake_k if a.stake_k is not None else 0))
    week_acts = [a for a in acts if a.urgency == URG_WEEK]
    overflow = week_acts[int(s.get("max_week_actions", 6)):]
    for a in overflow:
        a.urgency = -1   # secondary gap, listed but not actioned
    return acts


def since_yesterday(hist: pd.DataFrame, incomplete: set, s: dict) -> list[tuple[str, str]]:
    det = hist[hist.level == "detail"]
    dates = sorted(set(det.date))
    if len(dates) < 2:
        return []
    d1, d0 = dates[-1], dates[-2]
    out = []
    # data resolved
    for (d, site) in sorted(incomplete):
        if d == d0 and (d1, site) not in incomplete:
            ca = det[(det.date == d1) & (det.site == site)].ca.sum()
            out.append(("res", f"{site} : les ventes remontent de nouveau ({fmt_m(ca)} {weekday_fr(date.fromisoformat(d1))})."))
    a = det[det.date == d0].set_index(["rayon", "site"])
    b = det[det.date == d1].set_index(["rayon", "site"])
    th = -s["persist_k"] * 1000 * 5
    loss_a = (a.marge - a.marge_n1)
    loss_b = (b.marge - b.marge_n1)
    neg_b = b[(b.marge < 0) & (b.ca > 0)].index
    neg_a = a[(a.marge < 0) & (a.ca > 0)].index
    for k in neg_b:
        if k not in neg_a and (d1, k[1]) not in incomplete:
            out.append(("agg", f"{k[1]} : {k[0]} passe à marge négative."))
    for k in neg_a:
        if k not in neg_b and (d0, k[1]) not in incomplete:
            out.append(("res", f"{k[1]} : {k[0]} n'est plus vendue à perte."))
    new = [k for k in loss_b.index if loss_b[k] < th and (k not in loss_a.index or loss_a[k] > -s["persist_k"] * 1000)
           and (d0, k[1]) not in incomplete]
    for k in sorted(new, key=lambda k: loss_b[k])[:2]:
        rate = safe_div(b.loc[k].marge, b.loc[k].ca)
        out.append(("new", f"{k[0]} {k[1]} : {fmt_k(loss_b[k] / 1000)} de marge, taux {fmt_rate(rate)}."))
    return out[:5]


def what_works(ref: pd.DataFrame) -> list[str]:
    g = ref[(ref.dM > 0)].sort_values("dM", ascending=False).head(3)
    out = []
    for _, x in g.iterrows():
        driver = (f"tickets {fmt_pct(x.deb_var, 0)}, volume {fmt_k(x.effet_volume / 1000)}"
                  if x.effet_volume >= x.effet_taux else f"taux {fmt_rate(x.tm_n1)} → {fmt_rate(x.tm)}")
        out.append(f"<b class='pos'>{fmt_k(x.dM / 1000)}</b> · {esc(x.rayon)} {esc(x.site)} : {driver}.")
    return out


def copil_message(wt: dict, hors: dict, supeco_dM: float, acts: list[Action], av=None) -> str:
    hp = safe_div(hors.get("ca", 0), hors.get("ca_n1", 0)) - 1 if hors else float("nan")
    parts = [f"Semaine à date : PGC {fmt_pct(wt['ca_var'])} vs N-1"
             + (f" (reste du magasin {fmt_pct(hp)})" if hp == hp else "")
             + f" et {fmt_pct(wt['bgt_real'])} vs budget réel."]
    m = f"La marge {'progresse' if wt['dM'] >= 0 else 'recule'} de {fmt_k(abs(wt['dM']) / 1000, signed=False)}"
    if wt["dM"] < 0 and supeco_dM < 0 and abs(supeco_dM) >= abs(wt["dM"]) * 0.6:
        m += f", portée par les Supeco ({fmt_k(supeco_dM / 1000)})"
    parts.append(m + ".")
    if av is not None and av.promo_tot.get("ca_promo"):
        pt = av.promo_tot
        if pt["tm_promo"] == pt["tm_promo"] and pt["tm_promo"] < pt["tm_promo_n1"] - 0.02:
            parts.append(f"Le {dfr(av.day)}, les promos sortent à {fmt_rate(pt['tm_promo'])} de marge "
                         f"(N-1 {fmt_rate(pt['tm_promo_n1'])}), le hors promo à {fmt_rate(pt['tm_hp'])}.")
    today = [a for a in acts if a.urgency == URG_TODAY]
    if today:
        parts.append(f"{len(today)} décision{'s' if len(today) > 1 else ''} aujourd'hui : "
                     + " ; ".join(a.title for a in today[:4]) + ".")
    return " ".join(parts)


# ════════════════════════════════════════════════════════════════════════════
# ENGINE — articles (N-1 article is NOT iso-day: only rates are compared to N-1)
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class ArticleView:
    day: date
    lines: pd.DataFrame
    prev_day: date | None
    anomalies: pd.DataFrame
    loss_lines: pd.DataFrame
    losses: pd.DataFrame
    ruptures: pd.DataFrame
    rupture_source: str
    bulk: pd.DataFrame
    casse: pd.DataFrame
    promo: pd.DataFrame
    promo_tot: dict
    excluded_sites: list


def _rate(m, c):
    return m / c if c > 0 else float("nan")


def analyze_articles(today: ArticleExport, prev: ArticleExport | None, s: dict) -> ArticleView:
    d = today.lines.copy()
    d["tm"] = [_rate(m, c) for m, c in zip(d.marge, d.ca)]
    d["tm_n1"] = [_rate(m, c) for m, c in zip(d.marge_n1, d.ca_n1)]
    # rate effect on today's sales (N-1 volumes are not comparable, N-1 rates are)
    d["eff_taux"] = [((t - t1) * c) if (t == t and t1 == t1) else 0.0 for t, t1, c in zip(d.tm, d.tm_n1, d.ca)]
    d["promo"] = d.ca_promo > 0

    min_loss = s["article_min_loss_k"] * 1000
    anomalies = d[(d.ca > 0) & (d.tm < s["price_anomaly_rate"]) & (d.marge <= -min_loss)].sort_values("marge")
    loss_lines = d[(d.ca > 0) & (d.marge < 0) & (~d.index.isin(anomalies.index))].sort_values("marge")
    if len(loss_lines):
        losses = (loss_lines.groupby(["rayon", "code", "lib"])
                  .agg(n_sites=("site", "nunique"), sites=("site", lambda x: ", ".join(sorted(site_name_only(v) for v in set(x)))),
                       ca=("ca", "sum"), marge=("marge", "sum"),
                       marge_promo=("marge_promo", lambda x: x[x < 0].sum()), ca_promo=("ca_promo", "sum"))
                  .reset_index().sort_values("marge"))
    else:
        losses = pd.DataFrame(columns=["rayon", "code", "lib", "n_sites", "sites", "ca", "marge", "marge_promo", "ca_promo"])

    # Ruptures: sold at this site on the reference day, 0 today, still sold in >= N other sites today
    selling_today = d[d.qte > 0].groupby("code")["site"].apply(set).to_dict()
    today_qty = d.set_index(["code", "site"])["qte"].to_dict()
    rows = []
    if prev is not None and len(prev.lines):
        src = f"veille ({dfr(prev.day)})" if prev.day else "veille"
        ref = prev.lines[(prev.lines.qte > 0) & (prev.lines.ca >= s["rupture_min_ca_k"] * 1000)]
        for _, x in ref.iterrows():
            if today_qty.get((x.code, x.site), 0) > 0:
                continue
            others = selling_today.get(x.code, set()) - {x.site}
            if len(others) >= s["rupture_other_sites"]:
                rows.append({"rayon": x.rayon, "code": x.code, "lib": x.lib, "site": x.site, "ca_ref": x.ca,
                             "qte_ref": x.qte, "n_autres": len(others)})
    else:
        src = "N-1 (indicatif, non iso-jour)"
        ref = d[(d.qte <= 0) & (d.ca_n1 >= s["rupture_min_ca_k"] * 1000 * 2.5)]
        for _, x in ref.iterrows():
            others = selling_today.get(x.code, set()) - {x.site}
            if len(others) >= s["rupture_other_sites"]:
                rows.append({"rayon": x.rayon, "code": x.code, "lib": x.lib, "site": x.site, "ca_ref": x.ca_n1,
                             "qte_ref": x.qte_n1, "n_autres": len(others)})
    ruptures = pd.DataFrame(rows, columns=["rayon", "code", "lib", "site", "ca_ref", "qte_ref", "n_autres"]) \
        .sort_values("ca_ref", ascending=False)

    # Bulk sales confirmed at article level (quantity spike + low margin)
    if prev is not None and len(prev.lines):
        qref = prev.lines.set_index(["code", "site"])["qte"].to_dict()
        d["qte_ref"] = [qref.get((c, st_), 0.0) for c, st_ in zip(d.code, d.site)]
    else:
        d["qte_ref"] = d.qte_n1
    bulk = d[(d.ca > 0) & (d.qte >= s["bulk_min_qty"])
             & (d.qte >= s["bulk_qty_mult"] * d.qte_ref.clip(lower=1)) & (d.tm < s["bulk_max_rate"])] \
        .sort_values("ca", ascending=False)

    casse = d[d.casse < 0].sort_values("casse").head(20)

    pr = d.groupby("rayon")[["ca", "ca_promo", "marge_promo", "ca_promo_n1", "marge_promo_n1",
                             "ca_hp", "marge_hp", "ca_hp_n1", "marge_hp_n1", "marge", "marge_n1", "ca_n1"]].sum()
    pr = pr.reindex([r for r in RAYON_ORDER if r in pr.index])
    tot = pr.sum()
    promo_tot = {
        "ca": tot.ca, "marge": tot.marge, "ca_promo": tot.ca_promo, "poids": _rate(tot.ca_promo, tot.ca),
        "tm_promo": _rate(tot.marge_promo, tot.ca_promo), "tm_promo_n1": _rate(tot.marge_promo_n1, tot.ca_promo_n1),
        "tm_hp": _rate(tot.marge_hp, tot.ca_hp), "tm_hp_n1": _rate(tot.marge_hp_n1, tot.ca_hp_n1),
        "tm": _rate(tot.marge, tot.ca), "tm_n1": _rate(tot.marge_n1, tot.ca_n1),
        "marge_promo": tot.marge_promo,
    }
    return ArticleView(today.day, d, prev.day if prev else None, anomalies, loss_lines, losses, ruptures, src,
                       bulk, casse, pr.reset_index(), promo_tot, today.excluded_sites)


def _art_short(lib: str, n: int = 34) -> str:
    return lib if len(lib) <= n else lib[: n - 1] + "…"


def article_actions(av: ArticleView, s: dict, rec) -> list:
    acts = []
    dd = dfr(av.day)
    # Price anomalies (rate < -50 %) — one action, routed to the buyers concerned
    g = av.anomalies
    if len(g):
        w = g.iloc[0]
        rayons = [r for r in RAYON_ORDER if r in set(g.rayon)]
        owners = [("achat", r) for r in rayons]
        arts = [f"{_art_short(x.lib)} · {site_name_only(x.site)} : {fmt_k(x.marge / 1000)} pour "
                f"{fmt_k(x.ca / 1000, signed=False)} de CA (taux {fmt_rate(x.tm, 0)})" for _, x in g.head(6).iterrows()]
        by_r = "; ".join(f"{r} {len(g[g.rayon == r])}" for r in rayons)
        a = Action("prix", URG_TODAY, f"{len(g)} article{'s' if len(g) > 1 else ''} à prix anormal ({by_r})",
                   g.marge.sum() / 1000, f"marge du {dd}", f"Taux {_art_short(w.lib, 26)}",
                   ([f"N-1 {fmt_rate(w.tm_n1)}"] if w.tm_n1 == w.tm_n1 else []) + [f"{dd} {fmt_rate(w.tm, 0)}"],
                   f"Taux de marge sous {fmt_rate(s['price_anomaly_rate'], 0)} : le prix de vente est très en "
                   f"dessous du coût, sur {g.site.nunique()} magasin{'s' if g.site.nunique() > 1 else ''}.",
                   "Erreur de prix promo, de fiche article ou de coût chargé.",
                   "Corriger aujourd'hui le prix de vente (ou le coût) des articles listés et vérifier le "
                   "paramétrage des promos concernées. Liste complète dans l'Excel articles.",
                   owners, "Aujourd'hui 12h")
        a.articles, a.articles_label, a.scope = arts, f"Articles en cause · {dd}", "day"
        per = []
        for r in rayons:
            gr = g[g.rayon == r].head(3)
            per.append(f"{r} : " + ", ".join(f"{_art_short(x.lib, 28)} {site_name_only(x.site)} ({fmt_rate(x.tm, 0)})"
                                            for _, x in gr.iterrows()))
        a.message = (f"{rec.mentions(owners)}, {len(g)} articles à prix anormal le {dd} "
                     f"({fmt_k(g.marge.sum() / 1000)} de marge). " + " | ".join(per)
                     + ". Corrige les prix (ou les coûts) avant 12h, liste complète dans l'Excel articles.")
        a.tags = [("new", "Article")]
        acts.append(a)

    # Articles sold at a loss — one action per rayon
    mat = s["materiality_k"] * 1000
    for rayon, g in av.losses.groupby("rayon"):
        total = g.marge.sum()
        if abs(total) < mat:
            continue
        promo_part = g.marge_promo.sum()
        promo_share = promo_part / total if total else 0
        n_lines = len(av.loss_lines[av.loss_lines.rayon == rayon])
        arts = [f"{_art_short(x.lib)} ({x.n_sites} site{'s' if x.n_sites > 1 else ''}) : {fmt_k(x.marge / 1000)}"
                + (" · promo" if x.ca_promo > 0 else "") for _, x in g.head(5).iterrows()]
        owners = [("achat", rayon)]
        urg = URG_TODAY if abs(total) >= s["today_k"] * 1000 else URG_WEEK
        a = Action("perte_art", urg, f"{rayon} : {len(g)} articles vendus à perte",
                   total / 1000, f"marge du {dd}", "Part promo de la perte",
                   [f"{fmt_pct(promo_share, 0, signed=False)} promo", f"{fmt_pct(1 - promo_share, 0, signed=False)} hors promo"],
                   f"{n_lines} lignes article × magasin à marge négative le {dd}, dont {fmt_k(promo_part / 1000)} en promo.",
                   "Prix promo sous le coût : participation fournisseur non intégrée ou prix promo mal paramétré."
                   if promo_share >= 0.6 else
                   "Prix de vente sous le coût d'achat : hausse fournisseur non répercutée ou erreur de prix.",
                   "Revoir le prix de cession ou la participation fournisseur des promos listées, et corriger les "
                   "prix hors promo sous le coût.",
                   owners, "Aujourd'hui" if urg == URG_TODAY else "Cette semaine")
        a.articles, a.articles_label, a.scope = arts, f"Principales pertes · {dd}", "day"
        a.message = (f"{rec.mentions(owners)}, {rayon} : {len(g)} articles vendus à perte le {dd} "
                     f"({fmt_k(total / 1000)} de marge, dont {fmt_k(promo_part / 1000)} en promo). Priorités : "
                     + "; ".join(arts[:3]) + ". Revois les prix promo et la participation fournisseur.")
        a.tags = [("new", "Article")]
        acts.append(a)

    # Probable stock-outs — one action for Supply
    if len(av.ruptures):
        r = av.ruptures
        owners = [("supply", "")]
        arts = [f"{_art_short(x.lib)} · {site_name_only(x.site)} ({fmt_k(x.ca_ref / 1000, signed=False)}/jour, "
                f"vendu dans {x.n_autres} autres magasins)" for _, x in r.head(8).iterrows()]
        indic = r is not None and "N-1" in av.rupture_source
        a = Action("rupture", URG_WEEK if indic else URG_TODAY,
                   f"{len(r)} rupture{'s' if len(r) > 1 else ''} probable{'s' if len(r) > 1 else ''} "
                   f"sur {r.site.nunique()} magasin{'s' if r.site.nunique() > 1 else ''}",
                   -r.ca_ref.sum() / 1000, "CA/jour à risque", "Référence", [av.rupture_source],
                   f"Articles vendus d'habitude dans ces magasins, à 0 vente le {dd}, alors qu'ils se vendent dans "
                   f"au moins {s['rupture_other_sites']} autres magasins.",
                   "Rupture de stock, article absent du rayon ou blocage de commande.",
                   "Vérifier le stock réel et le linéaire, relancer les commandes ou organiser un transfert.",
                   owners, "Aujourd'hui" if not indic else "Cette semaine")
        a.articles, a.articles_label, a.scope = arts, f"Articles à 0 vente · {dd}", "day"
        a.message = (f"{rec.mentions(owners)}, {len(r)} ruptures probables le {dd} : " + "; ".join(arts[:5])
                     + ". Vérifie le stock et le linéaire, et relance les commandes.")
        a.tags = [("data", "Indicatif · N-1")] if indic else [("new", "Article")]
        acts.append(a)
    return acts


def attach_articles(acts: list, av: ArticleView) -> None:
    """Attach the responsible articles to rayon × site actions (same analysis day or latest article day)."""
    d = av.lines
    dd = dfr(av.day)
    rup = av.ruptures
    for a in acts:
        if not a.lines or a.kind not in ("perte", "gros", "taux", "taux_groupe", "volume"):
            continue
        keys = set(a.lines)
        sub = d[[(r, st_) in keys for r, st_ in zip(d.rayon, d.site)]]
        if sub.empty:
            continue
        items = []
        if a.kind == "perte":
            for _, x in sub[sub.marge < 0].sort_values("marge").head(5).iterrows():
                items.append(f"{_art_short(x.lib)} · {site_name_only(x.site)} : {fmt_k(x.marge / 1000)}"
                             + (" · promo" if x.promo else ""))
        elif a.kind in ("taux", "taux_groupe"):
            for _, x in sub[sub.eff_taux < 0].sort_values("eff_taux").head(5).iterrows():
                items.append(f"{_art_short(x.lib)} · {site_name_only(x.site)} : taux {fmt_rate(x.tm)} "
                             f"(N-1 {fmt_rate(x.tm_n1)}), {fmt_k(x.eff_taux / 1000)}" + (" · promo" if x.promo else ""))
        elif a.kind == "gros":
            b = av.bulk[[(r, st_) in keys for r, st_ in zip(av.bulk.rayon, av.bulk.site)]]
            for _, x in b.head(5).iterrows():
                items.append(f"{_art_short(x.lib)} · {site_name_only(x.site)} : {int(x.qte)} unités "
                             f"(réf. {int(x.qte_ref)}), taux {fmt_rate(x.tm)}")
        elif a.kind == "volume":
            rr = rup[[(r, st_) in keys for r, st_ in zip(rup.rayon, rup.site)]]
            for _, x in rr.head(5).iterrows():
                items.append(f"{_art_short(x.lib)} · {site_name_only(x.site)} : 0 vente "
                             f"(habituel {fmt_k(x.ca_ref / 1000, signed=False)}/jour)")
        if items:
            a.articles, a.articles_label = items, f"Articles en cause · {dd}"
            a.message += " Articles : " + "; ".join(items[:3]) + "."


def article_issues(av: ArticleView, hist: pd.DataFrame, latest: date) -> list:
    out = []
    if av.day and av.day != latest:
        out.append(Issue("g", "Export article", f"Export article du {dfr(av.day)}, export rayon du {dfr(latest)}",
                         f"Les articles cités dans les fiches sont ceux du {dfr(av.day)}."))
    det = hist[(hist.level == "detail") & (hist.date == (av.day.isoformat() if av.day else ""))]
    if len(det):
        a = av.lines.groupby("site").ca.sum()
        r = det.groupby("site").ca.sum()
        for site in r.index:
            if r[site] > 0 and abs(a.get(site, 0) - r[site]) / r[site] > 0.02:
                out.append(Issue("o", "Article ≠ rayon", f"{site} · {dfr(av.day)}",
                                 f"CA article {fmt_m(a.get(site, 0))} contre {fmt_m(r[site])} dans l'export rayon."))
    return out


def articles_excel(av: ArticleView) -> bytes:
    """One sheet per rayon (buyer) + one sheet for Supply. Values only, no formulas."""
    buf = io.BytesIO()

    def rows_for(rayon: str | None) -> pd.DataFrame:
        parts = []
        an = av.anomalies if rayon is None else av.anomalies[av.anomalies.rayon == rayon]
        for _, x in an.iterrows():
            parts.append(["Anomalie de prix", x.rayon, x.code, x.lib, x.site, x.ca, x.marge, x.tm, x.tm_n1, x.qte,
                          "Promo" if x.promo else "Hors promo"])
        ll = av.loss_lines if rayon is None else av.loss_lines[av.loss_lines.rayon == rayon]
        for _, x in ll.iterrows():
            parts.append(["Vendu à perte", x.rayon, x.code, x.lib, x.site, x.ca, x.marge, x.tm, x.tm_n1, x.qte,
                          "Promo" if x.promo else "Hors promo"])
        bb = av.bulk if rayon is None else av.bulk[av.bulk.rayon == rayon]
        for _, x in bb.iterrows():
            parts.append(["Vente en gros", x.rayon, x.code, x.lib, x.site, x.ca, x.marge, x.tm, x.tm_n1, x.qte,
                          f"Qté réf. {int(x.qte_ref)}"])
        cols = ["Type", "Rayon", "Code", "Article", "Magasin", "CA", "Marge", "Taux", "Taux N-1", "Qté", "Détail"]
        return pd.DataFrame(parts, columns=cols)

    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for r in RAYON_ORDER:
            df = rows_for(r)
            if len(df):
                df.to_excel(xw, sheet_name=r[:31], index=False)
        rp = av.ruptures.rename(columns={"rayon": "Rayon", "code": "Code", "lib": "Article", "site": "Magasin",
                                         "ca_ref": "CA habituel", "qte_ref": "Qté habituelle",
                                         "n_autres": "Vendu dans N autres magasins"})
        rp.to_excel(xw, sheet_name="Supply - Ruptures", index=False)
        cs = av.casse[["rayon", "code", "lib", "site", "casse", "casse_qte"]].rename(
            columns={"rayon": "Rayon", "code": "Code", "lib": "Article", "site": "Magasin",
                     "casse": "Casse (valeur)", "casse_qte": "Casse (qté)"})
        cs.to_excel(xw, sheet_name="Casse (info)", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 48)
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    hdr = ws.cell(1, c.column).value
                    if hdr in ("Taux", "Taux N-1") and isinstance(c.value, (int, float)):
                        c.number_format = "0.0%"
                    elif hdr in ("CA", "Marge", "CA habituel", "Casse (valeur)") and isinstance(c.value, (int, float)):
                        c.number_format = "#,##0"
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# HTML FRAGMENTS (shared by screen and export)
# ════════════════════════════════════════════════════════════════════════════

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&display=swap');
:root{--navy:#0A2540;--blue:#007AFF;--red:#FF3B30;--green:#34C759;--orange:#FF9500;--violet:#6A5ACD;
--bg:#F2F2F7;--ink:#1C2433;--muted:#6B7280;--line:#E3E5EC;--redbg:#FFEDEC;--greenbg:#E9F9EE;
--orangebg:#FFF4E5;--bluebg:#EAF3FF;--violetbg:#F1EEFF}
.sp,.sp *{font-family:'Nunito',-apple-system,'Segoe UI',sans-serif;box-sizing:border-box}
.sp .brief{border-radius:18px;padding:22px 24px;color:#fff;background:linear-gradient(120deg,#0A2540 0%,#1D4E89 55%,#6A5ACD 100%);display:flex;flex-wrap:wrap;gap:18px;align-items:flex-end;justify-content:space-between}
.sp .brief h1{margin:0 0 6px;font-size:15px;font-weight:700;opacity:.85;color:#fff}
.sp .brief .big{font-size:28px;font-weight:800;line-height:1.15;margin:0 0 12px;max-width:760px;color:#fff}
.sp .brief .big em{font-style:normal;color:#FFB4AE}
.sp .chips{display:flex;flex-wrap:wrap;gap:8px}
.sp .chip{background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.22);border-radius:20px;padding:5px 12px;font-size:13px;font-weight:700;color:#fff}
.sp .chip.bad{background:rgba(255,59,48,.28)}
.sp .counts{display:flex;gap:10px}
.sp .cnt2{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.22);border-radius:14px;padding:10px 14px;text-align:center;min-width:86px;color:#fff}
.sp .cnt2 b{display:block;font-size:26px;font-weight:800;line-height:1}
.sp .cnt2 span{font-size:12px;opacity:.85}
.sp .cnt2.r b{color:#FFB4AE}.sp .cnt2.o b{color:#FFD08A}.sp .cnt2.g b{color:#9FE8B4}
.sp .kline{display:flex;flex-wrap:wrap;gap:8px;margin:4px 0}
.sp .kline span{background:#fff;border:1px solid var(--line);border-radius:10px;padding:6px 10px;font-size:13px;font-weight:700;color:var(--ink)}
.sp .kline b{color:#D92D20}.sp .kline b.pos{color:#1E8E3E}
.sp .grp{font-size:13px;font-weight:800;color:var(--muted);margin:14px 0 6px;display:flex;align-items:center;gap:8px}
.sp .grp i{width:10px;height:10px;border-radius:50%;display:inline-block}
.sp .ac{background:#fff;border:1px solid var(--line);border-left:5px solid var(--red);border-radius:14px;padding:14px 16px;display:flex;flex-direction:column;gap:12px;margin-bottom:4px;color:var(--ink)}
.sp .ac.u1{border-left-color:var(--orange)}.sp .ac.u2{border-left-color:#8E96A8}
.sp .ac-top{display:flex;justify-content:space-between;gap:12px;align-items:flex-start}
.sp .ac-h{display:flex;gap:12px;align-items:flex-start}
.sp .num{width:28px;height:28px;border-radius:50%;background:var(--red);color:#fff;font-weight:800;display:grid;place-items:center;flex:none;font-size:14px}
.sp .ac.u1 .num{background:var(--orange)}.sp .ac.u2 .num{background:#8E96A8}
.sp .ac h3{margin:2px 0 6px;font-size:16.5px;font-weight:800;line-height:1.25;color:var(--ink);padding:0}
.sp .tags{display:flex;flex-wrap:wrap;gap:6px}
.sp .tg2{font-size:11.5px;font-weight:800;padding:2px 8px;border-radius:20px}
.sp .tg-pers{background:var(--redbg);color:#C8261C}.sp .tg-new{background:var(--bluebg);color:#0060D6}
.sp .tg-agg{background:#FFE1DE;color:#A21B12}.sp .tg-data{background:#ECEEF3;color:#4B5563}
.sp .stake{text-align:right;flex:none}
.sp .stake b{display:block;font-size:22px;font-weight:800;color:#D92D20;white-space:nowrap}
.sp .stake span{font-size:11.5px;color:var(--muted)}
.sp .stake.muted b{color:#9AA3B2}
.sp .trail{display:flex;flex-wrap:wrap;align-items:center;gap:6px;background:var(--bg);border-radius:10px;padding:8px 10px;font-size:13px}
.sp .trail .tl{font-weight:700;color:var(--muted);margin-right:4px}
.sp .trail .st{background:#fff;border:1px solid var(--line);border-radius:8px;padding:2px 8px;font-weight:700}
.sp .trail .st.hot{background:var(--redbg);border-color:#FFD2CE;color:#B42318}
.sp .ac.u2 .trail .st.hot{background:#ECEEF3;border-color:var(--line);color:var(--ink)}
.sp .trail .ar{color:#9AA3B2}
.sp .arts{border:1px dashed #D5D9E2;border-radius:10px;padding:8px 12px;font-size:13px}
.sp .arts b{font-size:11.5px;color:var(--muted);font-weight:800}
.sp .arts ul{margin:4px 0 0;padding-left:18px}.sp .arts li{margin:2px 0}
.sp .ac-body{margin:0;display:grid;grid-template-columns:1fr 1fr 1.3fr;gap:12px}
.sp .ac-body dt{font-size:11.5px;font-weight:800;color:var(--muted);margin-bottom:2px}
.sp .ac-body dd{margin:0;font-size:13.5px}
.sp .ac-body .do{background:var(--bluebg);border-radius:10px;padding:8px 10px}
.sp .ac-body .do dd{font-weight:700;color:var(--navy)}
.sp .ac-foot{display:flex;flex-wrap:wrap;align-items:center;gap:14px;border-top:1px solid var(--line);padding-top:10px;font-size:13px;font-weight:700}
.sp .ac-foot .due{color:var(--muted)}
.sp .mini{background:#fff;border:1px solid var(--line);border-radius:14px;padding:14px 16px;height:100%;color:var(--ink)}
.sp .mini h3{margin:0 0 8px;font-size:15px;font-weight:800;padding:0}
.sp .mini ul{margin:0;padding:0;list-style:none;display:flex;flex-direction:column;gap:8px;font-size:13.5px}
.sp .res{color:#1E8E3E;font-weight:800}.sp .newt{color:#0060D6;font-weight:800}.sp .aggt{color:#C8261C;font-weight:800}
.sp .pos{color:#1E8E3E}.sp .neg{color:#D92D20}
.sp .quote{background:var(--navy);color:#fff;border-radius:14px;padding:16px 18px;font-size:15.5px;margin:6px 0}
.sp .dq{background:#fff;border:1px solid #FFD8A8;border-left:4px solid var(--orange);border-radius:14px;padding:14px 16px;color:var(--ink)}
.sp .dq.ok{border-color:#CDEFD8;border-left-color:var(--green)}
.sp .dq b.h{font-size:15px}
.sp .dq ul{list-style:none;margin:10px 0 0;padding:0;display:flex;flex-direction:column;gap:8px}
.sp .dq li{display:flex;gap:10px;align-items:flex-start;font-size:13.5px}
.sp .dq li small{display:block;color:var(--muted);font-size:12.5px}
.sp .tg{flex:none;font-size:11.5px;font-weight:800;padding:2px 8px;border-radius:20px;white-space:nowrap}
.sp .tg.r{background:var(--redbg);color:#C8261C}.sp .tg.o{background:var(--orangebg);color:#9A5A00}
.sp .tg.g{background:var(--greenbg);color:#1E8E3E}.sp .tg.i{background:#ECEEF3;color:#4B5563}
.sp .kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
.sp .kpi{border-radius:14px;padding:14px;display:flex;flex-direction:column;gap:6px;color:var(--ink)}
.sp .kpi.b{background:var(--bluebg)}.sp .kpi.v{background:var(--violetbg)}
.sp .kpi .top{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--muted);font-weight:700}
.sp .ico{width:26px;height:26px;border-radius:50%;display:grid;place-items:center;color:#fff;font-size:13px;font-weight:800}
.sp .kpi.b .ico{background:var(--blue)}.sp .kpi.v .ico{background:var(--violet)}
.sp .kpi .val{font-size:24px;font-weight:800;white-space:nowrap}
.sp .pill{align-self:flex-start;font-size:12.5px;font-weight:800;padding:2px 9px;border-radius:20px}
.sp .pill.neg{background:var(--redbg);color:#C8261C}.sp .pill.pos{background:var(--greenbg);color:#1E8E3E}
.sp .kpi small{font-size:12px;color:var(--muted)}
.sp .note{background:var(--violetbg);border:1px solid #DCD5FF;border-radius:14px;padding:12px 16px;font-size:14px;color:var(--ink)}
.sp table.rt{width:100%;border-collapse:collapse;font-size:14px;background:#fff;border-radius:14px;overflow:hidden}
.sp .rt th{font-size:12px;color:var(--muted);text-align:right;font-weight:700;padding:8px;border-bottom:1px solid var(--line)}
.sp .rt th:first-child,.sp .rt td:first-child{text-align:left}
.sp .rt td{padding:9px 8px;border-bottom:1px solid var(--line);text-align:right;color:var(--ink)}
.sp .days{display:inline-flex;gap:3px;vertical-align:middle}
.sp .days i{width:14px;height:14px;border-radius:4px;display:inline-block}
.sp table.mx{border-collapse:separate;border-spacing:4px;font-size:12.5px}
.sp .mx thead th{height:86px;vertical-align:bottom;padding:0}
.sp .mx thead th span{display:inline-block;writing-mode:vertical-rl;transform:rotate(180deg);font-weight:700;color:var(--muted);white-space:nowrap}
.sp .mx tbody th{text-align:left;font-weight:800;padding-right:8px;white-space:nowrap;color:var(--ink)}
.sp .mx td{padding:0;white-space:nowrap}
.sp .mx td i{display:inline-block;width:12px;height:24px;border-radius:3px;margin-right:2px}
.sp .c-r3{background:#E5302A}.sp .c-r2{background:#FF7A70}.sp .c-r1{background:#FFC7C2}.sp .c-0{background:#E7E9EF}
.sp .c-g1{background:#BFEBCB}.sp .c-g2{background:#34C759}
.sp .c-na{background:repeating-linear-gradient(45deg,#D4D7E0 0 3px,#F2F2F7 3px 6px)}
.sp .legend{display:flex;flex-wrap:wrap;gap:12px;font-size:12px;color:var(--muted);margin-top:6px}
.sp .legend span{display:inline-flex;align-items:center;gap:6px}
.sp .sw{width:12px;height:12px;border-radius:3px;display:inline-block}
.sp .mxw{overflow-x:auto;background:#fff;border:1px solid var(--line);border-radius:14px;padding:12px}
@media (max-width:900px){.sp .ac-body{grid-template-columns:1fr}.sp .kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media (max-width:480px){.sp .brief .big{font-size:22px}.sp .stake b{font-size:18px}}
"""


def html_action(a: Action, i: int, rec: Recipients, with_copy: bool = False) -> str:
    tags = "".join(f'<span class="tg2 tg-{c}">{esc(l)}</span>' for c, l in a.tags)
    if a.stake_k is not None:
        stake = f'<div class="stake"><b>{fmt_k(a.stake_k)}</b><span>{esc(a.stake_label)}</span></div>'
    else:
        stake = '<div class="stake muted"><b>—</b><span>à qualifier</span></div>'
    trail = ""
    if a.trail:
        pills = []
        for j, v in enumerate(a.trail):
            cls = "st hot" if j == len(a.trail) - 1 else "st"
            pills.append(f'<span class="{cls}">{esc(v)}</span>')
        trail = (f'<div class="trail"><span class="tl">{esc(a.trail_label)}</span>'
                 + '<span class="ar">→</span>'.join(pills) + "</div>")
    arts = ""
    if a.articles:
        arts = (f'<div class="arts"><b>{esc(a.articles_label)}</b><ul>'
                + "".join(f"<li>{esc(x)}</li>" for x in a.articles) + "</ul></div>")
    copy = ""
    if with_copy:
        copy = (f'<button type="button" class="copy" data-msg="{esc(a.message)}">Copier le message</button>')
    return f"""
<article class="ac u{a.urgency}" data-owner="{' '.join(sorted(a.filters))}">
  <div class="ac-top"><div class="ac-h"><span class="num">{i}</span><div><h3>{esc(a.title)}</h3><div class="tags">{tags}</div></div></div>{stake}</div>
  {trail}{arts}
  <dl class="ac-body">
    <div><dt>Constat</dt><dd>{esc(a.fact)}</dd></div>
    <div><dt>Cause probable</dt><dd>{esc(a.cause)}</dd></div>
    <div class="do"><dt>Action</dt><dd>{esc(a.todo)}</dd></div>
  </dl>
  <div class="ac-foot"><span>👤 {esc(rec.owner_text(a.owners))}</span><span class="due">⏱ {esc(a.due)}</span>{copy}</div>
</article>"""


def html_brief(acts: list[Action], day_date: date | None, week_lab: str, n_pers: int, n_res: int,
               n_dq: int) -> str:
    cnt = {u: sum(1 for a in acts if a.urgency == u) for u in URG_LABELS}
    stake = sum(a.stake_k for a in acts if a.stake_k is not None and a.stake_k < 0
                and a.scope == "week" and a.kind != "rupture")
    title_d = f"{weekday_fr(day_date + timedelta(days=1))} {dfr(day_date + timedelta(days=1))}" if day_date else ""
    chips = [f'<span class="chip">{esc(week_lab)}</span>']
    if n_pers:
        chips.append(f'<span class="chip bad">{n_pers} écart{"s" if n_pers > 1 else ""} persistant{"s" if n_pers > 1 else ""}</span>')
    if n_res:
        chips.append(f'<span class="chip">{n_res} résolu{"s" if n_res > 1 else ""} depuis hier</span>')
    if n_dq:
        chips.append(f'<span class="chip bad">⚠ {n_dq} contrôle{"s" if n_dq > 1 else ""} de données</span>')
    return f"""
<header class="brief"><div>
  <h1>📊 Synthèse PGC · brief du {esc(title_d)}</h1>
  <p class="big">{len(acts)} action{'s' if len(acts) > 1 else ''} ce matin, dont {cnt[URG_TODAY]} à traiter aujourd'hui.
  <em>{fmt_k(stake, signed=False) if stake >= 0 else fmt_k(-stake, signed=False)} de marge en jeu</em> sur la semaine.</p>
  <div class="chips">{''.join(chips)}</div></div>
  <div class="counts"><div class="cnt2 r"><b>{cnt[URG_TODAY]}</b><span>Aujourd'hui</span></div>
  <div class="cnt2 o"><b>{cnt[URG_WEEK]}</b><span>Cette semaine</span></div>
  <div class="cnt2 g"><b>{cnt[URG_FOLLOW]}</b><span>À suivre</span></div></div>
</header>"""


def html_kline(wt: dict, hors: dict, dt: dict | None, av=None) -> str:
    hp = safe_div(hors.get("ca", 0), hors.get("ca_n1", 0)) - 1 if hors else float("nan")
    items = [f"CA semaine {fmt_m(wt['ca'])} · {fmt_pct(wt['ca_var'])} vs N-1"]
    if hp == hp:
        items.append(f"Reste du magasin <b class='{'pos' if hp >= 0 else ''}'>{fmt_pct(hp)}</b>")
    items.append(f"Budget réel <b>{fmt_pct(wt['bgt_real'])}</b>")
    items.append(f"Marge semaine <b class='{'pos' if wt['dM'] >= 0 else ''}'>{fmt_k(wt['dM'] / 1000)}</b>")
    if dt:
        items.append(f"Hier : budget réel <b>{fmt_pct(dt['bgt_real'])}</b>")
    if av is not None and av.promo_tot.get("ca_promo"):
        pt = av.promo_tot
        cls = "pos" if pt["tm_promo"] >= pt["tm_promo_n1"] else ""
        items.append(f"Promo {dfr(av.day)} : taux <b class='{cls}'>{fmt_rate(pt['tm_promo'])}</b> "
                     f"(N-1 {fmt_rate(pt['tm_promo_n1'])})")
    return '<div class="kline">' + "".join(f"<span>{x}</span>" for x in items) + "</div>"


def html_dq(issues: list[Issue]) -> str:
    shown = [i for i in issues if i.level in ("r", "o", "g")]
    if not shown:
        return ('<div class="dq ok"><b class="h">Contrôle des données</b> <span class="tg g">OK</span>'
                '<ul><li>Tous les magasins et rayons sont présents, aucune incohérence entre la semaine et les jours.</li></ul></div>')
    items = "".join(
        f'<li><span class="tg {i.level}">{esc(i.tag)}</span><div><b>{esc(i.title)}</b>'
        + (f"<small>{esc(i.detail)}</small>" if i.detail else "") + "</div></li>" for i in shown)
    n = sum(1 for i in shown if i.level in ("r", "o"))
    return (f'<div class="dq"><b class="h">Contrôle des données</b> '
            + (f'<span class="tg o">{n} point{"s" if n > 1 else ""} à vérifier</span>' if n else "")
            + f"<ul>{items}</ul></div>")


def rate_cell(v_k: float, zero: bool = False) -> str:
    if zero:
        return "c-na"
    if v_k <= -150:
        return "c-r3"
    if v_k <= -50:
        return "c-r2"
    if v_k < -10:
        return "c-r1"
    if v_k < 10:
        return "c-0"
    if v_k < 50:
        return "c-g1"
    return "c-g2"


def html_matrix(hist: pd.DataFrame, max_days: int = 7) -> str:
    det = hist[hist.level == "detail"]
    dates = sorted(set(det.date))[-max_days:]
    sites = [x for x in SITE_ORDER if x in set(det.site)] + sorted(set(det.site) - set(SITE_ORDER))
    head = "".join(f"<th><span>{esc(site_name_only(x))}</span></th>" for x in sites)
    body = ""
    for r in RAYON_ORDER:
        body += f"<tr><th scope='row'>{esc(r)}</th>"
        for sname in sites:
            cells = ""
            for d in dates:
                g = det[(det.date == d) & (det.rayon == r) & (det.site == sname)]
                if g.empty:
                    cells += '<i class="c-na" title="absent"></i>'
                    continue
                x = g.iloc[0]
                zero = x.ca == 0 and x.debit == 0
                v = (x.marge - x.marge_n1) / 1000
                tip = f"{r} · {sname} · {dfr(date.fromisoformat(d))} : " + ("sans vente" if zero else fmt_k(v))
                cells += f'<i class="{rate_cell(v, zero)}" title="{esc(tip)}"></i>'
            body += f"<td>{cells}</td>"
        body += "</tr>"
    legend = ('<div class="legend"><span><i class="sw c-r3"></i>≤ −150 k</span><span><i class="sw c-r2"></i>−150 à −50 k</span>'
              '<span><i class="sw c-r1"></i>−50 à −10 k</span><span><i class="sw c-0"></i>± 10 k</span>'
              '<span><i class="sw c-g1"></i>+10 à +50 k</span><span><i class="sw c-g2"></i>&gt; +50 k</span>'
              '<span><i class="sw c-na"></i>Sans vente</span></div>')
    days_lab = " · ".join(dfr(date.fromisoformat(d)) for d in dates)
    return (f'<div class="mxw"><table class="mx"><thead><tr><th></th>{head}</tr></thead><tbody>{body}</tbody></table>'
            f'{legend}<div class="legend">Une case par jour, dans l\'ordre : {esc(days_lab)}</div></div>')


def export_html(brief: str, kline: str, dq: str, acts: list[Action], rec: Recipients,
                since: list, works: list, copil: str) -> str:
    groups = ""
    i = 0
    sec = [a for a in acts if a.urgency < 0]
    for u in (URG_TODAY, URG_WEEK, URG_FOLLOW):
        ga = [a for a in acts if a.urgency == u]
        if not ga:
            continue
        st_ = sum(a.stake_k for a in ga if a.stake_k is not None and a.kind != "rupture")
        col = {0: "var(--red)", 1: "var(--orange)", 2: "#8E96A8"}[u]
        groups += (f'<p class="grp"><i style="background:{col}"></i>{URG_LABELS[u]}'
                   + (f" · {fmt_k(st_)} en jeu" if st_ else "") + "</p>")
        for a in ga:
            i += 1
            groups += html_action(a, i, rec, with_copy=True)
    since_html = "".join(f"<li>{_since_label(k)} · {esc(t)}</li>" for k, t in since) or "<li>Pas de journée précédente chargée.</li>"
    works_html = "".join(f"<li>{w}</li>" for w in works) or "<li>—</li>"
    return f"""<!doctype html><html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Synthèse PGC</title><style>{CSS}
body{{margin:0;background:#F2F2F7}} .wrap{{max-width:1100px;margin:0 auto;padding:20px 16px 40px;display:flex;flex-direction:column;gap:14px}}
.copy{{margin-left:auto;font:inherit;font-size:12.5px;font-weight:800;border:1.5px solid #0A2540;background:#fff;color:#0A2540;padding:4px 10px;border-radius:9px;cursor:pointer}}
.copy.ok{{background:#34C759;border-color:#34C759;color:#fff}}
.two{{display:grid;grid-template-columns:1fr 1fr;gap:12px}} @media(max-width:800px){{.two{{grid-template-columns:1fr}}}}
</style></head><body><div class="sp"><div class="wrap">
{brief}{kline}{dq}{groups}
{secondary_html(sec)}
<div class="two"><div class="mini"><h3>Depuis l'export d'hier</h3><ul>{since_html}</ul></div>
<div class="mini"><h3>Ce qui marche, à dupliquer</h3><ul>{works_html}</ul></div></div>
<blockquote class="quote">{esc(copil)}</blockquote>
</div></div>
<script>
document.querySelectorAll('.copy').forEach(function(b){{b.addEventListener('click',function(){{
  try{{navigator.clipboard.writeText(b.dataset.msg).then(function(){{b.classList.add('ok');b.textContent='Copié';
  setTimeout(function(){{b.classList.remove('ok');b.textContent='Copier le message';}},1500);}});}}catch(e){{}}
}});}});
</script></body></html>"""


def _since_label(k: str) -> str:
    return {"res": "<span class='res'>Résolu</span>", "new": "<span class='newt'>Nouveau</span>",
            "agg": "<span class='aggt'>Aggravé</span>"}.get(k, "<b>Stable</b>")


def history_csv(hist: pd.DataFrame) -> bytes:
    h = hist.copy()
    det = h.level == "detail"
    tm = (h.marge / h.ca).where(h.ca > 0, 0.0)
    tm1 = (h.marge_n1 / h.ca_n1).where(h.ca_n1 > 0, 0.0)
    h["effet_volume"] = ((h.ca - h.ca_n1) * (tm + tm1) / 2).where(det, 0.0).round(0)
    h["effet_taux"] = ((tm - tm1) * (h.ca + h.ca_n1) / 2).where(det, 0.0).round(0)
    for c in HISTORY_COLS:
        if c not in h.columns:
            h[c] = ""
    return h[HISTORY_COLS].to_csv(sep=";", index=False, decimal=",").encode("utf-8-sig")


def load_recipients(uploaded) -> pd.DataFrame | None:
    if uploaded is not None:
        try:
            return pd.read_csv(uploaded, sep=None, engine="python", encoding="utf-8-sig")
        except Exception:  # noqa: BLE001
            return None
    here = Path(__file__).resolve().parent
    for p in (here / "destinataires_pgc.csv", here.parent / "destinataires_pgc.csv",
              here.parent / "data" / "destinataires_pgc.csv"):
        if p.exists():
            try:
                return pd.read_csv(p, sep=None, engine="python", encoding="utf-8-sig")
            except Exception:  # noqa: BLE001
                return None
    return None


# ════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ════════════════════════════════════════════════════════════════════════════

PAGE_CSS = """
<style>
html, body, [class*="css"], .stMarkdown, .stText {font-family:'Nunito',-apple-system,'Segoe UI',sans-serif}
.stApp{background:#F2F2F7}
.block-container{padding-top:1.6rem;max-width:1250px}
[data-testid="stTabs"] [role="tablist"],.stTabs [data-baseweb="tab-list"]{background:#E4E5EB;border-radius:11px;padding:3px;gap:2px;width:fit-content;margin:10px 0 6px}
[data-testid="stTabs"] [role="tab"],.stTabs [data-baseweb="tab"]{border-radius:9px;padding:7px 18px !important;height:auto;font-weight:700;background:transparent}
[data-testid="stTabs"] [role="tab"] p{font-weight:700;font-size:15px}
[data-testid="stTabs"] [role="tab"][aria-selected="true"]{background:#fff;box-shadow:0 1px 3px rgba(10,37,64,.15)}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] p{color:#0A2540}
[data-baseweb="tab-highlight"],[data-baseweb="tab-border"]{display:none !important}
.sb-brand{display:flex;align-items:center;gap:10px;margin-bottom:6px}
.sb-brand .logo{width:36px;height:36px;border-radius:10px;background:linear-gradient(135deg,#0A2540,#1D4E89 60%,#6A5ACD);color:#fff;display:grid;place-items:center;font-weight:800}
.sb-brand b{display:block;font-size:15px}.sb-brand small{color:#6B7280;font-size:12px}
.alert-card{background:#EAF3FF;border:1px solid #CFE3FF;border-left:4px solid #007AFF;border-radius:14px;padding:14px 16px;margin-bottom:12px}
.col-required{display:inline-block;background:#fff;border:1px solid #E3E5EC;border-radius:8px;padding:3px 8px;margin:3px;font-size:13px;font-weight:700}
div.stDownloadButton > button{background:#0A2540;color:#fff;border-radius:11px;font-weight:800;border:0}
div.stDownloadButton > button:hover{background:#1D4E89;color:#fff}
</style>
"""


def sidebar() -> tuple:
    with st.sidebar:
        st.markdown('<div class="sb-brand"><div class="logo">SB</div><div><b>SmartBuyer Hub</b>'
                    '<small>Achats PGC · Carrefour CI</small></div></div>', unsafe_allow_html=True)
        st.markdown("### Import fichiers")
        files = st.file_uploader(
            "Dépose tous tes fichiers", type=["xlsx", "csv"], accept_multiple_files=True, key="all_files",
            help="Exports « Hier », « Cette semaine », export article, historique CSV, destinataires : "
                 "le module reconnaît chaque type tout seul.")
        with st.expander("Réglages des seuils"):
            s = dict(DEFAULT_SETTINGS)
            s["bulk_basket_mult"] = st.number_input("Vente en gros : panier N ≥ × panier N-1", 1.2, 10.0,
                                                    DEFAULT_SETTINGS["bulk_basket_mult"], 0.1)
            s["bulk_max_rate"] = st.number_input("Vente en gros : marge maximale (%)", 0.0, 20.0,
                                                 DEFAULT_SETTINGS["bulk_max_rate"] * 100, 0.5) / 100
            s["bulk_floor_rate"] = st.number_input("Plancher de marge B2B cité dans les consignes (%)", 0.0, 30.0,
                                                   DEFAULT_SETTINGS["bulk_floor_rate"] * 100, 0.5) / 100
            s["materiality_k"] = st.number_input("Seuil de matérialité d'une action (k FCFA)", 0.0, 1000.0,
                                                 DEFAULT_SETTINGS["materiality_k"], 5.0)
            s["today_k"] = st.number_input("Seuil « Aujourd'hui » (k FCFA)", 0.0, 5000.0,
                                           DEFAULT_SETTINGS["today_k"], 10.0)
            s["persist_k"] = st.number_input("Persistance : perte de marge par jour (k FCFA)", 0.0, 500.0,
                                             DEFAULT_SETTINGS["persist_k"], 5.0)
            s["traffic_drop"] = st.number_input("Chute de trafic : tickets < (%)", 0.0, 90.0,
                                                DEFAULT_SETTINGS["traffic_drop"] * 100, 5.0) / 100
            s["max_week_actions"] = int(st.number_input("Nombre maximum d'actions « Cette semaine »", 1, 30,
                                                        DEFAULT_SETTINGS["max_week_actions"], 1))
            st.markdown("**Articles**")
            s["price_anomaly_rate"] = st.number_input("Anomalie de prix : taux de marge < (%)", -500.0, 0.0,
                                                      DEFAULT_SETTINGS["price_anomaly_rate"] * 100, 5.0) / 100
            s["article_min_loss_k"] = st.number_input("Anomalie de prix : perte minimale (k FCFA)", 0.0, 500.0,
                                                      DEFAULT_SETTINGS["article_min_loss_k"], 1.0)
            s["rupture_min_ca_k"] = st.number_input("Rupture : CA habituel minimum (k FCFA / jour)", 0.0, 1000.0,
                                                    DEFAULT_SETTINGS["rupture_min_ca_k"], 5.0)
            s["rupture_other_sites"] = int(st.number_input("Rupture : vendu dans au moins N autres magasins", 1, 11,
                                                           DEFAULT_SETTINGS["rupture_other_sites"], 1))
            s["bulk_qty_mult"] = st.number_input("Vente en gros article : quantité ≥ × référence", 1.5, 20.0,
                                                 DEFAULT_SETTINGS["bulk_qty_mult"], 0.5)
            s["bulk_min_qty"] = st.number_input("Vente en gros article : quantité minimale", 1.0, 1000.0,
                                                DEFAULT_SETTINGS["bulk_min_qty"], 5.0)
        st.caption("Module 25 · Synthèse PGC — calculs Python, exports sans formule.")
    return files or [], s


KIND_LABEL = {"day": ("Hier", "#007AFF"), "week": ("Semaine", "#6A5ACD"), "article": ("Article", "#1E8E3E"),
              "history": ("Historique", "#0A2540"), "recipients": ("Destinataires", "#6B7280"),
              "unknown": ("Non reconnu", "#FF3B30")}


def sidebar_recap(recog: list, dates: list, wref, av) -> None:
    with st.sidebar:
        st.markdown("---")
        st.markdown("**Fichiers reconnus**")
        html_rows = ""
        for name, kind, detail in recog:
            lab, col = KIND_LABEL[kind]
            html_rows += (f'<div style="display:flex;gap:8px;align-items:flex-start;margin:4px 0;font-size:12.5px">'
                          f'<span style="background:{col};color:#fff;border-radius:6px;padding:1px 7px;font-weight:800;'
                          f'font-size:11px;white-space:nowrap">{lab}</span><span>{esc(name)}'
                          + (f'<br><span style="color:#6B7280">{esc(detail)}</span>' if detail else "") + "</span></div>")
        st.markdown(html_rows, unsafe_allow_html=True)
        if dates:
            st.markdown(f"**Jours rayon :** {', '.join(dfr(date.fromisoformat(d)) for d in dates[-10:])}")
        if wref:
            st.markdown(f"**{wref.label}**")
        if av:
            st.markdown(f"**Articles :** {weekday_fr(av.day)} {dfr(av.day)}"
                        + (f" (réf. ruptures : {dfr(av.prev_day)})" if av.prev_day else ""))



LANDING_CSS = """
<style>
.lp-title{display:flex;align-items:center;gap:10px;margin:8px 0 2px}
.lp-title h1{font-size:30px;font-weight:800;margin:0;padding:0;letter-spacing:-.5px;color:#1C2433}
.lp-title .emo{font-size:26px}
.lp-sub{color:#6B7280;font-size:14px;margin:0 0 18px;max-width:1100px;line-height:1.5}
.lp-info{background:#F3F7FE;border:1px solid #E1EAFB;border-left:5px solid #2F6FEB;border-radius:18px;
  padding:12px 18px;font-size:14px;line-height:1.55;color:#1C2433;margin-bottom:22px}
.lp-info b.t{display:block;font-size:14.5px;margin-bottom:2px}
.lp-lab{font-size:11px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:#6B7280;margin:0 0 10px}
.lp-card{background:#fff;border:1px solid #E7E9EF;border-radius:14px;padding:14px 18px;margin-bottom:10px}
.lp-card h4{margin:0 0 6px;font-size:15px;font-weight:800;color:#1C2433;padding:0}
.lp-card p{margin:0;font-size:13px;line-height:1.5;color:#3B4252}
.lp-rule{border-radius:14px;padding:12px 18px;margin-bottom:10px;border:1px solid}
.lp-rule .bd{display:inline-block;color:#fff;font-weight:800;font-size:12.5px;padding:4px 10px;border-radius:8px;margin-bottom:6px}
.lp-rule p{margin:0;color:#6B7280;font-size:13px}
.lp-rule p b{color:#1C2433}
.r-navy{background:#F2F6FD;border-color:#D9E4F7}.r-navy .bd{background:#1F3A5F}
.r-green{background:#F1FBF4;border-color:#D3EFDB}.r-green .bd{background:#1E5631}
.r-violet{background:#F6F2FD;border-color:#E3D8F6}.r-violet .bd{background:#5B2C83}
.r-red{background:#FEF3F2;border-color:#F8D7D3}.r-red .bd{background:#9B1C1C}
.lp-steps{background:#F1FBF4;border:1px solid #D3EFDB;border-left:5px solid #34C759;border-radius:18px;
  padding:12px 18px;font-size:13.5px;line-height:1.65;color:#1C2433}
.lp-steps ol{margin:0;padding-left:20px}.lp-steps li::marker{font-weight:800}
.lp-exp{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-bottom:18px}
.lp-exp .lp-card{margin:0}
.lp-exp .req{display:inline-block;font-size:10.5px;font-weight:800;padding:2px 8px;border-radius:20px;margin-left:6px;vertical-align:middle}
.req.on{background:#FFEDEC;color:#C8261C}.req.off{background:#ECEEF3;color:#4B5563}
.lp-exp code{background:#F2F2F7;border-radius:6px;padding:1px 5px;font-size:12px;color:#1C2433}
.lp-chips{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:24px}
.lp-chips span{background:#fff;border:1px solid #E3E5EC;border-radius:8px;padding:4px 10px;font-size:12px;font-weight:700;color:#1C2433}
@media (max-width:900px){.lp-exp{grid-template-columns:1fr}.lp-title h1{font-size:24px}}
</style>
"""


def landing(s: dict | None = None) -> None:
    s = s or DEFAULT_SETTINGS
    floor = _num(s["bulk_max_rate"] * 100, 0)
    st.markdown(LANDING_CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="lp-title"><span class="emo">📊</span><h1>Synthèse PGC</h1></div>'
        '<p class="lp-sub">Brief action du matin · contrôle des données · décomposition volume / taux / mix · '
        'semaine à date et jour par jour · consignes prêtes à envoyer</p>'
        '<div class="lp-info"><b class="t">ℹ️ À quoi sert ce module ?</b>'
        "Transformer les exports Power BI du matin en <b>liste d'actions chiffrées</b> : chaque écart de marge est "
        "expliqué (<b>effet volume</b>, <b>effet taux</b>, <b>effet mix</b>), priorisé par montant en jeu, attribué "
        "à un responsable et accompagné d'un <b>message prêt à copier</b>. Aucune connexion externe : tout part des "
        "fichiers chargés dans la barre latérale.</div>", unsafe_allow_html=True)

    c1, c2 = st.columns([1.15, 1], gap="large")
    with c1:
        st.markdown('<p class="lp-lab">Contenu du module</p>', unsafe_allow_html=True)
        cards = [
            ("🎯 Actions", "Les décisions du jour, triées par urgence (aujourd'hui, cette semaine, à suivre) puis "
                          "par marge en jeu. Responsable, échéance, étiquette « Nouveau » ou « Persistant » et "
                          "consigne à copier pour chaque fiche."),
            ("📅 Analyse semaine", "Cumul officiel de la semaine vs budget réel et N-1, PGC vs reste du magasin, "
                                  "rayons avec décomposition Bennet, sites et matrice rayons × sites jour par jour."),
            ("🗓️ État du jour", "La journée de la veille : urgences marge, ventes en gros, effet de base N-1, "
                                "rayons et sites."),
            ("🔎 Articles", "Avec l'export article : anomalies de prix, articles vendus à perte (promo / hors promo), "
                           "ruptures probables, ventes en gros confirmées, casse. Export Excel par acheteur."),
            ("🛡️ Contrôle des données", "Magasin absent ou à zéro, rayon manquant, jour non chargé, doublon, "
                                       "écart entre l'export semaine et la somme des jours, article ≠ rayon."),
        ]
        st.markdown("".join(f'<div class="lp-card"><h4>{t}</h4><p>{d}</p></div>' for t, d in cards),
                    unsafe_allow_html=True)
    with c2:
        st.markdown('<p class="lp-lab">Règles clés</p>', unsafe_allow_html=True)
        rules = [
            ("r-navy", "Budget réel", "Écart calculé <b>hors Supeco</b> (non budgétés), affiché à côté de l'écart brut."),
            ("r-green", "Vente en gros",
             f"Panier ≥ <b>×{_num(s['bulk_basket_mult'], 1)}</b> le N-1 et marge < <b>{floor} %</b>."),
            ("r-violet", "Persistance",
             f"Perte de marge > <b>{_num(s['persist_k'], 0)} k</b> par jour, <b>2 jours de suite</b>."),
            ("r-red", "Priorité « Aujourd'hui »",
             f"Marge négative, anomalie de prix (taux < <b>{fmt_rate(s['price_anomaly_rate'], 0)}</b>), "
             f"ou effet taux ≥ <b>{_num(s['today_k'], 0)} k</b>."),
            ("r-navy", "N-1 article",
             "Pas iso-jour : seuls les <b>taux de marge</b> sont comparés au N-1, jamais les volumes."),
        ]
        st.markdown("".join(f'<div class="lp-rule {c}"><span class="bd">{t}</span><p>{d}</p></div>'
                            for c, t, d in rules), unsafe_allow_html=True)
        st.markdown('<p class="lp-lab" style="margin-top:18px">Fonctionnement</p>'
                    '<div class="lp-steps"><ol>'
                    "<li>Glisse <b>tous tes fichiers</b> dans la zone de dépôt : le module reconnaît chaque type.</li>"
                    "<li>Au minimum l'export <b>« Hier »</b> ; ajoute <b>« Cette semaine »</b>, l'<b>export article</b> "
                    "et ton <b>historique CSV</b> quand tu les as.</li>"
                    "<li>Traite les actions dans l'ordre : aujourd'hui, cette semaine, à suivre.</li>"
                    "<li>Télécharge la synthèse HTML, l'Excel articles par acheteur et l'historique mis à jour.</li>"
                    "</ol></div>", unsafe_allow_html=True)

    st.markdown('<p class="lp-lab" style="margin-top:22px">Fichiers attendus</p>'
                '<div class="lp-exp">'
                '<div class="lp-card"><h4>Export « Hier » <span class="req on">Obligatoire</span></h4>'
                "<p>Power BI, type d'affichage <code>Hier</code>. La date des ventes est lue dans le nom "
                "du fichier (date d'export − 1 jour).</p></div>"
                '<div class="lp-card"><h4>Export « Cette semaine » <span class="req off">Facultatif</span></h4>'
                "<p>Type d'affichage <code>Cette Semaine</code>, du lundi à la veille. Le lundi, un export vide "
                "déclenche la clôture de la semaine précédente.</p></div>"
                '<div class="lp-card"><h4>Export article <span class="req off">Facultatif</span></h4>'
                "<p>Article × magasin sur un jour, date lue dans les filtres. Charge aussi celui de la veille "
                "pour fiabiliser les ruptures.</p></div>"
                '<div class="lp-card"><h4>Historique CSV <span class="req off">Facultatif</span></h4>'
                "<p><code>historique_synthese_pgc.csv</code> téléchargé en fin de session : une ligne par jour, "
                "rayon et site.</p></div></div>", unsafe_allow_html=True)
    cols = ["Département", "Rayon", "Site", "CA", "CA N-1", "Budget", "Marge", "Marge N-1",
            "Débit", "Débit N-1", "Volume", "Volume N-1"]
    st.markdown('<p class="lp-lab">Colonnes attendues</p><div class="lp-chips">'
                + "".join(f"<span>{c}</span>" for c in cols) + "</div>", unsafe_allow_html=True)
    st.info("Glisse tes fichiers dans la zone de dépôt de la barre latérale pour démarrer.")
    st.stop()


def kpi_html(items: list) -> str:
    out = '<div class="kpis">'
    for tint, ico, label, val, pill, pill_cls, small in items:
        out += (f'<div class="kpi {tint}"><div class="top"><span class="ico">{ico}</span>{esc(label)}</div>'
                f'<div class="val">{esc(val)}</div>'
                + (f'<span class="pill {pill_cls}">{esc(pill)}</span>' if pill else "")
                + (f"<small>{esc(small)}</small>" if small else "") + "</div>")
    return out + "</div>"


def code_block(txt: str) -> None:
    try:
        st.code(txt, language=None, wrap_lines=True)
    except TypeError:  # older Streamlit
        st.code(txt, language=None)


def sp(html_str: str) -> None:
    st.markdown(f'<div class="sp">{html_str}</div>', unsafe_allow_html=True)


def main() -> None:
    st.set_page_config(page_title="Synthèse PGC", page_icon="📊", layout="wide")
    st.markdown(PAGE_CSS + f"<style>{CSS}</style>", unsafe_allow_html=True)
    files, s = sidebar()

    # ── Read & classify every dropped file
    recog, days, week, history, rec_df, arts, errors = [], [], None, None, None, [], []
    for f in files:
        kind, obj = classify_file(f.getvalue(), f.name)
        if kind in ("day", "week"):
            ef = obj
            if ef.export_ts is None:
                man = st.sidebar.date_input(f"Date d'export de {f.name}", key=f"d_{f.name}")
                assign_dates(ef, man)
            else:
                assign_dates(ef)
            if kind == "day":
                days.append(ef)
                recog.append((f.name, kind, f"ventes du {dfr(ef.sales_date)}" if ef.sales_date else ""))
            else:
                if week is None or (ef.export_ts and week.export_ts and ef.export_ts > week.export_ts):
                    week = ef
                recog.append((f.name, kind, f"{dfr(ef.week_start)} → {dfr(ef.week_end)}" if ef.week_start else ""))
        elif kind == "article":
            if obj.day is None:
                man = st.sidebar.date_input(f"Date des ventes de {f.name}", key=f"a_{f.name}")
                obj.day = man
            arts.append(obj)
            recog.append((f.name, kind, f"ventes du {dfr(obj.day)} · {len(obj.lines):,} lignes".replace(",", " ")))
        elif kind == "history":
            history = obj if history is None else pd.concat([history, obj], ignore_index=True)
            recog.append((f.name, kind, f"{obj.date.nunique()} jours"))
        elif kind == "recipients":
            rec_df = obj
            recog.append((f.name, kind, f"{len(obj)} contacts"))
        else:
            errors.append(f"{f.name} : {obj}")
            recog.append((f.name, "unknown", str(obj)[:80]))

    if not days and history is None:
        if files:
            for e in errors:
                st.error(e)
            if arts or week is not None:
                st.warning("Ajoute au moins un export rayon « Hier » (ou ton historique CSV) : c'est la base de la synthèse.")
        sidebar_recap(recog, [], None, None)
        landing(s)
    for e in errors:
        st.error(e)

    hist, merge_issues = merge_history(history, days)
    if hist.empty or not len(hist[hist.level == "detail"]):
        st.warning("Aucune journée exploitable.")
        st.stop()

    rec = Recipients(rec_df if rec_df is not None else load_recipients(None))
    q_issues, incomplete = quality_checks(hist, week, s)
    issues = merge_issues + q_issues
    det = hist[hist.level == "detail"]
    dates = sorted(set(det.date))
    latest = date.fromisoformat(dates[-1])
    wref = build_week_ref(hist, week, latest, issues)

    # Articles: latest day = analysis, previous day (if loaded) = reference for stock-outs
    av = None
    if arts:
        arts = sorted([a for a in arts if a.day], key=lambda a: a.day)
        uniq = {a.day: a for a in arts}                       # last file wins per day
        days_a = sorted(uniq)
        today_a = uniq[days_a[-1]]
        prev_a = uniq[days_a[-2]] if len(days_a) > 1 else None
        av = analyze_articles(today_a, prev_a, s)
        issues += article_issues(av, hist, latest)
    sidebar_recap(recog, dates, wref, av)

    day_lines = analyze_lines(det[det.date == dates[-1]], s)
    day_tot = totals(day_lines)
    ref = analyze_lines(wref.lines, s) if wref else day_lines
    wt = totals(ref)
    hors = wref.hors_pgc if wref else {}
    stk = streaks(hist, incomplete, s)
    acts = build_actions(ref, day_lines, latest, stk, issues, s, rec,
                         traffic_streak=traffic_streaks(hist, incomplete, s))
    if av is not None:
        attach_articles(acts, av)
        acts = article_actions(av, s, rec) + acts
        acts.sort(key=lambda a: (a.urgency if a.urgency >= 0 else 9,
                                 a.stake_k if (a.stake_k is not None and a.kind != "rupture") else 0))
        for a in [a for a in acts if a.urgency == URG_WEEK][int(s.get("max_week_actions", 6)):]:
            a.urgency = -1
    main_acts = [a for a in acts if a.urgency >= 0]
    sec_acts = [a for a in acts if a.urgency < 0]
    since = since_yesterday(hist, incomplete, s)
    works = what_works(ref)
    supeco_dM = ref[ref.fmt == "Supeco"].dM.sum()
    copil = copil_message(wt, hors, supeco_dM, main_acts, av)
    n_pers = sum(1 for a in main_acts if any(c in ("pers", "agg") for c, _ in a.tags))
    n_res = sum(1 for k, _ in since if k == "res")
    n_dq = sum(1 for i in issues if i.level in ("r", "o"))

    brief = html_brief(main_acts, latest, wref.label if wref else "", n_pers, n_res, n_dq)
    kline = html_kline(wt, hors, day_tot, av)
    dq = html_dq(issues)
    sp(brief)

    tab_names = ["Actions", "Analyse semaine", "État du jour"] + (["Articles"] if av is not None else [])
    tabs = st.tabs(tab_names)
    t_act, t_sem, t_day = tabs[0], tabs[1], tabs[2]

    # ── Actions
    with t_act:
        sp(kline + dq)
        roles = ["Toutes"] + [ROLE_LABEL[r] for r in ("achat", "supply", "magasin", "cdg")
                              if any(r in a.filters for a in main_acts)]
        pills = getattr(st, "pills", None)
        choice = (pills("Responsable", roles, default="Toutes", key="owner_filter", label_visibility="collapsed")
                  if pills else st.radio("Responsable", roles, horizontal=True, label_visibility="collapsed"))
        choice = choice or "Toutes"
        inv = {v: k for k, v in ROLE_LABEL.items()}
        shown = [a for a in main_acts if choice == "Toutes" or inv[choice] in a.filters]
        if not shown:
            st.success("Aucune action pour ce responsable.")
        num = {id(a): i + 1 for i, a in enumerate(main_acts)}
        for u in (URG_TODAY, URG_WEEK, URG_FOLLOW):
            ga = [a for a in shown if a.urgency == u]
            if not ga:
                continue
            st_ = sum(a.stake_k for a in ga if a.stake_k is not None and a.kind != "rupture")
            col = {0: "var(--red)", 1: "var(--orange)", 2: "#8E96A8"}[u]
            sp(f'<p class="grp"><i style="background:{col}"></i>{URG_LABELS[u]}'
               + (f" · {fmt_k(st_)} de marge en jeu" if st_ else "") + "</p>")
            for a in ga:
                sp(html_action(a, num[id(a)], rec))
                with st.expander("✉️ Message à copier"):
                    code_block(a.message)
        sec_shown = [a for a in sec_acts if choice == "Toutes" or inv[choice] in a.filters]
        if sec_shown:
            sp(secondary_html(sec_shown))
        c1, c2 = st.columns(2)
        with c1:
            items = "".join(f"<li>{_since_label(k)} · {esc(t)}</li>" for k, t in since) \
                or "<li>Charge la journée précédente pour voir l'évolution.</li>"
            sp(f'<div class="mini"><h3>Depuis l\'export d\'hier</h3><ul>{items}</ul></div>')
        with c2:
            items = "".join(f"<li>{w}</li>" for w in works) or "<li>—</li>"
            sp(f'<div class="mini"><h3>Ce qui marche, à dupliquer</h3><ul>{items}</ul></div>')
        sp(f'<blockquote class="quote">{esc(copil)}</blockquote>')

        with st.expander("📋 Toutes les consignes, par destinataire"):
            by_person: dict = {}
            for a in shown:
                by_person.setdefault(rec.mentions(a.owners), []).append(a.message)
            txt = f"Synthèse PGC – consignes du {dfr(latest + timedelta(days=1))}\n\n" + "\n\n".join(
                "\n".join(f"{i}. {m}" for i, m in enumerate(msgs, 1)) for msgs in by_person.values())
            code_block(txt)

        stamp = (latest + timedelta(days=1)).isoformat()
        cols = st.columns(3 if av is not None else 2)
        with cols[0]:
            st.download_button("⬇️ Synthèse (HTML)",
                               export_html(brief, kline, dq, acts, rec, since, works, copil).encode("utf-8"),
                               f"synthese_pgc_{stamp}.html", "text/html", use_container_width=True)
        with cols[1]:
            st.download_button("⬇️ Historique (CSV)", history_csv(hist),
                               "historique_synthese_pgc.csv", "text/csv", use_container_width=True)
        if av is not None:
            with cols[2]:
                st.download_button("⬇️ Articles par acheteur (Excel)", articles_excel(av),
                                   f"articles_pgc_{av.day.isoformat()}.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)

    with t_sem:
        render_week(ref, wt, hors, hist, wref, s)
    with t_day:
        render_day(day_lines, day_tot, latest, hist, s, av)
    if av is not None:
        with tabs[3]:
            render_articles(av, s)


def _art_table(df: pd.DataFrame, cols: dict, pct: tuple = (), money: tuple = ()) -> None:
    if df is None or df.empty:
        st.caption("Aucun article.")
        return
    t = df[list(cols)].rename(columns=cols).copy()
    for c in money:
        if cols.get(c) in t.columns:
            t[cols[c]] = t[cols[c]].map(lambda v: fmt_k(v / 1000, signed=False) if v >= 0 else fmt_k(v / 1000))
    for c in pct:
        if cols.get(c) in t.columns:
            t[cols[c]] = t[cols[c]].map(fmt_rate)
    if "site" in cols:
        t[cols["site"]] = t[cols["site"]].map(site_name_only)
    st.dataframe(t, hide_index=True, use_container_width=True)


def render_articles(av: ArticleView, s: dict) -> None:
    st.markdown(f"#### Articles · {weekday_fr(av.day)} {dfr(av.day)}")
    st.caption("N-1 article non iso-jour : seuls les taux de marge sont comparés au N-1, jamais les volumes. "
               "Périmètre : 12 magasins"
               + (f" (exclus : {', '.join(av.excluded_sites)})" if av.excluded_sites else "") + ".")
    c1, c2 = st.columns(2)
    with c1:
        f_r = st.selectbox("Rayon", ["Tous"] + RAYON_ORDER, key="art_r")
    with c2:
        f_s = st.selectbox("Magasin", ["Tous"] + SITE_ORDER, key="art_s")

    def flt(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df
        out = df
        if f_r != "Tous" and "rayon" in out.columns:
            out = out[out.rayon == f_r]
        if f_s != "Tous" and "site" in out.columns:
            out = out[out.site == f_s]
        return out

    pt = av.promo_tot
    ll = flt(av.loss_lines)
    rp = flt(av.ruptures)
    sp(kpi_html([
        ("b", "€", "CA article du jour", fmt_m(pt["ca"]), "", "", f"Poids promo {fmt_pct(pt['poids'], 1, signed=False)}"),
        ("v", "%", "Taux promo", fmt_rate(pt["tm_promo"]), f"N-1 {fmt_rate(pt['tm_promo_n1'])}",
         "pos" if pt["tm_promo"] >= pt["tm_promo_n1"] else "neg", f"Marge promo {fmt_k(pt['marge_promo'] / 1000)}"),
        ("v", "%", "Taux hors promo", fmt_rate(pt["tm_hp"]), f"N-1 {fmt_rate(pt['tm_hp_n1'])}",
         "pos" if pt["tm_hp"] >= pt["tm_hp_n1"] else "neg", ""),
        ("b", "!", "Lignes à perte", f"{len(ll)}", fmt_k(ll.marge.sum() / 1000) if len(ll) else "", "neg",
         f"{len(rp)} rupture{'s' if len(rp) > 1 else ''} probable{'s' if len(rp) > 1 else ''}"),
    ]))

    st.markdown("##### Promo et hors promo par rayon")
    pr = av.promo.copy()
    if len(pr):
        show = pd.DataFrame({
            "Rayon": pr.rayon,
            "CA": [fmt_m(v) for v in pr.ca],
            "Poids promo": [fmt_pct(_rate(a, b), 1, signed=False) for a, b in zip(pr.ca_promo, pr.ca)],
            "Taux promo": [fmt_rate(_rate(a, b)) for a, b in zip(pr.marge_promo, pr.ca_promo)],
            "Taux promo N-1": [fmt_rate(_rate(a, b)) for a, b in zip(pr.marge_promo_n1, pr.ca_promo_n1)],
            "Taux hors promo": [fmt_rate(_rate(a, b)) for a, b in zip(pr.marge_hp, pr.ca_hp)],
            "Taux hors promo N-1": [fmt_rate(_rate(a, b)) for a, b in zip(pr.marge_hp_n1, pr.ca_hp_n1)],
            "Marge promo": [fmt_k(v / 1000) for v in pr.marge_promo],
        })
        st.dataframe(show, hide_index=True, use_container_width=True)

    st.markdown(f"##### Anomalies de prix (taux < {fmt_rate(s['price_anomaly_rate'], 0)})")
    _art_table(flt(av.anomalies), {"rayon": "Rayon", "lib": "Article", "site": "Magasin", "ca": "CA",
                                   "marge": "Marge", "tm": "Taux", "tm_n1": "Taux N-1", "qte": "Qté"},
               pct=("tm", "tm_n1"), money=("ca", "marge"))
    st.markdown("##### Articles vendus à perte")
    _art_table(ll, {"rayon": "Rayon", "lib": "Article", "site": "Magasin", "ca": "CA", "marge": "Marge",
                    "tm": "Taux", "tm_n1": "Taux N-1", "ca_promo": "CA promo"},
               pct=("tm", "tm_n1"), money=("ca", "marge", "ca_promo"))
    st.markdown(f"##### Ruptures probables · référence : {av.rupture_source}")
    _art_table(rp, {"rayon": "Rayon", "lib": "Article", "site": "Magasin", "ca_ref": "CA habituel",
                    "qte_ref": "Qté habituelle", "n_autres": "Vendu dans N autres magasins"}, money=("ca_ref",))
    st.markdown("##### Ventes en gros confirmées")
    _art_table(flt(av.bulk), {"rayon": "Rayon", "lib": "Article", "site": "Magasin", "qte": "Qté",
                              "qte_ref": "Qté réf.", "ca": "CA", "tm": "Taux"}, pct=("tm",), money=("ca",))
    st.markdown("##### Casse (information)")
    _art_table(flt(av.casse), {"rayon": "Rayon", "lib": "Article", "site": "Magasin", "casse": "Casse (valeur)",
                               "casse_qte": "Casse (qté)"}, money=("casse",))


def render_week(ref: pd.DataFrame, wt: dict, hors: dict, hist: pd.DataFrame, wref: WeekRef | None, s: dict) -> None:
    lab = wref.label if wref else "Semaine"
    src = {"export": "Chiffres officiels de l'export « Cette semaine ».",
           "history": "Reconstituée à partir des exports « Hier » chargés.",
           "closure": "Semaine précédente clôturée, reconstituée à partir des exports « Hier »."}.get(
        wref.source if wref else "", "")
    st.markdown(f"#### {lab}")
    st.caption(src + " Budget réel calculé hors Supeco.")
    sp(kpi_html([
        ("b", "€", "CA semaine", fmt_m(wt["ca"]), f"{fmt_pct(wt['ca_var'])} vs N-1",
         "pos" if wt["ca_var"] >= 0 else "neg", ""),
        ("b", "◎", "Vs budget réel", fmt_pct(wt["bgt_real"]), "", "",
         f"{fmt_m(wt['ca_budgeted'])} pour {fmt_m(wt['budget_real'])} · affiché {fmt_pct(wt['bgt_shown'])}"),
        ("v", "M", "Marge semaine", fmt_m(wt["marge"]), f"{fmt_k(wt['dM'] / 1000)} vs N-1",
         "pos" if wt["dM"] >= 0 else "neg", ""),
        ("v", "%", "Taux de marge", fmt_rate(wt["tm"]), f"{fmt_bps(wt['tm'] - wt['tm_n1'])} (N-1 {fmt_rate(wt['tm_n1'])})",
         "pos" if wt["tm"] >= wt["tm_n1"] else "neg", ""),
    ]))
    # PGC vs rest of store + formats
    hp = safe_div(hors.get("ca", 0), hors.get("ca_n1", 0)) - 1 if hors else float("nan")
    fm = ref.groupby("fmt")[["marge", "marge_n1"]].sum()
    fm_txt = " · ".join(f"{k} {fmt_k((r.marge - r.marge_n1) / 1000)}" for k, r in fm.iterrows())
    note = ""
    if hp == hp:
        gap = wt["ca_var"] - hp
        note += (f"<b>PGC vs reste du magasin :</b> {fmt_pct(wt['ca_var'])} contre {fmt_pct(hp)} "
                 f"({'+' if gap >= 0 else MINUS}{_num(abs(gap) * 100, 1)} pts). ")
    note += f"<b>Marge par format :</b> {fm_txt}."
    sp(f'<div class="note">{note}</div>')

    # Daily evolution
    det = hist[hist.level == "detail"]
    if go is not None and wref is not None:
        span = [(wref.start + timedelta(days=i)).isoformat() for i in range((wref.end - wref.start).days + 1)]
        g = det[det.date.isin(span)]
        rows = []
        for d, x in g.groupby("date"):
            b = x[~x.fmt.isin(UNBUDGETED_FORMATS)]
            rows.append({"lab": f"{weekday_fr(date.fromisoformat(d))[:3]} {date.fromisoformat(d).day}",
                         "ca": b.ca.sum() / 1e6, "bgt": b.budget.sum() / 1e6,
                         "tm": safe_div(x.marge.sum(), x.ca.sum()) * 100,
                         "tm1": safe_div(x.marge_n1.sum(), x.ca_n1.sum()) * 100, "implied": False})
        missing = [d for d in span if d not in set(g.date)]
        if wref.source == "export" and missing and rows:
            b = ref[~ref.fmt.isin(UNBUDGETED_FORMATS)]
            n = len(missing)
            ca_r = (b.ca.sum() / 1e6 - sum(r["ca"] for r in rows)) / n
            bg_r = (b.budget.sum() / 1e6 - sum(r["bgt"] for r in rows)) / n
            m_r = ref.marge.sum() - g.marge.sum()
            c_r = ref.ca.sum() - g.ca.sum()
            m1_r = ref.marge_n1.sum() - g.marge_n1.sum()
            c1_r = ref.ca_n1.sum() - g.ca_n1.sum()
            lab_m = "–".join(weekday_fr(date.fromisoformat(d))[:3] for d in (missing[0], missing[-1])) if n > 1 \
                else weekday_fr(date.fromisoformat(missing[0]))[:3]
            rows.insert(0, {"lab": f"{lab_m}*", "ca": ca_r, "bgt": bg_r, "tm": safe_div(m_r, c_r) * 100,
                            "tm1": safe_div(m1_r, c1_r) * 100, "implied": True})
        if rows:
            c1, c2 = st.columns(2)
            ev = pd.DataFrame(rows)
            with c1:
                fig = go.Figure()
                fig.add_bar(x=ev.lab, y=ev.bgt, name="Budget", marker_color="#C9D8EE")
                fig.add_bar(x=ev.lab, y=ev.ca, name="Réalisé",
                            marker_color=["#8DBDFF" if i else "#007AFF" for i in ev.implied],
                            text=[fmt_pct(c / b - 1 if b else float("nan"), 1) for c, b in zip(ev.ca, ev.bgt)],
                            textposition="outside")
                fig.update_layout(title="CA vs budget (hors Supeco, M FCFA)", barmode="group", height=320,
                                  margin=dict(l=10, r=10, t=40, b=10), plot_bgcolor="#fff", paper_bgcolor="#fff",
                                  font_family="Nunito", legend=dict(orientation="h", y=-0.15))
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = go.Figure()
                fig.add_scatter(x=ev.lab, y=ev.tm1, name="Taux N-1", mode="lines+markers",
                                line=dict(color="#9AA3B2", dash="dash"))
                fig.add_scatter(x=ev.lab, y=ev.tm, name="Taux N", mode="lines+markers+text",
                                line=dict(color="#6A5ACD", width=3),
                                text=[_num(v, 1) for v in ev.tm], textposition="top center")
                fig.update_layout(title="Taux de marge par jour (%)", height=320,
                                  margin=dict(l=10, r=10, t=40, b=10), plot_bgcolor="#fff", paper_bgcolor="#fff",
                                  font_family="Nunito", legend=dict(orientation="h", y=-0.15))
                st.plotly_chart(fig, use_container_width=True)
            if any(ev.implied):
                st.caption("* Jours non chargés : moyenne par jour déduite de l'export semaine.")

    # Rayons with Bennet
    st.markdown("#### Rayons · décomposition de l'écart de marge")
    br = bennet_group(ref, "rayon").set_index("rayon").reindex(RAYON_ORDER).dropna(how="all")
    rows = ""
    for r, x in br.iterrows():
        def cell(v):
            return f"<td class='{'pos' if v >= 0 else 'neg'}'>{fmt_k(v / 1000)}</td>"
        rows += (f"<tr><td><b>{esc(r)}</b></td><td class='{'pos' if x.ca_var >= 0 else 'neg'}'>{fmt_pct(x.ca_var)}</td>"
                 f"<td class='{'pos' if (x.bgt_real or 0) >= 0 else 'neg'}'>{fmt_pct(x.bgt_real)}</td>"
                 f"<td>{fmt_rate(x.tm_n1)} → <b>{fmt_rate(x.tm)}</b></td>"
                 f"{cell(x.effet_volume)}{cell(x.effet_taux)}{cell(x.effet_mix)}"
                 f"<td class='{'pos' if x.dM >= 0 else 'neg'}'><b>{fmt_k(x.dM / 1000)}</b></td></tr>")
    sp("<table class='rt'><thead><tr><th>Rayon</th><th>CA vs N-1</th><th>Vs budget réel</th><th>Taux N-1 → N</th>"
       "<th>Effet volume</th><th>Effet taux</th><th>Effet mix</th><th>Δ marge</th></tr></thead>"
       f"<tbody>{rows}</tbody></table>")
    st.caption("Décomposition symétrique : effet volume + effet taux + effet mix = Δ marge, sans résidu.")

    # Sites
    st.markdown("#### Sites · Δ marge vs N-1")
    bs = bennet_group(ref, "site").sort_values("dM")
    if go is not None:
        fig = go.Figure()
        fig.add_bar(y=bs.site, x=bs.effet_volume / 1000, name="Effet volume", orientation="h", marker_color="#8DBDFF")
        fig.add_bar(y=bs.site, x=(bs.effet_taux + bs.effet_mix) / 1000, name="Effet taux + mix", orientation="h",
                    marker_color="#6A5ACD")
        fig.add_scatter(y=bs.site, x=bs.dM / 1000, mode="markers+text", name="Δ marge",
                        marker=dict(color="#1C2433", size=9, symbol="diamond"),
                        text=[fmt_k(v / 1000) for v in bs.dM], textposition="middle right")
        fig.update_layout(barmode="relative", height=40 * len(bs) + 80, margin=dict(l=10, r=40, t=10, b=10),
                          plot_bgcolor="#fff", paper_bgcolor="#fff", font_family="Nunito",
                          xaxis_title="k FCFA", legend=dict(orientation="h", y=-0.08))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Matrice rayons × sites")
    st.caption("Δ marge vs N-1 par jour chargé (7 derniers jours au maximum).")
    sp(html_matrix(hist))


def render_day(day_lines: pd.DataFrame, dt: dict, latest: date, hist: pd.DataFrame, s: dict, av=None) -> None:
    st.markdown(f"#### {weekday_fr(latest).capitalize()} {dfr(latest)}")
    urg = day_lines[(day_lines.marge < 0) & (day_lines.ca > 0)]
    low = day_lines[(day_lines.tm < 0.05) & (day_lines.ca > 0) & (day_lines.tm_n1 > 0.10)]
    alerts = []
    for _, x in urg.iterrows():
        alerts.append(f"{x.rayon} {x.site} vendue à perte ({fmt_rate(x.tm)}).")
    for _, x in low.iterrows():
        if x.marge >= 0:
            alerts.append(f"{x.rayon} {x.site} à {fmt_rate(x.tm)} de marge ({fmt_rate(x.tm_n1)} en N-1).")
    if alerts:
        sp('<div class="dq" style="border-left-color:var(--red);border-color:#FFD2CE"><b class="h">Urgences marge</b>'
           "<ul>" + "".join(f"<li>{esc(a)}</li>" for a in alerts) + "</ul></div>")
    sp(kpi_html([
        ("b", "€", "CA du jour", fmt_m(dt["ca"]), f"{fmt_pct(dt['ca_var'])} vs N-1", "pos" if dt["ca_var"] >= 0 else "neg", ""),
        ("b", "◎", "Vs budget réel", fmt_pct(dt["bgt_real"]), "", "", f"Affiché {fmt_pct(dt['bgt_shown'])} avec les Supeco"),
        ("v", "M", "Marge du jour", fmt_m(dt["marge"]), f"{fmt_k(dt['dM'] / 1000)} vs N-1", "pos" if dt["dM"] >= 0 else "neg", ""),
        ("v", "%", "Taux de marge", fmt_rate(dt["tm"]), fmt_bps(dt["tm"] - dt["tm_n1"]), "pos" if dt["tm"] >= dt["tm_n1"] else "neg", ""),
    ]))
    mag = hist[(hist.level == "magasin") & (hist.date == latest.isoformat())]
    expl = []
    if len(mag) and mag.iloc[0].debit_n1:
        m = mag.iloc[0]
        dv = safe_div(m.debit, m.debit_n1) - 1
        pv = safe_div(m.ca / m.debit, m.ca_n1 / m.debit_n1) - 1
        expl.append(f"Tickets magasin {fmt_pct(dv)}, panier magasin {fmt_pct(pv)}.")
    bulk = day_lines[day_lines.bulk]
    if len(bulk):
        expl.append("Ventes en gros : " + ", ".join(f"{x.rayon} {x.site} ({fmt_rate(x.tm)})" for _, x in bulk.iterrows()) + ".")
    base = day_lines[day_lines.base_bulk]
    if len(base):
        expl.append(f"Effet de base N-1 : {fmt_m(base.ca_n1.sum())} de ventes en gros en N-1 ("
                    + ", ".join(f"{x.rayon} {site_name_only(x.site)}" for _, x in base.iterrows()) + ").")
    if av is not None and av.day == latest and av.promo_tot.get("ca_promo"):
        pt = av.promo_tot
        expl.append(f"Promo : {fmt_pct(pt['poids'], 0, signed=False)} du CA à {fmt_rate(pt['tm_promo'])} de marge "
                    f"(N-1 {fmt_rate(pt['tm_promo_n1'])}), hors promo à {fmt_rate(pt['tm_hp'])} "
                    f"(N-1 {fmt_rate(pt['tm_hp_n1'])}).")
    if expl:
        sp('<div class="note"><b>Ce qui explique la journée.</b> ' + " ".join(esc(e) for e in expl) + "</div>")
    br = bennet_group(day_lines, "rayon").set_index("rayon").reindex(RAYON_ORDER).dropna(how="all")
    rows = "".join(
        f"<tr><td><b>{esc(r)}</b></td><td class='{'pos' if x.ca_var >= 0 else 'neg'}'>{fmt_pct(x.ca_var)}</td>"
        f"<td>{fmt_pct(x.bgt_real)}</td><td>{fmt_rate(x.tm_n1)} → <b>{fmt_rate(x.tm)}</b></td>"
        f"<td class='{'pos' if x.effet_volume >= 0 else 'neg'}'>{fmt_k(x.effet_volume / 1000)}</td>"
        f"<td class='{'pos' if x.effet_taux + x.effet_mix >= 0 else 'neg'}'>{fmt_k((x.effet_taux + x.effet_mix) / 1000)}</td>"
        f"<td class='{'pos' if x.dM >= 0 else 'neg'}'><b>{fmt_k(x.dM / 1000)}</b></td></tr>"
        for r, x in br.iterrows())
    sp("<table class='rt'><thead><tr><th>Rayon</th><th>CA vs N-1</th><th>Vs budget réel</th><th>Taux N-1 → N</th>"
       f"<th>Effet volume</th><th>Effet taux + mix</th><th>Δ marge</th></tr></thead><tbody>{rows}</tbody></table>")
    st.markdown("#### Sites")
    bs = bennet_group(day_lines, "site").sort_values("dM")
    show = pd.DataFrame({
        "Site": bs.site, "CA": [fmt_m(v) for v in bs.ca], "vs N-1": [fmt_pct(v) for v in bs.ca_var],
        "vs budget": [fmt_pct(v) if v == v else "hors budget" for v in bs.bgt_real],
        "Taux N-1 → N": [f"{fmt_rate(a)} → {fmt_rate(b)}" for a, b in zip(bs.tm_n1, bs.tm)],
        "Δ marge": [fmt_k(v / 1000) for v in bs.dM],
    })
    st.dataframe(show, hide_index=True, use_container_width=True)


if __name__ == "__main__":
    main()
