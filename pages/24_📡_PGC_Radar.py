"""PGC Radar — pilotage hebdomadaire du département PGC.

Page Streamlit autonome pour SmartBuyer Hub. Les calculs sont réalisés en
Python à partir de l'onglet ``Export`` du classeur importé.
"""

from __future__ import annotations

import hashlib
import html
from io import BytesIO
from typing import Any

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="PGC Radar · SmartBuyer Hub",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ---------------------------------------------------------------------------
# Constantes métier et charte
# ---------------------------------------------------------------------------

PGC = "01 - PGC"
REAL_DEPARTMENTS = [
    "01 - PGC",
    "02 - PRODUITS FRAIS",
    "03 - BAZAR",
    "04 - EPCS",
    "06 - TEXTILE",
]
FORMAT_ORDER = ["Hyper", "Market", "Supeco"]
SITE_WEIGHT_THRESHOLD = 0.10
MATERIALITY_FLOOR = 1_000_000

CAUSES = [
    "Rupture stock fournisseur",
    "Rupture stock magasin",
    "Prix non compétitif",
    "Exécution magasin",
    "Action concurrente",
    "Effet promo / cannibalisation",
    "Saisonnalité",
    "Autre",
]

SITE_FORMATS = {
    "Hyper": ["Marcory", "Palmeraie", "Yopougon"],
    "Market": [
        "Riviera",
        "Kokoh Mall",
        "2 Plateaux",
        "7 Décembre",
        "Cité verte",
        "Aboboté",
    ],
    "Supeco": ["Niangon", "Terminus 47", "Toit rouge"],
}


def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --bg: #F2F2F7;
            --ink: #1C1C1E;
            --muted: #6E6E73;
            --line: #E5E5EA;
            --blue: #007AFF;
            --red: #FF3B30;
            --green: #34C759;
            --orange: #FF9500;
        }
        .stApp { background: var(--bg); color: var(--ink); }
        .block-container { max-width: 1380px; padding-top: 2rem; }
        [data-testid="stSidebar"] { background: var(--bg); }
        [data-testid="stSidebar"] section { padding-top: 1.4rem; }
        h1, h2, h3, p, div, span, label { font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "Helvetica Neue", Arial, sans-serif; }
        h1 { letter-spacing: -0.04em; }
        h2, h3 { letter-spacing: -0.025em; }
        .eyebrow { color: var(--muted); font-size: 0.70rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
        .muted { color: var(--muted); }
        .dashboard-card, .format-header, .site-row, .alert-shell, .empty-state {
            background: #FFFFFF; border: 1px solid var(--line); border-radius: 14px;
        }
        .dashboard-card { padding: 1.15rem 1.25rem; min-height: 145px; }
        .dashboard-card .label { color: var(--muted); font-size: .78rem; margin-bottom: .4rem; }
        .dashboard-card .value { font-size: 2.05rem; font-weight: 750; letter-spacing: -.04em; line-height: 1.1; }
        .dashboard-card .subvalue { color: var(--muted); font-size: .80rem; margin-top: .5rem; }
        .positive { color: var(--green) !important; }
        .negative { color: var(--red) !important; }
        .attention { color: var(--orange) !important; }
        .neutral { color: var(--ink) !important; }
        .budget-line { color: #8E8E93; font-size: .78rem; padding: .55rem .1rem .15rem; }
        .section-kicker { color: var(--muted); font-size: .76rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; margin: 1.5rem 0 .55rem; }
        .dept-row { display: grid; grid-template-columns: 1.8fr 1fr 1fr; gap: 1rem; align-items: center; padding: .8rem 1rem; background: #fff; border-bottom: 1px solid var(--line); }
        .dept-row:first-child { border-radius: 14px 14px 0 0; }
        .dept-row:last-child { border-radius: 0 0 14px 14px; border-bottom: 0; }
        .dept-row.pgc { border-left: 4px solid var(--blue); padding-left: .8rem; }
        .dept-row.total { font-weight: 700; border-top: 2px solid var(--line); }
        .dept-name { font-size: .92rem; font-weight: 650; }
        .dept-value { font-weight: 700; text-align: right; }
        .dept-var { text-align: right; font-weight: 700; }
        .format-header { display: grid; grid-template-columns: 2fr 1fr 1fr 1fr; gap: 1rem; align-items: center; padding: 1rem 1.15rem; margin-top: 1rem; }
        .format-title { font-size: 1.10rem; font-weight: 750; }
        .format-metric-label { color: var(--muted); font-size: .72rem; }
        .format-metric { font-weight: 750; font-size: 1.05rem; margin-top: .15rem; }
        .badge { display: inline-block; border-radius: 999px; padding: .27rem .62rem; font-size: .72rem; font-weight: 750; white-space: nowrap; }
        .badge-critique { color: #B42318; background: #FFE5E3; }
        .badge-attention { color: #9A5B00; background: #FFF0D5; }
        .badge-ok { color: #147A32; background: #DCF8E3; }
        .site-row { display: grid; grid-template-columns: 1.45fr 1fr 1fr 1.65fr; gap: 1rem; align-items: center; padding: .70rem 1.15rem; margin-top: .45rem; }
        .site-name { font-size: .88rem; font-weight: 650; }
        .site-meta { color: var(--muted); font-size: .74rem; }
        .site-value { font-size: .90rem; font-weight: 700; }
        .site-note { color: var(--muted); font-size: .78rem; text-align: right; }
        .alert-shell { overflow: hidden; }
        .alert-header, .alert-row { display: grid; grid-template-columns: 1.20fr 1.42fr .95fr 1.08fr 1.08fr .90fr .90fr 1.55fr; gap: .65rem; align-items: center; padding: .65rem .8rem; }
        .alert-header { background: #F8F8FA; color: var(--muted); font-size: .67rem; font-weight: 750; letter-spacing: .04em; text-transform: uppercase; }
        .alert-row { border-top: 1px solid var(--line); font-size: .79rem; }
        .alert-row .strong { font-weight: 700; }
        .alert-row .right { text-align: right; }
        .alert-row .grey { color: #8E8E93; }
        .alert-row .small { color: var(--muted); font-size: .74rem; }
        .empty-state { padding: 2.2rem; text-align: center; color: var(--muted); }
        .source-note { color: var(--muted); font-size: .75rem; margin-top: .8rem; }
        div[data-baseweb="tab-list"] { gap: .35rem; }
        button[data-baseweb="tab"] { border-radius: 999px; padding: .40rem .95rem; }
        button[data-baseweb="tab"][aria-selected="true"] { color: var(--blue); background: #E5F1FF; }
        @media (max-width: 900px) {
            .format-header { grid-template-columns: 1.3fr 1fr 1fr; }
            .format-header > div:last-child { grid-column: 1 / -1; }
            .alert-shell { overflow-x: auto; }
            .alert-header, .alert-row { min-width: 950px; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Import et préparation
# ---------------------------------------------------------------------------


def _parse_number(value: Any) -> float:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return np.nan
    if isinstance(value, (int, float, np.number)):
        return float(value)
    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not text or text.lower() in {"nan", "none", "null", "—", "-"}:
        return np.nan
    has_percent = "%" in text
    text = (
        text.replace("%", "")
        .replace("FCFA", "")
        .replace("€", "")
        .replace("$", "")
    )
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        number = float(text)
    except ValueError:
        return np.nan
    return number / 100 if has_percent else number


def _numericize(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    frame = frame.copy()
    for column in columns:
        if column in frame.columns:
            frame[column] = frame[column].map(_parse_number)
    return frame


def _ensure_column(frame: pd.DataFrame, canonical: str, aliases: list[str]) -> None:
    if canonical in frame.columns:
        return
    for alias in aliases:
        if alias in frame.columns:
            frame.rename(columns={alias: canonical}, inplace=True)
            return
    frame[canonical] = np.nan


def _derive_metrics(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    frame["delta_n1"] = frame["CA"] - frame["CA N-1"]
    frame["vs_n1"] = np.where(
        frame["CA N-1"].abs() > 0,
        frame["delta_n1"] / frame["CA N-1"],
        np.nan,
    )
    frame["delta_budget"] = frame["CA"] - frame["Budget"]
    frame["vs_budget"] = np.where(
        frame["Budget"].notna() & frame["Budget"].abs().gt(0),
        frame["delta_budget"] / frame["Budget"],
        np.nan,
    )
    frame["margin_rate_n1_calc"] = np.where(
        frame["CA N-1"].abs() > 0,
        frame["Marge N-1"] / frame["CA N-1"],
        np.nan,
    )
    frame["margin_rate_calc"] = np.where(
        frame["CA"].abs() > 0,
        frame["Marge"] / frame["CA"],
        np.nan,
    )
    # Alias utilisés par les vues ; les taux sont recalculés à partir des
    # montants afin de rester cohérents aux niveaux agrégés.
    frame["margin_rate_n1"] = frame["margin_rate_n1_calc"]
    frame["margin_rate"] = frame["margin_rate_calc"]
    frame["margin_delta_pts"] = (
        frame["margin_rate_calc"] - frame["margin_rate_n1_calc"]
    ) * 100
    frame["debit_vs_n1"] = frame["Vs N-1 (%) [débit]"]
    frame["panier_vs_n1"] = frame["Panier N Vs N-1"]
    return frame


def _severity(ca_vs_n1: Any, margin_delta_pts: Any) -> str:
    ca = float(ca_vs_n1) if pd.notna(ca_vs_n1) else 0.0
    margin = float(margin_delta_pts) if pd.notna(margin_delta_pts) else 0.0
    if ca < -0.05 or margin < -1.0:
        return "Critique"
    if ca < -0.02 or margin < -0.5:
        return "Attention"
    return "OK"


def _add_status(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    frame["status"] = [
        _severity(ca, margin)
        for ca, margin in zip(frame["vs_n1"], frame["margin_delta_pts"])
    ]
    return frame


def _format_from_site(site: Any) -> str:
    site_text = "" if pd.isna(site) else str(site)
    for format_name, names in SITE_FORMATS.items():
        if any(name.lower() in site_text.lower() for name in names):
            return format_name
    if "hyper" in site_text.lower():
        return "Hyper"
    if "market" in site_text.lower():
        return "Market"
    if "supeco" in site_text.lower():
        return "Supeco"
    return "Autre"


@st.cache_data(show_spinner=False)
def load_export(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    frame = pd.read_excel(BytesIO(file_bytes), sheet_name="Export")
    frame.columns = [str(column).replace("\u00a0", " ").strip() for column in frame.columns]

    # Le connecteur Excel peut suffixer la seconde colonne homonyme.
    _ensure_column(frame, "Vs N-1 (%) [débit]", ["Vs N-1 (%).1", "Vs N-1 (%) [débit]"])
    _ensure_column(frame, "Panier N Vs N-1", ["Panier N Vs N-1"])

    first_column = frame["Département"].astype("string") if "Département" in frame else pd.Series(dtype="string")
    footer = first_column.str.contains("Filtres appliqués", case=False, na=False)
    frame = frame.loc[~footer].copy()
    frame = frame.dropna(how="all")

    network_mask = frame["Département"].astype("string").str.strip().eq("Total")
    network = frame.loc[network_mask].tail(1).copy()
    frame = frame.loc[~network_mask].copy()

    # Les groupes du fichier source nécessitent un remplissage vers le bas.
    frame["Département"] = frame["Département"].ffill()
    frame["Rayon"] = frame["Rayon"].ffill()
    frame["Département"] = frame["Département"].astype(str).str.strip()
    frame["Rayon"] = frame["Rayon"].astype(str).str.strip()
    frame["Site"] = frame["Site"].where(frame["Site"].notna(), np.nan)

    numeric_columns = [
        "CA N-1", "Budget", "CA", "Poids", "Vs N-1 (%)", "Vs Bgt (%)",
        "Marge N-1", "Marge", "Taux de Marge N-1", "Taux de Marge",
        "Taux de Marge N Vs N-1", "Débit N-1", "Débit", "Vs N-1 (%) [débit]",
        "Panier N-1", "Panier", "Panier N Vs N-1", "Panier Qté N-1",
        "Panier Qté", "Panier Qté N Vs N-1", "Volume N-1", "Volume",
        "Volume N Vs N-1",
    ]
    frame = _numericize(frame, numeric_columns)
    network = _numericize(network, numeric_columns)

    departments = frame[frame["Département"].isin(REAL_DEPARTMENTS)].copy()
    summaries = departments[departments["Rayon"].eq("Total")].copy()
    pgc = departments[departments["Département"].eq(PGC)].copy()
    pgc_detail = pgc[
        ~pgc["Rayon"].eq("Total")
        & pgc["Site"].notna()
        & ~pgc["Site"].astype(str).str.strip().eq("Total")
    ].copy()
    pgc_detail["Format"] = pgc_detail["Site"].map(_format_from_site)
    pgc_detail = _derive_metrics(pgc_detail)
    pgc_detail = _add_status(pgc_detail)

    return {
        "detail": pgc_detail,
        "summaries": _derive_metrics(summaries),
        "network": _derive_metrics(network),
    }


# ---------------------------------------------------------------------------
# Présentation et agrégations
# ---------------------------------------------------------------------------


def _safe(value: Any) -> str:
    return html.escape(str(value))


def fmt_m(value: Any, signed: bool = False) -> str:
    if pd.isna(value):
        return "—"
    number = float(value) / 1_000_000
    rounded = int(round(abs(number)))
    sign = ""
    if signed:
        sign = "−" if number < 0 else "+" if number > 0 else ""
    elif number < 0:
        sign = "−"
    return f"{sign}{rounded:,}".replace(",", " ") + " M"


def fmt_pct(value: Any, decimals: int = 1, signed: bool = True) -> str:
    if pd.isna(value):
        return "—"
    number = float(value) * 100
    sign = ""
    if signed:
        sign = "−" if number < 0 else "+" if number > 0 else ""
    elif number < 0:
        sign = "−"
    return f"{sign}{abs(number):.{decimals}f}%"


def fmt_points(value: Any) -> str:
    if pd.isna(value):
        return "—"
    number = float(value)
    sign = "−" if number < 0 else "+" if number > 0 else ""
    return f"{sign}{abs(number):.1f} pt"


def trend_class(value: Any, invert: bool = False) -> str:
    if pd.isna(value) or float(value) == 0:
        return "neutral"
    positive = float(value) > 0
    if invert:
        positive = not positive
    return "positive" if positive else "negative"


def badge(status: str) -> str:
    css = {"Critique": "badge-critique", "Attention": "badge-attention", "OK": "badge-ok"}.get(status, "badge-ok")
    return f'<span class="badge {css}">{_safe(status)}</span>'


def aggregate(rows: pd.DataFrame, label: str = "") -> dict[str, Any]:
    if rows.empty:
        return {"label": label, "CA": np.nan, "CA N-1": np.nan, "Budget": np.nan, "Marge": np.nan, "Marge N-1": np.nan, "vs_n1": np.nan, "vs_budget": np.nan, "delta_n1": np.nan, "delta_budget": np.nan, "margin_delta_pts": np.nan, "status": "OK"}

    def total(column: str, min_count: int = 1) -> float:
        return rows[column].sum(min_count=min_count) if column in rows else np.nan

    result: dict[str, Any] = {
        "label": label,
        "CA": total("CA"),
        "CA N-1": total("CA N-1"),
        "Budget": total("Budget"),
        "Marge": total("Marge"),
        "Marge N-1": total("Marge N-1"),
    }
    result["delta_n1"] = result["CA"] - result["CA N-1"]
    result["vs_n1"] = result["delta_n1"] / result["CA N-1"] if result["CA N-1"] else np.nan
    result["delta_budget"] = result["CA"] - result["Budget"] if pd.notna(result["Budget"]) else np.nan
    result["vs_budget"] = result["delta_budget"] / result["Budget"] if pd.notna(result["Budget"]) and result["Budget"] else np.nan
    rate_n1 = result["Marge N-1"] / result["CA N-1"] if result["CA N-1"] else np.nan
    rate = result["Marge"] / result["CA"] if result["CA"] else np.nan
    result["margin_rate_n1"] = rate_n1
    result["margin_rate"] = rate
    result["margin_delta_pts"] = (rate - rate_n1) * 100
    result["status"] = _severity(result["vs_n1"], result["margin_delta_pts"])
    return result


def render_overview(data: dict[str, pd.DataFrame], file_name: str) -> None:
    summary_rows = data["summaries"]
    pgc_rows = summary_rows[summary_rows["Département"].eq(PGC)]
    pgc = pgc_rows.iloc[0].to_dict() if not pgc_rows.empty else aggregate(data["detail"], PGC)
    network_rows = data["network"]
    network = network_rows.iloc[0].to_dict() if not network_rows.empty else aggregate(summary_rows, "Total magasin")

    st.markdown("## Vue d’ensemble")
    st.caption(f"Pilotage hebdomadaire · fichier importé : {file_name}")

    ca_class = trend_class(pgc.get("vs_n1"))
    margin_class = trend_class(pgc.get("margin_delta_pts"))
    st.markdown(
        f"""
        <div class="dashboard-card" style="border-top: 4px solid #007AFF; min-height: 0;">
          <div class="eyebrow">Département PGC</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:2rem;margin-top:.5rem;">
            <div><div class="label">CA réalisé</div><div class="value {ca_class}">{fmt_m(pgc.get('CA'))}</div>
                 <div class="subvalue {ca_class}">{fmt_pct(pgc.get('vs_n1'))} vs N-1 · écart {fmt_m(pgc.get('delta_n1'), signed=True)}</div></div>
            <div><div class="label">Taux de marge</div><div class="value {margin_class}">{fmt_pct(pgc.get('margin_rate'), signed=False)}</div>
                 <div class="subvalue {margin_class}">{fmt_points(pgc.get('margin_delta_pts'))} vs N-1</div></div>
          </div>
          <div class="budget-line">Budget de référence : {fmt_m(pgc.get('Budget'))} · Vs budget : {fmt_pct(pgc.get('vs_budget'))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-kicker">Performance des départements</div>', unsafe_allow_html=True)
    rows_html = []
    for department in REAL_DEPARTMENTS:
        match = summary_rows[summary_rows["Département"].eq(department)]
        if match.empty:
            continue
        row = match.iloc[0]
        is_pgc = department == PGC
        classes = "pgc" if is_pgc else ""
        rows_html.append(
            f'<div class="dept-row {classes}"><div class="dept-name">{_safe(department)}</div>'
            f'<div class="dept-value">{fmt_m(row["CA"])}</div>'
            f'<div class="dept-var {trend_class(row["vs_n1"])}">{fmt_pct(row["vs_n1"])}</div></div>'
        )
    rows_html.append(
        f'<div class="dept-row total"><div class="dept-name">Total magasin</div>'
        f'<div class="dept-value">{fmt_m(network.get("CA"))}</div>'
        f'<div class="dept-var {trend_class(network.get("vs_n1"))}">{fmt_pct(network.get("vs_n1"))}</div></div>'
    )
    st.markdown(
        '<div class="dept-row" style="background:transparent;border:0;color:#6E6E73;font-size:.68rem;text-transform:uppercase;letter-spacing:.05em;">'
        '<div>Département</div><div style="text-align:right;">CA</div><div style="text-align:right;">Vs N-1</div></div>'
        + "".join(rows_html),
        unsafe_allow_html=True,
    )
    st.markdown('<div class="source-note">Les alertes sont pilotées uniquement par le CA vs N-1 et l’évolution du taux de marge en points. Le budget reste une référence secondaire.</div>', unsafe_allow_html=True)


def render_format(data: dict[str, pd.DataFrame]) -> None:
    detail = data["detail"]
    st.markdown("## Format")
    st.caption("Contribution à la sous-performance calculée séparément à chaque niveau d’agrégation.")

    format_aggregates = []
    for format_name in FORMAT_ORDER:
        rows = detail[detail["Format"].eq(format_name)]
        agg = aggregate(rows, format_name)
        agg["Format"] = format_name
        format_aggregates.append(agg)
    format_df = pd.DataFrame(format_aggregates)
    negative_pool = format_df.loc[format_df["delta_n1"] < 0, "delta_n1"].abs().sum()
    format_df["weight"] = np.where(
        (format_df["delta_n1"] < 0) & (negative_pool > 0),
        format_df["delta_n1"].abs() / negative_pool,
        0,
    )

    for format_name in FORMAT_ORDER:
        format_row = format_df[format_df["Format"].eq(format_name)].iloc[0]
        individual_sites = []
        remaining_weight = 0.0
        site_rows = []
        format_detail = detail[detail["Format"].eq(format_name)]
        for site, rows in format_detail.groupby("Site", sort=False):
            site_agg = aggregate(rows, str(site))
            site_agg["Site"] = site
            site_rows.append(site_agg)
        site_df = pd.DataFrame(site_rows)
        site_pool = site_df.loc[site_df["delta_n1"] < 0, "delta_n1"].abs().sum() if not site_df.empty else 0
        if not site_df.empty:
            site_df["weight"] = np.where(
                (site_df["delta_n1"] < 0) & (site_pool > 0),
                site_df["delta_n1"].abs() / site_pool,
                0,
            )
            site_df = site_df.sort_values(["weight", "delta_n1"], ascending=[False, True])
            individual_sites = site_df[site_df["weight"] > SITE_WEIGHT_THRESHOLD].to_dict("records")
            remaining_weight = float(site_df.loc[~site_df["Site"].isin([x["Site"] for x in individual_sites]), "weight"].sum())

        budget_text = fmt_pct(format_row.get("vs_budget"))
        st.markdown(
            f"""
            <div class="format-header">
              <div><div class="eyebrow">Format</div><div class="format-title">{_safe(format_name)}</div></div>
              <div><div class="format-metric-label">Vs N-1</div><div class="format-metric {trend_class(format_row.get('vs_n1'))}">{fmt_pct(format_row.get('vs_n1'))}</div></div>
              <div><div class="format-metric-label">Vs budget</div><div class="format-metric" style="color:#8E8E93;">{budget_text}</div></div>
              <div style="text-align:right;">{badge(format_row.get('status', 'OK'))}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        for site_row in individual_sites:
            st.markdown(
                f"""
                <div class="site-row">
                  <div><div class="site-name">{_safe(site_row['Site'])}</div><div class="site-meta">Site · {_safe(format_name)}</div></div>
                  <div class="site-value {trend_class(site_row.get('vs_n1'))}">{fmt_pct(site_row.get('vs_n1'))}</div>
                  <div class="site-value {trend_class(site_row.get('delta_n1'))}">{fmt_m(site_row.get('delta_n1'), signed=True)}</div>
                  <div class="site-note">{float(site_row.get('weight', 0)):.0%} du recul réseau</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        remaining_count = len(site_df) - len(individual_sites)
        if remaining_count:
            st.markdown(
                f'<div class="site-row"><div class="site-name">+ {remaining_count} autres sites</div><div></div><div></div><div class="site-note">poids cumulé {remaining_weight:.0%} du recul</div></div>',
                unsafe_allow_html=True,
            )


def render_alerts(data: dict[str, pd.DataFrame], signature: str) -> None:
    detail = data["detail"].copy()
    alerts = detail[
        detail["status"].isin(["Critique", "Attention"])
        & detail["delta_n1"].abs().ge(MATERIALITY_FLOOR)
    ].copy()
    alerts = alerts.sort_values(["delta_n1", "Marge"], ascending=[True, True])

    st.markdown("## Alertes à traiter")
    st.caption(f"{len(alerts)} ligne(s) Rayon × Site · filtre de matérialité : écart CA absolu ≥ {fmt_m(MATERIALITY_FLOOR)}")
    if alerts.empty:
        st.markdown('<div class="empty-state">Aucune alerte Rayon × Site ne dépasse le seuil de matérialité.</div>', unsafe_allow_html=True)
        return

    st.markdown(
        '<div class="alert-shell"><div class="alert-header"><div>Site</div><div>Rayon</div><div>Statut</div><div>Écart vs budget</div><div>Écart vs N-1</div><div>Débit</div><div>Panier</div><div>Cause</div></div></div>',
        unsafe_allow_html=True,
    )
    for index, row in alerts.reset_index(drop=True).iterrows():
        # La signature du fichier est incluse dans la clé : les causes repartent vierges au nouvel import.
        key = f"pgc_cause_{signature}_{index}"
        st.markdown('<div class="alert-shell">', unsafe_allow_html=True)
        columns = st.columns([1.20, 1.42, .95, 1.08, 1.08, .90, .90, 1.55], gap="small")
        with columns[0]:
            st.markdown(f'<div class="alert-row" style="display:block;border:0;padding:.65rem .2rem;"><div class="strong">{_safe(row["Site"])}</div><div class="small">{_safe(row["Format"])}</div></div>', unsafe_allow_html=True)
        with columns[1]:
            st.markdown(f'<div class="alert-row" style="display:block;border:0;padding:.65rem .2rem;"><div class="strong">{_safe(row["Rayon"])}</div><div class="small">écart CA {fmt_m(row["delta_n1"], signed=True)}</div></div>', unsafe_allow_html=True)
        with columns[2]:
            st.markdown(badge(row["status"]), unsafe_allow_html=True)
        with columns[3]:
            st.markdown(f'<div class="grey">{fmt_pct(row["vs_budget"])}</div>', unsafe_allow_html=True)
        with columns[4]:
            st.markdown(f'<div class="{trend_class(row["vs_n1"])}"><b>{fmt_pct(row["vs_n1"])}</b></div>', unsafe_allow_html=True)
        with columns[5]:
            st.markdown(f'<div class="{trend_class(row["debit_vs_n1"])}">{fmt_pct(row["debit_vs_n1"])}</div>', unsafe_allow_html=True)
        with columns[6]:
            st.markdown(f'<div class="{trend_class(row["panier_vs_n1"])}">{fmt_pct(row["panier_vs_n1"])}</div>', unsafe_allow_html=True)
        with columns[7]:
            st.selectbox(
                "Cause",
                CAUSES,
                key=key,
                label_visibility="collapsed",
            )
        st.markdown('</div>', unsafe_allow_html=True)


def main() -> None:
    inject_css()
    st.title("PGC Radar")
    st.markdown("La vigie hebdomadaire du chiffre d’affaires, de la marge et des alertes PGC.")

    with st.sidebar:
        st.markdown("### 📡 PGC Radar")
        st.caption("Importez le fichier hebdomadaire contenant l’onglet Export.")
        uploaded = st.file_uploader("Fichier Excel", type=["xlsx", "xls"], key="pgc_uploader")
        st.divider()
        st.caption("Alertes : CA vs N-1 et marge en points. Le budget est affiché comme référence secondaire.")

    if uploaded is None:
        st.markdown('<div class="empty-state"><div style="font-size:2.2rem;">📡</div><h3>Importez votre export hebdomadaire</h3><div>Les trois écrans Direction, Format et Alertes à traiter apparaîtront ici.</div></div>', unsafe_allow_html=True)
        st.stop()

    file_bytes = uploaded.getvalue()
    signature = hashlib.sha256(file_bytes).hexdigest()[:16]
    previous_signature = st.session_state.get("pgc_import_signature")
    if previous_signature != signature:
        for session_key in list(st.session_state):
            if str(session_key).startswith("pgc_cause_"):
                del st.session_state[session_key]
        st.session_state["pgc_import_signature"] = signature

    try:
        data = load_export(file_bytes)
    except Exception as exc:  # Affichage utilisateur propre en cas d’export mal formé.
        st.error(f"Impossible de lire l’onglet Export : {exc}")
        st.stop()

    tabs = st.tabs(["Vue d’ensemble", "Format", "Alertes à traiter"])
    with tabs[0]:
        render_overview(data, uploaded.name)
    with tabs[1]:
        render_format(data)
    with tabs[2]:
        render_alerts(data, signature)


if __name__ == "__main__":
    main()
