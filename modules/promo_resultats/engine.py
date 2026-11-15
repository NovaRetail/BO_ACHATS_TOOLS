"""SmartBuyer Hub · Résultats Promo — moteur de calcul et exports (V2, challenge des prévisions).

Entrées : extraction ventes PBI (hiérarchique, période = durée de la promo, montants HT)
          + liste promo (Code, Prévision Qté, Prix promo TTC, PA HT, TVA).
          CA prévu HT = Qté × Prix promo ÷ (1 + TVA) ; Marge prévue = Qté × (Prix promo HT − PA).
Sorties : fiabilité des prévisions par article, résultat vs N-1 (CA, marge en valeur et taux),
          contrôles, totaux, phrase de synthèse, classeur Excel (valeurs uniquement).
"""
from __future__ import annotations

import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import get_column_letter

NAVY = "0A2540"

# Fiabilité de la prévision (kind theme : ok / att / crit / info)
RELIABILITY = {
    "Fiable": {"kind": "ok", "bg": "E3F6E8"},
    "Surestimée": {"kind": "crit", "bg": "FFE3E1"},
    "Sous-estimée": {"kind": "att", "bg": "FFF0DB"},
    "Pas de vente": {"kind": "crit", "bg": "FFE3E1"},
    "Sans prévision": {"kind": "info", "bg": "EAEAEF"},
}
RELIABILITY_ORDER = list(RELIABILITY.keys())

# Lecture vs N-1 (sur le CA)
VS_N1 = {
    "Mieux que N-1": {"kind": "ok", "bg": "E3F6E8"},
    "Stable": {"kind": "info", "bg": "EAEAEF"},
    "Moins bien que N-1": {"kind": "crit", "bg": "FFE3E1"},
    "Pas de N-1": {"kind": "info", "bg": "EAEAEF"},
}
VS_N1_ORDER = list(VS_N1.keys())

# Sites dont le N-1 n'est pas comparable (conversions Supeco -> Market en 2026)
DEFAULT_CONVERTED_SITES = ["Market Cité verte", "Market Aboboté", "Market Toit rouge"]

SALES_REQUIRED = ["Article", "Site nom long", "CA", "Marge", "Qté Vente"]
N1_COLS = ["CA N-1", "Marge N-1", "Qté Vente N-1", "Marge Hors Promo N-1", "Marge Promo N-1", "CA Promo"]

PROMO_ALIASES = {
    "code": ["code", "code article", "article", "ref", "référence", "reference", "sku"],
    "label": ["libelle", "libellé", "designation", "désignation", "label", "nom"],
    "prev_qty": ["prevision qte", "prévision qté", "prévision qte", "prevision", "prévision", "prev qte", "qte prevue",
                 "qté prévue", "prevision qty", "qty", "qte", "qté", "quantite", "quantité"],
    "price": ["prix promo", "prix", "prix de vente promo", "pv promo", "prix vente", "pvc promo"],
    "prev_ca": ["prevision ca", "prévision ca", "ca prevu", "ca prévu"],
    "pa": ["pa", "prix d'achat", "prix achat", "pa ht", "prix d'achat ht", "pa net", "pa net ht", "cout", "coût"],
    "vat": ["tva", "taux tva", "tva %", "taux de tva", "vat"],
}


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def _nan(x) -> bool:
    return x is None or (isinstance(x, float) and np.isnan(x))


def fmt_int(x):
    return "—" if _nan(x) else f"{int(round(x)):,}".replace(",", " ")


def fmt_fcfa(x):
    if _nan(x):
        return "—"
    a = abs(x)
    if a >= 1e9:
        return f"{x / 1e9:.2f} Md".replace(".", ",")
    if a >= 1e6:
        return f"{x / 1e6:.2f} M".replace(".", ",")
    if a >= 1e3:
        return f"{x / 1e3:.0f} k"
    return fmt_int(x)


def fmt_pct(x, digits=1, sign=False):
    if _nan(x):
        return "—"
    s = f"{x * 100:+.{digits}f}" if sign else f"{x * 100:.{digits}f}"
    return s.replace(".", ",") + " %"


def fmt_pts(x):
    return "—" if _nan(x) else f"{x * 100:+.1f} pt".replace(".", ",")


def normalize_code(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return re.sub(r"\s+", "", s) or None


def norm_col(c):
    return re.sub(r"\s+", " ", str(c).strip().lower())


def safe_div(a, b):
    if b is None or (isinstance(b, float) and np.isnan(b)) or b == 0:
        return np.nan
    return a / b


# ----------------------------------------------------------------------------
# Parsing : extraction ventes
# ----------------------------------------------------------------------------
def read_any(file):
    name = file.name.lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(file)
    raw = file.getvalue()
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(raw), sep=None, engine="python", encoding=enc)
        except Exception:  # noqa: BLE001
            continue
    raise ValueError("Format de fichier non lu (xlsx ou csv attendu).")


def parse_filters(df):
    """Lit le bloc 'Filtres appliqués' présent en pied d'export PBI."""
    info = {"period_from": None, "period_to": None, "supplier": None, "sites": []}
    col = df.columns[0]
    txt = next((v for v in df[col].dropna().astype(str) if "Filtres appliqués" in v), None)
    if not txt:
        return info
    m = re.search(r"Fournisseur[^\n]*?est\s+(.+)", txt)
    if m:
        info["supplier"] = m.group(1).strip()
    m = re.search(r"Date est le ou après le (\d{2}/\d{2}/\d{4}) et est avant le (\d{2}/\d{2}/\d{4})", txt)
    if m:
        info["period_from"] = datetime.strptime(m.group(1), "%d/%m/%Y").date()
        info["period_to"] = datetime.strptime(m.group(2), "%d/%m/%Y").date() - pd.Timedelta(days=1)
    m = re.search(r"Site nom long est (.+)", txt)
    if m:
        info["sites"] = [s.strip() for s in re.split(r",\s*|\s+ou\s+", m.group(1).strip()) if s.strip()]
    return info


def parse_sales(file):
    df = read_any(file)
    missing = [c for c in SALES_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("colonnes manquantes -> " + ", ".join(missing))
    filters = parse_filters(df)
    d = df.copy()
    for c in ["Departement", "Rayon", "Famille", "Article"]:
        if c in d.columns:
            d[c] = d[c].ffill()
    d = d[d["Site nom long"].notna() & (d["Site nom long"].astype(str) != "Total")]
    d = d[d["Article"].notna() & (d["Article"].astype(str) != "Total")].copy()
    d["code"] = d["Article"].astype(str).str.extract(r"^\s*(\d+)")[0].map(normalize_code)
    d["label"] = d["Article"].astype(str).str.replace(r"^\s*\d+\s*-\s*", "", regex=True).str.strip()
    d["site"] = d["Site nom long"].astype(str).str.replace(r"^\s*\d+\s*-\s*", "", regex=True).str.strip()
    for c in SALES_REQUIRED[2:] + [c for c in N1_COLS if c in d.columns]:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0)
    # Marge N-1 absente de l'export PBI : reconstituée = Marge Hors Promo N-1 + Marge Promo N-1
    if "Marge N-1" not in d.columns:
        d["Marge N-1"] = (d["Marge Hors Promo N-1"] if "Marge Hors Promo N-1" in d.columns else 0.0) \
            + (d["Marge Promo N-1"] if "Marge Promo N-1" in d.columns else 0.0)
    for c in N1_COLS:
        if c not in d.columns:
            d[c] = 0.0
    keep = ["code", "label", "site", "CA", "Marge", "Qté Vente", "CA N-1", "Marge N-1", "Qté Vente N-1", "CA Promo"]
    return d[keep].reset_index(drop=True), filters


# ----------------------------------------------------------------------------
# Parsing : liste promo
# ----------------------------------------------------------------------------
def match_promo_columns(cols):
    normed = {norm_col(c): c for c in cols}
    found = {}
    for key, aliases in PROMO_ALIASES.items():
        for a in aliases:
            if a in normed and normed[a] not in found.values():
                found[key] = normed[a]
                break
    return found


def parse_rate(x):
    """TVA saisie en 18, 0,18, 18 % ou 0.18 -> 0.18."""
    if pd.isna(x):
        return np.nan
    v = pd.to_numeric(str(x).replace("%", "").replace(",", ".").strip(), errors="coerce")
    if pd.isna(v):
        return np.nan
    return v / 100 if v > 1 else v


def parse_promo(file, price_ttc=True):
    df = read_any(file).dropna(how="all")
    cols = match_promo_columns(df.columns)
    if "code" not in cols:
        raise ValueError("colonne « Code » introuvable.")
    if "prev_qty" not in cols:
        raise ValueError("colonne « Prévision Qté » introuvable.")
    if "price" not in cols and "prev_ca" not in cols:
        raise ValueError("colonne « Prix promo » introuvable (le CA prévu = Prévision Qté × Prix promo).")
    p = pd.DataFrame()
    p["code"] = df[cols["code"]].map(normalize_code)
    p["label_promo"] = df[cols["label"]].astype(str).str.strip() if "label" in cols else ""
    p["prev_qty"] = pd.to_numeric(df[cols["prev_qty"]], errors="coerce")
    p["price"] = pd.to_numeric(df[cols["price"]], errors="coerce") if "price" in cols else np.nan
    p["pa"] = pd.to_numeric(df[cols["pa"]], errors="coerce") if "pa" in cols else np.nan
    p["vat"] = df[cols["vat"]].map(parse_rate) if "vat" in cols else np.nan
    has_vat, has_pa = "vat" in cols, "pa" in cols
    if "price" in cols:
        divisor = (1 + p["vat"].fillna(0)) if (price_ttc and has_vat) else 1.0
        p["price_ht"] = p["price"] / divisor
        p["prev_ca"] = p["prev_qty"] * p["price_ht"]
        ca_mode = ("Qté × Prix promo ÷ (1 + TVA)" if (price_ttc and has_vat)
                   else "Qté × Prix promo (prix saisi HT)" if not price_ttc else "Qté × Prix promo (TVA absente : prix pris tel quel)")
    else:
        p["price_ht"] = np.nan
        p["prev_ca"] = pd.to_numeric(df[cols["prev_ca"]], errors="coerce")
        ca_mode = "CA prévu fourni tel quel"
    p["prev_marge"] = p["prev_qty"] * (p["price_ht"] - p["pa"]) if has_pa else np.nan
    p = p[p["code"].notna()].reset_index(drop=True)
    return p, {"columns": cols, "ca_mode": ca_mode, "has_pa": has_pa, "has_vat": has_vat, "price_ttc": price_ttc}


# ----------------------------------------------------------------------------
# Moteur de calcul
# ----------------------------------------------------------------------------
def compute(sales, promo, params):
    low, high, stable = params["low"], params["high"], params["stable_band"]
    converted = set(params["converted_sites"])
    codes = promo["code"].unique().tolist()
    s = sales[sales["code"].isin(codes)].copy()

    # Détail site (pour le N-1 non comparable et le drill-down export)
    site = s.copy()
    site["n1_comparable"] = ~site["site"].isin(converted)

    agg = s.groupby("code")[["CA", "Marge", "Qté Vente", "CA N-1", "Marge N-1", "Qté Vente N-1", "CA Promo"]].sum()
    agg["label"] = s.groupby("code")["label"].first()
    agg["n_sites"] = s.groupby("code").size()
    agg["n_sites_converted"] = s[s["site"].isin(converted)].groupby("code").size()

    prev = promo.groupby("code").agg(prev_qty=("prev_qty", "sum"), prev_ca=("prev_ca", "sum"), prev_marge=("prev_marge", "sum"),
                                     price=("price", "mean"), pa=("pa", "mean"), vat=("vat", "mean"),
                                     label_promo=("label_promo", "first"), n_rows=("code", "size"))
    for c in ["prev_qty", "prev_ca", "prev_marge", "price", "pa", "vat"]:
        prev[c] = prev[c].where(promo.groupby("code")[c].count() > 0, np.nan)
    art = prev.join(agg, how="left")
    art["found"] = art["CA"].notna()
    for c in ["CA", "Marge", "Qté Vente", "CA N-1", "Marge N-1", "Qté Vente N-1", "CA Promo", "n_sites", "n_sites_converted"]:
        art[c] = art[c].fillna(0)
    art["label"] = art["label"].fillna(art["label_promo"])

    # Fiabilité de la prévision
    art["real_qty"] = art.apply(lambda r: safe_div(r["Qté Vente"], r["prev_qty"]), axis=1)
    art["real_ca"] = art.apply(lambda r: safe_div(r["CA"], r["prev_ca"]), axis=1)
    art["gap_qty"] = art["Qté Vente"] - art["prev_qty"]
    art["gap_ca"] = art["CA"] - art["prev_ca"]
    art["ca_missed"] = (art["prev_ca"] - art["CA"]).clip(lower=0)  # CA promis non réalisé
    art["prev_rate"] = art.apply(lambda r: safe_div(r["prev_marge"], r["prev_ca"]), axis=1)
    art["real_marge"] = art.apply(lambda r: safe_div(r["Marge"], r["prev_marge"]), axis=1)
    art["gap_marge"] = art["Marge"] - art["prev_marge"]
    art["marge_missed"] = (art["prev_marge"] - art["Marge"]).clip(lower=0)

    def reliability(r):
        if pd.isna(r["prev_ca"]) and pd.isna(r["prev_qty"]):
            return "Sans prévision"
        if not r["found"] or r["Qté Vente"] <= 0:
            return "Pas de vente"
        rl = r["real_ca"] if not pd.isna(r["real_ca"]) else r["real_qty"]
        if pd.isna(rl):
            return "Sans prévision"
        if rl < low:
            return "Surestimée"
        if rl > high:
            return "Sous-estimée"
        return "Fiable"

    art["reliability"] = art.apply(reliability, axis=1)

    # Résultat vs N-1 : CA, marge en valeur et taux
    art["rate"] = art.apply(lambda r: safe_div(r["Marge"], r["CA"]), axis=1)
    art["rate_n1"] = art.apply(lambda r: safe_div(r["Marge N-1"], r["CA N-1"]), axis=1)
    art["ca_evo"] = art.apply(lambda r: safe_div(r["CA"] - r["CA N-1"], r["CA N-1"]), axis=1)
    art["marge_evo"] = art.apply(lambda r: safe_div(r["Marge"] - r["Marge N-1"], r["Marge N-1"]), axis=1)
    art["marge_delta"] = art["Marge"] - art["Marge N-1"]
    art["rate_delta"] = art["rate"] - art["rate_n1"]
    art["n1_partial"] = art["n_sites_converted"] > 0

    def vs_n1(r):
        if r["CA N-1"] <= 0:
            return "Pas de N-1"
        if r["ca_evo"] > stable:
            return "Mieux que N-1"
        if r["ca_evo"] < -stable:
            return "Moins bien que N-1"
        return "Stable"

    art["vs_n1"] = art.apply(vs_n1, axis=1)
    art = art.reset_index().sort_values("gap_ca", ascending=True, na_position="last").reset_index(drop=True)

    controls = {
        "duplicates": promo[promo.duplicated("code", keep=False)]["code"].unique().tolist(),
        "not_found": art[~art["found"]]["code"].tolist(),
        "converted_present": sorted(set(sales["site"]) & converted),
        "no_price": promo[promo["prev_ca"].isna()]["code"].tolist(),
        "no_pa": promo[promo["pa"].isna()]["code"].tolist() if "pa" in promo else [],
        "neg_prev_marge": promo[promo["prev_marge"] < 0]["code"].tolist(),
    }

    tot = {
        "n": len(art),
        "prev_qty": art["prev_qty"].sum(skipna=True), "qty": art["Qté Vente"].sum(),
        "prev_ca": art["prev_ca"].sum(skipna=True), "ca": art["CA"].sum(),
        "ca_missed": art["ca_missed"].sum(), "marge": art["Marge"].sum(),
        "prev_marge": art["prev_marge"].sum(min_count=1), "marge_missed": art["marge_missed"].sum(),
        "ca_n1": art["CA N-1"].sum(), "marge_n1": art["Marge N-1"].sum(),
        "sites": int(sales["site"].nunique()),
        "rel_counts": art["reliability"].value_counts().reindex(RELIABILITY_ORDER, fill_value=0),
        "n1_counts": art["vs_n1"].value_counts().reindex(VS_N1_ORDER, fill_value=0),
    }
    tot["real_qty"] = safe_div(tot["qty"], tot["prev_qty"])
    tot["real_ca"] = safe_div(tot["ca"], tot["prev_ca"])
    tot["rate"] = safe_div(tot["marge"], tot["ca"])
    tot["prev_rate"] = safe_div(tot["prev_marge"], tot["prev_ca"])
    tot["real_marge"] = safe_div(tot["marge"], tot["prev_marge"])
    tot["rate_n1"] = safe_div(tot["marge_n1"], tot["ca_n1"])
    tot["ca_evo"] = safe_div(tot["ca"] - tot["ca_n1"], tot["ca_n1"])
    tot["marge_evo"] = safe_div(tot["marge"] - tot["marge_n1"], tot["marge_n1"])
    return art, site, controls, tot


def synthesis_sentence(art, tot):
    rc = tot["rel_counts"]
    s = (f"{tot['n']} articles : CA prévu {fmt_fcfa(tot['prev_ca'])}, réalisé {fmt_fcfa(tot['ca'])} "
         f"({fmt_pct(tot['real_ca'], 0)}). {int(rc['Surestimée'])} prévision(s) surestimée(s), "
         f"{int(rc['Fiable'])} fiable(s), {int(rc['Sous-estimée'])} sous-estimée(s). "
         f"CA promis non réalisé : {fmt_fcfa(tot['ca_missed'])} FCFA.")
    if not _nan(tot["prev_marge"]):
        s += (f" Marge prévue {fmt_fcfa(tot['prev_marge'])} ({fmt_pct(tot['prev_rate'])}), réalisée {fmt_fcfa(tot['marge'])} "
              f"({fmt_pct(tot['rate'])}) : {fmt_pct(tot['real_marge'], 0)} de l'engagement.")
    if tot["ca_n1"] > 0:
        s += (f" Vs N-1 : CA {fmt_pct(tot['ca_evo'], 0, True)}, marge {fmt_pct(tot['marge_evo'], 0, True)} "
              f"({fmt_pct(tot['rate'])} vs {fmt_pct(tot['rate_n1'])}).")
    return s


# ----------------------------------------------------------------------------
# Export Excel (valeurs figées, aucune formule)
# ----------------------------------------------------------------------------
def build_excel(art, site, controls, tot, filters, meta, params, sentence):
    wb = Workbook()
    hdr_fill = PatternFill("solid", fgColor=NAVY)
    hdr_font = Font(bold=True, color="FFFFFF")
    thin = Border()
    fills = {k: PatternFill("solid", fgColor=v["bg"]) for k, v in {**RELIABILITY, **VS_N1}.items()}

    def write_table(ws, df, start_row=1, pct_cols=(), int_cols=(), badge_cols=()):
        for j, c in enumerate(df.columns, 1):
            cell = ws.cell(row=start_row, column=j, value=c)
            cell.fill, cell.font = hdr_fill, hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for i, row in enumerate(df.itertuples(index=False), start_row + 1):
            for j, val in enumerate(row, 1):
                if isinstance(val, float) and np.isnan(val):
                    val = None
                cell = ws.cell(row=i, column=j, value=val)
                name = df.columns[j - 1]
                if name in pct_cols:
                    cell.number_format = "0.0%"
                elif name in int_cols:
                    cell.number_format = "#,##0"
                if name in badge_cols and val in fills:
                    cell.fill = fills[val]
                    cell.font = Font(bold=True)
        for j, c in enumerate(df.columns, 1):
            ws.column_dimensions[get_column_letter(j)].width = max(12, min(42, len(str(c)) + 4))
        ws.column_dimensions["A"].width = 36
        ws.freeze_panes = ws.cell(row=start_row + 1, column=2)

    per = (f"{filters['period_from']:%d/%m/%Y} au {filters['period_to']:%d/%m/%Y}" if filters["period_from"] else "non lue")

    # --- Synthèse ---
    ws = wb.active
    ws.title = "Synthèse"
    ws["A1"] = "Résultats Promo · challenge des prévisions"
    ws["A1"].font = Font(bold=True, size=16, color=NAVY)
    ws["A2"] = f"Période : {per} · Fournisseur : {filters['supplier'] or 'non lu'} · CA prévu = {meta['ca_mode']} · généré le {datetime.now():%d/%m/%Y %H:%M}"
    ws["A4"] = sentence
    ws["A4"].font = Font(bold=True)
    ws["A4"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A4:F4")
    ws.row_dimensions[4].height = 48
    kpis = [("Articles", tot["n"], "#,##0"),
            ("Quantité prévue", tot["prev_qty"], "#,##0"), ("Quantité réalisée", tot["qty"], "#,##0"),
            ("Réalisation quantité", tot["real_qty"], "0%"),
            ("CA prévu (FCFA)", tot["prev_ca"], "#,##0"), ("CA réalisé (FCFA)", tot["ca"], "#,##0"),
            ("Réalisation CA", tot["real_ca"], "0%"), ("CA promis non réalisé (FCFA)", tot["ca_missed"], "#,##0"),
            ("Marge prévue (FCFA)", tot["prev_marge"], "#,##0"), ("Taux de marge prévu", tot["prev_rate"], "0.0%"),
            ("Réalisation marge", tot["real_marge"], "0%"), ("Marge promise non réalisée (FCFA)", tot["marge_missed"], "#,##0"),
            ("CA N-1 (FCFA)", tot["ca_n1"], "#,##0"), ("Évolution CA vs N-1", tot["ca_evo"], "+0.0%;-0.0%"),
            ("Marge réalisée (FCFA)", tot["marge"], "#,##0"), ("Taux de marge", tot["rate"], "0.0%"),
            ("Marge N-1 (FCFA)", tot["marge_n1"], "#,##0"), ("Taux de marge N-1", tot["rate_n1"], "0.0%"),
            ("Évolution marge vs N-1", tot["marge_evo"], "+0.0%;-0.0%")]
    r = 6
    for lab, nf in [("Indicateur", None), ("Valeur", None)]:
        pass
    ws.cell(row=r, column=1, value="Indicateur").fill = hdr_fill
    ws.cell(row=r, column=1).font = hdr_font
    ws.cell(row=r, column=2, value="Valeur").fill = hdr_fill
    ws.cell(row=r, column=2).font = hdr_font
    for lab, val, nf in kpis:
        r += 1
        ws.cell(row=r, column=1, value=lab)
        c = ws.cell(row=r, column=2, value=None if _nan(val) else val)
        c.number_format = nf
    r += 2
    ws.cell(row=r, column=1, value="Fiabilité des prévisions").fill = hdr_fill
    ws.cell(row=r, column=1).font = hdr_font
    ws.cell(row=r, column=2, value="Articles").fill = hdr_fill
    ws.cell(row=r, column=2).font = hdr_font
    for k in RELIABILITY_ORDER:
        r += 1
        ws.cell(row=r, column=1, value=k).fill = fills[k]
        ws.cell(row=r, column=2, value=int(tot["rel_counts"][k]))
    r += 2
    ws.cell(row=r, column=1, value="Résultat vs N-1 (CA)").fill = hdr_fill
    ws.cell(row=r, column=1).font = hdr_font
    ws.cell(row=r, column=2, value="Articles").fill = hdr_fill
    ws.cell(row=r, column=2).font = hdr_font
    for k in VS_N1_ORDER:
        r += 1
        ws.cell(row=r, column=1, value=k).fill = fills[k]
        ws.cell(row=r, column=2, value=int(tot["n1_counts"][k]))
    r += 2
    ws.cell(row=r, column=1, value="Paramètres").font = Font(bold=True)
    for k, lab in [("low", "Surestimée en dessous de"), ("high", "Sous-estimée au-delà de"), ("stable_band", "Bande « Stable » vs N-1 (±)")]:
        r += 1
        ws.cell(row=r, column=1, value=lab)
        ws.cell(row=r, column=2, value=params[k]).number_format = "0%"
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 18

    # --- Fiabilité des prévisions ---
    ws = wb.create_sheet("Fiabilité prévisions")
    cols = {"label": "Article", "code": "Code", "prev_qty": "Prévu Qté", "Qté Vente": "Réalisé Qté", "real_qty": "Réalisation Qté",
            "gap_qty": "Écart Qté", "price": "Prix promo", "pa": "PA", "vat": "TVA", "prev_ca": "CA prévu HT", "CA": "CA réalisé",
            "real_ca": "Réalisation CA", "gap_ca": "Écart CA", "ca_missed": "CA promis non réalisé",
            "prev_marge": "Marge prévue", "prev_rate": "Taux prévu", "Marge": "Marge réalisée", "rate": "Taux réalisé",
            "real_marge": "Réalisation marge", "gap_marge": "Écart marge", "reliability": "Fiabilité"}
    if not meta.get("has_pa"):
        for k in ["pa", "prev_marge", "prev_rate", "real_marge", "gap_marge"]:
            cols.pop(k)
    if not meta.get("has_vat"):
        cols.pop("vat")
    df = art[list(cols)].rename(columns=cols)
    write_table(ws, df, pct_cols=("Réalisation Qté", "Réalisation CA", "TVA", "Taux prévu", "Taux réalisé", "Réalisation marge"),
                int_cols=("Prévu Qté", "Réalisé Qté", "Écart Qté", "Prix promo", "PA", "CA prévu HT", "CA réalisé", "Écart CA",
                          "CA promis non réalisé", "Marge prévue", "Marge réalisée", "Écart marge"),
                badge_cols=("Fiabilité",))

    # --- Vs N-1 ---
    ws = wb.create_sheet("Vs N-1")
    cols = {"label": "Article", "code": "Code", "CA": "CA réalisé", "CA N-1": "CA N-1", "ca_evo": "Évo CA",
            "Marge": "Marge réalisée", "rate": "Taux de marge", "Marge N-1": "Marge N-1", "rate_n1": "Taux N-1",
            "marge_delta": "Écart marge", "marge_evo": "Évo marge", "rate_delta": "Écart taux (pts)",
            "Qté Vente": "Qté", "Qté Vente N-1": "Qté N-1", "n1_partial": "N-1 partiellement non comparable", "vs_n1": "Lecture"}
    df = art[list(cols)].rename(columns=cols)
    df["N-1 partiellement non comparable"] = df["N-1 partiellement non comparable"].map({True: "Oui", False: "Non"})
    write_table(ws, df, pct_cols=("Évo CA", "Taux de marge", "Taux N-1", "Évo marge", "Écart taux (pts)"),
                int_cols=("CA réalisé", "CA N-1", "Marge réalisée", "Marge N-1", "Écart marge", "Qté", "Qté N-1"), badge_cols=("Lecture",))

    # --- Détail sites ---
    ws = wb.create_sheet("Détail sites")
    cols = {"code": "Code", "label": "Article", "site": "Site", "Qté Vente": "Qté", "Qté Vente N-1": "Qté N-1", "CA": "CA",
            "CA N-1": "CA N-1", "Marge": "Marge", "Marge N-1": "Marge N-1", "n1_comparable": "N-1 comparable"}
    sdf = site[list(cols)].rename(columns=cols).sort_values(["Article", "CA"], ascending=[True, False])
    sdf["N-1 comparable"] = sdf["N-1 comparable"].map({True: "Oui", False: "Non (site converti)"})
    write_table(ws, sdf, int_cols=("Qté", "Qté N-1", "CA", "CA N-1", "Marge", "Marge N-1"))

    # --- Contrôles ---
    ws = wb.create_sheet("Contrôles")
    rows = [("Période", per), ("Fournisseur", filters["supplier"] or "non lu"), ("Sites", ", ".join(filters["sites"]) or "non lus"),
            ("CA prévu", meta["ca_mode"]), ("Codes trouvés", f"{int(art['found'].sum())} / {len(art)}"),
            ("Codes introuvables", ", ".join(controls["not_found"]) or "aucun"),
            ("Doublons", ", ".join(controls["duplicates"]) or "aucun"),
            ("Sans prix promo / CA prévu", ", ".join(controls["no_price"]) or "aucun"),
            ("Sans PA", ", ".join(controls["no_pa"]) or "aucun"),
            ("Marge prévue négative (prix promo HT < PA)", ", ".join(controls["neg_prev_marge"]) or "aucun"),
            ("Sites au N-1 non comparable", ", ".join(controls["converted_present"]) or "aucun")]
    write_table(ws, pd.DataFrame(rows, columns=["Contrôle", "Résultat"]))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
