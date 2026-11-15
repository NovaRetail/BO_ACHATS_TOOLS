"""SmartBuyer Hub · Résultats Promo — moteur de calcul et exports.

Entrées : extraction ventes PBI (hiérarchique) + liste promo (Code, Prévision Qté et/ou CA, Site*, Financement*).
Sorties : tables article / site, contrôles, totaux, phrase de synthèse, classeur Excel (valeurs uniquement).
"""
from __future__ import annotations

import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NAVY = "#0A2540"

# Verdicts : couleur de badge (kind theme : ok / att / crit / info) et action recommandée
VERDICTS = {
    "Prévision sous-estimée": {"kind": "info", "hex": "6A5ACD", "bg": "EFE9FF",
        "action": "Revoir la prévision à la hausse, sécuriser le stock."},
    "Succès": {"kind": "ok", "hex": "34C759", "bg": "E3F6E8",
        "action": "Reconduire la mécanique ; valider le stock de fin de promo."},
    "Conforme": {"kind": "ok", "hex": "007AFF", "bg": "E1EEFF",
        "action": "RAS : prévision fiable, à reconduire."},
    "Volume sans marge": {"kind": "att", "hex": "FF9500", "bg": "FFF0DB",
        "action": "Renégocier le financement fournisseur ou la profondeur de remise."},
    "Marge promo négative": {"kind": "crit", "hex": "FF3B30", "bg": "FFE3E1",
        "action": "Vente à perte en promo : vérifier le prix promo et le PMP."},
    "Sous-performance": {"kind": "crit", "hex": "FF3B30", "bg": "FFE3E1",
        "action": "Vérifier disponibilité, visibilité en magasin et prix ; ajuster la prévision."},
    "Promo non activée": {"kind": "att", "hex": "FF9500", "bg": "FFF0DB",
        "action": "Vérifier le paramétrage promo en caisse et en BDD."},
    "Pas de vente": {"kind": "info", "hex": "8E8E93", "bg": "EAEAEF",
        "action": "Vérifier le référencement et le stock."},
}
VERDICT_ORDER = list(VERDICTS.keys())

# Sites dont le N-1 n'est pas comparable (conversions Supeco -> Market en 2026)
DEFAULT_CONVERTED_SITES = ["Market Cité verte", "Market Aboboté", "Market Toit rouge"]

SALES_REQUIRED = ["Article", "Site nom long", "CA", "Marge", "CA Promo", "Marge Promo",
                  "CA Hors Promo", "Marge Hors Promo", "Qté Vente"]
SALES_OPTIONAL = ["CA N-1", "Marge N-1", "CA Promo N-1", "Qté Vente N-1", "CA Hors Promo N-1",
                  "Marge Hors Promo N-1", "Departement", "Rayon", "Famille"]

PROMO_ALIASES = {
    "code": ["code", "code article", "article", "ref", "référence", "reference", "sku"],
    "label": ["libelle", "libellé", "designation", "désignation", "label", "nom"],
    "prev_qty": ["prevision", "prévision", "prevision qte", "prévision qté", "prev qte", "qte prevue",
                 "qté prévue", "prevision qty", "qty", "qte", "qté", "quantite", "quantité"],
    "prev_ca": ["prevision ca", "prévision ca", "ca prevu", "ca prévu", "ca", "chiffre d'affaires"],
    "site": ["site", "site nom long", "magasin"],
    "funding": ["financement", "financement fournisseur", "budget promo", "participation", "rfa promo"],
}


def fmt_int(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    return f"{int(round(x)):,}".replace(",", " ")


def fmt_fcfa(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    a = abs(x)
    if a >= 1e9:
        return f"{x/1e9:.2f} Md".replace(".", ",")
    if a >= 1e6:
        return f"{x/1e6:.2f} M".replace(".", ",")
    if a >= 1e3:
        return f"{x/1e3:.0f} k"
    return fmt_int(x)


def fmt_pct(x, digits=1, sign=False):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    s = f"{x*100:+.{digits}f}" if sign else f"{x*100:.{digits}f}"
    return s.replace(".", ",") + " %"


def fmt_pts(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    return f"{x*100:+.1f} pt".replace(".", ",")


def normalize_code(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = re.sub(r"\s+", "", s)
    return s or None


def norm_col(c):
    return re.sub(r"\s+", " ", str(c).strip().lower())


def safe_div(a, b):
    return a / b if b not in (0, None) and not pd.isna(b) and b != 0 else np.nan


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
        except Exception:
            continue
    raise ValueError("Format de fichier non lu (xlsx ou csv attendu).")


def parse_filters(df):
    """Lit le bloc 'Filtres appliqués' présent en pied d'export PBI."""
    info = {"period_from": None, "period_to": None, "supplier": None, "sites": [], "raw": None}
    col = df.columns[0]
    txt = None
    for v in df[col].dropna().astype(str):
        if "Filtres appliqués" in v:
            txt = v
            break
    if not txt:
        return info
    info["raw"] = txt
    m = re.search(r"Fournisseur[^\n]*?est\s+(.+)", txt)
    if m:
        info["supplier"] = m.group(1).strip()
    m = re.search(r"Date est le ou après le (\d{2}/\d{2}/\d{4}) et est avant le (\d{2}/\d{2}/\d{4})", txt)
    if m:
        info["period_from"] = datetime.strptime(m.group(1), "%d/%m/%Y").date()
        d_to = datetime.strptime(m.group(2), "%d/%m/%Y").date()
        info["period_to"] = d_to - pd.Timedelta(days=1)
    m = re.search(r"Site nom long est (.+)", txt)
    if m:
        sites = re.split(r",\s*|\s+ou\s+", m.group(1).strip())
        info["sites"] = [s.strip() for s in sites if s.strip()]
    return info


def parse_sales(file):
    df = read_any(file)
    missing = [c for c in SALES_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("Extraction ventes : colonnes manquantes -> " + ", ".join(missing))
    filters = parse_filters(df)
    d = df.copy()
    for c in ["Departement", "Rayon", "Famille", "Article"]:
        if c in d.columns:
            d[c] = d[c].ffill()
    d = d[d["Site nom long"].notna() & (d["Site nom long"].astype(str) != "Total")].copy()
    d = d[d["Article"].notna() & (d["Article"].astype(str) != "Total")].copy()
    d["code"] = d["Article"].astype(str).str.extract(r"^\s*(\d+)")[0].map(normalize_code)
    d["label"] = d["Article"].astype(str).str.replace(r"^\s*\d+\s*-\s*", "", regex=True).str.strip()
    d["site"] = d["Site nom long"].astype(str).str.replace(r"^\s*\d+\s*-\s*", "", regex=True).str.strip()
    n1_cols = ["CA N-1", "Marge N-1", "CA Promo N-1", "Qté Vente N-1", "CA Hors Promo N-1", "Marge Hors Promo N-1"]
    for c in SALES_REQUIRED[2:] + [c for c in n1_cols if c in d.columns]:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0)
    # Marge N-1 absente de l'export PBI : reconstituée à partir de Marge HP N-1 + Marge Promo N-1
    if "Marge N-1" not in d.columns:
        if "Marge Hors Promo N-1" in d.columns and "Marge Promo N-1" in d.columns:
            d["Marge Promo N-1"] = pd.to_numeric(d["Marge Promo N-1"], errors="coerce").fillna(0.0)
            d["Marge N-1"] = d["Marge Hors Promo N-1"] + d["Marge Promo N-1"]
        else:
            d["Marge N-1"] = 0.0
    for c in n1_cols:
        if c not in d.columns:
            d[c] = 0.0
    num_cols = SALES_REQUIRED[2:] + n1_cols
    keep = ["code", "label", "site", "Departement", "Rayon", "Famille"] + num_cols
    keep = [c for c in keep if c in d.columns]
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
    # "ca" ne doit pas capter une colonne déjà prise par prev_qty
    if "prev_ca" in found and found["prev_ca"] == found.get("prev_qty"):
        del found["prev_ca"]
    return found


def parse_promo(file):
    df = read_any(file)
    df = df.dropna(how="all")
    cols = match_promo_columns(df.columns)
    if "code" not in cols:
        raise ValueError("Liste promo : colonne CODE introuvable (attendu : CODE ou Code article).")
    if "prev_qty" not in cols and "prev_ca" not in cols:
        raise ValueError("Liste promo : il faut au moins une colonne « Prévision » (quantité) ou « Prévision CA ».")
    p = pd.DataFrame()
    p["code"] = df[cols["code"]].map(normalize_code)
    p["label_promo"] = df[cols["label"]].astype(str).str.strip() if "label" in cols else ""
    p["prev_qty"] = pd.to_numeric(df[cols["prev_qty"]], errors="coerce") if "prev_qty" in cols else np.nan
    p["prev_ca"] = pd.to_numeric(df[cols["prev_ca"]], errors="coerce") if "prev_ca" in cols else np.nan
    p["site"] = df[cols["site"]].astype(str).str.replace(r"^\s*\d+\s*-\s*", "", regex=True).str.strip() if "site" in cols else None
    p["funding"] = pd.to_numeric(df[cols["funding"]], errors="coerce") if "funding" in cols else np.nan
    p = p[p["code"].notna()].reset_index(drop=True)
    meta = {
        "has_qty": "prev_qty" in cols and p["prev_qty"].notna().any(),
        "has_ca": "prev_ca" in cols and p["prev_ca"].notna().any(),
        "has_site": "site" in cols and p["site"].notna().any(),
        "has_funding": "funding" in cols and p["funding"].notna().any(),
        "columns": cols,
    }
    return p, meta


# ----------------------------------------------------------------------------
# Moteur de calcul
# ----------------------------------------------------------------------------
def compute(sales, promo, meta, params):
    low, high, over = params["low"], params["high"], params["over"]
    tol = params["margin_tol"]
    converted = set(params["converted_sites"])

    codes = promo["code"].unique().tolist()
    s = sales[sales["code"].isin(codes)].copy()

    # --- Détail site ---------------------------------------------------------
    site = s.copy()
    site["promo_active"] = site["CA Promo"] > 0
    site["rate_promo"] = site.apply(lambda r: safe_div(r["Marge Promo"], r["CA Promo"]), axis=1)
    site["rate_hp"] = site.apply(lambda r: safe_div(r["Marge Hors Promo"], r["CA Hors Promo"]), axis=1)
    site["qty_evo"] = site.apply(lambda r: safe_div(r["Qté Vente"] - r["Qté Vente N-1"], r["Qté Vente N-1"]), axis=1)
    site["n1_comparable"] = ~site["site"].isin(converted)

    # --- Agrégat article ----------------------------------------------------
    agg_cols = ["CA", "Marge", "CA Promo", "Marge Promo", "CA Hors Promo", "Marge Hors Promo", "Qté Vente",
                "CA N-1", "Marge N-1", "CA Promo N-1", "Qté Vente N-1"]
    art = s.groupby("code")[agg_cols].sum()
    art["label"] = s.groupby("code")["label"].first()
    art["n_sites_sold"] = s[s["Qté Vente"] > 0].groupby("code").size()
    art["n_sites_promo"] = s[s["CA Promo"] > 0].groupby("code").size()
    art["n_sites_total"] = s.groupby("code").size()
    # Hors promo sur les mêmes sites que la promo (comparaison juste)
    same = s[s["CA Promo"] > 0].groupby("code")[["CA Hors Promo", "Marge Hors Promo"]].sum()
    art["ca_hp_same"] = same["CA Hors Promo"]
    art["marge_hp_same"] = same["Marge Hors Promo"]
    art = art.fillna({"n_sites_sold": 0, "n_sites_promo": 0, "ca_hp_same": 0, "marge_hp_same": 0})

    # Prévisions (somme si plusieurs lignes par code, ex. par site)
    prev = promo.groupby("code").agg(prev_qty=("prev_qty", "sum"), prev_ca=("prev_ca", "sum"),
                                     funding=("funding", "sum"), label_promo=("label_promo", "first"),
                                     n_rows=("code", "size"))
    for c in ["prev_qty", "prev_ca", "funding"]:
        prev[c] = prev[c].where(promo.groupby("code")[c].count() > 0, np.nan)
    art = prev.join(art, how="left")
    art["found"] = art["CA"].notna()
    for c in agg_cols + ["n_sites_sold", "n_sites_promo", "n_sites_total", "ca_hp_same", "marge_hp_same"]:
        art[c] = art[c].fillna(0)
    art["label"] = art["label"].fillna(art["label_promo"])

    # Taux et marge
    art["rate_promo"] = art.apply(lambda r: safe_div(r["Marge Promo"], r["CA Promo"]), axis=1)
    art["rate_hp"] = art.apply(lambda r: safe_div(r["Marge Hors Promo"], r["CA Hors Promo"]), axis=1)
    art["rate_hp_same"] = art.apply(lambda r: safe_div(r["marge_hp_same"], r["ca_hp_same"]), axis=1)
    art["rate_total"] = art.apply(lambda r: safe_div(r["Marge"], r["CA"]), axis=1)
    art["rate_total_n1"] = art.apply(lambda r: safe_div(r["Marge N-1"], r["CA N-1"]), axis=1)
    # Référence hors promo : mêmes sites si dispo, sinon tout l'article
    art["rate_ref"] = art["rate_hp_same"].where(art["rate_hp_same"].notna(), art["rate_hp"])
    art["margin_delta_pts"] = art["rate_promo"] - art["rate_ref"]
    # Marge cédée = CA promo x (taux ref - taux promo), >0 = coût de la promo
    art["margin_ceded"] = (art["CA Promo"] * (art["rate_ref"] - art["rate_promo"])).where(art["CA Promo"] > 0, 0.0)
    art["margin_net"] = art["Marge Promo"] + art["funding"].fillna(0)
    art["marge_evo_n1"] = art["Marge"] - art["Marge N-1"]
    art["qty_evo_n1"] = art.apply(lambda r: safe_div(r["Qté Vente"] - r["Qté Vente N-1"], r["Qté Vente N-1"]), axis=1)
    art["promo_partial"] = (art["n_sites_promo"] > 0) & (art["n_sites_promo"] < art["n_sites_total"])

    # Réalisation
    art["real_qty"] = art.apply(lambda r: safe_div(r["Qté Vente"], r["prev_qty"]), axis=1)
    art["real_ca"] = art.apply(lambda r: safe_div(r["CA"], r["prev_ca"]), axis=1)
    art["real"] = art["real_qty"].where(art["real_qty"].notna(), art["real_ca"])
    art["gap_qty"] = art["Qté Vente"] - art["prev_qty"]
    art["gap_ca"] = art["CA"] - art["prev_ca"]

    # Verdict
    def verdict(r):
        if not r["found"] or r["Qté Vente"] <= 0:
            return "Pas de vente"
        rl = r["real"]
        if not pd.isna(rl) and rl < low:
            return "Sous-performance"
        if r["CA Promo"] <= 0:
            return "Promo non activée"
        if r["Marge Promo"] < 0:
            return "Marge promo négative"
        if pd.isna(rl):
            return "Conforme"
        if rl >= 1 and not pd.isna(r["margin_delta_pts"]) and r["margin_delta_pts"] < -tol:
            return "Volume sans marge"
        if rl > over:
            return "Prévision sous-estimée"
        if rl >= high:
            return "Succès"
        return "Conforme"

    art["verdict"] = art.apply(verdict, axis=1)
    art["action"] = art["verdict"].map(lambda v: VERDICTS[v]["action"])
    no_promo = (art["verdict"] == "Sous-performance") & (art["CA Promo"] <= 0)
    art.loc[no_promo, "action"] = "Aucun CA promo : vérifier d'abord l'activation de la promo (caisse, BDD), puis disponibilité et prévision."
    art["verdict_rank"] = art["verdict"].map(VERDICT_ORDER.index)
    art = art.reset_index().sort_values(["Marge", "CA"], ascending=False).reset_index(drop=True)

    # --- Contrôles -----------------------------------------------------------
    dup = promo[promo.duplicated("code", keep=False)]["code"].unique().tolist() if not meta["has_site"] else []
    not_found = art[~art["found"]]["code"].tolist()
    promo_not_listed = sales[(sales["CA Promo"] > 0) & (~sales["code"].isin(codes))]
    promo_not_listed = promo_not_listed.groupby(["code", "label"])["CA Promo"].sum().reset_index() \
        .sort_values("CA Promo", ascending=False)
    controls = {
        "duplicates": dup,
        "not_found": not_found,
        "promo_not_listed": promo_not_listed,
        "partial": art[art["promo_partial"]][["code", "label", "n_sites_promo", "n_sites_total"]],
        "converted_present": sorted(set(sales["site"]) & converted),
    }

    # --- Totaux --------------------------------------------------------------
    tot = {
        "n_articles": len(art),
        "qty": art["Qté Vente"].sum(), "prev_qty": art["prev_qty"].sum(skipna=True),
        "ca": art["CA"].sum(), "prev_ca": art["prev_ca"].sum(skipna=True),
        "marge": art["Marge"].sum(), "marge_promo": art["Marge Promo"].sum(),
        "ca_promo": art["CA Promo"].sum(), "ceded": art["margin_ceded"].sum(),
        "funding": art["funding"].sum(skipna=True) if meta["has_funding"] else np.nan,
        "marge_n1": art["Marge N-1"].sum(), "ca_n1": art["CA N-1"].sum(),
        "sites_promo": int(site[site["promo_active"]]["site"].nunique()),
        "sites_total": int(sales["site"].nunique()),
        "verdict_counts": art["verdict"].value_counts().reindex(VERDICT_ORDER, fill_value=0),
    }
    tot["real"] = safe_div(tot["qty"], tot["prev_qty"]) if meta["has_qty"] else safe_div(tot["ca"], tot["prev_ca"])
    tot["rate_promo"] = safe_div(tot["marge_promo"], tot["ca_promo"])
    tot["rate_hp"] = safe_div(art["Marge Hors Promo"].sum(), art["CA Hors Promo"].sum())
    tot["promo_share"] = safe_div(tot["ca_promo"], tot["ca"])
    return art, site, controls, tot


def synthesis_sentence(art, tot, meta):
    vc = tot["verdict_counts"]
    parts = []
    names = {"Prévision sous-estimée": "prévision{s} sous-estimée{s}", "Succès": "succès",
             "Conforme": "conforme{s}", "Volume sans marge": "volume{s} sans marge",
             "Marge promo négative": "marge{s} promo négative{s}", "Sous-performance": "sous-performance{s}",
             "Promo non activée": "promo{s} non activée{s}", "Pas de vente": "sans vente"}
    for v in VERDICT_ORDER:
        n = int(vc[v])
        if n:
            parts.append(f"{n} {names[v].format(s='s' if n > 1 else '')}")
    base = f"{tot['n_articles']} articles : " + ", ".join(parts) + "."
    base += f" {fmt_fcfa(tot['marge'])} FCFA de marge dégagée, dont {fmt_fcfa(tot['marge_promo'])} en promo."
    if tot["ceded"] > 0:
        base += f" La promo a coûté {fmt_fcfa(tot['ceded'])} FCFA de marge pour {fmt_fcfa(tot['ca_promo'])} de CA promo."
    return base


# ----------------------------------------------------------------------------
# Export Excel (valeurs figées, aucune formule)
# ----------------------------------------------------------------------------
def build_excel(art, site, controls, tot, filters, meta, params, sentence):
    wb = Workbook()
    hdr_fill = PatternFill("solid", fgColor=NAVY.lstrip("#"))
    hdr_font = Font(bold=True, color="FFFFFF", name="Calibri")
    thin = Side(style="thin", color="E5E5EA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    verdict_fill = {v: PatternFill("solid", fgColor=VERDICTS[v]["bg"].lstrip("#")) for v in VERDICTS}

    def write_table(ws, df, start_row=1, pct_cols=(), int_cols=(), verdict_col=None):
        for j, c in enumerate(df.columns, 1):
            cell = ws.cell(row=start_row, column=j, value=c)
            cell.fill, cell.font, cell.border = hdr_fill, hdr_font, border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for i, row in enumerate(df.itertuples(index=False), start_row + 1):
            for j, val in enumerate(row, 1):
                if isinstance(val, float) and np.isnan(val):
                    val = None
                cell = ws.cell(row=i, column=j, value=val)
                cell.border = border
                name = df.columns[j - 1]
                if name in pct_cols:
                    cell.number_format = "0.0%"
                elif name in int_cols:
                    cell.number_format = "#,##0"
                if verdict_col and name == verdict_col and val in verdict_fill:
                    cell.fill = verdict_fill[val]
                    cell.font = Font(bold=True)
        for j, c in enumerate(df.columns, 1):
            width = max(12, min(45, max(len(str(c)), *(len(str(v)) for v in df.iloc[:, j - 1].head(50))) + 2))
            ws.column_dimensions[get_column_letter(j)].width = width
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)

    # --- Synthèse ---
    ws = wb.active
    ws.title = "Synthèse"
    ws["A1"] = "Résultats Promo"
    ws["A1"].font = Font(bold=True, size=16, color=NAVY.lstrip("#"))
    per = f"{filters['period_from']:%d/%m/%Y} au {filters['period_to']:%d/%m/%Y}" if filters["period_from"] else "non lue"
    ws["A2"] = f"Période : {per} · Fournisseur : {filters['supplier'] or 'non lu'} · Généré le {datetime.now():%d/%m/%Y %H:%M}"
    ws["A4"] = sentence
    ws["A4"].font = Font(bold=True)
    ws["A4"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A4:H4")
    ws.row_dimensions[4].height = 45
    kpis = [
        ("Articles analysés", tot["n_articles"], "#,##0"),
        ("Quantité vendue", tot["qty"], "#,##0"),
        ("Quantité prévue", tot["prev_qty"] if meta["has_qty"] else None, "#,##0"),
        ("CA prévu (FCFA)", tot["prev_ca"] if meta["has_ca"] else None, "#,##0"),
        ("Taux de réalisation", tot["real"], "0.0%"),
        ("CA total (FCFA)", tot["ca"], "#,##0"),
        ("CA promo (FCFA)", tot["ca_promo"], "#,##0"),
        ("Poids promo dans le CA", tot["promo_share"], "0.0%"),
        ("Marge dégagée (FCFA)", tot["marge"], "#,##0"),
        ("dont marge promo (FCFA)", tot["marge_promo"], "#,##0"),
        ("Taux de marge promo", tot["rate_promo"], "0.0%"),
        ("Taux de marge hors promo", tot["rate_hp"], "0.0%"),
        ("Marge cédée par la promo (FCFA)", tot["ceded"], "#,##0"),
        ("Financement fournisseur (FCFA)", tot["funding"] if meta["has_funding"] else None, "#,##0"),
        ("Marge N-1 (FCFA)", tot["marge_n1"], "#,##0"),
        ("Sites avec promo active", f"{tot['sites_promo']} / {tot['sites_total']}", None),
    ]
    r = 6
    ws.cell(row=r, column=1, value="Indicateur").fill = hdr_fill
    ws.cell(row=r, column=1).font = hdr_font
    ws.cell(row=r, column=2, value="Valeur").fill = hdr_fill
    ws.cell(row=r, column=2).font = hdr_font
    for lab, val, nf in kpis:
        r += 1
        ws.cell(row=r, column=1, value=lab).border = border
        c = ws.cell(row=r, column=2, value=None if (isinstance(val, float) and np.isnan(val)) else val)
        c.border = border
        if nf:
            c.number_format = nf
    r += 2
    ws.cell(row=r, column=1, value="Verdict").fill = hdr_fill
    ws.cell(row=r, column=1).font = hdr_font
    ws.cell(row=r, column=2, value="Articles").fill = hdr_fill
    ws.cell(row=r, column=2).font = hdr_font
    for v in VERDICT_ORDER:
        r += 1
        c = ws.cell(row=r, column=1, value=v)
        c.fill, c.border = verdict_fill[v], border
        ws.cell(row=r, column=2, value=int(tot["verdict_counts"][v])).border = border
    r += 2
    ws.cell(row=r, column=1, value="Paramètres").font = Font(bold=True)
    for k, lab in [("low", "Seuil sous-performance"), ("high", "Seuil succès"), ("over", "Seuil prévision sous-estimée"),
                   ("margin_tol", "Tolérance marge (pts)")]:
        r += 1
        ws.cell(row=r, column=1, value=lab)
        ws.cell(row=r, column=2, value=params[k]).number_format = "0.0%"
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 20

    # --- Détail articles ---
    ws = wb.create_sheet("Détail articles")
    cols = {"code": "Code", "label": "Article", "prev_qty": "Prévision Qté", "Qté Vente": "Qté vendue",
            "real_qty": "Réalisation Qté", "gap_qty": "Écart Qté", "prev_ca": "Prévision CA", "CA": "CA total",
            "real_ca": "Réalisation CA", "gap_ca": "Écart CA", "CA Promo": "CA promo", "Marge": "Marge dégagée",
            "Marge Promo": "Marge promo", "rate_promo": "Taux marge promo", "rate_ref": "Taux marge hors promo (réf.)",
            "margin_delta_pts": "Écart taux (pts)", "margin_ceded": "Marge cédée", "funding": "Financement fournisseur",
            "margin_net": "Marge promo nette", "Marge N-1": "Marge N-1", "marge_evo_n1": "Évo marge vs N-1",
            "Qté Vente N-1": "Qté N-1", "qty_evo_n1": "Évo qté vs N-1", "n_sites_promo": "Sites promo",
            "n_sites_total": "Sites", "promo_partial": "Promo partielle", "verdict": "Verdict", "action": "Action recommandée"}
    if not meta["has_qty"]:
        for k in ["prev_qty", "real_qty", "gap_qty"]:
            cols.pop(k)
    if not meta["has_ca"]:
        for k in ["prev_ca", "real_ca", "gap_ca"]:
            cols.pop(k)
    if not meta["has_funding"]:
        for k in ["funding", "margin_net"]:
            cols.pop(k)
    df = art[list(cols)].rename(columns=cols)
    df["Promo partielle"] = df["Promo partielle"].map({True: "Oui", False: "Non"})
    write_table(ws, df, pct_cols=("Réalisation Qté", "Réalisation CA", "Taux marge promo", "Taux marge hors promo (réf.)",
                                  "Écart taux (pts)", "Évo qté vs N-1"),
                int_cols=("Prévision Qté", "Qté vendue", "Écart Qté", "Prévision CA", "CA total", "Écart CA", "CA promo",
                          "Marge dégagée", "Marge promo", "Marge cédée", "Financement fournisseur", "Marge promo nette",
                          "Marge N-1", "Évo marge vs N-1", "Qté N-1"), verdict_col="Verdict")

    # --- Détail sites ---
    ws = wb.create_sheet("Détail sites")
    scols = {"code": "Code", "label": "Article", "site": "Site", "Qté Vente": "Qté vendue", "Qté Vente N-1": "Qté N-1",
             "qty_evo": "Évo qté vs N-1", "CA": "CA total", "CA Promo": "CA promo", "Marge": "Marge dégagée",
             "Marge Promo": "Marge promo", "rate_promo": "Taux marge promo", "rate_hp": "Taux marge hors promo",
             "promo_active": "Promo active", "n1_comparable": "N-1 comparable"}
    sdf = site[list(scols)].rename(columns=scols).sort_values(["Article", "CA promo"], ascending=[True, False])
    sdf["Promo active"] = sdf["Promo active"].map({True: "Oui", False: "Non"})
    sdf["N-1 comparable"] = sdf["N-1 comparable"].map({True: "Oui", False: "Non (site converti)"})
    write_table(ws, sdf, pct_cols=("Évo qté vs N-1", "Taux marge promo", "Taux marge hors promo"),
                int_cols=("Qté vendue", "Qté N-1", "CA total", "CA promo", "Marge dégagée", "Marge promo"))

    # --- Contrôles ---
    ws = wb.create_sheet("Contrôles")
    rows = [("Périmètre", "Période", per), ("Périmètre", "Fournisseur", filters["supplier"] or "non lu"),
            ("Périmètre", "Sites dans l'extraction", ", ".join(filters["sites"]) or "non lus"),
            ("Liste promo", "Codes trouvés", f"{int(art['found'].sum())} / {len(art)}"),
            ("Liste promo", "Codes introuvables", ", ".join(controls["not_found"]) or "aucun"),
            ("Liste promo", "Doublons", ", ".join(controls["duplicates"]) or "aucun"),
            ("Liste promo", "Articles en promo non listés", str(len(controls["promo_not_listed"]))),
            ("Comparaisons", "Promo partielle (sites)", ", ".join(controls["partial"]["label"]) or "aucune"),
            ("Comparaisons", "Sites au N-1 non comparable", ", ".join(controls["converted_present"]) or "aucun")]
    cdf = pd.DataFrame(rows, columns=["Bloc", "Contrôle", "Résultat"])
    write_table(ws, cdf)
    if len(controls["promo_not_listed"]):
        r0 = len(cdf) + 4
        ws.cell(row=r0 - 1, column=1, value="Articles vendus en promo mais absents de la liste").font = Font(bold=True)
        pnl = controls["promo_not_listed"].rename(columns={"code": "Code", "label": "Article", "CA Promo": "CA promo"})
        write_table(ws, pnl, start_row=r0, int_cols=("CA promo",))

    # --- Bilan fournisseur (sans marge) ---
    ws = wb.create_sheet("Bilan fournisseur")
    ws["A1"] = f"Bilan promo · {filters['supplier'] or 'Fournisseur'} · {per}"
    ws["A1"].font = Font(bold=True, size=14, color=NAVY.lstrip("#"))
    ws["A2"] = "Document partageable : volumes, CA et sites. Aucune marge ni prix de revient."
    ws["A2"].font = Font(italic=True, color="5B6472")
    bcols = {"code": "Code", "label": "Article", "prev_qty": "Prévision Qté", "Qté Vente": "Qté vendue",
             "real_qty": "Réalisation", "Qté Vente N-1": "Qté N-1", "CA": "CA total (FCFA)", "CA Promo": "CA en promo (FCFA)",
             "n_sites_promo": "Sites en promo", "n_sites_total": "Sites", "verdict": "Constat"}
    if not meta["has_qty"]:
        for k in ["prev_qty", "real_qty"]:
            bcols.pop(k)
    bdf = art[list(bcols)].rename(columns=bcols)
    bdf["Constat"] = bdf["Constat"].replace({"Volume sans marge": "Succès", "Marge promo négative": "Succès"})
    write_table(ws, bdf, start_row=4, pct_cols=("Réalisation",),
                int_cols=("Prévision Qté", "Qté vendue", "Qté N-1", "CA total (FCFA)", "CA en promo (FCFA)"))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


