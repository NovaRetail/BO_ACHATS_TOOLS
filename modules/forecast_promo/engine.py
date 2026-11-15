"""SmartBuyer Hub · Forecast Promo — moteur de calcul et export.

Deux temps, un même format de fichiers :
  ① Avant la promo : la prévision de chaque acheteur est confrontée au CA N-1 de la même période
     (extraction PBI sur les dates de la promo à venir : les colonnes N-1 sont remplies).
  ② Après la promo : prévu vs réalisé, sans compensation entre articles, et comparaison à N-1.

Entrées : extraction ventes PBI hiérarchique (Departement › Rayon › Famille › Sous-Famille › Article › Site, montants HT)
          + liste promo (Code, Prévision Qté, Prix promo, PA, TVA).
Aucune dépendance Streamlit : la page gère le cache et l'affichage.
"""
from __future__ import annotations

import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

NAVY = "24235C"

LEVELS = [("dep", "Département"), ("ray", "Rayon"), ("fam", "Famille"), ("sf", "Sous-famille")]
LEVEL_SOURCES = {
    "dep": ["Departement", "Département"],
    "ray": ["Rayon"],
    "fam": ["Famille"],
    "sf": ["Sous-Famille", "Sous-famille", "Sous Famille", "Sous famille", "SousFamille"],
}
NOT_FOUND = "Introuvable dans l'extraction"
NO_SF = "Sous-famille non extraite"

# Sites dont le N-1 n'est pas comparable (conversions Supeco -> Market en 2026)
DEFAULT_CONVERTED_SITES = ["Market Cité verte", "Market Aboboté", "Market Toit rouge"]

SALES_REQUIRED = ["Article", "Site nom long", "CA", "Marge", "Qté Vente"]
N1_COLS = ["CA N-1", "Marge N-1", "Qté Vente N-1", "Marge Hors Promo N-1", "Marge Promo N-1"]

PROMO_ALIASES = {
    "code": ["code", "code article", "article", "ref", "référence", "reference", "sku"],
    "label": ["libelle", "libellé", "designation", "désignation", "label", "nom"],
    "prev_qty": ["prevision qte", "prévision qté", "prévision qte", "prevision", "prévision", "prev qte", "qte prevue",
                 "qté prévue", "prevision qty", "forecast", "qty", "qte", "qté", "quantite", "quantité"],
    "price": ["prix promo", "prix", "prix de vente promo", "pv promo", "prix vente", "pvc promo"],
    "pa": ["pa", "prix d'achat", "prix achat", "pa ht", "prix d'achat ht", "pa net", "pa net ht", "cout", "coût"],
    "vat": ["tva", "taux tva", "tva %", "taux de tva", "vat"],
}

# Statuts (kind du thème : ok / att / crit / info)
RESULT = {"Fiable": "ok", "Surestimée": "crit", "Sous-estimée": "att", "Pas de vente": "crit"}
CONTROL = {"Cohérente": "ok", "À surveiller": "att", "À justifier": "crit", "Sans historique": "att"}
VS_N1 = {"Mieux que N-1": "ok", "Stable": "info", "Moins bien que N-1": "crit", "Pas de N-1": "info"}
FILL = {"ok": "E1F3F3", "att": "FDF1E2", "crit": "FCE4E8", "info": "ECEDF4"}

LEXIQUE = [
    ("Prévision Qté", "Quantité que l'acheteur s'engage à vendre pendant la promo, tous magasins confondus."),
    ("Prix promo", "Prix de vente affiché pendant la promo, TTC par défaut."),
    ("PA", "Prix d'achat HT de l'article."),
    ("TVA", "Taux de TVA de l'article (18, 0,18 ou 18 %), sert à passer le prix promo en HT."),
    ("CA prévu HT", "Prévision Qté × Prix promo ÷ (1 + TVA). Calculé par le module, jamais saisi."),
    ("Marge prévue", "Prévision Qté × (Prix promo HT − PA)."),
    ("CA N-1", "CA réalisé sur les mêmes dates l'an dernier, dans les mêmes magasins."),
    ("Prévu / N-1", "CA prévu ÷ CA N-1. × 2 veut dire que l'acheteur promet deux fois les ventes de l'an dernier."),
    ("À surveiller · À justifier", "Prévision au-dessus de × 1,3 ou × 2 le N-1 (seuils réglables) : à faire expliquer avant validation."),
    ("Sans historique", "Article sans vente l'an dernier sur la période : la prévision n'a pas de point de comparaison."),
    ("Réalisation", "Réalisé ÷ prévu. 100 % = la prévision a été tenue."),
    ("Surestimée · Fiable · Sous-estimée", "Réalisation du CA sous 80 %, entre 80 et 120 %, au-dessus de 120 % (seuils réglables)."),
    ("CA promis non réalisé", "Somme des manques article par article. Un article qui dépasse ne compense pas un article gonflé."),
    ("Biais", "Part des articles surestimés ou sans vente dans un niveau. Orange dès 25 %, rouge dès 50 %."),
    ("Vs N-1", "Évolution du CA réalisé par rapport à l'an dernier, « stable » dans ± 5 % (réglable)."),
    ("N-1 non comparable", "Article vendu dans un magasin converti en 2026 : la comparaison avec l'an dernier est faussée."),
]


# ----------------------------------------------------------------------------
# Formats
# ----------------------------------------------------------------------------
def _nan(x) -> bool:
    return x is None or (isinstance(x, (float, np.floating)) and np.isnan(x))


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


def fmt_pct(x, digits=0, sign=False):
    if _nan(x) or np.isinf(x):
        return "—"
    s = f"{x * 100:+.{digits}f}" if sign else f"{x * 100:.{digits}f}"
    return s.replace(".", ",") + " %"


def fmt_ratio(x):
    return "—" if _nan(x) or np.isinf(x) else f"× {x:.1f}".replace(".", ",")


def div(a, b):
    """Division vectorisée ou scalaire, NaN si dénominateur nul."""
    if isinstance(a, pd.Series) or isinstance(b, pd.Series):
        b = pd.Series(b) if not isinstance(b, pd.Series) else b
        return (a / b.where(b != 0)).astype(float)
    return np.nan if (b is None or _nan(b) or b == 0) else a / b


def normalize_code(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return re.sub(r"\s+", "", s) or None


def strip_code(s):
    return re.sub(r"^\s*\d+\s*-\s*", "", str(s)).strip()


# ----------------------------------------------------------------------------
# Lecture
# ----------------------------------------------------------------------------
def read_any(data: bytes, name: str) -> pd.DataFrame:
    if name.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(io.BytesIO(data))
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(data), sep=None, engine="python", encoding=enc)
        except Exception:  # noqa: BLE001
            continue
    raise ValueError("format non lu (xlsx ou csv attendu).")


def parse_filters(df: pd.DataFrame) -> dict:
    """Lit le bloc « Filtres appliqués » en pied d'export PBI."""
    info = {"period_from": None, "period_to": None, "supplier": None, "sites": []}
    txt = next((v for v in df[df.columns[0]].dropna().astype(str) if "Filtres appliqués" in v), None)
    if not txt:
        return info
    m = re.search(r"Fournisseur[^\n]*?est\s+(.+)", txt)
    if m:
        info["supplier"] = m.group(1).strip()
    m = re.search(r"Date est le ou après le (\d{2}/\d{2}/\d{4}) et est avant le (\d{2}/\d{2}/\d{4})", txt)
    if m:
        info["period_from"] = datetime.strptime(m.group(1), "%d/%m/%Y").date()
        info["period_to"] = (datetime.strptime(m.group(2), "%d/%m/%Y") - pd.Timedelta(days=1)).date()
    m = re.search(r"Site nom long est (.+)", txt)
    if m:
        info["sites"] = [s.strip() for s in re.split(r",\s*|\s+ou\s+", m.group(1).strip()) if s.strip()]
    return info


def parse_sales(data: bytes, name: str):
    df = read_any(data, name)
    missing = [c for c in SALES_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("colonnes manquantes : " + ", ".join(missing) + ".")
    filters = parse_filters(df)
    level_cols = {}
    for key, srcs in LEVEL_SOURCES.items():
        col = next((c for c in srcs if c in df.columns), None)
        if col:
            level_cols[key] = col
    d = df.copy()
    for col in list(level_cols.values()) + ["Article"]:
        d[col] = d[col].ffill()
    d = d[d["Site nom long"].notna() & (d["Site nom long"].astype(str) != "Total") & (d["Article"].astype(str) != "Total")].copy()
    for col in level_cols.values():
        d = d[d[col].astype(str) != "Total"]
    out = pd.DataFrame({
        "code": d["Article"].astype(str).str.extract(r"^\s*(\d+)")[0].map(normalize_code),
        "label": d["Article"].map(strip_code),
        "site": d["Site nom long"].map(strip_code),
    })
    for key, _ in LEVELS:
        out[key] = d[level_cols[key]].map(strip_code) if key in level_cols else (NO_SF if key == "sf" else "—")
    for c in ["CA", "Marge", "Qté Vente"] + N1_COLS:
        out[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0) if c in d.columns else 0.0
    if "Marge N-1" not in d.columns:  # absente de l'export PBI : reconstituée
        out["Marge N-1"] = out["Marge Hors Promo N-1"] + out["Marge Promo N-1"]
    meta = {"has_sf": "sf" in level_cols, "levels_found": [lab for k, lab in LEVELS if k in level_cols]}
    return out.drop(columns=["Marge Hors Promo N-1", "Marge Promo N-1"]).reset_index(drop=True), filters, meta


def parse_rate(x):
    if pd.isna(x):
        return np.nan
    v = pd.to_numeric(str(x).replace("%", "").replace(",", ".").strip(), errors="coerce")
    return np.nan if pd.isna(v) else (v / 100 if v > 1 else v)


def parse_promo(data: bytes, name: str, price_ttc: bool = True):
    df = read_any(data, name).dropna(how="all")
    normed = {re.sub(r"\s+", " ", str(c).strip().lower()): c for c in df.columns}
    cols = {}
    for key, aliases in PROMO_ALIASES.items():
        hit = next((normed[a] for a in aliases if a in normed and normed[a] not in cols.values()), None)
        if hit is not None:
            cols[key] = hit
    for key, lab in [("code", "Code"), ("prev_qty", "Prévision Qté"), ("price", "Prix promo")]:
        if key not in cols:
            raise ValueError(f"colonne « {lab} » introuvable.")
    p = pd.DataFrame({"code": df[cols["code"]].map(normalize_code)})
    p["label_promo"] = df[cols["label"]].astype(str).str.strip() if "label" in cols else ""
    p["prev_qty"] = pd.to_numeric(df[cols["prev_qty"]], errors="coerce")
    p["price"] = pd.to_numeric(df[cols["price"]], errors="coerce")
    p["pa"] = pd.to_numeric(df[cols["pa"]], errors="coerce") if "pa" in cols else np.nan
    p["vat"] = df[cols["vat"]].map(parse_rate) if "vat" in cols else np.nan
    has_vat, has_pa = "vat" in cols, "pa" in cols
    p["price_ht"] = p["price"] / (1 + p["vat"].fillna(0)) if (price_ttc and has_vat) else p["price"]
    p["prev_ca"] = p["prev_qty"] * p["price_ht"]
    p["prev_marge"] = p["prev_qty"] * (p["price_ht"] - p["pa"]) if has_pa else np.nan
    p = p[p["code"].notna()].reset_index(drop=True)
    ca_mode = ("Qté × Prix promo ÷ (1 + TVA)" if (price_ttc and has_vat) else
               "Qté × Prix promo (prix déclaré HT)" if not price_ttc else "Qté × Prix promo (TVA absente)")
    return p, {"has_pa": has_pa, "has_vat": has_vat, "price_ttc": price_ttc, "ca_mode": ca_mode}


# ----------------------------------------------------------------------------
# Calculs (vectorisés)
# ----------------------------------------------------------------------------
def compute(sales: pd.DataFrame, promo: pd.DataFrame, params: dict):
    converted = set(params["converted_sites"])
    s = sales[sales["code"].isin(set(promo["code"]))]
    num = ["CA", "Marge", "Qté Vente", "CA N-1", "Marge N-1", "Qté Vente N-1"]
    agg = s.groupby("code")[num].sum()
    first = s.groupby("code")[["label"] + [k for k, _ in LEVELS]].first()
    conv = s[s["site"].isin(converted)].groupby("code").size().rename("n_conv")

    pv = promo.groupby("code").agg(prev_qty=("prev_qty", "sum"), prev_ca=("prev_ca", "sum"), prev_marge=("prev_marge", "sum"),
                                   price=("price", "mean"), pa=("pa", "mean"), label_promo=("label_promo", "first"))
    cnt = promo.groupby("code")[["prev_qty", "prev_ca", "prev_marge"]].count()
    for c in ["prev_qty", "prev_ca", "prev_marge"]:
        pv[c] = pv[c].where(cnt[c] > 0)

    a = pv.join(agg).join(first).join(conv)
    a["found"] = a["CA"].notna()
    a[num] = a[num].fillna(0.0)
    a["n_conv"] = a["n_conv"].fillna(0)
    a["label"] = a["label"].fillna(a["label_promo"])
    for k, _ in LEVELS:
        a[k] = a[k].fillna(NOT_FOUND)

    # ① Avant la promo : prévision vs N-1
    a["ratio"] = div(a["prev_ca"], a["CA N-1"])
    a["control"] = np.select([a["CA N-1"] <= 0, a["ratio"] > params["justify"], a["ratio"] > params["watch"]],
                             ["Sans historique", "À justifier", "À surveiller"], "Cohérente")

    # ② Après la promo : prévu vs réalisé
    a["real_qty"] = div(a["Qté Vente"], a["prev_qty"])
    a["real_ca"] = div(a["CA"], a["prev_ca"])
    a["real_marge"] = div(a["Marge"], a["prev_marge"])
    a["gap_ca"] = a["CA"] - a["prev_ca"]
    a["short"] = (a["prev_ca"] - a["CA"]).clip(lower=0).fillna(0)
    a["result"] = np.select([a["Qté Vente"] <= 0, a["real_ca"] < params["low"], a["real_ca"] > params["high"]],
                            ["Pas de vente", "Surestimée", "Sous-estimée"], "Fiable")
    a["over"] = a["result"].isin(["Surestimée", "Pas de vente"])
    a["rate"] = div(a["Marge"], a["CA"])
    a["prev_rate"] = div(a["prev_marge"], a["prev_ca"])
    a["rate_n1"] = div(a["Marge N-1"], a["CA N-1"])
    a["ca_evo"] = div(a["CA"] - a["CA N-1"], a["CA N-1"])
    a["marge_evo"] = div(a["Marge"] - a["Marge N-1"], a["Marge N-1"])
    b = params["stable"]
    a["vs_n1"] = np.select([a["CA N-1"] <= 0, a["ca_evo"] > b, a["ca_evo"] < -b], ["Pas de N-1", "Mieux que N-1", "Moins bien que N-1"], "Stable")
    a["n1_partial"] = a["n_conv"] > 0
    a = a.reset_index()

    tot = summarize(a)
    tot["sites"] = int(sales["site"].nunique())
    tot["phase"] = "avant" if tot["ca"] <= 0 else "après"
    controls = {
        "not_found": a.loc[~a["found"], "code"].tolist(),
        "duplicates": promo.loc[promo.duplicated("code", keep=False), "code"].unique().tolist(),
        "no_price": promo.loc[promo["prev_ca"].isna(), "code"].tolist(),
        "no_pa": promo.loc[promo["pa"].isna(), "code"].tolist(),
        "neg_marge": promo.loc[promo["prev_marge"] < 0, "code"].tolist(),
        "converted": sorted(set(sales["site"]) & converted),
    }
    return a, tot, controls


def summarize(d: pd.DataFrame) -> dict:
    """Agrégats d'un ensemble d'articles (total, ou un nœud de la matrice)."""
    t = {
        "n": len(d), "prev_qty": d["prev_qty"].sum(), "qty": d["Qté Vente"].sum(),
        "prev_ca": d["prev_ca"].sum(), "ca": d["CA"].sum(), "ca_n1": d["CA N-1"].sum(),
        "prev_marge": d["prev_marge"].sum(min_count=1), "marge": d["Marge"].sum(), "marge_n1": d["Marge N-1"].sum(),
        "short": d["short"].sum(), "over": int(d["over"].sum()),
        "justify": int((d["control"] == "À justifier").sum()), "watch": int((d["control"] == "À surveiller").sum()),
        "nohist": int((d["control"] == "Sans historique").sum()),
        "justify_ca": d.loc[d["control"] == "À justifier", "prev_ca"].sum(),
    }
    t["real_qty"] = div(t["qty"], t["prev_qty"])
    t["real_ca"] = div(t["ca"], t["prev_ca"])
    t["real_marge"] = div(t["marge"], t["prev_marge"])
    t["ratio"] = div(t["prev_ca"], t["ca_n1"])
    t["bias"] = div(t["over"], t["n"])
    t["rate"] = div(t["marge"], t["ca"])
    t["prev_rate"] = div(t["prev_marge"], t["prev_ca"])
    t["rate_n1"] = div(t["marge_n1"], t["ca_n1"])
    t["ca_evo"] = div(t["ca"] - t["ca_n1"], t["ca_n1"])
    t["marge_evo"] = div(t["marge"] - t["marge_n1"], t["marge_n1"])
    return t


def sentence_before(t) -> str:
    return (f"{t['n']} articles : CA prévu {fmt_fcfa(t['prev_ca'])} pour {fmt_fcfa(t['ca_n1'])} réalisés l'an dernier sur la même période "
            f"({fmt_ratio(t['ratio'])}). {t['justify']} prévision(s) à justifier, {t['watch']} à surveiller, {t['nohist']} sans historique. "
            f"Les prévisions à justifier portent {fmt_fcfa(t['justify_ca'])} de CA prévu ({fmt_pct(div(t['justify_ca'], t['prev_ca']))}).")


def sentence_after(t) -> str:
    s = (f"CA réalisé {fmt_fcfa(t['ca'])} pour {fmt_fcfa(t['prev_ca'])} prévus ({fmt_pct(t['real_ca'])}). "
         f"CA promis non réalisé : {fmt_fcfa(t['short'])} FCFA. Biais : {t['over']} article(s) sur {t['n']} surestimés ou sans vente "
         f"({fmt_pct(t['bias'])}).")
    if t["ca_n1"] > 0:
        s += f" Vs N-1 : CA {fmt_pct(t['ca_evo'], 0, True)}, marge {fmt_pct(t['marge_evo'], 0, True)}."
    return s


# ----------------------------------------------------------------------------
# Export Excel (valeurs uniquement, aucune formule)
# ----------------------------------------------------------------------------
def build_excel(a, tot, controls, filters, meta, params) -> bytes:
    wb = Workbook()
    hdr_fill, hdr_font = PatternFill("solid", fgColor=NAVY), Font(bold=True, color="FFFFFF")
    status_fill = {k: PatternFill("solid", fgColor=FILL[v]) for d in (RESULT, CONTROL, VS_N1) for k, v in d.items()}

    def table(ws, df, row=1, pct=(), num=(), ratio=(), status=()):
        for j, c in enumerate(df.columns, 1):
            cell = ws.cell(row=row, column=j, value=c)
            cell.fill, cell.font = hdr_fill, hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for i, rec in enumerate(df.itertuples(index=False), row + 1):
            for j, v in enumerate(rec, 1):
                if isinstance(v, (float, np.floating)) and (np.isnan(v) or np.isinf(v)):
                    v = None
                cell = ws.cell(row=i, column=j, value=v)
                name = df.columns[j - 1]
                if name in pct:
                    cell.number_format = "0.0%"
                elif name in num:
                    cell.number_format = "#,##0"
                elif name in ratio:
                    cell.number_format = '"× "0.0'
                if name in status and v in status_fill:
                    cell.fill, cell.font = status_fill[v], Font(bold=True)
        for j, c in enumerate(df.columns, 1):
            ws.column_dimensions[get_column_letter(j)].width = max(12, min(40, len(str(c)) + 4))
        ws.freeze_panes = ws.cell(row=row + 1, column=1)

    per = f"{filters['period_from']:%d/%m/%Y} au {filters['period_to']:%d/%m/%Y}" if filters["period_from"] else "non lue"
    hier = {k: lab for k, lab in LEVELS}

    ws = wb.active
    ws.title = "Synthèse"
    ws["A1"] = "Forecast Promo"
    ws["A1"].font = Font(bold=True, size=16, color=NAVY)
    ws["A2"] = f"Période : {per} · Fournisseur : {filters['supplier'] or 'non lu'} · CA prévu = {meta['ca_mode']} · {datetime.now():%d/%m/%Y %H:%M}"
    ws["A4"], ws["A5"] = "① Avant la promo : " + sentence_before(tot), "② Après la promo : " + sentence_after(tot)
    for r in (4, 5):
        ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(f"A{r}:F{r}")
        ws.row_dimensions[r].height = 46
    kpis = [("CA prévu HT", tot["prev_ca"], "#,##0"), ("CA N-1 même période", tot["ca_n1"], "#,##0"), ("Prévu / N-1", tot["ratio"], '"× "0.00'),
            ("Prévisions à justifier", tot["justify"], "0"), ("Prévisions à surveiller", tot["watch"], "0"), ("Sans historique", tot["nohist"], "0"),
            ("CA réalisé HT", tot["ca"], "#,##0"), ("Réalisation CA", tot["real_ca"], "0%"), ("CA promis non réalisé", tot["short"], "#,##0"),
            ("Biais (articles surestimés ou sans vente)", tot["bias"], "0%"), ("Marge prévue", tot["prev_marge"], "#,##0"),
            ("Marge réalisée", tot["marge"], "#,##0"), ("Taux de marge prévu", tot["prev_rate"], "0.0%"), ("Taux de marge réalisé", tot["rate"], "0.0%"),
            ("Marge N-1", tot["marge_n1"], "#,##0"), ("Taux de marge N-1", tot["rate_n1"], "0.0%"),
            ("Évolution CA vs N-1", tot["ca_evo"], "+0.0%;-0.0%"), ("Évolution marge vs N-1", tot["marge_evo"], "+0.0%;-0.0%")]
    for j, h in enumerate(["Indicateur", "Valeur"], 1):
        c = ws.cell(row=7, column=j, value=h)
        c.fill, c.font = hdr_fill, hdr_font
    for i, (lab, v, nf) in enumerate(kpis, 8):
        ws.cell(row=i, column=1, value=lab)
        c = ws.cell(row=i, column=2, value=None if _nan(v) else v)
        c.number_format = nf
    r = 8 + len(kpis) + 1
    ws.cell(row=r, column=1, value="Paramètres").font = Font(bold=True)
    for lab, v in [("À surveiller au-delà de (× N-1)", params["watch"]), ("À justifier au-delà de (× N-1)", params["justify"]),
                   ("Surestimée sous", params["low"]), ("Sous-estimée au-delà de", params["high"]), ("Stable vs N-1 (±)", params["stable"])]:
        r += 1
        ws.cell(row=r, column=1, value=lab)
        ws.cell(row=r, column=2, value=v)
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 42, 18

    base = {k: hier[k] for k, _ in LEVELS} | {"code": "Code", "label": "Article"}
    ws = wb.create_sheet("① Avant la promo")
    cols = base | {"prev_qty": "Prévision Qté", "price": "Prix promo", "pa": "PA", "prev_ca": "CA prévu HT",
                   "CA N-1": "CA N-1", "ratio": "Prévu / N-1", "control": "Contrôle"}
    table(ws, a.sort_values("ratio", ascending=False)[list(cols)].rename(columns=cols), num=("Prévision Qté", "Prix promo", "PA", "CA prévu HT", "CA N-1"),
          ratio=("Prévu / N-1",), status=("Contrôle",))

    ws = wb.create_sheet("② Après la promo")
    cols = base | {"prev_qty": "Prévision Qté", "Qté Vente": "Qté réalisée", "real_qty": "Réal. Qté", "prev_ca": "CA prévu HT", "CA": "CA réalisé",
                   "real_ca": "Réal. CA", "short": "CA promis non réalisé", "result": "Prévision", "prev_marge": "Marge prévue",
                   "Marge": "Marge réalisée", "prev_rate": "Taux prévu", "rate": "Taux réalisé", "CA N-1": "CA N-1", "ca_evo": "Évo CA vs N-1",
                   "Marge N-1": "Marge N-1", "rate_n1": "Taux N-1", "vs_n1": "Vs N-1", "n1_partial": "N-1 non comparable"}
    df = a.sort_values("short", ascending=False)[list(cols)].rename(columns=cols)
    df["N-1 non comparable"] = df["N-1 non comparable"].map({True: "Oui", False: ""})
    table(ws, df, pct=("Réal. Qté", "Réal. CA", "Taux prévu", "Taux réalisé", "Évo CA vs N-1", "Taux N-1"),
          num=("Prévision Qté", "Qté réalisée", "CA prévu HT", "CA réalisé", "CA promis non réalisé", "Marge prévue", "Marge réalisée", "CA N-1", "Marge N-1"),
          status=("Prévision", "Vs N-1"))

    ws = wb.create_sheet("Contrôles")
    rows = [("Période", per), ("Fournisseur", filters["supplier"] or "non lu"), ("Sites", ", ".join(filters["sites"]) or "non lus"),
            ("CA prévu", meta["ca_mode"]), ("Codes introuvables", ", ".join(controls["not_found"]) or "aucun"),
            ("Doublons", ", ".join(controls["duplicates"]) or "aucun"), ("Prix promo manquant", ", ".join(controls["no_price"]) or "aucun"),
            ("PA manquant", ", ".join(controls["no_pa"]) or "aucun"), ("Marge prévue négative", ", ".join(controls["neg_marge"]) or "aucune"),
            ("Sites au N-1 non comparable", ", ".join(controls["converted"]) or "aucun")]
    table(ws, pd.DataFrame(rows, columns=["Contrôle", "Résultat"]))
    ws.column_dimensions["B"].width = 80

    ws = wb.create_sheet("Lexique")
    table(ws, pd.DataFrame(LEXIQUE, columns=["Terme", "Définition"]))
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 34, 100

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
