"""SmartBuyer Hub · Forecast Promo — moteur de calcul et export.

Deux temps, un même format de fichiers :
  ① Avant la promo : la prévision de chaque acheteur (en quantité) est confrontée à une base, au choix :
     - base récente : ventes hors promo des semaines qui précèdent la promo (extraction de référence facultative),
       ramenées à la durée de la promo ;
     - N-1 : quantité vendue sur la même période l'an dernier (extraction PBI sur les dates de la promo à venir).
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
SPLIT_COLS = ["CA Promo", "CA Hors Promo", "CA Promo N-1", "CA Hors Promo N-1"]

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
CONTROL = {"À justifier": "crit", "À surveiller": "att", "Sous la base": "att", "Sans historique": "att", "Cohérente": "ok"}
CONTROL_ORDER = list(CONTROL)
BASE_LABEL = {"recent": "Base récente (ventes hors promo des semaines précédentes)", "n1": "N-1 même période"}
VS_N1 = {"Mieux que N-1": "ok", "Stable": "info", "Moins bien que N-1": "crit", "Pas de N-1": "info"}
FILL = {"ok": "E1F3F3", "att": "FDF1E2", "crit": "FCE4E8", "info": "ECEDF4"}

LEXIQUE = [
    ("Prévision Qté", "Quantité que l'acheteur s'engage à vendre pendant la promo, tous magasins confondus."),
    ("Prix promo", "Prix de vente affiché pendant la promo, TTC par défaut."),
    ("PA", "Prix d'achat HT de l'article."),
    ("TVA", "Taux de TVA de l'article (18, 0,18 ou 18 %), sert à passer le prix promo en HT."),
    ("CA prévu HT", "Prévision Qté × Prix promo ÷ (1 + TVA). Calculé par le module, jamais saisi."),
    ("Marge prévue", "Prévision Qté × (Prix promo HT − PA)."),
    ("Base", "Quantité que l'article vendrait sans promo sur la durée de l'opération. Base récente : ventes hors promo "
              "des semaines précédentes, ramenées à la durée de la promo. Base N-1 : quantité vendue l'an dernier sur les mêmes dates."),
    ("Multiplicateur", "Prévision Qté ÷ Base. × 2 veut dire que l'acheteur promet de doubler les ventes normales grâce à la promo."),
    ("À surveiller · À justifier", "Multiplicateur au-delà de × 1,5 ou × 2 (seuils réglables) : à faire expliquer avant validation."),
    ("Sous la base", "Prévision inférieure aux ventes normales : prudence excessive ou article en perte de vitesse."),
    ("Sans historique", "Aucune vente sur la base choisie : la prévision n'a pas de point de comparaison."),
    ("En promo N-1", "L'article était déjà en promo l'an dernier sur la même période (CA Promo N-1 > 0)."),
    ("CA N-1", "CA réalisé sur les mêmes dates l'an dernier, dans les mêmes magasins."),
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
    for c in ["CA", "Marge", "Qté Vente"] + N1_COLS + SPLIT_COLS:
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
# Base récente (extraction de référence facultative)
# ----------------------------------------------------------------------------
def period_days(filters: dict):
    if filters.get("period_from") and filters.get("period_to"):
        return (filters["period_to"] - filters["period_from"]).days + 1
    return None


def build_base(ref_sales: pd.DataFrame, ref_days: int, promo_days: int) -> pd.DataFrame:
    """Ventes hors promo de la période de référence, ramenées à la durée de la promo.
    L'export ne donne pas la quantité hors promo : elle est estimée au prorata du CA hors promo."""
    g = ref_sales.groupby("code")[["Qté Vente", "CA", "CA Hors Promo", "CA Promo"]].sum()
    share = (g["CA Hors Promo"] / g["CA"].where(g["CA"] > 0)).fillna(1.0).clip(0, 1)
    k = promo_days / ref_days
    return pd.DataFrame({"base_qty": g["Qté Vente"] * share * k, "base_ca": g["CA Hors Promo"] * k,
                         "ref_promo_share": (g["CA Promo"] / g["CA"].where(g["CA"] > 0)).fillna(0.0)})


# ----------------------------------------------------------------------------
# Calculs (vectorisés)
# ----------------------------------------------------------------------------
def compute(sales: pd.DataFrame, promo: pd.DataFrame, params: dict, base: pd.DataFrame | None = None):
    converted = set(params["converted_sites"])
    s = sales[sales["code"].isin(set(promo["code"]))]
    num = ["CA", "Marge", "Qté Vente", "CA N-1", "Marge N-1", "Qté Vente N-1", "CA Promo N-1", "CA Hors Promo N-1"]
    agg = s.groupby("code")[num].sum()
    first = s.groupby("code")[["label"] + [k for k, _ in LEVELS]].first()
    conv = s[s["site"].isin(converted)].groupby("code").size().rename("n_conv")

    pv = promo.groupby("code").agg(prev_qty=("prev_qty", "sum"), prev_ca=("prev_ca", "sum"), prev_marge=("prev_marge", "sum"),
                                   price=("price", "mean"), price_ht=("price_ht", "mean"), pa=("pa", "mean"),
                                   label_promo=("label_promo", "first"))
    cnt = promo.groupby("code")[["prev_qty", "prev_ca", "prev_marge"]].count()
    for c in ["prev_qty", "prev_ca", "prev_marge"]:
        pv[c] = pv[c].where(cnt[c] > 0)

    a = pv.join(agg).join(first).join(conv)
    a["found"] = a["CA"].notna() | a["CA N-1"].notna()
    a[num] = a[num].fillna(0.0)
    a["n_conv"] = a["n_conv"].fillna(0)
    a["label"] = a["label"].fillna(a["label_promo"])
    for k, _ in LEVELS:
        a[k] = a[k].fillna(NOT_FOUND)

    # ① Avant la promo : prévision (Qté) vs base
    use_recent = params.get("base_mode") == "recent" and base is not None
    if base is not None:
        a = a.join(base)
        a["base_qty"] = a["base_qty"].fillna(0.0)
    else:
        a["base_qty"], a["base_ca"], a["ref_promo_share"] = np.nan, np.nan, np.nan
    a["ref_qty"] = a["base_qty"] if use_recent else a["Qté Vente N-1"]
    a["mult"] = div(a["prev_qty"], a["ref_qty"])
    a["mult_n1"] = div(a["prev_qty"], a["Qté Vente N-1"])
    a["mult_recent"] = div(a["prev_qty"], a["base_qty"])
    a["promo_n1"] = a["CA Promo N-1"] > 0
    a["control"] = np.select([a["ref_qty"] <= 0, a["mult"] > params["justify"], a["mult"] > params["watch"], a["mult"] < 1],
                             ["Sans historique", "À justifier", "À surveiller", "Sous la base"], "Cohérente")
    a["control_rank"] = a["control"].map(CONTROL_ORDER.index)

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
    tot["base_mode"] = "recent" if use_recent else "n1"
    controls = {
        "not_found": a.loc[~a["found"], "code"].tolist(),
        "duplicates": promo.loc[promo.duplicated("code", keep=False), "code"].unique().tolist(),
        "no_price": promo.loc[promo["prev_ca"].isna(), "code"].tolist(),
        "no_pa": promo.loc[promo["pa"].isna(), "code"].tolist(),
        "neg_marge": promo.loc[promo["prev_marge"] < 0, "code"].tolist(),
        "converted": sorted(set(sales["site"]) & converted),
    }
    return a, tot, controls, sites_table(s, converted)


def sites_table(s: pd.DataFrame, converted: set) -> pd.DataFrame:
    g = s.groupby("site")[["CA", "Marge", "Qté Vente", "CA N-1", "Marge N-1"]].sum().reset_index()
    g["rate"] = div(g["Marge"], g["CA"])
    g["rate_n1"] = div(g["Marge N-1"], g["CA N-1"])
    g["ca_evo"] = div(g["CA"] - g["CA N-1"], g["CA N-1"])
    g["share"] = div(g["CA"], pd.Series(g["CA"].sum(), index=g.index))
    g["share_n1"] = div(g["CA N-1"], pd.Series(g["CA N-1"].sum(), index=g.index))
    g["n1_comparable"] = ~g["site"].isin(converted)
    return g


def summarize(d: pd.DataFrame) -> dict:
    """Agrégats d'un ensemble d'articles (total, ou un nœud de lecture)."""
    has_ref = d["ref_qty"] > 0
    t = {
        "n": len(d), "prev_qty": d["prev_qty"].sum(), "qty": d["Qté Vente"].sum(),
        "prev_ca": d["prev_ca"].sum(), "ca": d["CA"].sum(), "ca_n1": d["CA N-1"].sum(), "qty_n1": d["Qté Vente N-1"].sum(),
        "prev_marge": d["prev_marge"].sum(min_count=1), "marge": d["Marge"].sum(), "marge_n1": d["Marge N-1"].sum(),
        "short": d["short"].sum(), "over": int(d["over"].sum()),
        "ref_qty": d["ref_qty"].sum(), "prev_qty_ref": d.loc[has_ref, "prev_qty"].sum(), "ref_qty_ref": d.loc[has_ref, "ref_qty"].sum(),
        "justify": int((d["control"] == "À justifier").sum()), "watch": int((d["control"] == "À surveiller").sum()),
        "below": int((d["control"] == "Sous la base").sum()), "nohist": int((d["control"] == "Sans historique").sum()),
        "justify_ca": d.loc[d["control"] == "À justifier", "prev_ca"].sum(),
    }
    t["mult"] = div(t["prev_qty_ref"], t["ref_qty_ref"])
    t["real_qty"] = div(t["qty"], t["prev_qty"])
    t["real_ca"] = div(t["ca"], t["prev_ca"])
    t["real_marge"] = div(t["marge"], t["prev_marge"])
    t["bias"] = div(t["over"], t["n"])
    t["rate"] = div(t["marge"], t["ca"])
    t["prev_rate"] = div(t["prev_marge"], t["prev_ca"])
    t["rate_n1"] = div(t["marge_n1"], t["ca_n1"])
    t["ca_evo"] = div(t["ca"] - t["ca_n1"], t["ca_n1"])
    t["marge_evo"] = div(t["marge"] - t["marge_n1"], t["marge_n1"])
    return t


def sentence_before(t) -> str:
    base = "la base récente" if t.get("base_mode") == "recent" else "le N-1 de la même période"
    return (f"{t['n']} articles : {fmt_int(t['prev_qty'])} unités prévues, soit {fmt_ratio(t['mult'])} {base}. "
            f"{t['justify']} prévision(s) à justifier, {t['watch']} à surveiller, {t['below']} sous la base, {t['nohist']} sans historique. "
            f"Les prévisions à justifier portent {fmt_fcfa(t['justify_ca'])} de CA prévu ({fmt_pct(div(t['justify_ca'], t['prev_ca']))}).")


def sentence_after(t) -> str:
    s = (f"CA réalisé {fmt_fcfa(t['ca'])} pour {fmt_fcfa(t['prev_ca'])} prévus ({fmt_pct(t['real_ca'])}). "
         f"CA promis non réalisé : {fmt_fcfa(t['short'])} FCFA. Biais : {t['over']} article(s) sur {t['n']} surestimés ou sans vente "
         f"({fmt_pct(t['bias'])}).")
    if t["ca_n1"] > 0:
        s += f" Vs N-1 : CA {fmt_pct(t['ca_evo'], 0, True)}, marge {fmt_pct(t['marge_evo'], 0, True)}."
    return s


# ----------------------------------------------------------------------------
# Excel réunion achats (valeurs uniquement, aucune formule)
# ----------------------------------------------------------------------------
def build_excel(a, tot, controls, sites, filters, meta, params) -> bytes:
    wb = Workbook()
    hdr_fill, hdr_font = PatternFill("solid", fgColor=NAVY), Font(bold=True, color="FFFFFF")
    edit_fill = PatternFill("solid", fgColor="FFF8E6")
    status_fill = {k: PatternFill("solid", fgColor=FILL[v]) for d in (RESULT, CONTROL, VS_N1) for k, v in d.items()}
    before = tot["phase"] == "avant"
    per = f"{filters['period_from']:%d/%m/%Y} au {filters['period_to']:%d/%m/%Y}" if filters["period_from"] else "non lue"

    def table(ws, df, row=1, pct=(), num=(), ratio=(), status=(), edit=(), widths=None):
        for j, c in enumerate(df.columns, 1):
            cell = ws.cell(row=row, column=j, value=c)
            cell.fill, cell.font = hdr_fill, hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 32
        for i, rec in enumerate(df.itertuples(index=False), row + 1):
            for j, v in enumerate(rec, 1):
                if isinstance(v, (float, np.floating)) and (np.isnan(v) or np.isinf(v)):
                    v = None
                cell = ws.cell(row=i, column=j, value=v)
                name = df.columns[j - 1]
                if name in pct:
                    cell.number_format = "0%"
                elif name in num:
                    cell.number_format = "#,##0"
                elif name in ratio:
                    cell.number_format = '"× "0.0'
                if name in status and v in status_fill:
                    cell.fill, cell.font = status_fill[v], Font(bold=True)
                if name in edit:
                    cell.fill = edit_fill
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
        for j, c in enumerate(df.columns, 1):
            w = (widths or {}).get(c) or max(10, min(34, len(str(c)) + 3))
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.freeze_panes = ws.cell(row=row + 1, column=1)
        ws.auto_filter.ref = f"A{row}:{get_column_letter(len(df.columns))}{row + len(df)}"

    def head(ws, title, sub):
        ws["A1"] = title
        ws["A1"].font = Font(bold=True, size=15, color=NAVY)
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, color="6B7090")

    sub = (f"{'Réunion avant promo' if before else 'Réunion résultat promo'} · période {per} · fournisseur {filters['supplier'] or 'non lu'} · "
           f"généré le {datetime.now():%d/%m/%Y %H:%M}")
    hier = {"ray": "Rayon", "fam": "Famille", "sf": "Sous-famille"}
    has_sf = (a["sf"] != NO_SF).any()

    # --- Synthèse ---
    ws = wb.active
    ws.title = "Synthèse"
    head(ws, "Forecast Promo · " + ("contrôle avant promo" if before else "résultat promo"), sub)
    ws["A4"] = sentence_before(tot) if before else sentence_after(tot)
    ws["A4"].font = Font(bold=True)
    ws["A4"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A4:D4")
    ws.row_dimensions[4].height = 60
    if before:
        kpis = [("Articles", tot["n"], "0"), ("Base de comparaison", BASE_LABEL[tot["base_mode"]], None),
                ("Quantité prévue", tot["prev_qty"], "#,##0"), ("Base (quantité sans promo)", tot["ref_qty"], "#,##0"),
                ("Multiplicateur global", tot["mult"], '"× "0.00'), ("CA prévu HT", tot["prev_ca"], "#,##0"),
                ("Marge prévue", tot["prev_marge"], "#,##0"), ("Taux de marge prévu", tot["prev_rate"], "0.0%"),
                ("CA N-1 même période", tot["ca_n1"], "#,##0"), ("Taux de marge N-1", tot["rate_n1"], "0.0%"),
                ("À justifier", tot["justify"], "0"), ("CA prévu porté par les « à justifier »", tot["justify_ca"], "#,##0"),
                ("À surveiller", tot["watch"], "0"), ("Sous la base", tot["below"], "0"), ("Sans historique", tot["nohist"], "0")]
    else:
        kpis = [("Articles", tot["n"], "0"), ("CA prévu HT (forecast)", tot["prev_ca"], "#,##0"), ("CA réalisé HT", tot["ca"], "#,##0"),
                ("Réalisation CA", tot["real_ca"], "0%"), ("CA promis non réalisé", tot["short"], "#,##0"),
                ("Quantité prévue", tot["prev_qty"], "#,##0"), ("Quantité vendue", tot["qty"], "#,##0"),
                ("CA N-1", tot["ca_n1"], "#,##0"), ("Évolution CA vs N-1", tot["ca_evo"], "+0%;-0%"),
                ("Marge réalisée", tot["marge"], "#,##0"), ("Taux de marge", tot["rate"], "0.0%"), ("Taux de marge prévu", tot["prev_rate"], "0.0%"),
                ("Marge N-1", tot["marge_n1"], "#,##0"), ("Taux de marge N-1", tot["rate_n1"], "0.0%"),
                ("Évolution marge vs N-1", tot["marge_evo"], "+0%;-0%"), ("Biais (surestimés ou sans vente)", tot["bias"], "0%")]
    for j, h in enumerate(["Indicateur clé", "Valeur"], 1):
        c = ws.cell(row=6, column=j, value=h)
        c.fill, c.font = hdr_fill, hdr_font
    r = 6
    for lab, v, nf in kpis:
        r += 1
        ws.cell(row=r, column=1, value=lab)
        c = ws.cell(row=r, column=2, value=None if (not isinstance(v, str) and _nan(v)) else v)
        c.font = Font(bold=True)
        if nf:
            c.number_format = nf
    r += 2
    counts = a["control"].value_counts() if before else a["result"].value_counts()
    order = CONTROL_ORDER if before else list(RESULT)
    for j, h in enumerate(["Statut", "Articles"], 1):
        c = ws.cell(row=r, column=j, value=h)
        c.fill, c.font = hdr_fill, hdr_font
    for k in order:
        r += 1
        ws.cell(row=r, column=1, value=k).fill = status_fill[k]
        ws.cell(row=r, column=2, value=int(counts.get(k, 0)))
    r += 2
    ws.cell(row=r, column=1, value="Points de contrôle").font = Font(bold=True, color=NAVY)
    for lab, val in [("Codes introuvables", ", ".join(controls["not_found"]) or "aucun"), ("Doublons", ", ".join(controls["duplicates"]) or "aucun"),
                     ("Prix promo manquant", ", ".join(controls["no_price"]) or "aucun"), ("PA manquant", ", ".join(controls["no_pa"]) or "aucun"),
                     ("Marge prévue négative", ", ".join(controls["neg_marge"]) or "aucune"),
                     ("Sites au N-1 non comparable", ", ".join(controls["converted"]) or "aucun")]:
        r += 1
        ws.cell(row=r, column=1, value=lab)
        ws.cell(row=r, column=2, value=val)
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 42, 34

    # --- Articles ---
    ws = wb.create_sheet("Articles")
    lv = {k: hier[k] for k in (["ray", "fam", "sf"] if has_sf else ["ray", "fam"])}
    if before:
        d = a.sort_values(["control_rank", "mult"], ascending=[True, False], na_position="last").copy()
        d.insert(0, "rank", range(1, len(d) + 1))
        d["promo_n1_txt"] = d["promo_n1"].map({True: "Oui", False: ""})
        d["just"] = ""
        cols = {"rank": "Rang"} | lv | {"code": "Code", "label": "Article", "prev_qty": "Qté prévue", "ref_qty": "Base (Qté)",
                                         "mult": "Multiplicateur", "Qté Vente N-1": "Qté N-1", "promo_n1_txt": "En promo N-1",
                                         "price": "Prix promo", "pa": "PA", "prev_ca": "CA prévu HT", "prev_marge": "Marge prévue",
                                         "prev_rate": "Taux prévu", "control": "Statut", "just": "Justification acheteur"}
        table(ws, d[list(cols)].rename(columns=cols), row=4, pct=("Taux prévu",),
              num=("Qté prévue", "Base (Qté)", "Qté N-1", "Prix promo", "PA", "CA prévu HT", "Marge prévue"), ratio=("Multiplicateur",),
              status=("Statut",), edit=("Justification acheteur",), widths={"Article": 34, "Justification acheteur": 44, "Rang": 7, "Code": 11})
        head(ws, "Articles · du plus à challenger au moins", f"Base : {BASE_LABEL[tot['base_mode']]} · colonne jaune à remplir en réunion")
    else:
        d = a.sort_values(["short", "CA"], ascending=[False, False]).copy()
        d.insert(0, "rank", range(1, len(d) + 1))
        d["expl"] = ""
        cols = {"rank": "Rang"} | lv | {"code": "Code", "label": "Article", "prev_qty": "Qté prévue", "Qté Vente": "Qté vendue",
                                         "real_qty": "Réal. Qté", "prev_ca": "CA prévu HT", "CA": "CA réalisé", "real_ca": "Réal. CA",
                                         "short": "CA promis non réalisé", "CA N-1": "CA N-1", "ca_evo": "Évo vs N-1", "Marge": "Marge",
                                         "rate": "Taux de marge", "result": "Statut", "expl": "Explication acheteur"}
        table(ws, d[list(cols)].rename(columns=cols), row=4, pct=("Réal. Qté", "Réal. CA", "Évo vs N-1", "Taux de marge"),
              num=("Qté prévue", "Qté vendue", "CA prévu HT", "CA réalisé", "CA promis non réalisé", "CA N-1", "Marge"),
              status=("Statut",), edit=("Explication acheteur",), widths={"Article": 34, "Explication acheteur": 44, "Rang": 7, "Code": 11})
        head(ws, "Articles · classés par CA promis non réalisé", "Colonne jaune à remplir en réunion")

    # --- Magasins ---
    ws = wb.create_sheet("Magasins")
    m = sites.copy()
    m["comp"] = m["n1_comparable"].map({True: "", False: "Non (site converti)"})
    if before:
        m = m.sort_values("CA N-1", ascending=False)
        m.insert(0, "rank", range(1, len(m) + 1))
        cols = {"rank": "Rang", "site": "Magasin", "CA N-1": "CA N-1", "Marge N-1": "Marge N-1", "rate_n1": "Taux N-1",
                "share_n1": "Poids N-1", "comp": "N-1 comparable"}
        table(ws, m[list(cols)].rename(columns=cols), row=4, pct=("Taux N-1", "Poids N-1"), num=("CA N-1", "Marge N-1"),
              widths={"Magasin": 26, "N-1 comparable": 22})
        head(ws, "Magasins · poids des articles de la promo l'an dernier", "Pour anticiper implantation et quantités par site")
    else:
        m = m.sort_values("ca_evo", ascending=True, na_position="last")
        m.insert(0, "rank", range(1, len(m) + 1))
        cols = {"rank": "Rang", "site": "Magasin", "CA": "CA réalisé", "CA N-1": "CA N-1", "ca_evo": "Évo vs N-1", "Marge": "Marge",
                "rate": "Taux de marge", "rate_n1": "Taux N-1", "share": "Poids", "comp": "N-1 comparable"}
        table(ws, m[list(cols)].rename(columns=cols), row=4, pct=("Évo vs N-1", "Taux de marge", "Taux N-1", "Poids"),
              num=("CA réalisé", "CA N-1", "Marge"), widths={"Magasin": 26, "N-1 comparable": 22})
        head(ws, "Magasins · du plus fort recul vs N-1 au meilleur", f"Réseau : {fmt_pct(tot['ca_evo'], 0, True)} vs N-1")

    # --- Lexique ---
    ws = wb.create_sheet("Lexique")
    table(ws, pd.DataFrame(LEXIQUE, columns=["Terme", "Définition"]), row=1, widths={"Terme": 34, "Définition": 110})
    ws.auto_filter.ref = None

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
