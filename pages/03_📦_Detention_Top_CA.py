"""
03_📦_Detention_Top_CA.py — SmartBuyer Hub
Taux de détention Top CA · GOLD / SILVER · Articles Permanents
Source : Export PBI stock pivot (article × site) + Liste Top CA CSV
v3.0 — Migration PBI · Alertes ruptures/réappro · Scorecard GOLD/SILVER/Total
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Détention Top CA · SmartBuyer",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display",
                 "SF Pro Text", "Helvetica Neue", Arial, sans-serif !important;
    background-color: #F2F2F7;
}
.stApp { background: #F2F2F7; }
.main .block-container { padding-top: 1.8rem; max-width: 1300px; }
[data-testid="stSidebar"] { background: #FFFFFF !important; border-right: 0.5px solid #E5E5EA !important; }
[data-testid="stMetric"] { background: #FFFFFF !important; border: 0.5px solid #E5E5EA !important; border-radius: 12px !important; padding: 16px 18px !important; }
[data-testid="stMetricLabel"] { font-size: 11px !important; font-weight: 500 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 600 !important; color: #1C1C1E !important; }
[data-testid="stTabs"] button[role="tab"] { font-size: 13px !important; font-weight: 500 !important; padding: 8px 16px !important; color: #8E8E93 !important; border-bottom: 2px solid transparent !important; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color: #007AFF !important; border-bottom: 2px solid #007AFF !important; background: transparent !important; }
[data-testid="stTabs"] [role="tablist"] { border-bottom: 0.5px solid #E5E5EA !important; }
[data-testid="stDataFrame"] { border: 0.5px solid #E5E5EA !important; border-radius: 10px !important; }
[data-testid="stDataFrame"] th { background: #F2F2F7 !important; font-size: 11px !important; font-weight: 600 !important; color: #8E8E93 !important; text-transform: uppercase !important; }
[data-testid="stFileUploader"] { border: 1.5px dashed #D1D1D6 !important; border-radius: 10px !important; background: #F9F9FB !important; }
.stDownloadButton > button { background: #007AFF !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; font-size: 13px !important; padding: 10px 24px !important; width: 100% !important; }
hr { border-color: #E5E5EA !important; margin: 1rem 0 !important; }
.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 1.5rem; }
.section-label { font-size: 11px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 10px; }
.alert-card  { padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; font-size: 13px; line-height: 1.5; border-left: 3px solid; }
.alert-red   { background: #FFF2F2; border-color: #FF3B30; color: #3A0000; }
.alert-amber { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }
.alert-green { background: #F0FFF4; border-color: #34C759; color: #003A10; }
.alert-gray  { background: #F9F9FB; border-color: #C7C7CC; color: #3A3A3C; }
.col-required { background: #F0F8FF; border: 0.5px solid #B3D9FF; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; display: flex; align-items: flex-start; gap: 10px; }
.col-name { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
.sc-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 10px; margin-bottom: 16px; }
.sc-card { background: #FFFFFF; border: 0.5px solid #E5E5EA; border-radius: 12px; padding: 14px 16px; position: relative; }
.sc-card.ok   { border-color: #6EE7B7; background: #F0FFF4; }
.sc-card.warn { border-color: #FCD34D; background: #FFFBF0; }
.sc-card.ko   { border-color: #FECACA; background: #FFF2F2; }
.sc-dot { width: 8px; height: 8px; border-radius: 50%; position: absolute; top: 14px; right: 14px; }
.sc-name { font-size: 11px; font-weight: 600; color: #1C1C1E; margin-bottom: 10px; }
.sc-seg { display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px; font-size: 11px; }
.sc-seg-label { color: #8E8E93; display: flex; align-items: center; gap: 4px; }
.sc-seg-dot { width: 6px; height: 6px; border-radius: 50%; }
.sc-seg-right { display: flex; align-items: center; gap: 6px; }
.sc-seg-pct { font-weight: 600; }
.sc-seg-count { font-size: 10px; color: #8E8E93; }
.bar-track { height: 3px; background: #E5E5EA; border-radius: 2px; margin-bottom: 6px; }
.bar-fill { height: 3px; border-radius: 2px; }
.sc-divider { height: 0.5px; background: #E5E5EA; margin: 8px 0; }
.sc-total-row { display: flex; align-items: center; justify-content: space-between; }
.sc-total-label { font-size: 10px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: .04em; }
.sc-total-pct { font-size: 20px; font-weight: 700; }
.green-txt { color: #34C759; } .amber-txt { color: #FF9500; } .red-txt { color: #FF3B30; }
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def norm_code(s):
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(8)

def color_taux(v, cible):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "#8E8E93"
    return "#34C759" if v >= cible else "#FF9500" if v >= cible - 10 else "#FF3B30"

def cls_taux(v, cible):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "warn"
    return "ok" if v >= cible else "warn" if v >= cible - 10 else "ko"

def fmt_pct(v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "—"
    return f"{v:.1f}%"

def short_site(s):
    return str(s).split(" - ", 1)[-1].strip() if " - " in str(s) else str(s)

def short_code(s):
    return str(s).split(" - ", 1)[0].strip() if " - " in str(s) else str(s)

# ─── PARSERS ──────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_topca(byt, fname):
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            df = pd.read_csv(BytesIO(byt), sep=None, engine="python",
                             encoding=enc, dtype=str)
            df.columns = df.columns.str.strip().str.upper()
            code_col = next((c for c in df.columns if "CODE" in c or "ARTICLE" in c), df.columns[0])
            type_col = next((c for c in df.columns if "TYPE" in c), None)
            lib_col  = next((c for c in df.columns if "LIB" in c), None)
            out = pd.DataFrame()
            out["code"] = norm_code(df[code_col])
            out["type"] = df[type_col].str.strip().str.upper() if type_col else "GOLD"
            out["lib"]  = df[lib_col].astype(str).str.strip() if lib_col else ""
            out = out.dropna(subset=["code"])
            out = out[out["code"].str.match(r"^\d{8}$")]
            if len(out) > 0:
                return out
        except Exception:
            continue
    return pd.DataFrame(columns=["code", "type", "lib"])


@st.cache_data(show_spinner=False)
def load_pbi(byt, fname):
    """
    Parse l'export PBI pivot : Site nom court | 10202 - Palmeraie | ... | Total
    Ligne 0 = sous-entête 'Article','Stock','Stock'... à ignorer
    """
    df_raw = pd.read_excel(BytesIO(byt), dtype=str)

    # Colonnes sites (tout sauf 'Site nom court' et 'Total')
    site_cols = [c for c in df_raw.columns if c not in ["Site nom court", "Total"]]

    # Filtrer les lignes articles valides
    mask = (
        df_raw["Site nom court"].notna() &
        (~df_raw["Site nom court"].astype(str).str.strip().isin(
            ["Article", "Total", "nan", ""])) &
        (~df_raw["Site nom court"].astype(str).str.startswith("Filtres"))
    )
    df = df_raw[mask].copy().reset_index(drop=True)

    # Extraire code et libellé
    df["code"]    = df["Site nom court"].apply(short_code)
    df["code"]    = df["code"].str.strip().str.zfill(8)
    df["libelle"] = df["Site nom court"].apply(
        lambda s: str(s).split(" - ", 1)[-1].strip() if " - " in str(s) else str(s)
    )

    # Convertir stocks en numérique
    for col in site_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Construire la liste des sites avec leur libellé court
    sites_info = {col: short_site(col) for col in site_cols}

    return df, site_cols, sites_info


def compute(df_pbi, site_cols, sites_info, df_topca, type_filtre, cible, seuil_reappro):
    """
    Calcule pour chaque article Top CA × site :
    - stock
    - détenu (stock > 0)
    - réappro urgent (0 < stock < seuil)
    - alerte
    """
    # Filtrer Top CA par type
    if type_filtre != "Tous":
        topca_f = df_topca[df_topca["type"] == type_filtre].copy()
    else:
        topca_f = df_topca.copy()

    top_codes = set(topca_f["code"].unique())
    top_meta  = topca_f.drop_duplicates("code").set_index("code")[["type", "lib"]]

    # Filtrer PBI sur les codes Top CA
    df = df_pbi[df_pbi["code"].isin(top_codes)].copy()

    # Melt : article × site × stock
    df_long = df.melt(
        id_vars=["code", "libelle"],
        value_vars=site_cols,
        var_name="site_col",
        value_name="stock"
    )
    df_long["site"]   = df_long["site_col"].map(sites_info)
    df_long["type"]   = df_long["code"].map(lambda c: top_meta.get(c, {}).get("type", "?")
                                            if c in top_meta.index else "?")
    df_long["lib_topca"] = df_long["code"].map(lambda c: top_meta.get(c, {}).get("lib", "")
                                               if c in top_meta.index else "")

    # Calculs
    df_long["detenu"]        = df_long["stock"].notna() & (df_long["stock"] > 0)
    df_long["reappro_urgent"] = (
        df_long["stock"].notna() &
        (df_long["stock"] > 0) &
        (df_long["stock"] < seuil_reappro)
    )

    # Absents PBI
    absents = topca_f[~topca_f["code"].isin(df_pbi["code"].unique())].copy()

    return df_long, absents, top_codes, top_meta


def build_alertes(df_long, top_codes, top_meta, site_cols, seuil_reappro):
    """Construit les 4 DataFrames d'alertes."""
    n_sites = len(site_cols)

    # Agréger par article
    agg = df_long.groupby(["code", "lib_topca", "type"]).agg(
        nb_sites_detenu=("detenu", "sum"),
        nb_sites_total=("detenu", "count"),
        stock_total=("stock", lambda x: x.fillna(0).sum()),
    ).reset_index()
    agg["nb_sites_detenu"] = agg["nb_sites_detenu"].astype(int)

    # 1. Ruptures nettes
    ruptures_nettes = agg[agg["nb_sites_detenu"] == 0].copy()
    ruptures_nettes = ruptures_nettes.sort_values(
        ["type", "lib_topca"], ascending=[True, True]
    ).reset_index(drop=True)
    ruptures_nettes["Rang"] = range(1, len(ruptures_nettes) + 1)

    # 2. Réappro urgent (par article × site)
    reappro = df_long[df_long["reappro_urgent"]].copy()
    reappro = reappro[["code", "lib_topca", "type", "site", "stock"]].copy()
    reappro = reappro.sort_values(["type", "stock"], ascending=[True, True]).reset_index(drop=True)
    reappro["Rang"] = range(1, len(reappro) + 1)

    # 3. Ruptures partielles (détenu sur < 50% des sites, mais pas rupture nette)
    partielles = agg[
        (agg["nb_sites_detenu"] > 0) &
        (agg["nb_sites_detenu"] / agg["nb_sites_total"] < 0.5)
    ].copy()
    # Ajouter la liste des sites manquants
    def sites_manquants(code):
        rows = df_long[(df_long["code"] == code) & (~df_long["detenu"])]
        return ", ".join(sorted(rows["site"].dropna().unique().tolist()))
    partielles["sites_manquants"] = partielles["code"].apply(sites_manquants)
    partielles = partielles.sort_values(
        ["type", "nb_sites_detenu"], ascending=[True, True]
    ).reset_index(drop=True)
    partielles["Rang"] = range(1, len(partielles) + 1)
    partielles["taux_detenu"] = (
        partielles["nb_sites_detenu"] / partielles["nb_sites_total"] * 100
    ).round(1)

    return ruptures_nettes, reappro, partielles


def compute_taux(df_long, top_meta, site_cols, cible):
    """Taux de détention par site × type."""
    types_dispo = sorted(df_long["type"].unique())
    sites       = [s for s in df_long["site"].unique() if pd.notna(s)]

    rows = []
    for site in sorted(sites):
        s = df_long[df_long["site"] == site]
        row = {"site": site}
        # Par type
        for t in types_dispo:
            st = s[s["type"] == t]
            n  = len(st)
            d  = int(st["detenu"].sum())
            row[f"n_{t}"]    = n
            row[f"det_{t}"]  = d
            row[f"taux_{t}"] = round(d / n * 100, 1) if n > 0 else None
        # Total
        n_tot = len(s)
        d_tot = int(s["detenu"].sum())
        row["n_total"]    = n_tot
        row["det_total"]  = d_tot
        row["taux_total"] = round(d_tot / n_tot * 100, 1) if n_tot > 0 else None
        rows.append(row)

    return pd.DataFrame(rows), types_dispo


# ─── EXPORT EXCEL ─────────────────────────────────────────────────────────────
def gen_excel(taux_df, types_dispo, ruptures_nettes, reappro, partielles,
              absents, type_filtre, cible, seuil_reappro):
    wb  = Workbook()
    HDR = PatternFill("solid", fgColor="1C3557")
    GLD = PatternFill("solid", fgColor="FFFBF0")
    ODD = PatternFill("solid", fgColor="F7F7F7")
    EVN = PatternFill("solid", fgColor="FFFFFF")
    HF  = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    BF  = Font(bold=True, name="Calibri", size=10)
    NF  = Font(name="Calibri", size=10)
    CTR = Alignment(horizontal="center", vertical="center")
    LFT = Alignment(horizontal="left",   vertical="center")

    def header_row(ws, headers, widths):
        for i, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(row=1, column=i, value=h)
            c.fill = HDR; c.font = HF; c.alignment = CTR
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[1].height = 22
        ws.freeze_panes = "A2"

    def write_rows(ws, data_rows, gold_col=None):
        for ri, row in enumerate(data_rows, 2):
            bg = GLD if (gold_col and row[gold_col] == "GOLD") else (ODD if ri % 2 == 0 else EVN)
            fill = PatternFill("solid", fgColor=bg.fgColor)
            for ci, val in enumerate(row.values(), 1):
                c = ws.cell(row=ri, column=ci, value=val)
                c.fill = fill
                c.font = BF if ci == 1 else NF
                c.alignment = CTR if isinstance(val, (int, float)) else LFT
            ws.row_dimensions[ri].height = 18

    # ── Feuille 1 : Synthèse réseau ──────────────────────────────────────────
    ws1 = wb.active; ws1.title = "Synthèse réseau"
    headers1 = ["Magasin", "Réf total", "Détenus", "Taux total %"]
    widths1  = [24, 12, 12, 14]
    for t in types_dispo:
        headers1 += [f"Réf {t}", f"Détenus {t}", f"Taux {t} %"]
        widths1  += [12, 12, 12]
    header_row(ws1, headers1, widths1)
    for ri, (_, row) in enumerate(taux_df.iterrows(), 2):
        bg = ODD if ri % 2 == 0 else EVN
        fill = PatternFill("solid", fgColor=bg.fgColor)
        vals = [row["site"], row["n_total"], row["det_total"],
                row["taux_total"]]
        for t in types_dispo:
            vals += [row.get(f"n_{t}", 0), row.get(f"det_{t}", 0),
                     row.get(f"taux_{t}", None)]
        for ci, val in enumerate(vals, 1):
            c = ws1.cell(row=ri, column=ci, value=val)
            c.fill = fill; c.font = NF
            c.alignment = CTR if isinstance(val, (int, float)) else LFT
            if isinstance(val, float) and ci >= 4:
                c.number_format = "0.0"
        ws1.row_dimensions[ri].height = 18

    # ── Feuille 2 : Ruptures nettes ───────────────────────────────────────────
    ws2 = wb.create_sheet("Ruptures nettes")
    header_row(ws2,
        ["#", "Code article", "Libellé article", "Type",
         "Nb sites (stock=0)", "Nb sites total", "Action recommandée"],
        [5, 14, 40, 10, 18, 14, 28])
    for ri, (_, row) in enumerate(ruptures_nettes.iterrows(), 2):
        bg = GLD if row["type"] == "GOLD" else (ODD if ri % 2 == 0 else EVN)
        fill = PatternFill("solid", fgColor=bg.fgColor)
        vals = [row["Rang"], row["code"], row["lib_topca"], row["type"],
                int(row["nb_sites_total"] - row["nb_sites_detenu"]),
                int(row["nb_sites_total"]),
                "Commander en urgence"]
        for ci, val in enumerate(vals, 1):
            c = ws2.cell(row=ri, column=ci, value=val)
            c.fill = fill; c.font = NF
            c.alignment = CTR if ci in [1, 5, 6] else LFT
        ws2.row_dimensions[ri].height = 18

    # ── Feuille 3 : Réappro urgent ────────────────────────────────────────────
    ws3 = wb.create_sheet("Réappro urgent")
    header_row(ws3,
        ["#", "Code article", "Libellé article", "Type",
         "Magasin", "Stock actuel", f"Seuil réappro (< {seuil_reappro})", "Action recommandée"],
        [5, 14, 40, 10, 22, 14, 20, 26])
    for ri, (_, row) in enumerate(reappro.iterrows(), 2):
        bg = GLD if row["type"] == "GOLD" else (ODD if ri % 2 == 0 else EVN)
        fill = PatternFill("solid", fgColor=bg.fgColor)
        vals = [row["Rang"], row["code"], row["lib_topca"], row["type"],
                row["site"],
                int(row["stock"]) if pd.notna(row["stock"]) else 0,
                f"< {seuil_reappro}",
                "Réassort immédiat"]
        for ci, val in enumerate(vals, 1):
            c = ws3.cell(row=ri, column=ci, value=val)
            c.fill = fill; c.font = NF
            c.alignment = CTR if ci in [1, 6] else LFT
        ws3.row_dimensions[ri].height = 18

    # ── Feuille 4 : Ruptures partielles ───────────────────────────────────────
    ws4 = wb.create_sheet("Ruptures partielles")
    header_row(ws4,
        ["#", "Code article", "Libellé article", "Type",
         "Sites détenus", "Sites total", "Taux %", "Sites manquants"],
        [5, 14, 40, 10, 14, 12, 10, 60])
    for ri, (_, row) in enumerate(partielles.iterrows(), 2):
        bg = GLD if row["type"] == "GOLD" else (ODD if ri % 2 == 0 else EVN)
        fill = PatternFill("solid", fgColor=bg.fgColor)
        vals = [row["Rang"], row["code"], row["lib_topca"], row["type"],
                int(row["nb_sites_detenu"]), int(row["nb_sites_total"]),
                row["taux_detenu"], row["sites_manquants"]]
        for ci, val in enumerate(vals, 1):
            c = ws4.cell(row=ri, column=ci, value=val)
            c.fill = fill; c.font = NF
            c.alignment = CTR if ci in [1, 5, 6, 7] else LFT
            if ci == 7 and isinstance(val, float):
                c.number_format = "0.0"
        ws4.row_dimensions[ri].height = 18

    # ── Feuille 5 : Absents PBI ────────────────────────────────────────────────
    ws5 = wb.create_sheet("Absents PBI")
    header_row(ws5,
        ["Code article", "Libellé", "Type", "Vérification"],
        [14, 40, 10, 40])
    for ri, (_, row) in enumerate(absents.iterrows(), 2):
        bg = GLD if row["type"] == "GOLD" else (ODD if ri % 2 == 0 else EVN)
        fill = PatternFill("solid", fgColor=bg.fgColor)
        vals = [row["code"], row["lib"], row["type"],
                "Vérifier déréférencement ou code article"]
        for ci, val in enumerate(vals, 1):
            c = ws5.cell(row=ri, column=ci, value=val)
            c.fill = fill; c.font = NF
            c.alignment = CTR if ci == 3 else LFT
        ws5.row_dimensions[ri].height = 18

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:18px'>
  <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>🛍️ SmartBuyer</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · Équipe Achats</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Import fichiers</div>", unsafe_allow_html=True)
    st.markdown("**Liste Top CA** *(CSV · CODE ARTICLE · TYPE)*")
    f_topca = st.file_uploader("Top CA", type=["csv"], key="topca",
                                label_visibility="collapsed")
    st.markdown("**Export PBI stock** *(Excel pivot article × site)*")
    f_pbi   = st.file_uploader("PBI", type=["xlsx", "xls"], key="pbi",
                                label_visibility="collapsed")
    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Paramètres</div>", unsafe_allow_html=True)
    cible         = st.slider("Cible taux de détention (%)", 70, 100, 85, 1)
    seuil_reappro = st.number_input("Seuil réappro urgent (stock <)", 1, 50, 3, 1,
                                     help="Alerte si 0 < stock < seuil")


# ─── PAGE ─────────────────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>📦 Détention Top CA</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Articles permanents · GOLD / SILVER · Taux de détention · Ruptures · Réappro urgent</div>", unsafe_allow_html=True)

# ─── ÉCRAN D'ACCUEIL ──────────────────────────────────────────────────────────
if not f_topca or not f_pbi:
    st.markdown("---")
    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>ℹ️ À quoi sert ce module ?</strong><br>
  Vérifie la présence en magasin des articles Top CA et calcule le taux de détention
  par type (GOLD / SILVER) et par site. Détecte les ruptures et anticipe les réapprovisionnements urgents.
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Les 4 alertes</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    alertes_doc = [
        ("🔴", "Rupture nette",      "#FF3B30", "stock = 0 ou NaN sur tous les sites",
         "Commander en urgence"),
        ("🔵", "Réappro urgent",     "#007AFF", "0 < stock < seuil paramétrable",
         "Réassort immédiat avant épuisement"),
        ("🟡", "Rupture partielle",  "#FF9500", "détenu sur < 50% des sites",
         "Redistribuer ou approvisionner les sites manquants"),
        ("⚪", "Absent PBI",         "#8E8E93", "article Top CA absent de l'export PBI",
         "Vérifier déréférencement ou code article"),
    ]
    for i, (ico, titre, color, cond, action) in enumerate(alertes_doc):
        with (c1 if i % 2 == 0 else c2):
            st.markdown(f"""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;
            padding:14px;border-left:3px solid {color};margin-bottom:10px'>
  <div style='font-size:13px;font-weight:600;color:#1C1C1E;margin-bottom:6px'>{ico} {titre}</div>
  <div style='font-size:12px;color:#3A3A3C;margin-bottom:4px'>{cond}</div>
  <div style='font-size:11px;color:#8E8E93;font-style:italic'>→ {action}</div>
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Fichiers attendus</div>", unsafe_allow_html=True)
    cf1, cf2 = st.columns(2)
    with cf1:
        st.markdown("""
<div class='col-required'><div style='font-size:16px'>📋</div>
<div><div class='col-name'>Liste Top CA (.csv)</div>
<div class='col-desc'>Colonnes : CODE ARTICLE · TYPE (GOLD/SILVER/...)</div>
<div class='col-desc' style='color:#8E8E93;font-size:11px;margin-top:2px'>Le TYPE est flexible — toute valeur est acceptée</div>
</div></div>""", unsafe_allow_html=True)
    with cf2:
        st.markdown("""
<div class='col-required'><div style='font-size:16px'>📊</div>
<div><div class='col-name'>Export PBI stock (.xlsx)</div>
<div class='col-desc'>Pivot article × site · colonne Site nom court</div>
<div class='col-desc' style='color:#8E8E93;font-size:11px;margin-top:2px'>Valeur cellule = stock · NaN = absent</div>
</div></div>""", unsafe_allow_html=True)

    st.info("⬆️ Charge les deux fichiers dans la sidebar pour lancer l'analyse.")
    st.stop()

# ─── CHARGEMENT ───────────────────────────────────────────────────────────────
with st.spinner("Lecture des fichiers…"):
    df_topca          = load_topca(f_topca.read(), f_topca.name)
    df_pbi, site_cols, sites_info = load_pbi(f_pbi.read(),  f_pbi.name)

if df_topca.empty: st.error("Liste Top CA vide ou non reconnue."); st.stop()
if df_pbi.empty:   st.error("Export PBI vide ou non reconnu.");    st.stop()

# ─── SÉLECTEUR TYPE ───────────────────────────────────────────────────────────
types_dispo_all = sorted(df_topca["type"].unique())
options = ["Tous"] + types_dispo_all
with st.sidebar:
    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Filtre type</div>", unsafe_allow_html=True)
    type_filtre = st.radio("Type", options, horizontal=False, label_visibility="collapsed")

# ─── CALCUL ───────────────────────────────────────────────────────────────────
with st.spinner("Calcul des taux de détention…"):
    df_long, absents, top_codes, top_meta = compute(
        df_pbi, site_cols, sites_info, df_topca, type_filtre, cible, seuil_reappro)
    ruptures_nettes, reappro, partielles = build_alertes(
        df_long, top_codes, top_meta, site_cols, seuil_reappro)
    taux_df, types_seg = compute_taux(df_long, top_meta, site_cols, cible)

# Métriques globales
n_sites       = len(site_cols)
n_refs        = df_long["code"].nunique()
taux_moy      = taux_df["taux_total"].mean()
n_rupt_nettes = len(ruptures_nettes)
n_reappro     = len(reappro)
n_partielles  = len(partielles)
n_absents     = len(absents)

# KPIs par type pour la barre du haut
taux_par_type = {}
for t in types_seg:
    col = f"taux_{t}"
    if col in taux_df.columns:
        taux_par_type[t] = taux_df[col].mean()

# ─── KPIs ─────────────────────────────────────────────────────────────────────
st.markdown(f"<div class='section-label'>{n_sites} sites · {n_refs} références · {type_filtre}</div>", unsafe_allow_html=True)

kpi_cols = st.columns(2 + len(types_seg))
kpi_cols[0].metric("Réf analysées", str(n_refs),
                    f"{n_absents} absents PBI" if n_absents else "")
kpi_cols[1].metric("Taux réseau", fmt_pct(taux_moy),
                    f"{(taux_moy or 0) - cible:+.1f} pt vs cible" if taux_moy else "")
for i, t in enumerate(types_seg):
    tv = taux_par_type.get(t)
    kpi_cols[2 + i].metric(f"Taux {t}", fmt_pct(tv),
                            f"{(tv or 0) - cible:+.1f} pt" if tv else "")

# ─── ALERTES RÉSUMÉ ───────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Alertes</div>", unsafe_allow_html=True)

if n_rupt_nettes > 0:
    nb_gold_r = int((ruptures_nettes["type"] == "GOLD").sum())
    st.markdown(f"""
<div class='alert-card alert-red'>
  <strong>🔴 {n_rupt_nettes} rupture(s) nette(s)</strong> — stock = 0 sur tous les sites
  · {nb_gold_r} GOLD · {n_rupt_nettes - nb_gold_r} autres<br>
  <span style='font-size:12px;opacity:.85'>→ Commander en urgence.</span>
</div>""", unsafe_allow_html=True)

if n_reappro > 0:
    nb_gold_ra = int((reappro["type"] == "GOLD").sum())
    st.markdown(f"""
<div class='alert-card alert-blue'>
  <strong>🔵 {n_reappro} ligne(s) en réappro urgent</strong> — 0 &lt; stock &lt; {seuil_reappro} unités
  · {nb_gold_ra} GOLD<br>
  <span style='font-size:12px;opacity:.85'>→ Réassort immédiat avant épuisement.</span>
</div>""", unsafe_allow_html=True)

if n_partielles > 0:
    nb_gold_p = int((partielles["type"] == "GOLD").sum())
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>🟡 {n_partielles} rupture(s) partielle(s)</strong> — détenu sur &lt; 50% des sites
  · {nb_gold_p} GOLD<br>
  <span style='font-size:12px;opacity:.85'>→ Redistribuer ou approvisionner les sites manquants.</span>
</div>""", unsafe_allow_html=True)

if n_absents > 0:
    st.markdown(f"""
<div class='alert-card alert-gray'>
  <strong>⚪ {n_absents} article(s) Top CA absents du PBI</strong><br>
  <span style='font-size:12px;opacity:.85'>→ Vérifier déréférencement ou code article.</span>
</div>""", unsafe_allow_html=True)

if n_rupt_nettes == 0 and n_reappro == 0 and n_partielles == 0:
    st.markdown("<div class='alert-card alert-green'>✅ Aucune alerte — tous les sites au-dessus des seuils.</div>", unsafe_allow_html=True)

# ─── SCORECARD MAGASINS ───────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Scorecard magasins — taux de détention par type</div>", unsafe_allow_html=True)

# Couleurs par type
TYPE_COLORS = {
    "GOLD":   ("#FF9500", "#FF9500"),
    "SILVER": ("#8E8E93", "#636366"),
}
def type_color(t):
    return TYPE_COLORS.get(t, ("#007AFF", "#007AFF"))

sc_html = '<div class="sc-grid">'
for _, row in taux_df.sort_values("taux_total").iterrows():
    tv    = row["taux_total"]
    cls   = cls_taux(tv, cible)
    dot_c = "#34C759" if cls == "ok" else "#FF9500" if cls == "warn" else "#FF3B30"
    txt_c = "green-txt" if cls == "ok" else "amber-txt" if cls == "warn" else "red-txt"

    segs_html = ""
    for t in types_seg:
        tc, _ = type_color(t)
        tv_t  = row.get(f"taux_{t}")
        det_t = row.get(f"det_{t}", 0)
        n_t   = row.get(f"n_{t}", 0)
        pct_t = tv_t if tv_t and not np.isnan(tv_t) else 0
        tc_txt = "green-txt" if pct_t >= cible else "amber-txt" if pct_t >= cible - 10 else "red-txt"
        segs_html += f"""
<div class='sc-seg'>
  <div class='sc-seg-label'>
    <div class='sc-seg-dot' style='background:{tc}'></div>{t}
  </div>
  <div class='sc-seg-right'>
    <span class='sc-seg-pct {tc_txt}'>{fmt_pct(tv_t)}</span>
    <span class='sc-seg-count'>{int(det_t)}/{int(n_t)}</span>
  </div>
</div>
<div class='bar-track'><div class='bar-fill' style='width:{pct_t:.0f}%;background:{tc}'></div></div>"""

    sc_html += f"""
<div class='sc-card {cls}'>
  <div class='sc-dot' style='background:{dot_c}'></div>
  <div class='sc-name'>{row['site']}</div>
  {segs_html}
  <div class='sc-divider'></div>
  <div class='sc-total-row'>
    <span class='sc-total-label'>Total</span>
    <span class='sc-total-pct {txt_c}'>{fmt_pct(tv)}</span>
  </div>
  <div class='bar-track'><div class='bar-fill' style='width:{tv if tv else 0:.0f}%;background:{dot_c}'></div></div>
</div>"""

sc_html += "</div>"
st.markdown(sc_html, unsafe_allow_html=True)

# ─── TABS ─────────────────────────────────────────────────────────────────────
st.markdown("---")
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Synthèse réseau",
    "🚨 Alertes détaillées",
    "🔍 Détail article × site",
    "🚫 Absents PBI",
    "📥 Export Excel",
])

# ══ TAB 1 — SYNTHÈSE ══════════════════════════════════════════════════════════
with tab1:
    st.markdown("<div class='section-label'>Taux de détention par magasin et par type</div>", unsafe_allow_html=True)
    disp1 = taux_df[["site", "n_total", "det_total", "taux_total"]].copy()
    disp1.columns = ["Magasin", "Réf", "Détenus", "Taux total %"]
    for t in types_seg:
        disp1[f"Réf {t}"]    = taux_df[f"n_{t}"]
        disp1[f"Détenus {t}"] = taux_df[f"det_{t}"]
        disp1[f"Taux {t} %"] = taux_df[f"taux_{t}"].apply(fmt_pct)
    disp1["Taux total %"] = disp1["Taux total %"].apply(fmt_pct)
    disp1["Statut"] = taux_df["taux_total"].apply(
        lambda v: "🟢 OK" if v and v >= cible else "🟡 Surveiller" if v and v >= cible - 10 else "🔴 Action")
    st.dataframe(disp1.sort_values("Taux total %"), use_container_width=True, hide_index=True)

# ══ TAB 2 — ALERTES DÉTAILLÉES ════════════════════════════════════════════════
with tab2:
    al_sel = st.multiselect("Afficher",
        ["🔴 Ruptures nettes", "🔵 Réappro urgent",
         "🟡 Ruptures partielles", "⚪ Absents PBI"],
        default=["🔴 Ruptures nettes", "🔵 Réappro urgent",
                 "🟡 Ruptures partielles", "⚪ Absents PBI"])

    if "🔴 Ruptures nettes" in al_sel and not ruptures_nettes.empty:
        st.markdown(f"<div class='section-label'>🔴 Ruptures nettes — {len(ruptures_nettes)} article(s)</div>", unsafe_allow_html=True)
        d = ruptures_nettes[["Rang","code","lib_topca","type","nb_sites_detenu","nb_sites_total"]].copy()
        d.columns = ["#","Code","Libellé","Type","Sites détenus","Sites total"]
        st.dataframe(d, use_container_width=True, hide_index=True,
            column_config={"#": st.column_config.NumberColumn("#", width=50)})

    if "🔵 Réappro urgent" in al_sel and not reappro.empty:
        st.markdown(f"<div class='section-label'>🔵 Réappro urgent — {len(reappro)} ligne(s)</div>", unsafe_allow_html=True)
        d = reappro[["Rang","code","lib_topca","type","site","stock"]].copy()
        d.columns = ["#","Code","Libellé","Type","Magasin","Stock actuel"]
        d["Stock actuel"] = d["Stock actuel"].apply(lambda x: int(x) if pd.notna(x) else 0)
        st.dataframe(d, use_container_width=True, hide_index=True,
            column_config={"#": st.column_config.NumberColumn("#", width=50)})

    if "🟡 Ruptures partielles" in al_sel and not partielles.empty:
        st.markdown(f"<div class='section-label'>🟡 Ruptures partielles — {len(partielles)} article(s)</div>", unsafe_allow_html=True)
        d = partielles[["Rang","code","lib_topca","type","nb_sites_detenu","nb_sites_total","taux_detenu","sites_manquants"]].copy()
        d.columns = ["#","Code","Libellé","Type","Sites détenus","Sites total","Taux %","Sites manquants"]
        d["Taux %"] = d["Taux %"].apply(fmt_pct)
        st.dataframe(d, use_container_width=True, hide_index=True,
            column_config={
                "#": st.column_config.NumberColumn("#", width=50),
                "Sites manquants": st.column_config.TextColumn("Sites manquants", width="large"),
            })

    if "⚪ Absents PBI" in al_sel and not absents.empty:
        st.markdown(f"<div class='section-label'>⚪ Absents PBI — {len(absents)} article(s)</div>", unsafe_allow_html=True)
        d = absents[["code","lib","type"]].copy()
        d.columns = ["Code","Libellé","Type"]
        st.dataframe(d, use_container_width=True, hide_index=True)

# ══ TAB 3 — DÉTAIL ARTICLE × SITE ════════════════════════════════════════════
with tab3:
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        sel_type_det = st.selectbox("Type", ["Tous"] + types_seg, key="det_type")
    with fc2:
        sel_site_det = st.selectbox("Magasin", ["Tous"] + sorted(df_long["site"].dropna().unique().tolist()), key="det_site")
    with fc3:
        sel_alerte_det = st.selectbox("Statut",
            ["Tous", "Détenu", "Non détenu", "Réappro urgent"], key="det_al")

    df_det = df_long.copy()
    if sel_type_det != "Tous":   df_det = df_det[df_det["type"] == sel_type_det]
    if sel_site_det != "Tous":   df_det = df_det[df_det["site"] == sel_site_det]
    if sel_alerte_det == "Détenu":       df_det = df_det[df_det["detenu"]]
    elif sel_alerte_det == "Non détenu": df_det = df_det[~df_det["detenu"]]
    elif sel_alerte_det == "Réappro urgent": df_det = df_det[df_det["reappro_urgent"]]

    disp_det = df_det[["code","lib_topca","type","site","stock","detenu","reappro_urgent"]].copy()
    disp_det.columns = ["Code","Libellé","Type","Magasin","Stock","Détenu","Réappro urgent"]
    disp_det["Stock"]          = disp_det["Stock"].apply(lambda x: int(x) if pd.notna(x) else "—")
    disp_det["Détenu"]         = disp_det["Détenu"].map({True: "✅", False: "❌"})
    disp_det["Réappro urgent"] = disp_det["Réappro urgent"].map({True: "🔵", False: ""})

    st.markdown(f"<div style='font-size:12px;color:#8E8E93;margin-bottom:8px'>{len(disp_det):,} lignes</div>", unsafe_allow_html=True)
    st.dataframe(disp_det, use_container_width=True, hide_index=True)

# ══ TAB 4 — ABSENTS PBI ═══════════════════════════════════════════════════════
with tab4:
    if absents.empty:
        st.markdown("<div class='alert-card alert-green'>✅ Tous les articles Top CA sont présents dans l'export PBI.</div>", unsafe_allow_html=True)
    else:
        st.warning(f"⚠️ {len(absents)} référence(s) Top CA absentes du fichier PBI.")
        d = absents[["code","lib","type"]].copy()
        d.columns = ["Code article","Libellé","Type"]
        d["Vérification"] = "Vérifier déréférencement ou code article"
        st.dataframe(d, use_container_width=True, hide_index=True)

# ══ TAB 5 — EXPORT ════════════════════════════════════════════════════════════
with tab5:
    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>📋 5 feuilles Excel</strong><br>
  <strong>1. Synthèse réseau</strong> — taux GOLD / SILVER / Total par magasin<br>
  <strong>2. Ruptures nettes</strong> — stock = 0 sur tous les sites · GOLD en premier<br>
  <strong>3. Réappro urgent</strong> — 0 &lt; stock &lt; seuil · 1 ligne par article × site<br>
  <strong>4. Ruptures partielles</strong> — sites manquants listés<br>
  <strong>5. Absents PBI</strong> — articles Top CA introuvables
</div>""", unsafe_allow_html=True)

    st.caption(f"Périmètre : {n_refs} réf · {n_sites} sites · type : {type_filtre} · seuil réappro : {seuil_reappro}")

    if st.button("Générer l'export Excel", type="primary"):
        with st.spinner("Génération…"):
            buf = gen_excel(taux_df, types_seg, ruptures_nettes, reappro,
                            partielles, absents, type_filtre, cible, seuil_reappro)
        st.download_button(
            "⬇️ Télécharger",
            data=buf,
            file_name=f"SmartBuyer_Detention_{type_filtre}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ─── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(f"""
<div style='text-align:center;color:#C7C7CC;font-size:11px;padding:8px 0'>
  NovaRetail Solutions · SmartBuyer v2.3 · Détention Top CA · {n_refs} réf · {n_sites} sites
</div>""", unsafe_allow_html=True)
