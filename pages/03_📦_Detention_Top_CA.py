"""
03_📦_Detention_Top_CA.py — SmartBuyer Hub
Taux de détention Top CA · GOLD / SILVER · Flux IM/LO · Articles Permanents
Charte SmartBuyer v2
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

# ─── CHARTE SMARTBUYER ────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display",
                 "SF Pro Text", "Helvetica Neue", Arial, sans-serif !important;
    background-color: #F2F2F7;
}
.stApp { background: #F2F2F7; }
.main .block-container { padding-top: 1.8rem; max-width: 1200px; }
[data-testid="stSidebar"] { background: #F2F2F7 !important; border-right: 0.5px solid #D1D1D6 !important; }
[data-testid="stMetric"] { background: #FFFFFF !important; border: 0.5px solid #E5E5EA !important; border-radius: 12px !important; padding: 16px 18px !important; }
[data-testid="stMetricLabel"] { font-size: 11px !important; font-weight: 500 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 600 !important; color: #1C1C1E !important; letter-spacing: -0.02em !important; }
[data-testid="stMetricDelta"] { font-size: 12px !important; }
[data-testid="stTabs"] button[role="tab"] { font-size: 13px !important; font-weight: 500 !important; padding: 8px 16px !important; color: #8E8E93 !important; border-radius: 0 !important; border-bottom: 2px solid transparent !important; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color: #007AFF !important; border-bottom: 2px solid #007AFF !important; background: transparent !important; }
[data-testid="stTabs"] [role="tablist"] { border-bottom: 0.5px solid #E5E5EA !important; }
[data-testid="stDataFrame"] { border: 0.5px solid #E5E5EA !important; border-radius: 10px !important; }
[data-testid="stDataFrame"] th { background: #F2F2F7 !important; font-size: 11px !important; font-weight: 600 !important; color: #8E8E93 !important; text-transform: uppercase !important; letter-spacing: 0.04em !important; }
[data-testid="stFileUploader"] { border: 1.5px dashed #D1D1D6 !important; border-radius: 10px !important; background: #F9F9FB !important; }
.stDownloadButton > button { background: #007AFF !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; font-size: 13px !important; padding: 10px 24px !important; width: 100% !important; }
hr { border-color: #E5E5EA !important; margin: 1rem 0 !important; }

.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 0.5rem; }
.section-label { font-size: 11px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 10px; }

.alert-card  { padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; font-size: 13px; line-height: 1.5; border-left: 3px solid; }
.alert-red   { background: #FFF2F2; border-color: #FF3B30; color: #3A0000; }
.alert-amber { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }
.alert-green { background: #F0FFF4; border-color: #34C759; color: #003A10; }
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }

/* Sélecteur GOLD / SILVER */
.type-selector { display: flex; gap: 8px; margin-bottom: 16px; }
.type-btn { padding: 8px 20px; border-radius: 100px; font-size: 13px; font-weight: 600; cursor: pointer; border: 1.5px solid; }
.type-gold   { background: #FFF8E1; color: #B8860B; border-color: #FFD700; }
.type-silver { background: #F2F2F7; color: #636366; border-color: #C7C7CC; }
.type-all    { background: #E6F1FB; color: #007AFF; border-color: #007AFF; }

/* KPI GOLD highlight */
.kpi-gold { background: linear-gradient(135deg,#FFF8E1,#FFFBF0); border: 1px solid #FFD700; border-radius: 12px; padding: 16px 18px; }
.kpi-gold-label { font-size: 11px; font-weight: 500; color: #B8860B; text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 3px; }
.kpi-gold-value { font-size: 24px; font-weight: 700; color: #B8860B; letter-spacing: -0.02em; }
.kpi-gold-sub   { font-size: 12px; color: #996600; margin-top: 3px; font-weight: 500; }

.col-required { background: #F0F8FF; border: 0.5px solid #B3D9FF; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; display: flex; align-items: flex-start; gap: 10px; }
.col-name  { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc  { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
.col-ex    { font-size: 11px; color: #8E8E93; font-family: monospace; margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTES ───────────────────────────────────────────────────────────────
ETAT_LABELS = {
    "2": ("Actif",      "#34C759", "#F0FFF4"),
    "B": ("Bloqué",     "#FF3B30", "#FFF2F2"),
    "P": ("Permanent",  "#007AFF", "#E6F1FB"),
    "S": ("Saisonnier", "#FF9500", "#FFFBF0"),
    "F": ("Fin de vie", "#8E8E93", "#F2F2F7"),
}
TYPE_COLORS = {
    "GOLD":   ("#B8860B", "#FFF8E1", "#FFD700"),
    "SILVER": ("#636366", "#F2F2F7", "#C7C7CC"),
}
PGC = {"BOISSONS","DROGUERIE","PARFUMERIE HYGIENE","EPICERIE"}

def fmt(n):
    if pd.isna(n) or n is None: return "—"
    a = abs(n)
    if a >= 1_000_000: return f"{n/1_000_000:.1f} M"
    if a >= 1_000:     return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def norm_code(s):
    return s.astype(str).str.strip().str.replace(r"\.0$","",regex=True).str.zfill(8)

def color_taux(v, cible):
    if v is None: return "#8E8E93"
    return "#34C759" if v >= cible else "#FF9500" if v >= cible-10 else "#FF3B30"

def badge_etat(code):
    lbl, fg, bg = ETAT_LABELS.get(code, ("?","#8E8E93","#F2F2F7"))
    return f"<span style='background:{bg};color:{fg};padding:2px 8px;border-radius:100px;font-size:10px;font-weight:600'>{code} · {lbl}</span>"

def badge_type(t):
    fg, bg, bd = TYPE_COLORS.get(t, ("#8E8E93","#F2F2F7","#C7C7CC"))
    return f"<span style='background:{bg};color:{fg};border:1px solid {bd};padding:2px 8px;border-radius:100px;font-size:10px;font-weight:700'>{t}</span>"

# ─── PARSING ──────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_topca(byt, fname=""):
    """
    Lit la liste Top CA avec colonnes CODE ARTICLE, LIBELLÉ ARTICLE, TYPE.
    Gère UTF-8, latin-1, cp1252 automatiquement.
    """
    for encoding in ("utf-8-sig","utf-8","latin-1","cp1252"):
        try:
            raw = byt.decode(encoding, errors="strict")
            sep = ";" if raw.count(";") > raw.count(",") else ","
            df  = pd.read_csv(BytesIO(byt), sep=sep, encoding=encoding, dtype=str)
            # Normaliser les noms de colonnes
            df.columns = df.columns.str.strip().str.upper()
            # Code article : 1ère colonne ou colonne nommée CODE ARTICLE
            code_col = "CODE ARTICLE" if "CODE ARTICLE" in df.columns else df.columns[0]
            type_col = "TYPE" if "TYPE" in df.columns else None
            lib_col  = next((c for c in df.columns if "LIB" in c), None)

            result = pd.DataFrame()
            result["code"] = norm_code(df[code_col])
            result["type"] = df[type_col].str.strip().str.upper() if type_col else "GOLD"
            result["lib"]  = df[lib_col].astype(str).str.strip() if lib_col else ""
            result = result.dropna(subset=["code"])
            if len(result) > 0:
                return result
        except Exception:
            continue
    st.error("Erreur lecture liste Top CA — encodage non reconnu.")
    return pd.DataFrame(columns=["code","type","lib"])

@st.cache_data(show_spinner=False)
def load_stock(file_bytes, file_name):
    try:
        raw = pd.read_csv(BytesIO(file_bytes), sep=";", encoding="utf-8-sig",
                          dtype=str, low_memory=False)
    except Exception as e:
        st.error(f"Erreur lecture stock : {e}"); return pd.DataFrame()

    for col in ["Nouveau stock","Ral","Nb colis","PMP","Prix d'achat"]:
        if col in raw.columns:
            raw[col] = pd.to_numeric(raw[col], errors="coerce").fillna(0)

    raw["Code article"]      = norm_code(raw["Code article"])
    raw["Code etat"]         = raw["Code etat"].astype(str).str.strip().str.upper()
    raw["Code marketing"]    = raw.get("Code marketing",    pd.Series("?",index=raw.index)).astype(str).str.strip().str.upper()
    raw["Type saisonnalité"] = raw.get("Type saisonnalité", pd.Series("?",index=raw.index)).astype(str).str.strip().str.upper()

    if "Libellé rayon" in raw.columns:
        raw = raw[raw["Libellé rayon"].str.upper().isin(PGC)]
    return raw

# ─── CALCUL PRINCIPAL ────────────────────────────────────────────────────────
def compute(df_stock, df_topca, type_filtre, cible):
    """
    Périmètre : articles de la liste Top CA uniquement (filtrés par type si besoin).
    Dénominateur : tous les articles Top CA trouvés dans le stock.
    Numérateur   : articles avec stock > 0 (quel que soit le code état).
    Filtre       : Type saisonnalité = P (Permanent).
    """
    # Filtrer Top CA par type sélectionné
    if type_filtre != "Tous":
        df_topca_f = df_topca[df_topca["type"]==type_filtre].copy()
    else:
        df_topca_f = df_topca.copy()

    top_codes = set(df_topca_f["code"].unique())
    # Garder 1 seule ligne par code (prendre le premier TYPE en cas de doublon)
    top_meta  = df_topca_f.drop_duplicates(subset=["code"]).set_index("code")[["type","lib"]]
    top_map   = top_meta.to_dict("index")

    # Filtrer stock : permanents uniquement + articles de la liste
    df = df_stock[
        (df_stock["Code article"].isin(top_codes)) &
        (df_stock["Type saisonnalité"]=="P")
    ].copy()

    # Agréger par article × magasin
    agg_cols = {
        "code_etat":       ("Code etat",         lambda x: x.mode().iloc[0] if len(x) else "?"),
        "flux":            ("Code marketing",    lambda x: x.mode().iloc[0] if len(x) else "?"),
        "stock":           ("Nouveau stock",     "sum"),
        "ral":             ("Ral",               "sum"),
        "nb_colis":        ("Nb colis",          "first"),
        "lib_article":     ("Libellé article",   "first"),
        "lib_rayon":       ("Libellé rayon",     "first") if "Libellé rayon"  in df_stock.columns else ("Code article","first"),
        "lib_fournisseur": ("Nom fourn.",        "first") if "Nom fourn."     in df_stock.columns else ("Code article","first"),
    }
    grp = df.groupby(["Code article","Libellé site"]).agg(**agg_cols).reset_index()

    # Ajouter type et lib depuis Top CA
    grp["type_ca"]  = grp["Code article"].map(lambda c: top_map.get(c,{}).get("type","?"))
    grp["lib_topca"]= grp["Code article"].map(lambda c: top_map.get(c,{}).get("lib",""))

    # Détenu = stock > 0
    grp["detenu"] = grp["stock"] > 0
    grp["alerte"] = grp.apply(_alerte, axis=1)

    # Absents
    found      = set(df["Code article"].unique())
    absents_df = df_topca_f[~df_topca_f["code"].isin(found)].copy()

    return grp, absents_df, top_codes

def _alerte(row):
    if not row["detenu"]:
        if row["code_etat"] == "B": return "🔴 Bloqué — 0 stock"
        if row["ral"] > 0:          return "🚚 Relance en cours"
        return "🛒 Rupture"
    if row["nb_colis"] > 0 and 0 < row["stock"] < row["nb_colis"]:
        return "⚠️ Stock faible"
    if row["code_etat"] == "B":
        return "⚠️ Bloqué — stock résiduel"
    return "✅ OK"

def compute_taux(grp, top_codes, cible):
    """Taux par magasin et flux."""
    sites = sorted(grp["Libellé site"].unique())
    rows  = []
    for site in sites:
        s = grp[grp["Libellé site"]==site]
        for flux in ["ALL","IM","LO"]:
            sf = s if flux=="ALL" else s[s["flux"]==flux]
            n_total  = len(sf)
            n_detenu = int(sf["detenu"].sum())
            taux     = round(n_detenu/n_total*100,1) if n_total>0 else None

            # Par type
            sg = sf[sf["type_ca"]=="GOLD"]
            ss = sf[sf["type_ca"]=="SILVER"]
            t_gold   = round(sg["detenu"].sum()/len(sg)*100,1) if len(sg)>0 else None
            t_silver = round(ss["detenu"].sum()/len(ss)*100,1) if len(ss)>0 else None

            n_rupture = int((~sf["detenu"]).sum())
            n_bloque  = int((sf["code_etat"]=="B").sum())
            n_gold_rupt = int((sg["detenu"]==False).sum()) if len(sg)>0 else 0

            rows.append({
                "site":site,"flux":flux,
                "n_total":n_total,"n_detenu":n_detenu,"taux":taux,
                "taux_gold":t_gold,"taux_silver":t_silver,
                "n_rupture":n_rupture,"n_bloque":n_bloque,
                "n_gold_rupt":n_gold_rupt,
            })
    return pd.DataFrame(rows)

# ─── EXPORT EXCEL ────────────────────────────────────────────────────────────
def gen_excel(grp, taux_df, absents_df, type_filtre):
    wb    = Workbook()
    HDR_F = PatternFill("solid", fgColor="1C3557")
    HDR_T = Font(bold=True, color="FFFFFF", size=11)
    RED_F = PatternFill("solid", fgColor="FCE4E4")
    AMB_F = PatternFill("solid", fgColor="FEF3CD")
    GRN_F = PatternFill("solid", fgColor="D6F0D6")
    GLD_F = PatternFill("solid", fgColor="FFF8DC")
    NEU_F = PatternFill("solid", fgColor="FFFFFF")
    CTR   = Alignment(horizontal="center", vertical="center")

    def write_ws(ws, headers, rows, title):
        ws.append([title]); ws.cell(1,1).font=Font(bold=True,size=13)
        ws.append([]); ws.append(headers)
        for i,h in enumerate(headers,1):
            c=ws.cell(3,i); c.fill=HDR_F; c.font=HDR_T; c.alignment=CTR
        for row in rows: ws.append(row)
        for col in ws.columns:
            ws.column_dimensions[get_column_letter(col[0].column)].width=max(
                len(str(col[0].value or ""))+4,12)

    taux_all = taux_df[taux_df["flux"]=="ALL"]

    # Onglet 1 — Synthèse
    ws1 = wb.active; ws1.title="Synthèse réseau"
    rows1 = [[r["site"],r["n_total"],r["n_detenu"],r["taux"],
              r["taux_gold"],r["taux_silver"],r["n_rupture"],r["n_bloque"],r["n_gold_rupt"]]
             for _,r in taux_all.iterrows()]
    write_ws(ws1,
        ["Magasin","Réf analysées","Détenues","Taux %",
         "Taux GOLD %","Taux SILVER %","Ruptures","Bloqués (B)","Ruptures GOLD"],
        rows1, f"Synthèse détention — Top CA {type_filtre} · Permanents")
    for r in ws1.iter_rows(min_row=4, max_row=ws1.max_row):
        for ci in [4,5,6]:
            v=r[ci-1].value
            if isinstance(v,(int,float)):
                r[ci-1].fill=GRN_F if v>=85 else AMB_F if v>=70 else RED_F

    # Onglet 2 — IM vs LO
    ws2 = wb.create_sheet("IM vs LO")
    rows2=[[r["site"],r["flux"],r["n_total"],r["n_detenu"],r["taux"],
            r["taux_gold"],r["taux_silver"],r["n_rupture"]]
           for _,r in taux_df[taux_df["flux"]!="ALL"].iterrows()]
    write_ws(ws2,
        ["Magasin","Flux","Réf","Détenues","Taux %","Taux GOLD %","Taux SILVER %","Ruptures"],
        rows2,"IM vs LO · Permanents")
    for r in ws2.iter_rows(min_row=4,max_row=ws2.max_row):
        v=r[4].value
        if isinstance(v,(int,float)):
            r[4].fill=GRN_F if v>=85 else AMB_F if v>=70 else RED_F

    # Onglet 3 — GOLD ruptures (prioritaire)
    ws3 = wb.create_sheet("🥇 GOLD — Ruptures")
    gold_rupt = grp[(grp["type_ca"]=="GOLD")&(grp["detenu"]==False)].sort_values("Libellé site")
    rows3=[[r["Code article"],r["lib_topca"],r["Libellé site"],r["flux"],
            r["code_etat"],int(r["stock"]),int(r["ral"]),r["alerte"]]
           for _,r in gold_rupt.iterrows()]
    write_ws(ws3,
        ["Code","Libellé","Magasin","Flux","Code état","Stock","RAL","Alerte"],
        rows3,"Articles GOLD non détenus · Action prioritaire")
    for r in ws3.iter_rows(min_row=4,max_row=ws3.max_row):
        r[0].fill=GLD_F
        v=str(r[7].value or "")
        r[7].fill=RED_F if "🔴" in v or "🛒" in v else AMB_F if "⚠️" in v or "🚚" in v else NEU_F

    # Onglet 4 — Plan d'action complet
    ws4 = wb.create_sheet("Plan d'action")
    urgences = grp[grp["alerte"]!="✅ OK"].sort_values(["type_ca","alerte"])
    rows4=[[r["Code article"],r["lib_topca"],r["type_ca"],r["Libellé site"],
            r["flux"],r["code_etat"],int(r["stock"]),int(r["ral"]),r["alerte"]]
           for _,r in urgences.iterrows()]
    write_ws(ws4,
        ["Code","Libellé","Type","Magasin","Flux","Code état","Stock","RAL","Alerte"],
        rows4,"Plan d'action · GOLD en priorité")
    for r in ws4.iter_rows(min_row=4,max_row=ws4.max_row):
        if str(r[2].value)=="GOLD": r[2].fill=GLD_F
        v=str(r[8].value or "")
        r[8].fill=RED_F if "🔴" in v or "🛒" in v else AMB_F if "⚠️" in v or "🚚" in v else NEU_F

    # Onglet 5 — Absents ERP
    ws5 = wb.create_sheet("Absents ERP")
    rows5=[[r["code"],r["lib"],r["type"],"Absent ERP"] for _,r in absents_df.iterrows()]
    write_ws(ws5,["Code","Libellé","Type","Statut"],rows5,"Absents des extractions ERP")

    buf=BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:18px'>
  <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>🛍️ SmartBuyer</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · Équipe Achats</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Navigation</div>", unsafe_allow_html=True)
    st.page_link("app.py",                                  label="🏠  Accueil")
    st.page_link("pages/01_📊_Analyse_Scoring_ABC.py",      label="📊  Scoring ABC")
    st.page_link("pages/02_📈_Ventes_PBI.py",               label="📈  Ventes PBI")
    st.page_link("pages/03_📦_Detention_Top_CA.py",         label="📦  Détention Top CA")
    st.page_link("pages/04_💸_Performance_Promo.py",        label="💸  Performance Promo",  disabled=True)
    st.page_link("pages/05_🏪_Suivi_Implantation.py",       label="🏪  Suivi Implantation", disabled=True)
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Import fichiers</div>", unsafe_allow_html=True)
    f_topca  = st.file_uploader("Liste Top CA (CSV)", type=["csv","xlsx"], key="topca",
                                 help="Colonnes : CODE ARTICLE · LIBELLÉ ARTICLE · TYPE (GOLD/SILVER)")
    f_stock  = st.file_uploader("Stock consolidé (CSV)", type=["csv"], key="stock",
                                 help="Produit par consolider_stock.py · séparateur ; · UTF-8")
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Paramètres</div>", unsafe_allow_html=True)
    cible = st.slider("Cible taux de détention (%)", 70, 100, 85)
    st.markdown(f"""
<div style='background:#E6F1FB;border:0.5px solid #B3D9FF;border-radius:8px;
            padding:7px 11px;font-size:12px;color:#001A3A;margin-top:4px'>
  Cible : <strong>{cible}%</strong> · Articles <strong>Permanents</strong><br>
  Détenu = <strong>stock &gt; 0</strong> (tous codes état)
</div>""", unsafe_allow_html=True)

# ─── PAGE PRINCIPALE ──────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>📦 Détention Top CA</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Articles Permanents · Flux IM / LO · GOLD / SILVER · Détenu = stock &gt; 0</div>", unsafe_allow_html=True)

# ─── ÉCRAN D'ACCUEIL ─────────────────────────────────────────────────────────
if not f_topca or not f_stock:
    st.markdown("---")
    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>ℹ️ À quoi sert ce module ?</strong><br>
  Ce module vérifie la présence en magasin des articles <strong>Top CA permanents</strong>
  et calcule le taux de détention par niveau de priorité (<strong>GOLD / SILVER</strong>)
  et par flux (<strong>IM / LO</strong>).<br><br>
  Un article est considéré <strong>détenu dès qu'il a du stock &gt; 0</strong>,
  quel que soit son code état. Le code état est affiché pour information.
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Règles de calcul</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;
            padding:16px;border-left:3px solid #34C759;margin-bottom:10px'>
  <div style='font-size:13px;font-weight:600;color:#1C1C1E;margin-bottom:8px'>Calcul du taux</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:2'>
    <strong>Périmètre</strong> : articles de la liste Top CA uniquement<br>
    <strong>Filtre</strong> : Type saisonnalité = P (Permanent)<br>
    <strong>Dénominateur</strong> : tous articles Top CA trouvés dans le stock<br>
    <strong>Numérateur</strong> : articles avec stock &gt; 0 (tous codes état)
  </div>
</div>""", unsafe_allow_html=True)
    with c2:
        st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;
            padding:16px;border-left:3px solid #FFD700;margin-bottom:10px'>
  <div style='font-size:13px;font-weight:600;color:#1C1C1E;margin-bottom:8px'>Niveaux de priorité</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:2'>
    <span style='background:#FFF8E1;color:#B8860B;padding:2px 8px;border-radius:100px;
                 font-weight:700;font-size:11px;border:1px solid #FFD700'>GOLD</span>
    Articles prioritaires — alertes critiques<br>
    <span style='background:#F2F2F7;color:#636366;padding:2px 8px;border-radius:100px;
                 font-weight:700;font-size:11px;border:1px solid #C7C7CC'>SILVER</span>
    Articles secondaires — alertes standard<br>
    Le code état est affiché à titre informatif sur chaque article.
  </div>
</div>""", unsafe_allow_html=True)

    st.markdown("<div class='section-label'>Fichiers attendus</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
<div class='col-required'><div style='font-size:16px'>📋</div>
<div><div class='col-name'>Liste Top CA</div>
<div class='col-desc'>CSV avec en-tête · 3 colonnes</div>
<div class='col-ex'>CODE ARTICLE ; LIBELLÉ ARTICLE ; TYPE (GOLD/SILVER)</div></div></div>""",
        unsafe_allow_html=True)
    with c2:
        st.markdown("""
<div class='col-required'><div style='font-size:16px'>🏪</div>
<div><div class='col-name'>Stock consolidé ERP</div>
<div class='col-desc'>Produit par consolider_stock.py · 1 seul fichier · tous magasins</div>
<div class='col-ex'>stock_consolide_YYYYMMDD.csv</div></div></div>""",
        unsafe_allow_html=True)

    st.info("⬆️ Charge les deux fichiers dans la sidebar pour lancer l'analyse.")
    st.stop()

# ─── TRAITEMENT ───────────────────────────────────────────────────────────────
with st.spinner("Lecture des fichiers…"):
    df_topca = load_topca(f_topca.read(), f_topca.name)
    df_stock = load_stock(f_stock.read(), f_stock.name)

if df_stock.empty: st.error("Stock vide ou illisible."); st.stop()
if df_topca.empty: st.error("Liste Top CA vide."); st.stop()

# ─── SÉLECTEUR GOLD / SILVER ─────────────────────────────────────────────────
st.markdown("<div class='section-label'>Niveau d'analyse</div>", unsafe_allow_html=True)
types_dispo = sorted(df_topca["type"].unique())
options     = ["Tous"] + types_dispo
labels      = {"Tous":"🔵 Tous", "GOLD":"🥇 GOLD", "SILVER":"🥈 SILVER"}

type_filtre = st.radio(
    "Type à analyser",
    options,
    format_func=lambda x: labels.get(x, x),
    horizontal=True,
    label_visibility="collapsed",
)

# Sous-titre contextuel
n_type = len(df_topca[df_topca["type"]==type_filtre]) if type_filtre!="Tous" else len(df_topca)
if type_filtre == "GOLD":
    st.markdown(f"<div style='font-size:12px;color:#B8860B;background:#FFF8E1;padding:6px 12px;border-radius:8px;border:0.5px solid #FFD700;margin-bottom:12px;display:inline-block'>🥇 Analyse GOLD · <strong>{n_type}</strong> articles prioritaires</div>", unsafe_allow_html=True)
elif type_filtre == "SILVER":
    st.markdown(f"<div style='font-size:12px;color:#636366;background:#F2F2F7;padding:6px 12px;border-radius:8px;border:0.5px solid #C7C7CC;margin-bottom:12px;display:inline-block'>🥈 Analyse SILVER · <strong>{n_type}</strong> articles</div>", unsafe_allow_html=True)
else:
    n_gold   = len(df_topca[df_topca["type"]=="GOLD"])
    n_silver = len(df_topca[df_topca["type"]=="SILVER"])
    st.markdown(f"<div style='font-size:12px;color:#007AFF;background:#E6F1FB;padding:6px 12px;border-radius:8px;border:0.5px solid #B3D9FF;margin-bottom:12px;display:inline-block'>🔵 Tous · <strong>{n_gold}</strong> GOLD + <strong>{n_silver}</strong> SILVER = <strong>{n_type}</strong> articles</div>", unsafe_allow_html=True)

# ─── CALCUL ───────────────────────────────────────────────────────────────────
with st.spinner("Calcul des taux de détention…"):
    grp, absents_df, top_codes = compute(df_stock, df_topca, type_filtre, cible)
    taux_df = compute_taux(grp, top_codes, cible)

n_sites     = df_stock["Libellé site"].nunique()
taux_all    = taux_df[taux_df["flux"]=="ALL"]
taux_im     = taux_df[taux_df["flux"]=="IM"]
taux_lo     = taux_df[taux_df["flux"]=="LO"]
taux_moy    = taux_all["taux"].mean()
taux_im_moy = taux_im["taux"].mean()
taux_lo_moy = taux_lo["taux"].mean()
taux_gold   = taux_all["taux_gold"].mean()
n_urgences  = (grp["alerte"]!="✅ OK").sum()
n_gold_rupt = (grp[grp["type_ca"]=="GOLD"]["detenu"]==False).sum()

# ─── KPIs ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(f"<div class='section-label'>{n_sites} magasin(s) · {len(top_codes)} références · Permanents</div>", unsafe_allow_html=True)

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Réf analysées",  str(len(top_codes)), f"{len(absents_df)} absents ERP" if len(absents_df) else "")
k2.metric("Taux détention", f"{taux_moy:.1f}%" if taux_moy else "—",
          f"{taux_moy-cible:+.1f} pt vs cible" if taux_moy else "")
k3.metric("Taux IM · Import",
          f"{taux_im_moy:.1f}%" if taux_im_moy else "—",
          f"{taux_im_moy-cible:+.1f} pt" if taux_im_moy else "")
k4.metric("Taux LO · Local",
          f"{taux_lo_moy:.1f}%" if taux_lo_moy else "—",
          f"{taux_lo_moy-cible:+.1f} pt" if taux_lo_moy else "")

# KPI GOLD mis en avant
with k5:
    gold_color = "#34C759" if (taux_gold or 0)>=cible else "#FF9500" if (taux_gold or 0)>=cible-10 else "#FF3B30"
    gold_sub   = f"{(taux_gold or 0)-cible:+.1f} pt vs cible" if taux_gold else "—"
    st.markdown(f"""
<div class='kpi-gold'>
  <div class='kpi-gold-label'>🥇 Taux GOLD</div>
  <div class='kpi-gold-value' style='color:{gold_color}'>{f"{taux_gold:.1f}%" if taux_gold else "—"}</div>
  <div class='kpi-gold-sub'>{gold_sub} · {n_gold_rupt} ruptures</div>
</div>""", unsafe_allow_html=True)

# ─── ALERTES ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Alertes &amp; actions prioritaires</div>", unsafe_allow_html=True)

# Ruptures GOLD en priorité absolue
if n_gold_rupt > 0:
    sites_gold = grp[(grp["type_ca"]=="GOLD")&(grp["detenu"]==False)]["Libellé site"].value_counts()
    top3 = ", ".join([f"{s} ({n})" for s,n in sites_gold.head(3).items()])
    st.markdown(f"""
<div class='alert-card alert-red'>
  <strong>🥇 {n_gold_rupt} article(s) GOLD non détenus</strong> — {top3}<br>
  <span style='font-size:12px;opacity:.85'>→ Priorité absolue. Commander en urgence ou identifier substitut.</span>
</div>""", unsafe_allow_html=True)

# Magasins sous cible
sites_sous = taux_all[taux_all["taux"]<cible].sort_values("taux")
if not sites_sous.empty:
    liste = ", ".join([f"{r['site']} ({r['taux']:.0f}%)" for _,r in sites_sous.iterrows()])
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ {len(sites_sous)} magasin(s) sous la cible {cible}%</strong> — {liste}<br>
  <span style='font-size:12px;opacity:.85'>→ Analyser les ruptures par flux. IM : délai 4–8 sem. LO : réassort 48h.</span>
</div>""", unsafe_allow_html=True)

# IM sous cible
im_sous = taux_im[taux_im["taux"]<cible]
if not im_sous.empty:
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ Flux Import sous cible sur {len(im_sous)} magasin(s)</strong> — taux IM moyen {taux_im_moy:.1f}%<br>
  <span style='font-size:12px;opacity:.85'>→ Anticiper les commandes Import. Délai réappro 4–8 semaines.</span>
</div>""", unsafe_allow_html=True)

if len(absents_df) > 0:
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ {len(absents_df)} référence(s) Top CA absentes de toutes les extractions</strong><br>
  <span style='font-size:12px;opacity:.85'>→ Vérifier déréférencement ou erreur de code dans la liste Top CA.</span>
</div>""", unsafe_allow_html=True)

if n_urgences==0 and sites_sous.empty and n_gold_rupt==0:
    st.success("✅ Tous les magasins au-dessus de la cible · aucune rupture GOLD détectée.")

# ─── TABS ─────────────────────────────────────────────────────────────────────
st.markdown("---")
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Synthèse réseau",
    "🔄 IM vs LO",
    "🚨 Plan d'action",
    "🚫 Absents ERP"
])

# ═══ TAB 1 — SYNTHÈSE ═════════════════════════════════════════════════════════
with tab1:
    st.caption("Détenu = stock > 0 · Tous codes état · Permanents · " + ("GOLD + SILVER" if type_filtre=="Tous" else type_filtre))

    disp1 = taux_all[["site","n_total","n_detenu","taux","taux_gold","taux_silver","n_rupture","n_bloque","n_gold_rupt"]].copy()
    disp1.columns = ["Magasin","Réf","Détenues","Taux %","Taux GOLD %","Taux SILVER %","Ruptures","Bloqués (B)","Ruptures GOLD"]
    disp1 = disp1.sort_values("Taux %")
    for c in ["Taux %","Taux GOLD %","Taux SILVER %"]:
        disp1[c] = disp1[c].apply(lambda x: f"{x:.1f}%" if x is not None and not (isinstance(x,float) and np.isnan(x)) else "—")
    disp1["Statut"] = taux_all.sort_values("taux")["taux"].apply(
        lambda x: "🟢 OK" if x and x>=cible else ("🟡 Surveiller" if x and x>=cible-10 else "🔴 Action"))

    st.dataframe(disp1, use_container_width=True, hide_index=True)

    # Graphique SVG — barres arrondies modernes style Apple
    def render_chart(rows, cible, show_gold):
        n      = len(rows)
        W      = 420         # largeur zone barres
        LBL_W  = 190         # largeur labels
        ROW_H  = 50 if show_gold else 44
        BAR_G  = 14          # barre taux global
        BAR_g  = 9           # barre GOLD
        GAP    = 5            # écart entre les deux barres
        R      = 7
        SVG_W  = LBL_W + W + 80
        SVG_H  = n * ROW_H + 64
        SCALE  = W / 115.0

        def fill(v):
            if v is None: return "var(--color-border-tertiary)"
            if v >= cible:      return "url(#gg)"
            if v >= cible - 10: return "url(#ga)"
            return "url(#gr)"

        lines = [
            f'<svg xmlns="http://www.w3.org/2000/svg" width="100%" ',
            f'viewBox="0 0 {SVG_W} {SVG_H}" ',
            f'style="font-family:-apple-system,BlinkMacSystemFont,Helvetica Neue,Arial;background:transparent">',
            "<defs>",
            '<linearGradient id="gg" x1="0%" y1="0%" x2="100%" y2="0%">',
            '<stop offset="0%" stop-color="#34C759"/><stop offset="100%" stop-color="#28A745"/></linearGradient>',
            '<linearGradient id="ga" x1="0%" y1="0%" x2="100%" y2="0%">',
            '<stop offset="0%" stop-color="#FF9500"/><stop offset="100%" stop-color="#E68600"/></linearGradient>',
            '<linearGradient id="gr" x1="0%" y1="0%" x2="100%" y2="0%">',
            '<stop offset="0%" stop-color="#FF3B30"/><stop offset="100%" stop-color="#E62E24"/></linearGradient>',
            '<linearGradient id="gd" x1="0%" y1="0%" x2="100%" y2="0%">',
            '<stop offset="0%" stop-color="#FFD700"/><stop offset="100%" stop-color="#FFC200"/></linearGradient>',
            "</defs>",
        ]

        # Légende
        lx = LBL_W
        lines.append(f'<rect x="{lx}" y="10" width="14" height="10" rx="5" fill="url(#gg)" opacity=".9"/>')
        lines.append(f'<text x="{lx+18}" y="19" font-size="10" fill="var(--color-text-secondary)">Taux global</text>')
        if show_gold:
            lines.append(f'<rect x="{lx+110}" y="10" width="14" height="10" rx="5" fill="url(#gd)" opacity=".9"/>')
            lines.append(f'<text x="{lx+128}" y="19" font-size="10" fill="var(--color-text-secondary)">GOLD</text>')

        # Grille + labels axe
        for pct in [0, 20, 40, 60, 80, 100]:
            x = LBL_W + pct * SCALE
            lines.append(f'<line x1="{x:.1f}" y1="28" x2="{x:.1f}" y2="{SVG_H-26}" stroke="var(--color-border-tertiary)" stroke-width="0.5"/>')
            lines.append(f'<text x="{x:.1f}" y="38" text-anchor="middle" font-size="10" fill="var(--color-text-tertiary)" font-size="9">{pct}%</text>')

        # Ligne cible
        xc = LBL_W + cible * SCALE
        lines.append(f'<line x1="{xc:.1f}" y1="26" x2="{xc:.1f}" y2="{SVG_H-14}" stroke="#007AFF" stroke-width="1.5" stroke-dasharray="4,3"/>')
        lines.append(f'<text x="{xc:.1f}" y="{SVG_H-2}" text-anchor="middle" font-size="9" fill="#007AFF" font-weight="600">Cible {cible}%</text>')

        # Barres
        for i, row in enumerate(rows):
            y0   = 44 + i * ROW_H
            v    = row.get("taux") or 0
            vg   = row.get("taux_gold") or 0
            site = str(row.get("site", ""))[:24]

            # Calcul positions
            cy_global = y0 + (ROW_H - (BAR_G + (BAR_g + GAP if show_gold else 0))) // 2
            cy_gold   = cy_global + BAR_G + GAP

            # Label magasin — centré sur les deux barres
            label_y = cy_global + BAR_G // 2 + (BAR_g // 2 + GAP // 2 if show_gold else 0) + 4
            lines.append(
                f'<text x="{LBL_W-10}" y="{label_y}" text-anchor="end" font-size="11"'
                f' fill="var(--color-text-primary)" font-weight="500" font-size="10">{site}</text>'
            )

            def rr(bx, by, bw, bh, r, fill_ref, opacity=1.0):
                bw = max(float(bw), r * 2)
                return (
                    f'<rect x="{bx:.1f}" y="{by}" ',
                    f'width="{bw:.1f}" height="{bh}" rx="{r}" ry="{r}" ',
                    f'fill="{fill_ref}" opacity="{opacity}"/>',
                )

            # Track global
            lines.extend(rr(LBL_W, cy_global, W * 0.99, BAR_G, R, "var(--color-border-tertiary)", 0.35))
            # Barre global
            if v > 0:
                bw = v * SCALE
                lines.extend(rr(LBL_W, cy_global, bw, BAR_G, R, fill(v)))
                tx = LBL_W + bw + 6
                lines.append(
                    f'<text x="{tx:.1f}" y="{cy_global + BAR_G//2 + 4}"'
                    f' font-size="11" fill="var(--color-text-primary)" font-weight="600">{v:.1f}%</text>'
                )

            # Barre GOLD
            if show_gold:
                lines.extend(rr(LBL_W, cy_gold, W * 0.99, BAR_g, R - 2, "var(--color-border-tertiary)", 0.25))
                if vg > 0:
                    bwg = vg * SCALE
                    lines.extend(rr(LBL_W, cy_gold, bwg, BAR_g, R - 2, "url(#gd)", 0.85))
                    # Label GOLD
                    if bwg > 55:
                        lines.append(
                            f'<text x="{LBL_W + bwg / 2:.1f}" y="{cy_gold + BAR_g//2 + 3}"'
                            f' text-anchor="middle" font-size="9" fill="#7A5800" font-weight="600">{vg:.1f}%</text>'
                        )
                    else:
                        lines.append(
                            f'<text x="{LBL_W + bwg + 5:.1f}" y="{cy_gold + BAR_g//2 + 3}"'
                            f' font-size="9" fill="#B8860B" font-weight="600">{vg:.1f}%</text>'
                        )

        # Axe bas
        lines.append(f'<line x1="{LBL_W}" y1="{SVG_H-26}" x2="{LBL_W+W}" y2="{SVG_H-26}" stroke="var(--color-border-tertiary)" stroke-width="0.5"/>')
        lines.append(f'<text x="{LBL_W + W/2:.1f}" y="{SVG_H-12}" text-anchor="middle" font-size="10" fill="var(--color-text-tertiary)">Taux de détention (%)</text>')
        lines.append("</svg>")
        return "".join(lines)

    show_gold  = type_filtre in ("Tous", "GOLD") and taux_all["taux_gold"].notna().any()
    rows_chart = taux_all.sort_values("taux")[["site","taux","taux_gold","taux_silver"]].to_dict("records")
    svg_html   = render_chart(rows_chart, cible, show_gold)
    st.markdown(
        f'<div style="background:var(--color-background-primary);border:0.5px solid var(--color-border-tertiary);'
        f'border-radius:14px;padding:20px 16px 12px 16px;margin-top:12px">{svg_html}</div>',
        unsafe_allow_html=True,
    )

# ═══ TAB 2 — IM vs LO ═════════════════════════════════════════════════════════
with tab2:
    st.markdown("""
<div class='alert-card alert-blue'>
  Flux <strong>IM (Import)</strong> : délai réappro 4–8 semaines.
  Flux <strong>LO (Local)</strong> : réassort possible en 48h.
  Un écart important entre les deux signale des tensions à anticiper.
</div>""", unsafe_allow_html=True)

    tot_im  = taux_im["n_total"].sum(); det_im = taux_im["n_detenu"].sum()
    tot_lo  = taux_lo["n_total"].sum(); det_lo = taux_lo["n_detenu"].sum()
    g_im    = det_im/tot_im*100 if tot_im else 0
    g_lo    = det_lo/tot_lo*100 if tot_lo else 0

    c1, c2 = st.columns(2)
    for col_w, label, tg, tot, det, flux_col in [
        (c1,"IM · Import", g_im, tot_im, det_im, "#7C3AED"),
        (c2,"LO · Local",  g_lo, tot_lo, det_lo, "#007AFF"),
    ]:
        with col_w:
            cv = color_taux(tg, cible)
            st.markdown(f"""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-left:3px solid {flux_col};
            border-radius:12px;padding:16px 18px;margin-bottom:12px'>
  <div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;
              letter-spacing:.05em;margin-bottom:4px'>{label}</div>
  <div style='font-size:28px;font-weight:700;color:{cv};letter-spacing:-.02em'>{tg:.1f}%</div>
  <div style='font-size:12px;color:#8E8E93;margin-top:3px'>{int(det)} / {int(tot)} articles détenus · {n_sites} magasins</div>
  <div style='background:#E5E5EA;border-radius:3px;height:5px;margin-top:10px'>
    <div style='width:{min(tg,100):.0f}%;background:{flux_col};height:5px;border-radius:3px'></div>
  </div>
</div>""", unsafe_allow_html=True)

    try:
        pivot = taux_df[taux_df["flux"]!="ALL"].pivot_table(
            index="site", columns="flux",
            values=["n_total","n_detenu","taux","taux_gold"],
            aggfunc="first"
        ).reset_index()
        pivot.columns = ["Magasin",
                         "Réf IM","Réf LO",
                         "Détenues IM","Détenues LO",
                         "Taux GOLD IM","Taux GOLD LO",
                         "Taux IM %","Taux LO %"]
        pivot = pivot[["Magasin","Réf IM","Détenues IM","Taux IM %","Taux GOLD IM",
                        "Réf LO","Détenues LO","Taux LO %","Taux GOLD LO"]].sort_values("Taux IM %")
        for c in ["Taux IM %","Taux LO %","Taux GOLD IM","Taux GOLD LO"]:
            pivot[c] = pivot[c].apply(lambda x: f"{x:.1f}%" if pd.notna(x) and x is not None else "—")
        st.dataframe(pivot, use_container_width=True, hide_index=True)
    except Exception:
        st.dataframe(taux_df[taux_df["flux"]!="ALL"], use_container_width=True, hide_index=True)

# ═══ TAB 3 — PLAN D'ACTION ════════════════════════════════════════════════════
with tab3:
    urgences = grp[grp["alerte"]!="✅ OK"].copy()

    # Filtres inline
    fc1, fc2, fc3, fc4 = st.columns(4)
    with fc1:
        sel_site  = st.selectbox("Magasin", ["Tous"]+sorted(grp["Libellé site"].unique()))
    with fc2:
        sel_flux  = st.selectbox("Flux", ["Tous","IM","LO"])
    with fc3:
        types_opt = ["Tous"] + sorted(grp["type_ca"].unique())
        sel_type  = st.selectbox("Type", types_opt)
    with fc4:
        al_dispo  = sorted(urgences["alerte"].unique())
        sel_al    = st.multiselect("Alerte", al_dispo, default=al_dispo)

    if sel_site!="Tous": urgences=urgences[urgences["Libellé site"]==sel_site]
    if sel_flux!="Tous": urgences=urgences[urgences["flux"]==sel_flux]
    if sel_type!="Tous": urgences=urgences[urgences["type_ca"]==sel_type]
    if sel_al:           urgences=urgences[urgences["alerte"].isin(sel_al)]

    # Trier GOLD en premier, puis par alerte
    urgences = urgences.sort_values(["type_ca","alerte"], ascending=[True, True])

    st.markdown(f"<div style='font-size:12px;color:#8E8E93;margin-bottom:8px'><strong>{len(urgences)}</strong> article(s) nécessitant une action · GOLD en priorité</div>", unsafe_allow_html=True)

    if urgences.empty:
        st.success("✅ Aucune urgence sur la sélection.")
    else:
        disp3 = urgences[["Code article","lib_topca","type_ca","Libellé site",
                           "flux","code_etat","stock","ral","alerte"]].copy()
        disp3.columns = ["Code","Libellé","Type","Magasin","Flux","Code état","Stock","RAL","Alerte"]
        disp3["Stock"] = disp3["Stock"].apply(lambda x: int(x))
        disp3["RAL"]   = disp3["RAL"].apply(lambda x: int(x))
        st.dataframe(
            disp3, use_container_width=True, hide_index=True,
            column_config={
                "Type":      st.column_config.TextColumn("Type",      width="small"),
                "Flux":      st.column_config.TextColumn("Flux",      width="small"),
                "Code état": st.column_config.TextColumn("Code état", width="small"),
                "Stock":     st.column_config.NumberColumn("Stock",   format="%d"),
                "RAL":       st.column_config.NumberColumn("RAL",     format="%d"),
            }
        )

# ═══ TAB 4 — ABSENTS ERP ══════════════════════════════════════════════════════
with tab4:
    if absents_df.empty:
        st.success("✅ Toutes les références Top CA sont présentes dans au moins une extraction.")
    else:
        st.warning(f"⚠️ {len(absents_df)} référence(s) Top CA absentes de toutes les extractions ERP")
        disp4 = absents_df[["code","lib","type"]].copy()
        disp4.columns = ["Code article","Libellé","Type"]
        disp4["Action"] = "Vérifier référentiel ou déréférencement"
        st.dataframe(disp4, use_container_width=True, hide_index=True)

# ─── EXPORT ───────────────────────────────────────────────────────────────────
st.markdown("---")
with st.expander("📥 Export Excel — Synthèse · IM vs LO · GOLD Ruptures · Plan d'action · Absents ERP"):
    st.caption(f"5 onglets · {type_filtre} · Articles permanents · Détenu = stock > 0")
    if st.button("Générer le fichier Excel", type="primary"):
        with st.spinner("Génération…"):
            buf = gen_excel(grp, taux_df, absents_df, type_filtre)
        st.download_button(
            "⬇️ Télécharger",
            data=buf,
            file_name=f"SmartBuyer_Detention_{type_filtre}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
