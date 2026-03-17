"""
03_📦_Detention_Top_CA.py — SmartBuyer Hub
Taux de détention Top CA · Flux IM/LO · Code état par article × magasin
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
[data-testid="baseButton-primary"] { background: #007AFF !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; }
.stDownloadButton > button { background: #007AFF !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; font-size: 13px !important; padding: 10px 24px !important; width: 100% !important; }
hr { border-color: #E5E5EA !important; margin: 1rem 0 !important; }
.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 1.5rem; }
.section-label { font-size: 11px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 10px; }
.alert-card { padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; font-size: 13px; line-height: 1.5; border-left: 3px solid; }
.alert-red   { background: #FFF2F2; border-color: #FF3B30; color: #3A0000; }
.alert-amber { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }
.alert-green { background: #F0FFF4; border-color: #34C759; color: #003A10; }
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }
.col-required { background: #F0F8FF; border: 0.5px solid #B3D9FF; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; display: flex; align-items: flex-start; gap: 10px; }
.col-name    { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc    { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
.col-example { font-size: 11px; color: #8E8E93; font-family: monospace; margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTES ───────────────────────────────────────────────────────────────
ETAT_LABELS = {
    "2": ("Actif", "#34C759",  "#F0FFF4", "Inclus dans le taux"),
    "P": ("Permanent", "#007AFF", "#E6F1FB", "Exclu — signalé"),
    "S": ("Saisonnier", "#FF9500", "#FFFBF0", "Exclu — signalé"),
    "B": ("Bloqué", "#FF3B30",  "#FFF2F2", "Exclu — signalé"),
    "F": ("Fin de vie", "#8E8E93", "#F2F2F7", "Exclu — à retirer"),
    "1": ("Autre", "#8E8E93", "#F2F2F7", "Exclu"),
    "5": ("Autre", "#8E8E93", "#F2F2F7", "Exclu"),
    "6": ("Autre", "#8E8E93", "#F2F2F7", "Exclu"),
}
FLUX_LABELS = {"IM": ("Import", "#7C3AED", "#F0EEFF"), "LO": ("Local", "#007AFF", "#E6F1FB")}

def fmt(n):
    if pd.isna(n): return "—"
    a = abs(n)
    if a >= 1_000_000: return f"{n/1_000_000:.1f} M"
    if a >= 1_000:     return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def norm_code(s):
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(8)

# ─── PARSING ──────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_stock(files_bytes_names):
    dfs = []
    for byt, name in files_bytes_names:
        try:
            df = pd.read_csv(BytesIO(byt), sep=";", encoding="utf-8-sig",
                             dtype=str, low_memory=False)
            dfs.append(df)
        except Exception as e:
            st.warning(f"Erreur lecture {name} : {e}")
    if not dfs:
        return pd.DataFrame()
    raw = pd.concat(dfs, ignore_index=True)

    # Colonnes numériques
    for col in ["Nouveau stock", "Ral", "Nb colis", "Prix de vente", "Prix d'achat"]:
        if col in raw.columns:
            raw[col] = pd.to_numeric(raw[col], errors="coerce").fillna(0)

    raw["Code article"] = norm_code(raw["Code article"])
    raw["Code etat"]    = raw["Code etat"].astype(str).str.strip().str.upper()
    raw["Code marketing"] = raw["Code marketing"].astype(str).str.strip().str.upper() if "Code marketing" in raw.columns else "?"
    raw["Libellé marketing"] = raw.get("Libellé marketing", pd.Series("?", index=raw.index)).fillna("?")

    # Garder PGC uniquement
    PGC = {"BOISSONS","DROGUERIE","PARFUMERIE HYGIENE","EPICERIE"}
    if "Libellé rayon" in raw.columns:
        raw = raw[raw["Libellé rayon"].str.upper().isin(PGC)]

    return raw

@st.cache_data(show_spinner=False)
def load_topca(file_bytes):
    try:
        df = pd.read_csv(BytesIO(file_bytes), header=None, names=["Code article"], dtype=str)
    except Exception:
        df = pd.read_excel(BytesIO(file_bytes), header=None, names=["Code article"], dtype=str)
    df["Code article"] = norm_code(df["Code article"])
    return set(df["Code article"].dropna().unique())

# ─── CALCUL DÉTENTION ────────────────────────────────────────────────────────
def compute_detention(df_stock, top_codes):
    """
    Pour chaque article Top CA × magasin :
      - Récupère Code état, flux (IM/LO), stock
      - Taux détention = articles Code état=2 ET stock>0 / articles Code état=2
    """
    df = df_stock[df_stock["Code article"].isin(top_codes)].copy()

    # Agréger par article × magasin (garder le code état et le stock max)
    grp = df.groupby(["Code article","Libellé site","Code marketing","Libellé marketing"]).agg(
        code_etat       = ("Code etat",       lambda x: x.mode().iloc[0] if len(x) else "?"),
        stock           = ("Nouveau stock",   "sum"),
        ral             = ("Ral",             "sum"),
        nb_colis        = ("Nb colis",        "first"),
        lib_article     = ("Libellé article", "first"),
        lib_rayon       = ("Libellé rayon",   "first") if "Libellé rayon" in df.columns else ("Code article","first"),
        lib_fournisseur = ("Nom fourn.",      "first") if "Nom fourn." in df.columns else ("Code article","first"),
    ).reset_index()

    # Codes absents de toutes les extractions
    found = set(df["Code article"].unique())
    absents = sorted(top_codes - found)

    return grp, absents

def compute_taux(grp, top_codes):
    """Calcule les taux de détention par magasin et par flux — uniquement code état 2."""
    sites = sorted(grp["Libellé site"].unique())
    rows = []
    for site in sites:
        s = grp[grp["Libellé site"] == site]
        for flux in ["IM","LO","ALL"]:
            if flux == "ALL":
                sf = s
            else:
                sf = s[s["Code marketing"] == flux]

            actifs      = sf[sf["code_etat"] == "2"]
            n_actifs    = len(actifs)
            n_stock     = (actifs["stock"] > 0).sum()
            taux        = n_stock / n_actifs * 100 if n_actifs > 0 else None
            n_bloques   = (sf["code_etat"] == "B").sum()
            n_autres    = (sf["code_etat"].isin(["P","S","F","1","5","6"])).sum()
            n_rupture   = (actifs["stock"] <= 0).sum()
            n_faible    = ((actifs["stock"] > 0) & (actifs["stock"] < actifs["nb_colis"].replace(0, np.nan))).sum()

            rows.append({
                "site": site, "flux": flux,
                "n_top_ca": len(top_codes),
                "n_actifs": n_actifs,
                "n_stock_pos": int(n_stock),
                "taux": round(taux, 1) if taux is not None else None,
                "n_bloques": int(n_bloques),
                "n_autres_etats": int(n_autres),
                "n_rupture": int(n_rupture),
                "n_faible": int(n_faible),
            })
    return pd.DataFrame(rows)

def compute_alerte(row):
    if row["code_etat"] == "B":
        return "🔴 Bloqué"
    if row["code_etat"] == "F":
        return "⚪ Fin de vie"
    if row["code_etat"] not in ("2",):
        return f"🟡 État {row['code_etat']}"
    if row["stock"] <= 0 and row["ral"] <= 0:
        return "🛒 Rupture"
    if row["stock"] <= 0 and row["ral"] > 0:
        return "🚚 Relance"
    if row["nb_colis"] > 0 and row["stock"] < row["nb_colis"]:
        return "⚠️ Stock faible"
    return "✅ OK"

# ─── EXPORT EXCEL ────────────────────────────────────────────────────────────
def gen_excel(grp, taux_df, absents, top_codes):
    wb = Workbook()
    HDR_F = PatternFill("solid", fgColor="1C3557")
    HDR_T = Font(bold=True, color="FFFFFF", size=11)
    RED_F = PatternFill("solid", fgColor="FCE4E4")
    AMB_F = PatternFill("solid", fgColor="FEF3CD")
    GRN_F = PatternFill("solid", fgColor="D6F0D6")
    NEU_F = PatternFill("solid", fgColor="FFFFFF")
    CTR   = Alignment(horizontal="center", vertical="center")

    def write_ws(ws, headers, rows, title):
        ws.append([title]); ws.cell(1,1).font = Font(bold=True, size=13)
        ws.append([]); ws.append(headers)
        for i,h in enumerate(headers,1):
            c = ws.cell(3,i); c.fill=HDR_F; c.font=HDR_T; c.alignment=CTR
        for row in rows: ws.append(row)
        for col in ws.columns:
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(
                len(str(col[0].value or ""))+4, 12)

    # Onglet synthèse par magasin
    ws1 = wb.active; ws1.title = "Synthèse magasins"
    syn = taux_df[taux_df["flux"]=="ALL"].copy()
    rows1 = [[r["site"], len(top_codes), r["n_actifs"], r["n_stock_pos"],
              r["taux"], r["n_bloques"], r["n_autres_etats"], r["n_rupture"]]
             for _,r in syn.iterrows()]
    write_ws(ws1,
        ["Magasin","Réf Top CA","Actifs (état 2)","En stock","Taux %",
         "Bloqués (B)","Autres états","Ruptures"],
        rows1, f"Synthèse détention — {len(top_codes)} références Top CA")

    # Colorer taux
    for r in ws1.iter_rows(min_row=4, max_row=ws1.max_row):
        v = r[4].value
        if isinstance(v, (int,float)):
            r[4].fill = GRN_F if v >= 85 else AMB_F if v >= 70 else RED_F

    # Onglet IM vs LO
    ws2 = wb.create_sheet("IM vs LO")
    rows2 = []
    for _,r in taux_df[taux_df["flux"]!="ALL"].iterrows():
        rows2.append([r["site"],r["flux"],r["n_actifs"],r["n_stock_pos"],r["taux"],r["n_rupture"]])
    write_ws(ws2,
        ["Magasin","Flux","Actifs (état 2)","En stock","Taux %","Ruptures"],
        rows2, "Détention par flux IM / LO")
    for r in ws2.iter_rows(min_row=4, max_row=ws2.max_row):
        v = r[4].value
        if isinstance(v,(int,float)):
            r[4].fill = GRN_F if v>=85 else AMB_F if v>=70 else RED_F

    # Onglet Plan d'action
    ws3 = wb.create_sheet("Plan d'action")
    grp2 = grp.copy()
    grp2["Alerte"] = grp2.apply(compute_alerte, axis=1)
    urgences = grp2[grp2["Alerte"] != "✅ OK"].sort_values("Alerte")
    rows3 = [[r["Code article"],r["lib_article"],r["Libellé site"],
              r["Code marketing"],r["code_etat"],
              int(r["stock"]),int(r["ral"]),r["Alerte"]]
             for _,r in urgences.iterrows()]
    write_ws(ws3,
        ["Code","Libellé","Magasin","Flux","Code état","Stock","RAL","Alerte"],
        rows3, "Plan d'action — urgences détection")
    for r in ws3.iter_rows(min_row=4, max_row=ws3.max_row):
        v = str(r[7].value or "")
        r[7].fill = RED_F if "🔴" in v else AMB_F if "🛒" in v or "⚠️" in v or "🟡" in v else NEU_F

    # Onglet Absents ERP
    ws4 = wb.create_sheet("Absents ERP")
    write_ws(ws4,
        ["Code article","Statut"],
        [[c,"Absent de toutes les extractions"] for c in absents],
        "Références Top CA absentes des extractions ERP")

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
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
    f_topca  = st.file_uploader("Liste Top CA (CSV ou Excel)", type=["csv","xlsx"], key="topca")
    f_stocks = st.file_uploader("Extractions stock ERP (multi-CSV)",
                                 type=["csv"], accept_multiple_files=True, key="stocks")

    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px'>Paramètres</div>", unsafe_allow_html=True)
    cible_taux = st.slider("Cible taux de détention (%)", 70, 100, 85)
    st.markdown(f"<div style='font-size:12px;color:#007AFF;background:#E6F1FB;padding:7px 11px;border-radius:8px;border:0.5px solid #B3D9FF'>Cible : <strong>{cible_taux}%</strong></div>", unsafe_allow_html=True)

# ─── PAGE PRINCIPALE ──────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>📦 Détention Top CA</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Présence en magasin · Flux IM / LO · Code état 2 actif · 1 fichier par magasin</div>", unsafe_allow_html=True)

# ─── ÉCRAN D'ACCUEIL ─────────────────────────────────────────────────────────
if not f_topca or not f_stocks:
    st.markdown("---")
    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>ℹ️ À quoi sert ce module ?</strong><br>
  Ce module vérifie la <strong>présence en magasin des articles Top CA</strong> et calcule le taux de détention
  séparément pour les flux <strong>IM (Import)</strong> et <strong>LO (Local)</strong>.<br><br>
  Seuls les articles avec <strong>Code état = 2 (Actif)</strong> entrent dans le calcul.
  Les articles avec un autre code état sont signalés séparément par magasin.
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Règle de calcul du taux de détention</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:16px;border-left:3px solid #34C759'>
  <div style='font-size:13px;font-weight:600;color:#1C1C1E;margin-bottom:6px'>Inclus dans le calcul</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:1.8'>
    <span style='background:#F0FFF4;color:#34C759;padding:2px 8px;border-radius:100px;font-weight:600;font-size:11px'>Code état 2</span> — Article actif<br>
    Dénominateur : nb articles Top CA code état 2<br>
    Numérateur : nb articles Top CA code état 2 <strong>ET stock > 0</strong>
  </div>
</div>""", unsafe_allow_html=True)
    with c2:
        st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:16px;border-left:3px solid #FF3B30'>
  <div style='font-size:13px;font-weight:600;color:#1C1C1E;margin-bottom:6px'>Exclus du calcul — signalés</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:1.8'>
    <span style='background:#FFF2F2;color:#FF3B30;padding:2px 8px;border-radius:100px;font-weight:600;font-size:11px'>B</span> Bloqué · 
    <span style='background:#FFFBF0;color:#FF9500;padding:2px 8px;border-radius:100px;font-weight:600;font-size:11px'>P</span> Permanent<br>
    <span style='background:#FFFBF0;color:#FF9500;padding:2px 8px;border-radius:100px;font-weight:600;font-size:11px'>S</span> Saisonnier · 
    <span style='background:#F2F2F7;color:#8E8E93;padding:2px 8px;border-radius:100px;font-weight:600;font-size:11px'>F</span> Fin de vie<br>
    Un article peut être code 2 sur un magasin et B sur un autre.
  </div>
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Fichiers attendus</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
<div class='col-required'>
  <div style='font-size:16px'>📋</div>
  <div>
    <div class='col-name'>Liste Top CA</div>
    <div class='col-desc'>CSV ou Excel · 1 colonne · codes articles</div>
    <div class='col-example'>ex: 10002101, 14005975, 14006617…</div>
  </div>
</div>""", unsafe_allow_html=True)
    with c2:
        st.markdown("""
<div class='col-required'>
  <div style='font-size:16px'>🏪</div>
  <div>
    <div class='col-name'>Extractions stock ERP</div>
    <div class='col-desc'>1 CSV par magasin · séparateur ; · encodage UTF-8</div>
    <div class='col-example'>Extraction_stock_XXXXX_YYYYMMDD_GLOBAL.csv</div>
  </div>
</div>""", unsafe_allow_html=True)

    st.info("⬆️ Charge la liste Top CA et les extractions stock dans la sidebar pour lancer l'analyse.")
    st.stop()

# ─── TRAITEMENT ──────────────────────────────────────────────────────────────
with st.spinner("Lecture des fichiers…"):
    top_codes = load_topca(f_topca.read())
    files_bn  = tuple((f.read(), f.name) for f in f_stocks)
    df_stock  = load_stock(files_bn)

if df_stock.empty:
    st.error("Aucune donnée PGC lue depuis les extractions stock."); st.stop()
if not top_codes:
    st.error("Liste Top CA vide ou illisible."); st.stop()

with st.spinner("Calcul des taux de détention…"):
    grp, absents = compute_detention(df_stock, top_codes)
    grp["Alerte"] = grp.apply(compute_alerte, axis=1)
    taux_df = compute_taux(grp, top_codes)

n_sites   = df_stock["Libellé site"].nunique()
taux_all  = taux_df[taux_df["flux"]=="ALL"]
taux_im   = taux_df[taux_df["flux"]=="IM"]
taux_lo   = taux_df[taux_df["flux"]=="LO"]
taux_moy  = taux_all["taux"].mean()
taux_im_m = taux_im["taux"].mean()
taux_lo_m = taux_lo["taux"].mean()
n_urgences= (grp["Alerte"] != "✅ OK").sum()

# ─── KPIs ────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-label'>Indicateurs globaux · " + str(n_sites) + " magasin(s) · " + str(len(top_codes)) + " références Top CA</div>", unsafe_allow_html=True)
k1,k2,k3,k4,k5 = st.columns(5)
k1.metric("Réf Top CA",        str(len(top_codes)))
k2.metric("Taux détention moy",
          f"{taux_moy:.1f}%" if taux_moy else "—",
          f"cible {cible_taux}%")
k3.metric("Taux IM (Import)",
          f"{taux_im_m:.1f}%" if taux_im_m else "—",
          f"{taux_im_m-cible_taux:+.1f} pt vs cible" if taux_im_m else "")
k4.metric("Taux LO (Local)",
          f"{taux_lo_m:.1f}%" if taux_lo_m else "—",
          f"{taux_lo_m-cible_taux:+.1f} pt vs cible" if taux_lo_m else "")
k5.metric("Urgences", str(n_urgences), "articles à traiter")

# ─── ALERTES ─────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Alertes & actions prioritaires</div>", unsafe_allow_html=True)

sites_sous_cible = taux_all[taux_all["taux"] < cible_taux].sort_values("taux")
if not sites_sous_cible.empty:
    liste = ", ".join([f"{r['site']} ({r['taux']:.0f}%)" for _,r in sites_sous_cible.iterrows()])
    st.markdown(f"""
<div class='alert-card alert-red'>
  <strong>⚠️ {len(sites_sous_cible)} magasin(s) sous la cible {cible_taux}%</strong> — {liste}<br>
  <span style='font-size:12px;opacity:.85'>→ Vérifier les commandes en attente et les articles bloqués.</span>
</div>""", unsafe_allow_html=True)

n_bloques = (grp["code_etat"] == "B").sum()
if n_bloques > 0:
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ {n_bloques} ligne(s) Top CA avec code état B (Bloqué)</strong><br>
  <span style='font-size:12px;opacity:.85'>→ Ces articles sont exclus du taux de détention. Débloquer ou substituer.</span>
</div>""", unsafe_allow_html=True)

im_sous = taux_im[taux_im["taux"] < cible_taux]
if not im_sous.empty:
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ Flux Import sous cible sur {len(im_sous)} magasin(s)</strong> — taux IM moyen {taux_im_m:.1f}%<br>
  <span style='font-size:12px;opacity:.85'>→ Délai réapprovisionnement IM : 4–8 semaines. Anticiper les commandes.</span>
</div>""", unsafe_allow_html=True)

if len(absents) > 0:
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ {len(absents)} référence(s) Top CA absentes de toutes les extractions</strong><br>
  <span style='font-size:12px;opacity:.85'>→ Vérifier le référentiel article ou si déréférencement non planifié.</span>
</div>""", unsafe_allow_html=True)

if n_urgences == 0 and sites_sous_cible.empty:
    st.success("✅ Tous les magasins sont au-dessus de la cible — aucune urgence détectée.")

# ─── TABS ─────────────────────────────────────────────────────────────────────
st.markdown("---")
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Synthèse réseau", "🔄 IM vs LO", "🚨 Plan d'action", "🚫 Absents ERP"
])

# ═══ TAB 1 ═══════════════════════════════════════════════════════════════════
with tab1:
    disp1 = taux_all[["site","n_top_ca","n_actifs","n_stock_pos","taux",
                       "n_bloques","n_autres_etats","n_rupture"]].copy()
    disp1.columns = ["Magasin","Réf Top CA","Actifs (état 2)",
                     "En stock","Taux %","Bloqués (B)","Autres états","Ruptures"]
    disp1 = disp1.sort_values("Taux %")
    disp1["Taux %"] = disp1["Taux %"].apply(lambda x: f"{x:.1f}%" if x is not None else "—")
    disp1["Statut"] = taux_all.sort_values("taux")["taux"].apply(
        lambda x: "🟢 OK" if x is not None and x >= cible_taux
        else ("🟡 Surveiller" if x is not None and x >= cible_taux-10 else "🔴 Action requise"))
    st.dataframe(disp1, use_container_width=True, hide_index=True)

    # Mini graphique taux par magasin
    try:
        import plotly.graph_objects as go
        fig = go.Figure(go.Bar(
            x=taux_all.sort_values("taux")["taux"].tolist(),
            y=taux_all.sort_values("taux")["site"].tolist(),
            orientation="h",
            marker_color=[
                "#34C759" if (v or 0) >= cible_taux
                else "#FF9500" if (v or 0) >= cible_taux-10
                else "#FF3B30"
                for v in taux_all.sort_values("taux")["taux"]
            ],
            marker_line_width=0,
            text=[f"{v:.1f}%" if v else "—" for v in taux_all.sort_values("taux")["taux"]],
            textposition="outside",
        ))
        fig.add_vline(x=cible_taux, line_width=1.5, line_dash="dash", line_color="#007AFF",
                      annotation_text=f"Cible {cible_taux}%", annotation_position="top right")
        fig.update_layout(
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="-apple-system, Helvetica Neue", color="#3A3A3C", size=12),
            height=max(280, n_sites*48+80),
            margin=dict(t=20,b=20,l=10,r=80),
            xaxis=dict(title="Taux de détention (%)", ticksuffix="%",
                       showgrid=True, gridcolor="#F2F2F7", range=[0,110]),
            yaxis=dict(showgrid=False, title=""),
        )
        st.plotly_chart(fig, use_container_width=True)
    except ImportError:
        pass

# ═══ TAB 2 ═══════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>ℹ️ Lecture des flux</strong> — Le taux IM concerne les articles approvisionnés en Import (délai 4–8 sem.).
  Le taux LO concerne les articles approvisionnés en Local (réassort 48h). Un écart important entre les deux
  signale des tensions d'import à anticiper.
</div>""", unsafe_allow_html=True)

    # Tableau pivot IM vs LO par magasin
    pivot = taux_df[taux_df["flux"]!="ALL"].pivot_table(
        index="site", columns="flux",
        values=["n_actifs","n_stock_pos","taux","n_rupture"],
        aggfunc="first"
    ).reset_index()
    pivot.columns = ["Magasin",
                     "Actifs IM","Actifs LO",
                     "Ruptures IM","Ruptures LO",
                     "En stock IM","En stock LO",
                     "Taux IM %","Taux LO %"]

    # Réordonner
    pivot = pivot[["Magasin","Actifs IM","En stock IM","Taux IM %",
                   "Actifs LO","En stock LO","Taux LO %"]].sort_values("Taux IM %")
    pivot["Taux IM %"] = pivot["Taux IM %"].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "—")
    pivot["Taux LO %"] = pivot["Taux LO %"].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "—")
    st.dataframe(pivot, use_container_width=True, hide_index=True)

    # Totaux IM / LO
    c1,c2 = st.columns(2)
    with c1:
        tot_im = taux_im["n_actifs"].sum(); pres_im = taux_im["n_stock_pos"].sum()
        taux_im_global = pres_im/tot_im*100 if tot_im else 0
        color_im = "#34C759" if taux_im_global>=cible_taux else "#FF9500" if taux_im_global>=cible_taux-10 else "#FF3B30"
        st.markdown(f"""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:16px'>
  <div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px'>Flux IM · Import</div>
  <div style='font-size:28px;font-weight:700;color:{color_im};letter-spacing:-.02em'>{taux_im_global:.1f}%</div>
  <div style='font-size:12px;color:#8E8E93;margin-top:4px'>{pres_im} / {tot_im} lignes actives en stock · {n_sites} magasins</div>
  <div style='background:var(--color-border-tertiary,#E5E5EA);border-radius:3px;height:6px;margin-top:10px'>
    <div style='width:{min(taux_im_global,100):.0f}%;background:{color_im};height:6px;border-radius:3px'></div>
  </div>
</div>""", unsafe_allow_html=True)
    with c2:
        tot_lo = taux_lo["n_actifs"].sum(); pres_lo = taux_lo["n_stock_pos"].sum()
        taux_lo_global = pres_lo/tot_lo*100 if tot_lo else 0
        color_lo = "#34C759" if taux_lo_global>=cible_taux else "#FF9500" if taux_lo_global>=cible_taux-10 else "#FF3B30"
        st.markdown(f"""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:16px'>
  <div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px'>Flux LO · Local</div>
  <div style='font-size:28px;font-weight:700;color:{color_lo};letter-spacing:-.02em'>{taux_lo_global:.1f}%</div>
  <div style='font-size:12px;color:#8E8E93;margin-top:4px'>{pres_lo} / {tot_lo} lignes actives en stock · {n_sites} magasins</div>
  <div style='background:var(--color-border-tertiary,#E5E5EA);border-radius:3px;height:6px;margin-top:10px'>
    <div style='width:{min(taux_lo_global,100):.0f}%;background:{color_lo};height:6px;border-radius:3px'></div>
  </div>
</div>""", unsafe_allow_html=True)

# ═══ TAB 3 ═══════════════════════════════════════════════════════════════════
with tab3:
    urgences = grp[grp["Alerte"] != "✅ OK"].copy()

    fc1,fc2,fc3 = st.columns(3)
    with fc1:
        sites_filtre = ["Tous"] + sorted(grp["Libellé site"].unique())
        sel_site = st.selectbox("Magasin", sites_filtre)
    with fc2:
        flux_filtre = ["Tous","IM","LO"]
        sel_flux = st.selectbox("Flux", flux_filtre)
    with fc3:
        alertes_dispo = sorted(urgences["Alerte"].unique())
        sel_alerte = st.multiselect("Alerte", alertes_dispo, default=alertes_dispo)

    if sel_site != "Tous":    urgences = urgences[urgences["Libellé site"]==sel_site]
    if sel_flux != "Tous":    urgences = urgences[urgences["Code marketing"]==sel_flux]
    if sel_alerte:            urgences = urgences[urgences["Alerte"].isin(sel_alerte)]

    st.markdown(f"<div style='font-size:12px;color:#8E8E93;margin-bottom:8px'>{len(urgences)} article(s) nécessitant une action</div>", unsafe_allow_html=True)

    if urgences.empty:
        st.success("✅ Aucune urgence sur la sélection.")
    else:
        disp3 = urgences[["Code article","lib_article","Libellé site","Code marketing",
                           "code_etat","stock","ral","Alerte"]].copy()
        disp3.columns = ["Code","Libellé","Magasin","Flux","Code état","Stock","RAL","Alerte"]
        disp3["Stock"] = disp3["Stock"].apply(lambda x: int(x))
        disp3["RAL"]   = disp3["RAL"].apply(lambda x: int(x))
        st.dataframe(disp3.sort_values("Alerte"), use_container_width=True, hide_index=True,
                     column_config={
                         "Flux": st.column_config.TextColumn("Flux",   width="small"),
                         "Code état": st.column_config.TextColumn("État", width="small"),
                     })

# ═══ TAB 4 ═══════════════════════════════════════════════════════════════════
with tab4:
    if not absents:
        st.success("✅ Toutes les références Top CA sont présentes dans au moins une extraction ERP.")
    else:
        st.warning(f"⚠️ {len(absents)} référence(s) Top CA absentes de toutes les extractions")
        df_abs = pd.DataFrame({"Code article": absents, "Statut": "Absent ERP", "Action": "Vérifier référentiel ou déréférencement"})
        st.dataframe(df_abs, use_container_width=True, hide_index=True)

# ─── EXPORT ──────────────────────────────────────────────────────────────────
st.markdown("---")
with st.expander("📥 Export Excel — Synthèse · IM vs LO · Plan d'action · Absents ERP"):
    if st.button("Générer le fichier Excel", type="primary"):
        with st.spinner("Génération…"):
            buf = gen_excel(grp, taux_df, absents, top_codes)
        st.download_button(
            "⬇️ Télécharger",
            data=buf,
            file_name=f"SmartBuyer_Detention_TopCA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
