"""
03_📦_Detention_Top_CA.py — SmartBuyer Hub
Taux de détention Top CA · Articles Permanents · Flux IM/LO · Stock immobilisé code B
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
.alert-card  { padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; font-size: 13px; line-height: 1.5; border-left: 3px solid; }
.alert-red   { background: #FFF2F2; border-color: #FF3B30; color: #3A0000; }
.alert-amber { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }
.alert-green { background: #F0FFF4; border-color: #34C759; color: #003A10; }
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }
.kpi-immo { background: #FFF2F2; border: 1px solid #FFB3AE; border-radius: 12px; padding: 16px 18px; }
.kpi-immo-label { font-size: 11px; font-weight: 500; color: #FF3B30; text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 3px; }
.kpi-immo-value { font-size: 24px; font-weight: 700; color: #FF3B30; letter-spacing: -0.02em; }
.kpi-immo-sub   { font-size: 12px; color: #C62828; margin-top: 3px; font-weight: 500; }
.col-required { background: #F0F8FF; border: 0.5px solid #B3D9FF; border-radius: 8px; padding: 10px 14px; margin-bottom: 6px; display: flex; align-items: flex-start; gap: 10px; }
.col-name  { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc  { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
.col-ex    { font-size: 11px; color: #8E8E93; font-family: monospace; margin-top: 2px; }
.gauge-bg  { background: #E5E5EA; border-radius: 3px; height: 5px; margin-top: 10px; }
.gauge-fill { height: 5px; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ─────────────────────────────────────────────────────────────────
def fmt(n, suffix=""):
    if pd.isna(n): return "—"
    a = abs(n)
    s = f"{n/1_000_000:.1f} M" if a >= 1_000_000 else f"{int(n/1_000)} K" if a >= 1_000 else f"{int(n):,}"
    return s + suffix

def norm_code(s):
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(8)

def color_taux(v, cible):
    if v is None: return "#8E8E93"
    return "#34C759" if v >= cible else "#FF9500" if v >= cible - 10 else "#FF3B30"

def badge_etat(code):
    MAP = {"2": ("#E6F1FB","#007AFF"), "B": ("#FFF2F2","#FF3B30"),
           "P": ("#F2F2F7","#8E8E93"), "S": ("#FFFBF0","#FF9500"),
           "F": ("#F2F2F7","#8E8E93")}
    bg, fg = MAP.get(code, ("#F2F2F7","#8E8E93"))
    labels = {"2":"Actif","B":"Bloqué","P":"Permanent","S":"Saisonnier","F":"Fin de vie"}
    return f"<span style='background:{bg};color:{fg};padding:2px 8px;border-radius:100px;font-size:10px;font-weight:600'>{code} · {labels.get(code,code)}</span>"

# ─── PARSING ──────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_stock(file_bytes, file_name):
    """
    Accepte :
      - 1 fichier consolidé  : stock_consolide_YYYYMMDD.csv  (produit par consolider_stock.py)
      - 1 fichier individuel : Extraction_stock_XXXXX_GLOBAL.csv
    Séparateur ; · encodage UTF-8
    """
    try:
        raw = pd.read_csv(BytesIO(file_bytes), sep=";", encoding="utf-8-sig",
                          dtype=str, low_memory=False)
    except Exception as e:
        st.error(f"Erreur lecture {file_name} : {e}")
        return pd.DataFrame()

    num_cols = ["Nouveau stock","Ral","Nb colis","Prix d'achat","PMP","Prix de vente"]
    for col in num_cols:
        if col in raw.columns:
            raw[col] = pd.to_numeric(raw[col], errors="coerce").fillna(0)

    raw["Code article"]      = norm_code(raw["Code article"])
    raw["Code etat"]         = raw["Code etat"].astype(str).str.strip().str.upper()
    raw["Code marketing"]    = raw.get("Code marketing",    pd.Series("?", index=raw.index)).astype(str).str.strip().str.upper()
    raw["Type saisonnalité"] = raw.get("Type saisonnalité", pd.Series("?", index=raw.index)).astype(str).str.strip().str.upper()

    # Filtre PGC (au cas où le fichier contient d'autres rayons)
    PGC = {"BOISSONS","DROGUERIE","PARFUMERIE HYGIENE","EPICERIE"}
    if "Libellé rayon" in raw.columns:
        raw = raw[raw["Libellé rayon"].str.upper().isin(PGC)]

    return raw

@st.cache_data(show_spinner=False)
def load_topca(byt, fname=""):
    """
    Lit la liste Top CA.
    Accepte CSV avec ou sans en-tête, séparateur , ou ;
    Prend toujours la 1ère colonne comme code article.
    """
    try:
        raw = byt.decode("utf-8-sig", errors="replace")
        sep = ";" if raw.count(";") > raw.count(",") else ","
        df  = pd.read_csv(BytesIO(byt), sep=sep, encoding="utf-8-sig", dtype=str)
        # Prendre la 1ère colonne quelle que soit son nom
        df  = df.iloc[:, [0]].copy()
        df.columns = ["Code article"]
    except Exception as e:
        st.error(f"Erreur lecture liste Top CA : {e}")
        return set()
    df["Code article"] = norm_code(df["Code article"])
    return set(df["Code article"].dropna().unique())

# ─── CALCUL PRINCIPAL ────────────────────────────────────────────────────────
def compute(df_stock, top_codes):
    """
    Filtre : articles Top CA × Type saisonnalité = P (Permanent)
    Taux   : Code état = 2 uniquement
    Immo   : Code état = B × PMP
    """
    df = df_stock[df_stock["Code article"].isin(top_codes)].copy()

    # Agréger par article × magasin
    agg_cols = {
        "code_etat":        ("Code etat",        lambda x: x.mode().iloc[0] if len(x) else "?"),
        "saisonnalite":     ("Type saisonnalité", lambda x: x.mode().iloc[0] if len(x) else "?"),
        "flux":             ("Code marketing",    lambda x: x.mode().iloc[0] if len(x) else "?"),
        "stock":            ("Nouveau stock",     "sum"),
        "ral":              ("Ral",               "sum"),
        "nb_colis":         ("Nb colis",          "first"),
        "pmp":              ("PMP",               "first"),
        "lib_article":      ("Libellé article",   "first"),
        "lib_rayon":        ("Libellé rayon",     "first") if "Libellé rayon"    in df.columns else ("Code article","first"),
        "lib_fournisseur":  ("Nom fourn.",        "first") if "Nom fourn."       in df.columns else ("Code article","first"),
    }
    grp = df.groupby(["Code article","Libellé site"]).agg(**agg_cols).reset_index()

    # Valeur stock immobilisé (code B)
    grp["stock_immo"] = np.where(grp["code_etat"]=="B", grp["stock"] * grp["pmp"], 0)

    # Filtre Permanent pour les calculs de taux
    grp_perm = grp[grp["saisonnalite"]=="P"].copy()

    # Articles absents
    found   = set(df["Code article"].unique())
    absents = sorted(top_codes - found)

    # Nombre d'articles permanents dans le Top CA (trouvés)
    n_perm_top = grp_perm["Code article"].nunique()

    return grp, grp_perm, absents, n_perm_top

def compute_taux(grp_perm, top_codes, cible):
    """Taux de détention par magasin et flux — Code état 2 uniquement — Permanents uniquement."""
    sites = sorted(grp_perm["Libellé site"].unique())
    rows  = []
    for site in sites:
        s = grp_perm[grp_perm["Libellé site"]==site]
        for flux in ["ALL","IM","LO"]:
            sf      = s if flux=="ALL" else s[s["flux"]==flux]
            actifs  = sf[sf["code_etat"]=="2"]
            n_act   = len(actifs)
            n_stk   = int((actifs["stock"]>0).sum())
            taux    = round(n_stk/n_act*100, 1) if n_act > 0 else None
            n_blq   = int((sf["code_etat"]=="B").sum())
            immo    = sf["stock_immo"].sum()
            n_rupt  = int((actifs["stock"]<=0).sum())
            n_fbl   = int(((actifs["stock"]>0) & (actifs["stock"]<actifs["nb_colis"].replace(0,np.nan))).sum())
            rows.append({
                "site":site,"flux":flux,
                "n_perm": len(sf),
                "n_actifs":n_act,"n_stock":n_stk,"taux":taux,
                "n_bloques":n_blq,"stock_immo":round(immo),
                "n_rupture":n_rupt,"n_faible":n_fbl,
                "sous_cible": taux is not None and taux < cible,
            })
    return pd.DataFrame(rows)

def compute_alerte(row):
    if row["code_etat"]=="B":    return "🔴 Bloqué"
    if row["code_etat"]=="F":    return "⚪ Fin de vie"
    if row["code_etat"]!="2":    return f"🟡 État {row['code_etat']}"
    if row["stock"]<=0 and row["ral"]<=0: return "🛒 Rupture"
    if row["stock"]<=0 and row["ral"]>0:  return "🚚 Relance"
    if row["nb_colis"]>0 and 0<row["stock"]<row["nb_colis"]: return "⚠️ Stock faible"
    return "✅ OK"

# ─── EXPORT EXCEL ────────────────────────────────────────────────────────────
def gen_excel(grp, grp_perm, taux_df, absents, top_codes):
    wb    = Workbook()
    HDR_F = PatternFill("solid", fgColor="1C3557")
    HDR_T = Font(bold=True, color="FFFFFF", size=11)
    RED_F = PatternFill("solid", fgColor="FCE4E4")
    AMB_F = PatternFill("solid", fgColor="FEF3CD")
    GRN_F = PatternFill("solid", fgColor="D6F0D6")
    NEU_F = PatternFill("solid", fgColor="FFFFFF")
    ORG_F = PatternFill("solid", fgColor="FFE5CC")
    CTR   = Alignment(horizontal="center", vertical="center")

    def write_ws(ws, headers, rows, title, col_widths=None):
        ws.append([title]); ws.cell(1,1).font = Font(bold=True, size=13)
        ws.append([]); ws.append(headers)
        for i,h in enumerate(headers,1):
            c = ws.cell(3,i); c.fill=HDR_F; c.font=HDR_T; c.alignment=CTR
        for row in rows: ws.append(row)
        widths = col_widths or {}
        for ci, col in enumerate(ws.iter_cols(min_row=1, max_row=1), 1):
            hdr = str(ws.cell(3,ci).value or "")
            ws.column_dimensions[get_column_letter(ci)].width = widths.get(ci, max(len(hdr)+4, 12))

    # ── Onglet 1 : Synthèse magasins ─────────────────────────────────────────
    ws1 = wb.active; ws1.title = "Synthèse magasins"
    syn = taux_df[taux_df["flux"]=="ALL"].copy()
    rows1 = []
    for _,r in syn.iterrows():
        rows1.append([r["site"], len(top_codes), r["n_perm"], r["n_actifs"],
                      r["n_stock"], r["taux"], r["n_bloques"],
                      r["stock_immo"], r["n_rupture"]])
    write_ws(ws1,
        ["Magasin","Réf Top CA","Permanents","Actifs (état 2)",
         "En stock","Taux %","Bloqués (B)","Stock immobilisé FCFA","Ruptures"],
        rows1, f"Synthèse détention — {len(top_codes)} réf. Top CA · Permanents uniquement · Code état 2",
        {8: 24})
    for r in ws1.iter_rows(min_row=4, max_row=ws1.max_row):
        v = r[5].value  # Taux %
        if isinstance(v,(int,float)):
            r[5].fill = GRN_F if v>=85 else AMB_F if v>=70 else RED_F
        vi = r[7].value  # Stock immo
        if isinstance(vi,(int,float)) and vi > 0:
            r[7].fill = ORG_F; r[7].font = Font(bold=True, color="C62828")
            r[7].number_format = "#,##0"

    # ── Onglet 2 : IM vs LO ──────────────────────────────────────────────────
    ws2 = wb.create_sheet("IM vs LO")
    rows2 = []
    pivot = taux_df[taux_df["flux"]!="ALL"]
    for _,r in pivot.iterrows():
        rows2.append([r["site"], r["flux"], r["n_perm"], r["n_actifs"],
                      r["n_stock"], r["taux"], r["n_bloques"], r["stock_immo"], r["n_rupture"]])
    write_ws(ws2,
        ["Magasin","Flux","Permanents","Actifs (état 2)",
         "En stock","Taux %","Bloqués (B)","Stock immobilisé FCFA","Ruptures"],
        rows2, "Détention par flux IM / LO · Permanents · Code état 2", {8:24})
    for r in ws2.iter_rows(min_row=4, max_row=ws2.max_row):
        v = r[5].value
        if isinstance(v,(int,float)):
            r[5].fill = GRN_F if v>=85 else AMB_F if v>=70 else RED_F
        vi = r[7].value
        if isinstance(vi,(int,float)) and vi>0:
            r[7].fill = ORG_F; r[7].number_format="#,##0"

    # ── Onglet 3 : Plan d'action ──────────────────────────────────────────────
    ws3 = wb.create_sheet("Plan d'action")
    grp2 = grp.copy()
    grp2["Alerte"] = grp2.apply(compute_alerte, axis=1)
    urgences = grp2[grp2["Alerte"]!="✅ OK"].sort_values("Alerte")
    rows3 = [[r["Code article"],r["lib_article"],r["Libellé site"],
              r["flux"],r["code_etat"],r["saisonnalite"],
              int(r["stock"]),int(r["ral"]),
              int(r["stock_immo"]) if r["stock_immo"]>0 else "",
              r["Alerte"]]
             for _,r in urgences.iterrows()]
    write_ws(ws3,
        ["Code","Libellé","Magasin","Flux","Code état","Saisonnalité",
         "Stock","RAL","Stock immobilisé FCFA","Alerte"],
        rows3, "Plan d'action · Urgences détection", {2:36, 3:22, 9:24})
    for r in ws3.iter_rows(min_row=4, max_row=ws3.max_row):
        v = str(r[9].value or "")
        r[9].fill = RED_F if "🔴" in v else AMB_F if any(x in v for x in ["🛒","⚠️","🟡"]) else NEU_F
        vi = r[8].value
        if isinstance(vi,int) and vi>0:
            r[8].fill=ORG_F; r[8].number_format="#,##0"

    # ── Onglet 4 : Stock immobilisé détail ────────────────────────────────────
    ws4 = wb.create_sheet("Stock immobilisé (B)")
    immo_df = grp[grp["code_etat"]=="B"].sort_values("stock_immo", ascending=False)
    rows4 = [[r["Code article"],r["lib_article"],r["Libellé site"],
              r["flux"],r["saisonnalite"],int(r["stock"]),
              round(r["pmp"]),int(r["stock_immo"])]
             for _,r in immo_df.iterrows()]
    write_ws(ws4,
        ["Code","Libellé","Magasin","Flux","Saisonnalité",
         "Stock qté","PMP (FCFA)","Stock immobilisé FCFA"],
        rows4, "Détail stock immobilisé · Code état B · tous articles", {2:36,8:24})
    for r in ws4.iter_rows(min_row=4, max_row=ws4.max_row):
        vi = r[7].value
        if isinstance(vi,(int,float)) and vi>0:
            r[7].fill=ORG_F; r[7].font=Font(bold=True,color="C62828")
            r[7].number_format="#,##0"

    # ── Onglet 5 : Absents ERP ────────────────────────────────────────────────
    ws5 = wb.create_sheet("Absents ERP")
    write_ws(ws5,
        ["Code article","Statut","Action"],
        [[c,"Absent ERP","Vérifier référentiel ou déréférencement"] for c in absents],
        "Références Top CA absentes des extractions ERP")

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
    f_stocks = st.file_uploader(
        "Stock consolidé (CSV)",
        type=["csv"], key="stocks",
        help="Fichier produit par consolider_stock.py · séparateur ; · UTF-8"
    )
    st.markdown("---")

    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Paramètres</div>", unsafe_allow_html=True)
    cible = st.slider("Cible taux de détention (%)", 70, 100, 85)
    st.markdown(f"""
<div style='background:#E6F1FB;border:0.5px solid #B3D9FF;border-radius:8px;padding:7px 11px;font-size:12px;color:#001A3A;margin-bottom:10px'>
  Cible : <strong>{cible}%</strong> · Articles <strong>Permanents</strong> uniquement<br>
  Code état <strong>2</strong> dans le calcul · PMP pour la valorisation
</div>""", unsafe_allow_html=True)

# ─── PAGE PRINCIPALE ──────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>📦 Détention Top CA</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Articles Permanents · Flux IM / LO · Code état 2 · Stock immobilisé code B · Fichier consolidé</div>", unsafe_allow_html=True)

# ─── ÉCRAN D'ACCUEIL ─────────────────────────────────────────────────────────
if not f_topca or not f_stocks:
    st.markdown("---")
    st.markdown("""
<div class='alert-card alert-blue'>
  <strong>ℹ️ À quoi sert ce module ?</strong><br>
  Ce module vérifie la présence en magasin des <strong>articles Top CA permanents</strong>
  et calcule le taux de détention séparément pour les flux <strong>IM (Import)</strong> et <strong>LO (Local)</strong>.<br><br>
  Il identifie également le <strong>stock immobilisé en code état B</strong> (articles bloqués)
  valorisé au PMP — capital qui ne génère aucun CA.
</div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='section-label'>Règles de calcul</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:16px;border-left:3px solid #34C759;margin-bottom:10px'>
  <div style='font-size:13px;font-weight:600;color:#1C1C1E;margin-bottom:8px'>Inclus dans le taux</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:2'>
    <span style='background:#F0FFF4;color:#34C759;padding:2px 8px;border-radius:100px;font-weight:600;font-size:11px'>Type saisonnalité = P</span> Permanent<br>
    <span style='background:#E6F1FB;color:#007AFF;padding:2px 8px;border-radius:100px;font-weight:600;font-size:11px'>Code état = 2</span> Actif<br>
    Taux = articles en stock / articles actifs permanents
  </div>
</div>""", unsafe_allow_html=True)
    with c2:
        st.markdown("""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-radius:12px;padding:16px;border-left:3px solid #FF3B30;margin-bottom:10px'>
  <div style='font-size:13px;font-weight:600;color:#1C1C1E;margin-bottom:8px'>Exclus — signalés</div>
  <div style='font-size:12px;color:#3A3A3C;line-height:2'>
    <span style='background:#FFF2F2;color:#FF3B30;padding:2px 8px;border-radius:100px;font-weight:600;font-size:11px'>B · Bloqué</span> → stock immobilisé valorisé au PMP<br>
    <span style='background:#FFFBF0;color:#FF9500;padding:2px 8px;border-radius:100px;font-weight:600;font-size:11px'>S · Saisonnier</span> → exclu du calcul<br>
    <span style='background:#F2F2F7;color:#8E8E93;padding:2px 8px;border-radius:100px;font-weight:600;font-size:11px'>F · Fin de vie</span> → signalé séparément
  </div>
</div>""", unsafe_allow_html=True)

    st.markdown("<div class='section-label'>Fichiers attendus</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
<div class='col-required'><div style='font-size:16px'>📋</div>
<div><div class='col-name'>Liste Top CA</div>
<div class='col-desc'>CSV ou Excel · 1 colonne · codes articles sans en-tête</div>
<div class='col-ex'>ex: 10002101 / 14005975 / 14006617…</div></div></div>""", unsafe_allow_html=True)
    with c2:
        st.markdown("""
<div class='col-required'><div style='font-size:16px'>🏪</div>
<div><div class='col-name'>Stock consolidé ERP</div>
<div class='col-desc'>1 seul CSV · produit par <strong>consolider_stock.py</strong> · séparateur ; · UTF-8</div>
<div class='col-ex'>stock_consolide_YYYYMMDD.csv · tous magasins</div></div></div>""", unsafe_allow_html=True)

    st.info("⬆️ Charge la liste Top CA et les extractions stock dans la sidebar pour lancer l'analyse.")
    st.stop()

# ─── TRAITEMENT ───────────────────────────────────────────────────────────────
with st.spinner("Lecture des fichiers…"):
    top_codes = load_topca(f_topca.read(), f_topca.name)
    stock_bytes = f_stocks.read()
    df_stock  = load_stock(stock_bytes, f_stocks.name)

if df_stock.empty:
    st.error("Aucune donnée PGC lue."); st.stop()
if not top_codes:
    st.error("Liste Top CA vide."); st.stop()

with st.spinner("Calcul des taux de détention…"):
    grp, grp_perm, absents, n_perm_top = compute(df_stock, top_codes)
    grp["Alerte"]      = grp.apply(compute_alerte, axis=1)
    grp_perm["Alerte"] = grp_perm.apply(compute_alerte, axis=1)
    taux_df = compute_taux(grp_perm, top_codes, cible)

n_sites       = df_stock["Libellé site"].nunique()
taux_all      = taux_df[taux_df["flux"]=="ALL"]
taux_im       = taux_df[taux_df["flux"]=="IM"]
taux_lo       = taux_df[taux_df["flux"]=="LO"]
taux_moy      = taux_all["taux"].mean()
taux_im_moy   = taux_im["taux"].mean()
taux_lo_moy   = taux_lo["taux"].mean()
stock_immo_total = grp["stock_immo"].sum()
n_urgences    = (grp_perm["Alerte"]!="✅ OK").sum()
n_saisonniers = len(top_codes) - n_perm_top

# ─── KPIs ─────────────────────────────────────────────────────────────────────
st.markdown(f"<div class='section-label'>{n_sites} magasin(s) · {len(top_codes)} références Top CA · {n_perm_top} permanentes</div>", unsafe_allow_html=True)

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Réf permanentes",   str(n_perm_top),
          f"{n_saisonniers} saisonniers exclus" if n_saisonniers else "")
k2.metric("Taux détention moy",
          f"{taux_moy:.1f}%" if taux_moy else "—",
          f"{taux_moy-cible:+.1f} pt vs cible" if taux_moy else "")
k3.metric("Taux IM · Import",
          f"{taux_im_moy:.1f}%" if taux_im_moy else "—",
          f"{taux_im_moy-cible:+.1f} pt vs cible" if taux_im_moy else "")
k4.metric("Taux LO · Local",
          f"{taux_lo_moy:.1f}%" if taux_lo_moy else "—",
          f"{taux_lo_moy-cible:+.1f} pt vs cible" if taux_lo_moy else "")

# KPI Stock immobilisé — mis en avant avec style rouge
with k5:
    st.markdown(f"""
<div class='kpi-immo'>
  <div class='kpi-immo-label'>Stock immobilisé (B)</div>
  <div class='kpi-immo-value'>{fmt(stock_immo_total)}</div>
  <div class='kpi-immo-sub'>FCFA · code B · PMP</div>
</div>""", unsafe_allow_html=True)

# ─── ALERTES ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='section-label'>Alertes &amp; actions prioritaires</div>", unsafe_allow_html=True)

if stock_immo_total > 0:
    st.markdown(f"""
<div class='alert-card alert-red'>
  <strong>⚠️ {fmt(stock_immo_total)} FCFA immobilisés en code état B</strong> sur {n_sites} magasin(s)<br>
  <span style='font-size:12px;opacity:.85'>→ Ce stock bloqué ne génère aucun CA. Débloquer, liquider ou substituer en priorité.</span>
</div>""", unsafe_allow_html=True)

sites_sous = taux_all[taux_all["taux"]<cible].sort_values("taux")
if not sites_sous.empty:
    liste = ", ".join([f"{r['site']} ({r['taux']:.0f}%)" for _,r in sites_sous.iterrows()])
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ {len(sites_sous)} magasin(s) sous la cible {cible}%</strong> — {liste}<br>
  <span style='font-size:12px;opacity:.85'>→ Vérifier commandes IM en attente. Délai réappro Import : 4–8 semaines.</span>
</div>""", unsafe_allow_html=True)

im_sous = taux_im[taux_im["taux"]<cible]
if not im_sous.empty and len(im_sous) > len(sites_sous):
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ Flux Import sous cible sur {len(im_sous)} magasin(s)</strong> — taux IM moyen {taux_im_moy:.1f}%<br>
  <span style='font-size:12px;opacity:.85'>→ Anticiper les commandes Import. Le flux LO peut pallier en attendant.</span>
</div>""", unsafe_allow_html=True)

if absents:
    st.markdown(f"""
<div class='alert-card alert-amber'>
  <strong>⚠️ {len(absents)} référence(s) Top CA absentes de toutes les extractions</strong><br>
  <span style='font-size:12px;opacity:.85'>→ Vérifier déréférencement ou erreur de code dans la liste Top CA.</span>
</div>""", unsafe_allow_html=True)

if n_urgences == 0 and sites_sous.empty and stock_immo_total == 0:
    st.success("✅ Tous les magasins au-dessus de la cible · aucun stock immobilisé détecté.")

# ─── TABS ─────────────────────────────────────────────────────────────────────
st.markdown("---")
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Synthèse réseau", "🔄 IM vs LO", "🚨 Plan d'action", "🚫 Absents ERP"
])

# ═══ TAB 1 — SYNTHÈSE RÉSEAU ══════════════════════════════════════════════════
with tab1:
    st.caption("Articles permanents (Type saisonnalité = P) · Code état 2 dans le calcul · Stock immobilisé = code B × PMP")

    disp1 = taux_all[["site","n_perm","n_actifs","n_stock","taux",
                       "n_bloques","stock_immo","n_rupture"]].copy()
    disp1.columns = ["Magasin","Réf perm.","Actifs (état 2)",
                     "En stock","Taux %","Bloqués (B)",
                     "Stock immobilisé (FCFA)","Ruptures"]
    disp1 = disp1.sort_values("Taux %")
    disp1["Taux %"]                 = disp1["Taux %"].apply(lambda x: f"{x:.1f}%" if x else "—")
    disp1["Stock immobilisé (FCFA)"]= disp1["Stock immobilisé (FCFA)"].apply(fmt)
    disp1["Statut"] = taux_all.sort_values("taux")["taux"].apply(
        lambda x: "🟢 OK" if x and x>=cible else ("🟡 Surveiller" if x and x>=cible-10 else "🔴 Action"))

    st.dataframe(disp1, use_container_width=True, hide_index=True,
                 column_config={"Stock immobilisé (FCFA)": st.column_config.TextColumn(
                     "Stock immo. (FCFA)", help="Articles bloqués code B × PMP")})

    # Graphique taux par magasin
    try:
        import plotly.graph_objects as go
        sorted_sites = taux_all.sort_values("taux")
        colors = [color_taux(v, cible) for v in sorted_sites["taux"]]
        fig = go.Figure(go.Bar(
            x=sorted_sites["taux"].tolist(),
            y=sorted_sites["site"].tolist(),
            orientation="h",
            marker_color=colors, marker_line_width=0,
            text=[f"{v:.1f}%" if v else "—" for v in sorted_sites["taux"]],
            textposition="outside",
        ))
        fig.add_vline(x=cible, line_width=1.5, line_dash="dash", line_color="#007AFF",
                      annotation_text=f"Cible {cible}%", annotation_position="top right",
                      annotation_font_color="#007AFF")
        fig.update_layout(
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(family="-apple-system, Helvetica Neue", color="#3A3A3C", size=12),
            height=max(260, n_sites*50+80), margin=dict(t=20,b=20,l=10,r=80),
            xaxis=dict(title="Taux de détention (%)", ticksuffix="%",
                       showgrid=True, gridcolor="#F2F2F7", range=[0,110]),
            yaxis=dict(showgrid=False, title=""),
        )
        st.plotly_chart(fig, use_container_width=True)
    except ImportError:
        pass

# ═══ TAB 2 — IM vs LO ════════════════════════════════════════════════════════
with tab2:
    # Jauges IM / LO
    tot_im  = taux_im["n_actifs"].sum(); pres_im = taux_im["n_stock"].sum()
    tot_lo  = taux_lo["n_actifs"].sum(); pres_lo = taux_lo["n_stock"].sum()
    g_im    = pres_im/tot_im*100 if tot_im else 0
    g_lo    = pres_lo/tot_lo*100 if tot_lo else 0
    immo_im = taux_im["stock_immo"].sum()
    immo_lo = taux_lo["stock_immo"].sum()

    c1, c2 = st.columns(2)
    for col_w, label, taux_g, tot, pres, immo, flux_col in [
        (c1,"IM · Import",g_im,tot_im,pres_im,immo_im,"#7C3AED"),
        (c2,"LO · Local", g_lo,tot_lo,pres_lo,immo_lo,"#007AFF"),
    ]:
        with col_w:
            col_v = "#34C759" if taux_g>=cible else "#FF9500" if taux_g>=cible-10 else "#FF3B30"
            st.markdown(f"""
<div style='background:#FFFFFF;border:0.5px solid #E5E5EA;border-left:3px solid {flux_col};border-radius:12px;padding:16px 18px;margin-bottom:12px'>
  <div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px'>{label}</div>
  <div style='font-size:28px;font-weight:700;color:{col_v};letter-spacing:-.02em'>{taux_g:.1f}%</div>
  <div style='font-size:12px;color:#8E8E93;margin-top:3px'>{int(pres)} / {int(tot)} actifs permanents en stock · {n_sites} magasins</div>
  <div style='background:#E5E5EA;border-radius:3px;height:5px;margin-top:10px'>
    <div style='width:{min(taux_g,100):.0f}%;background:{flux_col};height:5px;border-radius:3px'></div>
  </div>
  <div style='font-size:12px;color:#FF3B30;margin-top:8px;font-weight:500'>Stock immo. : {fmt(immo)} FCFA</div>
</div>""", unsafe_allow_html=True)

    # Tableau IM vs LO
    try:
        pivot = taux_df[taux_df["flux"]!="ALL"].pivot_table(
            index="site", columns="flux",
            values=["n_actifs","n_stock","taux","stock_immo"],
            aggfunc="first"
        ).reset_index()
        pivot.columns = ["Magasin",
                         "Actifs IM","Actifs LO",
                         "En stock IM","En stock LO",
                         "Stock immo IM","Stock immo LO",
                         "Taux IM %","Taux LO %"]
        pivot = pivot[["Magasin","Actifs IM","En stock IM","Taux IM %",
                        "Actifs LO","En stock LO","Taux LO %",
                        "Stock immo IM","Stock immo LO"]].sort_values("Taux IM %")
        for c in ["Taux IM %","Taux LO %"]:
            pivot[c] = pivot[c].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "—")
        for c in ["Stock immo IM","Stock immo LO"]:
            pivot[c] = pivot[c].apply(fmt)
        st.dataframe(pivot, use_container_width=True, hide_index=True)
    except Exception:
        st.dataframe(taux_df[taux_df["flux"]!="ALL"], use_container_width=True, hide_index=True)

# ═══ TAB 3 — PLAN D'ACTION ════════════════════════════════════════════════════
with tab3:
    urgences = grp_perm[grp_perm["Alerte"]!="✅ OK"].copy()

    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        sel_site  = st.selectbox("Magasin", ["Tous"]+sorted(grp_perm["Libellé site"].unique()))
    with fc2:
        sel_flux  = st.selectbox("Flux", ["Tous","IM","LO"])
    with fc3:
        al_dispo  = sorted(urgences["Alerte"].unique())
        sel_al    = st.multiselect("Alerte", al_dispo, default=al_dispo)

    if sel_site!="Tous": urgences = urgences[urgences["Libellé site"]==sel_site]
    if sel_flux!="Tous": urgences = urgences[urgences["flux"]==sel_flux]
    if sel_al:           urgences = urgences[urgences["Alerte"].isin(sel_al)]

    st.markdown(f"<div style='font-size:12px;color:#8E8E93;margin-bottom:8px'>{len(urgences)} article(s) permanents nécessitant une action</div>", unsafe_allow_html=True)

    if urgences.empty:
        st.success("✅ Aucune urgence sur la sélection.")
    else:
        disp3 = urgences[["Code article","lib_article","Libellé site","flux",
                           "code_etat","saisonnalite","stock","ral","stock_immo","Alerte"]].copy()
        disp3.columns = ["Code","Libellé","Magasin","Flux","Code état",
                         "Saisonnalité","Stock","RAL","Stock immo (FCFA)","Alerte"]
        disp3["Stock"]            = disp3["Stock"].apply(lambda x: int(x))
        disp3["RAL"]              = disp3["RAL"].apply(lambda x: int(x))
        disp3["Stock immo (FCFA)"]= disp3["Stock immo (FCFA)"].apply(lambda x: fmt(x) if x>0 else "—")
        st.dataframe(disp3.sort_values("Alerte"), use_container_width=True, hide_index=True,
                     column_config={
                         "Flux":       st.column_config.TextColumn("Flux",  width="small"),
                         "Code état":  st.column_config.TextColumn("État",  width="small"),
                         "Saisonnalité": st.column_config.TextColumn("Saison",width="small"),
                     })

# ═══ TAB 4 — ABSENTS ERP ══════════════════════════════════════════════════════
with tab4:
    if not absents:
        st.success("✅ Toutes les références Top CA sont présentes dans au moins une extraction.")
    else:
        st.warning(f"⚠️ {len(absents)} référence(s) Top CA absentes de toutes les extractions ERP")
        df_abs = pd.DataFrame({
            "Code article": absents,
            "Statut": "Absent ERP",
            "Action suggérée": "Vérifier référentiel ou déréférencement non planifié"
        })
        st.dataframe(df_abs, use_container_width=True, hide_index=True)

# ─── EXPORT ───────────────────────────────────────────────────────────────────
st.markdown("---")
with st.expander("📥 Export Excel — Synthèse · IM vs LO · Plan d'action · Stock immobilisé · Absents ERP"):
    st.caption("5 onglets · Articles permanents · Code état 2 · Valorisation PMP")
    if st.button("Générer le fichier Excel", type="primary"):
        with st.spinner("Génération…"):
            buf = gen_excel(grp, grp_perm, taux_df, absents, top_codes)
        st.download_button(
            "⬇️ Télécharger",
            data=buf,
            file_name="SmartBuyer_Detention_TopCA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
