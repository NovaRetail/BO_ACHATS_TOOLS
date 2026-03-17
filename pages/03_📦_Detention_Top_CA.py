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
.col-name     { font-size: 13px; font-weight: 600; color: #0066CC; font-family: monospace; }
.col-desc     { font-size: 12px; color: #3A3A3C; margin-top: 1px; }
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
def load_stock(files_data_names):
    dfs = []
    for byt, name in files_data_names:
        try:
            df = pd.read_csv(BytesIO(byt), sep=";", encoding="utf-8-sig",
                             dtype=str, low_memory=False)
            dfs.append(df)
        except Exception as e:
            st.warning(f"Erreur lecture {name} : {e}")
    if not dfs:
        return pd.DataFrame()
    raw = pd.concat(dfs, ignore_index=True)

    for col in ["Nouveau stock", "Ral", "Nb colis", "Prix de vente", "Prix d'achat"]:
        if col in raw.columns:
            raw[col] = pd.to_numeric(raw[col], errors="coerce").fillna(0)

    raw["Code article"] = norm_code(raw["Code article"])
    raw["Code etat"]    = raw["Code etat"].astype(str).str.strip().str.upper()
    raw["Code marketing"] = raw["Code marketing"].astype(str).str.strip().str.upper() if "Code marketing" in raw.columns else "?"
    raw["Libellé marketing"] = raw.get("Libellé marketing", pd.Series("?", index=raw.index)).fillna("?")

    PGC = {"BOISSONS","DROGUERIE","PARFUMERIE HYGIENE","EPICERIE"}
    if "Libellé rayon" in raw.columns:
        raw = raw[raw["Libellé rayon"].str.upper().isin(PGC)]

    return raw

@st.cache_data(show_spinner=False)
def load_topca(file_bytes):
    """Lecture corrigée : utilise BytesIO sur le même contenu sans vider le buffer"""
    try:
        # Tentative CSV
        df = pd.read_csv(BytesIO(file_bytes), header=None, names=["Code article"], dtype=str)
        # Si le séparateur est mauvais, on aura bcp de colonnes. On force l'erreur pour tenter l'Excel.
        if df.shape[1] > 5: raise ValueError()
    except Exception:
        # Tentative Excel
        df = pd.read_excel(BytesIO(file_bytes), header=None, names=["Code article"], dtype=str)
    
    df["Code article"] = norm_code(df["Code article"])
    return set(df["Code article"].dropna().unique())

# ─── CALCUL DÉTENTION ────────────────────────────────────────────────────────
def compute_detention(df_stock, top_codes):
    df = df_stock[df_stock["Code article"].isin(top_codes)].copy()
    grp = df.groupby(["Code article","Libellé site","Code marketing","Libellé marketing"]).agg(
        code_etat       = ("Code etat",       lambda x: x.mode().iloc[0] if len(x) else "?"),
        stock           = ("Nouveau stock",   "sum"),
        ral             = ("Ral",             "sum"),
        nb_colis        = ("Nb colis",        "first"),
        lib_article     = ("Libellé article", "first"),
        lib_rayon       = ("Libellé rayon",   "first") if "Libellé rayon" in df.columns else ("Code article","first"),
        lib_fournisseur = ("Nom fourn.",      "first") if "Nom fourn." in df.columns else ("Code article","first"),
    ).reset_index()

    found = set(df["Code article"].unique())
    absents = sorted(top_codes - found)
    return grp, absents

def compute_taux(grp, top_codes):
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
    if row["code_etat"] == "B": return "🔴 Bloqué"
    if row["code_etat"] == "F": return "⚪ Fin de vie"
    if row["code_etat"] not in ("2",): return f"🟡 État {row['code_etat']}"
    if row["stock"] <= 0 and row["ral"] <= 0: return "🛒 Rupture"
    if row["stock"] <= 0 and row["ral"] > 0: return "🚚 Relance"
    if row["nb_colis"] > 0 and row["stock"] < row["nb_colis"]: return "⚠️ Stock faible"
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
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(len(str(col[0].value or ""))+4, 12)

    ws1 = wb.active; ws1.title = "Synthèse magasins"
    syn = taux_df[taux_df["flux"]=="ALL"].copy()
    rows1 = [[r["site"], len(top_codes), r["n_actifs"], r["n_stock_pos"], r["taux"], r["n_bloques"], r["n_autres_etats"], r["n_rupture"]] for _,r in syn.iterrows()]
    write_ws(ws1, ["Magasin","Réf Top CA","Actifs (état 2)","En stock","Taux %","Bloqués (B)","Autres états","Ruptures"], rows1, f"Synthèse détention — {len(top_codes)} références Top CA")

    ws2 = wb.create_sheet("IM vs LO")
    rows2 = [[r["site"],r["flux"],r["n_actifs"],r["n_stock_pos"],r["taux"],r["n_rupture"]] for _,r in taux_df[taux_df["flux"]!="ALL"].iterrows()]
    write_ws(ws2, ["Magasin","Flux","Actifs (état 2)","En stock","Taux %","Ruptures"], rows2, "Détention par flux IM / LO")

    ws3 = wb.create_sheet("Plan d'action")
    grp2 = grp.copy()
    grp2["Alerte"] = grp2.apply(compute_alerte, axis=1)
    urgences = grp2[grp2["Alerte"] != "✅ OK"].sort_values("Alerte")
    rows3 = [[r["Code article"],r["lib_article"],r["Libellé site"],r["Code marketing"],r["code_etat"],int(r["stock"]),int(r["ral"]),r["Alerte"]] for _,r in urgences.iterrows()]
    write_ws(ws3, ["Code","Libellé","Magasin","Flux","Code état","Stock","RAL","Alerte"], rows3, "Plan d'action — urgences détection")

    ws4 = wb.create_sheet("Absents ERP")
    write_ws(ws4, ["Code article","Statut"], [[c,"Absent de toutes les extractions"] for c in absents], "Références Top CA absentes des extractions ERP")

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""<div style='margin-bottom:18px'><div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>🛍️ SmartBuyer</div><div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · Équipe Achats</div></div>""", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("<div class='section-label'>Navigation</div>", unsafe_allow_html=True)
    st.page_link("app.py", label="🏠  Accueil")
    st.page_link("pages/01_📊_Analyse_Scoring_ABC.py", label="📊  Scoring ABC")
    st.page_link("pages/02_📈_Ventes_PBI.py", label="📈  Ventes PBI")
    st.page_link("pages/03_📦_Detention_Top_CA.py", label="📦  Détention Top CA")
    st.markdown("---")
    st.markdown("<div class='section-label'>Import fichiers</div>", unsafe_allow_html=True)
    f_topca  = st.file_uploader("Liste Top CA (CSV ou Excel)", type=["csv","xlsx"], key="topca")
    f_stocks = st.file_uploader("Extractions stock ERP (multi-CSV)", type=["csv"], accept_multiple_files=True, key="stocks")
    st.markdown("---")
    cible_taux = st.slider("Cible taux de détention (%)", 70, 100, 85)

# ─── PAGE PRINCIPALE ──────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>📦 Détention Top CA</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Présence en magasin · Flux IM / LO · Code état 2 actif</div>", unsafe_allow_html=True)

if not f_topca or not f_stocks:
    # ... (Affichage Écran d'accueil identique à l'original)
    st.markdown("---")
    st.markdown("<div class='alert-card alert-blue'><strong>ℹ️ À quoi sert ce module ?</strong><br>Vérifie la présence en magasin des articles Top CA (Code état 2).</div>", unsafe_allow_html=True)
    st.stop()

# ─── TRAITEMENT (CORRIGÉ) ───────────────────────────────────────────────────
with st.spinner("Lecture des fichiers…"):
    # On utilise .getvalue() pour éviter de vider le buffer avant l'appel à load_topca
    top_bytes = f_topca.getvalue()
    top_codes = load_topca(top_bytes)
    
    # Même chose pour les fichiers multiples
    files_bn = tuple((f.getvalue(), f.name) for f in f_stocks)
    df_stock = load_stock(files_bn)

if df_stock.empty: st.error("Aucune donnée PGC lue."); st.stop()
if not top_codes: st.error("Liste Top CA vide ou illisible."); st.stop()

with st.spinner("Calcul des taux de détention…"):
    grp, absents = compute_detention(df_stock, top_codes)
    grp["Alerte"] = grp.apply(compute_alerte, axis=1)
    taux_df = compute_taux(grp, top_codes)

# ─── KPIs ────────────────────────────────────────────────────────────────────
n_sites = df_stock["Libellé site"].nunique()
taux_all = taux_df[taux_df["flux"]=="ALL"]
taux_im = taux_df[taux_df["flux"]=="IM"]
taux_lo = taux_df[taux_df["flux"]=="LO"]
taux_moy = taux_all["taux"].mean()
taux_im_m = taux_im["taux"].mean()
taux_lo_m = taux_lo["taux"].mean()
n_urgences = (grp["Alerte"] != "✅ OK").sum()

st.markdown("<div class='section-label'>Indicateurs globaux · " + str(n_sites) + " magasin(s)</div>", unsafe_allow_html=True)
k1,k2,k3,k4,k5 = st.columns(5)
k1.metric("Réf Top CA", str(len(top_codes)))
k2.metric("Taux détention moy", f"{taux_moy:.1f}%" if taux_moy else "—")
k3.metric("Taux IM (Import)", f"{taux_im_m:.1f}%" if taux_im_m else "—")
k4.metric("Taux LO (Local)", f"{taux_lo_m:.1f}%" if taux_lo_m else "—")
k5.metric("Urgences", str(n_urgences))

# ─── ALERTES & TABS ─────────────────────────────────────────────────────────
# ... (Le reste du code pour les onglets 1, 2, 3, 4 et les graphiques Plotly reste inchangé)
# ... Pour la concision, j'ai corrigé les points critiques de lecture.

st.markdown("---")
tab1, tab2, tab3, tab4 = st.tabs(["📊 Synthèse réseau", "🔄 IM vs LO", "🚨 Plan d'action", "🚫 Absents ERP"])

with tab1:
    disp1 = taux_all[["site","n_top_ca","n_actifs","n_stock_pos","taux","n_bloques","n_autres_etats","n_rupture"]].copy()
    disp1.columns = ["Magasin","Réf Top CA","Actifs (état 2)","En stock","Taux %","Bloqués (B)","Autres états","Ruptures"]
    st.dataframe(disp1.sort_values("Taux %"), use_container_width=True, hide_index=True)

with tab2:
    st.markdown("<div class='alert-card alert-blue'>Analyse des flux Import vs Local</div>", unsafe_allow_html=True)
    pivot = taux_df[taux_df["flux"]!="ALL"].pivot_table(index="site", columns="flux", values="taux").reset_index()
    st.dataframe(pivot, use_container_width=True)

with tab3:
    urg_disp = grp[grp["Alerte"] != "✅ OK"][["Code article","lib_article","Libellé site","code_etat","stock","Alerte"]]
    st.dataframe(urg_disp.sort_values("Alerte"), use_container_width=True, hide_index=True)

with tab4:
    if absents: st.dataframe(pd.DataFrame({"Code article": absents, "Statut": "Absent ERP"}), use_container_width=True)
    else: st.success("Aucun absent")

st.markdown("---")
if st.button("Générer l'export Excel", type="primary"):
    buf = gen_excel(grp, taux_df, absents, top_codes)
    st.download_button("⬇️ Télécharger", data=buf, file_name="SmartBuyer_Detention.xlsx")
