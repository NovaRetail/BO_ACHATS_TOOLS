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

# ─── CHARTE SMARTBUYER v2 (RESTAURÉE) ────────────────────────────────────────
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

# ─── PARSING (CORRIGÉ ET SÉCURISÉ) ───────────────────────────────────────────
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
    if not dfs: return pd.DataFrame()
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
    """Utilise getvalue() pour la stabilité et gère CSV/Excel proprement"""
    try:
        # Essai Excel (BytesIO frais à chaque tentative pour éviter l'erreur de curseur)
        df = pd.read_excel(BytesIO(file_bytes), header=None, names=["Code article"], dtype=str)
    except Exception:
        try:
            # Essai CSV
            df = pd.read_csv(BytesIO(file_bytes), header=None, names=["Code article"], dtype=str)
        except Exception:
            return set()
    
    if df.empty: return set()
    df["Code article"] = norm_code(df["Code article"])
    return set(df["Code article"].dropna().unique())

# ─── CALCULS (LOGIQUE INCHANGÉE) ─────────────────────────────────────────────
def compute_detention(df_stock, top_codes):
    df = df_stock[df_stock["Code article"].isin(top_codes)].copy()
    grp = df.groupby(["Code article","Libellé site","Code marketing","Libellé marketing"]).agg(
        code_etat       = ("Code etat",       lambda x: x.mode().iloc[0] if not x.empty else "?"),
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
            sf = s if flux == "ALL" else s[s["Code marketing"] == flux]
            actifs    = sf[sf["code_etat"] == "2"]
            n_actifs  = len(actifs)
            n_stock   = (actifs["stock"] > 0).sum()
            taux      = n_stock / n_actifs * 100 if n_actifs > 0 else None
            rows.append({
                "site": site, "flux": flux, "n_top_ca": len(top_codes),
                "n_actifs": n_actifs, "n_stock_pos": int(n_stock),
                "taux": round(taux, 1) if taux is not None else None,
                "n_bloques": int((sf["code_etat"] == "B").sum()),
                "n_autres_etats": int((sf["code_etat"].isin(["P","S","F","1","5","6"])).sum()),
                "n_rupture": int((actifs["stock"] <= 0).sum()),
                "n_faible": int(((actifs["stock"] > 0) & (actifs["stock"] < actifs["nb_colis"].replace(0, np.nan))).sum()),
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

# ─── EXPORT EXCEL (RESTAURÉ) ────────────────────────────────────────────────
def gen_excel(grp, taux_df, absents, top_codes):
    wb = Workbook()
    HDR_F = PatternFill("solid", fgColor="1C3557")
    HDR_T = Font(bold=True, color="FFFFFF", size=11)
    def write_ws(ws, headers, rows, title):
        ws.append([title]); ws.append([]); ws.append(headers)
        for i,h in enumerate(headers,1):
            c = ws.cell(3,i); c.fill=HDR_F; c.font=HDR_T; c.alignment=Alignment(horizontal="center")
        for row in rows: ws.append(row)
    
    ws1 = wb.active; ws1.title = "Synthèse"
    syn = taux_df[taux_df["flux"]=="ALL"]
    write_ws(ws1, ["Magasin","Réf Top CA","Actifs","En stock","Taux %","Ruptures"], 
             [[r["site"], len(top_codes), r["n_actifs"], r["n_stock_pos"], r["taux"], r["n_rupture"]] for _,r in syn.iterrows()], "Synthèse Réseau")
    
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ─── SIDEBAR (RESTAURÉE) ─────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""<div style='margin-bottom:18px'><div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>🛍️ SmartBuyer</div></div>""", unsafe_allow_html=True)
    st.markdown("---")
    f_topca = st.file_uploader("Liste Top CA", type=["csv","xlsx"])
    f_stocks = st.file_uploader("Extractions stock ERP", type=["csv"], accept_multiple_files=True)
    cible_taux = st.slider("Cible taux de détention (%)", 70, 100, 85)

# ─── PAGE PRINCIPALE (RESTAURÉE) ─────────────────────────────────────────────
st.markdown("<div class='page-title'>📦 Détention Top CA</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Présence en magasin · Flux IM / LO · Code état 2 actif</div>", unsafe_allow_html=True)

if not f_topca or not f_stocks:
    # Écran d'accueil stylisé (Inchangé)
    st.markdown("<div class='alert-card alert-blue'><strong>ℹ️ À quoi sert ce module ?</strong><br>Vérifie la présence en magasin des articles Top CA actifs.</div>", unsafe_allow_html=True)
    st.stop()

# ─── TRAITEMENT (FIXED) ──────────────────────────────────────────────────────
with st.spinner("Lecture des fichiers..."):
    # FIX: Utilisation de getvalue() pour la stabilité
    top_bytes = f_topca.getvalue() 
    top_codes = load_topca(top_bytes)
    
    stocks_data = tuple((f.getvalue(), f.name) for f in f_stocks)
    df_stock = load_stock(stocks_data)

if df_stock.empty or not top_codes:
    st.error("Erreur de données."); st.stop()

grp, absents = compute_detention(df_stock, top_codes)
grp["Alerte"] = grp.apply(compute_alerte, axis=1)
taux_df = compute_taux(grp, top_codes)

# ─── AFFICHAGE KPIs STYLISÉS (RESTAURÉS) ────────────────────────────────────
n_sites = df_stock["Libellé site"].nunique()
taux_all = taux_df[taux_df["flux"]=="ALL"]
taux_moy = taux_all["taux"].mean()

st.markdown(f"<div class='section-label'>Indicateurs globaux · {n_sites} magasin(s)</div>", unsafe_allow_html=True)
k1, k2, k3 = st.columns(3)
k1.metric("Réf Top CA", len(top_codes))
k2.metric("Taux détention moy", f"{taux_moy:.1f}%", f"cible {cible_taux}%")
k3.metric("Urgences", (grp["Alerte"] != "✅ OK").sum())

# ─── TABS STYLISÉS (RESTAURÉS) ───────────────────────────────────────────────
st.markdown("---")
tab1, tab2, tab3, tab4 = st.tabs(["📊 Synthèse réseau", "🔄 IM vs LO", "🚨 Plan d'action", "🚫 Absents ERP"])

with tab1:
    disp1 = taux_all[["site","n_actifs","n_stock_pos","taux","n_rupture"]].copy()
    disp1.columns = ["Magasin","Actifs (état 2)","En stock","Taux %","Ruptures"]
    st.dataframe(disp1.sort_values("Taux %"), use_container_width=True, hide_index=True)

with tab2:
    st.markdown("<div class='alert-card alert-blue'>Comparaison des flux Import et Local</div>", unsafe_allow_html=True)
    pivot = taux_df[taux_df["flux"]!="ALL"].pivot_table(index="site", columns="flux", values="taux").reset_index()
    st.dataframe(pivot, use_container_width=True, hide_index=True)

with tab3:
    urgences = grp[grp["Alerte"] != "✅ OK"].sort_values("Alerte")
    st.dataframe(urgences[["Code article","lib_article","Libellé site","Alerte","stock"]], use_container_width=True, hide_index=True)

with tab4:
    if absents:
        st.dataframe(pd.DataFrame({"Code article": absents, "Statut": "Absent ERP"}), use_container_width=True, hide_index=True)
    else:
        st.success("Toutes les références sont présentes.")

# ─── EXPORT (RESTAURÉ) ──────────────────────────────────────────────────────
st.markdown("---")
if st.button("Générer l'export Excel", type="primary"):
    buf = gen_excel(grp, taux_df, absents, top_codes)
    st.download_button("⬇️ Télécharger", buf, "SmartBuyer_Detention.xlsx")
