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
}

def norm_code(s):
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(8)

# ─── PARSING ──────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_stock(files_data_names):
    dfs = []
    for content, name in files_data_names:
        try:
            # On essaye de lire avec le séparateur point-virgule (standard ERP Carrefour)
            df = pd.read_csv(BytesIO(content), sep=";", encoding="utf-8-sig",
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
    
    # Garder PGC uniquement
    PGC = {"BOISSONS","DROGUERIE","PARFUMERIE HYGIENE","EPICERIE"}
    if "Libellé rayon" in raw.columns:
        raw = raw[raw["Libellé rayon"].str.upper().isin(PGC)]

    return raw

@st.cache_data(show_spinner=False)
def load_topca(file_content):
    """Lit le fichier Top CA soit en CSV soit en Excel."""
    try:
        # Tentative CSV (si c'est un Excel, read_csv lèvera souvent une erreur d'encodage ou de parsing)
        df = pd.read_csv(BytesIO(file_content), header=None, names=["Code article"], dtype=str)
        if df.shape[1] > 10: raise ValueError("Probablement pas un CSV à une colonne")
    except Exception:
        # Tentative Excel
        try:
            df = pd.read_excel(BytesIO(file_content), header=None, names=["Code article"], dtype=str)
        except Exception as e:
            st.error(f"Impossible de lire le fichier Top CA : {e}")
            return set()
            
    df["Code article"] = norm_code(df["Code article"])
    return set(df["Code article"].dropna().unique())

# ─── CALCULS (FONCTIONS INCHANGÉES MAIS RE-DÉCLARÉES) ────────────────────────
def compute_detention(df_stock, top_codes):
    df = df_stock[df_stock["Code article"].isin(top_codes)].copy()
    grp = df.groupby(["Code article","Libellé site","Code marketing"]).agg(
        code_etat       = ("Code etat",       lambda x: x.mode().iloc[0] if len(x) else "?"),
        stock           = ("Nouveau stock",   "sum"),
        ral             = ("Ral",             "sum"),
        nb_colis        = ("Nb colis",        "first"),
        lib_article     = ("Libellé article", "first"),
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
                "n_autres_etats": int((sf["code_etat"].isin(["P","S","F"])).sum()),
                "n_rupture": int((actifs["stock"] <= 0).sum()),
            })
    return pd.DataFrame(rows)

def compute_alerte(row):
    if row["code_etat"] == "B": return "🔴 Bloqué"
    if row["code_etat"] == "F": return "⚪ Fin de vie"
    if row["code_etat"] not in ("2",): return f"🟡 État {row['code_etat']}"
    if row["stock"] <= 0 and row["ral"] <= 0: return "🛒 Rupture"
    if row["stock"] <= 0 and row["ral"] > 0: return "🚚 Relance"
    return "✅ OK"

# ─── EXPORT EXCEL (LOGIQUE SIMPLIFIÉE) ────────────────────────────────────────
def gen_excel(grp, taux_df, absents, top_codes):
    wb = Workbook()
    ws1 = wb.active; ws1.title = "Synthèse"
    for r in dataframe_to_rows(taux_df, index=False, header=True): ws1.append(r)
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf

from openpyxl.utils.dataframe import dataframe_to_rows

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("<b style='font-size:18px'>🛍️ SmartBuyer</b>", unsafe_allow_html=True)
    st.markdown("---")
    f_topca  = st.file_uploader("Liste Top CA", type=["csv","xlsx"])
    f_stocks = st.file_uploader("Extractions ERP", type=["csv"], accept_multiple_files=True)
    cible_taux = st.slider("Cible (%)", 70, 100, 85)

# ─── LOGIQUE PRINCIPALE ───────────────────────────────────────────────────────
st.markdown("<div class='page-title'>📦 Détention Top CA</div>", unsafe_allow_html=True)

if not f_topca or not f_stocks:
    st.info("Veuillez charger les fichiers dans la barre latérale.")
    st.stop()

# CORRECTION ICI : Utilisation de getvalue() pour ne pas vider le buffer
with st.spinner("Analyse en cours..."):
    # Lecture du Top CA
    top_content = f_topca.getvalue()
    top_codes = load_topca(top_content)
    
    # Lecture des stocks
    stocks_data = tuple((f.getvalue(), f.name) for f in f_stocks)
    df_stock = load_stock(stocks_data)

if df_stock.empty or not top_codes:
    st.error("Données invalides. Vérifiez vos fichiers.")
    st.stop()

# Suite du traitement
grp, absents = compute_detention(df_stock, top_codes)
grp["Alerte"] = grp.apply(compute_alerte, axis=1)
taux_df = compute_taux(grp, top_codes)

# Affichage des KPIs
taux_moy = taux_df[taux_df["flux"]=="ALL"]["taux"].mean()
c1, c2, c3 = st.columns(3)
c1.metric("Articles Top CA", len(top_codes))
c2.metric("Taux Moyen", f"{taux_moy:.1f}%")
c3.metric("Urgences", (grp["Alerte"] != "✅ OK").sum())

# Onglets d'affichage
t1, t2, t3 = st.tabs(["📊 Synthèse", "🚨 Plan d'action", "🚫 Absents"])
with t1:
    st.dataframe(taux_df[taux_df["flux"]=="ALL"], use_container_width=True, hide_index=True)
with t2:
    st.dataframe(grp[grp["Alerte"] != "✅ OK"], use_container_width=True, hide_index=True)
with t3:
    st.write(f"Articles non trouvés dans l'ERP : {len(absents)}")
    st.write(absents)

# Export
buf = gen_excel(grp, taux_df, absents, top_codes)
st.download_button("📥 Télécharger l'export Excel", buf, "Rapport_Detention.xlsx")
