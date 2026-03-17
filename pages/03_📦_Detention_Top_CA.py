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

# ─── CHARTE SMARTBUYER (Inchangée) ───────────────────────────────────────────
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
[data-testid="stTabs"] button[role="tab"] { font-size: 13px !important; font-weight: 500 !important; padding: 8px 16px !important; color: #8E8E93 !important; border-radius: 0 !important; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color: #007AFF !important; border-bottom: 2px solid #007AFF !important; background: transparent !important; }
[data-testid="stDataFrame"] { border: 0.5px solid #E5E5EA !important; border-radius: 10px !important; }
.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 1.5rem; }
.section-label { font-size: 11px; font-weight: 600; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 10px; }
.alert-card { padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; font-size: 13px; line-height: 1.5; border-left: 3px solid; }
.alert-red   { background: #FFF2F2; border-color: #FF3B30; color: #3A0000; }
.alert-blue  { background: #F0F8FF; border-color: #007AFF; color: #001A3A; }
.alert-amber { background: #FFFBF0; border-color: #FF9500; color: #3A2000; }
</style>
""", unsafe_allow_html=True)

# ─── UTILITAIRES ──────────────────────────────────────────────────────────────
def norm_code(s):
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(8)

# ─── PARSING (CORRIGÉ) ────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_topca(file_bytes):
    """Lecture sécurisée des bytes du Top CA"""
    try:
        # On tente d'abord l'Excel (plus commun pour le Top CA)
        df = pd.read_excel(BytesIO(file_bytes), header=None, names=["Code article"], dtype=str)
    except Exception:
        try:
            # Si échec, on tente le CSV
            df = pd.read_csv(BytesIO(file_bytes), header=None, names=["Code article"], dtype=str, sep=None, engine='python')
        except Exception as e:
            st.error(f"Erreur de lecture du fichier Top CA : {e}")
            return set()
    
    if "Code article" in df.columns:
        return set(norm_code(df["Code article"].dropna()).unique())
    return set()

@st.cache_data(show_spinner=False)
def load_stock(files_list):
    """Lecture sécurisée des extractions ERP"""
    dfs = []
    for content, name in files_list:
        try:
            df = pd.read_csv(BytesIO(content), sep=";", encoding="utf-8-sig", dtype=str, low_memory=False)
            dfs.append(df)
        except Exception as e:
            st.warning(f"Erreur sur {name} : {e}")
    
    if not dfs: return pd.DataFrame()
    raw = pd.concat(dfs, ignore_index=True)

    for col in ["Nouveau stock", "Ral", "Nb colis"]:
        if col in raw.columns:
            raw[col] = pd.to_numeric(raw[col], errors="coerce").fillna(0)

    raw["Code article"] = norm_code(raw["Code article"])
    raw["Code etat"] = raw["Code etat"].astype(str).str.strip().str.upper()
    raw["Code marketing"] = raw["Code marketing"].astype(str).str.strip().str.upper() if "Code marketing" in raw.columns else "LO"
    
    # Filtre PGC
    PGC = {"BOISSONS","DROGUERIE","PARFUMERIE HYGIENE","EPICERIE"}
    if "Libellé rayon" in raw.columns:
        raw = raw[raw["Libellé rayon"].str.upper().isin(PGC)]
    return raw

# ─── CALCULS ──────────────────────────────────────────────────────────────────
def compute_detention(df_stock, top_codes):
    df = df_stock[df_stock["Code article"].isin(top_codes)].copy()
    if df.empty: return pd.DataFrame(), sorted(list(top_codes))
    
    grp = df.groupby(["Code article","Libellé site","Code marketing"]).agg(
        code_etat   = ("Code etat", lambda x: x.mode().iloc[0] if not x.empty else "?"),
        stock       = ("Nouveau stock", "sum"),
        ral         = ("Ral", "sum"),
        nb_colis    = ("Nb colis", "first"),
        lib_article = ("Libellé article", "first"),
    ).reset_index()
    
    absents = sorted(list(top_codes - set(df["Code article"].unique())))
    return grp, absents

def compute_taux(grp, top_codes):
    rows = []
    sites = grp["Libellé site"].unique()
    for site in sites:
        s = grp[grp["Libellé site"] == site]
        for flux in ["IM", "LO", "ALL"]:
            sf = s if flux == "ALL" else s[s["Code marketing"] == flux]
            actifs = sf[sf["code_etat"] == "2"]
            n_actifs = len(actifs)
            n_stock = (actifs["stock"] > 0).sum()
            taux = (n_stock / n_actifs * 100) if n_actifs > 0 else None
            rows.append({
                "site": site, "flux": flux, "n_actifs": n_actifs,
                "n_stock_pos": int(n_stock), "taux": taux,
                "n_bloques": (sf["code_etat"] == "B").sum(),
                "n_rupture": (actifs["stock"] <= 0).sum()
            })
    return pd.DataFrame(rows)

def compute_alerte(row):
    if row["code_etat"] == "B": return "🔴 Bloqué"
    if row["code_etat"] == "F": return "⚪ Fin de vie"
    if row["code_etat"] != "2": return f"🟡 État {row['code_etat']}"
    if row["stock"] <= 0: return "🛒 Rupture"
    return "✅ OK"

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🛍️ SmartBuyer Hub")
    f_topca = st.file_uploader("Liste Top CA (CSV/XLSX)", type=["csv", "xlsx"])
    f_stocks = st.file_uploader("Stocks ERP (CSV)", type=["csv"], accept_multiple_files=True)
    cible_taux = st.slider("Cible (%)", 70, 100, 85)

# ─── MAIN ─────────────────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>📦 Détention Top CA</div>", unsafe_allow_html=True)

if f_topca and f_stocks:
    with st.spinner("Analyse..."):
        # Utilisation de getvalue() pour la stabilité du cache
        top_codes = load_topca(f_topca.getvalue())
        stocks_data = tuple((f.getvalue(), f.name) for f in f_stocks)
        df_stock = load_stock(stocks_data)

    if not df_stock.empty and top_codes:
        grp, absents = compute_detention(df_stock, top_codes)
        grp["Alerte"] = grp.apply(compute_alerte, axis=1)
        taux_df = compute_taux(grp, top_codes)
        
        # Affichage (identique à ton design)
        t_all = taux_df[taux_df["flux"] == "ALL"]
        st.metric("Taux Moyen Réseau", f"{t_all['taux'].mean():.1f}%", f"Cible {cible_taux}%")
        
        tab1, tab2, tab3 = st.tabs(["📊 Synthèse", "🚨 Alertes", "🚫 Absents"])
        with tab1:
            st.dataframe(t_all, use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(grp[grp["Alerte"] != "✅ OK"], use_container_width=True, hide_index=True)
        with tab3:
            st.write(f"{len(absents)} articles non trouvés dans l'ERP")
            st.write(absents)
    else:
        st.warning("Vérifiez le contenu des fichiers (colonnes 'Code article', 'Code etat', etc.)")
else:
    st.info("Veuillez charger les fichiers requis.")
