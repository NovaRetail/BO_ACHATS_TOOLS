import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import date
import plotly.express as px
import plotly.graph_objects as go

# ══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Dashboard Implantation COPIL",
    page_icon="🏪",
    layout="wide",
    initial_sidebar_state="expanded"
)

TODAY = date.today()
TODAY_STR = TODAY.strftime("%d/%m/%Y")
TODAY_FILE = TODAY.strftime("%Y%m%d")

# ══════════════════════════════════════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

:root{
  --bg:#f6f8fc;
  --surface:#ffffff;
  --border:#e2e8f0;
  --text:#0f172a;
  --muted:#64748b;

  --green:#059669;
  --blue:#2563eb;
  --cyan:#0284c7;
  --amber:#d97706;
  --red:#dc2626;

  --green-bg:#ecfdf5;
  --blue-bg:#eff6ff;
  --cyan-bg:#f0f9ff;
  --amber-bg:#fffbeb;
  --red-bg:#fef2f2;
}

html, body, [class*="css"]{
  font-family:'Inter', sans-serif !important;
  color:var(--text) !important;
}
.main, section[data-testid="stMain"]{
  background:linear-gradient(180deg,#f8fafc 0%, #f6f8fc 100%) !important;
}
.block-container{
  max-width:1600px !important;
  padding-top:1rem !important;
  padding-bottom:2.5rem !important;
}
header[data-testid="stHeader"], #MainMenu, footer{
  display:none !important;
}
section[data-testid="stSidebar"]{
  background:linear-gradient(180deg,#ffffff 0%, #f8fbff 100%) !important;
  border-right:1px solid var(--border) !important;
}
.topbar{
  background:linear-gradient(135deg,#0f172a,#1e293b);
  border-radius:24px;
  padding:24px 28px;
  margin-bottom:20px;
  color:#fff;
  box-shadow:0 18px 40px rgba(15,23,42,.18);
}
.topbar-title{
  font-size:28px;
  font-weight:900;
  line-height:1.05;
}
.topbar-sub{
  margin-top:6px;
  color:#cbd5e1;
  font-size:12px;
}
.kpi-grid{
  display:grid;
  grid-template-columns:repeat(4,minmax(0,1fr));
  gap:14px;
  margin-bottom:18px;
}
.kpi-card{
  background:#fff;
  border:1px solid var(--border);
  border-radius:18px;
  padding:18px;
  box-shadow:0 8px 24px rgba(15,23,42,.06);
}
.kpi-label{
  font-size:11px;
  text-transform:uppercase;
  letter-spacing:.12em;
  color:var(--muted);
  font-weight:800;
  margin-bottom:10px;
}
.kpi-value{
  font-size:38px;
  font-weight:900;
  line-height:1;
}
.kpi-sub{
  margin-top:8px;
  font-size:12px;
  color:var(--muted);
}
.banner{
  border-radius:18px;
  padding:16px 18px;
  border:1px solid;
  margin-bottom:18px;
  display:flex;
  justify-content:space-between;
  gap:18px;
  flex-wrap:wrap;
  align-items:center;
}
.banner.red{background:var(--red-bg);border-color:#fecaca;}
.banner.green{background:var(--green-bg);border-color:#a7f3d0;}
.banner.blue{background:var(--blue-bg);border-color:#bfdbfe;}
.banner.amber{background:var(--amber-bg);border-color:#fcd34d;}
.banner-title{font-size:16px;font-weight:800;}
.banner-sub{font-size:12px;color:#64748b;margin-top:4px;}
.banner-big{font-size:42px;font-weight:900;line-height:1;}

.section-title{
  margin:24px 0 10px 0;
  font-size:12px;
  text-transform:uppercase;
  letter-spacing:.14em;
  color:var(--muted);
  font-weight:900;
}

.rag-grid{
  display:grid;
  grid-template-columns:repeat(auto-fill,minmax(270px,1fr));
  gap:16px;
  margin-bottom:24px;
}
.rag-card{
  background:#fff;
  border:1px solid var(--border);
  border-radius:22px;
  padding:18px;
  min-height:160px;
  box-shadow:0 8px 24px rgba(15,23,42,.06);
}
.rag-card.good{
  background:linear-gradient(180deg,#f7fffb 0%, #ecfdf5 100%);
}
.rag-card.mid{
  background:linear-gradient(180deg,#fffef7 0%, #fffbeb 100%);
}
.rag-card.bad{
  background:linear-gradient(180deg,#fff8f8 0%, #fef2f2 100%);
}
.rag-top{
  display:flex;
  justify-content:space-between;
  align-items:flex-start;
  gap:12px;
}
.rag-name{
  font-size:14px;
  font-weight:800;
  color:#0f172a;
  max-width:78%;
  white-space:normal;
  word-break:break-word;
}
.rag-chip{
  padding:6px 10px;
  border-radius:999px;
  font-size:10px;
  font-weight:900;
}
.rag-chip.good{background:#d1fae5;color:#065f46;}
.rag-chip.mid{background:#fef3c7;color:#92400e;}
.rag-chip.bad{background:#fee2e2;color:#991b1b;}
.rag-pct{
  font-size:46px;
  font-weight:900;
  line-height:1;
  margin-top:14px;
}
.rag-progress{
  margin-top:14px;
  height:10px;
  border-radius:999px;
  background:#e5e7eb;
  overflow:hidden;
}
.rag-fill{
  height:100%;
  border-radius:999px;
}
.rag-foot{
  display:flex;
  justify-content:space-between;
  gap:10px;
  flex-wrap:wrap;
  margin-top:12px;
  font-size:11px;
  color:#64748b;
}
.comment-box{
  background:#fff;
  border:1px solid var(--border);
  border-radius:18px;
  padding:18px;
  box-shadow:0 8px 24px rgba(15,23,42,.06);
  line-height:1.6;
}
.stDownloadButton > button{
  width:100% !important;
  border:none !important;
  border-radius:14px !important;
  padding:12px 16px !important;
  background:linear-gradient(135deg,#0f172a,#1e293b) !important;
  color:#fff !important;
  font-weight:800 !important;
}
@media (max-width:1200px){
  .kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr));}
}
@media (max-width:780px){
  .kpi-grid{grid-template-columns:1fr;}
  .rag-grid{grid-template-columns:1fr;}
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = (
        df.columns.astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.replace("ï»¿", "", regex=False)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )
    return df

def fix_encoding_cols(df: pd.DataFrame) -> pd.DataFrame:
    repaired = []
    for c in df.columns:
        try:
            if "Ã" in c or "ï»¿" in c:
                repaired.append(c.encode("latin1").decode("utf-8", errors="ignore"))
            else:
                repaired.append(c)
        except Exception:
            repaired.append(c)
    df.columns = repaired
    return df

def read_csv_robust(file_bytes: bytes) -> pd.DataFrame:
    for enc in ("utf-8-sig", "latin1", "cp1252"):
        try:
            return pd.read_csv(io.BytesIO(file_bytes), sep=";", encoding=enc, low_memory=False)
        except Exception:
            continue
    return pd.read_csv(io.BytesIO(file_bytes), sep=None, engine="python", low_memory=False)

@st.cache_data(show_spinner=False)
def load_implant(file_bytes: bytes, filename: str):
    df = read_csv_robust(file_bytes)
    df = fix_encoding_cols(df)
    df = normalize_columns(df)
    df.columns = [c.upper() for c in df.columns]

    if "ARTICLE" not in df.columns:
        return None, f"Colonne ARTICLE absente dans {filename}"

    rename_map = {
        "LIBELLÉ ARTICLE": "LIBELLE ARTICLE",
        "LIBELLÉ FOURNISSEUR ORIGINE": "LIBELLE FOURNISSEUR ORIGINE"
    }
    df = df.rename(columns=rename_map)

    optional_cols = [
        "LIBELLE ARTICLE",
        "MODE APPRO",
        "DATE LIV.",
        "SEMAINE RECEPTION",
        "LIBELLE FOURNISSEUR ORIGINE"
    ]
    for c in optional_cols:
        if c not in df.columns:
            df[c] = ""

    df["SKU"] = (
        df["ARTICLE"]
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
        .str.zfill(8)
        .str[-8:]
    )
    df = df[df["SKU"].str.match(r"^\d{8}$", na=False)].copy()

    df["ORIGINE"] = np.where(
        df["MODE APPRO"].astype(str).str.upper().str.contains("IMPORT", na=False),
        "IM",
        "LO"
    )
    df["SEMAINE RECEPTION"] = df["SEMAINE RECEPTION"].astype(str).str.strip().replace("nan", "")

    return df.drop_duplicates(subset=["SKU"]).copy(), None

@st.cache_data(show_spinner=False)
def load_stock(file_bytes: bytes, filename: str):
    df = read_csv_robust(file_bytes)
    df = fix_encoding_cols(df)
    df = normalize_columns(df)

    required = {"Libellé site", "Code article", "Libellé rayon", "Libellé famille", "Nouveau stock", "Ral"}
    missing = required - set(df.columns)
    if missing:
        return None, f"{filename} : colonnes manquantes {missing}"

    optional = [
        "Code etat", "Code marketing", "Libellé marketing",
        "Nom fourn.", "Four.", "Libellé article"
    ]
    for c in optional:
        if c not in df.columns:
            df[c] = ""

    df["SKU"] = (
        df["Code article"]
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
        .str.zfill(8)
        .str[-8:]
    )
    df["Nouveau stock"] = pd.to_numeric(df["Nouveau stock"], errors="coerce").fillna(0)
    df["Ral"] = pd.to_numeric(df["Ral"], errors="coerce").fillna(0)

    return df.copy(), None

def sem_sort(s):
    try:
        return int(str(s).strip().upper().replace("S", ""))
    except Exception:
        return 999

def safe_pct(num, den):
    return int(round(num / den * 100, 0)) if den else 0

def classify_taux(t):
    if t >= 80:
        return "good", "ON TRACK", "#059669", "linear-gradient(90deg,#10b981,#059669)"
    if t >= 65:
        return "mid", "WATCH", "#d97706", "linear-gradient(90deg,#f59e0b,#d97706)"
    return "bad", "RISK", "#dc2626", "linear-gradient(90deg,#ef4444,#dc2626)"

def diagnostic_statut(taux, alerte_share):
    if taux >= 80 and alerte_share < 10:
        return "Situation maîtrisée"
    if taux < 65 and alerte_share >= 20:
        return "Risque fort de non-implantation"
    if taux < 80 and alerte_share >= 10:
        return "Situation sous tension"
    return "Situation à surveiller"

def build_copil_comment(avg_impl, cal, ca, top_rayon_alertes, im_taux, lo_taux, worst_store):
    parts = []

    if avg_impl >= 80:
        parts.append(f"Le niveau global d’implantation ressort à {avg_impl}%, ce qui traduit une situation globalement maîtrisée sur le périmètre suivi.")
    elif avg_impl >= 65:
        parts.append(f"Le niveau global d’implantation ressort à {avg_impl}%, ce qui montre une situation sous tension nécessitant un suivi rapproché.")
    else:
        parts.append(f"Le niveau global d’implantation ressort à {avg_impl}%, ce qui traduit un niveau critique d’exécution en magasin.")

    if cal > 0:
        parts.append(f"On dénombre {cal} alertes sans mouvement, qui constituent le principal point de vigilance opérationnel.")
    if ca > 0:
        parts.append(f"{ca} lignes sont encore en attente de livraison, ce qui indique qu’une partie du retard reste liée au pipeline d’approvisionnement.")

    if im_taux > lo_taux:
        parts.append(f"Le flux IMPORT affiche une meilleure exécution ({im_taux}%) que le flux LOCAL ({lo_taux}%).")
    elif lo_taux > im_taux:
        parts.append(f"Le flux LOCAL affiche une meilleure exécution ({lo_taux}%) que le flux IMPORT ({im_taux}%).")
    else:
        parts.append(f"Les flux IMPORT et LOCAL affichent un niveau d’exécution comparable ({im_taux}% / {lo_taux}%).")

    if top_rayon_alertes:
        parts.append(f"Le rayon le plus exposé est {top_rayon_alertes}.")
    if worst_store:
        parts.append(f"Le magasin le plus en difficulté à date est {worst_store}.")

    parts.append("La priorité COPIL doit porter sur les articles sans mouvement, la sécurisation des livraisons en cours et la levée ciblée des alertes sur les magasins les moins performants.")

    return " ".join(parts)

def build_export_excel(detail_df, pivot_df, destructeurs_df, top_mags_df, rayon_df, origine_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail_df.to_excel(writer, index=False, sheet_name="Detail")
        pivot_df.to_excel(writer, index=False, sheet_name="Magasins")
        top_mags_df.to_excel(writer, index=False, sheet_name="Top magasins")
        destructeurs_df.to_excel(writer, index=False, sheet_name="Destructeurs")
        rayon_df.to_excel(writer, index=False, sheet_name="Rayons")
        origine_df.to_excel(writer, index=False, sheet_name="IM vs LO")
    output.seek(0)
    return output.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# TOPBAR
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="topbar">
  <div class="topbar-title">Dashboard Implantation COPIL</div>
  <div class="topbar-sub">Pilotage réseau · Détention · Alertes · Analyse COPIL · {TODAY_STR}</div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📁 Chargement")
    implant_file = st.file_uploader("Fichier IMPLANT", type=["csv"], key="implant")
    stock_files = st.file_uploader("Fichier(s) STOCK", type=["csv"], accept_multiple_files=True, key="stocks")

if not implant_file:
    st.info("Charge le fichier IMPLANT.")
    st.stop()

if not stock_files:
    st.info("Charge au moins un fichier STOCK.")
    st.stop()

with st.spinner("Lecture du fichier IMPLANT…"):
    implant_df, err_implant = load_implant(implant_file.read(), implant_file.name)

if err_implant:
    st.error(err_implant)
    st.stop()

frames = []
with st.spinner(f"Lecture de {len(stock_files)} fichier(s) STOCK…"):
    for f in stock_files:
        df_tmp, err_stock = load_stock(f.read(), f.name)
        if err_stock:
            st.error(err_stock)
        else:
            frames.append(df_tmp)

if not frames:
    st.error("Aucun fichier STOCK valide.")
    st.stop()

stock_df = pd.concat(frames, ignore_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# FILTRES
# ══════════════════════════════════════════════════════════════════════════════
rayons = sorted(stock_df["Libellé rayon"].dropna().astype(str).unique().tolist())
familles = sorted(stock_df["Libellé famille"].dropna().astype(str).unique().tolist())
magasins = sorted(stock_df["Libellé site"].dropna().astype(str).unique().tolist())
origines = sorted(implant_df["ORIGINE"].dropna().astype(str).unique().tolist())
semaines = sorted(
    [s for s in implant_df["SEMAINE RECEPTION"].unique().tolist() if str(s).strip() not in ("", "nan")],
    key=sem_sort
)

with st.sidebar:
    st.markdown("---")
    st.markdown("### 🔎 Filtres")
    mag_sel = st.multiselect("Magasins", magasins, default=magasins)
    rayon_sel = st.multiselect("Rayons", rayons, default=rayons)
    famille_sel = st.multiselect("Familles", familles, default=familles)
    origine_sel = st.multiselect("Origine", origines, default=origines)
    semaine_sel = st.multiselect("Semaine réception", semaines, default=semaines if semaines else [])

if not mag_sel:
    st.warning("Sélectionne au moins un magasin.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# PREP DATA
# ══════════════════════════════════════════════════════════════════════════════
stock_scope = stock_df[stock_df["Libellé site"].isin(mag_sel)].copy()

implant_scope = implant_df[implant_df["ORIGINE"].isin(origine_sel)].copy()
if semaine_sel:
    implant_scope = implant_scope[implant_scope["SEMAINE RECEPTION"].isin(semaine_sel)].copy()

implant_sku = implant_scope["SKU"].unique().tolist()

sku_attr = (
    stock_scope[stock_scope["SKU"].isin(implant_sku)]
    .sort_values(["SKU"])
    .groupby("SKU", as_index=False)
    .agg({
        "Libellé rayon": "first",
        "Libellé famille": "first",
        "Code marketing": "first",
        "Libellé article": "first"
    })
)

implant_scope = implant_scope.merge(sku_attr, on="SKU", how="left")
implant_scope["Libellé rayon"] = implant_scope["Libellé rayon"].fillna("Non classé")
implant_scope["Libellé famille"] = implant_scope["Libellé famille"].fillna("Non classé")

implant_scope = implant_scope[
    implant_scope["Libellé rayon"].isin(rayon_sel) &
    implant_scope["Libellé famille"].isin(famille_sel)
].copy()

if implant_scope.empty:
    st.warning("Aucun article implant ne correspond aux filtres.")
    st.stop()

sku_scope = implant_scope["SKU"].unique().tolist()
total_sku = len(sku_scope)

base_df = pd.MultiIndex.from_product(
    [mag_sel, sku_scope], names=["Magasin", "SKU"]
).to_frame(index=False)

stock_join = stock_scope[stock_scope["SKU"].isin(sku_scope)].copy()
stock_join = stock_join.rename(columns={
    "Libellé site": "Magasin",
    "Nouveau stock": "Stock",
    "Ral": "RAL"
})

cols_stock = [
    "Magasin", "SKU", "Stock", "RAL", "Libellé rayon", "Libellé famille",
    "Libellé article", "Code etat", "Code marketing", "Libellé marketing",
    "Nom fourn.", "Four."
]
stock_join = stock_join[cols_stock].drop_duplicates(subset=["Magasin", "SKU"])

detail_df = base_df.merge(stock_join, on=["Magasin", "SKU"], how="left")
detail_df = detail_df.merge(
    implant_scope[[
        "SKU", "LIBELLE ARTICLE", "MODE APPRO", "DATE LIV.", "SEMAINE RECEPTION",
        "ORIGINE", "LIBELLE FOURNISSEUR ORIGINE"
    ]].drop_duplicates(subset=["SKU"]),
    on="SKU",
    how="left"
)

detail_df["Stock"] = detail_df["Stock"].fillna(0)
detail_df["RAL"] = detail_df["RAL"].fillna(0)
detail_df["Code etat"] = detail_df["Code etat"].fillna("").astype(str)
detail_df["Libellé rayon"] = detail_df["Libellé rayon"].fillna("Non classé")
detail_df["Libellé famille"] = detail_df["Libellé famille"].fillna("Non classé")
detail_df["Libellé article"] = detail_df["Libellé article"].fillna(detail_df["LIBELLE ARTICLE"]).fillna("")

conds = [
    detail_df["Stock"] > 0,
    (detail_df["Stock"] <= 0) & (detail_df["RAL"] > 0)
]
choices = ["Implantation Terminée", "En Attente Livraison"]
detail_df["Statut"] = np.select(conds, choices, default="Alerte Aucun Mouvement")

# ══════════════════════════════════════════════════════════════════════════════
# KPI GLOBAUX
# ══════════════════════════════════════════════════════════════════════════════
pivot = (
    detail_df.groupby(["Magasin", "Statut"]).size()
    .unstack(fill_value=0)
    .reset_index()
)

for c in ["Implantation Terminée", "En Attente Livraison", "Alerte Aucun Mouvement"]:
    if c not in pivot.columns:
        pivot[c] = 0

pivot["Total"] = total_sku
pivot["Taux (%)"] = ((pivot["Implantation Terminée"] / pivot["Total"]) * 100).round(0).astype(int)
pivot["Alerte (%)"] = ((pivot["Alerte Aucun Mouvement"] / pivot["Total"]) * 100).round(1)
pivot["Diagnostic"] = pivot.apply(
    lambda x: diagnostic_statut(x["Taux (%)"], x["Alerte (%)"]),
    axis=1
)

ct = int(pivot["Implantation Terminée"].sum())
ca = int(pivot["En Attente Livraison"].sum())
cal = int(pivot["Alerte Aucun Mouvement"].sum())
total_cells = len(mag_sel) * total_sku
avg_impl = int(round(pivot["Taux (%)"].mean(), 0)) if not pivot.empty else 0

# ══════════════════════════════════════════════════════════════════════════════
# ANALYSES COMPLÉMENTAIRES
# ══════════════════════════════════════════════════════════════════════════════
top_magasins_alerte = (
    pivot[["Magasin", "Alerte Aucun Mouvement", "En Attente Livraison", "Implantation Terminée", "Taux (%)", "Diagnostic"]]
    .sort_values(["Alerte Aucun Mouvement", "En Attente Livraison", "Taux (%)"], ascending=[False, False, True])
    .reset_index(drop=True)
)

articles_destructeurs = (
    detail_df[detail_df["Statut"] == "Alerte Aucun Mouvement"]
    .groupby(["SKU", "Libellé article", "Libellé rayon", "Libellé famille", "ORIGINE"], as_index=False)
    .agg(
        Nb_Magasins_Alertes=("Magasin", "nunique"),
        Total_Stock=("Stock", "sum"),
        Total_RAL=("RAL", "sum")
    )
    .sort_values(["Nb_Magasins_Alertes", "Total_Stock"], ascending=[False, True])
    .reset_index(drop=True)
)

rayon_perf = (
    detail_df.groupby(["Libellé rayon", "Statut"]).size()
    .unstack(fill_value=0)
    .reset_index()
)
for c in ["Implantation Terminée", "En Attente Livraison", "Alerte Aucun Mouvement"]:
    if c not in rayon_perf.columns:
        rayon_perf[c] = 0
rayon_perf["Total"] = (
    rayon_perf["Implantation Terminée"] +
    rayon_perf["En Attente Livraison"] +
    rayon_perf["Alerte Aucun Mouvement"]
)
rayon_perf["Taux (%)"] = ((rayon_perf["Implantation Terminée"] / rayon_perf["Total"]) * 100).round(0).astype(int)
rayon_perf["Alerte (%)"] = ((rayon_perf["Alerte Aucun Mouvement"] / rayon_perf["Total"]) * 100).round(1)
rayon_perf["Diagnostic"] = rayon_perf.apply(
    lambda x: diagnostic_statut(x["Taux (%)"], x["Alerte (%)"]),
    axis=1
)
rayon_perf = rayon_perf.sort_values(["Alerte Aucun Mouvement", "Taux (%)"], ascending=[False, True]).reset_index(drop=True)

origine_perf = (
    detail_df.groupby(["ORIGINE", "Statut"]).size()
    .unstack(fill_value=0)
    .reset_index()
)
for c in ["Implantation Terminée", "En Attente Livraison", "Alerte Aucun Mouvement"]:
    if c not in origine_perf.columns:
        origine_perf[c] = 0
origine_perf["Total"] = (
    origine_perf["Implantation Terminée"] +
    origine_perf["En Attente Livraison"] +
    origine_perf["Alerte Aucun Mouvement"]
)
origine_perf["Taux (%)"] = ((origine_perf["Implantation Terminée"] / origine_perf["Total"]) * 100).round(0).astype(int)
origine_perf = origine_perf.sort_values("ORIGINE").reset_index(drop=True)

im_taux = int(origine_perf.loc[origine_perf["ORIGINE"] == "IM", "Taux (%)"].iloc[0]) if "IM" in origine_perf["ORIGINE"].values else 0
lo_taux = int(origine_perf.loc[origine_perf["ORIGINE"] == "LO", "Taux (%)"].iloc[0]) if "LO" in origine_perf["ORIGINE"].values else 0

top_rayon_alertes = rayon_perf.iloc[0]["Libellé rayon"] if not rayon_perf.empty else ""
worst_store = top_magasins_alerte.iloc[0]["Magasin"] if not top_magasins_alerte.empty else ""

copil_comment = build_copil_comment(
    avg_impl=avg_impl,
    cal=cal,
    ca=ca,
    top_rayon_alertes=top_rayon_alertes,
    im_taux=im_taux,
    lo_taux=lo_taux,
    worst_store=worst_store
)

export_bytes = build_export_excel(
    detail_df=detail_df,
    pivot_df=pivot,
    destructeurs_df=articles_destructeurs,
    top_mags_df=top_magasins_alerte,
    rayon_df=rayon_perf,
    origine_df=origine_perf
)

# ══════════════════════════════════════════════════════════════════════════════
# KPI HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card">
    <div class="kpi-label">Taux moyen réseau</div>
    <div class="kpi-value" style="color:#2563eb">{avg_impl}%</div>
    <div class="kpi-sub">{len(mag_sel)} magasins · {total_sku} SKU implant</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">Implantation terminée</div>
    <div class="kpi-value" style="color:#059669">{ct}</div>
    <div class="kpi-sub">{safe_pct(ct,total_cells)}% du périmètre</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">En attente livraison</div>
    <div class="kpi-value" style="color:#0284c7">{ca}</div>
    <div class="kpi-sub">{safe_pct(ca,total_cells)}% du périmètre</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">Alertes aucun mouvement</div>
    <div class="kpi-value" style="color:#dc2626">{cal}</div>
    <div class="kpi-sub">{safe_pct(cal,total_cells)}% du périmètre</div>
  </div>
</div>
""", unsafe_allow_html=True)

if cal > 0 or ca > 0:
    st.markdown(f"""
    <div class="banner red">
      <div>
        <div class="banner-title" style="color:#dc2626">⚠️ Réseau sous tension</div>
        <div class="banner-sub">{cal} alertes sans mouvement · {ca} lignes en attente livraison</div>
      </div>
      <div class="banner-big" style="color:#dc2626">{cal + ca}</div>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <div class="banner green">
      <div>
        <div class="banner-title" style="color:#059669">✅ Réseau sous contrôle</div>
        <div class="banner-sub">Aucune alerte critique sur le périmètre filtré.</div>
      </div>
      <div class="banner-big" style="color:#059669">0</div>
    </div>
    """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# COMMENTAIRE AUTO COPIL
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">Commentaire automatique COPIL</div>', unsafe_allow_html=True)
st.markdown(f'<div class="comment-box">{copil_comment}</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SCORECARDS MAGASINS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">Scorecard magasins</div>', unsafe_allow_html=True)

rag_html = '<div class="rag-grid">'
for _, row in pivot.sort_values(["Alerte Aucun Mouvement", "Taux (%)"], ascending=[False, True]).iterrows():
    taux = int(row["Taux (%)"])
    cls, lbl, color_hex, progress = classify_taux(taux)
    rag_html += f"""
    <div class="rag-card {cls}">
      <div class="rag-top">
        <div class="rag-name">{row['Magasin']}</div>
        <div class="rag-chip {cls}">{lbl}</div>
      </div>
      <div class="rag-pct" style="color:{color_hex}">{taux}%</div>
      <div class="rag-progress">
        <div class="rag-fill" style="width:{min(taux,100)}%;background:{progress};"></div>
      </div>
      <div class="rag-foot">
        <div>{int(row['Implantation Terminée'])} terminés</div>
        <div>{int(row['En Attente Livraison'])} attente</div>
        <div>{int(row['Alerte Aucun Mouvement'])} alertes</div>
      </div>
    </div>
    """
rag_html += "</div>"
st.markdown(rag_html, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Vue COPIL",
    "🚨 Top magasins en alerte",
    "📉 Articles destructeurs",
    "🏷️ Analyse rayons",
    "📋 Détail & Export"
])

PLOTLY_LAYOUT = dict(
    paper_bgcolor="#ffffff",
    plot_bgcolor="#ffffff",
    font=dict(color="#64748b", size=12),
    margin=dict(l=20, r=20, t=50, b=20)
)

with tab1:
    c1, c2 = st.columns([3, 2])

    with c1:
        mel = pivot.melt(
            id_vars="Magasin",
            value_vars=["Implantation Terminée", "En Attente Livraison", "Alerte Aucun Mouvement"],
            var_name="Statut",
            value_name="N"
        )
        fig = px.bar(
            mel,
            x="Magasin",
            y="N",
            color="Statut",
            barmode="stack",
            color_discrete_map={
                "Implantation Terminée": "#059669",
                "En Attente Livraison": "#0284c7",
                "Alerte Aucun Mouvement": "#dc2626"
            },
            title="Situation par magasin"
        )
        fig.update_traces(textposition="inside", texttemplate="%{y}", textfont_color="white")
        fig.update_layout(**PLOTLY_LAYOUT, height=430, legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig_d = go.Figure(go.Pie(
            labels=["Terminée", "Attente", "Alerte"],
            values=[ct, ca, cal],
            hole=0.68,
            marker=dict(colors=["#059669", "#0284c7", "#dc2626"], line=dict(color="#fff", width=3))
        ))
        fig_d.add_annotation(
            text=f"<b>{avg_impl}%</b><br>implanté",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=20, color="#0f172a")
        )
        fig_d.update_layout(**PLOTLY_LAYOUT, height=430, title="Répartition globale")
        st.plotly_chart(fig_d, use_container_width=True)

    st.markdown("### Analyse IM vs LO")
    st.dataframe(origine_perf, use_container_width=True, hide_index=True)

    fig_orig = go.Figure(go.Bar(
        x=origine_perf["ORIGINE"],
        y=origine_perf["Taux (%)"],
        marker=dict(color=["#2563eb" if x == "IM" else "#059669" for x in origine_perf["ORIGINE"]]),
        text=origine_perf["Taux (%)"].astype(str) + "%",
        textposition="outside"
    ))
    fig_orig.update_layout(**PLOTLY_LAYOUT, height=350, title="Taux d’implantation IM vs LO")
    st.plotly_chart(fig_orig, use_container_width=True)

with tab2:
    st.markdown("### Top magasins en alerte")
    st.dataframe(top_magasins_alerte.head(15), use_container_width=True, hide_index=True)

    fig_alert = go.Figure(go.Bar(
        x=top_magasins_alerte.head(10)["Magasin"],
        y=top_magasins_alerte.head(10)["Alerte Aucun Mouvement"],
        marker=dict(color="#dc2626"),
        text=top_magasins_alerte.head(10)["Alerte Aucun Mouvement"],
        textposition="outside"
    ))
    fig_alert.update_layout(**PLOTLY_LAYOUT, height=420, title="Top 10 magasins - alertes aucun mouvement")
    st.plotly_chart(fig_alert, use_container_width=True)

with tab3:
    st.markdown("### Articles destructeurs")
    st.dataframe(articles_destructeurs.head(20), use_container_width=True, hide_index=True)

    top_graph = articles_destructeurs.head(10).copy()
    if not top_graph.empty:
        top_graph["Label"] = top_graph["SKU"] + " - " + top_graph["Libellé article"].astype(str).str[:35]
        fig_dest = go.Figure(go.Bar(
            x=top_graph["Nb_Magasins_Alertes"],
            y=top_graph["Label"],
            orientation="h",
            marker=dict(color="#dc2626"),
            text=top_graph["Nb_Magasins_Alertes"],
            textposition="outside"
        ))
        fig_dest.update_layout(**PLOTLY_LAYOUT, height=500, title="Top 10 articles destructeurs")
        st.plotly_chart(fig_dest, use_container_width=True)

with tab4:
    st.markdown("### Performance par rayon")
    st.dataframe(rayon_perf, use_container_width=True, hide_index=True)

    fig_rayon = go.Figure(go.Bar(
        x=rayon_perf.head(12)["Libellé rayon"],
        y=rayon_perf.head(12)["Alerte Aucun Mouvement"],
        marker=dict(color="#dc2626"),
        text=rayon_perf.head(12)["Alerte Aucun Mouvement"],
        textposition="outside"
    ))
    fig_rayon.update_layout(**PLOTLY_LAYOUT, height=420, title="Rayons les plus exposés aux alertes")
    st.plotly_chart(fig_rayon, use_container_width=True)

    fig_rayon_taux = go.Figure(go.Bar(
        x=rayon_perf.sort_values("Taux (%)", ascending=True).head(12)["Libellé rayon"],
        y=rayon_perf.sort_values("Taux (%)", ascending=True).head(12)["Taux (%)"],
        marker=dict(color="#d97706"),
        text=rayon_perf.sort_values("Taux (%)", ascending=True).head(12)["Taux (%)"].astype(str) + "%",
        textposition="outside"
    ))
    fig_rayon_taux.update_layout(**PLOTLY_LAYOUT, height=420, title="Rayons les moins performants")
    st.plotly_chart(fig_rayon_taux, use_container_width=True)

with tab5:
    st.markdown("### Détail opérationnel")
    st.dataframe(
        detail_df.sort_values(["Magasin", "Statut", "Libellé rayon", "Libellé famille"]).reset_index(drop=True),
        use_container_width=True,
        hide_index=True,
        height=520
    )

    st.markdown('<div class="section-title">Export Excel COPIL</div>', unsafe_allow_html=True)
    st.download_button(
        label="📥 Télécharger le pack Excel COPIL",
        data=export_bytes,
        file_name=f"dashboard_implantation_copil_{TODAY_FILE}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
