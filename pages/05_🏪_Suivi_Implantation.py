import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import date
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="Dashboard Implantation COMEX",
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
:root{
  --bg:#f6f8fc;
  --card:#ffffff;
  --border:#e2e8f0;
  --text:#0f172a;
  --muted:#64748b;
  --green:#059669;
  --blue:#2563eb;
  --amber:#d97706;
  --red:#dc2626;
  --green-bg:#ecfdf5;
  --blue-bg:#eff6ff;
  --amber-bg:#fffbeb;
  --red-bg:#fef2f2;
}
html, body, [class*="css"]{
  font-family: Inter, sans-serif !important;
}
.main, section[data-testid="stMain"]{
  background: linear-gradient(180deg,#f8fafc 0%, #f6f8fc 100%) !important;
}
.block-container{
  max-width: 1600px !important;
  padding-top: 1rem !important;
  padding-bottom: 2rem !important;
}
header[data-testid="stHeader"], #MainMenu, footer{
  display:none !important;
}
.topbar{
  background:linear-gradient(135deg,#0f172a,#1e293b);
  border-radius:24px;
  padding:22px 26px;
  margin-bottom:18px;
  color:white;
  box-shadow:0 12px 32px rgba(15,23,42,.20);
}
.topbar-title{
  font-size:28px;
  font-weight:800;
  line-height:1.1;
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
  background:white;
  border:1px solid var(--border);
  border-radius:18px;
  padding:18px;
  box-shadow:0 8px 24px rgba(15,23,42,.06);
}
.kpi-label{
  font-size:11px;
  text-transform:uppercase;
  letter-spacing:.12em;
  font-weight:800;
  color:var(--muted);
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
.section-title{
  margin:24px 0 10px 0;
  font-size:12px;
  text-transform:uppercase;
  letter-spacing:.14em;
  font-weight:900;
  color:var(--muted);
}
.rag-grid{
  display:grid;
  grid-template-columns:repeat(auto-fill,minmax(270px,1fr));
  gap:16px;
  margin-bottom:24px;
}
.rag-card{
  background:white;
  border:1px solid var(--border);
  border-radius:22px;
  padding:18px;
  min-height:160px;
  box-shadow:0 8px 24px rgba(15,23,42,.06);
}
.rag-card.good{background:linear-gradient(180deg,#f7fffb 0%, #ecfdf5 100%);}
.rag-card.mid{background:linear-gradient(180deg,#fffef7 0%, #fffbeb 100%);}
.rag-card.bad{background:linear-gradient(180deg,#fff8f8 0%, #fef2f2 100%);}
.rag-top{
  display:flex;
  justify-content:space-between;
  gap:12px;
  align-items:flex-start;
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
.banner{
  border-radius:18px;
  padding:16px 18px;
  border:1px solid;
  margin-bottom:18px;
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap:18px;
  flex-wrap:wrap;
}
.banner.red{background:var(--red-bg);border-color:#fecaca;}
.banner.blue{background:var(--blue-bg);border-color:#bfdbfe;}
.banner.green{background:var(--green-bg);border-color:#a7f3d0;}
.banner-title{font-size:16px;font-weight:800;}
.banner-sub{font-size:12px;color:#64748b;margin-top:4px;}
.banner-big{font-size:42px;font-weight:900;line-height:1;}
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

    required = {"ARTICLE"}
    missing = required - set(df.columns)
    if missing:
        return None, f"Colonnes manquantes dans IMPLANT : {missing}. Colonnes détectées : {list(df.columns)[:12]}"

    rename_map = {
        "LIBELLÉ ARTICLE": "LIBELLE ARTICLE",
        "LIBELLÉ FOURNISSEUR ORIGINE": "LIBELLE FOURNISSEUR ORIGINE"
    }
    df = df.rename(columns=rename_map)

    optional = [
        "LIBELLE ARTICLE",
        "FOURNISSEUR D'ORIGINE",
        "LIBELLE FOURNISSEUR ORIGINE",
        "MODE APPRO",
        "DATE CDE",
        "DATE LIV.",
        "SEMAINE RECEPTION"
    ]
    for c in optional:
        if c not in df.columns:
            df[c] = ""

    df["SKU"] = df["ARTICLE"].astype(str).str.strip().str.replace(".0", "", regex=False).str.zfill(8).str[-8:]
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
        return None, f"Colonnes manquantes dans STOCK : {missing}. Colonnes détectées : {list(df.columns)[:20]}"

    keep_optional = [
        "Code etat", "Code marketing", "Libellé marketing",
        "Nom fourn.", "Four.", "Site", "Ray", "Famille"
    ]
    for c in keep_optional:
        if c not in df.columns:
            df[c] = ""

    df["SKU"] = df["Code article"].astype(str).str.strip().str.replace(".0", "", regex=False).str.zfill(8).str[-8:]
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

def build_export_excel(detail_df, pivot_df, destructeurs_df, magasin_alert_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail_df.to_excel(writer, index=False, sheet_name="Detail")
        pivot_df.to_excel(writer, index=False, sheet_name="Vue magasins")
        destructeurs_df.to_excel(writer, index=False, sheet_name="Articles destructeurs")
        magasin_alert_df.to_excel(writer, index=False, sheet_name="Top magasins")
    output.seek(0)
    return output.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# TOPBAR
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="topbar">
  <div class="topbar-title">Dashboard Implantation COMEX</div>
  <div class="topbar-sub">Nouvelles références · Détention magasins · Alertes réseau · {TODAY_STR}</div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📁 Chargement des fichiers")
    implant_file = st.file_uploader("Fichier IMPLANT", type=["csv"], key="implant")
    stock_files = st.file_uploader("Fichier(s) STOCK", type=["csv"], accept_multiple_files=True, key="stocks")

if not implant_file:
    st.info("Charge le fichier IMPLANT.")
    st.stop()

if not stock_files:
    st.info("Charge au moins un fichier STOCK.")
    st.stop()

with st.spinner("Lecture IMPLANT…"):
    implant_df, err_implant = load_implant(implant_file.read(), implant_file.name)

if err_implant:
    st.error(err_implant)
    st.stop()

stock_frames = []
with st.spinner(f"Lecture de {len(stock_files)} fichier(s) STOCK…"):
    for f in stock_files:
        df_tmp, err_stock = load_stock(f.read(), f.name)
        if err_stock:
            st.error(f"{f.name} : {err_stock}")
        else:
            stock_frames.append(df_tmp)

if not stock_frames:
    st.error("Aucun fichier STOCK valide.")
    st.stop()

stock_df = pd.concat(stock_frames, ignore_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# FILTRES
# ══════════════════════════════════════════════════════════════════════════════
rayons = sorted(stock_df["Libellé rayon"].dropna().astype(str).unique().tolist())
familles = sorted(stock_df["Libellé famille"].dropna().astype(str).unique().tolist())
magasins = sorted(stock_df["Libellé site"].dropna().astype(str).unique().tolist())
origines = sorted(implant_df["ORIGINE"].dropna().astype(str).unique().tolist())
semaines = sorted([s for s in implant_df["SEMAINE RECEPTION"].unique().tolist() if str(s).strip() not in ("", "nan")], key=sem_sort)

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
# Limiter le stock aux magasins choisis
stock_scope = stock_df[stock_df["Libellé site"].isin(mag_sel)].copy()

# Articles implant à retenir
implant_scope = implant_df[implant_df["ORIGINE"].isin(origine_sel)].copy()
if semaine_sel:
    implant_scope = implant_scope[implant_scope["SEMAINE RECEPTION"].isin(semaine_sel)].copy()

implant_sku = implant_scope["SKU"].unique().tolist()

# On récupère les attributs rayon/famille depuis le stock pour les SKU de l'implant
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

# filtres rayon/famille sur le périmètre implant
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

# base magasin × SKU implant
base_df = pd.MultiIndex.from_product(
    [mag_sel, sku_scope], names=["Magasin", "SKU"]
).to_frame(index=False)

# stock à joindre
stock_join = stock_scope[stock_scope["SKU"].isin(sku_scope)].copy()

stock_join = stock_join.rename(columns={
    "Libellé site": "Magasin",
    "Nouveau stock": "Stock",
    "Ral": "RAL"
})

cols_stock_join = [
    "Magasin", "SKU", "Stock", "RAL", "Libellé rayon", "Libellé famille",
    "Libellé article", "Code etat", "Code marketing", "Libellé marketing",
    "Nom fourn.", "Four."
]
stock_join = stock_join[cols_stock_join].drop_duplicates(subset=["Magasin", "SKU"])

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

# statut
conds = [
    detail_df["Stock"] > 0,
    (detail_df["Stock"] <= 0) & (detail_df["RAL"] > 0)
]
choices = ["Implantation Terminée", "En Attente Livraison"]
detail_df["Statut"] = np.select(conds, choices, default="Alerte Aucun Mouvement")

# ══════════════════════════════════════════════════════════════════════════════
# KPI
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

ct = int(pivot["Implantation Terminée"].sum())
ca = int(pivot["En Attente Livraison"].sum())
cal = int(pivot["Alerte Aucun Mouvement"].sum())
total_cells = len(mag_sel) * total_sku
avg_impl = int(round(pivot["Taux (%)"].mean(), 0)) if not pivot.empty else 0

# top magasins en alerte
top_magasins_alerte = (
    pivot[["Magasin", "Alerte Aucun Mouvement", "En Attente Livraison", "Implantation Terminée", "Taux (%)"]]
    .sort_values(["Alerte Aucun Mouvement", "En Attente Livraison"], ascending=False)
    .reset_index(drop=True)
)

# articles destructeurs
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

top20_destructeurs = articles_destructeurs.head(20).copy()

# export
export_bytes = build_export_excel(detail_df, pivot, articles_destructeurs, top_magasins_alerte)

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
        <div class="banner-sub">{cal} alertes sans mouvement · {ca} en attente livraison</div>
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
# SCORECARDS MAGASINS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">Scorecard magasins</div>', unsafe_allow_html=True)

rag_html = '<div class="rag-grid">'
for _, row in pivot.sort_values("Taux (%)", ascending=False).iterrows():
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
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Vue COMEX",
    "🚨 Top magasins en alerte",
    "📉 Articles destructeurs",
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

    st.dataframe(
        pivot.sort_values("Taux (%)", ascending=False).reset_index(drop=True),
        use_container_width=True,
        hide_index=True
    )

with tab2:
    st.markdown("### Top magasins en alerte")
    st.dataframe(
        top_magasins_alerte.head(15),
        use_container_width=True,
        hide_index=True
    )

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
    st.dataframe(
        top20_destructeurs,
        use_container_width=True,
        hide_index=True
    )

    top_graph = top20_destructeurs.head(10).copy()
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
    st.markdown("### Détail opérationnel")
    st.dataframe(
        detail_df.sort_values(["Magasin", "Statut", "Libellé rayon", "Libellé famille"]).reset_index(drop=True),
        use_container_width=True,
        hide_index=True,
        height=520
    )

    st.markdown('<div class="section-title">Export Excel</div>', unsafe_allow_html=True)
    st.download_button(
        label="📥 Télécharger le pack Excel COMEX",
        data=export_bytes,
        file_name=f"dashboard_implantation_comex_{TODAY_FILE}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
