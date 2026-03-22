import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
from datetime import date

st.set_page_config(page_title="Suivi Implantation", layout="wide")

TODAY = date.today().strftime("%d %b %Y")

# ==============================
# HEADER
# ==============================
st.markdown(f"""
<div style="background:#0f1729;color:white;padding:14px 20px;
border-radius:8px;margin-bottom:20px;display:flex;justify-content:space-between">
<b>📦 Suivi Implantation — Data Quality Safe</b>
<span style="color:#94a3b8">{TODAY}</span>
</div>
""", unsafe_allow_html=True)

# ==============================
# AUTO CLEAN
# ==============================
def clean_columns(df):
    df.columns = [
        unicodedata.normalize('NFKD', str(c))
        .encode('ascii', 'ignore')
        .decode('utf-8')
        .strip()
        .upper()
        for c in df.columns
    ]
    return df

def find_column(df, keywords):
    for col in df.columns:
        for k in keywords:
            if k in col:
                return col
    return None

def auto_map_columns(df):
    return {
        "ARTICLE": find_column(df, ["ARTICLE"]),
        "SITE": find_column(df, ["SITE", "MAGASIN"]),
        "STOCK": find_column(df, ["STOCK"]),
        "RAL": find_column(df, ["RAL"]),
    }

def safe_sku(series):
    return (
        series.fillna("")
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
        .str.zfill(8)
    )

# ==============================
# INTRO
# ==============================
st.markdown("""
<div style="background:#f8fafc;border:1px solid #e2e8f0;
padding:18px;border-radius:10px;margin-bottom:20px">

<b>ℹ️ Module intelligent :</b><br>
Correction automatique des fichiers (colonnes, accents, formats)

<b>Statuts :</b>
<ul>
<li>Implanté = stock > 0</li>
<li>Attente = stock = 0 & RAL > 0</li>
<li>Alerte = stock = 0 & RAL = 0</li>
</ul>

</div>
""", unsafe_allow_html=True)

# ==============================
# FILE LOADER
# ==============================
def read_file(file):
    if file.name.endswith(".xlsx"):
        return pd.read_excel(file)
    return pd.read_csv(file, sep=None, engine="python", encoding="latin1")

# ==============================
# SIDEBAR
# ==============================
with st.sidebar:
    st.header("📁 Fichiers")
    t1_file = st.file_uploader("T1")
    stock_files = st.file_uploader("Stocks", accept_multiple_files=True)

# ==============================
# LOAD T1
# ==============================
if not t1_file:
    st.stop()

t1 = read_file(t1_file)
t1 = clean_columns(t1)

map_t1 = auto_map_columns(t1)

if not map_t1["ARTICLE"]:
    st.error(f"❌ Colonne ARTICLE introuvable\n{list(t1.columns)}")
    st.stop()

t1 = t1.rename(columns={map_t1["ARTICLE"]: "ARTICLE"})
t1["SKU"] = safe_sku(t1["ARTICLE"])

st.success("✅ T1 OK")

# ==============================
# LOAD STOCK
# ==============================
if not stock_files:
    st.stop()

dfs = []

for f in stock_files:
    df = read_file(f)
    df = clean_columns(df)

    mapping = auto_map_columns(df)

    if not mapping["ARTICLE"]:
        continue

    df = df.rename(columns={
        mapping["ARTICLE"]: "ARTICLE",
        mapping["SITE"]: "SITE",
        mapping["STOCK"]: "STOCK",
        mapping["RAL"]: "RAL"
    })

    # SAFE CONVERSION
    df["ARTICLE"] = df["ARTICLE"].fillna("").astype(str)

    df["SKU"] = safe_sku(df["ARTICLE"])
    df["STOCK"] = pd.to_numeric(df.get("STOCK", 0), errors="coerce").fillna(0)
    df["RAL"] = pd.to_numeric(df.get("RAL", 0), errors="coerce").fillna(0)

    dfs.append(df)

if not dfs:
    st.error("❌ Aucun fichier stock valide")
    st.stop()

stock = pd.concat(dfs)

st.success("✅ Stock OK")

# ==============================
# MERGE
# ==============================
df = stock.merge(t1[["SKU"]], on="SKU", how="inner")

if df.empty:
    st.warning("⚠️ Aucun matching SKU")
    st.stop()

# ==============================
# STATUT
# ==============================
df["Statut"] = np.select(
    [
        df["STOCK"] > 0,
        (df["STOCK"] == 0) & (df["RAL"] > 0)
    ],
    ["Implanté", "Attente"],
    default="Alerte"
)

# ==============================
# KPI
# ==============================
total = len(df)
ok = (df["Statut"] == "Implanté").sum()
att = (df["Statut"] == "Attente").sum()
alert = (df["Statut"] == "Alerte").sum()

pct = int(ok / total * 100)

c1, c2, c3 = st.columns(3)
c1.metric("Implanté", ok)
c2.metric("Attente", att)
c3.metric("Alerte", alert)

st.progress(pct / 100)

# ==============================
# PAR MAGASIN
# ==============================
pivot = (
    df.groupby(["SITE", "Statut"])
    .size()
    .unstack(fill_value=0)
    .reset_index()
)

pivot["Total"] = pivot.sum(axis=1, numeric_only=True)
pivot["Taux (%)"] = (pivot.get("Implanté", 0) / pivot["Total"] * 100).round(0)

pivot["Taux (%)"] = pivot["Taux (%)"].astype(str) + "%"

st.subheader("🏪 Magasins")
st.dataframe(pivot, use_container_width=True)

# ==============================
# ALERTES
# ==============================
st.subheader("🚨 Alertes")

df_alert = df[df["Statut"] == "Alerte"]

if df_alert.empty:
    st.success("RAS")
else:
    st.dataframe(df_alert, use_container_width=True)
