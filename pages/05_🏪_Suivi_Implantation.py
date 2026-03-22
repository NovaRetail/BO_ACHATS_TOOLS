import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
import io
from datetime import date

# ==============================
# CONFIG
# ==============================
st.set_page_config(page_title="Suivi Implantation", layout="wide")

TODAY = date.today().strftime("%d %b %Y")

# ==============================
# HEADER
# ==============================
st.markdown(f"""
<div style="background:#0f1729;color:white;padding:14px 20px;
border-radius:8px;margin-bottom:20px;display:flex;justify-content:space-between">
<b>📦 Suivi Implantation — Data Quality Ready</b>
<span style="color:#94a3b8">{TODAY}</span>
</div>
""", unsafe_allow_html=True)

# ==============================
# AUTO CLEAN DATA MODULE
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
    mapping = {}

    mapping["ARTICLE"] = find_column(df, ["ARTICLE", "CODE"])
    mapping["SITE"] = find_column(df, ["SITE", "MAGASIN"])
    mapping["STOCK"] = find_column(df, ["STOCK"])
    mapping["RAL"] = find_column(df, ["RAL"])

    return mapping

# ==============================
# INTRO UX
# ==============================
st.markdown("""
<div style="background:#f8fafc;border:1px solid #e2e8f0;
padding:20px;border-radius:10px;margin-bottom:20px">

<h4>📊 Module Suivi Implantation — Version Data Intelligent</h4>

<b>Objectif :</b><br>
Suivre en temps réel l’implantation des nouvelles références en magasin.

<hr>

<b>🧠 Intelligence intégrée :</b><br>
Ce module corrige automatiquement les erreurs fichiers :
<ul>
<li>Colonnes mal nommées</li>
<li>Accents / encodage Excel</li>
<li>Formats incohérents</li>
</ul>

<hr>

<b>📈 Indicateurs calculés :</b>
<ul>
<li>✅ Taux d’implantation réel</li>
<li>🚚 Retards logistiques (RAL)</li>
<li>🚨 Blocages supply</li>
<li>🏪 Performance par magasin</li>
</ul>

<hr>

<b>🎯 Lecture métier :</b>
<ul>
<li><b>Implanté :</b> stock > 0</li>
<li><b>Attente :</b> stock = 0 & RAL > 0</li>
<li><b>Alerte :</b> stock = 0 & RAL = 0</li>
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
    st.header("📁 Chargement")
    t1_file = st.file_uploader("Fichier T1", type=["csv", "xlsx"])
    stock_files = st.file_uploader("Stocks magasins", accept_multiple_files=True)

# ==============================
# LOAD T1
# ==============================
if not t1_file:
    st.info("Charge le fichier T1")
    st.stop()

t1 = read_file(t1_file)
t1 = clean_columns(t1)

map_t1 = auto_map_columns(t1)

if not map_t1["ARTICLE"]:
    st.error(f"❌ Impossible de détecter la colonne ARTICLE\n{list(t1.columns)}")
    st.stop()

t1 = t1.rename(columns={map_t1["ARTICLE"]: "ARTICLE"})
t1["SKU"] = t1["ARTICLE"].astype(str).str.zfill(8)

st.success("✅ T1 chargé et nettoyé automatiquement")

# ==============================
# LOAD STOCK
# ==============================
if not stock_files:
    st.info("Charge les fichiers stock")
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

    df["SKU"] = df["ARTICLE"].astype(str).str.zfill(8)
    df["STOCK"] = pd.to_numeric(df.get("STOCK", 0), errors="coerce").fillna(0)
    df["RAL"] = pd.to_numeric(df.get("RAL", 0), errors="coerce").fillna(0)

    dfs.append(df)

if not dfs:
    st.error("Aucun fichier stock exploitable")
    st.stop()

stock = pd.concat(dfs)

st.success("✅ Stocks chargés et nettoyés automatiquement")

# ==============================
# MERGE
# ==============================
df = stock.merge(t1[["SKU"]], on="SKU", how="inner")

# ==============================
# STATUT
# ==============================
df["Statut"] = np.select(
    [
        df["STOCK"] > 0,
        (df["STOCK"] == 0) & (df["RAL"] > 0)
    ],
    [
        "Implanté",
        "Attente"
    ],
    default="Alerte"
)

# ==============================
# KPI
# ==============================
total = len(df)
ok = (df["Statut"] == "Implanté").sum()
att = (df["Statut"] == "Attente").sum()
alert = (df["Statut"] == "Alerte").sum()

pct = int(ok / total * 100) if total > 0 else 0

c1, c2, c3 = st.columns(3)
c1.metric("✅ Implanté", ok)
c2.metric("🚚 Attente", att)
c3.metric("🚨 Alerte", alert)

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

st.subheader("🏪 Performance magasins")
st.dataframe(pivot, use_container_width=True)

# ==============================
# ALERTES
# ==============================
st.subheader("🚨 Articles bloqués")

df_alert = df[df["Statut"] == "Alerte"]

if df_alert.empty:
    st.success("Aucune alerte")
else:
    st.dataframe(df_alert, use_container_width=True)
