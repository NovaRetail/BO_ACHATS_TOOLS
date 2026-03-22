import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import date

# ==============================
# CONFIG
# ==============================
st.set_page_config(
    page_title="Suivi Implantation",
    layout="wide"
)

TODAY = date.today().strftime("%d %b %Y")

# ==============================
# HEADER
# ==============================
st.markdown(f"""
<div style="background:#0f1729;color:white;padding:14px 20px;
border-radius:8px;margin-bottom:20px;display:flex;justify-content:space-between">
<b>📦 Suivi Implantation</b>
<span style="color:#94a3b8">{TODAY}</span>
</div>
""", unsafe_allow_html=True)

# ==============================
# INTRO (AJOUT IMPORTANT)
# ==============================
st.markdown("""
<div style="background:#f8fafc;border:1px solid #e2e8f0;
padding:18px;border-radius:10px;margin-bottom:20px">

<b>ℹ️ À quoi sert ce module ?</b><br><br>

Ce module analyse l’implantation des nouvelles références en magasin.

Il croise :
• le fichier <b>T1</b> (nouvelles références)  
• les <b>stocks magasins</b>  

Objectifs :

1️⃣ Mesurer le taux d’implantation  
2️⃣ Identifier les blocages  
3️⃣ Suivre les retards logistiques  
4️⃣ Prioriser les actions terrain  

</div>
""", unsafe_allow_html=True)

# ==============================
# HELPERS
# ==============================
def read_file(file):
    if file.name.endswith(".xlsx"):
        return pd.read_excel(file)
    return pd.read_csv(file, sep=None, engine="python", encoding="latin1")

def normalize(df):
    df.columns = df.columns.str.upper().str.strip()
    return df

# ==============================
# SIDEBAR
# ==============================
with st.sidebar:
    st.title("📁 Chargement")
    t1_file = st.file_uploader("T1", type=["csv", "xlsx"])
    stock_files = st.file_uploader("Stocks", accept_multiple_files=True)

# ==============================
# LOAD T1
# ==============================
if not t1_file:
    st.info("Charge le fichier T1")
    st.stop()

t1 = normalize(read_file(t1_file))

if "ARTICLE" not in t1.columns:
    st.error("Colonne ARTICLE manquante")
    st.stop()

t1["SKU"] = t1["ARTICLE"].astype(str).str.zfill(8)

# ==============================
# LOAD STOCK
# ==============================
if not stock_files:
    st.info("Charge les fichiers stock")
    st.stop()

dfs = []
for f in stock_files:
    df = normalize(read_file(f))
    if "CODE ARTICLE" not in df.columns:
        continue

    df["SKU"] = df["CODE ARTICLE"].astype(str).str.zfill(8)
    df["NOUVEAU STOCK"] = pd.to_numeric(df.get("NOUVEAU STOCK", 0), errors="coerce").fillna(0)
    df["RAL"] = pd.to_numeric(df.get("RAL", 0), errors="coerce").fillna(0)

    dfs.append(df)

if not dfs:
    st.error("Aucun fichier stock valide")
    st.stop()

stock = pd.concat(dfs)

# ==============================
# MERGE
# ==============================
df = stock.merge(t1[["SKU"]], on="SKU", how="inner")

# ==============================
# STATUT
# ==============================
df["Statut"] = np.select(
    [
        df["NOUVEAU STOCK"] > 0,
        (df["NOUVEAU STOCK"] == 0) & (df["RAL"] > 0)
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
    df.groupby(["LIBELLÉ SITE", "Statut"])
    .size()
    .unstack(fill_value=0)
    .reset_index()
)

pivot["Total"] = pivot.sum(axis=1, numeric_only=True)
pivot["Taux (%)"] = (pivot.get("Implanté", 0) / pivot["Total"] * 100).round(0).astype(int)

st.subheader("📊 Par magasin")

pivot_display = pivot.copy()
pivot_display["Taux (%)"] = pivot_display["Taux (%)"].astype(str) + "%"

st.dataframe(pivot_display, use_container_width=True)

# ==============================
# ALERTES
# ==============================
st.subheader("🚨 Alertes")

df_alert = df[df["Statut"] == "Alerte"]

if df_alert.empty:
    st.success("Aucune alerte")
else:
    st.dataframe(df_alert, use_container_width=True)
