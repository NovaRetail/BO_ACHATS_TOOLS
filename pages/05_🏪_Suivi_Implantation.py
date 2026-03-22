import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
import io
from datetime import date

st.set_page_config(page_title="Suivi Implantation", layout="wide")

TODAY = date.today().strftime("%d %b %Y")

# ==============================
# HEADER
# ==============================
st.markdown(f"""
<div style="background:#0f1729;color:white;padding:14px 20px;
border-radius:8px;margin-bottom:20px;display:flex;justify-content:space-between">
<b>📦 Suivi Implantation — Data Quality Monitor</b>
<span style="color:#94a3b8">{TODAY}</span>
</div>
""", unsafe_allow_html=True)

# ==============================
# CLEAN COLUMNS
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

# ==============================
# REMOVE DUPLICATES
# ==============================
def remove_duplicate_columns(df):
    return df.loc[:, ~df.columns.duplicated()]

# ==============================
# FIND COLUMN
# ==============================
def find_column(df, keywords):
    for col in df.columns:
        for k in keywords:
            if k in col:
                return col
    return None

# ==============================
# MAPPING SAFE
# ==============================
def auto_map_columns(df):
    return {
        "ARTICLE": find_column(df, ["ARTICLE"]),
        "SITE": find_column(df, ["SITE", "MAGASIN"]),
        "STOCK": find_column(df, ["STOCK"]),
        "RAL": find_column(df, ["RAL"]),
    }

# ==============================
# SAFE SKU (ULTRA ROBUST)
# ==============================
def safe_sku(series):
    if isinstance(series, pd.DataFrame):
        series = series.iloc[:, 0]

    if hasattr(series, "ndim") and series.ndim > 1:
        series = series.iloc[:, 0]

    result = []
    for val in series:
        try:
            v = str(val)
            v = v.replace(".0", "")
            v = v.strip()
            v = v.zfill(8)
        except:
            v = "00000000"
        result.append(v)

    return pd.Series(result)

# ==============================
# DATA HEALTH CHECK
# ==============================
def data_health_check(df, name="Fichier"):
    score = 100
    issues = []

    # colonnes dupliquées
    dup_cols = df.columns[df.columns.duplicated()].tolist()
    if dup_cols:
        issues.append(f"Colonnes dupliquées : {dup_cols}")
        score -= 20

    # colonnes vides
    empty_cols = df.columns[df.isna().all()].tolist()
    if empty_cols:
        issues.append(f"Colonnes vides : {empty_cols}")
        score -= 10

    # taille dataset
    if df.shape[0] == 0:
        issues.append("Fichier vide")
        score -= 50

    return score, issues

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
    t1_file = st.file_uploader("Fichier T1")
    stock_files = st.file_uploader("Stocks", accept_multiple_files=True)

# ==============================
# LOAD T1
# ==============================
if not t1_file:
    st.stop()

t1 = read_file(t1_file)
t1 = clean_columns(t1)
t1 = remove_duplicate_columns(t1)

score, issues = data_health_check(t1, "T1")

st.subheader("🧠 Data Quality T1")
st.metric("Score qualité", f"{score}%")

if issues:
    for i in issues:
        st.warning(i)
else:
    st.success("Fichier propre")

map_t1 = auto_map_columns(t1)

if map_t1["ARTICLE"] is None:
    st.error("❌ Colonne ARTICLE non détectée")
    st.write(list(t1.columns))
    st.stop()

t1 = t1.rename(columns={map_t1["ARTICLE"]: "ARTICLE"})
t1["SKU"] = safe_sku(t1["ARTICLE"])

# ==============================
# LOAD STOCK
# ==============================
dfs = []

for f in stock_files:
    df = read_file(f)
    df = clean_columns(df)
    df = remove_duplicate_columns(df)

    score, issues = data_health_check(df, f.name)

    st.subheader(f"🧠 Data Quality {f.name}")
    st.metric("Score", f"{score}%")

    for i in issues:
        st.warning(i)

    mapping = auto_map_columns(df)

    if mapping["ARTICLE"] is None:
        st.error(f"{f.name} ignoré (pas de colonne ARTICLE)")
        continue

    df = df.rename(columns={
        mapping["ARTICLE"]: "ARTICLE",
        mapping["SITE"]: "SITE",
        mapping["STOCK"]: "STOCK",
        mapping["RAL"]: "RAL"
    })

    df["SKU"] = safe_sku(df["ARTICLE"])
    df["STOCK"] = pd.to_numeric(df.get("STOCK", 0), errors="coerce").fillna(0)
    df["RAL"] = pd.to_numeric(df.get("RAL", 0), errors="coerce").fillna(0)

    dfs.append(df)

if not dfs:
    st.error("❌ Aucun fichier exploitable")
    st.stop()

stock = pd.concat(dfs)

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

pct = int(ok / total * 100) if total > 0 else 0

c1, c2, c3 = st.columns(3)
c1.metric("Implanté", ok)
c2.metric("Attente", att)
c3.metric("Alerte", alert)

st.progress(pct / 100)

# ==============================
# TABLE MAGASIN
# ==============================
pivot = (
    df.groupby(["SITE", "Statut"])
    .size()
    .unstack(fill_value=0)
    .reset_index()
)

pivot["Total"] = pivot.sum(axis=1, numeric_only=True)
pivot["Taux (%)"] = (pivot.get("Implanté", 0) / pivot["Total"] * 100).round(0)

st.dataframe(pivot, use_container_width=True)

# ==============================
# EXPORT EXCEL
# ==============================
def build_excel(df, pivot):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Detail", index=False)
        pivot.to_excel(writer, sheet_name="Magasins", index=False)
        df[df["Statut"] == "Alerte"].to_excel(writer, sheet_name="Alertes", index=False)

    return output.getvalue()

excel_file = build_excel(df, pivot)

st.download_button(
    label="📥 Télécharger Excel",
    data=excel_file,
    file_name="rapport_implantation.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
