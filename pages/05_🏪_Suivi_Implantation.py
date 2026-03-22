import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
import io

st.set_page_config(page_title="Suivi Implantation", layout="wide")

# ==============================
# NORMALISATION COLONNES
# ==============================
def normalize_columns(df):
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
# SAFE SKU (corrigé)
# ==============================
def safe_sku(series):
    if isinstance(series, pd.DataFrame):
        series = series.iloc[:, 0]

    return series.apply(lambda x: str(x).replace(".0", "").strip().zfill(8)[:8])

# ==============================
# LECTURE FICHIER
# ==============================
def read_file(file):
    if file.name.endswith(".xlsx"):
        return pd.read_excel(file)
    return pd.read_csv(file, sep=None, engine="python", encoding="latin1")

# ==============================
# LOAD T1
# ==============================
def load_t1(file):
    df = read_file(file)
    df = normalize_columns(df)

    # 🔥 fix colonnes dupliquées
    df = df.loc[:, ~df.columns.duplicated()]

    if "CODE ARTICLE" not in df.columns:
        st.error("❌ Colonne CODE ARTICLE manquante dans T1")
        st.stop()

    df["SKU"] = safe_sku(df["CODE ARTICLE"])
    return df

# ==============================
# LOAD STOCK
# ==============================
def load_stock(files):
    dfs = []

    for f in files:
        df = read_file(f)
        df = normalize_columns(df)

        # 🔥 fix colonnes dupliquées
        df = df.loc[:, ~df.columns.duplicated()]

        if "CODE ARTICLE" not in df.columns:
            st.warning(f"{f.name} ignoré (pas de CODE ARTICLE)")
            continue

        df["SKU"] = safe_sku(df["CODE ARTICLE"])

        df["STOCK"] = pd.to_numeric(df.get("STOCK", 0), errors="coerce").fillna(0)
        df["RAL"] = pd.to_numeric(df.get("RAL", 0), errors="coerce").fillna(0)

        dfs.append(df)

    if not dfs:
        st.error("❌ Aucun fichier exploitable")
        st.stop()

    return pd.concat(dfs)

# ==============================
# DATA QUALITY (discret)
# ==============================
def compute_dq(df):
    score = 100
    issues = []

    dup = df.columns[df.columns.duplicated()]
    if len(dup) > 0:
        score -= 10
        issues.append("colonnes dupliquées")

    null_rate = df.isna().mean().mean()
    if null_rate > 0.15:
        score -= 15
        issues.append("valeurs manquantes")

    if len(df) == 0:
        score -= 50
        issues.append("dataset vide")

    return max(score, 0), issues

# ==============================
# SIDEBAR
# ==============================
with st.sidebar:
    st.header("📁 Chargement")
    t1_file = st.file_uploader("Fichier T1")
    stock_files = st.file_uploader("Stocks", accept_multiple_files=True)

if not t1_file:
    st.stop()

# ==============================
# LOAD DATA
# ==============================
t1 = load_t1(t1_file)
df_stock = load_stock(stock_files)

# ==============================
# DATA QUALITY CARD
# ==============================
dq_score, dq_issues = compute_dq(df_stock)

color = "#059669" if dq_score >= 85 else "#0284c7" if dq_score >= 60 else "#dc2626"
bg    = "#ecfdf5" if dq_score >= 85 else "#f0f9ff" if dq_score >= 60 else "#fef2f2"

st.markdown(f"""
<div style="
background:{bg};
border:1px solid #e2e8f0;
border-left:4px solid {color};
border-radius:8px;
padding:10px 14px;
margin-bottom:12px;
max-width:320px;
">
    <div style="font-size:10px;color:#64748b;font-weight:700;">
        DATA QUALITY
    </div>
    <div style="font-size:22px;font-weight:800;color:{color};">
        {dq_score}%
    </div>
    <div style="font-size:10px;color:#64748b;">
        {" · ".join(dq_issues) if dq_issues else "OK"}
    </div>
</div>
""", unsafe_allow_html=True)

# ==============================
# MERGE
# ==============================
df = df_stock.merge(t1[["SKU"]], on="SKU", how="inner")

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

c1, c2, c3 = st.columns(3)
c1.metric("Implanté", ok)
c2.metric("Attente", att)
c3.metric("Alerte", alert)

st.progress(ok / total if total > 0 else 0)

# ==============================
# TABLE MAGASINS
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
# ALERTES
# ==============================
st.subheader("🚨 Alertes")

df_alert = df[df["Statut"] == "Alerte"]

if df_alert.empty:
    st.success("Aucune alerte")
else:
    st.dataframe(df_alert, use_container_width=True)

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
