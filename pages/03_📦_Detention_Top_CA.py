"""
03_📦_Detention_Top_CA.py — SmartBuyer Hub
Version FULL ROBUST (CSV + Excel + encoding auto)
"""

import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Détention Top CA", layout="wide")

# ─────────────────────────────────────────────────────────────
# UTILS
# ─────────────────────────────────────────────────────────────
def fmt(n):
    if pd.isna(n): return "—"
    if abs(n) >= 1_000_000: return f"{n/1_000_000:.1f} M"
    if abs(n) >= 1_000: return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def norm_code(s):
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(8)

def read_any(file):
    """Lecture intelligente : Excel + CSV + encoding + séparateur"""
    name = file.name.lower()

    # ── Excel direct
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(file, dtype=str)

    # ── CSV : test séparateurs + encoding
    for sep in [";", ",", "\t"]:
        for enc in ["utf-8", "cp1252", "latin1"]:
            try:
                file.seek(0)
                df = pd.read_csv(file, sep=sep, encoding=enc, dtype=str)
                if df.shape[1] > 0:
                    return df
            except:
                continue

    raise ValueError("❌ Fichier illisible : mauvais format ou corrompu")

# ─────────────────────────────────────────────────────────────
# LOADERS
# ─────────────────────────────────────────────────────────────
@st.cache_data
def load_topca(file):
    df = read_any(file)

    # prend première colonne uniquement
    df = df.iloc[:, [0]]
    df.columns = ["Code article"]

    df["Code article"] = norm_code(df["Code article"])

    return set(df["Code article"])

@st.cache_data
def load_stock(files):
    dfs = []
    for f in files:
        df = read_any(f)
        dfs.append(df)

    df = pd.concat(dfs, ignore_index=True)

    # sécurisation colonnes
    for col in ["Nouveau stock","Ral","Nb colis","Prix d'achat"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0

    df["Code article"] = norm_code(df["Code article"])
    df["Code etat"] = df.get("Code etat", "2").astype(str).str.strip().str.upper()

    return df

# ─────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────
st.title("📦 Détention Top CA")

f_top = st.file_uploader("📥 Liste Top CA (CSV ou Excel)")
f_stock = st.file_uploader("📥 Stocks ERP", accept_multiple_files=True)

if not f_top or not f_stock:
    st.info("⬆️ Charge les fichiers pour lancer l’analyse")
    st.stop()

# ─────────────────────────────────────────────────────────────
# DATA
# ─────────────────────────────────────────────────────────────
df_stock = load_stock(f_stock)
top_codes = load_topca(f_top)

grp = df_stock[df_stock["Code article"].isin(top_codes)].copy()

grp = grp.groupby(["Code article","Libellé site"]).agg({
    "Code etat":"first",
    "Nouveau stock":"sum",
    "Ral":"sum",
    "Nb colis":"first",
    "Prix d'achat":"first",
    "Libellé article":"first"
}).reset_index()

grp.columns = ["Code article","Magasin","code_etat","stock","ral","nb_colis","prix","lib_article"]

# ─────────────────────────────────────────────────────────────
# CALCULS
# ─────────────────────────────────────────────────────────────
grp["Alerte"] = np.where(grp["stock"]<=0,"🔴 Rupture","🟢 OK")

df_actifs = grp[grp["code_etat"]=="2"]
taux = (df_actifs["stock"]>0).mean()*100 if len(df_actifs)>0 else 0

grp["val"] = grp["stock"] * grp["prix"]

val_total = grp["val"].sum()

df_b = grp[grp["code_etat"]=="B"]
val_b = df_b["val"].sum()
pct_b = (val_b/val_total*100) if val_total>0 else 0

# ─────────────────────────────────────────────────────────────
# KPI
# ─────────────────────────────────────────────────────────────
st.markdown("### 📊 KPIs")

k1,k2,k3 = st.columns(3)

k1.metric("Taux détention", f"{taux:.1f}%")
k2.metric("€ Stock total", fmt(val_total))
k3.metric("€ Stock bloqué", fmt(val_b), f"{pct_b:.1f}%")

# ─────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📊 Analyse", "🚨 Urgences"])

with tab1:
    st.subheader("🔴 Top 10 bloqués")

    top_b = (
        df_b.groupby("Code article")
        .agg({"val":"sum","stock":"sum"})
        .sort_values("val",ascending=False)
        .head(10)
    )

    st.dataframe(top_b)

    st.subheader("🟢 Top 10 actifs")

    top_a = (
        df_actifs.groupby("Code article")
        .agg({"val":"sum","stock":"sum"})
        .sort_values("val",ascending=False)
        .head(10)
    )

    st.dataframe(top_a)

with tab2:
    st.subheader("Ruptures")
    st.dataframe(grp[grp["stock"]<=0])

    st.subheader("Bloqués")
    st.dataframe(df_b)

st.success("✅ Analyse terminée")
