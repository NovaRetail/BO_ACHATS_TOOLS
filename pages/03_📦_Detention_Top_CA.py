"""
03_📦_Detention_Top_CA.py — SmartBuyer Hub
Version robuste avec gestion encoding + KPI stock bloqué
"""

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
    layout="wide"
)

# ─── UTILS ───────────────────────────────────────────────────────────────────
def fmt(n):
    if pd.isna(n): return "—"
    if abs(n) >= 1_000_000: return f"{n/1_000_000:.1f} M"
    if abs(n) >= 1_000: return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def norm_code(s):
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(8)

def read_file_safe(file, sep=","):
    """Lecture robuste CSV avec fallback encoding"""
    for enc in ["utf-8", "cp1252", "latin1"]:
        try:
            file.seek(0)
            return pd.read_csv(file, sep=sep, dtype=str, encoding=enc)
        except:
            continue
    raise ValueError("Impossible de lire le fichier (encoding)")

# ─── LOADERS ─────────────────────────────────────────────────────────────────
@st.cache_data
def load_topca(file):
    df = read_file_safe(file, sep=",")
    df.columns = ["Code article"]
    df["Code article"] = norm_code(df["Code article"])
    return set(df["Code article"])

@st.cache_data
def load_stock(files):
    dfs = []
    for f in files:
        df = read_file_safe(f, sep=";")
        dfs.append(df)

    df = pd.concat(dfs, ignore_index=True)

    # sécurisation colonnes
    for col in ["Nouveau stock","Ral","Nb colis","Prix d'achat"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Code article"] = norm_code(df["Code article"])
    df["Code etat"] = df["Code etat"].astype(str).str.strip().str.upper()

    return df

# ─── APP ─────────────────────────────────────────────────────────────────────
st.title("📦 Détention Top CA")

f_top = st.file_uploader("Liste Top CA")
f_stock = st.file_uploader("Fichiers stock", accept_multiple_files=True)

if not f_top or not f_stock:
    st.info("Charge les fichiers pour démarrer")
    st.stop()

# ─── DATA ────────────────────────────────────────────────────────────────────
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

# ─── CALCULS ─────────────────────────────────────────────────────────────────
grp["Alerte"] = np.where(grp["stock"]<=0,"🔴 Rupture","🟢 OK")

df_actifs = grp[grp["code_etat"]=="2"]
taux_moy = (df_actifs["stock"]>0).mean()*100 if len(df_actifs)>0 else 0

# ─── VALORISATION ────────────────────────────────────────────────────────────
grp["val"] = grp["stock"] * grp["prix"]

val_total = grp["val"].sum()

df_b = grp[grp["code_etat"]=="B"]
val_b = df_b["val"].sum()
pct_b = (val_b/val_total*100) if val_total>0 else 0

df_b_top = df_b[df_b["Code article"].isin(top_codes)]

# ─── KPI ─────────────────────────────────────────────────────────────────────
st.markdown("### 📊 Indicateurs")

k1,k2,k3 = st.columns(3)

k1.metric("Taux détention",
          f"{taux_moy:.1f}%")

k2.metric("€ Stock total",
          fmt(val_total))

# couleur intelligente
if pct_b > 15:
    delta = "🔴 Critique"
elif pct_b > 5:
    delta = "🟡 À surveiller"
else:
    delta = "🟢 OK"

k3.metric("€ Stock bloqué",
          fmt(val_b),
          f"{pct_b:.1f}%")

# ─── TABS ────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📊 Analyse", "🚨 Urgences"])

# ═══════════════════════════════════════════════════════════════
with tab1:

    st.subheader("🔴 Top 10 bloqués — Top CA")

    top_b_top = (
        df_b_top.groupby(["Code article","lib_article"])
        .agg({"val":"sum","stock":"sum"})
        .reset_index()
        .sort_values("val",ascending=False)
        .head(10)
    )
    top_b_top["val"] = top_b_top["val"].apply(fmt)

    st.dataframe(top_b_top, use_container_width=True)

    st.subheader("🔴 Top 10 bloqués — Global")

    top_b = (
        df_b.groupby(["Code article","lib_article"])
        .agg({"val":"sum","stock":"sum"})
        .reset_index()
        .sort_values("val",ascending=False)
        .head(10)
    )
    top_b["val"] = top_b["val"].apply(fmt)

    st.dataframe(top_b, use_container_width=True)

    st.subheader("🟢 Top 10 actifs")

    top_a = (
        df_actifs.groupby(["Code article","lib_article"])
        .agg({"val":"sum","stock":"sum"})
        .reset_index()
        .sort_values("val",ascending=False)
        .head(10)
    )
    top_a["val"] = top_a["val"].apply(fmt)

    st.dataframe(top_a, use_container_width=True)

# ═══════════════════════════════════════════════════════════════
with tab2:

    st.subheader("Articles en rupture")

    rupture = grp[grp["stock"]<=0]

    st.dataframe(rupture, use_container_width=True)

    st.subheader("Articles bloqués")

    st.dataframe(df_b, use_container_width=True)

# ─── FIN ──────────────────────────────────────────────────────
st.success("✅ Analyse prête")
