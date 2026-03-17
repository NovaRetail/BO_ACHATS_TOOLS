"""
03_📦_Detention_Top_CA.py — SmartBuyer Hub
Taux de détention Top CA · Flux IM/LO · Code état par article × magasin
Charte SmartBuyer v2
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
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── FONCTIONS UTILITAIRES ───────────────────────────────────────────────────
def fmt(n):
    if pd.isna(n): return "—"
    if abs(n) >= 1_000_000: return f"{n/1_000_000:.1f} M"
    if abs(n) >= 1_000: return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def norm_code(s):
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(8)

# ─── DATA ────────────────────────────────────────────────────────────────────
@st.cache_data
def load_stock(files):
    dfs = []
    for f in files:
        df = pd.read_csv(f, sep=";", dtype=str)
        dfs.append(df)
    df = pd.concat(dfs)

    for col in ["Nouveau stock","Ral","Nb colis","Prix d'achat"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Code article"] = norm_code(df["Code article"])
    df["Code etat"] = df["Code etat"].str.strip().str.upper()

    return df

@st.cache_data
def load_topca(f):
    df = pd.read_csv(f, header=None, names=["Code article"])
    df["Code article"] = norm_code(df["Code article"])
    return set(df["Code article"])

# ─── APP ─────────────────────────────────────────────────────────────────────
st.title("📦 Détention Top CA")

f_top = st.file_uploader("Top CA")
f_stock = st.file_uploader("Stocks", accept_multiple_files=True)

if not f_top or not f_stock:
    st.stop()

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
grp["Alerte"] = np.where(grp["stock"]<=0,"Rupture","OK")

taux = grp[grp["code_etat"]=="2"]
taux_moy = (taux["stock"]>0).mean()*100

# ─── VALORISATION ────────────────────────────────────────────────────────────
grp["val"] = grp["stock"] * grp["prix"]

val_total = grp["val"].sum()

df_b = grp[grp["code_etat"]=="B"]
val_b = df_b["val"].sum()
pct_b = (val_b/val_total*100) if val_total>0 else 0

df_b_top = df_b[df_b["Code article"].isin(top_codes)]
df_actifs = grp[grp["code_etat"]=="2"]

# ─── KPI ─────────────────────────────────────────────────────────────────────
k1,k2 = st.columns(2)

k1.metric("Taux détention", f"{taux_moy:.1f}%")

k2.metric("Stock bloqué",
          fmt(val_b),
          f"{pct_b:.1f}%")

# ─── ANALYSE ─────────────────────────────────────────────────────────────────
tab1,tab2 = st.tabs(["Analyse","Urgences"])

with tab1:

    st.subheader("🔴 Top bloqués")

    top_b = df_b.groupby("Code article").agg({
        "val":"sum",
        "stock":"sum"
    }).sort_values("val",ascending=False).head(10)

    st.dataframe(top_b)

    st.subheader("🟢 Top actifs")

    top_a = df_actifs.groupby("Code article").agg({
        "val":"sum",
        "stock":"sum"
    }).sort_values("val",ascending=False).head(10)

    st.dataframe(top_a)

with tab2:
    st.dataframe(grp[grp["Alerte"]!="OK"])
