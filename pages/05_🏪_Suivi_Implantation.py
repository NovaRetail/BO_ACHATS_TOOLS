import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(layout="wide", page_title="Dashboard COPIL")

# =========================
# LOADERS
# =========================
def load_implant(file):
    df = pd.read_csv(file, sep=";", encoding="latin1")
    df.columns = df.columns.str.strip().str.upper()

    df["SKU"] = (
        df["ARTICLE"].astype(str)
        .str.replace(".0", "")
        .str.zfill(8)
    )

    df["ORIGINE"] = np.where(
        df["MODE APPRO"].str.contains("IMPORT", na=False),
        "IM", "LO"
    )

    return df

def load_stock(files):
    dfs = []
    for f in files:
        df = pd.read_csv(f, sep=";", encoding="latin1")
        df.columns = df.columns.str.strip()

        df["SKU"] = (
            df["Code article"].astype(str)
            .str.replace(".0", "")
            .str.zfill(8)
        )

        df["Stock"] = pd.to_numeric(df["Nouveau stock"], errors="coerce").fillna(0)
        df["RAL"] = pd.to_numeric(df["Ral"], errors="coerce").fillna(0)

        dfs.append(df)

    return pd.concat(dfs, ignore_index=True)

# =========================
# UI SIDEBAR
# =========================
st.sidebar.title("Chargement")

implant_file = st.sidebar.file_uploader("IMPLANT", type="csv")
stock_files = st.sidebar.file_uploader("STOCK", type="csv", accept_multiple_files=True)

if not implant_file or not stock_files:
    st.stop()

implant = load_implant(implant_file)
stock = load_stock(stock_files)

# =========================
# DATA PREP
# =========================
skus = implant["SKU"].unique()

base = pd.MultiIndex.from_product(
    [stock["Libellé site"].unique(), skus],
    names=["Magasin", "SKU"]
).to_frame(index=False)

stock = stock.rename(columns={
    "Libellé site": "Magasin"
})

df = base.merge(stock[["Magasin","SKU","Stock","RAL"]], on=["Magasin","SKU"], how="left")
df["Stock"] = df["Stock"].fillna(0)
df["RAL"] = df["RAL"].fillna(0)

df["Statut"] = np.select(
    [
        df["Stock"] > 0,
        (df["Stock"] == 0) & (df["RAL"] > 0)
    ],
    [
        "OK",
        "ATTENTE"
    ],
    default="ALERTE"
)

# =========================
# KPI
# =========================
pivot = df.groupby(["Magasin","Statut"]).size().unstack(fill_value=0)

for col in ["OK","ATTENTE","ALERTE"]:
    if col not in pivot.columns:
        pivot[col] = 0

pivot["Total"] = pivot.sum(axis=1)
pivot["Taux (%)"] = (pivot["OK"] / pivot["Total"] * 100).round(0).astype(int)

# =========================
# HEADER
# =========================
st.markdown("## Dashboard Implantation COPIL")

# =========================
# SCORECARDS (FIXED)
# =========================
st.markdown("### Scorecard magasins")

cards = '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px;">'

for _, row in pivot.sort_values("Taux (%)").iterrows():

    taux = int(row["Taux (%)"])

    if taux >= 80:
        color = "#059669"
        bg = "#ecfdf5"
    elif taux >= 65:
        color = "#d97706"
        bg = "#fffbeb"
    else:
        color = "#dc2626"
        bg = "#fef2f2"

    cards += f"""
    <div style="
        background:{bg};
        padding:14px;
        border-radius:10px;
        border:1px solid #e2e8f0;
    ">
        <div style="font-size:12px;font-weight:600;">
            {row.name}
        </div>

        <div style="
            font-size:28px;
            font-weight:800;
            color:{color};
        ">
            {taux}%
        </div>

        <div style="font-size:11px;color:#64748b;">
            {row['OK']} OK · {row['ALERTE']} alertes
        </div>
    </div>
    """

cards += "</div>"

st.markdown(cards, unsafe_allow_html=True)

# =========================
# TOP ALERTES
# =========================
st.markdown("### Top magasins en alerte")

top = pivot.sort_values("ALERTE", ascending=False).reset_index()

st.dataframe(top.head(10), use_container_width=True)

# =========================
# ARTICLES DESTRUCTEURS
# =========================
st.markdown("### Articles destructeurs")

destructeurs = (
    df[df["Statut"]=="ALERTE"]
    .groupby("SKU")
    .agg(Alertes=("Magasin","count"))
    .sort_values("Alertes", ascending=False)
)

st.dataframe(destructeurs.head(20), use_container_width=True)

# =========================
# EXPORT
# =========================
output = io.BytesIO()
with pd.ExcelWriter(output) as writer:
    df.to_excel(writer, sheet_name="Detail")
    pivot.to_excel(writer, sheet_name="Magasins")

st.download_button(
    "Télécharger Excel",
    data=output.getvalue(),
    file_name="COPIL.xlsx"
)
