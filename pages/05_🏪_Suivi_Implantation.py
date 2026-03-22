import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(layout="wide", page_title="Dashboard COPIL")

# =========================
# STYLE SIMPLE & PRO
# =========================
st.markdown("""
<style>
.box {
    border-radius:12px;
    padding:18px;
    margin-bottom:12px;
    border:1px solid #e5e7eb;
    background:#f9fafb;
}
.title {
    font-weight:700;
    font-size:16px;
}
.small {
    font-size:13px;
    color:#6b7280;
}
</style>
""", unsafe_allow_html=True)

# =========================
# MENU
# =========================
page = st.sidebar.radio("Navigation", ["Accueil", "Analyse COPIL"])

# =========================
# PAGE ACCUEIL
# =========================
if page == "Accueil":

    st.title("📊 Module Suivi Implantation")

    st.markdown("""
    <div class="box">
    <div class="title">À quoi sert ce module ?</div>
    Ce module permet de suivre la performance d’implantation des articles dans le réseau magasins.
    <br><br>
    Il permet d’identifier rapidement :
    <ul>
        <li>Les magasins en difficulté</li>
        <li>Les articles non implantés</li>
        <li>Les blocages supply (RAL)</li>
        <li>Les priorités opérationnelles</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 🎯 Indicateurs clés")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="box">
        <b>📦 Taux d’implantation</b><br>
        % des articles disponibles en magasin
        <br><br>
        <span class="small">OK / Total articles</span>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="box">
        <b>🚨 Alertes</b><br>
        Articles sans stock ni commande
        <br><br>
        <span class="small">Risque rupture ou oubli implantation</span>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="box">
        <b>⏳ Attente livraison</b><br>
        Articles commandés mais non reçus
        <br><br>
        <span class="small">Impact supply chain</span>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="box">
        <b>🏪 Performance magasins</b><br>
        Classement des magasins selon leur exécution
        <br><br>
        <span class="small">Permet de prioriser les actions terrain</span>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("### 📂 Fichiers attendus")

    st.markdown("""
    <div class="box">
    <b>1. Fichier IMPLANT</b><br>
    Liste des articles à implanter<br><br>
    
    <b>2. Fichier STOCK</b><br>
    Export ERP avec stock et RAL
    </div>
    """, unsafe_allow_html=True)

    st.info("➡️ Charge les fichiers dans la sidebar puis va dans 'Analyse COPIL'")

# =========================
# PAGE ANALYSE
# =========================
if page == "Analyse COPIL":

    st.title("📊 Dashboard Implantation COPIL")

    # ========= LOAD =========
    implant_file = st.sidebar.file_uploader("IMPLANT", type="csv")
    stock_files = st.sidebar.file_uploader("STOCK", type="csv", accept_multiple_files=True)

    if not implant_file or not stock_files:
        st.warning("Charge les fichiers pour lancer l’analyse")
        st.stop()

    # ========= LOAD FUNCTIONS =========
    def clean(df):
        df.columns = df.columns.str.replace("\ufeff", "").str.strip()
        return df

    def find(df, names):
        for n in names:
            for c in df.columns:
                if n.lower() in c.lower():
                    return c
        return None

    # ========= LOAD IMPLANT =========
    implant = pd.read_csv(implant_file, sep=";", encoding="latin1")
    implant = clean(implant)

    col_article = find(implant, ["article"])
    implant["SKU"] = implant[col_article].astype(str).str.zfill(8)

    # ========= LOAD STOCK =========
    dfs = []
    for f in stock_files:
        df = pd.read_csv(f, sep=";", encoding="latin1")
        df = clean(df)

        col_article = find(df, ["article"])
        col_stock = find(df, ["stock"])
        col_ral = find(df, ["ral"])
        col_mag = find(df, ["site"])

        df["SKU"] = df[col_article].astype(str).str.zfill(8)
        df["Stock"] = pd.to_numeric(df[col_stock], errors="coerce").fillna(0)
        df["RAL"] = pd.to_numeric(df[col_ral], errors="coerce").fillna(0)
        df["Magasin"] = df[col_mag]

        dfs.append(df)

    stock = pd.concat(dfs)

    # ========= DATA =========
    base = pd.MultiIndex.from_product(
        [stock["Magasin"].unique(), implant["SKU"].unique()],
        names=["Magasin","SKU"]
    ).to_frame(index=False)

    df = base.merge(stock[["Magasin","SKU","Stock","RAL"]], how="left")

    df["Stock"] = df["Stock"].fillna(0)
    df["RAL"] = df["RAL"].fillna(0)

    df["Statut"] = np.select(
        [
            df["Stock"] > 0,
            (df["Stock"] == 0) & (df["RAL"] > 0)
        ],
        ["OK","ATTENTE"],
        default="ALERTE"
    )

    pivot = df.groupby(["Magasin","Statut"]).size().unstack(fill_value=0)

    for c in ["OK","ATTENTE","ALERTE"]:
        if c not in pivot.columns:
            pivot[c] = 0

    pivot["Total"] = pivot.sum(axis=1)
    pivot["Taux (%)"] = (pivot["OK"]/pivot["Total"]*100).round(0)

    # =========================
    # SCORECARDS FIX
    # =========================
    st.subheader("Scorecard magasins")

    cards = '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px;">'

    for magasin, row in pivot.sort_values("Taux (%)").iterrows():

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
        <div style="background:{bg};padding:14px;border-radius:10px;">
            <b>{magasin}</b>
            <div style="font-size:28px;color:{color};font-weight:800;">
                {taux}%
            </div>
            <div style="font-size:12px;">
                {row['OK']} OK · {row['ALERTE']} alertes
            </div>
        </div>
        """

    cards += "</div>"

    st.markdown(cards, unsafe_allow_html=True)

    # =========================
    # TOP ALERTES
    # =========================
    st.subheader("Top magasins en alerte")
    st.dataframe(pivot.sort_values("ALERTE", ascending=False).head(10))

    # =========================
    # DESTRUCTEURS
    # =========================
    st.subheader("Articles destructeurs")

    destructeurs = (
        df[df["Statut"]=="ALERTE"]
        .groupby("SKU")
        .size()
        .sort_values(ascending=False)
    )

    st.dataframe(destructeurs.head(20))

    # =========================
    # EXPORT
    # =========================
    output = io.BytesIO()
    with pd.ExcelWriter(output) as writer:
        df.to_excel(writer, sheet_name="Detail")
        pivot.to_excel(writer, sheet_name="Magasins")

    st.download_button(
        "📥 Télécharger Excel",
        data=output.getvalue(),
        file_name="copil.xlsx"
    )
