import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
from datetime import date
import io

# Configuration de la page
st.set_page_config(page_title="Data Quality Safe - Implantation", layout="wide", page_icon="📦")

# --- FONCTIONS UTILES ---
def clean_column_name(col):
    """Nettoie les noms de colonnes : majuscules, sans accents, sans espaces inutiles."""
    col = str(col).strip().upper()
    col = unicodedata.normalize('NFKD', col).encode('ascii', 'ignore').decode('utf-8')
    return col

@st.cache_data
def load_data(file):
    """Lecture intelligente des fichiers CSV, XLSX ou XLSB."""
    try:
        if file.name.endswith('.csv'):
            # Détection automatique du séparateur (souvent ';' dans tes fichiers)
            return pd.read_csv(file, sep=None, engine='python', encoding='latin1')
        elif file.name.endswith('.xlsb'):
            return pd.read_excel(file, engine='pyxlsb')
        else:
            return pd.read_excel(file)
    except Exception as e:
        st.error(f"Erreur lors de la lecture de {file.name}: {e}")
        return None

def format_sku(series):
    """Formatage rigoureux des codes articles sur 8 chiffres."""
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.zfill(8)

# --- INTERFACE ---
st.markdown(f"""
<div style="background:#0f1729;color:white;padding:15px;border-radius:10px;display:flex;justify-content:space-between;align-items:center;">
    <h2 style="margin:0;">📦 Suivi Implantation</h2>
    <span style="opacity:0.8;">Mise à jour : {date.today().strftime("%d/%m/%Y")}</span>
</div>
""", unsafe_allow_html=True)

# Barre latérale pour l'import
with st.sidebar:
    st.header("📂 Chargement des données")
    t1_input = st.file_uploader("Fichier Référentiel (T1 / IMPLANT)", type=["csv", "xlsx", "xlsb"])
    stock_inputs = st.file_uploader("Fichiers Extraction Stock", type=["csv", "xlsx", "xlsb"], accept_multiple_files=True)
    
    st.divider()
    st.info("💡 Les colonnes sont détectées automatiquement même si les noms varient légèrement.")

if not t1_input or not stock_inputs:
    st.warning("Veuillez charger le fichier référentiel et au moins un fichier de stock pour commencer.")
    st.stop()

# --- TRAITEMENT DU RÉFÉRENTIEL (T1) ---
with st.spinner("Analyse du référentiel..."):
    df_t1_raw = load_data(t1_input)
    if df_t1_raw is not None:
        df_t1_raw.columns = [clean_column_name(c) for c in df_t1_raw.columns]
        # Recherche de la colonne article
        col_art_t1 = next((c for c in df_t1_raw.columns if "ARTICLE" in c), None)
        
        if col_art_t1:
            df_t1_raw['SKU'] = format_sku(df_t1_raw[col_art_t1])
            list_sku_ref = df_t1_raw[['SKU']].drop_duplicates()
        else:
            st.error("Colonne 'ARTICLE' introuvable dans le fichier référentiel.")
            st.stop()

# --- TRAITEMENT DES STOCKS ---
all_stocks = []
for f in stock_inputs:
    df_s = load_data(f)
    if df_s is not None:
        df_s.columns = [clean_column_name(c) for c in df_s.columns]
        
        # Mapping dynamique basé sur tes fichiers
        mapping = {
            'SKU': next((c for c in df_s.columns if "ARTICLE" in c), None),
            'SITE': next((c for c in df_s.columns if any(k in c for k in ["SITE", "MAGASIN", "LIBELLE SITE"])), "SITE_INCONNU"),
            'STOCK': next((c for c in df_s.columns if "STOCK" in c), None),
            'RAL': next((c for c in df_s.columns if "RAL" in c), None)
        }
        
        if mapping['SKU']:
            df_s['SKU'] = format_sku(df_s[mapping['SKU']])
            # Conversion numérique sécurisée
            for col in ['STOCK', 'RAL']:
                if mapping[col] in df_s.columns:
                    df_s[col] = pd.to_numeric(df_s[mapping[col]], errors='coerce').fillna(0)
                else:
                    df_s[col] = 0
            
            # On ne garde que les colonnes utiles pour la fusion
            cols_to_keep = ['SKU', mapping['SITE'], 'STOCK', 'RAL']
            all_stocks.append(df_s[[c for c in cols_to_keep if c in df_s.columns]].rename(columns={mapping['SITE']: 'SITE'}))

if not all_stocks:
    st.error("Aucune donnée de stock valide trouvée.")
    st.stop()

df_final = pd.concat(all_stocks)
# On filtre uniquement sur les articles présents dans le référentiel T1
df_final = df_final.merge(list_sku_ref, on="SKU", how="inner")

# --- CALCUL DES STATUTS ---
conditions = [
    (df_final['STOCK'] > 0),
    (df_final['STOCK'] <= 0) & (df_final['RAL'] > 0)
]
choices = ["Implanté", "Attente"]
df_final['STATUT'] = np.select(conditions, choices, default="Alerte")

# --- DASHBOARD ---
tab1, tab2 = st.tabs(["📊 Synthèse", "🔍 Détail des Alertes"])

with tab1:
    # Métriques globales
    c1, c2, c3, c4 = st.columns(4)
    total = len(df_final)
    imp = (df_final['STATUT'] == "Implanté").sum()
    att = (df_final['STATUT'] == "Attente").sum()
    ale = (df_final['STATUT'] == "Alerte").sum()
    taux = (imp / total * 100) if total > 0 else 0
    
    c1.metric("Total SKU", total)
    c2.metric("Implanté ✅", imp, f"{taux:.1f}%")
    c3.metric("Attente ⏳", att)
    c4.metric("Alerte 🚨", ale)
    
    st.progress(taux / 100)
    
    # Vue par magasin
    st.subheader("Performance par Site")
    pivot = df_final.groupby(['SITE', 'STATUT']).size().unstack(fill_value=0).reset_index()
    if "Implanté" in pivot.columns:
        pivot['Taux %'] = (pivot['Implanté'] / pivot.sum(axis=1, numeric_only=True) * 100).round(1)
    st.dataframe(pivot.sort_values(by="Taux %", ascending=True), use_container_width=True)

with tab2:
    st.subheader("Articles en Alerte (Stock 0 / RAL 0)")
    df_alertes = df_final[df_final['STATUT'] == "Alerte"]
    
    if not df_alertes.empty:
        # --- BOUTON DE TÉLÉCHARGEMENT ---
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_alertes.to_excel(writer, index=False, sheet_name='Alertes')
        
        st.download_button(
            label="📥 Télécharger les Alertes en Excel",
            data=buffer.getvalue(),
            file_name=f"Alertes_Implantation_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.dataframe(df_alertes, use_container_width=True)
    else:
        st.success("Félicitations ! Aucune alerte détectée.")
