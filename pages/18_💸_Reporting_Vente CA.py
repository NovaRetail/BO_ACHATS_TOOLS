# ==============================================================================
# 0. CONFIGURATION STREAMLIT & PALETTE APPLE / iOS LIGHT
# ==============================================================================
import os
import re
import sys
import io
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Flag global pour le style des cartes KPI (3D double ombre vs Flat)
KPI_STYLE_3D = True

# Configuration de la page Streamlit
st.set_page_config(
    page_title="Reporting Vente CA",
    page_icon="💸",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Injection CSS Style iOS Light
CUSTOM_CSS = """
<style>
    /* Global Background */
    .stApp {
        background-color: #F2F2F7;
        font-family: -apple-system, "SF Pro Display", "Segoe UI", Roboto, sans-serif;
    }
    
    /* Custom Card Style */
    .ios-card {
        background-color: #FFFFFF;
        border-radius: 14px;
        padding: 16px 20px;
        margin-bottom: 16px;
        border: 1px solid rgba(0,0,0,0.04);
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
    }
    
    .ios-card-3d {
        background-color: #FFFFFF;
        border-radius: 14px;
        padding: 16px 20px;
        margin-bottom: 16px;
        border-top: 4px solid #007AFF;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04), 0 8px 16px rgba(0,122,255,0.06);
    }
    
    .ios-card-3d-alert {
        background-color: #FFFFFF;
        border-radius: 14px;
        padding: 16px 20px;
        margin-bottom: 16px;
        border-top: 4px solid #FF3B30;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04), 0 8px 16px rgba(255,59,48,0.08);
    }
    
    .ios-card-3d-success {
        background-color: #FFFFFF;
        border-radius: 14px;
        padding: 16px 20px;
        margin-bottom: 16px;
        border-top: 4px solid #34C759;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04), 0 8px 16px rgba(52,199,89,0.08);
    }

    /* Metric Labels & Values */
    .kpi-title {
        font-size: 0.82rem;
        font-weight: 600;
        color: #8E8E93;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #1C1C1E;
    }
    .kpi-sub {
        font-size: 0.85rem;
        font-weight: 500;
        margin-top: 4px;
    }

    /* Badges de Sévérité */
    .badge-critique {
        background-color: #FFE5E5;
        color: #D32F2F;
        padding: 4px 10px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.8rem;
        display: inline-block;
    }
    .badge-majeur {
        background-color: #FFF3E0;
        color: #E65100;
        padding: 4px 10px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.8rem;
        display: inline-block;
    }
    .badge-modere {
        background-color: #FFFDE7;
        color: #F57F17;
        padding: 4px 10px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.8rem;
        display: inline-block;
    }
    .badge-ok {
        background-color: #E8F5E9;
        color: #2E7D32;
        padding: 4px 10px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.8rem;
        display: inline-block;
    }

    /* Avatar Rayon */
    .rayon-avatar {
        width: 48px;
        height: 48px;
        border-radius: 50%;
        background-color: #007AFF;
        color: white;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 1.1rem;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ==============================================================================
# 1. I/O ROBUSTE ET PARSING DES DONNÉES
# ==============================================================================
REQUIRED_COLUMNS = [
    'Société', 'Rayon', 'Site', 'CA N-1', 'Budget', 'CA', 'Poids', 'Vs N-1 (%)',
    'Vs Bgt (%)', 'Marge N-1', 'Marge', 'Taux de Marge N-1', 'Taux de Marge',
    'Taux de Marge N Vs N-1', 'Débit N-1', 'Débit', 'Vs N-1 (%).1',
    'Panier N-1', 'Panier', 'Panier N Vs N-1', 'Panier Qté N-1', 'Panier Qté',
    'Panier Qté N Vs N-1', 'Volume N-1', 'Volume', 'Volume N Vs N-1'
]

def split_code_libelle(val):
    """Sépare une chaîne 'CODE - Libellé' en tuple (Code, Libellé)."""
    if pd.isna(val) or val is None:
        return np.nan, np.nan
    val_str = str(val).strip()
    if ' - ' in val_str:
        parts = val_str.split(' - ', 1)
        return parts[0].strip(), parts[1].strip()
    return np.nan, val_str

def extract_format(site_libelle):
    """Extrait le format magasin (Hyper, Market, Supeco) depuis le libellé."""
    if pd.isna(site_libelle):
        return "Inconnu"
    lib = str(site_libelle).upper()
    if "SUPER" in lib or "SUPECO" in lib:
        if "SUPECO" in lib:
            return "Supeco"
        return "Market"
    elif "MARKET" in lib:
        return "Market"
    elif "HYPER" in lib:
        return "Hyper"
    return "Autre"

def _load_data_impl(file_input):
    """Implémentation pure du chargement de données (sans dépendance st.cache)."""
    if file_input is None:
        return None, None, None

    # Lecture buffer ou file path
    try:
        if isinstance(file_input, (str, os.PathLike)):
            if file_input.endswith('.csv'):
                df_raw = pd.read_csv(file_input, sep=None, engine='python', encoding='utf-8')
            else:
                df_raw = pd.read_excel(file_input, sheet_name='Export')
        else:
            filename = getattr(file_input, 'name', '')
            if filename.endswith('.csv'):
                try:
                    df_raw = pd.read_csv(file_input, sep=None, engine='python', encoding='utf-8')
                except UnicodeDecodeError:
                    file_input.seek(0)
                    df_raw = pd.read_csv(file_input, sep=None, engine='python', encoding='latin-1')
            else:
                df_raw = pd.read_excel(file_input, sheet_name='Export')
    except Exception as e:
        raise ValueError(f"Erreur lors de la lecture du fichier : {str(e)}")

    # Nettoyage des lignes parasites
    df_clean = df_raw.dropna(how='all').copy()
    
    # Exclusion footer "Filtres appliqués" et ligne Société == "Total"
    if 'Société' in df_clean.columns:
        df_clean = df_clean[df_clean['Société'] != 'Total']
        df_clean = df_clean[~df_clean['Société'].astype(str).str.startswith('Filtres appliqués', na=False)]
    
    # Vérification des colonnes nécessaires
    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df_clean.columns]
    if missing_cols:
        raise KeyError(f"Colonnes obligatoires manquantes dans la feuille Export : {missing_cols}")

    # Conversion des colonnes numériques
    num_cols = [
        'CA N-1', 'Budget', 'CA', 'Poids', 'Vs N-1 (%)', 'Vs Bgt (%)',
        'Marge N-1', 'Marge', 'Taux de Marge N-1', 'Taux de Marge',
        'Débit N-1', 'Débit', 'Panier N-1', 'Panier', 'Volume N-1', 'Volume'
    ]
    for col in num_cols:
        if col in df_clean.columns:
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce')

    # Parsing Code/Libellé et Formats
    df_clean['Rayon_Code'], df_clean['Rayon_Lib'] = zip(*df_clean['Rayon'].apply(split_code_libelle))
    df_clean['Site_Code'], df_clean['Site_Lib'] = zip(*df_clean['Site'].apply(split_code_libelle))
    df_clean['Format'] = df_clean['Site_Lib'].apply(extract_format)

    # Séparation aux 3 niveaux de granularité
    # 1. Global / Société
    df_global = df_clean[(df_clean['Rayon'] == 'Total') & (df_clean['Site'].isna())].copy()
    
    # 2. Rayon (Toutes enseignes)
    df_rayon = df_clean[(df_clean['Rayon'] != 'Total') & (df_clean['Site'] == 'Total')].copy()

    # 3. Couple Magasin x Rayon
    df_couple = df_clean[
        (df_clean['Rayon'] != 'Total') & 
        (df_clean['Site'].notna()) & 
        (df_clean['Site'] != 'Total')
    ].copy()

    return df_global, df_rayon, df_couple

@st.cache_data
def load_data(file_input):
    return _load_data_impl(file_input)


# ==============================================================================
# 2. MOTEUR DE CALCUL DES FLOPS ET SÉVÉRITÉ
# ==============================================================================
def compute_flops(df_couple, seuil_ca=-0.10, seuil_bgt=-0.10, seuil_marge=-0.8):
    """
    Calcule la détection des flops au niveau Couple Magasin x Rayon.
    - seuil_ca, seuil_bgt: en ratio décimal (ex: -0.10 pour -10%)
    - seuil_marge: en points de pourcentage (ex: -0.8 pt)
    """
    if df_couple is None or df_couple.empty:
        return pd.DataFrame()

    df = df_couple.copy()

    # Calcul Delta Marge en points de pourcentage
    df['Delta_Marge_pt'] = (df['Taux de Marge'] - df['Taux de Marge N-1']) * 100

    def evaluate_row(row):
        ca_n1 = row.get('CA N-1', np.nan)
        ca_n = row.get('CA', np.nan)
        vs_n1 = row.get('Vs N-1 (%)', np.nan)
        vs_bgt = row.get('Vs Bgt (%)', np.nan)
        bgt = row.get('Budget', np.nan)
        delta_marge = row.get('Delta_Marge_pt', np.nan)

        # C4: Rupture / Fermeture (CA NaN ou 0 alors que CA N-1 > 0)
        c4 = (pd.isna(ca_n) or ca_n == 0) and (not pd.isna(ca_n1) and ca_n1 > 0)
        
        if c4:
            return pd.Series({
                'C1': False, 'C2': False, 'C3': False, 'C4': True,
                'Nb_KO': 4, 'Nb_Applicable': 4, 'Score': '4/4',
                'Severite': 'Critique', 'Emoji': '🔴'
            })

        # C1: Décrochage CA vs N-1
        c1_app = not pd.isna(vs_n1)
        c1 = (vs_n1 <= seuil_ca) if c1_app else False

        # C2: Écart vs Budget (Ignoré si Budget NaN)
        c2_app = not pd.isna(bgt) and not pd.isna(vs_bgt)
        c2 = (vs_bgt <= seuil_bgt) if c2_app else False

        # C3: Dégradation de marge
        c3_app = not pd.isna(delta_marge)
        c3 = (delta_marge <= seuil_marge) if c3_app else False

        nb_app = sum([c1_app, c2_app, c3_app])
        nb_ko = sum([c1, c2, c3])

        # Qualification Séverité
        if nb_ko >= 2:
            sev = 'Flop majeur'
            emoji = '🟠'
        elif nb_ko == 1:
            sev = 'Flop modéré'
            emoji = '🟡'
        else:
            sev = 'OK'
            emoji = '🟢'

        score_str = f"{nb_ko}/{nb_app}" if nb_app > 0 else "0/0"

        return pd.Series({
            'C1': c1, 'C2': c2, 'C3': c3, 'C4': False,
            'Nb_KO': nb_ko, 'Nb_Applicable': nb_app, 'Score': score_str,
            'Severite': sev, 'Emoji': emoji
        })

    eval_df = df.apply(evaluate_row, axis=1)
    res = pd.concat([df, eval_df], axis=1)
    
    # 1. Analytique : CA à risque (Pareto)
    res['CA_a_risque'] = (res['CA N-1'] - res['CA'].fillna(0)).clip(lower=0)
    
    # 2. Analytique : Benchmark par pairs de format
    res['Vs_N1_Pairs_Mean'] = res.groupby(['Format', 'Rayon'])['Vs N-1 (%)'].transform(
        lambda x: (x.sum() - x) / (x.count() - 1) if x.count() > 1 else np.nan
    )
    res['Ecart_vs_pairs'] = res['Vs N-1 (%)'] - res['Vs_N1_Pairs_Mean']

    # 3. Analytique : Décomposition Trafic / Panier
    def compute_decomp(r):
        d_n1, d_n = r.get('Débit N-1'), r.get('Débit')
        p_n1, p_n = r.get('Panier N-1'), r.get('Panier')
        if pd.isna(d_n1) or pd.isna(d_n) or pd.isna(p_n1) or pd.isna(p_n):
            return np.nan, np.nan, "n/a"
        
        eff_trafic = (d_n - d_n1) * ((p_n + p_n1) / 2)
        eff_panier = (p_n - p_n1) * ((d_n + d_n1) / 2)
        
        if abs(eff_trafic) > abs(eff_panier):
            moteur = "Trafic (Fréquentation)"
        else:
            moteur = "Panier moyen"
        return eff_trafic, eff_panier, moteur

    res[['Effet_Trafic', 'Effet_Panier', 'Moteur_Perte']] = res.apply(
        compute_decomp, axis=1, result_type='expand'
    )

    return res


# ==============================================================================
# 3. MOTEUR DE COMMENTAIRE AUTOMATIQUE DE RENTABILITÉ
# ==============================================================================
def generate_rayon_comment(vs_n1, delta_marge_pt, vs_bgt, seuil_ca=-0.10, seuil_marge=-0.8):
    """Génère le commentaire métier croisant CA, Marge et Budget au niveau Rayon."""
    if pd.isna(vs_n1):
        return "Données insuffisantes pour établir un diagnostic."

    # Axe CA
    if vs_n1 >= 0:
        axe_ca = "croissance"
    elif vs_n1 > seuil_ca:
        axe_ca = "recul"
    else:
        axe_ca = "decrochage"

    # Axe Marge
    seuil_marge_abs = abs(seuil_marge)
    if pd.isna(delta_marge_pt):
        axe_marge = "stable"
    elif delta_marge_pt >= seuil_marge_abs:
        axe_marge = "amelioration"
    elif delta_marge_pt <= seuil_marge:
        axe_marge = "degradation"
    else:
        axe_marge = "stable"

    # Matrice 3x3 des commentaires
    matrice = {
        ("croissance", "amelioration"): "Forte dynamique commerciale portée par une excellente expansion des volumes et une rentabilité renforcée.",
        ("croissance", "stable"): "Solide performance du chiffre d'affaires maintenant une rentabilité conforme aux standards.",
        ("croissance", "degradation"): "Croissance tirée par l'activité au détriment du taux de marge : vigilance requis sur le mix produits.",
        ("recul", "amelioration"): "Légère contraction du chiffre d'affaires compensée par un mix plus contributif et une marge préservée.",
        ("recul", "stable"): "Activité commerciale en léger retrait, préservant ses équilibres de marge globale.",
        ("recul", "degradation"): "Effet ciseau défavorable : effritement simultané des ventes et de la rentabilité opérationnelle.",
        ("decrochage", "amelioration"): "Décrochage de CA compensé par un pilotage marge défensif : moins de volume, marge préservée.",
        ("decrochage", "stable"): "Perte de volume significative sans dégradation du taux de marge unitaires : perte de parts de marché à corriger.",
        ("decrochage", "degradation"): "Alerte majeure : effondrement combiné des volumes d'affaires et de la rentabilité brute."
    }

    base_text = matrice.get((axe_ca, axe_marge), "Évolution à surveiller.")

    # Suffixe Budget
    if pd.isna(vs_bgt):
        suffixe = "Pas d'objectif budgétaire alloué."
    elif vs_bgt >= 0:
        suffixe = "Objectif budget atteint."
    elif vs_bgt >= -0.05:
        suffixe = "En léger retard sur le budget target."
    else:
        suffixe = "Nettement sous l'objectif budgétaire."

    return f"{base_text} | {suffixe}"


# ==============================================================================
# 4. HELPERS D'AFFICHAGE ET ANALYTIQUE UI
# ==============================================================================
def render_kpi_card(title, value, sub="", delta_color="neutral"):
    """Génère une carte KPI au design iOS 3D ou Flat."""
    card_class = "ios-card-3d" if KPI_STYLE_3D else "ios-card"
    if delta_color == "alert":
        card_class = "ios-card-3d-alert" if KPI_STYLE_3D else "ios-card"
    elif delta_color == "success":
        card_class = "ios-card-3d-success" if KPI_STYLE_3D else "ios-card"

    html = f"""
    <div class="{card_class}">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

def parse_dataframe_selection(selection_event):
    """Parse de manière ultra-robuste le retour d'un st.dataframe(on_select='rerun')."""
    if not selection_event:
        return []
    
    # Cas 1: Objet avec attribut .selection.rows
    if hasattr(selection_event, "selection"):
        sel_obj = getattr(selection_event, "selection")
        if hasattr(sel_obj, "rows"):
            return getattr(sel_obj, "rows")
        elif isinstance(sel_obj, dict) and "rows" in sel_obj:
            return sel_obj["rows"]

    # Cas 2: Dictionnaire brut {"selection": {"rows": [...]}}
    if isinstance(selection_event, dict):
        sel_dict = selection_event.get("selection", {})
        if isinstance(sel_dict, dict):
            return sel_dict.get("rows", [])
            
    return []


# ==============================================================================
# 5. APPLICATION STREAMLIT (MAIN RENDERING)
# ==============================================================================
def main():
    # Sidebar: Initialisation et upload
    st.sidebar.image("https://img.icons8.com/color/96/analytics.png", width=64)
    st.sidebar.title("Paramètres & Data")

    uploaded_file = st.sidebar.file_uploader(
        "Charger l'export Ventes (.xlsx ou .csv)",
        type=["xlsx", "csv"]
    )

    # Si pas de fichier chargé : Écran d'accueil
    if uploaded_file is None:
        st.title("💸 Reporting Vente CA — Pilotage & Orientation Acheteurs")
        st.markdown("""
        <div class="ios-card">
            <h3>Bienvenue dans le module de Reporting Vente CA</h3>
            <p>Cet outil offre un point de situation commercial rapide, du niveau global au couple <b>Magasin × Rayon</b>, avec un focus sur la détection automatique des <b>Flops</b> et la priorisation des plans d'action.</p>
            <h4>Fonctionnalités clés :</h4>
            <ul>
                <li>⚡ <b>Vue d'ensemble 10s :</b> Prise de température globale et Steering Wheel 4 cadrans par rayon.</li>
                <li>🎯 <b>Détection des Flops :</b> Identification selon 4 critères métier (CA, Budget, Marge, Rupture C4).</li>
                <li>📋 <b>Briefs Acheteurs :</b> Synthèse automatique copiable par rayon avec matrice de rentabilité.</li>
                <li>📊 <b>Analytique poussée :</b> Logique Pareto du CA à risque, décomposition Trafic/Panier et benchmark par pairs de format.</li>
                <li>📤 <b>Export COPIL :</b> Génération d'un classeur multi-onglets complet en un seul clic.</li>
            </ul>
            <p><i>👈 Veuillez charger votre fichier <b>data.xlsx</b> (feuille Export) dans la barre latérale pour lancer l'analyse.</i></p>
        </div>
        """, unsafe_allow_html=True)
        return

    # Chargement robuste des données
    try:
        df_global, df_rayon, df_couple = load_data(uploaded_file)
    except Exception as e:
        st.error(f"⚠️ {str(e)}")
        return

    # Sidebar: Presets et Sliders de Seuils
    st.sidebar.markdown("---")
    st.sidebar.subheader("🎯 Seuils d'Alerte")

    if 'seuil_ca' not in st.session_state:
        st.session_state.seuil_ca = -10.0
    if 'seuil_bgt' not in st.session_state:
        st.session_state.seuil_bgt = -10.0
    if 'seuil_marge' not in st.session_state:
        st.session_state.seuil_marge = -0.8

    col_p1, col_p2, col_p3 = st.sidebar.columns(3)
    if col_p1.button("Strict"):
        st.session_state.seuil_ca = -5.0
        st.session_state.seuil_bgt = -5.0
        st.session_state.seuil_marge = -0.5
    if col_p2.button("Standard"):
        st.session_state.seuil_ca = -10.0
        st.session_state.seuil_bgt = -10.0
        st.session_state.seuil_marge = -0.8
    if col_p3.button("Souple"):
        st.session_state.seuil_ca = -15.0
        st.session_state.seuil_bgt = -15.0
        st.session_state.seuil_marge = -1.2

    seuil_ca_val = st.sidebar.slider("Decrochage CA (%)", -30.0, 0.0, st.session_state.seuil_ca, step=1.0) / 100.0
    seuil_bgt_val = st.sidebar.slider("Écart Budget (%)", -30.0, 0.0, st.session_state.seuil_bgt, step=1.0) / 100.0
    seuil_marge_val = st.sidebar.slider("Dégradation Marge (pt)", -3.0, 0.0, st.session_state.seuil_marge, step=0.1)

    # Sidebar: Filtres métier
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filtres")

    societes = df_couple['Société'].unique().tolist()
    selected_soc = st.sidebar.selectbox("Société", societes)

    rayons_available = df_couple['Rayon_Lib'].dropna().unique().tolist()
    selected_rayons = st.sidebar.multiselect("Rayons", rayons_available, default=rayons_available)

    formats_available = df_couple['Format'].dropna().unique().tolist()
    selected_formats = st.sidebar.multiselect("Formats", formats_available, default=formats_available)

    # Filtre Magasin dépendant du format
    magasins_filtered = df_couple[df_couple['Format'].isin(selected_formats)]['Site_Lib'].dropna().unique().tolist()
    selected_magasins = st.sidebar.multiselect("Magasins", magasins_filtered, default=magasins_filtered)

    # Calcul des flops avec les seuils dynamiques
    df_evaluated = compute_flops(df_couple, seuil_ca_val, seuil_bgt_val, seuil_marge_val)

    # Appliquer les filtres
    df_filtered = df_evaluated[
        (df_evaluated['Société'] == selected_soc) &
        (df_evaluated['Rayon_Lib'].isin(selected_rayons)) &
        (df_evaluated['Format'].isin(selected_formats)) &
        (df_evaluated['Site_Lib'].isin(selected_magasins))
    ].copy()

    # HEADER + BANDEAU KPI GLOBAL
    col_head, col_btn = st.columns([3, 1])
    with col_head:
        st.title(f"Point de situation — {selected_soc}")
    with col_btn:
        st.write("") # Spacing
        # Bouton d'export global unique
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_global.to_excel(writer, sheet_name='KPI Global', index=False)
            df_rayon.to_excel(writer, sheet_name='Par Rayon', index=False)
            df_filtered.to_excel(writer, sheet_name='Flops', index=False)
            df_couple.to_excel(writer, sheet_name='Données brutes', index=False)
        
        st.download_button(
            label="📤 Export COPIL (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"COPIL_Reporting_Vente_CA_{selected_soc}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Top KPI Global Cards
    if not df_global.empty:
        glob = df_global.iloc[0]
        c1, c2, c3, c4, c5 = st.columns(5)
        
        ca_tot = glob.get('CA', 0)
        vs_n1_tot = glob.get('Vs N-1 (%)', 0)
        vs_bgt_tot = glob.get('Vs Bgt (%)', 0)
        marge_tot = glob.get('Marge', 0)
        taux_marge_tot = glob.get('Taux de Marge', 0)

        with c1:
            render_kpi_card("CA Total", f"{ca_tot:,.0f} M".replace(',', ' '), "Réseau global")
        with c2:
            color = "success" if vs_n1_tot >= 0 else "alert"
            render_kpi_card("Vs N-1", f"{vs_n1_tot*100:+.1f} %", "Évolution CA", color)
        with c3:
            color = "success" if vs_bgt_tot >= 0 else "alert"
            render_kpi_card("Vs Budget", f"{vs_bgt_tot*100:+.1f} %" if not pd.isna(vs_bgt_tot) else "n/a", "Atteinte bgt", color)
        with c4:
            render_kpi_card("Marge Brute", f"{marge_tot:,.0f} M".replace(',', ' '), "Valeur FCFA")
        with c5:
            render_kpi_card("Taux de Marge", f"{taux_marge_tot*100:.1f} %", "Taux moyen")

    # Bandeau d'alerte Ruptures Critiques
    nb_critiques = len(df_filtered[df_filtered['Severite'] == 'Critique'])
    if nb_critiques > 0:
        st.error(f"🚨 **ALERTE COPIL :** {nb_critiques} rupture(s) totale(s) / fermeture(s) (Critères C4) détectée(s) dans le périmètre filtré !")

    st.markdown("---")

    # ONGLETS DE L'APPLICATION
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🎯 Vue d'ensemble", "🏷️ Par Rayon", "🚩 Flops", "📖 Méthodologie", "📤 Export"
    ])

    # --------------------------------------------------------------------------
    # TAB 1: VUE D'ENSEMBLE
    # --------------------------------------------------------------------------
    with tab1:
        st.subheader("🔥 Top 5 points d'attention (Priorisation CA à risque)")
        top_5_flops = df_filtered[df_filtered['Severite'] != 'OK'].sort_values('CA_a_risque', ascending=False).head(5)
        
        if top_5_flops.empty:
            st.success("🎉 Aucun flop détecté dans le périmètre sélectionné !")
        else:
            cols = st.columns(len(top_5_flops))
            for idx, (_, flop) in enumerate(top_5_flops.iterrows()):
                with cols[idx]:
                    st.markdown(f"""
                    <div class="ios-card-3d-alert">
                        <div><b>{flop['Emoji']} {flop['Site_Lib']}</b></div>
                        <div style="color: #8E8E93; font-size:0.8rem;">{flop['Rayon_Lib']}</div>
                        <hr style="margin: 8px 0;">
                        <div><b>Perte CA :</b> {flop['CA_a_risque']:,.0f} FCFA</div>
                        <div><b>Score KO :</b> {flop['Score']}</div>
                        <div><b>Vs N-1 :</b> {flop['Vs N-1 (%)']*100:+.1f}%</div>
                    </div>
                    """, unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("📊 Point de situation par rayon (Steering Wheel)")

        # Synthèse par Rayon sur 4 Cadrans
        for r_lib in selected_rayons:
            df_r_couples = df_filtered[df_filtered['Rayon_Lib'] == r_lib]
            df_r_total = df_rayon[df_rayon['Rayon_Lib'] == r_lib]

            if df_r_total.empty:
                continue

            r_data = df_r_total.iloc[0]
            vs_n1_r = r_data.get('Vs N-1 (%)', 0)
            vs_bgt_r = r_data.get('Vs Bgt (%)', 0)
            delta_m_r = (r_data.get('Taux de Marge', 0) - r_data.get('Taux de Marge N-1', 0)) * 100
            debit_vs = r_data.get('Vs N-1 (%).1', 0)
            panier_vs = r_data.get('Panier N Vs N-1', 0)

            st.markdown(f"#### 🏷️ Rayon : {r_lib}")
            sw1, sw2, sw3, sw4 = st.columns(4)

            with sw1:
                st.markdown(f"""
                <div class="ios-card">
                    <b>👤 Client (Trafic & Panier)</b><br>
                    • Débit (Trafic) : <b style="color:{'#34C759' if debit_vs >= 0 else '#FF3B30'}">{debit_vs*100:+.1f}%</b><br>
                    • Panier Moyen : <b style="color:{'#34C759' if panier_vs >= 0 else '#FF3B30'}">{panier_vs*100:+.1f}%</b>
                </div>
                """, unsafe_allow_html=True)

            with sw2:
                st.markdown(f"""
                <div class="ios-card">
                    <b>💰 Finance (CA & Marge)</b><br>
                    • CA vs N-1 : <b style="color:{'#34C759' if vs_n1_r >= 0 else '#FF3B30'}">{vs_n1_r*100:+.1f}%</b><br>
                    • Δ Marge : <b style="color:{'#34C759' if delta_m_r >= 0 else '#FF3B30'}">{delta_m_r:+.2f} pt</b>
                </div>
                """, unsafe_allow_html=True)

            with sw3:
                nb_alertes = len(df_r_couples[df_r_couples['Severite'] != 'OK'])
                st.markdown(f"""
                <div class="ios-card">
                    <b>📊 Activité & Alertes</b><br>
                    • Points d'attention : <b>{nb_alertes}</b><br>
                    • Nb Magasins : <b>{len(df_r_couples)}</b>
                </div>
                """, unsafe_allow_html=True)

            with sw4:
                st.markdown("""
                <div class="ios-card" style="opacity: 0.6; background-color: #E5E5EA;">
                    <b>⚙️ Opérations</b><br>
                    <i>À connecter plus tard (Données dispo/rupture)</i>
                </div>
                """, unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("🏪 Magasins les plus en difficulté (Agrégation)")
        mag_summary = df_filtered.groupby('Site_Lib').agg(
            Nb_Critiques=('C4', 'sum'),
            Nb_Flops=('Severite', lambda x: (x != 'OK').sum()),
            CA_A_Risque_Total=('CA_a_risque', 'sum')
        ).reset_index().sort_values(by=['Nb_Critiques', 'Nb_Flops'], ascending=False)
        st.dataframe(mag_summary, use_container_width=True)

    # --------------------------------------------------------------------------
    # TAB 2: PAR RAYON (BRIEF ACHETEUR)
    # --------------------------------------------------------------------------
    with tab2:
        selected_r_brief = st.selectbox("Sélectionner un rayon pour le brief :", selected_rayons)
        
        df_r_couples = df_filtered[df_filtered['Rayon_Lib'] == selected_r_brief]
        df_r_total = df_rayon[df_rayon['Rayon_Lib'] == selected_r_brief]

        if not df_r_total.empty:
            r_row = df_r_total.iloc[0]
            
            # Header avec Avatar Trigramme
            trigramme = selected_r_brief[:3].upper()
            c_av, c_info = st.columns([1, 8])
            with c_av:
                st.markdown(f'<div class="rayon-avatar">{trigramme}</div>', unsafe_allow_html=True)
            with c_info:
                st.subheader(f"Rayon {selected_r_brief} — {len(df_r_couples)} magasins")

            # Bandeau 6 KPIs
            k1, k2, k3, k4, k5, k6 = st.columns(6)
            with k1:
                render_kpi_card("CA Rayon", f"{r_row['CA']:,.0f} M".replace(',', ' '))
            with k2:
                render_kpi_card("Marge", f"{r_row['Marge']:,.0f} M".replace(',', ' '))
            with k3:
                render_kpi_card("Poids CA", f"{r_row['Poids']*100:.1f} %")
            with k4:
                render_kpi_card("Critiques", f"{len(df_r_couples[df_r_couples['Severite']=='Critique'])}", delta_color="alert")
            with k5:
                render_kpi_card("Flops Maj.", f"{len(df_r_couples[df_r_couples['Severite']=='Flop majeur'])}")
            with k6:
                render_kpi_card("Flops Mod.", f"{len(df_r_couples[df_r_couples['Severite']=='Flop modéré'])}")

            # Tops Magasins côte à côte
            st.markdown("### 🏆 Tops & Flops Magasins")
            col_t1, col_t2 = st.columns(2)
            with col_t1:
                st.markdown("<b>Top 3 Contributeurs CA :</b>", unsafe_allow_html=True)
                top_ca = df_r_couples.sort_values('CA', ascending=False).head(3)
                st.dataframe(top_ca[['Site_Lib', 'CA', 'Vs N-1 (%)']], use_container_width=True)
            with col_t2:
                st.markdown("<b>Top 3 Progression Vs N-1 :</b>", unsafe_allow_html=True)
                top_prog = df_r_couples.sort_values('Vs N-1 (%)', ascending=False).head(3)
                st.dataframe(top_prog[['Site_Lib', 'Vs N-1 (%)', 'CA']], use_container_width=True)

            # Génération du Brief Automatique Copiable
            vs_n1_val = r_row.get('Vs N-1 (%)', np.nan)
            delta_m_pt = (r_row.get('Taux de Marge', 0) - r_row.get('Taux de Marge N-1', 0)) * 100
            vs_bgt_val = r_row.get('Vs Bgt (%)', np.nan)
            
            comm_auto = generate_rayon_comment(vs_n1_val, delta_m_pt, vs_bgt_val, seuil_ca_val, seuil_marge_val)

            brief_text = f"""POINT RAYON {selected_r_brief.upper()} — Date: Aujourd'hui
SYNTHÈSE
  CA      : {r_row.get('CA', 0):,.0f} FCFA ({vs_n1_val*100:+.1f}%)
  Marge   : {r_row.get('Marge', 0):,.0f} FCFA ({r_row.get('Taux de Marge', 0)*100:.1f}% · {delta_m_pt:+.2f} pt)
  Qté     : {r_row.get('Volume', 0):,.0f} ({r_row.get('Volume N Vs N-1', 0)*100:+.1f}%)
  Débit   : {r_row.get('Vs N-1 (%).1', 0)*100:+.1f}%
  Panier  : {r_row.get('Panier N Vs N-1', 0)*100:+.1f}%
  [COMMENTAIRE] : {comm_auto}

TOP MAGASINS
  CA          : {", ".join([f"{r['Site_Lib']} ({r['CA']:,.0f})" for _, r in top_ca.iterrows()])}
  Progression : {", ".join([f"{r['Site_Lib']} ({r['Vs N-1 (%)']*100:+.1f}%)" for _, r in top_prog.iterrows()])}

POINTS D'ATTENTION ({len(df_r_couples[df_r_couples['Severite']!='OK'])})
"""
            for _, flop in df_r_couples[df_r_couples['Severite']!='OK'].iterrows():
                brief_text += f"  - [{flop['Severite']}] {flop['Site_Lib']} : CA {flop['CA']:,.0f} FCFA (Vs N-1 {flop['Vs N-1 (%)']*100:+.1f}%) | Marge {flop['Delta_Marge_pt']:+.2f} pt\n"

            st.markdown("### 📝 Brief Prêt à Partager (Copiable)")
            st.code(brief_text, language="text")

    # --------------------------------------------------------------------------
    # TAB 3: FLOPS (TABLE MAÎTRE-DÉTAL)
    # --------------------------------------------------------------------------
    with tab3:
        st.subheader("🚩 Consultation Maître-Détail des Flops")
        
        col_f1, col_f2 = st.columns([2, 1])
        with col_f1:
            search_query = st.text_input("🔍 Rechercher un magasin ou rayon :", "")
        with col_f2:
            sev_filter = st.multiselect("Filtrer la sévérité :", ["Critique", "Flop majeur", "Flop modéré", "OK"], default=["Critique", "Flop majeur", "Flop modéré"])

        # Filtrage dynamique
        df_table = df_filtered[df_filtered['Severite'].isin(sev_filter)].copy()
        if search_query:
            df_table = df_table[
                df_table['Site_Lib'].str.contains(search_query, case=False, na=False) |
                df_table['Rayon_Lib'].str.contains(search_query, case=False, na=False)
            ]

        # Formatting pour st.dataframe
        df_display = df_table[[
            'Emoji', 'Severite', 'Site_Lib', 'Rayon_Lib', 'CA', 'Vs N-1 (%)',
            'Vs Bgt (%)', 'Delta_Marge_pt', 'CA_a_risque', 'Score'
        ]].rename(columns={
            'Site_Lib': 'Magasin', 'Rayon_Lib': 'Rayon', 'Delta_Marge_pt': 'Δ Marge (pt)'
        })

        col_left, col_right = st.columns([2, 1])
        with col_left:
            st.markdown("<i>Sélectionnez une ligne pour ouvrir le panneau de détail latéral :</i>", unsafe_allow_html=True)
            selection = st.dataframe(
                df_display,
                on_select="rerun",
                selection_mode="single-row",
                use_container_width=True
            )

        # Panneau de Détail Latéral
        with col_right:
            st.markdown("### 📋 Detail du Flop")
            selected_rows = parse_dataframe_selection(selection)
            
            if selected_rows:
                selected_idx = selected_rows[0]
                row_detail = df_table.iloc[selected_idx]

                st.markdown(f"#### {row_detail['Emoji']} {row_detail['Site_Lib']}")
                st.markdown(f"**Rayon :** {row_detail['Rayon_Lib']} | **Format :** {row_detail['Format']}")
                st.markdown(f"**Sévérité :** {row_detail['Severite']} (Score {row_detail['Score']})")
                st.markdown("---")

                # Statut des 4 critères
                st.markdown("**Statut des Critères :**")
                c1_icon = "❌" if row_detail['C1'] else "✅"
                c2_icon = "n/a" if pd.isna(row_detail['Budget']) else ("❌" if row_detail['C2'] else "✅")
                c3_icon = "n/a" if pd.isna(row_detail['Delta_Marge_pt']) else ("❌" if row_detail['C3'] else "✅")
                c4_icon = "🚨 KO (Rupture)" if row_detail['C4'] else "✅ OK"

                st.write(f"- C1 (Decrochage CA <= {seuil_ca_val*100:.0f}%) : {c1_icon}")
                st.write(f"- C2 (Ecart Budget <= {seuil_bgt_val*100:.0f}%) : {c2_icon}")
                st.write(f"- C3 (Degradation Marge <= {seuil_marge_val:.1f} pt) : {c3_icon}")
                st.write(f"- C4 (Rupture / Fermeture) : {c4_icon}")
                st.markdown("---")

                # Analytique : Trafic vs Panier
                st.markdown("**Décomposition Trafic / Panier :**")
                eff_t = row_detail['Effet_Trafic']
                eff_p = row_detail['Effet_Panier']
                if pd.isna(eff_t):
                    st.write("Données insuffisantes.")
                else:
                    st.write(f"- Effet Trafic : {eff_t:,.0f} FCFA")
                    st.write(f"- Effet Panier : {eff_p:,.0f} FCFA")
                    st.info(f"💡 Perte principalement tirée par : **{row_detail['Moteur_Perte']}**")

                # Benchmark Pairs
                st.markdown("---")
                st.markdown("**Benchmark Pairs de Format :**")
                ecart_p = row_detail['Ecart_vs_pairs']
                if pd.isna(ecart_p):
                    st.write("Pas assez de pairs pour comparer.")
                else:
                    st.write(f"Écart vs moyenne des pairs : **{ecart_p*100:+.1f} pt**")
                    if ecart_p < -0.05:
                        st.warning("⚠️ Sous-performance spécifique au magasin.")
                    else:
                        st.info("ℹ️ Tendance alignée avec le réseau/format.")
            else:
                st.info("👉 Cliquez sur une ligne du tableau pour afficher le diagnostic complet.")

    # --------------------------------------------------------------------------
    # TAB 4: MÉTHODOLOGIE
    # --------------------------------------------------------------------------
    with tab4:
        st.subheader("📖 Méthodologie & Règles Métier")
        st.markdown("""
        ### 1. Périmètre & Niveaux de Données
        L'application segmente automatiquement l'export Excel (`Export`) en 3 niveaux :
        - **Global / Société :** Ligne unique `Rayon == "Total"` et `Site == NaN`.
        - **Rayon (Toutes enseignes) :** Lignes `Rayon != "Total"` et `Site == "Total"`.
        - **Couple Magasin × Rayon :** Lignes `Rayon != "Total"` et `Site` renseigné (niveau de détection des Flops).

        ---
        ### 2. Les 4 Critères de Détection des Flops
        1. **C1 — Décrochage CA vs N-1 :** $Vs\ N-1\ (\%)\le seuil\_ca$ (Défaut : $-10\%$).
        2. **C2 — Écart vs Budget :** $Vs\ Bgt\ (\%)\le seuil\_bgt$ (Défaut : $-10\%$). Ignoré si Budget est NaN (ex. Supeco).
        3. **C3 — Dégradation de marge :** $\Delta Marge\ (pt) \le seuil\_marge$ (Défaut : $-0.8\text{ pt}$).
        4. **C4 — Rupture / Fermeture :** CA NaN ou 0 alors que CA N-1 > 0. Prioritaire (Score 4/4, Sévérité **Critique**).

        ---
        ### 3. Analytique Avancée
        - **CA à risque (Logique Pareto) :** $CA\_a\_risque = \max(0, CA_{N-1} - CA)$. Perte nette en FCFA.
        - **Benchmark par pairs de format :** Écart de performance du couple par rapport à la moyenne des magasins du même Format sur le même Rayon.
        - **Décomposition Trafic / Panier :** 
          $$\text{Effet Trafic} = (\text{Débit}_N - \text{Débit}_{N-1}) \times \frac{\text{Panier}_N + \text{Panier}_{N-1}}{2}$$
          $$\text{Effet Panier} = (\text{Panier}_N - \text{Panier}_{N-1}) \times \frac{\text{Débit}_N + \text{Débit}_{N-1}}{2}$$
        """)

    # --------------------------------------------------------------------------
    # TAB 5: EXPORT
    # --------------------------------------------------------------------------
    with tab5:
        st.subheader("📤 Export des Données COPIL")
        st.write("Téléchargez l'intégralité des analyses et données brutes au format Excel multi-onglets :")
        st.download_button(
            label="📥 Télécharger le Classeur Excel COPIL (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"COPIL_Reporting_Vente_CA_{selected_soc}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# ==============================================================================
# 6. BLOC DE TESTS UNITAIRES
# ==============================================================================
def run_all_tests():
    """Exécute les tests unitaires métier sans dépendance Streamlit."""
    print("=== DÉBUT DU RUN DES TESTS UNITAIRES METIER ===")

    # Test 1: Split code / libellé
    code, lib = split_code_libelle("010 - BOISSON")
    assert code == "010" and lib == "BOISSON", f"Test 1 Échoué: {code}, {lib}"

    # Test 2: Extraction Format
    fmt = extract_format("10605 - Supeco Aboboté")
    assert fmt == "Supeco", f"Test 2 Échoué: {fmt}"

    # Test 3: Évaluation Flop C4 (Priorité Rupture)
    df_dummy = pd.DataFrame([{
        'Société': 'ADIALEA RCI', 'Rayon': '010 - BOISSON', 'Site': '10605 - Supeco Aboboté',
        'Rayon_Lib': 'BOISSON', 'Format': 'Supeco', 'Site_Lib': 'Supeco Aboboté',
        'CA N-1': 1000.0, 'Budget': np.nan, 'CA': np.nan, 'Vs N-1 (%)': -1.0,
        'Vs Bgt (%)': np.nan, 'Taux de Marge N-1': 0.20, 'Taux de Marge': np.nan,
        'Débit N-1': 100, 'Débit': np.nan, 'Panier N-1': 10, 'Panier': np.nan
    }])
    res_dummy = compute_flops(df_dummy)
    assert res_dummy.iloc[0]['Severite'] == 'Critique', f"Test 3 Échoué: {res_dummy.iloc[0]['Severite']}"
    assert res_dummy.iloc[0]['Score'] == '4/4', f"Test 3 Score Échoué: {res_dummy.iloc[0]['Score']}"

    # Test 4: Flop majeur (2 critères) + Budget NaN ignoré
    df_dummy2 = pd.DataFrame([{
        'Société': 'ADIALEA RCI', 'Rayon': '010 - BOISSON', 'Site': '10601 - Supeco Niangon',
        'Rayon_Lib': 'BOISSON', 'Format': 'Supeco', 'Site_Lib': 'Supeco Niangon',
        'CA N-1': 1000.0, 'Budget': np.nan, 'CA': 750.0, 'Vs N-1 (%)': -0.25,
        'Vs Bgt (%)': np.nan, 'Taux de Marge N-1': 0.20, 'Taux de Marge': 0.18, # -2 pt
        'Débit N-1': 100, 'Débit': 90, 'Panier N-1': 10, 'Panier': 8.33
    }])
    res_dummy2 = compute_flops(df_dummy2, seuil_ca=-0.10, seuil_marge=-0.8)
    assert res_dummy2.iloc[0]['Severite'] == 'Flop majeur', f"Test 4 Échoué: {res_dummy2.iloc[0]['Severite']}"
    assert res_dummy2.iloc[0]['Score'] == '2/2', f"Test 4 Score Échoué: {res_dummy2.iloc[0]['Score']}"

    # Test 5: Commentaire Rentabilité Matrice 3x3
    comm = generate_rayon_comment(vs_n1=-0.15, delta_marge_pt=1.5, vs_bgt=0.02, seuil_ca=-0.10, seuil_marge=-0.8)
    assert "Décrochage de CA compensé" in comm, f"Test 5 Échoué: {comm}"

    print("✅ TOUS LES TESTS UNITAIRES SONT PASSÉS AVEC SUCCÈS !")


# ==============================================================================
# 7. POINT D'ENTRÉE DU SCRIPT
# ==============================================================================
if __name__ == "__main__":
    if os.environ.get("RUN_DASHBOARD_TESTS") == "1":
        run_all_tests()
    else:
        main()
