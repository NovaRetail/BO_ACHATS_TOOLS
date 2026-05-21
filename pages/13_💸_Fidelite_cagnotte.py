import streamlit as st
import pandas as pd
import numpy as np
import re
from io import StringIO

# ─────────────────────────────────────────────
# CHARTE VISUELLE SmartBuyer Hub
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Suivi Fidélité · SmartBuyer Hub",
    page_icon="🏷️",
    layout="wide",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', 'Calibri', -apple-system, BlinkMacSystemFont, 'SF Pro Text', sans-serif;
    background-color: #F2F2F7;
}

/* Main background */
.stApp { background-color: #F2F2F7; }

/* Remove default padding */
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1400px; }

/* Sidebar */
[data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E5E5EA; }

/* KPI Cards */
.kpi-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 24px; }
.kpi-card {
    background: #FFFFFF;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08), 0 1px 2px rgba(0,0,0,0.04);
    border: 1px solid #E5E5EA;
}
.kpi-label { font-size: 11px; font-weight: 500; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
.kpi-value { font-size: 22px; font-weight: 700; color: #1C1C1E; line-height: 1.1; }
.kpi-value.blue { color: #007AFF; }
.kpi-value.red { color: #FF3B30; }
.kpi-value.green { color: #34C759; }
.kpi-sub { font-size: 11px; color: #8E8E93; margin-top: 4px; }

/* Period badge */
.period-badge {
    background: #EAF4FF;
    border: 1px solid #B8D9FF;
    border-radius: 10px;
    padding: 12px 20px;
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 16px;
}
.period-label { font-size: 11px; font-weight: 600; color: #007AFF; text-transform: uppercase; letter-spacing: 0.5px; }
.period-value { font-size: 15px; font-weight: 700; color: #007AFF; }
.period-meta { font-size: 12px; color: #3A3A3C; }

/* Section headers */
.section-header {
    font-size: 13px;
    font-weight: 600;
    color: #3A3A3C;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 12px;
    margin-top: 4px;
    padding-bottom: 6px;
    border-bottom: 2px solid #007AFF;
    display: inline-block;
}

/* Dataframe styling override */
.stDataFrame { border-radius: 10px; overflow: hidden; }

/* Tab styling */
[data-testid="stTabs"] [role="tab"] {
    font-size: 13px;
    font-weight: 500;
    padding: 8px 16px;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    color: #007AFF;
    border-bottom: 2px solid #007AFF;
}

/* Upload zone */
.upload-zone {
    background: #FFFFFF;
    border: 2px dashed #C7C7CC;
    border-radius: 12px;
    padding: 20px;
    text-align: center;
    margin-bottom: 16px;
}

/* Info box */
.info-box {
    background: #F2F2F7;
    border-left: 3px solid #007AFF;
    border-radius: 4px;
    padding: 10px 14px;
    font-size: 12px;
    color: #3A3A3C;
    margin-bottom: 12px;
}

/* Metric pill */
.pill-positive { background: #E8FAF0; color: #1A7F3C; border-radius: 6px; padding: 2px 8px; font-size: 12px; font-weight: 600; }
.pill-negative { background: #FFF0EE; color: #C0392B; border-radius: 6px; padding: 2px 8px; font-size: 12px; font-weight: 600; }

/* Hide streamlit branding */
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def parse_number(val):
    """Convert French-formatted number string to float."""
    if pd.isna(val) or str(val).strip() in ['', 'NaN', 'nan']:
        return np.nan
    s = str(val).replace('\xa0', '').replace(' ', '').replace(',', '.')
    try:
        return float(s)
    except:
        return np.nan

def fmt_xof(val, show_sign=False):
    """Format number as XOF with French spacing."""
    if pd.isna(val):
        return '—'
    val = int(round(val))
    sign = '+' if (show_sign and val > 0) else ''
    return f"{sign}{val:,}".replace(',', ' ') + ' XOF'

def fmt_num(val):
    if pd.isna(val): return '—'
    return f"{val:,.0f}".replace(',', ' ')

def fmt_pct(val):
    if pd.isna(val): return '—'
    return f"{val:.1f}%"

def extract_article_id(article_str):
    """Extract numeric article ID from '12001277 - CITRON MEYER' → 12001277"""
    if pd.isna(article_str):
        return None
    m = re.match(r'^(\d+)', str(article_str).strip())
    return int(m.group(1)) if m else None

def parse_period_from_lines(lines):
    """Scan all lines to find date range pattern."""
    for line in lines:
        m = re.search(r'après le (\d{2}/\d{2}/\d{4}).*?avant le (\d{2}/\d{2}/\d{4})', line)
        if m:
            d1 = pd.to_datetime(m.group(1), dayfirst=True)
            d2 = pd.to_datetime(m.group(2), dayfirst=True)
            return d1, d2
    return None, None

def get_semaine_mois(date_debut, date_fin):
    """Derive semaine ISO and mois label from dates."""
    if date_debut is None:
        return '—', '—'
    sem = f"S{date_debut.isocalendar()[1]}"
    mois_map = {1:'Janvier',2:'Février',3:'Mars',4:'Avril',5:'Mai',6:'Juin',
                7:'Juillet',8:'Août',9:'Septembre',10:'Octobre',11:'Novembre',12:'Décembre'}
    mois = mois_map.get(date_debut.month, str(date_debut.month))
    return sem, f"{mois} {date_debut.year}"

def load_ventes_csv(file_obj):
    """Load a ventes CSV file, extract period, return clean dataframe."""
    content = file_obj.read().decode('latin1')
    lines = content.split('\n')
    
    date_debut, date_fin = parse_period_from_lines(lines)
    sem, mois = get_semaine_mois(date_debut, date_fin)
    
    # Parse main data (skip rows after the last data row)
    data_lines = []
    header_found = False
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if 'Filtres appliqués' in stripped:
            break
        if 'Site nom long' in stripped:
            header_found = True
        if header_found:
            data_lines.append(stripped)
    
    if not data_lines:
        return None, date_debut, date_fin, sem, mois
    
    csv_str = '\n'.join(data_lines)
    df = pd.read_csv(StringIO(csv_str), sep=';', encoding='utf-8', on_bad_lines='skip')
    
    # Clean columns
    df.columns = [c.strip() for c in df.columns]
    
    # Parse numerics
    for col in ['CA', 'Marge', 'Qté Vente']:
        if col in df.columns:
            df[col] = df[col].apply(parse_number)
    
    # Keep only article-level rows (not totals)
    df = df[
        df['Article'].notna() &
        ~df['Article'].astype(str).str.strip().isin(['Total', 'NaN', '']) &
        df['Site nom long'].notna() &
        ~df['Site nom long'].astype(str).str.strip().isin(['Total', 'NaN', ''])
    ].copy()
    
    # Add temporal columns
    df['Date Début'] = date_debut.strftime('%d/%m/%Y') if date_debut else '—'
    df['Date Fin'] = date_fin.strftime('%d/%m/%Y') if date_fin else '—'
    df['Semaine'] = sem
    df['Mois'] = mois
    df['_mois_num'] = date_debut.month if date_debut else 0
    
    # Extract article ID
    df['_article_id'] = df['Article'].apply(extract_article_id)
    
    return df, date_debut, date_fin, sem, mois


def load_fidelite_csv(file_obj):
    """Load the fidélité reference list."""
    content = file_obj.read().decode('latin1')
    df = pd.read_csv(StringIO(content), sep=None, engine='python', encoding='utf-8')
    df.columns = [c.strip() for c in df.columns]
    # Normalize mois
    mois_map = {'mai': 'Mai', 'avril': 'Avril', 'mars': 'Mars', 'juin': 'Juin',
                'janvier': 'Janvier', 'février': 'Février', 'juillet': 'Juillet',
                'août': 'Août', 'septembre': 'Septembre', 'octobre': 'Octobre',
                'novembre': 'Novembre', 'décembre': 'Décembre'}
    if 'Mois' in df.columns:
        df['Mois'] = df['Mois'].astype(str).str.strip().str.lower().map(mois_map).fillna(df['Mois'])
    df = df.dropna(subset=['Article', 'Cagnotte'])
    df['Article'] = df['Article'].astype(int)
    df['Cagnotte'] = pd.to_numeric(df['Cagnotte'], errors='coerce')
    return df


def color_marge(val):
    if pd.isna(val) or val == '—':
        return ''
    try:
        n = float(str(val).replace(' ', '').replace('XOF', '').replace(',', '.'))
        if n < 0:
            return 'color: #FF3B30; font-weight: 600'
        elif n > 0:
            return 'color: #34C759; font-weight: 600'
    except:
        pass
    return ''

def color_poids(val):
    if pd.isna(val) or val == '—':
        return ''
    try:
        n = float(str(val).replace('%','').replace(',','.'))
        if n >= 30:
            return 'background-color: #E8FAF0; color: #1A7F3C; font-weight: 600'
        elif n >= 15:
            return 'background-color: #FFF9E6; color: #7D5A00; font-weight: 600'
        else:
            return 'background-color: #FFF0EE; color: #C0392B; font-weight: 600'
    except:
        pass
    return ''


# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
col_logo, col_title = st.columns([1, 11])
with col_logo:
    st.markdown("""
    <div style="width:44px;height:44px;background:#007AFF;border-radius:10px;
    display:flex;align-items:center;justify-content:center;margin-top:4px;">
    <span style="font-size:22px;">🏷️</span></div>
    """, unsafe_allow_html=True)
with col_title:
    st.markdown("""
    <div style="padding-top:6px;">
    <span style="font-size:18px;font-weight:700;color:#1C1C1E;">SmartBuyer Hub</span>
    <span style="font-size:14px;color:#8E8E93;margin-left:10px;">· Suivi Fidélité · Investissement vs Performance</span>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<hr style='margin:8px 0 16px 0;border:none;border-top:1px solid #E5E5EA;'>", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SIDEBAR — UPLOAD
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📂 Chargement des données")
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    
    st.markdown("**Fichiers ventes** *(multi-upload)*")
    ventes_files = st.file_uploader(
        "CSV extractions ventes",
        type=['csv'],
        accept_multiple_files=True,
        key="ventes_upload",
        label_visibility="collapsed"
    )
    
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    st.markdown("**Liste Fidélité** *(référentiel)*")
    fidelite_file = st.file_uploader(
        "CSV liste fidélité",
        type=['csv'],
        accept_multiple_files=False,
        key="fidelite_upload",
        label_visibility="collapsed"
    )
    
    st.markdown("<hr style='margin:16px 0;border:none;border-top:1px solid #E5E5EA;'>")
    st.markdown("""
    <div style='font-size:11px;color:#8E8E93;line-height:1.6;'>
    📌 <b>Format attendu :</b><br>
    • Ventes : export PBI (séparateur <code>;</code>)<br>
    • Fidélité : Article / Cagnotte / Mois<br>
    • Plusieurs semaines cumulables
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────
if not ventes_files or not fidelite_file:
    st.markdown("""
    <div style='background:#FFFFFF;border-radius:16px;padding:48px;text-align:center;
    border:1px solid #E5E5EA;box-shadow:0 1px 3px rgba(0,0,0,0.06);'>
        <div style='font-size:48px;margin-bottom:16px;'>🏷️</div>
        <div style='font-size:20px;font-weight:700;color:#1C1C1E;margin-bottom:8px;'>
            Suivi Fidélité
        </div>
        <div style='font-size:14px;color:#8E8E93;max-width:400px;margin:0 auto;'>
            Chargez vos fichiers ventes et votre liste fidélité dans le panneau gauche pour démarrer l'analyse.
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# Load ventes
all_dfs = []
periods = []
for f in ventes_files:
    df_v, d1, d2, sem, mois = load_ventes_csv(f)
    if df_v is not None and len(df_v) > 0:
        all_dfs.append(df_v)
        periods.append({'file': f.name, 'date_debut': d1, 'date_fin': d2, 'sem': sem, 'mois': mois})

if not all_dfs:
    st.error("Aucune donnée valide trouvée dans les fichiers ventes.")
    st.stop()

df_ventes = pd.concat(all_dfs, ignore_index=True)

# Load fidélité
df_fidelite = load_fidelite_csv(fidelite_file)

# ─────────────────────────────────────────────
# FILTERS (top bar)
# ─────────────────────────────────────────────
mois_disponibles = sorted(df_ventes['Mois'].dropna().unique().tolist())
mois_fidelite = sorted(df_fidelite['Mois'].dropna().unique().tolist())

col_f1, col_f2, col_f3 = st.columns([2, 2, 4])
with col_f1:
    mois_sel = st.selectbox("Mois actif", options=mois_disponibles, index=0)
with col_f2:
    rayons_dispo = sorted(df_ventes['Rayon'].dropna().unique().tolist())
    rayon_sel = st.selectbox("Rayon", options=['Tous'] + rayons_dispo)
with col_f3:
    sites_dispo = sorted(df_ventes['Site nom long'].dropna().unique().tolist())
    sites_sel = st.multiselect("Magasins", options=sites_dispo, default=[])


# ─────────────────────────────────────────────
# FILTER VENTES
# ─────────────────────────────────────────────
df_mois = df_ventes[df_ventes['Mois'] == mois_sel].copy()

# Extract mois label from selected (e.g. "Mai 2026" → "Mai")
mois_court = mois_sel.split(' ')[0]  # "Mai"

# Get fidélité for this mois
# Try exact match, else try substring
df_fid_mois = df_fidelite[df_fidelite['Mois'] == mois_court]
if len(df_fid_mois) == 0:
    # Fallback: all fidelité
    df_fid_mois = df_fidelite.copy()

# Apply site filter
if sites_sel:
    df_mois = df_mois[df_mois['Site nom long'].isin(sites_sel)]

# Apply rayon filter
if rayon_sel != 'Tous':
    df_mois = df_mois[df_mois['Rayon'] == rayon_sel]


# ─────────────────────────────────────────────
# JOIN: ventes × fidélité
# ─────────────────────────────────────────────
df_joined = df_mois.merge(
    df_fid_mois[['Article', 'Cagnotte']].rename(columns={'Article': '_article_id', 'Cagnotte': 'Cagnotte_unit'}),
    on='_article_id',
    how='left'
)
df_joined['est_fidelite'] = df_joined['Cagnotte_unit'].notna()
df_joined['Total Cagnotte'] = df_joined['Cagnotte_unit'] * df_joined['Qté Vente']


# ─────────────────────────────────────────────
# PERIOD BADGE
# ─────────────────────────────────────────────
# Find period for selected mois
period_info = next((p for p in periods if mois_sel in p['mois']), periods[0] if periods else None)
if period_info and period_info['date_debut']:
    d1_str = period_info['date_debut'].strftime('%d/%m/%Y')
    d2_str = period_info['date_fin'].strftime('%d/%m/%Y')
    sem_str = period_info['sem']
    mois_str = period_info['mois']
else:
    d1_str = d2_str = sem_str = mois_str = '—'

nb_files = len(ventes_files)
st.markdown(f"""
<div class="period-badge">
    <div>
        <div class="period-label">Période détectée</div>
        <div class="period-value">{d1_str} → {d2_str} · {sem_str} · {mois_str}</div>
    </div>
    <div style="margin-left:auto;font-size:12px;color:#3A3A3C;">
        {nb_files} fichier(s) chargé(s) · {len(df_fid_mois)} articles fidélité ({mois_court})
    </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# KPI COMPUTATION
# ─────────────────────────────────────────────
df_fid_only = df_joined[df_joined['est_fidelite']].copy()

budget_cagnotte = df_fid_only['Total Cagnotte'].sum()
ca_fidelite = df_fid_only['CA'].sum()
marge_fidelite = df_fid_only['Marge'].sum()

# Articles actifs vs total périmètre
arts_avec_ventes = df_fid_only[df_fid_only['CA'].notna() & (df_fid_only['CA'] > 0)]['_article_id'].nunique()
arts_total_perimetre = df_fid_mois['Article'].nunique()

# Couverture réseau
sites_actifs = df_fid_only[df_fid_only['CA'] > 0]['Site nom long'].nunique()
sites_total = df_mois['Site nom long'].nunique()

# KPI display
marge_class = "red" if marge_fidelite < 0 else "green"
ca_display = fmt_xof(ca_fidelite)
marge_display = fmt_xof(marge_fidelite)
budget_display = fmt_xof(budget_cagnotte)

st.markdown(f"""
<div class="kpi-grid">
    <div class="kpi-card">
        <div class="kpi-label">Budget Cagnotte</div>
        <div class="kpi-value blue">{budget_display}</div>
        <div class="kpi-sub">Investissement fidélité</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">CA Fidélité</div>
        <div class="kpi-value">{ca_display}</div>
        <div class="kpi-sub">Articles en programme</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">Marge Fidélité</div>
        <div class="kpi-value {marge_class}">{marge_display}</div>
        <div class="kpi-sub">Impact marge programme</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">Articles Actifs</div>
        <div class="kpi-value blue">{arts_avec_ventes} <span style="font-size:14px;color:#8E8E93;">/ {arts_total_perimetre}</span></div>
        <div class="kpi-sub">Avec ventes > 0 / Total périmètre {mois_court}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">Couverture Réseau</div>
        <div class="kpi-value blue">{sites_actifs} <span style="font-size:14px;color:#8E8E93;">/ {sites_total}</span></div>
        <div class="kpi-sub">Sites avec ventes fidélité</div>
    </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs([
    "📊 Récap Financier",
    "🔍 Récap Détail Article × Site",
    "📋 Drill-down Granulaire"
])


# ═══════════════════════════════════════════
# ONGLET 1 — RÉCAP FINANCIER (Site × Rayon × Famille)
# ═══════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-header">Récap Financier · Investissement vs Performance · ' + mois_court + '</div>', unsafe_allow_html=True)
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.markdown(f"""
    <div class="info-box">
    Agrégat <b>Site × Rayon × Famille</b> — comparaison CA/Marge fidélité vs global · Poids du programme
    </div>
    """, unsafe_allow_html=True)

    # All ventes for ratio global (same filters, no fidelite filter)
    df_global = df_mois.copy()

    # Aggregate global by Site × Rayon × Famille
    grp_global = df_global.groupby(['Site nom long', 'Rayon', 'Famille'], as_index=False).agg(
        CA_Globale=('CA', 'sum'),
        Marge_Globale=('Marge', 'sum'),
    )

    # Aggregate fidélité only
    grp_fid = df_fid_only.groupby(['Site nom long', 'Rayon', 'Famille'], as_index=False).agg(
        CA_Fidelite=('CA', 'sum'),
        Marge_Fidelite=('Marge', 'sum'),
        Qte_Vente=('Qté Vente', 'sum'),
        Total_Cagnotte=('Total Cagnotte', 'sum'),
    )

    df_recap = grp_fid.merge(grp_global, on=['Site nom long', 'Rayon', 'Famille'], how='left')

    # Compute poids
    df_recap['Poids Fidélité %'] = np.where(
        df_recap['CA_Globale'] > 0,
        (df_recap['CA_Fidelite'] / df_recap['CA_Globale'] * 100).round(1),
        np.nan
    )
    df_recap['Poids Marge Fidélité %'] = np.where(
        df_recap['Marge_Globale'].abs() > 0,
        (df_recap['Marge_Fidelite'] / df_recap['Marge_Globale'] * 100).round(1),
        np.nan
    )

    # TOTAL row
    total_row = {
        'Site nom long': 'TOTAL',
        'Rayon': '—',
        'Famille': '—',
        'CA_Fidelite': df_recap['CA_Fidelite'].sum(),
        'CA_Globale': df_recap['CA_Globale'].sum(),
        'Poids Fidélité %': (df_recap['CA_Fidelite'].sum() / df_recap['CA_Globale'].sum() * 100) if df_recap['CA_Globale'].sum() > 0 else np.nan,
        'Marge_Fidelite': df_recap['Marge_Fidelite'].sum(),
        'Marge_Globale': df_recap['Marge_Globale'].sum(),
        'Poids Marge Fidélité %': (df_recap['Marge_Fidelite'].sum() / df_recap['Marge_Globale'].sum() * 100) if df_recap['Marge_Globale'].sum() != 0 else np.nan,
        'Qte_Vente': df_recap['Qte_Vente'].sum(),
        'Total_Cagnotte': df_recap['Total_Cagnotte'].sum(),
    }
    df_total = pd.DataFrame([total_row])
    df_recap_display = pd.concat([df_recap, df_total], ignore_index=True)

    # Format for display
    def fmt_recap(df):
        d = df.copy()
        for col in ['CA_Fidelite', 'CA_Globale', 'Marge_Fidelite', 'Marge_Globale', 'Total_Cagnotte']:
            d[col] = d[col].apply(lambda x: fmt_num(x) if not pd.isna(x) else '—')
        d['Qte_Vente'] = d['Qte_Vente'].apply(lambda x: f"{x:,.1f}".replace(',', ' ') if not pd.isna(x) else '—')
        d['Poids Fidélité %'] = d['Poids Fidélité %'].apply(fmt_pct)
        d['Poids Marge Fidélité %'] = d['Poids Marge Fidélité %'].apply(fmt_pct)
        return d

    df_display = fmt_recap(df_recap_display)
    df_display = df_display.rename(columns={
        'Site nom long': 'Site',
        'CA_Fidelite': 'CA Fidélité',
        'CA_Globale': 'CA Globale',
        'Marge_Fidelite': 'Marge Fidélité',
        'Marge_Globale': 'Marge Globale',
        'Qte_Vente': 'Qté Vente',
        'Total_Cagnotte': 'Total Cagnotte',
    })

    # Style
    def style_recap(row):
        if row['Site'] == 'TOTAL':
            return ['background-color: #EAF4FF; font-weight: 700; color: #007AFF;'] * len(row)
        return [''] * len(row)

    styled = df_display.style.apply(style_recap, axis=1)
    for col in ['Marge Fidélité', 'Marge Globale']:
        styled = styled.map(color_marge, subset=[col])
    for col in ['Poids Fidélité %', 'Poids Marge Fidélité %']:
        styled = styled.map(color_poids, subset=[col])

    st.dataframe(styled, use_container_width=True, height=500, hide_index=True)


# ═══════════════════════════════════════════
# ONGLET 2 — RÉCAP DÉTAIL Article × Site
# ═══════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-header">Récap Détail · Article × Magasin · ' + mois_court + '</div>', unsafe_allow_html=True)
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.markdown("""
    <div class="info-box">
    Vue Article × Site avec totaux par article — articles en programme fidélité uniquement
    </div>
    """, unsafe_allow_html=True)

    # Build article × site table
    grp_art_site = df_fid_only.groupby(
        ['Date Début', 'Date Fin', 'Semaine', 'Mois', 'Rayon', 'Famille', 'Article', '_article_id', 'Site nom long'],
        as_index=False
    ).agg(
        CA=('CA', 'sum'),
        Marge=('Marge', 'sum'),
        Qte=('Qté Vente', 'sum'),
        Cagnotte_unit=('Cagnotte_unit', 'first'),
        Total_Cagnotte=('Total Cagnotte', 'sum'),
    )

    # Build totals per article
    grp_art_total = df_fid_only.groupby(
        ['Date Début', 'Date Fin', 'Semaine', 'Mois', 'Rayon', 'Famille', 'Article', '_article_id'],
        as_index=False
    ).agg(
        CA=('CA', 'sum'),
        Marge=('Marge', 'sum'),
        Qte=('Qté Vente', 'sum'),
        Cagnotte_unit=('Cagnotte_unit', 'first'),
        Total_Cagnotte=('Total Cagnotte', 'sum'),
    )
    grp_art_total['Site nom long'] = 'TOTAL'
    nb_sites = df_fid_only.groupby('_article_id')['Site nom long'].nunique().reset_index(name='nb_sites')
    grp_art_total = grp_art_total.merge(nb_sites, on='_article_id', how='left')
    grp_art_total['Site nom long'] = grp_art_total.apply(
        lambda r: f"TOTAL · {int(r['nb_sites'])} mag." if not pd.isna(r['nb_sites']) else 'TOTAL', axis=1
    )

    # Interleave: for each article, site rows then total
    rows = []
    for art_id in grp_art_site['_article_id'].unique():
        sub = grp_art_site[grp_art_site['_article_id'] == art_id].copy()
        sub['_is_total'] = False
        tot = grp_art_total[grp_art_total['_article_id'] == art_id].copy()
        tot['_is_total'] = True
        rows.append(sub)
        rows.append(tot)

    if rows:
        df_detail = pd.concat(rows, ignore_index=True)
    else:
        df_detail = pd.DataFrame()

    if len(df_detail) > 0:
        # Format
        df_det_disp = df_detail[[
            'Date Début','Date Fin','Semaine','Mois','Rayon','Famille','Article',
            'Site nom long','CA','Marge','Qte','Cagnotte_unit','Total_Cagnotte','_is_total'
        ]].copy()
        df_det_disp = df_det_disp.rename(columns={
            'Site nom long': 'Site / Magasin',
            'Qte': 'Qté Vente',
            'Cagnotte_unit': 'Cagnotte/unité',
            'Total_Cagnotte': 'Total Cagnotte',
        })
        for col in ['CA','Marge','Total Cagnotte']:
            df_det_disp[col] = df_det_disp[col].apply(lambda x: fmt_num(x) if not pd.isna(x) else '—')
        df_det_disp['Qté Vente'] = df_det_disp['Qté Vente'].apply(lambda x: f"{x:,.1f}".replace(',', ' ') if not pd.isna(x) else '—')
        df_det_disp['Cagnotte/unité'] = df_det_disp['Cagnotte/unité'].apply(lambda x: fmt_num(x) if not pd.isna(x) else '—')

        _is_total = df_det_disp.pop('_is_total')

        def style_detail(row):
            idx = row.name
            if _is_total.iloc[idx]:
                return ['background-color: #EAF4FF; font-weight: 700; color: #007AFF;'] * len(row)
            return [''] * len(row)

        styled2 = df_det_disp.style.apply(style_detail, axis=1)
        styled2 = styled2.map(color_marge, subset=['Marge'])
        st.dataframe(styled2, use_container_width=True, height=600, hide_index=True)
    else:
        st.info("Aucune donnée fidélité pour le mois et les filtres sélectionnés.")


# ═══════════════════════════════════════════
# ONGLET 3 — DRILL-DOWN GRANULAIRE
# ═══════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-header">Drill-down Granulaire · Toutes lignes · ' + mois_court + '</div>', unsafe_allow_html=True)
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        rayons_dd = sorted(df_fid_only['Rayon'].dropna().unique().tolist())
        rayon_dd = st.selectbox("Rayon", ['Tous'] + rayons_dd, key='dd_rayon')
    with col_f2:
        if rayon_dd != 'Tous':
            fam_list = sorted(df_fid_only[df_fid_only['Rayon'] == rayon_dd]['Famille'].dropna().unique().tolist())
        else:
            fam_list = sorted(df_fid_only['Famille'].dropna().unique().tolist())
        famille_dd = st.selectbox("Famille", ['Toutes'] + fam_list, key='dd_famille')
    with col_f3:
        sites_dd = sorted(df_fid_only['Site nom long'].dropna().unique().tolist())
        site_dd = st.selectbox("Magasin", ['Tous'] + sites_dd, key='dd_site')

    df_drill = df_fid_only.copy()
    if rayon_dd != 'Tous':
        df_drill = df_drill[df_drill['Rayon'] == rayon_dd]
    if famille_dd != 'Toutes':
        df_drill = df_drill[df_drill['Famille'] == famille_dd]
    if site_dd != 'Tous':
        df_drill = df_drill[df_drill['Site nom long'] == site_dd]

    st.markdown(f"""
    <div class="info-box">
    {len(df_drill):,} lignes · Articles en programme fidélité uniquement · 
    CA total : <b>{fmt_num(df_drill['CA'].sum())} XOF</b> · 
    Cagnotte totale : <b>{fmt_num(df_drill['Total Cagnotte'].sum())} XOF</b>
    </div>
    """, unsafe_allow_html=True)

    df_drill_disp = df_drill[[
        'Date Début','Date Fin','Semaine','Mois','Site nom long','Rayon','Famille',
        'Article','CA','Marge','Qté Vente','Cagnotte_unit','Total Cagnotte'
    ]].copy().rename(columns={
        'Site nom long': 'Site',
        'Cagnotte_unit': 'Cagnotte/unité',
    })

    for col in ['CA','Marge','Total Cagnotte']:
        df_drill_disp[col] = df_drill_disp[col].apply(lambda x: fmt_num(x) if not pd.isna(x) else '—')
    df_drill_disp['Qté Vente'] = df_drill_disp['Qté Vente'].apply(lambda x: f"{x:,.1f}".replace(',', ' ') if not pd.isna(x) else '—')
    df_drill_disp['Cagnotte/unité'] = df_drill_disp['Cagnotte/unité'].apply(lambda x: fmt_num(x) if not pd.isna(x) else '—')

    styled3 = df_drill_disp.style.map(color_marge, subset=['Marge'])
    st.dataframe(styled3, use_container_width=True, height=600, hide_index=True)
