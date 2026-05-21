import streamlit as st
import pandas as pd
import numpy as np
import re
from io import StringIO, BytesIO

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Fidélité Cagnotte · SmartBuyer Hub",
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
.stApp { background-color: #F2F2F7; }
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1400px; }
[data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E5E5EA; }

.kpi-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 20px; }
.kpi-card {
    background: #FFFFFF; border-radius: 12px; padding: 16px 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08); border: 1px solid #E5E5EA;
}
.kpi-label { font-size: 11px; font-weight: 500; color: #8E8E93; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
.kpi-value { font-size: 21px; font-weight: 700; color: #1C1C1E; line-height: 1.1; }
.kpi-value.blue { color: #007AFF; }
.kpi-value.red { color: #FF3B30; }
.kpi-value.green { color: #34C759; }
.kpi-sub { font-size: 11px; color: #8E8E93; margin-top: 4px; }

.period-badge {
    background: #EAF4FF; border: 1px solid #B8D9FF; border-radius: 10px;
    padding: 12px 20px; margin-bottom: 16px; display: flex; align-items: center; gap: 16px;
}
.period-label { font-size: 11px; font-weight: 600; color: #007AFF; text-transform: uppercase; letter-spacing: 0.5px; }
.period-value { font-size: 15px; font-weight: 700; color: #007AFF; }

.section-header {
    font-size: 13px; font-weight: 600; color: #3A3A3C;
    text-transform: uppercase; letter-spacing: 0.5px;
    margin-bottom: 12px; margin-top: 4px; padding-bottom: 6px;
    border-bottom: 2px solid #007AFF; display: inline-block;
}

.alert-card {
    background: #FFFFFF; border-radius: 10px; padding: 14px 18px;
    border: 1px solid #E5E5EA; margin-bottom: 8px;
}
.alert-title { font-size: 12px; font-weight: 600; color: #1C1C1E; margin-bottom: 4px; }
.alert-badge-red { display:inline-block; background:#FFF0EE; color:#C0392B; border-radius:6px; padding:2px 8px; font-size:11px; font-weight:600; }
.alert-badge-orange { display:inline-block; background:#FFF9E6; color:#7D5A00; border-radius:6px; padding:2px 8px; font-size:11px; font-weight:600; }

.info-box {
    background: #F2F2F7; border-left: 3px solid #007AFF; border-radius: 4px;
    padding: 10px 14px; font-size: 12px; color: #3A3A3C; margin-bottom: 12px;
}
.stDataFrame { border-radius: 10px; overflow: hidden; }
[data-testid="stTabs"] [role="tab"] { font-size: 13px; font-weight: 500; padding: 8px 16px; }
[data-testid="stTabs"] [role="tab"][aria-selected="true"] { color: #007AFF; border-bottom: 2px solid #007AFF; }
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def parse_number(val):
    if pd.isna(val) or str(val).strip() in ['', 'NaN', 'nan']:
        return np.nan
    s = str(val).replace('\xa0', '').replace(' ', '').replace(',', '.')
    try:
        return float(s)
    except:
        return np.nan

def fmt_xof(val):
    if pd.isna(val): return '—'
    return f"{int(round(val)):,}".replace(',', ' ') + ' XOF'

def fmt_num(val):
    if pd.isna(val): return '—'
    return f"{val:,.0f}".replace(',', ' ')

def fmt_pct(val):
    if pd.isna(val): return '—'
    return f"{val:.1f}%"

def extract_article_id(s):
    if pd.isna(s): return None
    m = re.match(r'^(\d+)', str(s).strip())
    return int(m.group(1)) if m else None

def parse_period_from_lines(lines):
    for line in lines:
        m = re.search(r'après le (\d{2}/\d{2}/\d{4}).*?avant le (\d{2}/\d{2}/\d{4})', line)
        if m:
            return pd.to_datetime(m.group(1), dayfirst=True), pd.to_datetime(m.group(2), dayfirst=True)
    return None, None

def get_semaine_mois(d1):
    if d1 is None: return '—', '—'
    mois_map = {1:'Janvier',2:'Février',3:'Mars',4:'Avril',5:'Mai',6:'Juin',
                7:'Juillet',8:'Août',9:'Septembre',10:'Octobre',11:'Novembre',12:'Décembre'}
    return f"S{d1.isocalendar()[1]}", f"{mois_map[d1.month]} {d1.year}"

def load_ventes_csv(file_obj):
    content = file_obj.read().decode('latin1')
    lines = content.split('\n')
    d1, d2 = parse_period_from_lines(lines)
    sem, mois = get_semaine_mois(d1)
    data_lines = []
    header_found = False
    for line in lines:
        s = line.strip()
        if not s: continue
        if 'Filtres appliqués' in s: break
        if 'Site nom long' in s: header_found = True
        if header_found: data_lines.append(s)
    if not data_lines:
        return None, d1, d2, sem, mois
    df = pd.read_csv(StringIO('\n'.join(data_lines)), sep=';', encoding='utf-8', on_bad_lines='skip')
    df.columns = [c.strip() for c in df.columns]
    for col in ['CA', 'Marge', 'Qté Vente']:
        if col in df.columns:
            df[col] = df[col].apply(parse_number)
    df = df[
        df['Article'].notna() &
        ~df['Article'].astype(str).str.strip().isin(['Total', 'NaN', '']) &
        df['Site nom long'].notna() &
        ~df['Site nom long'].astype(str).str.strip().isin(['Total', 'NaN', ''])
    ].copy()
    df['Date Début'] = d1.strftime('%d/%m/%Y') if d1 else '—'
    df['Date Fin']   = d2.strftime('%d/%m/%Y') if d2 else '—'
    df['Semaine'] = sem
    df['Mois']    = mois
    df['_article_id'] = df['Article'].apply(extract_article_id)
    return df, d1, d2, sem, mois

def load_fidelite_csv(file_obj):
    content = file_obj.read().decode('latin1')
    df = pd.read_csv(StringIO(content), sep=None, engine='python', encoding='utf-8')
    df.columns = [c.strip() for c in df.columns]
    mois_map = {'mai':'Mai','avril':'Avril','mars':'Mars','juin':'Juin','janvier':'Janvier',
                'février':'Février','juillet':'Juillet','août':'Août','septembre':'Septembre',
                'octobre':'Octobre','novembre':'Novembre','décembre':'Décembre'}
    if 'Mois' in df.columns:
        df['Mois'] = df['Mois'].astype(str).str.strip().str.lower().map(mois_map).fillna(df['Mois'])
    df = df.dropna(subset=['Article','Cagnotte'])
    df['Article']  = df['Article'].astype(int)
    df['Cagnotte'] = pd.to_numeric(df['Cagnotte'], errors='coerce')
    return df

def to_excel(df_dict):
    """Export dict of {sheet_name: dataframe} to Excel bytes."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for sheet, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
    return buf.getvalue()

def color_marge(val):
    try:
        n = float(str(val).replace(' ','').replace('XOF','').replace(',','.'))
        if n < 0: return 'color:#FF3B30;font-weight:600'
        if n > 0: return 'color:#34C759;font-weight:600'
    except: pass
    return ''

def color_poids(val):
    try:
        n = float(str(val).replace('%','').replace(',','.'))
        if n >= 30: return 'background-color:#E8FAF0;color:#1A7F3C;font-weight:600'
        if n >= 15: return 'background-color:#FFF9E6;color:#7D5A00;font-weight:600'
        return 'background-color:#FFF0EE;color:#C0392B;font-weight:600'
    except: pass
    return ''


# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
c1, c2 = st.columns([1, 11])
with c1:
    st.markdown("""<div style="width:44px;height:44px;background:#007AFF;border-radius:10px;
    display:flex;align-items:center;justify-content:center;margin-top:4px;">
    <span style="font-size:22px;">🏷️</span></div>""", unsafe_allow_html=True)
with c2:
    st.markdown("""<div style="padding-top:6px;">
    <span style="font-size:18px;font-weight:700;color:#1C1C1E;">SmartBuyer Hub</span>
    <span style="font-size:14px;color:#8E8E93;margin-left:10px;">· Fidélité Cagnotte · Investissement vs Performance</span>
    </div>""", unsafe_allow_html=True)
st.markdown("<hr style='margin:8px 0 16px 0;border:none;border-top:1px solid #E5E5EA;'>", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📂 Chargement des données")
    st.markdown("**Fichiers ventes** *(multi-upload)*")
    ventes_files = st.file_uploader("ventes", type=['csv'], accept_multiple_files=True,
                                     key="ventes_upload", label_visibility="collapsed")
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    st.markdown("**Liste Fidélité** *(référentiel)*")
    fidelite_file = st.file_uploader("fidelite", type=['csv'], accept_multiple_files=False,
                                      key="fidelite_upload", label_visibility="collapsed")
    st.markdown("---")
    st.markdown("""<div style='font-size:11px;color:#8E8E93;line-height:1.8;'>
    <div style='font-size:12px;font-weight:600;color:#3A3A3C;margin-bottom:6px;'>📋 Format attendu</div>
    <b>Ventes :</b> export Power BI (séparateur <code>;</code>)<br>
    <b>Fidélité :</b> Article / Cagnotte / Mois<br>
    <b>Multi-semaines :</b> fichiers empilés automatiquement
    <div style='margin-top:10px;padding:8px;background:#F2F2F7;border-radius:6px;font-size:10px;color:#636366;'>
    💡 Période détectée automatiquement depuis les métadonnées PBI
    </div></div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# LANDING PAGE
# ─────────────────────────────────────────────
if not ventes_files or not fidelite_file:
    st.markdown("""
    <div style='background:#FFFFFF;border-radius:16px;padding:40px 48px;
    border:1px solid #E5E5EA;box-shadow:0 1px 3px rgba(0,0,0,0.06);margin-bottom:20px;'>
        <div style='display:flex;align-items:flex-start;gap:24px;'>
            <div style='font-size:48px;line-height:1;margin-top:4px;'>🏷️</div>
            <div style='flex:1;'>
                <div style='font-size:22px;font-weight:700;color:#1C1C1E;margin-bottom:6px;'>Fidélité Cagnotte</div>
                <div style='font-size:14px;color:#636366;line-height:1.6;max-width:680px;'>
                    Pilotez la performance de votre programme de fidélité : mesurez l'investissement cagnotte,
                    le poids des articles en programme sur votre CA et votre marge, et identifiez les magasins
                    et familles où le programme génère de la valeur réelle.
                </div>
            </div>
        </div>
    </div>""", unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    cards = [
        ("🏠", "Synthèse Exécutive", "Vue COPIL : top familles, performance réseau, alertes articles sans ventes et familles en marge négative."),
        ("📊", "Récap Financier", "Agrégat Site × Rayon × Famille — CA et Marge fidélité vs global, poids du programme, budget cagnotte."),
        ("🔍", "Récap Détail", "Vue Article × Magasin — cagnotte unitaire, total cagnotte, CA et marge par site."),
        ("📋", "Drill-down", "Toutes les lignes article × site × période, filtrable par Rayon, Famille et Magasin."),
    ]
    for col, (ico, title, desc) in zip([c1, c2, c3, c4], cards):
        with col:
            st.markdown(f"""<div style='background:#FFFFFF;border-radius:12px;padding:20px 24px;
            border:1px solid #E5E5EA;box-shadow:0 1px 3px rgba(0,0,0,0.06);min-height:170px;'>
                <div style='font-size:24px;margin-bottom:10px;'>{ico}</div>
                <div style='font-size:14px;font-weight:600;color:#1C1C1E;margin-bottom:6px;'>{title}</div>
                <div style='font-size:12px;color:#636366;line-height:1.6;'>{desc}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    st.markdown("""<div style='background:#EAF4FF;border-radius:12px;padding:20px 24px;border:1px solid #B8D9FF;'>
        <div style='font-size:13px;font-weight:600;color:#007AFF;margin-bottom:10px;'>🚀 Pour démarrer</div>
        <div style='display:grid;grid-template-columns:1fr 1fr;gap:16px;'>
            <div style='font-size:12px;color:#3A3A3C;line-height:1.7;'>
                <b>① Fichiers ventes</b> (panneau gauche)<br>
                Chargez une ou plusieurs extractions Power BI CSV (séparateur <code>;</code>).
                Plusieurs semaines peuvent être chargées simultanément.
            </div>
            <div style='font-size:12px;color:#3A3A3C;line-height:1.7;'>
                <b>② Liste Fidélité</b> (panneau gauche)<br>
                Chargez le référentiel des articles en programme avec les colonnes
                <code>Article</code> / <code>Cagnotte</code> / <code>Mois</code>.
            </div>
        </div>
    </div>""", unsafe_allow_html=True)
    st.stop()


# ─────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────
all_dfs, periods = [], []
for f in ventes_files:
    df_v, d1, d2, sem, mois = load_ventes_csv(f)
    if df_v is not None and len(df_v) > 0:
        all_dfs.append(df_v)
        periods.append({'file': f.name, 'date_debut': d1, 'date_fin': d2, 'sem': sem, 'mois': mois})

if not all_dfs:
    st.error("Aucune donnée valide trouvée dans les fichiers ventes.")
    st.stop()

df_ventes  = pd.concat(all_dfs, ignore_index=True)
df_fidelite = load_fidelite_csv(fidelite_file)


# ─────────────────────────────────────────────
# GLOBAL FILTERS (top bar)
# ─────────────────────────────────────────────
mois_disponibles = sorted(df_ventes['Mois'].dropna().unique().tolist())
cf1, cf2, cf3 = st.columns([2, 2, 4])
with cf1:
    mois_sel = st.selectbox("Mois actif", options=mois_disponibles, index=0)
with cf2:
    rayon_sel = st.selectbox("Rayon", options=['Tous'] + sorted(df_ventes['Rayon'].dropna().unique().tolist()))
with cf3:
    sites_sel = st.multiselect("Magasins", options=sorted(df_ventes['Site nom long'].dropna().unique().tolist()), default=[])


# ─────────────────────────────────────────────
# FILTER & JOIN
# ─────────────────────────────────────────────
mois_court = mois_sel.split(' ')[0]
df_mois = df_ventes[df_ventes['Mois'] == mois_sel].copy()

df_fid_mois = df_fidelite[df_fidelite['Mois'] == mois_court]
if len(df_fid_mois) == 0:
    df_fid_mois = df_fidelite.copy()

if sites_sel:
    df_mois = df_mois[df_mois['Site nom long'].isin(sites_sel)]
if rayon_sel != 'Tous':
    df_mois = df_mois[df_mois['Rayon'] == rayon_sel]

df_joined = df_mois.merge(
    df_fid_mois[['Article','Cagnotte']].rename(columns={'Article':'_article_id','Cagnotte':'Cagnotte_unit'}),
    on='_article_id', how='left'
)
df_joined['est_fidelite'] = df_joined['Cagnotte_unit'].notna()
df_joined['Total Cagnotte'] = df_joined['Cagnotte_unit'] * df_joined['Qté Vente']
df_fid_only = df_joined[df_joined['est_fidelite']].copy()


# ─────────────────────────────────────────────
# PERIOD BADGE
# ─────────────────────────────────────────────
pi = next((p for p in periods if mois_sel in p['mois']), periods[0] if periods else None)
d1_str = pi['date_debut'].strftime('%d/%m/%Y') if pi and pi['date_debut'] else '—'
d2_str = pi['date_fin'].strftime('%d/%m/%Y')   if pi and pi['date_fin']   else '—'
sem_str  = pi['sem']  if pi else '—'
mois_str = pi['mois'] if pi else '—'

st.markdown(f"""<div class="period-badge">
    <div>
        <div class="period-label">Période détectée</div>
        <div class="period-value">{d1_str} → {d2_str} · {sem_str} · {mois_str}</div>
    </div>
    <div style="margin-left:auto;font-size:12px;color:#3A3A3C;">
        {len(ventes_files)} fichier(s) chargé(s) · {len(df_fid_mois)} articles fidélité ({mois_court})
    </div>
</div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# KPI COMPUTATION
# ─────────────────────────────────────────────
budget_cagnotte = df_fid_only['Total Cagnotte'].sum()
ca_fidelite     = df_fid_only['CA'].sum()
marge_fidelite  = df_fid_only['Marge'].sum()
arts_actifs     = df_fid_only[df_fid_only['CA'] > 0]['_article_id'].nunique()
arts_perimetre  = df_fid_mois['Article'].nunique()
sites_actifs    = df_fid_only[df_fid_only['CA'] > 0]['Site nom long'].nunique()
sites_total     = df_mois['Site nom long'].nunique()
marge_class     = "red" if marge_fidelite < 0 else "green"

st.markdown(f"""<div class="kpi-grid">
    <div class="kpi-card">
        <div class="kpi-label">Budget Cagnotte</div>
        <div class="kpi-value blue">{fmt_xof(budget_cagnotte)}</div>
        <div class="kpi-sub">Investissement fidélité</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">CA Fidélité</div>
        <div class="kpi-value">{fmt_xof(ca_fidelite)}</div>
        <div class="kpi-sub">Articles en programme</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">Marge Fidélité</div>
        <div class="kpi-value {marge_class}">{fmt_xof(marge_fidelite)}</div>
        <div class="kpi-sub">Impact marge programme</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">Articles Actifs</div>
        <div class="kpi-value blue">{arts_actifs} <span style="font-size:14px;color:#8E8E93;">/ {arts_perimetre}</span></div>
        <div class="kpi-sub">Avec ventes &gt; 0 · périmètre {mois_court}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">Couverture Réseau</div>
        <div class="kpi-value blue">{sites_actifs} <span style="font-size:14px;color:#8E8E93;">/ {sites_total}</span></div>
        <div class="kpi-sub">Sites avec ventes fidélité</div>
    </div>
</div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab0, tab1, tab2, tab3, tab4 = st.tabs([
    "🏠 Synthèse Exécutive",
    "📊 Récap Financier",
    "🔍 Récap Détail Article × Site",
    "📋 Drill-down Granulaire",
    "🗃️ Raw Data",
])


# ═══════════════════════════════════════════════════════
# ONGLET 0 — SYNTHÈSE EXÉCUTIVE
# ═══════════════════════════════════════════════════════
with tab0:
    st.markdown(f'<div class="section-header">Synthèse Exécutive · {mois_court} · Vue COPIL</div>', unsafe_allow_html=True)
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

    # ── Bloc A : Top 5 Familles par budget cagnotte ──────────────
    col_a, col_b = st.columns([1, 1])

    with col_a:
        st.markdown("**🏆 Top 5 Familles · Budget Cagnotte investi**")
        grp_fam = df_fid_only.groupby('Famille', as_index=False).agg(
            CA_Fidelite=('CA', 'sum'),
            Marge_Fidelite=('Marge', 'sum'),
            Qte=('Qté Vente', 'sum'),
            Budget_Cagnotte=('Total Cagnotte', 'sum'),
        ).sort_values('Budget_Cagnotte', ascending=False).head(5)

        # Global CA per famille for poids
        grp_fam_global = df_mois.groupby('Famille', as_index=False).agg(CA_Globale=('CA','sum'))
        grp_fam = grp_fam.merge(grp_fam_global, on='Famille', how='left')
        grp_fam['Poids %'] = np.where(grp_fam['CA_Globale']>0,
            (grp_fam['CA_Fidelite']/grp_fam['CA_Globale']*100).round(1), np.nan)

        df_top5 = grp_fam[['Famille','CA_Fidelite','Marge_Fidelite','Budget_Cagnotte','Poids %']].copy()
        # shorten famille label
        df_top5['Famille'] = df_top5['Famille'].apply(
            lambda x: re.sub(r'^\d+\s*-\s*', '', str(x))[:30])
        df_top5_disp = df_top5.copy()
        df_top5_disp['CA_Fidelite']    = df_top5_disp['CA_Fidelite'].apply(fmt_num)
        df_top5_disp['Marge_Fidelite'] = df_top5_disp['Marge_Fidelite'].apply(fmt_num)
        df_top5_disp['Budget_Cagnotte']= df_top5_disp['Budget_Cagnotte'].apply(fmt_num)
        df_top5_disp['Poids %']        = df_top5_disp['Poids %'].apply(fmt_pct)
        df_top5_disp = df_top5_disp.rename(columns={
            'CA_Fidelite':'CA Fidélité','Marge_Fidelite':'Marge','Budget_Cagnotte':'Cagnotte','Poids %':'Poids CA %'})
        styled_top5 = df_top5_disp.style.map(color_marge, subset=['Marge']).map(color_poids, subset=['Poids CA %'])
        st.dataframe(styled_top5, use_container_width=True, hide_index=True)

    with col_b:
        st.markdown("**🏪 Performance Réseau · Site**")
        grp_site = df_fid_only.groupby('Site nom long', as_index=False).agg(
            CA_Fidelite=('CA','sum'),
            Marge_Fidelite=('Marge','sum'),
            Budget_Cagnotte=('Total Cagnotte','sum'),
            Nb_Articles=('_article_id','nunique'),
        ).sort_values('CA_Fidelite', ascending=False)
        grp_site_global = df_mois.groupby('Site nom long', as_index=False).agg(CA_Globale=('CA','sum'))
        grp_site = grp_site.merge(grp_site_global, on='Site nom long', how='left')
        grp_site['Poids %'] = np.where(grp_site['CA_Globale']>0,
            (grp_site['CA_Fidelite']/grp_site['CA_Globale']*100).round(1), np.nan)

        grp_site['Site'] = grp_site['Site nom long'].apply(
            lambda x: re.sub(r'^\d+\s*-\s*', '', str(x)))
        df_site_disp = grp_site[['Site','CA_Fidelite','Marge_Fidelite','Budget_Cagnotte','Poids %','Nb_Articles']].copy()
        df_site_disp['CA_Fidelite']    = df_site_disp['CA_Fidelite'].apply(fmt_num)
        df_site_disp['Marge_Fidelite'] = df_site_disp['Marge_Fidelite'].apply(fmt_num)
        df_site_disp['Budget_Cagnotte']= df_site_disp['Budget_Cagnotte'].apply(fmt_num)
        df_site_disp['Poids %']        = df_site_disp['Poids %'].apply(fmt_pct)
        df_site_disp = df_site_disp.rename(columns={
            'CA_Fidelite':'CA Fidélité','Marge_Fidelite':'Marge','Budget_Cagnotte':'Cagnotte',
            'Poids %':'Poids CA %','Nb_Articles':'Nb Art.'})
        styled_site = df_site_disp.style.map(color_marge, subset=['Marge']).map(color_poids, subset=['Poids CA %'])
        st.dataframe(styled_site, use_container_width=True, hide_index=True)

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ── Bloc B : Alertes ─────────────────────────────────────────
    col_c, col_d = st.columns([1, 1])

    with col_c:
        # Articles sans ventes (0 ROI)
        arts_zero = df_fid_only[df_fid_only['CA'].isna() | (df_fid_only['CA'] == 0)]['_article_id'].unique()
        # Distinct article names
        arts_zero_names = (df_fid_only[df_fid_only['_article_id'].isin(arts_zero)]
                           [['_article_id','Article']].drop_duplicates()
                           .sort_values('Article'))
        n_zero = len(arts_zero_names)
        st.markdown(f"""
        <div class="alert-card">
            <div class="alert-title">⚠️ Articles sans ventes · Cagnotte à ROI nul</div>
            <div style='margin-bottom:8px;'>
                <span class="alert-badge-red">{n_zero} article(s)</span>
                <span style='font-size:11px;color:#8E8E93;margin-left:8px;'>présents en programme mais 0 vente sur la période</span>
            </div>
        </div>""", unsafe_allow_html=True)
        if n_zero > 0:
            arts_zero_disp = arts_zero_names[['Article']].copy()
            arts_zero_disp['Article'] = arts_zero_disp['Article'].apply(
                lambda x: re.sub(r'^\d+\s*-\s*', '', str(x))[:50])
            st.dataframe(arts_zero_disp, use_container_width=True, hide_index=True, height=220)
        else:
            st.success("✅ Tous les articles en programme ont généré des ventes.")

    with col_d:
        # Familles marge fidélité négative
        grp_fam_marge = df_fid_only.groupby('Famille', as_index=False).agg(
            Marge_Fidelite=('Marge','sum'),
            CA_Fidelite=('CA','sum'),
            Budget_Cagnotte=('Total Cagnotte','sum'),
        )
        fam_neg = grp_fam_marge[grp_fam_marge['Marge_Fidelite'] < 0].sort_values('Marge_Fidelite')
        n_neg = len(fam_neg)
        st.markdown(f"""
        <div class="alert-card">
            <div class="alert-title">🔴 Familles · Marge Fidélité négative</div>
            <div style='margin-bottom:8px;'>
                <span class="alert-badge-orange">{n_neg} famille(s)</span>
                <span style='font-size:11px;color:#8E8E93;margin-left:8px;'>où le programme dégrade la marge</span>
            </div>
        </div>""", unsafe_allow_html=True)
        if n_neg > 0:
            fam_neg_disp = fam_neg.copy()
            fam_neg_disp['Famille'] = fam_neg_disp['Famille'].apply(
                lambda x: re.sub(r'^\d+\s*-\s*', '', str(x))[:35])
            fam_neg_disp['Marge_Fidelite']  = fam_neg_disp['Marge_Fidelite'].apply(fmt_num)
            fam_neg_disp['CA_Fidelite']     = fam_neg_disp['CA_Fidelite'].apply(fmt_num)
            fam_neg_disp['Budget_Cagnotte'] = fam_neg_disp['Budget_Cagnotte'].apply(fmt_num)
            fam_neg_disp = fam_neg_disp.rename(columns={
                'Marge_Fidelite':'Marge','CA_Fidelite':'CA Fidélité','Budget_Cagnotte':'Cagnotte'})
            styled_neg = fam_neg_disp.style.map(color_marge, subset=['Marge'])
            st.dataframe(styled_neg, use_container_width=True, hide_index=True, height=220)
        else:
            st.success("✅ Aucune famille en marge négative.")

    # ── Export Synthèse ─────────────────────────────────────────
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    synth_export = {
        'Top5 Familles': grp_fam[['Famille','CA_Fidelite','Marge_Fidelite','Budget_Cagnotte','Poids %']].rename(columns={
            'CA_Fidelite':'CA Fidélité','Marge_Fidelite':'Marge Fidélité','Budget_Cagnotte':'Budget Cagnotte'}),
        'Performance Réseau': grp_site[['Site nom long','CA_Fidelite','Marge_Fidelite','Budget_Cagnotte','Poids %','Nb_Articles']].rename(columns={
            'CA_Fidelite':'CA Fidélité','Marge_Fidelite':'Marge Fidélité','Budget_Cagnotte':'Budget Cagnotte','Nb_Articles':'Nb Articles'}),
        'Alertes Articles 0 vente': arts_zero_names,
        'Alertes Marge Négative': fam_neg[['Famille','CA_Fidelite','Marge_Fidelite','Budget_Cagnotte']].rename(columns={
            'CA_Fidelite':'CA Fidélité','Marge_Fidelite':'Marge Fidélité','Budget_Cagnotte':'Budget Cagnotte'}),
    }
    st.download_button(
        label="⬇️ Exporter Synthèse Excel",
        data=to_excel(synth_export),
        file_name=f"Synthese_Fidelite_{mois_court}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ═══════════════════════════════════════════════════════
# ONGLET 1 — RÉCAP FINANCIER (Mois · Site × Rayon × Famille)
# ═══════════════════════════════════════════════════════
with tab1:
    st.markdown(f'<div class="section-header">Récap Financier · {mois_court} · Site × Rayon × Famille</div>', unsafe_allow_html=True)
    st.markdown("""<div class="info-box">
    Agrégat <b>Site × Rayon × Famille</b> — CA / Marge fidélité vs global · Poids du programme · Budget cagnotte
    </div>""", unsafe_allow_html=True)

    # Global aggregation
    grp_global = df_mois.groupby(['Site nom long','Rayon','Famille'], as_index=False).agg(
        CA_Globale=('CA','sum'), Marge_Globale=('Marge','sum'))
    grp_fid = df_fid_only.groupby(['Site nom long','Rayon','Famille'], as_index=False).agg(
        CA_Fidelite=('CA','sum'), Marge_Fidelite=('Marge','sum'),
        Qte_Vente=('Qté Vente','sum'), Total_Cagnotte=('Total Cagnotte','sum'))

    df_recap = grp_fid.merge(grp_global, on=['Site nom long','Rayon','Famille'], how='left')
    df_recap['Poids Fidélité %'] = np.where(df_recap['CA_Globale']>0,
        (df_recap['CA_Fidelite']/df_recap['CA_Globale']*100).round(1), np.nan)
    df_recap['Poids Marge %'] = np.where(df_recap['Marge_Globale'].abs()>0,
        (df_recap['Marge_Fidelite']/df_recap['Marge_Globale']*100).round(1), np.nan)

    # Add Mois column first
    df_recap.insert(0, 'Mois', mois_court)

    # Format display (no total row)
    df_recap_disp = df_recap.copy()
    for col in ['CA_Fidelite','CA_Globale','Marge_Fidelite','Marge_Globale','Total_Cagnotte']:
        df_recap_disp[col] = df_recap_disp[col].apply(fmt_num)
    df_recap_disp['Qte_Vente']      = df_recap_disp['Qte_Vente'].apply(lambda x: f"{x:,.1f}".replace(',', ' ') if not pd.isna(x) else '—')
    df_recap_disp['Poids Fidélité %'] = df_recap_disp['Poids Fidélité %'].apply(fmt_pct)
    df_recap_disp['Poids Marge %']    = df_recap_disp['Poids Marge %'].apply(fmt_pct)
    df_recap_disp = df_recap_disp.rename(columns={
        'Site nom long':'Site','CA_Fidelite':'CA Fidélité','CA_Globale':'CA Globale',
        'Marge_Fidelite':'Marge Fidélité','Marge_Globale':'Marge Globale',
        'Qte_Vente':'Qté Vente','Total_Cagnotte':'Total Cagnotte',
        'Poids Marge %':'Poids Marge Fidélité %'})

    styled1 = df_recap_disp.style\
        .map(color_marge, subset=['Marge Fidélité','Marge Globale'])\
        .map(color_poids, subset=['Poids Fidélité %','Poids Marge Fidélité %'])
    st.dataframe(styled1, use_container_width=True, height=500, hide_index=True)

    # Export
    export1 = df_recap.rename(columns={
        'Site nom long':'Site','CA_Fidelite':'CA Fidélité','CA_Globale':'CA Globale',
        'Marge_Fidelite':'Marge Fidélité','Marge_Globale':'Marge Globale',
        'Qte_Vente':'Qté Vente','Total_Cagnotte':'Total Cagnotte',
        'Poids Marge %':'Poids Marge Fidélité %'})
    st.download_button(
        label="⬇️ Exporter Récap Financier Excel",
        data=to_excel({'Récap Financier': export1}),
        file_name=f"Recap_Financier_Fidelite_{mois_court}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ═══════════════════════════════════════════════════════
# ONGLET 2 — RÉCAP DÉTAIL Article × Site
# ═══════════════════════════════════════════════════════
with tab2:
    st.markdown(f'<div class="section-header">Récap Détail · Article × Magasin · {mois_court}</div>', unsafe_allow_html=True)
    st.markdown("""<div class="info-box">Articles en programme fidélité uniquement · Agrégat Article × Site</div>""",
                unsafe_allow_html=True)

    grp_art_site = df_fid_only.groupby(
        ['Date Début','Date Fin','Semaine','Mois','Rayon','Famille','Article','_article_id','Site nom long'],
        as_index=False
    ).agg(CA=('CA','sum'), Marge=('Marge','sum'), Qte=('Qté Vente','sum'),
          Cagnotte_unit=('Cagnotte_unit','first'), Total_Cagnotte=('Total Cagnotte','sum'))

    if len(grp_art_site) > 0:
        df_det_disp = grp_art_site[[
            'Date Début','Date Fin','Semaine','Mois','Rayon','Famille','Article',
            'Site nom long','CA','Marge','Qte','Cagnotte_unit','Total_Cagnotte'
        ]].copy().rename(columns={
            'Site nom long':'Site / Magasin','Qte':'Qté Vente',
            'Cagnotte_unit':'Cagnotte/unité','Total_Cagnotte':'Total Cagnotte'})
        for col in ['CA','Marge','Total Cagnotte']:
            df_det_disp[col] = df_det_disp[col].apply(fmt_num)
        df_det_disp['Qté Vente']       = df_det_disp['Qté Vente'].apply(lambda x: f"{x:,.1f}".replace(',', ' ') if not pd.isna(x) else '—')
        df_det_disp['Cagnotte/unité']  = df_det_disp['Cagnotte/unité'].apply(fmt_num)

        styled2 = df_det_disp.style.map(color_marge, subset=['Marge'])
        st.dataframe(styled2, use_container_width=True, height=580, hide_index=True)

        # Export
        export2 = grp_art_site[[
            'Date Début','Date Fin','Semaine','Mois','Rayon','Famille','Article',
            'Site nom long','CA','Marge','Qte','Cagnotte_unit','Total_Cagnotte'
        ]].rename(columns={'Site nom long':'Site','Qte':'Qté Vente',
                            'Cagnotte_unit':'Cagnotte/unité','Total_Cagnotte':'Total Cagnotte'})
        st.download_button(
            label="⬇️ Exporter Récap Détail Excel",
            data=to_excel({'Récap Détail': export2}),
            file_name=f"Recap_Detail_Fidelite_{mois_court}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Aucune donnée fidélité pour les filtres sélectionnés.")


# ═══════════════════════════════════════════════════════
# ONGLET 3 — DRILL-DOWN GRANULAIRE
# ═══════════════════════════════════════════════════════
with tab3:
    st.markdown(f'<div class="section-header">Drill-down Granulaire · {mois_court}</div>', unsafe_allow_html=True)

    cd1, cd2, cd3 = st.columns(3)
    with cd1:
        rayons_dd = sorted(df_fid_only['Rayon'].dropna().unique().tolist())
        rayon_dd  = st.selectbox("Rayon", ['Tous'] + rayons_dd, key='dd_rayon')
    with cd2:
        fam_src = df_fid_only[df_fid_only['Rayon']==rayon_dd] if rayon_dd!='Tous' else df_fid_only
        fam_list = sorted(fam_src['Famille'].dropna().unique().tolist())
        famille_dd = st.selectbox("Famille", ['Toutes'] + fam_list, key='dd_famille')
    with cd3:
        site_dd = st.selectbox("Magasin", ['Tous'] + sorted(df_fid_only['Site nom long'].dropna().unique().tolist()), key='dd_site')

    df_drill = df_fid_only.copy()
    if rayon_dd   != 'Tous':    df_drill = df_drill[df_drill['Rayon']==rayon_dd]
    if famille_dd != 'Toutes':  df_drill = df_drill[df_drill['Famille']==famille_dd]
    if site_dd    != 'Tous':    df_drill = df_drill[df_drill['Site nom long']==site_dd]

    st.markdown(f"""<div class="info-box">
    {len(df_drill):,} lignes · CA : <b>{fmt_num(df_drill['CA'].sum())} XOF</b> · 
    Cagnotte : <b>{fmt_num(df_drill['Total Cagnotte'].sum())} XOF</b>
    </div>""", unsafe_allow_html=True)

    df_drill_disp = df_drill[[
        'Date Début','Date Fin','Semaine','Mois','Site nom long','Rayon','Famille',
        'Article','CA','Marge','Qté Vente','Cagnotte_unit','Total Cagnotte'
    ]].copy().rename(columns={'Site nom long':'Site','Cagnotte_unit':'Cagnotte/unité'})
    for col in ['CA','Marge','Total Cagnotte']:
        df_drill_disp[col] = df_drill_disp[col].apply(fmt_num)
    df_drill_disp['Qté Vente']      = df_drill_disp['Qté Vente'].apply(lambda x: f"{x:,.1f}".replace(',', ' ') if not pd.isna(x) else '—')
    df_drill_disp['Cagnotte/unité'] = df_drill_disp['Cagnotte/unité'].apply(fmt_num)

    styled3 = df_drill_disp.style.map(color_marge, subset=['Marge'])
    st.dataframe(styled3, use_container_width=True, height=580, hide_index=True)

    # Export
    export3 = df_drill[[
        'Date Début','Date Fin','Semaine','Mois','Site nom long','Rayon','Famille',
        'Article','CA','Marge','Qté Vente','Cagnotte_unit','Total Cagnotte'
    ]].rename(columns={'Site nom long':'Site','Cagnotte_unit':'Cagnotte/unité'})
    st.download_button(
        label="⬇️ Exporter Drill-down Excel",
        data=to_excel({'Drill-down': export3}),
        file_name=f"Drilldown_Fidelite_{mois_court}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ═══════════════════════════════════════════════════════
# ONGLET 4 — RAW DATA (consolidation toutes extractions)
# ═══════════════════════════════════════════════════════
with tab4:
    st.markdown(f'<div class="section-header">Raw Data · Consolidation PBI · Toutes périodes</div>', unsafe_allow_html=True)
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

    # Raw = df_ventes complet (toutes semaines chargées, pas de filtre mois)
    # Rejoindre avec fidélité sur _article_id × mois
    # On fait un join global : pour chaque ligne, chercher la cagnotte du mois correspondant
    df_raw = df_ventes.copy()
    df_raw['_mois_court'] = df_raw['Mois'].apply(lambda x: str(x).split(' ')[0] if pd.notna(x) else '')

    # Join fidélité : article_id + mois_court
    df_fid_all = df_fidelite.copy()
    df_fid_all = df_fid_all.rename(columns={'Article': '_article_id', 'Cagnotte': 'Cagnotte/unité', 'Mois': '_mois_fid'})

    df_raw = df_raw.merge(
        df_fid_all[['_article_id', 'Cagnotte/unité', '_mois_fid']],
        left_on=['_article_id', '_mois_court'],
        right_on=['_article_id', '_mois_fid'],
        how='left'
    ).drop(columns=['_mois_fid', '_mois_court', '_article_id'], errors='ignore')

    df_raw['Total Cagnotte'] = df_raw['Cagnotte/unité'] * df_raw['Qté Vente']

    # Restreindre aux articles fidélité uniquement
    df_raw = df_raw[df_raw['Cagnotte/unité'].notna()].copy()

    # Stats
    nb_lignes   = len(df_raw)
    nb_semaines = df_raw['Semaine'].nunique()
    nb_sites    = df_raw['Site nom long'].nunique()
    nb_articles = df_raw['Article'].nunique()
    periodes_str = ' · '.join(sorted(df_raw['Semaine'].dropna().unique().tolist()))

    st.markdown(f"""<div class="info-box">
    Articles <b>programme fidélité uniquement</b> · <b>{nb_lignes:,}</b> lignes · 
    <b>{nb_semaines}</b> semaine(s) : {periodes_str} · <b>{nb_sites}</b> sites · <b>{nb_articles}</b> articles distincts
    </div>""", unsafe_allow_html=True)

    # Filtres rapides
    cr1, cr2 = st.columns(2)
    with cr1:
        sem_list   = sorted(df_raw['Semaine'].dropna().unique().tolist())
        sem_filter = st.multiselect("Semaine(s)", options=sem_list, default=[], key='raw_sem')
    with cr2:
        site_raw = st.selectbox("Magasin", ['Tous'] + sorted(df_raw['Site nom long'].dropna().unique().tolist()), key='raw_site')

    df_raw_f = df_raw.copy()
    if sem_filter:
        df_raw_f = df_raw_f[df_raw_f['Semaine'].isin(sem_filter)]
    if site_raw != 'Tous':
        df_raw_f = df_raw_f[df_raw_f['Site nom long'] == site_raw]

    st.markdown(f"""<div style='font-size:12px;color:#636366;margin-bottom:8px;'>
    {len(df_raw_f):,} lignes affichées</div>""", unsafe_allow_html=True)

    # Display columns
    cols_raw = ['Date Début','Date Fin','Semaine','Mois','Site nom long','Rayon',
                'Famille','Article','CA','Marge','Qté Vente','Cagnotte/unité','Total Cagnotte']
    df_raw_disp = df_raw_f[[c for c in cols_raw if c in df_raw_f.columns]].copy()

    # Format numerics for display
    for col in ['CA','Marge','Total Cagnotte']:
        if col in df_raw_disp.columns:
            df_raw_disp[col] = df_raw_disp[col].apply(fmt_num)
    if 'Qté Vente' in df_raw_disp.columns:
        df_raw_disp['Qté Vente'] = df_raw_disp['Qté Vente'].apply(
            lambda x: f"{x:,.1f}".replace(',', ' ') if not pd.isna(x) else '—')
    if 'Cagnotte/unité' in df_raw_disp.columns:
        df_raw_disp['Cagnotte/unité'] = df_raw_disp['Cagnotte/unité'].apply(fmt_num)

    # Highlight fidélité rows
    def style_raw(row):
        if row.get('Cagnotte/unité', '—') not in ['—', '', None] and str(row.get('Cagnotte/unité', '—')) != '—':
            return ['background-color:#F0F8FF;'] * len(row)
        return [''] * len(row)

    styled_raw = df_raw_disp.style\
        .apply(style_raw, axis=1)\
        .map(color_marge, subset=['Marge'] if 'Marge' in df_raw_disp.columns else [])
    st.dataframe(styled_raw, use_container_width=True, height=600, hide_index=True)

    # Export — raw values (not formatted strings)
    df_raw_export = df_raw_f[[c for c in cols_raw if c in df_raw_f.columns]].copy()
    # Keep numerics as numbers for Excel
    for col in ['CA','Marge','Qté Vente','Cagnotte/unité','Total Cagnotte']:
        if col in df_raw_export.columns:
            df_raw_export[col] = pd.to_numeric(df_raw_export[col], errors='coerce')

    st.download_button(
        label="⬇️ Exporter Raw Data Excel",
        data=to_excel({'Raw Data': df_raw_export}),
        file_name=f"RawData_Fidelite_Consolide.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
