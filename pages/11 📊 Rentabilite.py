import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Rentabilité · SmartBuyer", page_icon="📊", layout="wide")

st.markdown("""
<style>
/* ── Base ── */
body, .stApp { background:#F2F2F7 !important; font-family:-apple-system,'SF Pro Display','Helvetica Neue',Arial,sans-serif; }
[data-testid="stSidebar"] { background:#FFFFFF !important; border-right:0.5px solid #E5E5EA; }
[data-testid="stSidebar"] * { font-family:-apple-system,'SF Pro Display',Arial,sans-serif; }

/* ── Header ── */
.sb-header { font-size:28px; font-weight:700; color:#1C1C1E; letter-spacing:-0.03em; margin-bottom:2px; }
.sb-sub    { font-size:13px; color:#8E8E93; margin-top:0; margin-bottom:16px; font-weight:400; }

/* ── KPI cards ── */
div[data-testid="stMetric"] > div {
    background:#FFFFFF !important; border-radius:16px !important;
    padding:16px 18px !important; border:0.5px solid #E5E5EA !important;
    box-shadow:0 1px 4px rgba(0,0,0,0.05) !important;
    transition:box-shadow .15s ease;
}
div[data-testid="stMetric"] > div:hover { box-shadow:0 3px 12px rgba(0,0,0,0.09) !important; }
div[data-testid="stMetric"] label { font-size:11px !important; font-weight:600 !important;
    color:#8E8E93 !important; text-transform:uppercase; letter-spacing:0.05em; }
div[data-testid="stMetric"] [data-testid="stMetricValue"] { font-size:26px !important;
    font-weight:700 !important; color:#1C1C1E !important; letter-spacing:-0.02em; }

/* ── Sections ── */
.section-title {
    font-size:13px; font-weight:600; color:#1C1C1E;
    border-left:3px solid #007AFF; padding-left:10px;
    margin:20px 0 10px; letter-spacing:-0.01em;
}

/* ── Message boxes ── */
.info-box {
    background:#F0F7FF; border-radius:14px; padding:14px 18px;
    border-left:3px solid #007AFF; font-size:13px; color:#1C1C1E;
    margin-bottom:14px; line-height:1.6;
}
.warn-box {
    background:#FFFBF0; border-radius:14px; padding:14px 18px;
    border-left:3px solid #FF9500; font-size:13px; color:#1C1C1E;
    margin-bottom:14px; line-height:1.6;
}
.ok-box {
    background:#F0FFF4; border-radius:14px; padding:14px 18px;
    border-left:3px solid #34C759; font-size:13px; color:#1C1C1E;
    margin-bottom:14px; line-height:1.6;
}

/* ── Commentaires automatiques ── */
.commentaire {
    background:#FFFFFF; border-radius:12px; padding:12px 16px;
    border:0.5px solid #FFD580; border-left:3px solid #FF9500;
    font-size:13px; color:#1C1C1E; margin-bottom:8px;
    line-height:1.6; box-shadow:0 1px 3px rgba(255,149,0,0.08);
}

/* ── Tabs ── */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    background:#F2F2F7; border-radius:12px; padding:4px; gap:2px;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    border-radius:9px; font-size:13px; font-weight:500;
    color:#8E8E93; padding:7px 16px; background:transparent;
}
[data-testid="stTabs"] [aria-selected="true"] {
    background:#FFFFFF !important; color:#1C1C1E !important;
    box-shadow:0 1px 4px rgba(0,0,0,0.10);
}

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border-radius:14px !important; overflow:hidden;
    border:0.5px solid #E5E5EA !important; }

/* ── Buttons ── */
[data-testid="stButton"] button, [data-testid="stDownloadButton"] button {
    border-radius:12px !important; font-weight:600 !important;
    font-size:13px !important; border:0.5px solid #E5E5EA !important;
    background:#FFFFFF !important; color:#007AFF !important;
    padding:8px 18px !important; transition:all .15s ease !important;
}
[data-testid="stButton"] button:hover { background:#F0F7FF !important;
    box-shadow:0 2px 8px rgba(0,122,255,0.15) !important; }

/* ── Sidebar nav ── */
[data-testid="stSidebarNav"] { display:none; }
.sidebar-nav-item { padding:6px 10px; border-radius:9px; font-size:13px;
    color:#3C3C43; cursor:pointer; display:block; }
.sidebar-nav-item:hover { background:#F2F2F7; }

/* ── Expander ── */
[data-testid="stExpander"] { border-radius:14px !important;
    border:0.5px solid #E5E5EA !important; overflow:hidden; }
[data-testid="stExpander"] summary {
    font-size:13px !important; font-weight:600 !important; color:#1C1C1E !important;
    padding:12px 16px !important; background:#FAFAFA !important;
}

/* ── Slider ── */
[data-testid="stSlider"] label { font-size:12px !important; color:#8E8E93 !important; font-weight:600 !important; }
[data-testid="stSlider"] [data-baseweb="slider"] div[role="slider"] {
    background:#007AFF !important; border:2px solid #FFFFFF !important;
    box-shadow:0 2px 6px rgba(0,122,255,0.35) !important;
}

/* ── Selectbox ── */
[data-testid="stSelectbox"] label { font-size:12px !important; color:#8E8E93 !important; font-weight:600 !important; }
[data-testid="stSelectbox"] > div > div { border-radius:10px !important;
    border:0.5px solid #E5E5EA !important; }

/* ── File uploader ── */
[data-testid="stFileUploader"] { border-radius:14px !important; }
[data-testid="stFileUploader"] label { font-size:12px !important; color:#8E8E93 !important; font-weight:600 !important; }
[data-testid="stFileUploader"] > div { border-radius:14px !important;
    border:1.5px dashed #C7C7CC !important; background:#FAFAFA !important; }
</style>
""", unsafe_allow_html=True)

ACHETEURS = {'BOISSONS':'Acheteur Boissons','EPICERIE':'Acheteur Épicerie','DROGUERIE':'Acheteur DPH','PARFUMERIE HYGIENE':'Acheteur DPH'}
PLANCHERS = {'Produit d appel':0.10,'Valeur ajoutee':0.25,'PH Droguerie':0.22,'Coeur de gamme':0.18}
SEG_LABELS = {'Produit d appel':'Appel','Valeur ajoutee':'Val. aj.','PH Droguerie':'PH/Drog','Coeur de gamme':'Cœur'}
PRODUITS_APPEL = ['RIZ LONG','HUILES','LAITS','EAUX PLATES','EAUX GAZEUSES','SUCRES ET LEVURES','FARINES CEREALES','LEGUMES SEC','PATES LONGUES','PATES COURTES','SEMOULES COUSCOUS','BOUILLON AIDE CULINAIRE','SAUCES FROIDES']
VALEUR_AJOUTEE = ['BIO','CHIPS','SOINS DU CORPS','SOINS DU VISAGE','PANSEMENT ET COMPLEMENTS','PRODUITS DU MONDE','SELS ET EPICES','SNACKING','DERMO COSMETIQUE','DIET','MAQUILLAGE']
ORDRE_RAYONS = ['BOISSONS','EPICERIE','DROGUERIE','PARFUMERIE HYGIENE']
TOLERANCE    = 0.015
COLS_REQUIRED = ['Rayon','Sous Famille','CA','Marge','%Marge','CA N-1','%Vs N-1.1']

def fp(v, sign=True):
    try:
        if pd.isna(v): return '—'
        return f"{v:+.1%}" if sign else f"{v:.1%}"
    except: return '—'

def fk(v):
    try:
        if pd.isna(v) or v == 0: return '—'
        return f"{v/1000:+,.0f} K"
    except: return '—'

def cs(v):
    if '✅' in str(v): return 'background:#E3F9E5;color:#1B6B1B;font-weight:600'
    if '🟡' in str(v): return 'background:#FFF3E0;color:#B25000;font-weight:600'
    if '🔴' in str(v): return 'background:#FFE5E5;color:#CC0000;font-weight:600'
    return ''

def cd(v):
    try:
        x = float(str(v).replace('%','').replace('+','').replace(' K','').replace(',','').replace('—','').strip())
        if x >= -1.5: return 'color:#1B6B1B;font-weight:600'
        if x >= -3.0: return 'color:#B25000;font-weight:600'
        return 'color:#CC0000;font-weight:600'
    except: return ''

def _segment_vec(sf_s, ray_s):
    sf  = sf_s.str.upper().fillna('')
    ray = ray_s.str.upper().fillna('')
    seg = pd.Series('Coeur de gamme', index=sf.index)
    seg[ray.isin(['DROGUERIE','PARFUMERIE HYGIENE'])] = 'PH Droguerie'
    mask_a = sf.apply(lambda s: any(p in s for p in PRODUITS_APPEL))
    seg[mask_a & (seg == 'Coeur de gamme')] = 'Produit d appel'
    mask_v = sf.apply(lambda s: any(v in s for v in VALEUR_AJOUTEE))
    seg[mask_v & (seg == 'Coeur de gamme')] = 'Valeur ajoutee'
    return seg

def _statut_vec(dev):
    s = pd.Series('⚪ N/A', index=dev.index)
    nn = dev.notna()
    s[nn & (dev >= -TOLERANCE)]                         = '✅ OK'
    s[nn & (dev < -TOLERANCE) & (dev >= -TOLERANCE*2)]  = '🟡 Vigilance'
    s[nn & (dev < -TOLERANCE*2)]                        = '🔴 Action requise'
    return s

def _que_faire_vec(df):
    tx  = df['%Marge']
    dev = df['Dev_N1_pts']
    seg = df['Segment']
    ca_med = df.groupby('Rayon_court')['CA'].transform('median').fillna(1)
    gros   = df['CA'] > ca_med * 1.5
    c = pd.Series('✅ RAS — surveiller la tendance', index=df.index)
    c[dev.notna() & (dev < -TOLERANCE*2)]       = '🔍 Optimiser mix promo / conditions fournisseur'
    c[dev.notna() & (dev < -0.05)]              = '📋 Révision conditions achat + audit promos en cours'
    c[dev.notna() & (dev < -0.05) & gros]       = '📞 Volume élevé + marge en chute — renégociation urgente fournisseur'
    c[dev.notna() & (dev < -0.10)]              = '📞 Convocation fournisseur — analyse prix achat vs marché'
    c[seg == 'Produit d appel']                 = '📋 Négocier remise arrière ou ristourne volume fournisseur'
    c[(seg=='Produit d appel') & dev.notna() & (dev < -0.05)] = "📞 Produit d'appel sous pression — remise arrière + revue franco fournisseur"
    c[tx.notna() & (tx < 0)]                   = '🚨 Marge négative — bloquer la promo, revoir le prix de cession immédiatement'
    return c

def _impact_score_vec(df):
    ca_med = df.groupby('Rayon_court')['CA'].transform('median').replace(0,1)
    poids  = (df['CA'] / ca_med).clip(0.5, 3.0)
    return (df['Dev_N1_FCFA'].abs() * poids).round(0)

def _commentaires_auto(df):
    rouge = df[df['Statut']=='🔴 Action requise'].nlargest(6,'Impact_Score')
    out = []
    for _,r in rouge.iterrows():
        sf=r['SF_court']; ray=r['Rayon_court']; tx=r['%Marge']
        dev=r['Dev_N1_pts']; fcfa=r['Dev_N1_FCFA']; seg=r['Segment']; vol=r['CA']
        if pd.notna(tx) and tx < 0:
            out.append(f"🚨 **{sf}** ({ray}) — marge négative à {tx:.1%}. Arrêt immédiat des promos déficitaires.")
        elif pd.notna(dev) and dev < -0.10:
            out.append(f"🔴 **{sf}** ({ray}) — effondrement de {dev:+.1%} vs N-1 ({fcfa:+,.0f} FCFA). Convocation fournisseur urgente.")
        elif seg == 'Produit d appel':
            out.append(f"🟠 **{sf}** ({ray}) — produit d'appel à {tx:.1%} ({dev:+.1%} vs N-1). Négocier remise arrière fournisseur.")
        else:
            vtxt = f", gros volume ({vol/1e6:.1f}M FCFA)" if vol > 5e6 else ""
            out.append(f"🟠 **{sf}** ({ray}{vtxt}) — {dev:+.1%} vs N-1 ({fcfa:+,.0f} FCFA). Réviser conditions achat.")
    return out

def _detect_periode(df_raw):
    col_a = df_raw.iloc[:,0].dropna().astype(str)
    last  = col_a.iloc[-1] if len(col_a) else ''
    dates = re.findall(r'\d{2}/\d{2}/\d{4}', last)
    if len(dates) >= 2:   return f"{dates[0]} → {dates[1]}"
    elif len(dates) == 1: return dates[0]
    return 'Période inconnue'

def _validate(df, filename):
    missing = [c for c in COLS_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"**{filename}** : colonnes manquantes → `{'`, `'.join(missing)}`")

@st.cache_data(show_spinner=False)
def load_referentiel(override_bytes=None):
    if override_bytes:
        try:
            ref = pd.read_excel(BytesIO(override_bytes))
            if 'Cible' in ref.columns: return ref
        except Exception: pass
    repo_path = os.path.join(os.path.dirname(__file__),'..','data','referentiel_cibles.csv')
    if os.path.exists(repo_path): return pd.read_csv(repo_path)
    return None

@st.cache_data(show_spinner=False)
def load_extraction(file_bytes: bytes, filename: str, ref_bytes=None):
    raw = BytesIO(file_bytes)
    df_raw  = pd.read_excel(raw, header=None)
    periode = _detect_periode(df_raw)
    raw.seek(0)
    df = pd.read_excel(raw)
    _validate(df, filename)
    df = df[
        (df['Sous Famille'] != 'Total') &
        df['Rayon'].str.startswith('000', na=False) &
        ~df['Rayon'].str.contains('CIGARETTE', na=False) &
        df['CA'].notna() & (df['CA'] > 0)
    ].copy()
    df['SF_court']    = df['Sous Famille'].str.extract(r'\d+ - (.+)')[0]
    df['Rayon_court'] = df['Rayon'].str.extract(r'- (.+)')[0]
    df['Acheteur']    = df['Rayon_court'].map(ACHETEURS)
    df['Segment']     = _segment_vec(df['SF_court'], df['Rayon_court'])
    df['Plancher']    = df['Segment'].map(PLANCHERS)
    valid_n1   = df['%Vs N-1.1'].notna() & (df['%Vs N-1.1'] != -1) & (df['CA N-1'] > 0)
    df['Marge_N1'] = np.where(valid_n1, df['Marge'] / (1 + df['%Vs N-1.1']), np.nan)
    df['CA_N1']    = df['CA N-1']
    df['Tx_N1']    = np.where(valid_n1 & (df['CA_N1']>0), df['Marge_N1']/df['CA_N1'], np.nan)
    ref = load_referentiel(ref_bytes)
    if ref is not None and 'Cible' in ref.columns:
        df = df.merge(
            ref[['Rayon','Famille','Cible','Plancher']].rename(columns={'Rayon':'Rayon_court','Famille':'SF_court','Cible':'Cible_ref','Plancher':'Plancher_ref'}),
            on=['Rayon_court','SF_court'], how='left')
        df['Cible']    = df['Cible_ref'].fillna(df['Plancher'])
        df['Plancher'] = df['Plancher_ref'].fillna(df['Plancher'])
    else:
        df['Cible'] = np.where(df['Tx_N1'].notna(), np.maximum(df['Tx_N1']*1.02, df['Plancher']), df['Plancher'])
    df['Source_cible']   = np.where(df['Tx_N1'].notna(), 'N-1×1,02', 'Plancher (nouveauté)')
    df['Dev_N1_pts']     = df['%Marge'] - df['Tx_N1']
    df['Dev_N1_FCFA']    = df['Dev_N1_pts'] * df['CA']
    df['Dev_Cible_pts']  = df['%Marge'] - df['Cible']
    df['Dev_Cible_FCFA'] = df['Dev_Cible_pts'] * df['CA']
    df['Statut']         = _statut_vec(df['Dev_N1_pts'])
    df['Impact_Score']   = _impact_score_vec(df)
    df['Que_faire']      = _que_faire_vec(df)
    _ord = {'🔴 Action requise':0,'🟡 Vigilance':1,'✅ OK':2,'⚪ N/A':3}
    df['_ord_statut'] = df['Statut'].map(_ord)
    df['_ord_rayon']  = df['Rayon_court'].map({r:i for i,r in enumerate(ORDRE_RAYONS)})
    df['Periode']  = periode
    df['Fichier']  = filename
    return df

def export_excel(df_all, periodes):
    wb = Workbook()
    def fill(h): return PatternFill('solid',start_color=h,end_color=h)
    def thin():
        s=Side(style='thin',color='FFB0B0B0')
        return Border(left=s,right=s,top=s,bottom=s)
    C_HDR='FF1C3A5C'; C_R='FFFFD6D6'; C_O='FFFFF3CC'; C_G='FFD6F5D6'; C_L='FFF2F2F7'; C_W='FFFFFFFF'
    HDRS=['Rayon','Famille','Segment','Source cible','Acheteur','CA (FCFA)','Taux actuel','Taux N-1','Cible','Plancher','Dév. vs N-1 (pts)','Dév. vs Cible (pts)','Marge Δ FCFA (vs N-1)','Score Impact','Statut','Que faire ?']
    WIDTHS=[20,30,12,14,18,14,11,11,11,11,16,16,18,14,16,45]
    for i_p,periode in enumerate(periodes):
        df = df_all[df_all['Periode']==periode].sort_values(['_ord_statut','Impact_Score'],ascending=[True,False])
        safe = periode.replace('/','').replace('→','_').replace(' ','')[:28]
        ws   = wb.active if i_p==0 else wb.create_sheet(safe)
        ws.title = safe; ws.sheet_view.showGridLines = False
        t=ws.cell(row=1,column=1,value=f'SUIVI RENTABILITÉ — {periode}')
        t.font=Font(name='Arial',bold=True,size=12,color='FFFFFFFF'); t.fill=fill(C_HDR)
        t.alignment=Alignment(horizontal='center')
        ws.merge_cells(f'A1:{get_column_letter(len(HDRS))}1'); ws.row_dimensions[1].height=24
        for j,(h,w) in enumerate(zip(HDRS,WIDTHS),1):
            c=ws.cell(row=2,column=j,value=h)
            c.font=Font(name='Arial',bold=True,size=9,color='FFFFFFFF'); c.fill=fill('FF2E5C8A')
            c.alignment=Alignment(horizontal='center',wrap_text=True); c.border=thin()
            ws.column_dimensions[get_column_letter(j)].width=w
        ws.row_dimensions[2].height=32
        for i,(_,r) in enumerate(df.iterrows(),3):
            stat=r.get('Statut','')
            bg=C_R if '🔴' in str(stat) else (C_O if '🟡' in str(stat) else (C_G if '✅' in str(stat) else (C_L if i%2==0 else C_W)))
            vals=[r.get('Rayon_court',''),r.get('SF_court',''),SEG_LABELS.get(r.get('Segment',''),r.get('Segment','')),r.get('Source_cible',''),r.get('Acheteur',''),r.get('CA',None),r.get('%Marge',None),r.get('Tx_N1',None),r.get('Cible',None),r.get('Plancher',None),r.get('Dev_N1_pts',None),r.get('Dev_Cible_pts',None),r.get('Dev_N1_FCFA',None),r.get('Impact_Score',None),stat,r.get('Que_faire','')]
            for j,v in enumerate(vals,1):
                c=ws.cell(row=i,column=j,value=v); c.fill=fill(bg); c.border=thin()
                c.font=Font(name='Arial',size=9)
                c.alignment=Alignment(horizontal='center' if j in {7,8,9,10,11,12,15} else 'left',wrap_text=(j==16))
                if j==6:   c.number_format='#,##0'
                elif j==13: c.number_format='+#,##0;-#,##0;-'
                elif j==14: c.number_format='#,##0'
                elif j in {7,8,9,10}: c.number_format='0.0%'
                elif j in {11,12}:    c.number_format='+0.0%;-0.0%;-'
            ws.row_dimensions[i].height=16
        ws.freeze_panes='A3'
        ws.auto_filter.ref=f'A2:{get_column_letter(len(HDRS))}{1+len(df)}'
    buf=BytesIO(); wb.save(buf); buf.seek(0); return buf

# ── SIDEBAR ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("<div style='margin-bottom:16px'><div style='font-size:20px;font-weight:700;color:#1C1C1E'>🛍️ SmartBuyer</div><div style='font-size:11px;color:#8E8E93'>Hub analytique · Équipe Achats</div></div>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px'>Navigation</div>", unsafe_allow_html=True)
    for path,label in [("app.py","🏠  Accueil"),("pages/01_📊_Analyse_Scoring_ABC.py","📊  Scoring ABC"),("pages/02_📈_Ventes_PBI.py","📈  Ventes PBI"),("pages/03_📦_Detention_Top_CA.py","📦  Détention Top CA"),("pages/04_💸_Performance_Promo.py","💸  Performance Promo"),("pages/05_🏪_Suivi_Implantation.py","🏪  Suivi Implantation"),("pages/06_💸_Marges_Negatives.py","💸  Marges Négatives"),("pages/09_📦_OTIF.py","📦  OTIF"),("pages/10_📉_OOS_Ruptures.py","📉  OOS Ruptures"),("pages/11_📊_Rentabilite.py","📊  Rentabilité")]:
        try: st.page_link(path, label=label)
        except Exception: st.markdown(f"- {label}")
    st.markdown("---")
    uploaded_files = st.file_uploader("Extraction(s) PBI",type=['xlsx'],accept_multiple_files=True,help=f"Colonnes obligatoires : {', '.join(COLS_REQUIRED)}")
    ref_override   = st.file_uploader("Référentiel cibles (optionnel)",type=['xlsx'],help="Laissez vide → référentiel embarqué.")
    st.markdown("---")
    seuil_fcfa = st.slider("Impact min (K FCFA)",min_value=0,max_value=2000,value=0,step=50,help="Masquer les familles sous ce seuil de marge perdue")
    st.markdown("---")
    st.markdown("<div style='font-size:11px;color:#8E8E93;line-height:2'>🔴 &lt; −3 pts vs N-1<br>🟡 −1,5 à −3 pts<br>✅ &gt; −1,5 pt</div>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("<div style='font-size:10px;color:#C7C7CC;text-align:center'>NovaRetail Solutions · SmartBuyer v2.2</div>", unsafe_allow_html=True)

# ── ÉCRAN VIDE ─────────────────────────────────────────────────────────────────
if not uploaded_files:
    st.markdown('<div class="sb-header">📊 Rentabilité — Suivi Déviation Marge</div>', unsafe_allow_html=True)
    st.markdown('<div class="sb-sub">Pilotage hebdomadaire par famille · acheteur · magasin</div>', unsafe_allow_html=True)
    st.markdown('<div class="info-box">👈 <strong>Chargez une ou plusieurs extractions PBI</strong> dans la barre latérale.</div>', unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1:
        st.markdown("**Colonnes obligatoires**")
        st.markdown("\n".join([f"- `{c}`" for c in COLS_REQUIRED]))
    with c2:
        st.markdown("**Ce module permet de :**")
        st.markdown("- 📊 Déviation marge vs N-1 et vs Cible\n- 🎯 Score d'impact financier pour prioriser\n- 🔴 Action recommandée par famille et situation\n- 📅 Détection des dégradations persistantes\n- 📤 Export Excel par acheteur")
    st.stop()

# ── CHARGEMENT ─────────────────────────────────────────────────────────────────
ref_bytes = ref_override.read() if ref_override else None
all_dfs, errors = [], []
for f in uploaded_files:
    raw = f.read()
    try:    all_dfs.append(load_extraction(raw, f.name, ref_bytes))
    except ValueError as e: errors.append(str(e))
    except Exception as e:  errors.append(f"Erreur inattendue **{f.name}** : {e}")
for err in errors: st.sidebar.error(err)
if not all_dfs:
    st.error("Aucun fichier valide. Vérifiez les colonnes obligatoires.")
    st.stop()

df_all = pd.concat(all_dfs, ignore_index=True)
periodes_dispo = sorted(df_all['Periode'].unique(), reverse=True)

with st.sidebar:
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px'>Périodes chargées</div>", unsafe_allow_html=True)
    for p in periodes_dispo:
        st.markdown(f"<span style='background:#E3F0FF;color:#185FA5;border-radius:20px;padding:3px 10px;font-size:11px;font-weight:500'>{p}</span>", unsafe_allow_html=True)
    st.markdown("")
    periode_sel = st.selectbox("Période active", periodes_dispo, label_visibility='collapsed')

df = df_all[df_all['Periode']==periode_sel].copy()
if seuil_fcfa > 0:
    df = df[(df['Dev_N1_FCFA'].abs() >= seuil_fcfa*1000) | df['Dev_N1_FCFA'].isna()]

n_nouv = (df_all[df_all['Periode']==periode_sel]['Source_cible']=='Plancher (nouveauté)').sum()
st.markdown('<div class="sb-header">📊 Rentabilité — Suivi Déviation Marge</div>', unsafe_allow_html=True)
st.markdown(f'<div class="sb-sub">Période : <strong>{periode_sel}</strong> · {len(periodes_dispo)} période(s) · {n_nouv} nouveauté(s) comparée(s) sur cible segment uniquement</div>', unsafe_allow_html=True)

tab_reseau, tab_acheteur, tab_magasin, tab_tendance = st.tabs(["📊 Réseau","👤 Acheteur","🏪 Magasin","📅 Tendance"])

# ── TAB RÉSEAU ─────────────────────────────────────────────────────────────────
with tab_reseau:
    ca_t=df['CA'].sum(); mg_t=df['Marge'].sum(); tx_t=mg_t/ca_t if ca_t>0 else 0
    mn1_t=df['Marge_N1'].sum(); cn1_t=df['CA_N1'].sum(); tx_n1=mn1_t/cn1_t if cn1_t>0 else 0
    dev_t=tx_t-tx_n1; n_r=(df['Statut']=='🔴 Action requise').sum()
    n_o=(df['Statut']=='🟡 Vigilance').sum(); n_v=(df['Statut']=='✅ OK').sum()

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    with c1: st.metric("Taux actuel",  f"{tx_t:.1%}", fp(dev_t))
    with c2: st.metric("Taux N-1",     f"{tx_n1:.1%}")
    with c3: st.metric("Marge Δ FCFA", fk(df['Dev_N1_FCFA'].sum()))
    with c4: st.metric("🔴 Action",    n_r, f"sur {len(df)} familles", delta_color="off")
    with c5: st.metric("🟡 Vigilance", n_o, delta_color="off")
    with c6: st.metric("✅ OK",        n_v, delta_color="off")

    commentaires = _commentaires_auto(df)
    if commentaires:
        with st.expander("💬 Analyse automatique — points de blocage", expanded=True):
            for c in commentaires:
                st.markdown(f'<div class="commentaire">{c}</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Synthèse par rayon</div>', unsafe_allow_html=True)
    rows_r = []
    for rayon in ORDRE_RAYONS:
        sub=df[df['Rayon_court']==rayon]
        if len(sub)==0: continue
        ca_r=sub['CA'].sum(); mg_r=sub['Marge'].sum(); tx_r=mg_r/ca_r if ca_r>0 else 0
        mn1_r=sub['Marge_N1'].sum(); cn1_r=sub['CA_N1'].sum(); tn1_r=mn1_r/cn1_r if cn1_r>0 else 0
        cib_r=(sub['Cible']*sub['CA']).sum()/ca_r if ca_r>0 else 0
        rows_r.append({'Rayon':rayon,'Acheteur':ACHETEURS.get(rayon,'—'),'CA (K)':f"{ca_r/1000:,.0f}",'Taux actuel':fp(tx_r,False),'Taux N-1':fp(tn1_r,False),'Dév. vs N-1':fp(tx_r-tn1_r),'Cible':fp(cib_r,False),'Dév. vs Cible':fp(tx_r-cib_r),'Marge Δ FCFA':fk(sub['Dev_N1_FCFA'].sum()),'🔴':(sub['Statut']=='🔴 Action requise').sum(),'🟡':(sub['Statut']=='🟡 Vigilance').sum(),'✅':(sub['Statut']=='✅ OK').sum(),'Statut':'🔴 Action requise' if (sub['Statut']=='🔴 Action requise').sum()>2 else ('🟡 Vigilance' if (sub['Statut']=='🟡 Vigilance').sum()>2 else '✅ OK')})
    st.dataframe(pd.DataFrame(rows_r).style.map(cs,subset=['Statut']).map(cd,subset=['Dév. vs N-1','Dév. vs Cible']),use_container_width=True,hide_index=True,height=210)

    st.markdown('<div class="section-title">🔴 Familles — triées par score d\'impact financier</div>', unsafe_allow_html=True)
    rouge=df[df['Statut']=='🔴 Action requise'].sort_values('Impact_Score',ascending=False)
    if len(rouge)==0:
        st.markdown('<div class="ok-box">✅ Aucune famille en alerte rouge.</div>', unsafe_allow_html=True)
    else:
        disp=rouge[['Rayon_court','SF_court','Segment','Acheteur','CA','%Marge','Tx_N1','Dev_N1_pts','Dev_N1_FCFA','Impact_Score','Cible','Source_cible','Dev_Cible_pts','Statut','Que_faire']].copy()
        disp.columns=['Rayon','Famille','Segment','Acheteur','CA (FCFA)','Taux act.','Taux N-1','Dév. N-1','Marge Δ FCFA','Score Impact','Cible','Source cible','Dév. Cible','Statut','Que faire ?']
        for col in ['CA (FCFA)','Score Impact']: disp[col]=disp[col].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else '—')
        for col in ['Taux act.','Taux N-1','Cible']: disp[col]=disp[col].apply(lambda x: fp(x,False))
        for col in ['Dév. N-1','Dév. Cible']: disp[col]=disp[col].apply(fp)
        disp['Marge Δ FCFA']=disp['Marge Δ FCFA'].apply(fk)
        disp['Segment']=disp['Segment'].apply(lambda x: SEG_LABELS.get(x,x))
        st.dataframe(disp.style.map(cs,subset=['Statut']).map(cd,subset=['Dév. N-1','Dév. Cible']),use_container_width=True,hide_index=True,height=420)

    with st.expander(f"🟡 Familles en vigilance ({n_o})"):
        orange=df[df['Statut']=='🟡 Vigilance'].sort_values('Impact_Score',ascending=False)
        d2=orange[['Rayon_court','SF_court','Segment','Acheteur','%Marge','Tx_N1','Dev_N1_pts','Dev_N1_FCFA','Impact_Score','Que_faire']].copy()
        d2.columns=['Rayon','Famille','Segment','Acheteur','Taux act.','Taux N-1','Dév. N-1','Marge Δ FCFA','Score Impact','Que faire ?']
        for col in ['Taux act.','Taux N-1']: d2[col]=d2[col].apply(lambda x: fp(x,False))
        d2['Dév. N-1']=d2['Dév. N-1'].apply(fp); d2['Marge Δ FCFA']=d2['Marge Δ FCFA'].apply(fk)
        d2['Score Impact']=d2['Score Impact'].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else '—')
        d2['Segment']=d2['Segment'].apply(lambda x: SEG_LABELS.get(x,x))
        st.dataframe(d2.style.map(cd,subset=['Dév. N-1']),use_container_width=True,hide_index=True)

    st.markdown("")
    col_exp,_=st.columns([1,3])
    with col_exp:
        if st.button("📤 Exporter rapport Excel",use_container_width=True):
            buf=export_excel(df_all,periodes_dispo)
            st.download_button("⬇️ Télécharger",data=buf,file_name=f"Rentabilite_{periode_sel.replace('/','').replace('→','_').replace(' ','')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

# ── TAB ACHETEUR ───────────────────────────────────────────────────────────────
with tab_acheteur:
    acheteurs=sorted(df['Acheteur'].dropna().unique())
    ach_sel=st.selectbox("Acheteur",acheteurs,key='ach_sel')
    df_ach=df[df['Acheteur']==ach_sel].sort_values(['_ord_statut','Impact_Score'],ascending=[True,False])
    ca_a=df_ach['CA'].sum(); mg_a=df_ach['Marge'].sum(); tx_a=mg_a/ca_a if ca_a>0 else 0
    mn1_a=df_ach['Marge_N1'].sum(); cn1_a=df_ach['CA_N1'].sum(); tn1_a=mn1_a/cn1_a if cn1_a>0 else 0
    cib_a=(df_ach['Cible']*df_ach['CA']).sum()/ca_a if ca_a>0 else 0
    n_r_a=(df_ach['Statut']=='🔴 Action requise').sum()

    c1,c2,c3,c4,c5=st.columns(5)
    with c1: st.metric("Taux réalisé",f"{tx_a:.1%}",fp(tx_a-tn1_a))
    with c2: st.metric("Taux N-1",f"{tn1_a:.1%}")
    with c3: st.metric("Marge Δ",fk(df_ach['Dev_N1_FCFA'].sum()))
    with c4: st.metric("🔴 Alertes",n_r_a,f"sur {len(df_ach)} familles",delta_color="off")
    with c5: st.metric("Cible moy.",f"{cib_a:.1%}")

    if n_r_a > 0:
        top3=df_ach[df_ach['Statut']=='🔴 Action requise'].nlargest(3,'Impact_Score')['SF_court'].tolist()
        st.markdown(f'<div class="warn-box">🎯 <strong>Briefing {ach_sel} :</strong> {n_r_a} alerte(s) rouge(s).<br>Priorités : <strong>{" · ".join(top3)}</strong><br>👇 Tableau trié par score d\'impact — colonne <em>Que faire ?</em> = action à mener.</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="ok-box">✅ <strong>{ach_sel}</strong> — Aucune alerte rouge. Surveiller les familles en vigilance.</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Toutes les familles — triées par priorité et impact</div>', unsafe_allow_html=True)
    da=df_ach[['Rayon_court','SF_court','Segment','CA','%Marge','Tx_N1','Dev_N1_pts','Dev_N1_FCFA','Impact_Score','Cible','Plancher','Source_cible','Dev_Cible_pts','Statut','Que_faire']].copy()
    da.columns=['Rayon','Famille','Segment','CA (FCFA)','Taux act.','Taux N-1','Dév. N-1','Marge Δ FCFA','Score Impact','Cible','Plancher','Source cible','Dév. Cible','Statut','Que faire ?']
    da['CA (FCFA)']=da['CA (FCFA)'].apply(lambda x: f"{x:,.0f}")
    da['Score Impact']=da['Score Impact'].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else '—')
    for col in ['Taux act.','Taux N-1','Cible','Plancher']: da[col]=da[col].apply(lambda x: fp(x,False))
    for col in ['Dév. N-1','Dév. Cible']: da[col]=da[col].apply(fp)
    da['Marge Δ FCFA']=da['Marge Δ FCFA'].apply(fk)
    da['Segment']=da['Segment'].apply(lambda x: SEG_LABELS.get(x,x))
    st.dataframe(da.style.map(cs,subset=['Statut']).map(cd,subset=['Dév. N-1','Dév. Cible']),use_container_width=True,hide_index=True,height=580)

    col_ea,_=st.columns([1,3])
    with col_ea:
        if st.button("📤 Export acheteur",use_container_width=True):
            buf=export_excel(df_all[df_all['Acheteur']==ach_sel],[periode_sel])
            st.download_button("⬇️ Télécharger",data=buf,file_name=f"Rentabilite_{ach_sel.replace(' ','_')}_{periode_sel.replace('/','').replace('→','_').replace(' ','')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

# ── TAB MAGASIN ────────────────────────────────────────────────────────────────
with tab_magasin:
    has_site='Site nom long' in df.columns and df['Site nom long'].notna().any()
    if not has_site:
        st.markdown('<div class="info-box">ℹ️ Extraction au niveau réseau — pas de détail magasin.<br>Relancez l\'extraction PBI en ajoutant la dimension <strong>Site nom long</strong>.</div>', unsafe_allow_html=True)
    else:
        sites=sorted([s for s in df['Site nom long'].dropna().unique() if s not in ['Total','']])
        cf1,cf2,cf3=st.columns(3)
        with cf1: site_sel=st.selectbox("Magasin",['Tous']+sites,key='site_sel')
        with cf2: rayon_f=st.selectbox("Rayon",['Tous']+ORDRE_RAYONS,key='rayon_f')
        with cf3: stat_f=st.selectbox("Statut",['Tous','🔴 Action requise','🟡 Vigilance','✅ OK'],key='stat_f')
        df_mag=df.copy()
        if site_sel!='Tous': df_mag=df_mag[df_mag['Site nom long']==site_sel]
        if rayon_f!='Tous':  df_mag=df_mag[df_mag['Rayon_court']==rayon_f]
        if stat_f!='Tous':   df_mag=df_mag[df_mag['Statut']==stat_f]
        if site_sel=='Tous':
            rows_s=[]
            for site in sites:
                sub_s=df[df['Site nom long']==site]; ca_s=sub_s['CA'].sum()
                if ca_s==0: continue
                mg_s=sub_s['Marge'].sum(); tx_s=mg_s/ca_s
                mn1_s=sub_s['Marge_N1'].sum(); cn1_s=sub_s['CA_N1'].sum(); tn1_s=mn1_s/cn1_s if cn1_s>0 else 0
                dev_s=tx_s-tn1_s
                rows_s.append({'Magasin':site,'CA (K)':f"{ca_s/1000:,.0f}",'Taux act.':fp(tx_s,False),'Taux N-1':fp(tn1_s,False),'Dév. vs N-1':fp(dev_s),'Marge Δ FCFA':fk(sub_s['Dev_N1_FCFA'].sum()),'🔴':(sub_s['Statut']=='🔴 Action requise').sum(),'🟡':(sub_s['Statut']=='🟡 Vigilance').sum(),'✅':(sub_s['Statut']=='✅ OK').sum(),'Statut':'🔴 Action requise' if dev_s<-TOLERANCE*2 else ('🟡 Vigilance' if dev_s<-TOLERANCE else '✅ OK')})
            st.dataframe(pd.DataFrame(rows_s).sort_values('Statut').style.map(cs,subset=['Statut']).map(cd,subset=['Dév. vs N-1']),use_container_width=True,hide_index=True)
        else:
            dm=df_mag.sort_values(['_ord_statut','Impact_Score'],ascending=[True,False])
            d3=dm[['Rayon_court','SF_court','Segment','Acheteur','%Marge','Tx_N1','Dev_N1_pts','Dev_N1_FCFA','Impact_Score','Statut','Que_faire']].copy()
            d3.columns=['Rayon','Famille','Segment','Acheteur','Taux act.','Taux N-1','Dév. N-1','Marge Δ FCFA','Score Impact','Statut','Que faire ?']
            for col in ['Taux act.','Taux N-1']: d3[col]=d3[col].apply(lambda x: fp(x,False))
            d3['Dév. N-1']=d3['Dév. N-1'].apply(fp); d3['Marge Δ FCFA']=d3['Marge Δ FCFA'].apply(fk)
            d3['Score Impact']=d3['Score Impact'].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else '—')
            d3['Segment']=d3['Segment'].apply(lambda x: SEG_LABELS.get(x,x))
            st.dataframe(d3.style.map(cs,subset=['Statut']).map(cd,subset=['Dév. N-1']),use_container_width=True,hide_index=True,height=520)

# ── TAB TENDANCE ───────────────────────────────────────────────────────────────
with tab_tendance:
    if len(periodes_dispo) < 2:
        st.markdown('<div class="info-box">Chargez au moins 2 extractions pour voir l\'évolution dans le temps.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="section-title">Évolution globale par période</div>', unsafe_allow_html=True)
        rows_t=[]
        for p in sorted(periodes_dispo):
            dp=df_all[df_all['Periode']==p]; ca_p=dp['CA'].sum(); mg_p=dp['Marge'].sum()
            tx_p=mg_p/ca_p if ca_p>0 else 0
            mn1_p=dp['Marge_N1'].sum(); cn1_p=dp['CA_N1'].sum(); tn1_p=mn1_p/cn1_p if cn1_p>0 else 0
            rows_t.append({'Période':p,'Taux réalisé':fp(tx_p,False),'Taux N-1':fp(tn1_p,False),'Dév. vs N-1':fp(tx_p-tn1_p),'Marge Δ FCFA':fk(dp['Dev_N1_FCFA'].sum()),'🔴 Alertes':(dp['Statut']=='🔴 Action requise').sum(),'🟡 Vigilance':(dp['Statut']=='🟡 Vigilance').sum(),'✅ OK':(dp['Statut']=='✅ OK').sum()})
        st.dataframe(pd.DataFrame(rows_t).style.map(cd,subset=['Dév. vs N-1']),use_container_width=True,hide_index=True)

        st.markdown('<div class="section-title">Dégradations persistantes — rouge sur toutes les périodes</div>', unsafe_allow_html=True)
        rouge_sets={p:set(df_all[df_all['Periode']==p][df_all[df_all['Periode']==p]['Statut']=='🔴 Action requise'].apply(lambda r: f"{r['Rayon_court']}|{r['SF_court']}",axis=1)) for p in periodes_dispo}
        persistants=set.intersection(*rouge_sets.values()) if rouge_sets else set()
        if persistants:
            st.markdown(f'<div class="warn-box">⚠️ <strong>{len(persistants)} famille(s)</strong> en rouge sur toutes les {len(periodes_dispo)} périodes — problèmes structurels à traiter en réunion fournisseur.</div>', unsafe_allow_html=True)
            rows_p=[]
            for key in sorted(persistants):
                rayon,sf=key.split('|',1)
                sub=df[(df['Rayon_court']==rayon)&(df['SF_court']==sf)]
                if len(sub):
                    r0=sub.iloc[0]
                    rows_p.append({'Rayon':rayon,'Famille':sf,'Acheteur':r0.get('Acheteur','—'),'Taux actuel':fp(r0.get('%Marge'),False),'Score Impact':f"{r0.get('Impact_Score',0):,.0f}",'Que faire ?':r0.get('Que_faire','—')})
            st.dataframe(pd.DataFrame(rows_p).sort_values('Score Impact',ascending=False),use_container_width=True,hide_index=True)
        else:
            st.markdown('<div class="ok-box">✅ Aucune famille en rouge sur toutes les périodes.</div>', unsafe_allow_html=True)
