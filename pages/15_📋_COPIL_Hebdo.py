"""
15_📋_COPIL_Hebdo.py — Module COPIL Hebdo · SmartBuyer Hub
Vue réseau (Rayon) + Destructeurs/Performeurs (Article) + Marge Négative par Site.
V2 : design Apple clair modernisé · dédoublonnage réseau (export à la maille Article × Site)
     · nouvelle vue Marge Négative par Site (cockpit + export Excel, 3e feuille).
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# ============================================================
# CONFIG & CHARTE (Apple clair — V2)
# ============================================================
st.set_page_config(page_title="COPIL Hebdo", page_icon="📋", layout="wide")

BLUE = "#007AFF"
GREEN = "#34C759"
RED = "#FF3B30"
AMBER = "#FF9500"
DARK = "#1D1D1F"
GREY = "#86868B"
BG = "#F5F5F7"

# ============================================================
# 🎯 CIBLES DE MARGE PAR RAYON — identiques au module 14_💰_Marge.py
# ============================================================
CIBLES_DEFAUT = {
    "BOISSONS": 19.5,
    "DROGUERIE": 25.0,
    "PARFUMERIE HYGIENE": 29.0,
    "EPICERIE": 16.0,
}
CIBLE_FALLBACK = 23.5

# Colonnes candidates pour la dimension Site dans l'export PBI (auto-détection)
SITE_CANDIDATES = ["Site", "Magasin", "Code Site", "Libellé Site", "Nom Site",
                   "Site de vente", "Etablissement", "Établissement", "Store"]

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {{
    font-family: -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', 'Inter', sans-serif;
    -webkit-font-smoothing: antialiased;
}}
.stApp {{ background: {BG}; }}
.block-container {{ padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1300px; }}
[data-testid="stSidebar"] {{
    background: rgba(255,255,255,0.85);
    backdrop-filter: blur(20px);
    border-right: 1px solid rgba(0,0,0,0.06);
}}
[data-testid="stSidebar"] .block-container {{ padding-top: 2rem; }}
hr {{ border: none !important; border-top: 1px solid rgba(0,0,0,0.07) !important; margin: 1.1rem 0 !important; }}

/* ---------- Titres ---------- */
.page-title {{
    font-size: 32px; font-weight: 800; color: {DARK};
    letter-spacing: -0.035em; margin: 0; line-height: 1.15;
}}
.page-caption {{ font-size: 14px; color: {GREY}; margin-top: 5px; margin-bottom: 1.4rem; letter-spacing: -0.01em; }}
.section-label {{
    display: flex; align-items: center; gap: 8px;
    font-size: 11px; font-weight: 700; color: {GREY};
    text-transform: uppercase; letter-spacing: 0.09em; margin: 26px 0 12px;
}}
.section-label::before {{
    content: ""; width: 8px; height: 8px; border-radius: 3px;
    background: {BLUE}; display: inline-block;
}}

/* ---------- Cartes ---------- */
.card, .kpi-card, .recap-card {{
    background: #FFFFFF; border-radius: 18px;
    border: 1px solid rgba(0,0,0,0.06);
    box-shadow: 0 2px 12px rgba(0,0,0,0.045);
}}
.card {{ padding: 16px 20px; margin-bottom: 12px; }}

.kpi-card {{ padding: 18px 20px; transition: transform .15s ease, box-shadow .15s ease; }}
.kpi-card:hover {{ transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.08); }}
.kpi-label {{
    font-size: 10.5px; font-weight: 600; color: {GREY};
    text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 6px;
}}
.kpi-value {{ font-size: 26px; font-weight: 700; color: {DARK}; letter-spacing: -0.03em; line-height: 1.05; }}
.kpi-sub {{ margin-top: 8px; font-size: 12px; color: {GREY}; }}
.pill {{
    display: inline-block; padding: 3px 10px; border-radius: 999px;
    font-size: 11.5px; font-weight: 600; letter-spacing: -0.01em;
}}
.pill.pos {{ background: rgba(52,199,89,0.13); color: #1A7A3A; }}
.pill.neg {{ background: rgba(255,59,48,0.12); color: #C62A22; }}
.pill.neutral {{ background: rgba(0,0,0,0.05); color: {GREY}; }}

/* ---------- Hero récap ---------- */
.recap-card {{
    background: linear-gradient(135deg, #FFFFFF 0%, #F2F8FF 100%);
    border: 1px solid #E1EEFF;
    box-shadow: 0 6px 24px rgba(0,64,221,0.07);
    padding: 18px 24px; margin-bottom: 20px;
}}
.recap-line1 {{ font-size: 15.5px; font-weight: 700; color: {DARK}; letter-spacing: -0.015em; line-height: 1.65; }}
.recap-line2 {{
    font-size: 13px; color: {DARK}; margin-top: 10px; padding-top: 10px;
    border-top: 1px solid rgba(0,64,221,0.09); line-height: 1.6;
}}
.recap-line2 b {{ color: {BLUE}; }}

/* ---------- Info box (état vide) ---------- */
.info-box {{
    background: #FFFFFF; border: 1px solid rgba(0,0,0,0.06);
    border-radius: 20px; padding: 24px 28px; margin-bottom: 24px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.045);
}}
.info-box .it {{ font-size: 17px; font-weight: 700; color: {DARK}; letter-spacing: -0.02em; margin-bottom: 10px; }}
.info-box .ip {{ font-size: 13.5px; color: #3A3A3C; line-height: 1.65; }}
.info-box .iq {{ margin-top: 14px; font-size: 13px; color: #3A3A3C; line-height: 1.95; }}

/* ---------- Badges de section (Tab 2) ---------- */
.badge {{
    display: inline-block; padding: 5px 14px; border-radius: 999px;
    font-size: 12px; font-weight: 600; letter-spacing: -0.01em; margin: 14px 0 6px;
}}

/* ---------- Tabs façon segmented control ---------- */
.stTabs [data-baseweb="tab-list"] {{
    background: #E9E9EB; padding: 3px; border-radius: 12px;
    gap: 2px; width: fit-content;
}}
.stTabs [data-baseweb="tab"] {{
    border-radius: 9px; padding: 6px 18px; background: transparent;
    font-weight: 600; font-size: 13.5px; color: {DARK};
}}
.stTabs [aria-selected="true"] {{
    background: #FFFFFF !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.12);
}}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display: none; }}

/* ---------- Dataframes & widgets ---------- */
[data-testid="stDataFrame"] {{
    border: 1px solid rgba(0,0,0,0.06); border-radius: 14px;
    overflow: hidden; box-shadow: 0 1px 6px rgba(0,0,0,0.035);
}}
.stDownloadButton button, .stButton button {{
    background: {BLUE}; color: #FFFFFF; border: none;
    border-radius: 12px; padding: 0.62rem 1.5rem;
    font-weight: 600; font-size: 14px; letter-spacing: -0.01em;
    box-shadow: 0 2px 10px rgba(0,122,255,0.28);
    transition: background .15s ease, transform .1s ease;
}}
.stDownloadButton button:hover, .stButton button:hover {{
    background: #0A66D0; color: #FFFFFF; transform: translateY(-1px);
}}
[data-testid="stExpander"] {{
    background: #FFFFFF; border: 1px solid rgba(0,0,0,0.06) !important;
    border-radius: 14px !important; box-shadow: 0 1px 6px rgba(0,0,0,0.035);
}}
[data-testid="stFileUploader"] section {{
    background: #FFFFFF; border: 1.5px dashed rgba(0,122,255,0.35);
    border-radius: 14px;
}}
@media (prefers-reduced-motion: reduce) {{
    .kpi-card, .stDownloadButton button, .stButton button {{ transition: none; }}
}}
</style>
""", unsafe_allow_html=True)

# ============================================================
# HELPERS DE FORMAT (convention SmartBuyer)
# ============================================================
def fmt(n):
    if n is None or (isinstance(n, float) and (pd.isna(n) or not np.isfinite(n))):
        return "—"
    a = abs(n)
    if a >= 1_000_000: return f"{n/1_000_000:.1f} M"
    if a >= 1_000:     return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def fmt_pct(v, dec=1):
    if v is None or pd.isna(v): return "—"
    return f"{v:.{dec}f}%"

def fmt_delta(v):
    if v is None or pd.isna(v): return "—"
    return f"{v:+.1f} pts"

def rayon_key(libelle):
    s = str(libelle).upper()
    for k in CIBLES_DEFAUT:
        if k in s:
            return k
    return None

def kpi_card(label, value, sub=None, sub_class="neutral"):
    sub_html = f"<div class='kpi-sub'><span class='pill {sub_class}'>{sub}</span></div>" if sub else ""
    return (f"<div class='kpi-card'><div class='kpi-label'>{label}</div>"
            f"<div class='kpi-value'>{value}</div>{sub_html}</div>")

def _wavg(values, weights):
    """Moyenne pondérée robuste (NaN-safe). Renvoie NaN si aucun poids valide."""
    v = pd.Series(values, dtype=float)
    w = pd.Series(weights, dtype=float)
    m = v.notna() & w.notna() & (w > 0)
    if not m.any() or w[m].sum() == 0:
        return np.nan
    return float((v[m] * w[m]).sum() / w[m].sum())

# ============================================================
# CHARGEMENT — EXPORT ARTICLE UNIQUE (Rayon → Famille → Sous Famille → Article [× Site])
# ============================================================
@st.cache_data(show_spinner=False)
def load_export(file_bytes):
    raw = pd.read_excel(io.BytesIO(file_bytes))
    raw.columns = [str(c).lstrip('\ufeff').strip() for c in raw.columns]

    perimetre = None
    note_rows = raw[raw['Rayon'].astype(str).str.startswith('Filtres', na=False)]
    if not note_rows.empty:
        perimetre = str(note_rows.iloc[0]['Rayon'])

    df = raw[raw['Rayon'].notna()].copy()
    df = df[~df['Rayon'].astype(str).str.startswith('Filtres')]
    if 'Article' not in df.columns:
        df['Article'] = np.nan

    # Auto-détection de la colonne Site (export à la maille Article × Site)
    site_col = None
    for cand in SITE_CANDIDATES:
        if cand in df.columns:
            site_col = cand
            break
    if site_col is None:
        for c in df.columns:
            cu = str(c).upper()
            if 'SITE' in cu or 'MAGASIN' in cu:
                site_col = c
                break
    return df, perimetre, site_col

# ============================================================
# AGRÉGATS RÉSEAU (robustes mono/multi-site)
# ============================================================
def kpis_globaux_rayon(df):
    g = df[df['Rayon'] == 'Total']
    if g.empty:
        return None
    # Si l'export contient un Total par site, on somme (mono-site : identique).
    ca = g['CA'].sum()
    ca_n1 = g['CA N-1'].sum()
    marge = g['Marge'].sum()
    evol_marge_pct = g.get('%Vs N-1.1', pd.Series(np.nan, index=g.index))
    with np.errstate(divide='ignore', invalid='ignore'):
        marge_n1_rows = (g['Marge'] / (1 + evol_marge_pct)).replace([np.inf, -np.inf], np.nan)
    marge_n1 = marge_n1_rows.sum() if marge_n1_rows.notna().any() else np.nan
    qte = g.get('Qté Vente', pd.Series(np.nan, index=g.index)).sum()
    qte_n1 = g.get('Qté Vente N-1', pd.Series(np.nan, index=g.index)).sum()
    poids_promo = _wavg(g.get('%CA Poids Promo', pd.Series(np.nan, index=g.index)), g['CA'])
    casse = g.get('Casse (Valeur)', pd.Series(np.nan, index=g.index)).sum()
    return {
        'ca': ca, 'ca_n1': ca_n1, 'evol_ca': ca/ca_n1 - 1 if ca_n1 else np.nan,
        'marge': marge, 'marge_n1': marge_n1,
        'tx_marge': marge/ca*100 if ca else np.nan,
        'tx_marge_n1': marge_n1/ca_n1*100 if ca_n1 and pd.notna(marge_n1) else np.nan,
        'qte': qte, 'qte_n1': qte_n1,
        'poids_promo': poids_promo * 100 if pd.notna(poids_promo) else np.nan,
        'casse': casse,
    }

def perf_par_rayon(df, cibles):
    sub = df[(df['Famille'] == 'Total') & (df['Rayon'] != 'Total')].copy()
    sub['Rayon_aff'] = sub['Rayon'].astype(str).str.split(' - ').str[-1].str.strip()
    if 'Qté Vente' not in sub.columns: sub['Qté Vente'] = np.nan
    if 'Qté Vente N-1' not in sub.columns: sub['Qté Vente N-1'] = np.nan
    agg = sub.groupby('Rayon_aff', as_index=False).agg(
        CA=('CA', 'sum'), CA_N1=('CA N-1', 'sum'), Marge=('Marge', 'sum'),
        Qte=('Qté Vente', 'sum'), Qte_N1=('Qté Vente N-1', 'sum'))
    rows = []
    for _, r in agg.iterrows():
        key = rayon_key(r['Rayon_aff'])
        cible = cibles.get(key, CIBLE_FALLBACK) if key else CIBLE_FALLBACK
        tx = r['Marge']/r['CA']*100 if r['CA'] else np.nan
        rows.append({
            'Rayon': r['Rayon_aff'],
            'CA': r['CA'],
            'Évol CA %': (r['CA']/r['CA_N1']-1)*100 if r['CA_N1'] else np.nan,
            'Évol Qté %': (r['Qte']/r['Qte_N1']-1)*100 if r['Qte_N1'] else np.nan,
            'Taux Marge %': tx, 'Objectif %': cible,
            'Écart (pts)': tx - cible if pd.notna(tx) else np.nan,
        })
    return pd.DataFrame(rows).sort_values('CA', ascending=False)

def family_metrics(df):
    """Métriques niveau Famille, agrégées réseau (robuste aux exports multi-site)."""
    sub = df[(df['Sous Famille'] == 'Total') & (df['Famille'] != 'Total') & (df['Rayon'] != 'Total')].copy()
    sub['Rayon_aff'] = sub['Rayon'].astype(str).str.split(' - ').str[-1].str.strip()
    sub['Famille_aff'] = sub['Famille'].astype(str).str.split(' - ').str[-1].str.strip()
    for c in ['CA', 'CA N-1', 'Marge']:
        sub[c] = sub[c].fillna(0)
    for c in ['Qté Vente', 'Qté Vente N-1', 'Casse (Valeur)', '%CA Poids Promo',
              '%Marge Promo', '%Marge Hors Promo', '%Vs N-1.1']:
        if c not in sub.columns:
            sub[c] = np.nan

    # Marge N-1 par ligne (ratio non sommable, à calculer AVANT agrégation)
    with np.errstate(divide='ignore', invalid='ignore'):
        sub['_marge_n1'] = (sub['Marge'] / (1 + sub['%Vs N-1.1'])).replace([np.inf, -np.inf], np.nan)
    # CA promo par ligne (pour re-pondérer les % promo après agrégation)
    sub['_ca_promo'] = sub['%CA Poids Promo'] * sub['CA']
    sub['_ca_hp'] = sub['CA'] - sub['_ca_promo'].fillna(0)
    sub['_mp_w'] = sub['%Marge Promo'] * sub['_ca_promo']
    sub['_mhp_w'] = sub['%Marge Hors Promo'] * sub['_ca_hp']

    agg = sub.groupby(['Rayon_aff', 'Famille_aff'], as_index=False).agg(**{
        'CA': ('CA', 'sum'), 'CA N-1': ('CA N-1', 'sum'), 'Marge': ('Marge', 'sum'),
        '_marge_n1': ('_marge_n1', 'sum'),
        'Qté Vente': ('Qté Vente', 'sum'), 'Qté Vente N-1': ('Qté Vente N-1', 'sum'),
        'Casse (Valeur)': ('Casse (Valeur)', 'sum'),
        '_ca_promo': ('_ca_promo', 'sum'), '_ca_hp': ('_ca_hp', 'sum'),
        '_mp_w': ('_mp_w', 'sum'), '_mhp_w': ('_mhp_w', 'sum'),
    })
    agg['Perte CA'] = agg['CA'] - agg['CA N-1']
    agg['Évol CA %'] = np.where(agg['CA N-1'] > 0, (agg['CA']/agg['CA N-1']-1)*100, np.nan)
    agg['Tx Marge %'] = np.where(agg['CA'] > 0, agg['Marge']/agg['CA']*100, np.nan)
    agg['Tx Marge N-1 %'] = np.where(agg['CA N-1'] > 0, agg['_marge_n1']/agg['CA N-1']*100, np.nan)
    agg['Écart Tx Marge (pts)'] = agg['Tx Marge %'] - agg['Tx Marge N-1 %']
    agg['Évol Qté %'] = np.where(agg['Qté Vente N-1'] > 0, (agg['Qté Vente']/agg['Qté Vente N-1']-1)*100, np.nan)
    with np.errstate(divide='ignore', invalid='ignore'):
        agg['%CA Poids Promo'] = np.where(agg['CA'] > 0, agg['_ca_promo']/agg['CA'], np.nan)
        agg['%Marge Promo'] = np.where(agg['_ca_promo'] > 0, agg['_mp_w']/agg['_ca_promo'], np.nan)
        agg['%Marge Hors Promo'] = np.where(agg['_ca_hp'] > 0, agg['_mhp_w']/agg['_ca_hp'], np.nan)
        agg['%Casse (Valeur)'] = np.where(agg['CA'] > 0, agg['Casse (Valeur)']/agg['CA'], np.nan)
    return agg.drop(columns=['_marge_n1', '_ca_promo', '_ca_hp', '_mp_w', '_mhp_w'])

def build_headline(k, perf, fam):
    evol_ca = k['evol_ca'] * 100 if pd.notna(k['evol_ca']) else np.nan
    evo_tx = k['tx_marge'] - k['tx_marge_n1'] if pd.notna(k['tx_marge_n1']) else np.nan
    evol_qte = (k['qte']/k['qte_n1'] - 1) * 100 if k['qte_n1'] else np.nan
    pct_casse = k['casse']/k['ca']*100 if k['ca'] else np.nan

    line1 = (
        f"CA {fmt(k['ca'])} FCFA ({evol_ca:+.1f}%) &nbsp;·&nbsp; "
        f"Marge {fmt(k['marge'])} FCFA, taux {k['tx_marge']:.1f}% ({fmt_delta(evo_tx)}) &nbsp;·&nbsp; "
        f"Qté {fmt(k['qte'])} ({evol_qte:+.1f}%) &nbsp;·&nbsp; "
        f"Casse {pct_casse:.2f}% du CA &nbsp;·&nbsp; "
        f"Promo {k['poids_promo']:.1f}% du CA"
    )
    bits = []
    if perf is not None and perf['Écart (pts)'].notna().any():
        wr = perf.loc[perf['Écart (pts)'].idxmin()]
        bits.append(f"rayon à surveiller : <b>{wr['Rayon']}</b> ({fmt_delta(wr['Écart (pts)'])} vs objectif)")
    if fam is not None and len(fam) and fam['Perte CA'].notna().any():
        wf = fam.loc[fam['Perte CA'].idxmin()]
        fam_qte = wf.get('Évol Qté %', np.nan)
        fam_txt = (f"famille à risque : <b>{wf['Famille_aff']}</b> — "
                   f"CA {fmt(wf['CA'])} FCFA ({wf['Évol CA %']:+.1f}%), "
                   f"marge {wf['Tx Marge %']:.1f}%")
        if pd.notna(fam_qte):
            fam_txt += f", qté {fam_qte:+.1f}%"
        bits.append(fam_txt)
    line2 = "📌 " + " &nbsp;·&nbsp; ".join(bits) if bits else ""
    return line1, line2

def top_familles(df, n=5, by='perte_ca'):
    sub = family_metrics(df)
    if by == 'perte_ca':
        out = sub.nsmallest(n, 'Perte CA')[['Rayon_aff','Famille_aff','CA','CA N-1','Perte CA','Tx Marge %']]
    elif by == 'casse':
        out = sub.nsmallest(n, 'Casse (Valeur)')[['Rayon_aff','Famille_aff','CA','Casse (Valeur)','%Casse (Valeur)']]
    elif by == 'promo':
        mat = sub[sub['CA'] > 1_000_000]
        out = mat.nlargest(n, '%CA Poids Promo')[['Rayon_aff','Famille_aff','CA','%CA Poids Promo','%Marge Promo','%Marge Hors Promo']]
    return out.reset_index(drop=True)

def top_flop_table(sub, metric, n, mode, cols, ca_floor=0, directional=True):
    base = sub[sub['CA'] > ca_floor] if ca_floor else sub
    base = base[base[metric].notna()]
    if directional:
        base = base[base[metric] > 0] if mode == 'top' else base[base[metric] < 0]
    out = base.nlargest(n, metric) if mode == 'top' else base.nsmallest(n, metric)
    return out[['Rayon_aff','Famille_aff'] + cols].reset_index(drop=True)

# ============================================================
# VUE ARTICLE — 🔧 FIX DOUBLONS
# L'export PBI est à la maille Article × Site : un même article vendu dans
# plusieurs magasins produit plusieurs lignes. On agrège au niveau réseau
# AVANT tout classement, sinon les Top/Flop contiennent des doublons.
# ============================================================
def prep_articles(df):
    art = df[df['Article'].notna() & (df['Article'] != 'Total') & (df['Rayon'] != 'Total')].copy()
    art['Rayon_aff'] = art['Rayon'].astype(str).str.split(' - ').str[-1].str.strip()
    art['Famille_aff'] = art['Famille'].astype(str).str.split(' - ').str[-1].str.strip()
    art['SousFamille_aff'] = art['Sous Famille'].astype(str).str.split(' - ').str[-1].str.strip()
    art['Article_aff'] = art['Article'].astype(str)
    for c in ['CA', 'CA N-1', 'Marge', 'Qté Vente', 'Qté Vente N-1']:
        if c not in art.columns: art[c] = 0
        art[c] = art[c].fillna(0)

    # Marge N-1 recalculée AVANT agrégation (le % N-1 est un ratio par ligne)
    evol_marge = art.get('%Vs N-1.1', pd.Series(np.nan, index=art.index))
    with np.errstate(divide='ignore', invalid='ignore'):
        art['Marge N-1 (calc)'] = (art['Marge'] / (1 + evol_marge)).replace([np.inf, -np.inf], np.nan)

    # 🔧 Agrégation réseau : 1 article = 1 ligne, tous sites confondus
    group_cols = ['Rayon_aff', 'Famille_aff', 'SousFamille_aff', 'Article_aff']
    sum_cols = ['CA', 'CA N-1', 'Marge', 'Marge N-1 (calc)', 'Qté Vente', 'Qté Vente N-1']
    art = art.groupby(group_cols, as_index=False)[sum_cols].sum(min_count=1)
    art['Marge N-1 (calc)'] = art['Marge N-1 (calc)'].where(art['Marge N-1 (calc)'].notna(), np.nan)
    for c in ['CA', 'CA N-1', 'Marge', 'Qté Vente', 'Qté Vente N-1']:
        art[c] = art[c].fillna(0)

    art['Tx Marge %'] = np.where(art['CA'] > 0, art['Marge']/art['CA']*100, np.nan)
    art['Tx Marge N-1 % (calc)'] = np.where(art['CA N-1'] > 0, art['Marge N-1 (calc)']/art['CA N-1']*100, np.nan)
    art['Écart Tx Marge (pts)'] = art['Tx Marge %'] - art['Tx Marge N-1 % (calc)']
    art['Gain Marge (FCFA)'] = art['Marge'] - art['Marge N-1 (calc)']
    art['Variation Qté'] = art['Qté Vente'] - art['Qté Vente N-1']
    art['Perte CA (FCFA)'] = art['CA'] - art['CA N-1']
    return art

def destructeurs_performeurs(art, n=15, seuil_ca=100_000):
    res = {}
    res['A_marge_neg'] = art[art['Marge'] < 0].nsmallest(n, 'Marge')
    pos = art[art['Marge'] >= 0].copy()
    deg = pos[pos['Écart Tx Marge (pts)'].notna() & (pos['Écart Tx Marge (pts)'] < 0)]
    res['B_degrad_marge'] = deg.nsmallest(n, 'Écart Tx Marge (pts)')
    mat = art[art['CA'] > seuil_ca]
    gain = mat[mat['Gain Marge (FCFA)'].notna() & (mat['Gain Marge (FCFA)'] > 0)]
    res['C_perf_gain_marge'] = gain.nlargest(n, 'Gain Marge (FCFA)')
    mat_n1 = mat[mat['CA N-1'] > 0].assign(_evol=lambda d: d['CA']/d['CA N-1']-1)
    mat_n1_pos = mat_n1[mat_n1['_evol'] > 0]
    res['D_croissance_ca'] = mat_n1_pos.nlargest(n, '_evol')
    baisse = art[art['Perte CA (FCFA)'] < 0]
    res['E_baisse_ca'] = baisse.nsmallest(n, 'Perte CA (FCFA)')
    hausse_q = art[art['Variation Qté'] > 0]
    res['F_hausse_qte'] = hausse_q.nlargest(n, 'Variation Qté')
    baisse_q = art[art['Variation Qté'] < 0]
    res['G_baisse_qte'] = baisse_q.nsmallest(n, 'Variation Qté')
    return res

# ============================================================
# 🆕 MARGE NÉGATIVE PAR SITE
# Exploite la dimension Site de l'export (celle qui créait les doublons).
# ============================================================
def marge_negative_par_site(df, site_col):
    """Synthèse par site : CA site, nb articles en marge négative, marge négative
    cumulée, poids vs CA site. Renvoie None si la colonne Site est absente."""
    if not site_col or site_col not in df.columns:
        return None
    art = df[df['Article'].notna() & (df['Article'] != 'Total') & (df['Rayon'] != 'Total')].copy()
    art = art[art[site_col].notna() & (art[site_col].astype(str) != 'Total')]
    if art.empty:
        return None
    art['CA'] = art['CA'].fillna(0)
    art['Marge'] = art['Marge'].fillna(0)

    ca_site = art.groupby(site_col, as_index=False)['CA'].sum().rename(columns={'CA': 'CA Site'})
    neg = art[art['Marge'] < 0]
    g = neg.groupby(site_col, as_index=False).agg(
        **{'Nb Articles Nég.': ('Marge', 'count'), 'Marge Négative (FCFA)': ('Marge', 'sum')})
    out = ca_site.merge(g, on=site_col, how='left')
    out['Nb Articles Nég.'] = out['Nb Articles Nég.'].fillna(0).astype(int)
    out['Marge Négative (FCFA)'] = out['Marge Négative (FCFA)'].fillna(0)
    out['% Marge Nég / CA Site'] = np.where(out['CA Site'] > 0,
                                            out['Marge Négative (FCFA)']/out['CA Site']*100, np.nan)
    out = out.rename(columns={site_col: 'Site'})
    return out.sort_values('Marge Négative (FCFA)').reset_index(drop=True)

def _site_format(site):
    """Format magasin dérivé du libellé site (Hyper / Market / Supeco)."""
    s = str(site).upper()
    if 'HYPER' in s: return 'Hyper'
    if 'MARKET' in s: return 'Market'
    if 'SUPECO' in s: return 'Supeco'
    return '—'

def detail_marge_neg_site(df, site_col):
    """Détail EXHAUSTIF des articles en marge négative, ligne article × site
    (format du récap de référence 'Marges Négatives par Site').
    Colonnes : Code Art. / Article / Rayon / Famille / Magasin / Format /
    CA / Marge / Tx Marge % / Qté — trié par libellé article."""
    if not site_col or site_col not in df.columns:
        return None
    art = df[df['Article'].notna() & (df['Article'] != 'Total') & (df['Rayon'] != 'Total')].copy()
    art = art[art[site_col].notna() & (art[site_col].astype(str) != 'Total')]
    art['CA'] = art['CA'].fillna(0)
    art['Marge'] = art['Marge'].fillna(0)
    art = art[art['Marge'] < 0].copy()
    if art.empty:
        return None
    art['Rayon_aff'] = art['Rayon'].astype(str).str.split(' - ').str[-1].str.strip()
    art['Famille_aff'] = art['Famille'].astype(str).str.split(' - ').str[-1].str.strip()
    # Séparation "code - libellé" de l'export PBI
    art_str = art['Article'].astype(str)
    has_sep = art_str.str.contains(' - ')
    art['Code Art.'] = np.where(has_sep, art_str.str.split(' - ').str[0].str.strip(), '')
    art['Article_aff'] = np.where(has_sep,
                                  art_str.str.split(' - ', n=1).str[1].str.strip(), art_str)
    art['Site'] = art[site_col].astype(str).str.split(' - ').str[-1].str.strip()
    art['Format'] = art['Site'].map(_site_format)
    art['Tx Marge %'] = np.where(art['CA'] > 0, art['Marge']/art['CA']*100, np.nan)
    qte = art.get('Qté Vente', pd.Series(np.nan, index=art.index))
    art['Qté'] = qte.fillna(0)
    return art.sort_values('Article_aff')[
        ['Code Art.', 'Article_aff', 'Rayon_aff', 'Famille_aff', 'Site', 'Format',
         'CA', 'Marge', 'Tx Marge %', 'Qté']].reset_index(drop=True)

DISPLAY_COLS = ['Rayon_aff','Famille_aff','SousFamille_aff','Article_aff']
DISPLAY_RENAME = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article'}

def show_table(d, extra_cols, fmt_map=None):
    if d.empty:
        st.caption("Aucune ligne ne correspond à ce critère sur la période.")
        return
    cols = DISPLAY_COLS + extra_cols
    disp = d[cols].rename(columns=DISPLAY_RENAME).reset_index(drop=True)
    disp.index = disp.index + 1
    if fmt_map:
        for c, f in fmt_map.items():
            if c in disp.columns:
                disp[c] = disp[c].map(f)
    st.dataframe(disp, use_container_width=True)

# ============================================================
# EXPORT EXCEL — 3 feuilles : Dashboard COPIL · Destructeurs & Performeurs ·
# Marge Nég par Site
# ============================================================
def build_excel_full(k, perf, fam, n_top, art_res, perimetre, mns=None, mns_detail=None):
    BLUE_H = "FF007AFF"; DARK_H = "FF1D1D1F"; RED_H = "FFFF3B30"; GREEN_H = "FF34C759"
    WHITE_H = "FFFFFFFF"; LGREY_H = "FFF2F2F7"; ARIAL = "Arial"
    thin = Side(style="thin", color="FFD1D1D6")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    ACC = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
    ACC_SIGNED = '_-* +#,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
    QTY = "#,##0"
    PTS = '+0.00" pts";-0.00" pts"'
    PCT = "0.0%"
    PCT2 = "0.00%"

    def fmt_value(kind, v):
        if kind == 'pct100' and v is not None and pd.notna(v):
            return v / 100
        return v

    def fmt_code(kind):
        return {'amount': ACC, 'amount_signed': ACC_SIGNED, 'qty': QTY,
                'pct100': PCT, 'pct1': PCT, 'pts': PTS}.get(kind, "General")

    def section_bar(ws, row, ncols, text, color=DARK_H):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        c = ws.cell(row=row, column=1, value=text)
        c.font = Font(name=ARIAL, bold=True, size=11, color=WHITE_H)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 20

    def header_row(ws, row, labels):
        for i, lbl in enumerate(labels, start=1):
            c = ws.cell(row=row, column=i, value=lbl)
            c.font = Font(name=ARIAL, bold=True, size=10, color=WHITE_H)
            c.fill = PatternFill("solid", fgColor=BLUE_H)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def data_row(ws, row, values, zebra=False, left_cols=()):
        fillc = LGREY_H if zebra else "FFFFFFFF"
        for i, v in enumerate(values, start=1):
            c = ws.cell(row=row, column=i, value=v)
            c.font = Font(name=ARIAL, size=10)
            c.border = box
            c.fill = PatternFill("solid", fgColor=fillc)
            c.alignment = Alignment(horizontal="left" if i in left_cols else "center")

    def autosize(ws, widths):
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    wb = Workbook()

    # ============== FEUILLE 1 : DASHBOARD COPIL ==============
    ws = wb.active
    ws.title = "Dashboard COPIL"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:H1")
    ws["A1"] = "DASHBOARD COPIL HEBDO — PGC (Épicerie · Boissons · Droguerie · Parfumerie Hygiène)"
    ws["A1"].font = Font(name=ARIAL, bold=True, size=14, color=WHITE_H)
    ws["A1"].fill = PatternFill("solid", fgColor=BLUE_H)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26
    ws["A2"] = "Période :"; ws["A2"].font = Font(name=ARIAL, bold=True, size=10)
    ws["B2"] = "Voir export source"; ws["B2"].font = Font(name=ARIAL, size=10, color="FF0000FF", bold=True)
    if perimetre:
        ws["A3"] = "Périmètre :"; ws["A3"].font = Font(name=ARIAL, size=9, italic=True, color="FF86868B")
        ws["B3"] = perimetre.replace("\n", " ")[:200]
        ws["B3"].font = Font(name=ARIAL, size=9, italic=True, color="FF86868B")

    r = 5
    section_bar(ws, r, 8, "1.  VUE D'ENSEMBLE RÉSEAU"); r += 1
    header_row(ws, r, ["Indicateur", "Cette semaine", "N-1", "Évolution"]); r += 1
    evo_tx = k['tx_marge'] - k['tx_marge_n1'] if pd.notna(k['tx_marge_n1']) else None
    evol_qte = (k['qte']/k['qte_n1']-1) if k['qte_n1'] else None
    pct_casse = k['casse']/k['ca'] if k['ca'] else None
    kpi_rows = [
        ("CA (FCFA)",          k['ca'],            k['ca_n1'],   k['evol_ca'], 'amount', 'pct1'),
        ("Marge (FCFA)",       k['marge'],         k['marge_n1'],None,         'amount', None),
        ("Taux de marge",      k['tx_marge']/100,  (k['tx_marge_n1']/100 if pd.notna(k['tx_marge_n1']) else None), evo_tx, 'pct1', 'pts'),
        ("Qté vendue",         k['qte'],           k['qte_n1'],  evol_qte,     'qty',    'pct1'),
        ("Poids Promo (% CA)", k['poids_promo']/100 if pd.notna(k['poids_promo']) else None, None, None, 'pct1', None),
        ("Casse (FCFA)",       k['casse'],         None,         pct_casse,    'amount', 'pct1'),
    ]
    r0kpi = r
    for i, (label, v, n1, evo, kind_v, kind_evo) in enumerate(kpi_rows):
        zebra = i % 2 == 1
        fillc = LGREY_H if zebra else "FFFFFFFF"
        for col in range(1, 5):
            cc = ws.cell(row=r, column=col)
            cc.fill = PatternFill("solid", fgColor=fillc); cc.border = box; cc.font = Font(name=ARIAL, size=10)
        ws.cell(row=r, column=1, value=label).alignment = Alignment(horizontal="left", indent=1)
        ws.cell(row=r, column=2, value=v); ws.cell(row=r, column=2).number_format = fmt_code(kind_v)
        if n1 is not None and pd.notna(n1):
            ws.cell(row=r, column=3, value=n1); ws.cell(row=r, column=3).number_format = fmt_code(kind_v)
        if evo is not None and pd.notna(evo):
            ws.cell(row=r, column=4, value=evo)
            ws.cell(row=r, column=4).number_format = fmt_code(kind_evo)
        r += 1
    ws.conditional_formatting.add(f"D{r0kpi}:D{r-1}", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="FFFFD6D4")))
    ws.conditional_formatting.add(f"D{r0kpi}:D{r-1}", CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PatternFill("solid", fgColor="FFD7F5DE")))
    r += 1

    section_bar(ws, r, 8, "2.  PERFORMANCE PAR RAYON VS OBJECTIFS MARGE (MÉTI)"); r += 1
    header_row(ws, r, ["Rayon", "CA (FCFA)", "Évol CA", "Évol Qté", "Taux Marge", "Objectif Méti", "Écart (pts)"]); r += 1
    r0r = r
    for i, (_, row_) in enumerate(perf.iterrows()):
        data_row(ws, r, [row_['Rayon'], row_['CA'],
                          (row_['Évol CA %']/100 if pd.notna(row_['Évol CA %']) else None),
                          (row_['Évol Qté %']/100 if pd.notna(row_['Évol Qté %']) else None),
                          (row_['Taux Marge %']/100 if pd.notna(row_['Taux Marge %']) else None),
                          row_['Objectif %']/100, row_['Écart (pts)']], zebra=(i%2==1), left_cols=(1,))
        for col, fmt_ in [(2,ACC),(3,PCT),(4,PCT),(5,PCT2),(6,PCT),(7,PTS)]:
            ws.cell(row=r, column=col).number_format = fmt_
        r += 1
    ws.conditional_formatting.add(f"G{r0r}:G{r-1}", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="FFFFD6D4")))
    ws.conditional_formatting.add(f"G{r0r}:G{r-1}", CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PatternFill("solid", fgColor="FFD7F5DE")))
    r += 1

    # ---- 🆕 3. MARGE NÉGATIVE PAR SITE (synthèse) ----
    if mns is not None and not mns.empty:
        section_bar(ws, r, 8, "3.  MARGE NÉGATIVE PAR SITE", color=RED_H); r += 1
        header_row(ws, r, ["Site", "CA Site (FCFA)", "Nb Articles Nég.", "Marge Négative (FCFA)", "% vs CA Site"]); r += 1
        r0s = r
        for i, (_, row_) in enumerate(mns.iterrows()):
            data_row(ws, r, [row_['Site'], row_['CA Site'], row_['Nb Articles Nég.'],
                              row_['Marge Négative (FCFA)'],
                              (row_['% Marge Nég / CA Site']/100 if pd.notna(row_['% Marge Nég / CA Site']) else None)],
                      zebra=(i % 2 == 1), left_cols=(1,))
            for col, fmt_ in [(2, ACC), (3, QTY), (4, ACC), (5, PCT2)]:
                ws.cell(row=r, column=col).number_format = fmt_
            r += 1
        ws.conditional_formatting.add(f"D{r0s}:D{r-1}",
            CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="FFFFD6D4")))
        r += 1

    section_bar(ws, r, 8, "4.  DÉTAIL PAR FAMILLE — TOUTES FAMILLES (sans objectif)"); r += 1
    header_row(ws, r, ["Rayon", "Famille", "CA", "Évol CA %", "Marge", "Taux Marge %", "Qté Vente", "Évol Qté %"]); r += 1
    fam_sorted = fam.sort_values(['Rayon_aff', 'CA'], ascending=[True, False])
    for i, (_, row_) in enumerate(fam_sorted.iterrows()):
        data_row(ws, r, [row_['Rayon_aff'], row_['Famille_aff'], row_['CA'],
                          (row_['Évol CA %']/100 if pd.notna(row_['Évol CA %']) else None), row_['Marge'],
                          (row_['Tx Marge %']/100 if pd.notna(row_['Tx Marge %']) else None), row_['Qté Vente'],
                          (row_['Évol Qté %']/100 if pd.notna(row_['Évol Qté %']) else None)], zebra=(i%2==1), left_cols=(1,2))
        for col, fmt_ in [(3,ACC),(4,PCT),(5,ACC),(6,PCT),(7,QTY),(8,PCT)]:
            ws.cell(row=r, column=col).number_format = fmt_
        r += 1
    r += 1

    def top_section(title, dframe, cols_map, kinds, color=RED_H):
        nonlocal r
        section_bar(ws, r, 8, title, color=color); r += 1
        header_row(ws, r, ["Rang"] + list(cols_map.values())); r += 1
        keys = list(cols_map.keys())
        for i, (_, row_) in enumerate(dframe.iterrows()):
            vals = [i+1] + [fmt_value(kinds.get(c, 'amount'), row_[c]) for c in keys]
            data_row(ws, r, vals, zebra=(i%2==1), left_cols=(2,3))
            for j, ck in enumerate(keys, start=2):
                ws.cell(row=r, column=j).number_format = fmt_code(kinds.get(ck, 'amount'))
            r += 1
        r += 1

    top_section(f"5.  FLOP {n_top} — PLUS FORTE BAISSE DE CA (par Famille)",
                top_flop_table(fam, 'Perte CA', n_top, 'flop', ['CA','CA N-1','Évol CA %','Perte CA','Tx Marge %']),
                {'Rayon_aff':'Rayon','Famille_aff':'Famille','CA':'CA (FCFA)','CA N-1':'CA N-1','Évol CA %':'Évol %','Perte CA':'Perte (FCFA)','Tx Marge %':'Taux Marge'},
                {'CA':'amount','CA N-1':'amount','Évol CA %':'pct100','Perte CA':'amount','Tx Marge %':'pct100'})

    top_section(f"6.  TOP {n_top} — MEILLEUR GAIN DE CA (par Famille)",
                top_flop_table(fam, 'Perte CA', n_top, 'top', ['CA','CA N-1','Évol CA %','Perte CA','Tx Marge %']),
                {'Rayon_aff':'Rayon','Famille_aff':'Famille','CA':'CA (FCFA)','CA N-1':'CA N-1','Évol CA %':'Évol %','Perte CA':'Gain (FCFA)','Tx Marge %':'Taux Marge'},
                {'CA':'amount','CA N-1':'amount','Évol CA %':'pct100','Perte CA':'amount','Tx Marge %':'pct100'},
                color=GREEN_H)

    top_section(f"7.  FLOP {n_top} — CASSE EN VALEUR (par Famille)",
                top_familles_for_excel(fam, n_top, 'casse'),
                {'Rayon_aff':'Rayon','Famille_aff':'Famille','CA':'CA (FCFA)','Casse (Valeur)':'Casse (FCFA)','%Casse (Valeur)':'% Casse'},
                {'CA':'amount','Casse (Valeur)':'amount','%Casse (Valeur)':'pct1'})

    top_section(f"8.  TOP {n_top} — POIDS PROMO LE PLUS ÉLEVÉ (Famille, CA > 1M)",
                top_familles_for_excel(fam, n_top, 'promo'),
                {'Rayon_aff':'Rayon','Famille_aff':'Famille','CA':'CA (FCFA)','%CA Poids Promo':'Poids Promo','%Marge Promo':'Tx M. Promo','%Marge Hors Promo':'Tx M. HP'},
                {'CA':'amount','%CA Poids Promo':'pct1','%Marge Promo':'pct1','%Marge Hors Promo':'pct1'},
                color="FFFF9500")

    autosize(ws, {'A':22,'B':26,'C':16,'D':13,'E':16,'F':13,'G':13,'H':13})
    ws.freeze_panes = "A6"

    # ============== FEUILLE 2 : DESTRUCTEURS & PERFORMEURS ==============
    ws2 = wb.create_sheet("Destructeurs Performeurs")
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells("A1:I1")
    ws2["A1"] = "DESTRUCTEURS & PERFORMEURS — NIVEAU ARTICLE (agrégé réseau, tous sites)"
    ws2["A1"].font = Font(name=ARIAL, bold=True, size=14, color=WHITE_H)
    ws2["A1"].fill = PatternFill("solid", fgColor=BLUE_H)
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws2.row_dimensions[1].height = 26

    r2 = 3
    def art_section(title, dframe, cols_map, kinds, color):
        nonlocal r2
        ws2.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=9)
        c = ws2.cell(row=r2, column=1, value=title)
        c.font = Font(name=ARIAL, bold=True, size=11, color=WHITE_H)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws2.row_dimensions[r2].height = 20
        r2 += 1
        for i, lbl in enumerate(["Rang"] + list(cols_map.values()), start=1):
            cell = ws2.cell(row=r2, column=i, value=lbl)
            cell.font = Font(name=ARIAL, bold=True, size=10, color=WHITE_H)
            cell.fill = PatternFill("solid", fgColor=BLUE_H)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        r2 += 1
        keys = list(cols_map.keys())
        for i, (_, row_) in enumerate(dframe.iterrows()):
            vals = [i+1] + [fmt_value(kinds.get(c, 'amount'), row_.get(c, None)) for c in keys]
            fillc = LGREY_H if i % 2 == 1 else "FFFFFFFF"
            for j, v in enumerate(vals, start=1):
                cell = ws2.cell(row=r2, column=j, value=v)
                cell.font = Font(name=ARIAL, size=10)
                cell.border = box
                cell.fill = PatternFill("solid", fgColor=fillc)
                cell.alignment = Alignment(horizontal="left" if j in (2,3,4,5) else "center")
                if j >= 6:
                    cell.number_format = fmt_code(kinds.get(keys[j-2], 'amount'))
            r2 += 1
        r2 += 1

    cm_neg = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','CA':'CA','Marge':'Marge','Tx Marge %':'Taux Marge'}
    art_section(f"A.  FLOP {n_top} — ARTICLES EN MARGE NÉGATIVE", art_res['A_marge_neg'], cm_neg,
                {'CA':'amount','Marge':'amount','Tx Marge %':'pct100'}, RED_H)

    cm_deg = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','CA':'CA','Tx Marge %':'Taux Marge','Écart Tx Marge (pts)':'Écart pts'}
    art_section(f"B.  FLOP {n_top} — DÉGRADATION DU TAUX DE MARGE (marge encore positive)", art_res['B_degrad_marge'], cm_deg,
                {'CA':'amount','Tx Marge %':'pct100','Écart Tx Marge (pts)':'pts'}, RED_H)

    cm_gain = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','CA':'CA','Gain Marge (FCFA)':'Gain Marge','Tx Marge %':'Taux Marge'}
    art_section(f"C.  TOP {n_top} — PERFORMEURS : GAIN DE MARGE EN VALEUR", art_res['C_perf_gain_marge'], cm_gain,
                {'CA':'amount','Gain Marge (FCFA)':'amount','Tx Marge %':'pct100'}, GREEN_H)

    d4 = art_res['D_croissance_ca'].copy()
    if not d4.empty:
        d4['Évol CA %'] = (d4['CA']/d4['CA N-1']-1)*100
    cm_croi = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','CA':'CA','CA N-1':'CA N-1','Évol CA %':'Évol %','Tx Marge %':'Taux Marge'}
    art_section(f"D.  TOP {n_top} — PLUS FORTE CROISSANCE DE CA", d4, cm_croi,
                {'CA':'amount','CA N-1':'amount','Évol CA %':'pct100','Tx Marge %':'pct100'}, GREEN_H)

    cm_baisse = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','CA':'CA','CA N-1':'CA N-1','Perte CA (FCFA)':'Perte (FCFA)'}
    art_section(f"E.  FLOP {n_top} — PLUS FORTE BAISSE DE CA", art_res['E_baisse_ca'], cm_baisse,
                {'CA':'amount','CA N-1':'amount','Perte CA (FCFA)':'amount'}, RED_H)

    cm_qte = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','Qté Vente':'Qté Vente','Qté Vente N-1':'Qté N-1','Variation Qté':'Variation','CA':'CA'}
    kinds_qte = {'Qté Vente':'qty','Qté Vente N-1':'qty','Variation Qté':'amount_signed','CA':'amount'}
    art_section(f"F.  TOP {n_top} — PLUS FORTE HAUSSE DE QUANTITÉ VENDUE", art_res['F_hausse_qte'], cm_qte, kinds_qte, GREEN_H)
    art_section(f"G.  FLOP {n_top} — PLUS FORTE BAISSE DE QUANTITÉ VENDUE", art_res['G_baisse_qte'], cm_qte, kinds_qte, RED_H)

    autosize(ws2, {'A':7,'B':20,'C':24,'D':22,'E':38,'F':13,'G':13,'H':13,'I':13})

    # ============== 🆕 FEUILLE 3 : MARGES NÉGATIVES PAR SITE ==============
    # Format aligné sur le récap de référence : bandeau KPI + tableau plat
    # EXHAUSTIF (toutes les lignes article × site en marge négative).
    if mns_detail is not None and not mns_detail.empty:
        NAVY = "FF1B2A4A"; NAVY2 = "FF2E4B7A"; INK = "FF1A1A2E"
        AMBER_F = "FFFFF3E0"; REDT_F = "FFFDECEA"; RED_TX = "FFC0392B"
        CAL = "Calibri"

        ws3 = wb.create_sheet("Marges Négatives par Site")
        ws3.sheet_view.showGridLines = False

        ws3.merge_cells("A1:K1")
        ws3["A1"] = "ARTICLES À MARGE NÉGATIVE — DÉTAIL PAR MAGASIN"
        ws3["A1"].font = Font(name=CAL, bold=True, size=13, color=WHITE_H)
        ws3["A1"].fill = PatternFill("solid", fgColor=NAVY)
        ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws3.row_dimensions[1].height = 24

        ws3.merge_cells("A2:K2")
        per_txt = (perimetre.replace("\n", " ")[:150] if perimetre else "Voir export source")
        ws3["A2"] = f"  Période : {per_txt}"
        ws3["A2"].font = Font(name=CAL, size=9, color="FFAABBCC")
        ws3["A2"].fill = PatternFill("solid", fgColor=NAVY)
        ws3["A2"].alignment = Alignment(horizontal="left", vertical="center")

        # ---- Bandeau KPI (lignes 4-5) ----
        pertes = mns_detail['Marge'].sum()
        kpi_data = [
            ("Lignes article × site", f"{len(mns_detail):,}"),
            ("Articles distincts", f"{mns_detail['Code Art.'].replace('', np.nan).fillna(mns_detail['Article_aff']).nunique():,}"),
            ("Magasins touchés", f"{mns_detail['Site'].nunique():,}"),
            ("Pertes nettes (FCFA)", f"{abs(pertes):,.0f}"),
            ("CA exposé (FCFA)", f"{mns_detail['CA'].sum():,.0f}"),
        ]
        for j, (lbl, val) in enumerate(kpi_data, start=1):
            c4 = ws3.cell(row=4, column=j, value=lbl)
            c4.font = Font(name=CAL, bold=True, size=9, color=WHITE_H)
            c4.fill = PatternFill("solid", fgColor=NAVY2)
            c4.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c5 = ws3.cell(row=5, column=j, value=val)
            c5.font = Font(name=CAL, bold=True, size=12, color=INK)
            c5.fill = PatternFill("solid", fgColor=WHITE_H)
            c5.border = box
            c5.alignment = Alignment(horizontal="center", vertical="center")
        ws3.row_dimensions[4].height = 22
        ws3.row_dimensions[5].height = 20

        # ---- En-tête tableau (ligne 7) ----
        headers3 = ["#", "Code Art.", "Article", "Rayon", "Famille", "Magasin",
                    "Format", "CA (FCFA)", "Marge (FCFA)", "Tx Marge", "Qté"]
        for j, lbl in enumerate(headers3, start=1):
            c = ws3.cell(row=7, column=j, value=lbl)
            c.font = Font(name=CAL, bold=True, size=10, color=WHITE_H)
            c.fill = PatternFill("solid", fgColor=NAVY2)
            c.alignment = Alignment(horizontal="center", vertical="center")

        # ---- Données (exhaustif, trié par article) ----
        r3 = 8
        for i, (_, row_) in enumerate(mns_detail.iterrows(), start=1):
            tx = row_['Tx Marge %']
            severe = pd.notna(tx) and tx < -5.0
            fillc = REDT_F if severe else AMBER_F
            code = row_['Code Art.']
            vals = [i,
                    int(code) if str(code).isdigit() else (code or None),
                    row_['Article_aff'], row_['Rayon_aff'], row_['Famille_aff'],
                    row_['Site'], row_['Format'],
                    row_['CA'], row_['Marge'],
                    (tx/100 if pd.notna(tx) else None), row_['Qté']]
            for j, v in enumerate(vals, start=1):
                cell = ws3.cell(row=r3, column=j, value=v)
                cell.fill = PatternFill("solid", fgColor=fillc)
                cell.border = box
                if j == 10:  # Tx Marge en rouge gras
                    cell.font = Font(name=CAL, size=10, bold=True, color=RED_TX)
                    cell.number_format = "0.0%"
                    cell.alignment = Alignment(horizontal="center")
                elif j in (8, 9, 11):
                    cell.font = Font(name=CAL, size=10, color=INK)
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right")
                elif j == 1:
                    cell.font = Font(name=CAL, size=10, color=INK)
                    cell.alignment = Alignment(horizontal="center")
                else:
                    cell.font = Font(name=CAL, size=10, color=INK)
                    cell.alignment = Alignment(horizontal="left")
            r3 += 1

        autosize(ws3, {'A':5,'B':12,'C':38,'D':18,'E':24,'F':28,'G':10,'H':14,'I':14,'J':10,'K':8})
        ws3.freeze_panes = "A8"
        ws3.auto_filter.ref = f"A7:K{r3-1}"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def top_familles_for_excel(fam, n, by):
    if by == 'casse':
        sub = fam[fam['Casse (Valeur)'].notna()]
        out = sub.nsmallest(n, 'Casse (Valeur)')[['Rayon_aff','Famille_aff','CA','Casse (Valeur)','%Casse (Valeur)']]
    elif by == 'promo':
        mat = fam[fam['CA'] > 1_000_000]
        out = mat.nlargest(n, '%CA Poids Promo')[['Rayon_aff','Famille_aff','CA','%CA Poids Promo','%Marge Promo','%Marge Hors Promo']]
    return out.reset_index(drop=True)

# ============================================================
# INTERFACE
# ============================================================
st.markdown("<div class='page-title'>COPIL Hebdo</div>"
            "<div class='page-caption'>📋 Vue réseau PGC · Marge négative par site · Destructeurs & Performeurs · "
            "objectifs marge alignés Méti · une seule extraction à charger</div>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 📥 Import")
    up = st.file_uploader("Export Article hebdo (.xlsx)", type=['xlsx'], key="up_export")
    st.caption("L'extraction Article contient déjà les sous-totaux Rayon et Famille — "
               "tout le module en est dérivé, rien d'autre à charger.")
    st.markdown("---")
    st.markdown("##### ⚙️ Paramètres")
    n_top = st.slider("Nombre de lignes par classement", 5, 30, 15)
    seuil_ca = st.number_input("Seuil CA mini — performeurs/croissance (FCFA)", 0, 5_000_000, 100_000, step=10_000)
    st.markdown("---")
    st.caption("SmartBuyer Hub · Module COPIL Hebdo · V2")

if up is None:
    st.markdown(
        f"<div class='info-box'>"
        f"<div class='it'>ℹ️ À quoi sert ce module ?</div>"
        f"<div class='ip'>Ce module prépare le point hebdo réseau pour le COPIL à partir de l'export PBI "
        f"<b>Rayon → Famille → Sous-Famille → Article</b> (avec ou sans colonne Site). "
        f"Un seul fichier à charger chaque semaine, dans la barre latérale.</div>"
        f"<div class='iq'>"
        f"<b>Dashboard COPIL</b> — CA, marge, quantités vs N-1 · performance par rayon vs objectifs Méti · "
        f"marge négative par site · top familles en baisse de CA / casse / poids promo<br>"
        f"<b>Destructeurs &amp; Performeurs</b> — articles agrégés réseau (sans doublon multi-site) : marge négative, "
        f"dégradation de taux, gain de marge, croissance/baisse de CA et de quantité</div>"
        f"</div>", unsafe_allow_html=True)
    st.stop()

df, perimetre, site_col = load_export(up.getvalue())
art = prep_articles(df)
mns = marge_negative_par_site(df, site_col)
mns_detail = detail_marge_neg_site(df, site_col)

if perimetre:
    with st.expander("🔎 Périmètre détecté dans le fichier"):
        st.code(perimetre, language=None)

tab1, tab2 = st.tabs(["📋 Dashboard COPIL", "💥 Destructeurs & Performeurs"])

# ---------------- TAB 1 : DASHBOARD COPIL (RAYON + SITE) ----------------
with tab1:
    k = kpis_globaux_rayon(df)
    if k is None:
        st.error("Ligne de total réseau ('Total') introuvable dans l'export — vérifiez le fichier.")
    else:
        perf = perf_par_rayon(df, CIBLES_DEFAUT)
        fam = family_metrics(df)
        line1, line2 = build_headline(k, perf, fam)
        st.markdown(
            f"<div class='recap-card'><div class='recap-line1'>{line1}</div>"
            + (f"<div class='recap-line2'>{line2}</div>" if line2 else "")
            + "</div>", unsafe_allow_html=True)

        st.markdown("<div class='section-label'>Vue d'ensemble réseau</div>", unsafe_allow_html=True)
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.markdown(kpi_card("CA", f"{fmt(k['ca'])} FCFA",
                    f"{k['evol_ca']*100:+.1f}% vs N-1", "pos" if k['evol_ca'] >= 0 else "neg"), unsafe_allow_html=True)
        c2.markdown(kpi_card("Marge", f"{fmt(k['marge'])} FCFA"), unsafe_allow_html=True)
        evo_tx = k['tx_marge'] - k['tx_marge_n1'] if pd.notna(k['tx_marge_n1']) else np.nan
        c3.markdown(kpi_card("Taux de marge", fmt_pct(k['tx_marge'], 2),
                    fmt_delta(evo_tx), "pos" if (pd.notna(evo_tx) and evo_tx >= 0) else "neg"), unsafe_allow_html=True)
        evo_qte = (k['qte']/k['qte_n1']-1)*100 if k['qte_n1'] else np.nan
        c4.markdown(kpi_card("Qté vendue", fmt(k['qte']),
                    f"{evo_qte:+.1f}% vs N-1" if pd.notna(evo_qte) else None,
                    "pos" if (pd.notna(evo_qte) and evo_qte >= 0) else "neg"), unsafe_allow_html=True)
        pct_casse = k['casse']/k['ca']*100 if k['ca'] else np.nan
        c5.markdown(kpi_card("Casse", f"{fmt(k['casse'])} FCFA", fmt_pct(pct_casse, 2) + " du CA", "neutral"), unsafe_allow_html=True)

        st.markdown("<div class='section-label'>Performance par rayon vs objectifs marge (Méti)</div>", unsafe_allow_html=True)
        disp = perf.copy()
        for c in ['Évol CA %', 'Évol Qté %', 'Taux Marge %', 'Objectif %']:
            disp[c] = disp[c].map(lambda v: fmt_pct(v))
        disp['Écart (pts)'] = perf['Écart (pts)'].map(fmt_delta)
        disp['CA'] = perf['CA'].map(lambda v: fmt(v))
        st.dataframe(disp, use_container_width=True, hide_index=True)

        # ---- 🆕 MARGE NÉGATIVE PAR SITE ----
        st.markdown("<div class='section-label'>Marge négative par site</div>", unsafe_allow_html=True)
        if mns is not None:
            total_neg = mns['Marge Négative (FCFA)'].sum()
            nb_art = int(mns['Nb Articles Nég.'].sum())
            pire = mns.iloc[0] if len(mns) else None
            m1, m2, m3 = st.columns(3)
            m1.markdown(kpi_card("Marge négative réseau", f"{fmt(total_neg)} FCFA",
                                 f"{nb_art} lignes article × site", "neg"), unsafe_allow_html=True)
            if pire is not None:
                m2.markdown(kpi_card("Site le plus touché", str(pire['Site']),
                                     f"{fmt(pire['Marge Négative (FCFA)'])} FCFA", "neg"), unsafe_allow_html=True)
                m3.markdown(kpi_card("Poids max vs CA site", fmt_pct(abs(pire['% Marge Nég / CA Site']), 2),
                                     str(pire['Site']), "neutral"), unsafe_allow_html=True)
            disp_mns = mns.copy()
            disp_mns['CA Site'] = disp_mns['CA Site'].map(fmt)
            disp_mns['Marge Négative (FCFA)'] = disp_mns['Marge Négative (FCFA)'].map(fmt)
            disp_mns['% Marge Nég / CA Site'] = disp_mns['% Marge Nég / CA Site'].map(lambda v: fmt_pct(v, 2))
            st.dataframe(disp_mns, use_container_width=True, hide_index=True)

            if mns_detail is not None and not mns_detail.empty:
                with st.expander(f"🔍 Détail exhaustif — {len(mns_detail)} lignes article × site en marge négative"):
                    site_opts = ["Tous les sites"] + sorted(mns_detail['Site'].unique())
                    site_sel = st.selectbox("Site", site_opts, key="site_neg_sel")
                    d = mns_detail if site_sel == "Tous les sites" else mns_detail[mns_detail['Site'] == site_sel]
                    d = d.rename(columns={'Rayon_aff': 'Rayon', 'Famille_aff': 'Famille', 'Article_aff': 'Article'})
                    d = d.copy()
                    d['CA'] = d['CA'].map(fmt); d['Marge'] = d['Marge'].map(fmt)
                    d['Tx Marge %'] = d['Tx Marge %'].map(lambda v: fmt_pct(v))
                    d['Qté'] = d['Qté'].map(lambda v: f"{v:,.0f}".replace(",", " "))
                    st.dataframe(d[['Code Art.', 'Article', 'Rayon', 'Famille', 'Site', 'Format',
                                    'CA', 'Marge', 'Tx Marge %', 'Qté']],
                                 use_container_width=True, hide_index=True, height=420)
        else:
            st.caption("⚠️ Aucune colonne Site/Magasin détectée dans l'export — vue par site indisponible. "
                       "Ajoutez la dimension Site à l'export PBI pour l'activer.")

        st.markdown("<div class='section-label'>Détail par famille — toutes familles (sans objectif)</div>", unsafe_allow_html=True)
        fam_disp = fam[['Rayon_aff','Famille_aff','CA','Évol CA %','Marge','Tx Marge %','Qté Vente','Évol Qté %']].copy()
        fam_disp = fam_disp.rename(columns={'Rayon_aff':'Rayon','Famille_aff':'Famille','Tx Marge %':'Taux Marge %'})
        fam_disp = fam_disp.sort_values(['Rayon','CA'], ascending=[True, False])
        for c in ['CA','Marge','Qté Vente']:
            fam_disp[c] = fam_disp[c].map(fmt)
        for c in ['Évol CA %','Taux Marge %','Évol Qté %']:
            fam_disp[c] = fam_disp[c].map(lambda v: fmt_pct(v))
        st.dataframe(fam_disp, use_container_width=True, hide_index=True, height=420)
        st.caption("Pas d'objectif au niveau Famille (le cadrage marge est piloté au niveau Rayon) — vue CA / marge / quantité uniquement.")

        st.markdown("<div class='section-label'>Top & Flop par famille</div>", unsafe_allow_html=True)

        def pair(title_top, title_flop, metric, cols, fmt_map, ca_floor=0, directional=True):
            cA, cB = st.columns(2)
            for cc, mode, title, emoji in [(cA, 'top', title_top, '🟢'), (cB, 'flop', title_flop, '🔴')]:
                with cc:
                    st.markdown(f"**{emoji} {title}**")
                    t = top_flop_table(fam, metric, n_top, mode, cols, ca_floor, directional)
                    t = t.rename(columns={'Rayon_aff':'Rayon','Famille_aff':'Famille'})
                    if t.empty:
                        st.caption("Aucune famille ne respecte ce sens sur la période.")
                    else:
                        for c, f in fmt_map.items():
                            if c in t.columns: t[c] = t[c].map(f)
                        st.dataframe(t, hide_index=True, use_container_width=True)

        st.markdown("##### 📈 Évolution du CA (classement en valeur FCFA)")
        pair(f"Top {n_top} — Meilleur gain de CA", f"Flop {n_top} — Plus forte perte de CA",
             'Perte CA', ['CA','CA N-1','Évol CA %','Tx Marge %'],
             {'CA': fmt, 'CA N-1': fmt, 'Évol CA %': lambda v: fmt_pct(v), 'Tx Marge %': lambda v: fmt_pct(v)})

        st.markdown("##### 💰 Taux de marge")
        pair(f"Top {n_top} — Meilleur taux de marge", f"Flop {n_top} — Taux de marge le plus faible",
             'Tx Marge %', ['CA','Marge','Tx Marge %'],
             {'CA': fmt, 'Marge': fmt, 'Tx Marge %': lambda v: fmt_pct(v)}, ca_floor=1_000_000, directional=False)

        st.markdown("##### 📊 Évolution du taux de marge (pts vs N-1)")
        pair(f"Top {n_top} — Meilleure progression", f"Flop {n_top} — Plus forte dégradation",
             'Écart Tx Marge (pts)', ['CA','Tx Marge %','Écart Tx Marge (pts)'],
             {'CA': fmt, 'Tx Marge %': lambda v: fmt_pct(v), 'Écart Tx Marge (pts)': fmt_delta}, ca_floor=1_000_000)

        st.markdown("##### 💔 Casse & 🎯 Poids promo")
        cC, cD = st.columns(2)
        with cC:
            st.markdown(f"**🔴 Flop {n_top} — Casse en valeur**")
            t = top_familles(df, n_top, 'casse').rename(
                columns={'Rayon_aff':'Rayon','Famille_aff':'Famille','Casse (Valeur)':'Casse','%Casse (Valeur)':'% Casse'})
            t['CA'] = t['CA'].map(fmt); t['Casse'] = t['Casse'].map(fmt)
            t['% Casse'] = t['% Casse'].map(lambda v: fmt_pct(v*100, 2) if pd.notna(v) else "—")
            st.dataframe(t, hide_index=True, use_container_width=True)
        with cD:
            st.markdown(f"**🟠 Top {n_top} — Poids promo (CA&gt;1M)**")
            t = top_familles(df, n_top, 'promo').rename(
                columns={'Rayon_aff':'Rayon','Famille_aff':'Famille','%CA Poids Promo':'Poids Promo',
                         '%Marge Promo':'Tx M. Promo','%Marge Hors Promo':'Tx M. HP'})
            t['CA'] = t['CA'].map(fmt)
            for c in ['Poids Promo','Tx M. Promo','Tx M. HP']:
                t[c] = t[c].map(lambda v: fmt_pct(v*100) if pd.notna(v) else "—")
            st.dataframe(t, hide_index=True, use_container_width=True)

        st.markdown("<div class='section-label'>Export</div>", unsafe_allow_html=True)
        art_res_export = destructeurs_performeurs(art, n=n_top, seuil_ca=seuil_ca)
        xls = build_excel_full(k, perf, fam, n_top, art_res_export, perimetre, mns, mns_detail)
        st.download_button("📥 Télécharger le récap complet COPIL (.xlsx)", xls,
                            file_name="COPIL_Hebdo.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        nb_feuilles = "3 feuilles : Dashboard COPIL, Destructeurs & Performeurs, Marges Négatives par Site" \
            if (mns_detail is not None and not mns_detail.empty) \
            else "2 feuilles : Dashboard COPIL et Destructeurs & Performeurs"
        st.caption(f"Le fichier contient {nb_feuilles}.")

# ---------------- TAB 2 : DESTRUCTEURS & PERFORMEURS (ARTICLE, AGRÉGÉ RÉSEAU) ----------------
with tab2:
    st.caption(f"{len(art):,} articles (agrégés réseau, sans doublon multi-site) · "
               f"seuil de matérialité : {fmt(seuil_ca)} FCFA (modifiable dans la barre latérale)".replace(",", " "))
    res = destructeurs_performeurs(art, n=n_top, seuil_ca=seuil_ca)

    st.markdown(f"<span class='badge' style='background:rgba(255,59,48,0.12);color:#C62A22'>A · Marge négative</span>", unsafe_allow_html=True)
    show_table(res['A_marge_neg'], ['CA','Marge','Tx Marge %','Qté Vente'],
               {'CA': fmt, 'Marge': fmt, 'Tx Marge %': lambda v: fmt_pct(v), 'Qté Vente': fmt})

    st.markdown(f"<span class='badge' style='background:rgba(255,59,48,0.12);color:#C62A22'>B · Dégradation du taux de marge (marge encore positive)</span>", unsafe_allow_html=True)
    show_table(res['B_degrad_marge'], ['CA','Tx Marge %','Écart Tx Marge (pts)'],
               {'CA': fmt, 'Tx Marge %': lambda v: fmt_pct(v), 'Écart Tx Marge (pts)': fmt_delta})

    st.markdown(f"<span class='badge' style='background:rgba(52,199,89,0.13);color:#1A7A3A'>C · Performeurs — gain de marge en valeur</span>", unsafe_allow_html=True)
    show_table(res['C_perf_gain_marge'], ['CA','Gain Marge (FCFA)','Tx Marge %'],
               {'CA': fmt, 'Gain Marge (FCFA)': fmt, 'Tx Marge %': lambda v: fmt_pct(v)})

    st.markdown(f"<span class='badge' style='background:rgba(52,199,89,0.13);color:#1A7A3A'>D · Plus forte croissance de CA</span>", unsafe_allow_html=True)
    d4 = res['D_croissance_ca'].copy()
    if not d4.empty:
        d4['Évol CA %'] = (d4['CA']/d4['CA N-1']-1)*100
    show_table(d4, ['CA','CA N-1','Évol CA %','Tx Marge %'],
               {'CA': fmt, 'CA N-1': fmt, 'Évol CA %': lambda v: fmt_pct(v), 'Tx Marge %': lambda v: fmt_pct(v)})
    st.caption("⚠️ Une forte évolution % peut refléter un effet de base (article quasi absent en N-1) plutôt qu'une vraie dynamique.")

    st.markdown(f"<span class='badge' style='background:rgba(255,59,48,0.12);color:#C62A22'>E · Plus forte baisse de CA</span>", unsafe_allow_html=True)
    show_table(res['E_baisse_ca'], ['CA','CA N-1','Perte CA (FCFA)'],
               {'CA': fmt, 'CA N-1': fmt, 'Perte CA (FCFA)': fmt})

    st.markdown(f"<span class='badge' style='background:rgba(52,199,89,0.13);color:#1A7A3A'>F · Plus forte hausse de quantité vendue</span>", unsafe_allow_html=True)
    show_table(res['F_hausse_qte'], ['Qté Vente','Qté Vente N-1','Variation Qté','CA'],
               {'Qté Vente': fmt, 'Qté Vente N-1': fmt, 'Variation Qté': lambda v: f"{v:+,.0f}".replace(",", " "), 'CA': fmt})

    st.markdown(f"<span class='badge' style='background:rgba(255,59,48,0.12);color:#C62A22'>G · Plus forte baisse de quantité vendue</span>", unsafe_allow_html=True)
    show_table(res['G_baisse_qte'], ['Qté Vente','Qté Vente N-1','Variation Qté','CA'],
               {'Qté Vente': fmt, 'Qté Vente N-1': fmt, 'Variation Qté': lambda v: f"{v:+,.0f}".replace(",", " "), 'CA': fmt})
