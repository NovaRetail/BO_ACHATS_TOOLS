"""
Rapport Implantation v2.1 PRODUCTION
────────────────────────────────────
✅ Adapté aux fichiers réels Carrefour CI
✅ T1: CSV semicolon separator
✅ PBI: Rayon | Article | Stock×N sites
✅ Ultra-robuste sorted()
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import date
import re

TODAY = date.today()
TODAY_STR = TODAY.strftime("%d %b %Y")
TODAY_FILE = TODAY.strftime("%Y%m%d")

st.set_page_config(page_title="Rapport Implantation · Carrefour", layout="wide", initial_sidebar_state="expanded")

# DESIGN SYSTEM (réduit pour clarté)
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
:root {
  --bg:#f0f2f8; --surface:#fff; --border:#e2e8f4; --text:#0f1729; --muted:#64748b;
  --accent:#2563eb; --green:#059669; --blue:#0284c7; --red:#dc2626; --gold:#b45309; --radius:10px;
}
html,body,[class*="css"]{font-family:'Inter',sans-serif!important;background:var(--bg)!important;color:var(--text)!important;}
.main,section[data-testid="stMain"]{background:var(--bg)!important;}
.block-container{padding:0 2rem 4rem!important;max-width:1520px;}
header[data-testid="stHeader"],#MainMenu,footer{display:none!important;}
.topbar{background:var(--text);margin:0 -2rem 24px;padding:14px 28px;display:flex;align-items:center;justify-content:space-between;}
.topbar-left{display:flex;align-items:center;gap:14px;}
.topbar-icon{width:38px;height:38px;border-radius:9px;background:linear-gradient(135deg,#3b82f6,#60a5fa);display:flex;align-items:center;justify-content:center;font-size:20px;}
.topbar-title{font-size:17px;font-weight:700;color:#fff;}
.topbar-sub{font-size:11px;color:#94a3b8;margin-top:1px;}
.topbar-pill{background:rgba(255,255,255,.08);color:#94a3b8;border:1px solid rgba(255,255,255,.12);border-radius:6px;padding:4px 12px;font-size:11px;font-weight:500;}
.info-banner{background:#f0f9ff;border:1px solid #bae6fd;border-radius:10px;padding:12px 16px;font-size:13px;color:#0284c7;margin-bottom:14px;}
.sh{font-size:10px;font-weight:700;text-transform:uppercase;color:var(--muted);margin:22px 0 12px;padding-bottom:8px;border-bottom:1px solid var(--border);}
section[data-testid="stSidebar"]{background:#fff!important;border-right:1px solid var(--border)!important;}
.stDownloadButton>button{background:linear-gradient(135deg,#0f1729,#1e293b)!important;color:#fff!important;border:none!important;border-radius:10px!important;font-weight:700!important;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS ULTRA-SAFE
# ══════════════════════════════════════════════════════════════════════════════

def safe_sorted_list(items):
    """Trier une liste quelconque sans risque"""
    try:
        if not items:
            return []
        clean = []
        for x in items:
            if pd.notna(x):
                s = str(x).strip()
                if s and s not in ('nan', '', 'NaN'):
                    clean.append(s)
        return sorted(set(clean))
    except:
        return list(set([str(x) for x in items if pd.notna(x)]))

def load_t1_csv(file_bytes, filename):
    """Charger T1 CSV avec bons paramètres"""
    buf = io.BytesIO(file_bytes)
    try:
        # Essayer semicolon d'abord (format Carrefour)
        df = pd.read_csv(buf, sep=';', encoding='latin1')
    except:
        try:
            buf.seek(0)
            df = pd.read_csv(buf, sep=',', encoding='utf-8')
        except:
            buf.seek(0)
            df = pd.read_csv(buf, sep=None, engine='python', encoding='latin1')
    
    # Normaliser colonnes
    df.columns = df.columns.str.strip().str.upper()
    
    # Parser SKU
    if 'ARTICLE' in df.columns:
        df['SKU'] = df['ARTICLE'].astype(str).str.strip().str.zfill(8).str.slice(0, 8)
        df = df[df['SKU'].str.match(r'^\d{8}$', na=False)].copy()
    
    # Ajouter colonnes manquantes
    for col in ['LIBELLÉ ARTICLE', 'MODE APPRO', 'SEMAINE RECEPTION']:
        if col not in df.columns:
            df[col] = ''
    
    df['ORIGINE'] = df['MODE APPRO'].apply(lambda m: 'IM' if 'IMPORT' in str(m).upper() else 'LO')
    return df, None

def load_pbi_excel(file_bytes, filename):
    """Charger PBI Excel — Structure réelle: Rayon | Article | Stock×N"""
    buf = io.BytesIO(file_bytes)
    df_raw = pd.read_excel(buf, header=None)
    
    if len(df_raw) < 3:
        return None, "Fichier PBI trop court"
    
    # Row 0: Sites
    sites_raw = df_raw.iloc[0, 2:].tolist()  # Skip Col 0 (Rayon) et Col 1 (Article)
    
    results = []
    for idx in range(2, len(df_raw)):  # Skip headers
        row = df_raw.iloc[idx]
        
        # Col 0: Rayon, Col 1: Article
        rayon = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        article_raw = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        
        if not article_raw or article_raw == 'Total':
            continue
        
        # Parse article: "10000119 - 4X25CL BOIS,EN,RED BULL MM"
        if ' - ' in article_raw:
            sku, lib = article_raw.split(' - ', 1)
            sku = sku.strip().zfill(8)
            lib = lib.strip()
        else:
            sku = article_raw[:8].zfill(8)
            lib = article_raw
        
        # Stock par site (Col 2+)
        for site_idx, site_raw in enumerate(sites_raw, start=2):
            if pd.isna(site_raw):
                continue
            
            site_str = str(site_raw).strip()
            if ' - ' in site_str:
                code, nom = site_str.split(' - ', 1)
                code = code.strip()
                nom = nom.strip()
            else:
                code = site_str
                nom = site_str
            
            try:
                stock = int(float(row.iloc[site_idx])) if pd.notna(row.iloc[site_idx]) else 0
            except:
                stock = 0
            
            results.append({
                'SKU': sku,
                'Libellé article': lib,
                'Code site': code,
                'Libellé site': nom,
                'Stock': stock,
            })
    
    if not results:
        return None, "Aucun article parsé"
    
    return pd.DataFrame(results), None

# ══════════════════════════════════════════════════════════════════════════════
# TOPBAR
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div style="background:#0f1729;margin:0 -2rem 24px;padding:14px 28px;display:flex;align-items:center;justify-content:space-between;">
  <div style="display:flex;align-items:center;gap:14px;">
    <div style="width:38px;height:38px;border-radius:9px;background:linear-gradient(135deg,#3b82f6,#60a5fa);display:flex;align-items:center;justify-content:center;font-size:20px;">📋</div>
    <div>
      <div style="font-size:17px;font-weight:700;color:#fff;">Rapport Implantation</div>
      <div style="font-size:11px;color:#94a3b8;margin-top:1px;">Suivi Stock PBI · Rupture Commune</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:12px;color:#60a5fa;font-size:12px;">
    {TODAY_STR} · v2.1 PROD
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📁 Chargement")
    st.divider()
    st.markdown("**T1 Flux (CSV)**")
    t1_file = st.file_uploader("T1", type=["csv"], key="t1", label_visibility="collapsed")
    st.markdown("**Stock PBI (XLSX)**")
    pbi_file = st.file_uploader("PBI", type=["xlsx"], key="pbi", label_visibility="collapsed")

# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT
# ══════════════════════════════════════════════════════════════════════════════
if not t1_file:
    st.markdown('<div class="info-banner">⬆️ Charge T1 CSV pour démarrer</div>', unsafe_allow_html=True)
    st.stop()

with st.spinner("Lecture T1…"):
    t1_raw, err = load_t1_csv(t1_file.read(), t1_file.name)
    if err:
        st.error(f"❌ T1: {err}")
        st.stop()

st.success(f"✅ T1 chargé: {len(t1_raw)} articles")

if not pbi_file:
    st.markdown('<div class="info-banner">⬆️ Charge PBI XLSX</div>', unsafe_allow_html=True)
    st.stop()

with st.spinner("Parsing PBI…"):
    df_stock, err = load_pbi_excel(pbi_file.read(), pbi_file.name)
    if err:
        st.error(f"❌ PBI: {err}")
        st.stop()

st.success(f"✅ PBI chargé: {len(df_stock)} lignes stock")

# ══════════════════════════════════════════════════════════════════════════════
# FILTRES
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.divider()
    st.markdown("### 🔍 Filtres")
    
    try:
        magasins_list = safe_sorted_list(df_stock['Libellé site'].unique())
        st.write(f"✅ {len(magasins_list)} magasins trouvés")
    except Exception as e:
        st.error(f"Erreur magasins: {e}")
        magasins_list = []
    
    mag_sel = st.multiselect("Magasins", magasins_list, default=magasins_list[:3] if magasins_list else [])

if not mag_sel:
    st.warning("Sélectionne au moins un magasin")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# AFFICHAGE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sh">DONNÉES CHARGÉES</div>', unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("T1 Articles", len(t1_raw))
with col2:
    st.metric("Stock Lignes", len(df_stock))
with col3:
    st.metric("Magasins Actifs", len(mag_sel))

st.markdown('<div class="sh">APERÇU STOCK</div>', unsafe_allow_html=True)
preview = df_stock[df_stock['Libellé site'].isin(mag_sel)].head(20)
st.dataframe(preview[['SKU', 'Libellé article', 'Libellé site', 'Stock']], use_container_width=True, hide_index=True)

st.markdown('<div class="sh">TOP ARTICLES PAR STOCK</div>', unsafe_allow_html=True)
top_stock = df_stock.groupby(['SKU', 'Libellé article'])['Stock'].sum().reset_index().nlargest(10, 'Stock')
st.dataframe(top_stock, use_container_width=True, hide_index=True)

st.success("✅ v2.1 PRODUCTION — Prêt pour développement")
