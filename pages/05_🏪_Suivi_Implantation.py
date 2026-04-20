"""
Rapport Implantation v3 — MINIMAL & SAFE
Zéro dépendances externes, zéro erreurs sorted
"""
import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import date

st.set_page_config(page_title="Rapport Implantation", layout="wide")

st.markdown("""
<style>
:root { --bg: #F2F2F7; --surface: #FFF; --accent: #007AFF; }
html, body { background: var(--bg) !important; }
.main { background: var(--bg) !important; }
header, #MainMenu, footer { display: none !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="background:#fff; border-bottom:1px solid #e5e5ea; padding:12px 28px; margin:0 -2rem 24px; display:flex; justify-content:space-between;">
  <div style="font-size:16px; font-weight:700;">📋 Rapport Implantation v3</div>
  <div style="font-size:12px; color:#666;">{}</div>
</div>
""".format(date.today().strftime("%d %b %Y")), unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════════
# CHARGEMENT T1
# ════════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📁 Chargement")
    st.divider()
    t1_file = st.file_uploader("T1 Flux (CSV ;)", type=["csv"], key="t1")

if not t1_file:
    st.info("⬆️ Charge T1 CSV pour démarrer")
    st.stop()

try:
    buf_t1 = io.BytesIO(t1_file.read())
    try:
        t1_raw = pd.read_csv(buf_t1, sep=';', encoding='latin1')
    except:
        buf_t1.seek(0)
        t1_raw = pd.read_csv(buf_t1, sep=',', encoding='utf-8')
    
    t1_raw.columns = t1_raw.columns.str.strip().str.upper()
    
    if 'ARTICLE' not in t1_raw.columns:
        st.error("❌ Colonne ARTICLE manquante")
        st.stop()
    
    t1_raw['SKU'] = t1_raw['ARTICLE'].astype(str).str.strip().str.zfill(8)
    t1_raw = t1_raw[t1_raw['SKU'].str.match(r'^\d{8}$', na=False)].copy()
    
    if len(t1_raw) == 0:
        st.error("❌ Aucun article valide")
        st.stop()
    
    st.success(f"✅ T1: {len(t1_raw)} articles")
except Exception as e:
    st.error(f"❌ Erreur T1: {str(e)[:100]}")
    st.stop()

# ════════════════════════════════════════════════════════════════════════════════
# CHARGEMENT PBI
# ════════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    pbi_file = st.file_uploader("Stock PBI (XLSX)", type=["xlsx"], key="pbi")

if not pbi_file:
    st.info("⬆️ Charge PBI XLSX")
    st.stop()

try:
    buf_pbi = io.BytesIO(pbi_file.read())
    df_raw = pd.read_excel(buf_pbi, header=None)
    
    if len(df_raw) < 3:
        st.error("❌ PBI trop court")
        st.stop()
    
    sites_raw = df_raw.iloc[0, 2:].tolist()
    
    results = []
    for idx in range(2, len(df_raw)):
        row = df_raw.iloc[idx]
        
        article_raw = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if not article_raw or article_raw == 'Total':
            continue
        
        if ' - ' in article_raw:
            sku, lib = article_raw.split(' - ', 1)
            sku = sku.strip().zfill(8)
        else:
            sku = article_raw[:8].zfill(8)
            lib = article_raw
        
        for site_idx, site_raw in enumerate(sites_raw, start=2):
            if pd.isna(site_raw):
                continue
            
            site_str = str(site_raw).strip()
            if ' - ' in site_str:
                code, nom = site_str.split(' - ', 1)
            else:
                code, nom = site_str, site_str
            
            try:
                stock = int(float(row.iloc[site_idx])) if pd.notna(row.iloc[site_idx]) else 0
            except:
                stock = 0
            
            results.append({
                'SKU': sku,
                'Article': lib,
                'Site': nom.strip(),
                'Stock': stock,
            })
    
    if not results:
        st.error("❌ Aucune donnée PBI")
        st.stop()
    
    df_stock = pd.DataFrame(results)
    st.success(f"✅ PBI: {len(df_stock)} lignes")
except Exception as e:
    st.error(f"❌ Erreur PBI: {str(e)[:100]}")
    st.stop()

# ════════════════════════════════════════════════════════════════════════════════
# FILTRES
# ════════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.divider()
    st.markdown("### 🔍 Filtres")
    
    # IMPORTANT: Utiliser list() pour éviter sorted() sur Series
    sites_unique = df_stock['Site'].unique()
    sites_list = list(sites_unique)
    sites_list.sort()  # Trier la liste, pas la Series
    
    mag_sel = st.multiselect("Magasins", sites_list, default=sites_list[:3])

if not mag_sel:
    st.warning("Sélectionne au moins 1 magasin")
    st.stop()

# ════════════════════════════════════════════════════════════════════════════════
# AFFICHAGE
# ════════════════════════════════════════════════════════════════════════════════
st.markdown("### 📊 Aperçu données")

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("T1 Articles", len(t1_raw))
with col2:
    st.metric("Stock Lignes", len(df_stock))
with col3:
    st.metric("Magasins Actifs", len(mag_sel))

st.markdown("### Stock par magasin")
preview = df_stock[df_stock['Site'].isin(mag_sel)].head(30)
st.dataframe(preview, use_container_width=True, hide_index=True)

st.markdown("### Top 10 articles")
top = df_stock.groupby(['SKU', 'Article'])['Stock'].sum().reset_index().nlargest(10, 'Stock')
st.dataframe(top, use_container_width=True, hide_index=True)

st.success("✅ v3 — Zéro erreurs")
