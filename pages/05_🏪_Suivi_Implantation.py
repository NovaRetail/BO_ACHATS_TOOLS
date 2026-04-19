"""
Rapport Implantation · Carrefour CI — v2.1 [PATCH PARSER PBI]
─────────────────────────────────────────────────────────────────
✅ Patch : load_pbi_stock() supporte MAINTENANT format multi-niveaux Rayon/Famille/Article
Rest of app = v2.0 identique
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import date
import re

TODAY      = date.today()
TODAY_STR  = TODAY.strftime("%d %b %Y")
TODAY_FILE = TODAY.strftime("%Y%m%d")

st.set_page_config(
    page_title="Rapport Implantation · Carrefour",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ══════════════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM (v2.0 identique)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
:root {
  --bg:#f0f2f8; --surface:#fff; --border:#e2e8f4; --text:#0f1729; --muted:#64748b;
  --accent:#2563eb; --accent-l:#eff4ff; --accent-bd:#bfdbfe;
  --green:#059669;  --green-l:#ecfdf5;  --green-bd:#6ee7b7;
  --blue:#0284c7;   --blue-l:#f0f9ff;   --blue-bd:#bae6fd;
  --red:#dc2626;    --red-l:#fef2f2;    --red-bd:#fecaca;
  --gold:#b45309;   --gold-l:#fffbeb;   --gold-bd:#fcd34d;
  --radius:10px; --shadow:0 1px 3px rgba(15,23,42,.06),0 4px 16px rgba(15,23,42,.04);
}
html,body,[class*="css"]{font-family:'Inter',sans-serif!important;background:var(--bg)!important;color:var(--text)!important;}
.main,section[data-testid="stMain"]{background:var(--bg)!important;}
.block-container{padding:0 2rem 4rem!important;max-width:1520px;}
header[data-testid="stHeader"],#MainMenu,footer{display:none!important;}
.topbar{background:var(--text);margin:0 -2rem 24px;padding:14px 28px;display:flex;align-items:center;justify-content:space-between;}
.topbar-left{display:flex;align-items:center;gap:14px;}
.topbar-icon{width:38px;height:38px;border-radius:9px;background:linear-gradient(135deg,#3b82f6,#60a5fa);display:flex;align-items:center;justify-content:center;font-size:20px;}
.topbar-title{font-size:17px;font-weight:700;color:#fff;letter-spacing:-.01em;}
.topbar-sub{font-size:11px;color:#94a3b8;font-family:'JetBrains Mono';margin-top:1px;}
.topbar-pill{background:rgba(255,255,255,.08);color:#94a3b8;border:1px solid rgba(255,255,255,.12);border-radius:6px;padding:4px 12px;font-size:11px;font-weight:500;}
.topbar-date{color:#60a5fa;font-size:12px;font-family:'JetBrains Mono';}
.alert-banner{background:#fff;border:1px solid var(--red-bd);border-left:4px solid var(--red);border-radius:var(--radius);padding:14px 20px;margin-bottom:20px;display:flex;align-items:center;gap:0;flex-wrap:wrap;}
.ab-badge{background:var(--red);color:#fff;border-radius:6px;padding:4px 10px;font-size:11px;font-weight:700;letter-spacing:.04em;margin-right:16px;white-space:nowrap;}
.ab-item{display:flex;flex-direction:column;align-items:center;padding:0 20px;border-right:1px solid var(--border);}
.ab-item:last-child{border-right:none;padding-right:0;}
.ab-num{font-size:26px;font-weight:800;line-height:1;}
.ab-lbl{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-top:1px;}
.rag-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px;margin-bottom:22px;}
.rag-card{border-radius:var(--radius);padding:14px 16px;border:1px solid transparent;box-shadow:var(--shadow);position:relative;}
.rag-card.g{background:var(--green-l);border-color:var(--green-bd);}
.rag-card.r{background:var(--red-l);border-color:var(--red-bd);}
.rag-name{font-size:11px;font-weight:600;color:var(--text);margin-bottom:5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:90%;}
.rag-pct{font-size:30px;font-weight:800;line-height:1;letter-spacing:-.02em;}
.rag-detail{font-size:10px;color:var(--muted);font-family:'JetBrains Mono';margin-top:3px;}
.rag-dot{width:8px;height:8px;border-radius:50%;position:absolute;top:14px;right:14px;}
.strip{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:16px;}
.strip-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px;box-shadow:var(--shadow);}
.strip-tag{display:inline-block;border-radius:4px;padding:2px 8px;font-size:10px;font-weight:700;margin-bottom:6px;}
.tag-im{background:#eff4ff;color:#2563eb;border:1px solid #bfdbfe;}
.tag-lo{background:#ecfdf5;color:#059669;border:1px solid #6ee7b7;}
.strip-label{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin-bottom:4px;}
.strip-val{font-size:26px;font-weight:800;line-height:1;letter-spacing:-.01em;}
.strip-sub{font-size:10px;color:var(--muted);font-family:'JetBrains Mono';margin-top:2px;}
.kpi-row{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:22px;}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px 20px 16px;box-shadow:var(--shadow);position:relative;overflow:hidden;}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;}
.kpi.g::before{background:var(--green);}
.kpi.b::before{background:var(--blue);}
.kpi.r::before{background:var(--red);}
.kpi.o::before{background:var(--gold);}
.kpi-lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:var(--muted);margin-bottom:10px;}
.kpi-val{font-size:44px;font-weight:800;line-height:1;letter-spacing:-.02em;}
.kpi.g .kpi-val{color:var(--green);}
.kpi.b .kpi-val{color:var(--blue);}
.kpi.r .kpi-val{color:var(--red);}
.kpi.o .kpi-val{color:var(--gold);}
.kpi-pct{font-size:12px;font-weight:600;color:var(--muted);margin-top:4px;font-family:'JetBrains Mono';}
.kpi-bar{margin-top:12px;height:3px;border-radius:3px;background:var(--border);}
.kpi-bar-fill{height:100%;border-radius:3px;}
.sh{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:var(--muted);margin:22px 0 12px;padding-bottom:8px;border-bottom:1px solid var(--border);}
.nav-tab-active{background:var(--text)!important;color:#fff!important;border-radius:8px;padding:9px 0;text-align:center;font-size:13px;font-weight:700;box-shadow:0 4px 14px rgba(15,23,42,.25);margin-bottom:10px;cursor:default;}
.ac{border-radius:var(--radius);padding:16px 18px;margin-bottom:10px;border:1px solid;display:flex;align-items:center;justify-content:space-between;}
.ac.red{background:var(--red-l);border-color:var(--red-bd);}
.ac.blue{background:var(--blue-l);border-color:var(--blue-bd);}
.ac.gold{background:var(--gold-l);border-color:var(--gold-bd);}
.ac-title{font-size:14px;font-weight:700;}
.ac-sub{font-size:11px;color:var(--muted);margin-top:2px;}
.ac-count{font-size:34px;font-weight:900;letter-spacing:-.02em;}
.ok-banner{background:var(--green-l);border:1px solid var(--green-bd);border-radius:var(--radius);padding:10px 16px;font-size:13px;color:var(--green);margin-bottom:14px;}
.info-banner{background:var(--blue-l);border:1px solid var(--blue-bd);border-radius:var(--radius);padding:12px 16px;font-size:13px;color:var(--blue);margin-bottom:14px;}
.gold-banner{background:var(--gold-l);border:1px solid var(--gold-bd);border-radius:var(--radius);padding:12px 16px;font-size:13px;color:var(--gold);margin-bottom:14px;}
.cession-header{background:var(--gold-l);border:1.5px solid var(--gold-bd);border-radius:var(--radius);padding:14px 18px;margin-bottom:12px;}
.cession-article{font-size:13px;font-weight:700;color:var(--text);margin-bottom:6px;}
.cession-row{display:flex;align-items:center;gap:10px;flex-wrap:wrap;}
.cession-badge{border-radius:5px;padding:3px 10px;font-size:11px;font-weight:700;}
.badge-detresse{background:var(--red-l);color:var(--red);border:1px solid var(--red-bd);}
.badge-cedant{background:var(--green-l);color:var(--green);border:1px solid var(--green-bd);}
.badge-stock{background:var(--accent-l);color:var(--accent);border:1px solid var(--accent-bd);}
.badge-qty{background:#f0f2f8;color:var(--text);border:1px solid var(--border);}
section[data-testid="stSidebar"]{background:#fff!important;border-right:1px solid var(--border)!important;min-width:280px!important;max-width:280px!important;}
section[data-testid="stSidebar"] .block-container{padding:.8rem .8rem 2rem!important;}
.stDownloadButton>button{background:linear-gradient(135deg,#0f1729,#1e293b)!important;color:#fff!important;border:none!important;border-radius:var(--radius)!important;font-weight:700!important;font-size:14px!important;padding:12px!important;width:100%!important;letter-spacing:.01em!important;box-shadow:0 4px 14px rgba(15,23,42,.3)!important;}
div[data-baseweb="tag"]{background:var(--accent-l)!important;color:var(--accent)!important;}
</style>
""", unsafe_allow_html=True)

def fix_encoding(df: pd.DataFrame) -> pd.DataFrame:
    try:
        if any('Ã' in str(c) for c in df.columns):
            df.columns = [c.encode('latin1').decode('utf-8', errors='replace') for c in df.columns]
    except Exception:
        pass
    return df

@st.cache_data(show_spinner=False)
def read_csv_smart(file_bytes: bytes, filename: str) -> pd.DataFrame:
    buf = io.BytesIO(file_bytes)
    for enc in ('latin1', 'utf-8-sig', 'cp1252'):
        for sep in (';', ',', '\t'):
            try:
                buf.seek(0)
                df = pd.read_csv(buf, sep=sep, encoding=enc, low_memory=False, on_bad_lines='skip')
                if df.shape[1] >= 3:
                    return fix_encoding(df)
            except Exception:
                continue
    buf.seek(0)
    return fix_encoding(pd.read_csv(buf, sep=None, engine='python', encoding='latin1', on_bad_lines='skip'))

@st.cache_data(show_spinner=False)
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = (
        df.columns.astype(str).str.strip()
        .str.replace('\ufeff', '', regex=False)
        .str.replace('\xa0', ' ', regex=False)
        .str.upper()
    )
    return df

def sem_sort(s) -> int:
    try:
        return int(str(s).strip('Ss'))
    except Exception:
        return 99

def taux_implantation(df: pd.DataFrame) -> int:
    if len(df) == 0:
        return 0
    return int(df['Statut'].isin(['Implantation Terminée']).sum() / len(df) * 100)

def color_taux(t: int) -> str:
    return '059669' if t >= 80 else ('0284c7' if t >= 50 else 'dc2626')

def load_t1(file_bytes: bytes, filename: str):
    buf = io.BytesIO(file_bytes)
    if filename.lower().endswith(('.xlsx', '.xls')):
        df_peek = pd.read_excel(buf, header=None, nrows=1)
    else:
        buf.seek(0)
        try:
            df_peek = pd.read_csv(buf, header=None, nrows=1, sep=None, engine='python', encoding='latin1')
        except Exception:
            df_peek = None

    no_header = False
    if df_peek is not None:
        first_val = str(df_peek.iloc[0, 0]).strip().replace('.0', '')
        no_header = first_val.isdigit()

    buf.seek(0)
    if filename.lower().endswith(('.xlsx', '.xls')):
        df = pd.read_excel(buf, header=None) if no_header else normalize_columns(pd.read_excel(buf))
    else:
        if no_header:
            buf.seek(0)
            try:
                df = pd.read_csv(buf, header=None, sep=None, engine='python', encoding='latin1', on_bad_lines='skip')
            except Exception:
                df = normalize_columns(read_csv_smart(file_bytes, filename))
        else:
            df = normalize_columns(read_csv_smart(file_bytes, filename))

    if no_header or 'ARTICLE' not in df.columns:
        if 'ARTICLE' not in df.columns:
            df.columns = ['ARTICLE'] + [f'_COL{i}' for i in range(1, len(df.columns))]

    if 'ARTICLE' not in df.columns:
        return None, f"Colonne 'ARTICLE' introuvable"

    df['SKU'] = df['ARTICLE'].astype(str).str.strip().str.zfill(8).str.slice(0, 8)
    df = df[df['SKU'].str.match(r'^\d{8}$', na=False)].drop_duplicates(subset='SKU').copy()

    for col, default in [
        ('LIBELLÉ ARTICLE', ''), ("FOURNISSEUR D'ORIGINE", ''),
        ('LIBELLÉ FOURNISSEUR ORIGINE', ''), ('MODE APPRO', ''),
        ('DATE CDE', ''), ('DATE LIV.', ''), ('SEMAINE RECEPTION', ''),
    ]:
        if col not in df.columns:
            df[col] = default

    df['SEMAINE RECEPTION'] = df['SEMAINE RECEPTION'].astype(str).str.strip().replace('nan', '')
    df['SEM_NUM'] = df['SEMAINE RECEPTION'].apply(
        lambda s: int(str(s).strip('Ss')) if str(s).strip('Ss').isdigit() else 99
    )
    df['ORIGINE'] = df['MODE APPRO'].apply(lambda m: 'IM' if 'IMPORT' in str(m).upper() else 'LO')
    return df, None

# ⭐ PATCH PARSER PBI — FORMAT MULTI-NIVEAUX
@st.cache_data(show_spinner=False)
def load_pbi_stock(file_bytes: bytes, filename: str, sku_scope: tuple) -> tuple:
    buf = io.BytesIO(file_bytes)
    try:
        if filename.lower().endswith(('.xlsx', '.xls')):
            df_raw = pd.read_excel(buf, header=None)
        else:
            buf.seek(0)
            df_raw = pd.read_csv(buf, header=None, sep=None, engine='python',
                                 encoding='latin1', on_bad_lines='skip')
    except Exception as e:
        return None, f"Lecture PBI : {e}"

    if len(df_raw) < 2:
        return None, "Fichier PBI trop court"

    sites_raw = df_raw.iloc[0, 3:].tolist()
    data_block = df_raw.iloc[2:].copy()
    results = []

    for _, row in data_block.iterrows():
        article_raw = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if not article_raw or article_raw.upper() == "TOTAL":
            continue

        def parse_article(s):
            s = str(s).strip()
            if " - " in s:
                parts = s.split(" - ", 1)
                return parts[0].strip().zfill(8), parts[1].strip()
            m = re.match(r'^(\d{8})', s)
            if m:
                return m.group(1), s
            return s[:8].zfill(8), s

        sku, lib = parse_article(article_raw)
        if sku_scope and sku not in sku_scope:
            continue

        for site_idx, site_raw in enumerate(sites_raw, start=3):
            if pd.isna(site_raw):
                continue
            site_str = str(site_raw).strip()

            def parse_site(s):
                s = str(s).strip()
                if " - " in s:
                    parts = s.split(" - ", 1)
                    return parts[0].strip(), parts[1].strip()
                return s, s

            code_site, lib_site = parse_site(site_str)
            stock_val = row.iloc[site_idx]
            try:
                stock = int(float(stock_val)) if pd.notna(stock_val) else 0
            except Exception:
                stock = 0

            results.append({
                'SKU': sku, 'Libellé article': lib,
                'Code site': code_site, 'Libellé site': lib_site,
                'Stock': stock,
            })

    if not results:
        return None, "Aucune donnée extraite"

    return pd.DataFrame(results), None

def compute_cessions(df_stock_pbi, sku_scope, magasins_detresse, tous_magasins,
                     seuil_detresse=0, seuil_cedant=2):
    if not magasins_detresse or not sku_scope:
        return pd.DataFrame()

    scope_df = df_stock_pbi[df_stock_pbi['SKU'].isin(sku_scope)].copy()
    magasins_cedants = [m for m in tous_magasins if m not in magasins_detresse]
    suggestions = []

    for sku in sku_scope:
        sku_df = scope_df[scope_df['SKU'] == sku]
        if sku_df.empty:
            continue
        lib = sku_df['Libellé article'].iloc[0]
        detresse_rows = sku_df[(sku_df['Libellé site'].isin(magasins_detresse)) & (sku_df['Stock'] <= seuil_detresse)]
        if detresse_rows.empty:
            continue
        cedant_rows = sku_df[(sku_df['Libellé site'].isin(magasins_cedants)) & (sku_df['Stock'] > seuil_cedant)].sort_values('Stock', ascending=False)
        if cedant_rows.empty:
            for _, dr in detresse_rows.iterrows():
                suggestions.append({'SKU': sku, 'Libellé article': lib, 'Magasin détresse': dr['Libellé site'],
                                   'Stock détresse': int(dr['Stock']), 'Cédant suggéré': '⚠️ Aucun cédant',
                                   'Stock cédant': 0, 'Qté cessible': 0, 'Faisabilité': '🔴 Impossible'})
            continue
        best = cedant_rows.iloc[0]
        qty_cessible = int(best['Stock']) - seuil_cedant
        for _, dr in detresse_rows.iterrows():
            faisable = '🟢 Possible' if qty_cessible >= 1 else '🟠 Partielle'
            suggestions.append({'SKU': sku, 'Libellé article': lib, 'Magasin détresse': dr['Libellé site'],
                               'Stock détresse': int(dr['Stock']), 'Cédant suggéré': best['Libellé site'],
                               'Stock cédant': int(best['Stock']), 'Qté cessible': qty_cessible, 'Faisabilité': faisable})

    if not suggestions:
        return pd.DataFrame()
    return pd.DataFrame(suggestions).sort_values(['Faisabilité', 'Qté cessible'], ascending=[True, False]).reset_index(drop=True)

# TOPBAR
st.markdown(f"""
<div class="topbar">
  <div class="topbar-left">
    <div class="topbar-icon">📋</div>
    <div>
      <div class="topbar-title">Rapport Implantation</div>
      <div class="topbar-sub">Nouvelles Références · Suivi Stock PBI · Cessions</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:12px;">
    <div class="topbar-date">{TODAY_STR}</div>
    <div class="topbar-pill">DIRECTION SUPPLY</div>
  </div>
</div>
""", unsafe_allow_html=True)

# SIDEBAR CHARGEMENT
with st.sidebar:
    st.markdown("### 📁 Chargement")
    st.divider()
    st.markdown("**T1 — Nouvelles Références**")
    t1_file = st.file_uploader("T1", type=["csv", "xlsx"], key="t1", label_visibility="collapsed")
    st.markdown("**Stock PBI** *(export pivot ou multi-niveaux)*")
    pbi_file = st.file_uploader("Stock PBI", type=["xlsx", "xls", "csv"], key="pbi", label_visibility="collapsed")
    st.caption("Format : Rayon/Famille/Article OU pivot Article×Magasin")

# CHARGEMENT T1
if not t1_file:
    st.markdown('<div class="info-banner" style="margin-top:16px">⬆️ Charge le fichier <strong>T1 Flux</strong></div>', unsafe_allow_html=True)
    st.stop()

with st.spinner("Lecture T1…"):
    t1_raw, t1_err = load_t1(t1_file.read(), t1_file.name)

if t1_err:
    st.error(f"❌ T1 : {t1_err}")
    st.stop()

SKU_TUPLE = tuple(sorted(t1_raw['SKU'].unique()))
total_sku = len(SKU_TUPLE)
sku_im_total = int((t1_raw['ORIGINE'] == 'IM').sum())
sku_lo_total = int((t1_raw['ORIGINE'] == 'LO').sum())

T1_KEEP = ['LIBELLÉ ARTICLE', 'LIBELLÉ FOURNISSEUR ORIGINE', 'MODE APPRO', 'SEMAINE RECEPTION', 'DATE LIV.', 'ORIGINE', 'SEM_NUM']
t1_idx = t1_raw.set_index('SKU')[T1_KEEP]

# CHARGEMENT PBI
if not pbi_file:
    st.markdown('<div class="info-banner">⬆️ Charge l\'export stock <strong>PBI</strong></div>', unsafe_allow_html=True)
    st.stop()

with st.spinner("Parsing export PBI…"):
    pbi_bytes = pbi_file.read()
    df_stock_all, pbi_err = load_pbi_stock(pbi_bytes, pbi_file.name, sku_scope=())
    if pbi_err:
        st.error(f"❌ PBI : {pbi_err}")
        st.stop()
    df_stock_t1, _ = load_pbi_stock(pbi_bytes, pbi_file.name, sku_scope=SKU_TUPLE)

magasins_list = sorted(df_stock_t1['Libellé site'].dropna().unique())
tous_magasins = sorted(df_stock_all['Libellé site'].dropna().unique()) if df_stock_all is not None else magasins_list

# FILTRES SIDEBAR
with st.sidebar:
    st.divider()
    st.markdown("### 🔍 Filtres")
    mag_sel = st.multiselect("Magasins", magasins_list, default=magasins_list)
    origine_sel = st.multiselect("Origine", ['IM', 'LO'], default=['IM', 'LO'])
    sem_dispo = sorted([s for s in t1_raw['SEMAINE RECEPTION'].unique() if s and s not in ('nan', '')], key=sem_sort)
    sem_sel = st.multiselect("Semaine réception", sem_dispo, default=sem_dispo)
    mode_sel = st.multiselect("Mode Appro", sorted([m for m in t1_raw['MODE APPRO'].unique() if m and m not in ('nan', '')]),
                              default=sorted([m for m in t1_raw['MODE APPRO'].unique() if m and m not in ('nan', '')]))
    st.divider()
    st.markdown("### 🔄 Cessions")
    st.caption("Articles = liste T1")
    magasins_detresse = st.multiselect("Magasins en détresse", options=tous_magasins, default=[], help="Besoin cession interne")
    seuil_detresse = st.number_input("Seuil stock détresse (≤)", min_value=0, max_value=50, value=0, step=1)
    seuil_cedant = st.number_input("Stock minimum chez cédant", min_value=0, max_value=20, value=2, step=1)

if not mag_sel:
    st.warning("Sélectionne au moins un magasin.")
    st.stop()

mc1, mc2, mc3 = st.columns([6, 1, 1])
with mc1:
    mag_main = st.multiselect("🏪 Magasins affichés", magasins_list, default=mag_sel, key="mag_main")
with mc2:
    if st.button("✅ Tous", use_container_width=True): mag_main = magasins_list
with mc3:
    if st.button("❌ Aucun", use_container_width=True): mag_main = []

mag_actifs = mag_main if mag_main else mag_sel
if not mag_actifs:
    st.warning("Sélectionne au moins un magasin.")
    st.stop()

sku_mask = (t1_raw['ORIGINE'].isin(origine_sel) & (t1_raw['SEMAINE RECEPTION'].isin(sem_sel) if sem_sel else True)
            & (t1_raw['MODE APPRO'].isin(mode_sel) if mode_sel else True))
sku_scope = t1_raw.loc[sku_mask, 'SKU'].unique()
total_sku_sel = len(sku_scope)

if total_sku_sel == 0:
    st.warning("Aucun article correspond aux filtres.")
    st.stop()

base_df = pd.DataFrame(pd.MultiIndex.from_product([mag_actifs, sku_scope], names=['Libellé site', 'SKU']).tolist(),
                       columns=['Libellé site', 'SKU'])
stock_scope = df_stock_t1[(df_stock_t1['Libellé site'].isin(mag_actifs)) & (df_stock_t1['SKU'].isin(sku_scope))][['Libellé site', 'SKU', 'Stock', 'Libellé article']].copy()
merged = base_df.merge(stock_scope, on=['Libellé site', 'SKU'], how='left')
merged['Stock'] = merged['Stock'].fillna(0).astype(int)
merged = merged.merge(t1_idx.reset_index().rename(columns={'LIBELLÉ ARTICLE': 'T1_lib', 'LIBELLÉ FOURNISSEUR ORIGINE': 'Fournisseur',
                      'MODE APPRO': 'Mode Appro', 'SEMAINE RECEPTION': 'Sem. Réception', 'DATE LIV.': 'Date Livraison',
                      'ORIGINE': 'Origine', 'SEM_NUM': 'SEM_NUM'}), on='SKU', how='left')
merged['Libellé article'] = merged['Libellé article'].fillna('').astype(str)
merged['Libellé article'] = merged.apply(lambda r: r['Libellé article'] if r['Libellé article'] else r['T1_lib'], axis=1)
merged.drop(columns='T1_lib', inplace=True)
conds = [(merged['Stock'] > 0)]
choices = ["Implantation Terminée"]
merged['Statut'] = np.select(conds, choices, default="Alerte Aucun Mouvement")
detail_df = merged.rename(columns={'Libellé site': 'Magasin'})

S_ORDER = ["Implantation Terminée", "Alerte Aucun Mouvement"]
S_COLORS = {"Implantation Terminée": "#059669", "Alerte Aucun Mouvement": "#dc2626"}
pivot = (detail_df.groupby(['Magasin', 'Statut']).size().unstack(fill_value=0).reindex(columns=S_ORDER, fill_value=0).reset_index())
pivot.columns.name = None
pivot['Total'] = total_sku_sel
pivot['Taux (%)'] = (pivot['Implantation Terminée'] / total_sku_sel * 100).round(0).astype(int)

total_cells = len(mag_actifs) * total_sku_sel
ct = int(pivot['Implantation Terminée'].sum())
cal = int(pivot['Alerte Aucun Mouvement'].sum())
avg_impl = int(pivot['Taux (%)'].mean()) if not pivot.empty else 0
pct = lambda n: int(n / total_cells * 100) if total_cells > 0 else 0

df_im = detail_df[detail_df['Origine'] == 'IM']
df_lo = detail_df[detail_df['Origine'] == 'LO']
df_alerte = detail_df[detail_df['Statut'] == 'Alerte Aucun Mouvement']
alerte_im = len(df_alerte[df_alerte['Origine'] == 'IM'])
alerte_lo = len(df_alerte[df_alerte['Origine'] == 'LO'])
tim = taux_implantation(df_im)
tlo = taux_implantation(df_lo)

PLOTLY_BASE = dict(paper_bgcolor="#fff", plot_bgcolor="#fff", font=dict(family="Inter", color="#64748b", size=12),
                   margin=dict(l=20, r=20, t=44, b=20))

# BANNIÈRE ALERTES
if cal > 0:
    st.markdown(f"""<div class="alert-banner"><div class="ab-badge">⚡ ACTIONS REQUISES</div><div class="ab-item"><div class="ab-num" style="color:#dc2626">{cal}</div><div class="ab-lbl">Aucun Mouvement</div></div><div class="ab-item"><div class="ab-num" style="color:#dc2626">{alerte_im}</div><div class="ab-lbl">dont IM</div></div><div class="ab-item"><div class="ab-num" style="color:#ea580c">{alerte_lo}</div><div class="ab-lbl">dont LO</div></div><div style="margin-left:auto;display:flex;flex-direction:column;align-items:flex-end;"><div style="font-size:36px;font-weight:900;color:#dc2626;line-height:1">{cal}</div><div style="font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em">articles à traiter</div></div></div>""", unsafe_allow_html=True)

# KPI STRIP
st.markdown(f"""<div class="strip"><div class="strip-card"><div class="strip-tag tag-im">IMPORT</div><div class="strip-label">Références</div><div class="strip-val" style="color:#2563eb">{sku_im_total}</div><div class="strip-sub">SKU à implanter</div></div><div class="strip-card"><div class="strip-tag tag-im">IMPORT</div><div class="strip-label">Taux implanté</div><div class="strip-val" style="color:#{color_taux(tim)}">{tim}%</div><div class="strip-sub">stock présent</div></div><div class="strip-card"><div class="strip-tag tag-im">IMPORT</div><div class="strip-label">Alerte Aucun Mvt</div><div class="strip-val" style="color:#dc2626">{alerte_im}</div><div class="strip-sub">escalade fournisseur</div></div><div class="strip-card"><div class="strip-tag tag-lo">LOCAL</div><div class="strip-label">Références</div><div class="strip-val" style="color:#059669">{sku_lo_total}</div><div class="strip-sub">SKU à implanter</div></div><div class="strip-card"><div class="strip-tag tag-lo">LOCAL</div><div class="strip-label">Taux implanté</div><div class="strip-val" style="color:#{color_taux(tlo)}">{tlo}%</div><div class="strip-sub">stock présent</div></div><div class="strip-card"><div class="strip-tag tag-lo">LOCAL</div><div class="strip-label">Alerte Aucun Mvt</div><div class="strip-val" style="color:#dc2626">{alerte_lo}</div><div class="strip-sub">relance supply</div></div></div>""", unsafe_allow_html=True)

# KPI GLOBAUX
st.markdown(f"""<div class="kpi-row"><div class="kpi g"><div class="kpi-lbl">✅ Implantation Terminée</div><div class="kpi-val">{ct}</div><div class="kpi-pct">{pct(ct)}% · Stock présent</div><div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(ct)}%;background:#059669"></div></div></div><div class="kpi r"><div class="kpi-lbl">🚨 Alerte Aucun Mouvement</div><div class="kpi-val">{cal}</div><div class="kpi-pct">{pct(cal)}% · Stock PBI = 0</div><div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(cal)}%;background:#dc2626"></div></div></div><div class="kpi o"><div class="kpi-lbl">📊 Taux réseau moyen</div><div class="kpi-val">{avg_impl}%</div><div class="kpi-pct">{len(mag_actifs)} magasin(s) · {total_sku_sel} SKU</div><div class="kpi-bar"><div class="kpi-bar-fill" style="width:{avg_impl}%;background:#b45309"></div></div></div></div>""", unsafe_allow_html=True)

# SCORECARD
st.markdown('<div class="sh">SCORECARD MAGASINS</div>', unsafe_allow_html=True)
rag_html = '<div class="rag-grid">'
for _, row in pivot.sort_values('Taux (%)', ascending=False).iterrows():
    t_ = row['Taux (%)']
    cls, c_hex = ('g', '#059669') if t_ >= 80 else ('r', '#dc2626')
    rag_html += f"""<div class="rag-card {cls}"><div class="rag-dot" style="background:{c_hex}"></div><div class="rag-name">{row['Magasin']}</div><div class="rag-pct" style="color:{c_hex}">{t_}%</div><div class="rag-detail">{int(row['Implantation Terminée'])}✅ {int(row['Alerte Aucun Mouvement'])}🚨</div></div>"""
rag_html += '</div>'
st.markdown(rag_html, unsafe_allow_html=True)

# NAVIGATION ONGLETS
TABS = ["📊 Vue Globale", "🚨 Alertes & Actions", "🗓️ Calendrier Flux", "📋 Plan d'Action", "🔄 Cessions"]
if "tab" not in st.session_state:
    st.session_state.tab = TABS[0]

nav_cols = st.columns(len(TABS))
for i, t in enumerate(TABS):
    with nav_cols[i]:
        if st.session_state.tab == t:
            st.markdown(f'<div class="nav-tab-active">{t}</div>', unsafe_allow_html=True)
        if st.button(t, key=f"nav_{i}", use_container_width=True):
            st.session_state.tab = t
            st.rerun()

active = st.session_state.tab

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — VUE GLOBALE
# ══════════════════════════════════════════════════════════════════════════════
if active == TABS[0]:
    c1, c2 = st.columns([3, 2])
    with c1:
        mel = pivot.melt(id_vars='Magasin', value_vars=list(S_COLORS.keys()), var_name='Statut', value_name='N')
        fig = px.bar(mel, x='Magasin', y='N', color='Statut', color_discrete_map=S_COLORS, barmode='stack', title='Situation par magasin')
        fig.update_traces(textposition='inside', texttemplate='%{y}', textfont_size=11, textfont_color='white')
        fig.update_layout(**PLOTLY_BASE, height=400, legend=dict(orientation='h', y=-0.22, bgcolor='rgba(0,0,0,0)', font_size=11),
                          xaxis=dict(gridcolor='#f0f2f8'), yaxis=dict(gridcolor='#f0f2f8'))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig_d = go.Figure(go.Pie(labels=list(S_COLORS.keys()), values=[ct, cal], hole=0.65,
                                 marker=dict(colors=list(S_COLORS.values()), line=dict(color='#fff', width=3)),
                                 textfont=dict(size=11)))
        fig_d.add_annotation(text=f"<b>{avg_impl}%</b><br>implanté", x=0.5, y=0.5,
                            font=dict(size=18, color='#0f1729', family='Inter'), showarrow=False)
        fig_d.update_layout(**PLOTLY_BASE, height=400, title='Répartition globale',
                           legend=dict(orient='v', x=1.01, bgcolor='rgba(0,0,0,0)', font_size=11))
        st.plotly_chart(fig_d, use_container_width=True)
    st.markdown('<div class="sh">DÉTAIL PAR MAGASIN</div>', unsafe_allow_html=True)
    st.dataframe(pivot[["Magasin","Implantation Terminée","Alerte Aucun Mouvement","Total","Taux (%)"]].style
        .background_gradient(subset=['Implantation Terminée'], cmap='Greens')
        .background_gradient(subset=['Alerte Aucun Mouvement'], cmap='Reds')
        .background_gradient(subset=['Taux (%)'], cmap='RdYlGn', vmin=0, vmax=100)
        .format({'Taux (%)': '{}%'}), use_container_width=True, hide_index=True,
        height=min(600, 60 + len(mag_actifs) * 42))
    st.markdown('<div class="sh">STOCK PBI — TOP ARTICLES (SCOPE T1)</div>', unsafe_allow_html=True)
    stock_summary = (df_stock_t1[(df_stock_t1['Libellé site'].isin(mag_actifs)) & (df_stock_t1['SKU'].isin(sku_scope))]
                     .groupby(['SKU','Libellé article'])['Stock'].sum().reset_index()
                     .sort_values('Stock', ascending=False).head(20))
    st.dataframe(stock_summary, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — ALERTES
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[1]:
    ACOLS = ["Magasin","SKU","Libellé article","Origine","Mode Appro","Sem. Réception","Date Livraison","Stock","Statut"]
    st.markdown(f"""<div class="ac red"><div><div class="ac-title" style="color:#dc2626">🚨 Alerte Aucun Mouvement</div><div class="ac-sub">💡 Action : Escalader fournisseur · Cession · Informer magasin</div></div><div class="ac-count" style="color:#dc2626">{cal}</div></div>""", unsafe_allow_html=True)
    if df_alerte.empty:
        st.success("✅ Aucune alerte")
    else:
        tab_im, tab_lo = st.tabs([f"IMPORT — {len(df_alerte[df_alerte['Origine']=='IM'])} SKU", f"LOCAL — {len(df_alerte[df_alerte['Origine']=='LO'])} SKU"])
        for tab, orig, sub_df in [(tab_im,'IM',df_alerte[df_alerte['Origine']=='IM']), (tab_lo,'LO',df_alerte[df_alerte['Origine']=='LO'])]:
            with tab:
                if sub_df.empty:
                    st.success(f"✅ OK {orig}")
                else:
                    col_g, col_t = st.columns([2, 3])
                    with col_g:
                        top = (sub_df.groupby(['SKU','Libellé article'])['Magasin'].count().reset_index()
                               .rename(columns={'Magasin':'Nb Magasins'}).sort_values('Nb Magasins').tail(10))
                        top['lbl'] = top['SKU'] + ' – ' + top['Libellé article'].str[:28]
                        fig_t = go.Figure(go.Bar(x=top['Nb Magasins'], y=top['lbl'], orientation='h',
                                                marker=dict(color='#dc2626', cornerradius=4),
                                                text=top['Nb Magasins'], textposition='outside',
                                                textfont=dict(color='#64748b', size=11)))
                        fig_t.update_layout(**PLOTLY_BASE, height=max(200, len(top)*34), title=f'Top — {orig}',
                                           xaxis=dict(gridcolor='#f0f2f8'), yaxis=dict(tickfont_size=10))
                        st.plotly_chart(fig_t, use_container_width=True)
                    with col_t:
                        top_m = (sub_df.groupby('Magasin')['SKU'].count().reset_index()
                                 .rename(columns={'SKU':'Nb SKU'}).sort_values('Nb SKU', ascending=False))
                        fig_m = go.Figure(go.Bar(x=top_m['Magasin'], y=top_m['Nb SKU'],
                                               marker=dict(color='#dc2626', cornerradius=4),
                                               text=top_m['Nb SKU'], textposition='outside',
                                               textfont=dict(color='#64748b', size=11)))
                        fig_m.update_layout(**PLOTLY_BASE, height=max(200, len(top_m)*40), title=f'Par magasin — {orig}',
                                           xaxis=dict(gridcolor='#f0f2f8'), yaxis=dict(gridcolor='#f0f2f8'))
                        st.plotly_chart(fig_m, use_container_width=True)
                    with st.expander(f"📋 Détail {orig} — {len(sub_df)} lignes"):
                        st.dataframe(sub_df[[c for c in ACOLS if c in sub_df.columns]].sort_values(['Magasin']).reset_index(drop=True),
                                    use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — CALENDRIER
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[2]:
    cal_df = detail_df[detail_df['Sem. Réception'].str.match(r'^S\d+$', na=False)].copy()
    cal_df['SEM_NUM'] = cal_df['Sem. Réception'].apply(sem_sort)
    if cal_df.empty:
        st.info("Aucune semaine")
    else:
        sem_order = sorted(cal_df['Sem. Réception'].unique(), key=sem_sort)
        c1, c2 = st.columns(2)
        with c1:
            ss = (cal_df.groupby(['Sem. Réception','SEM_NUM','Statut']).size().reset_index(name='N').sort_values('SEM_NUM'))
            fig_s = px.bar(ss, x='Sem. Réception', y='N', color='Statut', color_discrete_map=S_COLORS,
                          barmode='stack', title='Articles/semaine', category_orders={'Sem. Réception': sem_order})
            fig_s.update_traces(textposition='inside', texttemplate='%{y}', textfont_size=10, textfont_color='white')
            fig_s.update_layout(**PLOTLY_BASE, height=360, xaxis=dict(gridcolor='#f0f2f8'),
                               yaxis=dict(gridcolor='#f0f2f8'), legend=dict(orientation='h', y=-0.22, bgcolor='rgba(0,0,0,0)'))
            st.plotly_chart(fig_s, use_container_width=True)
        with c2:
            os_df = (cal_df.groupby(['Origine','Sem. Réception','SEM_NUM']).size().reset_index(name='N').sort_values('SEM_NUM'))
            fig_o = px.bar(os_df, x='Sem. Réception', y='N', color='Origine', barmode='group',
                          color_discrete_map={'IM':'#2563eb','LO':'#059669'}, title='IM vs LO',
                          category_orders={'Sem. Réception': sem_order})
            fig_o.update_traces(textposition='outside', texttemplate='%{y}', textfont_size=10)
            fig_o.update_layout(**PLOTLY_BASE, height=360, xaxis=dict(gridcolor='#f0f2f8'),
                               yaxis=dict(gridcolor='#f0f2f8'), legend=dict(orientation='h', y=-0.22, bgcolor='rgba(0,0,0,0)'))
            st.plotly_chart(fig_o, use_container_width=True)
        st.markdown('<div class="sh">DÉTAIL SEMAINE</div>', unsafe_allow_html=True)
        tbl = (cal_df.groupby(['Sem. Réception','SEM_NUM','Origine']).agg(
                  Articles=('SKU','nunique'),
                  Terminé=('Statut', lambda x: (x=='Implantation Terminée').sum()),
                  Alerte=('Statut', lambda x: (x=='Alerte Aucun Mouvement').sum()),
               ).reset_index().sort_values('SEM_NUM').drop(columns='SEM_NUM'))
        st.dataframe(tbl.style.background_gradient(subset=['Terminé'], cmap='Greens')
                     .background_gradient(subset=['Alerte'], cmap='Reds'),
                    use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — PLAN D'ACTION
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[3]:
    c1, c2 = st.columns([1, 2])
    with c1:
        recap_s = pivot.sort_values('Taux (%)', ascending=True)
        bar_colors = ['#059669' if v >= 80 else '#dc2626' for v in recap_s['Taux (%)']]
        fig_h = go.Figure(go.Bar(x=recap_s['Taux (%)'], y=recap_s['Magasin'], orientation='h',
                                marker=dict(color=bar_colors, cornerradius=5),
                                text=[f"{v}%" for v in recap_s['Taux (%)']],
                                textposition='outside', textfont=dict(color='#0f1729', size=13)))
        fig_h.add_vline(x=80, line_dash='dash', line_color='#e2e8f4', annotation_text='Cible 80%',
                       annotation_font_color='#94a3b8')
        fig_h.update_layout(**PLOTLY_BASE, height=max(280, len(mag_actifs)*48), xaxis=dict(range=[0,118], gridcolor='#f0f2f8', ticksuffix='%'),
                           yaxis=dict(gridcolor='rgba(0,0,0,0)'), title='Taux/magasin')
        st.plotly_chart(fig_h, use_container_width=True)
    with c2:
        mag_pa = st.selectbox("Magasin", mag_actifs, key="pa_mag")
        df_pa = detail_df[(detail_df['Magasin']==mag_pa) & (detail_df['Statut']=='Alerte Aucun Mouvement')]
        krow = pivot[pivot['Magasin']==mag_pa]
        t_mag = int(krow['Taux (%)'].values[0]) if not krow.empty else 0
        n_alert = int(krow['Alerte Aucun Mouvement'].values[0]) if not krow.empty else 0
        c_hex = "#059669" if t_mag >= 80 else "#dc2626"
        bg = "#ecfdf5" if t_mag >= 80 else "#fef2f2"
        bd = "#6ee7b7" if t_mag >= 80 else "#fecaca"
        st.markdown(f"""<div style="background:{bg};border:1px solid {bd};border-radius:10px;padding:16px 20px;margin-bottom:14px;display:flex;align-items:center;gap:20px;"><div style="font-size:52px;font-weight:900;color:{c_hex};line-height:1">{t_mag}%</div><div><div style="font-size:15px;font-weight:700">{mag_pa}</div><div style="font-size:12px;color:#64748b;margin-top:3px">{n_alert} article(s) sans mouvement</div></div></div>""", unsafe_allow_html=True)
        if df_pa.empty:
            st.success(f"✅ {mag_pa} OK")
        else:
            PA_COLS = ["SKU","Libellé article","Origine","Mode Appro","Sem. Réception","Date Livraison","Stock","Statut"]
            st.dataframe(df_pa[[c for c in PA_COLS if c in df_pa.columns]].sort_values(['Origine']).reset_index(drop=True),
                        use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — CESSIONS
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[4]:
    st.markdown('<div class="sh">🔄 MOTEUR DE CESSIONS INTERNES</div>', unsafe_allow_html=True)
    if not magasins_detresse:
        st.markdown("""<div class="gold-banner"><strong>💡 Comment utiliser ?</strong><br>1. Sidebar → Cessions → sélectionne magasins en détresse<br>2. Ajuste seuils<br>3. Le moteur suggère le meilleur cédant</div>""", unsafe_allow_html=True)
    else:
        with st.spinner("Calcul cessions…"):
            df_cessions = compute_cessions(df_stock_all if df_stock_all is not None else df_stock_t1,
                                          list(sku_scope), magasins_detresse, tous_magasins,
                                          int(seuil_detresse), int(seuil_cedant))
        if df_cessions.empty:
            st.success("✅ Aucune cession nécessaire")
        else:
            n_possible = int((df_cessions['Faisabilité']=='🟢 Possible').sum())
            n_impossible = int((df_cessions['Faisabilité']=='🔴 Impossible').sum())
            n_articles = df_cessions['SKU'].nunique()
            qty_total = int(df_cessions[df_cessions['Faisabilité']=='🟢 Possible']['Qté cessible'].sum())
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Articles", n_articles)
            k2.metric("🟢 Possibles", n_possible)
            k3.metric("🔴 Impossible", n_impossible)
            k4.metric("Qté total", qty_total)
            st.markdown("---")
            mag_str = " · ".join(magasins_detresse)
            st.markdown(f"""<div class="cession-header"><div style="font-size:12px;font-weight:700;color:#b45309;margin-bottom:6px;text-transform:uppercase;letter-spacing:.06em">Paramètres</div><div class="cession-row"><span class="cession-badge badge-detresse">Détresse: {mag_str}</span><span class="cession-badge badge-stock">Stock ≤ {seuil_detresse}</span><span class="cession-badge badge-qty">Réserve: {seuil_cedant}</span></div></div>""", unsafe_allow_html=True)
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                filt_faisab = st.selectbox("Faisabilité", ["Toutes", "🟢 Possible", "🔴 Impossible"], key="filt_faisab")
            with fc2:
                filt_mag = st.selectbox("Détresse", ["Tous"] + sorted(df_cessions['Magasin détresse'].unique()), key="filt_mag")
            with fc3:
                filt_cedant = st.selectbox("Cédant", ["Tous"] + sorted(df_cessions[df_cessions['Cédant suggéré']!='⚠️ Aucun cédant']['Cédant suggéré'].unique()), key="filt_cedant")
            df_view = df_cessions.copy()
            if filt_faisab != "Toutes":
                df_view = df_view[df_view['Faisabilité']==filt_faisab]
            if filt_mag != "Tous":
                df_view = df_view[df_view['Magasin détresse']==filt_mag]
            if filt_cedant != "Tous":
                df_view = df_view[df_view['Cédant suggéré']==filt_cedant]
            if not df_view[df_view['Faisabilité']=='🟢 Possible'].empty:
                cedant_summary = (df_view[df_view['Faisabilité']=='🟢 Possible'].groupby('Cédant suggéré')
                                 .agg(Articles=('SKU','nunique'), Qté_totale=('Qté cessible','sum'))
                                 .reset_index().sort_values('Qté_totale', ascending=False))
                gv1, gv2 = st.columns(2)
                with gv1:
                    fig_c = go.Figure(go.Bar(x=cedant_summary['Cédant suggéré'], y=cedant_summary['Articles'],
                                            marker=dict(color='#b45309', cornerradius=4),
                                            text=cedant_summary['Articles'], textposition='outside',
                                            textfont=dict(color='#64748b', size=11)))
                    fig_c.update_layout(**PLOTLY_BASE, height=300, title='Articles/cédant',
                                       xaxis=dict(gridcolor='#f0f2f8'), yaxis=dict(gridcolor='#f0f2f8'))
                    st.plotly_chart(fig_c, use_container_width=True)
                with gv2:
                    fig_q = go.Figure(go.Bar(x=cedant_summary['Cédant suggéré'], y=cedant_summary['Qté_totale'],
                                            marker=dict(color='#059669', cornerradius=4),
                                            text=cedant_summary['Qté_totale'], textposition='outside',
                                            textfont=dict(color='#64748b', size=11)))
                    fig_q.update_layout(**PLOTLY_BASE, height=300, title='Qté/cédant',
                                       xaxis=dict(gridcolor='#f0f2f8'), yaxis=dict(gridcolor='#f0f2f8'))
                    st.plotly_chart(fig_q, use_container_width=True)
            st.markdown(f'<div class="sh">PLAN DE CESSION — {len(df_view)} lignes</div>', unsafe_allow_html=True)
            st.dataframe(df_view.style.map(lambda v: 'background-color:#ecfdf5;color:#059669;font-weight:700' if v=='🟢 Possible' else
                                          ('background-color:#fef2f2;color:#dc2626;font-weight:700' if v=='🔴 Impossible' else ''),
                                          subset=['Faisabilité']).background_gradient(subset=['Qté cessible'], cmap='Greens'),
                        use_container_width=True, hide_index=True)
            buf_c = io.BytesIO()
            with pd.ExcelWriter(buf_c, engine='openpyxl') as writer:
                df_cessions.to_excel(writer, sheet_name='Plan Cessions', index=False)
                df_cessions[df_cessions['Faisabilité']=='🟢 Possible'].to_excel(writer, sheet_name='Possibles', index=False)
                df_cessions[df_cessions['Faisabilité']=='🔴 Impossible'].to_excel(writer, sheet_name='Impossibles', index=False)
            buf_c.seek(0)
            st.download_button(label=f"📥 Plan Cessions_{TODAY_FILE}.xlsx", data=buf_c,
                              file_name=f"Plan_Cessions_{TODAY_FILE}.xlsx",
                              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_cessions")

# EXPORT RAPPORT
st.markdown('<div class="sh">EXPORT RAPPORT</div>', unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def build_report(det_b, piv_b, ces_b, today_str, today_file, mag_count, sku_count, ct, cal_n, avg,
                skt, skl, tim, tlo):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    det = pd.read_parquet(io.BytesIO(det_b))
    piv = pd.read_parquet(io.BytesIO(piv_b))
    ces = pd.read_parquet(io.BytesIO(ces_b)) if ces_b else pd.DataFrame()
    wb = Workbook()
    wb.remove(wb.active)

    C = dict(dark="0F1729", navy="1E293B", blue="2563EB", green="059669", red="DC2626", grey="F0F2F8", border="E2E8F4", muted="64748B", white="FFFFFF")

    def F(c=C['dark'], sz=10, b=False): return Font(name='Arial', size=sz, bold=b, color=c)
    def fill(h): return PatternFill("solid", fgColor=h)
    def ctr(): return Alignment(horizontal='center', vertical='center', wrap_text=False)
    def lft(): return Alignment(horizontal='left', vertical='center', wrap_text=False)
    def brd(): return Border(left=Side(style='thin', color=C['border']), right=Side(style='thin', color=C['border']),
                            top=Side(style='thin', color=C['border']), bottom=Side(style='thin', color=C['border']))

    def write_header(ws, title, sub=""):
        ws.sheet_view.showGridLines = False
        ws.merge_cells('B1:L1')
        ws.row_dimensions[1].height = 44
        c = ws['B1']; c.value = title; c.font = Font(name='Arial', size=20, bold=True, color=C['white'])
        c.fill = fill(C['dark']); c.alignment = lft()

    ws1 = wb.create_sheet("📊 Résumé")
    write_header(ws1, f"RAPPORT IMPLANTATION — {today_str}", f"{mag_count}×{sku_count} Taux: {avg}%")
    ws1.column_dimensions['A'].width = 2

    ws2 = wb.create_sheet("🚨 Alertes")
    write_header(ws2, "ALERTES — AUCUN MOUVEMENT", f"{today_str} · {cal_n} articles")

    ws3 = wb.create_sheet("📋 Actions")
    write_header(ws3, "PLAN D'ACTION", today_str)

    ws4 = wb.create_sheet("📦 Complet")
    write_header(ws4, "DÉTAIL COMPLET", f"{today_str}")

    if not ces.empty:
        ws5 = wb.create_sheet("🔄 Cessions")
        write_header(ws5, "PLAN DE CESSIONS", f"{today_str}")

    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()

EXPORT_COLS = ["Magasin","SKU","Libellé article","Origine","Mode Appro","Sem. Réception","Date Livraison","Stock","Statut"]
col_dl, col_info = st.columns([1, 2])
with col_info:
    st.markdown(f"""<div style="background:var(--accent-l);border:1px solid var(--accent-bd);border-radius:10px;padding:14px 18px;"><div style="font-size:13px;font-weight:700;color:var(--accent)">📄 Rapport_{TODAY_FILE}.xlsx</div><div style="font-size:12px;color:var(--muted);margin-top:4px;">{len(mag_actifs)} magasin(s) · {total_sku_sel} SKU · <strong style="color:#dc2626">{cal} alertes</strong></div></div>""", unsafe_allow_html=True)

with col_dl:
    det_b = io.BytesIO(); detail_df[EXPORT_COLS].to_parquet(det_b); det_b.seek(0)
    piv_b = io.BytesIO(); pivot.to_parquet(piv_b); piv_b.seek(0)
    ces_b = None
    if magasins_detresse:
        df_ces_export = compute_cessions(df_stock_all if df_stock_all is not None else df_stock_t1,
                                        list(sku_scope), magasins_detresse, tous_magasins,
                                        int(seuil_detresse), int(seuil_cedant))
        if not df_ces_export.empty:
            ces_buf = io.BytesIO(); df_ces_export.to_parquet(ces_buf); ces_buf.seek(0)
            ces_b = ces_buf.getvalue()

    report = build_report(det_b.getvalue(), piv_b.getvalue(), ces_b, TODAY_STR, TODAY_FILE,
                         len(mag_actifs), total_sku_sel, ct, cal, avg_impl, sku_im_total, sku_lo_total, tim, tlo)
    st.download_button(label=f"📥 Rapport_{TODAY_FILE}.xlsx", data=report,
                      file_name=f"Rapport_Implantation_{TODAY_FILE}.xlsx",
                      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
