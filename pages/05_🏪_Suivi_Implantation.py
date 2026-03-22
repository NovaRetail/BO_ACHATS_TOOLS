import streamlit as st
st.set_page_config(page_title="Suivi Implantation · SmartBuyer", page_icon="🏪", layout="wide")
st.title("🏪 Suivi Implantation")
st.info("Module en cours de développement.")
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import date

TODAY      = date.today()
TODAY_STR  = TODAY.strftime("%d %b %Y")
TODAY_FILE = TODAY.strftime("%Y%m%d")

st.set_page_config(
    page_title="Rapport Implantation · Carrefour",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ══════════════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM
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
  --radius:10px; --shadow:0 1px 3px rgba(15,23,42,.06),0 4px 16px rgba(15,23,42,.04);
}
html,body,[class*="css"]{font-family:'Inter',sans-serif!important;background:var(--bg)!important;color:var(--text)!important;}
.main,section[data-testid="stMain"]{background:var(--bg)!important;}
.block-container{padding:0 2rem 4rem!important;max-width:1520px;}
header[data-testid="stHeader"],#MainMenu,footer{display:none!important;}

/* ─── TOPBAR ─── */
.topbar{background:var(--text);margin:0 -2rem 24px;padding:14px 28px;
  display:flex;align-items:center;justify-content:space-between;}
.topbar-left{display:flex;align-items:center;gap:14px;}
.topbar-icon{width:38px;height:38px;border-radius:9px;background:linear-gradient(135deg,#3b82f6,#60a5fa);
  display:flex;align-items:center;justify-content:center;font-size:20px;}
.topbar-title{font-size:17px;font-weight:700;color:#fff;letter-spacing:-.01em;}
.topbar-sub{font-size:11px;color:#94a3b8;font-family:'JetBrains Mono';margin-top:1px;}
.topbar-pill{background:rgba(255,255,255,.08);color:#94a3b8;border:1px solid rgba(255,255,255,.12);
  border-radius:6px;padding:4px 12px;font-size:11px;font-weight:500;}
.topbar-date{color:#60a5fa;font-size:12px;font-family:'JetBrains Mono';}

/* ─── ALERT BANNER ─── */
.alert-banner{background:#fff;border:1px solid var(--red-bd);border-left:4px solid var(--red);
  border-radius:var(--radius);padding:14px 20px;margin-bottom:20px;
  display:flex;align-items:center;gap:0;flex-wrap:wrap;}
.ab-badge{background:var(--red);color:#fff;border-radius:6px;padding:4px 10px;
  font-size:11px;font-weight:700;letter-spacing:.04em;margin-right:16px;white-space:nowrap;}
.ab-item{display:flex;flex-direction:column;align-items:center;padding:0 20px;
  border-right:1px solid var(--border);}
.ab-item:last-child{border-right:none;padding-right:0;}
.ab-num{font-size:26px;font-weight:800;line-height:1;}
.ab-lbl{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-top:1px;}

/* ─── RAG SCORECARD ─── */
.rag-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px;margin-bottom:22px;}
.rag-card{border-radius:var(--radius);padding:14px 16px;border:1px solid transparent;
  box-shadow:var(--shadow);position:relative;}
.rag-card.g{background:var(--green-l);border-color:var(--green-bd);}
.rag-card.r{background:var(--red-l);border-color:var(--red-bd);}
.rag-name{font-size:11px;font-weight:600;color:var(--text);margin-bottom:5px;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:90%;}
.rag-pct{font-size:30px;font-weight:800;line-height:1;letter-spacing:-.02em;}
.rag-detail{font-size:10px;color:var(--muted);font-family:'JetBrains Mono';margin-top:3px;}
.rag-dot{width:8px;height:8px;border-radius:50%;position:absolute;top:14px;right:14px;}

/* ─── KPI STRIP ─── */
.strip{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:16px;}
.strip-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
  padding:14px 16px;box-shadow:var(--shadow);}
.strip-tag{display:inline-block;border-radius:4px;padding:2px 8px;font-size:10px;font-weight:700;margin-bottom:6px;}
.tag-im{background:#eff4ff;color:#2563eb;border:1px solid #bfdbfe;}
.tag-lo{background:#ecfdf5;color:#059669;border:1px solid #6ee7b7;}
.strip-label{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin-bottom:4px;}
.strip-val{font-size:26px;font-weight:800;line-height:1;letter-spacing:-.01em;}
.strip-sub{font-size:10px;color:var(--muted);font-family:'JetBrains Mono';margin-top:2px;}

/* ─── KPI CARDS ─── */
.kpi-row{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:22px;}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
  padding:20px 20px 16px;box-shadow:var(--shadow);position:relative;overflow:hidden;}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;}
.kpi.g::before{background:var(--green);}
.kpi.b::before{background:var(--blue);}
.kpi.r::before{background:var(--red);}
.kpi-lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:var(--muted);margin-bottom:10px;}
.kpi-val{font-size:44px;font-weight:800;line-height:1;letter-spacing:-.02em;}
.kpi.g .kpi-val{color:var(--green);}
.kpi.b .kpi-val{color:var(--blue);}
.kpi.r .kpi-val{color:var(--red);}
.kpi-pct{font-size:12px;font-weight:600;color:var(--muted);margin-top:4px;font-family:'JetBrains Mono';}
.kpi-bar{margin-top:12px;height:3px;border-radius:3px;background:var(--border);}
.kpi-bar-fill{height:100%;border-radius:3px;}

/* ─── SECTION HEADERS ─── */
.sh{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;
  color:var(--muted);margin:22px 0 12px;padding-bottom:8px;border-bottom:1px solid var(--border);}

/* ─── TABS ─── */
.nav-tab-active{background:var(--text)!important;color:#fff!important;
  border-radius:8px;padding:9px 0;text-align:center;font-size:13px;font-weight:700;
  box-shadow:0 4px 14px rgba(15,23,42,.25);margin-bottom:10px;cursor:default;}

/* ─── ALERT CARD ─── */
.ac{border-radius:var(--radius);padding:16px 18px;margin-bottom:10px;
  border:1px solid;display:flex;align-items:center;justify-content:space-between;}
.ac.red{background:var(--red-l);border-color:var(--red-bd);}
.ac.blue{background:var(--blue-l);border-color:var(--blue-bd);}
.ac-title{font-size:14px;font-weight:700;}
.ac-sub{font-size:11px;color:var(--muted);margin-top:2px;}
.ac-count{font-size:34px;font-weight:900;letter-spacing:-.02em;}

/* ─── MISC ─── */
.ok-banner{background:var(--green-l);border:1px solid var(--green-bd);border-radius:var(--radius);
  padding:10px 16px;font-size:13px;color:var(--green);margin-bottom:14px;}
.info-banner{background:var(--blue-l);border:1px solid var(--blue-bd);border-radius:var(--radius);
  padding:12px 16px;font-size:13px;color:var(--blue);margin-bottom:14px;}

section[data-testid="stSidebar"]{background:#fff!important;border-right:1px solid var(--border)!important;
  min-width:280px!important;max-width:280px!important;}
section[data-testid="stSidebar"] .block-container{padding:.8rem .8rem 2rem!important;}

.stDownloadButton>button{
  background:linear-gradient(135deg,#0f1729,#1e293b)!important;color:#fff!important;
  border:none!important;border-radius:var(--radius)!important;font-weight:700!important;
  font-size:14px!important;padding:12px!important;width:100%!important;
  letter-spacing:.01em!important;box-shadow:0 4px 14px rgba(15,23,42,.3)!important;}
div[data-baseweb="tag"]{background:var(--accent-l)!important;color:var(--accent)!important;}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def fix_encoding(df: pd.DataFrame) -> pd.DataFrame:
    try:
        if any('Ã' in str(c) for c in df.columns):
            df.columns = [
                c.encode('latin1').decode('utf-8', errors='replace') for c in df.columns
            ]
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
                df = pd.read_csv(
                    buf, sep=sep, encoding=enc,
                    low_memory=False, on_bad_lines='skip'
                )
                if df.shape[1] >= 3:
                    return fix_encoding(df)
            except Exception:
                continue
    buf.seek(0)
    df = pd.read_csv(
        buf, sep=None, engine='python',
        encoding='latin1', on_bad_lines='skip'
    )
    return fix_encoding(df)


@st.cache_data(show_spinner=False)
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace('\ufeff', '', regex=False)
        .str.replace('\xa0', ' ', regex=False)
        .str.upper()
    )
    return df


def load_t1(file_bytes: bytes, filename: str):
    buf = io.BytesIO(file_bytes)
    if filename.lower().endswith(('.xlsx', '.xls')):
        df_peek = pd.read_excel(buf, header=None, nrows=1)
    else:
        buf.seek(0)
        try:
            df_peek = pd.read_csv(buf, header=None, nrows=1, sep=None,
                                  engine='python', encoding='latin1')
        except Exception:
            df_peek = None

    no_header = False
    if df_peek is not None:
        first_val = str(df_peek.iloc[0, 0]).strip().replace('.0', '')
        no_header = first_val.isdigit()

    buf.seek(0)
    if filename.lower().endswith(('.xlsx', '.xls')):
        if no_header:
            df = pd.read_excel(buf, header=None)
        else:
            df = pd.read_excel(buf)
            df = normalize_columns(df)
    else:
        if no_header:
            buf.seek(0)
            try:
                df = pd.read_csv(buf, header=None, sep=None,
                                 engine='python', encoding='latin1',
                                 on_bad_lines='skip')
            except Exception:
                buf.seek(0)
                df = read_csv_smart(file_bytes, filename)
                df = normalize_columns(df)
        else:
            df = read_csv_smart(file_bytes, filename)
            df = normalize_columns(df)

    if no_header or 'ARTICLE' not in df.columns:
        if 'ARTICLE' not in df.columns:
            cols = ['ARTICLE'] + [f'_COL{i}' for i in range(1, len(df.columns))]
            df.columns = cols

    if 'ARTICLE' not in df.columns:
        found = ', '.join(df.columns.astype(str).tolist()[:10])
        return None, f"Colonne 'ARTICLE' introuvable. Colonnes détectées : {found}"

    df['SKU'] = (
        df['ARTICLE'].astype(str).str.strip()
        .str.zfill(8).str.slice(0, 8)
    )
    df = (
        df[df['SKU'].str.match(r'^\d{8}$', na=False)]
        .drop_duplicates(subset='SKU')
        .copy()
    )

    optional_cols = [
        ('LIBELLÉ ARTICLE', ''),
        ("FOURNISSEUR D'ORIGINE", ''),
        ('LIBELLÉ FOURNISSEUR ORIGINE', ''),
        ('MODE APPRO', ''),
        ('DATE CDE', ''),
        ('DATE LIV.', ''),
        ('SEMAINE RECEPTION', ''),
    ]
    for col, default in optional_cols:
        if col not in df.columns:
            df[col] = default

    df['SEMAINE RECEPTION'] = (
        df['SEMAINE RECEPTION'].astype(str).str.strip().replace('nan', '')
    )
    df['SEM_NUM'] = df['SEMAINE RECEPTION'].apply(
        lambda s: int(str(s).strip('Ss')) if str(s).strip('Ss').isdigit() else 99
    )
    df['ORIGINE'] = df['MODE APPRO'].apply(
        lambda m: 'IM' if 'IMPORT' in str(m).upper() else 'LO'
    )
    return df, None


@st.cache_data(show_spinner=False)
def load_stock(file_bytes: bytes, filename: str, sku_tuple: tuple):
    buf = io.BytesIO(file_bytes)
    if filename.lower().endswith(('.xlsx', '.xls')):
        df = pd.read_excel(buf)
    else:
        df = read_csv_smart(file_bytes, filename)
    df = fix_encoding(df)
    df = normalize_columns(df)

    required = {'LIBELLÉ SITE', 'CODE ARTICLE', 'NOUVEAU STOCK', 'RAL'}
    missing = required - set(df.columns)
    if missing:
        found = ', '.join(df.columns.tolist()[:10])
        return None, f"Colonnes manquantes : {missing}. Colonnes détectées : {found}"

    optional_stock = ('FOUR.', 'NOM FOURN.', 'LIBELLÉ ARTICLE', 'CODE ETAT', 'CODE MARKETING')
    for col in optional_stock:
        if col not in df.columns:
            df[col] = ''

    df['SKU'] = (
        df['CODE ARTICLE'].astype(str).str.strip()
        .str.zfill(8).str.slice(0, 8)
    )
    df = df[df['SKU'].isin(sku_tuple)].copy()
    df['NOUVEAU STOCK'] = pd.to_numeric(df['NOUVEAU STOCK'], errors='coerce').fillna(0)
    df['RAL']           = pd.to_numeric(df['RAL'],           errors='coerce').fillna(0)

    for col in optional_stock:
        df[col] = df[col].astype(str).str.strip().replace('nan', '')

    df = df.rename(columns={
        'LIBELLÉ SITE':    'Libellé site',
        'CODE ARTICLE':    'Code article',
        'NOUVEAU STOCK':   'Nouveau stock',
        'RAL':             'Ral',
        'FOUR.':           'Four.',
        'NOM FOURN.':      'Nom fourn.',
        'LIBELLÉ ARTICLE': 'Libellé article',
        'CODE ETAT':       'Code etat',
        'CODE MARKETING':  'Code marketing',
    })

    return df.drop_duplicates(subset=['Libellé site', 'SKU']), None


def sem_sort(s) -> int:
    try:
        return int(str(s).strip('Ss'))
    except Exception:
        return 99


def taux_implantation(df: pd.DataFrame) -> int:
    """Pourcentage d'articles implantés (Terminée uniquement)."""
    if len(df) == 0:
        return 0
    done = df['Statut'].isin(['Implantation Terminée']).sum()
    return int(done / len(df) * 100)


def color_taux(t: int) -> str:
    return '059669' if t >= 80 else ('0284c7' if t >= 50 else 'dc2626')


# ══════════════════════════════════════════════════════════════════════════════
# TOPBAR
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="topbar">
  <div class="topbar-left">
    <div class="topbar-icon">📋</div>
    <div>
      <div class="topbar-title">Rapport Implantation</div>
      <div class="topbar-sub">Nouvelles Références · Suivi Stock &amp; Flux</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:12px;">
    <div class="topbar-date">{TODAY_STR}</div>
    <div class="topbar-pill">DIRECTION SUPPLY</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — CHARGEMENT
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📁 Chargement")
    st.divider()
    st.markdown("**T1 — Nouvelles Références**")
    t1_file = st.file_uploader(
        "T1", type=["csv", "xlsx"], key="t1", label_visibility="collapsed"
    )
    st.markdown("**Extractions Stock** *(multi-fichiers)*")
    stock_files = st.file_uploader(
        "Stock", type=["csv", "xlsx"],
        accept_multiple_files=True, key="stocks", label_visibility="collapsed"
    )

# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT T1
# ══════════════════════════════════════════════════════════════════════════════
if not t1_file:
    st.markdown(
        '<div class="info-banner" style="margin-top:16px">⬆️ '
        'Charge le fichier <strong>T1 Flux</strong> dans la sidebar pour démarrer.</div>',
        unsafe_allow_html=True
    )
    st.stop()

with st.spinner("Lecture T1…"):
    t1_raw, t1_err = load_t1(t1_file.read(), t1_file.name)

if t1_err:
    st.error(f"❌ T1 : {t1_err}")
    st.stop()

SKU_TUPLE    = tuple(sorted(t1_raw['SKU'].unique()))
total_sku    = len(SKU_TUPLE)
sku_im_total = int((t1_raw['ORIGINE'] == 'IM').sum())
sku_lo_total = int((t1_raw['ORIGINE'] == 'LO').sum())

T1_KEEP = [
    'LIBELLÉ ARTICLE', 'LIBELLÉ FOURNISSEUR ORIGINE',
    'MODE APPRO', 'SEMAINE RECEPTION', 'DATE LIV.', 'ORIGINE', 'SEM_NUM'
]
t1_idx = t1_raw.set_index('SKU')[T1_KEEP]

# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT STOCK
# ══════════════════════════════════════════════════════════════════════════════
if not stock_files:
    st.markdown(
        '<div class="info-banner">⬆️ Charge les extractions stock dans la sidebar.</div>',
        unsafe_allow_html=True
    )
    st.stop()

frames = []
with st.spinner(f"Lecture {len(stock_files)} fichier(s)…"):
    for uf in stock_files:
        raw = uf.read()
        df_tmp, err = load_stock(raw, uf.name, SKU_TUPLE)
        if err:
            st.error(f"❌ `{uf.name}` : {err}")
        else:
            frames.append(df_tmp)

if not frames:
    st.error("Aucun fichier stock valide.")
    st.stop()

df_stock = (
    pd.concat(frames, ignore_index=True)
    .drop_duplicates(subset=['Libellé site', 'SKU'])
)
magasins_list = sorted(df_stock['Libellé site'].dropna().unique())


# ══════════════════════════════════════════════════════════════════════════════
# FILTRES SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.divider()
    st.markdown("### 🔍 Filtres")
    mag_sel     = st.multiselect("Magasins", magasins_list, default=magasins_list)
    origine_sel = st.multiselect("Origine", ['IM', 'LO'], default=['IM', 'LO'])

    sem_dispo = sorted(
        [s for s in t1_raw['SEMAINE RECEPTION'].unique() if s and s not in ('nan', '')],
        key=sem_sort
    )
    sem_sel = st.multiselect("Semaine réception", sem_dispo, default=sem_dispo)

    mode_dispo = sorted(
        [m for m in t1_raw['MODE APPRO'].unique() if m and m not in ('nan', '')]
    )
    mode_sel = st.multiselect("Mode Appro", mode_dispo, default=mode_dispo)

if not mag_sel:
    st.warning("Sélectionne au moins un magasin.")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# SÉLECTEUR MAGASINS PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
mc1, mc2, mc3 = st.columns([6, 1, 1])
with mc1:
    mag_main = st.multiselect(
        "🏪 Magasins affichés", magasins_list,
        default=mag_sel, key="mag_main"
    )
with mc2:
    if st.button("✅ Tous",  use_container_width=True):
        mag_main = magasins_list
with mc3:
    if st.button("❌ Aucun", use_container_width=True):
        mag_main = []

mag_actifs = mag_main if mag_main else mag_sel
if not mag_actifs:
    st.warning("Sélectionne au moins un magasin.")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# CALCUL VECTORISÉ — UNE SEULE PASSE
# ══════════════════════════════════════════════════════════════════════════════
sku_mask = (
    t1_raw['ORIGINE'].isin(origine_sel)
    & (t1_raw['SEMAINE RECEPTION'].isin(sem_sel)  if sem_sel  else True)
    & (t1_raw['MODE APPRO'].isin(mode_sel)         if mode_sel else True)
)
sku_scope     = t1_raw.loc[sku_mask, 'SKU'].unique()
total_sku_sel = len(sku_scope)

if total_sku_sel == 0:
    st.warning("Aucun article ne correspond aux filtres.")
    st.stop()

base_df = pd.DataFrame(
    pd.MultiIndex.from_product(
        [mag_actifs, sku_scope], names=['Libellé site', 'SKU']
    ).tolist(),
    columns=['Libellé site', 'SKU']
)

stock_scope = df_stock[
    df_stock['Libellé site'].isin(mag_actifs) & df_stock['SKU'].isin(sku_scope)
]

merged = base_df.merge(
    stock_scope[[
        'Libellé site', 'SKU', 'Nouveau stock', 'Ral',
        'Code etat', 'Code marketing', 'Libellé article'
    ]],
    on=['Libellé site', 'SKU'], how='left'
)
merged['Nouveau stock'] = merged['Nouveau stock'].fillna(0)
merged['Ral']           = merged['Ral'].fillna(0)
merged['Code etat']     = merged['Code etat'].fillna('').astype(str)

merged = merged.merge(
    t1_idx.reset_index().rename(columns={
        'LIBELLÉ ARTICLE':             'T1_lib',
        'LIBELLÉ FOURNISSEUR ORIGINE': 'Fournisseur',
        'MODE APPRO':                  'Mode Appro',
        'SEMAINE RECEPTION':           'Sem. Réception',
        'DATE LIV.':                   'Date Livraison',
        'ORIGINE':                     'Origine',
        'SEM_NUM':                     'SEM_NUM',
    }),
    on='SKU', how='left'
)

merged['Libellé article'] = merged['Libellé article'].fillna('').astype(str)
merged['Libellé article'] = merged.apply(
    lambda r: r['Libellé article'] if r['Libellé article'] else r['T1_lib'], axis=1
)
merged.drop(columns='T1_lib', inplace=True)

# ── STATUTS : Partielle supprimée — stock > 0 (RAL indifférent) = Terminée ──
conds = [
    (merged['Nouveau stock'] > 0),                                        # stock présent → Terminée
    (merged['Nouveau stock'] == 0) & (merged['Ral'] > 0),                # 0 stock + RAL → Attente
]
choices = ["Implantation Terminée", "En Attente Livraison"]
merged['Statut']     = np.select(conds, choices, default="Alerte Aucun Mouvement")
merged['Etat Actif'] = merged['Code etat'] == '2'

detail_df = merged.rename(columns={
    'Libellé site': 'Magasin',
    'Nouveau stock': 'Stock',
    'Ral': 'RAL',
})


# ══════════════════════════════════════════════════════════════════════════════
# AGRÉGATS
# ══════════════════════════════════════════════════════════════════════════════
# Ordre et couleurs à 3 statuts (Partielle supprimée)
S_ORDER  = [
    "Implantation Terminée",
    "En Attente Livraison",
    "Alerte Aucun Mouvement"
]
S_COLORS = {
    "Implantation Terminée":  "#059669",
    "En Attente Livraison":   "#0284c7",
    "Alerte Aucun Mouvement": "#dc2626",
}

pivot = (
    detail_df.groupby(['Magasin', 'Statut']).size()
    .unstack(fill_value=0)
    .reindex(columns=S_ORDER, fill_value=0)
    .reset_index()
)
pivot.columns.name = None
pivot['Total']    = total_sku_sel
pivot['Taux (%)'] = (
    pivot['Implantation Terminée'] / total_sku_sel * 100
).round(0).astype(int)

total_cells = len(mag_actifs) * total_sku_sel
ct  = int(pivot['Implantation Terminée'].sum())
ca  = int(pivot['En Attente Livraison'].sum())
cal = int(pivot['Alerte Aucun Mouvement'].sum())
avg_impl = int(pivot['Taux (%)'].mean()) if not pivot.empty else 0
pct = lambda n: int(n / total_cells * 100) if total_cells > 0 else 0

df_im = detail_df[detail_df['Origine'] == 'IM']
df_lo = detail_df[detail_df['Origine'] == 'LO']

df_attente = detail_df[detail_df['Statut'] == 'En Attente Livraison']
df_alerte  = detail_df[detail_df['Statut'] == 'Alerte Aucun Mouvement']

attente_im = df_attente[df_attente['Origine'] == 'IM']
attente_lo = df_attente[df_attente['Origine'] == 'LO']
alerte_im  = df_alerte[df_alerte['Origine']  == 'IM']
alerte_lo  = df_alerte[df_alerte['Origine']  == 'LO']

im_alerte     = int((df_im['Statut'] == 'Alerte Aucun Mouvement').sum())
lo_alerte     = int((df_lo['Statut'] == 'Alerte Aucun Mouvement').sum())
total_actions = ca + cal

tim = taux_implantation(df_im)
tlo = taux_implantation(df_lo)


# ══════════════════════════════════════════════════════════════════════════════
# BANNIÈRE ACTIONS CRITIQUES
# ══════════════════════════════════════════════════════════════════════════════
if total_actions > 0:
    st.markdown(f"""
    <div class="alert-banner">
      <div class="ab-badge">⚡ ACTIONS REQUISES</div>
      <div class="ab-item">
        <div class="ab-num" style="color:#dc2626">{cal}</div>
        <div class="ab-lbl">Aucun Mouvement</div>
      </div>
      <div class="ab-item">
        <div class="ab-num" style="color:#dc2626">{len(alerte_im)}</div>
        <div class="ab-lbl">dont IM</div>
      </div>
      <div class="ab-item">
        <div class="ab-num" style="color:#ea580c">{len(alerte_lo)}</div>
        <div class="ab-lbl">dont LO</div>
      </div>
      <div class="ab-item">
        <div class="ab-num" style="color:#0284c7">{ca}</div>
        <div class="ab-lbl">Attente Livraison</div>
      </div>
      <div class="ab-item">
        <div class="ab-num" style="color:#0284c7">{len(attente_im)}</div>
        <div class="ab-lbl">dont IM</div>
      </div>
      <div class="ab-item">
        <div class="ab-num" style="color:#0284c7">{len(attente_lo)}</div>
        <div class="ab-lbl">dont LO</div>
      </div>
      <div style="margin-left:auto;display:flex;flex-direction:column;align-items:flex-end;">
        <div style="font-size:36px;font-weight:900;color:#dc2626;line-height:1">{total_actions}</div>
        <div style="font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em">articles à traiter</div>
      </div>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# KPI STRIP IM / LO
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="strip">
  <div class="strip-card">
    <div class="strip-tag tag-im">IMPORT</div>
    <div class="strip-label">Références</div>
    <div class="strip-val" style="color:#2563eb">{sku_im_total}</div>
    <div class="strip-sub">SKU à implanter</div>
  </div>
  <div class="strip-card">
    <div class="strip-tag tag-im">IMPORT</div>
    <div class="strip-label">Taux implanté</div>
    <div class="strip-val" style="color:#{color_taux(tim)}">{tim}%%</div>
    <div class="strip-sub">stock présent en magasin</div>
  </div>
  <div class="strip-card">
    <div class="strip-tag tag-im">IMPORT</div>
    <div class="strip-label">Alerte Aucun Mvt</div>
    <div class="strip-val" style="color:#dc2626">{im_alerte}</div>
    <div class="strip-sub">escalade fournisseur</div>
  </div>
  <div class="strip-card">
    <div class="strip-tag tag-lo">LOCAL</div>
    <div class="strip-label">Références</div>
    <div class="strip-val" style="color:#059669">{sku_lo_total}</div>
    <div class="strip-sub">SKU à implanter</div>
  </div>
  <div class="strip-card">
    <div class="strip-tag tag-lo">LOCAL</div>
    <div class="strip-label">Taux implanté</div>
    <div class="strip-val" style="color:#{color_taux(tlo)}">{tlo}%%</div>
    <div class="strip-sub">stock présent en magasin</div>
  </div>
  <div class="strip-card">
    <div class="strip-tag tag-lo">LOCAL</div>
    <div class="strip-label">Alerte Aucun Mvt</div>
    <div class="strip-val" style="color:#dc2626">{lo_alerte}</div>
    <div class="strip-sub">relance supply locale</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# KPI GLOBAUX — 3 cartes (Partielle supprimée)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="kpi-row">
  <div class="kpi g">
    <div class="kpi-lbl">✅ Implantation Terminée</div>
    <div class="kpi-val">{ct}</div>
    <div class="kpi-pct">{pct(ct)}%% · Stock présent en magasin</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(ct)}%%;background:#059669"></div></div>
  </div>
  <div class="kpi b">
    <div class="kpi-lbl">🚚 En Attente Livraison</div>
    <div class="kpi-val">{ca}</div>
    <div class="kpi-pct">{pct(ca)}%% · RAL présent · À relancer</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(ca)}%%;background:#0284c7"></div></div>
  </div>
  <div class="kpi r">
    <div class="kpi-lbl">🚨 Alerte Aucun Mouvement</div>
    <div class="kpi-val">{cal}</div>
    <div class="kpi-pct">{pct(cal)}%% · Stock = 0 · RAL = 0</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(cal)}%%;background:#dc2626"></div></div>
  </div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SCORECARD RAG PAR MAGASIN
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sh">SCORECARD MAGASINS</div>', unsafe_allow_html=True)

rag_html = '<div class="rag-grid">'
for _, row in pivot.sort_values('Taux (%)', ascending=False).iterrows():
    t_ = row['Taux (%)']
    cls   = 'g' if t_ >= 80 else 'r'
    c_hex = '#059669' if t_ >= 80 else '#dc2626'
    rag_html += f"""
    <div class="rag-card {cls}">
      <div class="rag-dot" style="background:{c_hex}"></div>
      <div class="rag-name">{row['Magasin']}</div>
      <div class="rag-pct" style="color:{c_hex}">{t_}%</div>
      <div class="rag-detail">
        {int(row['Implantation Terminée'])}✅
        {int(row['Alerte Aucun Mouvement'])}🚨
      </div>
    </div>"""
rag_html += '</div>'
st.markdown(rag_html, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# NAVIGATION ONGLETS
# ══════════════════════════════════════════════════════════════════════════════
TABS = ["📊 Vue Globale", "🚨 Alertes & Actions", "🗓️ Calendrier Flux", "📋 Plan d'Action"]

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

PLOTLY_BASE = dict(
    paper_bgcolor="#fff", plot_bgcolor="#fff",
    font=dict(family="Inter", color="#64748b", size=12),
    margin=dict(l=20, r=20, t=44, b=20)
)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — VUE GLOBALE
# ══════════════════════════════════════════════════════════════════════════════
if active == TABS[0]:
    c1, c2 = st.columns([3, 2])

    with c1:
        mel = pivot.melt(
            id_vars='Magasin', value_vars=list(S_COLORS.keys()),
            var_name='Statut', value_name='N'
        )
        fig = px.bar(
            mel, x='Magasin', y='N', color='Statut',
            color_discrete_map=S_COLORS, barmode='stack',
            title='Situation par magasin'
        )
        fig.update_traces(
            textposition='inside', texttemplate='%{y}',
            textfont_size=11, textfont_color='white'
        )
        fig.update_layout(
            **PLOTLY_BASE, height=400,
            legend=dict(orientation='h', y=-0.22, bgcolor='rgba(0,0,0,0)', font_size=11),
            xaxis=dict(gridcolor='#f0f2f8'),
            yaxis=dict(gridcolor='#f0f2f8')
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig_d = go.Figure(go.Pie(
            labels=list(S_COLORS.keys()),
            values=[ct, ca, cal],
            hole=0.65,
            marker=dict(colors=list(S_COLORS.values()), line=dict(color='#fff', width=3)),
            textfont=dict(size=11)
        ))
        fig_d.add_annotation(
            text=f"<b>{avg_impl}%</b><br>implanté",
            x=0.5, y=0.5,
            font=dict(size=18, color='#0f1729', family='Inter'),
            showarrow=False
        )
        fig_d.update_layout(
            **PLOTLY_BASE, height=400, title='Répartition globale',
            legend=dict(orientation='v', x=1.01, bgcolor='rgba(0,0,0,0)', font_size=11)
        )
        st.plotly_chart(fig_d, use_container_width=True)

    st.markdown('<div class="sh">DÉTAIL PAR MAGASIN</div>', unsafe_allow_html=True)
    disp_cols = [
        "Magasin", "Implantation Terminée",
        "En Attente Livraison", "Alerte Aucun Mouvement", "Total", "Taux (%)"
    ]
    st.dataframe(
        pivot[disp_cols].style
        .background_gradient(subset=['Implantation Terminée'],  cmap='Greens')
        .background_gradient(subset=['Alerte Aucun Mouvement'], cmap='Reds')
        .background_gradient(subset=['En Attente Livraison'],   cmap='Blues')
        .background_gradient(subset=['Taux (%)'], cmap='RdYlGn', vmin=0, vmax=100)
        .format({'Taux (%)': '{}%'}),
        use_container_width=True, hide_index=True,
        height=min(600, 60 + len(mag_actifs) * 42)
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — ALERTES & ACTIONS
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[1]:
    filt = st.radio(
        "", ["Toutes les alertes", "🚨 Aucun Mouvement", "🚚 Attente Livraison"],
        horizontal=True, key="afilt", label_visibility="collapsed"
    )

    ACOLS = [
        "Magasin", "SKU", "Libellé article", "Origine", "Mode Appro",
        "Sem. Réception", "Date Livraison", "Code etat", "Stock", "RAL", "Statut"
    ]

    ALERT_SECTIONS = {
        "🚨 Aucun Mouvement": (
            df_alerte, "#dc2626",
            "Aucun Mouvement — Stock = 0 · RAL = 0",
            "Escalader fournisseur · Vérifier commande · Informer magasin",
            "red"
        ),
        "🚚 Attente Livraison": (
            df_attente, "#0284c7",
            "En Attente Livraison — RAL présent · Stock = 0",
            "Confirmer date livraison · Préparer réception magasin",
            "blue"
        ),
    }

    for key, (df_a, hex_color, title_txt, action_txt, css_cls) in ALERT_SECTIONS.items():
        if filt not in ("Toutes les alertes", key):
            continue

        st.markdown(f"""
        <div class="ac {css_cls}">
          <div>
            <div class="ac-title" style="color:{hex_color}">{title_txt}</div>
            <div class="ac-sub">💡 Action : {action_txt}</div>
          </div>
          <div class="ac-count" style="color:{hex_color}">{len(df_a)}</div>
        </div>""", unsafe_allow_html=True)

        if df_a.empty:
            st.success("✅ Aucune alerte dans cette catégorie")
            continue

        tab_im, tab_lo = st.tabs([
            f"IMPORT — {(df_a['Origine'] == 'IM').sum()} SKU",
            f"LOCAL  — {(df_a['Origine'] == 'LO').sum()} SKU",
        ])

        for tab, orig, sub_df in [
            (tab_im, 'IM', df_a[df_a['Origine'] == 'IM']),
            (tab_lo, 'LO', df_a[df_a['Origine'] == 'LO']),
        ]:
            with tab:
                if sub_df.empty:
                    st.success(f"✅ Aucune alerte {orig}")
                    continue

                col_g, col_t = st.columns([2, 3])

                with col_g:
                    top = (
                        sub_df.groupby(['SKU', 'Libellé article'])['Magasin'].count()
                        .reset_index().rename(columns={'Magasin': 'Nb Magasins'})
                        .sort_values('Nb Magasins').tail(10)
                    )
                    top['lbl'] = top['SKU'] + ' – ' + top['Libellé article'].str[:28]
                    fig_t = go.Figure(go.Bar(
                        x=top['Nb Magasins'], y=top['lbl'], orientation='h',
                        marker=dict(color=hex_color, cornerradius=4),
                        text=top['Nb Magasins'], textposition='outside',
                        textfont=dict(color='#64748b', size=11)
                    ))
                    fig_t.update_layout(
                        **PLOTLY_BASE, height=max(200, len(top) * 34),
                        title=f'Top SKU impactés — {orig}',
                        xaxis=dict(gridcolor='#f0f2f8'),
                        yaxis=dict(tickfont_size=10)
                    )
                    st.plotly_chart(fig_t, use_container_width=True)

                with col_t:
                    top_m = (
                        sub_df.groupby('Magasin')['SKU'].count()
                        .reset_index().rename(columns={'SKU': 'Nb SKU'})
                        .sort_values('Nb SKU', ascending=False)
                    )
                    fig_m = go.Figure(go.Bar(
                        x=top_m['Magasin'], y=top_m['Nb SKU'],
                        marker=dict(color=hex_color, cornerradius=4),
                        text=top_m['Nb SKU'], textposition='outside',
                        textfont=dict(color='#64748b', size=11)
                    ))
                    fig_m.update_layout(
                        **PLOTLY_BASE, height=max(200, len(top_m) * 40),
                        title=f'Alertes par magasin — {orig}',
                        xaxis=dict(gridcolor='#f0f2f8'),
                        yaxis=dict(gridcolor='#f0f2f8')
                    )
                    st.plotly_chart(fig_m, use_container_width=True)

                with st.expander(
                    f"📋 Liste complète {orig} — {len(sub_df)} lignes", expanded=False
                ):
                    st.dataframe(
                        sub_df[ACOLS].sort_values(['Magasin', 'Sem. Réception'])
                        .reset_index(drop=True),
                        use_container_width=True, hide_index=True
                    )

        st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — CALENDRIER FLUX
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[2]:
    cal_df = detail_df[
        detail_df['Sem. Réception'].str.match(r'^S\d+$', na=False)
    ].copy()
    cal_df['SEM_NUM'] = cal_df['Sem. Réception'].apply(sem_sort)

    if cal_df.empty:
        st.info("Aucune donnée de semaine disponible.")
    else:
        sem_order = sorted(cal_df['Sem. Réception'].unique(), key=sem_sort)

        c1, c2 = st.columns(2)
        with c1:
            ss = (
                cal_df.groupby(['Sem. Réception', 'SEM_NUM', 'Statut']).size()
                .reset_index(name='N').sort_values('SEM_NUM')
            )
            fig_s = px.bar(
                ss, x='Sem. Réception', y='N', color='Statut',
                color_discrete_map=S_COLORS, barmode='stack',
                title='Articles par semaine & statut',
                category_orders={'Sem. Réception': sem_order}
            )
            fig_s.update_traces(
                textposition='inside', texttemplate='%{y}',
                textfont_size=10, textfont_color='white'
            )
            fig_s.update_layout(
                **PLOTLY_BASE, height=360,
                xaxis=dict(gridcolor='#f0f2f8'),
                yaxis=dict(gridcolor='#f0f2f8'),
                legend=dict(orientation='h', y=-0.22, bgcolor='rgba(0,0,0,0)')
            )
            st.plotly_chart(fig_s, use_container_width=True)

        with c2:
            os_df = (
                cal_df.groupby(['Origine', 'Sem. Réception', 'SEM_NUM']).size()
                .reset_index(name='N').sort_values('SEM_NUM')
            )
            fig_o = px.bar(
                os_df, x='Sem. Réception', y='N', color='Origine', barmode='group',
                color_discrete_map={'IM': '#2563eb', 'LO': '#059669'},
                title='IM vs LO par semaine',
                category_orders={'Sem. Réception': sem_order}
            )
            fig_o.update_traces(
                textposition='outside', texttemplate='%{y}', textfont_size=10
            )
            fig_o.update_layout(
                **PLOTLY_BASE, height=360,
                xaxis=dict(gridcolor='#f0f2f8'),
                yaxis=dict(gridcolor='#f0f2f8'),
                legend=dict(orientation='h', y=-0.22, bgcolor='rgba(0,0,0,0)')
            )
            st.plotly_chart(fig_o, use_container_width=True)

        st.markdown('<div class="sh">DÉTAIL PAR SEMAINE</div>', unsafe_allow_html=True)
        tbl = (
            cal_df.groupby(['Sem. Réception', 'SEM_NUM', 'Origine']).agg(
                Articles=('SKU', 'nunique'),
                Terminé=('Statut',  lambda x: (x == 'Implantation Terminée').sum()),
                Attente=('Statut',  lambda x: (x == 'En Attente Livraison').sum()),
                Alerte=('Statut',   lambda x: (x == 'Alerte Aucun Mouvement').sum()),
            )
            .reset_index().sort_values('SEM_NUM').drop(columns='SEM_NUM')
        )
        st.dataframe(
            tbl.style
            .background_gradient(subset=['Terminé'], cmap='Greens')
            .background_gradient(subset=['Alerte'],  cmap='Reds')
            .background_gradient(subset=['Attente'], cmap='Blues'),
            use_container_width=True, hide_index=True
        )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — PLAN D'ACTION
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[3]:
    c1, c2 = st.columns([1, 2])

    with c1:
        recap_s    = pivot.sort_values('Taux (%)', ascending=True)
        bar_colors = [
            '#059669' if v >= 80 else '#dc2626'
            for v in recap_s['Taux (%)']
        ]
        fig_h = go.Figure(go.Bar(
            x=recap_s['Taux (%)'], y=recap_s['Magasin'], orientation='h',
            marker=dict(color=bar_colors, cornerradius=5),
            text=[f"{v}%" for v in recap_s['Taux (%)']],
            textposition='outside',
            textfont=dict(color='#0f1729', size=13, family='Inter')
        ))
        fig_h.add_vline(
            x=80, line_dash='dash', line_color='#e2e8f4',
            annotation_text='Cible 80%', annotation_font_color='#94a3b8'
        )
        fig_h.update_layout(
            **PLOTLY_BASE, height=max(280, len(mag_actifs) * 48),
            xaxis=dict(range=[0, 118], gridcolor='#f0f2f8', ticksuffix='%'),
            yaxis=dict(gridcolor='rgba(0,0,0,0)'),
            title='Taux par magasin'
        )
        st.plotly_chart(fig_h, use_container_width=True)

    with c2:
        mag_pa = st.selectbox("Sélectionner un magasin", mag_actifs, key="pa_mag")
        df_pa  = detail_df[
            (detail_df['Magasin'] == mag_pa) &
            (detail_df['Statut'].isin(['Alerte Aucun Mouvement', 'En Attente Livraison']))
        ]
        krow    = pivot[pivot['Magasin'] == mag_pa]
        t_mag   = int(krow['Taux (%)'].values[0])  if not krow.empty else 0
        n_alert = int(krow['Alerte Aucun Mouvement'].values[0]) if not krow.empty else 0
        n_att   = int(krow['En Attente Livraison'].values[0])   if not krow.empty else 0

        c_hex = "#059669" if t_mag >= 80 else "#dc2626"
        bg    = "#ecfdf5" if t_mag >= 80 else "#fef2f2"
        bd    = "#6ee7b7" if t_mag >= 80 else "#fecaca"

        st.markdown(f"""
        <div style="background:{bg};border:1px solid {bd};border-radius:10px;
             padding:16px 20px;margin-bottom:14px;display:flex;align-items:center;gap:20px;">
          <div style="font-size:52px;font-weight:900;color:{c_hex};line-height:1">{t_mag}%%</div>
          <div>
            <div style="font-size:15px;font-weight:700;color:#0f1729">{mag_pa}</div>
            <div style="font-size:12px;color:#64748b;margin-top:3px">
              {n_alert} alertes aucun mvt · {n_att} en attente livraison
            </div>
          </div>
        </div>""", unsafe_allow_html=True)

        if df_pa.empty:
            st.success(f"✅ {mag_pa} — aucune action requise.")
        else:
            PA_COLS = [
                "SKU", "Libellé article", "Origine", "Mode Appro",
                "Sem. Réception", "Date Livraison", "Code etat",
                "Stock", "RAL", "Statut"
            ]
            st.dataframe(
                df_pa[PA_COLS].sort_values(['Statut', 'Origine', 'Sem. Réception'])
                .reset_index(drop=True),
                use_container_width=True, hide_index=True
            )


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sh">EXPORT</div>', unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def build_report(
    det_b: bytes, piv_b: bytes,
    today_str: str, today_file: str,
    mag_count: int, sku_count: int,
    ct: int, ca: int, cal_n: int,
    avg: int, skt: int, skl: int,
    tim: int, tlo: int
) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    det = pd.read_parquet(io.BytesIO(det_b))
    piv = pd.read_parquet(io.BytesIO(piv_b))

    wb = Workbook()
    wb.remove(wb.active)

    C = dict(
        dark="0F1729", navy="1E293B", blue="2563EB", blue_l="EFF4FF",
        green="059669", green_l="ECFDF5", green_bd="6EE7B7",
        blue2="0284C7", blue2_l="F0F9FF", blue2_bd="BAE6FD",
        red="DC2626", red_l="FEF2F2", red_bd="FECACA",
        grey="F0F2F8", border="E2E8F4", muted="64748B", white="FFFFFF"
    )

    def F(c=C['dark'], sz=10, b=False):
        return Font(name='Arial', size=sz, bold=b, color=c)

    def fill(h):
        return PatternFill("solid", fgColor=h)

    def ctr(w=False):
        return Alignment(horizontal='center', vertical='center', wrap_text=w)

    def lft(w=False):
        return Alignment(horizontal='left', vertical='center', wrap_text=w)

    def brd():
        t = Side(style='thin', color=C['border'])
        return Border(left=t, right=t, top=t, bottom=t)

    def write_header(ws, title, sub=""):
        ws.sheet_view.showGridLines = False
        ws.merge_cells('B1:L1')
        ws.row_dimensions[1].height = 44
        c = ws['B1']
        c.value = title
        c.font  = Font(name='Arial', size=20, bold=True, color=C['white'])
        c.fill  = fill(C['dark'])
        c.alignment = lft()
        if sub:
            ws.merge_cells('B2:L2')
            ws.row_dimensions[2].height = 20
            c2 = ws['B2']
            c2.value     = sub
            c2.font      = F(C['muted'], 9)
            c2.fill      = fill(C['grey'])
            c2.alignment = lft()

    def write_table(ws, df, r0, cols, hcol=C['dark']):
        ws.row_dimensions[r0].height = 22
        for ci, (h, k, w, al, cfn) in enumerate(cols, start=2):
            cell = ws.cell(r0, ci)
            cell.value     = h
            cell.font      = Font(name='Arial', size=9, bold=True, color=C['white'])
            cell.fill      = fill(hcol)
            cell.alignment = ctr(True)
            cell.border    = brd()
            ws.column_dimensions[get_column_letter(ci)].width = w

        for ri, (_, row) in enumerate(df.iterrows()):
            rr = r0 + 1 + ri
            ws.row_dimensions[rr].height = 18
            for ci, (_, k, _, al, cfn) in enumerate(cols, start=2):
                cell = ws.cell(rr, ci)
                val  = row[k] if k in row.index else ''
                cell.value = bool(val) if isinstance(val, np.bool_) else val
                if cfn:
                    bg, fc = cfn(val, row)
                else:
                    bg = C['white'] if ri % 2 == 0 else C['grey']
                    fc = C['dark']
                cell.font      = F(fc, 9, b=(ci == 2))
                cell.fill      = fill(bg)
                cell.border    = brd()
                cell.alignment = ctr() if al == 'c' else lft()

    # Palettes statuts (3 statuts)
    ss = {
        'Implantation Terminée':  (C['green_l'],  C['green']),
        'En Attente Livraison':   (C['blue2_l'],  C['blue2']),
        'Alerte Aucun Mouvement': (C['red_l'],    C['red']),
    }
    os_ = {
        'IM': (C['blue_l'],  C['blue']),
        'LO': (C['green_l'], C['green']),
    }

    def ts(v, r):
        bg = C['green_l']  if v >= 80 else C['red_l']
        fc = C['green']    if v >= 80 else C['red']
        return bg, fc

    # ── SHEET 1 : RÉSUMÉ EXÉCUTIF ─────────────────────────────────────────────
    ws1 = wb.create_sheet("📊 Résumé Exécutif")
    write_header(
        ws1, f"RAPPORT IMPLANTATION — {today_str}",
        f"{mag_count} magasin(s)  ·  {sku_count} SKU  ·  Taux moyen réseau : {avg}%"
    )
    ws1.column_dimensions['A'].width = 2

    # 3 KPI cards (Partielle supprimée)
    kpis = [
        ("TERMINÉ",          ct,    C['green'],  C['green_l']),
        ("ATTENTE LIV.",     ca,    C['blue2'],  C['blue2_l']),
        ("ALERTE AUCUN MVT", cal_n, C['red'],    C['red_l']),
    ]
    for i, (lbl, val, col, bg) in enumerate(kpis):
        ci    = 2 + i
        col_l = get_column_letter(ci)
        ws1.column_dimensions[col_l].width = 20
        denom = mag_count * sku_count
        share = int(val / denom * 100) if denom > 0 else 0
        for r_, v_, fs_, bg_, fc_ in [
            (4, lbl,                  8,  C['grey'], col),
            (5, val,                  32, bg,        col),
            (6, f"{share}% du total", 9,  bg,        col),
        ]:
            ws1.row_dimensions[r_].height = 14 if r_ == 4 else (40 if r_ == 5 else 18)
            cell = ws1.cell(r_, ci)
            cell.value     = v_
            cell.font      = Font(name='Arial', size=fs_, bold=True, color=fc_)
            cell.fill      = fill(bg_)
            cell.alignment = ctr()
            cell.border    = brd()

    ws1.row_dimensions[7].height = 10
    for i, (lbl, desc, col, bg) in enumerate([
        ("IMPORT", f"Taux implanté : {tim}%  ·  {skt} SKU", C['blue'],  C['blue_l']),
        ("LOCAL",  f"Taux implanté : {tlo}%  ·  {skl} SKU", C['green'], C['green_l']),
    ]):
        ci = 2 + i
        ws1.row_dimensions[8].height = 22
        c8 = ws1.cell(8, ci)
        c8.value     = f"{lbl}  ·  {desc}"
        c8.font      = Font(name='Arial', size=10, bold=True, color=col)
        c8.fill      = fill(bg)
        c8.alignment = ctr()
        c8.border    = brd()

    # Situation IM
    ws1.row_dimensions[10].height = 10
    ws1.row_dimensions[11].height = 22
    ws1.merge_cells('B11:F11')
    h_im = ws1['B11']
    h_im.value     = "SITUATION PAR SITE — IMPORT"
    h_im.font      = Font(name='Arial', size=11, bold=True, color=C['white'])
    h_im.fill      = fill(C['blue'])
    h_im.alignment = lft()

    piv_im = (
        det[det['Origine'] == 'IM']
        .groupby('Magasin')['Statut'].value_counts().unstack(fill_value=0)
    )
    for s in S_ORDER:
        if s not in piv_im.columns: piv_im[s] = 0
    piv_im = piv_im.reset_index()
    tot_im = det[det['Origine'] == 'IM']['SKU'].nunique()
    if tot_im > 0:
        piv_im['Taux IM (%)'] = (
            piv_im.get('Implantation Terminée', 0) / tot_im * 100
        ).round(0).astype(int)
    else:
        piv_im['Taux IM (%)'] = 0

    im_site_cols = [
        ("MAGASIN",      "Magasin",               22, 'l', None),
        ("TERMINÉ",      "Implantation Terminée",  12, 'c', None),
        ("ATTENTE LIV.", "En Attente Livraison",   12, 'c', None),
        ("ALERTE",       "Alerte Aucun Mouvement", 12, 'c', None),
        ("TAUX IM",      "Taux IM (%)",            12, 'c', ts),
    ]
    write_table(ws1, piv_im, 12, im_site_cols, hcol=C['blue'])

    # Situation LO
    r_lo = 12 + len(piv_im) + 2
    ws1.row_dimensions[r_lo].height = 10
    r_lo += 1
    ws1.row_dimensions[r_lo].height = 22
    ws1.merge_cells(f'B{r_lo}:F{r_lo}')
    h_lo = ws1[f'B{r_lo}']
    h_lo.value     = "SITUATION PAR SITE — LOCAL"
    h_lo.font      = Font(name='Arial', size=11, bold=True, color=C['white'])
    h_lo.fill      = fill(C['green'])
    h_lo.alignment = lft()

    piv_lo = (
        det[det['Origine'] == 'LO']
        .groupby('Magasin')['Statut'].value_counts().unstack(fill_value=0)
    )
    for s in S_ORDER:
        if s not in piv_lo.columns: piv_lo[s] = 0
    piv_lo = piv_lo.reset_index()
    tot_lo = det[det['Origine'] == 'LO']['SKU'].nunique()
    if tot_lo > 0:
        piv_lo['Taux LO (%)'] = (
            piv_lo.get('Implantation Terminée', 0) / tot_lo * 100
        ).round(0).astype(int)
    else:
        piv_lo['Taux LO (%)'] = 0

    lo_site_cols = [
        ("MAGASIN",      "Magasin",               22, 'l', None),
        ("TERMINÉ",      "Implantation Terminée",  12, 'c', None),
        ("ATTENTE LIV.", "En Attente Livraison",   12, 'c', None),
        ("ALERTE",       "Alerte Aucun Mouvement", 12, 'c', None),
        ("TAUX LO",      "Taux LO (%)",            12, 'c', ts),
    ]
    write_table(ws1, piv_lo, r_lo + 1, lo_site_cols, hcol=C['green'])

    # ── SHEET 2 : ALERTES ─────────────────────────────────────────────────────
    ws2 = wb.create_sheet("🚨 Alertes & Actions")
    write_header(
        ws2, "ALERTES & ACTIONS",
        f"{today_str}  ·  {cal_n} aucun mouvement  ·  {ca} en attente livraison"
    )
    ws2.column_dimensions['A'].width = 2
    cur = 4

    a_cols = [
        ("MAGASIN",    "Magasin",         22, 'l', None),
        ("SKU",        "SKU",             12, 'c', None),
        ("LIBELLÉ",    "Libellé article", 36, 'l', None),
        ("ORIGINE",    "Origine",         10, 'c', lambda v, r: os_.get(v, (C['grey'], C['dark']))),
        ("MODE APPRO", "Mode Appro",      16, 'l', None),
        ("SEM.",       "Sem. Réception",  10, 'c', None),
        ("DATE LIV.",  "Date Livraison",  14, 'c', None),
        ("STOCK",      "Stock",           10, 'c', None),
        ("RAL",        "RAL",             10, 'c', None),
        ("STATUT",     "Statut",          24, 'c', lambda v, r: ss.get(v, (C['grey'], C['dark']))),
    ]

    alert_sections = [
        (
            "🚨 ALERTE AUCUN MOUVEMENT — IMPORT",
            det[(det['Origine'] == 'IM') & (det['Statut'] == 'Alerte Aucun Mouvement')],
            C['red'], "Action : Escalader fournisseur IM · Vérifier bon de commande"
        ),
        (
            "🚨 ALERTE AUCUN MOUVEMENT — LOCAL",
            det[(det['Origine'] == 'LO') & (det['Statut'] == 'Alerte Aucun Mouvement')],
            C['red'], "Action : Relancer supply locale · Contacter fournisseur LO"
        ),
        (
            "🚚 EN ATTENTE LIVRAISON — IMPORT",
            det[(det['Origine'] == 'IM') & (det['Statut'] == 'En Attente Livraison')],
            C['blue2'], "Action : Confirmer ETA · Préparer réception import"
        ),
        (
            "🚚 EN ATTENTE LIVRAISON — LOCAL",
            det[(det['Origine'] == 'LO') & (det['Statut'] == 'En Attente Livraison')],
            C['green'], "Action : Confirmer livraison · Informer magasin"
        ),
    ]

    for title, df_a, col, action in alert_sections:
        ws2.row_dimensions[cur].height = 10
        cur += 1
        ws2.merge_cells(f'B{cur}:K{cur}')
        ws2.row_dimensions[cur].height = 24
        tc = ws2[f'B{cur}']
        tc.value     = f"{title}  —  {len(df_a)} article(s)"
        tc.font      = Font(name='Arial', size=11, bold=True, color=C['white'])
        tc.fill      = fill(col)
        tc.alignment = lft()
        cur += 1

        ws2.merge_cells(f'B{cur}:K{cur}')
        ws2.row_dimensions[cur].height = 18
        ac_cell = ws2[f'B{cur}']
        ac_cell.value     = action
        ac_cell.font      = Font(name='Arial', size=9, color=col)
        ac_cell.fill      = fill(C['grey'])
        ac_cell.alignment = lft()
        cur += 1

        if df_a.empty:
            ws2.merge_cells(f'B{cur}:K{cur}')
            ec = ws2[f'B{cur}']
            ec.value     = "✅ Aucune alerte"
            ec.font      = F(C['green'], 10, True)
            ec.fill      = fill(C['green_l'])
            ec.alignment = ctr()
            cur += 2
            continue

        write_table(ws2, df_a.head(500), cur, a_cols, hcol=col)
        cur += len(df_a.head(500)) + 2

    # ── SHEET 3 : PLAN D'ACTION ───────────────────────────────────────────────
    ws3 = wb.create_sheet("📋 Plan d'Action")
    write_header(
        ws3, "PLAN D'ACTION — ARTICLES À TRAITER",
        f"Aucun Mouvement + Attente Livraison  ·  {today_str}"
    )
    ws3.column_dimensions['A'].width = 2
    pa = (
        det[det['Statut'].isin(['Alerte Aucun Mouvement', 'En Attente Livraison'])]
        .sort_values(['Magasin', 'Origine', 'Statut', 'Sem. Réception'])
    )
    write_table(ws3, pa, 4, a_cols)

    # ── SHEET 4 : DÉTAIL COMPLET ──────────────────────────────────────────────
    ws4 = wb.create_sheet("📦 Détail Complet")
    write_header(ws4, "DÉTAIL COMPLET SKU × MAGASIN", f"Tous statuts  ·  {today_str}")
    ws4.column_dimensions['A'].width = 2
    write_table(ws4, det.sort_values(['Magasin', 'Statut']), 4, a_cols)

    # ── SHEET 5 : CALENDRIER ─────────────────────────────────────────────────
    ws5 = wb.create_sheet("🗓️ Calendrier Flux")
    write_header(ws5, "CALENDRIER FLUX PAR SEMAINE", today_str)
    ws5.column_dimensions['A'].width = 2
    cal_r = det[det['Sem. Réception'].str.match(r'^S\d+$', na=False)].copy()
    if not cal_r.empty:
        cal_r['SN'] = cal_r['Sem. Réception'].apply(sem_sort)
        tbl5 = (
            cal_r.groupby(['Sem. Réception', 'SN', 'Origine']).agg(
                Articles=('SKU', 'nunique'),
                Terminé=('Statut',  lambda x: (x == 'Implantation Terminée').sum()),
                Attente=('Statut',  lambda x: (x == 'En Attente Livraison').sum()),
                Alerte=('Statut',   lambda x: (x == 'Alerte Aucun Mouvement').sum()),
            )
            .reset_index().sort_values('SN').drop(columns='SN')
        )
        c5 = [
            ("SEMAINE",  "Sem. Réception", 12, 'c', None),
            ("ORIGINE",  "Origine",        10, 'c', lambda v, r: os_.get(v, (C['grey'], C['dark']))),
            ("ARTICLES", "Articles",       12, 'c', None),
            ("TERMINÉ",  "Terminé",        12, 'c', None),
            ("ATTENTE",  "Attente",        12, 'c', None),
            ("ALERTE",   "Alerte",         12, 'c', None),
        ]
        write_table(ws5, tbl5, 4, c5)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── BOUTON EXPORT ────────────────────────────────────────────────────────────
EXPORT_COLS = [
    "Magasin", "SKU", "Libellé article", "Origine", "Mode Appro",
    "Sem. Réception", "Date Livraison", "Code etat", "Etat Actif",
    "Stock", "RAL", "Statut"
]

col_dl, col_info = st.columns([1, 2])

with col_info:
    st.markdown(f"""
    <div style="background:var(--accent-l);border:1px solid var(--accent-bd);
         border-radius:var(--radius);padding:14px 18px;">
      <div style="font-size:13px;font-weight:700;color:var(--accent)">
        📄 Rapport Implantation_{TODAY_FILE}.xlsx
      </div>
      <div style="font-size:12px;color:var(--muted);margin-top:4px;">
        5 onglets · Résumé Exécutif (IM/LO) · Alertes &amp; Actions · Plan d'Action · Détail · Calendrier<br>
        <strong>{len(mag_actifs)} magasin(s)</strong> · <strong>{total_sku_sel} SKU</strong> ·
        <strong style="color:#dc2626">{ca + cal} articles à traiter</strong>
      </div>
    </div>""", unsafe_allow_html=True)

with col_dl:
    det_b = io.BytesIO()
    detail_df[EXPORT_COLS].to_parquet(det_b)
    det_b.seek(0)

    piv_b = io.BytesIO()
    pivot.to_parquet(piv_b)
    piv_b.seek(0)

    report = build_report(
        det_b.getvalue(), piv_b.getvalue(),
        TODAY_STR, TODAY_FILE,
        len(mag_actifs), total_sku_sel,
        ct, ca, cal, avg_impl,
        sku_im_total, sku_lo_total,
        taux_implantation(df_im), taux_implantation(df_lo)
    )
    st.download_button(
        label=f"📥 Rapport Implantation_{TODAY_FILE}.xlsx",
        data=report,
        file_name=f"Rapport_Implantation_{TODAY_FILE}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
