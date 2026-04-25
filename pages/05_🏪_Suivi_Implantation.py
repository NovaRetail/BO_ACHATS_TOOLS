"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  03_📦_Implantation.py  ·  SmartBuyer · NovaRetail Solutions               ║
║  Suivi Implantation Nouvelles Références — v3.0                             ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  Parser natif export PBI pivoté (magasins en colonnes)                      ║
║  Détection alertes : stock NaN / négatif / seuil                            ║
║  Filtrage CI vs CM, par rayon, flux IM/LO                                   ║
║  Export Excel enrichi · Cessions inter-magasins                             ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import re
from datetime import date

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES MÉTIER
# ─────────────────────────────────────────────────────────────────────────────
TODAY      = date.today()
TODAY_STR  = TODAY.strftime("%d %b %Y")
TODAY_FILE = TODAY.strftime("%Y%m%d")

# Préfixes codes magasins par pays
CI_PREFIX  = ("10",)          # Côte d'Ivoire
CM_PREFIX  = ("20",)          # Cameroun

# Couleurs charte SmartBuyer (identique aux autres modules)
C = {
    "bg"        : "#F2F2F7",
    "surface"   : "#FFFFFF",
    "border"    : "#E5E5EA",
    "text"      : "#1C1C1E",
    "muted"     : "#6D6D72",
    "blue"      : "#007AFF",
    "green"     : "#34C759",
    "red"       : "#FF3B30",
    "orange"    : "#FF9500",
    "purple"    : "#AF52DE",
    "blue_l"    : "#EFF4FF",
    "green_l"   : "#F0FFF4",
    "red_l"     : "#FFF2F0",
    "orange_l"  : "#FFFBEB",
}

STATUT_COLORS = {
    "✅ Implanté"          : C["green"],
    "🔴 Stock négatif"     : C["red"],
    "⚠️ Non implanté"      : C["orange"],
}

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Implantation · SmartBuyer",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CHARTE GRAPHIQUE — identique aux modules OTIF / Perf Hebdo
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

:root {{
  --bg:{C['bg']}; --surface:{C['surface']}; --border:{C['border']};
  --text:{C['text']}; --muted:{C['muted']};
  --blue:{C['blue']}; --green:{C['green']}; --red:{C['red']};
  --orange:{C['orange']}; --purple:{C['purple']};
  --blue-l:{C['blue_l']}; --green-l:{C['green_l']};
  --red-l:{C['red_l']}; --orange-l:{C['orange_l']};
  --radius:14px;
  --shadow:0 1px 3px rgba(0,0,0,.06), 0 4px 16px rgba(0,0,0,.04);
}}

html, body, [class*="css"] {{
  font-family:'Inter',sans-serif!important;
  background:var(--bg)!important;
  color:var(--text)!important;
}}
.main, section[data-testid="stMain"] {{ background:var(--bg)!important; }}
.block-container {{ padding:0 2rem 4rem!important; max-width:1480px; }}
header[data-testid="stHeader"], #MainMenu, footer {{ display:none!important; }}

/* TOPBAR */
.topbar {{
  background:{C['text']}; margin:0 -2rem 28px; padding:16px 28px;
  display:flex; align-items:center; justify-content:space-between;
}}
.topbar-icon {{
  width:40px; height:40px; border-radius:10px;
  background:linear-gradient(135deg,{C['blue']},{C['purple']});
  display:flex; align-items:center; justify-content:center; font-size:22px;
}}
.topbar-title {{ font-size:17px; font-weight:700; color:#fff; letter-spacing:-.01em; }}
.topbar-sub {{ font-size:11px; color:#8E8E93; font-family:'JetBrains Mono'; margin-top:2px; }}
.topbar-pill {{
  background:rgba(255,255,255,.08); color:#8E8E93;
  border:1px solid rgba(255,255,255,.12); border-radius:8px;
  padding:4px 14px; font-size:11px; font-weight:600;
}}
.topbar-date {{ color:{C['blue']}; font-size:12px; font-family:'JetBrains Mono'; }}

/* INTRO MODULE */
.module-intro {{
  background:var(--surface); border:1px solid var(--border);
  border-radius:var(--radius); padding:20px 24px; margin-bottom:24px;
  box-shadow:var(--shadow);
}}
.module-intro-title {{
  font-size:15px; font-weight:700; color:var(--text); margin-bottom:6px;
}}
.module-intro-body {{
  font-size:13px; color:var(--muted); line-height:1.6;
}}
.tag-badge {{
  display:inline-flex; align-items:center; gap:5px;
  background:var(--blue-l); color:var(--blue);
  border:1px solid #BFDBFE; border-radius:6px;
  padding:3px 10px; font-size:11px; font-weight:600; margin:2px;
}}
.tag-badge.green {{ background:var(--green-l); color:var(--green); border-color:#6EE7B7; }}
.tag-badge.red   {{ background:var(--red-l);   color:var(--red);   border-color:#FECACA; }}
.tag-badge.orange{{ background:var(--orange-l);color:var(--orange);border-color:#FCD34D; }}

/* KPI CARDS */
.kpi-grid {{ display:grid; grid-template-columns:repeat(5,1fr); gap:12px; margin-bottom:24px; }}
.kpi-card {{
  background:var(--surface); border:1px solid var(--border);
  border-radius:var(--radius); padding:18px 20px 14px;
  box-shadow:var(--shadow); position:relative; overflow:hidden;
}}
.kpi-card::before {{
  content:''; position:absolute; top:0; left:0; right:0; height:3px;
  border-radius:var(--radius) var(--radius) 0 0;
}}
.kpi-card.blue::before  {{ background:var(--blue);  }}
.kpi-card.green::before {{ background:var(--green); }}
.kpi-card.red::before   {{ background:var(--red);   }}
.kpi-card.orange::before{{ background:var(--orange);}}
.kpi-card.purple::before{{ background:var(--purple);}}
.kpi-label {{
  font-size:10px; font-weight:700; text-transform:uppercase;
  letter-spacing:.10em; color:var(--muted); margin-bottom:10px;
}}
.kpi-value {{
  font-size:38px; font-weight:800; line-height:1; letter-spacing:-.02em;
}}
.kpi-card.blue  .kpi-value {{ color:var(--blue);  }}
.kpi-card.green .kpi-value {{ color:var(--green); }}
.kpi-card.red   .kpi-value {{ color:var(--red);   }}
.kpi-card.orange.kpi-value {{ color:var(--orange);}}
.kpi-card.purple.kpi-value {{ color:var(--purple);}}
.kpi-sub {{
  font-size:11px; color:var(--muted); font-family:'JetBrains Mono'; margin-top:4px;
}}
.kpi-bar {{
  margin-top:12px; height:3px; border-radius:3px; background:var(--border);
}}
.kpi-bar-fill {{ height:100%; border-radius:3px; }}

/* SCORECARD MAGASINS */
.scorecard-grid {{
  display:grid; grid-template-columns:repeat(auto-fill,minmax(160px,1fr));
  gap:10px; margin-bottom:24px;
}}
.scorecard-card {{
  background:var(--surface); border:1px solid var(--border);
  border-radius:var(--radius); padding:14px 16px;
  box-shadow:var(--shadow); position:relative;
}}
.scorecard-card.ok  {{ border-color:#6EE7B7; background:var(--green-l); }}
.scorecard-card.ko  {{ border-color:#FECACA; background:var(--red-l);   }}
.scorecard-card.warn{{ border-color:#FCD34D; background:var(--orange-l);}}
.scorecard-dot {{
  width:8px; height:8px; border-radius:50%;
  position:absolute; top:14px; right:14px;
}}
.scorecard-name {{
  font-size:11px; font-weight:600; color:var(--text);
  margin-bottom:6px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
}}
.scorecard-pct  {{ font-size:28px; font-weight:800; line-height:1; }}
.scorecard-sub  {{ font-size:10px; color:var(--muted); font-family:'JetBrains Mono'; margin-top:2px; }}

/* ALERT BANNER */
.alert-banner {{
  background:#FFF; border:1px solid #FECACA; border-left:4px solid var(--red);
  border-radius:var(--radius); padding:16px 20px; margin-bottom:20px;
  display:flex; align-items:center; gap:16px; flex-wrap:wrap;
}}
.alert-pill {{
  background:var(--red); color:#fff; border-radius:6px;
  padding:4px 12px; font-size:11px; font-weight:700;
  letter-spacing:.05em; white-space:nowrap;
}}
.alert-item {{ display:flex; flex-direction:column; align-items:center; padding:0 16px; border-right:1px solid var(--border); }}
.alert-item:last-child {{ border-right:none; }}
.alert-num {{ font-size:28px; font-weight:800; line-height:1; }}
.alert-lbl {{ font-size:10px; font-weight:600; color:var(--muted); text-transform:uppercase; letter-spacing:.06em; margin-top:1px; }}

/* SECTION HEADER */
.sh {{
  font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:.12em;
  color:var(--muted); margin:24px 0 14px; padding-bottom:8px;
  border-bottom:1px solid var(--border);
}}

/* SIDEBAR */
section[data-testid="stSidebar"] {{
  background:#fff!important; border-right:1px solid var(--border)!important;
  min-width:270px!important; max-width:270px!important;
}}
section[data-testid="stSidebar"] .block-container {{
  padding:.6rem .8rem 2rem!important;
}}

/* DOWNLOAD BUTTON */
.stDownloadButton>button {{
  background:linear-gradient(135deg,{C['text']},{C['blue']})!important;
  color:#fff!important; border:none!important;
  border-radius:10px!important; font-weight:700!important;
  font-size:13px!important; padding:12px!important; width:100%!important;
  box-shadow:0 4px 12px rgba(0,122,255,.25)!important;
}}

/* INFO BANNERS */
.info-box {{
  border-radius:var(--radius); padding:14px 18px; margin-bottom:16px;
  border:1px solid; font-size:13px; line-height:1.6;
}}
.info-box.blue   {{ background:var(--blue-l);   border-color:#BFDBFE; color:#1D4ED8; }}
.info-box.green  {{ background:var(--green-l);  border-color:#6EE7B7; color:#065F46; }}
.info-box.orange {{ background:var(--orange-l); border-color:#FCD34D; color:#92400E; }}

/* TABS CUSTOM */
.nav-active {{
  background:var(--text)!important; color:#fff!important;
  border-radius:10px; padding:10px 0; text-align:center;
  font-size:13px; font-weight:700;
  box-shadow:0 4px 14px rgba(28,28,30,.2); margin-bottom:10px;
}}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def fmt_n(n: int) -> str:
    """Format entier avec séparateur espace insécable."""
    return f"{n:,}".replace(",", "\u202f")

def color_taux(t: float) -> str:
    if t >= 80: return C["green"]
    if t >= 50: return C["orange"]
    return C["red"]

def scorecard_cls(t: float) -> str:
    if t >= 80: return "ok"
    if t >= 50: return "warn"
    return "ko"


# ─────────────────────────────────────────────────────────────────────────────
# PARSER PBI — Export pivoté natif SmartBuyer
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_pbi_stock(file_bytes: bytes, filename: str) -> tuple[pd.DataFrame | None, str | None]:
    """
    Parse l'export PBI pivoté (format NovaRetail).

    Structure attendue :
        Ligne 0 : [en-tête label] | [Code - Nom magasin] × N | Total
        Ligne 1 : [Article]       | [Stock] × N              | Stock
        Ligne 2+ : données articles

    Retourne un DataFrame long :
        SKU | Libellé article | Code site | Libellé site | Stock
    """
    buf = io.BytesIO(file_bytes)
    try:
        if filename.lower().endswith((".xlsx", ".xls")):
            raw = pd.read_excel(buf, header=None, dtype=str)
        else:
            buf.seek(0)
            raw = pd.read_csv(buf, header=None, sep=None, engine="python",
                              encoding="latin1", on_bad_lines="skip", dtype=str)
    except Exception as e:
        return None, f"Lecture fichier : {e}"

    if raw.shape[0] < 3 or raw.shape[1] < 3:
        return None, "Fichier trop court ou mal formaté"

    # — Extraction des magasins (ligne 0, colonnes 1 à avant-dernière) —
    sites_raw = raw.iloc[0, 1:-1].tolist()   # exclure colonne Total

    # — Données articles (lignes 2+) —
    records = []
    for _, row in raw.iloc[2:].iterrows():
        art_raw = str(row.iloc[0]).strip()
        if not art_raw or art_raw.lower() in ("nan", "article", "total", "rayon"):
            continue

        # Parsing "10000119 - 4X25CL BOIS,EN,RED BULL MM"
        if " - " in art_raw:
            parts = art_raw.split(" - ", 1)
            sku = parts[0].strip().zfill(8)[:8]
            lib = parts[1].strip()
        else:
            m = re.match(r"^(\d{6,8})", art_raw)
            sku = m.group(1).zfill(8)[:8] if m else art_raw[:8].zfill(8)
            lib = art_raw

        if not re.match(r"^\d{8}$", sku):
            continue

        # Stock par magasin
        for col_idx, site_raw in enumerate(sites_raw, start=1):
            site_str = str(site_raw).strip()
            if not site_str or site_str.lower() == "nan":
                continue
            if " - " in site_str:
                code_site, lib_site = site_str.split(" - ", 1)
                code_site = code_site.strip()
                lib_site  = lib_site.strip()
            else:
                code_site = site_str
                lib_site  = site_str

            val = row.iloc[col_idx]
            try:
                stock = int(float(str(val).replace(",", "."))) if pd.notna(val) and str(val).strip().lower() not in ("nan", "") else None
            except (ValueError, TypeError):
                stock = None

            records.append({
                "SKU"           : sku,
                "Libellé article": lib,
                "Code site"     : code_site.strip(),
                "Libellé site"  : lib_site.strip(),
                "Stock"         : stock,
            })

    if not records:
        return None, "Aucune donnée extraite — vérifiez le format du fichier PBI"

    df = pd.DataFrame(records)
    return df, None


@st.cache_data(show_spinner=False)
def load_t1(file_bytes: bytes, filename: str) -> tuple[pd.DataFrame | None, str | None]:
    """
    Parse le fichier T1 (liste nouvelles références à implanter).
    Colonnes attendues : ARTICLE (code 8 chiffres), LIBELLÉ ARTICLE, MODE APPRO,
                         SEMAINE RECEPTION, LIBELLÉ FOURNISSEUR ORIGINE, DATE LIV.
    """
    buf = io.BytesIO(file_bytes)
    try:
        if filename.lower().endswith((".xlsx", ".xls")):
            df_raw = pd.read_excel(buf, header=None, dtype=str)
        else:
            buf.seek(0)
            df_raw = pd.read_csv(buf, header=None, sep=None, engine="python",
                                 encoding="latin1", on_bad_lines="skip", dtype=str)
    except Exception as e:
        return None, f"Lecture T1 : {e}"

    if df_raw.empty:
        return None, "Fichier T1 vide"

    # Détection auto header : si la première cellule est numérique → sans header
    first_val = str(df_raw.iloc[0, 0]).strip().replace(".0", "")
    has_header = not first_val.isdigit()

    if has_header:
        df_raw.columns = df_raw.iloc[0].astype(str).str.strip().str.upper()
        df_raw = df_raw.iloc[1:].reset_index(drop=True)
    else:
        df_raw.columns = ["ARTICLE"] + [f"_COL{i}" for i in range(1, len(df_raw.columns))]

    # Normaliser les noms de colonnes
    df_raw.columns = (
        df_raw.columns.astype(str)
        .str.strip()
        .str.upper()
        .str.replace("\ufeff", "", regex=False)
        .str.replace("\xa0", " ", regex=False)
    )

    if "ARTICLE" not in df_raw.columns:
        return None, "Colonne ARTICLE introuvable dans le fichier T1"

    df_raw["SKU"] = (
        df_raw["ARTICLE"].astype(str).str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.zfill(8).str[:8]
    )
    df_raw = df_raw[df_raw["SKU"].str.match(r"^\d{8}$", na=False)].drop_duplicates("SKU").copy()

    # Colonnes optionnelles avec valeurs par défaut
    defaults = {
        "LIBELLÉ ARTICLE"              : "",
        "LIBELLÉ FOURNISSEUR ORIGINE"  : "",
        "MODE APPRO"                   : "",
        "SEMAINE RECEPTION"            : "",
        "DATE LIV."                    : "",
    }
    for col, default in defaults.items():
        if col not in df_raw.columns:
            df_raw[col] = default

    df_raw["SEMAINE RECEPTION"] = df_raw["SEMAINE RECEPTION"].astype(str).str.strip().replace("nan", "")
    def _sem_to_num(s):
        s = str(s).strip()
        cleaned = re.sub(r"[Ss]", "", s)
        return int(cleaned) if cleaned.isdigit() else 99

    df_raw["SEM_NUM"] = df_raw["SEMAINE RECEPTION"].apply(_sem_to_num)
    df_raw["ORIGINE"] = df_raw["MODE APPRO"].apply(
        lambda m: "IM" if "IMPORT" in str(m).upper() else "LO"
    )

    return df_raw, None


# ─────────────────────────────────────────────────────────────────────────────
# TOPBAR
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="topbar">
  <div style="display:flex;align-items:center;gap:14px;">
    <div class="topbar-icon">📦</div>
    <div>
      <div class="topbar-title">Suivi Implantation Nouvelles Références</div>
      <div class="topbar-sub">Parser PBI natif · Alertes stock · Cessions inter-magasins</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:12px;">
    <div class="topbar-date">{TODAY_STR}</div>
    <div class="topbar-pill">v3.0 · SmartBuyer</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# BLOC INTRO MÉTIER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="module-intro">
  <div class="module-intro-title">📦 À quoi sert ce module ?</div>
  <div class="module-intro-body">
    Ce module contrôle l'implantation physique des <strong>nouvelles références T1</strong> dans les magasins du réseau.
    Après chaque vague de lancement, il croise le fichier T1 (liste des articles à implanter) avec l'export
    PBI de stock pour détecter en temps réel les articles <strong>non encore présents</strong> (stock NaN),
    en <strong>stock négatif</strong> (écart d'inventaire à corriger) ou correctement <strong>implantés</strong> (stock &gt; 0).<br><br>
    <strong>Cas d'usage :</strong>
    <span class="tag-badge">📋 Suivi COPIL lancement</span>
    <span class="tag-badge green">✅ Validation implantation rayon</span>
    <span class="tag-badge red">🚨 Escalade fournisseur stock négatif</span>
    <span class="tag-badge orange">🔄 Cession inter-magasins</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — CHARGEMENT
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📁 Fichiers")
    st.divider()

    st.markdown("**① Liste T1 — Nouvelles Références**")
    st.caption("Fichier ERP : colonnes ARTICLE, MODE APPRO, SEMAINE RECEPTION…")
    t1_file = st.file_uploader("T1", type=["csv", "xlsx", "xls"],
                               key="t1", label_visibility="collapsed")

    st.markdown("**② Export PBI — Stocks par magasin**")
    st.caption("Export pivoté : lignes = articles, colonnes = magasins, valeurs = Stock")
    pbi_file = st.file_uploader("PBI Stock", type=["xlsx", "xls", "csv"],
                                key="pbi", label_visibility="collapsed")

# — Gates de chargement —
if not t1_file:
    st.markdown('<div class="info-box blue">⬆️ <strong>Étape 1</strong> — Charge le fichier T1 (Nouvelles Références) dans la sidebar.</div>', unsafe_allow_html=True)
    st.stop()

with st.spinner("Lecture T1…"):
    t1_df, t1_err = load_t1(t1_file.read(), t1_file.name)

if t1_err or t1_df is None:
    st.error(f"❌ T1 : {t1_err}")
    st.stop()

if not pbi_file:
    st.markdown(f'<div class="info-box blue">✅ T1 chargé — <strong>{len(t1_df):,}</strong> références. ⬆️ <strong>Étape 2</strong> — Charge maintenant l\'export PBI Stock.</div>', unsafe_allow_html=True)
    st.stop()

with st.spinner("Parsing export PBI…"):
    pbi_bytes = pbi_file.read()
    df_stock, pbi_err = parse_pbi_stock(pbi_bytes, pbi_file.name)

if pbi_err or df_stock is None:
    st.error(f"❌ PBI : {pbi_err}")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — FILTRES
# ─────────────────────────────────────────────────────────────────────────────
# Détecter les pays
all_sites    = df_stock["Libellé site"].unique().tolist()
all_codes    = df_stock["Code site"].unique().tolist()
sites_ci     = sorted([s for s in all_sites if any(
                   df_stock.loc[df_stock["Libellé site"]==s, "Code site"].str.startswith(p).any()
                   for p in CI_PREFIX)])
sites_cm     = sorted([s for s in all_sites if any(
                   df_stock.loc[df_stock["Libellé site"]==s, "Code site"].str.startswith(p).any()
                   for p in CM_PREFIX)])
sites_autres = sorted([s for s in all_sites if s not in sites_ci and s not in sites_cm])

with st.sidebar:
    st.divider()
    st.markdown("## 🔍 Filtres")

    # Pays / périmètre
    pays_opt = []
    if sites_ci: pays_opt.append("🇨🇮 Côte d'Ivoire")
    if sites_cm: pays_opt.append("🇨🇲 Cameroun")
    if sites_autres: pays_opt.append("Autres")
    pays_sel = st.multiselect("Périmètre", pays_opt,
                              default=["🇨🇮 Côte d'Ivoire"] if "🇨🇮 Côte d'Ivoire" in pays_opt else pays_opt)

    # Construction liste magasins selon pays
    mag_pool = []
    if "🇨🇮 Côte d'Ivoire" in pays_sel: mag_pool += sites_ci
    if "🇨🇲 Cameroun"       in pays_sel: mag_pool += sites_cm
    if "Autres"             in pays_sel: mag_pool += sites_autres
    mag_pool = sorted(set(mag_pool))

    mag_sel = st.multiselect("Magasins", mag_pool, default=mag_pool)

    # Flux
    orig_sel = st.multiselect("Flux", ["IM", "LO"], default=["IM", "LO"])

    # Semaine
    sem_dispo = sorted(
        [s for s in t1_df["SEMAINE RECEPTION"].unique()
         if s and s not in ("nan", "", "99")],
        key=lambda s: int(re.sub(r"[Ss]", "", s)) if re.match(r"^[Ss]?\d+$", s.strip()) else 99
    )
    sem_sel = st.multiselect("Semaine réception", sem_dispo, default=sem_dispo)

    st.divider()
    st.markdown("## 🔄 Cessions")
    mag_detresse = st.multiselect(
        "Magasins en détresse (stock ≤ seuil)",
        options=mag_pool, default=[]
    )
    seuil_det = st.number_input("Seuil stock détresse (≤)", 0, 50, 0, 1)
    reserve   = st.number_input("Réserve magasin cédant (≥)", 0, 50, 2, 1)

if not mag_sel:
    st.warning("⚠️ Sélectionne au moins un magasin.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# CONSTRUCTION DU DATASET PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────
# Filtre SKU selon origine + semaine
mask = (
    t1_df["ORIGINE"].isin(orig_sel)
    & (t1_df["SEMAINE RECEPTION"].isin(sem_sel) if sem_sel else pd.Series(True, index=t1_df.index))
)
t1_scope = t1_df[mask].copy()
sku_scope = t1_scope["SKU"].unique()

if len(sku_scope) == 0:
    st.warning("Aucun SKU correspondant aux filtres sélectionnés.")
    st.stop()

# Référentiel T1 : index par SKU
t1_ref = t1_scope.set_index("SKU")[[
    "LIBELLÉ ARTICLE", "LIBELLÉ FOURNISSEUR ORIGINE",
    "MODE APPRO", "SEMAINE RECEPTION", "DATE LIV.", "ORIGINE", "SEM_NUM"
]]

# Grille complète magasins × SKU (produit cartésien)
grid = pd.DataFrame(
    pd.MultiIndex.from_product(
        [mag_sel, sku_scope], names=["Libellé site", "SKU"]
    ).tolist(),
    columns=["Libellé site", "SKU"]
)

# Jointure avec les stocks PBI
stock_scope = df_stock[
    df_stock["Libellé site"].isin(mag_sel) &
    df_stock["SKU"].isin(sku_scope)
][["Libellé site", "SKU", "Stock", "Libellé article"]].copy()

merged = grid.merge(stock_scope, on=["Libellé site", "SKU"], how="left")

# Jointure avec T1
merged = merged.merge(
    t1_ref.reset_index().rename(columns={
        "LIBELLÉ ARTICLE"             : "T1_lib",
        "LIBELLÉ FOURNISSEUR ORIGINE" : "Fournisseur",
        "MODE APPRO"                  : "Mode Appro",
        "SEMAINE RECEPTION"           : "Sem. Réception",
        "DATE LIV."                   : "Date Livraison",
        "ORIGINE"                     : "Origine",
        "SEM_NUM"                     : "SEM_NUM",
    }),
    on="SKU", how="left"
)

# Libellé article : PBI > T1
merged["Libellé article"] = merged["Libellé article"].fillna("").astype(str)
merged["Libellé article"] = merged.apply(
    lambda r: r["Libellé article"] if r["Libellé article"] else r.get("T1_lib", ""), axis=1
)
merged.drop(columns=["T1_lib"], errors="ignore", inplace=True)
merged.rename(columns={"Libellé site": "Magasin"}, inplace=True)

# Statut métier
def statut(row):
    s = row["Stock"]
    if pd.isna(s):   return "⚠️ Non implanté"
    if s < 0:        return "🔴 Stock négatif"
    return "✅ Implanté"

merged["Statut"] = merged.apply(statut, axis=1)

# ─────────────────────────────────────────────────────────────────────────────
# MÉTRIQUES RÉSEAU
# ─────────────────────────────────────────────────────────────────────────────
total_cells  = len(merged)
n_implante   = int((merged["Statut"] == "✅ Implanté").sum())
n_non_impl   = int((merged["Statut"] == "⚠️ Non implanté").sum())
n_negatif    = int((merged["Statut"] == "🔴 Stock négatif").sum())
n_sku        = len(sku_scope)
n_mag        = len(mag_sel)
taux_reseau  = int(n_implante / total_cells * 100) if total_cells else 0

n_sku_im     = int((t1_scope["ORIGINE"] == "IM").sum())
n_sku_lo     = int((t1_scope["ORIGINE"] == "LO").sum())

pct = lambda n: int(n / total_cells * 100) if total_cells else 0

# ─────────────────────────────────────────────────────────────────────────────
# BANNIÈRE ALERTES
# ─────────────────────────────────────────────────────────────────────────────
if n_non_impl + n_negatif > 0:
    st.markdown(f"""
    <div class="alert-banner">
      <div class="alert-pill">⚡ ACTIONS REQUISES</div>
      <div class="alert-item">
        <div class="alert-num" style="color:{C['orange']}">{fmt_n(n_non_impl)}</div>
        <div class="alert-lbl">Non implanté</div>
      </div>
      <div class="alert-item">
        <div class="alert-num" style="color:{C['red']}">{fmt_n(n_negatif)}</div>
        <div class="alert-lbl">Stock négatif</div>
      </div>
      <div style="margin-left:auto;font-size:13px;color:{C['muted']};">
        {n_mag} magasins · {fmt_n(n_sku)} SKUs · {fmt_n(total_cells)} combinaisons
      </div>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# KPI GRID (5 cartes)
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card green">
    <div class="kpi-label">✅ Implanté</div>
    <div class="kpi-value" style="color:{C['green']}">{fmt_n(n_implante)}</div>
    <div class="kpi-sub">{pct(n_implante)}% du réseau</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(n_implante)}%;background:{C['green']}"></div></div>
  </div>
  <div class="kpi-card orange">
    <div class="kpi-label">⚠️ Non implanté</div>
    <div class="kpi-value" style="color:{C['orange']}">{fmt_n(n_non_impl)}</div>
    <div class="kpi-sub">{pct(n_non_impl)}% — à traiter</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(n_non_impl)}%;background:{C['orange']}"></div></div>
  </div>
  <div class="kpi-card red">
    <div class="kpi-label">🔴 Stock négatif</div>
    <div class="kpi-value" style="color:{C['red']}">{fmt_n(n_negatif)}</div>
    <div class="kpi-sub">{pct(n_negatif)}% — écart inventaire</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(n_negatif)}%;background:{C['red']}"></div></div>
  </div>
  <div class="kpi-card blue">
    <div class="kpi-label">📊 Taux réseau</div>
    <div class="kpi-value" style="color:{color_taux(taux_reseau)}">{taux_reseau}%</div>
    <div class="kpi-sub">{n_mag} mag × {fmt_n(n_sku)} SKU</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{taux_reseau}%;background:{C['blue']}"></div></div>
  </div>
  <div class="kpi-card purple">
    <div class="kpi-label">🔀 Flux IM / LO</div>
    <div class="kpi-value" style="color:{C['purple']}">{n_sku_im}<span style="font-size:18px;font-weight:500;color:{C['muted']}"> / </span>{n_sku_lo}</div>
    <div class="kpi-sub">Import · Local</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SCORECARD PAR MAGASIN
# ─────────────────────────────────────────────────────────────────────────────
pivot_mag = (
    merged.groupby(["Magasin", "Statut"])
    .size().unstack(fill_value=0)
    .reindex(columns=list(STATUT_COLORS.keys()), fill_value=0)
    .reset_index()
)
pivot_mag.columns.name = None
pivot_mag["Total"]    = n_sku
pivot_mag["Taux (%)"] = (pivot_mag.get("✅ Implanté", 0) / n_sku * 100).round(0).astype(int)

st.markdown('<div class="sh">SCORECARD MAGASINS</div>', unsafe_allow_html=True)
rag_html = '<div class="scorecard-grid">'
for _, row in pivot_mag.sort_values("Taux (%)", ascending=False).iterrows():
    t_ = row["Taux (%)"]
    cls = scorecard_cls(t_)
    col = color_taux(t_)
    ok_ = int(row.get("✅ Implanté", 0))
    ko_ = int(row.get("⚠️ Non implanté", 0))
    neg = int(row.get("🔴 Stock négatif", 0))
    rag_html += f"""
    <div class="scorecard-card {cls}">
      <div class="scorecard-dot" style="background:{col}"></div>
      <div class="scorecard-name">{row['Magasin']}</div>
      <div class="scorecard-pct" style="color:{col}">{t_}%</div>
      <div class="scorecard-sub">{ok_}✅ {ko_}⚠️ {neg}🔴</div>
    </div>"""
rag_html += "</div>"
st.markdown(rag_html, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# NAVIGATION ONGLETS
# ─────────────────────────────────────────────────────────────────────────────
TABS = ["📊 Vue Réseau", "⚠️ Non Implantés", "🔴 Stocks Négatifs", "🗓️ Calendrier", "🔄 Cessions", "📥 Export"]
if "impl_tab" not in st.session_state:
    st.session_state.impl_tab = TABS[0]

nav_cols = st.columns(len(TABS))
for i, t in enumerate(TABS):
    with nav_cols[i]:
        if st.session_state.impl_tab == t:
            st.markdown(f'<div class="nav-active">{t}</div>', unsafe_allow_html=True)
        if st.button(t, key=f"impl_nav_{i}", use_container_width=True):
            st.session_state.impl_tab = t
            st.rerun()

active = st.session_state.impl_tab

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — VUE RÉSEAU
# ══════════════════════════════════════════════════════════════════════════════
if active == TABS[0]:
    c1, c2 = st.columns([3, 2])

    with c1:
        mel = pivot_mag.melt(
            id_vars="Magasin",
            value_vars=[s for s in STATUT_COLORS if s in pivot_mag.columns],
            var_name="Statut", value_name="N"
        )
        fig = px.bar(
            mel, x="Magasin", y="N", color="Statut",
            color_discrete_map=STATUT_COLORS, barmode="stack",
            title="Situation par magasin"
        )
        fig.update_traces(
            textposition="inside", texttemplate="%{y}",
            textfont=dict(size=11, color="white")
        )
        fig.update_layout(
            paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=420,
            font=dict(family="Inter", color=C["muted"], size=12),
            margin=dict(l=10, r=10, t=44, b=20),
            legend=dict(orientation="h", y=-0.25, bgcolor="rgba(0,0,0,0)"),
            xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]),
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        labels = ["✅ Implanté", "⚠️ Non implanté", "🔴 Stock négatif"]
        values = [n_implante, n_non_impl, n_negatif]
        colors = [C["green"], C["orange"], C["red"]]
        fig_d = go.Figure(go.Pie(
            labels=labels, values=values, hole=0.62,
            marker=dict(colors=colors, line=dict(color="#fff", width=3)),
        ))
        fig_d.add_annotation(
            text=f"<b>{taux_reseau}%</b><br>implanté",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=20, color=C["text"], family="Inter"),
        )
        fig_d.update_layout(
            paper_bgcolor=C["surface"], height=420,
            font=dict(family="Inter", color=C["muted"], size=12),
            margin=dict(l=10, r=10, t=44, b=20),
            legend=dict(orientation="v", x=1.01, bgcolor="rgba(0,0,0,0)"),
            title="Répartition réseau",
        )
        st.plotly_chart(fig_d, use_container_width=True)

    st.markdown('<div class="sh">TABLEAU SYNTHÈSE PAR MAGASIN</div>', unsafe_allow_html=True)
    cols_disp = ["Magasin"] + [c for c in STATUT_COLORS if c in pivot_mag.columns] + ["Total", "Taux (%)"]
    st.dataframe(
        pivot_mag[cols_disp].sort_values("Taux (%)", ascending=False).reset_index(drop=True)
        .style
        .background_gradient(subset=["✅ Implanté"] if "✅ Implanté" in pivot_mag.columns else [], cmap="Greens")
        .background_gradient(subset=["⚠️ Non implanté"] if "⚠️ Non implanté" in pivot_mag.columns else [], cmap="Oranges")
        .background_gradient(subset=["🔴 Stock négatif"] if "🔴 Stock négatif" in pivot_mag.columns else [], cmap="Reds")
        .format({"Taux (%)": "{}%"}),
        use_container_width=True, hide_index=True
    )

    # Répartition IM / LO
    st.markdown('<div class="sh">RÉPARTITION FLUX IM / LO PAR MAGASIN</div>', unsafe_allow_html=True)
    df_flux = merged[merged["Statut"] == "✅ Implanté"].groupby(["Magasin", "Origine"]).size().reset_index(name="N")
    if not df_flux.empty:
        fig_flux = px.bar(
            df_flux, x="Magasin", y="N", color="Origine",
            color_discrete_map={"IM": C["blue"], "LO": C["green"]},
            barmode="group", title="Articles implantés par flux",
        )
        fig_flux.update_layout(
            paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=320,
            font=dict(family="Inter", color=C["muted"], size=12),
            margin=dict(l=10, r=10, t=44, b=20),
            legend=dict(orientation="h", y=-0.28),
            xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]),
        )
        st.plotly_chart(fig_flux, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — NON IMPLANTÉS
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[1]:
    df_ni = merged[merged["Statut"] == "⚠️ Non implanté"].copy()

    if df_ni.empty:
        st.markdown('<div class="info-box green">✅ Tous les articles sont implantés dans les magasins sélectionnés !</div>', unsafe_allow_html=True)
    else:
        # Ruptures communes (tous magasins = 0)
        sku_counts = df_ni.groupby("SKU")["Magasin"].count()
        sku_rupture_totale = sku_counts[sku_counts == len(mag_sel)].index.tolist()

        kc1, kc2, kc3 = st.columns(3)
        kc1.metric("⚠️ Lignes manquantes", fmt_n(len(df_ni)))
        kc2.metric("Articles distincts", fmt_n(len(df_ni["SKU"].unique())))
        kc3.metric("🔴 Rupture totale (tous mag)", len(sku_rupture_totale))

        if sku_rupture_totale:
            st.markdown(f'<div class="info-box orange">⚠️ <strong>{len(sku_rupture_totale)} article(s)</strong> absent(s) de TOUS les magasins — escalade critique recommandée.</div>', unsafe_allow_html=True)

        df_ni["Rupture totale"] = df_ni["SKU"].isin(sku_rupture_totale).map({True: "🔴 OUI", False: "—"})
        COLS = ["Magasin", "SKU", "Libellé article", "Origine", "Mode Appro",
                "Sem. Réception", "Fournisseur", "Rupture totale"]
        st.dataframe(
            df_ni[[c for c in COLS if c in df_ni.columns]]
            .sort_values(["Rupture totale", "Magasin"])
            .reset_index(drop=True),
            use_container_width=True, hide_index=True
        )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — STOCKS NÉGATIFS
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[2]:
    df_neg = merged[merged["Statut"] == "🔴 Stock négatif"].copy()

    if df_neg.empty:
        st.markdown('<div class="info-box green">✅ Aucun stock négatif détecté.</div>', unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="info-box orange">
          <strong>📌 Stock négatif = écart d'inventaire</strong> — un stock négatif indique que des articles
          ont été vendus/sortis sans réception ERP correspondante, ou que l'inventaire n'a pas été mis à jour.
          Action : régularisation inventaire + vérification des entrées ERP.
        </div>
        """, unsafe_allow_html=True)

        kc1, kc2 = st.columns(2)
        kc1.metric("🔴 Lignes stock négatif", fmt_n(len(df_neg)))
        kc2.metric("Articles distincts", fmt_n(len(df_neg["SKU"].unique())))

        COLS = ["Magasin", "SKU", "Libellé article", "Origine", "Stock", "Mode Appro", "Fournisseur"]
        st.dataframe(
            df_neg[[c for c in COLS if c in df_neg.columns]]
            .sort_values("Stock")  # Les plus négatifs en premier
            .reset_index(drop=True),
            use_container_width=True, hide_index=True
        )

        # Distribution des valeurs négatives
        fig_neg = px.histogram(
            df_neg, x="Stock", nbins=40,
            title="Distribution des stocks négatifs",
            color_discrete_sequence=[C["red"]]
        )
        fig_neg.update_layout(
            paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=280,
            font=dict(family="Inter", color=C["muted"], size=12),
            margin=dict(l=10, r=10, t=44, b=20),
        )
        st.plotly_chart(fig_neg, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — CALENDRIER
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[3]:
    cal_df = merged[merged["Sem. Réception"].str.match(r"^[Ss]?\d+$", na=False)].copy()

    if cal_df.empty:
        st.info("Aucune semaine de réception renseignée dans le fichier T1.")
    else:
        cal_agg = cal_df.groupby("Sem. Réception").agg(
            Implanté     =("Statut", lambda x: (x == "✅ Implanté").sum()),
            Non_implanté =("Statut", lambda x: (x == "⚠️ Non implanté").sum()),
            Stock_négatif=("Statut", lambda x: (x == "🔴 Stock négatif").sum()),
            SKU_distincts =("SKU", "nunique"),
        ).reset_index().rename(columns={"Non_implanté": "Non implanté", "Stock_négatif": "Stock négatif"})
        cal_agg["Taux (%)"] = (cal_agg["Implanté"] / (cal_agg["Implanté"] + cal_agg["Non implanté"] + cal_agg["Stock négatif"]) * 100).round(0).astype(int)

        st.markdown('<div class="sh">PROGRESSION PAR SEMAINE DE RÉCEPTION</div>', unsafe_allow_html=True)
        fig_cal = px.bar(
            cal_agg.melt(id_vars="Sem. Réception",
                         value_vars=["Implanté", "Non implanté", "Stock négatif"]),
            x="Sem. Réception", y="value", color="variable",
            color_discrete_map={"Implanté": C["green"], "Non implanté": C["orange"], "Stock négatif": C["red"]},
            barmode="stack", title="Statut par semaine de réception"
        )
        fig_cal.update_layout(
            paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=360,
            font=dict(family="Inter", color=C["muted"], size=12),
            margin=dict(l=10, r=10, t=44, b=20),
            legend=dict(orientation="h", y=-0.25),
            xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]),
        )
        st.plotly_chart(fig_cal, use_container_width=True)
        st.dataframe(cal_agg, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — CESSIONS
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[4]:
    st.markdown('<div class="sh">🔄 MOTEUR CESSIONS INTER-MAGASINS</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="info-box blue">
      <strong>💡 Comment utiliser ?</strong><br>
      Dans la sidebar, sélectionne les <strong>magasins en détresse</strong> (stock insuffisant) et le seuil de déclenchement.
      Le moteur identifie les magasins cédants disposant d'un stock excédentaire et propose un plan de cession article par article.
    </div>
    """, unsafe_allow_html=True)

    if not mag_detresse:
        st.markdown('<div class="info-box orange">⬅️ Sélectionne des magasins en détresse dans la sidebar pour activer le moteur de cessions.</div>', unsafe_allow_html=True)
    else:
        mag_cedants = [m for m in mag_sel if m not in mag_detresse]
        suggestions = []

        for sku in sku_scope:
            sku_df = df_stock[df_stock["SKU"] == sku].copy()
            if sku_df.empty:
                continue
            lib = sku_df["Libellé article"].iloc[0]

            det_rows = sku_df[
                sku_df["Libellé site"].isin(mag_detresse) &
                (sku_df["Stock"].fillna(0) <= seuil_det)
            ]
            if det_rows.empty:
                continue

            ced_rows = sku_df[
                sku_df["Libellé site"].isin(mag_cedants) &
                (sku_df["Stock"].fillna(0) > reserve)
            ].sort_values("Stock", ascending=False)

            for _, dr in det_rows.iterrows():
                if ced_rows.empty:
                    suggestions.append({
                        "SKU": sku, "Libellé article": lib,
                        "Magasin détresse": dr["Libellé site"],
                        "Stock détresse": int(dr["Stock"]) if pd.notna(dr["Stock"]) else "NaN",
                        "Cédant suggéré": "⚠️ Aucun cédant disponible",
                        "Stock cédant": 0, "Qté cessible": 0,
                        "Faisabilité": "🔴 Impossible",
                    })
                else:
                    best = ced_rows.iloc[0]
                    qty  = int(best["Stock"]) - reserve
                    suggestions.append({
                        "SKU": sku, "Libellé article": lib,
                        "Magasin détresse": dr["Libellé site"],
                        "Stock détresse": int(dr["Stock"]) if pd.notna(dr["Stock"]) else "NaN",
                        "Cédant suggéré": best["Libellé site"],
                        "Stock cédant": int(best["Stock"]),
                        "Qté cessible": qty,
                        "Faisabilité": "🟢 Possible" if qty >= 1 else "🟠 Partielle",
                    })

        if not suggestions:
            st.success("✅ Aucune cession nécessaire selon les critères actuels.")
        else:
            df_cess = pd.DataFrame(suggestions).sort_values(
                ["Faisabilité", "Qté cessible"], ascending=[True, False]
            ).reset_index(drop=True)

            c1, c2, c3 = st.columns(3)
            c1.metric("Cessions possibles", int((df_cess["Faisabilité"] == "🟢 Possible").sum()))
            c2.metric("Impossible", int((df_cess["Faisabilité"] == "🔴 Impossible").sum()))
            c3.metric("Articles concernés", df_cess["SKU"].nunique())

            st.dataframe(df_cess, use_container_width=True, hide_index=True)

            buf_c = io.BytesIO()
            with pd.ExcelWriter(buf_c, engine="openpyxl") as writer:
                df_cess.to_excel(writer, sheet_name="Plan Cessions", index=False)
            buf_c.seek(0)
            st.download_button(
                f"📥 Télécharger Plan_Cessions_{TODAY_FILE}.xlsx",
                data=buf_c,
                file_name=f"Plan_Cessions_{TODAY_FILE}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — EXPORT EXCEL
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[5]:
    st.markdown('<div class="sh">📥 EXPORT EXCEL ENRICHI</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="info-box blue">
      L'export contient 3 feuilles : <strong>Synthèse réseau</strong> (scorecard magasins),
      <strong>Détail complet</strong> (toutes les lignes avec statut) et
      <strong>Alertes prioritaires</strong> (non implantés + stocks négatifs uniquement).
    </div>
    """, unsafe_allow_html=True)

    if st.button("🔨 Générer l'export Excel", type="primary", use_container_width=False):
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        buf_x = io.BytesIO()

        with pd.ExcelWriter(buf_x, engine="openpyxl") as writer:
            # Feuille 1 — Synthèse
            cols_synth = ["Magasin"] + [c for c in STATUT_COLORS if c in pivot_mag.columns] + ["Total", "Taux (%)"]
            pivot_mag[cols_synth].sort_values("Taux (%)", ascending=False).to_excel(
                writer, sheet_name="Synthèse Réseau", index=False
            )

            # Feuille 2 — Détail
            cols_det = ["Magasin", "SKU", "Libellé article", "Origine", "Mode Appro",
                        "Sem. Réception", "Fournisseur", "Stock", "Statut"]
            merged[[c for c in cols_det if c in merged.columns]].to_excel(
                writer, sheet_name="Détail Complet", index=False
            )

            # Feuille 3 — Alertes
            df_alerte = merged[merged["Statut"].isin(["⚠️ Non implanté", "🔴 Stock négatif"])].copy()
            df_alerte[[c for c in cols_det if c in df_alerte.columns]].sort_values(
                ["Statut", "Magasin"]
            ).to_excel(writer, sheet_name="Alertes Prioritaires", index=False)

            # Mise en forme simple
            wb = writer.book
            FILL_HEADER = PatternFill("solid", fgColor="1C1C1E")
            FONT_HEADER = Font(bold=True, color="FFFFFF", name="Arial", size=11)
            ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

            for shname in wb.sheetnames:
                ws = wb[shname]
                # En-têtes
                for cell in ws[1]:
                    cell.fill   = FILL_HEADER
                    cell.font   = FONT_HEADER
                    cell.alignment = ALIGN_CENTER
                # Largeur colonnes auto
                for col in ws.columns:
                    max_len = max(
                        (len(str(cell.value)) for cell in col if cell.value), default=10
                    )
                    ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 50)
                # Freeze row 1
                ws.freeze_panes = "A2"
                # Coloriser colonne Statut si présente
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if str(cell.value) == "✅ Implanté":
                            cell.fill = PatternFill("solid", fgColor="D1FAE5")
                            cell.font = Font(color="065F46", name="Arial", size=10)
                        elif str(cell.value) == "⚠️ Non implanté":
                            cell.fill = PatternFill("solid", fgColor="FEF3C7")
                            cell.font = Font(color="92400E", name="Arial", size=10)
                        elif str(cell.value) == "🔴 Stock négatif":
                            cell.fill = PatternFill("solid", fgColor="FEE2E2")
                            cell.font = Font(color="991B1B", name="Arial", size=10)

        buf_x.seek(0)
        st.download_button(
            label=f"📥 Implantation_T1_{TODAY_FILE}.xlsx",
            data=buf_x,
            file_name=f"Implantation_T1_{TODAY_FILE}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success(f"✅ Export généré — {fmt_n(len(merged))} lignes · 3 feuilles")

# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    f'<div style="text-align:center;font-size:11px;color:{C["muted"]};font-family:JetBrains Mono;">'
    f'SmartBuyer · NovaRetail Solutions · Implantation v3.0 · {TODAY_STR}'
    f'</div>',
    unsafe_allow_html=True
)
