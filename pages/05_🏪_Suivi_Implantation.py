"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  03_📦_Implantation.py  ·  SmartBuyer · NovaRetail Solutions               ║
║  Suivi Implantation + Supply Nouvelles Références — v4.0                    ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  ① T1 liste nouvelles références                                            ║
║  ② Stock ERP multi-magasins (code site extrait du nom de fichier)           ║
║  ③ RAL Livraisons multi-magasins (optionnel → active onglet Supply)         ║
║  Situations RAL : 38=En attente Supply · 40=En transit · 50=Réception cours ║
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
TODAY      = pd.Timestamp(date.today())
TODAY_STR  = date.today().strftime("%d %b %Y")
TODAY_FILE = date.today().strftime("%Y%m%d")

SITUATION_LABEL = {
    38: ("🕐", "En attente Supply"),
    40: ("🚚", "En transit"),
    50: ("📥", "Réception en cours"),
}

C = {
    "bg"      : "#F2F2F7",
    "surface" : "#FFFFFF",
    "border"  : "#E5E5EA",
    "text"    : "#1C1C1E",
    "muted"   : "#6D6D72",
    "blue"    : "#007AFF",
    "green"   : "#34C759",
    "red"     : "#FF3B30",
    "orange"  : "#FF9500",
    "purple"  : "#AF52DE",
    "teal"    : "#5AC8FA",
    "blue_l"  : "#EFF4FF",
    "green_l" : "#F0FFF4",
    "red_l"   : "#FFF2F0",
    "orange_l": "#FFFBEB",
    "purple_l": "#F5F0FF",
}

STATUT_COLORS = {
    "✅ Implanté"      : C["green"],
    "🔴 Stock négatif" : C["red"],
    "⚠️ Non implanté"  : C["orange"],
}

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Implantation & Supply · SmartBuyer",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CHARTE GRAPHIQUE SMARTBUYER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
:root {{
  --bg:{C['bg']}; --surface:{C['surface']}; --border:{C['border']};
  --text:{C['text']}; --muted:{C['muted']};
  --blue:{C['blue']}; --green:{C['green']}; --red:{C['red']};
  --orange:{C['orange']}; --purple:{C['purple']}; --teal:{C['teal']};
  --blue-l:{C['blue_l']}; --green-l:{C['green_l']};
  --red-l:{C['red_l']}; --orange-l:{C['orange_l']}; --purple-l:{C['purple_l']};
  --radius:14px; --shadow:0 1px 3px rgba(0,0,0,.06),0 4px 16px rgba(0,0,0,.04);
}}
html,body,[class*="css"]{{font-family:'Inter',sans-serif!important;background:var(--bg)!important;color:var(--text)!important;}}
.main,section[data-testid="stMain"]{{background:var(--bg)!important;}}
.block-container{{padding:0 2rem 4rem!important;max-width:1480px;}}
header[data-testid="stHeader"],#MainMenu,footer{{display:none!important;}}
.topbar{{background:var(--text);margin:0 -2rem 28px;padding:16px 28px;display:flex;align-items:center;justify-content:space-between;}}
.topbar-icon{{width:40px;height:40px;border-radius:10px;background:linear-gradient(135deg,{C['blue']},{C['purple']});display:flex;align-items:center;justify-content:center;font-size:22px;}}
.topbar-title{{font-size:17px;font-weight:700;color:#fff;letter-spacing:-.01em;}}
.topbar-sub{{font-size:11px;color:#8E8E93;font-family:'JetBrains Mono';margin-top:2px;}}
.topbar-pill{{background:rgba(255,255,255,.08);color:#8E8E93;border:1px solid rgba(255,255,255,.12);border-radius:8px;padding:4px 14px;font-size:11px;font-weight:600;}}
.topbar-date{{color:{C['blue']};font-size:12px;font-family:'JetBrains Mono';}}
.module-intro{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px 24px;margin-bottom:24px;box-shadow:var(--shadow);}}
.module-intro-title{{font-size:15px;font-weight:700;color:var(--text);margin-bottom:6px;}}
.module-intro-body{{font-size:13px;color:var(--muted);line-height:1.6;}}
.tag-badge{{display:inline-flex;align-items:center;gap:5px;background:var(--blue-l);color:var(--blue);border:1px solid #BFDBFE;border-radius:6px;padding:3px 10px;font-size:11px;font-weight:600;margin:2px;}}
.tag-badge.green{{background:var(--green-l);color:var(--green);border-color:#6EE7B7;}}
.tag-badge.red{{background:var(--red-l);color:var(--red);border-color:#FECACA;}}
.tag-badge.orange{{background:var(--orange-l);color:var(--orange);border-color:#FCD34D;}}
.tag-badge.purple{{background:var(--purple-l);color:var(--purple);border-color:#D8B4FE;}}
.sh{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:var(--muted);margin:24px 0 14px;padding-bottom:8px;border-bottom:1px solid var(--border);}}
.sh-supply{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:var(--purple);margin:24px 0 14px;padding-bottom:8px;border-bottom:2px solid var(--purple);}}
.partie-label{{display:inline-flex;align-items:center;gap:8px;background:var(--text);color:#fff;border-radius:8px;padding:6px 16px;font-size:12px;font-weight:700;letter-spacing:.05em;margin-bottom:16px;}}
.partie-label.supply{{background:linear-gradient(135deg,{C['purple']},{C['blue']});}}
.kpi-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:24px;}}
.kpi-card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:18px 20px 14px;box-shadow:var(--shadow);position:relative;overflow:hidden;}}
.kpi-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:var(--radius) var(--radius) 0 0;}}
.kpi-card.blue::before{{background:var(--blue);}} .kpi-card.green::before{{background:var(--green);}}
.kpi-card.red::before{{background:var(--red);}}   .kpi-card.orange::before{{background:var(--orange);}}
.kpi-card.purple::before{{background:var(--purple);}} .kpi-card.teal::before{{background:var(--teal);}}
.kpi-label{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.10em;color:var(--muted);margin-bottom:10px;}}
.kpi-value{{font-size:36px;font-weight:800;line-height:1;letter-spacing:-.02em;}}
.kpi-card.blue .kpi-value{{color:var(--blue);}} .kpi-card.green .kpi-value{{color:var(--green);}}
.kpi-card.red .kpi-value{{color:var(--red);}}   .kpi-card.orange .kpi-value{{color:var(--orange);}}
.kpi-card.purple .kpi-value{{color:var(--purple);}} .kpi-card.teal .kpi-value{{color:var(--teal);}}
.kpi-sub{{font-size:11px;color:var(--muted);font-family:'JetBrains Mono';margin-top:4px;}}
.kpi-bar{{margin-top:12px;height:3px;border-radius:3px;background:var(--border);}}
.kpi-bar-fill{{height:100%;border-radius:3px;}}
.scorecard-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:10px;margin-bottom:24px;}}
.scorecard-card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px;box-shadow:var(--shadow);position:relative;}}
.scorecard-card.ok{{border-color:#6EE7B7;background:var(--green-l);}}
.scorecard-card.warn{{border-color:#FCD34D;background:var(--orange-l);}}
.scorecard-card.ko{{border-color:#FECACA;background:var(--red-l);}}
.scorecard-dot{{width:8px;height:8px;border-radius:50%;position:absolute;top:14px;right:14px;}}
.scorecard-name{{font-size:11px;font-weight:600;color:var(--text);margin-bottom:6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:88%;}}
.scorecard-pct{{font-size:28px;font-weight:800;line-height:1;}}
.scorecard-sub{{font-size:10px;color:var(--muted);font-family:'JetBrains Mono';margin-top:3px;}}
.alert-banner{{background:#FFF;border:1px solid #FECACA;border-left:4px solid var(--red);border-radius:var(--radius);padding:16px 20px;margin-bottom:20px;display:flex;align-items:center;gap:16px;flex-wrap:wrap;}}
.alert-pill{{background:var(--red);color:#fff;border-radius:6px;padding:4px 12px;font-size:11px;font-weight:700;letter-spacing:.05em;white-space:nowrap;}}
.alert-item{{display:flex;flex-direction:column;align-items:center;padding:0 16px;border-right:1px solid var(--border);}}
.alert-item:last-child{{border-right:none;}}
.alert-num{{font-size:26px;font-weight:800;line-height:1;}}
.alert-lbl{{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-top:1px;}}
.info-box{{border-radius:var(--radius);padding:14px 18px;margin-bottom:16px;border:1px solid;font-size:13px;line-height:1.6;}}
.info-box.blue{{background:var(--blue-l);border-color:#BFDBFE;color:#1D4ED8;}}
.info-box.green{{background:var(--green-l);border-color:#6EE7B7;color:#065F46;}}
.info-box.orange{{background:var(--orange-l);border-color:#FCD34D;color:#92400E;}}
.info-box.purple{{background:var(--purple-l);border-color:#D8B4FE;color:#6B21A8;}}
.nav-active{{background:var(--text)!important;color:#fff!important;border-radius:10px;padding:10px 0;text-align:center;font-size:13px;font-weight:700;box-shadow:0 4px 14px rgba(28,28,30,.2);margin-bottom:10px;}}
.nav-active-supply{{background:linear-gradient(135deg,{C['purple']},{C['blue']})!important;color:#fff!important;border-radius:10px;padding:10px 0;text-align:center;font-size:13px;font-weight:700;box-shadow:0 4px 14px rgba(175,82,222,.3);margin-bottom:10px;}}
.supply-divider{{background:linear-gradient(135deg,{C['purple']}22,{C['blue']}22);border:1px solid {C['purple']}44;border-radius:var(--radius);padding:16px 20px;margin:32px 0 20px;display:flex;align-items:center;gap:12px;}}
.supply-divider-icon{{width:36px;height:36px;border-radius:9px;background:linear-gradient(135deg,{C['purple']},{C['blue']});display:flex;align-items:center;justify-content:center;font-size:18px;}}
.supply-divider-title{{font-size:15px;font-weight:700;color:var(--text);}}
.supply-divider-sub{{font-size:12px;color:var(--muted);margin-top:2px;}}
section[data-testid="stSidebar"]{{background:#fff!important;border-right:1px solid var(--border)!important;min-width:270px!important;max-width:270px!important;}}
section[data-testid="stSidebar"] .block-container{{padding:.6rem .8rem 2rem!important;}}
.stDownloadButton>button{{background:linear-gradient(135deg,{C['text']},{C['blue']})!important;color:#fff!important;border:none!important;border-radius:10px!important;font-weight:700!important;font-size:13px!important;padding:12px!important;width:100%!important;box-shadow:0 4px 12px rgba(0,122,255,.25)!important;}}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def fmt_n(n) -> str:
    try:
        return f"{int(n):,}".replace(",", "\u202f")
    except Exception:
        return str(n)

def color_taux(t: float) -> str:
    if t >= 80: return C["green"]
    if t >= 50: return C["orange"]
    return C["red"]

def scorecard_cls(t: float) -> str:
    if t >= 80: return "ok"
    if t >= 50: return "warn"
    return "ko"

def extract_site_code(filename: str) -> str:
    m = re.search(r'(1\d{4}|2\d{4})', filename)
    return m.group(1) if m else None

def extract_date_from_filename(filename: str) -> pd.Timestamp:
    m = re.search(r'(\d{8})', filename)
    if not m:
        return TODAY
    s = m.group(1)
    for fmt in ('%Y%m%d', '%d%m%Y'):
        try:
            return pd.Timestamp(pd.to_datetime(s, format=fmt))
        except Exception:
            continue
    return TODAY

def retard_badge(jours: int) -> str:
    if jours <= 0:  return "✅ À temps"
    if jours <= 7:  return f"🟡 {jours}j"
    if jours <= 30: return f"🟠 {jours}j"
    return f"🔴 {jours}j critique"

def action_supply(stock, ral_actif, situation_principale) -> str:
    stock = 0 if pd.isna(stock) else float(stock)
    ral   = 0 if pd.isna(ral_actif) else float(ral_actif)
    if stock > 0 and ral > 0: return "✅ Implanté + réassort en cours"
    if stock > 0:              return "✅ Implanté"
    if stock < 0 and ral > 0: return "🔧 Régulariser + livraison en cours"
    if stock < 0:              return "🔧 Régulariser inventaire"
    if ral > 0:
        sit = int(situation_principale) if pd.notna(situation_principale) else 38
        if sit == 50: return "📥 Réception en cours"
        if sit == 40: return "🚚 En transit"
        return "🕐 En attente livraison"
    return "🛒 Passer commande"

def _sem_to_num(s: str) -> int:
    cleaned = re.sub(r"[Ss]", "", str(s).strip())
    return int(cleaned) if cleaned.isdigit() else 99


# ─────────────────────────────────────────────────────────────────────────────
# PARSERS
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_t1(file_bytes: bytes, filename: str):
    buf = io.BytesIO(file_bytes)
    try:
        df = pd.read_excel(buf, header=None, dtype=str) if filename.lower().endswith((".xlsx", ".xls")) \
             else pd.read_csv(buf, header=None, sep=None, engine="python", encoding="latin1",
                              dtype=str, on_bad_lines="skip")
    except Exception as e:
        return None, f"Lecture T1 : {e}"

    first = str(df.iloc[0, 0]).strip().replace(".0", "")
    has_header = not first.isdigit()
    if has_header:
        df.columns = df.iloc[0].astype(str).str.strip().str.upper()
        df = df.iloc[1:].reset_index(drop=True)
    else:
        df.columns = ["ARTICLE"] + [f"_COL{i}" for i in range(1, len(df.columns))]

    df.columns = (df.columns.astype(str).str.strip().str.upper()
                  .str.replace("\ufeff", "", regex=False).str.replace("\xa0", " ", regex=False))

    if "ARTICLE" not in df.columns:
        return None, "Colonne ARTICLE introuvable dans le fichier T1"

    df["SKU"] = (df["ARTICLE"].astype(str).str.strip()
                 .str.replace(r"\.0$", "", regex=True).str.zfill(8).str[:8])
    df = df[df["SKU"].str.match(r"^\d{8}$", na=False)].drop_duplicates("SKU").copy()

    defaults = {"LIBELLÉ ARTICLE": "", "LIBELLÉ FOURNISSEUR ORIGINE": "",
                "MODE APPRO": "", "SEMAINE RECEPTION": "", "DATE LIV.": ""}
    for col, val in defaults.items():
        if col not in df.columns:
            df[col] = val

    df["SEMAINE RECEPTION"] = df["SEMAINE RECEPTION"].astype(str).str.strip().replace("nan", "")
    df["SEM_NUM"]  = df["SEMAINE RECEPTION"].apply(_sem_to_num)
    df["ORIGINE"]  = df["MODE APPRO"].apply(lambda m: "IM" if "IMPORT" in str(m).upper() else "LO")
    return df, None


@st.cache_data(show_spinner=False)
def parse_stock_erp(files_bytes: list, filenames: list):
    frames, errors = [], []
    for fb, fn in zip(files_bytes, filenames):
        code_site = extract_site_code(fn)
        if not code_site:
            errors.append(f"⚠️ Code magasin introuvable dans '{fn}' — ignoré")
            continue
        try:
            df = pd.read_csv(io.BytesIO(fb), sep=";", encoding="latin1",
                             low_memory=False, on_bad_lines="skip")
        except Exception as e:
            errors.append(f"❌ Lecture '{fn}' : {e}")
            continue

        for col in df.select_dtypes(include="object").columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r"\t+", "", regex=True)

        if "Code article" not in df.columns:
            errors.append(f"⚠️ Colonne 'Code article' absente dans '{fn}'")
            continue

        df["Code site"] = code_site
        df["SKU"] = (df["Code article"].astype(str).str.strip()
                     .str.replace(r"\.0$", "", regex=True).str.zfill(8).str[:8])
        df = df[df["SKU"].str.match(r"^\d{8}$", na=False)].copy()
        frames.append(df)

    if not frames:
        return None, errors or ["Aucun fichier stock valide"]
    return pd.concat(frames, ignore_index=True), errors


@st.cache_data(show_spinner=False)
def parse_ral(files_bytes: list, filenames: list):
    frames, errors = [], []
    for fb, fn in zip(files_bytes, filenames):
        code_site       = extract_site_code(fn)
        date_extraction = extract_date_from_filename(fn)
        if not code_site:
            errors.append(f"⚠️ Code magasin introuvable dans '{fn}' — ignoré")
            continue
        try:
            df = pd.read_csv(io.BytesIO(fb), sep=";", encoding="latin1",
                             low_memory=False, on_bad_lines="skip")
        except Exception as e:
            errors.append(f"❌ Lecture '{fn}' : {e}")
            continue

        for col in df.select_dtypes(include="object").columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r"\t+", "", regex=True)

        if "Code art." not in df.columns:
            errors.append(f"⚠️ Colonne 'Code art.' absente dans '{fn}'")
            continue

        df["Code site"]       = code_site
        df["Date extraction"] = date_extraction
        df["SKU"] = (df["Code art."].astype(str).str.strip()
                     .str.replace(r"\.0$", "", regex=True).str.zfill(8).str[:8])
        df = df[df["SKU"].str.match(r"^\d{8}$", na=False)].copy()

        df["Situation"] = pd.to_numeric(df["Situation"], errors="coerce")
        df = df[df["Situation"].isin([38, 40, 50])].copy()

        df["Date Reception"] = pd.to_datetime(df["Date Reception"], format="%d/%m/%Y", errors="coerce")
        df["En retard"]      = df["Date Reception"] < date_extraction
        df["Jours retard"]   = (date_extraction - df["Date Reception"]).dt.days.clip(lower=0)
        df["Jours retard"]   = df["Jours retard"].where(df["En retard"], 0)
        df["RAL"]            = pd.to_numeric(df["RAL"], errors="coerce").fillna(0)
        frames.append(df)

    if not frames:
        return None, errors or ["Aucun fichier RAL valide"]

    raw = pd.concat(frames, ignore_index=True)
    agg = (raw.groupby(["SKU", "Code site"]).agg(
        RAL_actif            =("RAL", "sum"),
        Nb_commandes         =("Commande", "nunique"),
        Situation_principale =("Situation", lambda x: x.value_counts().index[0]),
        Prochaine_livraison  =("Date Reception", "min"),
        Commandes_en_retard  =("En retard", "sum"),
        Jours_retard_max     =("Jours retard", "max"),
    ).reset_index())
    agg["Prochaine_livraison"] = agg["Prochaine_livraison"].dt.strftime("%d/%m/%Y")
    return agg, errors


# ─────────────────────────────────────────────────────────────────────────────
# TOPBAR
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="topbar">
  <div style="display:flex;align-items:center;gap:14px;">
    <div class="topbar-icon">📦</div>
    <div>
      <div class="topbar-title">Suivi Implantation & Supply · Nouvelles Références</div>
      <div class="topbar-sub">T1 · Stock ERP · RAL Livraisons · Multi-magasins</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:12px;">
    <div class="topbar-date">{TODAY_STR}</div>
    <div class="topbar-pill">v4.0 · SmartBuyer</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# INTRO MÉTIER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="module-intro">
  <div class="module-intro-title">📦 À quoi sert ce module ?</div>
  <div class="module-intro-body">
    Ce module suit l'implantation des <strong>nouvelles références T1</strong> dans le réseau Carrefour CI,
    depuis la commande fournisseur jusqu'à la mise en rayon. Il croise la liste T1, le stock ERP par magasin
    et les commandes en attente (RAL) pour donner une <strong>vision complète en temps réel</strong> :
    où en est chaque article, dans chaque magasin, et quelle action prendre.<br><br>
    <span class="tag-badge">📋 Taux d'implantation réseau</span>
    <span class="tag-badge green">✅ Suivi mise en rayon</span>
    <span class="tag-badge orange">🛒 Articles à commander</span>
    <span class="tag-badge purple">🚚 Suivi livraisons & retards</span>
    <span class="tag-badge red">🔧 Régularisations inventaire</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — CHARGEMENT
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📁 Fichiers")
    st.divider()
    st.markdown("**① T1 — Nouvelles Références**")
    st.caption("Liste articles à implanter (ERP)")
    t1_file = st.file_uploader("T1", type=["csv", "xlsx", "xls"],
                               key="t1", label_visibility="collapsed")
    st.markdown("**② Stock ERP** *(multi-magasins)*")
    st.caption("Un fichier par magasin · code site extrait automatiquement")
    stk_files = st.file_uploader("Stock ERP", type=["csv"], key="stk",
                                 label_visibility="collapsed", accept_multiple_files=True)
    st.markdown("**③ RAL Livraisons** *(optionnel · multi-magasins)*")
    st.caption("Active l'onglet Supply · un fichier par magasin")
    ral_files = st.file_uploader("RAL", type=["csv"], key="ral",
                                 label_visibility="collapsed", accept_multiple_files=True)

# ─────────────────────────────────────────────────────────────────────────────
# GATES CHARGEMENT
# ─────────────────────────────────────────────────────────────────────────────
if not t1_file:
    st.markdown('<div class="info-box blue">⬆️ <strong>Étape 1</strong> — Charge le fichier T1 dans la sidebar.</div>', unsafe_allow_html=True)
    st.stop()

with st.spinner("Lecture T1…"):
    t1_df, t1_err = parse_t1(t1_file.read(), t1_file.name)

if t1_err or t1_df is None:
    st.error(f"❌ T1 : {t1_err}")
    st.stop()

if not stk_files:
    st.markdown(f'<div class="info-box blue">✅ T1 chargé — <strong>{len(t1_df):,}</strong> références. ⬆️ <strong>Étape 2</strong> — Charge les fichiers Stock ERP.</div>', unsafe_allow_html=True)
    st.stop()

with st.spinner(f"Parsing Stock ERP ({len(stk_files)} fichier(s))…"):
    stk_bytes = [f.read() for f in stk_files]
    stk_names = [f.name for f in stk_files]
    df_stock, stk_errors = parse_stock_erp(stk_bytes, stk_names)

for e in stk_errors:
    st.warning(e)

if df_stock is None:
    st.error("❌ Aucun fichier stock valide chargé.")
    st.stop()

# RAL optionnel
df_ral_agg, ral_errors, supply_active = None, [], False
if ral_files:
    with st.spinner(f"Parsing RAL ({len(ral_files)} fichier(s))…"):
        ral_bytes = [f.read() for f in ral_files]
        ral_names = [f.name for f in ral_files]
        df_ral_agg, ral_errors = parse_ral(ral_bytes, ral_names)
    for e in ral_errors:
        st.warning(e)
    if df_ral_agg is not None:
        supply_active = True

# ─────────────────────────────────────────────────────────────────────────────
# RÉFÉRENTIEL MAGASINS
# ─────────────────────────────────────────────────────────────────────────────
if "Libellé site" in df_stock.columns:
    site_ref = (df_stock[["Code site", "Libellé site"]]
                .drop_duplicates("Code site")
                .set_index("Code site")["Libellé site"].to_dict())
else:
    site_ref = {c: c for c in df_stock["Code site"].unique()}

all_codes = sorted(df_stock["Code site"].unique().tolist())

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — FILTRES
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.divider()
    st.markdown("## 🔍 Filtres")
    mag_labels    = sorted([site_ref.get(c, c) for c in all_codes])
    mag_sel_labels = st.multiselect("Magasins", mag_labels, default=mag_labels)
    mag_sel_codes  = [c for c in all_codes if site_ref.get(c, c) in mag_sel_labels]
    orig_sel       = st.multiselect("Flux", ["IM", "LO"], default=["IM", "LO"])

    def _sem_sort(s):
        cleaned = re.sub(r"[Ss]", "", str(s).strip())
        return int(cleaned) if cleaned.isdigit() else 99

    sem_dispo = sorted(
        [s for s in t1_df["SEMAINE RECEPTION"].unique()
         if str(s).strip() not in ("nan", "", "99")],
        key=_sem_sort
    )
    sem_sel = st.multiselect("Semaine réception", sem_dispo, default=sem_dispo)

    rayons_dispo = []
    if "Libellé rayon" in df_stock.columns:
        rayons_dispo = sorted([r for r in df_stock["Libellé rayon"].dropna().unique()
                               if str(r).strip() not in ("nan", "")])
    rayon_sel = st.multiselect("Rayon", rayons_dispo, default=rayons_dispo) if rayons_dispo else []

    st.divider()
    st.markdown("## 🔄 Cessions")
    mag_detresse = st.multiselect("Magasins en détresse", mag_labels, default=[])
    seuil_det    = st.number_input("Seuil stock (≤)", 0, 50, 0, 1)
    reserve      = st.number_input("Réserve cédant (≥)", 0, 50, 2, 1)

if not mag_sel_codes:
    st.warning("⚠️ Sélectionne au moins un magasin.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# DATASET IMPLANTATION
# ─────────────────────────────────────────────────────────────────────────────
mask = t1_df["ORIGINE"].isin(orig_sel)
if sem_sel:
    mask = mask & t1_df["SEMAINE RECEPTION"].isin(sem_sel)
t1_scope  = t1_df[mask].copy()
sku_scope = t1_scope["SKU"].unique()

if len(sku_scope) == 0:
    st.warning("Aucun SKU correspondant aux filtres.")
    st.stop()

stk_filt = df_stock[df_stock["Code site"].isin(mag_sel_codes) &
                    df_stock["SKU"].isin(sku_scope)].copy()

if rayon_sel and "Libellé rayon" in stk_filt.columns:
    stk_filt = stk_filt[stk_filt["Libellé rayon"].isin(rayon_sel)]

stk_cols = ["SKU", "Code site", "Nouveau stock"]
for opt in ["Libellé article", "Nom fourn.", "Libellé rayon", "Code etat", "Date dernière entrée"]:
    if opt in stk_filt.columns:
        stk_cols.append(opt)
stk_filt = stk_filt[stk_cols].copy()
stk_filt["Nouveau stock"] = pd.to_numeric(stk_filt["Nouveau stock"], errors="coerce")

grid   = pd.DataFrame(
    pd.MultiIndex.from_product([mag_sel_codes, sku_scope], names=["Code site", "SKU"]).tolist(),
    columns=["Code site", "SKU"]
)
merged = grid.merge(stk_filt, on=["Code site", "SKU"], how="left")

t1_ref = t1_scope.set_index("SKU")[[
    "LIBELLÉ ARTICLE", "LIBELLÉ FOURNISSEUR ORIGINE",
    "MODE APPRO", "SEMAINE RECEPTION", "DATE LIV.", "ORIGINE", "SEM_NUM"
]].rename(columns={
    "LIBELLÉ ARTICLE"             : "T1_lib",
    "LIBELLÉ FOURNISSEUR ORIGINE" : "Fournisseur T1",
    "MODE APPRO"                  : "Mode Appro",
    "SEMAINE RECEPTION"           : "Sem. Réception",
    "DATE LIV."                   : "Date Livraison",
    "ORIGINE"                     : "Origine",
    "SEM_NUM"                     : "SEM_NUM",
})
merged = merged.merge(t1_ref.reset_index(), on="SKU", how="left")

if "Libellé article" not in merged.columns:
    merged["Libellé article"] = ""
merged["Libellé article"] = merged["Libellé article"].fillna("").astype(str)
merged["Libellé article"] = merged.apply(
    lambda r: r["Libellé article"] if r["Libellé article"] else r.get("T1_lib", ""), axis=1
)
merged.drop(columns=["T1_lib"], errors="ignore", inplace=True)
merged["Magasin"] = merged["Code site"].map(site_ref).fillna(merged["Code site"])

def get_statut(s):
    if pd.isna(s): return "⚠️ Non implanté"
    v = float(s)
    if v < 0: return "🔴 Stock négatif"
    if v > 0: return "✅ Implanté"
    return "⚠️ Non implanté"

merged["Statut"] = merged["Nouveau stock"].apply(get_statut)

# ─────────────────────────────────────────────────────────────────────────────
# MÉTRIQUES
# ─────────────────────────────────────────────────────────────────────────────
n_sku       = len(sku_scope)
n_mag       = len(mag_sel_codes)
total_cells = len(merged)
n_impl      = int((merged["Statut"] == "✅ Implanté").sum())
n_non_impl  = int((merged["Statut"] == "⚠️ Non implanté").sum())
n_neg       = int((merged["Statut"] == "🔴 Stock négatif").sum())
taux_reseau = int(n_impl / total_cells * 100) if total_cells else 0
n_sku_im    = int((t1_scope["ORIGINE"] == "IM").sum())
n_sku_lo    = int((t1_scope["ORIGINE"] == "LO").sum())
pct         = lambda n: int(n / total_cells * 100) if total_cells else 0

pivot_mag = (merged.groupby(["Magasin", "Statut"]).size().unstack(fill_value=0)
             .reindex(columns=list(STATUT_COLORS.keys()), fill_value=0).reset_index())
pivot_mag.columns.name = None
pivot_mag["Total"]    = n_sku
pivot_mag["Taux (%)"] = (pivot_mag.get("✅ Implanté", 0) / n_sku * 100).round(0).astype(int)

# ─────────────────────────────────────────────────────────────────────────────
# BANNIÈRE + KPIs IMPLANTATION
# ─────────────────────────────────────────────────────────────────────────────
if n_non_impl + n_neg > 0:
    st.markdown(f"""
    <div class="alert-banner">
      <div class="alert-pill">⚡ ACTIONS REQUISES</div>
      <div class="alert-item"><div class="alert-num" style="color:{C['orange']}">{fmt_n(n_non_impl)}</div><div class="alert-lbl">Non implanté</div></div>
      <div class="alert-item"><div class="alert-num" style="color:{C['red']}">{fmt_n(n_neg)}</div><div class="alert-lbl">Stock négatif</div></div>
      <div style="margin-left:auto;font-size:12px;color:{C['muted']};">{n_mag} magasin(s) · {fmt_n(n_sku)} SKUs · {fmt_n(total_cells)} combinaisons</div>
    </div>""", unsafe_allow_html=True)

st.markdown('<div class="partie-label">📦 PARTIE 1 — IMPLANTATION</div>', unsafe_allow_html=True)
st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card green"><div class="kpi-label">✅ Implanté</div><div class="kpi-value">{fmt_n(n_impl)}</div><div class="kpi-sub">{pct(n_impl)}% du réseau</div><div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(n_impl)}%;background:{C['green']}"></div></div></div>
  <div class="kpi-card orange"><div class="kpi-label">⚠️ Non implanté</div><div class="kpi-value">{fmt_n(n_non_impl)}</div><div class="kpi-sub">{pct(n_non_impl)}% — à traiter</div><div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(n_non_impl)}%;background:{C['orange']}"></div></div></div>
  <div class="kpi-card red"><div class="kpi-label">🔴 Stock négatif</div><div class="kpi-value">{fmt_n(n_neg)}</div><div class="kpi-sub">{pct(n_neg)}% — écart inventaire</div><div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(n_neg)}%;background:{C['red']}"></div></div></div>
  <div class="kpi-card blue"><div class="kpi-label">📊 Taux réseau</div><div class="kpi-value" style="color:{color_taux(taux_reseau)}">{taux_reseau}%</div><div class="kpi-sub">{n_mag} mag × {fmt_n(n_sku)} SKU</div><div class="kpi-bar"><div class="kpi-bar-fill" style="width:{taux_reseau}%;background:{C['blue']}"></div></div></div>
  <div class="kpi-card purple"><div class="kpi-label">🔀 Flux IM / LO</div><div class="kpi-value" style="font-size:28px">{n_sku_im}<span style="font-size:16px;color:{C['muted']}"> / </span>{n_sku_lo}</div><div class="kpi-sub">Import · Local</div></div>
</div>""", unsafe_allow_html=True)

# SCORECARD
st.markdown('<div class="sh">SCORECARD MAGASINS</div>', unsafe_allow_html=True)
rag_html = '<div class="scorecard-grid">'
for _, row in pivot_mag.sort_values("Taux (%)", ascending=False).iterrows():
    t_  = row["Taux (%)"]
    col = color_taux(t_)
    rag_html += f"""<div class="scorecard-card {scorecard_cls(t_)}">
      <div class="scorecard-dot" style="background:{col}"></div>
      <div class="scorecard-name">{row['Magasin']}</div>
      <div class="scorecard-pct" style="color:{col}">{t_}%</div>
      <div class="scorecard-sub">{int(row.get('✅ Implanté',0))}✅ {int(row.get('⚠️ Non implanté',0))}⚠️ {int(row.get('🔴 Stock négatif',0))}🔴</div>
    </div>"""
rag_html += "</div>"
st.markdown(rag_html, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# ONGLETS IMPLANTATION
# ─────────────────────────────────────────────────────────────────────────────
TABS_IMPL = ["📊 Vue Réseau", "⚠️ Non Implantés", "🔴 Stocks Négatifs",
             "🗓️ Calendrier", "🔄 Cessions", "📥 Export"]
if "impl_tab" not in st.session_state:
    st.session_state.impl_tab = TABS_IMPL[0]

nav_cols = st.columns(len(TABS_IMPL))
for i, t in enumerate(TABS_IMPL):
    with nav_cols[i]:
        if st.session_state.impl_tab == t:
            st.markdown(f'<div class="nav-active">{t}</div>', unsafe_allow_html=True)
        if st.button(t, key=f"impl_nav_{i}", use_container_width=True):
            st.session_state.impl_tab = t
            st.rerun()

active_impl = st.session_state.impl_tab

# ── TAB 1 — VUE RÉSEAU ───────────────────────────────────────────────────────
if active_impl == TABS_IMPL[0]:
    c1, c2 = st.columns([3, 2])
    with c1:
        mel = pivot_mag.melt(id_vars="Magasin",
                             value_vars=[s for s in STATUT_COLORS if s in pivot_mag.columns],
                             var_name="Statut", value_name="N")
        fig = px.bar(mel, x="Magasin", y="N", color="Statut",
                     color_discrete_map=STATUT_COLORS, barmode="stack", title="Situation par magasin")
        fig.update_traces(textposition="inside", texttemplate="%{y}", textfont=dict(size=11, color="white"))
        fig.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=400,
                          font=dict(family="Inter", color=C["muted"], size=12),
                          margin=dict(l=10, r=10, t=44, b=20),
                          legend=dict(orientation="h", y=-0.28, bgcolor="rgba(0,0,0,0)"),
                          xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig_d = go.Figure(go.Pie(
            labels=["✅ Implanté", "⚠️ Non implanté", "🔴 Stock négatif"],
            values=[n_impl, n_non_impl, n_neg], hole=0.62,
            marker=dict(colors=[C["green"], C["orange"], C["red"]], line=dict(color="#fff", width=3))
        ))
        fig_d.add_annotation(text=f"<b>{taux_reseau}%</b><br>implanté", x=0.5, y=0.5,
                             showarrow=False, font=dict(size=20, color=C["text"], family="Inter"))
        fig_d.update_layout(paper_bgcolor=C["surface"], height=400,
                            font=dict(family="Inter", color=C["muted"], size=12),
                            margin=dict(l=10, r=10, t=44, b=20),
                            legend=dict(orientation="v", x=1.01), title="Répartition réseau")
        st.plotly_chart(fig_d, use_container_width=True)

    st.markdown('<div class="sh">SYNTHÈSE PAR MAGASIN</div>', unsafe_allow_html=True)
    cols_d = ["Magasin"] + [c for c in STATUT_COLORS if c in pivot_mag.columns] + ["Total", "Taux (%)"]
    st.dataframe(pivot_mag[cols_d].sort_values("Taux (%)", ascending=False).reset_index(drop=True)
                 .style
                 .background_gradient(subset=["✅ Implanté"] if "✅ Implanté" in pivot_mag.columns else [], cmap="Greens")
                 .background_gradient(subset=["⚠️ Non implanté"] if "⚠️ Non implanté" in pivot_mag.columns else [], cmap="Oranges")
                 .background_gradient(subset=["🔴 Stock négatif"] if "🔴 Stock négatif" in pivot_mag.columns else [], cmap="Reds")
                 .format({"Taux (%)": "{}%"}),
                 use_container_width=True, hide_index=True)

    df_flux = merged[merged["Statut"] == "✅ Implanté"].groupby(["Magasin", "Origine"]).size().reset_index(name="N")
    if not df_flux.empty:
        st.markdown('<div class="sh">FLUX IM / LO — ARTICLES IMPLANTÉS</div>', unsafe_allow_html=True)
        fig_flux = px.bar(df_flux, x="Magasin", y="N", color="Origine",
                          color_discrete_map={"IM": C["blue"], "LO": C["green"]},
                          barmode="group", title="Articles implantés par flux")
        fig_flux.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=300,
                               font=dict(family="Inter", color=C["muted"], size=12),
                               margin=dict(l=10, r=10, t=44, b=20),
                               legend=dict(orientation="h", y=-0.3),
                               xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]))
        st.plotly_chart(fig_flux, use_container_width=True)

# ── TAB 2 — NON IMPLANTÉS + RUPTURES TOTALES ─────────────────────────────────
elif active_impl == TABS_IMPL[1]:
    df_ni = merged[merged["Statut"] == "⚠️ Non implanté"].copy()

    sku_sans_stock  = (merged[merged["Statut"] != "✅ Implanté"]
                       .groupby("SKU")["Magasin"].count())
    sku_rupture_tot = sku_sans_stock[sku_sans_stock == n_mag].index.tolist()
    n_rc_im = int(t1_scope[t1_scope["SKU"].isin(sku_rupture_tot) & (t1_scope["ORIGINE"] == "IM")]["SKU"].nunique())
    n_rc_lo = int(t1_scope[t1_scope["SKU"].isin(sku_rupture_tot) & (t1_scope["ORIGINE"] == "LO")]["SKU"].nunique())

    if sku_rupture_tot:
        st.markdown(f"""
        <div class="alert-banner">
          <div class="alert-pill">🚨 RUPTURE TOTALE RÉSEAU</div>
          <div class="alert-item"><div class="alert-num" style="color:{C['red']}">{len(sku_rupture_tot)}</div><div class="alert-lbl">SKUs absents partout</div></div>
          <div class="alert-item"><div class="alert-num" style="color:{C['blue']}">{n_rc_im}</div><div class="alert-lbl">Flux IM</div></div>
          <div class="alert-item"><div class="alert-num" style="color:{C['green']}">{n_rc_lo}</div><div class="alert-lbl">Flux LO</div></div>
          <div style="margin-left:auto;font-size:12px;color:{C['muted']};">Aucun stock positif sur les {n_mag} magasin(s) — escalade critique</div>
        </div>""", unsafe_allow_html=True)

    sub1, sub2 = st.tabs([f"⚠️ Non implantés ({len(df_ni)})",
                          f"🔴 Ruptures totales ({len(sku_rupture_tot)} SKU)"])
    with sub1:
        if df_ni.empty:
            st.markdown('<div class="info-box green">✅ Tous les articles sont implantés !</div>', unsafe_allow_html=True)
        else:
            k1, k2, k3 = st.columns(3)
            k1.metric("⚠️ Lignes manquantes", fmt_n(len(df_ni)))
            k2.metric("Articles distincts",   fmt_n(df_ni["SKU"].nunique()))
            k3.metric("dont Rupture totale",  len(sku_rupture_tot))
            df_ni["Rupture totale"] = df_ni["SKU"].isin(sku_rupture_tot).map({True: "🔴 OUI", False: "—"})
            COLS = ["Magasin", "SKU", "Libellé article", "Origine", "Mode Appro",
                    "Sem. Réception", "Fournisseur T1", "Rupture totale"]
            st.dataframe(df_ni[[c for c in COLS if c in df_ni.columns]]
                         .sort_values(["Rupture totale", "Magasin"]).reset_index(drop=True),
                         use_container_width=True, hide_index=True)

    with sub2:
        if not sku_rupture_tot:
            st.markdown('<div class="info-box green">✅ Aucune rupture totale détectée.</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="info-box orange">📌 <strong>Rupture totale</strong> = aucun stock positif sur l\'ensemble des {n_mag} magasin(s). Action : escalade fournisseur + vérification commande ERP.</div>', unsafe_allow_html=True)
            k1, k2 = st.columns(2)
            k1.metric("SKUs en rupture totale", len(sku_rupture_tot))
            k2.metric("IM / LO", f"{n_rc_im} / {n_rc_lo}")
            df_rc = (merged[merged["SKU"].isin(sku_rupture_tot)]
                     [["SKU", "Libellé article", "Origine", "Fournisseur T1", "Mode Appro", "Sem. Réception"]]
                     .drop_duplicates("SKU").sort_values("Origine").reset_index(drop=True))
            st.dataframe(df_rc, use_container_width=True, hide_index=True)

# ── TAB 3 — STOCKS NÉGATIFS ──────────────────────────────────────────────────
elif active_impl == TABS_IMPL[2]:
    df_neg = merged[merged["Statut"] == "🔴 Stock négatif"].copy()
    if df_neg.empty:
        st.markdown('<div class="info-box green">✅ Aucun stock négatif détecté.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="info-box orange">📌 <strong>Stock négatif = écart d\'inventaire</strong> — articles sortis sans réception ERP. Action : régularisation inventaire magasin.</div>', unsafe_allow_html=True)
        k1, k2 = st.columns(2)
        k1.metric("🔴 Lignes stock négatif", fmt_n(len(df_neg)))
        k2.metric("Articles distincts",      fmt_n(df_neg["SKU"].nunique()))
        COLS = ["Magasin", "SKU", "Libellé article", "Origine", "Nouveau stock", "Mode Appro", "Fournisseur T1"]
        st.dataframe(df_neg[[c for c in COLS if c in df_neg.columns]]
                     .sort_values("Nouveau stock").reset_index(drop=True),
                     use_container_width=True, hide_index=True)
        fig_neg = px.histogram(df_neg, x="Nouveau stock", nbins=40,
                               title="Distribution des stocks négatifs",
                               color_discrete_sequence=[C["red"]])
        fig_neg.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=260,
                              font=dict(family="Inter", color=C["muted"], size=12),
                              margin=dict(l=10, r=10, t=44, b=20))
        st.plotly_chart(fig_neg, use_container_width=True)

# ── TAB 4 — CALENDRIER ───────────────────────────────────────────────────────
elif active_impl == TABS_IMPL[3]:
    cal_df = merged[merged["Sem. Réception"].str.match(r"^[Ss]?\d+$", na=False)].copy()
    if cal_df.empty:
        st.info("Aucune semaine de réception renseignée dans le T1.")
    else:
        cal_agg = cal_df.groupby("Sem. Réception").agg(
            Implanté     =("Statut", lambda x: (x == "✅ Implanté").sum()),
            Non_implanté =("Statut", lambda x: (x == "⚠️ Non implanté").sum()),
            Stock_négatif=("Statut", lambda x: (x == "🔴 Stock négatif").sum()),
            SKU_distincts=("SKU", "nunique"),
        ).reset_index().rename(columns={"Non_implanté": "Non implanté", "Stock_négatif": "Stock négatif"})
        cal_agg["Taux (%)"] = (cal_agg["Implanté"] /
                               (cal_agg["Implanté"] + cal_agg["Non implanté"] + cal_agg["Stock négatif"]) * 100
                               ).round(0).astype(int)
        fig_cal = px.bar(
            cal_agg.melt(id_vars="Sem. Réception",
                         value_vars=["Implanté", "Non implanté", "Stock négatif"]),
            x="Sem. Réception", y="value", color="variable",
            color_discrete_map={"Implanté": C["green"], "Non implanté": C["orange"], "Stock négatif": C["red"]},
            barmode="stack", title="Statut par semaine de réception"
        )
        fig_cal.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=360,
                              font=dict(family="Inter", color=C["muted"], size=12),
                              margin=dict(l=10, r=10, t=44, b=20),
                              legend=dict(orientation="h", y=-0.25),
                              xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]))
        st.plotly_chart(fig_cal, use_container_width=True)
        st.dataframe(cal_agg, use_container_width=True, hide_index=True)

# ── TAB 5 — CESSIONS ─────────────────────────────────────────────────────────
elif active_impl == TABS_IMPL[4]:
    st.markdown('<div class="sh">🔄 MOTEUR CESSIONS INTER-MAGASINS</div>', unsafe_allow_html=True)
    if not mag_detresse:
        st.markdown('<div class="info-box blue">⬅️ Sélectionne des magasins en détresse dans la sidebar.</div>', unsafe_allow_html=True)
    else:
        mag_det_codes = [c for c in all_codes if site_ref.get(c, c) in mag_detresse]
        mag_ced_codes = [c for c in mag_sel_codes if c not in mag_det_codes]
        suggestions   = []
        for sku in sku_scope:
            sku_df = df_stock[df_stock["SKU"] == sku].copy()
            if sku_df.empty: continue
            lib = sku_df["Libellé article"].iloc[0] if "Libellé article" in sku_df.columns else sku
            sku_df["Nouveau stock"] = pd.to_numeric(sku_df["Nouveau stock"], errors="coerce").fillna(0)
            det_rows = sku_df[sku_df["Code site"].isin(mag_det_codes) & (sku_df["Nouveau stock"] <= seuil_det)]
            if det_rows.empty: continue
            ced_rows = sku_df[sku_df["Code site"].isin(mag_ced_codes) & (sku_df["Nouveau stock"] > reserve)].sort_values("Nouveau stock", ascending=False)
            for _, dr in det_rows.iterrows():
                if ced_rows.empty:
                    suggestions.append({"SKU": sku, "Libellé article": lib,
                        "Magasin détresse": site_ref.get(dr["Code site"], dr["Code site"]),
                        "Stock détresse": int(dr["Nouveau stock"]), "Cédant suggéré": "⚠️ Aucun",
                        "Stock cédant": 0, "Qté cessible": 0, "Faisabilité": "🔴 Impossible"})
                else:
                    best = ced_rows.iloc[0]
                    qty  = int(best["Nouveau stock"]) - reserve
                    suggestions.append({"SKU": sku, "Libellé article": lib,
                        "Magasin détresse": site_ref.get(dr["Code site"], dr["Code site"]),
                        "Stock détresse": int(dr["Nouveau stock"]),
                        "Cédant suggéré": site_ref.get(best["Code site"], best["Code site"]),
                        "Stock cédant": int(best["Nouveau stock"]), "Qté cessible": qty,
                        "Faisabilité": "🟢 Possible" if qty >= 1 else "🟠 Partielle"})
        if not suggestions:
            st.success("✅ Aucune cession nécessaire selon les critères.")
        else:
            df_cess = pd.DataFrame(suggestions).sort_values(["Faisabilité", "Qté cessible"],
                                                            ascending=[True, False]).reset_index(drop=True)
            k1, k2, k3 = st.columns(3)
            k1.metric("🟢 Possible",   int((df_cess["Faisabilité"] == "🟢 Possible").sum()))
            k2.metric("🔴 Impossible", int((df_cess["Faisabilité"] == "🔴 Impossible").sum()))
            k3.metric("Articles",      df_cess["SKU"].nunique())
            st.dataframe(df_cess, use_container_width=True, hide_index=True)
            buf_c = io.BytesIO()
            with pd.ExcelWriter(buf_c, engine="openpyxl") as w:
                df_cess.to_excel(w, sheet_name="Plan Cessions", index=False)
            buf_c.seek(0)
            st.download_button(f"📥 Plan_Cessions_{TODAY_FILE}.xlsx", data=buf_c,
                               file_name=f"Plan_Cessions_{TODAY_FILE}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── TAB 6 — EXPORT IMPLANTATION ──────────────────────────────────────────────
elif active_impl == TABS_IMPL[5]:
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    st.markdown('<div class="info-box blue">3 feuilles : <strong>Synthèse réseau</strong> · <strong>Détail complet</strong> · <strong>Alertes prioritaires</strong></div>', unsafe_allow_html=True)
    if st.button("🔨 Générer Export Implantation", type="primary"):
        buf_x = io.BytesIO()
        COLS_DET = ["Magasin", "SKU", "Libellé article", "Origine", "Mode Appro",
                    "Sem. Réception", "Fournisseur T1", "Nouveau stock", "Statut"]
        with pd.ExcelWriter(buf_x, engine="openpyxl") as writer:
            cols_s = ["Magasin"] + [c for c in STATUT_COLORS if c in pivot_mag.columns] + ["Total", "Taux (%)"]
            pivot_mag[cols_s].sort_values("Taux (%)", ascending=False).to_excel(writer, sheet_name="Synthèse Réseau", index=False)
            merged[[c for c in COLS_DET if c in merged.columns]].to_excel(writer, sheet_name="Détail Complet", index=False)
            df_al = merged[merged["Statut"].isin(["⚠️ Non implanté", "🔴 Stock négatif"])]
            df_al[[c for c in COLS_DET if c in df_al.columns]].sort_values(["Statut", "Magasin"]).to_excel(
                writer, sheet_name="Alertes Prioritaires", index=False)
            wb = writer.book
            FH = PatternFill("solid", fgColor="1C1C1E")
            FT = Font(bold=True, color="FFFFFF", name="Arial", size=11)
            for sn in wb.sheetnames:
                ws = wb[sn]
                for cell in ws[1]:
                    cell.fill = FH; cell.font = FT
                    cell.alignment = Alignment(horizontal="center")
                for col in ws.columns:
                    ws.column_dimensions[get_column_letter(col[0].column)].width = min(
                        max((len(str(c.value)) for c in col if c.value), default=10) + 4, 50)
                ws.freeze_panes = "A2"
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
        st.download_button(f"📥 Implantation_T1_{TODAY_FILE}.xlsx", data=buf_x,
                           file_name=f"Implantation_T1_{TODAY_FILE}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success(f"✅ Export généré — {fmt_n(len(merged))} lignes · 3 feuilles")


# ═════════════════════════════════════════════════════════════════════════════
# PARTIE 2 — SUPPLY
# ═════════════════════════════════════════════════════════════════════════════
if not supply_active:
    st.markdown(f"""
    <div class="supply-divider">
      <div class="supply-divider-icon">🚚</div>
      <div>
        <div class="supply-divider-title">Partie Supply — non activée</div>
        <div class="supply-divider-sub">Charge les fichiers RAL (③) dans la sidebar pour activer le suivi des livraisons</div>
      </div>
    </div>""", unsafe_allow_html=True)
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# DATASET SUPPLY
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="supply-divider">
  <div class="supply-divider-icon">🚚</div>
  <div>
    <div class="supply-divider-title">Partie 2 — Supply · Suivi Livraisons & Actions</div>
    <div class="supply-divider-sub">Stock ERP × RAL actif (38 · 40 · 50) · {len(ral_files)} fichier(s) RAL</div>
  </div>
</div>""", unsafe_allow_html=True)

st.markdown('<div class="partie-label supply">🚚 PARTIE 2 — SUPPLY</div>', unsafe_allow_html=True)

supply_df = merged.merge(df_ral_agg, on=["SKU", "Code site"], how="left")
supply_df["RAL_actif"]           = supply_df["RAL_actif"].fillna(0)
supply_df["Nb_commandes"]        = supply_df["Nb_commandes"].fillna(0).astype(int)
supply_df["Commandes_en_retard"] = supply_df["Commandes_en_retard"].fillna(0).astype(int)
supply_df["Jours_retard_max"]    = supply_df["Jours_retard_max"].fillna(0).astype(int)
supply_df["Action Supply"]       = supply_df.apply(
    lambda r: action_supply(r["Nouveau stock"], r["RAL_actif"], r.get("Situation_principale")), axis=1)
supply_df["Retard Badge"]        = supply_df["Jours_retard_max"].apply(retard_badge)

# ─────────────────────────────────────────────────────────────────────────────
# KPI SUPPLY
# ─────────────────────────────────────────────────────────────────────────────
n_attente   = int(supply_df["Action Supply"].str.contains("En attente|En transit|Réception", na=False).sum())
n_commander = int((supply_df["Action Supply"] == "🛒 Passer commande").sum())
n_reg       = int(supply_df["Action Supply"].str.contains("Régulariser", na=False).sum())
n_retard    = int((supply_df["Commandes_en_retard"] > 0).sum())
retard_max  = int(supply_df["Jours_retard_max"].max()) if len(supply_df) else 0

st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card blue"><div class="kpi-label">🕐 En attente livraison</div><div class="kpi-value">{fmt_n(n_attente)}</div><div class="kpi-sub">RAL actif confirmé</div></div>
  <div class="kpi-card orange"><div class="kpi-label">🛒 À commander</div><div class="kpi-value">{fmt_n(n_commander)}</div><div class="kpi-sub">Aucun RAL · action acheteur</div></div>
  <div class="kpi-card red"><div class="kpi-label">🔧 À régulariser</div><div class="kpi-value">{fmt_n(n_reg)}</div><div class="kpi-sub">Écart inventaire magasin</div></div>
  <div class="kpi-card purple"><div class="kpi-label">⚠️ Lignes en retard</div><div class="kpi-value">{fmt_n(n_retard)}</div><div class="kpi-sub">Date ETA dépassée</div></div>
  <div class="kpi-card teal"><div class="kpi-label">🚨 Retard max</div><div class="kpi-value">{retard_max}j</div><div class="kpi-sub">{"🔴 Critique" if retard_max > 30 else ("🟠 Modéré" if retard_max > 7 else "🟡 Léger")}</div></div>
</div>""", unsafe_allow_html=True)

# SCORECARD SUPPLY
st.markdown('<div class="sh-supply">SCORECARD SUPPLY PAR MAGASIN</div>', unsafe_allow_html=True)
sup_pivot = (supply_df.groupby("Magasin")["Action Supply"]
             .value_counts().unstack(fill_value=0).reset_index())
sup_pivot.columns.name = None
sc_html = '<div class="scorecard-grid">'
for _, row in sup_pivot.iterrows():
    att = sum(int(row.get(k, 0)) for k in ["🕐 En attente livraison", "🚚 En transit", "📥 Réception en cours"])
    cmd = int(row.get("🛒 Passer commande", 0))
    reg = sum(int(row.get(k, 0)) for k in ["🔧 Régulariser inventaire", "🔧 Régulariser + livraison en cours"])
    ok  = sum(int(row.get(k, 0)) for k in ["✅ Implanté", "✅ Implanté + réassort en cours"])
    cls = "ok" if cmd == 0 and reg == 0 else ("warn" if cmd > 0 else "ko")
    sc_html += f'<div class="scorecard-card {cls}"><div class="scorecard-name">{row["Magasin"]}</div><div class="scorecard-sub">🕐{att} 🛒{cmd} 🔧{reg} ✅{ok}</div></div>'
sc_html += "</div>"
st.markdown(sc_html, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# ONGLETS SUPPLY
# ─────────────────────────────────────────────────────────────────────────────
TABS_SUP = ["🎯 Plan d'action", "🕐 Suivi livraisons", "🛒 À commander", "📥 Export Supply"]
if "sup_tab" not in st.session_state:
    st.session_state.sup_tab = TABS_SUP[0]

nav_sup = st.columns(len(TABS_SUP))
for i, t in enumerate(TABS_SUP):
    with nav_sup[i]:
        if st.session_state.sup_tab == t:
            st.markdown(f'<div class="nav-active-supply">{t}</div>', unsafe_allow_html=True)
        if st.button(t, key=f"sup_nav_{i}", use_container_width=True):
            st.session_state.sup_tab = t
            st.rerun()

active_sup = st.session_state.sup_tab
COLS_SUP   = ["Magasin", "SKU", "Libellé article", "Origine", "Nouveau stock",
              "RAL_actif", "Nb_commandes", "Prochaine_livraison", "Retard Badge", "Action Supply"]

# ── TAB SUP 1 — PLAN D'ACTION ────────────────────────────────────────────────
if active_sup == TABS_SUP[0]:
    st.markdown('<div class="sh-supply">PLAN D\'ACTION COMPLET — SKU × MAGASIN</div>', unsafe_allow_html=True)
    fc1, fc2 = st.columns(2)
    with fc1:
        act_sel = st.multiselect("Filtrer par action",
                                 sorted(supply_df["Action Supply"].unique()),
                                 default=sorted(supply_df["Action Supply"].unique()))
    with fc2:
        mag_sup_sel = st.multiselect("Filtrer par magasin",
                                     sorted(supply_df["Magasin"].unique()),
                                     default=sorted(supply_df["Magasin"].unique()))
    df_plan = supply_df[supply_df["Action Supply"].isin(act_sel) &
                        supply_df["Magasin"].isin(mag_sup_sel)]
    st.dataframe(df_plan[[c for c in COLS_SUP if c in df_plan.columns]]
                 .rename(columns={"RAL_actif": "RAL", "Nb_commandes": "Nb cdes",
                                  "Prochaine_livraison": "Proch. livraison", "Retard Badge": "Retard"})
                 .sort_values(["Action Supply", "Magasin"]).reset_index(drop=True),
                 use_container_width=True, hide_index=True)

    act_count = supply_df["Action Supply"].value_counts().reset_index()
    act_count.columns = ["Action", "N"]
    fig_act = px.bar(act_count, x="N", y="Action", orientation="h",
                     title="Répartition des actions Supply", color="Action",
                     color_discrete_map={
                         "🛒 Passer commande": C["orange"], "🕐 En attente livraison": C["blue"],
                         "🚚 En transit": C["teal"], "📥 Réception en cours": C["green"],
                         "🔧 Régulariser inventaire": C["red"],
                         "🔧 Régulariser + livraison en cours": "#FF6B6B",
                         "✅ Implanté + réassort en cours": "#34C759", "✅ Implanté": "#30D158",
                     })
    fig_act.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=320,
                          font=dict(family="Inter", color=C["muted"], size=12),
                          margin=dict(l=10, r=10, t=44, b=20), showlegend=False,
                          xaxis=dict(gridcolor=C["bg"]))
    st.plotly_chart(fig_act, use_container_width=True)

# ── TAB SUP 2 — SUIVI LIVRAISONS ─────────────────────────────────────────────
elif active_sup == TABS_SUP[1]:
    df_liv = supply_df[supply_df["RAL_actif"] > 0].copy()
    st.markdown('<div class="sh-supply">COMMANDES EN COURS (SITUATIONS 38 · 40 · 50)</div>', unsafe_allow_html=True)
    if df_liv.empty:
        st.markdown('<div class="info-box green">✅ Aucune commande en cours pour les articles T1.</div>', unsafe_allow_html=True)
    else:
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Lignes avec RAL",     fmt_n(len(df_liv)))
        k2.metric("SKUs distincts",      fmt_n(df_liv["SKU"].nunique()))
        k3.metric("RAL total (unités)",  fmt_n(int(df_liv["RAL_actif"].sum())))
        k4.metric("Lignes en retard",    fmt_n(int((df_liv["Commandes_en_retard"] > 0).sum())))

        df_liv["Tranche retard"] = pd.cut(df_liv["Jours_retard_max"],
                                          bins=[-1, 0, 7, 30, 999],
                                          labels=["✅ À temps", "🟡 1-7j", "🟠 8-30j", "🔴 >30j"])
        ret_dist = df_liv["Tranche retard"].value_counts().reset_index()
        ret_dist.columns = ["Tranche", "N"]
        fig_ret = px.bar(ret_dist, x="Tranche", y="N", color="Tranche",
                         color_discrete_map={"✅ À temps": C["green"], "🟡 1-7j": "#D97706",
                                             "🟠 8-30j": C["orange"], "🔴 >30j": C["red"]},
                         title="Commandes par tranche de retard")
        fig_ret.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=280,
                              font=dict(family="Inter", color=C["muted"], size=12),
                              margin=dict(l=10, r=10, t=44, b=20), showlegend=False,
                              xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]))
        st.plotly_chart(fig_ret, use_container_width=True)

        COLS_LIV = ["Magasin", "SKU", "Libellé article", "Origine", "Nouveau stock",
                    "RAL_actif", "Nb_commandes", "Prochaine_livraison", "Jours_retard_max",
                    "Retard Badge", "Action Supply"]
        st.dataframe(df_liv[[c for c in COLS_LIV if c in df_liv.columns]]
                     .rename(columns={"RAL_actif": "RAL", "Nb_commandes": "Nb cdes",
                                      "Prochaine_livraison": "Proch. livraison",
                                      "Jours_retard_max": "Retard (j)", "Retard Badge": "Statut retard"})
                     .sort_values("Retard (j)", ascending=False).reset_index(drop=True),
                     use_container_width=True, hide_index=True)

# ── TAB SUP 3 — À COMMANDER ──────────────────────────────────────────────────
elif active_sup == TABS_SUP[2]:
    df_cmd = supply_df[supply_df["Action Supply"] == "🛒 Passer commande"].copy()
    st.markdown('<div class="sh-supply">ARTICLES SANS STOCK ET SANS COMMANDE EN COURS</div>', unsafe_allow_html=True)
    st.markdown('<div class="info-box orange">🛒 Ces articles T1 n\'ont <strong>aucun stock positif</strong> et <strong>aucune commande active</strong> dans l\'ERP. Action acheteur : passer commande fournisseur.</div>', unsafe_allow_html=True)
    if df_cmd.empty:
        st.markdown('<div class="info-box green">✅ Tous les articles non implantés ont une commande en cours.</div>', unsafe_allow_html=True)
    else:
        k1, k2, k3 = st.columns(3)
        k1.metric("🛒 À commander",   fmt_n(len(df_cmd)))
        k2.metric("SKUs distincts",   fmt_n(df_cmd["SKU"].nunique()))
        k3.metric("Flux IM",          fmt_n(int((df_cmd["Origine"] == "IM").sum())))
        COLS_CMD = ["Magasin", "SKU", "Libellé article", "Origine", "Mode Appro",
                    "Fournisseur T1", "Sem. Réception", "Nouveau stock"]
        st.dataframe(df_cmd[[c for c in COLS_CMD if c in df_cmd.columns]]
                     .sort_values(["Origine", "Magasin"]).reset_index(drop=True),
                     use_container_width=True, hide_index=True)
        if "Fournisseur T1" in df_cmd.columns:
            four_count = (df_cmd.groupby("Fournisseur T1")["SKU"].nunique()
                          .sort_values(ascending=False).reset_index())
            four_count.columns = ["Fournisseur", "SKU à commander"]
            st.markdown('<div class="sh-supply">PAR FOURNISSEUR</div>', unsafe_allow_html=True)
            st.dataframe(four_count, use_container_width=True, hide_index=True)

# ── TAB SUP 4 — EXPORT SUPPLY ────────────────────────────────────────────────
elif active_sup == TABS_SUP[3]:
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    st.markdown('<div class="info-box purple">3 feuilles : <strong>Plan d\'action</strong> · <strong>Suivi livraisons</strong> · <strong>À commander</strong></div>', unsafe_allow_html=True)
    if st.button("🔨 Générer Export Supply", type="primary"):
        buf_s = io.BytesIO()
        COLS_EXP = ["Magasin", "SKU", "Libellé article", "Origine", "Nouveau stock",
                    "RAL_actif", "Nb_commandes", "Prochaine_livraison",
                    "Jours_retard_max", "Retard Badge", "Action Supply"]
        ACTION_FILLS = {
            "🛒 Passer commande"                 : ("FEF3C7", "92400E"),
            "🕐 En attente livraison"            : ("EFF4FF", "1D4ED8"),
            "🚚 En transit"                      : ("F0FFFE", "0E7490"),
            "📥 Réception en cours"              : ("F0FFF4", "065F46"),
            "🔧 Régulariser inventaire"          : ("FEE2E2", "991B1B"),
            "🔧 Régulariser + livraison en cours": ("FEE2E2", "991B1B"),
            "✅ Implanté + réassort en cours"    : ("D1FAE5", "065F46"),
            "✅ Implanté"                        : ("D1FAE5", "065F46"),
        }
        with pd.ExcelWriter(buf_s, engine="openpyxl") as writer:
            supply_df[[c for c in COLS_EXP if c in supply_df.columns]].sort_values(
                ["Action Supply", "Magasin"]).to_excel(writer, sheet_name="Plan Action", index=False)
            supply_df[supply_df["RAL_actif"] > 0][[c for c in COLS_EXP if c in supply_df.columns]].sort_values(
                "Jours_retard_max", ascending=False).to_excel(writer, sheet_name="Suivi Livraisons", index=False)
            COLS_CMD = ["Magasin", "SKU", "Libellé article", "Origine", "Mode Appro", "Fournisseur T1", "Sem. Réception"]
            supply_df[supply_df["Action Supply"] == "🛒 Passer commande"][
                [c for c in COLS_CMD if c in supply_df.columns]].to_excel(writer, sheet_name="À Commander", index=False)
            wb = writer.book
            FH = PatternFill("solid", fgColor="1C1C1E")
            FT = Font(bold=True, color="FFFFFF", name="Arial", size=11)
            for sn in wb.sheetnames:
                ws = wb[sn]
                for cell in ws[1]:
                    cell.fill = FH; cell.font = FT
                    cell.alignment = Alignment(horizontal="center")
                for col in ws.columns:
                    ws.column_dimensions[get_column_letter(col[0].column)].width = min(
                        max((len(str(c.value)) for c in col if c.value), default=10) + 4, 50)
                ws.freeze_panes = "A2"
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if str(cell.value) in ACTION_FILLS:
                            bg, fg = ACTION_FILLS[str(cell.value)]
                            cell.fill = PatternFill("solid", fgColor=bg)
                            cell.font = Font(color=fg, name="Arial", size=10)
        buf_s.seek(0)
        st.download_button(f"📥 Supply_T1_{TODAY_FILE}.xlsx", data=buf_s,
                           file_name=f"Supply_T1_{TODAY_FILE}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success(f"✅ Export Supply — {fmt_n(len(supply_df))} lignes · 3 feuilles")

# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    f'<div style="text-align:center;font-size:11px;color:{C["muted"]};font-family:JetBrains Mono;">'
    f'SmartBuyer · NovaRetail Solutions · Implantation & Supply v4.0 · {TODAY_STR}'
    f'</div>',
    unsafe_allow_html=True
)
