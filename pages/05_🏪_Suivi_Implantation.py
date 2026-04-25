"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  03_📦_Implantation.py  ·  SmartBuyer · NovaRetail Solutions               ║
║  Suivi Implantation Nouvelles Références — v6.0                             ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║  SOURCES DE DONNÉES                                                          ║
║                                                                              ║
║  ① T1 — Liste nouvelles références                                          ║
║     Format : CSV ou Excel                                                   ║
║     Colonnes : ARTICLE (8 chiffres) · MODE APPRO · SEMAINE RECEPTION        ║
║                LIBELLÉ FOURNISSEUR ORIGINE · DATE LIV.                      ║
║                                                                              ║
║  ② Stock light consolidé — 1 fichier tous magasins                          ║
║     Format  : CSV latin1 · séparateur ';'                                   ║
║     Nommage : stock_DDMM_light.csv  (date extraite automatiquement)         ║
║     Colonnes clés :                                                          ║
║       Site           → code magasin (5 chiffres)                            ║
║       Libellé site   → nom magasin                                          ║
║       Code article   → SKU 8 chiffres                                       ║
║       Libellé article                                                        ║
║       Nouveau stock  → stock actuel (peut être négatif)                     ║
║       Ral            → quantité en commande (RAL global ERP)                ║
║       Code etat      → 2=Actif · P=Purge (exclu) · B/S/F/6/5=Anomalie     ║
║       Code marketing → IM ou LO (flux approvisionnement)                    ║
║       Nom fourn.     → fournisseur ERP                                      ║
║       Libellé rayon  → rayon pour filtres                                   ║
║       Libellé famille                                                        ║
║       Qté sortie     → ventes période (couverture en jours)                 ║
║       Pcb            → conditionnement                                       ║
║       Date dernière entrée → dernière réception physique                    ║
║                                                                              ║
║  LOGIQUE ALERTES                                                             ║
║    ✅ Implanté              Code etat=2 · Stock > 0                         ║
║    🔵 Appro en cours        Code etat=2 · Stock = 0 · RAL > 0              ║
║    🛒 Passer commande       Code etat=2 · Stock = 0 · RAL = 0              ║
║    🔧 Régulariser + appro   Code etat=2 · Stock < 0 · RAL > 0              ║
║    🔧 Régulariser           Code etat=2 · Stock < 0 · RAL = 0              ║
║    🚩 Anomalie référ.       Code etat ≠ 2 (hors P)                         ║
║    ⚪ Non référencé         Article absent du fichier stock                  ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import re
from datetime import date, datetime

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────────────────────
TODAY      = pd.Timestamp(date.today())
TODAY_STR  = date.today().strftime("%d %b %Y")
TODAY_FILE = date.today().strftime("%Y%m%d")

STOCK_COLS = [
    "Site", "Libellé site", "Code article", "Libellé article",
    "Nouveau stock", "Ral", "Code etat", "Code marketing",
    "Nom fourn.", "Libellé rayon", "Libellé famille",
    "Qté sortie", "Pcb", "Date dernière entrée"
]

ETAT_ACTIF    = "2"
ETAT_PURGE    = "P"
ETAT_ANOMALIE = {"B", "S", "F", "6", "5", "1"}
ETAT_LABEL    = {
    "B": "Rayon générique", "S": "Suspendu",
    "F": "Fin de vie",      "6": "Déréférencé", "5": "Autre",
}

# Alertes et couleurs
ALERTES = {
    "✅ Implanté"            : "#34C759",
    "🔵 Appro en cours"      : "#007AFF",
    "🛒 Passer commande"     : "#FF9500",
    "🔧 Régulariser + appro" : "#FF6B6B",
    "🔧 Régulariser"         : "#FF3B30",
    "🚩 Anomalie référ."     : "#FFD60A",
    "⚪ Non référencé"       : "#8E8E93",
}

ACTION_LABEL = {
    "✅ Implanté"            : "—",
    "🔵 Appro en cours"      : "Accélérer livraison",
    "🛒 Passer commande"     : "Passer commande fournisseur",
    "🔧 Régulariser + appro" : "Régulariser inventaire + livraison en cours",
    "🔧 Régulariser"         : "Régulariser inventaire magasin",
    "🚩 Anomalie référ."     : "Vérifier référencement magasin",
    "⚪ Non référencé"       : "Article non dans le stock ERP",
}

C = {
    "bg"      : "#F2F2F7", "surface" : "#FFFFFF",
    "border"  : "#E5E5EA", "text"    : "#1C1C1E",
    "muted"   : "#6D6D72", "blue"    : "#007AFF",
    "green"   : "#34C759", "red"     : "#FF3B30",
    "orange"  : "#FF9500", "purple"  : "#AF52DE",
    "teal"    : "#5AC8FA", "yellow"  : "#FFD60A",
    "blue_l"  : "#EFF4FF", "green_l" : "#F0FFF4",
    "red_l"   : "#FFF2F0", "orange_l": "#FFFBEB",
    "purple_l": "#F5F0FF",
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
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
:root {{
  --bg:{C['bg']}; --surface:{C['surface']}; --border:{C['border']};
  --text:{C['text']}; --muted:{C['muted']};
  --blue:{C['blue']}; --green:{C['green']}; --red:{C['red']};
  --orange:{C['orange']}; --purple:{C['purple']}; --teal:{C['teal']};
  --yellow:{C['yellow']};
  --blue-l:{C['blue_l']}; --green-l:{C['green_l']};
  --red-l:{C['red_l']}; --orange-l:{C['orange_l']}; --purple-l:{C['purple_l']};
  --radius:14px; --shadow:0 1px 3px rgba(0,0,0,.06),0 4px 16px rgba(0,0,0,.04);
}}
html,body,[class*="css"]{{font-family:'Inter',sans-serif!important;background:var(--bg)!important;color:var(--text)!important;}}
.main,section[data-testid="stMain"]{{background:var(--bg)!important;}}
.block-container{{padding:0 2rem 4rem!important;max-width:1480px;}}
header[data-testid="stHeader"],#MainMenu,footer{{display:none!important;}}

/* TOPBAR */
.topbar{{background:var(--text);margin:0 -2rem 28px;padding:16px 28px;display:flex;align-items:center;justify-content:space-between;}}
.topbar-icon{{width:40px;height:40px;border-radius:10px;background:linear-gradient(135deg,{C['blue']},{C['purple']});display:flex;align-items:center;justify-content:center;font-size:22px;}}
.topbar-left{{display:flex;align-items:center;gap:14px;}}
.topbar-title{{font-size:17px;font-weight:700;color:#fff;letter-spacing:-.01em;}}
.topbar-sub{{font-size:11px;color:#8E8E93;font-family:'JetBrains Mono';margin-top:2px;}}
.topbar-pill{{background:rgba(255,255,255,.08);color:#8E8E93;border:1px solid rgba(255,255,255,.12);border-radius:8px;padding:4px 14px;font-size:11px;font-weight:600;}}
.topbar-date{{color:{C['blue']};font-size:12px;font-family:'JetBrains Mono';}}

/* INTRO */
.module-intro{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px 24px;margin-bottom:20px;box-shadow:var(--shadow);}}
.module-intro-title{{font-size:15px;font-weight:700;color:var(--text);margin-bottom:6px;}}
.module-intro-body{{font-size:13px;color:var(--muted);line-height:1.6;}}
.tag-badge{{display:inline-flex;align-items:center;gap:5px;background:var(--blue-l);color:var(--blue);border:1px solid #BFDBFE;border-radius:6px;padding:3px 10px;font-size:11px;font-weight:600;margin:2px;}}
.tag-badge.green{{background:var(--green-l);color:var(--green);border-color:#6EE7B7;}}
.tag-badge.orange{{background:var(--orange-l);color:var(--orange);border-color:#FCD34D;}}
.tag-badge.red{{background:var(--red-l);color:var(--red);border-color:#FECACA;}}

/* SH */
.sh{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:var(--muted);margin:22px 0 12px;padding-bottom:8px;border-bottom:1px solid var(--border);}}

/* KPI */
.kpi-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:22px;}}
.kpi-grid-4{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:22px;}}
.kpi-card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:18px 20px 14px;box-shadow:var(--shadow);position:relative;overflow:hidden;}}
.kpi-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:var(--radius) var(--radius) 0 0;}}
.kpi-card.blue::before{{background:var(--blue);}}   .kpi-card.green::before{{background:var(--green);}}
.kpi-card.red::before{{background:var(--red);}}     .kpi-card.orange::before{{background:var(--orange);}}
.kpi-card.purple::before{{background:var(--purple);}} .kpi-card.yellow::before{{background:var(--yellow);}}
.kpi-card.teal::before{{background:var(--teal);}}
.kpi-label{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.10em;color:var(--muted);margin-bottom:10px;}}
.kpi-value{{font-size:36px;font-weight:800;line-height:1;letter-spacing:-.02em;}}
.kpi-card.blue .kpi-value{{color:var(--blue);}}   .kpi-card.green .kpi-value{{color:var(--green);}}
.kpi-card.red .kpi-value{{color:var(--red);}}     .kpi-card.orange .kpi-value{{color:var(--orange);}}
.kpi-card.purple .kpi-value{{color:var(--purple);}} .kpi-card.yellow .kpi-value{{color:#B8860B;}}
.kpi-card.teal .kpi-value{{color:var(--teal);}}
.kpi-sub{{font-size:11px;color:var(--muted);font-family:'JetBrains Mono';margin-top:4px;}}
.kpi-bar{{margin-top:12px;height:3px;border-radius:3px;background:var(--border);}}
.kpi-bar-fill{{height:100%;border-radius:3px;}}

/* SCORECARD */
.scorecard-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:10px;margin-bottom:22px;}}
.scorecard-card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px;box-shadow:var(--shadow);position:relative;}}
.scorecard-card.ok{{border-color:#6EE7B7;background:var(--green-l);}}
.scorecard-card.warn{{border-color:#FCD34D;background:var(--orange-l);}}
.scorecard-card.ko{{border-color:#FECACA;background:var(--red-l);}}
.scorecard-dot{{width:8px;height:8px;border-radius:50%;position:absolute;top:14px;right:14px;}}
.scorecard-name{{font-size:11px;font-weight:600;color:var(--text);margin-bottom:6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:88%;}}
.scorecard-pct{{font-size:28px;font-weight:800;line-height:1;}}
.scorecard-sub{{font-size:10px;color:var(--muted);font-family:'JetBrains Mono';margin-top:3px;}}

/* ALERT BANNER */
.alert-banner{{background:#FFF;border:1px solid #FECACA;border-left:4px solid var(--red);border-radius:var(--radius);padding:14px 20px;margin-bottom:18px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;}}
.alert-pill{{background:var(--red);color:#fff;border-radius:6px;padding:4px 12px;font-size:11px;font-weight:700;letter-spacing:.05em;white-space:nowrap;}}
.alert-item{{display:flex;flex-direction:column;align-items:center;padding:0 14px;border-right:1px solid var(--border);}}
.alert-item:last-child{{border-right:none;}}
.alert-num{{font-size:24px;font-weight:800;line-height:1;}}
.alert-lbl{{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-top:1px;}}

/* INFO BOXES */
.info-box{{border-radius:var(--radius);padding:14px 18px;margin-bottom:14px;border:1px solid;font-size:13px;line-height:1.6;}}
.info-box.blue{{background:var(--blue-l);border-color:#BFDBFE;color:#1D4ED8;}}
.info-box.green{{background:var(--green-l);border-color:#6EE7B7;color:#065F46;}}
.info-box.orange{{background:var(--orange-l);border-color:#FCD34D;color:#92400E;}}
.info-box.yellow{{background:#FFFDE7;border-color:#FDD835;color:#795548;}}

/* VALIDATION BOX */
.val-box{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:12px 20px;margin-bottom:18px;box-shadow:var(--shadow);display:flex;align-items:center;gap:16px;flex-wrap:wrap;}}
.val-item{{display:flex;flex-direction:column;align-items:center;padding:0 14px;border-right:1px solid var(--border);}}
.val-item:last-child{{border-right:none;padding-right:0;}}
.val-num{{font-size:20px;font-weight:800;line-height:1;}}
.val-lbl{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-top:2px;}}

/* ALERTE ACTION CARD */
.action-row{{display:flex;align-items:center;gap:8px;padding:8px 0;border-bottom:1px solid var(--border);}}
.action-row:last-child{{border-bottom:none;}}
.action-badge{{display:inline-flex;align-items:center;border-radius:6px;padding:3px 10px;font-size:11px;font-weight:700;white-space:nowrap;}}

/* NAV */
.nav-active{{background:var(--text)!important;color:#fff!important;border-radius:10px;padding:10px 0;text-align:center;font-size:13px;font-weight:700;box-shadow:0 4px 14px rgba(28,28,30,.2);margin-bottom:10px;}}

/* SIDEBAR */
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

def extract_date_from_filename(filename: str) -> pd.Timestamp:
    m = re.search(r'(\d{8})', filename)
    if m:
        for fmt in ('%Y%m%d', '%d%m%Y'):
            try:
                return pd.Timestamp(pd.to_datetime(m.group(1), format=fmt))
            except Exception:
                continue
    m2 = re.search(r'_(\d{4})[_\.]', filename)
    if m2:
        s = m2.group(1)
        try:
            d, mo = int(s[:2]), int(s[2:])
            return pd.Timestamp(datetime(date.today().year, mo, d))
        except Exception:
            pass
    return TODAY

def _sem_to_num(s: str) -> int:
    cleaned = re.sub(r"[Ss]", "", str(s).strip())
    return int(cleaned) if cleaned.isdigit() else 99

def _sem_sort(s) -> int:
    cleaned = re.sub(r"[Ss]", "", str(s).strip())
    return int(cleaned) if cleaned.isdigit() else 99

def get_alerte(row) -> str:
    """Logique d'alerte métier centrale."""
    etat  = str(row.get("Code etat", "")).strip()
    stock = row.get("Nouveau stock")
    ral   = float(row.get("Ral", 0) or 0)

    # Article absent du stock (left join → ligne vide)
    if not etat or etat == "nan":
        return "⚪ Non référencé"
    # Purge — ne devrait pas arriver (filtré au parsing) mais sécurité
    if etat == ETAT_PURGE:
        return "⚪ Non référencé"
    # Anomalie référencement
    if etat in ETAT_ANOMALIE:
        return "🚩 Anomalie référ."
    # Actif (Code etat = 2)
    if pd.isna(stock):
        return "⚪ Non référencé"
    s = float(stock)
    if s > 0:
        return "✅ Implanté"
    if s < 0 and ral > 0:
        return "🔧 Régulariser + appro"
    if s < 0:
        return "🔧 Régulariser"
    # stock = 0
    if ral > 0:
        return "🔵 Appro en cours"
    return "🛒 Passer commande"


# ─────────────────────────────────────────────────────────────────────────────
# PARSERS
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_t1(file_bytes: bytes, filename: str):
    buf = io.BytesIO(file_bytes)
    try:
        df = pd.read_excel(buf, header=None, dtype=str) \
             if filename.lower().endswith((".xlsx", ".xls")) \
             else pd.read_csv(buf, header=None, sep=None, engine="python",
                              encoding="latin1", dtype=str, on_bad_lines="skip")
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
                  .str.replace("\ufeff", "", regex=False)
                  .str.replace("\xa0", " ", regex=False))

    if "ARTICLE" not in df.columns:
        return None, "Colonne ARTICLE introuvable dans le fichier T1"

    df["SKU"] = (df["ARTICLE"].astype(str).str.strip()
                 .str.replace(r"\.0$", "", regex=True).str.zfill(8).str[:8])
    df = df[df["SKU"].str.match(r"^\d{8}$", na=False)].drop_duplicates("SKU").copy()

    defaults = {
        "LIBELLÉ ARTICLE": "", "LIBELLÉ FOURNISSEUR ORIGINE": "",
        "MODE APPRO": "", "SEMAINE RECEPTION": "", "DATE LIV.": ""
    }
    for col, val in defaults.items():
        if col not in df.columns:
            df[col] = val

    df["SEMAINE RECEPTION"] = df["SEMAINE RECEPTION"].astype(str).str.strip().replace("nan", "")
    df["SEM_NUM"] = df["SEMAINE RECEPTION"].apply(_sem_to_num)
    df["ORIGINE"] = df["MODE APPRO"].apply(
        lambda m: "IM" if "IMPORT" in str(m).upper() else "LO"
    )
    return df, None


@st.cache_data(show_spinner=False)
def parse_stock(file_bytes: bytes, filename: str, sku_scope: tuple):
    """
    Parse le stock light consolidé.
    Performance : usecols + dtype + filtre P + filtre SKU T1 immédiat.
    """
    buf = io.BytesIO(file_bytes)
    try:
        df = pd.read_csv(
            buf, sep=";", encoding="latin1", low_memory=False,
            on_bad_lines="skip", usecols=STOCK_COLS,
            dtype={
                "Site": str, "Code article": str, "Code etat": str,
                "Code marketing": str, "Libellé site": str,
                "Libellé article": str, "Nom fourn.": str,
                "Libellé rayon": str, "Libellé famille": str,
            }
        )
    except Exception as e:
        return None, f"Lecture stock : {e}"

    # Nettoyage libellés
    for col in ["Libellé site", "Libellé article", "Nom fourn.",
                "Libellé rayon", "Code etat", "Code marketing"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r"\t+", "", regex=True)

    df["SKU"]      = (df["Code article"].astype(str).str.strip()
                      .str.replace(r"\.0$", "", regex=True).str.zfill(8).str[:8])
    df["Code site"] = df["Site"].astype(str).str.strip()

    # Filtre purge silencieux
    df = df[df["Code etat"] != ETAT_PURGE].copy()

    # Filtre SKU T1 — réduction mémoire majeure
    if sku_scope:
        df = df[df["SKU"].isin(sku_scope)].copy()

    # Numériques
    df["Nouveau stock"] = pd.to_numeric(df["Nouveau stock"], errors="coerce")
    df["Ral"]           = pd.to_numeric(df["Ral"], errors="coerce").fillna(0).astype(int)
    df["Qté sortie"]    = pd.to_numeric(df["Qté sortie"], errors="coerce").fillna(0)

    # Flux
    df["Origine"] = df["Code marketing"].apply(
        lambda m: "IM" if str(m).strip().upper() == "IM" else "LO"
    )
    return df, None


# ─────────────────────────────────────────────────────────────────────────────
# TOPBAR
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="topbar">
  <div class="topbar-left">
    <div class="topbar-icon">📦</div>
    <div>
      <div class="topbar-title">Suivi Implantation · Nouvelles Références</div>
      <div class="topbar-sub">T1 · Stock ERP consolidé · Alertes · Cessions</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:12px;">
    <div class="topbar-date">{TODAY_STR}</div>
    <div class="topbar-pill">v6.0 · SmartBuyer</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# INTRO
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="module-intro">
  <div class="module-intro-title">📦 À quoi sert ce module ?</div>
  <div class="module-intro-body">
    Suivi en temps réel de l'implantation des <strong>nouvelles références T1</strong> dans le réseau.
    À partir de la liste T1 et du stock ERP consolidé, le module calcule le <strong>taux d'implantation</strong>
    par magasin et réseau, détecte les alertes article par article, et propose des <strong>actions concrètes</strong> :
    accélérer une livraison, passer une commande, régulariser un inventaire.<br><br>
    <span class="tag-badge green">✅ Taux d'implantation</span>
    <span class="tag-badge">🔵 Appro à accélérer</span>
    <span class="tag-badge orange">🛒 Commandes à passer</span>
    <span class="tag-badge red">🔧 Inventaires à régulariser</span>
  </div>
</div>
""", unsafe_allow_html=True)

with st.expander("📋 Structure des fichiers attendus", expanded=False):
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**① T1 — Nouvelles Références** (CSV ou Excel)")
        st.code("ARTICLE · MODE APPRO · SEMAINE RECEPTION\nLIBELLÉ FOURNISSEUR ORIGINE · DATE LIV.", language="text")
        st.caption("Code article 8 chiffres · MODE APPRO contient 'IMPORT' pour IM")
    with c2:
        st.markdown("**② Stock light consolidé** (CSV latin1 · sep ';' · nommage : stock_DDMM_light.csv)")
        st.code("Site · Libellé site · Code article · Libellé article\nNouveau stock · Ral · Code etat · Code marketing\nNom fourn. · Libellé rayon · Qté sortie", language="text")
        st.caption("Code etat : 2=Actif · P=Purge (exclu) · B/S/F=Anomalie · Code marketing : IM ou LO")

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📁 Fichiers")
    st.divider()
    st.markdown("**① T1 — Nouvelles Références**")
    t1_file = st.file_uploader("T1", type=["csv", "xlsx", "xls"],
                               key="t1", label_visibility="collapsed")
    st.markdown("**② Stock light consolidé**")
    st.caption("1 fichier · tous magasins · nommage stock_DDMM_light.csv")
    stk_file = st.file_uploader("Stock", type=["csv"],
                                key="stk", label_visibility="collapsed")

# ─────────────────────────────────────────────────────────────────────────────
# GATES
# ─────────────────────────────────────────────────────────────────────────────
if not t1_file:
    st.markdown('<div class="info-box blue">⬆️ <strong>Étape 1</strong> — Charge le fichier T1 dans la sidebar.</div>', unsafe_allow_html=True)
    st.stop()

with st.spinner("Lecture T1…"):
    t1_df, t1_err = parse_t1(t1_file.read(), t1_file.name)

if t1_err or t1_df is None:
    st.error(f"❌ T1 : {t1_err}")
    st.stop()

if not stk_file:
    st.markdown(f'<div class="info-box blue">✅ T1 chargé — <strong>{len(t1_df):,}</strong> références. ⬆️ <strong>Étape 2</strong> — Charge le fichier Stock light consolidé.</div>', unsafe_allow_html=True)
    st.stop()

# Date du stock
date_stock     = extract_date_from_filename(stk_file.name)
date_stock_str = date_stock.strftime("%d %b %Y")
age_stock      = (TODAY - date_stock).days

SKU_SCOPE = tuple(sorted(t1_df["SKU"].unique()))

with st.spinner("Parsing Stock ERP…"):
    df_stock, stk_err = parse_stock(stk_file.read(), stk_file.name, SKU_SCOPE)

if stk_err or df_stock is None:
    st.error(f"❌ Stock : {stk_err}")
    st.stop()

# Alerte données anciennes
if age_stock > 7:
    st.warning(f"⚠️ Données stock du **{date_stock_str}** — {age_stock} jours. Recharge un fichier plus récent.")

# ─────────────────────────────────────────────────────────────────────────────
# RÉFÉRENTIEL MAGASINS
# ─────────────────────────────────────────────────────────────────────────────
site_ref  = (df_stock[["Code site", "Libellé site"]]
             .drop_duplicates("Code site")
             .set_index("Code site")["Libellé site"].to_dict())
all_codes = sorted(site_ref.keys())

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — FILTRES
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.divider()
    st.markdown("## 🔍 Filtres")

    mag_labels     = sorted([site_ref.get(c, c) for c in all_codes])
    mag_sel_labels = st.multiselect("Magasins", mag_labels, default=mag_labels)
    mag_sel_codes  = [c for c in all_codes if site_ref.get(c, c) in mag_sel_labels]

    orig_sel = st.multiselect("Flux", ["IM", "LO"], default=["IM", "LO"])

    sem_dispo = sorted(
        [s for s in t1_df["SEMAINE RECEPTION"].unique()
         if str(s).strip() not in ("nan", "", "99")],
        key=_sem_sort
    )
    sem_sel = st.multiselect("Semaine réception", sem_dispo, default=sem_dispo)

    rayons_dispo = sorted([r for r in df_stock["Libellé rayon"].dropna().unique()
                           if str(r).strip() not in ("nan", "")])
    rayon_sel = st.multiselect("Rayon", rayons_dispo, default=rayons_dispo) \
                if rayons_dispo else []

    st.divider()
    st.markdown("## 🔄 Cessions")
    mag_detresse = st.multiselect("Magasins en détresse", mag_labels, default=[])
    seuil_det       = st.number_input("Seuil stock (≤)", 0, 50, 0, 1)
    min_1pcb        = st.toggle("Qté min = 1 PCB", value=True,
                                help="Si activé, ne propose que les cessions où la quantité cessible ≥ 1 PCB")

if not mag_sel_codes:
    st.warning("⚠️ Sélectionne au moins un magasin.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# FILTRES SKU
# ─────────────────────────────────────────────────────────────────────────────
mask = t1_df["ORIGINE"].isin(orig_sel)
if sem_sel:
    mask = mask & t1_df["SEMAINE RECEPTION"].isin(sem_sel)
t1_scope  = t1_df[mask].copy()
sku_scope = t1_scope["SKU"].unique()

if len(sku_scope) == 0:
    st.warning("Aucun SKU correspondant aux filtres.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# VALIDATION CROISÉE T1 × STOCK
# ─────────────────────────────────────────────────────────────────────────────
sku_dans_stock = set(df_stock[df_stock["Code site"].isin(mag_sel_codes)]["SKU"].unique())
n_sku          = len(sku_scope)
n_sku_trouves  = len([s for s in sku_scope if s in sku_dans_stock])
n_sku_absents  = n_sku - n_sku_trouves
n_mag          = len(mag_sel_codes)

st.markdown(f"""
<div class="val-box">
  <div style="font-size:13px;font-weight:700;color:{C['text']};margin-right:4px;">📋 T1 × Stock</div>
  <div class="val-item"><div class="val-num" style="color:{C['blue']}">{fmt_n(n_sku)}</div><div class="val-lbl">SKUs T1</div></div>
  <div class="val-item"><div class="val-num" style="color:{C['green']}">{fmt_n(n_sku_trouves)}</div><div class="val-lbl">Trouvés</div></div>
  <div class="val-item"><div class="val-num" style="color:{C['red'] if n_sku_absents > 0 else C['green']}">{fmt_n(n_sku_absents)}</div><div class="val-lbl">Absents stock</div></div>
  <div class="val-item"><div class="val-num" style="color:{C['muted']}">{n_mag}</div><div class="val-lbl">Magasins</div></div>
  <div style="margin-left:auto;font-size:11px;color:{C['muted']};">
    Données stock : <strong style="color:{C['blue']}">{date_stock_str}</strong>
    {"&nbsp;·&nbsp;<span style='color:#FF3B30'>⚠️ " + str(age_stock) + "j</span>" if age_stock > 7 else ""}
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTRUCTION DATASET
# ─────────────────────────────────────────────────────────────────────────────
stk_filt = df_stock[
    df_stock["Code site"].isin(mag_sel_codes) &
    df_stock["SKU"].isin(sku_scope)
].copy()

if rayon_sel:
    stk_filt = stk_filt[stk_filt["Libellé rayon"].isin(rayon_sel)]

# Grille SKU × magasin
grid = pd.DataFrame(
    pd.MultiIndex.from_product(
        [mag_sel_codes, sku_scope], names=["Code site", "SKU"]
    ).tolist(), columns=["Code site", "SKU"]
)

KEEP = ["Code site", "SKU", "Nouveau stock", "Ral", "Code etat",
        "Origine", "Libellé article", "Nom fourn.",
        "Libellé rayon", "Libellé famille", "Qté sortie"]
merged = grid.merge(stk_filt[[c for c in KEEP if c in stk_filt.columns]],
                    on=["Code site", "SKU"], how="left")

# Référentiel T1
t1_ref = t1_scope.set_index("SKU")[[
    "LIBELLÉ ARTICLE", "LIBELLÉ FOURNISSEUR ORIGINE",
    "MODE APPRO", "SEMAINE RECEPTION", "DATE LIV.", "ORIGINE", "SEM_NUM"
]].rename(columns={
    "LIBELLÉ ARTICLE"             : "T1_lib",
    "LIBELLÉ FOURNISSEUR ORIGINE" : "Fournisseur T1",
    "MODE APPRO"                  : "Mode Appro",
    "SEMAINE RECEPTION"           : "Sem. Réception",
    "DATE LIV."                   : "Date Livraison",
    "ORIGINE"                     : "Origine_T1",
    "SEM_NUM"                     : "SEM_NUM",
})
merged = merged.merge(t1_ref.reset_index(), on="SKU", how="left")

# Libellé article et flux
if "Libellé article" not in merged.columns:
    merged["Libellé article"] = ""
merged["Libellé article"] = merged["Libellé article"].fillna("").astype(str)
merged["Libellé article"] = merged.apply(
    lambda r: r["Libellé article"] if r["Libellé article"] else r.get("T1_lib", ""), axis=1
)
merged["Origine"] = merged.apply(
    lambda r: r.get("Origine") if pd.notna(r.get("Origine")) and str(r.get("Origine")) not in ("nan","")
    else r.get("Origine_T1", "LO"), axis=1
)
merged.drop(columns=["T1_lib", "Origine_T1"], errors="ignore", inplace=True)
merged["Magasin"] = merged["Code site"].map(site_ref).fillna(merged["Code site"])
merged["Code etat"] = merged["Code etat"].fillna("").astype(str)
merged["Ral"]       = pd.to_numeric(merged["Ral"], errors="coerce").fillna(0)

# Alertes
merged["Alerte"] = merged.apply(get_alerte, axis=1)
merged["Action"] = merged["Alerte"].map(ACTION_LABEL)

# ─────────────────────────────────────────────────────────────────────────────
# MÉTRIQUES
# ─────────────────────────────────────────────────────────────────────────────
# Taux sur actifs (Code etat = 2) uniquement
merged_actif = merged[merged["Code etat"].str.strip() == ETAT_ACTIF]
n_base_taux  = len(merged_actif)
n_impl       = int((merged["Alerte"] == "✅ Implanté").sum())
n_appro      = int((merged["Alerte"] == "🔵 Appro en cours").sum())
n_cmd        = int((merged["Alerte"] == "🛒 Passer commande").sum())
n_reg_appro  = int((merged["Alerte"] == "🔧 Régulariser + appro").sum())
n_reg        = int((merged["Alerte"] == "🔧 Régulariser").sum())
n_anomalie   = int((merged["Alerte"] == "🚩 Anomalie référ.").sum())
n_non_ref    = int((merged["Alerte"] == "⚪ Non référencé").sum())
taux_reseau  = int(n_impl / n_base_taux * 100) if n_base_taux else 0
n_sku_im     = int((t1_scope["ORIGINE"] == "IM").sum())
n_sku_lo     = int((t1_scope["ORIGINE"] == "LO").sum())
total_cells  = len(merged)
pct          = lambda n: int(n / total_cells * 100) if total_cells else 0

# Pivot magasin
def taux_mag(mag):
    dm = merged_actif[merged_actif["Magasin"] == mag]
    if len(dm) == 0: return 0
    return int((dm["Alerte"] == "✅ Implanté").sum() / len(dm) * 100)

pivot_mag = (
    merged.groupby(["Magasin", "Alerte"]).size()
    .unstack(fill_value=0)
    .reindex(columns=list(ALERTES.keys()), fill_value=0)
    .reset_index()
)
pivot_mag.columns.name = None
pivot_mag["Taux (%)"] = pivot_mag["Magasin"].apply(taux_mag)

# Matrice rayon × magasin
rayon_pivot = pd.DataFrame()
if "Libellé rayon" in merged.columns:
    try:
        rayon_pivot = (
            merged_actif.groupby(["Libellé rayon", "Magasin"])
            .apply(lambda x: int((x["Alerte"] == "✅ Implanté").sum() / len(x) * 100)
                   if len(x) > 0 else 0)
            .reset_index(name="Taux (%)")
            .pivot(index="Libellé rayon", columns="Magasin", values="Taux (%)")
            .fillna(0).astype(int)
        )
    except Exception:
        rayon_pivot = pd.DataFrame()

# ─────────────────────────────────────────────────────────────────────────────
# BANNIÈRE
# ─────────────────────────────────────────────────────────────────────────────
n_actions = n_appro + n_cmd + n_reg + n_reg_appro + n_anomalie
if n_actions > 0:
    st.markdown(f"""
    <div class="alert-banner">
      <div class="alert-pill">⚡ ACTIONS</div>
      <div class="alert-item"><div class="alert-num" style="color:{C['blue']}">{fmt_n(n_appro)}</div><div class="alert-lbl">Appro en cours</div></div>
      <div class="alert-item"><div class="alert-num" style="color:{C['orange']}">{fmt_n(n_cmd)}</div><div class="alert-lbl">À commander</div></div>
      <div class="alert-item"><div class="alert-num" style="color:{C['red']}">{fmt_n(n_reg + n_reg_appro)}</div><div class="alert-lbl">Régulariser</div></div>
      <div class="alert-item"><div class="alert-num" style="color:#B8860B">{fmt_n(n_anomalie)}</div><div class="alert-lbl">Anomalies</div></div>
      <div style="margin-left:auto;font-size:11px;color:{C['muted']};">{n_mag} mag · {fmt_n(n_sku)} SKU</div>
    </div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# KPI
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card green">
    <div class="kpi-label">✅ Implanté</div>
    <div class="kpi-value">{fmt_n(n_impl)}</div>
    <div class="kpi-sub">{pct(n_impl)}% du réseau</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(n_impl)}%;background:{C['green']}"></div></div>
  </div>
  <div class="kpi-card blue">
    <div class="kpi-label">📊 Taux réseau</div>
    <div class="kpi-value" style="color:{color_taux(taux_reseau)}">{taux_reseau}%</div>
    <div class="kpi-sub">sur {fmt_n(n_base_taux)} refs actives</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{taux_reseau}%;background:{C['blue']}"></div></div>
  </div>
  <div class="kpi-card teal">
    <div class="kpi-label">🔵 Appro en cours</div>
    <div class="kpi-value">{fmt_n(n_appro)}</div>
    <div class="kpi-sub">RAL > 0 · accélérer</div>
  </div>
  <div class="kpi-card orange">
    <div class="kpi-label">🛒 À commander</div>
    <div class="kpi-value">{fmt_n(n_cmd)}</div>
    <div class="kpi-sub">Stock 0 · RAL 0</div>
  </div>
  <div class="kpi-card red">
    <div class="kpi-label">🔧 À régulariser</div>
    <div class="kpi-value">{fmt_n(n_reg + n_reg_appro)}</div>
    <div class="kpi-sub">Stock négatif</div>
  </div>
</div>
<div class="kpi-grid-4" style="margin-top:-8px;">
  <div class="kpi-card yellow">
    <div class="kpi-label">🚩 Anomalies référ.</div>
    <div class="kpi-value">{fmt_n(n_anomalie)}</div>
    <div class="kpi-sub">Code etat ≠ 2 · hors taux</div>
  </div>
  <div class="kpi-card purple">
    <div class="kpi-label">🔀 Flux IM / LO</div>
    <div class="kpi-value" style="font-size:24px">{n_sku_im} <span style="font-size:15px;color:{C['muted']}">/ </span>{n_sku_lo}</div>
    <div class="kpi-sub">Import · Local</div>
  </div>
  <div class="kpi-card" style="border-top:3px solid {C['muted']};">
    <div class="kpi-label">⚪ Non référencés</div>
    <div class="kpi-value" style="color:{C['muted']}">{fmt_n(n_non_ref)}</div>
    <div class="kpi-sub">Absents du stock ERP</div>
  </div>
  <div class="kpi-card" style="border-top:3px solid {C['muted']};">
    <div class="kpi-label">📅 Données stock</div>
    <div class="kpi-value" style="font-size:18px;color:{C['red'] if age_stock>7 else C['blue']}">{date_stock_str}</div>
    <div class="kpi-sub">{"⚠️ " + str(age_stock) + "j — recharger" if age_stock > 7 else str(age_stock) + "j"}</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SCORECARD
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="sh">SCORECARD MAGASINS</div>', unsafe_allow_html=True)
sc_html = '<div class="scorecard-grid">'
for _, row in pivot_mag.sort_values("Taux (%)", ascending=False).iterrows():
    t_   = row["Taux (%)"]
    col  = color_taux(t_)
    impl = int(row.get("✅ Implanté", 0))
    app  = int(row.get("🔵 Appro en cours", 0))
    cmd  = int(row.get("🛒 Passer commande", 0))
    reg  = int(row.get("🔧 Régulariser", 0)) + int(row.get("🔧 Régulariser + appro", 0))
    an   = int(row.get("🚩 Anomalie référ.", 0))
    sc_html += f"""
    <div class="scorecard-card {scorecard_cls(t_)}">
      <div class="scorecard-dot" style="background:{col}"></div>
      <div class="scorecard-name">{row['Magasin']}</div>
      <div class="scorecard-pct" style="color:{col}">{t_}%</div>
      <div class="scorecard-sub">{impl}✅ {app}🔵 {cmd}🛒 {reg}🔧 {an}🚩</div>
    </div>"""
sc_html += "</div>"
st.markdown(sc_html, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# ONGLETS
# ─────────────────────────────────────────────────────────────────────────────
TABS = ["📋 COPIL", "📊 Vue Réseau", "🚨 Alertes", "🔄 Cessions", "📥 Export"]
if "impl_tab" not in st.session_state:
    st.session_state.impl_tab = TABS[0]

nav_cols = st.columns(len(TABS))
for i, t in enumerate(TABS):
    with nav_cols[i]:
        if st.session_state.impl_tab == t:
            st.markdown(f'<div class="nav-active">{t}</div>', unsafe_allow_html=True)
        if st.button(t, key=f"nav_{i}", use_container_width=True):
            st.session_state.impl_tab = t
            st.rerun()

active = st.session_state.impl_tab

# ══════════════════════════════════════════════════════════════════════════════
# COPIL
# ══════════════════════════════════════════════════════════════════════════════
if active == TABS[0]:
    st.markdown('<div class="sh">SYNTHÈSE DIRECTION — 30 SECONDES</div>', unsafe_allow_html=True)

    # Synthèse magasins
    cols_aff = ["Magasin"] + [c for c in ALERTES if c in pivot_mag.columns] + ["Taux (%)"]
    st.dataframe(
        pivot_mag[cols_aff].sort_values("Taux (%)", ascending=False).reset_index(drop=True)
        .style.format({"Taux (%)": "{}%"}),
        use_container_width=True, hide_index=True
    )

    # Matrice rayon × magasin
    if not rayon_pivot.empty:
        st.markdown('<div class="sh">TAUX D\'IMPLANTATION PAR RAYON × MAGASIN</div>', unsafe_allow_html=True)
        def color_cell(val):
            if val >= 80: return "background-color:#D1FAE5;color:#065F46;font-weight:700"
            if val >= 50: return "background-color:#FEF3C7;color:#92400E;font-weight:700"
            return "background-color:#FEE2E2;color:#991B1B;font-weight:700"
        st.dataframe(rayon_pivot.style.map(color_cell).format("{}%"),
                     use_container_width=True)

    # Top priorités par criticité
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sh">🛒 TOP ARTICLES À COMMANDER</div>', unsafe_allow_html=True)
        df_cmd = merged[merged["Alerte"] == "🛒 Passer commande"].groupby("SKU").agg(
            Libellé=("Libellé article","first"),
            Fournisseur=("Fournisseur T1","first"),
            Origine=("Origine","first"),
            Nb_magasins=("Magasin","count"),
        ).reset_index().sort_values("Nb_magasins", ascending=False).head(10)
        if df_cmd.empty:
            st.success("✅ Aucun article sans commande")
        else:
            st.dataframe(df_cmd.rename(columns={"Nb_magasins":"Mag. sans stock"}),
                         use_container_width=True, hide_index=True)

    with c2:
        st.markdown('<div class="sh">🔵 TOP APPROS À ACCÉLÉRER</div>', unsafe_allow_html=True)
        df_acc = merged[merged["Alerte"] == "🔵 Appro en cours"].groupby("SKU").agg(
            Libellé=("Libellé article","first"),
            Fournisseur=("Fournisseur T1","first"),
            Origine=("Origine","first"),
            Nb_magasins=("Magasin","count"),
            RAL_total=("Ral","sum"),
        ).reset_index().sort_values("Nb_magasins", ascending=False).head(10)
        if df_acc.empty:
            st.success("✅ Aucune appro en attente")
        else:
            st.dataframe(df_acc.rename(columns={"Nb_magasins":"Mag. en attente","RAL_total":"RAL total"}),
                         use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# VUE RÉSEAU
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[1]:
    c1, c2 = st.columns([3, 2])
    with c1:
        # Barres empilées par magasin (alertes actionables)
        alertes_aff = [a for a in ALERTES if a in pivot_mag.columns and a != "✅ Implanté"]
        mel = pivot_mag.melt(
            id_vars="Magasin",
            value_vars=["✅ Implanté"] + alertes_aff,
            var_name="Alerte", value_name="N"
        )
        fig = px.bar(mel, x="Magasin", y="N", color="Alerte",
                     color_discrete_map=ALERTES, barmode="stack",
                     title="Situation par magasin")
        fig.update_traces(textposition="inside", texttemplate="%{y}",
                          textfont=dict(size=10, color="white"))
        fig.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=400,
                          font=dict(family="Inter", color=C["muted"], size=12),
                          margin=dict(l=10,r=10,t=44,b=20),
                          legend=dict(orientation="h", y=-0.3, bgcolor="rgba(0,0,0,0)"),
                          xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        labels = list(ALERTES.keys())
        vals   = [int(pivot_mag.get(a, pd.Series([0])).sum()) if a in pivot_mag.columns else 0
                  for a in labels]
        fig_d = go.Figure(go.Pie(
            labels=labels, values=vals, hole=0.62,
            marker=dict(colors=list(ALERTES.values()),
                        line=dict(color="#fff", width=3))
        ))
        fig_d.add_annotation(
            text=f"<b>{taux_reseau}%</b><br>implanté",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=20, color=C["text"], family="Inter")
        )
        fig_d.update_layout(paper_bgcolor=C["surface"], height=400,
                            font=dict(family="Inter", color=C["muted"], size=12),
                            margin=dict(l=10,r=10,t=44,b=20),
                            legend=dict(orientation="v", x=1.01),
                            title="Répartition réseau")
        st.plotly_chart(fig_d, use_container_width=True)

    # Matrice rayon
    if not rayon_pivot.empty:
        st.markdown('<div class="sh">TAUX PAR RAYON × MAGASIN</div>', unsafe_allow_html=True)
        def color_cell(val):
            if val >= 80: return "background-color:#D1FAE5;color:#065F46;font-weight:700"
            if val >= 50: return "background-color:#FEF3C7;color:#92400E;font-weight:700"
            return "background-color:#FEE2E2;color:#991B1B;font-weight:700"
        st.dataframe(rayon_pivot.style.map(color_cell).format("{}%"),
                     use_container_width=True)

    # Flux IM/LO
    df_flux = merged[merged["Alerte"]=="✅ Implanté"].groupby(["Magasin","Origine"]).size().reset_index(name="N")
    if not df_flux.empty:
        st.markdown('<div class="sh">FLUX IM / LO — ARTICLES IMPLANTÉS</div>', unsafe_allow_html=True)
        fig_flux = px.bar(df_flux, x="Magasin", y="N", color="Origine",
                          color_discrete_map={"IM":C["blue"],"LO":C["green"]},
                          barmode="group", title="Articles implantés par flux")
        fig_flux.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=280,
                               font=dict(family="Inter", color=C["muted"], size=12),
                               margin=dict(l=10,r=10,t=44,b=20),
                               legend=dict(orientation="h",y=-0.3),
                               xaxis=dict(gridcolor=C["bg"]),yaxis=dict(gridcolor=C["bg"]))
        st.plotly_chart(fig_flux, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# ALERTES
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[2]:
    st.markdown('<div class="sh">TOUTES LES ALERTES AVEC ACTION</div>', unsafe_allow_html=True)

    # Filtre alerte
    alertes_dispo = [a for a in ALERTES if a != "✅ Implanté" and
                     (merged["Alerte"] == a).any()]
    alerte_sel = st.multiselect(
        "Filtrer par alerte", alertes_dispo, default=alertes_dispo,
        format_func=lambda a: f"{a} — {ACTION_LABEL[a]}"
    )

    mag_alerte_sel = st.multiselect(
        "Filtrer par magasin", sorted(merged["Magasin"].unique()),
        default=sorted(merged["Magasin"].unique())
    )

    df_al = merged[
        merged["Alerte"].isin(alerte_sel) &
        merged["Magasin"].isin(mag_alerte_sel)
    ].copy()

    if df_al.empty:
        st.markdown('<div class="info-box green">✅ Aucune alerte pour les filtres sélectionnés.</div>', unsafe_allow_html=True)
    else:
        # Résumé par alerte
        for alerte in alerte_sel:
            df_a = df_al[df_al["Alerte"] == alerte]
            if df_a.empty: continue
            color = ALERTES.get(alerte, C["muted"])
            action = ACTION_LABEL.get(alerte, "")
            st.markdown(f"""
            <div style="background:{color}18;border:1px solid {color}44;border-left:4px solid {color};
                        border-radius:10px;padding:10px 16px;margin-bottom:8px;
                        display:flex;align-items:center;justify-content:space-between;">
              <div>
                <span style="font-size:14px;font-weight:700;color:{C['text']}">{alerte}</span>
                <span style="font-size:12px;color:{C['muted']};margin-left:12px;">→ {action}</span>
              </div>
              <span style="font-size:22px;font-weight:800;color:{color}">{fmt_n(len(df_a))}</span>
            </div>""", unsafe_allow_html=True)

        # Tableau principal
        COLS_AL = ["Magasin","SKU","Libellé article","Origine","Code etat",
                   "Nouveau stock","Ral","Mode Appro","Sem. Réception",
                   "Fournisseur T1","Alerte","Action"]
        st.dataframe(
            df_al[[c for c in COLS_AL if c in df_al.columns]]
            .sort_values(["Alerte","Magasin"]).reset_index(drop=True),
            use_container_width=True, hide_index=True
        )

        # Ruptures totales réseau (non implanté sur tous les magasins)
        df_non_impl = merged[merged["Alerte"].isin([
            "🛒 Passer commande","🔵 Appro en cours","🔧 Régulariser","🔧 Régulariser + appro"
        ])]
        sku_counts = df_non_impl.groupby("SKU")["Magasin"].count()
        sku_rupt   = sku_counts[sku_counts == n_mag].index.tolist()

        if sku_rupt:
            st.markdown(f"""
            <div class="info-box orange" style="margin-top:14px;">
              🚨 <strong>{len(sku_rupt)} article(s) en alerte sur TOUS les magasins</strong> — escalade critique recommandée.
            </div>""", unsafe_allow_html=True)
            df_rupt = (merged[merged["SKU"].isin(sku_rupt)]
                       [["SKU","Libellé article","Origine","Fournisseur T1","Alerte"]]
                       .drop_duplicates("SKU").sort_values("Alerte").reset_index(drop=True))
            st.dataframe(df_rupt, use_container_width=True, hide_index=True)

        # Anomalies référencement
        if n_anomalie > 0 and "🚩 Anomalie référ." in alerte_sel:
            st.markdown('<div class="sh">🚩 ANOMALIES RÉFÉRENCEMENT — DÉTAIL</div>', unsafe_allow_html=True)
            df_an_agg = (
                merged[merged["Alerte"]=="🚩 Anomalie référ."]
                .groupby("SKU").agg(
                    Libellé=("Libellé article","first"),
                    Fournisseur=("Fournisseur T1","first"),
                    Origine=("Origine","first"),
                    Nb_magasins=("Magasin","count"),
                    Codes_etat=("Code etat", lambda x: ", ".join(sorted(set(x.dropna().astype(str))))),
                ).reset_index()
            )
            df_an_agg["Criticité"] = df_an_agg.apply(
                lambda r: r["Nb_magasins"] * (2 if r["Origine"] == "IM" else 1), axis=1
            )
            df_an_agg["Signification"] = df_an_agg["Codes_etat"].apply(
                lambda x: " · ".join([f"{e}={ETAT_LABEL.get(e,e)}" for e in x.split(", ")])
            )
            st.dataframe(
                df_an_agg[["SKU","Libellé","Origine","Fournisseur","Nb_magasins","Signification","Criticité"]]
                .sort_values("Criticité", ascending=False).reset_index(drop=True),
                use_container_width=True, hide_index=True
            )

# ══════════════════════════════════════════════════════════════════════════════
# CESSIONS
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[3]:
    st.markdown('<div class="sh">🔄 MOTEUR CESSIONS INTER-MAGASINS</div>', unsafe_allow_html=True)

    # Suggestions automatiques ruptures totales
    df_non_impl_all = merged[merged["Alerte"].isin([
        "🛒 Passer commande","🔵 Appro en cours"
    ])]
    sku_rupt_tot = (df_non_impl_all.groupby("SKU")["Magasin"].count()
                    .pipe(lambda s: s[s == n_mag]).index.tolist())

    if sku_rupt_tot:
        st.markdown(f'<div class="info-box blue">🤖 <strong>{len(sku_rupt_tot)} article(s)</strong> sans stock sur tous les magasins — suggestions automatiques de cession activées.</div>', unsafe_allow_html=True)
        auto_sugg = []
        for sku in sku_rupt_tot:
            sku_df = df_stock[df_stock["SKU"] == sku].copy()
            sku_df["Nouveau stock"] = pd.to_numeric(sku_df["Nouveau stock"], errors="coerce").fillna(0)
            sku_df["Pcb"] = pd.to_numeric(sku_df.get("Pcb", 1), errors="coerce").fillna(1).clip(lower=1)
            best = sku_df.copy()
            best["Reserve_2pcb"] = (best["Pcb"] * 2).astype(int)
            best = best[best["Nouveau stock"] > best["Reserve_2pcb"]].sort_values("Nouveau stock", ascending=False)
            lib  = sku_df["Libellé article"].iloc[0] if len(sku_df) else sku
            if not best.empty:
                b   = best.iloc[0]
                reserve_art = int(b["Reserve_2pcb"])
                qty = int(b["Nouveau stock"]) - reserve_art
                auto_sugg.append({
                    "SKU": sku, "Libellé": lib,
                    "Cédant": site_ref.get(b["Code site"], b["Code site"]),
                    "Stock cédant": int(b["Nouveau stock"]),
                    "Qté cessible": qty,
                    "Type": "🤖 Suggestion auto"
                })
        if auto_sugg:
            st.dataframe(pd.DataFrame(auto_sugg), use_container_width=True, hide_index=True)
        else:
            st.info("Aucun magasin cédant disponible pour les ruptures totales.")

    if not mag_detresse:
        st.markdown('<div class="info-box blue">⬅️ Sélectionne des magasins en détresse dans la sidebar pour le moteur de cessions manuel.</div>', unsafe_allow_html=True)
    else:
        mag_det_codes = [c for c in all_codes if site_ref.get(c,c) in mag_detresse]
        mag_ced_codes = [c for c in mag_sel_codes if c not in mag_det_codes]
        suggestions   = []

        for sku in sku_scope:
            sku_df = df_stock[df_stock["SKU"] == sku].copy()
            if sku_df.empty: continue
            lib = sku_df["Libellé article"].iloc[0] if "Libellé article" in sku_df.columns else sku
            sku_df["Nouveau stock"] = pd.to_numeric(sku_df["Nouveau stock"], errors="coerce").fillna(0)
            sku_df["Pcb"] = pd.to_numeric(sku_df.get("Pcb", 1), errors="coerce").fillna(1).clip(lower=1)

            det_rows = sku_df[sku_df["Code site"].isin(mag_det_codes) &
                              (sku_df["Nouveau stock"] <= seuil_det)]
            if det_rows.empty: continue

            # Réserve cédant = 2 PCB par article
            sku_df["Reserve_2pcb"] = (sku_df["Pcb"] * 2).astype(int)
            ced_rows = sku_df[sku_df["Code site"].isin(mag_ced_codes)].copy()
            ced_rows = ced_rows[ced_rows["Nouveau stock"] > ced_rows["Reserve_2pcb"]].sort_values("Nouveau stock", ascending=False)

            for _, dr in det_rows.iterrows():
                if ced_rows.empty:
                    suggestions.append({
                        "SKU": sku, "Libellé article": lib,
                        "Magasin détresse": site_ref.get(dr["Code site"], dr["Code site"]),
                        "Stock détresse": int(dr["Nouveau stock"]),
                        "Cédant suggéré": "⚠️ Aucun cédant",
                        "Stock cédant": 0, "Qté cessible": 0,
                        "Réserve (2 PCB)": 0,
                        "Faisabilité": "🔴 Impossible"
                    })
                else:
                    best        = ced_rows.iloc[0]
                    reserve_art = int(best["Reserve_2pcb"])
                    qty         = int(best["Nouveau stock"]) - reserve_art
                    suggestions.append({
                        "SKU": sku, "Libellé article": lib,
                        "Magasin détresse": site_ref.get(dr["Code site"], dr["Code site"]),
                        "Stock détresse": int(dr["Nouveau stock"]),
                        "Cédant suggéré": site_ref.get(best["Code site"], best["Code site"]),
                        "Stock cédant": int(best["Nouveau stock"]),
                        "Réserve (2 PCB)": reserve_art,
                        "Qté cessible": qty,
                        "Faisabilité": "🟢 Possible" if qty >= 1 else "🟠 Partielle"
                    })

        if not suggestions:
            st.success("✅ Aucune cession nécessaire selon les critères.")
        else:
            df_all  = pd.DataFrame(suggestions)
            n_poss  = int((df_all["Faisabilité"]=="🟢 Possible").sum())
            n_imp   = int((df_all["Faisabilité"]=="🔴 Impossible").sum())

            df_cess = df_all[df_all["Faisabilité"]=="🟢 Possible"].copy()

            # Filtre minimum 1 PCB cessible
            if min_1pcb and "Réserve (2 PCB)" in df_cess.columns:
                # PCB = Réserve / 2
                df_cess = df_cess[df_cess["Qté cessible"] >= (df_cess["Réserve (2 PCB)"] / 2).clip(lower=1)]

            df_cess = df_cess.sort_values("Qté cessible", ascending=False).reset_index(drop=True)

            k1, k2, k3 = st.columns(3)
            k1.metric("🟢 Cessions possibles", n_poss)
            k2.metric("🔴 Impossible (masqué)", n_imp)
            k3.metric("Articles cessibles",     df_cess["SKU"].nunique() if not df_cess.empty else 0)

            st.dataframe(df_cess, use_container_width=True, hide_index=True)

            buf_c = io.BytesIO()
            with pd.ExcelWriter(buf_c, engine="openpyxl") as w:
                df_cess.to_excel(w, sheet_name="Plan Cessions", index=False)
            buf_c.seek(0)
            st.download_button(f"📥 Plan_Cessions_{TODAY_FILE}.xlsx", data=buf_c,
                               file_name=f"Plan_Cessions_{TODAY_FILE}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════════════════
elif active == TABS[4]:
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    st.markdown('<div class="info-box blue">3 feuilles : <strong>Synthèse réseau</strong> · <strong>Détail complet</strong> · <strong>Alertes & Actions</strong></div>', unsafe_allow_html=True)

    if st.button("🔨 Générer Export", type="primary"):
        ALERTE_FILLS = {
            "✅ Implanté"            : ("D1FAE5","065F46"),
            "🔵 Appro en cours"      : ("DBEAFE","1D4ED8"),
            "🛒 Passer commande"     : ("FEF3C7","92400E"),
            "🔧 Régulariser + appro" : ("FEE2E2","991B1B"),
            "🔧 Régulariser"         : ("FEE2E2","991B1B"),
            "🚩 Anomalie référ."     : ("FFFDE7","795548"),
            "⚪ Non référencé"       : ("F3F4F6","374151"),
        }
        COLS_DET = ["Magasin","SKU","Libellé article","Origine","Code etat",
                    "Nouveau stock","Ral","Mode Appro","Sem. Réception",
                    "Fournisseur T1","Alerte","Action"]

        buf_x = io.BytesIO()
        with pd.ExcelWriter(buf_x, engine="openpyxl") as writer:
            # Feuille 1 : Synthèse
            cols_s = ["Magasin"] + [c for c in ALERTES if c in pivot_mag.columns] + ["Taux (%)"]
            pivot_mag[cols_s].sort_values("Taux (%)", ascending=False).to_excel(
                writer, sheet_name="Synthèse Réseau", index=False)

            # Feuille 2 : Détail complet
            merged[[c for c in COLS_DET if c in merged.columns]].to_excel(
                writer, sheet_name="Détail Complet", index=False)

            # Feuille 3 : Alertes uniquement
            df_alertes_exp = merged[merged["Alerte"] != "✅ Implanté"]
            df_alertes_exp[[c for c in COLS_DET if c in df_alertes_exp.columns]].sort_values(
                ["Alerte","Magasin"]).to_excel(writer, sheet_name="Alertes & Actions", index=False)

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
                        v = str(cell.value)
                        if v in ALERTE_FILLS:
                            bg, fg = ALERTE_FILLS[v]
                            cell.fill = PatternFill("solid", fgColor=bg)
                            cell.font = Font(color=fg, name="Arial", size=10)

        buf_x.seek(0)
        st.download_button(
            f"📥 Implantation_T1_{TODAY_FILE}.xlsx", data=buf_x,
            file_name=f"Implantation_T1_{TODAY_FILE}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success(f"✅ Export — {fmt_n(len(merged))} lignes · 3 feuilles")

# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    f'<div style="text-align:center;font-size:11px;color:{C["muted"]};font-family:JetBrains Mono;">'
    f'SmartBuyer · NovaRetail Solutions · Implantation v6.0 · {TODAY_STR} · Données : {date_stock_str}'
    f'</div>',
    unsafe_allow_html=True
)
