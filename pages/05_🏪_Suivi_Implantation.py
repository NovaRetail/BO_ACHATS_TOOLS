"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  03_📦_Implantation.py  ·  SmartBuyer · NovaRetail Solutions               ║
║  Suivi Implantation + Supply Nouvelles Références — v5.0                    ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║  SOURCES DE DONNÉES                                                          ║
║  ① T1  — Liste nouvelles références (ERP)                                   ║
║       Colonnes : ARTICLE, MODE APPRO, SEMAINE RECEPTION,                    ║
║                  LIBELLÉ FOURNISSEUR ORIGINE, DATE LIV.                     ║
║                                                                              ║
║  ② Stock light consolidé — 1 fichier tous magasins                          ║
║       Format  : CSV latin1 séparateur ';'                                   ║
║       Nommage : stock_DDMM_light.csv (date extraite auto)                   ║
║       Colonnes utiles : Site, Libellé site, Code article,                   ║
║                         Libellé article, Nouveau stock, Ral,                ║
║                         Code etat, Code marketing (IM/LO),                  ║
║                         Nom fourn., Libellé rayon, Qté sortie,              ║
║                         Pcb, Date dernière entrée                           ║
║       Code etat : 2=Actif (analysé) · P=Purge (exclu)                      ║
║                   B/S/F/6/5=Anomalie (alerte)                               ║
║                                                                              ║
║  ③ RAL dédié — optionnel · multi-fichiers (un par magasin)                  ║
║       Format  : CSV latin1 séparateur ';'                                   ║
║       Nommage : C450_EXTR_RAL_DATE_SITE_GLOBAL.csv                         ║
║       Code magasin + date extraits automatiquement du nom de fichier        ║
║       Colonnes : Code art., RAL, Commande, Situation, Date Reception        ║
║       Situations : 38=En attente Supply · 40=En transit                     ║
║                    50=Réception en cours                                     ║
║       Active la Partie Supply avec : prochaine date livraison,              ║
║                                      nb commandes, retards                  ║
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

# Colonnes à lire dans le stock light (usecols)
STOCK_COLS = [
    "Site", "Libellé site", "Code article", "Libellé article",
    "Nouveau stock", "Ral", "Code etat", "Code marketing",
    "Nom fourn.", "Libellé rayon", "Libellé famille",
    "Qté sortie", "Pcb", "Date dernière entrée"
]

# Code etat
ETAT_ACTIF    = "2"
ETAT_PURGE    = "P"
ETAT_ANOMALIE = {"B", "S", "F", "6", "5", "1"}

ETAT_LABEL = {
    "B": "Rayon générique",
    "S": "Suspendu",
    "F": "Fin de vie",
    "6": "Déréférencé",
    "5": "Autre",
    "1": "Autre",
}

# Charte SmartBuyer
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

STATUT_COLORS = {
    "✅ Implanté"            : C["green"],
    "🔴 Stock négatif"       : C["red"],
    "⚠️ Non implanté"        : C["orange"],
    "🚩 Anomalie référ."     : C["yellow"],
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
.topbar-title{{font-size:17px;font-weight:700;color:#fff;letter-spacing:-.01em;}}
.topbar-sub{{font-size:11px;color:#8E8E93;font-family:'JetBrains Mono';margin-top:2px;}}
.topbar-pill{{background:rgba(255,255,255,.08);color:#8E8E93;border:1px solid rgba(255,255,255,.12);border-radius:8px;padding:4px 14px;font-size:11px;font-weight:600;}}
.topbar-date{{color:{C['blue']};font-size:12px;font-family:'JetBrains Mono';}}

/* INTRO */
.module-intro{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px 24px;margin-bottom:24px;box-shadow:var(--shadow);}}
.module-intro-title{{font-size:15px;font-weight:700;color:var(--text);margin-bottom:6px;}}
.module-intro-body{{font-size:13px;color:var(--muted);line-height:1.6;}}
.tag-badge{{display:inline-flex;align-items:center;gap:5px;background:var(--blue-l);color:var(--blue);border:1px solid #BFDBFE;border-radius:6px;padding:3px 10px;font-size:11px;font-weight:600;margin:2px;}}
.tag-badge.green{{background:var(--green-l);color:var(--green);border-color:#6EE7B7;}}
.tag-badge.red{{background:var(--red-l);color:var(--red);border-color:#FECACA;}}
.tag-badge.orange{{background:var(--orange-l);color:var(--orange);border-color:#FCD34D;}}
.tag-badge.purple{{background:var(--purple-l);color:var(--purple);border-color:#D8B4FE;}}

/* FORMAT DOC BOX */
.format-box{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:16px 20px;margin-bottom:16px;box-shadow:var(--shadow);}}
.format-box-title{{font-size:12px;font-weight:700;color:var(--text);margin-bottom:8px;display:flex;align-items:center;gap:6px;}}
.format-col{{display:inline-flex;align-items:center;gap:4px;background:var(--bg);border:1px solid var(--border);border-radius:5px;padding:2px 8px;font-size:11px;font-family:'JetBrains Mono';color:var(--text);margin:2px;}}
.format-col.key{{background:var(--blue-l);border-color:#BFDBFE;color:{C['blue']};}}
.format-col.info{{background:var(--green-l);border-color:#6EE7B7;color:{C['green']};}}

/* SH */
.sh{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:var(--muted);margin:24px 0 14px;padding-bottom:8px;border-bottom:1px solid var(--border);}}
.sh-supply{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:var(--purple);margin:24px 0 14px;padding-bottom:8px;border-bottom:2px solid var(--purple);}}
.sh-copil{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:{C['teal']};margin:24px 0 14px;padding-bottom:8px;border-bottom:2px solid {C['teal']};}}

/* PARTIE LABELS */
.partie-label{{display:inline-flex;align-items:center;gap:8px;background:var(--text);color:#fff;border-radius:8px;padding:6px 16px;font-size:12px;font-weight:700;letter-spacing:.05em;margin-bottom:16px;}}
.partie-label.supply{{background:linear-gradient(135deg,{C['purple']},{C['blue']});}}
.partie-label.copil{{background:linear-gradient(135deg,{C['teal']},{C['blue']});}}

/* KPI */
.kpi-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:24px;}}
.kpi-grid-4{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:24px;}}
.kpi-grid-3{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:16px;}}
.kpi-card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:18px 20px 14px;box-shadow:var(--shadow);position:relative;overflow:hidden;}}
.kpi-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:var(--radius) var(--radius) 0 0;}}
.kpi-card.blue::before{{background:var(--blue);}} .kpi-card.green::before{{background:var(--green);}}
.kpi-card.red::before{{background:var(--red);}}   .kpi-card.orange::before{{background:var(--orange);}}
.kpi-card.purple::before{{background:var(--purple);}} .kpi-card.teal::before{{background:var(--teal);}}
.kpi-card.yellow::before{{background:var(--yellow);}}
.kpi-label{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.10em;color:var(--muted);margin-bottom:10px;}}
.kpi-value{{font-size:36px;font-weight:800;line-height:1;letter-spacing:-.02em;}}
.kpi-card.blue .kpi-value{{color:var(--blue);}} .kpi-card.green .kpi-value{{color:var(--green);}}
.kpi-card.red .kpi-value{{color:var(--red);}}   .kpi-card.orange .kpi-value{{color:var(--orange);}}
.kpi-card.purple .kpi-value{{color:var(--purple);}} .kpi-card.teal .kpi-value{{color:var(--teal);}}
.kpi-card.yellow .kpi-value{{color:#B8860B;}}
.kpi-sub{{font-size:11px;color:var(--muted);font-family:'JetBrains Mono';margin-top:4px;}}
.kpi-bar{{margin-top:12px;height:3px;border-radius:3px;background:var(--border);}}
.kpi-bar-fill{{height:100%;border-radius:3px;}}

/* SCORECARD */
.scorecard-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:10px;margin-bottom:24px;}}
.scorecard-card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px;box-shadow:var(--shadow);position:relative;}}
.scorecard-card.ok{{border-color:#6EE7B7;background:var(--green-l);}}
.scorecard-card.warn{{border-color:#FCD34D;background:var(--orange-l);}}
.scorecard-card.ko{{border-color:#FECACA;background:var(--red-l);}}
.scorecard-dot{{width:8px;height:8px;border-radius:50%;position:absolute;top:14px;right:14px;}}
.scorecard-name{{font-size:11px;font-weight:600;color:var(--text);margin-bottom:6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:88%;}}
.scorecard-pct{{font-size:28px;font-weight:800;line-height:1;}}
.scorecard-sub{{font-size:10px;color:var(--muted);font-family:'JetBrains Mono';margin-top:3px;}}

/* ALERT BANNERS */
.alert-banner{{background:#FFF;border:1px solid #FECACA;border-left:4px solid var(--red);border-radius:var(--radius);padding:16px 20px;margin-bottom:20px;display:flex;align-items:center;gap:16px;flex-wrap:wrap;}}
.alert-pill{{background:var(--red);color:#fff;border-radius:6px;padding:4px 12px;font-size:11px;font-weight:700;letter-spacing:.05em;white-space:nowrap;}}
.alert-item{{display:flex;flex-direction:column;align-items:center;padding:0 16px;border-right:1px solid var(--border);}}
.alert-item:last-child{{border-right:none;}}
.alert-num{{font-size:26px;font-weight:800;line-height:1;}}
.alert-lbl{{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-top:1px;}}

/* INFO BOXES */
.info-box{{border-radius:var(--radius);padding:14px 18px;margin-bottom:16px;border:1px solid;font-size:13px;line-height:1.6;}}
.info-box.blue{{background:var(--blue-l);border-color:#BFDBFE;color:#1D4ED8;}}
.info-box.green{{background:var(--green-l);border-color:#6EE7B7;color:#065F46;}}
.info-box.orange{{background:var(--orange-l);border-color:#FCD34D;color:#92400E;}}
.info-box.purple{{background:var(--purple-l);border-color:#D8B4FE;color:#6B21A8;}}
.info-box.yellow{{background:#FFFDE7;border-color:#FDD835;color:#795548;}}

/* VALIDATION BOX */
.validation-box{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 20px;margin-bottom:20px;box-shadow:var(--shadow);display:flex;align-items:center;gap:20px;flex-wrap:wrap;}}
.vbox-item{{display:flex;flex-direction:column;align-items:center;padding:0 16px;border-right:1px solid var(--border);}}
.vbox-item:last-child{{border-right:none;}}
.vbox-num{{font-size:22px;font-weight:800;line-height:1;}}
.vbox-lbl{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-top:2px;}}

/* SUPPLY DIVIDER */
.supply-divider{{background:linear-gradient(135deg,{C['purple']}22,{C['blue']}22);border:1px solid {C['purple']}44;border-radius:var(--radius);padding:16px 20px;margin:32px 0 20px;display:flex;align-items:center;gap:12px;}}
.supply-divider-icon{{width:36px;height:36px;border-radius:9px;background:linear-gradient(135deg,{C['purple']},{C['blue']});display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0;}}
.supply-divider-title{{font-size:15px;font-weight:700;color:var(--text);}}
.supply-divider-sub{{font-size:12px;color:var(--muted);margin-top:2px;}}

/* NAV */
.nav-active{{background:var(--text)!important;color:#fff!important;border-radius:10px;padding:10px 0;text-align:center;font-size:13px;font-weight:700;box-shadow:0 4px 14px rgba(28,28,30,.2);margin-bottom:10px;}}
.nav-active-supply{{background:linear-gradient(135deg,{C['purple']},{C['blue']})!important;color:#fff!important;border-radius:10px;padding:10px 0;text-align:center;font-size:13px;font-weight:700;box-shadow:0 4px 14px rgba(175,82,222,.3);margin-bottom:10px;}}

/* SIDEBAR */
section[data-testid="stSidebar"]{{background:#fff!important;border-right:1px solid var(--border)!important;min-width:270px!important;max-width:270px!important;}}
section[data-testid="stSidebar"] .block-container{{padding:.6rem .8rem 2rem!important;}}
.stDownloadButton>button{{background:linear-gradient(135deg,{C['text']},{C['blue']})!important;color:#fff!important;border:none!important;border-radius:10px!important;font-weight:700!important;font-size:13px!important;padding:12px!important;width:100%!important;box-shadow:0 4px 12px rgba(0,122,255,.25)!important;}}

/* COPIL CARD */
.copil-card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:16px 20px;box-shadow:var(--shadow);margin-bottom:12px;}}
.copil-title{{font-size:13px;font-weight:700;color:var(--text);margin-bottom:10px;}}

/* MATRICE */
.matrix-cell{{text-align:center;font-size:12px;font-weight:700;border-radius:6px;padding:4px 8px;}}
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

def fmt_pct(v: float) -> str:
    return f"{v:.1f}%"

def color_taux(t: float) -> str:
    if t >= 80: return C["green"]
    if t >= 50: return C["orange"]
    return C["red"]

def scorecard_cls(t: float) -> str:
    if t >= 80: return "ok"
    if t >= 50: return "warn"
    return "ko"

def extract_date_from_filename(filename: str) -> pd.Timestamp:
    """Extrait date depuis nom fichier. stock_2504_light → 25/04/2026."""
    # Format YYYYMMDD
    m = re.search(r'(\d{8})', filename)
    if m:
        for fmt in ('%Y%m%d', '%d%m%Y'):
            try:
                return pd.Timestamp(pd.to_datetime(m.group(1), format=fmt))
            except Exception:
                continue
    # Format DDMM (ex: 2504)
    m2 = re.search(r'_(\d{4})_', filename)
    if m2:
        s = m2.group(1)
        try:
            d, mo = int(s[:2]), int(s[2:])
            yr = date.today().year
            return pd.Timestamp(datetime(yr, mo, d))
        except Exception:
            pass
    return TODAY

def extract_site_code(filename: str) -> str:
    m = re.search(r'(1\d{4}|2\d{4})', filename)
    return m.group(1) if m else None

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

def _sem_sort(s) -> int:
    cleaned = re.sub(r"[Ss]", "", str(s).strip())
    return int(cleaned) if cleaned.isdigit() else 99

def couverture_jours(stock, qte_sortie, nb_jours=30) -> float:
    """Couverture en jours = stock / (qte_sortie / nb_jours)."""
    try:
        qs = float(qte_sortie)
        if qs <= 0: return 999.0
        return round(float(stock) / (qs / nb_jours), 1)
    except Exception:
        return 999.0

def badge_couverture(j: float) -> str:
    if j >= 999: return "— j"
    if j < 7:    return f"🔴 {j}j"
    if j < 15:   return f"🟠 {j}j"
    return f"🟢 {j}j"


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
    df["SEM_NUM"]  = df["SEMAINE RECEPTION"].apply(_sem_to_num)
    df["ORIGINE"]  = df["MODE APPRO"].apply(
        lambda m: "IM" if "IMPORT" in str(m).upper() else "LO"
    )
    return df, None


@st.cache_data(show_spinner=False)
def parse_stock_light(file_bytes: bytes, filename: str, sku_scope: tuple):
    """
    Parse le fichier stock light consolidé.
    - usecols : 14 colonnes utiles uniquement
    - dtype forcés pour performance
    - Filtre P (purge) silencieux dès la lecture
    - Filtre SKU scope T1 immédiat
    """
    buf = io.BytesIO(file_bytes)
    try:
        df = pd.read_csv(
            buf, sep=";", encoding="latin1", low_memory=False,
            on_bad_lines="skip",
            usecols=STOCK_COLS,
            dtype={
                "Site"          : str,
                "Code article"  : str,
                "Code etat"     : str,
                "Code marketing": str,
                "Libellé site"  : str,
                "Libellé article": str,
                "Nom fourn."    : str,
                "Libellé rayon" : str,
                "Libellé famille": str,
            }
        )
    except Exception as e:
        return None, f"Lecture stock : {e}"

    # Nettoyage
    for col in ["Libellé site", "Libellé article", "Nom fourn.",
                "Libellé rayon", "Code etat", "Code marketing"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r"\t+", "", regex=True)

    # Normaliser SKU
    df["SKU"] = (df["Code article"].astype(str).str.strip()
                 .str.replace(r"\.0$", "", regex=True).str.zfill(8).str[:8])

    # Filtrer P (purge) silencieusement
    df = df[df["Code etat"] != ETAT_PURGE].copy()

    # Filtrer uniquement SKUs T1 — gain mémoire majeur
    if sku_scope:
        df = df[df["SKU"].isin(sku_scope)].copy()

    # Typage numériques
    df["Nouveau stock"] = pd.to_numeric(df["Nouveau stock"], errors="coerce")
    df["Ral"]           = pd.to_numeric(df["Ral"],           errors="coerce").fillna(0).astype(int)
    df["Qté sortie"]    = pd.to_numeric(df["Qté sortie"],    errors="coerce").fillna(0)
    df["Pcb"]           = pd.to_numeric(df["Pcb"],           errors="coerce").fillna(1)

    # Code site depuis colonne Site
    df["Code site"] = df["Site"].astype(str).str.strip()

    # Flux depuis Code marketing
    df["Origine"] = df["Code marketing"].apply(
        lambda m: "IM" if str(m).strip().upper() == "IM" else "LO"
    )

    return df, None


@st.cache_data(show_spinner=False)
def parse_ral(files_bytes: list, filenames: list):
    """
    Parse N fichiers RAL mono-magasin.
    Agrège par SKU × Code site :
      - Prochaine livraison (min Date Reception)
      - Nb commandes (nunique Commande)
      - Retard max (jours)
    """
    frames, errors = [], []

    for fb, fn in zip(files_bytes, filenames):
        code_site       = extract_site_code(fn)
        date_extraction = extract_date_from_filename(fn)

        if not code_site:
            errors.append(f"⚠️ Code magasin introuvable dans '{fn}'")
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

        df["Situation"]      = pd.to_numeric(df["Situation"], errors="coerce")
        df                   = df[df["Situation"].isin([38, 40, 50])].copy()
        df["Date Reception"] = pd.to_datetime(df["Date Reception"],
                                              format="%d/%m/%Y", errors="coerce")
        df["RAL"]            = pd.to_numeric(df["RAL"], errors="coerce").fillna(0)
        df["En retard"]      = df["Date Reception"] < date_extraction
        df["Jours retard"]   = (date_extraction - df["Date Reception"]).dt.days.clip(lower=0)
        df["Jours retard"]   = df["Jours retard"].where(df["En retard"], 0)

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

    agg["Prochaine_livraison_fmt"] = agg["Prochaine_livraison"].dt.strftime("%d/%m/%Y")
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
      <div class="topbar-sub">T1 · Stock ERP light · RAL Livraisons · Multi-magasins · v5.0</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:12px;">
    <div class="topbar-date">{TODAY_STR}</div>
    <div class="topbar-pill">v5.0 · SmartBuyer</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# INTRO MÉTIER + STRUCTURE FICHIERS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="module-intro">
  <div class="module-intro-title">📦 À quoi sert ce module ?</div>
  <div class="module-intro-body">
    Ce module suit l'implantation des <strong>nouvelles références T1</strong> dans le réseau Carrefour CI,
    depuis la commande fournisseur jusqu'à la mise en rayon. Il croise la liste T1 avec le stock ERP
    et les commandes en attente (RAL) pour une <strong>vision complète en temps réel</strong> :
    où en est chaque article, dans chaque magasin, et quelle action prendre.<br><br>
    <span class="tag-badge">📋 Taux d'implantation réseau</span>
    <span class="tag-badge green">✅ Suivi mise en rayon</span>
    <span class="tag-badge orange">🛒 Articles à commander</span>
    <span class="tag-badge purple">🚚 Suivi livraisons & retards</span>
    <span class="tag-badge red">🔧 Régularisations inventaire</span>
  </div>
</div>
""", unsafe_allow_html=True)

with st.expander("📋 Structure des fichiers attendus", expanded=False):
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""
        <div class="format-box">
          <div class="format-box-title">① T1 — Nouvelles Références</div>
          <div>Format : CSV ou Excel</div><br>
          <span class="format-col key">ARTICLE</span>
          <span class="format-col key">MODE APPRO</span>
          <span class="format-col key">SEMAINE RECEPTION</span>
          <span class="format-col">LIBELLÉ FOURNISSEUR</span>
          <span class="format-col">DATE LIV.</span><br><br>
          <div style="font-size:11px;color:#6D6D72;">
            Code article 8 chiffres · MODE APPRO contient "IMPORT" pour IM
          </div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class="format-box">
          <div class="format-box-title">② Stock light consolidé</div>
          <div>Format : CSV latin1 · sep ';'</div>
          <div style="font-size:11px;color:#6D6D72;margin-bottom:6px;">Nommage : stock_DDMM_light.csv</div>
          <span class="format-col key">Site</span>
          <span class="format-col key">Code article</span>
          <span class="format-col key">Nouveau stock</span>
          <span class="format-col key">Code etat</span>
          <span class="format-col info">Code marketing</span>
          <span class="format-col">Ral</span>
          <span class="format-col">Libellé site</span>
          <span class="format-col">Nom fourn.</span>
          <span class="format-col">Libellé rayon</span>
          <span class="format-col">Qté sortie</span><br><br>
          <div style="font-size:11px;color:#6D6D72;">
            Code etat : <strong>2</strong>=Actif · <strong>P</strong>=Purge (exclu) · B/S/F/6/5=Alerte<br>
            Code marketing : IM ou LO (flux approvisionnement)
          </div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown("""
        <div class="format-box">
          <div class="format-box-title">③ RAL Livraisons <span style="font-weight:400;font-size:11px;color:#6D6D72">(optionnel)</span></div>
          <div>Format : CSV latin1 · sep ';'</div>
          <div style="font-size:11px;color:#6D6D72;margin-bottom:6px;">Nommage : C450_EXTR_RAL_DATE_SITE_GLOBAL.csv</div>
          <span class="format-col key">Code art.</span>
          <span class="format-col key">RAL</span>
          <span class="format-col key">Commande</span>
          <span class="format-col key">Situation</span>
          <span class="format-col key">Date Reception</span><br><br>
          <div style="font-size:11px;color:#6D6D72;">
            Situation : <strong>38</strong>=En attente Supply · <strong>40</strong>=En transit · <strong>50</strong>=Réception en cours<br>
            Active la Partie Supply avec date livraison, nb commandes, retards
          </div>
        </div>
        """, unsafe_allow_html=True)

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
    st.markdown("**③ RAL Livraisons** *(optionnel)*")
    st.caption("Un fichier par magasin · active la Partie Supply")
    ral_files = st.file_uploader("RAL", type=["csv"], key="ral",
                                 label_visibility="collapsed",
                                 accept_multiple_files=True)

# ─────────────────────────────────────────────────────────────────────────────
# GATES
# ─────────────────────────────────────────────────────────────────────────────
if not t1_file:
    st.markdown('<div class="info-box blue">⬆️ <strong>Étape 1</strong> — Charge le fichier T1 (Nouvelles Références) dans la sidebar.</div>', unsafe_allow_html=True)
    st.stop()

with st.spinner("Lecture T1…"):
    t1_df, t1_err = parse_t1(t1_file.read(), t1_file.name)

if t1_err or t1_df is None:
    st.error(f"❌ T1 : {t1_err}")
    st.stop()

if not stk_file:
    st.markdown(f'<div class="info-box blue">✅ T1 chargé — <strong>{len(t1_df):,}</strong> références. ⬆️ <strong>Étape 2</strong> — Charge le fichier Stock light consolidé.</div>', unsafe_allow_html=True)
    st.stop()

# Date du stock depuis le nom du fichier
date_stock    = extract_date_from_filename(stk_file.name)
date_stock_str = date_stock.strftime("%d %b %Y")
age_stock     = (TODAY - date_stock).days

# Parser T1 d'abord pour avoir le scope SKU
SKU_SCOPE = tuple(sorted(t1_df["SKU"].unique()))

with st.spinner("Parsing Stock ERP…"):
    stk_bytes = stk_file.read()
    df_stock, stk_err = parse_stock_light(stk_bytes, stk_file.name, SKU_SCOPE)

if stk_err or df_stock is None:
    st.error(f"❌ Stock : {stk_err}")
    st.stop()

# RAL optionnel
df_ral_agg, ral_errors, supply_active = None, [], False
if ral_files:
    with st.spinner(f"Parsing RAL ({len(ral_files)} fichier(s))…"):
        ral_bytes_list = [f.read() for f in ral_files]
        ral_names_list = [f.name for f in ral_files]
        df_ral_agg, ral_errors = parse_ral(ral_bytes_list, ral_names_list)
    for e in ral_errors:
        st.warning(e)
    if df_ral_agg is not None:
        supply_active = True

# ─────────────────────────────────────────────────────────────────────────────
# RÉFÉRENTIEL MAGASINS
# ─────────────────────────────────────────────────────────────────────────────
site_ref = (df_stock[["Code site", "Libellé site"]]
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
    seuil_det    = st.number_input("Seuil stock (≤)", 0, 50, 0, 1)
    reserve      = st.number_input("Réserve cédant (≥)", 0, 50, 2, 1)

if not mag_sel_codes:
    st.warning("⚠️ Sélectionne au moins un magasin.")
    st.stop()

# Alerte données anciennes
if age_stock > 7:
    st.warning(f"⚠️ Données stock datées du **{date_stock_str}** — {age_stock} jours. Pense à recharger un fichier plus récent.")

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
sku_dans_stock  = set(df_stock[df_stock["Code site"].isin(mag_sel_codes)]["SKU"].unique())
sku_trouves     = [s for s in sku_scope if s in sku_dans_stock]
sku_absents     = [s for s in sku_scope if s not in sku_dans_stock]
n_sku           = len(sku_scope)
n_sku_trouves   = len(sku_trouves)
n_sku_absents   = len(sku_absents)
n_mag           = len(mag_sel_codes)

st.markdown(f"""
<div class="validation-box">
  <div style="font-size:13px;font-weight:700;color:{C['text']};margin-right:8px;">📋 Validation T1 × Stock</div>
  <div class="vbox-item">
    <div class="vbox-num" style="color:{C['blue']}">{fmt_n(n_sku)}</div>
    <div class="vbox-lbl">SKUs T1</div>
  </div>
  <div class="vbox-item">
    <div class="vbox-num" style="color:{C['green']}">{fmt_n(n_sku_trouves)}</div>
    <div class="vbox-lbl">Trouvés stock</div>
  </div>
  <div class="vbox-item">
    <div class="vbox-num" style="color:{C['red']}">{fmt_n(n_sku_absents)}</div>
    <div class="vbox-lbl">Absents stock</div>
  </div>
  <div class="vbox-item">
    <div class="vbox-num" style="color:{C['muted']}">{n_mag}</div>
    <div class="vbox-lbl">Magasins</div>
  </div>
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

# Grille cartésienne SKU × magasin
grid = pd.DataFrame(
    pd.MultiIndex.from_product([mag_sel_codes, sku_scope],
                               names=["Code site", "SKU"]).tolist(),
    columns=["Code site", "SKU"]
)
merged = grid.merge(stk_filt[[
    "Code site", "SKU", "Nouveau stock", "Ral", "Code etat",
    "Origine", "Libellé article", "Nom fourn.", "Libellé rayon",
    "Libellé famille", "Qté sortie", "Pcb", "Date dernière entrée"
]], on=["Code site", "SKU"], how="left")

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

# Libellé article : stock > T1
if "Libellé article" not in merged.columns:
    merged["Libellé article"] = ""
merged["Libellé article"] = merged["Libellé article"].fillna("").astype(str)
merged["Libellé article"] = merged.apply(
    lambda r: r["Libellé article"] if r["Libellé article"] else r.get("T1_lib", ""), axis=1
)
# Flux : stock > T1
if "Origine" not in merged.columns:
    merged["Origine"] = merged["Origine_T1"]
merged["Origine"] = merged.apply(
    lambda r: r["Origine"] if pd.notna(r["Origine"]) and r["Origine"] else r.get("Origine_T1", "LO"), axis=1
)
merged.drop(columns=["T1_lib", "Origine_T1"], errors="ignore", inplace=True)
merged["Magasin"] = merged["Code site"].map(site_ref).fillna(merged["Code site"])

# ─────────────────────────────────────────────────────────────────────────────
# STATUTS
# ─────────────────────────────────────────────────────────────────────────────
def get_statut(row):
    etat  = str(row.get("Code etat", "")).strip()
    stock = row.get("Nouveau stock")

    # Absent du stock = ligne vide après left join
    if etat == "" or pd.isna(etat) or etat == "nan":
        return "⚠️ Non implanté"
    # Anomalie référencement (hors scope mais alerté)
    if etat in ETAT_ANOMALIE:
        return "🚩 Anomalie référ."
    # Actif (Code etat = 2)
    if pd.isna(stock): return "⚠️ Non implanté"
    v = float(stock)
    if v > 0: return "✅ Implanté"
    if v < 0: return "🔴 Stock négatif"
    return "⚠️ Non implanté"

merged["Statut"] = merged.apply(get_statut, axis=1)

# Couverture en jours
merged["Couverture (j)"] = merged.apply(
    lambda r: couverture_jours(
        r.get("Nouveau stock", 0) or 0,
        r.get("Qté sortie", 0) or 0
    ), axis=1
)
merged["Badge couverture"] = merged["Couverture (j)"].apply(badge_couverture)

# ─────────────────────────────────────────────────────────────────────────────
# MÉTRIQUES
# ─────────────────────────────────────────────────────────────────────────────
# On calcule le taux uniquement sur Code etat = 2
merged_actif = merged[merged["Code etat"].astype(str).str.strip() == ETAT_ACTIF]
n_impl_actif = int((merged_actif["Statut"] == "✅ Implanté").sum())
n_base_taux  = len(merged_actif)  # dénominateur = actifs uniquement

total_cells = len(merged)
n_impl      = int((merged["Statut"] == "✅ Implanté").sum())
n_non_impl  = int((merged["Statut"] == "⚠️ Non implanté").sum())
n_neg       = int((merged["Statut"] == "🔴 Stock négatif").sum())
n_anomalie  = int((merged["Statut"] == "🚩 Anomalie référ.").sum())
taux_reseau = int(n_impl_actif / n_base_taux * 100) if n_base_taux else 0
n_sku_im    = int((t1_scope["ORIGINE"] == "IM").sum())
n_sku_lo    = int((t1_scope["ORIGINE"] == "LO").sum())
pct         = lambda n: int(n / total_cells * 100) if total_cells else 0

pivot_mag = (
    merged.groupby(["Magasin", "Statut"])
    .size().unstack(fill_value=0)
    .reindex(columns=list(STATUT_COLORS.keys()), fill_value=0)
    .reset_index()
)
pivot_mag.columns.name = None

# Taux par magasin (sur actifs uniquement)
def taux_magasin(mag):
    df_m = merged[(merged["Magasin"] == mag) &
                  (merged["Code etat"].astype(str).str.strip() == ETAT_ACTIF)]
    if len(df_m) == 0: return 0
    return int((df_m["Statut"] == "✅ Implanté").sum() / len(df_m) * 100)

pivot_mag["Taux (%)"] = pivot_mag["Magasin"].apply(taux_magasin)

# Taux par rayon × magasin
if "Libellé rayon" in merged.columns:
    rayon_pivot = (
        merged[merged["Code etat"].astype(str).str.strip() == ETAT_ACTIF]
        .groupby(["Libellé rayon", "Magasin"])
        .apply(lambda x: int((x["Statut"] == "✅ Implanté").sum() / len(x) * 100)
               if len(x) > 0 else 0)
        .reset_index(name="Taux (%)")
        .pivot(index="Libellé rayon", columns="Magasin", values="Taux (%)")
        .fillna(0).astype(int)
    )
else:
    rayon_pivot = pd.DataFrame()

# ─────────────────────────────────────────────────────────────────────────────
# BANNIÈRE ALERTES
# ─────────────────────────────────────────────────────────────────────────────
if n_non_impl + n_neg + n_anomalie > 0:
    st.markdown(f"""
    <div class="alert-banner">
      <div class="alert-pill">⚡ ACTIONS REQUISES</div>
      <div class="alert-item"><div class="alert-num" style="color:{C['orange']}">{fmt_n(n_non_impl)}</div><div class="alert-lbl">Non implanté</div></div>
      <div class="alert-item"><div class="alert-num" style="color:{C['red']}">{fmt_n(n_neg)}</div><div class="alert-lbl">Stock négatif</div></div>
      <div class="alert-item"><div class="alert-num" style="color:{C['yellow']}">{fmt_n(n_anomalie)}</div><div class="alert-lbl">Anomalie référ.</div></div>
      <div style="margin-left:auto;font-size:12px;color:{C['muted']};">{n_mag} mag · {fmt_n(n_sku)} SKU · {fmt_n(total_cells)} combinaisons</div>
    </div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# KPI IMPLANTATION
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="partie-label">📦 PARTIE 1 — IMPLANTATION</div>', unsafe_allow_html=True)

st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card green">
    <div class="kpi-label">✅ Implanté</div>
    <div class="kpi-value">{fmt_n(n_impl)}</div>
    <div class="kpi-sub">{pct(n_impl)}% du réseau</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(n_impl)}%;background:{C['green']}"></div></div>
  </div>
  <div class="kpi-card orange">
    <div class="kpi-label">⚠️ Non implanté</div>
    <div class="kpi-value">{fmt_n(n_non_impl)}</div>
    <div class="kpi-sub">{pct(n_non_impl)}% — à traiter</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(n_non_impl)}%;background:{C['orange']}"></div></div>
  </div>
  <div class="kpi-card red">
    <div class="kpi-label">🔴 Stock négatif</div>
    <div class="kpi-value">{fmt_n(n_neg)}</div>
    <div class="kpi-sub">{pct(n_neg)}% — écart inventaire</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{pct(n_neg)}%;background:{C['red']}"></div></div>
  </div>
  <div class="kpi-card blue">
    <div class="kpi-label">📊 Taux réseau</div>
    <div class="kpi-value" style="color:{color_taux(taux_reseau)}">{taux_reseau}%</div>
    <div class="kpi-sub">sur {fmt_n(n_base_taux)} refs actives</div>
    <div class="kpi-bar"><div class="kpi-bar-fill" style="width:{taux_reseau}%;background:{C['blue']}"></div></div>
  </div>
  <div class="kpi-card yellow">
    <div class="kpi-label">🚩 Anomalies référ.</div>
    <div class="kpi-value">{fmt_n(n_anomalie)}</div>
    <div class="kpi-sub">Code etat ≠ 2 · hors taux</div>
  </div>
</div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SCORECARD MAGASINS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="sh">SCORECARD MAGASINS</div>', unsafe_allow_html=True)
rag_html = '<div class="scorecard-grid">'
for _, row in pivot_mag.sort_values("Taux (%)", ascending=False).iterrows():
    t_  = row["Taux (%)"]
    col = color_taux(t_)
    ok_ = int(row.get("✅ Implanté", 0))
    ko_ = int(row.get("⚠️ Non implanté", 0))
    neg = int(row.get("🔴 Stock négatif", 0))
    an_ = int(row.get("🚩 Anomalie référ.", 0))
    rag_html += f"""<div class="scorecard-card {scorecard_cls(t_)}">
      <div class="scorecard-dot" style="background:{col}"></div>
      <div class="scorecard-name">{row['Magasin']}</div>
      <div class="scorecard-pct" style="color:{col}">{t_}%</div>
      <div class="scorecard-sub">{ok_}✅ {ko_}⚠️ {neg}🔴 {an_}🚩</div>
    </div>"""
rag_html += "</div>"
st.markdown(rag_html, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# ONGLETS IMPLANTATION
# ─────────────────────────────────────────────────────────────────────────────
TABS_IMPL = [
    "📋 COPIL", "📊 Vue Réseau", "🚩 Référencement",
    "⚠️ Non Implantés", "🔴 Stocks Négatifs",
    "🗓️ Calendrier", "🔄 Cessions", "📥 Export"
]
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

# ══════════════════════════════════════════════════════════════════════════════
# TAB COPIL
# ══════════════════════════════════════════════════════════════════════════════
if active_impl == TABS_IMPL[0]:
    st.markdown('<div class="sh-copil">SYNTHÈSE COPIL — VUE DIRECTION</div>', unsafe_allow_html=True)

    # Top alertes
    df_alertes_top = merged[merged["Statut"].isin(["⚠️ Non implanté","🔴 Stock négatif"])] \
                     .groupby("SKU").agg(
                         Libellé=("Libellé article","first"),
                         Origine=("Origine","first"),
                         Nb_magasins=("Magasin","count"),
                     ).reset_index()
    df_alertes_top["Score"] = df_alertes_top.apply(
        lambda r: r["Nb_magasins"] * (2 if r["Origine"] == "IM" else 1), axis=1
    )
    df_alertes_top = df_alertes_top.sort_values("Score", ascending=False).head(5)

    # Top à commander
    df_cmd_top = merged[merged["Statut"] == "⚠️ Non implanté"] \
                 .groupby("SKU").agg(
                     Libellé=("Libellé article","first"),
                     Fournisseur=("Fournisseur T1","first"),
                     Nb_magasins=("Magasin","count"),
                 ).reset_index().sort_values("Nb_magasins", ascending=False).head(5)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="copil-card"><div class="copil-title">🔴 Top 5 alertes critiques</div>', unsafe_allow_html=True)
        if df_alertes_top.empty:
            st.success("✅ Aucune alerte")
        else:
            st.dataframe(
                df_alertes_top[["SKU","Libellé","Origine","Nb_magasins","Score"]]
                .rename(columns={"Nb_magasins":"Mag. touchés"}),
                use_container_width=True, hide_index=True
            )
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="copil-card"><div class="copil-title">🛒 Top 5 à commander</div>', unsafe_allow_html=True)
        if df_cmd_top.empty:
            st.success("✅ Tous les articles ont une commande")
        else:
            st.dataframe(
                df_cmd_top[["SKU","Libellé","Fournisseur","Nb_magasins"]]
                .rename(columns={"Nb_magasins":"Mag. sans stock"}),
                use_container_width=True, hide_index=True
            )
        st.markdown('</div>', unsafe_allow_html=True)

    # Matrice Rayon × Magasin
    if not rayon_pivot.empty:
        st.markdown('<div class="sh-copil">TAUX D\'IMPLANTATION PAR RAYON × MAGASIN</div>', unsafe_allow_html=True)

        def color_cell(val):
            if val >= 80: bg, color = "#D1FAE5", "#065F46"
            elif val >= 50: bg, color = "#FEF3C7", "#92400E"
            else: bg, color = "#FEE2E2", "#991B1B"
            return f"background-color:{bg};color:{color};font-weight:700"

        st.dataframe(
            rayon_pivot.style.applymap(color_cell).format("{}%"),
            use_container_width=True
        )

    # Synthèse magasins
    st.markdown('<div class="sh-copil">SYNTHÈSE PAR MAGASIN</div>', unsafe_allow_html=True)
    cols_s = ["Magasin"] + [c for c in STATUT_COLORS if c in pivot_mag.columns] + ["Taux (%)"]
    st.dataframe(
        pivot_mag[cols_s].sort_values("Taux (%)", ascending=False).reset_index(drop=True)
        .style.format({"Taux (%)": "{}%"}),
        use_container_width=True, hide_index=True
    )

# ══════════════════════════════════════════════════════════════════════════════
# TAB VUE RÉSEAU
# ══════════════════════════════════════════════════════════════════════════════
elif active_impl == TABS_IMPL[1]:
    c1, c2 = st.columns([3, 2])
    with c1:
        mel = pivot_mag.melt(
            id_vars="Magasin",
            value_vars=[s for s in STATUT_COLORS if s in pivot_mag.columns],
            var_name="Statut", value_name="N"
        )
        fig = px.bar(mel, x="Magasin", y="N", color="Statut",
                     color_discrete_map=STATUT_COLORS, barmode="stack",
                     title="Situation par magasin")
        fig.update_traces(textposition="inside", texttemplate="%{y}",
                          textfont=dict(size=11, color="white"))
        fig.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=400,
                          font=dict(family="Inter", color=C["muted"], size=12),
                          margin=dict(l=10,r=10,t=44,b=20),
                          legend=dict(orientation="h", y=-0.28, bgcolor="rgba(0,0,0,0)"),
                          xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig_d = go.Figure(go.Pie(
            labels=["✅ Implanté","⚠️ Non implanté","🔴 Stock négatif","🚩 Anomalie"],
            values=[n_impl, n_non_impl, n_neg, n_anomalie], hole=0.62,
            marker=dict(colors=[C["green"],C["orange"],C["red"],C["yellow"]],
                        line=dict(color="#fff", width=3))
        ))
        fig_d.add_annotation(text=f"<b>{taux_reseau}%</b><br>implanté",
                             x=0.5, y=0.5, showarrow=False,
                             font=dict(size=20, color=C["text"], family="Inter"))
        fig_d.update_layout(paper_bgcolor=C["surface"], height=400,
                            font=dict(family="Inter", color=C["muted"], size=12),
                            margin=dict(l=10,r=10,t=44,b=20),
                            legend=dict(orientation="v", x=1.01),
                            title="Répartition réseau")
        st.plotly_chart(fig_d, use_container_width=True)

    # Matrice rayon × magasin
    if not rayon_pivot.empty:
        st.markdown('<div class="sh">TAUX PAR RAYON × MAGASIN</div>', unsafe_allow_html=True)
        def color_cell(val):
            if val >= 80: bg, color = "#D1FAE5", "#065F46"
            elif val >= 50: bg, color = "#FEF3C7", "#92400E"
            else: bg, color = "#FEE2E2", "#991B1B"
            return f"background-color:{bg};color:{color};font-weight:700"
        st.dataframe(rayon_pivot.style.applymap(color_cell).format("{}%"),
                     use_container_width=True)

    # Flux IM/LO
    df_flux = merged[merged["Statut"]=="✅ Implanté"].groupby(["Magasin","Origine"]).size().reset_index(name="N")
    if not df_flux.empty:
        st.markdown('<div class="sh">FLUX IM / LO — ARTICLES IMPLANTÉS</div>', unsafe_allow_html=True)
        fig_flux = px.bar(df_flux, x="Magasin", y="N", color="Origine",
                          color_discrete_map={"IM":C["blue"],"LO":C["green"]},
                          barmode="group", title="Articles implantés par flux")
        fig_flux.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=300,
                               font=dict(family="Inter", color=C["muted"], size=12),
                               margin=dict(l=10,r=10,t=44,b=20),
                               legend=dict(orientation="h", y=-0.3),
                               xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]))
        st.plotly_chart(fig_flux, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB RÉFÉRENCEMENT
# ══════════════════════════════════════════════════════════════════════════════
elif active_impl == TABS_IMPL[2]:
    df_an = merged[merged["Statut"] == "🚩 Anomalie référ."].copy()
    st.markdown('<div class="sh">🚩 ANOMALIES RÉFÉRENCEMENT — CODE ETAT ≠ 2</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="info-box yellow">
      <strong>📌 Règle métier</strong> — Tout article présent dans la liste T1 doit avoir
      <strong>Code etat = 2</strong> (actif) sur chaque magasin. Un code différent signale un problème
      de référencement à corriger avant toute implantation. Ces articles sont exclus du taux d'implantation.
    </div>""", unsafe_allow_html=True)

    if df_an.empty:
        st.markdown('<div class="info-box green">✅ Aucune anomalie de référencement détectée.</div>', unsafe_allow_html=True)
    else:
        k1, k2, k3 = st.columns(3)
        k1.metric("🚩 Lignes anomalie", fmt_n(len(df_an)))
        k2.metric("SKUs distincts",     fmt_n(df_an["SKU"].nunique()))
        k3.metric("Magasins touchés",   df_an["Magasin"].nunique())

        # Criticité = nb magasins × flux
        df_an_agg = df_an.groupby("SKU").agg(
            Libellé=("Libellé article","first"),
            Fournisseur=("Fournisseur T1","first"),
            Origine=("Origine","first"),
            Nb_magasins=("Magasin","count"),
            Codes_etat=("Code etat", lambda x: ", ".join(sorted(set(x.dropna().astype(str))))),
        ).reset_index()
        df_an_agg["Criticité"] = df_an_agg.apply(
            lambda r: r["Nb_magasins"] * (2 if r["Origine"] == "IM" else 1), axis=1
        )
        df_an_agg = df_an_agg.sort_values("Criticité", ascending=False).reset_index(drop=True)
        df_an_agg["Code etat"] = df_an_agg["Codes_etat"].map(
            lambda x: " · ".join([f"{e}={ETAT_LABEL.get(e,e)}" for e in x.split(", ")])
        )

        st.dataframe(df_an_agg[["SKU","Libellé","Origine","Fournisseur","Nb_magasins","Code etat","Criticité"]]
                     .rename(columns={"Nb_magasins":"Mag. touchés"}),
                     use_container_width=True, hide_index=True)

        # Par fournisseur
        four_an = df_an.groupby("Fournisseur T1")["SKU"].nunique().sort_values(ascending=False).reset_index()
        four_an.columns = ["Fournisseur","SKUs en anomalie"]
        st.markdown('<div class="sh">PAR FOURNISSEUR</div>', unsafe_allow_html=True)
        st.dataframe(four_an, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB NON IMPLANTÉS
# ══════════════════════════════════════════════════════════════════════════════
elif active_impl == TABS_IMPL[3]:
    df_ni = merged[merged["Statut"] == "⚠️ Non implanté"].copy()

    # Ruptures totales
    sku_sans_stock  = (merged[merged["Statut"] != "✅ Implanté"]
                       .groupby("SKU")["Magasin"].count())
    sku_rupture_tot = sku_sans_stock[sku_sans_stock == n_mag].index.tolist()
    n_rc_im = int(t1_scope[t1_scope["SKU"].isin(sku_rupture_tot) &
                            (t1_scope["ORIGINE"] == "IM")]["SKU"].nunique())
    n_rc_lo = int(t1_scope[t1_scope["SKU"].isin(sku_rupture_tot) &
                            (t1_scope["ORIGINE"] == "LO")]["SKU"].nunique())

    if sku_rupture_tot:
        st.markdown(f"""
        <div class="alert-banner">
          <div class="alert-pill">🚨 RUPTURE TOTALE RÉSEAU</div>
          <div class="alert-item"><div class="alert-num" style="color:{C['red']}">{len(sku_rupture_tot)}</div><div class="alert-lbl">SKUs absents partout</div></div>
          <div class="alert-item"><div class="alert-num" style="color:{C['blue']}">{n_rc_im}</div><div class="alert-lbl">Flux IM</div></div>
          <div class="alert-item"><div class="alert-num" style="color:{C['green']}">{n_rc_lo}</div><div class="alert-lbl">Flux LO</div></div>
          <div style="margin-left:auto;font-size:12px;color:{C['muted']};">Aucun stock positif sur les {n_mag} magasin(s)</div>
        </div>""", unsafe_allow_html=True)

    sub1, sub2 = st.tabs([f"⚠️ Non implantés ({len(df_ni)})",
                          f"🔴 Ruptures totales ({len(sku_rupture_tot)} SKU)"])
    with sub1:
        if df_ni.empty:
            st.markdown('<div class="info-box green">✅ Tous les articles sont implantés !</div>', unsafe_allow_html=True)
        else:
            k1, k2, k3 = st.columns(3)
            k1.metric("⚠️ Lignes", fmt_n(len(df_ni)))
            k2.metric("SKUs distincts", fmt_n(df_ni["SKU"].nunique()))
            k3.metric("dont Rupture totale", len(sku_rupture_tot))
            df_ni["Rupture totale"] = df_ni["SKU"].isin(sku_rupture_tot).map({True:"🔴 OUI",False:"—"})
            COLS = ["Magasin","SKU","Libellé article","Origine","Mode Appro",
                    "Sem. Réception","Fournisseur T1","Rupture totale"]
            st.dataframe(df_ni[[c for c in COLS if c in df_ni.columns]]
                         .sort_values(["Rupture totale","Magasin"]).reset_index(drop=True),
                         use_container_width=True, hide_index=True)

    with sub2:
        if not sku_rupture_tot:
            st.markdown('<div class="info-box green">✅ Aucune rupture totale.</div>', unsafe_allow_html=True)
        else:
            k1, k2 = st.columns(2)
            k1.metric("SKUs rupture totale", len(sku_rupture_tot))
            k2.metric("IM / LO", f"{n_rc_im} / {n_rc_lo}")
            df_rc = (merged[merged["SKU"].isin(sku_rupture_tot)]
                     [["SKU","Libellé article","Origine","Fournisseur T1","Mode Appro","Sem. Réception"]]
                     .drop_duplicates("SKU").sort_values("Origine").reset_index(drop=True))
            st.dataframe(df_rc, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB STOCKS NÉGATIFS
# ══════════════════════════════════════════════════════════════════════════════
elif active_impl == TABS_IMPL[4]:
    df_neg = merged[merged["Statut"] == "🔴 Stock négatif"].copy()
    st.markdown('<div class="sh">🔴 STOCKS NÉGATIFS — ÉCARTS INVENTAIRE</div>', unsafe_allow_html=True)
    st.markdown('<div class="info-box orange">📌 <strong>Stock négatif = écart d\'inventaire</strong> — articles sortis sans réception ERP correspondante. Action : régularisation inventaire magasin.</div>', unsafe_allow_html=True)
    if df_neg.empty:
        st.markdown('<div class="info-box green">✅ Aucun stock négatif détecté.</div>', unsafe_allow_html=True)
    else:
        k1, k2 = st.columns(2)
        k1.metric("🔴 Lignes", fmt_n(len(df_neg)))
        k2.metric("SKUs distincts", fmt_n(df_neg["SKU"].nunique()))
        COLS = ["Magasin","SKU","Libellé article","Origine","Nouveau stock","Mode Appro","Fournisseur T1"]
        st.dataframe(df_neg[[c for c in COLS if c in df_neg.columns]]
                     .sort_values("Nouveau stock").reset_index(drop=True),
                     use_container_width=True, hide_index=True)
        fig_neg = px.histogram(df_neg, x="Nouveau stock", nbins=30,
                               title="Distribution des stocks négatifs",
                               color_discrete_sequence=[C["red"]])
        fig_neg.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=240,
                              font=dict(family="Inter", color=C["muted"], size=12),
                              margin=dict(l=10,r=10,t=44,b=20))
        st.plotly_chart(fig_neg, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB CALENDRIER
# ══════════════════════════════════════════════════════════════════════════════
elif active_impl == TABS_IMPL[5]:
    cal_df = merged[merged["Sem. Réception"].str.match(r"^[Ss]?\d+$", na=False)].copy()
    if cal_df.empty:
        st.info("Aucune semaine de réception renseignée dans le T1.")
    else:
        cal_agg = cal_df.groupby("Sem. Réception").agg(
            Implanté     =("Statut", lambda x: (x=="✅ Implanté").sum()),
            Non_implanté =("Statut", lambda x: (x=="⚠️ Non implanté").sum()),
            Stock_négatif=("Statut", lambda x: (x=="🔴 Stock négatif").sum()),
            SKU_distincts=("SKU","nunique"),
        ).reset_index().rename(columns={"Non_implanté":"Non implanté","Stock_négatif":"Stock négatif"})
        cal_agg["Taux (%)"] = (cal_agg["Implanté"] /
                               (cal_agg["Implanté"]+cal_agg["Non implanté"]+cal_agg["Stock négatif"]) * 100
                               ).round(0).astype(int)
        fig_cal = px.bar(
            cal_agg.melt(id_vars="Sem. Réception",
                         value_vars=["Implanté","Non implanté","Stock négatif"]),
            x="Sem. Réception", y="value", color="variable",
            color_discrete_map={"Implanté":C["green"],"Non implanté":C["orange"],"Stock négatif":C["red"]},
            barmode="stack", title="Statut par semaine de réception"
        )
        fig_cal.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=360,
                              font=dict(family="Inter", color=C["muted"], size=12),
                              margin=dict(l=10,r=10,t=44,b=20),
                              legend=dict(orientation="h", y=-0.25),
                              xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]))
        st.plotly_chart(fig_cal, use_container_width=True)
        st.dataframe(cal_agg, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB CESSIONS
# ══════════════════════════════════════════════════════════════════════════════
elif active_impl == TABS_IMPL[6]:
    st.markdown('<div class="sh">🔄 MOTEUR CESSIONS INTER-MAGASINS</div>', unsafe_allow_html=True)

    # Cession automatique si rupture totale
    if sku_rupture_tot:
        st.markdown(f'<div class="info-box orange">🔄 <strong>{len(sku_rupture_tot)} article(s) en rupture totale</strong> — suggestion automatique de cession activée.</div>', unsafe_allow_html=True)

    if not mag_detresse and not sku_rupture_tot:
        st.markdown('<div class="info-box blue">⬅️ Sélectionne des magasins en détresse dans la sidebar, ou charge un T1 avec des ruptures totales pour activer les suggestions automatiques.</div>', unsafe_allow_html=True)
    else:
        mag_det_codes = [c for c in all_codes if site_ref.get(c,c) in mag_detresse]

        # Suggestions automatiques ruptures totales
        if sku_rupture_tot and not mag_detresse:
            # Le magasin avec le plus de stock par SKU devient cédant
            auto_sugg = []
            for sku in sku_rupture_tot:
                sku_df = df_stock[df_stock["SKU"]==sku].copy()
                sku_df["Nouveau stock"] = pd.to_numeric(sku_df["Nouveau stock"],errors="coerce").fillna(0)
                best = sku_df[sku_df["Nouveau stock"] > reserve].sort_values("Nouveau stock",ascending=False)
                lib  = sku_df["Libellé article"].iloc[0] if "Libellé article" in sku_df.columns and len(sku_df) else sku
                if not best.empty:
                    b   = best.iloc[0]
                    qty = int(b["Nouveau stock"]) - reserve
                    auto_sugg.append({"SKU":sku,"Libellé":lib,
                        "Cédant":site_ref.get(b["Code site"],b["Code site"]),
                        "Stock cédant":int(b["Nouveau stock"]),
                        "Qté cessible":qty,"Type":"🤖 Auto (rupture totale)"})
            if auto_sugg:
                st.markdown('<div class="sh">SUGGESTIONS AUTOMATIQUES — RUPTURES TOTALES</div>', unsafe_allow_html=True)
                st.dataframe(pd.DataFrame(auto_sugg), use_container_width=True, hide_index=True)

        # Cessions manuelles
        if mag_detresse:
            mag_ced_codes = [c for c in mag_sel_codes if c not in mag_det_codes]
            suggestions   = []
            for sku in sku_scope:
                sku_df = df_stock[df_stock["SKU"]==sku].copy()
                if sku_df.empty: continue
                lib = sku_df["Libellé article"].iloc[0] if "Libellé article" in sku_df.columns else sku
                sku_df["Nouveau stock"] = pd.to_numeric(sku_df["Nouveau stock"],errors="coerce").fillna(0)
                det_rows = sku_df[sku_df["Code site"].isin(mag_det_codes) &
                                  (sku_df["Nouveau stock"] <= seuil_det)]
                if det_rows.empty: continue
                ced_rows = sku_df[sku_df["Code site"].isin(mag_ced_codes) &
                                  (sku_df["Nouveau stock"] > reserve)].sort_values("Nouveau stock",ascending=False)
                for _, dr in det_rows.iterrows():
                    if ced_rows.empty:
                        suggestions.append({"SKU":sku,"Libellé article":lib,
                            "Magasin détresse":site_ref.get(dr["Code site"],dr["Code site"]),
                            "Stock détresse":int(dr["Nouveau stock"]),
                            "Cédant suggéré":"⚠️ Aucun","Stock cédant":0,
                            "Qté cessible":0,"Faisabilité":"🔴 Impossible"})
                    else:
                        best = ced_rows.iloc[0]
                        qty  = int(best["Nouveau stock"]) - reserve
                        suggestions.append({"SKU":sku,"Libellé article":lib,
                            "Magasin détresse":site_ref.get(dr["Code site"],dr["Code site"]),
                            "Stock détresse":int(dr["Nouveau stock"]),
                            "Cédant suggéré":site_ref.get(best["Code site"],best["Code site"]),
                            "Stock cédant":int(best["Nouveau stock"]),"Qté cessible":qty,
                            "Faisabilité":"🟢 Possible" if qty>=1 else "🟠 Partielle"})
            if suggestions:
                df_cess = pd.DataFrame(suggestions).sort_values(
                    ["Faisabilité","Qté cessible"],ascending=[True,False]).reset_index(drop=True)
                k1,k2,k3 = st.columns(3)
                k1.metric("🟢 Possible",   int((df_cess["Faisabilité"]=="🟢 Possible").sum()))
                k2.metric("🔴 Impossible", int((df_cess["Faisabilité"]=="🔴 Impossible").sum()))
                k3.metric("Articles",      df_cess["SKU"].nunique())
                st.dataframe(df_cess, use_container_width=True, hide_index=True)
                buf_c = io.BytesIO()
                with pd.ExcelWriter(buf_c, engine="openpyxl") as w:
                    df_cess.to_excel(w, sheet_name="Plan Cessions", index=False)
                buf_c.seek(0)
                st.download_button(f"📥 Plan_Cessions_{TODAY_FILE}.xlsx", data=buf_c,
                                   file_name=f"Plan_Cessions_{TODAY_FILE}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.success("✅ Aucune cession nécessaire selon les critères.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB EXPORT
# ══════════════════════════════════════════════════════════════════════════════
elif active_impl == TABS_IMPL[7]:
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    st.markdown('<div class="info-box blue">3 feuilles : <strong>Synthèse réseau</strong> · <strong>Détail complet</strong> · <strong>Alertes prioritaires</strong></div>', unsafe_allow_html=True)
    if st.button("🔨 Générer Export Implantation", type="primary"):
        buf_x = io.BytesIO()
        COLS_DET = ["Magasin","SKU","Libellé article","Origine","Code etat","Mode Appro",
                    "Sem. Réception","Fournisseur T1","Nouveau stock","Statut","Badge couverture"]
        STATUT_FILLS = {
            "✅ Implanté"       : ("D1FAE5","065F46"),
            "⚠️ Non implanté"  : ("FEF3C7","92400E"),
            "🔴 Stock négatif"  : ("FEE2E2","991B1B"),
            "🚩 Anomalie référ.": ("FFFDE7","795548"),
        }
        with pd.ExcelWriter(buf_x, engine="openpyxl") as writer:
            cols_s = ["Magasin"]+[c for c in STATUT_COLORS if c in pivot_mag.columns]+["Taux (%)"]
            pivot_mag[cols_s].sort_values("Taux (%)",ascending=False).to_excel(
                writer, sheet_name="Synthèse Réseau", index=False)
            merged[[c for c in COLS_DET if c in merged.columns]].to_excel(
                writer, sheet_name="Détail Complet", index=False)
            df_al = merged[merged["Statut"].isin(["⚠️ Non implanté","🔴 Stock négatif","🚩 Anomalie référ."])]
            df_al[[c for c in COLS_DET if c in df_al.columns]].sort_values(
                ["Statut","Magasin"]).to_excel(writer, sheet_name="Alertes Prioritaires", index=False)
            wb = writer.book
            FH = PatternFill("solid",fgColor="1C1C1E")
            FT = Font(bold=True,color="FFFFFF",name="Arial",size=11)
            for sn in wb.sheetnames:
                ws = wb[sn]
                for cell in ws[1]:
                    cell.fill=FH; cell.font=FT
                    cell.alignment=Alignment(horizontal="center")
                for col in ws.columns:
                    ws.column_dimensions[get_column_letter(col[0].column)].width = min(
                        max((len(str(c.value)) for c in col if c.value),default=10)+4,50)
                ws.freeze_panes="A2"
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if str(cell.value) in STATUT_FILLS:
                            bg, fg = STATUT_FILLS[str(cell.value)]
                            cell.fill = PatternFill("solid",fgColor=bg)
                            cell.font = Font(color=fg,name="Arial",size=10)
        buf_x.seek(0)
        st.download_button(f"📥 Implantation_T1_{TODAY_FILE}.xlsx", data=buf_x,
                           file_name=f"Implantation_T1_{TODAY_FILE}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success(f"✅ Export — {fmt_n(len(merged))} lignes · 3 feuilles")


# ═════════════════════════════════════════════════════════════════════════════
# PARTIE 2 — SUPPLY
# ═════════════════════════════════════════════════════════════════════════════
if not supply_active:
    st.markdown(f"""
    <div class="supply-divider">
      <div class="supply-divider-icon">🚚</div>
      <div>
        <div class="supply-divider-title">Partie Supply — non activée</div>
        <div class="supply-divider-sub">
          Charge les fichiers RAL (③) dans la sidebar pour activer le suivi des livraisons.
          Le RAL estimé (colonne Ral du stock) est disponible dans le tableau Détail.
        </div>
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
    <div class="supply-divider-sub">Stock ERP × RAL actif (38·40·50) · {len(ral_files)} fichier(s) RAL</div>
  </div>
</div>""", unsafe_allow_html=True)

st.markdown('<div class="partie-label supply">🚚 PARTIE 2 — SUPPLY</div>', unsafe_allow_html=True)

supply_df = merged.merge(df_ral_agg, on=["SKU","Code site"], how="left")
supply_df["RAL_actif"]           = supply_df["RAL_actif"].fillna(0)
supply_df["Nb_commandes"]        = supply_df["Nb_commandes"].fillna(0).astype(int)
supply_df["Commandes_en_retard"] = supply_df["Commandes_en_retard"].fillna(0).astype(int)
supply_df["Jours_retard_max"]    = supply_df["Jours_retard_max"].fillna(0).astype(int)
supply_df["Action Supply"]       = supply_df.apply(
    lambda r: action_supply(r["Nouveau stock"], r["RAL_actif"],
                            r.get("Situation_principale")), axis=1)
supply_df["Retard Badge"]        = supply_df["Jours_retard_max"].apply(retard_badge)

# ─────────────────────────────────────────────────────────────────────────────
# KPI SUPPLY
# ─────────────────────────────────────────────────────────────────────────────
n_attente   = int(supply_df["Action Supply"].str.contains("En attente|En transit|Réception",na=False).sum())
n_commander = int((supply_df["Action Supply"]=="🛒 Passer commande").sum())
n_reg       = int(supply_df["Action Supply"].str.contains("Régulariser",na=False).sum())
n_retard    = int((supply_df["Commandes_en_retard"]>0).sum())
retard_max  = int(supply_df["Jours_retard_max"].max()) if len(supply_df) else 0

st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card blue"><div class="kpi-label">🕐 En attente livraison</div><div class="kpi-value">{fmt_n(n_attente)}</div><div class="kpi-sub">RAL actif confirmé</div></div>
  <div class="kpi-card orange"><div class="kpi-label">🛒 À commander</div><div class="kpi-value">{fmt_n(n_commander)}</div><div class="kpi-sub">Aucun RAL · acheteur</div></div>
  <div class="kpi-card red"><div class="kpi-label">🔧 À régulariser</div><div class="kpi-value">{fmt_n(n_reg)}</div><div class="kpi-sub">Écart inventaire</div></div>
  <div class="kpi-card purple"><div class="kpi-label">⚠️ Lignes en retard</div><div class="kpi-value">{fmt_n(n_retard)}</div><div class="kpi-sub">Date ETA dépassée</div></div>
  <div class="kpi-card teal"><div class="kpi-label">🚨 Retard max</div><div class="kpi-value">{retard_max}j</div><div class="kpi-sub">{"🔴 Critique" if retard_max>30 else ("🟠 Modéré" if retard_max>7 else "🟡 Léger")}</div></div>
</div>""", unsafe_allow_html=True)

# SCORECARD SUPPLY
st.markdown('<div class="sh-supply">SCORECARD SUPPLY PAR MAGASIN</div>', unsafe_allow_html=True)
sup_pivot = (supply_df.groupby("Magasin")["Action Supply"]
             .value_counts().unstack(fill_value=0).reset_index())
sup_pivot.columns.name = None
sc_html = '<div class="scorecard-grid">'
for _, row in sup_pivot.iterrows():
    att = sum(int(row.get(k,0)) for k in ["🕐 En attente livraison","🚚 En transit","📥 Réception en cours"])
    cmd = int(row.get("🛒 Passer commande",0))
    reg = sum(int(row.get(k,0)) for k in ["🔧 Régulariser inventaire","🔧 Régulariser + livraison en cours"])
    ok  = sum(int(row.get(k,0)) for k in ["✅ Implanté","✅ Implanté + réassort en cours"])
    cls = "ok" if cmd==0 and reg==0 else ("warn" if cmd>0 else "ko")
    sc_html += f'<div class="scorecard-card {cls}"><div class="scorecard-name">{row["Magasin"]}</div><div class="scorecard-sub">🕐{att} 🛒{cmd} 🔧{reg} ✅{ok}</div></div>'
sc_html += "</div>"
st.markdown(sc_html, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# ONGLETS SUPPLY
# ─────────────────────────────────────────────────────────────────────────────
TABS_SUP = ["🎯 Plan d'action","🕐 Suivi livraisons","🛒 À commander","📥 Export Supply"]
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
COLS_SUP   = ["Magasin","SKU","Libellé article","Origine","Nouveau stock",
              "RAL_actif","Nb_commandes","Prochaine_livraison_fmt",
              "Retard Badge","Badge couverture","Action Supply"]

# ── PLAN D'ACTION ─────────────────────────────────────────────────────────────
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
                 .rename(columns={"RAL_actif":"RAL","Nb_commandes":"Nb cdes",
                                  "Prochaine_livraison_fmt":"Proch. livraison",
                                  "Retard Badge":"Retard","Badge couverture":"Couverture"})
                 .sort_values(["Action Supply","Magasin"]).reset_index(drop=True),
                 use_container_width=True, hide_index=True)

    act_count = supply_df["Action Supply"].value_counts().reset_index()
    act_count.columns = ["Action","N"]
    fig_act = px.bar(act_count, x="N", y="Action", orientation="h",
                     title="Répartition des actions Supply", color="Action",
                     color_discrete_map={
                         "🛒 Passer commande":C["orange"],
                         "🕐 En attente livraison":C["blue"],
                         "🚚 En transit":C["teal"],
                         "📥 Réception en cours":C["green"],
                         "🔧 Régulariser inventaire":C["red"],
                         "🔧 Régulariser + livraison en cours":"#FF6B6B",
                         "✅ Implanté + réassort en cours":"#34C759",
                         "✅ Implanté":"#30D158",
                     })
    fig_act.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=300,
                          font=dict(family="Inter", color=C["muted"], size=12),
                          margin=dict(l=10,r=10,t=44,b=20), showlegend=False,
                          xaxis=dict(gridcolor=C["bg"]))
    st.plotly_chart(fig_act, use_container_width=True)

# ── SUIVI LIVRAISONS ──────────────────────────────────────────────────────────
elif active_sup == TABS_SUP[1]:
    df_liv = supply_df[supply_df["RAL_actif"]>0].copy()
    st.markdown('<div class="sh-supply">COMMANDES EN COURS — SITUATIONS 38 · 40 · 50</div>', unsafe_allow_html=True)
    if df_liv.empty:
        st.markdown('<div class="info-box green">✅ Aucune commande en cours.</div>', unsafe_allow_html=True)
    else:
        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Lignes avec RAL",    fmt_n(len(df_liv)))
        k2.metric("SKUs distincts",     fmt_n(df_liv["SKU"].nunique()))
        k3.metric("RAL total (unités)", fmt_n(int(df_liv["RAL_actif"].sum())))
        k4.metric("Lignes en retard",   fmt_n(int((df_liv["Commandes_en_retard"]>0).sum())))

        df_liv["Tranche retard"] = pd.cut(
            df_liv["Jours_retard_max"], bins=[-1,0,7,30,999],
            labels=["✅ À temps","🟡 1-7j","🟠 8-30j","🔴 >30j"]
        )
        ret_dist = df_liv["Tranche retard"].value_counts().reset_index()
        ret_dist.columns = ["Tranche","N"]
        fig_ret = px.bar(ret_dist, x="Tranche", y="N", color="Tranche",
                         color_discrete_map={"✅ À temps":C["green"],"🟡 1-7j":"#D97706",
                                             "🟠 8-30j":C["orange"],"🔴 >30j":C["red"]},
                         title="Commandes par tranche de retard")
        fig_ret.update_layout(paper_bgcolor=C["surface"], plot_bgcolor=C["surface"], height=260,
                              font=dict(family="Inter", color=C["muted"], size=12),
                              margin=dict(l=10,r=10,t=44,b=20), showlegend=False,
                              xaxis=dict(gridcolor=C["bg"]), yaxis=dict(gridcolor=C["bg"]))
        st.plotly_chart(fig_ret, use_container_width=True)

        COLS_LIV = ["Magasin","SKU","Libellé article","Origine","Nouveau stock",
                    "RAL_actif","Nb_commandes","Prochaine_livraison_fmt",
                    "Jours_retard_max","Retard Badge","Action Supply"]
        st.dataframe(df_liv[[c for c in COLS_LIV if c in df_liv.columns]]
                     .rename(columns={"RAL_actif":"RAL","Nb_commandes":"Nb cdes",
                                      "Prochaine_livraison_fmt":"Proch. livraison",
                                      "Jours_retard_max":"Retard (j)","Retard Badge":"Statut retard"})
                     .sort_values("Retard (j)",ascending=False).reset_index(drop=True),
                     use_container_width=True, hide_index=True)

# ── À COMMANDER ───────────────────────────────────────────────────────────────
elif active_sup == TABS_SUP[2]:
    df_cmd = supply_df[supply_df["Action Supply"]=="🛒 Passer commande"].copy()
    st.markdown('<div class="sh-supply">ARTICLES SANS STOCK ET SANS COMMANDE EN COURS</div>', unsafe_allow_html=True)
    st.markdown('<div class="info-box orange">🛒 Ces articles T1 n\'ont <strong>aucun stock positif</strong> et <strong>aucune commande active</strong>. Action acheteur : passer commande fournisseur.</div>', unsafe_allow_html=True)
    if df_cmd.empty:
        st.markdown('<div class="info-box green">✅ Tous les articles non implantés ont une commande en cours.</div>', unsafe_allow_html=True)
    else:
        k1,k2,k3 = st.columns(3)
        k1.metric("🛒 À commander", fmt_n(len(df_cmd)))
        k2.metric("SKUs distincts", fmt_n(df_cmd["SKU"].nunique()))
        k3.metric("Flux IM",        fmt_n(int((df_cmd["Origine"]=="IM").sum())))
        COLS_CMD = ["Magasin","SKU","Libellé article","Origine","Mode Appro",
                    "Fournisseur T1","Sem. Réception","Nouveau stock"]
        st.dataframe(df_cmd[[c for c in COLS_CMD if c in df_cmd.columns]]
                     .sort_values(["Origine","Magasin"]).reset_index(drop=True),
                     use_container_width=True, hide_index=True)
        if "Fournisseur T1" in df_cmd.columns:
            four_count = (df_cmd.groupby("Fournisseur T1")["SKU"].nunique()
                          .sort_values(ascending=False).reset_index())
            four_count.columns = ["Fournisseur","SKU à commander"]
            st.markdown('<div class="sh-supply">PAR FOURNISSEUR</div>', unsafe_allow_html=True)
            st.dataframe(four_count, use_container_width=True, hide_index=True)

# ── EXPORT SUPPLY ─────────────────────────────────────────────────────────────
elif active_sup == TABS_SUP[3]:
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    st.markdown('<div class="info-box purple">3 feuilles : <strong>Plan d\'action</strong> · <strong>Suivi livraisons</strong> · <strong>À commander</strong></div>', unsafe_allow_html=True)
    if st.button("🔨 Générer Export Supply", type="primary"):
        ACTION_FILLS = {
            "🛒 Passer commande"                 : ("FEF3C7","92400E"),
            "🕐 En attente livraison"            : ("EFF4FF","1D4ED8"),
            "🚚 En transit"                      : ("F0FFFE","0E7490"),
            "📥 Réception en cours"              : ("F0FFF4","065F46"),
            "🔧 Régulariser inventaire"          : ("FEE2E2","991B1B"),
            "🔧 Régulariser + livraison en cours": ("FEE2E2","991B1B"),
            "✅ Implanté + réassort en cours"    : ("D1FAE5","065F46"),
            "✅ Implanté"                        : ("D1FAE5","065F46"),
        }
        COLS_EXP = ["Magasin","SKU","Libellé article","Origine","Nouveau stock",
                    "RAL_actif","Nb_commandes","Prochaine_livraison_fmt",
                    "Jours_retard_max","Retard Badge","Badge couverture","Action Supply"]
        COLS_CMD = ["Magasin","SKU","Libellé article","Origine","Mode Appro","Fournisseur T1","Sem. Réception"]

        buf_s = io.BytesIO()
        with pd.ExcelWriter(buf_s, engine="openpyxl") as writer:
            supply_df[[c for c in COLS_EXP if c in supply_df.columns]].sort_values(
                ["Action Supply","Magasin"]).to_excel(writer,sheet_name="Plan Action",index=False)
            supply_df[supply_df["RAL_actif"]>0][[c for c in COLS_EXP if c in supply_df.columns]].sort_values(
                "Jours_retard_max",ascending=False).to_excel(writer,sheet_name="Suivi Livraisons",index=False)
            supply_df[supply_df["Action Supply"]=="🛒 Passer commande"][
                [c for c in COLS_CMD if c in supply_df.columns]].to_excel(
                writer,sheet_name="À Commander",index=False)
            wb = writer.book
            FH = PatternFill("solid",fgColor="1C1C1E")
            FT = Font(bold=True,color="FFFFFF",name="Arial",size=11)
            for sn in wb.sheetnames:
                ws = wb[sn]
                for cell in ws[1]:
                    cell.fill=FH; cell.font=FT
                    cell.alignment=Alignment(horizontal="center")
                for col in ws.columns:
                    ws.column_dimensions[get_column_letter(col[0].column)].width = min(
                        max((len(str(c.value)) for c in col if c.value),default=10)+4,50)
                ws.freeze_panes="A2"
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if str(cell.value) in ACTION_FILLS:
                            bg,fg = ACTION_FILLS[str(cell.value)]
                            cell.fill = PatternFill("solid",fgColor=bg)
                            cell.font = Font(color=fg,name="Arial",size=10)
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
    f'SmartBuyer · NovaRetail Solutions · Implantation & Supply v5.0 · {TODAY_STR} · Données : {date_stock_str}'
    f'</div>',
    unsafe_allow_html=True
)
