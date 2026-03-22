import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import date

# ══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Rapport Implantation · Carrefour",
    page_icon="🏪",
    layout="wide",
    initial_sidebar_state="expanded"
)

TODAY = date.today()
TODAY_STR = TODAY.strftime("%d %b %Y")
TODAY_FILE = TODAY.strftime("%Y%m%d")

# ══════════════════════════════════════════════════════════════════════════════
# CSS PREMIUM
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;700&display=swap');

:root{
  --bg:#f4f7fb;
  --bg-2:#eef3f9;
  --surface:#ffffff;
  --surface-soft:#f8fbff;
  --border:#dde6f0;
  --text:#0f172a;
  --muted:#64748b;

  --navy:#0b1b33;
  --navy-2:#142845;
  --blue:#2563eb;
  --blue-2:#3b82f6;
  --cyan:#0284c7;
  --green:#059669;
  --amber:#d97706;
  --red:#dc2626;

  --green-bg:#ecfdf5;
  --blue-bg:#eff6ff;
  --cyan-bg:#f0f9ff;
  --amber-bg:#fffbeb;
  --red-bg:#fef2f2;

  --green-bd:#a7f3d0;
  --blue-bd:#bfdbfe;
  --cyan-bd:#bae6fd;
  --amber-bd:#fcd34d;
  --red-bd:#fecaca;

  --shadow-sm:0 1px 2px rgba(15,23,42,.05);
  --shadow-md:0 8px 24px rgba(15,23,42,.07);
  --shadow-lg:0 18px 48px rgba(15,23,42,.12);

  --radius-sm:14px;
  --radius-md:20px;
  --radius-lg:24px;
}

html, body, [class*="css"] {
  font-family:'Inter',sans-serif !important;
  color:var(--text) !important;
}

body{
  background:linear-gradient(180deg,#f8fbff 0%, #f4f7fb 60%, #eef3f9 100%);
}

.main, section[data-testid="stMain"]{
  background:transparent !important;
}

.block-container{
  max-width:1600px !important;
  padding-top:1rem !important;
  padding-bottom:3rem !important;
}

header[data-testid="stHeader"], #MainMenu, footer{
  display:none !important;
}

section[data-testid="stSidebar"]{
  background:linear-gradient(180deg,#ffffff 0%, #f8fbff 100%) !important;
  border-right:1px solid var(--border) !important;
}
section[data-testid="stSidebar"] .block-container{
  padding-top:1rem !important;
}

/* TOPBAR */
.topbar{
  position:relative;
  overflow:hidden;
  border-radius:28px;
  padding:24px 28px;
  margin-bottom:22px;
  background:
    radial-gradient(circle at top right, rgba(59,130,246,.28), transparent 30%),
    radial-gradient(circle at left bottom, rgba(2,132,199,.18), transparent 26%),
    linear-gradient(135deg, #081528 0%, #0f1f38 40%, #162a46 100%);
  border:1px solid rgba(255,255,255,.08);
  box-shadow:0 18px 50px rgba(8,21,40,.24);
  color:#fff;
}
.topbar::after{
  content:"";
  position:absolute;
  inset:0;
  background:linear-gradient(180deg, rgba(255,255,255,.05), rgba(255,255,255,0));
  pointer-events:none;
}
.topbar-row{
  position:relative;
  z-index:2;
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap:24px;
  flex-wrap:wrap;
}
.topbar-left{
  display:flex;
  align-items:center;
  gap:18px;
}
.topbar-icon{
  width:62px;
  height:62px;
  border-radius:18px;
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:30px;
  background:linear-gradient(135deg, #3b82f6, #60a5fa);
  box-shadow:
    inset 0 1px 0 rgba(255,255,255,.25),
    0 12px 28px rgba(37,99,235,.28);
}
.topbar-title{
  font-size:28px;
  font-weight:900;
  line-height:1.05;
  letter-spacing:-.03em;
}
.topbar-sub{
  margin-top:6px;
  font-size:12px;
  color:#c9d6eb;
  font-family:'JetBrains Mono', monospace;
}
.topbar-meta{
  display:flex;
  gap:10px;
  flex-wrap:wrap;
  align-items:center;
}
.badge{
  border:1px solid rgba(255,255,255,.14);
  background:rgba(255,255,255,.07);
  color:#eef4ff;
  border-radius:999px;
  padding:8px 13px;
  font-size:12px;
  font-weight:700;
  backdrop-filter:blur(8px);
}
.date-pill{
  padding:8px 12px;
  border-radius:999px;
  background:rgba(59,130,246,.13);
  border:1px solid rgba(147,197,253,.24);
  color:#bfdbfe;
  font-size:12px;
  font-family:'JetBrains Mono', monospace;
  font-weight:700;
}

/* KPI */
.kpi-grid{
  display:grid;
  grid-template-columns:repeat(4,minmax(0,1fr));
  gap:16px;
  margin-bottom:18px;
}
.kpi-card{
  position:relative;
  overflow:hidden;
  background:linear-gradient(180deg, #ffffff 0%, #fbfdff 100%);
  border:1px solid var(--border);
  border-radius:22px;
  padding:18px 18px 16px 18px;
  box-shadow:var(--shadow-md);
}
.kpi-card::before{
  content:"";
  position:absolute;
  left:0; right:0; top:0;
  height:5px;
}
.kpi-card.green::before{background:linear-gradient(90deg,#10b981,#059669);}
.kpi-card.cyan::before{background:linear-gradient(90deg,#38bdf8,#0284c7);}
.kpi-card.red::before{background:linear-gradient(90deg,#ef4444,#dc2626);}
.kpi-card.blue::before{background:linear-gradient(90deg,#3b82f6,#2563eb);}

.kpi-label{
  font-size:11px;
  text-transform:uppercase;
  letter-spacing:.12em;
  color:var(--muted);
  font-weight:800;
  margin-bottom:12px;
}
.kpi-value{
  font-size:40px;
  line-height:1;
  font-weight:900;
  letter-spacing:-.03em;
}
.kpi-sub{
  margin-top:9px;
  font-size:12px;
  color:var(--muted);
  font-family:'JetBrains Mono', monospace;
}

/* BANNER */
.banner{
  border-radius:22px;
  padding:18px 20px;
  margin-bottom:18px;
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap:20px;
  flex-wrap:wrap;
  border:1px solid;
  box-shadow:var(--shadow-md);
}
.banner.red{
  background:linear-gradient(180deg,#fff7f7 0%, var(--red-bg) 100%);
  border-color:var(--red-bd);
}
.banner.blue{
  background:linear-gradient(180deg,#f7fbff 0%, var(--blue-bg) 100%);
  border-color:var(--blue-bd);
}
.banner.amber{
  background:linear-gradient(180deg,#fffdf6 0%, var(--amber-bg) 100%);
  border-color:var(--amber-bd);
}
.banner.green{
  background:linear-gradient(180deg,#f7fffb 0%, var(--green-bg) 100%);
  border-color:var(--green-bd);
}
.banner-title{
  font-size:16px;
  font-weight:900;
}
.banner-sub{
  margin-top:5px;
  font-size:12px;
  color:var(--muted);
}
.banner-big{
  font-size:44px;
  font-weight:900;
  line-height:1;
  letter-spacing:-.03em;
}

/* STRIP */
.strip{
  display:grid;
  grid-template-columns:repeat(6,minmax(0,1fr));
  gap:14px;
  margin:16px 0 24px 0;
}
.strip-card{
  background:linear-gradient(180deg,#ffffff 0%, #fbfdff 100%);
  border:1px solid var(--border);
  border-radius:18px;
  padding:14px;
  box-shadow:var(--shadow-sm);
}
.strip-tag{
  display:inline-block;
  padding:5px 9px;
  border-radius:999px;
  font-size:10px;
  font-weight:900;
  letter-spacing:.08em;
  margin-bottom:10px;
}
.tag-im{
  background:var(--blue-bg);
  border:1px solid var(--blue-bd);
  color:var(--blue);
}
.tag-lo{
  background:var(--green-bg);
  border:1px solid var(--green-bd);
  color:var(--green);
}
.strip-label{
  font-size:11px;
  text-transform:uppercase;
  letter-spacing:.08em;
  color:var(--muted);
  font-weight:800;
}
.strip-value{
  margin-top:8px;
  font-size:30px;
  line-height:1;
  font-weight:900;
  letter-spacing:-.03em;
}
.strip-sub{
  margin-top:6px;
  font-size:11px;
  color:var(--muted);
  font-family:'JetBrains Mono', monospace;
}

/* SECTIONS */
.section-title{
  margin:26px 0 12px 0;
  font-size:12px;
  text-transform:uppercase;
  letter-spacing:.16em;
  color:var(--muted);
  font-weight:900;
}

/* SCORECARD STORES */
.rag-grid{
  display:grid;
  grid-template-columns:repeat(auto-fill,minmax(270px,1fr));
  gap:16px;
  margin-bottom:24px;
}
.rag-card{
  position:relative;
  overflow:hidden;
  min-height:168px;
  border-radius:24px;
  padding:18px 18px 16px 18px;
  border:1px solid;
  box-shadow:var(--shadow-md);
  transition:transform .18s ease, box-shadow .18s ease, border-color .18s ease;
}
.rag-card:hover{
  transform:translateY(-3px);
  box-shadow:var(--shadow-lg);
}
.rag-card::before{
  content:"";
  position:absolute;
  inset:0;
  pointer-events:none;
  background:linear-gradient(180deg, rgba(255,255,255,.30), rgba(255,255,255,0));
}
.rag-card.good{
  background:linear-gradient(180deg,#f7fffb 0%, #ecfdf5 100%);
  border-color:var(--green-bd);
}
.rag-card.mid{
  background:linear-gradient(180deg,#fffef7 0%, #fffbeb 100%);
  border-color:var(--amber-bd);
}
.rag-card.bad{
  background:linear-gradient(180deg,#fff8f8 0%, #fef2f2 100%);
  border-color:var(--red-bd);
}
.rag-top{
  display:flex;
  align-items:flex-start;
  justify-content:space-between;
  gap:14px;
  margin-bottom:14px;
}
.rag-name{
  font-size:14px;
  line-height:1.3;
  font-weight:900;
  color:#0f172a;
  max-width:82%;
  white-space:normal;
  word-break:break-word;
}
.rag-chip{
  padding:6px 10px;
  border-radius:999px;
  font-size:10px;
  line-height:1;
  font-weight:900;
  letter-spacing:.08em;
  border:1px solid;
  white-space:nowrap;
}
.rag-chip.good{
  color:var(--green);
  background:rgba(5,150,105,.08);
  border-color:rgba(5,150,105,.18);
}
.rag-chip.mid{
  color:var(--amber);
  background:rgba(217,119,6,.08);
  border-color:rgba(217,119,6,.18);
}
.rag-chip.bad{
  color:var(--red);
  background:rgba(220,38,38,.08);
  border-color:rgba(220,38,38,.18);
}
.rag-main{
  display:flex;
  align-items:flex-end;
  gap:14px;
}
.rag-pct{
  font-size:46px;
  line-height:.92;
  font-weight:900;
  letter-spacing:-.05em;
}
.rag-progress{
  position:relative;
  height:10px;
  border-radius:999px;
  background:rgba(255,255,255,.75);
  border:1px solid rgba(148,163,184,.16);
  overflow:hidden;
  margin-top:14px;
}
.rag-progress-fill{
  height:100%;
  border-radius:999px;
}
.rag-foot{
  margin-top:13px;
  display:flex;
  justify-content:space-between;
  gap:10px;
  flex-wrap:wrap;
}
.rag-mini{
  font-size:11px;
  color:var(--muted);
  font-family:'JetBrains Mono', monospace;
}
.rag-rank{
  position:absolute;
  right:16px;
  bottom:14px;
  font-size:11px;
  font-weight:800;
  color:#94a3b8;
}

/* TABS */
.stTabs [data-baseweb="tab-list"]{
  gap:8px;
}
.stTabs [data-baseweb="tab"]{
  background:#fff;
  border:1px solid var(--border);
  border-radius:14px;
  padding:10px 16px;
  box-shadow:none;
}
.stTabs [aria-selected="true"]{
  background:linear-gradient(135deg,#0b1b33,#142845) !important;
  color:#fff !important;
  border-color:#0b1b33 !important;
}

/* BUTTONS */
.stDownloadButton > button{
  width:100% !important;
  border:none !important;
  border-radius:16px !important;
  padding:12px 16px !important;
  background:linear-gradient(135deg,#0b1b33,#142845) !important;
  color:#fff !important;
  font-weight:900 !important;
  box-shadow:0 10px 26px rgba(11,27,51,.22) !important;
}

/* METRIC */
div[data-testid="stMetric"]{
  background:linear-gradient(180deg,#ffffff 0%, #fbfdff 100%);
  border:1px solid var(--border);
  border-radius:18px;
  padding:12px 10px;
  box-shadow:var(--shadow-sm);
}

/* TABLE-LIKE BLOCK */
.soft-card{
  background:linear-gradient(180deg,#ffffff 0%, #fbfdff 100%);
  border:1px solid var(--border);
  border-radius:22px;
  padding:16px;
  box-shadow:var(--shadow-sm);
}

@media (max-width:1250px){
  .kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr));}
  .strip{grid-template-columns:repeat(2,minmax(0,1fr));}
}
@media (max-width:780px){
  .kpi-grid{grid-template-columns:1fr;}
  .strip{grid-template-columns:1fr;}
  .rag-grid{grid-template-columns:1fr;}
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def fix_encoding(df: pd.DataFrame) -> pd.DataFrame:
    try:
        if any("Ã" in str(c) for c in df.columns):
            df.columns = [c.encode("latin1").decode("utf-8", errors="replace") for c in df.columns]
    except Exception:
        pass
    return df


@st.cache_data(show_spinner=False)
def read_csv_smart(file_bytes: bytes, filename: str) -> pd.DataFrame:
    buf = io.BytesIO(file_bytes)

    for enc in ("latin1", "utf-8-sig", "cp1252"):
        for sep in (";", ",", "\t"):
            try:
                buf.seek(0)
                df = pd.read_csv(
                    buf,
                    sep=sep,
                    encoding=enc,
                    low_memory=False,
                    on_bad_lines="skip"
                )
                if df.shape[1] >= 3:
                    return fix_encoding(df)
            except Exception:
                continue

    buf.seek(0)
    df = pd.read_csv(
        buf,
        sep=None,
        engine="python",
        encoding="latin1",
        on_bad_lines="skip"
    )
    return fix_encoding(df)


@st.cache_data(show_spinner=False)
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\ufeff", "", regex=False)
        .str.replace("\xa0", " ", regex=False)
        .str.upper()
    )
    return df


@st.cache_data(show_spinner=False)
def load_t1(file_bytes: bytes, filename: str):
    buf = io.BytesIO(file_bytes)

    if filename.lower().endswith((".xlsx", ".xls")):
        df_peek = pd.read_excel(buf, header=None, nrows=1)
    else:
        buf.seek(0)
        try:
            df_peek = pd.read_csv(buf, header=None, nrows=1, sep=None, engine="python", encoding="latin1")
        except Exception:
            df_peek = None

    no_header = False
    if df_peek is not None and not df_peek.empty:
        first_val = str(df_peek.iloc[0, 0]).strip().replace(".0", "")
        no_header = first_val.isdigit()

    buf.seek(0)
    if filename.lower().endswith((".xlsx", ".xls")):
        if no_header:
            df = pd.read_excel(buf, header=None)
        else:
            df = pd.read_excel(buf)
            df = normalize_columns(df)
    else:
        if no_header:
            try:
                buf.seek(0)
                df = pd.read_csv(buf, header=None, sep=None, engine="python", encoding="latin1", on_bad_lines="skip")
            except Exception:
                df = read_csv_smart(file_bytes, filename)
                df = normalize_columns(df)
        else:
            df = read_csv_smart(file_bytes, filename)
            df = normalize_columns(df)

    if no_header or "ARTICLE" not in df.columns:
        if "ARTICLE" not in df.columns:
            cols = ["ARTICLE"] + [f"_COL{i}" for i in range(1, len(df.columns))]
            df.columns = cols

    if "ARTICLE" not in df.columns:
        found = ", ".join(df.columns.astype(str).tolist()[:10])
        return None, f"Colonne 'ARTICLE' introuvable. Colonnes détectées : {found}"

    df["SKU"] = df["ARTICLE"].astype(str).str.strip().str.zfill(8).str.slice(0, 8)
    df = df[df["SKU"].str.match(r"^\\d{8}$", na=False)].drop_duplicates(subset="SKU").copy()

    optional_cols = [
        ("LIBELLÉ ARTICLE", ""),
        ("FOURNISSEUR D'ORIGINE", ""),
        ("LIBELLÉ FOURNISSEUR ORIGINE", ""),
        ("MODE APPRO", ""),
        ("DATE CDE", ""),
        ("DATE LIV.", ""),
        ("SEMAINE RECEPTION", ""),
    ]
    for col, default in optional_cols:
        if col not in df.columns:
            df[col] = default

    df["SEMAINE RECEPTION"] = df["SEMAINE RECEPTION"].astype(str).str.strip().replace("nan", "")
    df["SEM_NUM"] = df["SEMAINE RECEPTION"].apply(
        lambda s: int(str(s).strip("Ss")) if str(s).strip("Ss").isdigit() else 99
    )
    df["ORIGINE"] = df["MODE APPRO"].apply(lambda m: "IM" if "IMPORT" in str(m).upper() else "LO")
    return df, None


@st.cache_data(show_spinner=False)
def load_stock(file_bytes: bytes, filename: str, sku_tuple: tuple):
    buf = io.BytesIO(file_bytes)

    if filename.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(buf)
    else:
        df = read_csv_smart(file_bytes, filename)

    df = fix_encoding(df)
    df = normalize_columns(df)

    required = {"LIBELLÉ SITE", "CODE ARTICLE", "NOUVEAU STOCK", "RAL"}
    missing = required - set(df.columns)
    if missing:
        found = ", ".join(df.columns.tolist()[:10])
        return None, f"Colonnes manquantes : {missing}. Colonnes détectées : {found}"

    optional_stock = ("FOUR.", "NOM FOURN.", "LIBELLÉ ARTICLE", "CODE ETAT", "CODE MARKETING")
    for col in optional_stock:
        if col not in df.columns:
            df[col] = ""

    df["SKU"] = df["CODE ARTICLE"].astype(str).str.strip().str.zfill(8).str.slice(0, 8)
    df = df[df["SKU"].isin(sku_tuple)].copy()

    df["NOUVEAU STOCK"] = pd.to_numeric(df["NOUVEAU STOCK"], errors="coerce").fillna(0)
    df["RAL"] = pd.to_numeric(df["RAL"], errors="coerce").fillna(0)

    for col in optional_stock:
        df[col] = df[col].astype(str).str.strip().replace("nan", "")

    df = df.rename(columns={
        "LIBELLÉ SITE": "Libellé site",
        "CODE ARTICLE": "Code article",
        "NOUVEAU STOCK": "Nouveau stock",
        "RAL": "Ral",
        "FOUR.": "Four.",
        "NOM FOURN.": "Nom fourn.",
        "LIBELLÉ ARTICLE": "Libellé article",
        "CODE ETAT": "Code etat",
        "CODE MARKETING": "Code marketing",
    })

    return df.drop_duplicates(subset=["Libellé site", "SKU"]), None


def sem_sort(s) -> int:
    try:
        return int(str(s).strip("Ss"))
    except Exception:
        return 99


def taux_implantation(df: pd.DataFrame) -> int:
    if len(df) == 0:
        return 0
    done = df["Statut"].eq("Implantation Terminée").sum()
    return int(done / len(df) * 100)


def safe_pct(num: int, den: int) -> int:
    return int(num / den * 100) if den else 0


def color_taux(t: int) -> str:
    if t >= 80:
        return "#059669"
    if t >= 65:
        return "#d97706"
    return "#dc2626"


def status_level(t: int) -> str:
    if t >= 80:
        return "good"
    if t >= 65:
        return "mid"
    return "bad"


def status_label(t: int) -> str:
    if t >= 80:
        return "ON TRACK"
    if t >= 65:
        return "WATCH"
    return "RISK"


def progress_color(t: int) -> str:
    if t >= 80:
        return "linear-gradient(90deg,#10b981,#059669)"
    if t >= 65:
        return "linear-gradient(90deg,#f59e0b,#d97706)"
    return "linear-gradient(90deg,#ef4444,#dc2626)"


def prep_display_table(df: pd.DataFrame, percent_cols=None):
    out = df.copy()
    percent_cols = percent_cols or []
    for col in percent_cols:
        if col in out.columns:
            out[col] = out[col].astype(int).astype(str) + "%"
    return out


def build_direction_summary(avg_impl, ct, ca, cal, tim, tlo):
    if avg_impl >= 80:
        status = "🟢 Situation maîtrisée"
        color = "#059669"
    elif avg_impl >= 60:
        status = "🟠 Situation sous tension"
        color = "#d97706"
    else:
        status = "🔴 Situation critique"
        color = "#dc2626"

    reco = []
    if cal > 0:
        reco.append("Escalader immédiatement les articles sans mouvement.")
    if ca > 0:
        reco.append("Sécuriser les dates de livraison et confirmer les ETA fournisseurs.")
    if tim < 70:
        reco.append("Prioriser le flux IMPORT : suivi fournisseurs, transit et dédouanement.")
    if tlo < 70:
        reco.append("Renforcer les relances sur les fournisseurs locaux.")
    if not reco:
        reco.append("Maintenir le rythme d’implantation et surveiller les magasins proches du seuil cible.")

    return status, color, reco


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def build_report_excel(
    detail_df: pd.DataFrame,
    pivot_df: pd.DataFrame,
    calendar_df: pd.DataFrame,
    top_alerts_df: pd.DataFrame,
    today_str: str,
    avg_impl: int,
    sku_im_total: int,
    sku_lo_total: int,
    tim: int,
    tlo: int,
    ct: int,
    ca: int,
    cal: int
) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)

    C = dict(
        dark="0F172A", grey="F8FAFC", border="E2E8F0", white="FFFFFF", muted="64748B",
        green="059669", green_bg="ECFDF5",
        blue="2563EB", blue_bg="EFF6FF",
        cyan="0284C7", cyan_bg="F0F9FF",
        red="DC2626", red_bg="FEF2F2",
        amber="D97706", amber_bg="FFFBEB"
    )

    thin = Side(style="thin", color=C["border"])

    def font(color=C["dark"], size=10, bold=False):
        return Font(name="Arial", size=size, bold=bold, color=color)

    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def border():
        return Border(left=thin, right=thin, top=thin, bottom=thin)

    def center(wrap=False):
        return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

    def left(wrap=False):
        return Alignment(horizontal="left", vertical="center", wrap_text=wrap)

    def write_header(ws, title, sub=""):
        ws.sheet_view.showGridLines = False
        ws.merge_cells("B1:L1")
        ws["B1"] = title
        ws["B1"].font = font(C["white"], 20, True)
        ws["B1"].fill = fill(C["dark"])
        ws["B1"].alignment = left()
        ws.row_dimensions[1].height = 34

        if sub:
            ws.merge_cells("B2:L2")
            ws["B2"] = sub
            ws["B2"].font = font(C["muted"], 10, False)
            ws["B2"].fill = fill(C["grey"])
            ws["B2"].alignment = left()
            ws.row_dimensions[2].height = 22

    def write_kpi_box(ws, cell_ref, title_txt, value_txt, fill_color, font_color="FFFFFF"):
        cell = ws[cell_ref]
        cell.value = f"{title_txt}\\n{value_txt}"
        cell.font = Font(name="Arial", size=12, bold=True, color=font_color)
        cell.fill = fill(fill_color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border()

    def write_table(ws, df, start_row, columns, header_fill="0F172A"):
        row = start_row
        for i, (display, key, width, align) in enumerate(columns, start=2):
            c = ws.cell(row, i)
            c.value = display
            c.font = font(C["white"], 9, True)
            c.fill = fill(header_fill)
            c.alignment = center(True)
            c.border = border()
            ws.column_dimensions[get_column_letter(i)].width = width

        for r_idx, (_, rec) in enumerate(df.iterrows(), start=row + 1):
            for i, (_, key, _, align) in enumerate(columns, start=2):
                c = ws.cell(r_idx, i)
                val = rec[key] if key in rec.index else ""
                c.value = val
                c.font = font(C["dark"], 9, False)
                c.fill = fill("FFFFFF" if (r_idx - row) % 2 else "F8FAFC")
                c.alignment = center() if align == "c" else left()
                c.border = border()

    status, _, reco = build_direction_summary(avg_impl, ct, ca, cal, tim, tlo)

    ws0 = wb.create_sheet("Synthèse Direction")
    write_header(ws0, f"SYNTHÈSE DIRECTION — {today_str}", "Lecture exécutive du niveau d’implantation réseau")

    ws0.row_dimensions[4].height = 45
    ws0.row_dimensions[5].height = 45

    ws0.merge_cells("B4:C5")
    ws0.merge_cells("D4:E5")
    ws0.merge_cells("F4:G5")
    ws0.merge_cells("H4:I5")

    write_kpi_box(ws0, "B4", "Taux Réseau", f"{avg_impl}%", C["blue"])
    write_kpi_box(ws0, "D4", "Terminée", f"{ct}", C["green"])
    write_kpi_box(ws0, "F4", "Attente Livraison", f"{ca}", C["cyan"])
    write_kpi_box(ws0, "H4", "Alertes Critiques", f"{cal}", C["red"])

    ws0.merge_cells("B8:I8")
    ws0["B8"] = f"Diagnostic : {status}"
    ws0["B8"].font = font(C["dark"], 14, True)
    ws0["B8"].fill = fill(C["amber_bg"] if "tension" in status.lower() else (C["red_bg"] if "critique" in status.lower() else C["green_bg"]))
    ws0["B8"].alignment = left()
    ws0["B8"].border = border()

    ws0["B10"] = "Indicateurs par origine"
    ws0["B10"].font = font(C["dark"], 12, True)

    ind_df = pd.DataFrame([
        ["Références IMPORT", sku_im_total],
        ["Taux IM (%)", tim],
        ["Références LOCAL", sku_lo_total],
        ["Taux LO (%)", tlo],
    ], columns=["Indicateur", "Valeur"])

    write_table(
        ws0,
        ind_df,
        11,
        [("Indicateur", "Indicateur", 30, "l"), ("Valeur", "Valeur", 12, "c")],
        header_fill=C["blue"]
    )

    start_reco_row = 18
    ws0["B17"] = "Recommandations prioritaires"
    ws0["B17"].font = font(C["dark"], 12, True)
    for i, r in enumerate(reco, start=start_reco_row):
        ws0.merge_cells(f"B{i}:I{i}")
        ws0[f"B{i}"] = f"• {r}"
        ws0[f"B{i}"].font = font(C["dark"], 10, False)
        ws0[f"B{i}"].alignment = left(True)
        ws0[f"B{i}"].fill = fill(C["grey"])
        ws0[f"B{i}"].border = border()

    if not top_alerts_df.empty:
        ws0["B23"] = "Top articles à risque"
        ws0["B23"].font = font(C["dark"], 12, True)
        write_table(
            ws0,
            top_alerts_df,
            24,
            [
                ("SKU", "SKU", 12, "c"),
                ("Libellé article", "Libellé article", 38, "l"),
                ("Nb magasins", "Nb Magasins", 14, "c"),
            ],
            header_fill=C["red"]
        )

    ws1 = wb.create_sheet("Résumé Exécutif")
    write_header(
        ws1,
        f"RAPPORT IMPLANTATION — {today_str}",
        f"Taux moyen réseau : {avg_impl}%"
    )

    summary = pd.DataFrame([
        ["Implantation Terminée", ct],
        ["En Attente Livraison", ca],
        ["Alerte Aucun Mouvement", cal],
        ["Références IMPORT", sku_im_total],
        ["Références LOCAL", sku_lo_total],
        ["Taux IM (%)", tim],
        ["Taux LO (%)", tlo],
    ], columns=["Indicateur", "Valeur"])

    write_table(
        ws1,
        summary,
        4,
        [
            ("Indicateur", "Indicateur", 32, "l"),
            ("Valeur", "Valeur", 16, "c"),
        ],
        header_fill=C["dark"]
    )

    ws2 = wb.create_sheet("Vue Magasins")
    write_header(ws2, "VUE PAR MAGASIN", today_str)
    write_table(
        ws2,
        pivot_df,
        4,
        [
            ("Magasin", "Magasin", 28, "l"),
            ("Terminée", "Implantation Terminée", 14, "c"),
            ("Attente", "En Attente Livraison", 14, "c"),
            ("Alerte", "Alerte Aucun Mouvement", 14, "c"),
            ("Total", "Total", 12, "c"),
            ("Taux (%)", "Taux (%)", 12, "c"),
        ],
        header_fill=C["blue"]
    )

    ws3 = wb.create_sheet("Alertes")
    write_header(ws3, "ALERTES & ACTIONS", today_str)
    alerts = detail_df[detail_df["Statut"] != "Implantation Terminée"].copy()
    write_table(
        ws3,
        alerts,
        4,
        [
            ("Magasin", "Magasin", 24, "l"),
            ("SKU", "SKU", 12, "c"),
            ("Libellé article", "Libellé article", 40, "l"),
            ("Origine", "Origine", 10, "c"),
            ("Mode Appro", "Mode Appro", 18, "l"),
            ("Sem. Réception", "Sem. Réception", 14, "c"),
            ("Date Livraison", "Date Livraison", 14, "c"),
            ("Code etat", "Code etat", 12, "c"),
            ("Stock", "Stock", 10, "c"),
            ("RAL", "RAL", 10, "c"),
            ("Statut", "Statut", 24, "c"),
        ],
        header_fill=C["red"]
    )

    ws4 = wb.create_sheet("Plan Action")
    write_header(ws4, "PLAN D'ACTION", today_str)
    pa = detail_df[detail_df["Statut"].isin(["Alerte Aucun Mouvement", "En Attente Livraison"])].copy()
    pa["Action recommandée"] = np.where(
        pa["Statut"].eq("Alerte Aucun Mouvement"),
        "Escalader / vérifier commande / informer magasin",
        "Confirmer date livraison / préparer réception"
    )
    write_table(
        ws4,
        pa,
        4,
        [
            ("Magasin", "Magasin", 24, "l"),
            ("SKU", "SKU", 12, "c"),
            ("Libellé article", "Libellé article", 40, "l"),
            ("Origine", "Origine", 10, "c"),
            ("Statut", "Statut", 24, "c"),
            ("Action recommandée", "Action recommandée", 44, "l"),
        ],
        header_fill=C["amber"]
    )

    ws5 = wb.create_sheet("Calendrier Flux")
    write_header(ws5, "CALENDRIER FLUX", today_str)
    write_table(
        ws5,
        calendar_df,
        4,
        [
            ("Sem. Réception", "Sem. Réception", 16, "c"),
            ("Origine", "Origine", 10, "c"),
            ("Articles", "Articles", 12, "c"),
            ("Terminé", "Terminé", 12, "c"),
            ("Attente", "Attente", 12, "c"),
            ("Alerte", "Alerte", 12, "c"),
        ],
        header_fill=C["cyan"]
    )

    ws6 = wb.create_sheet("Détail Complet")
    write_header(ws6, "DETAIL COMPLET", today_str)
    write_table(
        ws6,
        detail_df,
        4,
        [
            ("Magasin", "Magasin", 24, "l"),
            ("SKU", "SKU", 12, "c"),
            ("Libellé article", "Libellé article", 40, "l"),
            ("Origine", "Origine", 10, "c"),
            ("Mode Appro", "Mode Appro", 18, "l"),
            ("Sem. Réception", "Sem. Réception", 14, "c"),
            ("Date Livraison", "Date Livraison", 14, "c"),
            ("Code etat", "Code etat", 12, "c"),
            ("Stock", "Stock", 10, "c"),
            ("RAL", "RAL", 10, "c"),
            ("Statut", "Statut", 24, "c"),
        ],
        header_fill=C["dark"]
    )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# TOPBAR
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="topbar">
  <div class="topbar-row">
    <div class="topbar-left">
      <div class="topbar-icon">🏪</div>
      <div>
        <div class="topbar-title">Rapport Implantation</div>
        <div class="topbar-sub">Suivi Nouvelles Références · Réseau Magasins · Stock & Flux</div>
      </div>
    </div>
    <div class="topbar-meta">
      <div class="badge">DIRECTION SUPPLY</div>
      <div class="badge">OPERATIONS REVIEW</div>
      <div class="date-pill">{TODAY_STR}</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — CHARGEMENT
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📁 Chargement")
    st.markdown("Charge d'abord le fichier T1 puis les extractions stock.")
    t1_file = st.file_uploader("T1 — Nouvelles Références", type=["csv", "xlsx"], key="t1")
    stock_files = st.file_uploader(
        "Extractions Stock (multi-fichiers)",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        key="stocks"
    )

if not t1_file:
    st.info("Charge le fichier T1 pour démarrer.")
    st.stop()

with st.spinner("Lecture du fichier T1…"):
    t1_raw, t1_err = load_t1(t1_file.read(), t1_file.name)

if t1_err:
    st.error(f"T1 : {t1_err}")
    st.stop()

SKU_TUPLE = tuple(sorted(t1_raw["SKU"].unique()))
sku_im_total = int((t1_raw["ORIGINE"] == "IM").sum())
sku_lo_total = int((t1_raw["ORIGINE"] == "LO").sum())

T1_KEEP = [
    "LIBELLÉ ARTICLE",
    "LIBELLÉ FOURNISSEUR ORIGINE",
    "MODE APPRO",
    "SEMAINE RECEPTION",
    "DATE LIV.",
    "ORIGINE",
    "SEM_NUM"
]
t1_idx = t1_raw.set_index("SKU")[T1_KEEP]

if not stock_files:
    st.info("Charge les extractions stock dans la barre latérale.")
    st.stop()

frames = []
with st.spinner(f"Lecture de {len(stock_files)} fichier(s) stock…"):
    for uf in stock_files:
        raw = uf.read()
        df_tmp, err = load_stock(raw, uf.name, SKU_TUPLE)
        if err:
            st.error(f"{uf.name} : {err}")
        else:
            frames.append(df_tmp)

if not frames:
    st.error("Aucun fichier stock valide.")
    st.stop()

df_stock = pd.concat(frames, ignore_index=True).drop_duplicates(subset=["Libellé site", "SKU"])
magasins_list = sorted(df_stock["Libellé site"].dropna().unique())

# ══════════════════════════════════════════════════════════════════════════════
# FILTRES
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("---")
    st.markdown("### 🔍 Filtres")
    mag_sel = st.multiselect("Magasins", magasins_list, default=magasins_list)
    origine_sel = st.multiselect("Origine", ["IM", "LO"], default=["IM", "LO"])

    sem_dispo = sorted(
        [s for s in t1_raw["SEMAINE RECEPTION"].unique() if s and s not in ("nan", "")],
        key=sem_sort
    )
    sem_sel = st.multiselect("Semaine réception", sem_dispo, default=sem_dispo)

    mode_dispo = sorted([m for m in t1_raw["MODE APPRO"].unique() if m and m not in ("nan", "")])
    mode_sel = st.multiselect("Mode Appro", mode_dispo, default=mode_dispo)

if not mag_sel:
    st.warning("Sélectionne au moins un magasin.")
    st.stop()

mc1, mc2, mc3 = st.columns([6, 1, 1])
with mc1:
    mag_main = st.multiselect("🏪 Magasins affichés", magasins_list, default=mag_sel, key="mag_main")
with mc2:
    if st.button("Tous", use_container_width=True):
        st.session_state["mag_main"] = magasins_list
with mc3:
    if st.button("Aucun", use_container_width=True):
        st.session_state["mag_main"] = []

mag_actifs = st.session_state.get("mag_main", mag_main if mag_main else mag_sel)
if not mag_actifs:
    st.warning("Sélectionne au moins un magasin.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# CALCULS
# ══════════════════════════════════════════════════════════════════════════════
sku_mask = (
    t1_raw["ORIGINE"].isin(origine_sel)
    & (t1_raw["SEMAINE RECEPTION"].isin(sem_sel) if sem_sel else True)
    & (t1_raw["MODE APPRO"].isin(mode_sel) if mode_sel else True)
)
sku_scope = t1_raw.loc[sku_mask, "SKU"].unique()
total_sku_sel = len(sku_scope)

if total_sku_sel == 0:
    st.warning("Aucun article ne correspond aux filtres.")
    st.stop()

base_df = pd.DataFrame(
    pd.MultiIndex.from_product([mag_actifs, sku_scope], names=["Libellé site", "SKU"]).tolist(),
    columns=["Libellé site", "SKU"]
)

stock_scope = df_stock[df_stock["Libellé site"].isin(mag_actifs) & df_stock["SKU"].isin(sku_scope)]

merged = base_df.merge(
    stock_scope[["Libellé site", "SKU", "Nouveau stock", "Ral", "Code etat", "Code marketing", "Libellé article"]],
    on=["Libellé site", "SKU"],
    how="left"
)
merged["Nouveau stock"] = merged["Nouveau stock"].fillna(0)
merged["Ral"] = merged["Ral"].fillna(0)
merged["Code etat"] = merged["Code etat"].fillna("").astype(str)

merged = merged.merge(
    t1_idx.reset_index().rename(columns={
        "LIBELLÉ ARTICLE": "T1_lib",
        "LIBELLÉ FOURNISSEUR ORIGINE": "Fournisseur",
        "MODE APPRO": "Mode Appro",
        "SEMAINE RECEPTION": "Sem. Réception",
        "DATE LIV.": "Date Livraison",
        "ORIGINE": "Origine",
        "SEM_NUM": "SEM_NUM",
    }),
    on="SKU",
    how="left"
)

merged["Libellé article"] = merged["Libellé article"].fillna("").astype(str)
merged["Libellé article"] = np.where(
    merged["Libellé article"].eq(""),
    merged["T1_lib"],
    merged["Libellé article"]
)
merged.drop(columns="T1_lib", inplace=True)

conds = [
    merged["Nouveau stock"] > 0,
    (merged["Nouveau stock"] == 0) & (merged["Ral"] > 0),
]
choices = ["Implantation Terminée", "En Attente Livraison"]
merged["Statut"] = np.select(conds, choices, default="Alerte Aucun Mouvement")
merged["Etat Actif"] = merged["Code etat"] == "2"

detail_df = merged.rename(columns={
    "Libellé site": "Magasin",
    "Nouveau stock": "Stock",
    "Ral": "RAL",
})

S_ORDER = ["Implantation Terminée", "En Attente Livraison", "Alerte Aucun Mouvement"]
S_COLORS = {
    "Implantation Terminée": "#059669",
    "En Attente Livraison": "#0284c7",
    "Alerte Aucun Mouvement": "#dc2626",
}

pivot = (
    detail_df.groupby(["Magasin", "Statut"]).size()
    .unstack(fill_value=0)
    .reindex(columns=S_ORDER, fill_value=0)
    .reset_index()
)
pivot.columns.name = None
pivot["Total"] = total_sku_sel
pivot["Taux (%)"] = (pivot["Implantation Terminée"] / total_sku_sel * 100).round(0).astype(int)

total_cells = len(mag_actifs) * total_sku_sel
ct = int(pivot["Implantation Terminée"].sum())
ca = int(pivot["En Attente Livraison"].sum())
cal = int(pivot["Alerte Aucun Mouvement"].sum())
avg_impl = int(pivot["Taux (%)"].mean()) if not pivot.empty else 0

df_im = detail_df[detail_df["Origine"] == "IM"]
df_lo = detail_df[detail_df["Origine"] == "LO"]
df_attente = detail_df[detail_df["Statut"] == "En Attente Livraison"]
df_alerte = detail_df[detail_df["Statut"] == "Alerte Aucun Mouvement"]

attente_im = df_attente[df_attente["Origine"] == "IM"]
attente_lo = df_attente[df_attente["Origine"] == "LO"]
alerte_im = df_alerte[df_alerte["Origine"] == "IM"]
alerte_lo = df_alerte[df_alerte["Origine"] == "LO"]

im_alerte = int((df_im["Statut"] == "Alerte Aucun Mouvement").sum())
lo_alerte = int((df_lo["Statut"] == "Alerte Aucun Mouvement").sum())
total_actions = ca + cal

tim = taux_implantation(df_im)
tlo = taux_implantation(df_lo)

status_txt, status_color, reco_list = build_direction_summary(avg_impl, ct, ca, cal, tim, tlo)

top_pb = (
    df_alerte.groupby(["SKU", "Libellé article"])["Magasin"]
    .count()
    .reset_index()
    .rename(columns={"Magasin": "Nb Magasins"})
    .sort_values("Nb Magasins", ascending=False)
    .head(10)
)

# ══════════════════════════════════════════════════════════════════════════════
# KPI HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card green">
    <div class="kpi-label">Implantation Terminée</div>
    <div class="kpi-value" style="color:#059669">{ct}</div>
    <div class="kpi-sub">{safe_pct(ct,total_cells)}% du total réseau</div>
  </div>
  <div class="kpi-card cyan">
    <div class="kpi-label">En Attente Livraison</div>
    <div class="kpi-value" style="color:#0284c7">{ca}</div>
    <div class="kpi-sub">{safe_pct(ca,total_cells)}% du total réseau</div>
  </div>
  <div class="kpi-card red">
    <div class="kpi-label">Alerte Aucun Mouvement</div>
    <div class="kpi-value" style="color:#dc2626">{cal}</div>
    <div class="kpi-sub">{safe_pct(cal,total_cells)}% du total réseau</div>
  </div>
  <div class="kpi-card blue">
    <div class="kpi-label">Taux Moyen Réseau</div>
    <div class="kpi-value" style="color:#2563eb">{avg_impl}%</div>
    <div class="kpi-sub">{len(mag_actifs)} magasins · {total_sku_sel} SKU</div>
  </div>
</div>
""", unsafe_allow_html=True)

if total_actions > 0:
    st.markdown(f"""
    <div class="banner red">
      <div>
        <div class="banner-title" style="color:#dc2626">⚠️ Actions requises</div>
        <div class="banner-sub">
          {cal} sans mouvement · {ca} en attente livraison · IM alertes : {len(alerte_im)} · LO alertes : {len(alerte_lo)}
        </div>
      </div>
      <div class="banner-big" style="color:#dc2626">{total_actions}</div>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <div class="banner green">
      <div>
        <div class="banner-title" style="color:#059669">✅ Réseau sous contrôle</div>
        <div class="banner-sub">Aucune action urgente détectée sur le périmètre filtré.</div>
      </div>
      <div class="banner-big" style="color:#059669">0</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown(f"""
<div class="strip">
  <div class="strip-card">
    <div class="strip-tag tag-im">IMPORT</div>
    <div class="strip-label">Références</div>
    <div class="strip-value" style="color:#2563eb">{sku_im_total}</div>
    <div class="strip-sub">SKU à implanter</div>
  </div>
  <div class="strip-card">
    <div class="strip-tag tag-im">IMPORT</div>
    <div class="strip-label">Taux implanté</div>
    <div class="strip-value" style="color:{color_taux(tim)}">{tim}%</div>
    <div class="strip-sub">stock présent</div>
  </div>
  <div class="strip-card">
    <div class="strip-tag tag-im">IMPORT</div>
    <div class="strip-label">Alertes</div>
    <div class="strip-value" style="color:#dc2626">{im_alerte}</div>
    <div class="strip-sub">à escalader</div>
  </div>
  <div class="strip-card">
    <div class="strip-tag tag-lo">LOCAL</div>
    <div class="strip-label">Références</div>
    <div class="strip-value" style="color:#059669">{sku_lo_total}</div>
    <div class="strip-sub">SKU à implanter</div>
  </div>
  <div class="strip-card">
    <div class="strip-tag tag-lo">LOCAL</div>
    <div class="strip-label">Taux implanté</div>
    <div class="strip-value" style="color:{color_taux(tlo)}">{tlo}%</div>
    <div class="strip-sub">stock présent</div>
  </div>
  <div class="strip-card">
    <div class="strip-tag tag-lo">LOCAL</div>
    <div class="strip-label">Alertes</div>
    <div class="strip-value" style="color:#dc2626">{lo_alerte}</div>
    <div class="strip-sub">à relancer</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SCORECARD MAGASINS — VERSION PREMIUM
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">Scorecard magasins</div>', unsafe_allow_html=True)

pivot_sorted = pivot.sort_values("Taux (%)", ascending=False).reset_index(drop=True)
rag_html = '<div class="rag-grid">'

for i, row in pivot_sorted.iterrows():
    t_ = int(row["Taux (%)"])
    level = status_level(t_)
    label = status_label(t_)
    pct_color = color_taux(t_)
    prog = progress_color(t_)

    rag_html += f"""
    <div class="rag-card {level}">
      <div class="rag-top">
        <div class="rag-name">{row['Magasin']}</div>
        <div class="rag-chip {level}">{label}</div>
      </div>

      <div class="rag-main">
        <div class="rag-pct" style="color:{pct_color}">{t_}%</div>
      </div>

      <div class="rag-progress">
        <div class="rag-progress-fill" style="width:{min(t_,100)}%; background:{prog};"></div>
      </div>

      <div class="rag-foot">
        <div class="rag-mini">{int(row['Implantation Terminée'])} terminés</div>
        <div class="rag-mini">{int(row['En Attente Livraison'])} attente</div>
        <div class="rag-mini">{int(row['Alerte Aucun Mouvement'])} alertes</div>
      </div>

      <div class="rag-rank">RANK #{i+1}</div>
    </div>
    """

rag_html += "</div>"
st.markdown(rag_html, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab0, tab1, tab2, tab3, tab4 = st.tabs([
    "🧠 Synthèse Direction",
    "📊 Vue Globale",
    "🚨 Alertes & Actions",
    "🗓️ Calendrier Flux",
    "📋 Plan d'Action"
])

PLOTLY_BASE = dict(
    paper_bgcolor="#ffffff",
    plot_bgcolor="#ffffff",
    font=dict(family="Inter", color="#64748b", size=12),
    margin=dict(l=20, r=20, t=50, b=20)
)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 0 — SYNTHÈSE DIRECTION
# ══════════════════════════════════════════════════════════════════════════════
with tab0:
    st.markdown('<div class="section-title">Synthèse exécutive</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Taux réseau", f"{avg_impl}%")
    c2.metric("Alertes critiques", cal)
    c3.metric("En attente livraison", ca)

    banner_class = "green" if "maîtrisée" in status_txt.lower() else ("amber" if "tension" in status_txt.lower() else "red")

    st.markdown(f"""
    <div class="banner {banner_class}">
      <div>
        <div class="banner-title" style="color:{status_color}">{status_txt}</div>
        <div class="banner-sub">{cal} articles sans mouvement · {ca} en attente livraison · {avg_impl}% implanté</div>
      </div>
      <div class="banner-big" style="color:{status_color}">{avg_impl}%</div>
    </div>
    """, unsafe_allow_html=True)

    left, right = st.columns([3, 2])

    with left:
        st.markdown("### 🔎 Top problèmes")
        if top_pb.empty:
            st.success("Aucun article critique détecté.")
        else:
            st.dataframe(top_pb, use_container_width=True, hide_index=True)

    with right:
        st.markdown("### 🎯 Recommandations")
        for r in reco_list:
            st.markdown(f"- {r}")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — VUE GLOBALE
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    c1, c2 = st.columns([3, 2])

    with c1:
        mel = pivot.melt(
            id_vars="Magasin",
            value_vars=list(S_COLORS.keys()),
            var_name="Statut",
            value_name="N"
        )
        fig = px.bar(
            mel,
            x="Magasin",
            y="N",
            color="Statut",
            color_discrete_map=S_COLORS,
            barmode="stack",
            title="Situation par magasin"
        )
        fig.update_traces(
            textposition="inside",
            texttemplate="%{y}",
            textfont_size=11,
            textfont_color="white"
        )
        fig.update_layout(
            **PLOTLY_BASE,
            height=430,
            legend=dict(orientation="h", y=-0.2),
            xaxis=dict(gridcolor="#f1f5f9"),
            yaxis=dict(gridcolor="#f1f5f9"),
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig_d = go.Figure(go.Pie(
            labels=list(S_COLORS.keys()),
            values=[ct, ca, cal],
            hole=0.68,
            marker=dict(colors=list(S_COLORS.values()), line=dict(color="#fff", width=4)),
            textfont=dict(size=11)
        ))
        fig_d.add_annotation(
            text=f"<b>{avg_impl}%</b><br>implanté",
            x=0.5, y=0.5,
            font=dict(size=20, color="#0f172a", family="Inter"),
            showarrow=False
        )
        fig_d.update_layout(
            **PLOTLY_BASE,
            height=430,
            title="Répartition globale",
            legend=dict(orientation="v", x=1.01)
        )
        st.plotly_chart(fig_d, use_container_width=True)

    st.markdown('<div class="section-title">Détail par magasin</div>', unsafe_allow_html=True)
    disp_cols = [
        "Magasin",
        "Implantation Terminée",
        "En Attente Livraison",
        "Alerte Aucun Mouvement",
        "Total",
        "Taux (%)"
    ]
    pivot_display = prep_display_table(pivot[disp_cols], percent_cols=["Taux (%)"])
    st.dataframe(
        pivot_display,
        use_container_width=True,
        hide_index=True,
        height=min(600, 60 + len(mag_actifs) * 42)
    )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — ALERTES & ACTIONS
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    filt = st.radio(
        "Filtre alertes",
        ["Toutes les alertes", "🚨 Aucun Mouvement", "🚚 Attente Livraison"],
        horizontal=True
    )

    ACOLS = [
        "Magasin", "SKU", "Libellé article", "Origine", "Mode Appro",
        "Sem. Réception", "Date Livraison", "Code etat", "Stock", "RAL", "Statut"
    ]

    ALERT_SECTIONS = {
        "🚨 Aucun Mouvement": (
            df_alerte,
            "#dc2626",
            "Aucun mouvement — Stock = 0 · RAL = 0",
            "Escalader fournisseur · Vérifier commande · Informer magasin",
        ),
        "🚚 Attente Livraison": (
            df_attente,
            "#0284c7",
            "En attente livraison — RAL présent · Stock = 0",
            "Confirmer date livraison · Préparer réception magasin",
        ),
    }

    for key, (df_a, hex_color, title_txt, action_txt) in ALERT_SECTIONS.items():
        if filt not in ("Toutes les alertes", key):
            continue

        banner_type = "red" if "Aucun" in title_txt else "blue"
        st.markdown(f"""
        <div class="banner {banner_type}">
          <div>
            <div class="banner-title" style="color:{hex_color}">{title_txt}</div>
            <div class="banner-sub">Action : {action_txt}</div>
          </div>
          <div class="banner-big" style="color:{hex_color}">{len(df_a)}</div>
        </div>
        """, unsafe_allow_html=True)

        if df_a.empty:
            st.success("Aucune ligne dans cette catégorie.")
            continue

        left, right = st.columns([2, 3])

        with left:
            top = (
                df_a.groupby(["SKU", "Libellé article"])["Magasin"]
                .count()
                .reset_index()
                .rename(columns={"Magasin": "Nb Magasins"})
                .sort_values("Nb Magasins", ascending=False)
                .head(10)
            )
            if not top.empty:
                top["lbl"] = top["SKU"] + " – " + top["Libellé article"].astype(str).str[:30]
                fig_t = go.Figure(go.Bar(
                    x=top["Nb Magasins"],
                    y=top["lbl"],
                    orientation="h",
                    marker=dict(color=hex_color),
                    text=top["Nb Magasins"],
                    textposition="outside"
                ))
                fig_t.update_layout(
                    **PLOTLY_BASE,
                    height=max(260, len(top) * 36),
                    title="Top SKU impactés",
                    xaxis=dict(gridcolor="#f1f5f9"),
                    yaxis=dict(tickfont_size=10)
                )
                st.plotly_chart(fig_t, use_container_width=True)

        with right:
            top_m = (
                df_a.groupby("Magasin")["SKU"]
                .count()
                .reset_index()
                .rename(columns={"SKU": "Nb SKU"})
                .sort_values("Nb SKU", ascending=False)
            )
            fig_m = go.Figure(go.Bar(
                x=top_m["Magasin"],
                y=top_m["Nb SKU"],
                marker=dict(color=hex_color),
                text=top_m["Nb SKU"],
                textposition="outside"
            ))
            fig_m.update_layout(
                **PLOTLY_BASE,
                height=max(260, len(top_m) * 40),
                title="Alertes par magasin",
                xaxis=dict(gridcolor="#f1f5f9"),
                yaxis=dict(gridcolor="#f1f5f9")
            )
            st.plotly_chart(fig_m, use_container_width=True)

        with st.expander(f"Voir le détail — {len(df_a)} ligne(s)", expanded=False):
            st.dataframe(
                df_a[ACOLS].sort_values(["Magasin", "Sem. Réception"]).reset_index(drop=True),
                use_container_width=True,
                hide_index=True
            )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — CALENDRIER FLUX
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    cal_df = detail_df[detail_df["Sem. Réception"].astype(str).str.match(r"^S\\d+$", na=False)].copy()

    if cal_df.empty:
        st.info("Aucune donnée de semaine disponible.")
        calendar_export_df = pd.DataFrame(columns=["Sem. Réception", "Origine", "Articles", "Terminé", "Attente", "Alerte"])
    else:
        cal_df["SEM_NUM"] = cal_df["Sem. Réception"].apply(sem_sort)
        sem_order = sorted(cal_df["Sem. Réception"].unique(), key=sem_sort)

        c1, c2 = st.columns(2)

        with c1:
            ss = (
                cal_df.groupby(["Sem. Réception", "SEM_NUM", "Statut"])
                .size()
                .reset_index(name="N")
                .sort_values("SEM_NUM")
            )
            fig_s = px.bar(
                ss,
                x="Sem. Réception",
                y="N",
                color="Statut",
                color_discrete_map=S_COLORS,
                barmode="stack",
                category_orders={"Sem. Réception": sem_order},
                title="Articles par semaine & statut"
            )
            fig_s.update_traces(textposition="inside", texttemplate="%{y}", textfont_size=10, textfont_color="white")
            fig_s.update_layout(
                **PLOTLY_BASE,
                height=390,
                xaxis=dict(gridcolor="#f1f5f9"),
                yaxis=dict(gridcolor="#f1f5f9"),
                legend=dict(orientation="h", y=-0.2)
            )
            st.plotly_chart(fig_s, use_container_width=True)

        with c2:
            os_df = (
                cal_df.groupby(["Origine", "Sem. Réception", "SEM_NUM"])
                .size()
                .reset_index(name="N")
                .sort_values("SEM_NUM")
            )
            fig_o = px.bar(
                os_df,
                x="Sem. Réception",
                y="N",
                color="Origine",
                barmode="group",
                color_discrete_map={"IM": "#2563eb", "LO": "#059669"},
                category_orders={"Sem. Réception": sem_order},
                title="IM vs LO par semaine"
            )
            fig_o.update_traces(textposition="outside", texttemplate="%{y}", textfont_size=10)
            fig_o.update_layout(
                **PLOTLY_BASE,
                height=390,
                xaxis=dict(gridcolor="#f1f5f9"),
                yaxis=dict(gridcolor="#f1f5f9"),
                legend=dict(orientation="h", y=-0.2)
            )
            st.plotly_chart(fig_o, use_container_width=True)

        st.markdown('<div class="section-title">Détail par semaine</div>', unsafe_allow_html=True)
        calendar_export_df = (
            cal_df.groupby(["Sem. Réception", "SEM_NUM", "Origine"]).agg(
                Articles=("SKU", "nunique"),
                Terminé=("Statut", lambda x: (x == "Implantation Terminée").sum()),
                Attente=("Statut", lambda x: (x == "En Attente Livraison").sum()),
                Alerte=("Statut", lambda x: (x == "Alerte Aucun Mouvement").sum()),
            )
            .reset_index()
            .sort_values("SEM_NUM")
            .drop(columns="SEM_NUM")
        )
        st.dataframe(calendar_export_df, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — PLAN D'ACTION
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    c1, c2 = st.columns([1, 2])

    with c1:
        recap_s = pivot.sort_values("Taux (%)", ascending=True)
        bar_colors = ["#059669" if v >= 80 else ("#d97706" if v >= 65 else "#dc2626") for v in recap_s["Taux (%)"]]

        fig_h = go.Figure(go.Bar(
            x=recap_s["Taux (%)"],
            y=recap_s["Magasin"],
            orientation="h",
            marker=dict(color=bar_colors),
            text=[f"{v}%" for v in recap_s["Taux (%)"]],
            textposition="outside"
        ))
        fig_h.add_vline(
            x=80,
            line_dash="dash",
            line_color="#cbd5e1",
            annotation_text="Cible 80%",
            annotation_font_color="#64748b"
        )
        fig_h.update_layout(
            **PLOTLY_BASE,
            height=max(300, len(mag_actifs) * 48),
            xaxis=dict(range=[0, 118], gridcolor="#f1f5f9", ticksuffix="%"),
            yaxis=dict(gridcolor="rgba(0,0,0,0)"),
            title="Taux par magasin"
        )
        st.plotly_chart(fig_h, use_container_width=True)

    with c2:
        mag_pa = st.selectbox("Sélectionner un magasin", mag_actifs, key="pa_mag")
        df_pa = detail_df[
            (detail_df["Magasin"] == mag_pa) &
            (detail_df["Statut"].isin(["Alerte Aucun Mouvement", "En Attente Livraison"]))
        ]
        krow = pivot[pivot["Magasin"] == mag_pa]
        t_mag = int(krow["Taux (%)"].values[0]) if not krow.empty else 0
        n_alert = int(krow["Alerte Aucun Mouvement"].values[0]) if not krow.empty else 0
        n_att = int(krow["En Attente Livraison"].values[0]) if not krow.empty else 0

        c_hex = "#059669" if t_mag >= 80 else ("#d97706" if t_mag >= 65 else "#dc2626")
        bg = "#ecfdf5" if t_mag >= 80 else ("#fffbeb" if t_mag >= 65 else "#fef2f2")
        bd = "#a7f3d0" if t_mag >= 80 else ("#fcd34d" if t_mag >= 65 else "#fecaca")

        st.markdown(f"""
        <div style="background:{bg};border:1px solid {bd};border-radius:22px;padding:18px 20px;margin-bottom:14px;display:flex;align-items:center;gap:20px;box-shadow:0 6px 18px rgba(15,23,42,.06);">
          <div style="font-size:54px;font-weight:900;color:{c_hex};line-height:1">{t_mag}%</div>
          <div>
            <div style="font-size:16px;font-weight:900;color:#0f172a">{mag_pa}</div>
            <div style="font-size:12px;color:#64748b;margin-top:4px">{n_alert} alertes aucun mouvement · {n_att} en attente livraison</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        if df_pa.empty:
            st.success(f"{mag_pa} — aucune action requise.")
        else:
            PA_COLS = [
                "SKU", "Libellé article", "Origine", "Mode Appro",
                "Sem. Réception", "Date Livraison", "Code etat",
                "Stock", "RAL", "Statut"
            ]
            st.dataframe(
                df_pa[PA_COLS].sort_values(["Statut", "Origine", "Sem. Réception"]).reset_index(drop=True),
                use_container_width=True,
                hide_index=True
            )

# ══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">Export opérationnel</div>', unsafe_allow_html=True)

export_bytes = build_report_excel(
    detail_df=detail_df.copy(),
    pivot_df=pivot.copy(),
    calendar_df=calendar_export_df.copy() if "calendar_export_df" in locals() else pd.DataFrame(),
    top_alerts_df=top_pb.copy(),
    today_str=TODAY_STR,
    avg_impl=avg_impl,
    sku_im_total=sku_im_total,
    sku_lo_total=sku_lo_total,
    tim=tim,
    tlo=tlo,
    ct=ct,
    ca=ca,
    cal=cal
)

col_dl1, col_dl2 = st.columns([2, 1])

with col_dl1:
    st.download_button(
        label="📥 Télécharger le rapport Excel corporate",
        data=export_bytes,
        file_name=f"rapport_implantation_corporate_{TODAY_FILE}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col_dl2:
    st.caption("Contenu : Synthèse Direction · Résumé Exécutif · Vue magasins · Alertes · Plan d’action · Calendrier · Détail complet")
