import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import date

# ══════════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Rapport Implantation · Carrefour",
    page_icon="🏪",
    layout="wide",
    initial_sidebar_state="expanded"
)

TODAY = date.today()
TODAY_STR = TODAY.strftime("%d %b %Y")
TODAY_FILE = TODAY.strftime("%Y%m%d")

# ══════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM
# ══════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

:root{
  --bg:#f5f7fb;
  --surface:#ffffff;
  --surface-2:#f8fafc;
  --border:#e2e8f0;
  --text:#0f172a;
  --muted:#64748b;

  --green:#059669;
  --green-bg:#ecfdf5;
  --green-bd:#a7f3d0;

  --blue:#2563eb;
  --blue-bg:#eff6ff;
  --blue-bd:#bfdbfe;

  --cyan:#0284c7;
  --cyan-bg:#f0f9ff;
  --cyan-bd:#bae6fd;

  --red:#dc2626;
  --red-bg:#fef2f2;
  --red-bd:#fecaca;

  --amber:#d97706;
  --amber-bg:#fffbeb;
  --amber-bd:#fcd34d;

  --violet:#7c3aed;
  --violet-bg:#f5f3ff;
  --violet-bd:#ddd6fe;

  --shadow-sm:0 1px 2px rgba(15,23,42,.05);
  --shadow-md:0 6px 18px rgba(15,23,42,.08);
  --shadow-lg:0 18px 40px rgba(15,23,42,.10);
  --radius:20px;
}

html, body, [class*="css"]{
  font-family:'Inter', sans-serif !important;
  color:var(--text) !important;
}

.main, section[data-testid="stMain"]{
  background:
    radial-gradient(circle at top left, rgba(59,130,246,.08), transparent 28%),
    radial-gradient(circle at top right, rgba(16,185,129,.08), transparent 24%),
    linear-gradient(180deg,#f8fafc 0%, #f5f7fb 100%) !important;
}

.block-container{
  max-width:1550px !important;
  padding-top:1rem !important;
  padding-bottom:3rem !important;
}

header[data-testid="stHeader"], #MainMenu, footer{
  display:none !important;
}

section[data-testid="stSidebar"]{
  background:#ffffff !important;
  border-right:1px solid var(--border) !important;
}
section[data-testid="stSidebar"] .block-container{
  padding-top:1rem !important;
}

.topbar{
  background:
    linear-gradient(135deg, rgba(15,23,42,.98), rgba(30,41,59,.98)),
    linear-gradient(135deg, #0f172a, #1e293b);
  border:1px solid rgba(255,255,255,.08);
  border-radius:24px;
  padding:22px 26px;
  margin-bottom:22px;
  color:white;
  box-shadow:0 20px 50px rgba(15,23,42,.22);
  position:relative;
  overflow:hidden;
}
.topbar:before{
  content:"";
  position:absolute;
  inset:0;
  background:
    radial-gradient(circle at 15% 20%, rgba(59,130,246,.22), transparent 22%),
    radial-gradient(circle at 90% 20%, rgba(16,185,129,.16), transparent 18%);
  pointer-events:none;
}
.topbar-row{
  position:relative;
  z-index:2;
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap:20px;
  flex-wrap:wrap;
}
.topbar-left{
  display:flex;
  align-items:center;
  gap:16px;
}
.topbar-icon{
  width:56px;
  height:56px;
  border-radius:18px;
  background:linear-gradient(135deg,#3b82f6,#60a5fa);
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:28px;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.22), 0 10px 24px rgba(59,130,246,.22);
}
.topbar-title{
  font-size:26px;
  font-weight:900;
  line-height:1.05;
  color:#ffffff;
  letter-spacing:-0.02em;
}
.topbar-sub{
  font-size:12px;
  color:#cbd5e1;
  margin-top:5px;
  font-weight:500;
}
.topbar-meta{
  display:flex;
  gap:10px;
  flex-wrap:wrap;
  align-items:center;
}
.badge{
  border:1px solid rgba(255,255,255,.14);
  background:rgba(255,255,255,.08);
  color:#e2e8f0;
  border-radius:999px;
  padding:8px 12px;
  font-size:12px;
  font-weight:700;
}
.date-pill{
  color:#93c5fd;
  font-size:12px;
  font-weight:700;
}

.kpi-grid{
  display:grid;
  grid-template-columns:repeat(4,minmax(0,1fr));
  gap:16px;
  margin-bottom:18px;
}
.kpi-card{
  background:rgba(255,255,255,.9);
  backdrop-filter: blur(6px);
  border:1px solid rgba(226,232,240,.9);
  border-radius:22px;
  padding:20px 20px 18px 20px;
  box-shadow:var(--shadow-md);
  position:relative;
  overflow:hidden;
}
.kpi-card:before{
  content:'';
  position:absolute;
  left:0;
  top:0;
  right:0;
  height:5px;
}
.kpi-card.green:before{background:linear-gradient(90deg,#10b981,#059669);}
.kpi-card.blue:before{background:linear-gradient(90deg,#3b82f6,#2563eb);}
.kpi-card.cyan:before{background:linear-gradient(90deg,#06b6d4,#0284c7);}
.kpi-card.red:before{background:linear-gradient(90deg,#ef4444,#dc2626);}
.kpi-label{
  font-size:11px;
  text-transform:uppercase;
  letter-spacing:.12em;
  color:var(--muted);
  font-weight:800;
  margin-bottom:10px;
}
.kpi-value{
  font-size:42px;
  font-weight:900;
  line-height:1;
  letter-spacing:-.03em;
}
.kpi-sub{
  margin-top:8px;
  font-size:12px;
  color:var(--muted);
}

.strip{
  display:grid;
  grid-template-columns:repeat(6,minmax(0,1fr));
  gap:14px;
  margin:16px 0 24px 0;
}
.strip-card{
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:18px;
  padding:16px;
  box-shadow:var(--shadow-sm);
}
.strip-tag{
  display:inline-block;
  padding:5px 9px;
  border-radius:999px;
  font-size:10px;
  font-weight:800;
  letter-spacing:.08em;
  margin-bottom:10px;
}
.tag-im{
  background:var(--blue-bg);
  color:var(--blue);
  border:1px solid var(--blue-bd);
}
.tag-lo{
  background:var(--green-bg);
  color:var(--green);
  border:1px solid var(--green-bd);
}
.strip-label{
  font-size:11px;
  text-transform:uppercase;
  color:var(--muted);
  font-weight:800;
  letter-spacing:.08em;
}
.strip-value{
  font-size:28px;
  font-weight:900;
  line-height:1;
  margin-top:6px;
}
.strip-sub{
  font-size:11px;
  color:var(--muted);
  margin-top:6px;
}

.banner{
  border-radius:20px;
  padding:18px 20px;
  border:1px solid;
  margin-bottom:20px;
  display:flex;
  justify-content:space-between;
  gap:18px;
  flex-wrap:wrap;
  align-items:center;
  box-shadow:var(--shadow-sm);
}
.banner.red{
  background:linear-gradient(180deg,#fff,#fef2f2);
  border-color:var(--red-bd);
}
.banner.blue{
  background:linear-gradient(180deg,#fff,#eff6ff);
  border-color:var(--blue-bd);
}
.banner.amber{
  background:linear-gradient(180deg,#fff,#fffbeb);
  border-color:var(--amber-bd);
}
.banner.green{
  background:linear-gradient(180deg,#fff,#ecfdf5);
  border-color:var(--green-bd);
}
.banner-title{
  font-size:17px;
  font-weight:900;
}
.banner-sub{
  font-size:12px;
  color:var(--muted);
  margin-top:5px;
}
.banner-big{
  font-size:44px;
  font-weight:900;
  line-height:1;
  letter-spacing:-.03em;
}

.section-title{
  margin:26px 0 12px 0;
  font-size:12px;
  font-weight:900;
  text-transform:uppercase;
  letter-spacing:.16em;
  color:var(--muted);
}

.card{
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:20px;
  padding:16px;
  box-shadow:var(--shadow-sm);
}

/* Scorecards magasins */
.rag-grid{
  display:grid;
  grid-template-columns:repeat(auto-fit,minmax(260px,1fr));
  gap:16px;
  margin-bottom:26px;
}
.rag-card{
  position:relative;
  border-radius:22px;
  padding:18px 18px 16px 18px;
  border:1px solid;
  box-shadow:var(--shadow-md);
  min-height:152px;
  overflow:hidden;
  transition:transform .18s ease, box-shadow .18s ease;
}
.rag-card:hover{
  transform:translateY(-2px);
  box-shadow:var(--shadow-lg);
}
.rag-card:before{
  content:"";
  position:absolute;
  left:0;
  top:0;
  right:0;
  height:5px;
}
.rag-card.good{
  background:linear-gradient(180deg,#ffffff 0%, #ecfdf5 100%);
  border-color:var(--green-bd);
}
.rag-card.good:before{
  background:linear-gradient(90deg,#10b981,#059669);
}
.rag-card.mid{
  background:linear-gradient(180deg,#ffffff 0%, #fffbeb 100%);
  border-color:var(--amber-bd);
}
.rag-card.mid:before{
  background:linear-gradient(90deg,#f59e0b,#d97706);
}
.rag-card.bad{
  background:linear-gradient(180deg,#ffffff 0%, #fef2f2 100%);
  border-color:var(--red-bd);
}
.rag-card.bad:before{
  background:linear-gradient(90deg,#ef4444,#dc2626);
}
.rag-top{
  display:flex;
  justify-content:space-between;
  align-items:flex-start;
  gap:10px;
}
.rag-name{
  font-size:13px;
  font-weight:800;
  line-height:1.25;
  color:#0f172a;
  flex:1;
  word-break:break-word;
}
.rag-chip{
  padding:5px 9px;
  border-radius:999px;
  font-size:10px;
  font-weight:900;
  letter-spacing:.08em;
  white-space:nowrap;
}
.rag-chip.good{
  background:#d1fae5;
  color:#065f46;
}
.rag-chip.mid{
  background:#fef3c7;
  color:#92400e;
}
.rag-chip.bad{
  background:#fee2e2;
  color:#991b1b;
}
.rag-main{
  display:flex;
  align-items:flex-end;
  justify-content:space-between;
  margin-top:16px;
  gap:12px;
}
.rag-pct{
  font-size:40px;
  font-weight:900;
  line-height:1;
  letter-spacing:-.03em;
}
.rag-kpis{
  text-align:right;
}
.rag-mini{
  font-size:11px;
  color:var(--muted);
  margin-bottom:4px;
}
.rag-divider{
  height:1px;
  background:rgba(148,163,184,.18);
  margin:14px 0 10px 0;
}
.rag-footer{
  display:flex;
  justify-content:space-between;
  gap:10px;
  align-items:center;
}
.rag-sub{
  font-size:11px;
  color:var(--muted);
}
.rag-alert{
  font-size:11px;
  font-weight:800;
  padding:6px 10px;
  border-radius:999px;
}

.stTabs [data-baseweb="tab-list"]{
  gap:8px;
}
.stTabs [data-baseweb="tab"]{
  background:#fff;
  border:1px solid var(--border);
  border-radius:12px;
  padding:10px 16px;
}
.stTabs [aria-selected="true"]{
  background:linear-gradient(135deg,#0f172a,#1e293b) !important;
  color:#fff !important;
  border-color:#0f172a !important;
}

.stDownloadButton > button{
  width:100% !important;
  border:none !important;
  background:linear-gradient(135deg,#0f172a,#1e293b) !important;
  color:#fff !important;
  font-weight:800 !important;
  border-radius:14px !important;
  padding:12px 16px !important;
  box-shadow:0 10px 24px rgba(15,23,42,.18) !important;
}

div[data-testid="stMetric"]{
  background:#fff;
  border:1px solid var(--border);
  border-radius:16px;
  padding:10px 8px;
  box-shadow:var(--shadow-sm);
}

@media (max-width:1200px){
  .kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr));}
  .strip{grid-template-columns:repeat(2,minmax(0,1fr));}
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════
def fix_encoding(df: pd.DataFrame) -> pd.DataFrame:
    try:
        if any("Ã" in str(c) for c in df.columns):
            df.columns = [
                c.encode("latin1").decode("utf-8", errors="replace")
                for c in df.columns
            ]
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
            df_peek = pd.read_csv(
                buf,
                header=None,
                nrows=1,
                sep=None,
                engine="python",
                encoding="latin1"
            )
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
                df = pd.read_csv(
                    buf,
                    header=None,
                    sep=None,
                    engine="python",
                    encoding="latin1",
                    on_bad_lines="skip"
                )
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
    df = df[df["SKU"].str.match(r"^\d{8}$", na=False)].drop_duplicates(subset="SKU").copy()

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

    df["SEMAINE RECEPTION"] = (
        df["SEMAINE RECEPTION"]
        .astype(str)
        .str.strip()
        .replace("nan", "")
    )

    df["SEM_NUM"] = df["SEMAINE RECEPTION"].apply(
        lambda s: int(str(s).strip("Ss")) if str(s).strip("Ss").isdigit() else 99
    )

    df["ORIGINE"] = df["MODE APPRO"].apply(
        lambda m: "IM" if "IMPORT" in str(m).upper() else "LO"
    )

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


def sem_sort(value) -> int:
    try:
        return int(str(value).strip("Ss"))
    except Exception:
        return 99


def taux_implantation(df: pd.DataFrame) -> int:
    if len(df) == 0:
        return 0
    done = df["Statut"].eq("Implantation Terminée").sum()
    return int(done / len(df) * 100)


def color_taux(t: int) -> str:
    if t >= 80:
        return "#059669"
    if t >= 60:
        return "#d97706"
    return "#dc2626"


def safe_pct(num: int, den: int) -> int:
    return int(num / den * 100) if den else 0


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
        reco.append("Prioriser le flux IMPORT : suivi fournisseurs, transit, dédouanement.")
    if tlo < 70:
        reco.append("Renforcer les relances sur les fournisseurs locaux.")
    if not reco:
        reco.append("Maintenir le rythme d’implantation et surveiller les magasins sous la cible.")

    return status, color, reco


def scorecard_class(taux: int) -> str:
    if taux >= 80:
        return "good"
    if taux >= 60:
        return "mid"
    return "bad"


def scorecard_label(taux: int) -> str:
    if taux >= 80:
        return "ON TRACK"
    if taux >= 60:
        return "WATCH"
    return "AT RISK"


def scorecard_color(taux: int) -> str:
    if taux >= 80:
        return "#059669"
    if taux >= 60:
        return "#d97706"
    return "#dc2626"


def render_store_scorecards(pivot_df: pd.DataFrame) -> str:
    html = '<div class="rag-grid">'
    ordered = pivot_df.sort_values(["Taux (%)", "Alerte Aucun Mouvement"], ascending=[False, True])

    for _, row in ordered.iterrows():
        taux = int(row["Taux (%)"])
        cls = scorecard_class(taux)
        label = scorecard_label(taux)
        color = scorecard_color(taux)

        total = int(row["Total"])
        termines = int(row["Implantation Terminée"])
        attente = int(row["En Attente Livraison"])
        alertes = int(row["Alerte Aucun Mouvement"])

        if cls == "good":
            alert_badge = f'<span class="rag-alert" style="background:#d1fae5;color:#065f46;">{alertes} alertes</span>'
        elif cls == "mid":
            alert_badge = f'<span class="rag-alert" style="background:#fef3c7;color:#92400e;">{alertes} alertes</span>'
        else:
            alert_badge = f'<span class="rag-alert" style="background:#fee2e2;color:#991b1b;">{alertes} alertes</span>'

        html += f"""
        <div class="rag-card {cls}">
          <div class="rag-top">
            <div class="rag-name">{row['Magasin']}</div>
            <div class="rag-chip {cls}">{label}</div>
          </div>

          <div class="rag-main">
            <div class="rag-pct" style="color:{color}">{taux}%</div>
            <div class="rag-kpis">
              <div class="rag-mini"><b>{termines}</b> / {total} terminés</div>
              <div class="rag-mini">{attente} attente livraison</div>
            </div>
          </div>

          <div class="rag-divider"></div>

          <div class="rag-footer">
            <div class="rag-sub">Performance implantation magasin</div>
            {alert_badge}
          </div>
        </div>
        """

    html += "</div>"
    return html


# ══════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════════════
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
        cell.value = f"{title_txt}\n{value_txt}"
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
    ws0.row_dimensions[6].height = 45

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
    ws0["B8"].fill = fill(
        C["amber_bg"] if "tension" in status.lower()
        else (C["red_bg"] if "critique" in status.lower() else C["green_bg"])
    )
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
    write_header(ws1, f"RAPPORT IMPLANTATION — {today_str}", f"Taux moyen réseau : {avg_impl}%")

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
        [("Indicateur", "Indicateur", 32, "l"), ("Valeur", "Valeur", 16, "c")],
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
            ("Libellé article", "Libell
