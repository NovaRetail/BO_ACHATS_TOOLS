import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Perf Hebdo — SmartBuyer",
    page_icon="📊",
    layout="wide",
)

# ─── CHARTE SMARTBUYER (identique aux autres modules) ────────────────────────
COLORS = {
    "Épicerie": "#FF9500",
    "Boissons": "#007AFF",
    "DPH":      "#AF52DE",
}
RAYON_MAP = {
    "00014 - EPICERIE":           "Épicerie",
    "00010 - BOISSONS":           "Boissons",
    "00012 - PARFUMERIE HYGIENE": "DPH",
    "00011 - DROGUERIE":          "DPH",
}

st.markdown("""
<style>
/* ── Fond global ── */
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
.main .block-container          { background: #F2F2F7 !important; }

/* ── Sidebar : light, pas sombre ── */
[data-testid="stSidebar"] > div:first-child { background: #FFFFFF !important; }
[data-testid="stSidebar"] section           { background: #FFFFFF !important; }
[data-testid="stSidebar"] *                 { color: #1C1C1E !important; }
[data-testid="stSidebar"] hr               { border-color: #E5E5EA !important; }
[data-testid="stSidebar"] .stFileUploader  { background: #F9F9FB !important; border-radius: 10px; }

/* ── Typographie globale ── */
html, body, [class*="css"]  { font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", sans-serif; }

/* ── KPI cards ── */
.kpi-card {
    background: #FFFFFF;
    border-radius: 14px;
    border: 0.5px solid #E5E5EA;
    padding: 16px 18px;
    margin-bottom: 0;
}
.kpi-label { font-size: 11px; color: #8E8E93; font-weight: 600;
             text-transform: uppercase; letter-spacing: .5px; margin-bottom: 6px; }
.kpi-val   { font-size: 24px; font-weight: 700; line-height: 1; margin-bottom: 4px; }
.kpi-sub   { font-size: 11px; color: #8E8E93; }

/* ── Rayon cards ── */
.rayon-card {
    background: #FFFFFF;
    border-radius: 14px;
    border: 0.5px solid #E5E5EA;
    padding: 16px 18px;
}
.rayon-name { font-size: 13px; font-weight: 700; margin-bottom: 4px; }
.rayon-ca   { font-size: 20px; font-weight: 700; margin-bottom: 2px; }
.rayon-sub  { font-size: 11px; color: #8E8E93; }

/* ── Section labels ── */
.section-label {
    font-size: 11px; font-weight: 700; color: #8E8E93;
    text-transform: uppercase; letter-spacing: .5px;
    margin: 20px 0 10px;
}

/* ── Alertes ── */
.alert-orange {
    background: #FFF3E0; border-left: 3px solid #FF9500;
    border-radius: 0 8px 8px 0; padding: 10px 14px;
    font-size: 12px; color: #7A4100; margin-bottom: 12px;
}

/* ── Export block ── */
.export-card {
    background: #FFFFFF; border-radius: 14px;
    border: 0.5px solid #E5E5EA; padding: 18px 22px;
    margin-top: 16px;
}
.export-title { font-size: 14px; font-weight: 700; color: #1C1C1E; margin-bottom: 3px; }
.export-sub   { font-size: 12px; color: #8E8E93; }

/* ── Tabs : override Streamlit ── */
[data-testid="stTabs"] [role="tablist"] { background: #E5E5EA; border-radius: 10px; padding: 3px; gap: 0; }
[data-testid="stTabs"] [role="tab"]     { border-radius: 8px; font-size: 12px; font-weight: 500; color: #8E8E93; border: none; background: transparent; }
[data-testid="stTabs"] [aria-selected="true"] { background: #FFFFFF !important; color: #1C1C1E !important; font-weight: 600; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; border: 0.5px solid #E5E5EA; }

/* ── Bouton primaire ── */
.stDownloadButton > button, .stButton > button[kind="primary"] {
    background: #007AFF !important; color: #FFFFFF !important;
    border: none !important; border-radius: 10px !important;
    font-weight: 600 !important; font-size: 13px !important;
}
.stDownloadButton > button:hover { background: #0066DD !important; }

/* ── segmented_control ── */
[data-testid="stSegmentedControl"] { background: #E5E5EA !important; border-radius: 10px; padding: 3px; }
[data-testid="stSegmentedControl"] label { font-size: 12px !important; font-weight: 500 !important; }

/* ── Remove top padding ── */
.block-container { padding-top: 1.5rem !important; }
</style>
""", unsafe_allow_html=True)

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/5/5b/Carrefour_logo.svg/120px-Carrefour_logo.svg.png", width=90)
    st.markdown("### SmartBuyer Hub")
    st.caption("Carrefour Côte d'Ivoire")
    st.divider()

    uploaded = st.file_uploader(
        "📂 Extraction PBI",
        type=["xlsx", "csv"],
        help="Déposez votre extraction PBI hebdomadaire (.xlsx ou .csv)",
    )

    if uploaded:
        st.success(f"✓ {uploaded.name}")
    else:
        st.info("Aucun fichier chargé")

    st.divider()
    st.caption("Module · Perf Hebdo")
    st.caption("SmartBuyer Hub v2.0")

# ─── CHARGEMENT ───────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(file_bytes, file_name):
    if file_name.endswith(".csv"):
        df = pd.read_csv(BytesIO(file_bytes), encoding="latin-1")
    else:
        df = pd.read_excel(BytesIO(file_bytes), engine="openpyxl")
    df.columns = df.columns.str.strip()

    def clean(s):
        if pd.isna(s): return ""
        m = re.match(r"^\d+ - (.+)$", str(s))
        return m.group(1) if m else str(s)

    arts = df[df["Article"].notna() & (df["Article"] != "Total")].copy()
    arts["art_label"]   = arts["Article"].apply(clean)
    arts["rayon_label"] = arts["Rayon"].apply(
        lambda x: RAYON_MAP.get(str(x).strip(), clean(x))
    )

    rayon_tots = df[df["Famille"] == "Total"].copy()
    rayon_tots["rayon_label"] = rayon_tots["Rayon"].apply(
        lambda x: RAYON_MAP.get(str(x).strip(), clean(x))
    )
    rayon_tots = rayon_tots.groupby("rayon_label", as_index=False).agg(
        CA=("CA","sum"), Marge=("Marge","sum"), Casse=("Casse (Valeur)","sum")
    )
    rayon_tots["%Marge"] = rayon_tots["Marge"] / rayon_tots["CA"]
    return arts, rayon_tots

def compute_kpis(arts):
    ca    = arts["CA"].sum()
    marge = arts["Marge"].sum()
    casse = arts["Casse (Valeur)"].sum()
    nb_neg = int((arts["Marge"] < 0).sum())
    nb_casse = int((arts["Casse (Valeur)"] < 0).sum())
    return ca, marge, marge/ca if ca else 0, casse, nb_neg, nb_casse

def top_ca(arts, n=10):
    return arts.nlargest(n,"CA")[["art_label","rayon_label","CA","Marge","%Marge","Qté Vente"]].reset_index(drop=True)
def top_marge(arts, n=10):
    return arts.nlargest(n,"Marge")[["art_label","rayon_label","CA","Marge","%Marge"]].reset_index(drop=True)
def top_promo(arts, n=10):
    return (arts[arts["CA HT Promo"]>0]
            .nlargest(n,"CA HT Promo")[["art_label","rayon_label","CA HT Promo","Marge Promo","%CA Poids Promo"]]
            .reset_index(drop=True))
def flop_marge(arts, n=15):
    return (arts[arts["Marge"]<0]
            .nsmallest(n,"Marge")[["art_label","rayon_label","CA","Marge","%Marge"]]
            .reset_index(drop=True))
def top_casse(arts, n=10):
    return (arts[arts["Casse (Valeur)"].notna()&(arts["Casse (Valeur)"]<0)]
            .nsmallest(n,"Casse (Valeur)")[["art_label","rayon_label","Casse (Valeur)","Casse (Qté)"]]
            .reset_index(drop=True))

def show_df(df, rename_map, num_cols=(), pct_cols=(), neg_cols=(), green_cols=()):
    df = df.rename(columns=rename_map).copy()
    df.index = range(1, len(df)+1)
    style = df.style
    for c in num_cols:
        c2 = rename_map.get(c, c)
        if c2 in df.columns:
            style = style.format({c2: lambda v: f"{v:,.0f}".replace(",", " ")})
    for c in pct_cols:
        c2 = rename_map.get(c, c)
        if c2 in df.columns:
            style = style.format({c2: "{:.1%}"})
    for c in neg_cols:
        c2 = rename_map.get(c, c)
        if c2 in df.columns:
            style = style.map(lambda v: "color:#FF3B30;font-weight:600" if isinstance(v,(int,float)) and v<0 else "", subset=[c2])
    for c in green_cols:
        c2 = rename_map.get(c, c)
        if c2 in df.columns:
            style = style.map(lambda v: "color:#34C759;font-weight:600" if isinstance(v,(int,float)) and v>0 else "", subset=[c2])
    st.dataframe(style, use_container_width=True, height=400)

# ─── EXCEL EXPORT (identique au build_v3) ─────────────────────────────────────
def hdr_fill(h): return PatternFill("solid", fgColor=h)
def mk_border(s="thin",c="D1D1D6"):
    x=Side(style=s,color=c); return Border(left=x,right=x,top=x,bottom=x)
def mk_bottom(c="D1D1D6"): return Border(bottom=Side(style="thin",color=c))
def xfnt(bold=False,color="1C1C1E",size=10): return Font(bold=bold,color=color,size=size,name="Arial")
def xaln(h="left",v="center"): return Alignment(horizontal=h,vertical=v)

def xl_hdr(ws,row,col,title,color,ncols):
    ws.merge_cells(start_row=row,start_column=col,end_row=row,end_column=col+ncols-1)
    c=ws.cell(row=row,column=col,value=f"  {title}")
    c.font=Font(bold=True,color="FFFFFF",size=11,name="Arial")
    c.fill=hdr_fill(color); c.alignment=xaln("left")
    ws.row_dimensions[row].height=24

def xl_table(ws,sr,sc,headers,rows,widths,hc,num_cols=None,pct_cols=None,
             neg_cols=None,green_cols=None,rank=True,rayon_col=None):
    if num_cols  is None: num_cols=[]
    if pct_cols  is None: pct_cols=[]
    if neg_cols  is None: neg_cols=[]
    if green_cols is None: green_cols=[]
    ah=(["#"]+headers) if rank else headers
    aw=([4]+widths)    if rank else widths
    for j,(h,w) in enumerate(zip(ah,aw)):
        col=sc+j
        c=ws.cell(row=sr,column=col,value=h)
        c.font=Font(bold=True,color="FFFFFF",size=10,name="Arial")
        c.fill=hdr_fill(hc); c.alignment=xaln("center"); c.border=mk_border("thin","555555")
        ws.column_dimensions[get_column_letter(col)].width=w
    ws.row_dimensions[sr].height=22
    off=1 if rank else 0
    for i,rd in enumerate(rows):
        r=sr+1+i
        ws.row_dimensions[r].height=19
        bg=hdr_fill("F9F9FB" if i%2 else "FFFFFF")
        if rank:
            c=ws.cell(row=r,column=sc,value=i+1)
            c.font=Font(color="AAAAAA",size=9,name="Arial")
            c.fill=bg; c.alignment=xaln("center"); c.border=mk_bottom()
        for j,val in enumerate(rd):
            col=sc+j+off
            if rayon_col is not None and j==rayon_col:
                label,rcolor=val
                c=ws.cell(row=r,column=col,value=label)
                c.fill=bg; c.border=mk_bottom()
                c.font=Font(bold=True,color=rcolor,size=10,name="Arial")
                c.alignment=xaln("center"); continue
            c=ws.cell(row=r,column=col,value=val)
            c.fill=bg; c.border=mk_bottom()
            if j in num_cols:
                c.number_format='#,##0'; c.alignment=xaln("right"); c.font=xfnt()
            elif j in pct_cols:
                c.number_format='0.0%'; c.alignment=xaln("center"); c.font=xfnt()
            else:
                c.alignment=xaln("left"); c.font=xfnt()
            if j in neg_cols   and isinstance(val,(int,float)) and val<0:
                c.font=Font(bold=True,color="FF3B30",size=10,name="Arial")
            if j in green_cols and isinstance(val,(int,float)) and val>0:
                c.font=Font(bold=True,color="34C759",size=10,name="Arial")
    return sr+1+len(rows)

def sp(ws,row,h=12): ws.row_dimensions[row].height=h; return row+1

def generate_excel(arts, rayon_tots):
    wb = Workbook()

    def rc_rows(df, art_col, rayon_col, *extra):
        return [tuple([r[art_col], (r[rayon_col], COLORS.get(r[rayon_col],"555555"))] +
                      [r[c] for c in extra]) for _,r in df.iterrows()]

    def simple_rows(df, cols):
        return [tuple(r[c] for c in cols) for _,r in df.iterrows()]

    tca   = top_ca(arts)
    tmg   = top_marge(arts)
    tpr   = top_promo(arts)
    tfl   = flop_marge(arts)
    tcs   = top_casse(arts)
    ca_t,mg_t,pct_t,cs_t,nb_neg,nb_cs = compute_kpis(arts)
    nb_art = len(arts)
    RCOL_W = 16

    # ── SYNTHÈSE RÉSEAU ──────────────────────────────────────────────────────
    ws0 = wb.active; ws0.title = "📊 Synthèse Réseau"
    ws0.sheet_view.showGridLines = False
    ws0.row_dimensions[1].height = 8

    ws0.merge_cells("A2:J2")
    ws0["A2"] = "PERFORMANCE COMMERCIALE HEBDOMADAIRE — RÉSEAU CARREFOUR CÔTE D'IVOIRE"
    ws0["A2"].font = Font(bold=True,color="FFFFFF",size=15,name="Arial")
    ws0["A2"].fill = hdr_fill("1C1C1E"); ws0["A2"].alignment = xaln("center")
    ws0.row_dimensions[2].height = 36

    ws0.merge_cells("A3:J3")
    ws0["A3"] = f"Extraction PBI · {nb_art:,} articles actifs · Semaine en cours".replace(",","")
    ws0["A3"].font = Font(color="8E8E93",size=10,name="Arial")
    ws0["A3"].fill = hdr_fill("F2F2F7"); ws0["A3"].alignment = xaln("center")
    ws0.row_dimensions[3].height = 20; ws0.row_dimensions[4].height = 10

    kpis_net = [
        ("CA TOTAL RÉSEAU", f"{ca_t/1e6:.1f} M FCFA", "",                  "007AFF", 1),
        ("MARGE BRUTE",     f"{mg_t/1e6:.1f} M FCFA", f"{pct_t:.1%} du CA","34C759", 3),
        ("CASSE RÉSEAU",    f"{cs_t/1e6:.2f} M FCFA", f"{nb_cs} articles", "FF3B30", 5),
        ("MARGES NÉGATIVES",f"{nb_neg} articles",      "à corriger",        "FF9500", 7),
    ]
    for label,val,sub,color,col in kpis_net:
        ec=col+1
        for r in range(5,10): ws0.merge_cells(start_row=r,start_column=col,end_row=r,end_column=ec)
        ws0.cell(row=5,column=col).fill=hdr_fill(color); ws0.row_dimensions[5].height=5
        c=ws0.cell(row=6,column=col,value=label)
        c.font=Font(bold=True,color="FFFFFF",size=9,name="Arial")
        c.fill=hdr_fill(color); c.alignment=xaln("center"); ws0.row_dimensions[6].height=16
        c=ws0.cell(row=7,column=col,value=val)
        c.font=Font(bold=True,color="FFFFFF",size=13,name="Arial")
        c.fill=hdr_fill(color); c.alignment=xaln("center"); ws0.row_dimensions[7].height=26
        c=ws0.cell(row=8,column=col,value=sub)
        c.font=Font(color="CCCCCC",size=9,name="Arial")
        c.fill=hdr_fill(color); c.alignment=xaln("center"); ws0.row_dimensions[8].height=16
        ws0.cell(row=9,column=col).fill=hdr_fill(color); ws0.row_dimensions[9].height=5

    ws0.row_dimensions[10].height = 14

    # récap rayons
    xl_hdr(ws0,11,1,"RÉCAPITULATIF PAR RAYON","3A3A3C",7)
    rh=["RAYON","CA (FCFA)","MARGE (FCFA)","% MARGE","CASSE (FCFA)","ART. ACTIFS"]
    rw=[26,18,18,10,16,12]
    for j,(h,w) in enumerate(zip(rh,rw)):
        c=ws0.cell(row=12,column=1+j,value=h)
        c.font=Font(bold=True,color="FFFFFF",size=10,name="Arial")
        c.fill=hdr_fill("3A3A3C"); c.alignment=xaln("center"); c.border=mk_border("thin","555555")
        ws0.column_dimensions[get_column_letter(1+j)].width=w
    ws0.row_dimensions[12].height=22
    nb_arts_r=arts.groupby("rayon_label").size().to_dict()
    for i,(_,row) in enumerate(rayon_tots.iterrows()):
        r=13+i; ws0.row_dimensions[r].height=20
        bg=hdr_fill("F9F9FB" if i%2 else "FFFFFF")
        vals=[row["rayon_label"],row["CA"],row["Marge"],row["%Marge"],row["Casse"],nb_arts_r.get(row["rayon_label"],0)]
        fmts=[None,'#,##0','#,##0','0.0%','#,##0','#,##0']
        for j,(v,fmt) in enumerate(zip(vals,fmts)):
            c=ws0.cell(row=r,column=1+j,value=v)
            c.fill=bg; c.border=mk_bottom()
            if fmt: c.number_format=fmt
            c.alignment=xaln("right" if isinstance(v,(int,float)) else "left")
            c.font=xfnt(bold=(j==0),color=COLORS.get(str(v),"1C1C1E") if j==0 else "1C1C1E")
            if j==4 and isinstance(v,(int,float)) and v<0:
                c.font=Font(bold=True,color="FF3B30",size=10,name="Arial")
    r_tot=13+len(rayon_tots); ws0.row_dimensions[r_tot].height=22
    for j,(v,fmt,clr) in enumerate([
        ("TOTAL",None,"1C1C1E"),(ca_t,'#,##0',"007AFF"),
        (mg_t,'#,##0',"34C759"),(pct_t,'0.0%',"1C1C1E"),
        (cs_t,'#,##0',"FF3B30"),(nb_art,'#,##0',"1C1C1E")
    ]):
        c=ws0.cell(row=r_tot,column=1+j,value=v)
        c.fill=hdr_fill("F2F2F7"); c.border=mk_border("medium","AAAAAA")
        if fmt: c.number_format=fmt
        c.font=Font(bold=True,color=clr,size=10,name="Arial")
        c.alignment=xaln("right" if j>0 else "left")

    cur=r_tot+1; cur=sp(ws0,cur,18)

    # classements tous rayons
    xl_hdr(ws0,cur,1,"TOP 10 CA — TOUS RAYONS CONFONDUS","1C1C1E",8); cur+=1
    cur=xl_table(ws0,cur,1,["ARTICLE","RAYON","CA (FCFA)","MARGE (FCFA)","% MARGE","QTÉ VENDUE"],
        rc_rows(tca,"art_label","rayon_label","CA","Marge","%Marge","Qté Vente"),
        [40,RCOL_W,16,16,10,12],"1C1C1E",num_cols=[2,3,5],pct_cols=[4],green_cols=[3],neg_cols=[3],rayon_col=1)
    cur=sp(ws0,cur,14)

    xl_hdr(ws0,cur,1,"TOP 10 MARGE BRUTE — TOUS RAYONS CONFONDUS","34C759",7); cur+=1
    cur=xl_table(ws0,cur,1,["ARTICLE","RAYON","CA (FCFA)","MARGE (FCFA)","% MARGE"],
        rc_rows(tmg,"art_label","rayon_label","CA","Marge","%Marge"),
        [40,RCOL_W,16,16,10],"34C759",num_cols=[2,3],pct_cols=[4],green_cols=[3],rayon_col=1)
    cur=sp(ws0,cur,14)

    xl_hdr(ws0,cur,1,"TOP 10 VENTES PROMO — TOUS RAYONS CONFONDUS","FF9500",7); cur+=1
    cur=xl_table(ws0,cur,1,["ARTICLE","RAYON","CA PROMO (FCFA)","MARGE PROMO (FCFA)","POIDS PROMO"],
        rc_rows(tpr,"art_label","rayon_label","CA HT Promo","Marge Promo","%CA Poids Promo"),
        [40,RCOL_W,18,18,12],"FF9500",num_cols=[2,3],pct_cols=[4],green_cols=[3],rayon_col=1)
    cur=sp(ws0,cur,14)

    xl_hdr(ws0,cur,1,"FLOP 15 MARGES NÉGATIVES — TOUS RAYONS CONFONDUS","FF3B30",7); cur+=1
    cur=xl_table(ws0,cur,1,["ARTICLE","RAYON","CA (FCFA)","MARGE (FCFA)","% MARGE"],
        rc_rows(tfl,"art_label","rayon_label","CA","Marge","%Marge"),
        [40,RCOL_W,14,14,12],"FF3B30",num_cols=[2,3],pct_cols=[4],neg_cols=[3,4],rayon_col=1)
    cur=sp(ws0,cur,14)

    xl_hdr(ws0,cur,1,"TOP 10 CASSE — TOUS RAYONS CONFONDUS","555555",6); cur+=1
    xl_table(ws0,cur,1,["ARTICLE","RAYON","CASSE VALEUR (FCFA)","CASSE QTÉ"],
        rc_rows(tcs,"art_label","rayon_label","Casse (Valeur)","Casse (Qté)"),
        [40,RCOL_W,20,12],"555555",num_cols=[2,3],neg_cols=[2,3],rayon_col=1)
    ws0.freeze_panes="A4"

    # ── ONGLETS PAR RAYON ────────────────────────────────────────────────────
    for rayon,color in COLORS.items():
        arts_r=arts[arts["rayon_label"]==rayon]
        if arts_r.empty: continue
        ca_r=arts_r["CA"].sum(); mg_r=arts_r["Marge"].sum()
        pct_r=mg_r/ca_r if ca_r else 0
        cs_r=arts_r["Casse (Valeur)"].sum()
        neg_r=int((arts_r["Marge"]<0).sum()); cas_r=int((arts_r["Casse (Valeur)"]<0).sum())
        tca_r=top_ca(arts_r); tmg_r=top_marge(arts_r)
        tpr_r=top_promo(arts_r); tfl_r=flop_marge(arts_r); tcs_r=top_casse(arts_r)

        ws=wb.create_sheet(f"📋 {rayon}")
        ws.sheet_view.showGridLines=False; ws.row_dimensions[1].height=8

        ws.merge_cells("A2:G2")
        ws["A2"]=f"CLASSEMENTS SEMAINE — {rayon.upper()}"
        ws["A2"].font=Font(bold=True,color="FFFFFF",size=14,name="Arial")
        ws["A2"].fill=hdr_fill(color.replace("#","")); ws["A2"].alignment=xaln("center")
        ws.row_dimensions[2].height=32

        ws.merge_cells("A3:G3")
        ws["A3"]=(f"CA : {ca_r:,.0f} FCFA  |  Marge : {mg_r:,.0f} FCFA ({pct_r:.1%})"
                  f"  |  Casse : {cs_r:,.0f} FCFA  |  Marges neg. : {neg_r} art.").replace(",","")
        ws["A3"].font=Font(color=color.replace("#",""),size=10,name="Arial",bold=True)
        ws["A3"].fill=hdr_fill("F9F9FB"); ws["A3"].alignment=xaln("center")
        ws.row_dimensions[3].height=22; cur=5

        def sr(df,cols): return [tuple(r[c] for c in cols) for _,r in df.iterrows()]

        xl_hdr(ws,cur,1,f"TOP 10 CA — {rayon.upper()}",color.replace("#",""),7); cur+=1
        cur=xl_table(ws,cur,1,["ARTICLE","CA (FCFA)","MARGE (FCFA)","% MARGE","QTÉ VENDUE"],
            sr(tca_r,["art_label","CA","Marge","%Marge","Qté Vente"]),
            [42,16,16,10,12],color.replace("#",""),num_cols=[1,2,4],pct_cols=[3],green_cols=[2],neg_cols=[2,3])
        cur=sp(ws,cur,14)

        xl_hdr(ws,cur,1,f"TOP 10 MARGE — {rayon.upper()}","34C759",6); cur+=1
        cur=xl_table(ws,cur,1,["ARTICLE","CA (FCFA)","MARGE (FCFA)","% MARGE"],
            sr(tmg_r,["art_label","CA","Marge","%Marge"]),
            [44,16,16,10],"34C759",num_cols=[1,2],pct_cols=[3],green_cols=[2])
        cur=sp(ws,cur,14)

        xl_hdr(ws,cur,1,f"TOP PROMO — {rayon.upper()}","FF9500",6); cur+=1
        if not tpr_r.empty:
            cur=xl_table(ws,cur,1,["ARTICLE","CA PROMO","MARGE PROMO","POIDS PROMO"],
                sr(tpr_r,["art_label","CA HT Promo","Marge Promo","%CA Poids Promo"]),
                [44,18,18,12],"FF9500",num_cols=[1,2],pct_cols=[3],green_cols=[2])
        cur=sp(ws,cur,14)

        xl_hdr(ws,cur,1,f"FLOP MARGES NÉGATIVES — {rayon.upper()}","FF3B30",6); cur+=1
        ws.merge_cells(start_row=cur,start_column=1,end_row=cur,end_column=6)
        note=ws.cell(row=cur,column=1,value=f"⚠  {neg_r} articles en perte — vérifier PA, promos, conditions fournisseurs.")
        note.font=Font(italic=True,color="FF3B30",size=9,name="Arial")
        note.fill=hdr_fill("FFF5F5"); note.alignment=xaln("left"); ws.row_dimensions[cur].height=17; cur+=1
        cur=xl_table(ws,cur,1,["ARTICLE","CA (FCFA)","MARGE (FCFA)","% MARGE"],
            sr(tfl_r,["art_label","CA","Marge","%Marge"]),
            [44,14,14,12],"FF3B30",num_cols=[1,2],pct_cols=[3],neg_cols=[2,3])
        cur=sp(ws,cur,14)

        xl_hdr(ws,cur,1,f"TOP CASSE — {rayon.upper()}","555555",5); cur+=1
        ws.merge_cells(start_row=cur,start_column=1,end_row=cur,end_column=5)
        note2=ws.cell(row=cur,column=1,value=f"Casse totale : {cs_r:,.0f} FCFA  ·  {cas_r} articles".replace(",",""))
        note2.font=Font(italic=True,color="555555",size=9,name="Arial")
        note2.fill=hdr_fill("F2F2F7"); note2.alignment=xaln("left"); ws.row_dimensions[cur].height=17; cur+=1
        if not tcs_r.empty:
            xl_table(ws,cur,1,["ARTICLE","CASSE VALEUR (FCFA)","CASSE QTÉ"],
                sr(tcs_r,["art_label","Casse (Valeur)","Casse (Qté)"]),
                [46,20,12],"555555",num_cols=[1,2],neg_cols=[1,2])
        ws.freeze_panes="A4"

    wb.active=wb["📊 Synthèse Réseau"]
    buf=BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ─── RENDU ────────────────────────────────────────────────────────────────────
st.markdown("## 📊 Performance commerciale hebdo")

st.markdown("""
Ce module génère le **rapport de performance commerciale hebdomadaire** à partir de votre extraction PBI.

Il calcule automatiquement, pour l'ensemble du réseau et par rayon, les indicateurs clés de la semaine :
**Top CA**, **Top masse de marge**, **Meilleures ventes en promotion**, **Flop marges négatives** et **Top casse**.

Le bouton **Exporter Excel** produit un fichier structuré en 4 onglets — Synthèse réseau + un onglet par rayon (Épicerie, Boissons, DPH) — prêt à être partagé avec l'équipe d'acheteurs.
""")

st.divider()

if not uploaded:
    st.info("⬆ Chargez votre extraction PBI dans la sidebar pour démarrer.")
    st.stop()

with st.spinner("Chargement de l'extraction..."):
    arts, rayon_tots = load_data(uploaded.getvalue(), uploaded.name)

ca_tot, marge_tot, pct_tot, casse_tot, nb_neg, nb_cs = compute_kpis(arts)
st.caption(f"Fichier : **{uploaded.name}** · {len(arts):,} articles actifs".replace(",",""))

# KPIs
c1,c2,c3,c4 = st.columns(4)
c1.markdown(f'<div class="kpi-card"><div class="kpi-label">CA Total Réseau</div><div class="kpi-val" style="color:#007AFF">{ca_tot/1e6:.1f} M</div><div class="kpi-sub">FCFA</div></div>', unsafe_allow_html=True)
c2.markdown(f'<div class="kpi-card"><div class="kpi-label">Marge Brute</div><div class="kpi-val" style="color:#34C759">{marge_tot/1e6:.1f} M</div><div class="kpi-sub">{pct_tot:.1%} du CA</div></div>', unsafe_allow_html=True)
c3.markdown(f'<div class="kpi-card"><div class="kpi-label">Casse</div><div class="kpi-val" style="color:#FF3B30">{casse_tot/1e6:.2f} M</div><div class="kpi-sub">{nb_cs} articles touchés</div></div>', unsafe_allow_html=True)
c4.markdown(f'<div class="kpi-card"><div class="kpi-label">Marges Négatives</div><div class="kpi-val" style="color:#FF9500">{nb_neg}</div><div class="kpi-sub">articles en perte</div></div>', unsafe_allow_html=True)

st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

# Rayon cards
r1,r2,r3 = st.columns(3)
ca_max = rayon_tots["CA"].max()
for col_ui, rayon in zip([r1,r2,r3], ["Épicerie","Boissons","DPH"]):
    row = rayon_tots[rayon_tots["rayon_label"]==rayon]
    if row.empty: continue
    rv = row.iloc[0]
    color = COLORS[rayon]
    pct_bar = int(rv["CA"]/ca_max*100)
    col_ui.markdown(f"""<div class="rayon-card">
      <div class="rayon-name" style="color:{color}">{rayon}</div>
      <div class="rayon-ca" style="color:{color}">{rv['CA']/1e6:.1f} M FCFA</div>
      <div class="rayon-sub">Marge {rv['%Marge']:.1%} &nbsp;·&nbsp; Casse {rv['Casse']/1e3:.0f} K FCFA</div>
      <div style="height:4px;background:#E5E5EA;border-radius:2px;margin-top:10px;overflow:hidden">
        <div style="height:4px;width:{pct_bar}%;background:{color};border-radius:2px"></div>
      </div>
    </div>""", unsafe_allow_html=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# Alerte
critiques = int((arts["Marge"] < -100_000).sum())
if critiques > 0:
    st.markdown(f'<div class="alert-orange">⚠ &nbsp;<strong>{critiques} article{"s" if critiques>1 else ""}</strong> avec marge &lt; -100 000 FCFA — action corrective recommandée avant la semaine prochaine.</div>', unsafe_allow_html=True)

# Filtre rayon
st.markdown('<div class="section-label">Classements</div>', unsafe_allow_html=True)
rayon_filtre = st.segmented_control(
    "Rayon", ["Tous rayons","Épicerie","Boissons","DPH"],
    default="Tous rayons", label_visibility="collapsed"
)
arts_f = arts if rayon_filtre=="Tous rayons" else arts[arts["rayon_label"]==rayon_filtre]

# Onglets
RENAME = {
    "art_label":"Article","rayon_label":"Rayon","CA":"CA (FCFA)","Marge":"Marge (FCFA)",
    "%Marge":"% Marge","Qté Vente":"Qté vendue","CA HT Promo":"CA Promo (FCFA)",
    "Marge Promo":"Marge Promo (FCFA)","%CA Poids Promo":"Poids Promo",
    "Casse (Valeur)":"Casse valeur (FCFA)","Casse (Qté)":"Casse qté",
}

tab1,tab2,tab3,tab4,tab5 = st.tabs(["🏆 Top CA","💚 Top Marge","🎯 Top Promo","🔴 Flop Marges","🗑️ Casse"])

with tab1:
    show_df(top_ca(arts_f), RENAME, num_cols=["CA","Marge","Qté Vente"], pct_cols=["%Marge"],
            neg_cols=["Marge"], green_cols=["Marge"])
with tab2:
    show_df(top_marge(arts_f), RENAME, num_cols=["CA","Marge"], pct_cols=["%Marge"], green_cols=["Marge"])
with tab3:
    df_pr=top_promo(arts_f)
    if df_pr.empty: st.info("Aucun article promotionnel sur ce périmètre.")
    else: show_df(df_pr, RENAME, num_cols=["CA HT Promo","Marge Promo"], pct_cols=["%CA Poids Promo"], green_cols=["Marge Promo"])
with tab4:
    df_fl=flop_marge(arts_f)
    st.warning(f"{len(df_fl)} articles en marge négative sur ce périmètre.")
    show_df(df_fl, RENAME, num_cols=["CA","Marge"], pct_cols=["%Marge"], neg_cols=["Marge","%Marge"])
with tab5:
    df_cs=top_casse(arts_f)
    if df_cs.empty: st.info("Aucune casse enregistrée sur ce périmètre.")
    else: show_df(df_cs, RENAME, num_cols=["Casse (Valeur)","Casse (Qté)"], neg_cols=["Casse (Valeur)","Casse (Qté)"])

# Export
st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
st.divider()
ex1,ex2 = st.columns([3,1])
with ex1:
    st.markdown('<div class="export-card"><div class="export-title">Export Excel — Rapport complet</div><div class="export-sub">1 fichier · 4 onglets · Synthèse réseau + classements tous rayons + 1 onglet par rayon</div></div>', unsafe_allow_html=True)
with ex2:
    buf = generate_excel(arts, rayon_tots)
    st.download_button(
        label="⬇ Télécharger le rapport",
        data=buf,
        file_name="Perf_Hebdo_SmartBuyer.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
