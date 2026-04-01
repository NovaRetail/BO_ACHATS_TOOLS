import streamlit as st
import pandas as pd
import numpy as np
import io
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config (doit être le PREMIER appel Streamlit) ───────────────────────
st.set_page_config(page_title="Ruptures de Stock", page_icon="📦", layout="wide")

# ── Charte SmartBuyer ─────────────────────────────────────────────────────────
PL_BG   = "rgba(0,0,0,0)"
PL_FONT = "#1C1C1E"
PL_GRID = "#E5E5EA"
RED     = "#FF3B30"
GREEN   = "#34C759"
ORANGE  = "#FF9500"
BLUE    = "#007AFF"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background:#F2F2F7; }
[data-testid="stHeader"]           { background:transparent; }
section[data-testid="stSidebar"]   { background:#FFFFFF; border-right:1px solid #E5E5EA; }
.block-container { padding:1.5rem 2rem 2rem; max-width:1400px; }

.kpi-card { background:#FFFFFF; border-radius:12px; padding:18px 20px;
            border:1px solid #E5E5EA; }
.kpi-label { font-size:11px; font-weight:600; text-transform:uppercase;
             letter-spacing:.06em; color:#8E8E93; margin-bottom:4px; }
.kpi-value { font-size:28px; font-weight:600; line-height:1; margin-bottom:4px; }
.kpi-sub   { font-size:12px; color:#8E8E93; }
.kpi-danger  .kpi-value { color:#C0392B; }
.kpi-success .kpi-value { color:#196F3D; }
.kpi-info    .kpi-value { color:#007AFF; }
.kpi-warn    .kpi-value { color:#B7850A; }

.section-header { font-size:13px; font-weight:600; color:#1C1C1E;
                  text-transform:uppercase; letter-spacing:.05em;
                  margin:1.2rem 0 .6rem 0; }
.tag-cmd  { background:#FDECEA; color:#C0392B; padding:3px 10px;
            border-radius:20px; font-size:11px; font-weight:600; }
.tag-cess { background:#E8F8F0; color:#196F3D; padding:3px 10px;
            border-radius:20px; font-size:11px; font-weight:600; }
.badge-critique { background:#FDECEA; color:#C0392B; padding:2px 8px;
                  border-radius:10px; font-size:11px; font-weight:600; }
.badge-eleve    { background:#FEF3E2; color:#B7850A; padding:2px 8px;
                  border-radius:10px; font-size:11px; font-weight:600; }
.badge-ok       { background:#E8F8F0; color:#196F3D; padding:2px 8px;
                  border-radius:10px; font-size:11px; font-weight:600; }
.card-wrap { background:#FFFFFF; border-radius:12px; border:1px solid #E5E5EA;
             padding:14px 16px; }
</style>
""", unsafe_allow_html=True)

# ── Constantes ────────────────────────────────────────────────────────────────
MAG_COLS = [
    'Palmeraie', 'Yopougon', 'Kokoh Mall', 'Golf', 'Riviera',
    'II Plateaux', 'Marcory', 'Cité verte', '7 Décembre'
]

def is_rupture(val):
    return pd.isna(val) or val <= 0

# ── Chargement ────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None,
                       engine='openpyxl')
    data = df.iloc[2:].copy()
    data.columns = [
        'Rayon', 'Famille', 'Sous_Famille', 'Article',
        'Palmeraie', 'Yopougon', 'Kokoh Mall', 'Golf', 'Riviera',
        'II Plateaux', 'Marcory', 'Cité verte', '7 Décembre', 'Total'
    ]
    arts = data[
        data['Article'].notna() &
        data['Article'].astype(str).str.contains(' - ', regex=False)
    ].copy().reset_index(drop=True)

    for col in MAG_COLS:
        arts[col] = pd.to_numeric(arts[col], errors='coerce')
    for col in ['Rayon', 'Famille', 'Sous_Famille']:
        arts[col] = arts[col].astype(str).apply(
            lambda x: x.split(' - ', 1)[-1].strip() if ' - ' in x else x
        )
    arts['Code_Article'] = arts['Article'].astype(str).str.split(' - ').str[0].str.strip()
    arts['Libelle']      = arts['Article'].astype(str).apply(
        lambda x: ' - '.join(x.split(' - ')[1:]).strip()
    )
    return arts

@st.cache_data(show_spinner=False)
def compute_plan(file_bytes: bytes, seuil: int) -> pd.DataFrame:
    arts = load_data(file_bytes)
    rows = []
    for _, r in arts.iterrows():
        rups   = [m for m in MAG_COLS if is_rupture(r[m])]
        dispos = {m: int(r[m]) for m in MAG_COLS
                  if not is_rupture(r[m]) and r[m] > seuil}
        if not rups:
            continue
        if dispos:
            max_mag   = max(dispos, key=dispos.get)
            action    = 'VOIR CESSION'
            site_ref  = max_mag
            stock_ref = dispos[max_mag]
        else:
            action    = 'COMMANDER'
            site_ref  = '—'
            stock_ref = 0
        rows.append({
            'Rayon':            r['Rayon'],
            'Famille':          r['Famille'],
            'Sous_Famille':     r['Sous_Famille'],
            'Code_Article':     r['Code_Article'],
            'Libellé':          r['Libelle'],
            'Nb_Ruptures':      len(rups),
            'Magasins_Rupture': ', '.join(rups),
            'Action':           action,
            'Site_Donneur':     site_ref,
            'Stock_Disponible': stock_ref if action == 'VOIR CESSION' else None,
        })
    return pd.DataFrame(rows).sort_values(
        ['Nb_Ruptures', 'Action'], ascending=[False, True]
    ).reset_index(drop=True)

@st.cache_data(show_spinner=False)
def to_excel_bytes(df: pd.DataFrame, seuil: int) -> bytes:
    def bd():
        s = Side(style='thin', color='CCCCCC')
        return Border(left=s, right=s, top=s, bottom=s)
    def fill(h):
        return PatternFill('solid', fgColor=h)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Plan Action'

    ws.merge_cells('A1:I1')
    ws['A1'] = f"PLAN D'ACTION RUPTURES — Seuil cession : {seuil} unités"
    ws['A1'].font      = Font(name='Calibri', bold=True, size=13, color='FFFFFF')
    ws['A1'].fill      = fill('1B3A6B')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    hdrs   = ['Rayon','Famille','Code Article','Libellé','Nb Rupt.',
              'Magasins en Rupture','Action','Site Donneur','Stock Dispo']
    widths = [16, 24, 12, 42, 9, 55, 16, 18, 11]
    for j, (h, w) in enumerate(zip(hdrs, widths), 1):
        c = ws.cell(row=3, column=j, value=h)
        c.font      = Font(name='Calibri', bold=True, size=9, color='FFFFFF')
        c.fill      = fill('1B3A6B')
        c.alignment = Alignment(horizontal='center', vertical='center',
                                wrap_text=True)
        c.border    = bd()
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[3].height = 22

    for i, row in df.iterrows():
        r      = i + 4
        is_cmd = row['Action'] == 'COMMANDER'
        bg     = ('FEF9F9' if is_cmd else 'F9FEF9') if i % 2 == 0 \
                 else ('FEF3F3' if is_cmd else 'F3FEF5')
        stock_val = (int(row['Stock_Disponible'])
                     if pd.notna(row['Stock_Disponible']) and row['Stock_Disponible']
                     else '')
        vals = [row['Rayon'], row['Famille'], row['Code_Article'], row['Libellé'],
                row['Nb_Ruptures'], row['Magasins_Rupture'],
                '🔴 COMMANDER' if is_cmd else '🟢 VOIR CESSION',
                row['Site_Donneur'], stock_val]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = Font(
                name='Calibri', size=9, bold=(j == 7),
                color=('C0392B' if (j == 7 and is_cmd) else
                       '196F3D' if (j in [7, 8, 9] and not is_cmd) else '333333')
            )
            c.fill = fill(
                'FDECEA' if (j == 7 and is_cmd) else
                'EAF7EC' if (j == 7 and not is_cmd) else bg
            )
            c.alignment = Alignment(
                horizontal='center' if j in [3, 5, 7, 8, 9] else 'left',
                vertical='center', wrap_text=(j == 6)
            )
            c.border = bd()
        ws.row_dimensions[r].height = 15

    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = f'A3:I{len(df) + 3}'

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Paramètres")
    uploaded = st.file_uploader(
        "Export stock (xlsx)", type=["xlsx"],
        help="Export ERP avec une colonne par magasin"
    )
    st.markdown("---")
    seuil = st.slider(
        "Seuil minimum cession (unités)",
        min_value=0, max_value=500, value=100, step=10,
        help="Le site donneur doit avoir PLUS que ce stock pour qu'une cession soit proposée"
    )
    st.markdown(
        f'<div style="background:#E8F0FE;color:#1A56DB;border-radius:8px;'
        f'padding:6px 14px;font-size:12px;font-weight:500;margin-top:6px">'
        f'Cession possible si stock donneur &gt; <b>{seuil}</b> unités</div>',
        unsafe_allow_html=True
    )
    st.markdown("---")
    st.markdown("**SmartBuyer** · Module Ruptures")
    st.markdown('<span style="font-size:11px;color:#8E8E93">v1.0 — DPH</span>',
                unsafe_allow_html=True)

# ── Guard upload ──────────────────────────────────────────────────────────────
st.markdown("## 📦 Ruptures de stock")

if uploaded is None:
    st.info("👈 Chargez votre export stock dans le panneau gauche pour démarrer.")
    st.stop()

# ── Calculs ───────────────────────────────────────────────────────────────────
with st.spinner("Analyse en cours…"):
    file_bytes = uploaded.read()
    arts       = load_data(file_bytes)
    plan_df    = compute_plan(file_bytes, seuil)

nb_total = len(arts)
nb_rupt  = len(plan_df)
nb_cmd   = int((plan_df['Action'] == 'COMMANDER').sum())
nb_cess  = int((plan_df['Action'] == 'VOIR CESSION').sum())

# ── KPIs ──────────────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)

def kpi(col, label, value, sub, css_cls):
    with col:
        st.markdown(
            f'<div class="kpi-card {css_cls}">'
            f'<div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value:,}</div>'
            f'<div class="kpi-sub">{sub}</div>'
            f'</div>',
            unsafe_allow_html=True
        )

kpi(c1, "Références analysées", nb_total, f"sur {len(MAG_COLS)} magasins",   "kpi-info")
kpi(c2, "Refs en rupture",       nb_rupt,  "≥ 1 magasin concerné",            "kpi-danger")
kpi(c3, "À commander",           nb_cmd,   "aucun stock réseau",               "kpi-danger")
kpi(c4, "Cessions possibles",    nb_cess,  f"stock donneur > {seuil} unités", "kpi-success")

st.markdown("<div style='margin-top:1rem'></div>", unsafe_allow_html=True)

# ── Graphiques côte à côte ────────────────────────────────────────────────────
col_left, col_right = st.columns(2)

# Magasins — barres horizontales Plotly
with col_left:
    st.markdown('<div class="section-header">Ruptures par magasin</div>',
                unsafe_allow_html=True)
    mag_rows = []
    for mag in MAG_COLS:
        nb  = int(sum(1 for v in arts[mag] if is_rupture(v)))
        pct = nb / nb_total * 100
        mag_rows.append({'Magasin': mag, 'Ruptures': nb, 'Pct': pct})
    mag_df = (pd.DataFrame(mag_rows)
                .sort_values('Ruptures', ascending=True))

    colors = ['#E24B4A' if p >= 40 else ('#EF9F27' if p >= 25 else '#63B967')
              for p in mag_df['Pct']]
    fig_mag = go.Figure(go.Bar(
        x=mag_df['Ruptures'], y=mag_df['Magasin'],
        orientation='h',
        text=[f"{p:.0f}%" for p in mag_df['Pct']],
        textposition='outside',
        marker_color=colors,
        hovertemplate='%{y}<br>%{x} ruptures<extra></extra>',
    ))
    fig_mag.update_layout(
        paper_bgcolor=PL_BG, plot_bgcolor=PL_BG,
        margin=dict(l=0, r=60, t=10, b=10),
        height=280,
        xaxis=dict(showgrid=True, gridcolor=PL_GRID, zeroline=False,
                   tickfont=dict(size=10, color=PL_FONT)),
        yaxis=dict(showgrid=False, tickfont=dict(size=11, color=PL_FONT)),
        font=dict(family='-apple-system, BlinkMacSystemFont, sans-serif',
                  color=PL_FONT),
        showlegend=False,
    )
    st.plotly_chart(fig_mag, use_container_width=True,
                    config={"displayModeBar": False})

# Familles — barres groupées Commander / Cession
with col_right:
    st.markdown('<div class="section-header">Ruptures par famille</div>',
                unsafe_allow_html=True)
    fam_rows = []
    for fam, grp in arts.groupby('Famille'):
        sub   = plan_df[plan_df['Famille'] == fam]
        cmd_n = int((sub['Action'] == 'COMMANDER').sum())
        ces_n = int((sub['Action'] == 'VOIR CESSION').sum())
        fam_rows.append({'Famille': fam, 'Commander': cmd_n, 'Cession': ces_n,
                         'Total': cmd_n + ces_n})
    fam_df = (pd.DataFrame(fam_rows)
                .sort_values('Total', ascending=True))

    fig_fam = go.Figure()
    fig_fam.add_trace(go.Bar(
        name='Commander', x=fam_df['Commander'], y=fam_df['Famille'],
        orientation='h', marker_color='#E24B4A',
        hovertemplate='%{y}<br>%{x} à commander<extra></extra>',
    ))
    fig_fam.add_trace(go.Bar(
        name='Voir cession', x=fam_df['Cession'], y=fam_df['Famille'],
        orientation='h', marker_color='#63B967',
        hovertemplate='%{y}<br>%{x} cessions possibles<extra></extra>',
    ))
    fig_fam.update_layout(
        barmode='stack',
        paper_bgcolor=PL_BG, plot_bgcolor=PL_BG,
        margin=dict(l=0, r=20, t=10, b=10),
        height=280,
        xaxis=dict(showgrid=True, gridcolor=PL_GRID, zeroline=False,
                   tickfont=dict(size=10, color=PL_FONT)),
        yaxis=dict(showgrid=False, tickfont=dict(size=11, color=PL_FONT)),
        legend=dict(orientation='h', yanchor='bottom', y=1.02,
                    font=dict(size=10)),
        font=dict(family='-apple-system, BlinkMacSystemFont, sans-serif',
                  color=PL_FONT),
    )
    st.plotly_chart(fig_fam, use_container_width=True,
                    config={"displayModeBar": False})

# ── Filtres + table ───────────────────────────────────────────────────────────
st.markdown('<div class="section-header">Plan d\'action — détail par référence</div>',
            unsafe_allow_html=True)

f1, f2, f3, f4 = st.columns([2, 2, 2, 1])
with f1:
    sel_fam = st.selectbox(
        "Famille",
        ['Toutes'] + sorted(plan_df['Famille'].dropna().unique().tolist())
    )
with f2:
    sel_action = st.selectbox(
        "Action",
        ['Toutes', 'COMMANDER', 'VOIR CESSION']
    )
with f3:
    sel_mag = st.selectbox(
        "Magasin en rupture",
        ['Tous'] + sorted(MAG_COLS)
    )
with f4:
    nb_min = st.number_input(
        "Nb rupt. min.", min_value=1, max_value=len(MAG_COLS), value=1
    )

view = plan_df.copy()
if sel_fam    != 'Toutes':    view = view[view['Famille'] == sel_fam]
if sel_action != 'Toutes':    view = view[view['Action']  == sel_action]
if sel_mag    != 'Tous':      view = view[view['Magasins_Rupture'].str.contains(sel_mag, regex=False)]
view = view[view['Nb_Ruptures'] >= nb_min]

display = view[[
    'Code_Article', 'Libellé', 'Famille', 'Nb_Ruptures',
    'Magasins_Rupture', 'Action', 'Site_Donneur', 'Stock_Disponible'
]].copy()
display['Action']           = display['Action'].apply(
    lambda a: '🔴 COMMANDER' if a == 'COMMANDER' else '🟢 VOIR CESSION'
)
display['Stock_Disponible'] = display['Stock_Disponible'].apply(
    lambda x: int(x) if pd.notna(x) and x else ''
)
display.columns = [
    'Code', 'Libellé', 'Famille', 'Nb rupt.',
    'Magasins en rupture', 'Action', 'Site donneur', 'Stock dispo'
]

st.markdown(
    f'<span style="font-size:12px;color:#8E8E93">'
    f'{len(view):,} référence(s) affichée(s)</span>',
    unsafe_allow_html=True
)
st.dataframe(
    display,
    use_container_width=True,
    height=420,
    column_config={
        'Code':                 st.column_config.TextColumn('Code',   width=100),
        'Libellé':              st.column_config.TextColumn('Libellé', width=280),
        'Famille':              st.column_config.TextColumn('Famille', width=160),
        'Nb rupt.':             st.column_config.NumberColumn('Nb rupt.', width=80,
                                    format="%d"),
        'Magasins en rupture':  st.column_config.TextColumn('Magasins en rupture',
                                    width=260),
        'Action':               st.column_config.TextColumn('Action',  width=150),
        'Site donneur':         st.column_config.TextColumn('Site donneur', width=120),
        'Stock dispo':          st.column_config.NumberColumn('Stock dispo', width=90,
                                    format="%d"),
    },
    hide_index=True,
)

# ── Export ────────────────────────────────────────────────────────────────────
st.markdown("---")
dl1, dl2, _ = st.columns([1, 1, 3])
with dl1:
    st.download_button(
        "⬇️ Exporter le plan complet",
        data=to_excel_bytes(plan_df, seuil),
        file_name="plan_action_ruptures.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with dl2:
    cmd_df = plan_df[plan_df['Action'] == 'COMMANDER'].copy()
    st.download_button(
        "⬇️ À commander uniquement",
        data=to_excel_bytes(cmd_df, seuil),
        file_name="a_commander.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
