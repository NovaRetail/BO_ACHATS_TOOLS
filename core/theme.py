"""Charte graphique commune du SmartBuyer Hub (charte CFAO Consumer, style tableau de bord BI, comme le Flash Quotidien).

Usage dans un module :
    from core import theme as T
    T.apply()                                   # CSS commun, à appeler juste après st.set_page_config
    T.sidebar_brand("📊", "Revue CODIR")
    T.page_header("📊", "Revue CODIR", "Sous-titre")
    T.landing_hero("Question clé ?", "Explication")
    T.hero("📊 Revue CODIR", "S42 · Vs N-1", ["CA 947,0 M", ("3 alertes", "alert")])
    T.kpi_row([T.kpi("CA", "947,0 M", "−3,7 %", "neg", tone="blue", icon="💰"), ...])
"""
from __future__ import annotations

import re

import streamlit as st

# ------------------------------------------------------------------ jetons (charte CFAO Consumer, style tableau de bord BI)
NAVY, ORANGE, TEAL, RED = '#24235C', '#F09F39', '#08868D', '#E62244'
BLEU, ROUGE, VERT, VIOLET = '#4B4AA8', RED, TEAL, '#4B4AA8'
NOIR, GRIS, GRIS_TEXTE, GRIS_CLAIR, FOND = '#24235C', '#9A9DB5', '#6B7090', '#F3F4F8', '#F7F7FB'
INK, LIGNE = '#363C5A', '#E3E5EE'
ROUGE_TEXTE, VERT_TEXTE, ORANGE_TEXTE = '#C81E3C', '#077178', '#B4610A'
GRADIENT = 'linear-gradient(120deg, #1B1A4A 0%, #24235C 45%, #3B3A8F 100%)'
PLOTLY_FONT = 'Inter, system-ui, -apple-system, sans-serif'
SHADOW = '0 1px 2px rgba(20,20,60,.04), 0 8px 22px rgba(20,20,60,.05)'


def html(fragment: str) -> None:
    """Affiche du HTML sur une seule ligne (évite que Markdown interprète l'indentation)."""
    st.markdown(re.sub(r'\n\s*', ' ', fragment), unsafe_allow_html=True)


_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html, body, [class*="css"], .stApp, .stMarkdown, button, input, textarea {{ font-family: 'Inter', system-ui, -apple-system, sans-serif; }}
.stApp {{ background-color: {FOND}; }}
.block-container {{ padding-top: 2rem; max-width: 1360px; }}
[data-testid="stSidebar"] {{ background: #FFFFFF; border-right: 1px solid {LIGNE}; }}

/* onglets : contrôle segmenté */
div[data-testid="stTabs"] [role="tablist"], .stTabs [data-baseweb="tab-list"] {{
  background: {GRIS_CLAIR}; border-radius: 12px; padding: 5px; gap: 4px; border-bottom: none !important; box-shadow: none; }}
div[data-testid="stTabs"] [data-testid="stTab"], .stTabs [data-baseweb="tab"] {{
  flex: 1; justify-content: center; border-radius: 9px; height: 38px; background: transparent; border: none !important; margin: 0; }}
div[data-testid="stTabs"] [data-testid="stTab"] div {{ background: transparent !important; }}
div[data-testid="stTabs"] [data-testid="stTab"] p {{ font-weight: 600; color: {GRIS_TEXTE}; font-size: 14px; }}
div[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"] {{ background: #FFFFFF !important; box-shadow: 0 1px 3px rgba(20,20,60,.14); }}
div[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"] p {{ color: {NAVY} !important; font-weight: 700; }}
div[data-testid="stTabs"] [data-baseweb="tab-highlight"], div[data-testid="stTabs"] [data-baseweb="tab-border"],
div[data-testid="stTabs"] .react-aria-SelectionIndicator {{ display: none !important; }}

/* boutons */
.stDownloadButton button, .stButton button[kind="primary"] {{
  background: {NAVY} !important; color: #fff !important; border-radius: 10px !important; font-weight: 700 !important; border: none !important; min-height: 44px; }}
.stDownloadButton button:hover, .stButton button[kind="primary"]:hover {{ background: #34338A !important; }}
.stButton button {{ border-radius: 10px !important; font-weight: 600 !important; min-height: 44px; }}
[data-testid="stFileUploaderDropzone"] {{ border-radius: 12px; background: #FAFAFD; }}
[data-testid="stSidebar"] [data-testid="stAlert"] > div {{ padding: 8px 12px; border-radius: 10px; }}
[data-testid="stSidebar"] [data-testid="stAlert"] p {{ font-size: 12.5px; font-weight: 600; }}

/* en-têtes */
.page-title {{ font-size: 26px; font-weight: 800; color: {NAVY}; padding: 2px 0; letter-spacing: -0.01em; }}
.page-sub {{ font-size: 13.5px; color: {GRIS_TEXTE}; font-weight: 500; margin-bottom: 16px; }}
.section-title {{ display: flex; align-items: center; gap: 10px; font-size: 15px; font-weight: 700; color: {NAVY}; margin: 8px 0 12px 0; }}
.section-title:after {{ content: ""; flex: 1; height: 1px; background: {LIGNE}; }}
.card .section-title {{ margin-top: 0; }}
.side-brand {{ font-size: 18px; font-weight: 800; color: {NAVY}; }}
.side-sub {{ font-size: 12px; color: {GRIS_TEXTE}; font-weight: 500; margin-bottom: 14px; }}
.side-label {{ font-size: 11px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; color: {ORANGE_TEXTE}; margin: 16px 0 6px; }}
.side-foot {{ font-size: 11px; color: {GRIS}; font-weight: 500; margin-top: 18px; padding-top: 10px; border-top: 1px solid {LIGNE}; }}

/* bandeaux */
.hero, .landing-hero {{ position: relative; overflow: hidden; background: {GRADIENT}; border-radius: 18px; color: #fff;
  box-shadow: 0 12px 30px rgba(36,35,92,.22); margin-bottom: 18px; }}
.hero:before, .landing-hero:before {{ content: ""; position: absolute; inset: 0; opacity: .10;
  background-image: linear-gradient(#fff 1px, transparent 1px), linear-gradient(90deg, #fff 1px, transparent 1px);
  background-size: 28px 28px; -webkit-mask-image: linear-gradient(90deg, transparent 30%, #000 100%); mask-image: linear-gradient(90deg, transparent 30%, #000 100%); }}
.hero:after, .landing-hero:after {{ content: ""; position: absolute; right: -80px; top: -80px; width: 280px; height: 280px; border-radius: 50%;
  background: radial-gradient(circle, {ORANGE}55 0%, transparent 70%); }}
.hero > *, .landing-hero > * {{ position: relative; z-index: 1; }}
.hero {{ padding: 24px 28px; }}
.hero-eyebrow, .landing-eyebrow {{ font-size: 11px; letter-spacing: .18em; font-weight: 700; color: {ORANGE}; text-transform: uppercase; }}
.hero-top {{ display: flex; justify-content: space-between; align-items: flex-end; gap: 10px; flex-wrap: wrap; margin-top: 4px; }}
.hero-title {{ font-size: 30px; font-weight: 300; color: #fff; letter-spacing: .01em; }}
.hero-title b {{ font-weight: 800; }}
.hero-sub {{ font-size: 12.5px; color: #CFD0E6; font-weight: 500; }}
.hero-chips {{ display: flex; gap: 8px; margin-top: 16px; flex-wrap: wrap; }}
.chip {{ background: rgba(255,255,255,.10); border: 1px solid rgba(255,255,255,.18); color: #fff; font-size: 13px; font-weight: 600;
  padding: 6px 14px; border-radius: 999px; }}
.chip-alert {{ background: rgba(230,34,68,.30); border-color: rgba(255,120,140,.45); }}
.chip-ok {{ background: rgba(8,134,141,.38); border-color: rgba(120,220,225,.35); }}
.landing-hero {{ padding: 28px 32px; display: flex; gap: 28px; align-items: stretch; }}
.landing-main {{ flex: 1.4; }}
.landing-q {{ font-size: 30px; font-weight: 300; color: #fff; margin: 6px 0 8px; line-height: 1.15; }}
.landing-q b {{ font-weight: 800; }}
.landing-t {{ font-size: 14.5px; color: #CFD0E6; font-weight: 400; line-height: 1.55; max-width: 680px; }}
.landing-status {{ flex: 1; min-width: 280px; background: rgba(255,255,255,.07); border: 1px solid rgba(255,255,255,.16);
  border-radius: 14px; padding: 16px 18px; }}
.landing-status h6 {{ margin: 0 0 4px; font-size: 11px; letter-spacing: .12em; text-transform: uppercase; color: #B9BAD6; font-weight: 600; }}
.ls-prog {{ display: flex; align-items: baseline; gap: 8px; margin-bottom: 8px; }}
.ls-prog .n {{ font-size: 32px; font-weight: 800; color: #fff; }} .ls-prog .t {{ color: #B9BAD6; font-size: 13px; }}
.ls-bar {{ height: 6px; border-radius: 6px; background: rgba(255,255,255,.14); overflow: hidden; margin-bottom: 10px; }}
.ls-bar span {{ display: block; height: 100%; background: linear-gradient(90deg, {ORANGE}, #FFC27A); border-radius: 6px; }}
.ls-row {{ display: flex; align-items: center; gap: 10px; padding: 7px 0; font-size: 13.5px; color: #E7E8F5; border-top: 1px solid rgba(255,255,255,.08); }}
.ls-dot {{ width: 20px; height: 20px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 800; flex: none; }}
.ls-dot.ok {{ background: {TEAL}; color: #fff; }} .ls-dot.wait {{ border: 1.5px dashed #8E8FB8; }} .ls-dot.err {{ background: {RED}; color: #fff; }}
.ls-row .r {{ margin-left: auto; font-size: 11.5px; color: #B9BAD6; }}
.steps {{ display: grid; grid-template-columns: repeat(var(--n), 1fr); margin: 4px 0 22px; }}
.step {{ position: relative; padding-right: 14px; }}
.step:not(:last-child):after {{ content: ""; position: absolute; top: 17px; left: 44px; right: 8px; height: 2px; background: linear-gradient(90deg, {ORANGE}, {LIGNE}); }}
.step .num {{ width: 36px; height: 36px; border-radius: 10px; background: #fff; border: 2px solid {ORANGE}; color: {NAVY}; display: flex;
  align-items: center; justify-content: center; font-weight: 800; position: relative; z-index: 1; font-size: 13px; }}
.step .lab {{ font-weight: 700; color: {NAVY}; margin-top: 10px; font-size: 14.5px; }}
.step .txt {{ color: {GRIS_TEXTE}; font-size: 12.5px; line-height: 1.45; margin-top: 2px; padding-right: 10px; }}

/* cartes */
[data-testid="stPlotlyChart"] {{ background: #fff; border: 1px solid {LIGNE}; border-radius: 16px; padding: 0; overflow: hidden; box-shadow: {SHADOW}; }}
.card {{ background: #fff; border-radius: 16px; padding: 18px 20px; border: 1px solid {LIGNE}; margin-bottom: 14px; box-shadow: {SHADOW}; }}
.card-tall {{ height: 100%; }}
.rule-row {{ display: flex; justify-content: space-between; gap: 12px; padding: 8px 0; font-size: 13px; font-weight: 600; color: {INK}; border-bottom: 1px solid #F0F1F5; }}
.rule-row:last-child {{ border-bottom: none; }}
.rule-row span:last-child {{ text-align: right; }}
.col-chip {{ display: inline-block; background: {GRIS_CLAIR}; border: 1px solid {LIGNE}; color: {INK}; font-size: 11.5px; font-weight: 600; padding: 2px 8px; border-radius: 6px; margin: 3px 3px 0 0; }}
.col-group {{ font-size: 10.5px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; color: {GRIS_TEXTE}; margin: 12px 0 4px; }}
.note {{ font-size: 12px; color: {GRIS_TEXTE}; font-weight: 500; margin-top: 10px; line-height: 1.45; }}

/* KPI */
.kpi-row {{ display: grid; gap: 14px; margin-bottom: 18px; }}
.kpi-card {{ position: relative; overflow: hidden; border-radius: 16px; padding: 16px 18px 14px; display: flex; flex-direction: column; gap: 8px;
  background: #fff; border: 1px solid {LIGNE}; box-shadow: {SHADOW}; --acc: {NAVY}; }}
.kpi-card:before {{ content: ""; position: absolute; left: 0; top: 0; right: 0; height: 3px; background: var(--acc); }}
.kpi-blue {{ --acc: {NAVY}; }} .kpi-violet {{ --acc: {ORANGE}; }} .kpi-teal {{ --acc: {TEAL}; }}
.kpi-top {{ display: flex; justify-content: space-between; align-items: center; }}
.kpi-label {{ font-size: 12.5px; font-weight: 600; color: {GRIS_TEXTE}; }}
.kpi-ico {{ width: 30px; height: 30px; border-radius: 9px; display: flex; align-items: center; justify-content: center; font-size: 14px;
  background: {GRIS_CLAIR}; }}
.kpi-value {{ font-size: 28px; font-weight: 800; line-height: 1.05; color: {NAVY}; letter-spacing: -.01em; }}
.kpi-unit {{ font-size: 13px; font-weight: 600; color: {GRIS_TEXTE}; }}
.kpi-foot {{ display: flex; justify-content: space-between; align-items: center; gap: 6px; flex-wrap: wrap; }}
.kpi-ref {{ font-size: 11.5px; font-weight: 500; color: {GRIS_TEXTE}; }}
.pill {{ display: inline-block; font-size: 12px; font-weight: 700; padding: 3px 10px; border-radius: 999px; width: fit-content; }}
.pill-neg {{ background: {RED}14; color: {ROUGE_TEXTE}; }} .pill-pos {{ background: {TEAL}17; color: {VERT_TEXTE}; }} .pill-mut {{ background: {GRIS_CLAIR}; color: {GRIS_TEXTE}; }}

/* badges, alertes, barres */
.badge {{ font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 999px; white-space: nowrap; }}
.badge-crit {{ background: {RED}14; color: {ROUGE_TEXTE}; }} .badge-att {{ background: {ORANGE}1F; color: {ORANGE_TEXTE}; }}
.badge-ok {{ background: {TEAL}17; color: {VERT_TEXTE}; }} .badge-info {{ background: #24235C12; color: {NAVY}; }}
.alert-card {{ background: #fff; padding: 13px 16px; margin-bottom: 12px; border: 1px solid {LIGNE}; border-left: 4px solid {GRIS}; border-radius: 0 14px 14px 0; box-shadow: {SHADOW}; }}
.alert-crit {{ border-left-color: {RED}; background: #FFF7F8; }} .alert-att {{ border-left-color: {ORANGE}; background: #FFF9F1; }}
.alert-ok {{ border-left-color: {TEAL}; }} .alert-info {{ border-left-color: {NAVY}; }}
.alert-title {{ font-size: 14px; font-weight: 700; color: {NAVY}; }}
.alert-line {{ font-size: 12.5px; color: {INK}; font-weight: 400; margin-top: 3px; line-height: 1.5; }}
.dv-row {{ display: grid; grid-template-columns: minmax(0,1.7fr) minmax(0,1fr) 52px; gap: 10px; align-items: center; padding: 7px 0;
  font-size: 13px; font-weight: 600; color: {NAVY}; border-bottom: 1px solid #F3F4F8; }}
.dv-row:last-child {{ border-bottom: none; }}
.dv-row small {{ display: block; font-size: 11.5px; color: {GRIS_TEXTE}; font-weight: 500; }}
.dv-track {{ position: relative; height: 10px; background: {GRIS_CLAIR}; border-radius: 5px; }}
.dv-zero {{ position: absolute; left: 50%; top: -3px; bottom: -3px; width: 1px; background: #C9CBD8; }}
.dv-bar {{ position: absolute; top: 0; height: 10px; border-radius: 5px; }}
.tbl {{ width: 100%; border-collapse: separate; border-spacing: 0; font-size: 13px; }}
.tbl th {{ font-size: 10.5px; font-weight: 700; letter-spacing: .06em; text-transform: uppercase; color: {GRIS_TEXTE}; text-align: right;
  padding: 8px 8px; border-bottom: 1px solid {LIGNE}; background: #F7F7FB; }}
.tbl th:first-child, .tbl td:first-child {{ text-align: left; }}
.tbl td {{ text-align: right; padding: 9px 8px; font-weight: 600; color: {NAVY}; border-bottom: 1px solid #F0F1F5; }}
.tbl tr:last-child td {{ border-bottom: none; }}
.tbl td.hm {{ border-radius: 8px; border-bottom: 3px solid #fff; border-left: 3px solid #fff; font-weight: 700; }}
.neg {{ color: {ROUGE_TEXTE}; font-weight: 700; }} .pos {{ color: {VERT_TEXTE}; font-weight: 700; }} .mut {{ color: {GRIS_TEXTE}; font-weight: 500; }}
ul.bul {{ margin: 0; padding-left: 18px; }} ul.bul li {{ margin-bottom: 9px; font-size: 13.5px; line-height: 1.5; font-weight: 500; color: {INK}; }}
ul.bul li::marker {{ color: {ORANGE}; }}
@media (max-width: 900px) {{ .kpi-row {{ grid-template-columns: repeat(2, minmax(0,1fr)) !important; }} .landing-hero {{ flex-direction: column; }} }}
</style>
"""


def apply() -> None:
    st.markdown(_CSS, unsafe_allow_html=True)


# ------------------------------------------------------------------ composants
def sidebar_brand(icon: str, name: str, sub: str = 'SmartBuyer Hub') -> None:
    with st.sidebar:
        html(f"<div class='side-brand'>{icon} {name}</div><div class='side-sub'>{sub}</div>")


def sidebar_label(text: str) -> None:
    with st.sidebar:
        html(f"<div class='side-label'>{text}</div>")


def sidebar_footer(text: str) -> None:
    with st.sidebar:
        html(f"<div class='side-foot'>{text}</div>")


def page_header(icon: str, title: str, sub: str = '') -> None:
    html(f"<div class='page-title'>{icon} {title}</div>" + (f"<div class='page-sub'>{sub}</div>" if sub else ''))


def landing_hero(question: str, text: str, eyebrow: str = '', status: str = '') -> None:
    """status : HTML d'un panneau à droite (voir landing_status)."""
    eb = f"<div class='landing-eyebrow'>{eyebrow}</div>" if eyebrow else ''
    html(f"<div class='landing-hero'><div class='landing-main'>{eb}<div class='landing-q'>{question}</div>"
         f"<div class='landing-t'>{text}</div></div>{status}</div>")


def landing_status(title: str, rows, unit: str = 'fichiers reconnus', count: int | None = None) -> str:
    """rows : [(libellé, état 'ok' | 'err' | 'wait', texte à droite)] ; count : nombre de lignes comptées dans la jauge."""
    k = count or len(rows)
    n = sum(1 for _, s, _ in rows[:k] if s == 'ok')
    body = ''.join(f"<div class='ls-row'><div class='ls-dot {s}'>{'✓' if s == 'ok' else '✕' if s == 'err' else ''}</div>{lab}"
                   f"<span class='r'>{r}</span></div>" for lab, s, r in rows)
    return (f"<div class='landing-status'><h6>{title}</h6><div class='ls-prog'><span class='n'>{n}/{k}</span>"
            f"<span class='t'>{unit}</span></div><div class='ls-bar'><span style='width:{max(n / k * 100, 3):.0f}%'></span></div>{body}</div>")


def steps(items) -> None:
    """items : [(titre, texte)] — parcours numéroté."""
    body = ''.join(f"<div class='step'><div class='num'>{i}</div><div class='lab'>{a}</div><div class='txt'>{b}</div></div>"
                   for i, (a, b) in enumerate(items, 1))
    html(f"<div class='steps' style='--n:{len(items)}'>{body}</div>")


def hero(title: str, sub: str = '', chips=(), eyebrow: str = '') -> None:
    cs = ''
    for c in chips:
        txt, kind = (c, '') if isinstance(c, str) else c
        cs += f"<span class='chip {'chip-' + kind if kind else ''}'>{txt}</span>"
    eb = f"<div class='hero-eyebrow'>{eyebrow}</div>" if eyebrow else ''
    html(f"<div class='hero'>{eb}<div class='hero-top'><span class='hero-title'>{title}</span><span class='hero-sub'>{sub}</span></div>"
         f"<div class='hero-chips'>{cs}</div></div>")


def rules_card(title: str, rows, tall: bool = True) -> None:
    """rows : [(gauche, droite)] ; la droite peut contenir du HTML."""
    body = ''.join(f"<div class='rule-row'><span>{a}</span><span class='mut'>{b}</span></div>" for a, b in rows)
    html(f"<div class='card {'card-tall' if tall else ''}'><div class='section-title'>{title}</div>{body}</div>")


def columns_card(title: str, groups, note: str = '') -> None:
    """groups : [(nom du fichier, [colonnes])]."""
    body = ''
    for g, cols in groups:
        body += f"<div class='col-group'>{g}</div>" + ''.join(f"<span class='col-chip'>{c}</span>" for c in cols)
    html(f"<div class='card card-tall'><div class='section-title'>{title}</div>{body}" + (f"<div class='note'>{note}</div>" if note else '') + '</div>')


def kpi(label, value, delta='', sign='mut', ref='', tone='', icon='', unit='') -> str:
    ico = f"<div class='kpi-ico'>{icon}</div>" if icon else ''
    pill = f"<span class='pill pill-{sign}'>{delta}</span>" if delta else ''
    return (f"<div class='kpi-card {'kpi-' + tone if tone else ''}'><div class='kpi-top'><span class='kpi-label'>{label}</span>{ico}</div>"
            f"<div class='kpi-value'>{value} <span class='kpi-unit'>{unit}</span></div>"
            f"<div class='kpi-foot'>{pill}<span class='kpi-ref'>{ref}</span></div></div>")


def kpi_row(cards) -> None:
    html(f"<div class='kpi-row' style='grid-template-columns:repeat({len(cards)},minmax(0,1fr))'>{''.join(cards)}</div>")


def section(title: str) -> None:
    html(f"<div class='section-title'>{title}</div>")


def card(inner_html: str, title: str = '') -> None:
    html(f"<div class='card'>{f'<div class=section-title>{title}</div>' if title else ''}{inner_html}</div>")


def alert(title: str, line: str = '', kind: str = 'info') -> None:
    html(f"<div class='alert-card alert-{kind}'><div class='alert-title'>{title}</div>" + (f"<div class='alert-line'>{line}</div>" if line else '') + '</div>')


def badge(text: str, kind: str = 'info') -> str:
    return f"<span class='badge badge-{kind}'>{text}</span>"


def dv_row(name: str, sub: str, value: float, vmax: float, label: str, small: bool = False) -> str:
    """Barre horizontale de part et d'autre de zéro (écarts positifs / négatifs)."""
    w = abs(value) / vmax * 48 if vmax else 0
    col = '#C9CBD8' if small else (ROUGE if value < 0 else VERT)
    left = 50 - w if value < 0 else 50
    cls = 'mut' if small else ('neg' if value < 0 else 'pos')
    return (f"<div class='dv-row'><div>{name}{f'<small>{sub}</small>' if sub else ''}</div>"
            f"<div class='dv-track'><div class='dv-zero'></div><div class='dv-bar' style='left:{left}%;width:{w}%;background:{col}'></div></div>"
            f"<b class='{cls}' style='text-align:right'>{label}</b></div>")


def heat(value: float, scale: float, small: bool = False):
    """Fond et couleur de texte d'une cellule de heatmap (rouge = recul, vert = gain)."""
    if small:
        return GRIS_CLAIR, GRIS_TEXTE
    t = min(abs(value) / scale, 1) if scale else 0
    a = 0.12 + 0.78 * t
    bg = f'rgba(230,34,68,{a:.2f})' if value < 0 else f'rgba(8,134,141,{a:.2f})'
    fg = '#FFFFFF' if t > 0.45 else (ROUGE_TEXTE if value < 0 else VERT_TEXTE)
    return bg, fg


def plotly_layout(fig, height=340):
    fig.update_layout(height=height, margin=dict(l=14, r=14, t=18, b=14), plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF',
                      font=dict(family=PLOTLY_FONT, color=NOIR, size=13), showlegend=False)
    return fig
