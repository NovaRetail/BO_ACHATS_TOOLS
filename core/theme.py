"""Charte graphique commune du SmartBuyer Hub (rendu « premium », issu de PGC Radar).

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

# ------------------------------------------------------------------ jetons
BLEU, ROUGE, VERT, ORANGE, VIOLET = '#007AFF', '#FF3B30', '#34C759', '#FF9500', '#AF52DE'
NAVY = '#0A2540'
NOIR, GRIS, GRIS_TEXTE, GRIS_CLAIR, FOND = '#1C1C1E', '#8E8E93', '#6B6F80', '#F2F2F7', '#F5F6FA'
ROUGE_TEXTE, VERT_TEXTE, ORANGE_TEXTE = '#C82D24', '#1E8A45', '#B36200'
GRADIENT = 'linear-gradient(135deg, #0A2540 0%, #1D4E89 55%, #6A5ACD 100%)'
PLOTLY_FONT = 'Nunito, -apple-system, BlinkMacSystemFont, sans-serif'


def html(fragment: str) -> None:
    """Affiche du HTML sur une seule ligne (évite que Markdown interprète l'indentation)."""
    st.markdown(re.sub(r'\n\s*', ' ', fragment), unsafe_allow_html=True)


_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap');
html, body, [class*="css"], .stApp, .stMarkdown, button, input, textarea {{ font-family: 'Nunito', -apple-system, BlinkMacSystemFont, sans-serif; }}
.stApp {{ background-color: {FOND}; }}
.block-container {{ padding-top: 2.2rem; max-width: 1320px; }}
[data-testid="stSidebar"] {{ background: #FFFFFF; border-right: 0.5px solid #E5E5EA; }}

/* onglets : segmented control iOS */
div[data-testid="stTabs"] [role="tablist"], .stTabs [data-baseweb="tab-list"] {{
  background: #E9EAF0; border-radius: 12px; padding: 3px; gap: 0; border-bottom: none !important; }}
div[data-testid="stTabs"] [data-testid="stTab"], .stTabs [data-baseweb="tab"] {{
  flex: 1; justify-content: center; border-radius: 10px; height: 38px; background: transparent; border: none !important; }}
div[data-testid="stTabs"] [data-testid="stTab"] p {{ font-weight: 800; color: {GRIS_TEXTE}; font-size: 14px; }}
div[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"] {{ background: #FFFFFF !important; box-shadow: 0 1px 3px rgba(0,0,0,.08); }}
div[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"] p {{ color: {NOIR} !important; }}
div[data-testid="stTabs"] [data-baseweb="tab-highlight"], div[data-testid="stTabs"] [data-baseweb="tab-border"],
div[data-testid="stTabs"] .react-aria-SelectionIndicator {{ display: none !important; }}

/* boutons */
.stDownloadButton button, .stButton button[kind="primary"] {{
  background: {NAVY} !important; color: #fff !important; border-radius: 12px !important; font-weight: 800 !important; border: none !important; min-height: 44px; }}
.stButton button {{ border-radius: 12px !important; font-weight: 700 !important; min-height: 44px; }}
[data-testid="stFileUploaderDropzone"] {{ border-radius: 14px; background: #FAFAFC; }}

/* en-têtes */
.page-title {{ font-size: 26px; font-weight: 900; color: {NOIR}; padding: 2px 0; letter-spacing: -0.01em; }}
.page-sub {{ font-size: 13.5px; color: {GRIS_TEXTE}; font-weight: 600; margin-bottom: 16px; }}
.section-title {{ font-size: 15px; font-weight: 900; color: {NOIR}; margin: 6px 0 10px 0; }}
.side-brand {{ font-size: 18px; font-weight: 900; color: {NOIR}; }}
.side-sub {{ font-size: 12px; color: {GRIS_TEXTE}; font-weight: 600; margin-bottom: 14px; }}
.side-label {{ font-size: 11px; font-weight: 800; letter-spacing: .07em; text-transform: uppercase; color: {GRIS_TEXTE}; margin: 14px 0 6px; }}
.side-foot {{ font-size: 11px; color: {GRIS}; font-weight: 600; margin-top: 18px; padding-top: 10px; border-top: 0.5px solid #E5E5EA; }}

/* bandeau héro */
.hero {{ background: {GRADIENT}; border-radius: 18px; padding: 18px 22px; margin-bottom: 16px; }}
.hero-top {{ display: flex; justify-content: space-between; align-items: center; gap: 10px; flex-wrap: wrap; }}
.hero-title {{ font-size: 20px; font-weight: 900; color: #fff; }}
.hero-sub {{ font-size: 12px; color: rgba(255,255,255,0.78); font-weight: 700; }}
.hero-chips {{ display: flex; gap: 8px; margin-top: 12px; flex-wrap: wrap; }}
.chip {{ background: rgba(255,255,255,0.15); color: #fff; font-size: 12.5px; font-weight: 800; padding: 5px 13px; border-radius: 12px; }}
.chip-alert {{ background: rgba(255,90,80,0.42); }}
.chip-ok {{ background: rgba(52,199,89,0.38); }}
.landing-hero {{ background: {GRADIENT}; border-radius: 18px; padding: 22px 24px; margin-bottom: 16px; }}
.landing-q {{ font-size: 18px; font-weight: 900; color: #fff; margin-bottom: 6px; }}
.landing-t {{ font-size: 13.5px; color: rgba(255,255,255,0.88); font-weight: 600; line-height: 1.5; }}

/* cartes */
.card {{ background: #fff; border-radius: 16px; padding: 16px 18px; border: 1px solid rgba(0,0,0,0.04); margin-bottom: 14px; }}
.card-tall {{ height: 100%; }}
.rule-row {{ display: flex; justify-content: space-between; gap: 12px; padding: 8px 0; font-size: 13.5px; font-weight: 700; color: {NOIR}; border-bottom: 0.5px solid #F0F0F3; }}
.rule-row:last-child {{ border-bottom: none; }}
.rule-row span:last-child {{ text-align: right; }}
.col-chip {{ display: inline-block; background: {GRIS_CLAIR}; color: {NOIR}; font-size: 11.5px; font-weight: 700; padding: 3px 10px; border-radius: 10px; margin: 3px 3px 0 0; }}
.col-group {{ font-size: 12px; font-weight: 900; color: {NOIR}; margin: 10px 0 2px; }}
.note {{ font-size: 12px; color: {GRIS_TEXTE}; font-weight: 600; margin-top: 10px; line-height: 1.45; }}

/* KPI */
.kpi-row {{ display: grid; gap: 12px; margin-bottom: 16px; }}
.kpi-card {{ border-radius: 18px; padding: 15px 17px; display: flex; flex-direction: column; gap: 6px; background: #fff; border: 1px solid rgba(0,0,0,0.04); }}
.kpi-blue {{ background: #EAF2FF; border-color: transparent; }} .kpi-violet {{ background: #F3EDFC; border-color: transparent; }}
.kpi-top {{ display: flex; justify-content: space-between; align-items: center; }}
.kpi-label {{ font-size: 13px; font-weight: 800; color: {GRIS_TEXTE}; }}
.kpi-blue .kpi-label {{ color: #4A6FA5; }} .kpi-violet .kpi-label {{ color: #7A5FA8; }}
.kpi-ico {{ width: 34px; height: 34px; border-radius: 11px; display: flex; align-items: center; justify-content: center; font-size: 16px; background: {GRIS_CLAIR}; }}
.kpi-blue .kpi-ico {{ background: {BLEU}; }} .kpi-violet .kpi-ico {{ background: {VIOLET}; }}
.kpi-value {{ font-size: 29px; font-weight: 900; line-height: 1.05; color: {NOIR}; }}
.kpi-blue .kpi-value {{ color: #0C2D5B; }} .kpi-violet .kpi-value {{ color: #3E2467; }}
.kpi-unit {{ font-size: 14px; font-weight: 700; color: {GRIS_TEXTE}; }}
.kpi-foot {{ display: flex; justify-content: space-between; align-items: center; gap: 6px; flex-wrap: wrap; }}
.kpi-ref {{ font-size: 11.5px; font-weight: 700; color: {GRIS_TEXTE}; }}
.pill {{ display: inline-block; font-size: 12px; font-weight: 800; padding: 3px 11px; border-radius: 12px; width: fit-content; }}
.pill-neg {{ background: #FFE0DD; color: {ROUGE_TEXTE}; }} .pill-pos {{ background: #E4F8EA; color: {VERT_TEXTE}; }} .pill-mut {{ background: {GRIS_CLAIR}; color: {GRIS_TEXTE}; }}

/* badges, alertes, barres */
.badge {{ font-size: 11px; font-weight: 900; padding: 3px 11px; border-radius: 12px; white-space: nowrap; }}
.badge-crit {{ background: #FFE5E3; color: {ROUGE_TEXTE}; }} .badge-att {{ background: #FFF2DE; color: {ORANGE_TEXTE}; }}
.badge-ok {{ background: #E4F8EA; color: {VERT_TEXTE}; }} .badge-info {{ background: #EAF2FF; color: #1D4E89; }}
.alert-card {{ background: #fff; padding: 13px 16px; margin-bottom: 12px; border: 1px solid rgba(0,0,0,0.04); border-left: 4px solid {GRIS}; border-radius: 0 16px 16px 0; }}
.alert-crit {{ border-left-color: {ROUGE}; }} .alert-att {{ border-left-color: {ORANGE}; }} .alert-ok {{ border-left-color: {VERT}; }} .alert-info {{ border-left-color: {BLEU}; }}
.alert-title {{ font-size: 14px; font-weight: 900; color: {NOIR}; }}
.alert-line {{ font-size: 12.5px; color: {GRIS_TEXTE}; font-weight: 700; margin-top: 3px; line-height: 1.45; }}
.dv-row {{ display: grid; grid-template-columns: minmax(0,1.35fr) minmax(0,1fr) 62px; gap: 10px; align-items: center; padding: 6px 0; font-size: 13.5px; font-weight: 700; color: {NOIR}; }}
.dv-row small {{ display: block; font-size: 11.5px; color: {GRIS_TEXTE}; font-weight: 700; }}
.dv-track {{ position: relative; height: 12px; }}
.dv-zero {{ position: absolute; left: 50%; top: -4px; bottom: -4px; width: 1px; background: #C7C7CC; }}
.dv-bar {{ position: absolute; top: 1px; height: 10px; border-radius: 5px; }}
.tbl {{ width: 100%; border-collapse: separate; border-spacing: 0; font-size: 13.5px; }}
.tbl th {{ font-size: 11.5px; font-weight: 800; color: {GRIS_TEXTE}; text-align: right; padding: 6px 8px; border-bottom: 1px solid #ECECF0; }}
.tbl th:first-child, .tbl td:first-child {{ text-align: left; }}
.tbl td {{ text-align: right; padding: 9px 8px; font-weight: 700; color: {NOIR}; border-bottom: 0.5px solid #F0F0F3; }}
.tbl tr:last-child td {{ border-bottom: none; }}
.tbl td.hm {{ border-radius: 8px; border-bottom: 2px solid #fff; }}
.neg {{ color: {ROUGE_TEXTE}; font-weight: 800; }} .pos {{ color: {VERT_TEXTE}; font-weight: 800; }} .mut {{ color: {GRIS_TEXTE}; font-weight: 700; }}
ul.bul {{ margin: 0; padding-left: 18px; }} ul.bul li {{ margin-bottom: 8px; font-size: 14px; line-height: 1.5; font-weight: 600; color: {NOIR}; }}
ul.bul li::marker {{ color: {BLEU}; }}
@media (max-width: 900px) {{ .kpi-row {{ grid-template-columns: repeat(2, minmax(0,1fr)) !important; }} }}
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


def landing_hero(question: str, text: str) -> None:
    html(f"<div class='landing-hero'><div class='landing-q'>{question}</div><div class='landing-t'>{text}</div></div>")


def hero(title: str, sub: str = '', chips=()) -> None:
    cs = ''
    for c in chips:
        txt, kind = (c, '') if isinstance(c, str) else c
        cs += f"<span class='chip {'chip-' + kind if kind else ''}'>{txt}</span>"
    html(f"<div class='hero'><div class='hero-top'><span class='hero-title'>{title}</span><span class='hero-sub'>{sub}</span></div>"
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
    col = '#C7C7CC' if small else (ROUGE if value < 0 else VERT)
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
    a = 0.14 + 0.62 * t
    bg = f'rgba(255,59,48,{a:.2f})' if value < 0 else f'rgba(52,199,89,{a:.2f})'
    fg = '#FFFFFF' if t > 0.65 else (ROUGE_TEXTE if value < 0 else VERT_TEXTE)
    return bg, fg


def plotly_layout(fig, height=340):
    fig.update_layout(height=height, margin=dict(l=8, r=8, t=8, b=8), plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF',
                      font=dict(family=PLOTLY_FONT, color=NOIR, size=13), showlegend=False)
    return fig
