"""SmartBuyer Hub — Module 26 · Revue CODIR (rapport hebdomadaire à la Direction Générale).

Entrées : exports Power BI A (famille × site), B (rayon × site), C (site) ; classeurs « Causes et actions » complétés.
Sorties : classeur des chefs de groupe (valeurs, sans formule) et rapport PowerPoint 5 pages + annexe sur le masque.
Calculs : package revue_codir/ (engine, excel_cdg, ppt_report).
"""
from __future__ import annotations

import io
import os
import shutil
import subprocess
import tempfile
import warnings
from pathlib import Path

import numpy as np
import plotly.graph_objects as go
import streamlit as st

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from revue_codir import engine as E  # noqa: E402
from revue_codir import excel_cdg as X  # noqa: E402
from revue_codir import ppt_report as P  # noqa: E402

warnings.filterwarnings('ignore')
ROOT = Path(__file__).resolve().parent.parent
MASQUE = ROOT / 'revue_codir' / 'assets' / 'masque_cfao_consumer.pptx'  # non versionné (dépôt public)

st.set_page_config(page_title='Revue de performance', page_icon='📊', layout='wide')

NAVY, ORANGE = '#24235C', '#F19F39'
NEG, POS, SMALL = '#C0392B', '#1E8449', '#B5B9CC'

st.markdown(f"""
<style>
html, body, [class*="css"], .stMarkdown, .stText {{ font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif; }}
.stApp {{ background: #F2F3F7; }}
.block-container {{ padding-top: 3.2rem; max-width: 1280px; }}
.rp-bar {{ background:{NAVY}; color:#fff; padding:14px 22px; border-radius:4px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; }}
.rp-bar b {{ font-weight:700; letter-spacing:.02em; }}
.rp-bar span.l {{ font-weight:300; }}
.rp-meta {{ font-size:13px; color:#C5C8E0; }}
.stTabs [data-testid="stTabs"] {{ gap:4px; }}
[data-testid="stTab"] {{ font-size:15px; color:#4E5370; padding:10px 14px; }}
[data-testid="stTab"][aria-selected="true"] {{ color:{NAVY} !important; font-weight:700; border-bottom:3px solid {ORANGE} !important; }}
.slc {{ display:flex; gap:10px; flex-wrap:wrap; margin:6px 0 10px; }}
.sl {{ background:#fff; border:1px solid #D6D9E4; border-radius:4px; padding:6px 12px; }}
.sl small {{ display:block; font-size:10.5px; font-weight:700; letter-spacing:.08em; text-transform:uppercase; color:#4E5370; }}
.sl b {{ font-size:14px; color:{NAVY}; }}
.ctx {{ background:#fff; border:1px solid #D6D9E4; border-left:4px solid {ORANGE}; border-radius:4px; padding:10px 14px; font-size:14px; line-height:1.5; margin-bottom:12px; }}
.kpis {{ display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:12px; margin-bottom:14px; }}
.tile {{ background:#fff; border:1px solid #D6D9E4; border-radius:4px; padding:12px 14px; border-bottom:4px solid {NAVY}; }}
.tile small {{ font-size:11px; font-weight:700; letter-spacing:.08em; text-transform:uppercase; color:#4E5370; }}
.tile .v {{ font-size:26px; font-weight:700; color:{NAVY}; line-height:1.15; }}
.tile .d {{ display:flex; justify-content:space-between; font-size:13px; border-top:1px solid #ECEDF2; padding-top:6px; margin-top:4px; }}
.neg {{ color:#B3261E; }} .pos {{ color:#1B7A45; }} .mut {{ color:#6B7090; }}
.vh {{ background:{NAVY}; color:#fff; padding:9px 14px; font-size:13.5px; text-transform:uppercase; letter-spacing:.02em; border-radius:4px 4px 0 0; }}
.vh span {{ font-weight:300; }}
.vb {{ background:#fff; border:1px solid #D6D9E4; border-top:0; border-radius:0 0 4px 4px; padding:12px 14px; margin-bottom:14px; }}
.fm li {{ margin-bottom:8px; font-size:14.5px; line-height:1.45; }}
.fm li::marker {{ color:{ORANGE}; }}
table.ot {{ width:100%; border-collapse:collapse; font-size:14px; }}
table.ot th {{ background:{ORANGE}; color:#fff; font-style:italic; text-align:right; padding:7px 8px; }}
table.ot th:first-child, table.ot td:first-child {{ text-align:left; }}
table.ot td {{ text-align:right; padding:8px; border-bottom:1px solid #fff; }}
table.ot tbody tr:nth-child(odd) td {{ background:#ECEDF2; }}
.cr {{ display:grid; grid-template-columns:minmax(0,1.3fr) minmax(0,1fr) 54px; gap:8px; align-items:center; font-size:13.5px; margin-bottom:8px; }}
.cr small {{ display:block; font-size:11.5px; color:#6B7090; }}
.bar {{ position:relative; height:14px; }}
.bar .z {{ position:absolute; left:50%; top:-4px; bottom:-4px; width:1px; background:#9AA0BC; }}
.bar .b {{ position:absolute; top:2px; height:10px; }}
.pt {{ font-size:11.5px; font-weight:700; letter-spacing:.06em; text-transform:uppercase; color:{NAVY}; margin:6px 0 8px; }}
.card-h {{ background:{NAVY}; color:#fff; padding:10px 14px; border-radius:4px 4px 0 0; display:flex; justify-content:space-between; align-items:center; }}
.card-h b {{ font-size:16px; text-transform:uppercase; }}
.tag {{ font-size:11.5px; font-weight:700; border-radius:12px; padding:3px 10px; }}
.tneg {{ background:#FBEAE8; color:#B3261E; }} .tpos {{ background:#E7F4EC; color:#1B7A45; }}
.card-b {{ background:#fff; border:1px solid #D6D9E4; border-top:0; padding:12px 14px; border-radius:0 0 4px 4px; }}
.cm {{ background:#F7F8FB; border:1px solid #C9CCDB; border-radius:4px; padding:10px 12px; font-size:13.5px; line-height:1.45; }}
.lex dt {{ font-weight:700; color:{NAVY}; margin-top:8px; }} .lex dd {{ margin:0 0 4px 0; font-size:14px; color:#2B3644; }}
@media (max-width: 900px) {{ .kpis {{ grid-template-columns:repeat(2,minmax(0,1fr)); }} }}
</style>
""", unsafe_allow_html=True)


# ------------------------------------------------------------------ helpers
def cls(v):
    return 'neg' if v < 0 else 'pos'


def bar_row(name, why, v, vmax):
    w = abs(v) / vmax * 48 if vmax else 0
    small = abs(v) < E.SEUIL_SIGNIF
    col = SMALL if small else (NEG if v < 0 else POS)
    left = 50 - w if v < 0 else 50
    return (f'<div class="cr"><div><b>{name}</b><small>{why}</small></div>'
            f'<div class="bar"><div class="z"></div><div class="b" style="left:{left}%;width:{w}%;background:{col}"></div></div>'
            f'<b class="{"mut" if small else cls(v)}" style="text-align:right">{E.fdm(v)}</b></div>')


def heat_css(v):
    if abs(v) < E.SEUIL_SIGNIF:
        return '#ECEDF2', '#6B7090'
    t = min(abs(v) / 5e6, 1)
    a = 0.15 + 0.65 * t
    bg = f'rgba(192,57,43,{a:.2f})' if v < 0 else f'rgba(30,132,73,{a:.2f})'
    return bg, ('#FFFFFF' if abs(v) > 3e6 else ('#8E1F17' if v < 0 else '#145A32'))


@st.cache_data(show_spinner=False)
def compute_from_bytes(blobs: tuple):
    files = []
    for name, data in blobs:
        b = io.BytesIO(data)
        b.name = name
        files.append(b)
    return E.compute(E.prepare(E.load_exports(files)))


def comments_from_bytes(blobs):
    files = [io.BytesIO(d) for _, d in blobs]
    return X.read_comments(files) if files else {}


# ------------------------------------------------------------------ bandeau
st.markdown('<div class="rp-bar"><div><b>REVUE DE PERFORMANCE</b> <span class="l">COMMERCIALE</span></div>'
            '<div class="rp-meta">Vs N-1 · M FCFA HT · 12 magasins</div></div>', unsafe_allow_html=True)

tabs = st.tabs(['Accueil et lexique', 'Synthèse exécutive', 'Drill-down secteurs', 'Rapport PowerPoint'])

# ------------------------------------------------------------------ 1. accueil
with tabs[0]:
    c1, c2 = st.columns([1.25, 1])
    with c1:
        st.markdown('<div class="vh"><b>Importer les exports Power BI de la semaine</b></div>', unsafe_allow_html=True)
        up = st.file_uploader('Exports A (famille × site), B (rayon × site) et C (site), dans n’importe quel ordre',
                              type=['xlsx'], accept_multiple_files=True, key='exports')
        cdg = st.file_uploader('Classeurs des chefs de groupe complétés (facultatif, au gel de 10:10)', type=['xlsx'],
                               accept_multiple_files=True, key='cdg')
    with c2:
        st.markdown('<div class="vh"><b>Le lundi</b></div><div class="vb">'
                    '<b>08:00</b> le pôle BDD dépose les 3 exports<br><b>08:15</b> analyse et classeur des chefs de groupe sur Teams<br>'
                    '<b>10:10</b> gel des « Causes et actions »<br><b>10:30</b> envoi du rapport au directeur<br>'
                    '<b>14:00</b> point de détail et plans d’action</div>', unsafe_allow_html=True)

    R = None
    if up:
        try:
            R = compute_from_bytes(tuple((f.name, f.getvalue()) for f in up))
            R = E.merge_comments(dict(R), comments_from_bytes([(f.name, f.getvalue()) for f in (cdg or [])]))
            st.session_state['R'] = R
        except E.ExportError as ex:
            st.error(str(ex))
        except Exception as ex:  # noqa: BLE001
            st.error(f'Lecture impossible : {ex}')
    if R:
        m = R['meta']
        st.success(f"Semaine {m['week']} détectée · {m['periode']} · données au {m['data_date']:%d/%m/%Y} 08:00")
        st.markdown('<div class="pt">Contrôles de cohérence</div>', unsafe_allow_html=True)
        for lab, ok, det in R['checks']:
            st.markdown(f"{'✅' if ok else '⚠️'} **{lab}** · {det}")
        if not all(ok for _, ok, _ in R['checks']):
            st.warning('Un contrôle n’est pas passé : vérifier les exports avant de diffuser le rapport.')
    else:
        st.info('Déposez les 3 exports pour lancer l’analyse.')

    st.markdown('### Lexique')
    LEX = [
        ('Indicateurs', [("Chiffre d'affaires (CA)", 'Ventes de la semaine, hors taxes, en millions de FCFA.'),
                         ('Marge', 'CA moins le coût d’achat des produits vendus. La marge N-1 additionne la marge hors promo et la marge promo.'),
                         ('Taux de marge', 'Marge ÷ CA : ce que l’on gagne sur 100 FCFA vendus.'),
                         ('Trafic (tickets)', 'Nombre de passages en caisse (débit).'),
                         ('Panier moyen', 'CA ÷ tickets : ce que dépense un client par passage.'),
                         ('Articles par panier', 'Volume vendu ÷ tickets.')]),
        ('Écarts et décomposition', [('Vs N-1', 'Comparaison avec la même semaine de l’année précédente.'),
                                     ('Écart (Δ)', 'Valeur N moins valeur N-1.'), ('Point (pt)', 'Écart entre deux taux : 22,1 % contre 21,6 % = +0,5 pt.'),
                                     ('Effet volume', '(CA N − CA N-1) × taux N-1. Inclut aussi l’effet prix.'),
                                     ('Effet taux', 'CA N × (taux N − taux N-1).'),
                                     ('Lecture', 'Effet volume + effet taux = écart de marge.')]),
        ('Visuels', [('Passage de marge', 'Cascade : marge N-1, contribution de chaque secteur, marge N (appelé aussi « bridge »).'),
                     ('Heatmap', 'Tableau coloré secteur × format : rouge = recul, vert = progression, gris = moins de 0,5 M FCFA.'),
                     ('Top contributeurs / principaux détracteurs', 'Rayons ou magasins qui ajoutent ou retirent le plus de marge. Les deux listes ne s’additionnent pas.'),
                     ('Causes et actions', 'Explication et action saisies par le chef de groupe avant 10:10.')]),
        ('Périmètre et règles', [('Secteurs', 'PGC ; Frais ; Non-alimentaire (bazar, EPCS, textile).'),
                                 ('Formats', 'Hyper : Palmeraie, Marcory, Yopougon · Market : Kokoh Mall, Riviera, 2 Plateaux, 7 Décembre · Supeco : Niangon, Terminus 47.'),
                                 ('Ex-Supeco', 'Anciens Supeco convertis en Market : ' + ', '.join(f"{n.replace('Market ', '')} le {d}" for n, d in E.CONVERSIONS) + '.'),
                                 ('Familles retenues', 'CA d’au moins 1 M FCFA en N ou N-1, classées par écart de marge en valeur.')]),
    ]
    lc = st.columns(2)
    for i, (t, items) in enumerate(LEX):
        with lc[i % 2]:
            st.markdown(f'<div class="vh"><b>{t}</b></div><div class="vb"><dl class="lex">' +
                        ''.join(f'<dt>{k}</dt><dd>{v}</dd>' for k, v in items) + '</dl></div>', unsafe_allow_html=True)

R = st.session_state.get('R')


def need_data():
    st.info('Importez d’abord les exports dans l’onglet Accueil.')


def slicers(R, extra=None):
    m = R['meta']
    items = [('Période', f"S{m['week']} · {m['start']:%d/%m}–{m['end']:%d/%m/%Y}"), ('Comparaison', 'Vs N-1'),
             ('Périmètre', '12 magasins, tous ouverts en N-1'), ('Unité', 'M FCFA HT'), ('Données au', f"{m['data_date']:%d/%m/%Y} 08:00")]
    st.markdown('<div class="slc">' + ''.join(f'<div class="sl"><small>{a}</small><b>{b}</b></div>' for a, b in items) + '</div>', unsafe_allow_html=True)


# ------------------------------------------------------------------ 2. synthèse
with tabs[1]:
    if not R:
        need_data()
    else:
        g, t = R['global'], R['texts']
        slicers(R)
        verif = ', '.join(f"{k} ({v})" for k, v in E.SITES_A_VERIFIER.items())
        st.markdown(f"<div class='ctx'><b>Contexte :</b> {t['contexte']}. <b>Périmètre :</b> Ex-Supeco convertis en Market : "
                    + ', '.join(f"{n.replace('Market ', '')} le {d[:5]}" for n, d in E.CONVERSIONS)
                    + f". <b style='color:#8A4B00'>À vérifier :</b> {verif}. <b>Référence :</b> Vs N-1 uniquement, budget suivi au reporting mensuel.</div>",
                    unsafe_allow_html=True)
        K = [("Chiffre d'affaires", E.fm(g['CA']), 'M FCFA', f"{E.fm(g['CA1'])} M", g['ca'], E.fpct(g['ca'])),
             ('Marge', E.fm(g['M']), 'M FCFA', f"{E.fm(g['M1'])} M", g['dM'], f"{E.fdm(g['dM'])} M"),
             ('Taux de marge', E.fnum(g['t'] * 100), '%', f"{E.fnum(g['t1'] * 100)} %", g['dt'], E.fpt(g['dt'])),
             ('Trafic (tickets)', E.fnum(g['T'], 0), '', E.fnum(g['T1'], 0), g['tk'], E.fpct(g['tk'])),
             ('Panier moyen', E.fnum(g['panier'], 0), 'FCFA', f"{E.fnum(g['panier1'], 0)} FCFA", g['pan'], E.fpct(g['pan']))]
        st.markdown('<div class="kpis">' + ''.join(
            f'<div class="tile" style="border-bottom-color:{NEG if d < 0 else POS}"><small>{l}</small>'
            f'<div class="v">{v} <span style="font-size:14px;font-weight:400;color:#4E5370">{u}</span></div>'
            f'<div class="d"><span class="mut">N-1 : {n1}</span><b class="{cls(d)}">{"▼" if d < 0 else "▲"} {ds}</b></div></div>'
            for l, v, u, n1, d, ds in K) + '</div>', unsafe_allow_html=True)

        a, b = st.columns([1.45, 1])
        with a:
            st.markdown('<div class="vh"><span>Passage de la marge</span> <b>N-1 → N par secteur</b></div>', unsafe_allow_html=True)
            steps = [('Marge N-1', g['M1'], 'tot')] + [(s, R['sect'][s]['dM'], 'd') for s in E.SECTEUR_ORDER] + [('Marge N', g['M'], 'tot')]
            xs, base, h, col, txt = [], [], [], [], []
            cur = g['M1']
            for n, v, k in steps:
                xs.append(n)
                if k == 'tot':
                    base.append(0)
                    h.append(v / 1e6)
                    col.append('#758CC0' if n == 'Marge N-1' else NAVY)
                    txt.append(E.fm(v))
                else:
                    base.append(min(cur, cur + v) / 1e6)
                    h.append(abs(v) / 1e6)
                    col.append(NEG if v < 0 else POS)
                    txt.append(E.fdm(v))
                    cur += v
            lo = min(b_ for b_ in base if b_ > 0) if any(b_ > 0 for b_ in base) else 0
            ymin = int(min(lo, g['M1'] / 1e6, g['M'] / 1e6) * 0.95 // 5 * 5)
            fig = go.Figure(go.Bar(x=xs, y=h, base=base, marker_color=col, text=txt, textposition='outside',
                                   textfont=dict(size=14, color='#1E2140'), hoverinfo='skip'))
            fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10), plot_bgcolor='#fff', paper_bgcolor='#fff',
                              yaxis=dict(range=[ymin, max(base[i] + h[i] for i in range(len(h))) * 1.02], visible=False),
                              xaxis=dict(tickfont=dict(size=13)), showlegend=False, bargap=0.35)
            st.plotly_chart(fig, width='stretch', config={'displayModeBar': False})
            st.caption(f'≈ Axe tronqué à {ymin} M FCFA pour rendre les écarts visibles.')
        with b:
            items = ''.join(f'<li>{x}</li>' for x in t['faits'] + [t['taux']])
            st.markdown(f'<div class="vh"><b>Faits marquants</b></div><div class="vb"><ul class="fm">{items}</ul></div>', unsafe_allow_html=True)

        a, b = st.columns(2)
        with a:
            rows = ''
            for s in E.SECTEUR_ORDER + ['Total']:
                vals = [R['mtx'][s][f] for f in E.FORMAT_ORDER] + [R['sect'][s]['dM']] if s != 'Total' else [R['fmt'][f]['dM'] for f in E.FORMAT_ORDER] + [g['dM']]
                cells = ''.join('<td style="background:{0};color:{1};font-weight:700">{2}</td>'.format(*heat_css(v), E.fdm(v)) for v in vals)
                rows += f'<tr><td style="font-weight:{700 if s == "Total" else 400}">{s}</td>{cells}</tr>'
            st.markdown('<div class="vh"><span>Écart de marge</span> <b>secteur × format</b> <span style="font-size:12px">· gris = moins de 0,5 M</span></div><div class="vb">'
                        '<table class="ot"><thead><tr><th>Secteur</th>' + ''.join(f'<th>{f}</th>' for f in E.FORMAT_ORDER) +
                        '<th>Total</th></tr></thead><tbody>' + rows + '</tbody></table></div>', unsafe_allow_html=True)
        with b:
            C = R['contrib']
            vmax = max(abs(r['v']) for r in C['rayon'] + C['site']) * 1.05
            html = '<div class="vh"><span>Top contributeurs et</span> <b>principaux détracteurs</b> <span style="font-size:12px">· non additifs</span></div><div class="vb">'
            html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">'
            for k, lab in [('rayon', 'Par rayon'), ('site', 'Par magasin')]:
                html += f'<div><div class="pt">{lab}</div>' + ''.join(bar_row(r['n'], r['why'], r['v'], vmax) for r in C[k]) + '</div>'
            html += '</div></div>'
            st.markdown(html, unsafe_allow_html=True)

# ------------------------------------------------------------------ 3. secteurs
with tabs[2]:
    if not R:
        need_data()
    else:
        slicers(R)
        cols = st.columns(3)
        for col, s in zip(cols, E.SECTEUR_ORDER):
            o = R['sect_detail'][s]
            fam = R['familles'][s]
            with col:
                tag = ('tneg', 'Marge sous N-1') if o['dM'] < 0 else ('tpos', 'Marge au-dessus de N-1')
                html = f'<div class="card-h"><div><b>{s}</b><div style="font-size:12px;color:#C5C8E0">Responsable : {E.CHEFS[s]}</div></div><span class="tag {tag[0]}">{tag[1]}</span></div><div class="card-b">'
                html += '<div class="pt">Chiffre d’affaires et marge (Vs N-1)</div><table class="ot"><thead><tr><th>M FCFA</th><th>Réel</th><th>N-1</th><th>Δ</th><th>Δ %</th></tr></thead><tbody>'
                html += f"<tr><td>CA</td><td><b>{E.fm(o['CA'])}</b></td><td class='mut'>{E.fm(o['CA1'])}</td><td class='{cls(o['dCA'])}'><b>{E.fdm(o['dCA'])}</b></td><td class='{cls(o['dCA'])}'>{E.fpct(o['ca'])}</td></tr>"
                html += f"<tr><td>Marge</td><td><b>{E.fm(o['M'])}</b></td><td class='mut'>{E.fm(o['M1'])}</td><td class='{cls(o['dM'])}'><b>{E.fdm(o['dM'])}</b></td><td class='{cls(o['dM'])}'>{E.fpct(o['M'] / o['M1'] - 1)}</td></tr>"
                html += f"<tr><td>Taux</td><td><b>{E.fnum(o['t'] * 100)} %</b></td><td class='mut'>{E.fnum(o['t1'] * 100)} %</td><td class='{cls(o['dt'])}'><b>{E.fpt(o['dt'])}</b></td><td></td></tr></tbody></table>"
                html += f"<div class='pt'>Lecture analyste</div><div style='font-size:14px;line-height:1.45'>{R['texts']['lecture'][s]}</div>"
                fv = fam['fmt']
                vm = max(abs(v) for v in fv.values()) * 1.1 or 1
                html += '<div class="pt">Écart de marge par format (Vs N-1)</div>' + ''.join(bar_row(f, '', fv[f], vm) for f in E.FORMAT_ORDER)
                sel = fam['top'][:3] + fam['flop'][:3][::-1]
                vm = max([abs(x['dM']) for x in sel] + [1]) * 1.1
                html += '<div class="pt">Familles : contributeurs et détracteurs</div>'
                html += ''.join(bar_row(x['label'], f"CA {E.fm(x['CA'])} M · dont volume {E.fdm(x['effV'])} · dont taux {E.fdm(x['effT'])}", x['dM'], vm) for x in sel)
                gen = R.get('general', {}).get(s)
                lines = [x for x in fam['flop'][:3] if x.get('cdg', {}).get('Commentaire') or x.get('cdg', {}).get('Action')]
                body = ''
                if gen:
                    body += f'<p style="margin:0 0 6px">{gen}</p>'
                for x in lines:
                    c = x['cdg']
                    body += f"<p style='margin:0 0 6px'><b>{x['label']}</b> · {c.get('Cause', '')} · {c.get('Commentaire', '')} <i>Action : {c.get('Action', '—')}</i></p>"
                if not body:
                    body = f'<i class="mut">En attente : présenté en séance par {E.CHEFS[s]}.</i>'
                html += f'<div class="pt">Causes et actions (chef de groupe)</div><div class="cm">{body}</div></div>'
                st.markdown(html, unsafe_allow_html=True)

# ------------------------------------------------------------------ 4. rapport
with tabs[3]:
    if not R:
        need_data()
    else:
        sem = f"S{R['meta']['week']}"
        st.markdown(f'<div class="vh"><b>Rapport à la Direction Générale · 5 pages + annexe</b></div><div class="vb">'
                    '1 Couverture · 2 Synthèse exécutive · 3 Performance globale · 4 Contributeurs et détracteurs · 5 Décisions et plan d’actions'
                    '<br>Annexe : A1 PGC · A2 Frais · A3 Non-alimentaire · A4 Lexique et méthode</div>', unsafe_allow_html=True)
        comp = X.completeness(R)
        tot = sum(a for a, _ in comp.values())
        st.markdown(f'<div class="pt">Causes et actions saisies : {tot} sur {sum(b for _, b in comp.values())} familles prioritaires</div>', unsafe_allow_html=True)
        st.markdown(' · '.join(f"**{s}** {a}/{b}" for s, (a, b) in comp.items()) +
                    '  \nSans saisie, le rapport indique « En attente : présenté en séance par [chef de groupe] ». Le rapport part à 10:30 quoi qu’il arrive.')

        c1, c2, c3 = st.columns(3)
        with c1:
            buf = io.BytesIO()
            X.write_workbook(R, buf)
            st.download_button(f'Classeur des chefs de groupe ({sem})', buf.getvalue(), file_name=f'Causes_et_actions_{sem}.xlsx',
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', width='stretch')
            st.caption('À déposer sur Teams à 08:15.')
        with c2:
            masque = str(MASQUE) if MASQUE.exists() else None
            if masque is None:
                mup = st.file_uploader('Masque PowerPoint de l\u2019entreprise (.pptx)', type=['pptx'], key='masque',
                                       help='Le masque n\u2019est pas stocké dans le dépôt GitHub (public). Il est demandé à chaque session.')
                if mup:
                    masque = io.BytesIO(mup.getvalue())
            pbuf = io.BytesIO()
            if masque is None:
                st.info('Déposez le masque PowerPoint pour générer le rapport.')
                st.stop()
            P.build_report(R, masque, pbuf)
            st.download_button(f'Rapport PowerPoint ({sem})', pbuf.getvalue(), file_name=f'Revue_performance_{sem}.pptx',
                               mime='application/vnd.openxmlformats-officedocument.presentationml.presentation', type='primary', width='stretch')
            st.caption('Généré sur le masque de l’entreprise.')
        with c3:
            soffice = shutil.which('soffice') or shutil.which('libreoffice')
            if soffice:
                if st.button('Préparer la version PDF', width='stretch'):
                    with tempfile.TemporaryDirectory() as d:
                        src = Path(d) / f'Revue_performance_{sem}.pptx'
                        src.write_bytes(pbuf.getvalue())
                        subprocess.run([soffice, '--headless', '--convert-to', 'pdf', '--outdir', d, str(src)], capture_output=True, timeout=180)
                        pdf = src.with_suffix('.pdf')
                        if pdf.exists():
                            st.session_state['pdf'] = pdf.read_bytes()
                if st.session_state.get('pdf'):
                    st.download_button(f'Rapport PDF ({sem})', st.session_state['pdf'], file_name=f'Revue_performance_{sem}.pdf',
                                       mime='application/pdf', width='stretch')
            else:
                st.caption('Version PDF : ouvrir le PowerPoint puis Fichier › Exporter › PDF (LibreOffice absent de ce poste).')
