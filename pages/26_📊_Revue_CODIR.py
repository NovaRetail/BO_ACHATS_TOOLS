"""SmartBuyer Hub — Module 26 · Revue CODIR (rapport hebdomadaire à la Direction Générale).

Entrées : exports Power BI A (famille × site), B (rayon × site), C (site) ; classeurs « Causes et actions » complétés.
Sorties : classeur des chefs de groupe (valeurs, sans formule) et rapport PowerPoint 5 pages + annexe sur le masque.
Calculs : package modules/revue_codir/ (engine, excel_cdg, ppt_report). Charte : core/theme.py.
"""
from __future__ import annotations

import io
import shutil
import subprocess
import sys
import tempfile
import warnings
from pathlib import Path

import plotly.graph_objects as go
import streamlit as st

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from core import theme as T  # noqa: E402
from modules.revue_codir import engine as E  # noqa: E402
from modules.revue_codir import excel_cdg as X  # noqa: E402
from modules.revue_codir import ppt_report as P  # noqa: E402

warnings.filterwarnings('ignore')
MASQUE = ROOT / 'modules' / 'revue_codir' / 'assets' / 'masque_cfao_consumer.pptx'  # non versionné (dépôt public)

st.set_page_config(page_title='Revue CODIR', page_icon='📊', layout='wide')
T.apply()


# ------------------------------------------------------------------ calcul (mis en cache)
@st.cache_data(show_spinner='Analyse des exports…')
def compute_from_bytes(blobs: tuple):
    files = []
    for name, data in blobs:
        b = io.BytesIO(data)
        b.name = name
        files.append(b)
    return E.compute(E.prepare(E.load_exports(files)))


def sgn(v):
    return 'neg' if v < 0 else 'pos'


# ------------------------------------------------------------------ barre latérale
T.sidebar_brand('📊', 'Revue CODIR')
T.sidebar_label('Import fichiers')
with st.sidebar:
    up = st.file_uploader('Exports Power BI A, B et C', type=['xlsx'], accept_multiple_files=True, key='exports',
                          help='A : famille × site · B : rayon × site · C : site. Ordre libre.')
    cdg = st.file_uploader('Causes et actions complétées (10:10)', type=['xlsx'], accept_multiple_files=True, key='cdg',
                           help='Classeurs des chefs de groupe, facultatif.')
    mup = None
    if not MASQUE.exists():
        mup = st.file_uploader('Masque PowerPoint (.pptx)', type=['pptx'], key='masque',
                               help='Non stocké dans le dépôt GitHub : à déposer à chaque session.')
T.sidebar_footer('Vs N-1 · M FCFA HT · 12 magasins<br>Casse et budget hors périmètre')

# ------------------------------------------------------------------ page d'accueil (avant import)
if not up:
    T.page_header('📊', 'Revue CODIR', 'Rapport hebdomadaire à la Direction Générale · entonnoir Vs N-1, de la semaine jusqu’aux familles')
    T.landing_hero('Le rapport du lundi, prêt pour 10:30.',
                   'Dépose les 3 exports Power BI de la semaine dans la barre latérale : le module contrôle leur cohérence, '
                   'calcule l’entonnoir Global → Secteur → Format → Famille, prépare le classeur des chefs de groupe '
                   'et génère le rapport PowerPoint sur le masque de l’entreprise.')
    c1, c2 = st.columns(2)
    with c1:
        T.rules_card('🧭 Ce que contient le module', [
            ('Synthèse exécutive', 'KPI, passage de marge, faits marquants'),
            ('Secteurs', 'PGC, Frais, Non-alimentaire : formats et familles'),
            ('Rapport et exports', 'PowerPoint 5 pages + annexe, classeur, PDF'),
            ('Lexique', 'définitions, formules, périmètre'),
        ])
    with c2:
        T.rules_card('⚙️ Le lundi', [
            ('08:00 · Pôle BDD', 'dépose les 3 exports'),
            ('08:15 · Pôle BDD', 'classeur des chefs de groupe sur Teams'),
            ('10:10 · Chefs de groupe', 'gel des « Causes et actions »'),
            ('10:30 · Coordinateur', 'envoi du rapport au directeur'),
            ('14:00 · Équipe', 'point de détail et plans d’action'),
        ])
    st.write('')
    c3, c4 = st.columns(2)
    with c3:
        T.rules_card('🛡️ Contrôles automatiques', [
            ('CA export A = CA export B', 'écart toléré 0,1 %'),
            ('Tickets des sites = réseau', 'export C complet'),
            ('12 magasins', 'dans chaque export'),
            ('Même semaine', 'lue dans « Filtres appliqués »'),
            ('Familles retenues', 'CA ≥ 1 M FCFA en N ou N-1'),
            ('Écart non significatif', '&lt; 0,5 M FCFA, affiché en gris'),
        ])
    with c4:
        T.columns_card('📄 Colonnes attendues', [
            ('Export A · famille × site', ['Departement', 'Rayon', 'Famille', 'Site nom long', 'CA', 'CA N-1', 'Marge',
                                           'Marge Hors Promo N-1', 'Marge Promo N-1', 'CA Promo', 'Qté Vente']),
            ('Export B · rayon × site', ['Département', 'Site', 'Rayon', 'CA', 'Débit', 'Débit N-1', 'Volume', 'Volume N-1']),
            ('Export C · site', ['Site', 'Débit', 'Débit N-1']),
        ], note='Les colonnes N-1 sont lues automatiquement. Le Budget, s’il est présent, est ignoré.')
    st.info('Dépose les 3 exports Power BI dans la barre latérale pour démarrer.')
    st.stop()

# ------------------------------------------------------------------ données
try:
    R = compute_from_bytes(tuple((f.name, f.getvalue()) for f in up))
except E.ExportError as ex:
    T.page_header('📊', 'Revue CODIR')
    T.alert('Exports incomplets ou non reconnus', str(ex), 'crit')
    st.stop()
comments = X.read_comments([io.BytesIO(f.getvalue()) for f in (cdg or [])]) if cdg else {}
R = E.merge_comments(dict(R), comments)
g, t, m = R['global'], R['texts'], R['meta']
sem = f"S{m['week']}"
checks_ok = all(ok for _, ok, _ in R['checks'])

T.hero(f'📊 Revue CODIR · {sem}', f"{m['periode']} · Vs N-1 · données au {m['data_date']:%d/%m/%Y} 08:00", [
    f"CA {E.fm(g['CA'])} M · {E.fpct(g['ca'])}",
    (f"Marge {E.fdm(g['dM'])} M", 'alert' if g['dM'] < 0 else 'ok'),
    f"Taux {E.fnum(g['t'] * 100)} % · {E.fpt(g['dt'])}",
    f"Trafic {E.fpct(g['tk'])}",
    (f"Panier {E.fpct(g['pan'])}", 'alert' if g['pan'] < 0 else 'ok'),
    ('Contrôles OK', 'ok') if checks_ok else ('Contrôle à vérifier', 'alert'),
])
if not checks_ok:
    T.alert('Un contrôle de cohérence n’est pas passé', ' · '.join(f'{lab} : {det}' for lab, ok, det in R['checks'] if not ok), 'crit')

tabs = st.tabs(['Synthèse exécutive', 'Secteurs', 'Rapport et exports', 'Lexique'])

# ------------------------------------------------------------------ 1. synthèse
with tabs[0]:
    st.write('')
    T.kpi_row([
        T.kpi("Chiffre d'affaires", E.fm(g['CA']), E.fpct(g['ca']), sgn(g['ca']), f"N-1 {E.fm(g['CA1'])} M", 'blue', '💰', 'M FCFA'),
        T.kpi('Marge', E.fm(g['M']), f"{E.fdm(g['dM'])} M", sgn(g['dM']), f"N-1 {E.fm(g['M1'])} M", 'violet', '📈', 'M FCFA'),
        T.kpi('Taux de marge', E.fnum(g['t'] * 100), E.fpt(g['dt']), sgn(g['dt']), f"N-1 {E.fnum(g['t1'] * 100)} %", 'violet', '％', '%'),
        T.kpi('Trafic', E.fnum(g['T'], 0), E.fpct(g['tk']), sgn(g['tk']), f"N-1 {E.fnum(g['T1'], 0)}", '', '🧾', 'tickets'),
        T.kpi('Panier moyen', E.fnum(g['panier'], 0), E.fpct(g['pan']), sgn(g['pan']), f"N-1 {E.fnum(g['panier1'], 0)}", '', '🛒', 'FCFA'),
    ])
    verif = ', '.join(f'{k} ({v})' for k, v in E.SITES_A_VERIFIER.items())
    T.alert('Contexte et périmètre',
            f"{t['contexte'][:1].upper() + t['contexte'][1:]}. Ex-Supeco convertis en Market : "
            + ', '.join(f"{n.replace('Market ', '')} le {d[:5]}" for n, d in E.CONVERSIONS)
            + f". À vérifier : {verif}. Référence : Vs N-1 uniquement, budget suivi au reporting mensuel.", 'info')

    a, b = st.columns([1.45, 1])
    with a:
        T.section('Passage de la marge N-1 → N par secteur')
        steps = [('Marge N-1', g['M1'], 'tot')] + [(s, R['sect'][s]['dM'], 'd') for s in E.SECTEUR_ORDER] + [('Marge N', g['M'], 'tot')]
        xs, base, h, col, txt = [], [], [], [], []
        cur = g['M1']
        for n, v, k in steps:
            xs.append(n)
            if k == 'tot':
                base.append(0)
                h.append(v / 1e6)
                col.append('#8FB8F5' if n == 'Marge N-1' else T.NAVY)
                txt.append(E.fm(v))
            else:
                base.append(min(cur, cur + v) / 1e6)
                h.append(abs(v) / 1e6)
                col.append(T.ROUGE if v < 0 else T.VERT)
                txt.append(E.fdm(v))
                cur += v
        pos_b = [b_ for b_ in base if b_ > 0]
        ymin = int(min(pos_b + [g['M1'] / 1e6, g['M'] / 1e6]) * 0.95 // 5 * 5)
        fig = go.Figure(go.Bar(x=xs, y=h, base=base, marker=dict(color=col, cornerradius=8), text=txt, textposition='outside',
                               textfont=dict(size=14, color=T.NOIR), hoverinfo='skip'))
        T.plotly_layout(fig, 330)
        fig.update_layout(yaxis=dict(range=[ymin, max(base[i] + h[i] for i in range(len(h))) * 1.02], visible=False),
                          xaxis=dict(tickfont=dict(size=13, color=T.GRIS_TEXTE)), bargap=0.38)
        with st.container(border=False):
            st.plotly_chart(fig, width='stretch', config={'displayModeBar': False})
        st.caption(f'≈ Axe tronqué à {ymin} M FCFA pour rendre les écarts visibles.')
    with b:
        T.section('Faits marquants')
        T.card('<ul class="bul">' + ''.join(f'<li>{x}</li>' for x in t['faits'] + [t['taux']]) + '</ul>')

    a, b = st.columns(2)
    with a:
        T.section('Écart de marge · secteur × format (M FCFA)')
        rows = ''
        for s in E.SECTEUR_ORDER + ['Total']:
            vals = ([R['mtx'][s][f] for f in E.FORMAT_ORDER] + [R['sect'][s]['dM']]) if s != 'Total' else ([R['fmt'][f]['dM'] for f in E.FORMAT_ORDER] + [g['dM']])
            cells = ''
            for v in vals:
                bg, fg = T.heat(v, 5e6, abs(v) < E.SEUIL_SIGNIF)
                cells += f"<td class='hm' style='background:{bg};color:{fg}'>{E.fdm(v)}</td>"
            rows += f"<tr><td style='font-weight:{900 if s == 'Total' else 700}'>{s}</td>{cells}</tr>"
        T.card("<table class='tbl'><thead><tr><th>Secteur</th>" + ''.join(f'<th>{f}</th>' for f in E.FORMAT_ORDER)
               + '<th>Total</th></tr></thead><tbody>' + rows + "</tbody></table><div class='note'>Gris : écart inférieur à 0,5 M FCFA.</div>")
    with b:
        T.section('Top contributeurs et principaux détracteurs')
        C = R['contrib']
        vmax = max(abs(r['v']) for r in C['rayon'] + C['site']) * 1.05
        inner = "<div style='display:grid;grid-template-columns:1fr 1fr;gap:18px'>"
        for k, lab in [('rayon', 'Par rayon'), ('site', 'Par magasin')]:
            inner += f"<div><div class='kpi-label' style='margin-bottom:4px'>{lab}</div>" + ''.join(
                T.dv_row(r['n'], r['why'], r['v'], vmax, E.fdm(r['v']), abs(r['v']) < E.SEUIL_SIGNIF) for r in C[k]) + '</div>'
        T.card(inner + "</div><div class='note'>Les deux listes ne s’additionnent pas : un rayon et un magasin peuvent porter la même perte.</div>")

# ------------------------------------------------------------------ 2. secteurs
with tabs[1]:
    st.write('')
    cols = st.columns(3)
    for col, s in zip(cols, E.SECTEUR_ORDER):
        o, fam = R['sect_detail'][s], R['familles'][s]
        with col:
            status = T.badge('Marge sous N-1', 'crit') if o['dM'] < 0 else T.badge('Marge au-dessus de N-1', 'ok')
            T.html(f"<div class='card'><div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:4px'>"
                   f"<span class='section-title' style='margin:0;font-size:18px'>{s}</span>{status}</div>"
                   f"<div class='mut' style='font-size:12.5px'>Responsable : {E.CHEFS[s]}</div></div>")
            T.kpi_row([
                T.kpi('CA', E.fm(o['CA']), E.fpct(o['ca']), sgn(o['ca']), '', 'blue', '', 'M'),
                T.kpi('Marge', E.fdm(o['dM']), E.fpt(o['dt']), sgn(o['dt']), f"taux {E.fnum(o['t'] * 100)} %", 'violet', '', 'M'),
            ])
            T.card(f"<div style='font-size:13.5px;line-height:1.5;font-weight:600'>{t['lecture'][s]}</div>", 'Lecture analyste')
            fv = fam['fmt']
            vm = max(abs(v) for v in fv.values()) * 1.1 or 1
            T.card(''.join(T.dv_row(f, '', fv[f], vm, E.fdm(fv[f]), abs(fv[f]) < E.SEUIL_SIGNIF) for f in E.FORMAT_ORDER),
                   'Écart de marge par format (Vs N-1)')
            sel = fam['top'][:3] + fam['flop'][:3][::-1]
            vm = max([abs(x['dM']) for x in sel] + [1]) * 1.1
            T.card(''.join(T.dv_row(x['label'], f"volume {E.fdm(x['effV'])} · taux {E.fdm(x['effT'])}", x['dM'], vm, E.fdm(x['dM']),
                                    abs(x['dM']) < E.SEUIL_SIGNIF) for x in sel), 'Familles · dont volume / dont taux')
            gen = R.get('general', {}).get(s)
            lines = [x for x in fam['flop'][:3] if x.get('cdg', {}).get('Commentaire') or x.get('cdg', {}).get('Action')]
            if gen or lines:
                body = (f'{gen}<br>' if gen else '') + '<br>'.join(
                    f"<b>{x['label']}</b> · {x['cdg'].get('Cause', '')} · {x['cdg'].get('Commentaire', '')} · Action : {x['cdg'].get('Action', '—')}" for x in lines)
                T.alert('Causes et actions', body, 'ok')
            else:
                T.alert('Causes et actions', f'En attente : présenté en séance par {E.CHEFS[s]}.', 'att')

# ------------------------------------------------------------------ 3. rapport et exports
with tabs[2]:
    st.write('')
    comp = X.completeness(R)
    done, total = sum(a for a, _ in comp.values()), sum(b for _, b in comp.values())
    T.kpi_row([
        T.kpi('Causes et actions saisies', f'{done} / {total}', 'familles prioritaires', 'pos' if done == total else 'mut', 'gel à 10:10', 'blue', '✍️'),
        *[T.kpi(s, f'{a} / {b}', 'complet' if a == b else 'à compléter', 'pos' if a == b else 'neg', E.CHEFS[s], '', '') for s, (a, b) in comp.items()],
    ])
    c1, c2 = st.columns([1, 1])
    with c1:
        T.rules_card('📑 Rapport à la Direction Générale', [
            ('1 · Couverture', 'semaine, période, date des données'),
            ('2 · Synthèse exécutive', 'où en est-on, pourquoi, que fait-on'),
            ('3 · Performance globale', 'passage de marge, secteur × format'),
            ('4 · Contributeurs et détracteurs', 'par rayon, par magasin'),
            ('5 · Décisions et plan d’actions', 'repris des chefs de groupe'),
            ('Annexe A1 à A4', 'PGC, Frais, Non-alimentaire, lexique'),
        ], tall=False)
    with c2:
        T.section('Téléchargements')
        buf = io.BytesIO()
        X.write_workbook(R, buf)
        st.download_button(f'⬇️ Classeur des chefs de groupe ({sem})', buf.getvalue(), file_name=f'Causes_et_actions_{sem}.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', width='stretch')
        st.caption('À déposer sur Teams à 08:15. Valeurs uniquement, aucune formule.')
        masque = str(MASQUE) if MASQUE.exists() else (io.BytesIO(mup.getvalue()) if mup else None)
        if masque is None:
            T.alert('Masque PowerPoint manquant', 'Dépose le masque .pptx dans la barre latérale pour générer le rapport.', 'att')
        else:
            pbuf = io.BytesIO()
            P.build_report(R, masque, pbuf)
            st.download_button(f'⬇️ Rapport PowerPoint ({sem})', pbuf.getvalue(), file_name=f'Revue_performance_{sem}.pptx',
                               mime='application/vnd.openxmlformats-officedocument.presentationml.presentation', width='stretch')
            st.caption('Le rapport part à 10:30 quoi qu’il arrive ; sans saisie, il indique « En attente : présenté en séance par [chef de groupe] ».')
            soffice = shutil.which('soffice') or shutil.which('libreoffice')
            if soffice and st.button('Préparer la version PDF', width='stretch'):
                with tempfile.TemporaryDirectory() as d:
                    src = Path(d) / f'Revue_performance_{sem}.pptx'
                    src.write_bytes(pbuf.getvalue())
                    subprocess.run([soffice, '--headless', '--convert-to', 'pdf', '--outdir', d, str(src)], capture_output=True, timeout=180)
                    if src.with_suffix('.pdf').exists():
                        st.session_state['pdf'] = src.with_suffix('.pdf').read_bytes()
            if st.session_state.get('pdf'):
                st.download_button(f'⬇️ Rapport PDF ({sem})', st.session_state['pdf'], file_name=f'Revue_performance_{sem}.pdf',
                                   mime='application/pdf', width='stretch')
            if not soffice:
                st.caption('Version PDF : ouvrir le PowerPoint puis Fichier › Exporter › PDF.')
    T.section('Contrôles de cohérence')
    for lab, ok, det in R['checks']:
        T.alert(lab, det, 'ok' if ok else 'crit')

# ------------------------------------------------------------------ 4. lexique
with tabs[3]:
    st.write('')
    LEX = [
        ('📏 Indicateurs', [("Chiffre d'affaires (CA)", 'ventes HT de la semaine, en M FCFA'),
                           ('Marge', 'CA − coût d’achat ; marge N-1 = hors promo + promo'),
                           ('Taux de marge', 'marge ÷ CA'), ('Trafic', 'nombre de tickets de caisse'),
                           ('Panier moyen', 'CA ÷ tickets, en FCFA'), ('Articles par panier', 'volume ÷ tickets')]),
        ('🧮 Écarts et décomposition', [('Vs N-1', 'même semaine de l’année précédente'), ('Écart (Δ)', 'valeur N − valeur N-1'),
                                       ('Point (pt)', 'écart entre deux taux'), ('Effet volume', '(CA N − CA N-1) × taux N-1, inclut l’effet prix'),
                                       ('Effet taux', 'CA N × (taux N − taux N-1)'), ('Lecture', 'effet volume + effet taux = écart de marge')]),
        ('📊 Visuels', [('Passage de marge', 'marge N-1, contribution de chaque secteur, marge N'),
                       ('Heatmap', 'rouge = recul, vert = gain, gris = moins de 0,5 M'),
                       ('Contributeurs / détracteurs', 'rayons ou magasins qui ajoutent ou retirent le plus de marge'),
                       ('Causes et actions', 'saisies par les chefs de groupe avant 10:10')]),
        ('🏬 Périmètre', [('Secteurs', 'PGC ; Frais ; Non-alimentaire (bazar, EPCS, textile)'),
                         ('Hyper', 'Palmeraie, Marcory, Yopougon'), ('Market', 'Kokoh Mall, Riviera, 2 Plateaux, 7 Décembre'),
                         ('Ex-Supeco', ', '.join(f"{n.replace('Market ', '')} ({d})" for n, d in E.CONVERSIONS)),
                         ('Supeco', 'Niangon, Terminus 47'), ('Familles retenues', 'CA ≥ 1 M FCFA en N ou N-1')]),
    ]
    c = st.columns(2)
    for i, (title, rows) in enumerate(LEX):
        with c[i % 2]:
            T.rules_card(title, rows, tall=False)
