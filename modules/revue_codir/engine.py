"""Moteur de calcul de la revue de performance hebdomadaire.

Entrées : les 3 exports Power BI de la semaine
  A  famille × site (CA, marge, promo, quantités, avec N-1)
  B  département × site × rayon (débit, volume, avec N-1)
  C  site (débit par magasin, avec N-1)
Sortie : un dictionnaire `R` qui alimente l'application, le PowerPoint et le classeur Excel.
Tous les calculs sont faits ici, une seule fois (source unique de vérité).
"""
from __future__ import annotations

import io
import re
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd

# ------------------------------------------------------------------ paramètres
FORMATS = {  # code site -> format
    '10202': 'Hyper', '10301': 'Hyper', '10203': 'Hyper',
    '10206': 'Market', '10208': 'Market', '10209': 'Market', '10705': 'Market',
    '10604': 'Ex-Supeco', '10605': 'Ex-Supeco', '10603': 'Ex-Supeco',
    '10601': 'Supeco', '10602': 'Supeco',
}
FORMAT_ORDER = ['Hyper', 'Market', 'Ex-Supeco', 'Supeco']
SECTEURS = {'01 - PGC': 'PGC', '02 - PRODUITS FRAIS': 'Frais', '03 - BAZAR': 'Non-alimentaire',
            '04 - EPCS': 'Non-alimentaire', '06 - TEXTILE': 'Non-alimentaire'}
SECTEUR_ORDER = ['PGC', 'Frais', 'Non-alimentaire']
CHEFS = {'PGC': 'Chef de groupe 1', 'Frais': 'Chef de groupe 2', 'Non-alimentaire': 'Chef de groupe 1'}
CONVERSIONS = [('Market Cité Verte', '01/04/2026'), ('Market Aboboté', '26/08/2026'), ('Market Toit Rouge', '22/10/2026')]
SEUIL_FAMILLE = 1e6        # CA N ou N-1 minimum pour entrer dans le Top/Flop (FCFA)
SEUIL_SIGNIF = 0.5e6       # écart de marge en dessous duquel l'écart est « non significatif » (FCFA)
TOLERANCE = 0.001          # écart toléré entre les totaux des exports (0,1 %)
SITES_A_VERIFIER = {'Market Toit Rouge': 'converti le 22/10'}

# ------------------------------------------------------------------ libellés
_WORDS = {'pate': 'pâte', 'feculent': 'féculent', 'patisserie': 'pâtisserie', 'cremerie': 'crèmerie', 'surgeles': 'surgelés',
          'legumes': 'légumes', 'electro': 'électro', 'menager': 'ménager', 'hygiene': 'hygiène', 'cafeteria': 'cafétéria',
          'aperitif': 'apéritif', 'sante': 'santé', 'bebe': 'bébé', 'epicerie': 'épicerie', 'decoration': 'décoration',
          'securite': 'sécurité', 'telephonie': 'téléphonie', 'equipement': 'équipement', 'detente': 'détente',
          'puericulture': 'puériculture', 'fete': 'fête', 'fetes': 'fêtes', 'noel': 'Noël', 'specialite': 'spécialité',
          'rotisserie': 'rôtisserie', 'qualite': 'qualité', 'preparation': 'préparation', 'seche': 'sèche', 'ecran': 'écran',
          'creme': 'crème', 'cuisines': 'cuisinés', 'vetement': 'vêtement', 'vetements': 'vêtements', 'a': 'à',
          'dejeuner': 'déjeuner', 'sucre': 'sucré', 'randonnee': 'randonnée', 'radiotelephonie': 'radiotéléphonie',
          'cafetaria': 'cafétaria', 'boissons': 'boissons', 'pieces': 'pièces', 'detachees': 'détachées'}
_FIX = {'PRODUITS FRAIS L.S.': 'Produits frais LS', 'BOISSON': 'Boissons', 'GEM': 'GEM'}


def label(s: str) -> str:
    """Libellé lisible : minuscules, accents rétablis, majuscule initiale."""
    s = str(s).strip()
    if s.upper() in _FIX:
        return _FIX[s.upper()]
    toks = re.split(r'(\W+)', s.lower())
    out = ''.join(_WORDS.get(t, t) for t in toks)
    out = re.sub(r'\s+', ' ', out).replace(' - ', '-')
    return out[:1].upper() + out[1:]


def _strip_code(s) -> str:
    return re.sub(r'^\d+\s*-\s*', '', str(s)).strip()


def site_label(s) -> str:
    t = _strip_code(s)
    return ' '.join(w if w.isupper() and len(w) <= 3 else (w[:1].upper() + w[1:]) for w in t.split(' '))


# ------------------------------------------------------------------ formats d'affichage (FR)
def fnum(v, d=1):
    return f'{abs(v):,.{d}f}'.replace(',', ' ').replace('.', ',')


def fsg(v, d=1):
    return ('+' if v >= 0 else '−') + fnum(v, d)


def fpct(v, d=1):
    return 'n.d.' if v is None or (isinstance(v, float) and not np.isfinite(v)) else fsg(v * 100, d) + ' %'


def fm(v, d=1):          # FCFA -> « 947,0 »
    return fnum(v / 1e6, d)


def fdm(v, d=1):         # écart FCFA -> « −3,4 »
    return fsg(v / 1e6, d)


def fpt(v, d=1):         # points de taux
    return fsg(v, d) + ' pt'


# ------------------------------------------------------------------ lecture des exports
class ExportError(Exception):
    pass


def _read(src) -> pd.DataFrame:
    if hasattr(src, 'read'):
        data = src.read()
        src = io.BytesIO(data)
    return pd.read_excel(src)


def _footer(df: pd.DataFrame) -> str:
    for v in df.iloc[::-1, 0].astype(str):
        if 'Filtres appliqu' in v:
            return v
    return ''


def identify(df: pd.DataFrame) -> str:
    cols = set(df.columns)
    if {'Famille', 'Site nom long', 'Marge Promo N-1'} <= cols:
        return 'A'
    if {'Département', 'Site', 'Rayon', 'Débit'} <= cols:
        return 'B'
    if {'Site', 'Débit', 'Débit N-1'} <= cols and 'Rayon' not in cols:
        return 'C'
    return '?'


REQUIRED = {
    'A': ['Departement', 'Rayon', 'Famille', 'Site nom long', 'CA', 'CA N-1', 'Marge', 'Marge Hors Promo N-1', 'Marge Promo N-1',
          'CA Promo', 'CA Promo N-1', 'Qté Vente', 'Qté Vente N-1'],
    'B': ['Département', 'Site', 'Rayon', 'CA', 'Débit', 'Débit N-1', 'Volume', 'Volume N-1'],
    'C': ['Site', 'Débit', 'Débit N-1'],
}
LABELS = {'A': 'Export A · Famille × Magasin', 'B': 'Export B · Rayon × Magasin', 'C': 'Export C · Magasin'}


def inspect_export(src) -> dict:
    """Diagnostic d'un fichier déposé : type reconnu, colonnes manquantes, période, lignes Total."""
    df = _read(src)
    kind = identify(df)
    per = period_from_footer(_footer(df))
    out = {'kind': kind, 'rows': len(df), 'period': per, 'has_footer': bool(_footer(df))}
    out['missing'] = {k: [c for c in cols if c not in df.columns] for k, cols in REQUIRED.items()}
    first = df.columns[0]
    out['has_total'] = bool((df[first].astype(str) == 'Total').any())
    return out


def period_from_footer(txt: str):
    dates = [datetime.strptime(d, '%d/%m/%Y').date() for d in re.findall(r'\d{2}/\d{2}/\d{4}', txt)]
    if not dates:
        return None
    start, end = min(dates), max(dates)
    if re.search(r'est avant le\s+' + end.strftime('%d/%m/%Y'), txt):
        end = end - timedelta(days=1)
    return start, end


def load_exports(files) -> dict:
    """files : liste de chemins ou de fichiers téléversés (ordre libre)."""
    found, errors = {}, []
    for f in files:
        df = _read(f)
        kind = identify(df)
        name = getattr(f, 'name', str(f))
        if kind == '?':
            errors.append(f'Fichier non reconnu : {name}')
            continue
        found[kind] = (df, name)
    for k, lib in [('A', 'A · famille × site'), ('B', 'B · département × site × rayon'), ('C', 'C · site')]:
        if k not in found:
            errors.append(f'Export manquant : {lib}')
    if errors:
        raise ExportError(' ; '.join(errors))
    return found


# ------------------------------------------------------------------ préparation
def prepare(found: dict) -> dict:
    a, _ = found['A']
    b, _ = found['B']
    c, _ = found['C']
    periods = {k: period_from_footer(_footer(found[k][0])) for k in 'ABC'}

    # --- A : famille × site
    a = a[a['Departement'].isin(SECTEURS)]
    a = a[a['Site nom long'].notna() & (a['Site nom long'] != 'Total') & a['Famille'].notna()
          & (a['Famille'] != 'Total') & (a['Rayon'] != 'Total')].copy()
    code = a['Site nom long'].astype(str).str[:5]
    fam = pd.DataFrame({
        'Secteur': a['Departement'].map(SECTEURS), 'Format': code.map(FORMATS), 'Site': a['Site nom long'].map(site_label),
        'Rayon': a['Rayon'].map(_strip_code), 'Famille': a['Famille'].map(_strip_code),
        'CA': a['CA'], 'CA1': a['CA N-1'], 'M': a['Marge'], 'M1': a['Marge Hors Promo N-1'].fillna(0) + a['Marge Promo N-1'].fillna(0),
        'Q': a['Qté Vente'], 'Q1': a['Qté Vente N-1'], 'CAP': a['CA Promo'], 'CAP1': a['CA Promo N-1'],
    }).fillna(0).reset_index(drop=True)
    unknown = sorted(set(code[fam['Format'].isna().values]))
    fam = fam[fam['Format'].notna()]

    # --- B : département × site × rayon
    b = b[b.iloc[:, 0].astype(str).str.contains('Filtres') == False]  # noqa: E712
    gt = b[b['Département'] == 'Total'].iloc[0]
    bb = b[b['Département'].isin(SECTEURS)].copy()
    bb['Secteur'] = bb['Département'].map(SECTEURS)
    ds = bb[(bb['Rayon'] == 'Total') & bb['Site'].notna() & (bb['Site'] != 'Total')]
    ds = pd.DataFrame({'Secteur': ds['Secteur'], 'Site': ds['Site'].map(site_label), 'Format': ds['Site'].astype(str).str[:5].map(FORMATS),
                       'V': ds['Volume'], 'V1': ds['Volume N-1'], 'CA': ds['CA']}).fillna(0)
    ray = bb[bb['Rayon'].notna() & (bb['Rayon'] != 'Total') & bb['Site'].notna() & (bb['Site'] != 'Total')]
    ray = pd.DataFrame({'Secteur': ray['Secteur'], 'Rayon': ray['Rayon'].map(_strip_code).replace({'BOISSON': 'BOISSONS'}),
                        'Site': ray['Site'].map(site_label), 'T': ray['Débit'], 'T1': ray['Débit N-1'],
                        'V': ray['Volume'], 'V1': ray['Volume N-1']}).fillna(0)

    # --- C : site
    cc = c[c['Site'].notna() & ~c['Site'].astype(str).str.contains('Filtres')]
    ctot = cc[cc['Site'] == 'Total']
    cc = cc[cc['Site'] != 'Total']
    st = pd.DataFrame({'Site': cc['Site'].map(site_label), 'Format': cc['Site'].astype(str).str[:5].map(FORMATS),
                       'T': cc['Débit'], 'T1': cc['Débit N-1']}).fillna(0)

    T, T1 = float(gt['Débit']), float(gt['Débit N-1'])
    VOL, VOL1 = float(gt['Volume']), float(gt['Volume N-1'])

    # --- contrôles
    checks = []
    caA, caB = fam['CA'].sum(), ds['CA'].sum()
    ecart = abs(caA - caB) / caB if caB else 1
    checks.append(('CA export A = CA export B', ecart <= TOLERANCE, f'{fm(caA)} M contre {fm(caB)} M (écart {fnum(ecart * 100, 3)} %)'))
    tks = st['T'].sum()
    checks.append(('Tickets des sites = tickets du réseau', abs(tks - T) <= max(1, T * TOLERANCE), f'{fnum(tks, 0)} contre {fnum(T, 0)}'))
    checks.append(('12 magasins dans chaque export', fam['Site'].nunique() == 12 and st['Site'].nunique() == 12,
                   f'A : {fam["Site"].nunique()} · C : {st["Site"].nunique()}'))
    ps = {k: v for k, v in periods.items() if v}
    same = len(set(ps.values())) == 1 and len(ps) == 3
    checks.append(('Même période dans les 3 exports', same,
                   ' · '.join(f'{k} : {v[0]:%d/%m}–{v[1]:%d/%m}' for k, v in ps.items()) or 'période introuvable'))
    if unknown:
        checks.append(('Codes magasin connus', False, 'inconnus : ' + ', '.join(unknown)))

    per = ps.get('A') or ps.get('B') or ps.get('C')
    if per is None:
        raise ExportError('Période introuvable dans le pied de page « Filtres appliqués » des exports.')
    start, end = per
    iso = start.isocalendar()
    monday_after = end + timedelta(days=(7 - end.weekday()) % 7 or 7)
    meta = {'start': start, 'end': end, 'week': iso[1], 'year': iso[0], 'data_date': monday_after,
            'periode': f'du {start:%d} au {end:%d} {_MOIS[end.month]} {end.year}' if start.month == end.month
            else f'du {start:%d/%m} au {end:%d/%m/%Y}'}
    return {'fam': fam, 'ray': ray, 'ds': ds, 'st': st, 'T': T, 'T1': T1, 'VOL': VOL, 'VOL1': VOL1,
            'checks': checks, 'meta': meta}


_MOIS = ['', 'janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']


# ------------------------------------------------------------------ calculs
def kpi(d: pd.DataFrame) -> dict:
    s = d[['CA', 'CA1', 'M', 'M1', 'Q', 'Q1', 'CAP', 'CAP1']].sum()
    o = {k: float(v) for k, v in s.items()}
    o['dCA'] = o['CA'] - o['CA1']
    o['dM'] = o['M'] - o['M1']
    o['ca'] = o['CA'] / o['CA1'] - 1 if o['CA1'] else np.nan
    o['t'] = o['M'] / o['CA'] if o['CA'] else np.nan
    o['t1'] = o['M1'] / o['CA1'] if o['CA1'] else np.nan
    o['dt'] = (o['t'] - o['t1']) * 100 if o['CA'] and o['CA1'] else np.nan
    o['effV'] = o['dCA'] * (o['t1'] if o['CA1'] else 0)
    o['effT'] = o['dM'] - o['effV']
    o['q'] = o['Q'] / o['Q1'] - 1 if o['Q1'] else np.nan
    o['px'] = (o['CA'] / o['Q']) / (o['CA1'] / o['Q1']) - 1 if o['Q'] and o['Q1'] else np.nan
    o['promo'] = o['CAP'] / o['CA'] if o['CA'] else np.nan
    o['promo1'] = o['CAP1'] / o['CA1'] if o['CA1'] else np.nan
    return o


def _traffic(o: dict, T, T1, V=None, V1=None) -> dict:
    o = dict(o)
    o['T'], o['T1'] = T, T1
    o['tk'] = T / T1 - 1 if T1 else np.nan
    o['panier'] = (o['CA'] / T) if T else np.nan
    o['panier1'] = (o['CA1'] / T1) if T1 else np.nan
    o['pan'] = o['panier'] / o['panier1'] - 1 if T and T1 and o['CA1'] else np.nan
    if V is not None and T and T1 and V1:
        o['art'] = V / T
        o['art1'] = V1 / T1
        o['artv'] = o['art'] / o['art1'] - 1
    return o


def compute(P: dict) -> dict:
    fam, ray, ds, st = P['fam'], P['ray'], P['ds'], P['st']
    T, T1 = P['T'], P['T1']
    R = {'meta': P['meta'], 'checks': P['checks']}

    g = _traffic(kpi(fam), T, T1, P['VOL'], P['VOL1'])
    R['global'] = g

    R['sect'] = {s: kpi(fam[fam.Secteur == s]) for s in SECTEUR_ORDER}
    R['fmt'] = {}
    for f in FORMAT_ORDER:
        q = st[st.Format == f]
        R['fmt'][f] = _traffic(kpi(fam[fam.Format == f]), q['T'].sum(), q['T1'].sum())
    R['mtx'] = {s: {f: kpi(fam[(fam.Secteur == s) & (fam.Format == f)])['dM'] for f in FORMAT_ORDER} for s in SECTEUR_ORDER}

    sites = []
    for x in sorted(fam.Site.unique()):
        q = st[st.Site == x]
        o = _traffic(kpi(fam[fam.Site == x]), q['T'].sum(), q['T1'].sum())
        o['Site'], o['Format'] = x, fam[fam.Site == x].Format.iloc[0]
        sites.append(o)
    R['sites'] = sorted(sites, key=lambda o: o['dM'])

    rays = []
    for (s, r), d in fam.groupby(['Secteur', 'Rayon']):
        q = ray[ray.Rayon == r]
        o = _traffic(kpi(d), q['T'].sum(), q['T1'].sum())
        o['Secteur'], o['Rayon'], o['label'] = s, r, label(r)
        rays.append(o)
    R['rayons'] = sorted(rays, key=lambda o: o['dM'])

    fams = []
    for (s, r, f), d in fam.groupby(['Secteur', 'Rayon', 'Famille']):
        o = kpi(d)
        o.update({'Secteur': s, 'Rayon': r, 'Famille': f, 'label': label(f), 'rlabel': label(r), 'key': f'{s}|{r}|{f}'})
        fams.append(o)
    F = pd.DataFrame(fams)
    F = F[F[['CA', 'CA1']].max(axis=1) >= SEUIL_FAMILLE]
    R['familles'] = {}
    for s in SECTEUR_ORDER:
        x = F[F.Secteur == s].sort_values('dM')
        R['familles'][s] = {
            'flop': x[x.dM < 0].head(5).to_dict('records'),
            'top': x[x.dM > 0].tail(5).iloc[::-1].to_dict('records'),
        }
        R['familles'][s]['fmt'] = {f: kpi(fam[(fam.Secteur == s) & (fam.Format == f)])['dM'] for f in FORMAT_ORDER}

    R['sect_detail'] = {}
    for s in SECTEUR_ORDER:
        d = fam[fam.Secteur == s]
        o = _traffic(kpi(d), T, T1, ds[ds.Secteur == s]['V'].sum(), ds[ds.Secteur == s]['V1'].sum())
        R['sect_detail'][s] = o

    R['contrib'] = _contributors(R)
    R['texts'] = _texts(R)
    return R


def _why_rayon(o):
    if abs(o['effT']) > abs(o['effV']):
        return f"taux {fpt(o['dt'])}"
    return f"CA {fpct(o['ca'])}"


def _why_site(o):
    if o['dM'] >= 0:
        return f"CA {fpct(o['ca'])}"
    if np.isfinite(o.get('tk', np.nan)) and np.isfinite(o.get('pan', np.nan)) and o['tk'] < o['pan']:
        return f"trafic {fpct(o['tk'])}"
    return f"panier {fpct(o['pan'])}"


def _contributors(R):
    sect_short = {'PGC': 'PGC', 'Frais': 'Frais', 'Non-alimentaire': 'Non-alim.'}
    rs = R['rayons']
    top = [o for o in sorted(rs, key=lambda o: -o['dM']) if o['dM'] > 0][:3]
    bot = [o for o in rs if o['dM'] < 0][:3][::-1]
    ray_rows = [{'n': o['label'], 'why': f"{sect_short[o['Secteur']]} · {_why_rayon(o)}", 'v': o['dM']} for o in top + bot]
    ss = R['sites']
    top = [o for o in sorted(ss, key=lambda o: -o['dM']) if o['dM'] > 0][:3]
    bot = [o for o in ss if o['dM'] < 0][:3][::-1]
    site_rows = [{'n': o['Site'], 'why': _why_site(o), 'v': o['dM']} for o in top + bot]
    return {'rayon': ray_rows, 'site': site_rows}


def _effect_word(o, pos):
    vol = abs(o['effV']) >= abs(o['effT'])
    if pos:
        return 'gagne en volume' if vol else 'gagne par le taux'
    return 'perd du volume' if vol else 'perd du taux'


def _texts(R):
    g = R['global']
    sect = R['sect']
    t = {}
    tk, pan = g['tk'], g['pan']
    if tk >= 0 and pan < 0:
        drv = 'le panier annule le trafic'
    elif tk < 0 and pan >= 0:
        drv = 'le trafic recule, le panier résiste'
    elif tk < 0 and pan < 0:
        drv = 'trafic et panier en baisse'
    else:
        drv = 'trafic et panier en hausse'
    t['p2_title'] = (f"marge {fdm(g['dM'])} M FCFA, {drv}")
    order = sorted(SECTEUR_ORDER, key=lambda s: sect[s]['dM'])
    worst, best = order[0], order[-1]
    if sect[best]['dM'] > 0 and sect[worst]['dM'] < 0:
        t['p3_title'] = f"le {best} compense, le {worst} pèse" if best != 'Non-alimentaire' else f"le Non-alimentaire compense, le {worst} pèse"
        if worst == 'Non-alimentaire':
            t['p3_title'] = f"le {best} compense, le Non-alimentaire pèse"
    elif sect[worst]['dM'] >= 0:
        t['p3_title'] = 'les trois secteurs progressent'
    else:
        t['p3_title'] = 'les trois secteurs reculent'
    wr = R['rayons'][0]
    ws = R['sites'][0]
    t['p4_title'] = f"{wr['label']} et {ws['Site']} pèsent le plus"
    neg = [s for s in order if sect[s]['dM'] < 0]
    pos = [s for s in order[::-1] if sect[s]['dM'] > 0]
    secteurs_line = ' ; '.join([f"{s} {fdm(sect[s]['dM'])} M" for s in pos] + [f"{s} {fdm(sect[s]['dM'])} M" for s in neg])
    t['faits'] = [
        f"Trafic {fpct(tk)}, panier {fpct(pan)} : CA {fpct(g['ca'])}.",
        f"Secteurs : {secteurs_line}.",
        (f"Pourquoi : {wr['label']} {_effect_word(wr, False)} ({fdm(wr['effV'] if abs(wr['effV']) >= abs(wr['effT']) else wr['effT'])} M) ; "
         f"cause à confirmer par le chef de groupe {wr['Secteur']}.") if wr['dM'] < 0 else
        f"Pourquoi : {wr['label']} est le rayon le plus en recul.",
    ]
    t['taux'] = f"Taux de marge {fnum(g['t'] * 100)} % ({fpt(g['dt'])})."
    t['sect_title'] = {}
    for s in SECTEUR_ORDER:
        o = sect[s]
        rs = [r for r in R['rayons'] if r['Secteur'] == s]
        if o['dM'] < 0:
            r = rs[0]
            t['sect_title'][s] = f"{fdm(o['dM'])} M, {r['label']} {_effect_word(r, False)}"
        else:
            r = max(rs, key=lambda r: r['dM'])
            t['sect_title'][s] = f"{fdm(o['dM'])} M, {r['label']} {_effect_word(r, True)}"
    sd = R['sect_detail']
    t['lecture'] = {}
    for s in SECTEUR_ORDER:
        o = sd[s]
        rs = [r for r in R['rayons'] if r['Secteur'] == s]
        key = rs[0] if o['dM'] < 0 else max(rs, key=lambda r: r['dM'])
        t['lecture'][s] = (f"CA {fpct(o['ca'])}, marge {fdm(o['dM'])} M (effet volume {fdm(o['effV'])} M, effet taux {fdm(o['effT'])} M). "
                           f"Rayon clé : {key['label']} ({fdm(key['dM'])} M). Volume {fpct(o['q'])}, prix moyen {fpct(o['px'])}.")
    promo = f"promo {fnum(g['promo'] * 100)} % du CA ({fnum(g['promo1'] * 100)} % en N-1)"
    t['contexte'] = promo
    return t


def run(files) -> dict:
    return compute(prepare(load_exports(files)))


# ------------------------------------------------------------------ causes et actions (relues depuis le classeur des chefs de groupe)
def merge_comments(R: dict, comments: dict | None) -> dict:
    """comments : {key: {...}} lu par excel_cdg.read_comments ; ajoute 'cdg' aux familles du Flop."""
    comments = comments or {}
    R['comments'] = comments
    for s in SECTEUR_ORDER:
        for f in R['familles'][s]['flop']:
            f['cdg'] = comments.get(f['key'], {})
    R['general'] = comments.get('__general__', {})
    return R
