"""Jeu de données fictif pour démonstration du rapport (aucun chiffre réel).

Transforme 3 exports réels en exports fictifs cohérents :
  - montants, marges et volumes rééchelonnés par secteur (N et N-1 avec des coefficients différents),
  - tickets rééchelonnés d'un même coefficient partout (les totaux restent justes),
  - magasins renommés.
Les contrôles de cohérence du moteur passent donc comme sur des exports réels.

Usage : python -m modules.revue_codir.demo export_A.xlsx export_B.xlsx export_C.xlsx masque.pptx sortie.pptx
"""
from __future__ import annotations

import io
import sys

import pandas as pd

from . import engine as E

NOMS = {
    '10202 - Hyper Palmeraie': '10202 - Hyper Atlantique', '10301 - Hyper Marcory': '10301 - Hyper Lagune',
    '10203 - Hyper Yopougon': '10203 - Hyper Plateau Ouest', '10206 - Market Kokoh Mall': '10206 - Market Centre',
    '10208 - Market Riviera': '10208 - Market Corniche', '10209 - Market 2 Plateaux': '10209 - Market Les Jardins',
    '10705 - Market 7 Décembre': '10705 - Market Boulevard', '10604 - Market Cité verte': '10604 - Market Les Palmiers',
    '10605 - Market Aboboté': '10605 - Market Les Acacias', '10603 - Market Toit rouge': '10603 - Market Bel Air',
    '10601 - Supeco Niangon': '10601 - Supeco Est', '10602 - Supeco Terminus 47': '10602 - Supeco Nord',
    '10606 - Supeco Square': '10606 - Supeco Square Ouest',
}
# (CA et volumes N, CA et volumes N-1, marge N, marge N-1) par secteur
COEF = {'PGC': (0.86, 0.80, 0.90, 0.82), 'Frais': (0.88, 0.90, 0.86, 0.90), 'Non-alimentaire': (0.79, 0.84, 0.82, 0.83)}
COEF_TICKETS = (0.91, 0.88)
CONVERSIONS_DEMO = [('Market Les Palmiers', '01/04/2026'), ('Market Les Acacias', '26/08/2026'), ('Market Bel Air', '22/10/2026')]
A_VERIFIER_DEMO = {'Market Bel Air': 'converti le 22/10'}


def _is_ratio(col: str) -> bool:
    return any(k in col for k in ('%', 'Taux', 'Panier', 'Poids', 'Vs', 'Evo'))


def _rename(df):
    for c in df.columns:
        if df[c].dtype == object or str(df[c].dtype).startswith('str'):
            df[c] = df[c].map(lambda v: _swap(v) if isinstance(v, str) else v)
    return df


def _swap(v: str) -> str:
    for a, b in NOMS.items():
        v = v.replace(a, b)
    return v


def _scale_rows(df, dept_col):
    num = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c]) and not _is_ratio(c)]
    for i, dep in df[dept_col].items():
        sect = E.SECTEURS.get(dep)
        if not sect:
            continue
        n, n1, m, m1 = COEF[sect]
        for c in num:
            if 'Débit' in c:
                f = COEF_TICKETS[1] if 'N-1' in c else COEF_TICKETS[0]
            elif 'Marge' in c:
                f = m1 if 'N-1' in c else m
            else:
                f = n1 if 'N-1' in c else n
            v = df.at[i, c]
            if pd.notna(v):
                df.at[i, c] = v * f
    return df


def _to_bytes(df, name):
    b = io.BytesIO()
    df.to_excel(b, index=False)
    b.seek(0)
    b.name = name
    return b


def fictive_exports(path_a, path_b, path_c):
    a = _rename(pd.read_excel(path_a))
    a = _scale_rows(a, 'Departement')
    b = _rename(pd.read_excel(path_b))
    b = _scale_rows(b, 'Département')
    gi = b.index[b['Département'] == 'Total']
    dep_tot = b[(b['Site'] == 'Total') & b['Département'].isin(E.SECTEURS)]
    for c in ('CA', 'CA N-1', 'Volume', 'Volume N-1', 'Marge', 'Marge N-1'):
        if c in b.columns:
            b.loc[gi, c] = dep_tot[c].sum()
    for c in ('Débit', 'Débit N-1'):
        b.loc[gi, c] = b.loc[gi, c] * (COEF_TICKETS[1] if 'N-1' in c else COEF_TICKETS[0])
    c_ = _rename(pd.read_excel(path_c))
    for col in ('Débit', 'Débit N-1'):
        f = COEF_TICKETS[1] if 'N-1' in col else COEF_TICKETS[0]
        c_[col] = c_[col] * f
    return [_to_bytes(a, 'A.xlsx'), _to_bytes(b, 'B.xlsx'), _to_bytes(c_, 'C.xlsx')]


def run_demo(path_a, path_b, path_c):
    E.CONVERSIONS[:] = CONVERSIONS_DEMO
    E.SITES_A_VERIFIER.clear()
    E.SITES_A_VERIFIER.update(A_VERIFIER_DEMO)
    R = E.run(fictive_exports(path_a, path_b, path_c))
    R['demo'] = True
    return E.merge_comments(R, {})


if __name__ == '__main__':
    from . import ppt_report as P
    pa, pb, pc, masque, out = sys.argv[1:6]
    R = run_demo(pa, pb, pc)
    P.build_report(R, masque, out)
    print('ok', out, [c for c in R['checks']])
