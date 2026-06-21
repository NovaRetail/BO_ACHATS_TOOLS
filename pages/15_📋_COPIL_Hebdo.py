"""
15_📋_COPIL_Hebdo.py — Module COPIL Hebdo · SmartBuyer Hub
Vue réseau (Rayon) + Destructeurs/Performeurs (Article) sur exports PowerBI hebdo.
Deux exports indépendants : Rayon→Famille→Sous-Famille, et Rayon→Famille→Sous-Famille→Article.
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# ============================================================
# CONFIG & CHARTE (identique au reste du Hub)
# ============================================================
st.set_page_config(page_title="COPIL Hebdo", page_icon="📋", layout="wide")

BLUE = "#007AFF"
GREEN = "#34C759"
RED = "#FF3B30"
AMBER = "#FF9500"
DARK = "#1C1C1E"
GREY = "#8E8E93"
BG = "#F2F2F7"

# ============================================================
# 🎯 CIBLES DE MARGE PAR RAYON — identiques au module 14_💰_Marge.py
# Clé = mot-clé contenu dans le libellé Rayon de l'export PBI.
# ============================================================
CIBLES_DEFAUT = {
    "BOISSONS": 19.5,
    "DROGUERIE": 25.0,
    "PARFUMERIE HYGIENE": 29.0,
    "EPICERIE": 16.0,
}
CIBLE_FALLBACK = 23.5

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html, body, [class*="css"] {{ font-family: 'Inter', -apple-system, 'SF Pro Display', BlinkMacSystemFont, Calibri, sans-serif; }}
.stApp {{ background-color: {BG}; }}
.block-container {{ padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1280px; }}
[data-testid="stSidebar"] {{ background-color: #FFFFFF; border-right: 1px solid #E5E5EA; }}
hr {{ border-color: #E5E5EA !important; margin: 1rem 0 !important; }}

.page-title {{ font-size: 28px; font-weight: 700; color: {DARK}; letter-spacing: -0.03em; margin: 0; }}
.page-caption {{ font-size: 13px; color: {GREY}; margin-top: 3px; margin-bottom: 1.5rem; }}
.section-label {{ font-size: 11px; font-weight: 600; color: {GREY};
                 text-transform: uppercase; letter-spacing: 0.07em; margin: 18px 0 10px; }}

.card {{ background:#FFFFFF; border-radius:12px; padding:16px 18px; margin-bottom:10px;
        border:1px solid #E5E5EA; box-shadow:0 1px 3px rgba(0,0,0,0.04); }}

.kpi-card {{ background:#FFFFFF; border-radius:12px; padding:16px 18px;
            border:1px solid #E5E5EA; box-shadow:0 1px 3px rgba(0,0,0,0.04); }}
.kpi-label {{ font-size:11px; font-weight:500; color:{GREY};
             text-transform:uppercase; letter-spacing:0.04em; margin-bottom:3px; }}
.kpi-value {{ font-size:24px; font-weight:700; color:{DARK}; letter-spacing:-0.02em; line-height:1.1; }}
.kpi-sub {{ font-size:12px; color:{GREY}; margin-top:3px; }}
.kpi-sub.pos {{ color:{GREEN}; }}
.kpi-sub.neg {{ color:{RED}; }}

.info-box {{ background:#F0F8FF; border-left:3px solid {BLUE}; border-radius:10px;
            padding:16px 20px; margin-bottom:24px; }}
.info-box .it {{ font-size:15px; font-weight:700; color:{DARK}; margin-bottom:10px; }}
.info-box .ip {{ font-size:13px; color:#1C1C1E; line-height:1.6; }}
.info-box .iq {{ margin-top:14px; font-size:13px; color:#1C1C1E; line-height:1.9; }}

.recap-card {{ background:#FFFFFF; border-radius:12px; padding:14px 18px; margin-bottom:18px;
              border:1px solid #E5E5EA; box-shadow:0 1px 3px rgba(0,0,0,0.04); }}
.recap-line1 {{ font-size:15px; font-weight:700; color:{DARK}; letter-spacing:-0.01em; line-height:1.5; }}
.recap-line2 {{ font-size:13px; color:{DARK}; margin-top:8px; padding-top:8px;
               border-top:1px solid #F0F0F2; line-height:1.5; }}
.recap-line2 b {{ color:{BLUE}; }}

.badge {{ display:inline-block; padding:2px 10px; border-radius:6px; font-size:11px; font-weight:600; }}
</style>
""", unsafe_allow_html=True)

# ============================================================
# HELPERS DE FORMAT (convention SmartBuyer — cf. 06_💸_Marges_Negatives.py)
# ============================================================
def fmt(n):
    if n is None or (isinstance(n, float) and (pd.isna(n) or not np.isfinite(n))):
        return "—"
    a = abs(n)
    if a >= 1_000_000: return f"{n/1_000_000:.1f} M"
    if a >= 1_000:     return f"{int(n/1_000)} K"
    return f"{int(n):,}"

def fmt_pct(v, dec=1):
    if v is None or pd.isna(v): return "—"
    return f"{v:.{dec}f}%"

def fmt_delta(v):
    if v is None or pd.isna(v): return "—"
    return f"{v:+.1f} pts"

def safe_rate(num, den):
    den = np.where(den == 0, np.nan, den)
    r = num / den * 100
    return np.where(np.isnan(r), 0.0, r)

def rayon_key(libelle):
    """Mot-clé de rapprochement vers CIBLES_DEFAUT (insensible au code/préfixe numérique)."""
    s = str(libelle).upper()
    for k in CIBLES_DEFAUT:
        if k in s:
            return k
    return None

def kpi_card(label, value, sub=None, sub_class=""):
    sub_html = f"<div class='kpi-sub {sub_class}'>{sub}</div>" if sub else ""
    return (f"<div class='kpi-card'><div class='kpi-label'>{label}</div>"
            f"<div class='kpi-value'>{value}</div>{sub_html}</div>")

# ============================================================
# CHARGEMENT — EXPORT ARTICLE UNIQUE (Rayon → Famille → Sous Famille → Article)
# Toute la hiérarchie (réseau / rayon / famille / article) est dérivée de ce
# seul fichier — c'est lui que l'on charge chaque semaine.
# ============================================================
@st.cache_data(show_spinner=False)
def load_export(file_bytes):
    raw = pd.read_excel(io.BytesIO(file_bytes))
    raw.columns = [str(c).lstrip('\ufeff') for c in raw.columns]

    # La note "Filtres appliqués : ..." (dernière ligne) décrit le périmètre exact
    # (sites, enseigne, période) — utile pour vérifier la cohérence semaine après semaine.
    perimetre = None
    note_rows = raw[raw['Rayon'].astype(str).str.startswith('Filtres', na=False)]
    if not note_rows.empty:
        perimetre = str(note_rows.iloc[0]['Rayon'])

    df = raw[raw['Rayon'].notna()].copy()
    df = df[~df['Rayon'].astype(str).str.startswith('Filtres')]
    if 'Article' not in df.columns:
        df['Article'] = np.nan
    return df, perimetre

def kpis_globaux_rayon(df):
    g = df[df['Rayon'] == 'Total']
    if g.empty:
        return None
    g = g.iloc[0]
    ca, ca_n1, marge = g['CA'], g['CA N-1'], g['Marge']
    evol_marge_pct = g.get('%Vs N-1.1', np.nan)
    marge_n1 = marge / (1 + evol_marge_pct) if pd.notna(evol_marge_pct) and (1 + evol_marge_pct) != 0 else np.nan
    return {
        'ca': ca, 'ca_n1': ca_n1, 'evol_ca': ca/ca_n1 - 1 if ca_n1 else np.nan,
        'marge': marge, 'marge_n1': marge_n1,
        'tx_marge': marge/ca*100 if ca else np.nan,
        'tx_marge_n1': marge_n1/ca_n1*100 if ca_n1 and pd.notna(marge_n1) else np.nan,
        'qte': g.get('Qté Vente', np.nan), 'qte_n1': g.get('Qté Vente N-1', np.nan),
        'poids_promo': g.get('%CA Poids Promo', np.nan) * 100,
        'casse': g.get('Casse (Valeur)', np.nan),
    }

def perf_par_rayon(df, cibles):
    rows = []
    sub = df[(df['Famille'] == 'Total') & (df['Rayon'] != 'Total')]
    for _, r in sub.iterrows():
        key = rayon_key(r['Rayon'])
        cible = cibles.get(key, CIBLE_FALLBACK) if key else CIBLE_FALLBACK
        ca, ca_n1 = r['CA'], r['CA N-1']
        qte, qte_n1 = r.get('Qté Vente', np.nan), r.get('Qté Vente N-1', np.nan)
        tx = r['Marge']/ca*100 if ca else np.nan
        rows.append({
            'Rayon': str(r['Rayon']).split(' - ')[-1].strip(),
            'CA': ca, 'Évol CA %': (ca/ca_n1-1)*100 if ca_n1 else np.nan,
            'Évol Qté %': (qte/qte_n1-1)*100 if qte_n1 else np.nan,
            'Taux Marge %': tx, 'Objectif %': cible, 'Écart (pts)': tx - cible if pd.notna(tx) else np.nan,
        })
    return pd.DataFrame(rows).sort_values('CA', ascending=False)

def family_metrics(df):
    """Prépare toutes les métriques Top/Flop au niveau Famille (Sous Famille == Total)."""
    sub = df[(df['Sous Famille'] == 'Total') & (df['Famille'] != 'Total') & (df['Rayon'] != 'Total')].copy()
    sub['Rayon_aff'] = sub['Rayon'].astype(str).str.split(' - ').str[-1].str.strip()
    sub['Famille_aff'] = sub['Famille'].astype(str).str.split(' - ').str[-1].str.strip()
    sub['CA'] = sub['CA'].fillna(0)
    sub['CA N-1'] = sub['CA N-1'].fillna(0)
    sub['Perte CA'] = sub['CA'] - sub['CA N-1']
    sub['Évol CA %'] = np.where(sub['CA N-1'] > 0, (sub['CA']/sub['CA N-1']-1)*100, np.nan)
    sub['Tx Marge %'] = np.where(sub['CA'] > 0, sub['Marge']/sub['CA']*100, np.nan)
    evol_marge_pct = sub.get('%Vs N-1.1', pd.Series(np.nan, index=sub.index))
    with np.errstate(divide='ignore', invalid='ignore'):
        marge_n1 = sub['Marge'] / (1 + evol_marge_pct)
    marge_n1 = marge_n1.replace([np.inf, -np.inf], np.nan)
    sub['Tx Marge N-1 %'] = np.where(sub['CA N-1'] > 0, marge_n1/sub['CA N-1']*100, np.nan)
    sub['Écart Tx Marge (pts)'] = sub['Tx Marge %'] - sub['Tx Marge N-1 %']
    qte = sub.get('Qté Vente', pd.Series(np.nan, index=sub.index))
    qte_n1 = sub.get('Qté Vente N-1', pd.Series(np.nan, index=sub.index))
    sub['Qté Vente'] = qte
    sub['Qté Vente N-1'] = qte_n1
    sub['Évol Qté %'] = np.where(qte_n1 > 0, (qte/qte_n1 - 1) * 100, np.nan)
    return sub

def build_headline(k, perf, fam):
    """Récap en 2 lignes : ligne 1 = synthèse complète des KPIs, ligne 2 = point d'attention prioritaire."""
    evol_ca = k['evol_ca'] * 100 if pd.notna(k['evol_ca']) else np.nan
    evo_tx = k['tx_marge'] - k['tx_marge_n1'] if pd.notna(k['tx_marge_n1']) else np.nan
    evol_qte = (k['qte']/k['qte_n1'] - 1) * 100 if k['qte_n1'] else np.nan
    pct_casse = k['casse']/k['ca']*100 if k['ca'] else np.nan

    line1 = (
        f"CA {fmt(k['ca'])} FCFA ({evol_ca:+.1f}%) &nbsp;·&nbsp; "
        f"Marge {fmt(k['marge'])} FCFA, taux {k['tx_marge']:.1f}% ({fmt_delta(evo_tx)}) &nbsp;·&nbsp; "
        f"Qté {fmt(k['qte'])} ({evol_qte:+.1f}%) &nbsp;·&nbsp; "
        f"Casse {pct_casse:.2f}% du CA &nbsp;·&nbsp; "
        f"Promo {k['poids_promo']:.1f}% du CA"
    )

    bits = []
    if perf is not None and perf['Écart (pts)'].notna().any():
        wr = perf.loc[perf['Écart (pts)'].idxmin()]
        bits.append(f"rayon à surveiller : <b>{wr['Rayon']}</b> ({fmt_delta(wr['Écart (pts)'])} vs objectif)")
    if fam is not None and len(fam) and fam['Perte CA'].notna().any():
        wf = fam.loc[fam['Perte CA'].idxmin()]
        fam_qte = wf.get('Évol Qté %', np.nan)
        fam_txt = (f"famille à risque : <b>{wf['Famille_aff']}</b> — "
                   f"CA {fmt(wf['CA'])} FCFA ({wf['Évol CA %']:+.1f}%), "
                   f"marge {wf['Tx Marge %']:.1f}%")
        if pd.notna(fam_qte):
            fam_txt += f", qté {fam_qte:+.1f}%"
        bits.append(fam_txt)
    line2 = "📌 " + " &nbsp;·&nbsp; ".join(bits) if bits else ""
    return line1, line2

def top_familles(df, n=5, by='perte_ca'):
    """Top/Flop N familles selon le critère choisi (conserve la version courte pour compatibilité)."""
    sub = family_metrics(df)
    if by == 'perte_ca':
        out = sub.nsmallest(n, 'Perte CA')[['Rayon_aff','Famille_aff','CA','CA N-1','Perte CA','Tx Marge %']]
    elif by == 'casse':
        out = sub.nsmallest(n, 'Casse (Valeur)')[['Rayon_aff','Famille_aff','CA','Casse (Valeur)','%Casse (Valeur)']]
    elif by == 'promo':
        mat = sub[sub['CA'] > 1_000_000]
        out = mat.nlargest(n, '%CA Poids Promo')[['Rayon_aff','Famille_aff','CA','%CA Poids Promo','%Marge Promo','%Marge Hors Promo']]
    return out.reset_index(drop=True)

def top_flop_table(sub, metric, n, mode, cols, ca_floor=0):
    """mode='top' -> nlargest, mode='flop' -> nsmallest. ca_floor filtre les familles trop petites (bruit)."""
    base = sub[sub['CA'] > ca_floor] if ca_floor else sub
    base = base[base[metric].notna()]
    out = base.nlargest(n, metric) if mode == 'top' else base.nsmallest(n, metric)
    return out[['Rayon_aff','Famille_aff'] + cols].reset_index(drop=True)

# ============================================================
# DÉRIVATION — VUE ARTICLE (à partir du même dataframe)
# ============================================================
def prep_articles(df):
    # lignes article réelles uniquement (pas les sous-totaux Rayon/Famille/Sous-Famille)
    art = df[df['Article'].notna() & (df['Article'] != 'Total') & (df['Rayon'] != 'Total')].copy()
    art['Rayon_aff'] = art['Rayon'].astype(str).str.split(' - ').str[-1].str.strip()
    art['Famille_aff'] = art['Famille'].astype(str).str.split(' - ').str[-1].str.strip()
    art['SousFamille_aff'] = art['Sous Famille'].astype(str).str.split(' - ').str[-1].str.strip()
    art['Article_aff'] = art['Article'].astype(str)
    art['CA'] = art['CA'].fillna(0)
    art['Marge'] = art['Marge'].fillna(0)
    art['Qté Vente'] = art['Qté Vente'].fillna(0)
    art['Qté Vente N-1'] = art['Qté Vente N-1'].fillna(0)
    art['CA N-1'] = art['CA N-1'].fillna(0)
    evol_marge = art.get('%Vs N-1.1', pd.Series(np.nan, index=art.index))
    with np.errstate(divide='ignore', invalid='ignore'):
        marge_n1 = art['Marge'] / (1 + evol_marge)
    marge_n1 = marge_n1.replace([np.inf, -np.inf], np.nan)
    art['Marge N-1 (calc)'] = marge_n1
    art['Tx Marge %'] = np.where(art['CA'] > 0, art['Marge']/art['CA']*100, np.nan)
    art['Tx Marge N-1 % (calc)'] = np.where(art['CA N-1'] > 0, marge_n1/art['CA N-1']*100, np.nan)
    art['Écart Tx Marge (pts)'] = art['Tx Marge %'] - art['Tx Marge N-1 % (calc)']
    art['Gain Marge (FCFA)'] = art['Marge'] - marge_n1
    art['Variation Qté'] = art['Qté Vente'] - art['Qté Vente N-1']
    art['Perte CA (FCFA)'] = art['CA'] - art['CA N-1']
    return art

def destructeurs_performeurs(art, n=15, seuil_ca=100_000):
    res = {}
    res['A_marge_neg'] = art[art['Marge'] < 0].nsmallest(n, 'Marge')
    pos = art[art['Marge'] >= 0].copy()
    res['B_degrad_marge'] = pos[pos['Écart Tx Marge (pts)'].notna()].nsmallest(n, 'Écart Tx Marge (pts)')
    mat = art[art['CA'] > seuil_ca]
    res['C_perf_gain_marge'] = mat.nlargest(n, 'Gain Marge (FCFA)')
    mat_n1 = mat[mat['CA N-1'] > 0].assign(_evol=lambda d: d['CA']/d['CA N-1']-1)
    res['D_croissance_ca'] = mat_n1.nlargest(n, '_evol')
    res['E_baisse_ca'] = art.nsmallest(n, 'Perte CA (FCFA)')
    res['F_hausse_qte'] = art.nlargest(n, 'Variation Qté')
    res['G_baisse_qte'] = art.nsmallest(n, 'Variation Qté')
    return res

DISPLAY_COLS = ['Rayon_aff','Famille_aff','SousFamille_aff','Article_aff']
DISPLAY_RENAME = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article'}

def show_table(d, extra_cols, fmt_map=None):
    if d.empty:
        st.caption("Aucune ligne ne correspond à ce critère sur la période.")
        return
    cols = DISPLAY_COLS + extra_cols
    disp = d[cols].rename(columns=DISPLAY_RENAME).reset_index(drop=True)
    disp.index = disp.index + 1
    if fmt_map:
        for c, f in fmt_map.items():
            if c in disp.columns:
                disp[c] = disp[c].map(f)
    st.dataframe(disp, use_container_width=True)

# ============================================================
# EXPORT EXCEL (valeurs figées — recalcul natif via Data_Semaine/Data_Articles
# conseillé pour le suivi hebdo continu, ce bouton sert aux exports ponctuels COPIL)
# ============================================================
def build_excel_full(k, perf, fam, n_top, art_res, perimetre):
    """Reproduit le classeur de référence : Dashboard COPIL (5+ sections) + Destructeurs & Performeurs (A-G)."""
    BLUE_H = "FF007AFF"; DARK_H = "FF1C1C1E"; RED_H = "FFFF3B30"; GREEN_H = "FF34C759"
    WHITE_H = "FFFFFFFF"; LGREY_H = "FFE5E5EA"; ARIAL = "Arial"
    thin = Side(style="thin", color="FFD1D1D6")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Format "Comptabilité" Excel natif, sans symbole monétaire (séparateur de milliers,
    # négatifs alignés, zéro affiché "-"). 'amount_signed' identique mais avec signe +/- explicite.
    ACC = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
    ACC_SIGNED = '_-* +#,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
    QTY = "#,##0"
    PTS = '+0.00" pts";-0.00" pts"'
    PCT = "0.0%"
    PCT2 = "0.00%"

    def fmt_value(kind, v):
        """Convertit v selon le type déclaré : pct100 -> fraction (v/100), sinon inchangé."""
        if kind == 'pct100' and v is not None and pd.notna(v):
            return v / 100
        return v

    def fmt_code(kind):
        return {'amount': ACC, 'amount_signed': ACC_SIGNED, 'qty': QTY,
                'pct100': PCT, 'pct1': PCT, 'pts': PTS}.get(kind, "General")

    def section_bar(ws, row, ncols, text, color=DARK_H):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        c = ws.cell(row=row, column=1, value=text)
        c.font = Font(name=ARIAL, bold=True, size=11, color=WHITE_H)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 20

    def header_row(ws, row, labels):
        for i, lbl in enumerate(labels, start=1):
            c = ws.cell(row=row, column=i, value=lbl)
            c.font = Font(name=ARIAL, bold=True, size=10, color=WHITE_H)
            c.fill = PatternFill("solid", fgColor=BLUE_H)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def data_row(ws, row, values, zebra=False, left_cols=()):
        fillc = LGREY_H if zebra else "FFFFFFFF"
        for i, v in enumerate(values, start=1):
            c = ws.cell(row=row, column=i, value=v)
            c.font = Font(name=ARIAL, size=10)
            c.border = box
            c.fill = PatternFill("solid", fgColor=fillc)
            c.alignment = Alignment(horizontal="left" if i in left_cols else "center")

    def autosize(ws, widths):
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    wb = Workbook()

    # ============== FEUILLE 1 : DASHBOARD COPIL ==============
    ws = wb.active
    ws.title = "Dashboard COPIL"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:H1")
    ws["A1"] = "DASHBOARD COPIL HEBDO — PGC (Épicerie · Boissons · Droguerie · Parfumerie Hygiène)"
    ws["A1"].font = Font(name=ARIAL, bold=True, size=14, color=WHITE_H)
    ws["A1"].fill = PatternFill("solid", fgColor=BLUE_H)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26
    ws["A2"] = "Période :"; ws["A2"].font = Font(name=ARIAL, bold=True, size=10)
    ws["B2"] = "Voir export source"; ws["B2"].font = Font(name=ARIAL, size=10, color="FF0000FF", bold=True)
    if perimetre:
        ws["A3"] = "Périmètre :"; ws["A3"].font = Font(name=ARIAL, size=9, italic=True, color="FF8E8E93")
        ws["B3"] = perimetre.replace("\n", " ")[:200]
        ws["B3"].font = Font(name=ARIAL, size=9, italic=True, color="FF8E8E93")

    r = 5
    section_bar(ws, r, 8, "1.  VUE D'ENSEMBLE RÉSEAU"); r += 1
    header_row(ws, r, ["Indicateur", "Cette semaine", "N-1", "Évolution"]); r += 1
    evo_tx = k['tx_marge'] - k['tx_marge_n1'] if pd.notna(k['tx_marge_n1']) else None
    evol_qte = (k['qte']/k['qte_n1']-1) if k['qte_n1'] else None
    pct_casse = k['casse']/k['ca'] if k['ca'] else None
    # (label, valeur, valeur N-1, évolution, kind_valeur, kind_évolution)
    kpi_rows = [
        ("CA (FCFA)",          k['ca'],            k['ca_n1'],   k['evol_ca'], 'amount', 'pct1'),
        ("Marge (FCFA)",       k['marge'],         k['marge_n1'],None,         'amount', None),
        ("Taux de marge",      k['tx_marge']/100,  (k['tx_marge_n1']/100 if pd.notna(k['tx_marge_n1']) else None), evo_tx, 'pct1', 'pts'),
        ("Qté vendue",         k['qte'],           k['qte_n1'],  evol_qte,     'qty',    'pct1'),
        ("Poids Promo (% CA)", k['poids_promo']/100, None,       None,         'pct1',   None),
        ("Casse (FCFA)",       k['casse'],         None,         pct_casse,    'amount', 'pct1'),
    ]
    r0kpi = r
    for i, (label, v, n1, evo, kind_v, kind_evo) in enumerate(kpi_rows):
        zebra = i % 2 == 1
        fillc = LGREY_H if zebra else "FFFFFFFF"
        for col in range(1, 5): ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=fillc); ws.cell(row=r, column=col).border = box; ws.cell(row=r, column=col).font = Font(name=ARIAL, size=10)
        ws.cell(row=r, column=1, value=label).alignment = Alignment(horizontal="left", indent=1)
        ws.cell(row=r, column=2, value=v); ws.cell(row=r, column=2).number_format = fmt_code(kind_v)
        if n1 is not None:
            ws.cell(row=r, column=3, value=n1); ws.cell(row=r, column=3).number_format = fmt_code(kind_v)
        if evo is not None:
            ws.cell(row=r, column=4, value=evo)
            ws.cell(row=r, column=4).number_format = fmt_code(kind_evo)
        r += 1
    ws.conditional_formatting.add(f"D{r0kpi}:D{r-1}", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="FFFFD6D4")))
    ws.conditional_formatting.add(f"D{r0kpi}:D{r-1}", CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PatternFill("solid", fgColor="FFD7F5DE")))
    r += 1

    section_bar(ws, r, 8, "2.  PERFORMANCE PAR RAYON VS OBJECTIFS MARGE (MÉTI)"); r += 1
    header_row(ws, r, ["Rayon", "CA (FCFA)", "Évol CA", "Évol Qté", "Taux Marge", "Objectif Méti", "Écart (pts)"]); r += 1
    r0r = r
    for i, (_, row_) in enumerate(perf.iterrows()):
        data_row(ws, r, [row_['Rayon'], row_['CA'], row_['Évol CA %']/100, row_['Évol Qté %']/100,
                          row_['Taux Marge %']/100, row_['Objectif %']/100, row_['Écart (pts)']], zebra=(i%2==1), left_cols=(1,))
        for col, fmt_ in [(2,ACC),(3,PCT),(4,PCT),(5,PCT2),(6,PCT),(7,PTS)]:
            ws.cell(row=r, column=col).number_format = fmt_
        r += 1
    ws.conditional_formatting.add(f"G{r0r}:G{r-1}", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="FFFFD6D4")))
    ws.conditional_formatting.add(f"G{r0r}:G{r-1}", CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PatternFill("solid", fgColor="FFD7F5DE")))
    r += 1

    section_bar(ws, r, 8, "3.  DÉTAIL PAR FAMILLE — TOUTES FAMILLES (sans objectif)"); r += 1
    header_row(ws, r, ["Rayon", "Famille", "CA", "Évol CA %", "Marge", "Taux Marge %", "Qté Vente", "Évol Qté %"]); r += 1
    fam_sorted = fam.sort_values(['Rayon_aff', 'CA'], ascending=[True, False])
    for i, (_, row_) in enumerate(fam_sorted.iterrows()):
        data_row(ws, r, [row_['Rayon_aff'], row_['Famille_aff'], row_['CA'],
                          (row_['Évol CA %']/100 if pd.notna(row_['Évol CA %']) else None), row_['Marge'],
                          (row_['Tx Marge %']/100 if pd.notna(row_['Tx Marge %']) else None), row_['Qté Vente'],
                          (row_['Évol Qté %']/100 if pd.notna(row_['Évol Qté %']) else None)], zebra=(i%2==1), left_cols=(1,2))
        for col, fmt_ in [(3,ACC),(4,PCT),(5,ACC),(6,PCT),(7,QTY),(8,PCT)]:
            ws.cell(row=r, column=col).number_format = fmt_
        r += 1
    r += 1

    def top_section(title, dframe, cols_map, kinds, color=RED_H):
        """kinds : dict {clé_colonne_source: kind} où kind ∈ amount/qty/pct100/pct1/pts."""
        nonlocal r
        section_bar(ws, r, 8, title, color=color); r += 1
        header_row(ws, r, ["Rang"] + list(cols_map.values())); r += 1
        keys = list(cols_map.keys())
        for i, (_, row_) in enumerate(dframe.iterrows()):
            vals = [i+1] + [fmt_value(kinds.get(c, 'amount'), row_[c]) for c in keys]
            data_row(ws, r, vals, zebra=(i%2==1), left_cols=(2,3))
            for j, ck in enumerate(keys, start=2):
                ws.cell(row=r, column=j).number_format = fmt_code(kinds.get(ck, 'amount'))
            r += 1
        r += 1

    top_section(f"4.  TOP {n_top} — PLUS FORTE BAISSE DE CA (par Famille)",
                top_flop_table(fam, 'Perte CA', n_top, 'flop', ['CA','CA N-1','Évol CA %','Perte CA','Tx Marge %']),
                {'Rayon_aff':'Rayon','Famille_aff':'Famille','CA':'CA (FCFA)','CA N-1':'CA N-1','Évol CA %':'Évol %','Perte CA':'Perte (FCFA)','Tx Marge %':'Taux Marge'},
                {'CA':'amount','CA N-1':'amount','Évol CA %':'pct100','Perte CA':'amount','Tx Marge %':'pct100'})

    top_section(f"5.  TOP {n_top} — MEILLEUR GAIN DE CA (par Famille)",
                top_flop_table(fam, 'Perte CA', n_top, 'top', ['CA','CA N-1','Évol CA %','Perte CA','Tx Marge %']),
                {'Rayon_aff':'Rayon','Famille_aff':'Famille','CA':'CA (FCFA)','CA N-1':'CA N-1','Évol CA %':'Évol %','Perte CA':'Gain (FCFA)','Tx Marge %':'Taux Marge'},
                {'CA':'amount','CA N-1':'amount','Évol CA %':'pct100','Perte CA':'amount','Tx Marge %':'pct100'},
                color=GREEN_H)

    top_section(f"6.  TOP {n_top} — CASSE EN VALEUR (par Famille)",
                top_familles_for_excel(fam, n_top, 'casse'),
                {'Rayon_aff':'Rayon','Famille_aff':'Famille','CA':'CA (FCFA)','Casse (Valeur)':'Casse (FCFA)','%Casse (Valeur)':'% Casse'},
                {'CA':'amount','Casse (Valeur)':'amount','%Casse (Valeur)':'pct1'})

    top_section(f"7.  TOP {n_top} — POIDS PROMO LE PLUS ÉLEVÉ (Famille, CA > 1M)",
                top_familles_for_excel(fam, n_top, 'promo'),
                {'Rayon_aff':'Rayon','Famille_aff':'Famille','CA':'CA (FCFA)','%CA Poids Promo':'Poids Promo','%Marge Promo':'Tx M. Promo','%Marge Hors Promo':'Tx M. HP'},
                {'CA':'amount','%CA Poids Promo':'pct1','%Marge Promo':'pct1','%Marge Hors Promo':'pct1'},
                color="FFFF9500")

    autosize(ws, {'A':22,'B':26,'C':16,'D':13,'E':16,'F':13,'G':13,'H':13})
    ws.freeze_panes = "A6"

    # ============== FEUILLE 2 : DESTRUCTEURS & PERFORMEURS ==============
    ws2 = wb.create_sheet("Destructeurs Performeurs")
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells("A1:I1")
    ws2["A1"] = "DESTRUCTEURS & PERFORMEURS — NIVEAU ARTICLE"
    ws2["A1"].font = Font(name=ARIAL, bold=True, size=14, color=WHITE_H)
    ws2["A1"].fill = PatternFill("solid", fgColor=BLUE_H)
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws2.row_dimensions[1].height = 26

    r2 = 3
    def art_section(title, dframe, cols_map, kinds, color):
        nonlocal r2
        ws2.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=9)
        c = ws2.cell(row=r2, column=1, value=title)
        c.font = Font(name=ARIAL, bold=True, size=11, color=WHITE_H)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws2.row_dimensions[r2].height = 20
        r2 += 1
        for i, lbl in enumerate(["Rang"] + list(cols_map.values()), start=1):
            cell = ws2.cell(row=r2, column=i, value=lbl)
            cell.font = Font(name=ARIAL, bold=True, size=10, color=WHITE_H)
            cell.fill = PatternFill("solid", fgColor=BLUE_H)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        r2 += 1
        keys = list(cols_map.keys())
        for i, (_, row_) in enumerate(dframe.iterrows()):
            vals = [i+1] + [fmt_value(kinds.get(c, 'amount'), row_.get(c, None)) for c in keys]
            fillc = LGREY_H if i % 2 == 1 else "FFFFFFFF"
            for j, v in enumerate(vals, start=1):
                cell = ws2.cell(row=r2, column=j, value=v)
                cell.font = Font(name=ARIAL, size=10)
                cell.border = box
                cell.fill = PatternFill("solid", fgColor=fillc)
                cell.alignment = Alignment(horizontal="left" if j in (2,3,4,5) else "center")
                if j >= 6:
                    cell.number_format = fmt_code(kinds.get(keys[j-2], 'amount'))
            r2 += 1
        r2 += 1

    cm_neg = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','CA':'CA','Marge':'Marge','Tx Marge %':'Taux Marge'}
    art_section(f"A.  TOP {n_top} — ARTICLES EN MARGE NÉGATIVE", art_res['A_marge_neg'], cm_neg,
                {'CA':'amount','Marge':'amount','Tx Marge %':'pct100'}, RED_H)

    cm_deg = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','CA':'CA','Tx Marge %':'Taux Marge','Écart Tx Marge (pts)':'Écart pts'}
    art_section(f"B.  TOP {n_top} — DÉGRADATION DU TAUX DE MARGE (marge encore positive)", art_res['B_degrad_marge'], cm_deg,
                {'CA':'amount','Tx Marge %':'pct100','Écart Tx Marge (pts)':'pts'}, RED_H)

    cm_gain = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','CA':'CA','Gain Marge (FCFA)':'Gain Marge','Tx Marge %':'Taux Marge'}
    art_section(f"C.  TOP {n_top} — PERFORMEURS : GAIN DE MARGE EN VALEUR", art_res['C_perf_gain_marge'], cm_gain,
                {'CA':'amount','Gain Marge (FCFA)':'amount','Tx Marge %':'pct100'}, GREEN_H)

    d4 = art_res['D_croissance_ca'].copy()
    if not d4.empty:
        d4['Évol CA %'] = (d4['CA']/d4['CA N-1']-1)*100
    cm_croi = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','CA':'CA','CA N-1':'CA N-1','Évol CA %':'Évol %','Tx Marge %':'Taux Marge'}
    art_section(f"D.  TOP {n_top} — PLUS FORTE CROISSANCE DE CA", d4, cm_croi,
                {'CA':'amount','CA N-1':'amount','Évol CA %':'pct100','Tx Marge %':'pct100'}, GREEN_H)

    cm_baisse = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','CA':'CA','CA N-1':'CA N-1','Perte CA (FCFA)':'Perte (FCFA)'}
    art_section(f"E.  TOP {n_top} — PLUS FORTE BAISSE DE CA", art_res['E_baisse_ca'], cm_baisse,
                {'CA':'amount','CA N-1':'amount','Perte CA (FCFA)':'amount'}, RED_H)

    cm_qte = {'Rayon_aff':'Rayon','Famille_aff':'Famille','SousFamille_aff':'Sous Famille','Article_aff':'Article','Qté Vente':'Qté Vente','Qté Vente N-1':'Qté N-1','Variation Qté':'Variation','CA':'CA'}
    kinds_qte = {'Qté Vente':'qty','Qté Vente N-1':'qty','Variation Qté':'amount_signed','CA':'amount'}
    art_section(f"F.  TOP {n_top} — PLUS FORTE HAUSSE DE QUANTITÉ VENDUE", art_res['F_hausse_qte'], cm_qte, kinds_qte, GREEN_H)
    art_section(f"G.  TOP {n_top} — PLUS FORTE BAISSE DE QUANTITÉ VENDUE", art_res['G_baisse_qte'], cm_qte, kinds_qte, RED_H)

    autosize(ws2, {'A':7,'B':20,'C':24,'D':22,'E':38,'F':13,'G':13,'H':13,'I':13})

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def top_familles_for_excel(fam, n, by):
    """Variante de top_familles qui part directement d'un family_metrics déjà calculé."""
    if by == 'casse':
        sub = fam[fam['Casse (Valeur)'].notna()]
        out = sub.nsmallest(n, 'Casse (Valeur)')[['Rayon_aff','Famille_aff','CA','Casse (Valeur)','%Casse (Valeur)']]
    elif by == 'promo':
        mat = fam[fam['CA'] > 1_000_000]
        out = mat.nlargest(n, '%CA Poids Promo')[['Rayon_aff','Famille_aff','CA','%CA Poids Promo','%Marge Promo','%Marge Hors Promo']]
    return out.reset_index(drop=True)

# ============================================================
# INTERFACE
# ============================================================
st.markdown("<div class='page-title'>📋 Module COPIL Hebdo</div>"
            "<div class='page-caption'>Vue réseau PGC (Rayon) + Destructeurs/Performeurs (Article) · "
            "objectifs marge alignés Méti · une seule extraction à charger</div>", unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 📥 Import")
    up = st.file_uploader("Export Article hebdo (.xlsx)", type=['xlsx'], key="up_export")
    st.caption("L'extraction Article contient déjà les sous-totaux Rayon et Famille — "
               "tout le module en est dérivé, rien d'autre à charger.")
    st.markdown("---")
    st.markdown("##### ⚙️ Paramètres")
    n_top = st.slider("Nombre de lignes par classement", 5, 30, 15)
    seuil_ca = st.number_input("Seuil CA mini — performeurs/croissance (FCFA)", 0, 5_000_000, 100_000, step=10_000)
    st.markdown("---")
    st.caption("SmartBuyer Hub · Module COPIL Hebdo")

if up is None:
    st.markdown(
        f"<div class='info-box'>"
        f"<div class='it'>ℹ️ À quoi sert ce module ?</div>"
        f"<div class='ip'>Ce module prépare le point hebdo réseau pour le COPIL à partir de l'export PBI "
        f"<b>Rayon → Famille → Sous-Famille → Article</b>. Un seul fichier à charger chaque semaine, "
        f"dans la barre latérale.</div>"
        f"<div class='iq'>"
        f"<b>Onglet Dashboard COPIL</b> — CA, marge, quantités vs N-1 · performance par rayon vs objectifs Méti · "
        f"top familles en baisse de CA / casse / poids promo<br>"
        f"<b>Onglet Destructeurs &amp; Performeurs</b> — articles en marge négative, dégradation de taux de marge, "
        f"gain de marge, croissance/baisse de CA et de quantité</div>"
        f"</div>", unsafe_allow_html=True)
    st.stop()

df, perimetre = load_export(up.getvalue())
art = prep_articles(df)

if perimetre:
    with st.expander("🔎 Périmètre détecté dans le fichier"):
        st.code(perimetre, language=None)

tab1, tab2 = st.tabs(["📋 Dashboard COPIL", "💥 Destructeurs & Performeurs"])

# ---------------- TAB 1 : DASHBOARD COPIL (RAYON) ----------------
with tab1:
    k = kpis_globaux_rayon(df)
    if k is None:
        st.error("Ligne de total réseau ('Total') introuvable dans l'export — vérifiez le fichier.")
    else:
        perf = perf_par_rayon(df, CIBLES_DEFAUT)
        fam = family_metrics(df)
        line1, line2 = build_headline(k, perf, fam)
        st.markdown(
            f"<div class='recap-card'><div class='recap-line1'>{line1}</div>"
            + (f"<div class='recap-line2'>{line2}</div>" if line2 else "")
            + "</div>", unsafe_allow_html=True)

        st.markdown("<div class='section-label'>VUE D'ENSEMBLE RÉSEAU</div>", unsafe_allow_html=True)
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.markdown(kpi_card("CA", f"{fmt(k['ca'])} FCFA",
                    f"{k['evol_ca']*100:+.1f}% vs N-1", "pos" if k['evol_ca'] >= 0 else "neg"), unsafe_allow_html=True)
        c2.markdown(kpi_card("Marge", f"{fmt(k['marge'])} FCFA"), unsafe_allow_html=True)
        evo_tx = k['tx_marge'] - k['tx_marge_n1'] if pd.notna(k['tx_marge_n1']) else np.nan
        c3.markdown(kpi_card("Taux de marge", fmt_pct(k['tx_marge'], 2),
                    fmt_delta(evo_tx), "pos" if (pd.notna(evo_tx) and evo_tx >= 0) else "neg"), unsafe_allow_html=True)
        evo_qte = (k['qte']/k['qte_n1']-1)*100 if k['qte_n1'] else np.nan
        c4.markdown(kpi_card("Qté vendue", fmt(k['qte']),
                    f"{evo_qte:+.1f}% vs N-1" if pd.notna(evo_qte) else None,
                    "pos" if (pd.notna(evo_qte) and evo_qte >= 0) else "neg"), unsafe_allow_html=True)
        pct_casse = k['casse']/k['ca']*100 if k['ca'] else np.nan
        c5.markdown(kpi_card("Casse", f"{fmt(k['casse'])} FCFA", fmt_pct(pct_casse, 2) + " du CA"), unsafe_allow_html=True)

        st.markdown("<div class='section-label'>PERFORMANCE PAR RAYON VS OBJECTIFS MARGE (MÉTI)</div>", unsafe_allow_html=True)
        disp = perf.copy()
        for c in ['Évol CA %', 'Évol Qté %', 'Taux Marge %', 'Objectif %']:
            disp[c] = disp[c].map(lambda v: fmt_pct(v))
        disp['Écart (pts)'] = perf['Écart (pts)'].map(fmt_delta)
        disp['CA'] = perf['CA'].map(lambda v: fmt(v))
        st.dataframe(disp, use_container_width=True, hide_index=True)

        st.markdown("<div class='section-label'>DÉTAIL PAR FAMILLE — TOUTES FAMILLES (SANS OBJECTIF)</div>", unsafe_allow_html=True)
        fam_disp = fam[['Rayon_aff','Famille_aff','CA','Évol CA %','Marge','Tx Marge %','Qté Vente','Évol Qté %']].copy()
        fam_disp = fam_disp.rename(columns={'Rayon_aff':'Rayon','Famille_aff':'Famille','Tx Marge %':'Taux Marge %'})
        fam_disp = fam_disp.sort_values(['Rayon','CA'], ascending=[True, False])
        for c in ['CA','Marge','Qté Vente']:
            fam_disp[c] = fam_disp[c].map(fmt)
        for c in ['Évol CA %','Taux Marge %','Évol Qté %']:
            fam_disp[c] = fam_disp[c].map(lambda v: fmt_pct(v))
        st.dataframe(fam_disp, use_container_width=True, hide_index=True, height=420)
        st.caption("Pas d'objectif au niveau Famille (le cadrage marge est piloté au niveau Rayon) — vue CA / marge / quantité uniquement.")

        st.markdown("<div class='section-label'>TOP & FLOP PAR FAMILLE</div>", unsafe_allow_html=True)

        def pair(title_top, title_flop, metric, cols, fmt_map, ca_floor=0):
            cA, cB = st.columns(2)
            with cA:
                st.markdown(f"**🟢 {title_top}**")
                t = top_flop_table(fam, metric, n_top, 'top', cols, ca_floor)
                t = t.rename(columns={'Rayon_aff':'Rayon','Famille_aff':'Famille'})
                for c, f in fmt_map.items():
                    if c in t.columns: t[c] = t[c].map(f)
                st.dataframe(t, hide_index=True, use_container_width=True)
            with cB:
                st.markdown(f"**🔴 {title_flop}**")
                t = top_flop_table(fam, metric, n_top, 'flop', cols, ca_floor)
                t = t.rename(columns={'Rayon_aff':'Rayon','Famille_aff':'Famille'})
                for c, f in fmt_map.items():
                    if c in t.columns: t[c] = t[c].map(f)
                st.dataframe(t, hide_index=True, use_container_width=True)

        st.markdown("##### 📈 Évolution du CA (classement en valeur FCFA)")
        pair(f"Top {n_top} — Meilleur gain de CA", f"Flop {n_top} — Plus forte perte de CA",
             'Perte CA', ['CA','CA N-1','Évol CA %','Tx Marge %'],
             {'CA': fmt, 'CA N-1': fmt, 'Évol CA %': lambda v: fmt_pct(v), 'Tx Marge %': lambda v: fmt_pct(v)})

        st.markdown("##### 💰 Taux de marge")
        pair(f"Top {n_top} — Meilleur taux de marge", f"Flop {n_top} — Taux de marge le plus faible",
             'Tx Marge %', ['CA','Marge','Tx Marge %'],
             {'CA': fmt, 'Marge': fmt, 'Tx Marge %': lambda v: fmt_pct(v)}, ca_floor=1_000_000)

        st.markdown("##### 📊 Évolution du taux de marge (pts vs N-1)")
        pair(f"Top {n_top} — Meilleure progression", f"Flop {n_top} — Plus forte dégradation",
             'Écart Tx Marge (pts)', ['CA','Tx Marge %','Écart Tx Marge (pts)'],
             {'CA': fmt, 'Tx Marge %': lambda v: fmt_pct(v), 'Écart Tx Marge (pts)': fmt_delta}, ca_floor=1_000_000)

        st.markdown("##### 💔 Casse & 🎯 Poids promo")
        cC, cD = st.columns(2)
        with cC:
            st.markdown(f"**🔴 Flop {n_top} — Casse en valeur**")
            t = top_familles(df, n_top, 'casse').rename(
                columns={'Rayon_aff':'Rayon','Famille_aff':'Famille','Casse (Valeur)':'Casse','%Casse (Valeur)':'% Casse'})
            t['CA'] = t['CA'].map(fmt); t['Casse'] = t['Casse'].map(fmt)
            t['% Casse'] = t['% Casse'].map(lambda v: fmt_pct(v*100, 2))
            st.dataframe(t, hide_index=True, use_container_width=True)
        with cD:
            st.markdown(f"**🟠 Top {n_top} — Poids promo (CA&gt;1M)**")
            t = top_familles(df, n_top, 'promo').rename(
                columns={'Rayon_aff':'Rayon','Famille_aff':'Famille','%CA Poids Promo':'Poids Promo',
                         '%Marge Promo':'Tx M. Promo','%Marge Hors Promo':'Tx M. HP'})
            t['CA'] = t['CA'].map(fmt)
            for c in ['Poids Promo','Tx M. Promo','Tx M. HP']:
                t[c] = t[c].map(lambda v: fmt_pct(v*100))
            st.dataframe(t, hide_index=True, use_container_width=True)

        st.markdown("<div class='section-label'>EXPORT</div>", unsafe_allow_html=True)
        art_res_export = destructeurs_performeurs(art, n=n_top, seuil_ca=seuil_ca)
        xls = build_excel_full(k, perf, fam, n_top, art_res_export, perimetre)
        st.download_button("📥 Télécharger le récap complet COPIL + Articles (.xlsx)", xls,
                            file_name="COPIL_Hebdo.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Le fichier contient 2 feuilles : Dashboard COPIL (réseau/rayon/famille) "
                   "et Destructeurs & Performeurs (article).")

# ---------------- TAB 2 : DESTRUCTEURS & PERFORMEURS (ARTICLE) ----------------
with tab2:
    st.caption(f"{len(art):,} lignes article chargées · seuil de matérialité : {fmt(seuil_ca)} FCFA "
               f"(modifiable dans la barre latérale)".replace(",", " "))
    res = destructeurs_performeurs(art, n=n_top, seuil_ca=seuil_ca)

    st.markdown(f"<span class='badge' style='background:#FFD6D4;color:{RED}'>A · Marge négative</span>", unsafe_allow_html=True)
    show_table(res['A_marge_neg'], ['CA','Marge','Tx Marge %','Qté Vente'],
               {'CA': fmt, 'Marge': fmt, 'Tx Marge %': lambda v: fmt_pct(v), 'Qté Vente': fmt})

    st.markdown(f"<span class='badge' style='background:#FFD6D4;color:{RED}'>B · Dégradation du taux de marge (marge encore positive)</span>", unsafe_allow_html=True)
    show_table(res['B_degrad_marge'], ['CA','Tx Marge %','Écart Tx Marge (pts)'],
               {'CA': fmt, 'Tx Marge %': lambda v: fmt_pct(v), 'Écart Tx Marge (pts)': fmt_delta})

    st.markdown(f"<span class='badge' style='background:#D7F5DE;color:#1A7A3A'>C · Performeurs — gain de marge en valeur</span>", unsafe_allow_html=True)
    show_table(res['C_perf_gain_marge'], ['CA','Gain Marge (FCFA)','Tx Marge %'],
               {'CA': fmt, 'Gain Marge (FCFA)': fmt, 'Tx Marge %': lambda v: fmt_pct(v)})

    st.markdown(f"<span class='badge' style='background:#D7F5DE;color:#1A7A3A'>D · Plus forte croissance de CA</span>", unsafe_allow_html=True)
    d4 = res['D_croissance_ca'].copy()
    if not d4.empty:
        d4['Évol CA %'] = (d4['CA']/d4['CA N-1']-1)*100
    show_table(d4, ['CA','CA N-1','Évol CA %','Tx Marge %'],
               {'CA': fmt, 'CA N-1': fmt, 'Évol CA %': lambda v: fmt_pct(v), 'Tx Marge %': lambda v: fmt_pct(v)})
    st.caption("⚠️ Une forte évolution % peut refléter un effet de base (article quasi absent en N-1) plutôt qu'une vraie dynamique.")

    st.markdown(f"<span class='badge' style='background:#FFD6D4;color:{RED}'>E · Plus forte baisse de CA</span>", unsafe_allow_html=True)
    show_table(res['E_baisse_ca'], ['CA','CA N-1','Perte CA (FCFA)'],
               {'CA': fmt, 'CA N-1': fmt, 'Perte CA (FCFA)': fmt})

    st.markdown(f"<span class='badge' style='background:#D7F5DE;color:#1A7A3A'>F · Plus forte hausse de quantité vendue</span>", unsafe_allow_html=True)
    show_table(res['F_hausse_qte'], ['Qté Vente','Qté Vente N-1','Variation Qté','CA'],
               {'Qté Vente': fmt, 'Qté Vente N-1': fmt, 'Variation Qté': lambda v: f"{v:+,.0f}".replace(",", " "), 'CA': fmt})

    st.markdown(f"<span class='badge' style='background:#FFD6D4;color:{RED}'>G · Plus forte baisse de quantité vendue</span>", unsafe_allow_html=True)
    show_table(res['G_baisse_qte'], ['Qté Vente','Qté Vente N-1','Variation Qté','CA'],
               {'Qté Vente': fmt, 'Qté Vente N-1': fmt, 'Variation Qté': lambda v: f"{v:+,.0f}".replace(",", " "), 'CA': fmt})
