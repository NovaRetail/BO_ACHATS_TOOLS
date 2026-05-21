"""
SmartBuyer Hub - Module Fidelite Cagnotte v2
============================================
Adapté à la nouvelle structure PBI plate :
  Col A : Site nom long
  Col B : Rayon  (ex: 00014 - EPICERIE, 00010 - BOISSONS, 00012 - PARFUMERIE HYGIENE...)
  Col C : Famille (ex: 00415 - CUISSON CULINAIRE)
  Col D : Article (ex: 15001234 - MON ARTICLE)
  Col E : CA  |  Col H : Marge  |  Col AB : Qté Vente

Totaux PBI :
  - Total PGC réseau : ligne où Site='Total', Rayon=NaN, Famille=NaN, Article=NaN
  - Total par Rayon  : ligne où Article=NaN et Famille=NaN et Rayon != 'Total' et Rayon non NaN
  - Lignes articles  : Article != 'Total' et Article non NaN et Article ne commence pas par code rayon
"""

import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# CHARTE SMARTBUYER HUB
# ============================================================
COULEUR_BLEU       = "007AFF"
COULEUR_GRIS_FOND  = "F2F2F7"
COULEUR_JAUNE_TOTAL = "FFF9E6"
COULEUR_BLEU_CLAIR = "E6F2FF"
COULEUR_TEXTE      = "1C1C1E"
COULEUR_ROUGE      = "D70015"
COULEUR_BORDURE    = "E5E5EA"
POLICE             = "Calibri"

MOIS_FR = {
    1:"Janvier", 2:"Fevrier", 3:"Mars", 4:"Avril",
    5:"Mai", 6:"Juin", 7:"Juillet", 8:"Aout",
    9:"Septembre", 10:"Octobre", 11:"Novembre", 12:"Decembre",
}
MOIS_NORMALIZE = {
    "janvier":"Janvier","fevrier":"Fevrier","mars":"Mars","avril":"Avril",
    "mai":"Mai","juin":"Juin","juillet":"Juillet","aout":"Aout",
    "septembre":"Septembre","octobre":"Octobre","novembre":"Novembre","decembre":"Decembre",
}

def normaliser_mois(s):
    if not isinstance(s, str): return ""
    s_clean = (s.strip().lower()
               .replace("é","e").replace("è","e").replace("ê","e")
               .replace("û","u").replace("à","a").replace("ç","c"))
    return MOIS_NORMALIZE.get(s_clean, s.strip().capitalize())

def normaliser_rayon(s):
    """Normalise le libellé rayon pour comparaison."""
    if not isinstance(s, str): return ""
    return (s.strip()
            .replace("é","e").replace("è","e").replace("ê","e")
            .replace("û","u").replace("à","a").replace("ç","c")
            .upper())

# ============================================================
# PARSING PBI — NOUVELLE STRUCTURE PLATE
# ============================================================
def extraire_periode_pbi_flat(df: pd.DataFrame) -> dict | None:
    """
    Cherche la ligne 'Filtres appliqués' dans la colonne 'Site nom long'
    et extrait la période.
    """
    pattern = re.compile(
        r"apr[eè]s\s+le\s+(\d{2}/\d{2}/\d{4})\s+et\s+est\s+avant\s+le\s+(\d{2}/\d{2}/\d{4})",
        re.IGNORECASE,
    )
    col = df["Site nom long"]
    for val in col.dropna():
        if "Filtres" in str(val):
            m = pattern.search(str(val))
            if m:
                d_deb = datetime.strptime(m.group(1), "%d/%m/%Y").date()
                d_fin_ex = datetime.strptime(m.group(2), "%d/%m/%Y").date()
                d_fin = (pd.Timestamp(d_fin_ex) - pd.Timedelta(days=1)).date()
                return {
                    "date_debut":  d_deb,
                    "date_fin":    d_fin,
                    "semaine":     f"S{d_fin.isocalendar().week:02d}",
                    "mois":        f"{MOIS_FR[d_fin.month]} {d_fin.year}",
                    "mois_court":  MOIS_FR[d_fin.month],
                }
    return None


def parser_pbi_flat(fichier) -> dict:
    """
    Parse un fichier PBI à structure plate.

    Retourne :
      - periode       : dict dates/semaine/mois
      - lignes        : DataFrame articles x magasins (avec Rayon, Famille, Code Article…)
      - totaux_pgc    : {CA, Marge, Qte} global réseau
      - totaux_rayon  : {rayon_norm: {CA, Marge, Qte, libelle}}
    """
    try:
        df = pd.read_excel(fichier, header=0, dtype=str)
    except Exception as e:
        return {"periode": None, "lignes": pd.DataFrame(),
                "totaux_pgc": None, "totaux_rayon": {},
                "erreur": str(e)}

    periode = extraire_periode_pbi_flat(df)
    if periode is None:
        return {"periode": None, "lignes": pd.DataFrame(),
                "totaux_pgc": None, "totaux_rayon": {},
                "erreur": "Bloc 'Filtres appliqués' introuvable"}

    def to_float(v):
        try:
            f = float(v)
            return f if f == f else 0.0
        except:
            return 0.0

    # --- Colonnes utiles ---
    # CA=col E (idx 4), Marge=col H (idx 7), Qte=col AB (idx 27)
    COL_CA    = df.columns[4]   # 'CA'
    COL_MARGE = df.columns[7]   # 'Marge'
    COL_QTE   = df.columns[27]  # 'Qté Vente'

    # Normalise les colonnes hiérarchiques
    df["_site"]    = df["Site nom long"].fillna("").astype(str).str.strip()
    df["_rayon"]   = df["Rayon"].fillna("").astype(str).str.strip()
    df["_famille"] = df["Famille"].fillna("").astype(str).str.strip()
    df["_article"] = df["Article"].fillna("").astype(str).str.strip()

    # --- Total PGC réseau ---
    # Ligne : site='Total', rayon='', famille='', article=''
    mask_pgc = (df["_site"] == "Total") & (df["_rayon"] == "") & \
               (df["_famille"] == "") & (df["_article"] == "")
    totaux_pgc = None
    if mask_pgc.any():
        row = df[mask_pgc].iloc[0]
        totaux_pgc = {
            "CA":   to_float(row[COL_CA]),
            "Marge":to_float(row[COL_MARGE]),
            "Qte":  to_float(row[COL_QTE]),
        }
    else:
        # Fallback : sommer les totaux site (rayon='Total')
        mask_site_tot = (
            df["_site"].str.match(r"^\d{4,6}\s*-\s*.+") &
            (df["_rayon"] == "Total") &
            (df["_famille"] == "") & (df["_article"] == "")
        )
        if mask_site_tot.any():
            totaux_pgc = {
                "CA":   df[mask_site_tot][COL_CA].apply(to_float).sum(),
                "Marge":df[mask_site_tot][COL_MARGE].apply(to_float).sum(),
                "Qte":  df[mask_site_tot][COL_QTE].apply(to_float).sum(),
            }

    # --- Totaux par Rayon (agrégés réseau, sommés sur tous les magasins) ---
    # Lignes : site = vrai magasin, rayon != 'Total', famille = 'Total', article = ''
    pat_site_real = re.compile(r"^\d{4,6}\s*-\s*.+")
    mask_rayon = (
        df["_site"].str.match(pat_site_real) &
        (df["_rayon"] != "") & (df["_rayon"] != "Total") &
        (df["_famille"] == "Total") & (df["_article"] == "")
    )
    totaux_rayon = {}
    for _, row in df[mask_rayon].iterrows():
        lib = row["_rayon"]
        rn  = normaliser_rayon(lib)
        ca  = to_float(row[COL_CA])
        mg  = to_float(row[COL_MARGE])
        qt  = to_float(row[COL_QTE])
        if rn in totaux_rayon:
            totaux_rayon[rn]["CA"]    += ca
            totaux_rayon[rn]["Marge"] += mg
            totaux_rayon[rn]["Qte"]   += qt
        else:
            totaux_rayon[rn] = {"CA": ca, "Marge": mg, "Qte": qt, "libelle": lib}

    # --- Lignes articles x magasins ---
    # Critères : site est un vrai magasin (non 'Total', non 'nan', non filtre)
    #            article non vide, non 'Total', commence par un code numérique
    pat_article = re.compile(r"^\d{7,9}\s*-\s*.+")
    pat_site    = re.compile(r"^\d{4,6}\s*-\s*.+")

    mask_articles = (
        df["_article"].str.match(pat_article) &
        df["_site"].str.match(pat_site)
    )

    df_art = df[mask_articles].copy()

    if df_art.empty:
        return {"periode": periode, "lignes": pd.DataFrame(),
                "totaux_pgc": totaux_pgc, "totaux_rayon": totaux_rayon,
                "erreur": None}

    # Extraction code article
    def split_code(s, sep="-"):
        parts = s.split(sep, 1)
        return parts[0].strip(), s.strip()

    df_art[["Code Article", "Article"]] = df_art["_article"].apply(
        lambda s: pd.Series(split_code(s))
    )
    df_art[["Code Site", "Magasin"]] = df_art["_site"].apply(
        lambda s: pd.Series(split_code(s))
    )

    df_art["Rayon"]         = df_art["_rayon"]
    df_art["Rayon_norm"]    = df_art["_rayon"].apply(normaliser_rayon)
    df_art["Sous Famille"]  = df_art["_famille"]
    df_art["CA"]    = df_art[COL_CA].apply(to_float)
    df_art["Marge"] = df_art[COL_MARGE].apply(to_float)
    df_art["Qte"]   = df_art[COL_QTE].apply(to_float)

    lignes = df_art[[
        "Rayon", "Rayon_norm", "Sous Famille",
        "Code Article", "Article",
        "Code Site", "Magasin",
        "CA", "Marge", "Qte",
    ]].reset_index(drop=True)

    return {
        "periode":      periode,
        "lignes":       lignes,
        "totaux_pgc":   totaux_pgc,
        "totaux_rayon": totaux_rayon,
        "erreur":       None,
    }


# ============================================================
# LISTE CSV
# ============================================================
def lire_liste_csv(fichier) -> pd.DataFrame:
    try:
        df = pd.read_csv(fichier, sep=";", dtype={"Article": str, "Cagnotte": float, "Mois": str})
    except Exception:
        fichier.seek(0)
        df = pd.read_csv(fichier, sep=",", dtype={"Article": str, "Cagnotte": float, "Mois": str})
    df["Article"]   = df["Article"].astype(str).str.strip()
    df["Mois_norm"] = df["Mois"].apply(normaliser_mois)
    return df


# ============================================================
# CONSTRUCTION RECAP
# ============================================================
def construire_recap_fichier(parsing: dict, liste_df: pd.DataFrame) -> dict:
    periode      = parsing["periode"]
    df_pbi       = parsing["lignes"]
    totaux_pgc   = parsing.get("totaux_pgc")
    totaux_rayon = parsing.get("totaux_rayon", {})

    vide = {
        "lignes_recap": pd.DataFrame(), "articles_attendus": [],
        "articles_trouves": [], "totaux_pgc": totaux_pgc,
        "totaux_rayon": totaux_rayon,
        "semaine": periode["semaine"] if periode else None,
        "mois":    periode["mois"]    if periode else None,
    }

    if periode is None or df_pbi.empty:
        return vide

    mois_court = periode["mois_court"]
    articles_mois   = liste_df[liste_df["Mois_norm"] == mois_court].copy()
    articles_codes  = articles_mois["Article"].tolist()
    articles_trouves = [a for a in articles_codes if a in df_pbi["Code Article"].values]

    df_filtre = df_pbi[df_pbi["Code Article"].isin(articles_codes)].copy()
    if df_filtre.empty:
        return {**vide, "articles_attendus": articles_codes}

    cag_map = dict(zip(articles_mois["Article"], articles_mois["Cagnotte"]))
    df_filtre["Cagnotte"]        = df_filtre["Code Article"].map(cag_map)
    df_filtre["Budget Cagnotte"] = df_filtre["Cagnotte"] * df_filtre["Qte"]
    df_filtre["Date Debut"]      = periode["date_debut"]
    df_filtre["Date Fin"]        = periode["date_fin"]
    df_filtre["Semaine"]         = periode["semaine"]
    df_filtre["Mois"]            = periode["mois"]

    df_filtre = df_filtre[[
        "Date Debut", "Date Fin", "Semaine", "Mois",
        "Rayon", "Rayon_norm", "Sous Famille",
        "Code Article", "Article",
        "Code Site", "Magasin",
        "CA", "Marge", "Qte", "Cagnotte", "Budget Cagnotte",
    ]]

    return {
        "lignes_recap":     df_filtre,
        "articles_attendus":articles_codes,
        "articles_trouves": articles_trouves,
        "totaux_pgc":       totaux_pgc,
        "totaux_rayon":     totaux_rayon,
        "semaine":          periode["semaine"],
        "mois":             periode["mois"],
    }


# ============================================================
# RECAP FINANCIER
# ============================================================
def construire_recap_financier(df_recap_all: pd.DataFrame,
                                totaux_rayon_par_semaine: dict) -> pd.DataFrame:
    """Agrège par Semaine x Article, calcule poids %CA/%Marge au RAYON."""
    cols = [
        "Semaine", "Mois", "Rayon", "Code Article", "Article",
        "Nb magasins", "Cagnotte/u", "Qte", "Budget cagnotte",
        "CA", "Marge", "CA Rayon", "Marge Rayon", "%CA", "%Marge",
    ]
    if df_recap_all.empty:
        return pd.DataFrame(columns=cols)

    grouped = df_recap_all.groupby(
        ["Semaine", "Mois", "Rayon", "Rayon_norm", "Code Article", "Article"],
        as_index=False
    ).agg(
        Nb_magasins=("Code Site", "nunique"),
        Cagnotte_u=("Cagnotte", "first"),
        CA=("CA", "sum"),
        Marge=("Marge", "sum"),
        Qte=("Qte", "sum"),
        Budget_cagnotte=("Budget Cagnotte", "sum"),
    ).rename(columns={
        "Nb_magasins": "Nb magasins",
        "Cagnotte_u":  "Cagnotte/u",
        "Budget_cagnotte": "Budget cagnotte",
    })

    def _get(row, key):
        sem = totaux_rayon_par_semaine.get(row["Semaine"], {})
        return sem.get(row["Rayon_norm"], {}).get(key, 0)

    grouped["CA Rayon"]    = grouped.apply(lambda r: _get(r, "CA"), axis=1)
    grouped["Marge Rayon"] = grouped.apply(lambda r: _get(r, "Marge"), axis=1)
    grouped["%CA"]         = grouped.apply(
        lambda r: r["CA"] / r["CA Rayon"] if r["CA Rayon"] else 0, axis=1)
    grouped["%Marge"]      = grouped.apply(
        lambda r: r["Marge"] / r["Marge Rayon"] if r["Marge Rayon"] else 0, axis=1)

    return grouped[cols]


# ============================================================
# STYLES OPENPYXL
# ============================================================
def _border():
    s = Side(style="thin", color=COULEUR_BORDURE)
    return Border(left=s, right=s, top=s, bottom=s)

def style_header(cell, couleur=None):
    cell.font      = Font(name=POLICE, size=10, bold=True, color="FFFFFF")
    cell.fill      = PatternFill("solid", fgColor=couleur or COULEUR_BLEU)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = _border()

def style_cell(cell, fond=None, bold=False, fmt=None, align="left"):
    cell.font      = Font(name=POLICE, size=10, bold=bold, color=COULEUR_TEXTE)
    if fond: cell.fill = PatternFill("solid", fgColor=fond)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if fmt: cell.number_format = fmt
    cell.border    = _border()


# ============================================================
# EXPORT EXCEL
# ============================================================
def exporter_excel(df_recap: pd.DataFrame, df_financier: pd.DataFrame,
                   kpis: dict, totaux_rayon_cumul: dict) -> BytesIO:
    wb = Workbook()

    # ── ONGLET 1 : RECAP ──────────────────────────────────────
    ws = wb.active
    ws.title = "Recap"

    headers = [
        "Date Debut","Date Fin","Semaine","Mois",
        "Rayon","Sous Famille",
        "Code Article","Article",
        "Code Site","Magasin",
        "CA","Marge","Qte","Cagnotte","Budget Cagnotte",
    ]
    widths = [12,12,9,14, 28,20, 13,36, 10,24, 13,13,10,11,15]
    for i,(h,w) in enumerate(zip(headers,widths),1):
        ws.column_dimensions[get_column_letter(i)].width = w
        c = ws.cell(row=1, column=i, value=h)
        style_header(c)
    ws.row_dimensions[1].height = 28

    if not df_recap.empty:
        df_s = df_recap.sort_values(["Rayon","Code Article","Magasin","Semaine"]).reset_index(drop=True)
        for r_idx, row in enumerate(df_s.itertuples(index=False), 2):
            vals = [
                row._4, row._5, row.Semaine, row.Mois,          # Date Debut/Fin/Sem/Mois
                row.Rayon, row._8,                               # Rayon, Sous Famille
                row._9, row.Article,                             # Code Art, Article
                row._11, row.Magasin,                            # Code Site, Magasin
                row.CA, row.Marge, row.Qte, row.Cagnotte, row._16,  # financier
            ]
            for c_idx, val in enumerate(vals, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                if c_idx in (1,2):   style_cell(cell, fond=COULEUR_BLEU_CLAIR, fmt="DD/MM/YYYY", align="center")
                elif c_idx in (3,4): style_cell(cell, fond=COULEUR_BLEU_CLAIR, align="center")
                elif c_idx == 5:     style_cell(cell, fond=COULEUR_BLEU_CLAIR, align="left", bold=True)
                elif c_idx in (11,12): style_cell(cell, fmt="#,##0;[Red]-#,##0", align="right")
                elif c_idx == 13:    style_cell(cell, fmt="#,##0.0", align="right")
                elif c_idx == 14:    style_cell(cell, fmt="#,##0", align="right")
                elif c_idx == 15:    style_cell(cell, fond=COULEUR_JAUNE_TOTAL, fmt="#,##0", align="right", bold=True)
                else:                style_cell(cell)

    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    # ── ONGLET 2 : RECAP FINANCIER ────────────────────────────
    ws2 = wb.create_sheet("Recap Financier")

    ws2["A1"] = "RECAP FINANCIER · FIDELITE CAGNOTTE"
    ws2["A1"].font = Font(name=POLICE, size=14, bold=True, color=COULEUR_BLEU)
    ws2.merge_cells("A1:G1")

    # KPIs globaux
    ws2["A3"] = "Indicateurs globaux"
    ws2["A3"].font = Font(name=POLICE, size=11, bold=True, color=COULEUR_TEXTE)

    kpi_items = [
        ("Budget cagnotte total",         kpis.get("budget_total",0),   "#,##0"),
        ("CA articles fidélité",          kpis.get("ca_total",0),       "#,##0"),
        ("Marge articles fidélité",       kpis.get("marge_total",0),    "#,##0;[Red]-#,##0"),
        ("Qté articles fidélité",         kpis.get("qte_total",0),      "#,##0.0"),
        ("CA total PGC (référence)",      kpis.get("ca_pgc_total",0),   "#,##0"),
        ("Marge totale PGC (référence)",  kpis.get("marge_pgc_total",0),"#,##0;[Red]-#,##0"),
        ("% CA fidélité / PGC",           kpis.get("pct_ca_pgc",0),     "0.00%"),
        ("% Marge fidélité / PGC",        kpis.get("pct_marge_pgc",0),  "0.00%"),
    ]
    for i,(lbl,val,fmt) in enumerate(kpi_items):
        r = 4+i
        ws2.cell(row=r, column=1, value=lbl).font = Font(name=POLICE, size=10, color=COULEUR_TEXTE)
        c = ws2.cell(row=r, column=2, value=val)
        c.font          = Font(name=POLICE, size=11, bold=True, color=COULEUR_BLEU)
        c.fill          = PatternFill("solid", fgColor=COULEUR_JAUNE_TOTAL if "%" in lbl else COULEUR_GRIS_FOND)
        c.number_format = fmt
        c.alignment     = Alignment(horizontal="right")

    # Poids par RAYON
    row_rt = 14
    ws2.cell(row=row_rt, column=1, value="Poids fidélité par RAYON").font = \
        Font(name=POLICE, size=11, bold=True, color=COULEUR_TEXTE)

    headers_r = ["Rayon","CA Fidélité","CA Rayon","%CA","Marge Fidélité","Marge Rayon","%Marge"]
    widths_r  = [30,16,16,10,16,16,10]
    for i,(h,w) in enumerate(zip(headers_r,widths_r),1):
        ws2.column_dimensions[get_column_letter(i)].width = w
        c = ws2.cell(row=row_rt+1, column=i, value=h)
        style_header(c)
    ws2.row_dimensions[row_rt+1].height = 28

    row_r = row_rt+2
    if not df_financier.empty:
        fam_art = df_financier.groupby("Rayon", as_index=False).agg(
            CA_fid=("CA","sum"), Marge_fid=("Marge","sum"))
        rayon_cumul = {}
        for sem, dico in totaux_rayon_cumul.items():
            for rn, tot in dico.items():
                lib = tot.get("libelle", rn)
                if lib not in rayon_cumul:
                    rayon_cumul[lib] = {"CA":0,"Marge":0}
                rayon_cumul[lib]["CA"]    += tot.get("CA",0)
                rayon_cumul[lib]["Marge"] += tot.get("Marge",0)

        for _, r in fam_art.iterrows():
            ray    = r["Rayon"]
            ca_f   = r["CA_fid"]; mg_f = r["Marge_fid"]
            ca_r   = rayon_cumul.get(ray,{}).get("CA",0)
            mg_r   = rayon_cumul.get(ray,{}).get("Marge",0)
            pct_ca = ca_f/ca_r if ca_r else 0
            pct_mg = mg_f/mg_r if mg_r else 0
            ws2.cell(row=row_r,column=1,value=ray)
            ws2.cell(row=row_r,column=2,value=ca_f)
            ws2.cell(row=row_r,column=3,value=ca_r)
            ws2.cell(row=row_r,column=4,value=pct_ca)
            ws2.cell(row=row_r,column=5,value=mg_f)
            ws2.cell(row=row_r,column=6,value=mg_r)
            ws2.cell(row=row_r,column=7,value=pct_mg)
            for ci in range(1,8):
                cell = ws2.cell(row=row_r, column=ci)
                if ci==1:      style_cell(cell,fond=COULEUR_BLEU_CLAIR,bold=True)
                elif ci in (4,7): style_cell(cell,fond=COULEUR_JAUNE_TOTAL,fmt="0.00%",align="right",bold=True)
                else:          style_cell(cell,fmt="#,##0;[Red]-#,##0",align="right")
            row_r += 1

    # Détail article x semaine
    row_t2 = row_r+2
    ws2.cell(row=row_t2,column=1,value="Détail par article et semaine").font = \
        Font(name=POLICE,size=11,bold=True,color=COULEUR_TEXTE)

    headers_f = [
        "Semaine","Mois","Rayon","Code Article","Article",
        "Nb magasins","Cagnotte/u","Qté","Budget cagnotte",
        "CA","Marge","CA Rayon","Marge Rayon","%CA","%Marge",
    ]
    widths_f = [10,14,28,13,36,12,12,11,16,14,14,15,15,9,9]
    for i,(h,w) in enumerate(zip(headers_f,widths_f),1):
        ws2.column_dimensions[get_column_letter(i)].width = max(
            ws2.column_dimensions[get_column_letter(i)].width or 0, w)
        c = ws2.cell(row=row_t2+1, column=i, value=h)
        style_header(c)
    ws2.row_dimensions[row_t2+1].height = 28

    row_d = row_t2+2
    if not df_financier.empty:
        df_fs = df_financier.sort_values(["Semaine","Rayon","Code Article"]).reset_index(drop=True)
        for _, r in df_fs.iterrows():
            vals = [
                r["Semaine"],r["Mois"],r["Rayon"],r["Code Article"],r["Article"],
                r["Nb magasins"],r["Cagnotte/u"],r["Qte"],r["Budget cagnotte"],
                r["CA"],r["Marge"],r["CA Rayon"],r["Marge Rayon"],r["%CA"],r["%Marge"],
            ]
            for ci,val in enumerate(vals,1):
                cell = ws2.cell(row=row_d, column=ci, value=val)
                if ci in (1,2,3): style_cell(cell,fond=COULEUR_BLEU_CLAIR,align="center" if ci<3 else "left",bold=(ci==3))
                elif ci==6:       style_cell(cell,fmt="#,##0",align="center")
                elif ci==7:       style_cell(cell,fmt="#,##0",align="right")
                elif ci==8:       style_cell(cell,fmt="#,##0.0",align="right")
                elif ci==9:       style_cell(cell,fond=COULEUR_JAUNE_TOTAL,fmt="#,##0",align="right",bold=True)
                elif ci in (10,11): style_cell(cell,fmt="#,##0;[Red]-#,##0",align="right")
                elif ci in (12,13): style_cell(cell,fond=COULEUR_GRIS_FOND,fmt="#,##0;[Red]-#,##0",align="right")
                elif ci in (14,15): style_cell(cell,fond=COULEUR_BLEU_CLAIR,fmt="0.00%",align="right",bold=True)
                else:             style_cell(cell)
            row_d += 1

        # Ligne total
        rT = row_d
        ws2.cell(row=rT,column=1,value="TOTAL")
        ws2.cell(row=rT,column=5,value=f"{len(df_fs)} ligne(s)")
        rh1 = row_t2+2
        for ci,col_letter,fmt,align in [
            (6,"F","#,##0","center"),(8,"H","#,##0.0","right"),
            (9,"I","#,##0","right"),(10,"J","#,##0;[Red]-#,##0","right"),
            (11,"K","#,##0;[Red]-#,##0","right"),
        ]:
            cell = ws2.cell(row=rT, column=ci,
                value=f"=SUM({col_letter}{rh1}:{col_letter}{row_d-1})")
            style_cell(cell,fond=COULEUR_JAUNE_TOTAL,bold=True,fmt=fmt,align=align)
        for ci,val in [(14,kpis.get("pct_ca_pgc",0)),(15,kpis.get("pct_marge_pgc",0))]:
            cell = ws2.cell(row=rT, column=ci, value=val)
            style_cell(cell,fond=COULEUR_JAUNE_TOTAL,bold=True,fmt="0.00%",align="right")
        for ci in [1,5]:
            cell = ws2.cell(row=rT, column=ci)
            style_cell(cell,fond=COULEUR_JAUNE_TOTAL,bold=True,align="center" if ci==1 else "left")

    ws2.freeze_panes = f"A{row_t2+2}"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ============================================================
# INTERFACE STREAMLIT
# ============================================================
def render_fidelite_cagnotte():
    st.markdown("""
    <style>
        .stApp { font-family: -apple-system,'SF Pro Display',Calibri,sans-serif; }
        .hdr { display:flex;align-items:center;gap:12px;padding:16px 0;margin-bottom:20px;
               border-bottom:1px solid #E5E5EA; }
        .hdr .ico { width:40px;height:40px;border-radius:10px;background:#007AFF;
                    display:flex;align-items:center;justify-content:center;
                    color:white;font-size:20px; }
        .hdr .t { font-size:20px;font-weight:600;color:#1C1C1E;margin:0; }
        .hdr .s { font-size:13px;color:#8E8E93;margin:0; }
        .pill  { background:#F2F2F7;border-radius:8px;padding:8px 12px;
                 margin-bottom:6px;font-size:13px; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="hdr">
      <div class="ico">$</div>
      <div>
        <p class="t">Fidélité Cagnotte</p>
        <p class="s">Suivi hebdomadaire · Investissement vs Performance (poids au Rayon)</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Étape 1 : Upload ──────────────────────────────────────
    st.markdown("**1 · Charger les fichiers**")
    col1, col2 = st.columns(2)
    with col1:
        fichiers_pbi = st.file_uploader(
            "Extractions PBI (xlsx, plusieurs possibles)",
            type=["xlsx"], accept_multiple_files=True,
            key="fidelite_cagnotte_files",
        )
    with col2:
        fichier_liste = st.file_uploader(
            "Liste articles (csv — colonnes : Article ; Cagnotte ; Mois)",
            type=["csv"], key="fidelite_cagnotte_liste",
        )

    if not fichiers_pbi or not fichier_liste:
        st.info("Charge au moins un fichier PBI (.xlsx) et le CSV liste pour démarrer.")
        return

    # ── Étape 2 : Parsing ─────────────────────────────────────
    liste_df = lire_liste_csv(fichier_liste)
    parsings = []
    semaines_vues = {}

    for f in fichiers_pbi:
        p = parser_pbi_flat(f)
        if p["periode"] is None:
            st.error(f"`{f.name}` : impossible de détecter la période. {p.get('erreur','')}")
            continue
        sem = p["periode"]["semaine"]
        if sem in semaines_vues:
            st.warning(f"Doublon : semaine **{sem}** dans `{semaines_vues[sem]}` et `{f.name}` — les deux seront empilés.")
        else:
            semaines_vues[sem] = f.name
        parsings.append({"fichier": f.name, "parsing": p})

    if not parsings:
        st.error("Aucun fichier PBI exploitable.")
        return

    # Affichage périodes
    st.markdown("**2 · Périodes et rayons détectés**")
    for item in parsings:
        per  = item["parsing"]["periode"]
        rpts = list(item["parsing"].get("totaux_rayon", {}).keys())
        rpts_str = ", ".join(rpts[:6]) + ("…" if len(rpts)>6 else "") if rpts else "aucun rayon"
        st.markdown(f"""
        <div class="pill">
          <strong>{item['fichier']}</strong> · {per['date_debut'].strftime('%d/%m/%Y')} → {per['date_fin'].strftime('%d/%m/%Y')}
          · <strong>{per['semaine']}</strong> · {per['mois']}
          <br><span style="color:#8E8E93;font-size:12px;">Rayons : {rpts_str}</span>
        </div>""", unsafe_allow_html=True)

    # ── Étape 3 : Recap ───────────────────────────────────────
    recap_parts = []
    manquants_global = {}
    totaux_pgc_par_sem = {}
    totaux_rayon_par_sem = {}

    for item in parsings:
        res = construire_recap_fichier(item["parsing"], liste_df)
        if not res["lignes_recap"].empty:
            recap_parts.append(res["lignes_recap"])
        manq = set(res["articles_attendus"]) - set(res["articles_trouves"])
        if manq:
            manquants_global[f"{item['fichier']} ({res['semaine']})"] = list(manq)

        sem = res.get("semaine")
        if sem:
            pgc = res.get("totaux_pgc") or {}
            if sem in totaux_pgc_par_sem:
                for k in ("CA","Marge","Qte"):
                    totaux_pgc_par_sem[sem][k] += pgc.get(k,0)
            else:
                totaux_pgc_par_sem[sem] = {k: pgc.get(k,0) for k in ("CA","Marge","Qte")}

            ray = res.get("totaux_rayon", {})
            if sem not in totaux_rayon_par_sem:
                totaux_rayon_par_sem[sem] = {}
            for rn, vals in ray.items():
                if rn in totaux_rayon_par_sem[sem]:
                    for k in ("CA","Marge","Qte"):
                        totaux_rayon_par_sem[sem][rn][k] += vals.get(k,0)
                else:
                    totaux_rayon_par_sem[sem][rn] = dict(vals)

    df_recap_all = pd.concat(recap_parts, ignore_index=True) if recap_parts else pd.DataFrame()

    if manquants_global:
        nb = sum(len(v) for v in manquants_global.values())
        with st.expander(f"⚠ Articles non trouvés dans certaines extractions ({nb})"):
            for fic, arts in manquants_global.items():
                st.markdown(f"**{fic}** : {', '.join(arts)}")

    if df_recap_all.empty:
        st.warning("Aucun article de la liste n'a été trouvé dans les extractions PBI pour les mois concernés.")
        return

    if not any(totaux_rayon_par_sem.values()):
        st.warning("⚠ Aucun total par rayon détecté — les poids %CA/%Marge seront à 0. Vérifie la structure du fichier.")

    # ── Étape 4 : KPIs ────────────────────────────────────────
    df_financier = construire_recap_financier(df_recap_all, totaux_rayon_par_sem)

    ca_pgc   = sum(v.get("CA",0)    for v in totaux_pgc_par_sem.values())
    marge_pgc = sum(v.get("Marge",0) for v in totaux_pgc_par_sem.values())
    ca_fid   = df_financier["CA"].sum()    if not df_financier.empty else 0
    marge_fid = df_financier["Marge"].sum() if not df_financier.empty else 0

    kpis = {
        "budget_total":   df_financier["Budget cagnotte"].sum() if not df_financier.empty else 0,
        "ca_total":       ca_fid,
        "marge_total":    marge_fid,
        "qte_total":      df_financier["Qte"].sum() if not df_financier.empty else 0,
        "ca_pgc_total":   ca_pgc,
        "marge_pgc_total":marge_pgc,
        "pct_ca_pgc":     ca_fid/ca_pgc if ca_pgc else 0,
        "pct_marge_pgc":  marge_fid/marge_pgc if marge_pgc else 0,
    }

    st.markdown("**3 · Synthèse globale**")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Budget cagnotte",       f"{kpis['budget_total']:,.0f}".replace(","," "))
    c2.metric("CA articles fidélité",  f"{kpis['ca_total']:,.0f}".replace(","," "))
    c3.metric("Marge articles fidélité",f"{kpis['marge_total']:,.0f}".replace(","," "))
    c4.metric("Qté vendue",            f"{kpis['qte_total']:,.1f}".replace(","," "))
    c5,c6,c7,c8 = st.columns(4)
    c5.metric("CA total PGC",         f"{kpis['ca_pgc_total']:,.0f}".replace(","," "))
    c6.metric("Marge totale PGC",     f"{kpis['marge_pgc_total']:,.0f}".replace(","," "))
    c7.metric("% CA fid. / PGC",      f"{kpis['pct_ca_pgc']*100:.2f}%")
    c8.metric("% Marge fid. / PGC",   f"{kpis['pct_marge_pgc']*100:.2f}%")

    # Poids par rayon
    st.markdown("**4 · Poids fidélité par RAYON**")
    rayon_cumul = {}
    for sem, dico in totaux_rayon_par_sem.items():
        for rn, tot in dico.items():
            lib = tot.get("libelle", rn)
            if lib not in rayon_cumul:
                rayon_cumul[lib] = {"CA":0,"Marge":0}
            rayon_cumul[lib]["CA"]    += tot.get("CA",0)
            rayon_cumul[lib]["Marge"] += tot.get("Marge",0)

    if not df_financier.empty and rayon_cumul:
        ray_art = df_financier.groupby("Rayon", as_index=False).agg(
            CA_fid=("CA","sum"), Marge_fid=("Marge","sum"))
        rows = []
        for _, r in ray_art.iterrows():
            ray = r["Rayon"]
            ca_r = rayon_cumul.get(ray,{}).get("CA",0)
            mg_r = rayon_cumul.get(ray,{}).get("Marge",0)
            rows.append({
                "Rayon": ray,
                "CA Fidélité": r["CA_fid"],
                "CA Rayon": ca_r,
                "%CA": r["CA_fid"]/ca_r if ca_r else 0,
                "Marge Fidélité": r["Marge_fid"],
                "Marge Rayon": mg_r,
                "%Marge": r["Marge_fid"]/mg_r if mg_r else 0,
            })
        st.dataframe(
            pd.DataFrame(rows).style.format({
                "CA Fidélité":"{:,.0f}","CA Rayon":"{:,.0f}","%CA":"{:.2%}",
                "Marge Fidélité":"{:,.0f}","Marge Rayon":"{:,.0f}","%Marge":"{:.2%}",
            }),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("Pas de totaux rayon disponibles.")

    st.markdown("**5 · Aperçu Recap Financier**")
    st.dataframe(df_financier, use_container_width=True, hide_index=True)

    with st.expander("Voir le Recap détaillé (Article × Magasin)"):
        df_aff = df_recap_all.drop(columns=["Rayon_norm"], errors="ignore")
        st.dataframe(df_aff, use_container_width=True, hide_index=True)

    # ── Étape 6 : Export ──────────────────────────────────────
    st.markdown("**6 · Télécharger**")
    buf = exporter_excel(df_recap_all, df_financier, kpis, totaux_rayon_par_sem)
    mois_u = df_recap_all["Mois"].unique() if not df_recap_all.empty else []
    nom = f"Fidelite_Cagnotte_{mois_u[0].replace(' ','_')}.xlsx" if len(mois_u)==1 \
          else "Fidelite_Cagnotte_multi_mois.xlsx"
    st.download_button(
        label=f"📥 Télécharger {nom}",
        data=buf, file_name=nom,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Point d'entrée standalone ──────────────────────────────
if __name__ == "__main__":
    st.set_page_config(page_title="Fidélité Cagnotte", layout="wide")
    render_fidelite_cagnotte()
