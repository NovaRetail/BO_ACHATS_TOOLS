"""
SmartBuyer Hub - Module Fidelite Cagnotte
=====================================
Suivi hebdomadaire des cagnottes article x magasin a partir d'extractions Power BI.

Fonctionnalites:
- Multi-upload de fichiers PBI (.xlsx)
- 1 fichier CSV liste articles (Article, Cagnotte, Mois)
- Detection auto des dates / semaine / mois dans le bloc "Filtres appliques"
- Filtrage par mois propre a chaque fichier
- Empilage multi-semaines
- Export Excel 2 onglets : Recap detaille + Recap Financier
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
COULEUR_BLEU = "007AFF"
COULEUR_GRIS_FOND = "F2F2F7"
COULEUR_JAUNE_TOTAL = "FFF9E6"
COULEUR_BLEU_CLAIR = "E6F2FF"
COULEUR_TEXTE = "1C1C1E"
COULEUR_ROUGE = "D70015"
COULEUR_BORDURE = "E5E5EA"
POLICE = "Calibri"

MOIS_FR = {
    1: "Janvier", 2: "Fevrier", 3: "Mars", 4: "Avril",
    5: "Mai", 6: "Juin", 7: "Juillet", 8: "Aout",
    9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Decembre",
}

# Normalisation pour comparaison avec la colonne Mois du CSV (qui peut avoir accents)
MOIS_NORMALIZE = {
    "janvier": "Janvier", "fevrier": "Fevrier", "fevrier": "Fevrier", "fevrier": "Fevrier",
    "mars": "Mars", "avril": "Avril", "mai": "Mai", "juin": "Juin",
    "juillet": "Juillet", "aout": "Aout", "aout": "Aout",
    "septembre": "Septembre", "octobre": "Octobre", "novembre": "Novembre",
    "decembre": "Decembre", "decembre": "Decembre",
}


def normaliser_mois(s: str) -> str:
    """Normalise un libelle mois (gere accents et casse)."""
    if not isinstance(s, str):
        return ""
    s_clean = (s.strip().lower()
               .replace("é", "e").replace("è", "e").replace("ê", "e")
               .replace("û", "u").replace("à", "a").replace("ç", "c"))
    return MOIS_NORMALIZE.get(s_clean, s.strip().capitalize())


# ============================================================
# PARSING PBI
# ============================================================
def extraire_periode_pbi(ws) -> dict:
    """
    Cherche dans toutes les cellules du fichier le bloc 'Filtres appliques'
    et extrait Date Debut / Date Fin.
    Pattern attendu: "Date est le ou apres le DD/MM/YYYY et est avant le DD/MM/YYYY"
    """
    pattern = re.compile(
        r"apr[eè]s\s+le\s+(\d{2}/\d{2}/\d{4})\s+et\s+est\s+avant\s+le\s+(\d{2}/\d{2}/\d{4})",
        re.IGNORECASE,
    )
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "Filtres appliqu" in cell.value:
                m = pattern.search(cell.value)
                if m:
                    d_deb = datetime.strptime(m.group(1), "%d/%m/%Y").date()
                    d_fin_exclusive = datetime.strptime(m.group(2), "%d/%m/%Y").date()
                    # Le filtre PBI est "avant le" donc la vraie fin est j-1
                    d_fin = pd.Timestamp(d_fin_exclusive) - pd.Timedelta(days=1)
                    d_fin = d_fin.date()
                    return {
                        "date_debut": d_deb,
                        "date_fin": d_fin,
                        "semaine": f"S{d_fin.isocalendar().week:02d}",
                        "mois": f"{MOIS_FR[d_fin.month]} {d_fin.year}",
                        "mois_court": MOIS_FR[d_fin.month],
                    }
    return None


def parser_pbi(fichier) -> dict:
    """
    Parse un fichier PBI et retourne :
    - periode : dict avec dates / semaine / mois
    - lignes : DataFrame avec les lignes article x magasin
    
    Logique :
    - Colonnes : A=Famille, B=Rayon, C=Sous Famille, D=Article, E=Site nom long,
                 F=CA, I=Marge, AC=Qte Vente
    - Lignes article : colonne D contient un libelle (ex "22001928 - KG CITRON MEYER")
                       et colonne E = "Total"
    - Lignes magasin : colonne E contient "CODE - Nom Magasin" (ex "10301 - Hyper Marcory")
    - On garde le contexte Rayon (B) et Sous Famille (C) en remontant les niveaux
    """
    wb = load_workbook(fichier, data_only=True)
    ws = wb.active
    
    periode = extraire_periode_pbi(ws)
    if periode is None:
        return {"periode": None, "lignes": pd.DataFrame(), "totaux_pgc": None, "erreur": "Bloc 'Filtres appliques' introuvable"}
    
    lignes = []
    totaux_pgc = None  # CA / Marge / Qte total perimetre PGC (lue sur la 1ere ligne Total)
    rayon_courant = None
    sous_famille_courante = None
    article_code_courant = None
    article_libelle_courant = None
    
    # Regex pour detecter un code article (en colonne D) : "12345678 - Libelle" ou "12345678"
    pattern_article = re.compile(r"^(\d{6,9})\s*[-–]?\s*(.*)$")
    # Regex pour detecter un site (en colonne E) : "10301 - Hyper Marcory"
    pattern_site = re.compile(r"^(\d{4,6})\s*[-–]\s*(.+)$")
    
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None for v in row):
            continue
        
        famille = row[0]
        rayon = row[1]
        sous_famille = row[2]
        article = row[3]
        site = row[4]
        ca = row[5]
        marge = row[8]
        qte = row[28] if len(row) > 28 else None
        
        # Capture des totaux PGC : 1ere ligne ou Rayon = "Total" (niveau le plus haut du tableau)
        if totaux_pgc is None and isinstance(rayon, str) and rayon.strip() == "Total":
            totaux_pgc = {
                "CA": ca if ca is not None else 0,
                "Marge": marge if marge is not None else 0,
                "Qte": qte if qte is not None else 0,
            }
        
        # Maj contexte Rayon
        if isinstance(rayon, str) and rayon != "Total" and rayon.strip():
            rayon_courant = rayon.strip()
        
        # Maj contexte Sous Famille
        if isinstance(sous_famille, str) and sous_famille != "Total" and sous_famille.strip():
            sous_famille_courante = sous_famille.strip()
        
        # Detection ligne article (D rempli, E = Total)
        if isinstance(article, str) and article.strip() and article.strip() != "Total":
            m = pattern_article.match(article.strip())
            if m and (site == "Total" or site is None):
                article_code_courant = m.group(1)
                article_libelle_courant = article.strip()
                continue
        
        # Detection ligne magasin (D vide, E = code-nom)
        if isinstance(site, str) and article_code_courant:
            m_site = pattern_site.match(site.strip())
            if m_site:
                code_site = m_site.group(1)
                nom_magasin = m_site.group(2).strip()
                lignes.append({
                    "Rayon": rayon_courant or "",
                    "Sous Famille": sous_famille_courante or "",
                    "Code Article": article_code_courant,
                    "Article": article_libelle_courant,
                    "Code Site": code_site,
                    "Magasin": nom_magasin,
                    "CA": ca if ca is not None else 0,
                    "Marge": marge if marge is not None else 0,
                    "Qte": qte if qte is not None else 0,
                })
    
    df = pd.DataFrame(lignes)
    return {"periode": periode, "lignes": df, "totaux_pgc": totaux_pgc, "erreur": None}


# ============================================================
# CONSTRUCTION DU RECAP
# ============================================================
def lire_liste_csv(fichier) -> pd.DataFrame:
    """Lit le CSV liste articles (separateur ; UTF-8)."""
    try:
        df = pd.read_csv(fichier, sep=";", dtype={"Article": str, "Cagnotte": float, "Mois": str})
    except Exception:
        fichier.seek(0)
        df = pd.read_csv(fichier, sep=",", dtype={"Article": str, "Cagnotte": float, "Mois": str})
    df["Article"] = df["Article"].astype(str).str.strip()
    df["Mois_norm"] = df["Mois"].apply(normaliser_mois)
    return df


def construire_recap_fichier(parsing: dict, liste_df: pd.DataFrame) -> dict:
    """
    Pour un fichier PBI parse, applique le filtre articles de la liste sur le mois,
    et construit les lignes du recap.
    Retourne :
    - lignes_recap : DataFrame des lignes article x magasin
    - articles_attendus / articles_trouves
    - totaux_pgc : dict {CA, Marge, Qte} totaux du perimetre PGC pour la semaine
    - semaine, mois : pour identifier la cle
    """
    periode = parsing["periode"]
    df_pbi = parsing["lignes"]
    totaux_pgc = parsing.get("totaux_pgc")
    
    if periode is None or df_pbi.empty:
        return {
            "lignes_recap": pd.DataFrame(),
            "articles_attendus": [], "articles_trouves": [],
            "totaux_pgc": totaux_pgc,
            "semaine": periode["semaine"] if periode else None,
            "mois": periode["mois"] if periode else None,
        }
    
    mois_court = periode["mois_court"]
    
    # Filtre liste articles sur le mois
    articles_mois = liste_df[liste_df["Mois_norm"] == mois_court].copy()
    articles_codes = articles_mois["Article"].tolist()
    
    # Articles presents dans le PBI
    articles_pbi = df_pbi["Code Article"].unique().tolist()
    articles_trouves = [a for a in articles_codes if a in articles_pbi]
    
    # Jointure : on garde les lignes du PBI dont le code article est dans la liste filtree
    df_filtre = df_pbi[df_pbi["Code Article"].isin(articles_codes)].copy()
    
    if df_filtre.empty:
        return {
            "lignes_recap": pd.DataFrame(),
            "articles_attendus": articles_codes,
            "articles_trouves": [],
            "totaux_pgc": totaux_pgc,
            "semaine": periode["semaine"],
            "mois": periode["mois"],
        }
    
    # Ajout de la cagnotte unitaire
    cag_map = dict(zip(articles_mois["Article"], articles_mois["Cagnotte"]))
    df_filtre["Cagnotte"] = df_filtre["Code Article"].map(cag_map)
    
    # Budget cagnotte par ligne magasin = Cagnotte unitaire x Qte vendue dans ce magasin
    df_filtre["Budget Cagnotte"] = df_filtre["Cagnotte"] * df_filtre["Qte"]
    
    # Ajout des colonnes temporelles
    df_filtre["Date Debut"] = periode["date_debut"]
    df_filtre["Date Fin"] = periode["date_fin"]
    df_filtre["Semaine"] = periode["semaine"]
    df_filtre["Mois"] = periode["mois"]
    
    # Reordonnancement
    df_filtre = df_filtre[[
        "Date Debut", "Date Fin", "Semaine", "Mois",
        "Rayon", "Sous Famille", "Code Article", "Article",
        "Code Site", "Magasin",
        "CA", "Marge", "Qte", "Cagnotte", "Budget Cagnotte",
    ]]
    
    return {
        "lignes_recap": df_filtre,
        "articles_attendus": articles_codes,
        "articles_trouves": articles_trouves,
        "totaux_pgc": totaux_pgc,
        "semaine": periode["semaine"],
        "mois": periode["mois"],
    }


# ============================================================
# EXPORT EXCEL
# ============================================================
def style_header(cell, couleur_fond=None):
    cell.font = Font(name=POLICE, size=10, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=couleur_fond or COULEUR_BLEU)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(
        left=Side(style="thin", color=COULEUR_BORDURE),
        right=Side(style="thin", color=COULEUR_BORDURE),
        top=Side(style="thin", color=COULEUR_BORDURE),
        bottom=Side(style="thin", color=COULEUR_BORDURE),
    )


def style_cell(cell, fond=None, bold=False, fmt=None, align="left"):
    cell.font = Font(name=POLICE, size=10, bold=bold, color=COULEUR_TEXTE)
    if fond:
        cell.fill = PatternFill("solid", fgColor=fond)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if fmt:
        cell.number_format = fmt
    cell.border = Border(
        left=Side(style="thin", color=COULEUR_BORDURE),
        right=Side(style="thin", color=COULEUR_BORDURE),
        top=Side(style="thin", color=COULEUR_BORDURE),
        bottom=Side(style="thin", color=COULEUR_BORDURE),
    )


def exporter_excel(df_recap: pd.DataFrame, df_financier: pd.DataFrame, kpis: dict) -> BytesIO:
    """Construit le fichier Excel 2 onglets : Recap + Recap Financier."""
    wb = Workbook()
    
    # ===== ONGLET 1 : RECAP =====
    ws = wb.active
    ws.title = "Recap"
    
    headers = [
        "Date Debut", "Date Fin", "Semaine", "Mois",
        "Rayon", "Sous Famille", "Code Article", "Article",
        "Code Site", "Magasin",
        "CA", "Marge", "Qte", "Cagnotte", "Budget Cagnotte",
    ]
    
    # Largeurs colonnes
    widths = [12, 12, 9, 14, 16, 18, 13, 36, 10, 24, 13, 13, 10, 11, 15]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    # Header
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        style_header(cell)
    ws.row_dimensions[1].height = 28
    
    # Donnees
    if not df_recap.empty:
        # Tri Article > Magasin > Semaine
        df_sorted = df_recap.sort_values(by=["Code Article", "Magasin", "Semaine"]).reset_index(drop=True)
        
        row_excel = 2
        for _, row in df_sorted.iterrows():
            values = [
                row["Date Debut"], row["Date Fin"], row["Semaine"], row["Mois"],
                row["Rayon"], row["Sous Famille"], row["Code Article"], row["Article"],
                row["Code Site"], row["Magasin"],
                row["CA"], row["Marge"], row["Qte"], row["Cagnotte"], row["Budget Cagnotte"],
            ]
            for col_idx, val in enumerate(values, start=1):
                cell = ws.cell(row=row_excel, column=col_idx, value=val)
                # Format selon colonne
                if col_idx in (1, 2):  # dates
                    style_cell(cell, fond=COULEUR_BLEU_CLAIR, fmt="DD/MM/YYYY", align="center")
                elif col_idx == 3:  # semaine
                    style_cell(cell, fond=COULEUR_BLEU_CLAIR, align="center")
                elif col_idx == 4:  # mois
                    style_cell(cell, fond=COULEUR_BLEU_CLAIR, align="center")
                elif col_idx in (11, 12):  # CA, Marge
                    style_cell(cell, fmt="#,##0;[Red]-#,##0", align="right")
                elif col_idx == 13:  # Qte
                    style_cell(cell, fmt="#,##0.0", align="right")
                elif col_idx == 14:  # Cagnotte unitaire
                    style_cell(cell, fmt="#,##0", align="right")
                elif col_idx == 15:  # Budget Cagnotte
                    style_cell(cell, fond=COULEUR_JAUNE_TOTAL, fmt="#,##0", align="right", bold=True)
                else:
                    style_cell(cell, align="left")
            row_excel += 1
    
    # Freeze + autofilter
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    
    # ===== ONGLET 2 : RECAP FINANCIER =====
    ws2 = wb.create_sheet("Recap Financier")
    
    # Bloc KPIs
    ws2.column_dimensions["A"].width = 26
    ws2.column_dimensions["B"].width = 18
    ws2.column_dimensions["C"].width = 2
    ws2.column_dimensions["D"].width = 26
    ws2.column_dimensions["E"].width = 18
    
    # Titre
    ws2["A1"] = "RECAP FINANCIER · FIDELITE CAGNOTTE"
    ws2["A1"].font = Font(name=POLICE, size=14, bold=True, color=COULEUR_BLEU)
    ws2.merge_cells("A1:E1")
    
    ws2["A3"] = "Indicateurs globaux"
    ws2["A3"].font = Font(name=POLICE, size=11, bold=True, color=COULEUR_TEXTE)
    
    kpi_items = [
        ("Budget cagnotte total", kpis.get("budget_total", 0), "#,##0"),
        ("CA articles fidelite", kpis.get("ca_total", 0), "#,##0"),
        ("Marge articles fidelite", kpis.get("marge_total", 0), "#,##0;[Red]-#,##0"),
        ("Qte articles fidelite", kpis.get("qte_total", 0), "#,##0.0"),
        ("CA total PGC (reference)", kpis.get("ca_pgc_total", 0), "#,##0"),
        ("Marge totale PGC (reference)", kpis.get("marge_pgc_total", 0), "#,##0;[Red]-#,##0"),
        ("% CA fidelite / PGC", kpis.get("pct_ca", 0), "0.00%"),
        ("% Marge fidelite / PGC", kpis.get("pct_marge", 0), "0.00%"),
    ]
    
    for i, (lbl, val, fmt) in enumerate(kpi_items):
        r = 4 + i
        ws2.cell(row=r, column=1, value=lbl).font = Font(name=POLICE, size=10, color=COULEUR_TEXTE)
        c = ws2.cell(row=r, column=2, value=val)
        # Mise en evidence des % pour la direction
        if lbl.startswith("%"):
            c.font = Font(name=POLICE, size=11, bold=True, color=COULEUR_BLEU)
            c.fill = PatternFill("solid", fgColor=COULEUR_JAUNE_TOTAL)
        else:
            c.font = Font(name=POLICE, size=11, bold=True, color=COULEUR_BLEU)
            c.fill = PatternFill("solid", fgColor=COULEUR_GRIS_FOND)
        c.number_format = fmt
        c.alignment = Alignment(horizontal="right")
    
    # Tableau par article x semaine
    row_titre = 14
    ws2.cell(row=row_titre, column=1, value="Detail par article et semaine").font = Font(
        name=POLICE, size=11, bold=True, color=COULEUR_TEXTE
    )
    
    headers_fin = [
        "Semaine", "Mois", "Code Article", "Article",
        "Nb magasins", "Cagnotte/u", "Qte", "Budget cagnotte",
        "CA", "Marge", "%CA", "%Marge",
    ]
    widths_fin = [10, 14, 13, 36, 13, 13, 11, 17, 14, 14, 10, 10]
    
    for i, w in enumerate(widths_fin, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    
    row_header = row_titre + 1
    for col_idx, h in enumerate(headers_fin, start=1):
        cell = ws2.cell(row=row_header, column=col_idx, value=h)
        style_header(cell)
    ws2.row_dimensions[row_header].height = 28
    
    row_data = row_header + 1
    if not df_financier.empty:
        df_fin_sorted = df_financier.sort_values(by=["Semaine", "Code Article"]).reset_index(drop=True)
        for _, r in df_fin_sorted.iterrows():
            values = [
                r["Semaine"], r["Mois"], r["Code Article"], r["Article"],
                r["Nb magasins"], r["Cagnotte/u"], r["Qte"], r["Budget cagnotte"],
                r["CA"], r["Marge"], r["%CA"], r["%Marge"],
            ]
            for col_idx, val in enumerate(values, start=1):
                cell = ws2.cell(row=row_data, column=col_idx, value=val)
                if col_idx == 1:
                    style_cell(cell, fond=COULEUR_BLEU_CLAIR, align="center")
                elif col_idx == 2:
                    style_cell(cell, fond=COULEUR_BLEU_CLAIR, align="center")
                elif col_idx == 5:  # Nb magasins
                    style_cell(cell, fmt="#,##0", align="center")
                elif col_idx == 6:  # Cagnotte/u
                    style_cell(cell, fmt="#,##0", align="right")
                elif col_idx == 7:  # Qte
                    style_cell(cell, fmt="#,##0.0", align="right")
                elif col_idx == 8:  # Budget cagnotte
                    style_cell(cell, fond=COULEUR_JAUNE_TOTAL, fmt="#,##0", align="right", bold=True)
                elif col_idx in (9, 10):  # CA, Marge
                    style_cell(cell, fmt="#,##0;[Red]-#,##0", align="right")
                elif col_idx in (11, 12):  # %CA, %Marge
                    style_cell(cell, fond=COULEUR_BLEU_CLAIR, fmt="0.00%", align="right", bold=True)
                else:
                    style_cell(cell, align="left")
            row_data += 1
        
        # Ligne TOTAL
        row_total = row_data
        ws2.cell(row=row_total, column=1, value="TOTAL")
        ws2.cell(row=row_total, column=4, value=f"{len(df_fin_sorted)} ligne(s)")
        # Formules de total (Nb magasins, Qte, Budget, CA, Marge)
        ws2.cell(row=row_total, column=5, value=f"=SUM(E{row_header+1}:E{row_data-1})")
        ws2.cell(row=row_total, column=7, value=f"=SUM(G{row_header+1}:G{row_data-1})")
        ws2.cell(row=row_total, column=8, value=f"=SUM(H{row_header+1}:H{row_data-1})")
        ws2.cell(row=row_total, column=9, value=f"=SUM(I{row_header+1}:I{row_data-1})")
        ws2.cell(row=row_total, column=10, value=f"=SUM(J{row_header+1}:J{row_data-1})")
        # %CA et %Marge en ligne TOTAL = poids global vs total PGC (depuis kpis)
        ws2.cell(row=row_total, column=11, value=kpis.get("pct_ca", 0))
        ws2.cell(row=row_total, column=12, value=kpis.get("pct_marge", 0))
        for col_idx in range(1, 13):
            cell = ws2.cell(row=row_total, column=col_idx)
            if col_idx == 1:
                style_cell(cell, fond=COULEUR_JAUNE_TOTAL, bold=True, align="center")
            elif col_idx == 5:  # Nb magasins
                style_cell(cell, fond=COULEUR_JAUNE_TOTAL, bold=True, fmt="#,##0", align="center")
            elif col_idx == 7:  # Qte
                style_cell(cell, fond=COULEUR_JAUNE_TOTAL, bold=True, fmt="#,##0.0", align="right")
            elif col_idx == 8:  # Budget cagnotte
                style_cell(cell, fond=COULEUR_JAUNE_TOTAL, bold=True, fmt="#,##0", align="right")
            elif col_idx in (9, 10):  # CA, Marge
                style_cell(cell, fond=COULEUR_JAUNE_TOTAL, bold=True, fmt="#,##0;[Red]-#,##0", align="right")
            elif col_idx in (11, 12):  # %CA, %Marge
                style_cell(cell, fond=COULEUR_JAUNE_TOTAL, bold=True, fmt="0.00%", align="right")
            else:
                style_cell(cell, fond=COULEUR_JAUNE_TOTAL, bold=True, align="left")
    
    # Freeze
    ws2.freeze_panes = f"A{row_header+1}"
    
    # Sauvegarde en memoire
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ============================================================
# CALCUL RECAP FINANCIER
# ============================================================
def construire_recap_financier(df_recap_all: pd.DataFrame, totaux_pgc_par_semaine: dict) -> pd.DataFrame:
    """
    Agrege le recap detaille par Semaine x Article.
    totaux_pgc_par_semaine : dict {semaine: {"CA": x, "Marge": y, "Qte": z}}
    Ajoute les colonnes %CA et %Marge (poids vs total perimetre PGC de la semaine).
    """
    cols = [
        "Semaine", "Mois", "Code Article", "Article",
        "Nb magasins", "Cagnotte/u", "Qte", "Budget cagnotte",
        "CA", "Marge", "%CA", "%Marge",
        "CA PGC semaine", "Marge PGC semaine",
    ]
    if df_recap_all.empty:
        return pd.DataFrame(columns=cols)
    
    grouped = df_recap_all.groupby(
        ["Semaine", "Mois", "Code Article", "Article"], as_index=False
    ).agg(
        Nb_magasins=("Code Site", "nunique"),
        Cagnotte_mag=("Cagnotte", "first"),
        CA=("CA", "sum"),
        Marge=("Marge", "sum"),
        Qte=("Qte", "sum"),
        Budget_cagnotte=("Budget Cagnotte", "sum"),
    )
    grouped = grouped.rename(columns={
        "Nb_magasins": "Nb magasins",
        "Cagnotte_mag": "Cagnotte/u",
        "Budget_cagnotte": "Budget cagnotte",
    })
    
    # Ajout des totaux PGC de la semaine pour chaque ligne
    grouped["CA PGC semaine"] = grouped["Semaine"].map(
        lambda s: (totaux_pgc_par_semaine.get(s) or {}).get("CA", 0)
    )
    grouped["Marge PGC semaine"] = grouped["Semaine"].map(
        lambda s: (totaux_pgc_par_semaine.get(s) or {}).get("Marge", 0)
    )
    
    # Poids (gestion division par zero)
    grouped["%CA"] = grouped.apply(
        lambda r: (r["CA"] / r["CA PGC semaine"]) if r["CA PGC semaine"] else 0, axis=1
    )
    grouped["%Marge"] = grouped.apply(
        lambda r: (r["Marge"] / r["Marge PGC semaine"]) if r["Marge PGC semaine"] else 0, axis=1
    )
    
    grouped = grouped[cols]
    return grouped


# ============================================================
# INTERFACE STREAMLIT
# ============================================================
def render_fidelite_cagnotte():
    """Module a appeler depuis SmartBuyer Hub."""
    # CSS charte Apple/iOS
    st.markdown("""
    <style>
        .stApp { font-family: -apple-system, 'SF Pro Display', Calibri, sans-serif; }
        .header-cagnotte {
            display: flex; align-items: center; gap: 12px;
            padding: 16px 0; margin-bottom: 20px;
            border-bottom: 1px solid #E5E5EA;
        }
        .header-cagnotte .icone {
            width: 40px; height: 40px; border-radius: 10px;
            background: #007AFF; display: flex; align-items: center;
            justify-content: center; color: white; font-size: 20px;
        }
        .header-cagnotte .titre { font-size: 20px; font-weight: 600; color: #1C1C1E; margin: 0; }
        .header-cagnotte .sous-titre { font-size: 13px; color: #8E8E93; margin: 0; }
        .bandeau-periode {
            background: #E6F2FF; border-radius: 10px;
            padding: 12px 16px; margin: 16px 0;
        }
        .bandeau-periode .lbl { font-size: 11px; color: #0040A0; letter-spacing: 0.5px; }
        .bandeau-periode .val { font-size: 14px; font-weight: 600; color: #007AFF; }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="header-cagnotte">
        <div class="icone">$</div>
        <div>
            <p class="titre">Fidelite Cagnotte</p>
            <p class="sous-titre">Suivi hebdomadaire · Investissement vs Performance</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # === ETAPE 1 : Upload ===
    st.markdown("**1 · Charger les fichiers**")
    col1, col2 = st.columns(2)
    with col1:
        fichiers_pbi = st.file_uploader(
            "Extractions PBI (xlsx, plusieurs possibles)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="fidelite_cagnotte_files",
        )
    with col2:
        fichier_liste = st.file_uploader(
            "Liste articles (csv)",
            type=["csv"],
            key="fidelite_cagnotte_liste",
        )
    
    if not fichiers_pbi or not fichier_liste:
        st.info("Charge au moins un fichier PBI et le CSV liste pour demarrer.")
        return
    
    # === ETAPE 2 : Parsing ===
    liste_df = lire_liste_csv(fichier_liste)
    
    parsings = []
    semaines_vues = {}
    
    for f in fichiers_pbi:
        p = parser_pbi(f)
        if p["periode"] is None:
            st.error(f"`{f.name}` : impossible de detecter la periode. {p.get('erreur', '')}")
            continue
        sem = p["periode"]["semaine"]
        if sem in semaines_vues:
            st.warning(f"Doublon detecte : la semaine **{sem}** est presente dans `{semaines_vues[sem]}` et `{f.name}`. Les deux seront empilees.")
        else:
            semaines_vues[sem] = f.name
        parsings.append({"fichier": f.name, "parsing": p})
    
    if not parsings:
        st.error("Aucun fichier PBI exploitable.")
        return
    
    # Affichage periodes detectees
    st.markdown("**2 · Periodes detectees**")
    for item in parsings:
        per = item["parsing"]["periode"]
        st.markdown(f"""
        <div style="background: #F2F2F7; border-radius: 8px; padding: 8px 12px; margin-bottom: 6px; font-size: 13px;">
            <strong>{item['fichier']}</strong> · {per['date_debut'].strftime('%d/%m/%Y')} → {per['date_fin'].strftime('%d/%m/%Y')} · {per['semaine']} · {per['mois']}
        </div>
        """, unsafe_allow_html=True)
    
    # === ETAPE 3 : Construction recap ===
    recap_parts = []
    articles_manquants_global = {}
    totaux_pgc_par_semaine = {}  # {semaine: {"CA":, "Marge":, "Qte":}}
    
    for item in parsings:
        res = construire_recap_fichier(item["parsing"], liste_df)
        if not res["lignes_recap"].empty:
            recap_parts.append(res["lignes_recap"])
        manquants = set(res["articles_attendus"]) - set(res["articles_trouves"])
        if manquants:
            sem = item["parsing"]["periode"]["semaine"]
            articles_manquants_global[f"{item['fichier']} ({sem})"] = list(manquants)
        # Collecte des totaux PGC pour le calcul du poids
        sem = res.get("semaine")
        tot = res.get("totaux_pgc")
        if sem and tot:
            # Si meme semaine chargee 2x : on cumule (cas de doublon utilisateur)
            if sem in totaux_pgc_par_semaine:
                totaux_pgc_par_semaine[sem] = {
                    "CA": totaux_pgc_par_semaine[sem]["CA"] + tot.get("CA", 0),
                    "Marge": totaux_pgc_par_semaine[sem]["Marge"] + tot.get("Marge", 0),
                    "Qte": totaux_pgc_par_semaine[sem]["Qte"] + tot.get("Qte", 0),
                }
            else:
                totaux_pgc_par_semaine[sem] = tot
    
    df_recap_all = pd.concat(recap_parts, ignore_index=True) if recap_parts else pd.DataFrame()
    
    # Alertes articles manquants
    if articles_manquants_global:
        with st.expander(f"⚠ Articles non trouves dans certaines extractions ({sum(len(v) for v in articles_manquants_global.values())})"):
            for fic, arts in articles_manquants_global.items():
                st.markdown(f"**{fic}** : {', '.join(arts)}")
    
    if df_recap_all.empty:
        st.warning("Aucun article de la liste n'a ete trouve dans les extractions PBI pour les mois concernes.")
        return
    
    # === ETAPE 4 : KPIs et recap financier ===
    df_financier = construire_recap_financier(df_recap_all, totaux_pgc_par_semaine)
    
    # Totaux PGC cumules sur toutes les semaines chargees
    ca_pgc_total = sum(v.get("CA", 0) for v in totaux_pgc_par_semaine.values())
    marge_pgc_total = sum(v.get("Marge", 0) for v in totaux_pgc_par_semaine.values())
    
    ca_fid = df_financier["CA"].sum()
    marge_fid = df_financier["Marge"].sum()
    
    kpis = {
        "budget_total": df_financier["Budget cagnotte"].sum(),
        "ca_total": ca_fid,
        "marge_total": marge_fid,
        "qte_total": df_financier["Qte"].sum(),
        "ca_pgc_total": ca_pgc_total,
        "marge_pgc_total": marge_pgc_total,
        "pct_ca": (ca_fid / ca_pgc_total) if ca_pgc_total else 0,
        "pct_marge": (marge_fid / marge_pgc_total) if marge_pgc_total else 0,
    }
    
    st.markdown("**3 · Synthese**")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Budget cagnotte", f"{kpis['budget_total']:,.0f}".replace(",", " "))
    c2.metric("CA articles fidelite", f"{kpis['ca_total']:,.0f}".replace(",", " "))
    c3.metric("Marge articles fidelite", f"{kpis['marge_total']:,.0f}".replace(",", " "))
    c4.metric("Qte vendue", f"{kpis['qte_total']:,.1f}".replace(",", " "))
    
    # Ligne 2 KPIs : poids fidelite / PGC
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("CA total PGC", f"{kpis['ca_pgc_total']:,.0f}".replace(",", " "))
    c6.metric("Marge totale PGC", f"{kpis['marge_pgc_total']:,.0f}".replace(",", " "))
    c7.metric("% CA fidelite / PGC", f"{kpis['pct_ca']*100:.2f}%")
    c8.metric("% Marge fidelite / PGC", f"{kpis['pct_marge']*100:.2f}%")
    
    # === ETAPE 5 : Apercu ===
    st.markdown("**4 · Apercu Recap Financier**")
    st.dataframe(df_financier, use_container_width=True, hide_index=True)
    
    with st.expander("Voir le Recap detaille (Article x Magasin)"):
        st.dataframe(df_recap_all, use_container_width=True, hide_index=True)
    
    # === ETAPE 6 : Export ===
    st.markdown("**5 · Telecharger**")
    buf = exporter_excel(df_recap_all, df_financier, kpis)
    
    # Nom de fichier
    mois_dedans = df_recap_all["Mois"].unique()
    if len(mois_dedans) == 1:
        nom_fichier = f"Fidelite_Cagnotte_{mois_dedans[0].replace(' ', '_')}.xlsx"
    else:
        nom_fichier = f"Fidelite_Cagnotte_multi_mois.xlsx"
    
    st.download_button(
        label=f"📥 Telecharger {nom_fichier}",
        data=buf,
        file_name=nom_fichier,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ============================================================
# POINT D'ENTREE
# ============================================================
if __name__ == "__main__":
    st.set_page_config(page_title="Fidelite Cagnotte", layout="wide")
    render_fidelite_cagnotte()
