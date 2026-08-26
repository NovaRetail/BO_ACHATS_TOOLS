"""
Suivi Rayon vs Pays — PGC
=========================
Compare la performance d'un rayon PGC à la Tendance Pays (tout le reseau)
et au total PGC, avec :
  - un historique archive dans Google Sheets (meme backend que le module
    Tasks Tracker existant),
  - upload en masse d'exports PBI historiques,
  - detection automatique de la date par nom de fichier, avec repli sur
    la date du jour et confirmation manuelle (voir guess_date_from_filename),
  - trois ecrans acheteur : Ma semaine, Ma tendance, Mes magasins a regarder.

A brancher dans bo_achats_tools comme n'importe quel autre module
(auto-discovery app.py). Renommer le prefixe numerique selon ton
prochain numero de module disponible.

Pre-requis avant premier lancement :
  1. Creer un Google Sheet nomme SHEET_NAME (ou changer la constante),
     partage avec le meme service account que Tasks Tracker.
  2. pip install gspread google-auth pandas openpyxl streamlit
  3. Verifier que EXPORT_COLUMNS correspond exactement a l'ordre des
     colonnes de ton export PBI (Departement -> Volume_VsN1_pct).

Limite connue : ce script n'a pas ete execute contre un vrai Google
Sheet ni deploye sur Streamlit Cloud dans cet environnement — relire
avant mise en prod, en particulier les noms de colonnes et la logique
de detection de date.
"""

import io
import re
from datetime import date, datetime

import numpy as np
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# ============================================================
# CHARTE VISUELLE (alignee sur SmartBuyer Hub)
# ============================================================
COLOR_BG = "#F2F2F7"
COLOR_BLUE = "#007AFF"
COLOR_RED = "#FF3B30"
COLOR_GREEN = "#34C759"
COLOR_ORANGE = "#FF9500"
RADIUS = "14px"

FLAG_COLOR = {"Rouge": COLOR_RED, "Orange": COLOR_ORANGE, "Vert": COLOR_GREEN}

# ============================================================
# CONFIG
# ============================================================
SHEET_NAME = "Journal_Suivi_Rayon_Pays"   # nom du Google Sheet a creer
WORKSHEET_NAME = "Journal"

# A ajuster : chemin du logo dans ton repo bo_achats_tools (ex: "assets/logo.png").
# st.logo() accepte un chemin local ou une URL. Si le fichier n'existe pas encore,
# l'appel est protege plus bas (main()) pour ne pas faire planter la page.
LOGO_PATH = "assets/logo.png"

RAYONS = [
    "Tous PGC",
    "010 - BOISSON",
    "011 - DROGUERIE",
    "012 - PARFUMERIE HYGIENE",
    "014 - EPICERIE",
]

# Ordre exact des colonnes d'un export PBI (Departement -> Volume_VsN1_pct).
# A verifier/ajuster si le format d'export change.
EXPORT_COLUMNS = [
    "Departement", "Rayon", "Site", "CA_N1", "Budget", "CA", "Poids",
    "VsN1_pct", "VsBgt_pct", "Marge_N1", "Marge", "TauxMarge_N1", "TauxMarge",
    "TauxMarge_VsN1", "Debit_N1", "Debit", "Debit_VsN1_pct",
    "Panier_N1", "Panier", "Panier_VsN1_pct", "PanierQte_N1", "PanierQte",
    "PanierQte_VsN1_pct", "Volume_N1", "Volume", "Volume_VsN1_pct",
]
JOURNAL_COLUMNS = EXPORT_COLUMNS + ["Departement_rempli", "Rayon_rempli", "Format", "RowType", "Date"]

SEUILS_FORMAT = {"Hyper": 0.02, "Market": 0.03, "Supeco": 0.04}
SEUIL_ARCHIVE_SUFFISANT = 8   # semaines avant de passer sur des seuils statistiques
SEUIL_ARCHIVE_CONFORTABLE = 13


# ============================================================
# GOOGLE SHEETS
# ============================================================
@st.cache_resource
def get_gspread_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return gspread.authorize(creds)


def get_journal_worksheet():
    client = get_gspread_client()
    sheet = client.open(SHEET_NAME)
    try:
        return sheet.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sheet.add_worksheet(title=WORKSHEET_NAME, rows=2000, cols=len(JOURNAL_COLUMNS))
        ws.append_row(JOURNAL_COLUMNS)
        return ws


@st.cache_data(ttl=300)
def load_journal() -> pd.DataFrame:
    ws = get_journal_worksheet()
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=JOURNAL_COLUMNS)
    df = pd.DataFrame(records)
    df["Date"] = pd.to_datetime(df["Date"]).dt.date
    numeric_cols = [c for c in EXPORT_COLUMNS if c not in ("Departement", "Rayon", "Site")]
    for c in numeric_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def archive_export(df: pd.DataFrame, export_date: date) -> tuple[bool, str]:
    journal = load_journal()
    if not journal.empty and export_date in set(journal["Date"]):
        return False, f"{export_date.strftime('%d/%m/%Y')} est deja archivee — rien fait."
    ws = get_journal_worksheet()
    out = df.copy()
    out["Date"] = export_date.isoformat()
    rows = out[JOURNAL_COLUMNS].fillna("").values.tolist()
    ws.append_rows(rows)
    load_journal.clear()
    return True, f"{len(rows)} lignes archivees pour le {export_date.strftime('%d/%m/%Y')}."


# ============================================================
# PARSING D'UN EXPORT PBI
# ============================================================
def detect_format(site) -> str:
    if pd.isna(site):
        return ""
    s = str(site)
    if "Hyper" in s:
        return "Hyper"
    if "Market" in s:
        return "Market"
    if "Supeco" in s:
        return "Supeco"
    return ""


def detect_row_type(row) -> str:
    if pd.isna(row["Site"]) or row["Site"] == "":
        return "Dept Total"
    if row["Site"] == "Total":
        return "Rayon Total"
    if row["Format"]:
        return "Site"
    return "Autre"


def parse_pbi_export(uploaded_file) -> pd.DataFrame:
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=0)
    raw = raw.iloc[:, : len(EXPORT_COLUMNS)]
    raw.columns = EXPORT_COLUMNS
    raw["Departement_rempli"] = raw["Departement"].ffill()
    raw["Rayon_rempli"] = raw["Rayon"].ffill()
    raw["Format"] = raw["Site"].apply(detect_format)
    raw["RowType"] = raw.apply(detect_row_type, axis=1)
    return raw


# ============================================================
# DETECTION DE DATE (repli : nom de fichier, sinon aujourd'hui)
# ============================================================
DATE_PATTERN = re.compile(r"(\d{4})[-_](\d{2})[-_](\d{2})")


def guess_date_from_filename(filename: str) -> date:
    """Tente de lire une date AAAA-MM-JJ dans le nom du fichier.
    A defaut, retourne la date du jour — a confirmer par l'utilisateur.
    A remplacer par une lecture directe d'un champ Date PBI des que ce
    champ existera dans l'export (voir en-tete du fichier)."""
    match = DATE_PATTERN.search(filename)
    if match:
        try:
            return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            pass
    return date.today()


# ============================================================
# CALCULS PARTAGES
# ============================================================
def pgc_and_pays_refs(journal: pd.DataFrame, target_date: date):
    day = journal[journal["Date"] == target_date]
    pgc = day[(day["Departement"] == "01 - PGC") & (day["Rayon"] == "Total")]
    pays = day[day["Departement"] == "Total"]
    return (
        pgc.iloc[0] if len(pgc) else None,
        pays.iloc[0] if len(pays) else None,
    )


def rayon_total_row(journal: pd.DataFrame, target_date: date, rayon: str):
    day = journal[journal["Date"] == target_date]
    if rayon == "Tous PGC":
        row = day[(day["Departement"] == "01 - PGC") & (day["Rayon"] == "Total")]
    else:
        row = day[(day["Rayon"] == rayon) & (day["Site"] == "Total")]
    return row.iloc[0] if len(row) else None


def site_table(journal: pd.DataFrame, target_date: date, rayon: str) -> pd.DataFrame:
    day = journal[(journal["Date"] == target_date) & (journal["RowType"] == "Site")]
    total_pgc_ca = day["CA"].sum()
    if rayon != "Tous PGC":
        day = day[day["Rayon"] == rayon]
    if day.empty:
        return pd.DataFrame()

    agg = day.groupby("Site").agg(
        Format=("Format", "first"),
        CA_N1=("CA_N1", "sum"), CA=("CA", "sum"),
        Marge_N1=("Marge_N1", "sum"), Marge=("Marge", "sum"),
        Debit_N1=("Debit_N1", "sum"), Debit=("Debit", "sum"),
        Volume_N1=("Volume_N1", "sum"), Volume=("Volume", "sum"),
    ).reset_index()

    pgc_row, pays_row = pgc_and_pays_refs(journal, target_date)
    agg["CA_vs_N1"] = agg["CA"] / agg["CA_N1"] - 1
    agg["Ecart_Pays"] = agg["CA_vs_N1"] - pays_row["VsN1_pct"]
    agg["Ecart_PGC"] = agg["CA_vs_N1"] - pgc_row["VsN1_pct"]
    agg["Poids_CA"] = agg["CA"] / total_pgc_ca
    agg["Debit_vs_N1"] = agg["Debit"] / agg["Debit_N1"] - 1
    agg["Volume_vs_N1"] = agg["Volume"] / agg["Volume_N1"] - 1
    agg["Panier_vs_N1"] = (agg["CA"] / agg["Debit"]) / (agg["CA_N1"] / agg["Debit_N1"]) - 1
    agg["PanierQte_vs_N1"] = (agg["Volume"] / agg["Debit"]) / (agg["Volume_N1"] / agg["Debit_N1"]) - 1
    agg["TauxMarge"] = agg["Marge"] / agg["CA"]
    agg["TauxMarge_N1"] = agg["Marge_N1"] / agg["CA_N1"]
    agg["DeltaMarge"] = agg["TauxMarge"] - agg["TauxMarge_N1"]

    def flag(r):
        seuil = SEUILS_FORMAT.get(r["Format"], 0.03)
        worst = min(r["Ecart_Pays"], r["Ecart_PGC"])
        if worst < -seuil:
            return "Rouge"
        if worst < -seuil * 0.6:
            return "Orange"
        return "Vert"

    agg["Flag"] = agg.apply(flag, axis=1)

    def contact(r):
        if r["Flag"] == "Vert":
            return pd.Series(["", ""])
        if r["Debit_vs_N1"] < -0.05 and abs(r["Panier_vs_N1"]) < 0.03:
            return pd.Series(["Magasin", "Trafic en forte baisse"])
        if r["Volume_vs_N1"] < 0 and r["PanierQte_vs_N1"] < -0.03 and abs(r["Debit_vs_N1"]) < 0.03:
            return pd.Series(["Supply", "Volume/rupture, trafic stable (piste a verifier)"])
        if r["DeltaMarge"] < -0.02 and r["CA_vs_N1"] > 0:
            return pd.Series(["Achat", "Marge brute en recul, CA en hausse"])
        # Aucune regle dominante : on donne les chiffres bruts plutot qu'une
        # phrase generique, pour que l'acheteur voie lui-meme ou regarder.
        cause = (
            f"Trafic {r['Debit_vs_N1']:+.0%}, panier {r['Panier_vs_N1']:+.0%}, "
            f"volume {r['Volume_vs_N1']:+.0%}, marge brute {r['DeltaMarge']*100:+.1f} pt "
            f"— pas de cause dominante, plusieurs facteurs a la fois"
        )
        return pd.Series(["A investiguer", cause])

    agg[["Qui_contacter", "Causes"]] = agg.apply(contact, axis=1)
    agg["Score"] = agg[["Ecart_Pays", "Ecart_PGC"]].min(axis=1).clip(upper=0).abs() * agg["Poids_CA"]
    agg = agg.sort_values("Score", ascending=False).reset_index(drop=True)
    agg.insert(0, "Priorite", range(1, len(agg) + 1))
    return agg


# ============================================================
# EXPORT EXCEL
# ============================================================
FLAG_FILL_HEX = {"Rouge": "FFE0DE", "Orange": "FFEBCC", "Vert": "DCF5E3"}


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Export", flags: list | None = None) -> bytes:
    """Genere un classeur Excel en memoire (pas de fichier sur disque).
    Si `flags` est fourni (une valeur Rouge/Orange/Vert par ligne, meme ordre
    que df), l'en-tete passe en gras et chaque ligne est coloree pour
    ressembler a l'ecran."""
    from openpyxl.styles import Font, PatternFill

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        for cell in ws[1]:
            cell.font = Font(bold=True)

        if flags:
            for row_idx, flag_value in enumerate(flags, start=2):
                hex_color = FLAG_FILL_HEX.get(flag_value)
                if not hex_color:
                    continue
                fill = PatternFill("solid", fgColor=hex_color)
                for cell in ws[row_idx]:
                    cell.fill = fill

        for col_cells in ws.columns:
            width = max(len(str(c.value)) for c in col_cells if c.value is not None) + 2
            ws.column_dimensions[col_cells[0].column_letter].width = min(width, 40)
    return buffer.getvalue()


# ============================================================
# ECRAN : ARCHIVER
# ============================================================
def page_upload():
    st.markdown("### Archiver des exports PBI")
    st.caption(
        "Un ou plusieurs fichiers a la fois. La date est devinee depuis le nom du fichier "
        "(ex: export_2026-08-19.xlsx) ou mise a aujourd'hui par defaut — verifie/corrige "
        "avant d'archiver, une seule fois par fichier."
    )
    files = st.file_uploader("Exports PBI (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    if not files:
        return

    to_archive = []
    for f in files:
        guessed = guess_date_from_filename(f.name)
        col1, col2 = st.columns([3, 2])
        with col1:
            st.write(f.name)
        with col2:
            confirmed = st.date_input(f"Date — {f.name}", value=guessed, key=f"date_{f.name}")
        to_archive.append((f, confirmed))

    if st.button("Archiver ces exports", type="primary"):
        for f, d in to_archive:
            try:
                df = parse_pbi_export(f)
                ok, msg = archive_export(df, d)
                (st.success if ok else st.warning)(msg)
            except Exception as exc:
                st.error(f"{f.name} : echec ({exc})")


# ============================================================
# ECRAN : MA SEMAINE
# ============================================================
def page_ma_semaine(journal: pd.DataFrame):
    st.markdown("## Ma semaine")
    rayon = st.selectbox("Rayon", RAYONS, key="week_rayon")
    d = journal["Date"].max()
    st.caption(f"Semaine du {d.strftime('%d/%m/%Y')}")

    row = rayon_total_row(journal, d, rayon)
    pgc_row, pays_row = pgc_and_pays_refs(journal, d)
    if row is None or pgc_row is None or pays_row is None:
        st.warning("Donnees manquantes pour ce rayon a cette date.")
        return

    ca_n1, ca_bgt, marge = row["VsN1_pct"], row["VsBgt_pct"], row["TauxMarge"]
    ecart_pays = ca_n1 - pays_row["VsN1_pct"]
    ecart_pgc = ca_n1 - pgc_row["VsN1_pct"]

    c1, c2, c3 = st.columns(3)
    c1.metric("CA vs N-1", f"{ca_n1:+.1%}")
    c2.metric("CA vs budget", f"{ca_bgt:+.1%}")
    c3.metric("Marge brute (taux)", f"{marge:.1%}")

    c4, c5 = st.columns(2)
    c4.metric("Ecart vs Tendance Pays", f"{ecart_pays:+.1%}".replace("%", " pts"))
    c5.metric("Ecart vs PGC", f"{ecart_pgc:+.1%}".replace("%", " pts"))

    st.caption(
        f"Tendance Pays (tout reseau) : {pays_row['VsN1_pct']:+.1%} vs N-1  ·  "
        f"PGC (total departement) : {pgc_row['VsN1_pct']:+.1%} vs N-1"
    )


# ============================================================
# ECRAN : MA TENDANCE
# ============================================================
def page_ma_tendance(journal: pd.DataFrame):
    st.markdown("## Ma tendance")
    rayon = st.selectbox("Rayon", RAYONS, key="trend_rayon")

    dates = sorted(journal["Date"].unique())
    values, pgc_values = [], []
    for d in dates:
        row = rayon_total_row(journal, d, rayon)
        pgc_row, _ = pgc_and_pays_refs(journal, d)
        values.append(row["VsN1_pct"] if row is not None else np.nan)
        pgc_values.append(pgc_row["VsN1_pct"] if pgc_row is not None else np.nan)

    trend_df = pd.DataFrame(
        {"Mon rayon": values, "PGC (reference)": pgc_values}, index=pd.Index(dates, name="Date")
    )
    st.line_chart(trend_df)

    n = len(dates)
    if n < SEUIL_ARCHIVE_SUFFISANT:
        st.caption(
            f"{n} semaine(s) archivee(s) — seuils encore provisoires. "
            f"Cible : {SEUIL_ARCHIVE_SUFFISANT} a {SEUIL_ARCHIVE_CONFORTABLE} semaines pour des seuils statistiques fiables."
        )
    else:
        recent = pd.Series(values[-SEUIL_ARCHIVE_CONFORTABLE:]).dropna()
        st.caption(
            f"Seuils calibrables sur {len(recent)} semaines — moyenne {recent.mean():+.1%}, "
            f"ecart-type {recent.std():.1%}."
        )


# ============================================================
# ECRAN : MES MAGASINS A REGARDER
# ============================================================
def page_mes_magasins(journal: pd.DataFrame):
    st.markdown("## Mes magasins a regarder")
    rayon = st.selectbox("Rayon", RAYONS, key="store_rayon")
    d = journal["Date"].max()
    table = site_table(journal, d, rayon)
    if table.empty:
        st.info("Pas de donnees site pour ce rayon a cette date.")
        return

    display = table[[
        "Priorite", "Site", "Format", "CA_vs_N1", "Ecart_Pays", "Ecart_PGC",
        "Poids_CA", "Qui_contacter", "Causes",
    ]].rename(columns={
        "CA_vs_N1": "CA vs N-1", "Ecart_Pays": "Ecart vs Pays", "Ecart_PGC": "Ecart vs PGC",
        "Poids_CA": "Poids CA pays", "Qui_contacter": "Piste a verifier",
    })
    flags = table["Flag"].tolist()

    def color_rows(row):
        hex_color = {"Rouge": "#FFE0DE", "Orange": "#FFEBCC", "Vert": "#DCF5E3"}.get(flags[row.name], "")
        return [f"background-color: {hex_color}" if hex_color else "" for _ in row]

    styled = display.style.apply(color_rows, axis=1).format({
        "CA vs N-1": "{:+.1%}", "Ecart vs Pays": "{:+.1%}",
        "Ecart vs PGC": "{:+.1%}", "Poids CA pays": "{:.1%}",
    })
    st.dataframe(styled, use_container_width=True, hide_index=True)
    st.caption("Piste a verifier = hypothese reconstruite a partir du volume et du panier, pas une mesure de rupture reelle — a confirmer sur le terrain.")

    export_df = display.copy()
    for col in ("CA vs N-1", "Ecart vs Pays", "Ecart vs PGC"):
        export_df[col] = export_df[col].apply(lambda v: f"{v:+.1%}")
    export_df["Poids CA pays"] = export_df["Poids CA pays"].apply(lambda v: f"{v:.1%}")
    export_df["Plan d'actions"] = ""
    excel_bytes = to_excel_bytes(export_df, sheet_name="Mes magasins", flags=flags)
    filename = f"mes_magasins_{rayon.replace(' ', '_')}_{d.strftime('%Y%m%d')}.xlsx"
    st.download_button(
        "Telecharger en Excel",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ============================================================
# MAIN
# ============================================================
def main():
    st.set_page_config(page_title="Suivi Rayon vs Pays", layout="wide")
    st.markdown(f"<style>.stApp {{ background-color: {COLOR_BG}; }}</style>", unsafe_allow_html=True)
    try:
        st.logo(LOGO_PATH)
    except Exception:
        pass  # logo optionnel : ne bloque pas la page si le chemin n'existe pas encore
    st.title("Suivi Rayon vs Pays")

    tab_upload, tab_week, tab_trend, tab_stores = st.tabs(
        ["Archiver un export", "Ma semaine", "Ma tendance", "Mes magasins a regarder"]
    )
    with tab_upload:
        page_upload()

    journal = load_journal()
    if journal.empty:
        with tab_week, tab_trend, tab_stores:
            st.info("Aucune donnee archivee pour l'instant — commence par l'onglet 'Archiver un export'.")
        return

    with tab_week:
        page_ma_semaine(journal)
    with tab_trend:
        page_ma_tendance(journal)
    with tab_stores:
        page_mes_magasins(journal)


if __name__ == "__main__":
    main()
