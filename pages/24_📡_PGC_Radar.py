import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="PGC Radar", page_icon="📡", layout="wide")

# ============================================================
# Constantes
# ============================================================

FORMAT_MAP = {
    "10301": "Hyper", "10202": "Hyper", "10203": "Hyper",
    "10705": "Market", "10208": "Market", "10209": "Market",
    "10206": "Market", "10605": "Market", "10604": "Market",
    "10601": "Supeco", "10602": "Supeco", "10603": "Supeco",
}
SITES_HORS_RATTACHEMENT = {"10604", "10605"}  # Cité Verte, Aboboté — conversion récente

SEUIL_CA_CRITIQUE = -0.05
SEUIL_CA_ATTENTION = -0.02
SEUIL_MARGE_CRITIQUE_PT = -1.0
SEUIL_MARGE_ATTENTION_PT = -0.5
MATERIALITE_NIVEAU4_FCFA = 1_000_000  # plancher en valeur — évite le bruit naturel à la maille Rayon x Site

CAUSES = [
    "— À sélectionner —", "Rupture stock fournisseur", "Rupture stock magasin",
    "Prix non compétitif", "Exécution magasin", "Action concurrente",
    "Effet promo / cannibalisation", "Saisonnalité", "Autre",
]
STATUTS = ["À investiguer", "Cause identifiée", "Plan en cours", "Résolu"]

NAVY = "24235C"
ORANGE = "F09F39"
RED_FILL, RED_TEXT = "FFC7CE", "9C0006"
AMBER_FILL, AMBER_TEXT = "FFEB9C", "9C6500"
GREEN_FILL, GREEN_TEXT = "C6EFCE", "006100"

# ============================================================
# Chargement et parsing — Fichier 1 (Département / Rayon / Site, avec Budget)
# ============================================================

def load_fichier1(file):
    df = pd.read_excel(file, sheet_name="Export", header=0)
    df = df[~df["Département"].astype(str).str.startswith("Filtres appliqués", na=False)]
    df = df.dropna(how="all")
    df["Département"] = df["Département"].ffill()
    df["Rayon"] = df["Rayon"].ffill()

    grand_total = df[df["Département"] == "Total"]
    ca_reseau = grand_total["CA"].sum()
    budget_reseau = grand_total["Budget"].sum()

    df = df[df["Département"] != "Total"]

    dept_level = df[df["Rayon"] == "Total"].copy()
    rayon_level = df[(df["Site"] == "Total") & (df["Rayon"] != "Total") & (df["Département"] == "01 - PGC")].copy()
    site_level = df[(df["Site"] != "Total") & (df["Rayon"] != "Total") & df["Site"].notna()
                     & (df["Département"] == "01 - PGC")].copy()
    site_level["code_site"] = site_level["Site"].str.split(" - ").str[0].str.strip()
    site_level["format"] = site_level["code_site"].map(FORMAT_MAP).fillna("Autre")

    return {
        "dept": dept_level, "rayon": rayon_level, "site": site_level,
        "ca_reseau": ca_reseau, "budget_reseau": budget_reseau,
    }


def compute_format_level(site_level, exclude_sites=None):
    df = site_level.copy()
    if exclude_sites:
        df = df[~df["code_site"].isin(exclude_sites)]
    agg = df.groupby("format").agg(
        nb_sites=("code_site", "nunique"),
        CA=("CA", "sum"), CA_N1=("CA N-1", "sum"),
        Budget=("Budget", lambda s: np.nan if s.isna().all() else s.sum()),
        Marge=("Marge", "sum"), Marge_N1=("Marge N-1", "sum"),
        Debit=("Débit", "sum"), Panier_CA=("CA", "sum"),
    ).reset_index()
    agg["Vs_N1"] = (agg["CA"] - agg["CA_N1"]) / agg["CA_N1"]
    agg["Vs_Bgt"] = np.where(agg["Budget"].isna(), np.nan, (agg["CA"] - agg["Budget"]) / agg["Budget"])
    agg["Taux_marge"] = agg["Marge"] / agg["CA"]
    agg["Taux_marge_N1"] = agg["Marge_N1"] / agg["CA_N1"]
    agg["Marge_delta_pt"] = (agg["Taux_marge"] - agg["Taux_marge_N1"]) * 100
    agg["Panier"] = agg["CA"] / agg["Debit"]
    order = {"Hyper": 0, "Market": 1, "Supeco": 2}
    agg["_ord"] = agg["format"].map(order)
    return agg.sort_values("_ord").drop(columns="_ord")


# ============================================================
# Chargement et parsing — Fichier 2 (détail article, courant + N-1)
# ============================================================

def load_fichier2(file):
    df = pd.read_excel(file, sheet_name="Export", header=0)
    df = df[~df["Departement"].astype(str).str.startswith("Filtres appliqués", na=False)]
    df = df.dropna(how="all")
    for col in ["Departement", "Site nom long", "Rayon", "Famille", "Sous Famille"]:
        df[col] = df[col].ffill()
    leaf = df[df["Article"].notna() & (df["Article"] != "Total")].copy()
    leaf["code_article"] = leaf["Article"].str.extract(r"^(\d+)\s*-")
    return leaf


def build_detail_articles(courant, n1, site_filtre, rayon_filtre):
    cur = courant[(courant["Site nom long"] == site_filtre) & (courant["Rayon"] == rayon_filtre)].copy()
    key = ["code_article", "Site nom long", "Rayon"]
    n1_slim = n1[key + ["CA"]].rename(columns={"CA": "CA_N1"})
    merged = cur.merge(n1_slim, on=key, how="left")
    merged["nouveau"] = merged["CA_N1"].isna()
    merged["Vs_N1"] = np.where(merged["nouveau"], np.nan, (merged["CA"] - merged["CA_N1"]) / merged["CA_N1"])
    total_ca = merged["CA"].sum()
    merged["Contribution"] = merged["CA"] / total_ca if total_ca else np.nan
    return merged.sort_values("CA", ascending=False)


# ============================================================
# Alertes
# ============================================================

def alert_level(vs_bgt, vs_n1, marge_delta_pt):
    candidates_pct = [v for v in [vs_bgt, vs_n1] if pd.notna(v)]
    worst_pct = min(candidates_pct) if candidates_pct else np.nan
    md = marge_delta_pt if pd.notna(marge_delta_pt) else 0
    critique = (pd.notna(worst_pct) and worst_pct < SEUIL_CA_CRITIQUE) or md < SEUIL_MARGE_CRITIQUE_PT
    attention = (pd.notna(worst_pct) and worst_pct < SEUIL_CA_ATTENTION) or md < SEUIL_MARGE_ATTENTION_PT
    if critique:
        return "Critique"
    if attention:
        return "Attention"
    return "OK"


ALERT_COLOR = {"Critique": "🔴", "Attention": "🟠", "OK": "🟢"}


# ============================================================
# Formatage
# ============================================================

def fmt_m(x):
    return "—" if pd.isna(x) else f"{x/1_000_000:,.0f} M".replace(",", " ")

def fmt_pct(x, digits=1):
    return "—" if pd.isna(x) else f"{x*100:+.{digits}f}%"

def fmt_pt(x, digits=1):
    return "—" if pd.isna(x) else f"{x:+.{digits}f} pt"

def fmt_int(x):
    return "—" if pd.isna(x) else f"{int(round(x)):,}".replace(",", " ")


# ============================================================
# Export Excel (sans formule — calcul 100% Python)
# ============================================================

def build_excel_export(f1, fmt_a, fmt_b, alertes_df):
    wb = Workbook()
    ws = wb.active
    ws.title = "PGC Radar"

    header_fill = PatternFill("solid", fgColor=NAVY)
    header_font = Font(color="FFFFFF", bold=True)

    def write_header(row, headers, start_col=1):
        for i, h in enumerate(headers):
            c = ws.cell(row=row, column=start_col + i, value=h)
            c.fill = header_fill
            c.font = header_font

    def color_for(alert):
        return {"Critique": (RED_FILL, RED_TEXT), "Attention": (AMBER_FILL, AMBER_TEXT),
                "OK": (GREEN_FILL, GREEN_TEXT)}[alert]

    r = 1
    ws.cell(row=r, column=1, value="Niveau 1 — Vue globale").font = Font(bold=True, color=NAVY)
    r += 1
    write_header(r, ["Département", "CA", "Budget", "Vs Budget", "Vs N-1", "Marge (pt vs N-1)"])
    r += 1
    for _, row in f1["dept"].iterrows():
        vs_n1, vs_bgt = row.get("Vs N-1 (%)"), row.get("Vs Bgt (%)")
        marge_pt = row.get("Taux de Marge N Vs N-1")
        alert = alert_level(vs_bgt, vs_n1, marge_pt)
        fill, text = color_for(alert)
        ws.cell(row=r, column=1, value=row["Département"])
        ws.cell(row=r, column=2, value=round(row["CA"] / 1_000_000, 1))
        ws.cell(row=r, column=3, value=round(row["Budget"] / 1_000_000, 1) if pd.notna(row["Budget"]) else None)
        c4 = ws.cell(row=r, column=4, value=round(vs_bgt, 4) if pd.notna(vs_bgt) else None)
        c4.number_format = "0.0%"
        c5 = ws.cell(row=r, column=5, value=round(vs_n1, 4) if pd.notna(vs_n1) else None)
        c5.number_format = "0.0%"
        c6 = ws.cell(row=r, column=6, value=round(marge_pt, 2) if pd.notna(marge_pt) else None)
        for col in (4, 5, 6):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=fill)
            ws.cell(row=r, column=col).font = Font(color=text)
        r += 1

    r += 2
    ws.cell(row=r, column=1, value="Niveau 2 — Format (rattaché)").font = Font(bold=True, color=NAVY)
    r += 1
    write_header(r, ["Format", "Nb sites", "CA", "Marge", "Vs Budget", "Vs N-1"])
    r += 1
    for _, row in fmt_a.iterrows():
        alert = alert_level(row["Vs_Bgt"], row["Vs_N1"], row["Marge_delta_pt"])
        fill, text = color_for(alert)
        ws.cell(row=r, column=1, value=row["format"])
        ws.cell(row=r, column=2, value=int(row["nb_sites"]))
        ws.cell(row=r, column=3, value=round(row["CA"] / 1_000_000, 1))
        ws.cell(row=r, column=4, value=round(row["Marge"] / 1_000_000, 1))
        c5 = ws.cell(row=r, column=5, value=round(row["Vs_Bgt"], 4) if pd.notna(row["Vs_Bgt"]) else None)
        c5.number_format = "0.0%"
        c6 = ws.cell(row=r, column=6, value=round(row["Vs_N1"], 4) if pd.notna(row["Vs_N1"]) else None)
        c6.number_format = "0.0%"
        for col in (5, 6):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=fill)
            ws.cell(row=r, column=col).font = Font(color=text)
        r += 1

    r += 2
    ws.cell(row=r, column=1, value="Niveau 4 — Sites en alerte").font = Font(bold=True, color=NAVY)
    r += 1
    write_header(r, ["Site", "Rayon", "Vs Budget", "Vs N-1", "Marge (pt)", "Alerte", "Cause", "Statut"])
    r += 1
    for _, row in alertes_df.iterrows():
        fill, text = color_for(row["Alerte"])
        ws.cell(row=r, column=1, value=row["Site"])
        ws.cell(row=r, column=2, value=row["Rayon"])
        c3 = ws.cell(row=r, column=3, value=round(row["Vs_Bgt"], 4) if pd.notna(row["Vs_Bgt"]) else None)
        c3.number_format = "0.0%"
        c4 = ws.cell(row=r, column=4, value=round(row["Vs_N1"], 4) if pd.notna(row["Vs_N1"]) else None)
        c4.number_format = "0.0%"
        ws.cell(row=r, column=5, value=round(row["Marge_pt"], 2) if pd.notna(row["Marge_pt"]) else None)
        c6 = ws.cell(row=r, column=6, value=row["Alerte"])
        c6.fill = PatternFill("solid", fgColor=fill)
        c6.font = Font(color=text, bold=True)
        ws.cell(row=r, column=7, value="")
        ws.cell(row=r, column=8, value="À investiguer")
        r += 1

    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ============================================================
# UI — Import
# ============================================================

st.markdown("## 📡 PGC Radar")
st.caption("Pilotage Budget vs Réel vs N-1 — Département PGC")

with st.sidebar:
    st.markdown("### Import fichiers")
    up_f1 = st.file_uploader("Export 1 — Département / Rayon / Site (avec Budget)", type=["xlsx"])
    up_f2_cur = st.file_uploader("Export 2 — Détail article, semaine en cours", type=["xlsx"])
    up_f2_n1 = st.file_uploader("Export 2 — Détail article, semaine N-1 (dates iso)", type=["xlsx"])

if not up_f1:
    st.info("Dépose l'export 1 (Département / Rayon / Site) pour démarrer. Les exports article sont nécessaires uniquement pour le Niveau 5.")
    st.stop()

f1 = load_fichier1(up_f1)
fmt_a = compute_format_level(f1["site"])
fmt_b = compute_format_level(f1["site"], exclude_sites=SITES_HORS_RATTACHEMENT)

# Alertes Niveau 4 — maille Rayon x Site, PGC uniquement
alertes = f1["site"].copy()
alertes["Vs_Bgt"] = alertes["Vs Bgt (%)"]
alertes["Vs_N1"] = alertes["Vs N-1 (%)"]
alertes["Marge_pt"] = alertes["Taux de Marge N Vs N-1"]
alertes["Alerte"] = alertes.apply(lambda r: alert_level(r["Vs_Bgt"], r["Vs_N1"], r["Marge_pt"]), axis=1)
alertes["ecart_valeur"] = np.where(alertes["Budget"].notna(), alertes["CA"] - alertes["Budget"],
                                    alertes["CA"] - alertes["CA N-1"])
alertes = alertes.rename(columns={"Site": "Site_full", "Rayon": "Rayon_full"})
alertes["Site"] = alertes["Site_full"]
alertes["Rayon"] = alertes["Rayon_full"]
# Plancher de matérialité : à la maille Rayon x Site, un petit dénominateur produit des
# écarts en % très volatils (marge notamment) sans poids réel — on ne remonte que les
# alertes dont l'écart en valeur dépasse le plancher.
alertes_en_cours = alertes[(alertes["Alerte"] != "OK")
                            & (alertes["ecart_valeur"].abs() >= MATERIALITE_NIVEAU4_FCFA)].sort_values("ecart_valeur")

tab_prio, tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["Priorités", "Vue globale", "Format", "Rayon", "Alertes", "Détail articles"]
)

# ---------------- Priorités ----------------
with tab_prio:
    st.markdown("#### 3 priorités de la semaine")
    top3 = alertes_en_cours.head(3)
    if top3.empty:
        st.success("Aucune alerte Critique ou Attention cette semaine.")
    for _, row in top3.iterrows():
        c1, c2 = st.columns([3, 1])
        c1.markdown(f"**{row['Site']} — {row['Rayon']}**  \n{ALERT_COLOR[row['Alerte']]} {row['Alerte']}")
        c2.metric("Écart", fmt_m(row["ecart_valeur"]))

# ---------------- Niveau 1 ----------------
with tab1:
    st.markdown("#### Vue globale — benchmark inter-département")
    pgc = f1["dept"][f1["dept"]["Département"] == "01 - PGC"].iloc[0]
    poids_ca = pgc["CA"] / f1["ca_reseau"] if f1["ca_reseau"] else np.nan
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("CA PGC", fmt_m(pgc["CA"]), fmt_pct(pgc["Vs N-1 (%)"]))
    c2.metric("Marge PGC", fmt_m(pgc["Marge"]), fmt_pt(pgc["Taux de Marge N Vs N-1"]))
    c3.metric("Poids CA magasin", f"{poids_ca*100:.0f}%")
    c4.metric("Vs Budget", fmt_pct(pgc["Vs Bgt (%)"]))

    rows = []
    for _, row in f1["dept"].iterrows():
        rows.append({"Département": row["Département"], "CA": fmt_m(row["CA"]),
                     "Vs N-1": fmt_pct(row["Vs N-1 (%)"]), "Vs Budget": fmt_pct(row["Vs Bgt (%)"])})
    rows.append({"Département": "Total magasin", "CA": fmt_m(f1["ca_reseau"]), "Vs N-1": "", "Vs Budget": ""})
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

# ---------------- Niveau 2 ----------------
with tab2:
    st.markdown("#### Format — Tableau A (rattaché)")
    disp_a = fmt_a.copy()
    disp_a["CA"] = disp_a["CA"].apply(fmt_m)
    disp_a["Marge"] = disp_a["Marge"].apply(fmt_m)
    disp_a["Débit"] = disp_a["Debit"].apply(fmt_int)
    disp_a["Panier"] = disp_a["Panier"].apply(lambda x: "—" if pd.isna(x) else f"{x:,.0f} F".replace(",", " "))
    disp_a["Vs N-1"] = disp_a["Vs_N1"].apply(fmt_pct)
    disp_a["Vs Budget"] = disp_a["Vs_Bgt"].apply(fmt_pct)
    st.dataframe(disp_a[["format", "nb_sites", "CA", "Marge", "Débit", "Panier", "Vs N-1", "Vs Budget"]],
                 hide_index=True, use_container_width=True)

    st.caption("Budget non disponible pour Supeco dans l'export source (colonne vide sur les 3 sites).")

    st.markdown("#### Format — Tableau B (hors Cité Verte / Aboboté)")
    st.caption("Exclut les 2 sites convertis en Market en cours d'année, pour un Vs N-1 comparable au même format.")
    disp_b = fmt_b.copy()
    disp_b["CA"] = disp_b["CA"].apply(fmt_m)
    disp_b["Vs N-1"] = disp_b["Vs_N1"].apply(fmt_pct)
    st.dataframe(disp_b[["format", "nb_sites", "CA", "Vs N-1"]], hide_index=True, use_container_width=True)

# ---------------- Niveau 3 ----------------
with tab3:
    st.markdown("#### Rayon PGC")
    disp_r = f1["rayon"].copy()
    disp_r["CA"] = disp_r["CA"].apply(fmt_m)
    disp_r["Marge"] = disp_r["Marge"].apply(fmt_m)
    disp_r["Débit"] = disp_r["Débit"].apply(fmt_int)
    disp_r["Panier"] = disp_r["Panier"].apply(lambda x: "—" if pd.isna(x) else f"{x:,.0f} F".replace(",", " "))
    disp_r["Vs N-1"] = disp_r["Vs N-1 (%)"].apply(fmt_pct)
    disp_r["Vs Budget"] = disp_r["Vs Bgt (%)"].apply(fmt_pct)
    st.dataframe(disp_r[["Rayon", "CA", "Marge", "Débit", "Panier", "Vs N-1", "Vs Budget"]],
                 hide_index=True, use_container_width=True)

# ---------------- Niveau 4 ----------------
with tab4:
    st.markdown("#### Sites en alerte")
    st.caption(f"Seuils : CA < {SEUIL_CA_CRITIQUE:.0%} critique / {SEUIL_CA_ATTENTION:.0%} attention · "
               f"Marge < {SEUIL_MARGE_CRITIQUE_PT:.1f} pt critique / {SEUIL_MARGE_ATTENTION_PT:.1f} pt attention · "
               f"écart en valeur ≥ {MATERIALITE_NIVEAU4_FCFA:,.0f} FCFA pour filtrer le bruit à cette maille.".replace(",", " "))
    if alertes_en_cours.empty:
        st.success("Aucun site en alerte cette semaine.")
    else:
        edit_df = alertes_en_cours[["Site", "Rayon", "Alerte", "Vs_Bgt", "Vs_N1", "Marge_pt"]].copy()
        edit_df["Vs Budget"] = edit_df["Vs_Bgt"].apply(fmt_pct)
        edit_df["Vs N-1"] = edit_df["Vs_N1"].apply(fmt_pct)
        edit_df["Marge (pt)"] = edit_df["Marge_pt"].apply(fmt_pt)
        edit_df["Cause"] = CAUSES[0]
        edit_df["Statut"] = STATUTS[0]
        st.data_editor(
            edit_df[["Site", "Rayon", "Alerte", "Vs Budget", "Vs N-1", "Marge (pt)", "Cause", "Statut"]],
            column_config={
                "Cause": st.column_config.SelectboxColumn(options=CAUSES),
                "Statut": st.column_config.SelectboxColumn(options=STATUTS),
            },
            hide_index=True, use_container_width=True, key="alert_editor",
        )

# ---------------- Niveau 5 ----------------
with tab5:
    st.markdown("#### Détail articles")
    if not (up_f2_cur and up_f2_n1):
        st.info("Dépose les deux exports article (semaine en cours + semaine N-1, dates iso) pour activer ce niveau.")
    else:
        art_cur = load_fichier2(up_f2_cur)
        art_n1 = load_fichier2(up_f2_n1)
        sites_dispo = sorted(art_cur["Site nom long"].dropna().unique())
        rayons_dispo = sorted(art_cur["Rayon"].dropna().unique())
        c1, c2 = st.columns(2)
        site_sel = c1.selectbox("Site", sites_dispo)
        rayon_sel = c2.selectbox("Rayon", rayons_dispo)

        detail = build_detail_articles(art_cur, art_n1, site_sel, rayon_sel)
        disp = detail.copy()
        disp["CA"] = disp["CA"].apply(fmt_m)
        disp["Marge"] = disp["Marge"].apply(fmt_m)
        disp["Vs N-1"] = disp.apply(lambda r: "Nouveau" if r["nouveau"] else fmt_pct(r["Vs_N1"]), axis=1)
        disp["Contribution"] = disp["Contribution"].apply(fmt_pct)
        st.dataframe(disp[["Article", "CA", "Vs N-1", "Marge", "Contribution"]].head(30),
                     hide_index=True, use_container_width=True)

        pct_nouveaux = detail["nouveau"].mean()
        st.caption(f"{pct_nouveaux*100:.0f}% des articles de ce périmètre sont nouveaux ou sans historique N-1 comparable à ce site.")

# ---------------- Qualité de la donnée ----------------
st.divider()
st.markdown("#### Qualité de la donnée")
nb_sites_sans_budget = f1["site"][f1["site"]["Budget"].isna()]["code_site"].nunique()
nb_sites_total = f1["site"]["code_site"].nunique()
q1, q2, q3 = st.columns(3)
q1.metric("Sites sans budget (Supeco)", f"{nb_sites_sans_budget} / {nb_sites_total}")
if up_f2_cur and up_f2_n1:
    key = ["code_article", "Site nom long", "Rayon"]
    _n1_codes = set(zip(art_n1["code_article"], art_n1["Site nom long"], art_n1["Rayon"]))
    _sans_n1 = ~art_cur[key].apply(tuple, axis=1).isin(_n1_codes)
    q2.metric("Articles sans correspondance N-1 (réseau)", f"{_sans_n1.mean()*100:.0f}%")
else:
    q2.metric("Articles sans correspondance N-1", "—")
q3.metric("Écart de périmètre fichier 1 / fichier 2", "~3% (hors Ecommerce, périmètre FOOD)")

st.divider()
st.markdown("#### Export Excel")
excel_buf = build_excel_export(f1, fmt_a, fmt_b, alertes)
st.download_button("Télécharger le fichier Excel", data=excel_buf,
                    file_name="PGC_Radar.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
