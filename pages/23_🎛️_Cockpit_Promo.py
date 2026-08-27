"""
23_🎛️_Cockpit_Promo.py — SmartBuyer Hub
=========================================
Cockpit Promo — alertes Disponibilité & Marge sur les exports PROMO (type PROMO.CSV).

Colonnes obligatoires — fichier PROMO (export ERP, séparateur ';') :
  Code site, Rayon, Libellé rayon, Libellé article, Code article, DPR,
  PV Promo, Taux TVA, PMP, Stock, RAL, Marge en cours, Four.
  (Quantité vendue, Montant vente HT, Montant achat : optionnelles —
   utilisées pour la marge réalisée si présentes.)

Logique des flags (calculée en Python — aucune formule dans l'export Excel) :
  - Disponibilité : Stock <= 0
      "Rupture sans réappro" si RAL <= 0, sinon "Réappro en cours"
  - Marge : Marge en cours < 0 -> "Vente à perte"
            0 <= Marge en cours < seuil -> "Marge faible"
            PMP manquant -> "PMP manquant" (si pas déjà classé Vente à perte/Marge faible)
  - Type alerte = "Disponibilité + Marge" si les deux flags sont actifs,
                  sinon "Disponibilité" ou "Marge" seuls.

Rien n'est exploité tant que les contrôles ne sont pas au vert.
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import date

# =========================================================
# CONFIG
# =========================================================
COLOR_BG = "#F2F2F7"
COLOR_BLUE = "#007AFF"
COLOR_RED = "#FF3B30"
COLOR_GREEN = "#34C759"
COLOR_ORANGE = "#FF9500"
RADIUS = "14px"

MARGE_FAIBLE_SEUIL = 10  # % — modifiable dans la sidebar

VALID_SITES = {"0010301", "0010202", "0010203"}  # à ajuster si le périmètre magasins change

REQUIRED_COLS = [
    "Code site", "Rayon", "Libellé rayon", "Libellé article", "Code article",
    "DPR", "PV Promo", "Taux TVA", "PMP", "Stock", "RAL", "Marge en cours", "Four.",
]
OPTIONAL_COLS = ["Quantité vendue", "Montant vente HT", "Montant achat"]

st.set_page_config(page_title="Cockpit Promo", page_icon="🎛️", layout="wide")

# =========================================================
# STYLE — charte SmartBuyer Hub
# =========================================================
st.markdown(f"""
<style>
.stApp {{ background-color: {COLOR_BG}; }}
.sb-logo {{
    display:flex; align-items:center; gap:10px; padding:14px 4px 18px 4px;
    border-bottom: 0.5px solid #D1D1D6; margin-bottom: 14px;
}}
.sb-logo-mark {{
    width:36px; height:36px; border-radius:10px; background:{COLOR_BLUE};
    color:#fff; display:flex; align-items:center; justify-content:center;
    font-weight:700; font-size:15px; font-family:-apple-system,'SF Pro Display',Arial,sans-serif;
}}
.sb-logo-text {{ font-size:15px; font-weight:600; color:#1C1C1E; line-height:1.2; }}
.sb-logo-sub {{ font-size:11px; color:#8E8E93; }}
.alert-card {{
    background:#fff; border-radius:{RADIUS}; padding:16px 20px; margin-bottom:16px;
    border-left: 4px solid {COLOR_BLUE};
}}
.metric-card {{
    background:#fff; border-radius:{RADIUS}; padding:14px 16px; text-align:left;
}}
.metric-label {{ font-size:11px; color:#8E8E93; margin-bottom:4px; }}
.metric-value {{ font-size:24px; font-weight:600; color:#1C1C1E; }}
.col-required {{
    background:#fff; border-radius:{RADIUS}; padding:10px 14px; margin-bottom:8px;
    font-size:13px; border-left: 3px solid {COLOR_BLUE};
}}
.col-optional {{ border-left: 3px solid #C7C7CC; color:#636366; }}
.how-step {{
    display:flex; gap:10px; align-items:flex-start; margin-bottom:14px;
}}
.how-num {{
    width:22px; height:22px; border-radius:50%; background:{COLOR_BLUE}; color:#fff;
    font-size:12px; font-weight:600; display:flex; align-items:center; justify-content:center;
    flex-shrink:0;
}}
</style>
""", unsafe_allow_html=True)

# =========================================================
# SIDEBAR — bloc marque + import
# =========================================================
with st.sidebar:
    st.markdown("""
    <div class="sb-logo">
        <div class="sb-logo-mark">CP</div>
        <div>
            <div class="sb-logo-text">Cockpit Promo</div>
            <div class="sb-logo-sub">SmartBuyer Hub</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    # Remplacer par st.logo("assets/logo.png") si un vrai fichier logo existe dans le repo.

    st.markdown("**Import fichiers**")
    uploaded = st.file_uploader("Export PROMO (.csv)", type=["csv"])

    st.markdown("---")
    seuil = st.slider("Seuil marge faible (%)", min_value=0, max_value=30, value=MARGE_FAIBLE_SEUIL, step=1)

    st.markdown("---")
    st.caption("SmartBuyer Hub · Module 23 · Cockpit Promo")

# =========================================================
# HELPERS
# =========================================================
def read_csv_robust(file) -> pd.DataFrame:
    """Lecture CSV avec repli d'encodage UTF-8 -> CP1252 -> Latin-1 (convention SmartBuyer)."""
    raw = file.read()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(raw), sep=";", encoding=enc, dtype=str)
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError("Impossible de décoder le fichier (UTF-8 / CP1252 / Latin-1 ont échoué).")


def to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.strip(), errors="coerce")


def compute_alerts(df: pd.DataFrame, seuil_marge_faible: float) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    df = df[df["Code site"].isin(VALID_SITES)].copy()

    for c in ["DPR", "PV Promo", "Taux TVA", "PMP", "Stock", "RAL", "Marge en cours",
              "Quantité vendue", "Montant vente HT", "Montant achat", "Four."]:
        if c in df.columns:
            df[c] = to_num(df[c])
        else:
            df[c] = np.nan

    df["Site"] = df["Code site"].str.lstrip("0").astype(int)
    df["Rayon_lib"] = df["Libellé rayon"].str.strip()
    df["Article_lib"] = df["Libellé article"].str.strip()
    df["Code_article"] = df["Code article"].astype(str).str.lstrip("0")

    df["stat_stock"] = np.where(
        (df["Stock"] <= 0) & (df["RAL"].fillna(0) <= 0), "Rupture sans réappro",
        np.where((df["Stock"] <= 0) & (df["RAL"].fillna(0) > 0), "Réappro en cours", "")
    )
    df["stat_marge"] = np.where(
        df["Marge en cours"] < 0, "Vente à perte",
        np.where(df["Marge en cours"] < seuil_marge_faible, "Marge faible", "")
    )
    df["pmp_manquant"] = df["PMP"].isna()
    df["flag_dispo"] = df["stat_stock"] != ""
    df["flag_marge"] = (df["stat_marge"] != "") | df["pmp_manquant"]

    df["type_alerte"] = np.where(
        df["flag_dispo"] & df["flag_marge"], "Disponibilité + Marge",
        np.where(df["flag_dispo"], "Disponibilité",
                 np.where(df["flag_marge"], "Marge", ""))
    )
    df["marge_realisee"] = df["Montant vente HT"] - df["Montant achat"]
    df["pmp_effectif"] = df["PMP"].fillna(df["DPR"])
    df["detail_marge"] = np.where(
        df["stat_marge"] != "", df["stat_marge"],
        np.where(df["pmp_manquant"], "PMP manquant", "")
    )
    return df


def build_excel(alerts: pd.DataFrame) -> bytes:
    """Génère le classeur — toutes les valeurs sont calculées en Python, aucune formule."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    FONT_NAME = "Arial"
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    title_font = Font(name=FONT_NAME, bold=True, size=14, color="1F4E78")
    normal_font = Font(name=FONT_NAME, size=10)
    bold_font = Font(name=FONT_NAME, size=10, bold=True)
    red_fill = PatternFill("solid", fgColor="F8CBCB")
    orange_fill = PatternFill("solid", fgColor="FCE4B6")
    grey_fill = PatternFill("solid", fgColor="E7E6E6")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_header(ws, row_idx, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.font, cell.fill, cell.border = header_font, header_fill, border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def autosize(ws, widths):
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

    # --- Onglet Alertes ---
    ws1 = wb.active
    ws1.title = "Alertes"
    ws1["A1"] = "Alertes Disponibilité & Marge — Cockpit Promo"
    ws1["A1"].font = title_font
    ws1["A2"] = f"{len(alerts)} lignes en alerte | Généré le {date.today().strftime('%d/%m/%Y')} | Valeurs calculées (pas de formule)"
    ws1["A2"].font = Font(name=FONT_NAME, italic=True, size=9, color="666666")

    headers1 = ["Site", "Rayon", "Libellé article", "Code article", "Stock", "RAL", "PMP", "PV Promo",
                "Taux TVA", "Marge %", "Qté vendue", "Marge réalisée (FCFA)", "Type alerte",
                "Détail stock", "Détail marge", "Fournisseur"]
    hdr_row = 4
    for i, h in enumerate(headers1, start=1):
        ws1.cell(row=hdr_row, column=i, value=h)
    style_header(ws1, hdr_row, len(headers1))

    r = hdr_row + 1
    for _, d in alerts.iterrows():
        pv, tva, pmp = d["PV Promo"], d["Taux TVA"], d["pmp_effectif"]
        ht = pv / (1 + tva / 100) if pv and tva is not None else None
        marge_pct = (ht - pmp) / ht if ht else None

        ws1.cell(row=r, column=1, value=int(d["Site"]))
        ws1.cell(row=r, column=2, value=d["Rayon_lib"])
        ws1.cell(row=r, column=3, value=d["Article_lib"])
        ws1.cell(row=r, column=4, value=d["Code_article"])
        ws1.cell(row=r, column=5, value=None if pd.isna(d["Stock"]) else d["Stock"])
        ws1.cell(row=r, column=6, value=0 if pd.isna(d["RAL"]) else d["RAL"])
        ws1.cell(row=r, column=7, value=pmp)
        ws1.cell(row=r, column=8, value=pv)
        ws1.cell(row=r, column=9, value=tva / 100 if tva is not None else None)
        ws1.cell(row=r, column=10, value=round(marge_pct, 5) if marge_pct is not None else None)
        ws1.cell(row=r, column=11, value=None if pd.isna(d["Quantité vendue"]) else d["Quantité vendue"])
        ws1.cell(row=r, column=12, value=None if pd.isna(d["marge_realisee"]) else round(d["marge_realisee"], 2))
        ws1.cell(row=r, column=13, value=d["type_alerte"])
        ws1.cell(row=r, column=14, value=d["stat_stock"])
        ws1.cell(row=r, column=15, value=d["detail_marge"])
        ws1.cell(row=r, column=16, value=None if pd.isna(d["Four."]) else int(d["Four."]))

        if d["type_alerte"] == "Disponibilité + Marge":
            fill = red_fill
        elif d["type_alerte"] == "Marge":
            fill = red_fill if d["stat_marge"] == "Vente à perte" else orange_fill
        else:
            fill = orange_fill

        for c in range(1, 17):
            cell = ws1.cell(row=r, column=c)
            cell.font, cell.border, cell.fill = normal_font, border, fill
            if c == 9: cell.number_format = "0%"
            if c == 10: cell.number_format = "0.0%"
            if c == 12: cell.number_format = "#,##0"
        r += 1

    last_row1 = r - 1
    ws1.auto_filter.ref = f"A{hdr_row}:P{last_row1}"
    ws1.freeze_panes = f"A{hdr_row + 1}"
    autosize(ws1, [8, 20, 32, 12, 8, 7, 9, 10, 9, 9, 9, 15, 18, 20, 16, 12])

    # --- Onglet Synthèse ---
    ws2 = wb.create_sheet("Synthèse")
    ws2["A1"] = "Synthèse par magasin — Cockpit Promo"
    ws2["A1"].font = title_font
    ws2["A2"] = "Compteurs calculés à la génération du fichier (valeurs figées)"
    ws2["A2"].font = Font(name=FONT_NAME, italic=True, size=9, color="666666")

    headers2 = ["Site", "Rupture sans réappro", "Réappro en cours", "Vente à perte",
                "Marge faible", "PMP manquant", "Dispo + Marge (cumul)"]
    hdr_row2 = 4
    for i, h in enumerate(headers2, start=1):
        ws2.cell(row=hdr_row2, column=i, value=h)
    style_header(ws2, hdr_row2, len(headers2))

    sites = sorted(alerts["Site"].unique())
    r = hdr_row2 + 1
    first_data_row = r
    totals = {h: 0 for h in headers2[1:]}
    for site in sites:
        sub = alerts[alerts["Site"] == site]
        vals = {
            "Rupture sans réappro": int((sub["stat_stock"] == "Rupture sans réappro").sum()),
            "Réappro en cours": int((sub["stat_stock"] == "Réappro en cours").sum()),
            "Vente à perte": int((sub["stat_marge"] == "Vente à perte").sum()),
            "Marge faible": int((sub["stat_marge"] == "Marge faible").sum()),
            "PMP manquant": int((sub["detail_marge"] == "PMP manquant").sum()),
            "Dispo + Marge (cumul)": int((sub["type_alerte"] == "Disponibilité + Marge").sum()),
        }
        ws2.cell(row=r, column=1, value=int(site))
        for i, h in enumerate(headers2[1:], start=2):
            ws2.cell(row=r, column=i, value=vals[h])
            totals[h] += vals[h]
        for c in range(1, 8):
            cell = ws2.cell(row=r, column=c)
            cell.font, cell.border = normal_font, border
        r += 1

    ws2.cell(row=r, column=1, value="TOTAL")
    for i, h in enumerate(headers2[1:], start=2):
        ws2.cell(row=r, column=i, value=totals[h])
    for c in range(1, 8):
        cell = ws2.cell(row=r, column=c)
        cell.font, cell.border, cell.fill = bold_font, border, grey_fill

    r += 3
    ws2.cell(row=r, column=1, value="Total alertes (lignes distinctes) :")
    ws2.cell(row=r, column=1).font = bold_font
    ws2.cell(row=r, column=4, value=len(alerts))
    ws2.cell(row=r, column=4).font = Font(name=FONT_NAME, bold=True, size=12, color="C00000")

    autosize(ws2, [10, 20, 16, 14, 14, 14, 20])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =========================================================
# LANDING PAGE (avant upload)
# =========================================================
if uploaded is None:
    st.markdown("""
    <div class="alert-card">
        <div style="font-size:15px; font-weight:600; color:#1C1C1E; margin-bottom:4px;">🎛️ Cockpit Promo</div>
        <div style="font-size:13px; color:#48484A;">
            Balaie un export PROMO et sort en un clic les lignes à risque de rupture et/ou de marge négative,
            avec un flag unique par ligne et une synthèse par magasin.
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Ce que fait le module**")
        st.markdown("""
- Calcule la marge % à partir de PMP, PV Promo et Taux TVA
- Flag **Disponibilité** : stock ≤ 0, avec ou sans réappro en cours (RAL)
- Flag **Marge** : marge négative, marge faible (seuil réglable), ou PMP manquant
- Un flag combiné **Disponibilité + Marge** pour les cas cumulés, prioritaires
- Synthèse par magasin en un coup d'œil
- Export Excel prêt à diffuser (2 onglets, valeurs figées)
        """)
    with col2:
        st.markdown("**Comment ça marche**")
        st.markdown("""
        <div class="how-step"><div class="how-num">1</div><div style="font-size:13px; padding-top:2px;">Dépose l'export PROMO (.csv) dans la sidebar</div></div>
        <div class="how-step"><div class="how-num">2</div><div style="font-size:13px; padding-top:2px;">Le module calcule les flags et affiche le cockpit</div></div>
        <div class="how-step"><div class="how-num">3</div><div style="font-size:13px; padding-top:2px;">Filtre par magasin / type d'alerte, puis télécharge l'Excel</div></div>
        """, unsafe_allow_html=True)

    st.markdown("**Colonnes attendues**")
    c1, c2 = st.columns(2)
    with c1:
        for col in REQUIRED_COLS:
            st.markdown(f'<div class="col-required">{col}</div>', unsafe_allow_html=True)
    with c2:
        st.caption("Optionnelles (marge réalisée)")
        for col in OPTIONAL_COLS:
            st.markdown(f'<div class="col-required col-optional">{col}</div>', unsafe_allow_html=True)

    st.info("Dépose ton fichier PROMO.csv dans la sidebar pour lancer l'analyse.")
    st.stop()

# =========================================================
# TRAITEMENT
# =========================================================
try:
    raw_df = read_csv_robust(uploaded)
except Exception as e:
    st.error(f"Lecture du fichier impossible : {e}")
    st.stop()

missing = [c for c in REQUIRED_COLS if c not in [x.strip() for x in raw_df.columns]]
if missing:
    st.error(f"Colonnes manquantes dans le fichier : {', '.join(missing)}")
    st.stop()

alerts = compute_alerts(raw_df, seuil)

# =========================================================
# DASHBOARD
# =========================================================
st.markdown(f'<div style="font-size:15px; font-weight:600; color:#1C1C1E; margin-bottom:12px;">🎛️ Cockpit Promo — {len(alerts)} lignes en alerte</div>', unsafe_allow_html=True)

n_dispo = int((alerts["type_alerte"] == "Disponibilité").sum())
n_marge = int((alerts["type_alerte"] == "Marge").sum())
n_cumul = int((alerts["type_alerte"] == "Disponibilité + Marge").sum())

k1, k2, k3, k4 = st.columns(4)
for col, label, value, color in [
    (k1, "Total alertes", len(alerts), "#1C1C1E"),
    (k2, "Disponibilité", n_dispo, COLOR_ORANGE),
    (k3, "Marge", n_marge, COLOR_RED),
    (k4, "Cumul (critique)", n_cumul, COLOR_RED),
]:
    col.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value" style="color:{color};">{value}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("###  ")
st.markdown("**Synthèse par magasin**")
sites = sorted(alerts["Site"].unique())
rows = []
for site in sites:
    sub = alerts[alerts["Site"] == site]
    rows.append({
        "Site": site,
        "Rupture sans réappro": int((sub["stat_stock"] == "Rupture sans réappro").sum()),
        "Réappro en cours": int((sub["stat_stock"] == "Réappro en cours").sum()),
        "Vente à perte": int((sub["stat_marge"] == "Vente à perte").sum()),
        "Marge faible": int((sub["stat_marge"] == "Marge faible").sum()),
        "PMP manquant": int((sub["detail_marge"] == "PMP manquant").sum()),
        "Dispo + Marge": int((sub["type_alerte"] == "Disponibilité + Marge").sum()),
    })
st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

st.markdown("**Détail des alertes**")
fcol1, fcol2 = st.columns(2)
site_filter = fcol1.selectbox("Site", ["Tous"] + [str(s) for s in sites])
type_filter = fcol2.selectbox("Type alerte", ["Tous", "Disponibilité", "Marge", "Disponibilité + Marge"])

view = alerts.copy()
if site_filter != "Tous":
    view = view[view["Site"] == int(site_filter)]
if type_filter != "Tous":
    view = view[view["type_alerte"] == type_filter]

display_cols = ["Site", "Rayon_lib", "Article_lib", "Stock", "RAL", "Marge en cours", "type_alerte", "detail_marge"]
st.dataframe(
    view[display_cols].rename(columns={
        "Rayon_lib": "Rayon", "Article_lib": "Article", "Marge en cours": "Marge %",
        "type_alerte": "Type alerte", "detail_marge": "Détail marge"
    }),
    use_container_width=True, hide_index=True
)

st.markdown("###  ")
excel_bytes = build_excel(alerts)
st.download_button(
    "Télécharger l'Excel (Alertes + Synthèse)",
    data=excel_bytes,
    file_name=f"Cockpit_Promo_{date.today().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
