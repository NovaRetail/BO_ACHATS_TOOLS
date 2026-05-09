import streamlit as st
import pandas as pd
import numpy as np
import json
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter

# ─── CONFIG PAGE ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Rentabilité · SmartBuyer",
    page_icon="📊",
    layout="wide"
)

# ─── CHARTE SMARTBUYER ──────────────────────────────────────────────────────
st.markdown("""
<style>
body, [data-testid="stAppViewContainer"] { background: #F2F2F7; }
[data-testid="stSidebar"] { background: #FFFFFF; border-right: 0.5px solid #E5E5EA; }
h1,h2,h3 { font-weight: 600; }
.block-container { padding-top: 1.5rem; }

.sb-card {
    background: #FFFFFF;
    border-radius: 14px;
    border: 0.5px solid #E5E5EA;
    padding: 1rem 1.25rem;
    margin-bottom: 0.75rem;
}
.sb-kpi-label { font-size: 12px; color: #8E8E93; margin-bottom: 4px; }
.sb-kpi-val   { font-size: 22px; font-weight: 600; color: #1C1C1E; }
.sb-kpi-val.up   { color: #34C759; }
.sb-kpi-val.down { color: #FF3B30; }
.sb-kpi-val.warn { color: #FF9500; }

.badge-green { background:#E8F8ED; color:#1A7A3A; border-radius:20px; padding:2px 10px; font-size:11px; font-weight:600; }
.badge-red   { background:#FFEAEA; color:#C0392B; border-radius:20px; padding:2px 10px; font-size:11px; font-weight:600; }
.badge-warn  { background:#FFF3E0; color:#B45309; border-radius:20px; padding:2px 10px; font-size:11px; font-weight:600; }
.badge-blue  { background:#E8F0FE; color:#1A56DB; border-radius:20px; padding:2px 10px; font-size:11px; font-weight:600; }

.intro-card {
    background: #EAF4FF;
    border-left: 4px solid #007AFF;
    border-radius: 0 14px 14px 0;
    padding: 1rem 1.25rem;
    margin-bottom: 1.5rem;
}
.kpi-guide {
    background: #FFFFFF;
    border-radius: 14px;
    border: 0.5px solid #E5E5EA;
    padding: 1rem 1.25rem;
    margin-bottom: 1rem;
}
.kpi-guide-title { font-size: 13px; font-weight: 600; color: #1C1C1E; margin-bottom: 4px; }
.kpi-guide-desc  { font-size: 12px; color: #48484A; line-height: 1.5; }
.kpi-guide-seuil { font-size: 11px; color: #8E8E93; margin-top: 6px; }
.num-badge { display:inline-flex; align-items:center; justify-content:center;
    width:22px; height:22px; border-radius:50%; font-size:11px; font-weight:600;
    margin-right:8px; flex-shrink:0; }
.n1 { background:#FFEAEA; color:#C0392B; }
.n2 { background:#FFF3E0; color:#B45309; }
.n3 { background:#E8F0FE; color:#1A56DB; }
.n4 { background:#F2F2F7; color:#48484A; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTES ─────────────────────────────────────────────────────────────
CONFIG_PATH = "config_seuils_rentabilite.json"

SEUILS_DEFAULT = {
    "00010 - BOISSONS":           {"evo_marge": -1.0, "ecart_hp_pro": 8.0, "poids_promo": 30.0, "casse_marge": 1.0},
    "00014 - EPICERIE":           {"evo_marge": -1.5, "ecart_hp_pro": 9.0, "poids_promo": 25.0, "casse_marge": 3.0},
    "00011 - DROGUERIE":          {"evo_marge": -1.0, "ecart_hp_pro": 7.0, "poids_promo": 30.0, "casse_marge": 1.5},
    "00012 - PARFUMERIE HYGIENE": {"evo_marge": -0.5, "ecart_hp_pro": 7.0, "poids_promo": 25.0, "casse_marge": 1.0},
}

COLS_REQUIRED = [
    "Rayon", "Sous Famille", "Article",
    "CA", "CA N-1", "Marge", "%Marge", "Evo %Marge",
    "CA Hors Promo", "Marge Hors Promo", "%Marge Hors Promo",
    "CA Promo", "Marge Promo", "%CA Poids Promo", "%Marge Promo", "Evo %Marge Promo",
    "Casse (Valeur)", "%Casse (Valeur)",
]

RAYON_COLORS = {
    "00010 - BOISSONS":           "#007AFF",
    "00014 - EPICERIE":           "#FF3B30",
    "00011 - DROGUERIE":          "#34C759",
    "00012 - PARFUMERIE HYGIENE": "#FF2D55",
}

# ─── HELPERS ────────────────────────────────────────────────────────────────
def load_seuils():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH) as f:
            return json.load(f)
    return SEUILS_DEFAULT.copy()

def save_seuils(s):
    with open(CONFIG_PATH, "w") as f:
        json.dump(s, f, indent=2)

def safe_num(v):
    try:
        return float(v) if pd.notna(v) else np.nan
    except:
        return np.nan

def fmt_ca(v):
    if pd.isna(v): return "—"
    if abs(v) >= 1e9: return f"{v/1e9:.1f} Md"
    if abs(v) >= 1e6: return f"{v/1e6:.1f} M"
    if abs(v) >= 1e3: return f"{v/1e3:.0f} K"
    return f"{v:.0f}"

def fmt_pct(v, decimals=2):
    if pd.isna(v): return "—"
    return f"{v*100:.{decimals}f}%"

def fmt_pp(v):
    if pd.isna(v): return "—"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v:.2f} pp"

def verdict(row, seuils, rayon_key):
    s = seuils.get(rayon_key, list(SEUILS_DEFAULT.values())[0])
    evo   = safe_num(row.get("Evo %Marge", np.nan))
    ecart = safe_num(row.get("ecart_hp_pro", np.nan))
    poids = safe_num(row.get("%CA Poids Promo", np.nan))
    cm    = safe_num(row.get("casse_marge", np.nan))

    alertes = []
    if not pd.isna(evo)   and evo   < s["evo_marge"]:                          alertes.append("Marge ↓")
    if not pd.isna(ecart) and ecart > s["ecart_hp_pro"]:                        alertes.append("Promo déstr.")
    if not pd.isna(poids) and poids*100 > s["poids_promo"] and not pd.isna(cm): alertes.append("Poids promo élevé")
    if not pd.isna(cm)    and cm    > s["casse_marge"]:                         alertes.append("Casse élevée")

    if not alertes:
        return "🟢 Sain"
    elif len(alertes) >= 2:
        return "🔴 Alerte"
    else:
        return "🟡 Surveiller"

def is_total_row(row):
    art = str(row.get("Article", ""))
    sf  = str(row.get("Sous Famille", ""))
    fam = str(row.get("Famille", ""))
    return "Total" in art or "Total" in sf or "Total" in fam

# ─── LECTURE FICHIER ────────────────────────────────────────────────────────
def load_data(uploaded):
    df = pd.read_excel(uploaded, dtype=str)
    # Nettoyage colonnes dupliquées (%Vs N-1.x)
    rename = {}
    vsn1_cols = [c for c in df.columns if c.startswith("%Vs N-1")]
    mapping = ["%Vs N-1 CA", "%Vs N-1 Marge", "%Vs N-1 CA HP", "%Vs N-1 Marge HP",
               "%Vs N-1 CA Promo", "%Vs N-1 Marge Promo", "%Vs N-1 Qté"]
    for i, c in enumerate(vsn1_cols):
        if i < len(mapping):
            rename[c] = mapping[i]
    df = df.rename(columns=rename)

    # Conversion numérique
    num_cols = [c for c in df.columns if c not in ["Departement","Rayon","Famille","Sous Famille","Article"]]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # Filtrer lignes métadonnées (Filtres appliqués…)
    if "Departement" in df.columns:
        df = df[~df["Departement"].str.contains("Filtres appliqués", na=False)]

    return df

def compute_metrics(df):
    df = df.copy()
    # Écart HP - Promo en pp
    df["ecart_hp_pro"] = (df["%Marge Hors Promo"].fillna(0) - df["%Marge Promo"].fillna(0)) * 100
    # Casse / Marge produite
    df["casse_marge"] = np.where(
        df["Marge"].notna() & (df["Marge"] != 0),
        (df["Casse (Valeur)"].fillna(0).abs() / df["Marge"].abs()) * 100,
        np.nan
    )
    return df

def get_rayon_totals(df):
    mask = (
        df["Article"].str.contains("Total", na=False) &
        df["Rayon"].notna() &
        ~df["Rayon"].str.contains("Total", na=False)
    )
    return df[mask].copy()

def get_sf_totals(df):
    mask = (
        df["Article"].str.contains("Total", na=False) &
        df["Sous Famille"].notna() &
        ~df["Sous Famille"].str.contains("Total", na=False) &
        ~df["Rayon"].str.contains("Total", na=False)
    )
    return df[mask].copy()

def get_articles(df):
    return df[~df["Article"].str.contains("Total", na=False) & df["CA"].notna()].copy()

# ─── EXPORT EXCEL ───────────────────────────────────────────────────────────
def make_excel(df_rayon, df_sf, df_articles, seuils, periode):
    wb = Workbook()

    # Styles
    BLUE   = "007AFF"; GREEN  = "34C759"; RED = "FF3B30"; ORANGE = "FF9500"
    GRAY_H = "F2F2F7"; WHITE  = "FFFFFF"; DARK = "1C1C1E"; MID = "8E8E93"

    def hdr_fill(hex_color): return PatternFill("solid", fgColor=hex_color)
    def font_bold(size=11, color=DARK, bold=True): return Font(name="Calibri", size=size, bold=bold, color=color)
    def font_reg(size=10, color=DARK): return Font(name="Calibri", size=size, color=color)
    def align_c(): return Alignment(horizontal="center", vertical="center", wrap_text=False)
    def align_l(): return Alignment(horizontal="left",   vertical="center")
    def thin_border():
        s = Side(style="thin", color="E5E5EA")
        return Border(left=s, right=s, top=s, bottom=s)

    def write_header_row(ws, row_num, headers, widths, fill_color=GRAY_H):
        for col, (h, w) in enumerate(zip(headers, widths), 1):
            cell = ws.cell(row=row_num, column=col, value=h)
            cell.fill   = hdr_fill(fill_color)
            cell.font   = font_bold(10, DARK)
            cell.alignment = align_c()
            cell.border = thin_border()
            ws.column_dimensions[get_column_letter(col)].width = w

    def color_cell(cell, value, threshold_good, threshold_bad, reverse=False):
        if pd.isna(value): return
        if not reverse:
            if value >= threshold_good: cell.font = Font(name="Calibri", size=10, color=GREEN, bold=True)
            elif value <= threshold_bad: cell.font = Font(name="Calibri", size=10, color=RED, bold=True)
        else:
            if value <= threshold_good: cell.font = Font(name="Calibri", size=10, color=GREEN, bold=True)
            elif value >= threshold_bad: cell.font = Font(name="Calibri", size=10, color=RED, bold=True)

    def verdict_fill(v):
        if "Sain"      in v: return hdr_fill("E8F8ED")
        if "Alerte"    in v: return hdr_fill("FFEAEA")
        if "Surveiller"in v: return hdr_fill("FFF3E0")
        return hdr_fill(WHITE)

    # ── ONGLET 1 : DASHBOARD ─────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "📊 Dashboard"
    ws1.sheet_view.showGridLines = False

    # Titre
    ws1.merge_cells("A1:J1")
    t = ws1["A1"]
    t.value = f"TABLEAU DE BORD RENTABILITÉ — {periode}"
    t.font  = font_bold(14, WHITE)
    t.fill  = hdr_fill(BLUE)
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[1].height = 28

    # KPIs globaux
    ws1.merge_cells("A2:J2")
    ws1.row_dimensions[2].height = 8

    ca_tot  = df_rayon["CA"].sum()
    mg_tot  = (df_rayon["Marge"].sum() / ca_tot) if ca_tot else np.nan
    evo_tot = df_rayon["Evo %Marge"].mean()
    casse_t = df_rayon["Casse (Valeur)"].sum()

    kpis = [
        ("CA Total", fmt_ca(ca_tot), ""),
        ("% Marge Global", f"{mg_tot*100:.2f}%" if not pd.isna(mg_tot) else "—", ""),
        ("Évolution Marge", fmt_pp(evo_tot), ""),
        ("Casse Totale", fmt_ca(casse_t), ""),
    ]
    kpi_cols = [1, 3, 5, 7]
    ws1.row_dimensions[3].height = 16
    ws1.row_dimensions[4].height = 20
    ws1.row_dimensions[5].height = 20
    ws1.row_dimensions[6].height = 8

    for i, (label, val, _) in enumerate(kpis):
        c = kpi_cols[i]
        ws1.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+1)
        ws1.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c+1)
        ws1.merge_cells(start_row=5, start_column=c, end_row=5, end_column=c+1)
        lc = ws1.cell(row=4, column=c, value=label)
        lc.font = font_reg(9, MID); lc.alignment = align_c()
        lc.fill = hdr_fill("F8F8FA")
        vc = ws1.cell(row=5, column=c, value=val)
        vc.font = font_bold(13, DARK); vc.alignment = align_c()
        vc.fill = hdr_fill("F8F8FA")
        if "Évolution" in label and not pd.isna(evo_tot):
            vc.font = Font(name="Calibri", size=13, bold=True,
                           color=GREEN if evo_tot >= 0 else RED)
        if "Casse" in label:
            vc.font = Font(name="Calibri", size=13, bold=True, color=RED)

    # Guide lecture KPIs
    ws1.row_dimensions[7].height = 18
    guide_title = ws1.cell(row=7, column=1, value="GUIDE DE LECTURE DES INDICATEURS CLÉS")
    guide_title.font = font_bold(9, MID)
    guide_title.alignment = align_l()
    ws1.merge_cells("A7:J7")

    guide_data = [
        ("1", "Évolution % Marge (pp)", "Thermomètre principal. Alerte si < seuil configuré par rayon.", BLUE),
        ("2", "Écart Marge HP − Marge Promo", "Impact réel de la promo. Critique si écart > seuil (promo destructrice).", ORANGE),
        ("3", "Poids Promo dans le CA", "Part du CA réalisée en promo. Risque si > seuil avec marge promo faible.", "1A56DB"),
        ("4", "Casse / Marge produite", "Impact casse sur rentabilité. Plus parlant que la valeur brute.", MID),
    ]
    for i, (num, titre, desc, col) in enumerate(guide_data):
        r = 8 + i
        ws1.row_dimensions[r].height = 16
        nc = ws1.cell(row=r, column=1, value=num)
        nc.font = font_bold(9, WHITE); nc.fill = hdr_fill(col)
        nc.alignment = align_c()
        ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        tc = ws1.cell(row=r, column=2, value=titre)
        tc.font = font_bold(9, DARK); tc.fill = hdr_fill("F8F8FA"); tc.alignment = align_l()
        ws1.merge_cells(start_row=r, start_column=4, end_row=r, end_column=10)
        dc = ws1.cell(row=r, column=4, value=desc)
        dc.font = font_reg(9, MID); dc.fill = hdr_fill("F8F8FA"); dc.alignment = align_l()

    # Séparateur
    ws1.row_dimensions[12].height = 12

    # Tableau par Rayon
    headers = ["Rayon","CA","% vs N-1","% Marge","Évo pp","% M HP","% M Promo",
               "Écart HP-Pro","Poids Promo","Casse/Marge","Verdict"]
    widths  = [28, 14, 10, 10, 10, 10, 10, 12, 12, 12, 14]
    write_header_row(ws1, 13, headers, widths, BLUE)
    for cell in ws1[13]:
        cell.font = font_bold(10, WHITE)

    for r_idx, (_, row) in enumerate(df_rayon.iterrows(), 14):
        ws1.row_dimensions[r_idx].height = 16
        fill = hdr_fill(WHITE) if r_idx % 2 == 0 else hdr_fill("F8F9FA")
        rayon_key = str(row.get("Rayon",""))
        verd = verdict(row, seuils, rayon_key)

        vals = [
            rayon_key,
            fmt_ca(row.get("CA", np.nan)),
            fmt_pct(row.get("%Vs N-1 CA", np.nan), 1),
            fmt_pct(row.get("%Marge", np.nan), 2),
            fmt_pp(row.get("Evo %Marge", np.nan)),
            fmt_pct(row.get("%Marge Hors Promo", np.nan), 2),
            fmt_pct(row.get("%Marge Promo", np.nan), 2),
            f"{row.get('ecart_hp_pro', np.nan):.1f} pp" if not pd.isna(row.get("ecart_hp_pro")) else "—",
            fmt_pct(row.get("%CA Poids Promo", np.nan), 1),
            f"{row.get('casse_marge', np.nan):.2f}%" if not pd.isna(row.get("casse_marge")) else "—",
            verd,
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_reg(10); cell.fill = fill
            cell.alignment = align_c() if c_idx > 1 else align_l()
            cell.border = thin_border()
            # Couleurs conditionnelles
            if c_idx == 5:  # Evo
                evo_v = row.get("Evo %Marge", np.nan)
                if not pd.isna(evo_v):
                    cell.font = Font(name="Calibri", size=10, bold=True,
                                     color=GREEN if evo_v >= 0 else RED)
            if c_idx == 11:  # Verdict
                cell.fill = verdict_fill(verd)
                cell.font = font_bold(10, DARK)

    ws1.freeze_panes = "A14"

    # ── ONGLET 2 : SOUS FAMILLE ──────────────────────────────────────────
    ws2 = wb.create_sheet("🏪 Sous Famille")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:K1")
    t2 = ws2["A1"]
    t2.value = f"RENTABILITÉ PAR SOUS-FAMILLE — {periode}"
    t2.font = font_bold(13, WHITE); t2.fill = hdr_fill(BLUE)
    t2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws2.row_dimensions[1].height = 26

    headers2 = ["Rayon","Sous Famille","CA","% vs N-1","% Marge","Évo pp",
                "% M HP","% M Promo","Écart HP-Pro","Poids Promo","Casse/Marge","Verdict"]
    widths2   = [24, 28, 13, 10, 10, 10, 10, 10, 12, 12, 12, 14]
    write_header_row(ws2, 2, headers2, widths2, BLUE)
    for cell in ws2[2]:
        cell.font = font_bold(10, WHITE)

    prev_rayon = None
    for r_idx, (_, row) in enumerate(df_sf.iterrows(), 3):
        ws2.row_dimensions[r_idx].height = 15
        rayon_key = str(row.get("Rayon",""))
        fill_bg = hdr_fill(WHITE) if r_idx % 2 == 0 else hdr_fill("F8F9FA")
        verd = verdict(row, seuils, rayon_key)

        vals = [
            rayon_key,
            str(row.get("Sous Famille","")),
            fmt_ca(row.get("CA")),
            fmt_pct(row.get("%Vs N-1 CA"), 1),
            fmt_pct(row.get("%Marge"), 2),
            fmt_pp(row.get("Evo %Marge")),
            fmt_pct(row.get("%Marge Hors Promo"), 2),
            fmt_pct(row.get("%Marge Promo"), 2),
            f"{row.get('ecart_hp_pro', np.nan):.1f} pp" if not pd.isna(row.get("ecart_hp_pro")) else "—",
            fmt_pct(row.get("%CA Poids Promo"), 1),
            f"{row.get('casse_marge', np.nan):.2f}%" if not pd.isna(row.get("casse_marge")) else "—",
            verd,
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_reg(10); cell.fill = fill_bg
            cell.alignment = align_c() if c_idx > 2 else align_l()
            cell.border = thin_border()
            if c_idx == 6:
                evo_v = row.get("Evo %Marge", np.nan)
                if not pd.isna(evo_v):
                    cell.font = Font(name="Calibri", size=10, bold=True,
                                     color=GREEN if evo_v >= 0 else RED)
            if c_idx == 12:
                cell.fill = verdict_fill(verd)
                cell.font = font_bold(10, DARK)

    ws2.freeze_panes = "A3"
    ws2.auto_filter.ref = f"A2:L{2+len(df_sf)}"

    # ── ONGLET 3 : ALERTES ───────────────────────────────────────────────
    ws3 = wb.create_sheet("🚨 Alertes")
    ws3.sheet_view.showGridLines = False

    ws3.merge_cells("A1:I1")
    t3 = ws3["A1"]
    t3.value = f"ALERTES — ARTICLES HORS SEUIL — {periode}"
    t3.font = font_bold(13, WHITE); t3.fill = hdr_fill(RED)
    t3.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws3.row_dimensions[1].height = 26

    # Alertes SF
    ws3.merge_cells("A2:I2")
    ws3["A2"].value = "ALERTES PAR SOUS-FAMILLE"
    ws3["A2"].font  = font_bold(10, MID)
    ws3.row_dimensions[2].height = 16

    headers3 = ["Rayon","Sous Famille","CA","% Marge","Évo pp","Écart HP-Pro","Poids Promo","Casse/Marge","Motif"]
    widths3   = [24, 28, 13, 10, 10, 12, 12, 12, 24]
    write_header_row(ws3, 3, headers3, widths3, "FF3B30")
    for cell in ws3[3]:
        cell.font = font_bold(10, WHITE)

    alertes_sf = []
    for _, row in df_sf.iterrows():
        rayon_key = str(row.get("Rayon",""))
        s = seuils.get(rayon_key, list(SEUILS_DEFAULT.values())[0])
        motifs = []
        evo   = safe_num(row.get("Evo %Marge"))
        ecart = safe_num(row.get("ecart_hp_pro"))
        poids = safe_num(row.get("%CA Poids Promo"))
        cm    = safe_num(row.get("casse_marge"))
        if not pd.isna(evo)   and evo < s["evo_marge"]:                        motifs.append(f"Evo marge {evo:+.2f}pp")
        if not pd.isna(ecart) and ecart > s["ecart_hp_pro"]:                   motifs.append(f"Écart HP-Pro {ecart:.1f}pp")
        if not pd.isna(poids) and poids*100 > s["poids_promo"]:                motifs.append(f"Promo {poids*100:.1f}%")
        if not pd.isna(cm)    and cm > s["casse_marge"]:                       motifs.append(f"Casse/Marge {cm:.2f}%")
        if motifs:
            alertes_sf.append({**row, "motif": " · ".join(motifs)})

    for r_idx, row in enumerate(alertes_sf, 4):
        ws3.row_dimensions[r_idx].height = 15
        vals = [
            str(row.get("Rayon","")),
            str(row.get("Sous Famille","")),
            fmt_ca(row.get("CA")),
            fmt_pct(row.get("%Marge"), 2),
            fmt_pp(row.get("Evo %Marge")),
            f"{row.get('ecart_hp_pro', np.nan):.1f} pp" if not pd.isna(row.get("ecart_hp_pro")) else "—",
            fmt_pct(row.get("%CA Poids Promo"), 1),
            f"{row.get('casse_marge', np.nan):.2f}%" if not pd.isna(row.get("casse_marge")) else "—",
            row.get("motif",""),
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_reg(10); cell.fill = hdr_fill("FFEAEA")
            cell.alignment = align_c() if c_idx > 2 else align_l()
            cell.border = thin_border()

    # Séparateur
    sep_r = 4 + len(alertes_sf) + 1
    ws3.row_dimensions[sep_r].height = 16
    ws3.merge_cells(f"A{sep_r}:I{sep_r}")
    ws3[f"A{sep_r}"].value = "FLOP 10 — ARTICLES DESTRUCTEURS DE MARGE"
    ws3[f"A{sep_r}"].font  = font_bold(10, MID)

    headers_f = ["Rayon","Sous Famille","Article","CA","% Marge","Évo pp","% M Promo","Casse","Statut"]
    widths_f  = [20, 20, 38, 13, 10, 10, 10, 13, 14]
    write_header_row(ws3, sep_r+1, headers_f, widths_f, ORANGE)
    for cell in ws3[sep_r+1]:
        cell.font = font_bold(10, WHITE)

    flop = df_articles.copy()
    flop = flop[flop["%Marge"].notna() & flop["CA"].notna()]
    flop = flop.sort_values("%Marge").head(10)

    for r_idx, (_, row) in enumerate(flop.iterrows(), sep_r+2):
        ws3.row_dimensions[r_idx].height = 15
        mg = safe_num(row.get("%Marge"))
        statut = "🔴 Critique" if not pd.isna(mg) and mg < 0.05 else "🟡 Surveiller"
        vals = [
            str(row.get("Rayon","")),
            str(row.get("Sous Famille","")),
            str(row.get("Article","")),
            fmt_ca(row.get("CA")),
            fmt_pct(mg, 2),
            fmt_pp(row.get("Evo %Marge")),
            fmt_pct(row.get("%Marge Promo"), 2),
            fmt_ca(row.get("Casse (Valeur)")),
            statut,
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_reg(10)
            cell.fill = hdr_fill("FFF3E0") if "Surveiller" in statut else hdr_fill("FFEAEA")
            cell.alignment = align_c() if c_idx > 3 else align_l()
            cell.border = thin_border()
            if c_idx == 5 and not pd.isna(mg):
                cell.font = Font(name="Calibri", size=10, bold=True,
                                 color=RED if mg < 0.05 else ORANGE)

    ws3.freeze_panes = "A4"

    # ── ONGLET 4 : SYNTHÈSE HEBDO ────────────────────────────────────────
    ws4 = wb.create_sheet("📋 Synthèse Hebdo")
    ws4.sheet_view.showGridLines = False

    ws4.merge_cells("A1:G1")
    t4 = ws4["A1"]
    t4.value = f"SYNTHÈSE HEBDOMADAIRE RENTABILITÉ — {periode}"
    t4.font = font_bold(13, WHITE); t4.fill = hdr_fill(DARK)
    t4.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws4.row_dimensions[1].height = 26

    # KPIs
    ws4.row_dimensions[2].height = 8
    kpi4 = [("CA Total PGC", fmt_ca(ca_tot)), ("% Marge Global", f"{mg_tot*100:.2f}%" if not pd.isna(mg_tot) else "—"),
            ("Évo Marge", fmt_pp(evo_tot)), ("Casse Totale", fmt_ca(casse_t))]
    ws4.row_dimensions[3].height = 14; ws4.row_dimensions[4].height = 20; ws4.row_dimensions[5].height = 8
    for i,(lbl,val) in enumerate(kpi4):
        c = i*2+1
        ws4.merge_cells(start_row=3,start_column=c,end_row=3,end_column=c+1)
        ws4.merge_cells(start_row=4,start_column=c,end_row=4,end_column=c+1)
        lc = ws4.cell(row=3,column=c,value=lbl)
        lc.font=font_reg(9,MID); lc.alignment=align_c(); lc.fill=hdr_fill("F8F8FA")
        vc = ws4.cell(row=4,column=c,value=val)
        vc.font=font_bold(12,DARK); vc.alignment=align_c(); vc.fill=hdr_fill("F8F8FA")
        if "Évo" in lbl and not pd.isna(evo_tot):
            vc.font=Font(name="Calibri",size=12,bold=True,color=GREEN if evo_tot>=0 else RED)

    # Tableau synthèse rayon
    headers4 = ["Rayon","CA","% Marge","Évo pp","Poids Promo","Casse/Marge","Verdict"]
    widths4   = [28, 14, 12, 12, 14, 14, 14]
    write_header_row(ws4, 6, headers4, widths4, DARK)
    for cell in ws4[6]: cell.font=font_bold(10,WHITE)

    for r_idx,(_, row) in enumerate(df_rayon.iterrows(),7):
        ws4.row_dimensions[r_idx].height=16
        rayon_key=str(row.get("Rayon",""))
        fill_bg=hdr_fill(WHITE) if r_idx%2==0 else hdr_fill("F8F9FA")
        verd=verdict(row,seuils,rayon_key)
        vals=[
            rayon_key,
            fmt_ca(row.get("CA")),
            fmt_pct(row.get("%Marge"),2),
            fmt_pp(row.get("Evo %Marge")),
            fmt_pct(row.get("%CA Poids Promo"),1),
            f"{row.get('casse_marge',np.nan):.2f}%" if not pd.isna(row.get("casse_marge")) else "—",
            verd,
        ]
        for c_idx,val in enumerate(vals,1):
            cell=ws4.cell(row=r_idx,column=c_idx,value=val)
            cell.font=font_reg(10); cell.fill=fill_bg
            cell.alignment=align_c() if c_idx>1 else align_l()
            cell.border=thin_border()
            if c_idx==4:
                evo_v=row.get("Evo %Marge",np.nan)
                if not pd.isna(evo_v):
                    cell.font=Font(name="Calibri",size=10,bold=True,color=GREEN if evo_v>=0 else RED)
            if c_idx==7:
                cell.fill=verdict_fill(verd); cell.font=font_bold(10,DARK)

    # Nb alertes
    nb_alertes=sum(1 for _,r in df_rayon.iterrows() if "Alerte" in verdict(r,seuils,str(r.get("Rayon",""))))
    r_note=7+len(df_rayon)+1
    ws4.row_dimensions[r_note].height=14
    ws4.merge_cells(f"A{r_note}:G{r_note}")
    note=ws4[f"A{r_note}"]
    note.value=f"⚠️  {nb_alertes} rayon(s) en alerte · Consulter onglet Alertes pour le détail article"
    note.font=Font(name="Calibri",size=10,color="B45309")
    note.fill=hdr_fill("FFF3E0"); note.alignment=align_l()

    ws4.freeze_panes="A7"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─── SIDEBAR ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    st.markdown("---")

    seuils = load_seuils()
    rayons_disponibles = list(SEUILS_DEFAULT.keys())

    st.markdown("**Seuils d'alerte par rayon**")
    st.caption("Valeurs déclenchant une alerte dans le dashboard")

    cols_h = st.columns([3,1,1,1,1])
    cols_h[0].markdown("<small style='color:#8E8E93'>Rayon</small>", unsafe_allow_html=True)
    cols_h[1].markdown("<small style='color:#8E8E93'>Evo pp</small>", unsafe_allow_html=True)
    cols_h[2].markdown("<small style='color:#8E8E93'>HP-Pro</small>", unsafe_allow_html=True)
    cols_h[3].markdown("<small style='color:#8E8E93'>Promo%</small>", unsafe_allow_html=True)
    cols_h[4].markdown("<small style='color:#8E8E93'>Casse%</small>", unsafe_allow_html=True)

    for rayon in rayons_disponibles:
        s = seuils.get(rayon, SEUILS_DEFAULT[rayon])
        short = rayon.split(" - ")[-1][:10]
        cols = st.columns([3,1,1,1,1])
        cols[0].markdown(f"<small>{short}</small>", unsafe_allow_html=True)
        s["evo_marge"]   = cols[1].number_input("", value=float(s["evo_marge"]),   step=0.5, key=f"evo_{rayon}",   label_visibility="collapsed")
        s["ecart_hp_pro"]= cols[2].number_input("", value=float(s["ecart_hp_pro"]),step=1.0, key=f"eco_{rayon}",   label_visibility="collapsed")
        s["poids_promo"] = cols[3].number_input("", value=float(s["poids_promo"]), step=5.0, key=f"poi_{rayon}",   label_visibility="collapsed")
        s["casse_marge"] = cols[4].number_input("", value=float(s["casse_marge"]), step=0.5, key=f"cas_{rayon}",   label_visibility="collapsed")
        seuils[rayon] = s

    if st.button("💾 Enregistrer", use_container_width=True):
        save_seuils(seuils)
        st.success("Seuils enregistrés ✓")

    if st.button("↺ Réinitialiser", use_container_width=True):
        seuils = SEUILS_DEFAULT.copy()
        save_seuils(seuils)
        st.rerun()

    st.markdown("---")
    st.markdown("**Colonnes requises**")
    st.caption("Le fichier PBI doit contenir au minimum :")
    st.markdown("\n".join([f"- `{c}`" for c in ["Rayon","Sous Famille","Article","CA","Marge","%Marge","Evo %Marge","CA Promo","Marge Promo","Casse (Valeur)"]]))

# ─── MAIN ───────────────────────────────────────────────────────────────────
st.markdown("# 📊 Rentabilité · Marge / Promo / Casse")

st.markdown("""
<div class="intro-card">
<b>Module de suivi de la rentabilité</b> — Upload ton export PBI mensuel pour générer automatiquement
le tableau de bord par Rayon et Sous-Famille avec verdicts, alertes et Flop 10 articles.<br>
<small style="color:#48484A">Hiérarchie attendue : Département → Rayon → Famille → Sous Famille → Article</small>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("📂 Déposer l'export PBI (.xlsx)", type=["xlsx"])

if not uploaded:
    # Guide KPIs en page d'accueil
    st.markdown("#### Guide des indicateurs clés")
    kpi_guide = [
        ("1","n1","Évolution % Marge (pp)",
         "Thermomètre principal. Mesure si on gagne ou perd de la rentabilité vs N-1, indépendamment du volume vendu.",
         "Alerte si < seuil configuré · Sain si ≥ 0 pp"),
        ("2","n2","Écart Marge HP − Marge Promo",
         "Impact réel de la politique promo. Un grand écart signifie que les promos coûtent plus qu'elles ne rapportent.",
         "Critique si écart > seuil par rayon · Vigilance entre 5 et 8 pp"),
        ("3","n3","Poids Promo dans le CA",
         "Part du chiffre d'affaires réalisée en promotion. Au-delà du seuil avec une marge promo faible, le rayon est sous perfusion promo.",
         "Risque si > seuil configuré ET Marge Promo < 10%"),
        ("4","n4","Casse / Marge produite",
         "Impact réel de la casse sur la rentabilité. Plus parlant que la valeur brute — rapportée à la marge générée.",
         "Critique si > seuil · OK si < 1% de la marge"),
    ]
    col1, col2 = st.columns(2)
    for i, (num, cls, titre, desc, seuil_txt) in enumerate(kpi_guide):
        col = col1 if i % 2 == 0 else col2
        col.markdown(f"""
        <div class="kpi-guide">
          <div style="display:flex;align-items:center;margin-bottom:6px">
            <span class="num-badge {cls}">{num}</span>
            <span class="kpi-guide-title">{titre}</span>
          </div>
          <div class="kpi-guide-desc">{desc}</div>
          <div class="kpi-guide-seuil">⚙ {seuil_txt}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div style="background:#F2F2F7;border-radius:14px;padding:1rem 1.25rem;margin-top:0.5rem;">
    <b>Verdicts automatiques</b><br>
    <span style="color:#1A7A3A">🟢 Sain</span> — aucun seuil déclenché &nbsp;|&nbsp;
    <span style="color:#B45309">🟡 Surveiller</span> — 1 seuil déclenché &nbsp;|&nbsp;
    <span style="color:#C0392B">🔴 Alerte</span> — 2 seuils ou plus déclenchés
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ─── TRAITEMENT ─────────────────────────────────────────────────────────────
with st.spinner("Chargement et calcul des métriques…"):
    try:
        df = load_data(uploaded)
        df = compute_metrics(df)
        df_rayon   = compute_metrics(get_rayon_totals(df))
        df_sf      = compute_metrics(get_sf_totals(df))
        df_articles= get_articles(df)
        periode    = uploaded.name.replace(".xlsx","").replace("data","").strip("_").strip()
        if not periode:
            periode = "Période"
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")
        st.stop()

if df_rayon.empty:
    st.warning("Aucune ligne 'Total' par Rayon détectée. Vérifier le format du fichier.")
    st.stop()

# ─── KPIs GLOBAUX ───────────────────────────────────────────────────────────
ca_tot  = df_rayon["CA"].sum()
mg_tot  = df_rayon["Marge"].sum() / ca_tot if ca_tot else np.nan
evo_tot = df_rayon["Evo %Marge"].mean()
casse_t = df_rayon["Casse (Valeur)"].sum()
nb_alertes = sum(1 for _,r in df_rayon.iterrows()
                 if "Alerte" in verdict(r, seuils, str(r.get("Rayon",""))))

k1,k2,k3,k4,k5 = st.columns(5)
k1.markdown(f'<div class="sb-card"><div class="sb-kpi-label">CA Total</div><div class="sb-kpi-val">{fmt_ca(ca_tot)}</div></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="sb-card"><div class="sb-kpi-label">% Marge Global</div><div class="sb-kpi-val {"warn" if not pd.isna(mg_tot) and mg_tot<0.20 else ""}">{fmt_pct(mg_tot,2)}</div></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="sb-card"><div class="sb-kpi-label">Évo Marge vs N-1</div><div class="sb-kpi-val {"up" if not pd.isna(evo_tot) and evo_tot>=0 else "down"}">{fmt_pp(evo_tot)}</div></div>', unsafe_allow_html=True)
k4.markdown(f'<div class="sb-card"><div class="sb-kpi-label">Casse Totale</div><div class="sb-kpi-val down">{fmt_ca(casse_t)}</div></div>', unsafe_allow_html=True)
k5.markdown(f'<div class="sb-card"><div class="sb-kpi-label">Rayons en Alerte</div><div class="sb-kpi-val {"down" if nb_alertes>0 else "up"}">{nb_alertes}</div></div>', unsafe_allow_html=True)

# ─── TABS ───────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["📊 Par Rayon","🏪 Sous Famille","🚨 Alertes","📋 Flop 10"])

with tab1:
    rows = []
    for _, row in df_rayon.iterrows():
        rayon_key = str(row.get("Rayon",""))
        verd = verdict(row, seuils, rayon_key)
        ecart = row.get("ecart_hp_pro", np.nan)
        cm    = row.get("casse_marge", np.nan)
        rows.append({
            "Rayon":        rayon_key,
            "CA":           fmt_ca(row.get("CA")),
            "% vs N-1":     fmt_pct(row.get("%Vs N-1 CA"), 1),
            "% Marge":      fmt_pct(row.get("%Marge"), 2),
            "Évo pp":       fmt_pp(row.get("Evo %Marge")),
            "% M HP":       fmt_pct(row.get("%Marge Hors Promo"), 2),
            "% M Promo":    fmt_pct(row.get("%Marge Promo"), 2),
            "Écart HP-Pro": f"{ecart:.1f} pp" if not pd.isna(ecart) else "—",
            "Poids Promo":  fmt_pct(row.get("%CA Poids Promo"), 1),
            "Casse/Marge":  f"{cm:.2f}%" if not pd.isna(cm) else "—",
            "Verdict":      verd,
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

with tab2:
    rows2 = []
    for _, row in df_sf.iterrows():
        rayon_key = str(row.get("Rayon",""))
        verd = verdict(row, seuils, rayon_key)
        ecart = row.get("ecart_hp_pro", np.nan)
        cm    = row.get("casse_marge", np.nan)
        rows2.append({
            "Rayon":        rayon_key,
            "Sous Famille": str(row.get("Sous Famille","")),
            "CA":           fmt_ca(row.get("CA")),
            "% Marge":      fmt_pct(row.get("%Marge"), 2),
            "Évo pp":       fmt_pp(row.get("Evo %Marge")),
            "% M HP":       fmt_pct(row.get("%Marge Hors Promo"), 2),
            "% M Promo":    fmt_pct(row.get("%Marge Promo"), 2),
            "Écart HP-Pro": f"{ecart:.1f} pp" if not pd.isna(ecart) else "—",
            "Poids Promo":  fmt_pct(row.get("%CA Poids Promo"), 1),
            "Casse/Marge":  f"{cm:.2f}%" if not pd.isna(cm) else "—",
            "Verdict":      verd,
        })
    st.dataframe(pd.DataFrame(rows2), use_container_width=True, hide_index=True)

with tab3:
    alertes = []
    for _, row in df_sf.iterrows():
        rayon_key = str(row.get("Rayon",""))
        s = seuils.get(rayon_key, list(SEUILS_DEFAULT.values())[0])
        motifs = []
        evo   = safe_num(row.get("Evo %Marge"))
        ecart = safe_num(row.get("ecart_hp_pro"))
        poids = safe_num(row.get("%CA Poids Promo"))
        cm    = safe_num(row.get("casse_marge"))
        if not pd.isna(evo)   and evo < s["evo_marge"]:          motifs.append(f"Evo {evo:+.2f}pp")
        if not pd.isna(ecart) and ecart > s["ecart_hp_pro"]:     motifs.append(f"Écart HP-Pro {ecart:.1f}pp")
        if not pd.isna(poids) and poids*100 > s["poids_promo"]:  motifs.append(f"Promo {poids*100:.1f}%")
        if not pd.isna(cm)    and cm > s["casse_marge"]:         motifs.append(f"Casse {cm:.2f}%")
        if motifs:
            alertes.append({
                "Rayon":        rayon_key,
                "Sous Famille": str(row.get("Sous Famille","")),
                "CA":           fmt_ca(row.get("CA")),
                "% Marge":      fmt_pct(row.get("%Marge"), 2),
                "Évo pp":       fmt_pp(evo),
                "Motifs":       " · ".join(motifs),
            })
    if alertes:
        st.dataframe(pd.DataFrame(alertes), use_container_width=True, hide_index=True)
    else:
        st.success("✅ Aucune alerte — tous les seuils sont respectés")

with tab4:
    flop = df_articles[df_articles["%Marge"].notna() & df_articles["CA"].notna()].copy()
    flop = flop.sort_values("%Marge").head(10)
    rows_f = []
    for _, row in flop.iterrows():
        mg = safe_num(row.get("%Marge"))
        rows_f.append({
            "Rayon":        str(row.get("Rayon","")),
            "Sous Famille": str(row.get("Sous Famille","")),
            "Article":      str(row.get("Article","")),
            "CA":           fmt_ca(row.get("CA")),
            "% Marge":      fmt_pct(mg, 2),
            "Évo pp":       fmt_pp(row.get("Evo %Marge")),
            "% M Promo":    fmt_pct(row.get("%Marge Promo"), 2),
            "Casse":        fmt_ca(row.get("Casse (Valeur)")),
            "Statut":       "🔴 Critique" if not pd.isna(mg) and mg < 0.05 else "🟡 Surveiller",
        })
    st.dataframe(pd.DataFrame(rows_f), use_container_width=True, hide_index=True)

# ─── EXPORT ─────────────────────────────────────────────────────────────────
st.markdown("---")
col_dl, col_info = st.columns([1,3])
with col_dl:
    if st.button("📥 Générer le rapport Excel", type="primary", use_container_width=True):
        with st.spinner("Génération en cours…"):
            buf = make_excel(df_rayon, df_sf, df_articles, seuils, periode)
        st.download_button(
            label="⬇️ Télécharger",
            data=buf,
            file_name=f"Rentabilite_{periode}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
with col_info:
    st.caption(f"4 onglets : Dashboard · Sous Famille · Alertes · Synthèse Hebdo — Période : {periode}")
