import streamlit as st
from pathlib import Path

st.set_page_config(page_title="NovaRetail Solutions", page_icon="🛍️", layout="wide")

# ── GARDE : n'affiche un lien que si la page existe vraiment ──────────────────
def lien(chemin: str, label: str):
    if Path(chemin).exists():
        st.page_link(chemin, label=label)

st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "Helvetica Neue", Arial, sans-serif !important;
    background-color: #F2F2F7;
}
.stApp { background: #F2F2F7; }
.main .block-container { padding-top: 2rem; max-width: 1300px; }
header[data-testid="stHeader"] { display: none !important; }
[data-testid="stSidebar"] { background: #FFFFFF !important; border-right: 0.5px solid #E5E5EA !important; }

.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 1.5rem; }
.section-label {
    font-size: 11px; font-weight: 600; color: #8E8E93;
    text-transform: uppercase; letter-spacing: 0.07em;
    margin: 1.25rem 0 0.5rem; padding-bottom: 6px;
    border-bottom: 0.5px solid #E5E5EA;
}
.kpi-bar { background: #FFFFFF; border-radius: 14px; border: 0.5px solid #E5E5EA; padding: 0.85rem 1.25rem; text-align: center; }
.kpi-bar-val   { font-size: 20px; font-weight: 700; color: #1C1C1E; }
.kpi-bar-label { font-size: 11px; color: #8E8E93; margin-top: 2px; }

/* Wrapper card autour du st.page_link */
.mod-wrap {
    border-radius: 12px;
    overflow: hidden;
    margin-bottom: 4px;
    border: 0.5px solid #E5E5EA;
    border-left-width: 3px;
    background: #FFFFFF;
}
.mod-wrap [data-testid="stPageLink"] {
    background: transparent !important;
    border: none !important;
    border-radius: 0 !important;
    display: block !important;
    width: 100% !important;
}
.mod-wrap [data-testid="stPageLink"]:hover {
    background: #F9F9FB !important;
}
.mod-wrap [data-testid="stPageLink"] p {
    font-size: 13px !important;
    font-weight: 500 !important;
    color: #1C1C1E !important;
    padding: 12px 14px !important;
    margin: 0 !important;
    white-space: pre-line !important;
    line-height: 1.5 !important;
}
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:16px;padding-bottom:16px;border-bottom:0.5px solid #E5E5EA'>
  <div style='font-size:17px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>NovaRetail Solutions</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:2px'>Hub analytique · Équipe Achats CI</div>
</div>""", unsafe_allow_html=True)
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px'>Modules</div>", unsafe_allow_html=True)
    lien("pages/01_📊_Analyse_Scoring_ABC.py", label="📊  Scoring ABC")
    lien("pages/02_📈_Ventes_PBI.py",          label="📈  Ventes PBI")
    lien("pages/03_📦_Detention_Top_CA.py",    label="📦  Détention Top CA")
    lien("pages/04_💸_Performance_Promo.py",   label="💸  Performance Promo")
    lien("pages/05_🏪_Suivi_Implantation.py",  label="🏪  Suivi Implantation")
    lien("pages/06_💸_Marges_Negatives.py",    label="💸  Marges Négatives")
    lien("pages/07_📈_OTIF.py",                label="📈  OTIF")
    lien("pages/08_📦_OOS.py",                 label="📦  Ruptures OOS")
    lien("pages/09_✅_Tasks_Trackers.py",      label="✅  Tasks Tracker")
    lien("pages/10_📊_Perf_Hebdo.py",          label="📊  Perf Hebdo")
    lien("pages/11_📊_Rentabilite.py",         label="📊  Rentabilité")
    lien("pages/12_🏪_Bascule_XD.py",          label="🏪  Bascule XD")
    lien("pages/13_💸_Fidelite_cagnotte.py",   label="🏷️  Fidélité Cagnotte")

# ── EN-TÊTE ──────────────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>NovaRetail Solutions</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Plateforme analytique · Équipe Achats Carrefour CI</div>", unsafe_allow_html=True)

# ── KPIs ─────────────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)
for col, val, label in [
    (k1, "13", "Modules actifs"), (k2, "4", "Rayons couverts"),
    (k3, "13", "Sites réseau"),   (k4, "v2.3", "Version"),
    (k5, "Mai 2026", "Dernière période"),
]:
    col.markdown(f'<div class="kpi-bar"><div class="kpi-bar-val">{val}</div><div class="kpi-bar-label">{label}</div></div>', unsafe_allow_html=True)

# ── MODULES ──────────────────────────────────────────────────────────────────
# (page, color, label_titre, desc)
SECTIONS = [
    {
        "label": "📈 Performance commerciale",
        "modules": [
            ("pages/01_📊_Analyse_Scoring_ABC.py", "#007AFF", "📊  Scoring ABC  ›",       "Classification articles · 5 règles de recommandation"),
            ("pages/02_📈_Ventes_PBI.py",           "#5E35B1", "📈  Ventes PBI  ›",        "Comparaison hebdo CA / Marge entre deux périodes"),
            ("pages/10_📊_Perf_Hebdo.py",           "#5E35B1", "📊  Perf Hebdo  ›",        "Top CA · Flop marges · Top casse par rayon"),
            ("pages/11_📊_Rentabilite.py",          "#C0392B", "💰  Rentabilité  ›",       "Cockpit direction · score santé 0–100 · vs N-1"),
            ("pages/06_💸_Marges_Negatives.py",     "#FF3B30", "💸  Marges Négatives  ›",  "Flop 100 · matrice rayon × magasin · casse"),
            ("pages/04_💸_Performance_Promo.py",    "#FF9500", "🎯  Perf Promo  ›",        "Poids promo · taux marge · dépendance"),
        ],
    },
    {
        "label": "🏪 Assortiment & stock",
        "modules": [
            ("pages/03_📦_Detention_Top_CA.py",    "#B45309", "📦  Détention Top CA  ›",  "GOLD / SILVER · IM / LO · articles permanents"),
            ("pages/08_📦_OOS.py",                  "#B45309", "📦  Ruptures OOS  ›",      "Commander ou Voir Cession · seuil paramétrable"),
            ("pages/05_🏪_Suivi_Implantation.py",   "#1A7A3A", "🏪  Implantation  ›",      "Taux T1 · alertes article × site · cessions"),
            ("pages/13_💸_Fidelite_cagnotte.py",    "#007AFF", "🏷️  Fidélité Cagnotte  ›", "Budget cagnotte · poids % programme"),
        ],
    },
    {
        "label": "🚚 Fournisseurs & organisation",
        "modules": [
            ("pages/07_📈_OTIF.py",                "#1A7A3A", "📈  OTIF  ›",              "Fill Rate · On Time · watchlist GOLD/SILVER"),
            ("pages/12_🏪_Bascule_XD.py",           "#5E35B1", "🏪  Bascule XD  ›",        "DL → Cross-Docking · plan lissage · 90 FCFA/colis"),
            ("pages/09_✅_Tasks_Trackers.py",       "#48484A", "✅  Tasks Tracker  ›",     "Kanban équipe · Grace, Carine, Yves · Google Sheets"),
        ],
    },
]

def mod_card(page, color, titre, desc):
    if not Path(page).exists():
        return
    st.markdown(f"<div class='mod-wrap' style='border-left-color:{color}'>", unsafe_allow_html=True)
    st.page_link(page, label=f"{titre}\n{desc}")
    st.markdown("</div>", unsafe_allow_html=True)

for section in SECTIONS:
    # ne garder que les modules dont le fichier existe encore
    modules = [m for m in section["modules"] if Path(m[0]).exists()]
    if not modules:
        continue
    st.markdown(f"<div class='section-label'>{section['label']}</div>", unsafe_allow_html=True)
    rows = [modules[i:i+3] for i in range(0, len(modules), 3)]
    for row in rows:
        cols = st.columns(3)
        for col, (page, color, titre, desc) in zip(cols, row):
            with col:
                mod_card(page, color, titre, desc)

# ── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div style="text-align:center;color:#C7C7CC;font-size:11px;padding:8px 0">NovaRetail Solutions · v2.3 · Carrefour Côte d\'Ivoire</div>', unsafe_allow_html=True)
