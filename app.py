import re
from pathlib import Path

import streamlit as st

st.set_page_config(page_title="NovaRetail Solutions", page_icon="🛍️", layout="wide")

# ════════════════════════════════════════════════════════════════════════════
#  REGISTRE DES MODULES
#  Clé = préfixe numérique du fichier (stable). Un module absent du dossier
#  pages/ n'apparaît nulle part ; un module présent mais inconnu ici s'affiche
#  quand même (rubrique « Autres ») avec un libellé déduit du nom de fichier.
# ════════════════════════════════════════════════════════════════════════════
PAGES_DIR = Path("pages")

# num : (nom affiché, description, couleur, section)
META = {
    "01": ("Scoring ABC",            "Classification articles · 5 règles de reco",       "#007AFF", "perf"),
    "02": ("Ventes PBI",             "Comparaison hebdo CA / Marge · 2 périodes",        "#5E35B1", "perf"),
    "03": ("Détention Top CA",       "GOLD / SILVER · IM / LO · articles permanents",    "#B45309", "assort"),
    "04": ("Performance Promo",      "Poids promo · taux marge · dépendance",            "#FF9500", "perf"),
    "05": ("Suivi Implantation",     "Taux T1 · alertes article × site · cessions",      "#1A7A3A", "assort"),
    "06": ("Marges Négatives",       "Flop 100 · matrice rayon × magasin · casse",       "#FF3B30", "perf"),
    "07": ("OTIF",                   "Fill Rate · On Time · watchlist GOLD/SILVER",      "#1A7A3A", "fourn"),
    "08": ("Ruptures OOS",           "Commander ou céder · seuil paramétrable",          "#B45309", "assort"),
    "09": ("Tasks Tracker",          "Kanban équipe · Google Sheets",                    "#48484A", "fourn"),
    "10": ("Perf Hebdo",             "Top CA · Flop marges · Top casse par rayon",       "#5E35B1", "perf"),
    "11": ("Rentabilité",            "Cockpit direction · score santé · vs N-1",         "#C0392B", "perf"),
    "12": ("Bascule XD",             "DL → Cross-Docking · plan lissage",                "#5E35B1", "fourn"),
    "13": ("Fidélité Cagnotte",      "Budget cagnotte · poids % programme",              "#007AFF", "assort"),
    "15": ("COPIL Hebdo",            "COPIL exécutif · Destructeurs / Performeurs",      "#C0392B", "reporting"),
    "16": ("Reporting Ventes",       "Reporting COPIL · 1 onglet par rayon",             "#007AFF", "reporting"),
    "17": ("Reporting Sous-Familles","Alertes marge Sous-Famille × Site · 3 paliers",    "#FF9500", "reporting"),
    "18": ("Reporting Vente CA",     "Détection CA / Flop · 4 critères C1–C4 · scoring",  "#FF3B30", "reporting"),
    "19": ("Reporting Promo",        "Poids promo · taux marge · dépendance",            "#34C759", "reporting"),
}

# ordre + libellés des rubriques (celles non vides s'affichent)
SECTIONS = [
    ("reporting", "📊 Reporting & COPIL"),
    ("perf",      "📈 Performance commerciale"),
    ("assort",    "🏪 Assortiment & stock"),
    ("fourn",     "🚚 Fournisseurs & organisation"),
    ("autres",    "🧩 Autres modules"),
]

DEFAULT_COLOR = "#48484A"


def discover_pages():
    """Scanne pages/ et retourne les modules réellement présents, enrichis du registre."""
    found = []
    for p in sorted(PAGES_DIR.glob("*.py")):
        if p.name.startswith("_"):          # _utils, _archive… ignorés
            continue
        parts = p.stem.split("_")
        if not parts or not parts[0].isdigit():
            continue
        num   = parts[0]
        emoji = parts[1] if len(parts) > 1 else "•"
        raw   = " ".join(parts[2:]) if len(parts) > 2 else p.stem
        name, desc, color, section = META.get(num, (raw, "", DEFAULT_COLOR, "autres"))
        found.append({
            "num": num, "emoji": emoji, "name": name, "desc": desc,
            "color": color, "section": section, "path": f"pages/{p.name}",
        })
    return found


PAGES = discover_pages()


def pages_of(section_key):
    return [m for m in PAGES if m["section"] == section_key]


# ── STYLE ────────────────────────────────────────────────────────────────────
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
/* on masque la nav auto de Streamlit : on gère la nôtre, groupée par thème */
[data-testid="stSidebarNav"] { display: none !important; }

.page-title   { font-size: 28px; font-weight: 700; color: #1C1C1E; letter-spacing: -0.03em; margin: 0; }
.page-caption { font-size: 13px; color: #8E8E93; margin-top: 3px; margin-bottom: 1.5rem; }
.section-label {
    font-size: 11px; font-weight: 600; color: #8E8E93;
    text-transform: uppercase; letter-spacing: 0.07em;
    margin: 1.4rem 0 0.6rem; padding-bottom: 6px;
    border-bottom: 0.5px solid #E5E5EA;
}
.kpi-bar { background: #FFFFFF; border-radius: 14px; border: 0.5px solid #E5E5EA; padding: 0.85rem 1.25rem; text-align: center; }
.kpi-bar-val   { font-size: 20px; font-weight: 700; color: #1C1C1E; }
.kpi-bar-label { font-size: 11px; color: #8E8E93; margin-top: 2px; }

/* libellés de rubrique dans la sidebar */
.side-group { font-size: 10px; font-weight: 700; color: #8E8E93; text-transform: uppercase; letter-spacing: .06em; margin: 14px 0 4px; }

/* Carte récap autour du st.page_link */
.mod-wrap {
    border-radius: 12px; overflow: hidden; margin-bottom: 8px;
    border: 0.5px solid #E5E5EA; border-left-width: 3px; background: #FFFFFF;
    height: 100%;
}
.mod-wrap [data-testid="stPageLink"] {
    background: transparent !important; border: none !important;
    border-radius: 0 !important; display: block !important; width: 100% !important;
}
.mod-wrap [data-testid="stPageLink"]:hover { background: #F9F9FB !important; }
.mod-wrap [data-testid="stPageLink"] p {
    font-size: 13px !important; font-weight: 500 !important; color: #1C1C1E !important;
    padding: 12px 14px !important; margin: 0 !important;
    white-space: pre-line !important; line-height: 1.5 !important;
}
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR : navigation groupée (auto) ──────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:8px;padding-bottom:16px;border-bottom:0.5px solid #E5E5EA'>
  <div style='font-size:17px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>NovaRetail Solutions</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:2px'>Hub analytique · Équipe Achats CI</div>
</div>""", unsafe_allow_html=True)

    for sec_key, sec_label in SECTIONS:
        group = pages_of(sec_key)
        if not group:
            continue
        st.markdown(f"<div class='side-group'>{sec_label}</div>", unsafe_allow_html=True)
        for m in group:
            st.page_link(m["path"], label=f"{m['emoji']}  {m['name']}")

# ── EN-TÊTE ──────────────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>NovaRetail Solutions</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Plateforme analytique · Équipe Achats Carrefour CI</div>", unsafe_allow_html=True)

# ── KPIs (compteur de modules automatique) ───────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)
for col, val, label in [
    (k1, str(len(PAGES)), "Modules actifs"), (k2, "4", "Rayons couverts"),
    (k3, "13", "Sites réseau"),              (k4, "v2.4", "Version"),
    (k5, "Mai 2026", "Dernière période"),
]:
    col.markdown(f'<div class="kpi-bar"><div class="kpi-bar-val">{val}</div><div class="kpi-bar-label">{label}</div></div>', unsafe_allow_html=True)

# ── BLOC PRINCIPAL : récap de toutes les apps ────────────────────────────────
def mod_card(m):
    st.markdown(f"<div class='mod-wrap' style='border-left-color:{m['color']}'>", unsafe_allow_html=True)
    titre = f"{m['emoji']}  {m['name']}  ›"
    label = f"{titre}\n{m['desc']}" if m["desc"] else titre
    st.page_link(m["path"], label=label)
    st.markdown("</div>", unsafe_allow_html=True)

for sec_key, sec_label in SECTIONS:
    group = pages_of(sec_key)
    if not group:
        continue
    st.markdown(f"<div class='section-label'>{sec_label}</div>", unsafe_allow_html=True)
    rows = [group[i:i + 3] for i in range(0, len(group), 3)]
    for row in rows:
        cols = st.columns(3)
        for col, m in zip(cols, row):
            with col:
                mod_card(m)

# ── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    f'<div style="text-align:center;color:#C7C7CC;font-size:11px;padding:8px 0">'
    f'NovaRetail Solutions · v2.4 · {len(PAGES)} modules · Carrefour Côte d\'Ivoire</div>',
    unsafe_allow_html=True,
)
