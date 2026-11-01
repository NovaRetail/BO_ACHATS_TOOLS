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

# num : (nom affiché, description, couleur, rubrique) — l'ordre ci-dessous est l'ordre d'affichage
META = {
    # Pilotage
    "27": ("Flash Quotidien",         "Ventes de la veille · secteurs · retours attendus · PowerPoint", "#F09F39", "pilotage"),
    "26": ("Revue CODIR",             "Rapport hebdo à la DG · entonnoir Vs N-1 · PowerPoint",  "#24235C", "pilotage"),
    "24": ("PGC Radar",               "Réel vs Budget vs N-1 · alertes à traiter",             "#5E35B1", "pilotage"),
    # Ventes et marge
    "18": ("Reporting Vente CA",      "CA en recul · 4 critères de détection · scoring",       "#FF3B30", "ventes"),
    "17": ("Reporting Sous-Familles", "Alertes marge sous-famille × magasin · 3 paliers",      "#FF9500", "ventes"),
    "22": ("Suivi Rayon vs Pays",     "Rayon vs tendance pays · magasins à regarder",          "#007AFF", "ventes"),
    "06": ("Marge négative",          "Ventes à perte article × magasin · Flop 100",           "#FF3B30", "ventes"),
    # Assortiment et fournisseurs
    "05": ("Suivi Implantation",      "Taux T1 · alertes article × magasin · cessions",        "#1A7A3A", "assort"),
    "07": ("OTIF",                    "Taux de service fournisseurs · plan d’action acheteurs", "#1A7A3A", "assort"),
}
ORDER = list(META)

# ordre + libellés des rubriques (celles non vides s'affichent)
SECTIONS = [
    ("pilotage", "📋 Pilotage"),
    ("ventes",   "📊 Ventes et marge"),
    ("assort",   "🏪 Assortiment et fournisseurs"),
    ("autres",   "🧩 Autres modules"),
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
    group = [m for m in PAGES if m["section"] == section_key]
    return sorted(group, key=lambda m: ORDER.index(m["num"]) if m["num"] in ORDER else 99)


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

/* Cartes de modules : st.container(border=True) */
[data-testid="stMain"] [data-testid="stVerticalBlockBorderWrapper"] {
    background: #FFFFFF; border-radius: 12px !important; border: 0.5px solid #E5E5EA !important;
    padding: 2px 4px 6px !important;
}
.mod-bar  { height: 3px; width: 34px; border-radius: 2px; margin: 6px 0 2px 10px; }
.mod-desc { font-size: 12px; color: #8E8E93; margin: -6px 0 2px 10px; line-height: 1.4; min-height: 34px; }
[data-testid="stMain"] [data-testid="stPageLink"] p { font-size: 14px !important; font-weight: 600 !important; color: #1C1C1E !important; }
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR : navigation groupée (auto) ──────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:8px;padding-bottom:16px;border-bottom:0.5px solid #E5E5EA'>
  <div style='font-size:17px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>NovaRetail Solutions</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:2px'>Hub analytique · Achats Carrefour CI</div>
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
st.markdown("<div class='page-caption'>Plateforme analytique · Achats et category management · Carrefour Côte d\'Ivoire</div>", unsafe_allow_html=True)

# ── KPIs (compteur de modules automatique) ───────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)
for col, val, label in [
    (k1, str(len(PAGES)), "Modules actifs"),
    (k2, str(sum(1 for k, _ in SECTIONS if pages_of(k))), "Rubriques"),
    (k3, "3", "Secteurs couverts"),
    (k4, "12", "Magasins"),
    (k5, "v2.5", "Version"),
]:
    col.markdown(f'<div class="kpi-bar"><div class="kpi-bar-val">{val}</div><div class="kpi-bar-label">{label}</div></div>', unsafe_allow_html=True)

# ── BLOC PRINCIPAL : récap de toutes les apps ────────────────────────────────
def mod_card(m):
    with st.container(border=True):
        st.markdown(f"<div class='mod-bar' style='background:{m['color']}'></div>", unsafe_allow_html=True)
        st.page_link(m["path"], label=f"{m['emoji']}  {m['name']}  ›")
        st.markdown(f"<div class='mod-desc'>{m['desc'] or '&nbsp;'}</div>", unsafe_allow_html=True)


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
    f'NovaRetail Solutions · v2.5 · {len(PAGES)} modules · Carrefour Côte d\'Ivoire</div>',
    unsafe_allow_html=True,
)
