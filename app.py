import streamlit as st

st.set_page_config(page_title="SmartBuyer", page_icon="🛍️", layout="wide")

st.markdown("""
<style>
body, [data-testid="stAppViewContainer"] { background: #F2F2F7; }
[data-testid="stSidebar"] { background: #FFFFFF; border-right: 0.5px solid #E5E5EA; }
.block-container { padding-top: 1.5rem; }

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

/* Card module — fond + bordures seulement, pas de hover interactif (géré par st.page_link) */
.mod-card {
    background: #FFFFFF;
    border: 0.5px solid #E5E5EA;
    border-radius: 12px;
    padding: 14px 16px;
    border-left: 3px solid;
    height: 100%;
}
.mod-title { font-size: 14px; font-weight: 600; color: #1C1C1E; margin-bottom: 6px; }
.mod-desc  { font-size: 12px; color: #3A3A3C; margin-bottom: 5px; line-height: 1.5; }
.mod-formula {
    font-size: 11px; font-family: monospace;
    background: #F9F9FB; padding: 3px 8px;
    border-radius: 6px; color: #8E8E93;
    display: inline-block; margin-bottom: 6px;
}
.mod-action { font-size: 11px; color: #8E8E93; font-style: italic; }
.badge { display: inline-block; padding: 2px 8px; border-radius: 20px; font-size: 10px; font-weight: 600; }
.b-blue  { background: #EAF4FF; color: #007AFF; }
.b-green { background: #E8F8ED; color: #1A7A3A; }
.b-amber { background: #FFF3E0; color: #B45309; }
.b-red   { background: #FFEAEA; color: #C0392B; }
.b-gray  { background: #F2F2F7; color: #48484A; }
.b-purp  { background: #F2F0FF; color: #5E35B1; }

/* Masquer le label du st.page_link pour n'afficher que la card */
[data-testid="stPageLink"] { text-decoration: none !important; }
[data-testid="stPageLink"] p { display: none !important; }
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR navigation ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<div style='margin-bottom:18px'>
  <div style='font-size:20px;font-weight:700;color:#1C1C1E;letter-spacing:-0.02em'>🛍️ SmartBuyer</div>
  <div style='font-size:11px;color:#8E8E93;margin-top:1px'>Hub analytique · Équipe Achats</div>
</div>""", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("<div style='font-size:11px;font-weight:600;color:#8E8E93;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px'>Navigation</div>", unsafe_allow_html=True)
    st.page_link("app.py",                                    label="🏠  Accueil")
    st.page_link("pages/01_📊_Analyse_Scoring_ABC.py",        label="📊  Scoring ABC")
    st.page_link("pages/02_📈_Ventes_PBI.py",                 label="📈  Ventes PBI")
    st.page_link("pages/03_📦_Detention_Top_CA.py",           label="📦  Détention Top CA")
    st.page_link("pages/04_💸_Performance_Promo.py",          label="💸  Performance Promo")
    st.page_link("pages/05_🏪_Suivi_Implantation.py",         label="🏪  Suivi Implantation")
    st.page_link("pages/06_💸_Marges_Negatives.py",           label="💸  Marges Négatives")
    st.page_link("pages/07_📈_OTIF.py",                       label="📈  OTIF")
    st.page_link("pages/08_📦_OOS.py",                        label="📦  Ruptures OOS")
    st.page_link("pages/09_✅_Tasks_Trackers.py",             label="✅  Tasks Tracker")
    st.page_link("pages/10_📊_Perf_Hebdo.py",                 label="📊  Perf Hebdo")
    st.page_link("pages/11_📊_Rentabilite.py",                label="📊  Rentabilité")
    st.page_link("pages/12_🏪_Bascule_XD.py",                 label="🏪  Bascule XD")
    st.page_link("pages/13_🏷️_Fidelite_Cagnotte.py",         label="🏷️  Fidélité Cagnotte")

# ── EN-TÊTE ──────────────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>🛍️ SmartBuyer Hub</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Plateforme analytique · Équipe Achats Carrefour CI · Clique sur un module pour l'ouvrir</div>", unsafe_allow_html=True)

# ── KPIs BARRE ───────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)
for col, val, label in [
    (k1, "13",       "Modules actifs"),
    (k2, "4",        "Rayons couverts"),
    (k3, "13",       "Sites réseau"),
    (k4, "v2.3",     "Version"),
    (k5, "Mai 2026", "Dernière période"),
]:
    col.markdown(f"""<div class="kpi-bar">
        <div class="kpi-bar-val">{val}</div>
        <div class="kpi-bar-label">{label}</div>
    </div>""", unsafe_allow_html=True)

# ── MODULES ──────────────────────────────────────────────────────────────────
# Chaque module : (page, icon, title, color, badge_cls, badge, desc, formula, action)
SECTIONS = [
    {
        "label": "📈 Performance commerciale",
        "modules": [
            ("pages/01_📊_Analyse_Scoring_ABC.py",  "📊", "Scoring ABC",    "#007AFF", "b-blue",  "Achats",
             "Classification articles par UBD · 5 règles de recommandation.",
             "ABC Qté / Vente / Marge × Pricing × Nouveauté",
             "→ Protéger, développer, arbitrer ou supprimer"),
            ("pages/02_📈_Ventes_PBI.py",            "📈", "Ventes PBI",     "#5E35B1", "b-purp",  "Performance",
             "Comparaison hebdo CA / Marge / Qté entre deux périodes PBI.",
             "Évol CA % = (CA S / CA réf − 1) × 100",
             "→ Alertes reculs > 10% et références sans vente"),
            ("pages/10_📊_Perf_Hebdo.py",            "📊", "Perf Hebdo",     "#5E35B1", "b-purp",  "Performance",
             "Top CA / Top Marge / Flop marges négatives / Top casse par rayon.",
             "Export PBI standard — 1 onglet Excel par rayon",
             "→ Rapport hebdo prêt à partager"),
            ("pages/11_📊_Rentabilite.py",           "💰", "Rentabilité",    "#C0392B", "b-red",   "Rentabilité",
             "Cockpit direction : taux marge vs N-1, cible par segment, score santé 0–100.",
             "Déviation = Tx Marge − Tx N-1 · Score santé / 100",
             "→ Briefing acheteur + plan de négociation"),
            ("pages/06_💸_Marges_Negatives.py",      "💸", "Marges Négatives","#FF3B30", "b-red",   "Marge",
             "Flop 100 destructeurs · matrice rayon × magasin · effet promo/casse.",
             "Marge < 0 · Δ HP−Promo · Tx Casse = Casse / CA × 100",
             "→ Identifier et corriger chaque fuite de valeur"),
            ("pages/04_💸_Performance_Promo.py",     "🎯", "Perf Promo",     "#FF9500", "b-amber", "Promo",
             "Poids promo · taux marge promo vs HP · efficacité · dépendance.",
             "Efficacité = CA Promo ÷ Nb jours",
             "→ Revoir les promos déficitaires avant renouvellement"),
        ],
    },
    {
        "label": "🏪 Assortiment & stock",
        "modules": [
            ("pages/03_📦_Detention_Top_CA.py",      "📦", "Détention Top CA","#B45309", "b-amber", "Stock",
             "Taux de détention GOLD / SILVER · flux IM / LO · articles permanents.",
             "Détenu = stock > 0 · Taux = Nb détenus / Nb refs × 100",
             "→ Commander en urgence les GOLD non détenus"),
            ("pages/08_📦_OOS.py",                   "📦", "Ruptures OOS",   "#B45309", "b-amber", "Stock",
             "Détection ruptures article × magasin. Commander ou Voir Cession.",
             "Cession si stock donneur > seuil paramétrable",
             "→ Liste Commander + plan cessions inter-magasins"),
            ("pages/05_🏪_Suivi_Implantation.py",    "🏪", "Implantation",   "#1A7A3A", "b-green", "Implant.",
             "Taux d'implantation T1 par magasin. Alertes article × site.",
             "Implanté = stock ≠ 0 · Appro = RAL > 0",
             "→ Accélérer livraisons, passer commandes, céder"),
            ("pages/13_🏷️_Fidelite_Cagnotte.py",    "🏷️", "Fidélité Cagnotte","#007AFF", "b-blue", "Fidélité",
             "Budget cagnotte vs CA et Marge fidélité. Poids % programme.",
             "Budget = Cagnotte/unité × Qté Vente",
             "→ Articles à ROI nul + familles marge négative"),
        ],
    },
    {
        "label": "🚚 Fournisseurs & organisation",
        "modules": [
            ("pages/07_📈_OTIF.py",                  "📈", "OTIF",           "#1A7A3A", "b-green", "Fournisseur",
             "Fill Rate / On Time / OTIF par fournisseur, magasin, article.",
             "Fill Rate = Qté reçue / Qté cde × 100",
             "→ Fiche fournisseur Excel + récap priorisation"),
            ("pages/12_🏪_Bascule_XD.py",            "🏪", "Bascule XD",     "#5E35B1", "b-purp",  "Logistique",
             "DL → Cross-Docking : candidats, plan lissage, coût 90 FCFA/colis.",
             "Candidat XD si valeur moy. livraison < seuil",
             "→ 5 onglets Excel : décisions, plan, BDD articles"),
            ("pages/09_✅_Tasks_Trackers.py",         "✅", "Tasks Tracker",  "#48484A", "b-gray",  "Organisation",
             "Kanban équipe · Grace, Carine, Yves · Google Sheets.",
             "À faire / En cours / Terminé",
             "→ Vue partagée mise à jour en temps réel"),
        ],
    },
]

def module_card(page, icon, title, color, badge_cls, badge, desc, formula, action):
    """Affiche une card cliquable via st.page_link() superposé."""
    st.markdown(f"""
<div class="mod-card" style="border-left-color:{color}">
  <div class="mod-title">{icon} {title}</div>
  <div class="mod-desc">{desc}</div>
  <div class="mod-formula">{formula}</div>
  <div style="display:flex;align-items:center;justify-content:space-between;margin-top:4px">
    <div class="mod-action">{action}</div>
    <span class="badge {badge_cls}">{badge}</span>
  </div>
</div>
""", unsafe_allow_html=True)
    st.page_link(page, label=f"Ouvrir {title} →")

for section in SECTIONS:
    st.markdown(f"<div class='section-label'>{section['label']}</div>", unsafe_allow_html=True)
    modules = section["modules"]
    # 3 colonnes
    rows = [modules[i:i+3] for i in range(0, len(modules), 3)]
    for row in rows:
        cols = st.columns(3)
        for col, mod in zip(cols, row):
            with col:
                module_card(*mod)

# ── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style='text-align:center;color:#C7C7CC;font-size:11px;padding:8px 0'>
    NovaRetail Solutions · SmartBuyer v2.3 · 13 modules · Carrefour Côte d'Ivoire
</div>
""", unsafe_allow_html=True)
