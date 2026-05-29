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
    text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 10px;
}

/* KPI barre haute */
.kpi-bar { background: #FFFFFF; border-radius: 14px; border: 0.5px solid #E5E5EA; padding: 0.85rem 1.25rem; text-align: center; }
.kpi-bar-val   { font-size: 20px; font-weight: 700; color: #1C1C1E; }
.kpi-bar-label { font-size: 11px; color: #8E8E93; margin-top: 2px; }

/* Cards modules — style Marges Négatives */
.module-card {
    background: #FFFFFF; border: 0.5px solid #E5E5EA; border-radius: 12px;
    padding: 16px; border-left: 3px solid; margin-bottom: 10px;
}
.module-card-header { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
.module-icon  { font-size: 18px; }
.module-title { font-size: 14px; font-weight: 600; color: #1C1C1E; }
.module-desc  { font-size: 12px; color: #3A3A3C; margin-bottom: 4px; }
.module-formula {
    font-size: 11px; font-family: monospace;
    background: #F9F9FB; padding: 4px 8px; border-radius: 6px;
    margin-bottom: 6px;
}
.module-action { font-size: 11px; color: #8E8E93; font-style: italic; }

/* Tag badges */
.badge { display: inline-block; padding: 2px 8px; border-radius: 20px; font-size: 10px; font-weight: 600; margin-top: 5px; }
.tag-achats  { background: #EAF4FF; color: #007AFF; }
.tag-fourn   { background: #E8F8ED; color: #1A7A3A; }
.tag-stock   { background: #FFF3E0; color: #B45309; }
.tag-perf    { background: #F2F0FF; color: #5E35B1; }
.tag-promo   { background: #FFEAEA; color: #C0392B; }
.tag-tasks   { background: #F2F2F7; color: #48484A; }
</style>
""", unsafe_allow_html=True)

# ── EN-TÊTE ──────────────────────────────────────────────────────────────────
st.markdown("<div class='page-title'>🛍️ SmartBuyer Hub</div>", unsafe_allow_html=True)
st.markdown("<div class='page-caption'>Plateforme analytique · Équipe Achats Carrefour CI</div>", unsafe_allow_html=True)

# ── KPIs BARRE ───────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)
for col, val, label in [
    (k1, "13",       "Modules actifs"),
    (k2, "4",        "Rayons couverts"),
    (k3, "13",       "Sites réseau"),
    (k4, "v2.3",     "Version"),
    (k5, "Mai 2026", "Dernière période"),
]:
    col.markdown(f"""
    <div class="kpi-bar">
        <div class="kpi-bar-val">{val}</div>
        <div class="kpi-bar-label">{label}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
st.markdown("---")

# ── MODULES — 3 sections, 2 colonnes chacune ─────────────────────────────────

SECTIONS = [
    {
        "label": "📈 Performance commerciale",
        "modules": [
            {
                "icon": "📊", "title": "Scoring ABC",
                "color": "#007AFF", "tag": ("tag-achats", "Achats"),
                "desc": "Classe chaque article en 5 catégories d'action relatives à son Unité de Besoin.",
                "formula": "ABC Qté / Vente / Marge × Pricing × Nouveauté < N mois",
                "action": "→ Protéger, développer, arbitrer ou supprimer",
            },
            {
                "icon": "📈", "title": "Ventes PBI",
                "color": "#5E35B1", "tag": ("tag-perf", "Performance"),
                "desc": "Compare les performances hebdomadaires CA / Marge / Qté entre deux périodes PBI.",
                "formula": "Évol CA % = (CA S / CA réf − 1) × 100",
                "action": "→ Alertes automatiques sur reculs > 10% et références sans vente",
            },
            {
                "icon": "📊", "title": "Performance Hebdo",
                "color": "#5E35B1", "tag": ("tag-perf", "Performance"),
                "desc": "Top CA / Top Marge / Flop marges négatives / Top casse par rayon et réseau.",
                "formula": "Export PBI standard — 1 onglet Excel par rayon",
                "action": "→ Rapport hebdo prêt à partager à l'équipe",
            },
            {
                "icon": "💰", "title": "Rentabilité",
                "color": "#C0392B", "tag": ("tag-promo", "Rentabilité"),
                "desc": "Cockpit direction : taux de marge vs N-1, cible par segment, score de santé 0–100.",
                "formula": "Déviation = Tx Marge − Tx N-1 · Score = 50% FR + 30% OT + 20% OTIF",
                "action": "→ Briefing acheteur, plan de négociation fournisseur",
            },
            {
                "icon": "💸", "title": "Marges Négatives",
                "color": "#FF3B30", "tag": ("tag-promo", "Marge"),
                "desc": "Diagnostic réseau : Flop 100 destructeurs, matrice rayon × magasin, effet promo/casse.",
                "formula": "Marge < 0 · Δ HP−Promo · Tx Casse = Casse / CA × 100",
                "action": "→ Identifier et corriger chaque fuite de valeur",
            },
            {
                "icon": "🎯", "title": "Performance Promo",
                "color": "#FF9500", "tag": ("tag-promo", "Promo"),
                "desc": "Poids promo, taux de marge promo vs hors promo, efficacité par jour, dépendance.",
                "formula": "Efficacité = CA Promo ÷ Nb jours · Dépendance = CA HP = 0",
                "action": "→ Revoir les promos déficitaires avant renouvellement",
            },
        ],
    },
    {
        "label": "🏪 Gestion assortiment & stock",
        "modules": [
            {
                "icon": "📦", "title": "Détention Top CA",
                "color": "#B45309", "tag": ("tag-stock", "Stock"),
                "desc": "Taux de détention des articles Top CA permanents GOLD / SILVER par flux IM / LO.",
                "formula": "Détenu = stock > 0 · Taux = Nb détenus / Nb refs permanentes × 100",
                "action": "→ Commander en urgence les GOLD non détenus",
            },
            {
                "icon": "📦", "title": "Ruptures OOS",
                "color": "#B45309", "tag": ("tag-stock", "Stock"),
                "desc": "Détection ruptures article × magasin. Plan d'action : Commander ou Voir Cession.",
                "formula": "Rupture = stock ≤ 0 · Cession si stock donneur > seuil paramétrable",
                "action": "→ Liste Commander + plan de cessions inter-magasins",
            },
            {
                "icon": "🏪", "title": "Suivi Implantation",
                "color": "#1A7A3A", "tag": ("tag-achats", "Implantation"),
                "desc": "Taux d'implantation des nouvelles références T1 par magasin. Alertes article × site.",
                "formula": "Implanté = stock ≠ 0 · Appro = RAL > 0 · Commander = stock 0 + RAL 0",
                "action": "→ Accélérer livraisons, passer commandes, planifier cessions",
            },
            {
                "icon": "🏷️", "title": "Fidélité Cagnotte",
                "color": "#007AFF", "tag": ("tag-achats", "Fidélité"),
                "desc": "Investissement cagnotte vs CA et Marge fidélité. Poids % programme par famille.",
                "formula": "Budget = Cagnotte/unité × Qté Vente · Poids = CA Fidélité / CA Global",
                "action": "→ Identifier les articles à ROI nul et familles en marge négative",
            },
        ],
    },
    {
        "label": "🚚 Fournisseurs & organisation",
        "modules": [
            {
                "icon": "📈", "title": "OTIF",
                "color": "#1A7A3A", "tag": ("tag-fourn", "Fournisseur"),
                "desc": "Fill Rate / On Time / OTIF par fournisseur, magasin et article. Watchlist GOLD/SILVER.",
                "formula": "Fill Rate = Qté reçue / Qté cde · OTIF = Complet ET à l'heure",
                "action": "→ Fiche fournisseur Excel, plan de relance, récap priorisation",
            },
            {
                "icon": "🏪", "title": "Bascule XD",
                "color": "#5E35B1", "tag": ("tag-fourn", "Logistique"),
                "desc": "Analyse DL → Cross-Docking : candidats XD, plan de lissage plateforme, coût à 90 FCFA/colis.",
                "formula": "Candidat XD si valeur moyenne livraison < seuil · Coût = Colis/mois × 90 FCFA",
                "action": "→ 5 onglets Excel : décisions, plan lissage, BDD articles",
            },
            {
                "icon": "✅", "title": "Tasks Tracker",
                "color": "#48484A", "tag": ("tag-tasks", "Organisation"),
                "desc": "Suivi des tâches équipe en mode Kanban. Responsables Grace, Carine, Yves.",
                "formula": "Google Sheets · Statuts : À faire / En cours / Terminé",
                "action": "→ Vue équipe partagée, mise à jour en temps réel",
            },
        ],
    },
]

for section in SECTIONS:
    st.markdown(f"<div class='section-label'>{section['label']}</div>", unsafe_allow_html=True)

    modules = section["modules"]
    # Disposition 2 colonnes
    pairs = [modules[i:i+2] for i in range(0, len(modules), 2)]
    for pair in pairs:
        cols = st.columns(2)
        for col, mod in zip(cols, pair):
            tag_cls, tag_label = mod["tag"]
            with col:
                st.markdown(f"""
<div class="module-card" style="border-left-color:{mod['color']}">
  <div class="module-card-header">
    <span class="module-icon">{mod['icon']}</span>
    <span class="module-title">{mod['title']}</span>
  </div>
  <div class="module-desc">{mod['desc']}</div>
  <div class="module-formula">{mod['formula']}</div>
  <div class="module-action">{mod['action']}</div>
  <span class="badge {tag_cls}">{tag_label}</span>
</div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

# ── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style='text-align:center;color:#C7C7CC;font-size:11px;padding:8px 0'>
    NovaRetail Solutions · SmartBuyer v2.3 · 13 modules actifs · Carrefour Côte d'Ivoire
</div>
""", unsafe_allow_html=True)
