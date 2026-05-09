import streamlit as st

st.set_page_config(page_title="SmartBuyer", page_icon="🛍️", layout="wide")

st.markdown("""
<style>
body, [data-testid="stAppViewContainer"] { background: #F2F2F7; }
[data-testid="stSidebar"] { background: #FFFFFF; border-right: 0.5px solid #E5E5EA; }
.block-container { padding-top: 1.5rem; }

.module-card {
    background: #FFFFFF;
    border-radius: 14px;
    border: 0.5px solid #E5E5EA;
    padding: 1rem 1.25rem;
    margin-bottom: 0.75rem;
    display: flex;
    align-items: flex-start;
    gap: 12px;
    transition: border-color 0.15s;
}
.module-card:hover { border-color: #007AFF; }
.module-icon {
    font-size: 22px;
    line-height: 1;
    margin-top: 2px;
    flex-shrink: 0;
}
.module-title {
    font-size: 14px;
    font-weight: 600;
    color: #1C1C1E;
    margin-bottom: 2px;
}
.module-desc {
    font-size: 12px;
    color: #8E8E93;
    line-height: 1.4;
}
.module-tag {
    display: inline-block;
    font-size: 10px;
    font-weight: 600;
    padding: 1px 7px;
    border-radius: 20px;
    margin-top: 5px;
}
.tag-achats    { background: #EAF4FF; color: #007AFF; }
.tag-fourn     { background: #E8F8ED; color: #1A7A3A; }
.tag-stock     { background: #FFF3E0; color: #B45309; }
.tag-perf      { background: #F2F0FF; color: #5E35B1; }
.tag-promo     { background: #FFEAEA; color: #C0392B; }
.tag-tasks     { background: #F2F2F7; color: #48484A; }

.section-label {
    font-size: 11px;
    font-weight: 600;
    color: #8E8E93;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin: 1.25rem 0 0.5rem;
}
.kpi-bar {
    background: #FFFFFF;
    border-radius: 14px;
    border: 0.5px solid #E5E5EA;
    padding: 0.85rem 1.25rem;
    text-align: center;
}
.kpi-bar-val   { font-size: 20px; font-weight: 700; color: #1C1C1E; }
.kpi-bar-label { font-size: 11px; color: #8E8E93; margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

# ── EN-TÊTE ──────────────────────────────────────────────────────────────────
st.markdown("<h1 style='font-size:30px;font-weight:700;color:#1C1C1E;margin-bottom:2px'>🛍️ SmartBuyer Hub</h1>", unsafe_allow_html=True)
st.markdown("<p style='color:#8E8E93;margin-top:0;margin-bottom:1rem;font-size:14px'>Plateforme analytique · Équipe Achats Carrefour CI</p>", unsafe_allow_html=True)

# ── KPIs BARRE ───────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)
for col, val, label in [
    (k1, "11", "Modules actifs"),
    (k2, "4", "Rayons couverts"),
    (k3, "3", "Hypers"),
    (k4, "v2.2", "Version"),
    (k5, "Avr 2026", "Dernière période"),
]:
    col.markdown(f"""
    <div class="kpi-bar">
        <div class="kpi-bar-val">{val}</div>
        <div class="kpi-bar-label">{label}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ── MODULES ──────────────────────────────────────────────────────────────────
MODULES = [
    {
        "cat": "📈 Performance commerciale",
        "items": [
            ("📊", "Analyse Scoring ABC",   "Classement articles · Pareto 80/95 · Priorisation assortiment",          "tag-achats", "Achats"),
            ("📈", "Ventes PBI",            "CA · Marge · Évolution hebdomadaire · Flash J-1",                         "tag-perf",   "Performance"),
            ("📊", "Performance Hebdo",     "Revue hebdo Épicerie / Boissons / DPH · Export Excel",                    "tag-perf",   "Performance"),
            ("💰", "Rentabilité",           "Marge / Promo / Casse · Verdicts par rayon · Alertes seuils",             "tag-promo",  "Rentabilité"),
            ("💸", "Marges Négatives",      "Diagnostic réseau · Flop 100 · Fuites de valeur · Alertes urgentes",      "tag-promo",  "Marge"),
            ("🎯", "Performance Promo",     "Suivi promos VSD vs PBI · Poids CA · Marge promo · Alertes écart",        "tag-promo",  "Promo"),
        ]
    },
    {
        "cat": "🏪 Gestion assortiment & stock",
        "items": [
            ("📦", "Detention Top CA",      "Taux de détention · Alertes réseau · Analyse IM / LO",                    "tag-stock",  "Stock"),
            ("📦", "OOS — Ruptures",        "Détection ruptures · Plan d'action · Cessions inter-magasins",            "tag-stock",  "Stock"),
            ("🏪", "Suivi Implantation",    "Taux implantation T1 · Statuts articles · Avancement réseau",             "tag-achats", "Implantation"),
        ]
    },
    {
        "cat": "🚚 Fournisseurs & tâches",
        "items": [
            ("📈", "OTIF",                  "Performance fournisseurs · Taux de service · Criticité · Watchlist",      "tag-fourn",  "Fournisseur"),
            ("✅", "Tasks Trackers",         "Suivi tâches équipe · Kanban · Responsables · Google Sheets",             "tag-tasks",  "Organisation"),
        ]
    },
]

for section in MODULES:
    st.markdown(f"<div class='section-label'>{section['cat']}</div>", unsafe_allow_html=True)
    cols = st.columns(3)
    for i, (icon, title, desc, tag_cls, tag_label) in enumerate(section["items"]):
        with cols[i % 3]:
            st.markdown(f"""
            <div class="module-card">
                <div class="module-icon">{icon}</div>
                <div>
                    <div class="module-title">{title}</div>
                    <div class="module-desc">{desc}</div>
                    <span class="module-tag {tag_cls}">{tag_label}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

# ── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
st.markdown("""
<div style='text-align:center;color:#C7C7CC;font-size:11px;padding:8px 0'>
    NovaRetail Solutions · SmartBuyer v2.2 · 11 modules actifs · Carrefour Côte d'Ivoire
</div>
""", unsafe_allow_html=True)
