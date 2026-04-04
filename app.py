import streamlit as st

st.set_page_config(page_title="SmartBuyer", page_icon="🛍️", layout="wide")

st.markdown("<h1 style='font-size:32px;font-weight:700;color:#1C1C1E'>🛍️ SmartBuyer</h1>", unsafe_allow_html=True)
st.markdown("<p style='color:#8E8E93;margin-top:-10px'>Hub analytique · Équipe Achats</p>", unsafe_allow_html=True)
st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.info("📊 **Scoring ABC**  \nClassement articles · Pareto · Priorités")
    st.info("📈 **Suivi Ventes PBI**  \nCA · Marge · Évolution hebdomadaire")
    st.info("📦 **Détention Top CA**  \nTaux de détention · Alertes réseau")
    st.info("💸 **Marges Négatives**  \nDiagnostic réseau · Flop 100 · Fuites de valeur")
    st.info("✅ **Tasks Tracker**  \nSuivi tâches · Kanban · Google Sheets")

with col2:
    st.info("💸 **Performance Promotion**  \nSuivi promos · Poids CA · Marge promo")
    st.info("🏪 **Suivi Implantation**  \nTaux implantation · Statuts · Avancement")
    st.info("📈 **OTIF**  \nPerformance fournisseurs · Taux de service · Criticité")
    st.info("📦 **OOS — Ruptures**  \nDétection ruptures · Plan d'action · Cessions")

st.markdown("---")
st.caption("NovaRetail Solutions · SmartBuyer v2.1 · 9 modules actifs")
