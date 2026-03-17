import streamlit as st

st.set_page_config(
    page_title="SmartBuyer",
    page_icon="🛍️",
    layout="wide",
)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #FAFAFA; }
[data-testid="stSidebar"] { background: #F2F2F7; border-right: 1px solid #E5E5EA; }
.main .block-container { padding-top: 2rem; max-width: 1100px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style='margin-bottom:2rem'>
  <div style='font-size:32px;font-weight:700;color:#1C1C1E;letter-spacing:-0.03em'>
    🛍️ SmartBuyer
  </div>
  <div style='font-size:15px;color:#8E8E93;margin-top:4px'>
    Hub analytique · Équipe Achats
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
<div style='background:#fff;border:1px solid #E5E5EA;border-radius:14px;padding:20px 24px;margin-bottom:12px'>
  <div style='font-size:24px;
