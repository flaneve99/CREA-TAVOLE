import streamlit as st

def setup_page():
    st.set_page_config(
        page_title="Gestione Cantiere",
        page_icon="🚧",
        layout="wide",
        initial_sidebar_state="expanded"
    )

def inject_custom_css():
    st.markdown("""
        <style>
        .stApp { background-color: #f8f9fa; }
        .main-header { font-size: 2rem; font-weight: 700; color: #1e3a8a; margin-bottom: 1rem; }
        </style>
    """, unsafe_allow_html=True)
