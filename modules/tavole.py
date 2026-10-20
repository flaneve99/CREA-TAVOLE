import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from modules.abachi import crea_excel_abaco

@st.cache_resource
def carica_immagine_tavola(percorso_immagine):
    """Carica l'immagine in cache per velocizzare l'app."""
    try:
        return Image.open(percorso_immagine).convert("RGB")
    except Exception:
        return None
