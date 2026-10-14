import io
import os
import pandas as pd
import streamlit as st
from PIL import Image

st.set_page_config(
    page_title="Gestione Cantieri Stradali - Abaco & Tavole", layout="wide"
)

st.title("🚧 Gestione Cantieri Stradali e Segnaletica")
st.write(
    "Seleziona la tavola, inserisci i dati chilometrici e scarica l'abaco in"
    " Excel e la tavola compilata."
)

# Selezione della Tavola
tavola_selezionata = st.selectbox(
    "Seleziona la Tavola di Cantiere",
    ["Tavola 39 - Cantiere mobile carreggiata a due corsie"],
)

if "39" in tavola_selezionata:
  st.subheader("📍 Dati e Progressive Chilometriche - Tavola 39")
  st.markdown(
      "*Direzione Sud, Viadotto Torre Guaceto 1 - Chiusura corsia di destra*"
  )

  col1, col2 = st.columns(2)

  with col1:
    km_inizio = st.text_input("KM Inizio Cantiere", "35+560")
    km_preavviso = st.text_input("KM Segnale Mobile Preavviso", "36+060")

  with col2:
    km_fine = st.text_input("KM Fine / Rientro", "36+600")
    distanza_coni = st.text_input("Interasse Coni / Dettagli", "12 m")

  # Pulsante per generare l'Abaco Excel
  if st.button("Genera Abaco Segnaletica (Excel)"):
    dati_abaco = {
        "SEGNALE": [
            "Preavviso Cantiere Mobile",
            "Segnale Mobile di Protezione",
            "Via Libera / Fine Cantiere",
        ],
        "FIGURA": ["FIG. 400", "FIG. 401", "FIG. 070"],
        "DESCRIZIONE": [
            "SEGNALE MOBILE DI PREAVVISO",
            "SEGNALE MOBILE DI PROTEZIONE",
            "VIA LIBERA",
        ],
        "PROGRESSIVA CHILOMETRICA": [km_preavviso, km_inizio, km_fine],
        "SEGNALETICA": [
            "SEGNALI TEMPORANEI",
            "SEGNALI TEMPORANEI",
            "SEGNALI DI PRESCRIZIONE",
        ],
    }

    df_risultato = pd.DataFrame(dati_abaco)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_risultato.to_excel(writer, index=False, sheet_name="Abaco Tavola 39")
    excel_data = output.getvalue()

    st.success("Abaco generato con successo!")

    st.download_button(
        label="📥 Scarica Excel Abaco Segnaletica",
        data=excel_data,
        file_name="Abaco_Segnaletica_Tavola_39.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  st.divider()
  st.subheader("👀 Anteprima Tavola 39")

  img_path = "TORRE GUACETO 1 DIR SUD.jpg"
  if os.path.exists(img_path):
    image = Image.open(img_path)
    st.image(
        image,
        caption=(
            f"Tavola 39 - Con progressive: {km_inizio}, {km_preavviso},"
            f" {km_fine}"
        ),
        use_container_width=True,
    )
  else:
    st.warning(
        f"⚠️ L'immagine '{img_path}' non è stata ancora caricata su GitHub."
        " Caricala per vederla nell'anteprima."
    )
