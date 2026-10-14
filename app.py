import io
import os
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

st.set_page_config(
    page_title="Gestione Cantieri Stradali - Abaco & Tavole", layout="wide"
)

st.title("🚧 Gestione Cantieri Stradali e Segnaletica")
st.write(
    "Inserisci i dati chilometrici: il sito aggiornerà l'Abaco Excel e"
    " genererà la Tavola in JPG personalizzata!"
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

  # --- GENERAZIONE FILE EXCEL (ABACO) ---
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

    st.success("Abaco Excel generato con successo!")
    st.download_button(
        label="📥 Scarica Excel Abaco Segnaletica",
        data=excel_data,
        file_name="Abaco_Segnaletica_Tavola_39.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  st.divider()
  st.subheader("🖼️ Anteprima Tavola 39 Personalizzata")

  img_path = "TORRE GUACETO 1 DIR SUD.jpg"
  if os.path.exists(img_path):
    base_img = Image.open(img_path).convert("RGB")
    draw = ImageDraw.Draw(base_img)

    width, height = base_img.size

    # Carichiamo un font scalato o di default
    try:
      font_grande = ImageFont.truetype(
          "arial.ttf", int(width * 0.04)
      )  # Proporzionale alla larghezza
      font_medio = ImageFont.truetype("arial.ttf", int(width * 0.03))
    except:
      font_grande = ImageFont.load_default()
      font_medio = font_grande

    # Coordinate percentuali o basate sulle dimensioni reali per coprire esattamente la scritta in alto "KM 36+600"
    # Sulla tua tavola, la scritta in alto si trova orientativamente nella parte superiore destra
    box_x1, box_y1, box_x2, box_y2 = (
        int(width * 0.40),
        int(height * 0.01),
        int(width * 0.75),
        int(height * 0.08),
    )

    # 1. Copriamo la vecchia scritta in alto con un rettangolo bianco pulito
    draw.rectangle([box_x1, box_y1, box_x2, box_y2], fill="white")

    # 2. Scriviamo il nuovo KM Fine
    draw.text(
        (box_x1 + 10, box_y1 + 5), f"KM {km_fine}", fill="red", font=font_grande
    )

    # Salviamo l'immagine modificata in memoria
    img_io = io.BytesIO()
    base_img.save(img_io, "JPEG")
    img_io.seek(0)

    # Mostriamo l'anteprima
    st.image(
        base_img,
        caption="Tavola 39 - Aggiornata con le progressive",
        use_container_width=True,
    )

    st.download_button(
        label="📥 Scarica Tavola in formato JPG",
        data=img_io,
        file_name="Tavola_39_Compilata.jpg",
        mime="image/jpeg",
    )

  else:
    st.warning(f"⚠️ L'immagine '{img_path}' non è stata trovata su GitHub.")
