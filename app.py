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

    # Carichiamo un font standard o di fallback ingrandito
    try:
      font_grande = ImageFont.truetype("arial.ttf", 45)
      font_medio = ImageFont.truetype("arial.ttf", 36)
    except:
      font_grande = ImageFont.load_default()
      font_medio = font_grande

    # 1. Copriamo la scritta in alto ("KM 36+600") con un rettangolo bianco e mettiamo la nuova
    # (Coordinate tarate sulle dimensioni reali della tavola)
    draw.rectangle([(420, 10), (750, 90)], fill="white")
    draw.text((430, 20), f"KM {km_fine}", fill="red", font=font_grande)

    # 2. Copriamo e aggiorniamo i KM intermedi e di inizio se necessario
    draw.rectangle([(640, 535), (820, 595)], fill="white")
    draw.text(
        (645, 542), f"KM {km_preavviso}", fill="red", font=font_medio
    )  # KM 36+060

    draw.rectangle([(640, 905), (820, 965)], fill="white")
    draw.text((645, 912), f"KM {km_inizio}", fill="red", font=font_medio)  # KM 35+560

    # Salviamo in memoria
    img_io = io.BytesIO()
    base_img.save(img_io, "JPEG")
    img_io.seek(0)

    # Mostriamo l'anteprima corretta
    st.image(
        base_img,
        caption="Tavola 39 - Con progressive aggiornate e pulite",
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
