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
    km_inizio = st.text_input("KM Inizio Cantiere (Basso)", "35+560")
    km_preavviso = st.text_input(
        "KM Segnale Mobile Preavviso (Centro)", "36+060"
    )

  with col2:
    km_fine = st.text_input("KM Fine / Rientro (Alto)", "36+600")
    distanza_coni = st.text_input("Interasse Coni / Dettagli", "12 m")

  # --- 1. GENERAZIONE FILE EXCEL (ABACO) ---
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
    # Apriamo l'immagine originale
    base_img = Image.open(img_path).convert("RGB")
    draw = ImageDraw.Draw(base_img)

    width, height = base_img.size

    # Proviamo a caricare un font standard grande, o usiamo un fallback
    try:
      # Scegliamo un font Arial/Sans-serif grande abbastanza (testa di default su arial)
      font_grande = ImageFont.truetype(
          "arial.ttf", int(width * 0.038)
      )  # Circa 60px per larghezza 1600
      font_medio = ImageFont.truetype("arial.ttf", int(width * 0.028))
    except:
      # Fallback su font di default searial.ttf non trovato (potrebbe apparire diverso)
      font_grande = ImageFont.load_default()
      font_medio = font_grande

    # NOTA: Qui calcoliamo e applichiamo rettangoli bianchi per coprire le vecchie scritte
    # e disegniamo sopra i nuovi testi. Le coordinate sono tarate sull'immagine 1653x2339.

    # 1. Copriamo e scriviamo il KM Fine in alto (KM 36+600)
    # Spostiamo il riquadro verso destra per non toccare la carreggiata stradale!
    # Coordinate X di pulizia: (520, 20) a (750, 90). Inseriamo testo nuovo più a destra.
    # Abbiamo spostato il riquadro e il testo a destra di circa 100 pixel.

    x_pulizia = 520
    draw.rectangle(
        [(x_pulizia, 15), (750, 95)], fill="white"
    )  # Riquadro pulizia alto
    draw.text(
        (x_pulizia + 20, 25), f"KM {km_fine}", fill="red", font=font_grande
    )  # Testo nuovo grande

    # 2. Copriamo e aggiorniamo i KM intermedi e di inizio
    # (Copriamo le coordinate X=640 e scriviamo sopra a X=650)
    draw.rectangle([(640, 535), (820, 595)], fill="white")
    draw.text((650, 542), f"KM {km_preavviso}", fill="red", font=font_medio)

    draw.rectangle([(640, 905), (820, 965)], fill="white")
    draw.text((650, 912), f"KM {km_inizio}", fill="red", font=font_medio)

    # Salviamo l'immagine modificata in memoria per il download
    img_io = io.BytesIO()
    base_img.save(img_io, "JPEG")
    img_io.seek(0)

    # Mostriamo l'anteprima corretta sul sito
    st.image(
        base_img,
        caption="Tavola 39 - Aggiornata con le progressive senza coprire la strada",
        use_container_width=True,
    )

    # Pulsante per scaricare la JPG compilata
    st.download_button(
        label="📥 Scarica Tavola in formato JPG",
        data=img_io,
        file_name="Tavola_39_Compilata.jpg",
        mime="image/jpeg",
    )

  else:
    st.warning(f"⚠️ L'immagine '{img_path}' non è stata trovata su GitHub.")
