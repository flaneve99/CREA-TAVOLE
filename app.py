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
    km_testata = st.text_input("KM Inizio Testata / Intermedio", "36+110")
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
    base_img = Image.open(img_path).convert("RGB")
    draw = ImageDraw.Draw(base_img)

    width, height = base_img.size

    # Carichiamo i font grandi e definitivi
    try:
      font_grande = ImageFont.truetype("arial.ttf", int(width * 0.045))
      font_medio = ImageFont.truetype("arial.ttf", int(width * 0.035))
    except:
      font_grande = ImageFont.load_default()
      font_medio = font_grande

    # 1. KM Fine (in alto) - Rettangolo di pulizia ben largo per cancellare ogni residuo vecchio
    box_fine = (
        int(width * 0.30),
        int(height * 0.00),
        int(width * 0.70),
        int(height * 0.07),
    )
    draw.rectangle(box_fine, fill="white")
    draw.text(
        (box_fine[0] + 10, box_fine[1] + 5),
        f"KM {km_fine}",
        fill="red",
        font=font_grande,
    )

    # 2. KM Testata / Intermedio (es. 36+110)
    box_testata = (
        int(width * 0.58),
        int(height * 0.52),
        int(width * 0.98),
        int(height * 0.61),
    )
    draw.rectangle(box_testata, fill="white")
    draw.text(
        (box_testata[0] + 10, box_testata[1] + 5),
        f"KM {km_testata}",
        fill="red",
        font=font_medio,
    )

    # 3. KM Preavviso (al centro)
    box_preavviso = (
        int(width * 0.58),
        int(height * 0.64),
        int(width * 0.98),
        int(height * 0.73),
    )
    draw.rectangle(box_preavviso, fill="white")
    draw.text(
        (box_preavviso[0] + 10, box_preavviso[1] + 5),
        f"KM {km_preavviso}",
        fill="red",
        font=font_medio,
    )

    # 4. KM Inizio (in basso)
    box_inizio = (
        int(width * 0.58),
        int(height * 0.90),
        int(width * 0.98),
        int(height * 0.99),
    )
    draw.rectangle(box_inizio, fill="white")
    draw.text(
        (box_inizio[0] + 10, box_inizio[1] + 5),
        f"KM {km_inizio}",
        fill="red",
        font=font_medio,
    )

    # Salviamo in memoria
    img_io = io.BytesIO()
    base_img.save(img_io, "JPEG")
    img_io.seek(0)

    # Mostriamo l'anteprima sul sito
    st.image(
        base_img,
        caption="Tavola 39 - Pulita e aggiornata",
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
