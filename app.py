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

  img_path = "TAVOLA 39.jpg"

  if os.path.exists(img_path):
    # Apriamo l'immagine
    orig_img = Image.open(img_path).convert("RGB")
    w_orig, h_orig = orig_img.size

    # TRUCCO INSIEME: ingrandiamo l'immagine temporaneamente del 200% (2x) prima di disegnare i testi,
    # così qualsiasi testo disegnato apparirà enorme e perfettamente definito quando viene riscalato!
    scale_factor = 2
    base_img = orig_img.resize(
        (w_orig * scale_factor, h_orig * scale_factor), Image.Resampling.LANCZOS
    )
    draw = ImageDraw.Draw(base_img)

    width, height = base_img.size

    # Usiamo il font di default moltiplicato per la scala ingrandita (apparirà grande)
    # Oppure proviamo a caricare un font con grandezza raddoppiata
    try:
      font_grande = ImageFont.truetype("arial.ttf", 110)
      font_medio = ImageFont.truetype("arial.ttf", 85)
    except:
      font_grande = ImageFont.load_default()
      font_medio = font_grande

    # Coordinate moltiplicate per il fattore di scala (scale_factor = 2)
    # 1. KM Fine (in alto)
    draw.text(
        (int(width * 0.60), int(height * 0.002)),
        f"KM {km_fine}",
        fill="red",
        font=font_grande,
    )

    # 2. KM Testata / Intermedio (es. 36+110)
    draw.text(
        (int(width * 0.59), int(height * 0.525)),
        f"KM {km_testata}",
        fill="red",
        font=font_medio,
    )

    # 3. KM Preavviso (al centro)
    draw.text(
        (int(width * 0.59), int(height * 0.645)),
        f"KM {km_preavviso}",
        fill="red",
        font=font_medio,
    )

    # 4. KM Inizio (in basso)
    draw.text(
        (int(width * 0.59), int(height * 0.905)),
        f"KM {km_inizio}",
        fill="red",
        font=font_medio,
    )

    # Riportiamo l'immagine alla dimensione originale pulita e definita
    final_img = base_img.resize((w_orig, h_orig), Image.Resampling.LANCZOS)

    # Salviamo in memoria
    img_io = io.BytesIO()
    final_img.save(img_io, "JPEG", quality=95)
    img_io.seek(0)

    # Mostriamo l'anteprima sul sito
    st.image(
        final_img,
        caption="Tavola 39 - Testi ingranditi con fattore di scala",
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
    st.warning(
        f"⚠️ L'immagine '{img_path}' non è stata trovata su GitHub. Assicurati"
        " di averla caricata con questo nome esatto."
    )
