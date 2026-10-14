import io
import os
import urllib.request
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from openpyxl.drawing.image import Image as OpenpyxlImage
import openpyxl

# ==================================================
# CONFIGURAZIONE STREAMLIT
# ==================================================
st.set_page_config(
    page_title="Gestione Cantieri Stradali - Abaco & Tavole",
    layout="wide",
)

st.title("🚧 Gestione Cantieri Stradali e Segnaletica")

st.write(
    "Inserisci le progressive chilometriche e l'ubicazione: il programma "
    "genererà l'Abaco Excel con le immagini dei segnali e la Tavola JPG personalizzata."
)


# ==================================================
# CARICAMENTO FONT COMPATIBILE CON STREAMLIT CLOUD
# ==================================================
@st.cache_resource
def load_custom_font(size):
  cartella_app = os.path.dirname(os.path.abspath(__file__))

  possibili_font = [
      os.path.join(cartella_app, "DejaVuSans-Bold.ttf"),
      "DejaVuSans-Bold.ttf",
      "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
      "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
      "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
      "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
      "C:/Windows/Fonts/arialbd.ttf",
      "/Library/Fonts/Arial Bold.ttf",
  ]

  for percorso in possibili_font:
    if os.path.isfile(percorso):
      try:
        return ImageFont.truetype(percorso, size)
      except (OSError, ValueError):
        continue

  url = (
      "https://raw.githubusercontent.com/"
      "dejavu-fonts/dejavu-fonts/master/"
      "src/DejaVuSans-Bold.ttf"
  )
  percorso_locale = "/tmp/DejaVuSans-Bold.ttf"

  try:
    urllib.request.urlretrieve(url, percorso_locale)
    return ImageFont.truetype(percorso_locale, size)
  except Exception:
    pass

  try:
    return ImageFont.load_default(size=size)
  except TypeError:
    return ImageFont.load_default()


# ==================================================
# SELEZIONE DELLA TAVOLA
# ==================================================
tavola_selezionata = st.selectbox(
    "Seleziona la Tavola di Cantiere",
    ["Tavola 39 - Cantiere mobile carreggiata a due corsie"],
)


if "39" in tavola_selezionata:

  st.subheader("📍 Dati e Progressive Chilometriche - Tavola 39")

  ubicazione_testo = st.text_area(
      "Ubicazione / Direzione (può essere divisa su più righe)",
      value="DIREZIONE SUD\nVIADOTTO TORRE\nGUACETO 1",
  )

  col1, col2 = st.columns(2)

  with col1:
    km_inizio = st.text_input(
        "KM Inizio Cantiere (Basso)",
        value="35+560",
    )

    km_preavviso = st.text_input(
        "KM Segnale Mobile Preavviso (Centro)",
        value="36+060",
    )

  with col2:
    km_fine = st.text_input(
        "KM Fine / Rientro (Alto)",
        value="36+600",
    )

    km_testata = st.text_input(
        "KM Inizio Testata / Intermedio",
        value="36+110",
    )

    distanza_coni = st.text_input(
        "Interasse Coni / Dettagli",
        value="12 m",
    )

  # ==================================================
  # GENERAZIONE FILE EXCEL (ABACO CON IMMAGINI NELLE CELLE)
  # ==================================================
  if st.button("Genera Abaco Segnaletica (Excel)"):

    dati_abaco = {
        "SEGNALE": [
            "Preavviso Cantiere Mobile",
            "Lavori in corso",
            "Fine Cantiere / Via Libera",
            "Segnale Mobile di Preavviso su Veicolo",
            "Segnale Mobile di Protezione su Veicolo",
        ],
        "FIGURA": [
            "FIG. II 399",
            "FIG. II 383",
            "FIG. II 070",
            "FIG. II 400",
            "FIG. II 401",
        ],
        "DESCRIZIONE": [
            "PREAVVISO DI CANTIERE MOBILE",
            "LAVORI IN CORSO",
            "FINE CANTIERE",
            "SEGNALE MOBILE DI PREAVVISO CON FRECCIA LUMINOSA",
            "VEICOLO DI PROTEZIONE CON PANNELLO A MESSAGGI VARIABILI",
        ],
        "PROGRESSIVA CHILOMETRICA": [
            km_preavviso,
            km_inizio,
            km_fine,
            km_preavviso,
            km_inizio,
        ],
        "SEGNALETICA": [
            "SEGNALI TEMPORANEI",
            "SEGNALI TEMPORANEI",
            "SEGNALI DI PRESCRIZIONE",
            "SEGNALI COMPLEMENTARI / TEMPORANEI",
            "SEGNALI COMPLEMENTARI / TEMPORANEI",
        ],
    }

    df_risultato = pd.DataFrame(dati_abaco)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_risultato.to_excel(writer, index=False, sheet_name="Abaco Tavola 39")

    output.seek(0)
    
    wb = openpyxl.load_workbook(output)
    ws = wb["Abaco Tavola 39"]

    cartella_app = os.path.dirname(os.path.abspath(__file__))
    dir_immagini = os.path.join(cartella_app, "immagini_segnali")

    # Mappa dei codici alle chiavi dei file immagine
    mappa_nomi = {
        "FIG. II 399": ["fig_399.png", "fig_399.jpg", "399.png", "399.jpg"],
        "FIG. II 383": ["fig_383.png", "fig_383.jpg", "383.png", "383.jpg"],
        "FIG. II 070": ["fig_070.png", "fig_070.jpg", "070.png", "070.jpg", "70.png"],
        "FIG. II 400": ["fig_400.png", "fig_400.jpg", "400.png", "400.jpg"],
        "FIG. II 401": ["fig_401.png", "fig_401.jpg", "401.png", "401.jpg"],
    }

    ws.column_dimensions['B'].width = 18

    for riga_idx in range(2, len(df_risultato) + 2):
      ws.row_dimensions[riga_idx].height = 50
      figura_val = ws.cell(row=riga_idx, column=2).value

      if figura_val in mappa_nomi:
        # Cerca quale dei file possibili esiste nella cartella
        file_trovato = None
        for nome_file in mappa_nomi[figura_val]:
          percorso_tentativo = os.path.join(dir_immagini, nome_file)
          if os.path.exists(percorso_tentativo):
            file_trovato = percorso_tentativo
            break
        
        if file_trovato:
          try:
            # Salviamo una versione temporanea standardizzata per openpyxl
            img_pil = Image.open(file_trovato).convert("RGB")
            temp_img_path = os.path.join("/tmp", f"temp_sig_{riga_idx}.jpg")
            img_pil.save(temp_img_path, "JPEG")

            img_obj = OpenpyxlImage(temp_img_path)
            img_obj.width = 40
            img_obj.height = 40
            
            ws.add_image(img_obj, f"B{riga_idx}")
          except Exception as e:
            print(f"Errore caricamento immagine riga {riga_idx}: {e}")

    final_output = io.BytesIO()
    wb.save(final_output)
    excel_data = final_output.getvalue()

    st.success("Abaco Excel generato con successo con tutte le immagini dei segnali!")

    st.download_button(
        label="📥 Scarica Excel Abaco Segnaletica con Immagini",
        data=excel_data,
        file_name="Abaco_Segnaletica_Tavola_39.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

  # ==================================================
  # GENERAZIONE TAVOLA JPG
  # ==================================================
  st.divider()

  st.subheader("🖼️ Anteprima Tavola 39 Personalizzata")

  cartella_app = os.path.dirname(os.path.abspath(__file__))
  img_path = os.path.join(cartella_app, "TAVOLA 39.jpg")

  if os.path.exists(img_path):

    try:
      with Image.open(img_path) as immagine_originale:
        base_img = immagine_originale.convert("RGB")

      width, height = base_img.size
      draw = ImageDraw.Draw(base_img)

      DIMENSIONE_FONT_KM = 60
      DIMENSIONE_FONT_UBICAZIONE = 69

      font_km = load_custom_font(DIMENSIONE_FONT_KM)
      font_ubicazione = load_custom_font(DIMENSIONE_FONT_UBICAZIONE)

      BORDO = 2

      etichette = [
          (
              int(width * 0.45),
              int(height * 0.015),
              f"KM {km_fine}",
          ),
          (
              int(width * 0.62),
              int(height * 0.525),
              f"KM {km_testata}",
          ),
          (
              int(width * 0.62),
              int(height * 0.645),
              f"KM {km_preavviso}",
          ),
          (
              int(width * 0.62),
              int(height * 0.905),
              f"KM {km_inizio}",
          ),
      ]

      for x, y, testo in etichette:
        draw.text(
            (x, y),
            testo,
            fill=(255, 0, 0),
            font=font_km,
            stroke_width=BORDO,
            stroke_fill=(255, 255, 255),
        )

      x_ubicazione = int(width * 0.65)
      y_ubicazione = int(height * 0.225)

      draw.multiline_text(
          (x_ubicazione, y_ubicazione),
          ubicazione_testo,
          fill=(255, 0, 0),
          font=font_ubicazione,
          align="center",
          spacing=6,
          stroke_width=BORDO,
          stroke_fill=(255, 255, 255),
      )

      img_io = io.BytesIO()
      base_img.save(
          img_io,
          format="JPEG",
          quality=100,
          subsampling=0,
      )
      jpg_data = img_io.getvalue()

      st.image(
          base_img,
          caption="Tavola 39 - Configurazione finale completata",
          use_container_width=True,
      )

      st.download_button(
          label="📥 Scarica Tavola in formato JPG",
          data=jpg_data,
          file_name="Tavola_39_Compilata.jpg",
          mime="image/jpeg",
      )

      st.success("Tavola generata correttamente con ubicazione ottimizzata.")

    except Exception as errore:
      st.error("Errore durante la generazione della tavola.")
      st.exception(errore)

  else:
    st.error(
        "Immagine 'TAVOLA 39.jpg' non trovata. "
        "Carica il file nella stessa cartella di app.py."
    )
