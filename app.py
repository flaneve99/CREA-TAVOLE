import io
import os
import urllib.request
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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
    "genererà l'Abaco Excel strutturato (con dettagli vuoti e immagini bloccate tramite TwoCellAnchor) e la Tavola JPG personalizzata."
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
  # GENERAZIONE FILE EXCEL (ABACO STRUTTURATO)
  # ==================================================
  if st.button("Genera Abaco Segnaletica (Excel)"):

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Abaco Tavola 39"

    # 1. Riga di intestazione principale unificata
    ws.merge_cells("A1:E1")
    titolo_principale = ws["A1"]
    titolo_principale.value = "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE"
    
    fill_verde_chiaro = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    fill_verde_scuro = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    
    font_titolo = Font(name="Arial", size=11, bold=True, color="000000")
    font_header = Font(name="Arial", size=10, bold=True, color="000000")
    font_testo = Font(name="Arial", size=10, color="000000")

    bordo_sottile = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    titolo_principale.fill = fill_verde_chiaro
    titolo_principale.font = font_titolo
    titolo_principale.alignment = Alignment(horizontal="center", vertical="center")

    for col in range(1, 6):
      ws.cell(row=1, column=col).border = bordo_sottile

    # 2. Riga delle colonne (DETTAGLI vuoto)
    headers = ["SEGNALE", "FIGURA", "DESCRIZIONE", "DETTAGLI", "SEGNALETICA"]
    for col_idx, header in enumerate(headers, 1):
      cell = ws.cell(row=2, column=col_idx)
      cell.value = header
      cell.fill = fill_verde_scuro
      cell.font = font_header
      cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
      cell.border = bordo_sottile

    # 3. Dati delle righe
    dati_righe = [
        ("fig_399.png", "FIG. II 399", "PREAVVISO DI CANTIERE MOBILE", "", "SEGNALI TEMPORANEI"),
        ("fig_383.png", "FIG. II 383", "LAVORI IN CORSO", "", "SEGNALI TEMPORANEI"),
        ("fig_070.png", "FIG. II 070", "FINE CANTIERE", "", "SEGNALI DI PRESCRIZIONE"),
        ("fig_400.png", "FIG. II 400", "SEGNALE MOBILE DI PREAVVISO CON FRECCIA LUMINOSA", "", "SEGNALI COMPLEMENTARI"),
        ("fig_401.png", "FIG. II 401", "VEICOLO DI PROTEZIONE CON PANNELLO", "", "SEGNALI COMPLEMENTARI"),
    ]

    cartella_app = os.path.dirname(os.path.abspath(__file__))
    dir_immagini = os.path.join(cartella_app, "immagini_segnali")

    # Dimensioni colonne e righe
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 25

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 25

    immagini_inserite = 0

    for i, row_data in enumerate(dati_righe, start=3):
      img_nome, figura, descrizione, dettagli, segnaletica = row_data
      
      ws.row_dimensions[i].height = 55

      c_fig = ws.cell(row=i, column=2, value=figura)
      c_desc = ws.cell(row=i, column=3, value=descrizione)
      c_det = ws.cell(row=i, column=4, value=dettagli)
      c_seg = ws.cell(row=i, column=5, value=segnaletica)

      for cell in [c_fig, c_desc, c_det, c_seg]:
        cell.font = font_testo
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = bordo_sottile

      c_segnale = ws.cell(row=i, column=1)
      c_segnale.border = bordo_sottile

      nomi_da_cercare = [img_nome, img_nome.replace(".png", ".jpg"), img_nome.replace(".jpg", ".png")]
      percorso_trovato = None
      
      for nome in nomi_da_cercare:
        tentativo = os.path.join(dir_immagini, nome)
        if os.path.exists(tentativo):
          percorso_trovato = tentativo
          break

      if percorso_trovato:
        try:
          img_pil = Image.open(percorso_trovato).convert("RGB")
          temp_path = os.path.join("/tmp", f"abaco_img_{i}.jpg")
          img_pil.save(temp_path, "JPEG")

          img_obj = OpenpyxlImage(temp_path)
          
          # ==========================================================
          # ANCORAGGIO A DUE CELLE (TwoCellAnchor) PER BLOCCARE L'IMMAGINE
          # ==========================================================
          marker_from = AnchorMarker(col=0, colOff=15000, row=i - 1, rowOff=10000)
          marker_to = AnchorMarker(col=1, colOff=-15000, row=i, rowOff=-10000)
          
          img_obj.anchor = TwoCellAnchor(
              editAs="twoCell",
              _from=marker_from,
              to=marker_to
          )
          
          ws.add_image(img_obj)
          immagini_inserite += 1
        except Exception as e:
          st.warning(f"Errore nel caricamento dell'immagine per {figura}: {e}")
      else:
        c_segnale.value = "[Img non trovata]"
        c_segnale.font = Font(name="Arial", size=8, color="FF0000")
        c_segnale.alignment = Alignment(horizontal="center", vertical="center")

    output = io.BytesIO()
    wb.save(output)
    excel_data = output.getvalue()

    st.success(f"Abaco Excel generato con successo! (Immagini bloccate con TwoCellAnchor: {immagini_inserite}/5)")

    st.download_button(
        label="📥 Scarica Excel Abaco Segnaletica Formattato",
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
