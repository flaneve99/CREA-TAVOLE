import io
import os
import urllib.request
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import streamlit as st

# ==================================================
# CONFIGURAZIONE STREAMLIT E PERCORSI
# ==================================================
st.set_page_config(
    page_title="Gestione Cantieri Stradali - Abaco & Tavole",
    layout="wide",
)

cartella_app = os.path.dirname(os.path.abspath(__file__))
dir_immagini = os.path.join(cartella_app, "immagini_segnali")

st.title("🚧 Gestione Cantieri Stradali e Segnaletica")

st.write(
    "Seleziona la Tavola di cantiere, inserisci i dati e le progressive: "
    "il programma genererà l'Abaco Excel strutturato e la Tavola JPG"
    " personalizzata."
)


# ==================================================
# CARICAMENTO FONT COMPATIBILE CON STREAMLIT CLOUD
# ==================================================
@st.cache_resource
def load_custom_font(size):
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
    [
        "Tavola 39 - Cantiere mobile carreggiata a due corsie",
        "Tavola 16 - Chiusura della corsia di marcia su carreggiata a due corsie",
        (
            "Tavola 66 - Lavori su carreggiata con senso unico alternato"
            " semaforico"
        ),
    ],
)


# ====================================================================
# TAVOLA 39
# ====================================================================
if "39" in tavola_selezionata:

  st.subheader("📍 Dati e Progressive Chilometriche - Tavola 39")

  ubicazione_testo = st.text_area(
      "Ubicazione / Direzione (può essere divisa su più righe)",
      value="DIREZIONE SUD\nVIADOTTO TORRE\nGUACETO 1",
  )

  col1, col2 = st.columns(2)

  with col1:
    km_inizio = st.text_input("KM Inizio Cantiere (Basso)", value="35+560")
    km_preavviso = st.text_input(
        "KM Segnale Mobile Preavviso (Centro)", value="36+060"
    )

  with col2:
    km_fine = st.text_input("KM Fine / Rientro (Alto)", value="36+600")
    km_testata = st.text_input("KM Inizio Testata / Intermedio", value="36+110")
    distanza_coni = st.text_input("Interasse Coni / Dettagli", value="12 m")

  if st.button("Genera Abaco Segnaletica (Excel) - Tavola 39"):

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Abaco Tavola 39"

    ws.merge_cells("A1:E1")
    titolo_principale = ws["A1"]
    titolo_principale.value = "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE"

    fill_verde_chiaro = PatternFill(
        start_color="92D050", end_color="92D050", fill_type="solid"
    )
    fill_verde_scuro = PatternFill(
        start_color="70AD47", end_color="70AD47", fill_type="solid"
    )

    font_titolo = Font(name="Arial", size=11, bold=True, color="000000")
    font_header = Font(name="Arial", size=10, bold=True, color="000000")
    font_testo = Font(name="Arial", size=10, color="000000")

    bordo_sottile = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    titolo_principale.fill = fill_verde_chiaro
    titolo_principale.font = font_titolo
    titolo_principale.alignment = Alignment(
        horizontal="center", vertical="center"
    )

    for col in range(1, 6):
      ws.cell(row=1, column=col).border = bordo_sottile

    headers = ["SEGNALE", "FIGURA", "DESCRIZIONE", "DETTAGLI", "SEGNALETICA"]
    for col_idx, header in enumerate(headers, 1):
      cell = ws.cell(row=2, column=col_idx)
      cell.value = header
      cell.fill = fill_verde_scuro
      cell.font = font_header
      cell.alignment = Alignment(
          horizontal="center", vertical="center", wrap_text=True
      )
      cell.border = bordo_sottile

    dati_righe = [
        (
            "fig_399.jpg",
            "FIG. II 399",
            "PREAVVISO DI CANTIERE MOBILE",
            "",
            "SEGNALI TEMPORANEI",
        ),
        ("fig_383.jpg", "FIG. II 383", "LAVORI IN CORSO", "", "SEGNALI TEMPORANEI"),
        (
            "fig_070.jpg",
            "FIG. II 070",
            "FINE CANTIERE",
            "",
            "SEGNALI DI PRESCRIZIONE",
        ),
        (
            "fig_400.jpg",
            "FIG. II 400",
            "SEGNALE MOBILE DI PREAVVISO CON FRECCIA LUMINOSA",
            "",
            "SEGNALI TEMPORANEI",
        ),
        (
            "fig_401.jpg",
            "FIG. II 401",
            "VEICOLO DI PROTEZIONE CON PANNELLO",
            "",
            "SEGNALI TEMPORANEI",
        ),
    ]

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 25

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
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = bordo_sottile

      c_segnale = ws.cell(row=i, column=1)
      c_segnale.border = bordo_sottile

      nomi_da_cercare = [
          img_nome,
          img_nome.replace(".jpg", ".png"),
          img_nome.replace(".png", ".jpg"),
      ]
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
          marker_from = AnchorMarker(
              col=0, colOff=150000, row=i - 1, rowOff=40000
          )
          marker_to = AnchorMarker(col=1, colOff=-150000, row=i, rowOff=-40000)
          img_obj.anchor = TwoCellAnchor(
              editAs="twoCell", _from=marker_from, to=marker_to
          )

          ws.add_image(img_obj)
          immagini_inserite += 1
        except Exception as e:
          st.warning(f"Errore caricamento immagine per {figura}: {e}")
      else:
        c_segnale.value = "[Img non trovata]"
        c_segnale.font = Font(name="Arial", size=8, color="FF0000")
        c_segnale.alignment = Alignment(horizontal="center", vertical="center")

    output = io.BytesIO()
    wb.save(output)
    excel_data = output.getvalue()

    st.success(
        "Abaco Excel generato con successo! (Immagini:"
        f" {immagini_inserite}/5)"
    )
    st.download_button(
        label="📥 Scarica Excel Abaco Segnaletica Formattato",
        data=excel_data,
        file_name="Abaco_Segnaletica_Tavola_39.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  st.divider()
  st.subheader("🖼️ Anteprima Tavola 39 Personalizzata")

  img_path = os.path.join(cartella_app, "TAVOLA 39.jpg")
  if os.path.exists(img_path):
    try:
      with Image.open(img_path) as immagine_originale:
        base_img = immagine_originale.convert("RGB")

      width, height = base_img.size
      draw = ImageDraw.Draw(base_img)
      font_km = load_custom_font(60)
      font_ubicazione = load_custom_font(69)
      BORDO = 2

      etichette = [
          (int(width * 0.45), int(height * 0.015), f"KM {km_fine}"),
          (int(width * 0.62), int(height * 0.525), f"KM {km_testata}"),
          (int(width * 0.62), int(height * 0.645), f"KM {km_preavviso}"),
          (int(width * 0.62), int(height * 0.905), f"KM {km_inizio}"),
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

      draw.multiline_text(
          (int(width * 0.65), int(height * 0.225)),
          ubicazione_testo,
          fill=(255, 0, 0),
          font=font_ubicazione,
          align="center",
          spacing=6,
          stroke_width=BORDO,
          stroke_fill=(255, 255, 255),
      )

      img_io = io.BytesIO()
      base_img.save(img_io, format="JPEG", quality=100, subsampling=0)
      jpg_data = img_io.getvalue()

      st.image(
          base_img,
          caption="Tavola 39 - Configurazione completata",
          use_container_width=True,
      )
      st.download_button(
          label="📥 Scarica Tavola in formato JPG",
          data=jpg_data,
          file_name="Tavola_39_Compilata.jpg",
          mime="image/jpeg",
      )
    except Exception as errore:
      st.error("Errore generazione Tavola 39.")
      st.exception(errore)
  else:
    st.error("Immagine 'TAVOLA 39.jpg' non trovata nella cartella.")


# ====================================================================
# TAVOLA 16
# ====================================================================
elif "16" in tavola_selezionata:

  st.subheader("📍 Dati e Progressive Chilometriche - Tavola 16")

  ubicazione_testo_16 = st.text_area(
      "Ubicazione / Direzione (Tavola 16)",
      value="DIREZIONE SUD\nVIADOTTO TORRE\nGUACETO 1",
  )

  col1, col2 = st.columns(2)

  with col1:
    km_fine_16 = st.text_input("KM Fine / Rientro (In alto)", value="36+600")
    km_preavviso_700 = st.text_input(
        "KM Preavviso 700m (In basso)", value="34+860"
    )
    km_limite_110 = st.text_input("KM Limite 110 km/h", value="35+100")
    km_limite_90 = st.text_input("KM Limite 90 km/h", value="35+220")
    km_preavviso_250 = st.text_input(
        "KM Preavviso Chiusura Corsia 250m", value="35+310"
    )

  with col2:
    km_limite_60 = st.text_input("KM Limite 60 km/h", value="35+430")
    km_preavviso_100 = st.text_input(
        "KM Preavviso Chiusura Corsia 100m", value="35+490"
    )
    km_inizio_testata = st.text_input(
        "KM Inizio Testata (Freccia blu)", value="35+560"
    )

  if st.button("Genera Abaco Segnaletica (Excel) - Tavola 16"):

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Abaco Tavola 16"

    ws.merge_cells("A1:E1")
    titolo_principale = ws["A1"]
    titolo_principale.value = (
        "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 16"
    )

    fill_verde_chiaro = PatternFill(
        start_color="92D050", end_color="92D050", fill_type="solid"
    )
    fill_verde_scuro = PatternFill(
        start_color="70AD47", end_color="70AD47", fill_type="solid"
    )

    font_titolo = Font(name="Arial", size=11, bold=True, color="000000")
    font_header = Font(name="Arial", size=10, bold=True, color="000000")
    font_testo = Font(name="Arial", size=10, color="000000")

    bordo_sottile = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    titolo_principale.fill = fill_verde_chiaro
    titolo_principale.font = font_titolo
    titolo_principale.alignment = Alignment(
        horizontal="center", vertical="center"
    )

    for col in range(1, 6):
      ws.cell(row=1, column=col).border = bordo_sottile

    headers = ["SEGNALE", "FIGURA", "DESCRIZIONE", "DETTAGLI", "SEGNALETICA"]
    for col_idx, header in enumerate(headers, 1):
      cell = ws.cell(row=2, column=col_idx)
      cell.value = header
      cell.fill = fill_verde_scuro
      cell.font = font_header
      cell.alignment = Alignment(
          horizontal="center", vertical="center", wrap_text=True
      )
      cell.border = bordo_sottile

    dati_righe_16 = [
        ("fig_383.jpg", "FIG. II 383", "LAVORI IN CORSO", "", "SEGNALI TEMPORANEI"),
        (
            "fig_082a.jpg",
            "FIG. II 082/a",
            "PASSAGGIO OBBLIGATORIO A SINISTRA",
            "",
            "SEGNALI DI OBBLIGO",
        ),
        (
            "fig_048.jpg",
            "FIG. II 048",
            "DIVIETO DI SORPASSO",
            "",
            "SEGNALI DI DIVIETO",
        ),
        (
            "fig_411a.jpg",
            "FIG. II 411/a",
            "SEGNALE DI CORSIA CHIUSA (CHIUSURA CORSIA DI DESTRA)",
            "",
            "SEGNALI TEMPORANEI",
        ),
        (
            "fig050_lim110.jpg",
            "FIG. II 050",
            "LIMITE DI VELOCITÀ 110 KM/H",
            "",
            "SEGNALI DI PRESCRIZIONE",
        ),
        (
            "fig050_lim90.jpg",
            "FIG. II 050",
            "LIMITE DI VELOCITÀ 90 KM/H",
            "",
            "SEGNALI DI PRESCRIZIONE",
        ),
        (
            "fig_050_lim60.jpg",
            "FIG. II 050",
            "LIMITE DI VELOCITÀ 60 KM/H",
            "",
            "SEGNALI DI PRESCRIZIONE",
        ),
        ("fig_070.jpg", "FIG. II 070", "VIA LIBERA", "", "SEGNALI DI DIVIETO"),
    ]

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 25

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 25

    immagini_inserite = 0

    for i, row_data in enumerate(dati_righe_16, start=3):
      img_nome, figura, descrizione, dettagli, segnaletica = row_data
      ws.row_dimensions[i].height = 55

      c_fig = ws.cell(row=i, column=2, value=figura)
      c_desc = ws.cell(row=i, column=3, value=descrizione)
      c_det = ws.cell(row=i, column=4, value=dettagli)
      c_seg = ws.cell(row=i, column=5, value=segnaletica)

      for cell in [c_fig, c_desc, c_det, c_seg]:
        cell.font = font_testo
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = bordo_sottile

      c_segnale = ws.cell(row=i, column=1)
      c_segnale.border = bordo_sottile

      nomi_da_cercare = [
          img_nome,
          img_nome.replace(".jpg", ".png"),
          img_nome.replace(".png", ".jpg"),
      ]
      percorso_trovato = None
      for nome in nomi_da_cercare:
        tentativo = os.path.join(dir_immagini, nome)
        if os.path.exists(tentativo):
          percorso_trovato = tentativo
          break

      if percorso_trovato:
        try:
          img_pil = Image.open(percorso_trovato).convert("RGB")
          temp_path = os.path.join("/tmp", f"abaco_16_img_{i}.jpg")
          img_pil.save(temp_path, "JPEG")

          img_obj = OpenpyxlImage(temp_path)
          marker_from = AnchorMarker(
              col=0, colOff=150000, row=i - 1, rowOff=40000
          )
          marker_to = AnchorMarker(col=1, colOff=-150000, row=i, rowOff=-40000)
          img_obj.anchor = TwoCellAnchor(
              editAs="twoCell", _from=marker_from, to=marker_to
          )

          ws.add_image(img_obj)
          immagini_inserite += 1
        except Exception as e:
          st.warning(f"Errore caricamento immagine per {figura}: {e}")
      else:
        c_segnale.value = "[Img non trovata]"
        c_segnale.font = Font(name="Arial", size=8, color="FF0000")
        c_segnale.alignment = Alignment(horizontal="center", vertical="center")

    output = io.BytesIO()
    wb.save(output)
    excel_data = output.getvalue()

    st.success(
        "Abaco Excel Tavola 16 generato! (Immagini inserite:"
        f" {immagini_inserite}/{len(dati_righe_16)})"
    )
    st.download_button(
        label="📥 Scarica Excel Abaco Tavola 16",
        data=excel_data,
        file_name="Abaco_Segnaletica_Tavola_16.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  st.divider()
  st.subheader("🖼️ Anteprima Tavola 16 Personalizzata")

  img_path_16 = os.path.join(cartella_app, "TAVOLA 16.jpg")
  if os.path.exists(img_path_16):
    try:
      with Image.open(img_path_16) as immagine_originale:
        base_img_16 = immagine_originale.convert("RGB")

      width_16, height_16 = base_img_16.size
      draw_16 = ImageDraw.Draw(base_img_16)

      font_km = load_custom_font(25)
      font_ub_16 = load_custom_font(30)
      BORDO = 2

      etichette_16 = [
          (int(width_16 * 0.34), int(height_16 * 0.07), f"KM {km_fine_16}"),
          (int(width_16 * 0.49), int(height_16 * 0.925), f"KM {km_preavviso_700}"),
          (int(width_16 * 0.49), int(height_16 * 0.865), f"KM {km_limite_110}"),
          (int(width_16 * 0.49), int(height_16 * 0.800), f"KM {km_limite_90}"),
          (int(width_16 * 0.49), int(height_16 * 0.735), f"KM {km_preavviso_250}"),
          (int(width_16 * 0.49), int(height_16 * 0.675), f"KM {km_limite_60}"),
          (int(width_16 * 0.49), int(height_16 * 0.610), f"KM {km_preavviso_100}"),
          (int(width_16 * 0.49), int(height_16 * 0.505), f"KM {km_inizio_testata}"),
      ]

      for x, y, testo in etichette_16:
        draw_16.text(
            (x, y),
            testo,
            fill=(255, 0, 0),
            font=font_km,
            stroke_width=BORDO,
            stroke_fill=(255, 255, 255),
        )

      x_ub = int(width_16 * 0.67)
      y_ub = int(height_16 * 0.27)
      draw_16.multiline_text(
          (x_ub, y_ub),
          ubicazione_testo_16,
          fill=(255, 0, 0),
          font=font_ub_16,
          align="center",
          spacing=6,
          stroke_width=BORDO,
          stroke_fill=(255, 255, 255),
      )

      img_io_16 = io.BytesIO()
      base_img_16.save(img_io_16, format="JPEG", quality=100, subsampling=0)
      jpg_data_16 = img_io_16.getvalue()

      st.image(
          base_img_16,
          caption="Tavola 16 - Configurazione completata",
          use_container_width=True,
      )
      st.download_button(
          label="📥 Scarica Tavola 16 in formato JPG",
          data=jpg_data_16,
          file_name="Tavola_16_Compilata.jpg",
          mime="image/jpeg",
      )
    except Exception as errore:
      st.error("Errore generazione Tavola 16.")
      st.exception(errore)
  else:
    st.error("Immagine 'TAVOLA 16.jpg' non trovata nella cartella.")


# ====================================================================
# TAVOLA 66
# ====================================================================
elif "66" in tavola_selezionata:

  st.subheader("📍 Dati e Progressive Chilometriche - Tavola 66")

  ubicazione_testo_66 = st.text_area(
      "Ubicazione / Direzione (Tavola 66)",
      value="DIREZIONE MATERA LATO DX\nS.P. 3",
  )

  st.write("### Lato Destro dell'Immagine")
  col1, col2 = st.columns(2)
  with col1:
    km_66_dx_preavviso = st.text_input(
        "KM Preavviso Lavori a 200m (Dx)", value="5+065"
    )
    km_66_dx_limite60 = st.text_input("KM Limite 60 (Dx)", value="5+110")
    km_66_dx_strettoia = st.text_input(
        "KM Strettoia Asimmetrica a Destra (Dx)", value="5+155"
    )
  with col2:
    km_66_dx_semaforo = st.text_input("KM Semaforo (Dx)", value="5+200")
    km_66_dx_chiusura = st.text_input("KM Chiusura Zona Lavori (Dx)", value="5+300")
    km_66_dx_fine = st.text_input("KM Fine Zona Lavori (Dx)", value="5+600")

  st.write("### Lato Sinistro dell'Immagine")
  col3, col4 = st.columns(2)
  with col3:
    km_66_sx_zona = st.text_input("KM Zona Lavori (Sx)", value="5+600")
    km_66_sx_semaforo = st.text_input("KM Semaforo (Sx)", value="5+700")
    km_66_sx_strettoia = st.text_input(
        "KM Strettoia / Altro segnale (Sx)", value="5+745"
    )
  with col4:
    km_66_sx_limite60 = st.text_input("KM Limite 60 / Divieto (Sx)", value="5+790")
    km_66_sx_preavviso = st.text_input(
        "KM Preavviso Lavori (Sx)", value="5+835"
    )

  if st.button("Genera Abaco Segnaletica (Excel) - Tavola 66"):

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Abaco Tavola 66"

    ws.merge_cells("A1:E1")
    titolo_principale = ws["A1"]
    titolo_principale.value = (
        "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 66"
    )

    fill_verde_chiaro = PatternFill(
        start_color="92D050", end_color="92D050", fill_type="solid"
    )
    fill_verde_scuro = PatternFill(
        start_color="70AD47", end_color="70AD47", fill_type="solid"
    )

    font_titolo = Font(name="Arial", size=11, bold=True, color="000000")
    font_header = Font(name="Arial", size=10, bold=True, color="000000")
    font_testo = Font(name="Arial", size=10, color="000000")

    bordo_sottile = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    titolo_principale.fill = fill_verde_chiaro
    titolo_principale.font = font_titolo
    titolo_principale.alignment = Alignment(
        horizontal="center", vertical="center"
    )

    for col in range(1, 6):
      ws.cell(row=1, column=col).border = bordo_sottile

    headers = ["SEGNALE", "FIGURA", "DESCRIZIONE", "DETTAGLI", "SEGNALETICA"]
    for col_idx, header in enumerate(headers, 1):
      cell = ws.cell(row=2, column=col_idx)
      cell.value = header
      cell.fill = fill_verde_scuro
      cell.font = font_header
      cell.alignment = Alignment(
          horizontal="center", vertical="center", wrap_text=True
      )
      cell.border = bordo_sottile

    dati_righe_66 = [
        (
            "fig_050_lim60.jpg",
            "FIG. II 050",
            "LIMITE DI VELOCITÀ 60 KM/H",
            "",
            "SEGNALI DI PRESCRIZIONE",
        ),
        (
            "fig_048.jpg",
            "FIG. II 048",
            "DIVIETO DI SORPASSO",
            "",
            "SEGNALI DI DIVIETO",
        ),
        (
            "fig_386.jpg",
            "FIG. II 386",
            "STRETTOIA ASIMMETRICA A DESTRA",
            "",
            "SEGNALI TEMPORANEI",
        ),
        (
            "fig_404.jpg",
            "FIG. II 404",
            "SEMAFORO",
            "",
            "SEGNALI TEMPORANEI",
        ),
        (
            "fig_383.jpg",
            "FIG. II 383",
            "ZONA LAVORI / SENSO UNICO ALTERNATO",
            "",
            "SEGNALI TEMPORANEI",
        ),
        (
            "fig_070.jpg",
            "FIG. II 070",
            "VIA LIBERA / FINE CANTIERE",
            "",
            "SEGNALI DI PRESCRIZIONE",
        ),
    ]

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 25

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 25

    immagini_inserite = 0

    for i, row_data in enumerate(dati_righe_66, start=3):
      img_nome, figura, descrizione, dettagli, segnaletica = row_data
      ws.row_dimensions[i].height = 55

      c_fig = ws.cell(row=i, column=2, value=figura)
      c_desc = ws.cell(row=i, column=3, value=descrizione)
      c_det = ws.cell(row=i, column=4, value=dettagli)
      c_seg = ws.cell(row=i, column=5, value=segnaletica)

      for cell in [c_fig, c_desc, c_det, c_seg]:
        cell.font = font_testo
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = bordo_sottile

      c_segnale = ws.cell(row=i, column=1)
      c_segnale.border = bordo_sottile

      nomi_da_cercare = [
          img_nome,
          img_nome.replace(".jpg", ".png"),
          img_nome.replace(".png", ".jpg"),
      ]
      percorso_trovato = None
      for nome in nomi_da_cercare:
        tentativo = os.path.join(dir_immagini, nome)
        if os.path.exists(tentativo):
          percorso_trovato = tentativo
          break

      if percorso_trovato:
        try:
          img_pil = Image.open(percorso_trovato).convert("RGB")
          temp_path = os.path.join("/tmp", f"abaco_66_img_{i}.jpg")
          img_pil.save(temp_path, "JPEG")

          img_obj = OpenpyxlImage(temp_path)
          marker_from = AnchorMarker(
              col=0, colOff=150000, row=i - 1, rowOff=40000
          )
          marker_to = AnchorMarker(col=1, colOff=-150000, row=i, rowOff=-40000)
          img_obj.anchor = TwoCellAnchor(
              editAs="twoCell", _from=marker_from, to=marker_to
          )

          ws.add_image(img_obj)
          immagini_inserite += 1
        except Exception as e:
          st.warning(f"Errore caricamento immagine per {figura}: {e}")
      else:
        c_segnale.value = "[Img non trovata]"
        c_segnale.font = Font(name="Arial", size=8, color="FF0000")
        c_segnale.alignment = Alignment(horizontal="center", vertical="center")

    output = io.BytesIO()
    wb.save(output)
    excel_data = output.getvalue()

    st.success(
        "Abaco Excel Tavola 66 generato! (Immagini inserite:"
        f" {immagini_inserite}/{len(dati_righe_66)})"
    )
    st.download_button(
        label="📥 Scarica Excel Abaco Tavola 66",
        data=excel_data,
        file_name="Abaco_Segnaletica_Tavola_66.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  st.divider()
  st.subheader("🖼️ Anteprima Tavola 66 Personalizzata")

  img_path_66 = os.path.join(cartella_app, "TAVOLA 66.jpg")
  if os.path.exists(img_path_66):
    try:
      with Image.open(img_path_66) as immagine_originale:
        base_img_66 = immagine_originale.convert("RGB")

      width_66, height_66 = base_img_66.size
      draw_66 = ImageDraw.Draw(base_img_66)

      font_km = load_custom_font(13)
      font_ub_66 = load_custom_font(20)
      BORDO = 2

      # Mappatura coordinate per Tavola 66 (lato destro e lato sinistro)
      etichette_66 = [
          # --- LATO DESTRO (dal basso verso l'alto) ---
          (
              int(width_66 * 0.65),
              int(height_66 * 0.880),
              f"KM {km_66_dx_preavviso}",
          ),
          (int(width_66 * 0.65), int(height_66 * 0.825), f"KM {km_66_dx_limite60}"),
          (
              int(width_66 * 0.65),
              int(height_66 * 0.780),
              f"KM {km_66_dx_strettoia}",
          ),
          (int(width_66 * 0.65), int(height_66 * 0.727), f"KM {km_66_dx_semaforo}"),
          (int(width_66 * 0.65), int(height_66 * 0.600), f"KM {km_66_dx_chiusura}"),
          (int(width_66 * 0.65), int(height_66 * 0.370), f"KM {km_66_dx_fine}"),
          # --- LATO SINISTRO ---
          (int(width_66 * 0.0085), int(height_66 * 0.107), f"KM {km_66_sx_preavviso}"),
          (int(width_66 * 0.0085), int(height_66 * 0.152), f"KM {km_66_sx_limite60}"),
          (int(width_66 * 0.0085), int(height_66 * 0.200), f"KM {km_66_sx_strettoia}"),
          (int(width_66 * 0.0085), int(height_66 * 0.242), f"KM {km_66_sx_semaforo}"),
          (int(width_66 * 0.0085), int(height_66 * 0.370), f"KM {km_66_sx_zona}"),
      ]
      for x, y, testo in etichette_66:
        draw_66.text(
            (x, y),
            testo,
            fill=(255, 0, 0),
            font=font_km,
            stroke_width=BORDO,
            stroke_fill=(255, 255, 255),
        )

      # Testo Ubicazione Tavola 66
      x_ub_66 = int(width_66 * 0.65)
      y_ub_66 = int(height_66 * 0.245)
      draw_66.multiline_text(
          (x_ub_66, y_ub_66),
          ubicazione_testo_66,
          fill=(255, 0, 0),
          font=font_ub_66,
          align="center",
          spacing=6,
          stroke_width=BORDO,
          stroke_fill=(255, 255, 255),
      )

      img_io_66 = io.BytesIO()
      base_img_66.save(img_io_66, format="JPEG", quality=100, subsampling=0)
      jpg_data_66 = img_io_66.getvalue()

      st.image(
          base_img_66,
          caption="Tavola 66 - Configurazione completata",
          use_container_width=True,
      )
      st.download_button(
          label="📥 Scarica Tavola 66 in formato JPG",
          data=jpg_data_66,
          file_name="Tavola_66_Compilata.jpg",
          mime="image/jpeg",
      )
    except Exception as errore:
      st.error("Errore generazione Tavola 66.")
      st.exception(errore)
  else:
    st.error("Immagine 'TAVOLA 66.jpg' non trovata nella cartella.")
