import calendar
import datetime
import io
from io import BytesIO
import os
import urllib.request
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import streamlit as st

# ==================================================
# CONFIGURAZIONE STREAMLIT E PERCORSi
# ==================================================
st.set_page_config(
    page_title="Gestione Cantieri Stradali - Abaco & Tavole",
    page_icon="🚧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==================================================
# STILE GRAFICO MODERNO E PROFESSIONALE (CSS)
# ==================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Sfondo Principale Scuro */
    .stApp { 
        background-color: #0B0F19 !important; 
        color: #F8FAFC !important;
    }
    
    /* Sidebar Custom Scura */
    [data-testid="stSidebar"] {
        background-color: #111827 !important;
        border-right: 1px solid #1F2937 !important;
    }
    
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3, 
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] .stCaption {
        color: #F8FAFC !important;
    }

    /* FIX CAMPI DI TESTO (text_input e text_area): Sfondo scuro e testo chiaro ben visibile */
    .stTextInput input, .stTextArea textarea {
        background-color: #1F2937 !important;
        color: #FFFFFF !important;
        border: 1px solid #374151 !important;
        border-radius: 8px !important;
    }

    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #2563EB !important;
        box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.2) !important;
    }

    /* Etichette sopra i campi di testo */
    .stTextInput label, .stTextArea label, .stSelectbox label {
        color: #E2E8F0 !important;
        font-weight: 500 !important;
    }

    /* Selectbox nella Sidebar e nella pagina */
    div[data-baseweb="select"] > div {
        background-color: #1F2937 !important;
        color: #FFFFFF !important;
        border-color: #374151 !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="select"] * {
        color: #FFFFFF !important;
    }

    /* Hero Banner Stile Dark Mode */
    .app-hero {
        padding: 2rem; 
        margin-bottom: 2rem;
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
        border-radius: 16px; 
        color: #FFFFFF;
        box-shadow: 0 10px 25px -5px rgba(37, 99, 235, 0.3);
    }
    .app-hero h1 { 
        color: #FFFFFF !important; 
        font-size: 2.1rem !important; 
        font-weight: 700 !important;
        margin: 0 0 0.5rem 0 !important; 
    }
    .app-hero p { 
        color: #E0E7FF !important; 
        margin: 0 !important; 
        font-size: 1.05rem !important; 
        opacity: 0.95;
    }

    /* Titoli generali */
    h1, h2, h3, h4 { 
        color: #F8FAFC !important; 
        font-weight: 600 !important; 
    }

    /* Pulsanti */
    div.stButton > button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        background-color: #2563EB !important;
        color: #FFFFFF !important;
        border: none !important;
        padding: 0.6rem 1.2rem !important;
        transition: all 0.2s ease-in-out !important;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.3) !important;
    }
    div.stButton > button:hover {
        background-color: #1D4ED8 !important;
        transform: translateY(-1px);
    }

    div.stDownloadButton > button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        background-color: #059669 !important;
        color: #FFFFFF !important;
        border: none !important;
        padding: 0.6rem 1.2rem !important;
        box-shadow: 0 4px 6px -1px rgba(5, 150, 105, 0.3) !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #047857 !important;
        transform: translateY(-1px);
    }
</style>
""", unsafe_allow_html=True)

cartella_app = os.path.dirname(os.path.abspath(__file__))
dir_immagini = os.path.join(cartella_app, "immagini_segnali")

# Header Hero
st.markdown("""
<div class="app-hero">
  <h1>🚧 Gestione Cantieri Stradali e Segnaletica</h1>
  <p>Configurazione avanzata delle tavole di cantiere, generazione abachi Excel e immagini personalizzate.</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### 🛠️ Configurazione Cantiere")
    
    # ➕ AGGIUNGI QUESTA SELEZIONE
    sezione_app = st.radio(
        "Seleziona Funzione:",
        ["Gestione Tavole & Abaco", "Programma Lavori Mensile"]
    )
    
    st.divider() # Riga divisoria per pulizia grafica
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
# =============================================================
# CONTENUTO PRINCIPALE DELL'APP
# =============================================================

if sezione_app == "Gestione Tavole & Abaco":
        
        # ---------------------------------------------------------
        # QUI LASCIA TUTTO IL TUO CODICE ESISTENTE SULLE TAVOLE/ABACO
        # (Indentalo semplicemente dentro questo "if")
        # ---------------------------------------------------------
    # ==================================================
    # SELEZIONE DELLA TAVOLA
    # ==================================================
    # ==================================================
    # SELEZIONE DELLA TAVOLA
    # ==================================================
    # Sostituisci il vecchio blocco selectbox con questo:
        tavola_selezionata = st.sidebar.selectbox("Tavola di cantiere", [
        "Tavola 16 - Chiusura della corsia di marcia su carreggiata a due corsie",
        "Tavola 17 - Chiusura della corsia di sorpasso su carreggiata a due corsie",
        "Tavola 18 - Chiusura della corsia di destra su carreggiata a tre corsie",
        "Tavola 19 - Chiusura della corsia di sorpasso su carreggiata a tre corsie",
        "Tavola 20 - Chiusura delle corsie di destra e centrale su carreggiata a tre corsie",
        "Tavola 21 - Chiusura della corsia centrale e di sorpasso su carreggiata a tre corsie",
        "Tavola 22 - Chiusura della corsia di destra su carreggiata a tre corsie priva di corsia di emergenza",
        "Tavola 23 - Chiusura delle corsie di destra e centrale su carreggiata a tre corsie priva di corsia di emergenza (Unificata)",
        "Tavola 24 - Chiusura delle corsie di sorpasso e centrale su carreggiata a tre corsie priva di corsia di emergenza e spartitraffico ridotto (Unificata)",
        "Tavola 25 - Deviazione con una sola corsia per senso di marcia su carreggiata a due corsie",
        "Tavola 37 - Restringimento della carreggiata su rampa a senso unico",
        "Tavola 39 - Cantiere mobile carreggiata a due corsie",
        "Tavola 40 - Cantiere mobile su carreggiata a due corsie chiusura della corsia di sorpasso",
        "Tavola 43 - Cantiere mobile su carreggiata a tre corsie chiusura della corsia di destra",
        "Tavola 44 - Cantiere mobile su carreggiata a tre corsie chiusura della corsia di sorpasso",
        "Tavola 66 - Lavori su carreggiata con senso unico alternato semaforico",
        "Tavola 66 NJ - Lavori con senso unico alternato e New Jersey in plastica",
        "CHIUSURA CON SVINCOLO NJ - Chiusura con Svincolo e New Jersey",
        "CORSIA DI EMERGENZA NJ - Chiusura della corsia di emergenza"
    ])
        st.sidebar.divider()
        st.sidebar.caption("Compila i campi nella pagina principale, poi genera i file.")
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
    # TAVOLA 40
    # ====================================================================
        elif "40" in tavola_selezionata:
        
          st.subheader("📍 Dati e Progressive Chilometriche - Tavola 40")
        
          ubicazione_testo_40 = st.text_area(
              "Ubicazione / Direzione (Tavola 40)",
              value="DIREZIONE NORD\nKM VARIA",
          )
        
          col1, col2 = st.columns(2)
        
          with col1:
            km_40_750 = st.text_input(
                "KM Chiusura corsia a 750m (Basso)", value="35+560"
            )
            km_40_500 = st.text_input("KM Chiusura corsia a 500m", value="35+810")
            km_40_preavviso = st.text_input(
                "KM Segnale mobile di preavviso (Centro)", value="36+060"
            )
        
          with col2:
            km_40_testata = st.text_input("KM Inizio Testata", value="36+110")
            km_40_vialibera = st.text_input("KM Via Libera / Fine (Alto)", value="36+600")
        
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 40"):
        
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 40"
        
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 40"
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
        
            dati_righe_40 = [
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
        
            for i, row_data in enumerate(dati_righe_40, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_40_img_{i}.jpg")
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
                "Abaco Excel Tavola 40 generato con successo! (Immagini:"
                f" {immagini_inserite}/5)"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Segnaletica Tavola 40",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_40.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 40 Personalizzata")
        
          img_path_40 = os.path.join(cartella_app, "TAVOLA 40.jpg")
          if os.path.exists(img_path_40):
            try:
              with Image.open(img_path_40) as immagine_originale:
                base_img_40 = immagine_originale.convert("RGB")
        
              width_40, height_40 = base_img_40.size
              draw_40 = ImageDraw.Draw(base_img_40)
              font_km = load_custom_font(25)
              font_ubicazione_40 = load_custom_font(30)
              BORDO = 2
        
              # Mappatura coordinate progressive lato destro (dal basso verso l'alto)
              etichette_40 = [
                  (int(width_40 * 0.64), int(height_40 * 0.935), f"KM {km_40_750}"),
                  (int(width_40 * 0.64), int(height_40 * 0.769), f"KM {km_40_500}"),
                  (
                      int(width_40 * 0.64),
                      int(height_40 * 0.608),
                      f"KM {km_40_preavviso}",
                  ),
                  (int(width_40 * 0.64), int(height_40 * 0.497), f"KM {km_40_testata}"),
                  (int(width_40 * 0.41), int(height_40 * 0.077), f"KM {km_40_vialibera}"),
              ]
        
              for x, y, testo in etichette_40:
                draw_40.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              draw_40.multiline_text(
                  (int(width_40 * 0.655), int(height_40 * 0.330)),
                  ubicazione_testo_40,
                  fill=(255, 0, 0),
                  font=font_ubicazione_40,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_40 = io.BytesIO()
              base_img_40.save(img_io_40, format="JPEG", quality=100, subsampling=0)
              jpg_data_40 = img_io_40.getvalue()
        
              st.image(
                  base_img_40,
                  caption="Tavola 40 - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 40 in formato JPG",
                  data=jpg_data_40,
                  file_name="Tavola_40_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 40.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 40.jpg' non trovata nella cartella.")
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
        # TAVOLA 17
        # ====================================================================
        elif "17" in tavola_selezionata:
          st.subheader("📍 Dati e Progressive Chilometriche - Tavola 17")
          ubicazione_testo_17 = st.text_area(
              "Ubicazione / Direzione (Tavola 17)", value="DIREZIONE NORD"
          )
          col1, col2 = st.columns(2)
          with col1:
            km_17_via_libera = st.text_input(
                "KM Via Libera / Fine (In alto)", value="814+900"
            )
            km_17_stop = st.text_input("KM Stop", value="815+050")
            km_17_fine_testata = st.text_input("KM Fine Testata", value="815+350")
            km_17_inizio_testata = st.text_input("KM Inizio Testata", value="815+450")
            km_17_preavviso_100 = st.text_input(
                "KM Preavviso Chiusura Corsia 100m", value="815+550"
            )
          with col2:
            km_17_limite_50_alta = st.text_input("KM Limite 50 (superiore)", value="815+670")
            km_17_preavviso_250 = st.text_input(
                "KM Preavviso Chiusura Corsia 250m", value="815+790"
            )
            km_17_limite_50_bassa = st.text_input("KM Limite 50 (inferiore)", value="815+910")
            km_17_limite_70 = st.text_input("KM Limite 70", value="816+030")
            km_17_preavviso_700 = st.text_input(
                "KM Preavviso Lavori 700m (In basso)", value="816+150"
            )
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 17"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 17"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 17"
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
            dati_righe_17 = [
                ("fig_383.jpg", "FIG. II 383", "LAVORI IN CORSO", "", "SEGNALI TEMPORANEI"),
                (
                    "fig_082_b.jpg",
                    "FIG. II 082/b",
                    "PASSAGGIO OBBLIGATORIO A DESTRA",
                    "",
                    "SEGNALETICA DI OBBLIGO",
                ),
                (
                    "fig050_lim50.jpg",
                    "FIG. II 050",
                    "LIMITE MASSIMO DI VELOCITA' 50km/h",
                    "",
                    "SEGNALI DI DIVIETO",
                ),
                (
                    "fig050_lim70.jpg",
                    "FIG. II 050",
                    "LIMITE MASSIMO DI VELOCITA' 70km/h",
                    "",
                    "SEGNALI DI DIVIETO",
                ),
                (
                    "fig_048.jpg",
                    "FIG. II 048",
                    "DIVIETO DI SORPASSO",
                    "",
                    "SEGNALI DI DIVIETO",
                ),
                (
                    "fig_411a_sx.jpg",
                    "FIG. II 411/a",
                    "SEGNALE DI CORSIA CHIUSA (CORSIA DI SINISTRA)",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_37.jpg",
                    "FIG. II 37",
                    "FERMARSI E DARE PRECEDENZA",
                    "",
                    "SEGNALI DI PRECEDENZA",
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
            for i, row_data in enumerate(dati_righe_17, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_17_img_{i}.jpg")
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
                "Abaco Excel Tavola 17 generato! (Immagini inserite:"
                f" {immagini_inserite}/{len(dati_righe_17)})"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Tavola 17",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_17.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 17 Personalizzata")
          img_path_17 = os.path.join(cartella_app, "TAVOLA 17.jpg")
          if os.path.exists(img_path_17):
            try:
              with Image.open(img_path_17) as immagine_originale:
                base_img_17 = immagine_originale.convert("RGB")
              width_17, height_17 = base_img_17.size
              draw_17 = ImageDraw.Draw(base_img_17)
              font_km = load_custom_font(17)
              font_ub_17 = load_custom_font(25)
              BORDO = 2
              etichette_17 = [
                  (
                      int(width_17 * 0.466),
                      int(height_17 * 0.009),
                      f"KM {km_17_via_libera}",
                  ),
                  (int(width_17 * 0.63), int(height_17 * 0.265), f"KM {km_17_stop}"),
                  (
                      int(width_17 * 0.63),
                      int(height_17 * 0.326),
                      f"KM {km_17_fine_testata}",
                  ),
                  (
                      int(width_17 * 0.63),
                      int(height_17 * 0.484),
                      f"KM {km_17_inizio_testata}",
                  ),
                  (
                      int(width_17 * 0.63),
                      int(height_17 * 0.605),
                      f"KM {km_17_preavviso_100}",
                  ),
                  (
                      int(width_17 * 0.63),
                      int(height_17 * 0.670),
                      f"KM {km_17_limite_50_alta}",
                  ),
                  (
                      int(width_17 * 0.63),
                      int(height_17 * 0.735),
                      f"KM {km_17_preavviso_250}",
                  ),
                  (
                      int(width_17 * 0.63),
                      int(height_17 * 0.805),
                      f"KM {km_17_limite_50_bassa}",
                  ),
                  (int(width_17 * 0.63), int(height_17 * 0.867), f"KM {km_17_limite_70}"),
                  (
                      int(width_17 * 0.63),
                      int(height_17 * 0.932),
                      f"KM {km_17_preavviso_700}",
                  ),
              ]
              for x, y, testo in etichette_17:
                draw_17.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
              x_ub_17 = int(width_17 * 0.645)
              y_ub_17 = int(height_17 * 0.375)
              draw_17.multiline_text(
                  (x_ub_17, y_ub_17),
                  ubicazione_testo_17,
                  fill=(255, 0, 0),
                  font=font_ub_17,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
              img_io_17 = io.BytesIO()
              base_img_17.save(img_io_17, format="JPEG", quality=100, subsampling=0)
              jpg_data_17 = img_io_17.getvalue()
              st.image(
                  base_img_17,
                  caption="Tavola 17 - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 17 in formato JPG",
                  data=jpg_data_17,
                  file_name="Tavola_17_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 17.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 17.jpg' non trovata nella cartella.")
        
        # ====================================================================
        # TAVOLA 25
        # ====================================================================
        elif "25" in tavola_selezionata:
          st.subheader("📍 Dati e Progressive Chilometriche - Tavola 25")
        
          col_ub1, col_ub2 = st.columns(2)
          with col_ub1:
            ubicazione_nord = st.text_area(
                "Ubicazione / Direzione Sinistra", value="DIREZIONE NORD"
            )
          with col_ub2:
            ubicazione_sud = st.text_area(
                "Ubicazione / Direzione Destra", value="DIREZIONE SUD"
            )
        
          st.write("### Lato Destro dell'Immagine (dal basso verso l'alto)")
          col1, col2 = st.columns(2)
          with col1:
            km_25_dx_limite60 = st.text_input("KM Limite 60 (Dx)", value="811+872")
            km_25_dx_limite40 = st.text_input("KM Limite 40 (Dx)", value="811+968")
          with col2:
            km_25_dx_rientro = st.text_input(
                "KM Rientro in carreggiata (Dx)", value="812+064"
            )
            km_25_dx_immissione = st.text_input(
                "KM Inizio immissione (Dx)", value="812+100"
            )
        
          st.write("### Lato Sinistro dell'Immagine (dall'alto verso il basso)")
          col3, col4 = st.columns(2)
          with col3:
            km_25_sx_lavori = st.text_input(
                "KM Lavori in corso (Sx)", value="812+862"
            )
            km_25_sx_limite90 = st.text_input("KM Limite 90 (Sx)", value="812+742")
            km_25_sx_corsia250 = st.text_input(
                "KM Chiusura corsia 250m (Sx)", value="812+622"
            )
            km_25_sx_limite60 = st.text_input(
                "KM Limite 60 (Sx)", value="812+502"
            )
            km_25_sx_corsia100 = st.text_input(
                "KM Chiusura corsia 100m (Sx)", value="812+382"
            )
          with col4:
            km_25_sx_testata = st.text_input(
                "KM Inizio testata (Sx)", value="812+322"
            )
            km_25_sx_obbligo = st.text_input(
                "KM Passaggio obbligatorio (Sx)", value="812+286"
            )
            km_25_sx_doppiosenso = st.text_input(
                "KM Doppio senso di circolazione (Sx)", value="812+178"
            )
            km_25_sx_corsiaunica = st.text_input(
                "KM Corsia unica (Sx)", value="812+142"
            )
            km_25_sx_immissione_altro = st.text_input(
                "KM Inizio immissione altro senso (Sx)", value="812+100"
            )
        
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 25"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 25"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 25"
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
        
            dati_righe_25 = [
                (
                    "fig_050_lim60.jpg",
                    "FIG. II 050",
                    "LIMITE DI VELOCITÀ 60 KM/H",
                    "",
                    "SEGNALI DI PRESCRIZIONE",
                ),
                (
                    "fig050_lim40.jpg",
                    "FIG. II 050",
                    "LIMITE MASSIMO DI VELOCITÀ 40 KM/H",
                    "",
                    "SEGNALI DI DIVIETO",
                ),
                (
                    "fig_412b.jpg",
                    "FIG. II 412/b",
                    "SEGNALE DI RIENTRO IN CARREGGIATA",
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
                (
                    "fig_387.jpg",
                    "FIG. II 387",
                    "DOPPIO SENSO DI CIRCOLAZIONE",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_411a_sx.jpg",
                    "FIG. II 411/a",
                    "SEGNALE DI CORSIA CHIUSA (CORSIA DI SINISTRA)",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_048.jpg",
                    "FIG. II 048",
                    "DIVIETO DI SORPASSO",
                    "",
                    "SEGNALI DI DIVIETO",
                ),
                (
                    "fig050_lim90.jpg",
                    "FIG. II 050",
                    "LIMITE DI VELOCITÀ 90 KM/H",
                    "",
                    "SEGNALI DI PRESCRIZIONE",
                ),
                (
                    "fig_383.jpg",
                    "FIG. II 383",
                    "LAVORI IN CORSO",
                    "",
                    "SEGNALI TEMPORANEI",
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
            for i, row_data in enumerate(dati_righe_25, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_25_img_{i}.jpg")
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
                "Abaco Excel Tavola 25 generato! (Immagini inserite:"
                f" {immagini_inserite}/{len(dati_righe_25)})"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Tavola 25",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_25.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 25 Personalizzata")
          img_path_25 = os.path.join(cartella_app, "TAVOLA 25.jpg")
          if os.path.exists(img_path_25):
            try:
              with Image.open(img_path_25) as immagine_originale:
                base_img_25 = immagine_originale.convert("RGB")
              width_25, height_25 = base_img_25.size
              draw_25 = ImageDraw.Draw(base_img_25)
              font_km = load_custom_font(32)
              font_ub_25 = load_custom_font(56)
              BORDO = 2
        
              etichette_25 = [
                  # --- LATO DESTRO (dal basso verso l'alto) ---
                  (
                      int(width_25 * 0.77),
                      int(height_25 * 0.928),
                      f"KM {km_25_dx_limite60}",
                  ),
                  (
                      int(width_25 * 0.77),
                      int(height_25 * 0.860),
                      f"KM {km_25_dx_limite40}",
                  ),
                  (
                      int(width_25 * 0.77),
                      int(height_25 * 0.785),
                      f"KM {km_25_dx_rientro}",
                  ),
                  (
                      int(width_25 * 0.77),
                      int(height_25 * 0.715),
                      f"KM {km_25_dx_immissione}",
                  ),
                  # --- LATO SINISTRO (dall'alto verso il basso) ---
                  (
                      int(width_25 * 0.195),
                      int(height_25 * 0.244),
                      f"KM {km_25_sx_lavori}",
                  ),
                  (
                      int(width_25 * 0.195),
                      int(height_25 * 0.2845),
                      f"KM {km_25_sx_limite90}",
                  ),
                  (
                      int(width_25 * 0.195),
                      int(height_25 * 0.326),
                      f"KM {km_25_sx_corsia250}",
                  ),
                  (
                      int(width_25 * 0.195),
                      int(height_25 * 0.366),
                      f"KM {km_25_sx_limite60}",
                  ),
                  (
                      int(width_25 * 0.195),
                      int(height_25 * 0.408),
                      f"KM {km_25_sx_corsia100}",
                  ),
                  (
                      int(width_25 * 0.195),
                      int(height_25 * 0.443),
                      f"KM {km_25_sx_testata}",
                  ),
                  (
                      int(width_25 *0.195),
                      int(height_25 * 0.478),
                      f"KM {km_25_sx_obbligo}",
                  ),
                  (
                      int(width_25 * 0.195),
                      int(height_25 * 0.580),
                      f"KM {km_25_sx_doppiosenso}",
                  ),
                  (
                      int(width_25 * 0.195),
                      int(height_25 * 0.615),
                      f"KM {km_25_sx_corsiaunica}",
                  ),
                  (
                      int(width_25 * 0.195),
                      int(height_25 * 0.715),
                      f"KM {km_25_sx_immissione_altro}",
                  ),
              ]
        
              for x, y, testo in etichette_25:
                draw_25.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              # Ubicazione Sinistra
              draw_25.multiline_text(
                  (int(width_25 * 0.34), int(height_25 * 0.08)),
                  ubicazione_nord,
                  fill=(255, 0, 0),
                  font=font_ub_25,
                  align="center",
                  spacing=4,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              # Ubicazione Destra
              draw_25.multiline_text(
                  (int(width_25 * 0.81), int(height_25 * 0.08)),
                  ubicazione_sud,
                  fill=(255, 0, 0),
                  font=font_ub_25,
                  align="center",
                  spacing=4,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_25 = io.BytesIO()
              base_img_25.save(img_io_25, format="JPEG", quality=100, subsampling=0)
              jpg_data_25 = img_io_25.getvalue()
              st.image(
                  base_img_25,
                  caption="Tavola 25 - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 25 in formato JPG",
                  data=jpg_data_25,
                  file_name="Tavola_25_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 25.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 25.jpg' non trovata nella cartella.")
        
        # ====================================================================
        # TAVOLA 66 NJ
        # ====================================================================
        elif "66 NJ" in tavola_selezionata:
          st.subheader(
              "📍 Dati e Progressive Chilometriche - Tavola 66 NJ (New Jersey)"
          )
          ubicazione_testo_66_nj = st.text_area(
              "Ubicazione / Direzione (Tavola 66 NJ)",
              value="VIADOTTO MARCHITELLO I\nS.P. 3",
          )
          st.write("### Lato Destro dell'Immagine")
          col1, col2 = st.columns(2)
          with col1:
            km_66nj_dx_preavviso = st.text_input(
                "KM Preavviso Lavori a 200m (Dx)", value="5+065"
            )
            km_66nj_dx_limite60 = st.text_input("KM Limite 60 (Dx)", value="5+110")
            km_66nj_dx_strettoia = st.text_input(
                "KM Strettoia Asimmetrica a Destra (Dx)", value="5+155"
            )
          with col2:
            km_66nj_dx_semaforo = st.text_input("KM Semaforo (Dx)", value="5+200")
            km_66nj_dx_chiusura = st.text_input(
                "KM Chiusura Zona Lavori (Dx)", value="5+300"
            )
            km_66nj_dx_fine = st.text_input("KM Fine Zona Lavori (Dx)", value="5+600")
        
          st.write("### Lato Sinistro dell'Immagine")
          col3, col4 = st.columns(2)
          with col3:
            km_66nj_sx_zona = st.text_input("KM Zona Lavori (Sx)", value="5+600")
            km_66nj_sx_semaforo = st.text_input("KM Semaforo (Sx)", value="5+700")
            km_66nj_sx_strettoia = st.text_input(
                "KM Strettoia / Altro segnale (Sx)", value="5+745"
            )
          with col4:
            km_66nj_sx_limite60 = st.text_input(
                "KM Limite 60 / Divieto (Sx)", value="5+790"
            )
            km_66nj_sx_preavviso = st.text_input(
                "KM Preavviso Lavori (Sx)", value="5+835"
            )
        
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 66 NJ"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 66 NJ"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 66 NJ"
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
        
            # Abaco Tavola 66 con l'aggiunta del New Jersey in plastica
            dati_righe_66nj = [
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
                    "newjerseyplastica.jpg",
                    "NEW JERSEY",
                    "NEW JERSEY IN PLASTICA CON RETE ROSSA",
                    "",
                    "PROTEZIONE CANTIERE",
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
            for i, row_data in enumerate(dati_righe_66nj, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_66nj_img_{i}.jpg")
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
                "Abaco Excel Tavola 66 NJ generato! (Immagini inserite:"
                f" {immagini_inserite}/{len(dati_righe_66nj)})"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Tavola 66 NJ",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_66_NJ.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 66 NJ Personalizzata")
          img_path_66nj = os.path.join(cartella_app, "TAVOLA 66 NJ.jpg")
          if os.path.exists(img_path_66nj):
            try:
              with Image.open(img_path_66nj) as immagine_originale:
                base_img_66nj = immagine_originale.convert("RGB")
              width_66nj, height_66nj = base_img_66nj.size
              draw_66nj = ImageDraw.Draw(base_img_66nj)
              font_km = load_custom_font(14)
              font_ub_66nj = load_custom_font(20)
              BORDO = 2
        
              # Stesse coordinate della Tavola 66
              etichette_66nj = [
                  (
                      int(width_66nj * 0.65),
                      int(height_66nj * 0.880),
                      f"KM {km_66nj_dx_preavviso}",
                  ),
                  (
                      int(width_66nj * 0.65),
                      int(height_66nj * 0.825),
                      f"KM {km_66nj_dx_limite60}",
                  ),
                  (
                      int(width_66nj * 0.65),
                      int(height_66nj * 0.780),
                      f"KM {km_66nj_dx_strettoia}",
                  ),
                  (
                      int(width_66nj * 0.65),
                      int(height_66nj * 0.727),
                      f"KM {km_66nj_dx_semaforo}",
                  ),
                  (
                      int(width_66nj * 0.65),
                      int(height_66nj * 0.600),
                      f"KM {km_66nj_dx_chiusura}",
                  ),
                  (
                      int(width_66nj * 0.65),
                      int(height_66nj * 0.370),
                      f"KM {km_66nj_dx_fine}",
                  ),
                  (
                      int(width_66nj * 0.0075),
                      int(height_66nj * 0.107),
                      f"KM {km_66nj_sx_preavviso}",
                  ),
                  (
                      int(width_66nj * 0.0075),
                      int(height_66nj * 0.152),
                      f"KM {km_66nj_sx_limite60}",
                  ),
                  (
                      int(width_66nj * 0.0075),
                      int(height_66nj * 0.200),
                      f"KM {km_66nj_sx_strettoia}",
                  ),
                  (
                      int(width_66nj * 0.0075),
                      int(height_66nj * 0.242),
                      f"KM {km_66nj_sx_semaforo}",
                  ),
                  (
                      int(width_66nj * 0.0075),
                      int(height_66nj * 0.370),
                      f"KM {km_66nj_sx_zona}",
                  ),
              ]
        
              for x, y, testo in etichette_66nj:
                draw_66nj.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              x_ub_66nj = int(width_66nj * 0.65)
              y_ub_66nj = int(height_66nj * 0.245)
              draw_66nj.multiline_text(
                  (x_ub_66nj, y_ub_66nj),
                  ubicazione_testo_66_nj,
                  fill=(255, 0, 0),
                  font=font_ub_66nj,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_66nj = io.BytesIO()
              base_img_66nj.save(img_io_66nj, format="JPEG", quality=100, subsampling=0)
              jpg_data_66nj = img_io_66nj.getvalue()
              st.image(
                  base_img_66nj,
                  caption="Tavola 66 NJ - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 66 NJ in formato JPG",
                  data=jpg_data_66nj,
                  file_name="Tavola_66_NJ_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 66 NJ.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 66 NJ.jpg' non trovata nella cartella.")
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
              font_km = load_custom_font(14)
              font_ub_66 = load_custom_font(20)
              BORDO = 2
              etichette_66 = [
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
                  (int(width_66 * 0.0075), int(height_66 * 0.107), f"KM {km_66_sx_preavviso}"),
                  (int(width_66 * 0.0075), int(height_66 * 0.152), f"KM {km_66_sx_limite60}"),
                  (int(width_66 * 0.0075), int(height_66 * 0.200), f"KM {km_66_sx_strettoia}"),
                  (int(width_66 * 0.0075), int(height_66 * 0.242), f"KM {km_66_sx_semaforo}"),
                  (int(width_66 * 0.0075), int(height_66 * 0.370), f"KM {km_66_sx_zona}"),
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
        
        # ====================================================================
        # TAVOLA 18
        # ====================================================================
        elif "18" in tavola_selezionata:
          st.subheader("📍 Dati e Progressive Chilometriche - Tavola 18")
          ubicazione_testo_18 = st.text_area(
              "Ubicazione / Direzione (Tavola 18)", value="DIREZIONE SUD"
          )
          st.write("### Lato Destro dell'Immagine (dal basso verso l'alto)")
          col1, col2 = st.columns(2)
          with col1:
            km_18_lavori = st.text_input("KM Lavori in corso", value="805+860")
            km_18_limite70 = st.text_input("KM Limite 70 km/h", value="805+980")
            km_18_limite50 = st.text_input("KM Limite 50 km/h", value="806+100")
            km_18_corsia250 = st.text_input("KM Chiusura corsia a 250 m", value="806+220")
            km_18_sorpasso = st.text_input("KM Divieto di sorpasso", value="806+340")
          with col2:
            km_18_corsia100 = st.text_input(
                "KM Chiusura corsia 100 metri", value="806+460"
            )
            km_18_inizio_chiusura = st.text_input("KM Inizio chiusura", value="806+556")
            km_18_ultima_chiusura = st.text_input("KM Ultima chiusura", value="806+700")
            km_18_via_libera = st.text_input("KM Via libera", value="807+300")
        
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 18"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 18"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 18"
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
        
            dati_righe_18 = [
                (
                    "fig_383.jpg",
                    "FIG. II 383",
                    "LAVORI IN CORSO",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig050_lim70.jpg",
                    "FIG. II 050",
                    "LIMITE DI VELOCITÀ 70 KM/H",
                    "",
                    "SEGNALI DI PRESCRIZIONE",
                ),
                (
                    "fig050_lim50.jpg",
                    "FIG. II 050",
                    "LIMITE DI VELOCITÀ 50 KM/H",
                    "",
                    "SEGNALI DI PRESCRIZIONE",
                ),
                (
                    "fig_411a.jpg",
                    "FIG. II 411/a",
                    "SEGNALE DI CORSIA CHIUSA (Chiusura corsia di destra)",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_052.jpg",
                    "FIG. II 052",
                    "DIVIETO DI SORPASSO PER VEICOLI DI MASSA > 3,5t",
                    "",
                    "SEGNALI DI DIVIETO",
                ),
                (
                    "fig_082a.jpg",
                    "FIG. II 082/a",
                    "PASSAGGIO OBBLIGATORIO A SINISTRA",
                    "",
                    "SEGNALI DI OBBLIGO",
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
            for i, row_data in enumerate(dati_righe_18, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_18_img_{i}.jpg")
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
                "Abaco Excel Tavola 18 generato! (Immagini inserite:"
                f" {immagini_inserite}/{len(dati_righe_18)})"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Tavola 18",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_18.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 18 Personalizzata")
          img_path_18 = os.path.join(cartella_app, "TAVOLA 18.jpg")
          if os.path.exists(img_path_18):
            try:
              with Image.open(img_path_18) as immagine_originale:
                base_img_18 = immagine_originale.convert("RGB")
              width_18, height_18 = base_img_18.size
              draw_18 = ImageDraw.Draw(base_img_18)
              font_km = load_custom_font(17)
              font_ub_18 = load_custom_font(26)
              BORDO = 2
        
              etichette_18 = [
                  (int(width_18 * 0.65), int(height_18 * 0.928), f"KM {km_18_lavori}"),
                  (int(width_18 * 0.65), int(height_18 * 0.860), f"KM {km_18_limite70}"),
                  (int(width_18 * 0.65), int(height_18 * 0.799), f"KM {km_18_limite50}"),
                  (
                      int(width_18 * 0.65),
                      int(height_18 * 0.735),
                      f"KM {km_18_corsia250}",
                  ),
                  (int(width_18 * 0.65), int(height_18 * 0.670), f"KM {km_18_sorpasso}"),
                  (
                      int(width_18 * 0.65),
                      int(height_18 * 0.605),
                      f"KM {km_18_corsia100}",
                  ),
                  (
                      int(width_18 * 0.65),
                      int(height_18 * 0.470),
                      f"KM {km_18_inizio_chiusura}",
                  ),
                  (
                      int(width_18 * 0.65),
                      int(height_18 * 0.270),
                      f"KM {km_18_ultima_chiusura}",
                  ),
                  (int(width_18 * 0.49), int(height_18 * 0.052), f"KM {km_18_via_libera}"),
              ]
        
              for x, y, testo in etichette_18:
                draw_18.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              draw_18.multiline_text(
                  (int(width_18 * 0.668), int(height_18 * 0.315)),
                  ubicazione_testo_18,
                  fill=(255, 0, 0),
                  font=font_ub_18,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_18 = io.BytesIO()
              base_img_18.save(img_io_18, format="JPEG", quality=100, subsampling=0)
              jpg_data_18 = img_io_18.getvalue()
              st.image(
                  base_img_18,
                  caption="Tavola 18 - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 18 in formato JPG",
                  data=jpg_data_18,
                  file_name="Tavola_18_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 18.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 18.jpg' non trovata nella cartella.")
        
        # ====================================================================
        # CORSIA DI EMERGENZA NJ
        # ====================================================================
        elif "CORSIA DI EMERGENZA NJ" in tavola_selezionata:
          st.subheader(
              "📍 Dati e Progressive Chilometriche - Chiusura Corsia di Emergenza NJ"
          )
          ubicazione_testo_em = st.text_area(
              "Ubicazione / Direzione", value="CARREGGIATA SUD"
          )
          st.write("### Progressive Chilometriche (dal basso verso l'alto)")
          col1, col2 = st.columns(2)
          with col1:
            km_em_rallentare = st.text_input("KM Rallentare (In basso)", value="813+950")
            km_em_lavori = st.text_input("KM Lavori in corso", value="814+050")
            km_em_passaggio = st.text_input("KM Passaggio obbligatorio", value="814+200")
          with col2:
            km_em_barriera = st.text_input("KM Barriera", value="814+250")
            km_em_fine = st.text_input("KM Fine / Inizio tratto (In alto)", value="814+450")
        
          if st.button("Genera Abaco Segnaletica (Excel) - Corsia di Emergenza NJ"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Corsia di Emergenza"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - CHIUSURA CORSIA DI"
                " EMERGENZA"
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
        
            dati_righe_em = [
                (
                    "fig_421.jpg",
                    "FIG. II 421",
                    "RALLENTARE OPERAI AL LAVORO",
                    "",
                    "SEGNALI DI PREAVVISO",
                ),
                (
                    "fig_383_n.jpg",
                    "FIG. II 383",
                    "LAVORI IN CORSO",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_082a.jpg",
                    "FIG. II 082/a",
                    "PASSAGGIO OBBLIGATORIO A SINISTRA",
                    "",
                    "SEGNALI DI OBBLIGO",
                ),
                (
                    "fig_392.jpg",
                    "FIG. II 392",
                    "BARRIERA NORMALE",
                    "",
                    "SEGNALI COMPLEMENTARI",
                ),
                (
                    "NEWJERSEY.jpg",
                    "NEW JERSEY",
                    "NEW JERSEY IN CLS A PROTEZIONE",
                    "",
                    "PROTEZIONE CANTIERE",
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
            for i, row_data in enumerate(dati_righe_em, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_em_img_{i}.jpg")
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
                "Abaco Excel Corsia di Emergenza generato! (Immagini inserite:"
                f" {immagini_inserite}/{len(dati_righe_em)})"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Corsia di Emergenza",
                data=excel_data,
                file_name="Abaco_Segnaletica_Corsia_Emergenza.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
          st.divider()
          st.subheader("🖼️ Anteprima Corsia di Emergenza NJ Personalizzata")
          img_path_em = os.path.join(cartella_app, "CORSIA DI EMERGENZA NJ.jpg")
          if os.path.exists(img_path_em):
            try:
              with Image.open(img_path_em) as immagine_originale:
                base_img_em = immagine_originale.convert("RGB")
              width_em, height_em = base_img_em.size
              draw_em = ImageDraw.Draw(base_img_em)
              font_km = load_custom_font(39)
              font_ub_em = load_custom_font(50)
              BORDO = 2
        
              # Coordinate provvisorie da calibrare in base alle posizioni sull'immagine pulita
              etichette_em = [
                  (int(width_em * 0.56), int(height_em * 0.885), f"KM {km_em_rallentare}"),
                  (int(width_em * 0.56), int(height_em * 0.7715), f"KM {km_em_lavori}"),
                  (int(width_em * 0.56), int(height_em * 0.615), f"KM {km_em_passaggio}"),
                  (int(width_em * 0.56), int(height_em * 0.553), f"KM {km_em_barriera}"),
                  (int(width_em * 0.38), int(height_em * 0.028), f"KM {km_em_fine}"),
              ]
        
              for x, y, testo in etichette_em:
                draw_em.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              # Testo fisso "CHIUSURA DELLA CORSIA DI EMERGENZA" già presente o da aggiungere se necessario, più l'ubicazione
              draw_em.multiline_text(
                  (int(width_em * 0.56), int(height_em * 0.420)),
                  ubicazione_testo_em,
                  fill=(255, 0, 0),
                  font=font_ub_em,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_em = io.BytesIO()
              base_img_em.save(img_io_em, format="JPEG", quality=100, subsampling=0)
              jpg_data_em = img_io_em.getvalue()
              st.image(
                  base_img_em,
                  caption="Corsia di Emergenza NJ - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola Corsia di Emergenza in formato JPG",
                  data=jpg_data_em,
                  file_name="Corsia_Emergenza_NJ_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola Corsia di Emergenza.")
              st.exception(errore)
          else:
            st.error("Immagine 'CORSIA DI EMERGENZA NJ.jpg' non trovata nella cartella.")
        
        # ====================================================================
        # TAVOLA 43
        # ====================================================================
        elif "43" in tavola_selezionata:
          st.subheader("📍 Dati e Progressive Chilometriche - Tavola 43")
          ubicazione_testo_43 = st.text_area(
              "Ubicazione / Direzione (Tavola 43)",
              value="DIREZIONE SUD\nCANTIERE MOBILE",
          )
          st.write("### Progressive Chilometriche (dal basso verso l'alto)")
          col1, col2 = st.columns(2)
          with col1:
            km_43_primo_preavviso = st.text_input(
                "KM Primo preavviso (Basso)", value="35+560"
            )
            km_43_secondo_preavviso = st.text_input("KM Secondo preavviso", value="36+060")
            km_43_passaggio = st.text_input("KM Passaggio obbligatorio", value="36+110")
          with col2:
            km_43_inizio = st.text_input("KM Inizio cantiere", value="36+400")
            km_43_vialibera = st.text_input("KM Via libera / Fine (Alto)", value="36+600")
        
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 43"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 43"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 43"
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
        
            dati_righe_43 = [
                (
                    "fig_399.jpg",
                    "FIG. II 399",
                    "PREAVVISO DI CANTIERE MOBILE",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_383.jpg",
                    "FIG. II 383",
                    "LAVORI IN CORSO",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
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
            for i, row_data in enumerate(dati_righe_43, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_43_img_{i}.jpg")
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
                "Abaco Excel Tavola 43 generato con successo! (Immagini:"
                f" {immagini_inserite}/5)"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Segnaletica Tavola 43",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_43.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 43 Personalizzata")
          img_path_43 = os.path.join(cartella_app, "TAVOLA 43.jpg")
          if os.path.exists(img_path_43):
            try:
              with Image.open(img_path_43) as immagine_originale:
                base_img_43 = immagine_originale.convert("RGB")
              width_43, height_43 = base_img_43.size
              draw_43 = ImageDraw.Draw(base_img_43)
              font_km = load_custom_font(45)
              font_ub_43 = load_custom_font(60)
              BORDO = 2
        
              # Coordinate mappate sulla Tavola 43 (dal basso verso l'alto)
              etichette_43 = [
                  (
                      int(width_43 * 0.64),
                      int(height_43 * 0.880),
                      f"KM {km_43_primo_preavviso}",
                  ),
                  (
                      int(width_43 * 0.64),
                      int(height_43 * 0.594),
                      f"KM {km_43_secondo_preavviso}",
                  ),
                  (int(width_43 * 0.64), int(height_43 * 0.420), f"KM {km_43_passaggio}"),
                  (int(width_43 * 0.627), int(height_43 * 0.274), f"KM {km_43_inizio}"),
                  (int(width_43 * 0.46), int(height_43 * 0.037), f"KM {km_43_vialibera}"),
              ]
        
              for x, y, testo in etichette_43:
                draw_43.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              draw_43.multiline_text(
                  (int(width_43 * 0.67), int(height_43 * 0.320)),
                  ubicazione_testo_43,
                  fill=(255, 0, 0),
                  font=font_ub_43,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_43 = io.BytesIO()
              base_img_43.save(img_io_43, format="JPEG", quality=100, subsampling=0)
              jpg_data_43 = img_io_43.getvalue()
              st.image(
                  base_img_43,
                  caption="Tavola 43 - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 43 in formato JPG",
                  data=jpg_data_43,
                  file_name="Tavola_43_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 43.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 43.jpg' non trovata nella cartella.")
        
        # ====================================================================
        # TAVOLA 44
        # ====================================================================
        elif "44" in tavola_selezionata:
          st.subheader("📍 Dati e Progressive Chilometriche - Tavola 44")
          ubicazione_testo_44 = st.text_area(
              "Ubicazione / Direzione (Tavola 44)",
              value="DIREZIONE SUD\nCANTIERE MOBILE",
          )
          st.write("### Progressive Chilometriche (dal basso verso l'alto)")
          col1, col2 = st.columns(2)
          with col1:
            km_44_primo_preavviso = st.text_input(
                "KM Primo preavviso (Basso)", value="35+560"
            )
            km_44_secondo_preavviso = st.text_input("KM Secondo preavviso", value="36+060")
            km_44_passaggio = st.text_input("KM Passaggio obbligatorio", value="36+110")
          with col2:
            km_44_inizio = st.text_input("KM Inizio cantiere", value="36+400")
            km_44_vialibera = st.text_input("KM Via libera / Fine (Alto)", value="36+600")
        
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 44"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 44"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 44"
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
        
            dati_righe_44 = [
                (
                    "fig_399.jpg",
                    "FIG. II 399",
                    "PREAVVISO DI CANTIERE MOBILE",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_383.jpg",
                    "FIG. II 383",
                    "LAVORI IN CORSO",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
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
            for i, row_data in enumerate(dati_righe_44, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_44_img_{i}.jpg")
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
                "Abaco Excel Tavola 44 generato con successo! (Immagini:"
                f" {immagini_inserite}/5)"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Segnaletica Tavola 44",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_44.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 44 Personalizzata")
          img_path_44 = os.path.join(cartella_app, "TAVOLA 44.jpg")
          if os.path.exists(img_path_44):
            try:
              with Image.open(img_path_44) as immagine_originale:
                base_img_44 = immagine_originale.convert("RGB")
              width_44, height_44 = base_img_44.size
              draw_44 = ImageDraw.Draw(base_img_44)
              font_km = load_custom_font(45)
              font_ub_44 = load_custom_font(65)
              BORDO = 2
        
              # Coordinate mappate sulla Tavola 44 (dal basso verso l'alto)
              etichette_44 = [
                  (
                      int(width_44 * 0.62),
                      int(height_44 * 0.895),
                      f"KM {km_44_primo_preavviso}",
                  ),
                  (
                      int(width_44 * 0.62),
                      int(height_44 * 0.735),
                      f"KM {km_44_secondo_preavviso}",
                  ),
                  (int(width_44 * 0.62), int(height_44 * 0.540), f"KM {km_44_passaggio}"),
                  (int(width_44 * 0.608), int(height_44 * 0.430), f"KM {km_44_inizio}"),
                  (int(width_44 * 0.435), int(height_44 * 0.040), f"KM {km_44_vialibera}"),
              ]
        
              for x, y, testo in etichette_44:
                draw_44.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              draw_44.multiline_text(
                  (int(width_44 * 0.65), int(height_44 * 0.250)),
                  ubicazione_testo_44,
                  fill=(255, 0, 0),
                  font=font_ub_44,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_44 = io.BytesIO()
              base_img_44.save(img_io_44, format="JPEG", quality=100, subsampling=0)
              jpg_data_44 = img_io_44.getvalue()
              st.image(
                  base_img_44,
                  caption="Tavola 44 - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 44 in formato JPG",
                  data=jpg_data_44,
                  file_name="Tavola_44_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 44.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 44.jpg' non trovata nella cartella.")
        
        # ====================================================================
        # TAVOLA 37
        # ====================================================================
        elif "37" in tavola_selezionata:
          st.subheader("📍 Dati e Progressive Chilometriche - Tavola 37")
          ubicazione_testo_37 = st.text_area(
              "Ubicazione / Direzione (Tavola 37)",
              value="DIREZIONE SUD\nRAMPA A SENSO UNICO",
          )
          st.write("### Progressive Chilometriche (dal basso verso l'alto)")
          col1, col2 = st.columns(2)
          with col1:
            km_37_lavori = st.text_input("KM Lavori in corso (Basso)", value="10+100")
            km_37_strettoia = st.text_input("KM Strettoia", value="10+150")
            km_37_passaggio = st.text_input("KM Passaggio obbligatorio", value="10+200")
          with col2:
            km_37_chiusura = st.text_input("KM Chiusura restringimento", value="10+250")
            km_37_vialibera = st.text_input("KM Via libera / Fine (Alto)", value="10+400")
        
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 37"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 37"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 37"
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
        
            dati_righe_37 = [
                (
                    "fig_383.jpg",
                    "FIG. II 383",
                    "LAVORI IN CORSO",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_386.jpg",
                    "FIG. II 386",
                    "STRETTOIA ASIMMETRICA",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_082a.jpg",
                    "FIG. II 082/a",
                    "PASSAGGIO OBBLIGATORIO A SINISTRA",
                    "",
                    "SEGNALI DI OBBLIGO",
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
            ws.column_dimensions["B"].width = 15
            ws.column_dimensions["C"].width = 40
            ws.column_dimensions["D"].width = 20
            ws.column_dimensions["E"].width = 25
            ws.row_dimensions[1].height = 30
            ws.row_dimensions[2].height = 25
            immagini_inserite = 0
            for i, row_data in enumerate(dati_righe_37, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_37_img_{i}.jpg")
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
                "Abaco Excel Tavola 37 generato con successo! (Immagini:"
                f" {immagini_inserite}/4)"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Segnaletica Tavola 37",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_37.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 37 Personalizzata")
          img_path_37 = os.path.join(cartella_app, "TAVOLA 37.jpg")
          if os.path.exists(img_path_37):
            try:
              with Image.open(img_path_37) as immagine_originale:
                base_img_37 = immagine_originale.convert("RGB")
              width_37, height_37 = base_img_37.size
              draw_37 = ImageDraw.Draw(base_img_37)
              font_km = load_custom_font(44)
              font_ub_37 = load_custom_font(58)
              BORDO = 2
        
              # Coordinate mappate sulla Tavola 37 (dal basso verso l'alto)
              etichette_37 = [
                  (int(width_37 * 0.62), int(height_37 * 0.870), f"KM {km_37_lavori}"),
                  (int(width_37 * 0.62), int(height_37 * 0.7925), f"KM {km_37_strettoia}"),
                  (int(width_37 * 0.62), int(height_37 * 0.705), f"KM {km_37_passaggio}"),
                  (int(width_37 * 0.62), int(height_37 * 0.490), f"KM {km_37_chiusura}"),
                  (int(width_37 * 0.48), int(height_37 * 0.0680), f"KM {km_37_vialibera}"),
              ]
        
              for x, y, testo in etichette_37:
                draw_37.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              draw_37.multiline_text(
                  (int(width_37 * 0.65), int(height_37 * 0.260)),
                  ubicazione_testo_37,
                  fill=(255, 0, 0),
                  font=font_ub_37,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_37 = io.BytesIO()
              base_img_37.save(img_io_37, format="JPEG", quality=100, subsampling=0)
              jpg_data_37 = img_io_37.getvalue()
              st.image(
                  base_img_37,
                  caption="Tavola 37 - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 37 in formato JPG",
                  data=jpg_data_37,
                  file_name="Tavola_37_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 37.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 37.jpg' non trovata nella cartella.")
        
        # ====================================================================
        # CHIUSURA CON SVINCOLO NJ
        # ====================================================================
        elif "CHIUSURA CON SVINCOLO NJ" in tavola_selezionata:
          st.subheader("📍 Dati e Progressive Chilometriche - Chiusura con Svincolo NJ")
          ubicazione_testo_svincolo = st.text_area(
              "Ubicazione / Direzione",
              value="S.S. 16\nSVINCOLO COZZE\nDIR. SUD",
          )
          st.write("### Progressive Chilometriche (dal basso verso l'alto)")
          col1, col2 = st.columns(2)
          with col1:
            km_sv_lavori = st.text_input("KM Lavori in corso (In basso)", value="829+076")
            km_sv_limite70 = st.text_input("KM Limite 70", value="828+956")
            km_sv_restringimento1 = st.text_input(
                "KM Restringimento carreggiata 1", value="828+836"
            )
          with col2:
            km_sv_limite50 = st.text_input("KM Limite 50", value="828+716")
            km_sv_restringimento2 = st.text_input(
                "KM Restringimento carreggiata 2", value="828+596"
            )
            km_sv_fine = st.text_input("KM Fine lavori (In alto)", value="827+750")
        
          if st.button("Genera Abaco Segnaletica (Excel) - Chiusura con Svincolo NJ"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Svincolo NJ"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - CHIUSURA CON SVINCOLO"
                " NJ"
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
        
            dati_righe_svincolo = [
                (
                    "fig_383.jpg",
                    "FIG. II 383",
                    "LAVORI IN CORSO",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_082a.jpg",
                    "FIG. II 082/a",
                    "PASSAGGIO OBBLIGATORIO A SINISTRA",
                    "",
                    "SEGNALI DI OBBLIGO",
                ),
                (
                    "fig_083.jpg",
                    "FIG. II 083",
                    "PASSAGGI CONSENTITI",
                    "",
                    "SEGNALI DI OBBLIGO",
                ),
                (
                    "NEWJERSEY.jpg",
                    "NEW JERSEY",
                    "NEW JERSEY IN CLS A PROTEZIONE",
                    "",
                    "PROTEZIONE CANTIERE",
                ),
                (
                    "fig_392.jpg",
                    "FIG. II 392",
                    "BARRIERA NORMALE",
                    "",
                    "SEGNALI COMPLEMENTARI",
                ),
                (
                    "fig050_lim40.jpg",
                    "FIG. II 050",
                    "LIMITE MASSIMO DI VELOCITÀ 40 KM/H",
                    "",
                    "SEGNALI DI DIVIETO",
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
                    "SEGNALE DI CORSIA CHIUSA",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_2101.jpg",
                    "FIG. 2101",
                    "RAMPA CHIUSA PER LAVORI",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_407.jpg",
                    "FIG. 407",
                    "SEGNALI DI DIREZIONE",
                    "",
                    "SEGNALI TEMPORANEI",
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
            for i, row_data in enumerate(dati_righe_svincolo, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_svincolo_img_{i}.jpg")
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
                "Abaco Excel Chiusura con Svincolo NJ generato! (Immagini inserite:"
                f" {immagini_inserite}/{len(dati_righe_svincolo)})"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Chiusura con Svincolo NJ",
                data=excel_data,
                file_name="Abaco_Segnaletica_Chiusura_Svincolo_NJ.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        
          st.divider()
          st.subheader("🖼️ Anteprima Chiusura con Svincolo NJ Personalizzata")
          img_path_sv = os.path.join(cartella_app, "CHIUSURA CON SVINCOLO NJ.jpg")
          if os.path.exists(img_path_sv):
            try:
              with Image.open(img_path_sv) as immagine_originale:
                base_img_sv = immagine_originale.convert("RGB")
              width_sv, height_sv = base_img_sv.size
              draw_sv = ImageDraw.Draw(base_img_sv)
              font_km = load_custom_font(20)
              font_ub_sv = load_custom_font(29)
              BORDO = 2
        
              # Mappatura coordinate progressive e ubicazione
              etichette_sv = [
                  (int(width_sv * 0.73), int(height_sv * 0.950), f"KM {km_sv_lavori}"),
                  (int(width_sv * 0.73), int(height_sv * 0.920), f"KM {km_sv_limite70}"),
                  (
                      int(width_sv * 0.73),
                      int(height_sv * 0.880),
                      f"KM {km_sv_restringimento1}",
                  ),
                  (int(width_sv * 0.75), int(height_sv * 0.838), f"KM {km_sv_limite50}"),
                  (
                      int(width_sv * 0.73),
                      int(height_sv * 0.7970),
                      f"KM {km_sv_restringimento2}",
                  ),
                  (int(width_sv * 0.675), int(height_sv * 0.009), f"KM {km_sv_fine}"),
              ]
        
              for x, y, testo in etichette_sv:
                draw_sv.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              draw_sv.multiline_text(
                  (int(width_sv * 0.565), int(height_sv * 0.110)),
                  ubicazione_testo_svincolo,
                  fill=(255, 0, 0),
                  font=font_ub_sv,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_sv = io.BytesIO()
              base_img_sv.save(img_io_sv, format="JPEG", quality=100, subsampling=0)
              jpg_data_sv = img_io_sv.getvalue()
              st.image(
                  base_img_sv,
                  caption="Chiusura con Svincolo NJ - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola Chiusura con Svincolo NJ in formato JPG",
                  data=jpg_data_sv,
                  file_name="Chiusura_Svincolo_NJ_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Chiusura con Svincolo NJ.")
              st.exception(errore)
          else:
            st.error("Immagine 'CHIUSURA CON SVINCOLO NJ.jpg' non trovata nella cartella.")
        
        # ====================================================================
        # TAVOLA 19
        # ====================================================================
        elif "Tavola 19" in tavola_selezionata:
          st.subheader("📍 Dati e Progressive Chilometriche - Tavola 19")
          ubicazione_testo_19 = st.text_area(
              "Ubicazione / Direzione (Tavola 19)", value="DIREZIONE SUD"
          )
          st.write("### Progressive Chilometriche (dal basso verso l'alto)")
          col1, col2 = st.columns(2)
          with col1:
            km_19_lavori = st.text_input(
                "KM Lavori in corso (In basso)", value="805+860"
            )
            km_19_limite110 = st.text_input("KM Limite 110 km/h", value="805+980")
            km_19_limite90 = st.text_input("KM Limite 90 km/h", value="806+100")
            km_19_restr1 = st.text_input("KM Restringimento carreggiata 1", value="806+220")
            km_19_sorpasso = st.text_input("KM Divieto di sorpasso", value="806+340")
          with col2:
            km_19_restr2 = st.text_input("KM Restringimento carreggiata 2", value="806+460")
            km_19_iniziotestata = st.text_input("KM Inizio testata", value="806+556")
            km_19_finetestata = st.text_input("KM Fine testata", value="806+700")
            km_19_vialibera = st.text_input(
                "KM Via libera / Fine (In alto)", value="807+300"
            )
        
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 19"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 19"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 19"
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
        
            dati_righe_19 = [
                (
                    "fig_383.jpg",
                    "FIG. II 383",
                    "LAVORI IN CORSO",
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
                    "fig_411a.jpg",
                    "FIG. II 411/a",
                    "SEGNALE DI CORSIA CHIUSA (Chiusura corsia di sinistra)",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_052.jpg",
                    "FIG. II 052",
                    "DIVIETO DI SORPASSO PER VEICOLI DI MASSA > 3,5t",
                    "",
                    "SEGNALI DI DIVIETO",
                ),
                (
                    "fig_082_b.jpg",
                    "FIG. II 082/b",
                    "PASSAGGIO OBBLIGATORIO A DESTRA",
                    "",
                    "SEGNALETICA DI OBBLIGO",
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
            for i, row_data in enumerate(dati_righe_19, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_19_img_{i}.jpg")
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
                "Abaco Excel Tavola 19 generato! (Immagini inserite:"
                f" {immagini_inserite}/{len(dati_righe_19)})"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Tavola 19",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_19.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 19 Personalizzata")
          img_path_19 = os.path.join(cartella_app, "TAVOLA 19.jpg")
          if os.path.exists(img_path_19):
            try:
              with Image.open(img_path_19) as immagine_originale:
                base_img_19 = immagine_originale.convert("RGB")
              width_19, height_19 = base_img_19.size
              draw_19 = ImageDraw.Draw(base_img_19)
              font_km = load_custom_font(22)
              font_ub_19 = load_custom_font(35)
              BORDO = 2
        
              # Mappatura coordinate progressive (speculare/adattata alla Tavola 19)
              etichette_19 = [
                  (int(width_19 * 0.53), int(height_19 * 0.905), f"KM {km_19_lavori}"),
                  (
                      int(width_19 * 0.53),
                      int(height_19 * 0.845),
                      f"KM {km_19_limite110}",
                  ),
                  (int(width_19 * 0.53), int(height_19 * 0.785), f"KM {km_19_limite90}"),
                  (int(width_19 * 0.53), int(height_19 * 0.724), f"KM {km_19_restr1}"),
                  (int(width_19 * 0.53), int(height_19 * 0.664), f"KM {km_19_sorpasso}"),
                  (int(width_19 * 0.53), int(height_19 * 0.600), f"KM {km_19_restr2}"),
                  (
                      int(width_19 * 0.53),
                      int(height_19 * 0.492),
                      f"KM {km_19_iniziotestata}",
                  ),
                  (
                      int(width_19 * 0.53),
                      int(height_19 * 0.287),
                      f"KM {km_19_finetestata}",
                  ),
                  (int(width_19 * 0.53), int(height_19 * 0.120), f"KM {km_19_vialibera}"),
              ]
        
              for x, y, testo in etichette_19:
                draw_19.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              draw_19.multiline_text(
                  (int(width_19 * 0.668), int(height_19 * 0.345)),
                  ubicazione_testo_19,
                  fill=(255, 0, 0),
                  font=font_ub_19,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_19 = io.BytesIO()
              base_img_19.save(img_io_19, format="JPEG", quality=100, subsampling=0)
              jpg_data_19 = img_io_19.getvalue()
              st.image(
                  base_img_19,
                  caption="Tavola 19 - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 19 in formato JPG",
                  data=jpg_data_19,
                  file_name="Tavola_19_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 19.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 19.jpg' non trovata nella cartella.")
        
        # ====================================================================
        # TAVOLA 20
        # ====================================================================
        elif "Tavola 20" in tavola_selezionata:
          st.subheader("📍 Dati e Progressive Chilometriche - Tavola 20")
          ubicazione_testo_20 = st.text_area(
              "Ubicazione / Direzione (Tavola 20)", value="DIREZIONE SUD"
          )
          st.write("### Progressive Chilometriche (dal basso verso l'alto)")
          col1, col2 = st.columns(2)
          with col1:
            km_20_lavori = st.text_input(
                "KM Lavori in corso (In basso)", value="805+860"
            )
            km_20_limite110 = st.text_input("KM Limite 110 km/h", value="805+980")
            km_20_limite90 = st.text_input("KM Limite 90 km/h", value="806+100")
            km_20_restr1 = st.text_input("KM Restringimento 1", value="806+220")
            km_20_sorpasso = st.text_input("KM Divieto di sorpasso", value="806+340")
            km_20_restr2 = st.text_input("KM Restringimento 2", value="806+460")
            km_20_iniziotestata1 = st.text_input(
                "KM Inizio prima testata", value="806+556"
            )
          with col2:
            km_20_finetestata1 = st.text_input("KM Fine prima testata", value="806+700")
            km_20_restr3 = st.text_input("KM Restringimento 3", value="806+850")
            km_20_limite60 = st.text_input("KM Limite 60 km/h", value="806+950")
            km_20_restr4 = st.text_input("KM Restringimento 4", value="807+050")
            km_20_iniziotestata2 = st.text_input(
                "KM Inizio seconda testata", value="807+150"
            )
            km_20_finetestata2 = st.text_input(
                "KM Fine seconda testata", value="807+250"
            )
            km_20_vialibera = st.text_input(
                "KM Via libera / Fine (In alto)", value="807+300"
            )
        
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 20"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 20"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 20"
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
        
            dati_righe_20 = [
                (
                    "fig_383.jpg",
                    "FIG. II 383",
                    "LAVORI IN CORSO",
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
                    "fig_411a.jpg",
                    "FIG. II 411/a",
                    "SEGNALE DI CORSIA CHIUSA",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_052.jpg",
                    "FIG. II 052",
                    "DIVIETO DI SORPASSO PER VEICOLI DI MASSA > 3,5t",
                    "",
                    "SEGNALI DI DIVIETO",
                ),
                (
                    "fig_050_lim60.jpg",
                    "FIG. II 050",
                    "LIMITE DI VELOCITÀ 60 KM/H",
                    "",
                    "SEGNALI DI PRESCRIZIONE",
                ),
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
            for i, row_data in enumerate(dati_righe_20, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_20_img_{i}.jpg")
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
                "Abaco Excel Tavola 20 generato! (Immagini inserite:"
                f" {immagini_inserite}/{len(dati_righe_20)})"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Tavola 20",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_20.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 20 Personalizzata")
          img_path_20 = os.path.join(cartella_app, "TAVOLA 20.jpg")
          if os.path.exists(img_path_20):
            try:
              with Image.open(img_path_20) as immagine_originale:
                base_img_20 = immagine_originale.convert("RGB")
              width_20, height_20 = base_img_20.size
              draw_20 = ImageDraw.Draw(base_img_20)
              font_km = load_custom_font(18)
              font_ub_20 = load_custom_font(30)
              BORDO = 2
        
              # Mappatura coordinate progressive per la Tavola 20
              etichette_20 = [
                  (int(width_20 * 0.49), int(height_20 * 0.924), f"KM {km_20_lavori}"),
                  (
                      int(width_20 * 0.49),
                      int(height_20 * 0.884),
                      f"KM {km_20_limite110}",
                  ),
                  (int(width_20 * 0.49), int(height_20 * 0.841), f"KM {km_20_limite90}"),
                  (int(width_20 * 0.49), int(height_20 * 0.800), f"KM {km_20_restr1}"),
                  (int(width_20 * 0.49), int(height_20 * 0.757), f"KM {km_20_sorpasso}"),
                  (int(width_20 * 0.49), int(height_20 * 0.716), f"KM {km_20_restr2}"),
                  (
                      int(width_20 * 0.49),
                      int(height_20 * 0.6435),
                      f"KM {km_20_iniziotestata1}",
                  ),
                  (
                      int(width_20 * 0.49),
                      int(height_20 * 0.5050),
                      f"KM {km_20_finetestata1}",
                  ),
                  (int(width_20 * 0.49), int(height_20 * 0.468), f"KM {km_20_restr3}"),
                  (int(width_20 * 0.49), int(height_20 * 0.420), f"KM {km_20_limite60}"),
                  (int(width_20 * 0.49), int(height_20 * 0.381), f"KM {km_20_restr4}"),
                  (
                      int(width_20 * 0.49),
                      int(height_20 * 0.310),
                      f"KM {km_20_iniziotestata2}",
                  ),
                  (
                      int(width_20 * 0.49),
                      int(height_20 * 0.16),
                      f"KM {km_20_finetestata2}",
                  ),
                  (int(width_20 * 0.49), int(height_20 * 0.044), f"KM {km_20_vialibera}"),
              ]
        
              for x, y, testo in etichette_20:
                draw_20.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              draw_20.multiline_text(
                  (int(width_20 * 0.670), int(height_20 * 0.350)),
                  ubicazione_testo_20,
                  fill=(255, 0, 0),
                  font=font_ub_20,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_20 = io.BytesIO()
              base_img_20.save(img_io_20, format="JPEG", quality=100, subsampling=0)
              jpg_data_20 = img_io_20.getvalue()
              st.image(
                  base_img_20,
                  caption="Tavola 20 - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 20 in formato JPG",
                  data=jpg_data_20,
                  file_name="Tavola_20_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 20.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 20.jpg' non trovata nella cartella.")
        
        # ====================================================================
        # TAVOLA 21
        # ====================================================================
        elif "Tavola 21" in tavola_selezionata:
          st.subheader("📍 Dati e Progressive Chilometriche - Tavola 21")
          ubicazione_testo_21 = st.text_area(
              "Ubicazione / Direzione (Tavola 21)", value="DIREZIONE SUD"
          )
          st.write("### Progressive Chilometriche (dal basso verso l'alto)")
          col1, col2 = st.columns(2)
          with col1:
            km_21_lavori = st.text_input(
                "KM Lavori in corso (In basso)", value="805+860"
            )
            km_21_limite110 = st.text_input("KM Limite 110 km/h", value="805+980")
            km_21_limite90 = st.text_input("KM Limite 90 km/h", value="806+100")
            km_21_restr1 = st.text_input("KM Restringimento 1", value="806+220")
            km_21_sorpasso = st.text_input("KM Divieto di sorpasso", value="806+340")
            km_21_restr2 = st.text_input("KM Restringimento 2", value="806+460")
            km_21_iniziotestata1 = st.text_input(
                "KM Inizio prima testata", value="806+556"
            )
          with col2:
            km_21_finetestata1 = st.text_input("KM Fine prima testata", value="806+700")
            km_21_restr3 = st.text_input("KM Restringimento 3", value="806+850")
            km_21_limite60 = st.text_input("KM Limite 60 km/h", value="806+950")
            km_21_restr4 = st.text_input("KM Restringimento 4", value="807+050")
            km_21_iniziotestata2 = st.text_input(
                "KM Inizio seconda testata", value="807+150"
            )
            km_21_finetestata2 = st.text_input(
                "KM Fine seconda testata", value="807+250"
            )
            km_21_vialibera = st.text_input(
                "KM Via libera / Fine (In alto)", value="807+300"
            )
        
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 21"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 21"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 21"
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
        
            # Abaco Tavola 21 (con figura 082/b al posto di 082/a)
            dati_righe_21 = [
                (
                    "fig_383.jpg",
                    "FIG. II 383",
                    "LAVORI IN CORSO",
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
                    "fig_411a.jpg",
                    "FIG. II 411/a",
                    "SEGNALE DI CORSIA CHIUSA",
                    "",
                    "SEGNALI TEMPORANEI",
                ),
                (
                    "fig_052.jpg",
                    "FIG. II 052",
                    "DIVIETO DI SORPASSO PER VEICOLI DI MASSA > 3,5t",
                    "",
                    "SEGNALI DI DIVIETO",
                ),
                (
                    "fig_050_lim60.jpg",
                    "FIG. II 050",
                    "LIMITE DI VELOCITÀ 60 KM/H",
                    "",
                    "SEGNALI DI PRESCRIZIONE",
                ),
                (
                    "fig_082_b.jpg",
                    "FIG. II 082/b",
                    "PASSAGGIO OBBLIGATORIO A DESTRA",
                    "",
                    "SEGNALETICA DI OBBLIGO",
                ),
                (
                    "fig_048.jpg",
                    "FIG. II 048",
                    "DIVIETO DI SORPASSO",
                    "",
                    "SEGNALI DI DIVIETO",
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
            for i, row_data in enumerate(dati_righe_21, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_21_img_{i}.jpg")
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
                "Abaco Excel Tavola 21 generato! (Immagini inserite:"
                f" {immagini_inserite}/{len(dati_righe_21)})"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Tavola 21",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_21.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 21 Personalizzata")
          img_path_21 = os.path.join(cartella_app, "TAVOLA 21.jpg")
          if os.path.exists(img_path_21):
            try:
              with Image.open(img_path_21) as immagine_originale:
                base_img_21 = immagine_originale.convert("RGB")
              width_21, height_21 = base_img_21.size
              draw_21 = ImageDraw.Draw(base_img_21)
              font_km = load_custom_font(18)
              font_ub_21 = load_custom_font(32)
              BORDO = 2
        
              # Mappatura coordinate progressive per la Tavola 21 (identica alla 20)
              etichette_21 = [
                  (int(width_21 * 0.50), int(height_21 * 0.944), f"KM {km_21_lavori}"),
                  (
                      int(width_21 * 0.50),
                      int(height_21 * 0.900),
                      f"KM {km_21_limite110}",
                  ),
                  (int(width_21 * 0.50), int(height_21 * 0.858), f"KM {km_21_limite90}"),
                  (int(width_21 * 0.50), int(height_21 * 0.814), f"KM {km_21_restr1}"),
                  (int(width_21 * 0.50), int(height_21 * 0.774), f"KM {km_21_sorpasso}"),
                  (int(width_21 * 0.50), int(height_21 * 0.732), f"KM {km_21_restr2}"),
                  (
                      int(width_21 * 0.50),
                      int(height_21 * 0.660),
                      f"KM {km_21_iniziotestata1}",
                  ),
                  (
                      int(width_21 * 0.50),
                      int(height_21 * 0.520),
                      f"KM {km_21_finetestata1}",
                  ),
                  (int(width_21 * 0.50), int(height_21 * 0.480), f"KM {km_21_restr3}"),
                  (int(width_21 * 0.50), int(height_21 * 0.430), f"KM {km_21_limite60}"),
                  (int(width_21 * 0.50), int(height_21 * 0.392), f"KM {km_21_restr4}"),
                  (
                      int(width_21 * 0.50),
                      int(height_21 * 0.320),
                      f"KM {km_21_iniziotestata2}",
                  ),
                  (
                      int(width_21 * 0.50),
                      int(height_21 * 0.1655),
                      f"KM {km_21_finetestata2}",
                  ),
                  (int(width_21 * 0.50), int(height_21 * 0.052), f"KM {km_21_vialibera}"),
              ]
        
              for x, y, testo in etichette_21:
                draw_21.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              draw_21.multiline_text(
                  (int(width_21 * 0.673), int(height_21 * 0.350)),
                  ubicazione_testo_21,
                  fill=(255, 0, 0),
                  font=font_ub_21,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_21 = io.BytesIO()
              base_img_21.save(img_io_21, format="JPEG", quality=100, subsampling=0)
              jpg_data_21 = img_io_21.getvalue()
              st.image(
                  base_img_21,
                  caption="Tavola 21 - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 21 in formato JPG",
                  data=jpg_data_21,
                  file_name="Tavola_21_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 21.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 21.jpg' non trovata nella cartella.")
        
        # ====================================================================
        # TAVOLA 22
        # ====================================================================
        elif "Tavola 22" in tavola_selezionata:
          st.subheader("📍 Dati e Progressive Chilometriche - Tavola 22")
          ubicazione_testo_22 = st.text_area(
              "Ubicazione / Direzione (Tavola 22)", value="DIREZIONE SUD"
          )
          st.write("### Progressive Chilometriche (dal basso verso l'alto)")
          col1, col2 = st.columns(2)
          with col1:
            km_22_lavori = st.text_input(
                "KM Lavori in corso / preavviso (In basso)", value="805+860"
            )
            km_22_limite110 = st.text_input("KM Limite 110 km/h", value="806+100")
            km_22_restr = st.text_input("KM Restringimento", value="806+340")
          with col2:
            km_22_limite80 = st.text_input("KM Limite 80 km/h", value="806+760")
            km_22_obbligo = st.text_input("KM Passaggio obbligatorio", value="806+856")
            km_22_vialibera = st.text_input(
                "KM Via libera / Fine (In alto)", value="807+000"
            )
        
          if st.button("Genera Abaco Segnaletica (Excel) - Tavola 22"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Abaco Tavola 22"
            ws.merge_cells("A1:E1")
            titolo_principale = ws["A1"]
            titolo_principale.value = (
                "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 22"
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
        
            # Abaco Tavola 22 (con le figure richieste)
            dati_righe_22 = [
                (
                    "fig_383.jpg",
                    "FIG. II 383",
                    "LAVORI IN CORSO",
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
                    "fig050_lim80.jpg",
                    "FIG. II 050",
                    "LIMITE MASSIMO DI VELOCITÀ 80 KM/H",
                    "",
                    "SEGNALI DI DIVIETO",
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
            for i, row_data in enumerate(dati_righe_22, start=3):
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
                  temp_path = os.path.join("/tmp", f"abaco_22_img_{i}.jpg")
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
                "Abaco Excel Tavola 22 generato! (Immagini inserite:"
                f" {immagini_inserite}/{len(dati_righe_22)})"
            )
            st.download_button(
                label="📥 Scarica Excel Abaco Tavola 22",
                data=excel_data,
                file_name="Abaco_Segnaletica_Tavola_22.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        
          st.divider()
          st.subheader("🖼️ Anteprima Tavola 22 Personalizzata")
          img_path_22 = os.path.join(cartella_app, "TAVOLA 22.jpg")
          if os.path.exists(img_path_22):
            try:
              with Image.open(img_path_22) as immagine_originale:
                base_img_22 = immagine_originale.convert("RGB")
              width_22, height_22 = base_img_22.size
              draw_22 = ImageDraw.Draw(base_img_22)
              font_km = load_custom_font(25)
              font_ub_22 = load_custom_font(33)
              BORDO = 2
        
              # Mappatura coordinate progressive per la Tavola 22 (dal basso verso l'alto)
              etichette_22 = [
                  (int(width_22 * 0.49), int(height_22 * 0.835), f"KM {km_22_lavori}"),
                  (int(width_22 * 0.49), int(height_22 * 0.692), f"KM {km_22_limite110}"),
                  (int(width_22 * 0.49), int(height_22 * 0.5050), f"KM {km_22_restr}"),
                  (int(width_22 * 0.49), int(height_22 * 0.2880), f"KM {km_22_limite80}"),
                  (int(width_22 * 0.49), int(height_22 * 0.2310), f"KM {km_22_obbligo}"),
                  (int(width_22 * 0.49), int(height_22 * 0.073), f"KM {km_22_vialibera}"),
              ]
        
              for x, y, testo in etichette_22:
                draw_22.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )
        
              draw_22.multiline_text(
                  (int(width_22 * 0.66), int(height_22 * 0.400)),
                  ubicazione_testo_22,
                  fill=(255, 0, 0),
                  font=font_ub_22,
                  align="center",
                  spacing=6,
                  stroke_width=BORDO,
                  stroke_fill=(255, 255, 255),
              )
        
              img_io_22 = io.BytesIO()
              base_img_22.save(img_io_22, format="JPEG", quality=100, subsampling=0)
              jpg_data_22 = img_io_22.getvalue()
              st.image(
                  base_img_22,
                  caption="Tavola 22 - Configurazione completata",
                  use_container_width=True,
              )
              st.download_button(
                  label="📥 Scarica Tavola 22 in formato JPG",
                  data=jpg_data_22,
                  file_name="Tavola_22_Compilata.jpg",
                  mime="image/jpeg",
              )
            except Exception as errore:
              st.error("Errore generazione Tavola 22.")
              st.exception(errore)
          else:
            st.error("Immagine 'TAVOLA 22.jpg' non trovata nella cartella.")
        # ====================================================================
        # TAVOLA 23 - UNIFICATA (PARTE INFERIORE + ALTA)
        # ====================================================================
        elif "Tavola 23" in tavola_selezionata:
            st.subheader("📍 Dati e Progressive Chilometriche - Tavola 23 (Unificata)")
            ubicazione_testo_23 = st.text_area("Ubicazione / Direzione (Tavola 23)", value="DIREZIONE SUD")
            
            st.write("### 🛣️ Progressive Chilometriche (In ordine dal basso verso l'alto)")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Sezione Inferiore (Preavviso e Approccio)**")
                km_23inf_lavori = st.text_input("KM Lavori in corso (In basso)", value="805+560")
                km_23inf_limite110 = st.text_input("KM Limite 110 km/h", value="805+800")
                km_23inf_restr = st.text_input("KM Restringimento", value="806+040")
                km_23inf_limite80 = st.text_input("KM Limite 80 km/h", value="806+460")
                km_23inf_obbligo = st.text_input("KM Passaggio obbligatorio", value="806+556")
                km_23inf_sorpasso = st.text_input("KM Divieto di sorpasso", value="806+650")
            with col2:
                st.markdown("**Sezione Superiore (Chiusura e Rientro)**")
                km_23_restr1 = st.text_input("KM Restringimento 1", value="806+860")
                km_23_limite60 = st.text_input("KM Limite 60 km/h", value="807+100")
                km_23_restr2 = st.text_input("KM Restringimento 2", value="807+340")
                km_23_iniziotestata = st.text_input("KM Inizio testata", value="807+556")
                km_23_finetestata = st.text_input("KM Fine testata", value="807+700")
                km_23_vialibera = st.text_input("KM Via libera / Fine (In alto)", value="808+000")
        
            if st.button("Genera Abaco Segnaletica (Excel) - Tavola 23 Unificata"):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Abaco Tavola 23 Completa"
                ws.merge_cells("A1:E1")
                
                titolo_principale = ws["A1"]
                titolo_principale.value = "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 23"
                
                fill_verde_chiaro = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
                fill_verde_scuro = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                font_titolo = Font(name="Arial", size=11, bold=True, color="000000")
                font_header = Font(name="Arial", size=10, bold=True, color="000000")
                font_testo = Font(name="Arial", size=10, color="000000")
                bordo_sottile = Border(
                    left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"),
                    top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000")
                )
                
                titolo_principale.fill = fill_verde_chiaro
                titolo_principale.font = font_titolo
                titolo_principale.alignment = Alignment(horizontal="center", vertical="center")
                
                for col in range(1, 6):
                    ws.cell(row=1, column=col).border = bordo_sottile
                    
                headers = ["SEGNALE", "FIGURA", "DESCRIZIONE", "DETTAGLI", "SEGNALETICA"]
                for col_idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=2, column=col_idx)
                    cell.value = header
                    cell.fill = fill_verde_scuro
                    cell.font = font_header
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.border = bordo_sottile
                    
                # Manteniamo la struttura estesa della parte inferiore (Abaco completo richiesto)
                dati_righe_23completa = [
                    ("fig_383.jpg", "FIG. II 383", "LAVORI IN CORSO", "", "SEGNALI TEMPORANEI"),
                    ("fig050_lim110.jpg", "FIG. II 050", "LIMITE DI VELOCITÀ 110 KM/H", "", "SEGNALI DI PRESCRIZIONE"),
                    ("fig050_lim80.jpg", "FIG. II 050", "LIMITE MASSIMO DI VELOCITÀ 80 KM/H", "", "SEGNALI DI DIVIETO"),
                    ("fig_050_lim60.jpg", "FIG. II 050", "LIMITE DI VELOCITÀ 60 KM/H", "", "SEGNALI DI PRESCRIZIONE"),
                    ("fig_048.jpg", "FIG. II 048", "DIVIETO DI SORPASSO", "", "SEGNALI DI DIVIETO"),
                    ("fig_411a.jpg", "FIG. II 411/a", "SEGNALE DI CORSIA CHIUSA", "", "SEGNALI TEMPORANEI"),
                    ("fig_082_b.jpg", "FIG. II 082/b", "PASSAGGIO OBBLIGATORIO A DESTRA", "", "SEGNALETICA DI OBBLIGO"),
                    ("fig_400.jpg", "FIG. II 400", "SEGNALE MOBILE DI PREAVVISO CON FRECCIA", "", "SEGNALI TEMPORANEI"),
                    ("fig_401.jpg", "FIG. II 401", "VEICOLO DI PROTEZIONE CON PANNELLO", "", "SEGNALI TEMPORANEI"),
                    ("fig_070.jpg", "FIG. II 070", "VIA LIBERA / FINE CANTIERE", "", "SEGNALI DI PRESCRIZIONE")
                ]
                
                ws.column_dimensions["A"].width = 18
                ws.column_dimensions["B"].width = 16
                ws.column_dimensions["C"].width = 45
                ws.column_dimensions["D"].width = 20
                ws.column_dimensions["E"].width = 25
                ws.row_dimensions[1].height = 30
                ws.row_dimensions[2].height = 25
                
                immagini_inserite = 0
                for i, row_data in enumerate(dati_righe_23completa, start=3):
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
                    
                    nomi_da_cercare = [img_nome, img_nome.replace(".jpg", ".png"), img_nome.replace(".png", ".jpg")]
                    percorso_trovato = None
                    for nome in nomi_da_cercare:
                        tentativo = os.path.join(dir_immagini, nome)
                        if os.path.exists(tentativo):
                            percorso_trovato = tentativo
                            break
                            
                    if percorso_trovato:
                        try:
                            img_pil = Image.open(percorso_trovato).convert("RGB")
                            temp_path = os.path.join("/tmp", f"abaco_23completa_{i}.jpg")
                            img_pil.save(temp_path, "JPEG")
                            img_obj = OpenpyxlImage(temp_path)
                            marker_from = AnchorMarker(col=0, colOff=150000, row=i-1, rowOff=40000)
                            marker_to = AnchorMarker(col=1, colOff=-150000, row=i, rowOff=-40000)
                            img_obj.anchor = TwoCellAnchor(editAs="twoCell", _from=marker_from, to=marker_to)
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
                st.success(f"Abaco Excel Unificato generato! (Immagini inserite: {immagini_inserite}/{len(dati_righe_23completa)})")
                st.download_button(
                    label="📥 Scarica Excel Abaco Segnaletica Tavola 23 Completa",
                    data=excel_data,
                    file_name="Abaco_Segnaletica_Tavola_23_Completa.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            st.divider()
            st.subheader("🖼️ Anteprima Tavola 23 Personalizzata (Unificata)")
            
            # Renderizziamo graficamente entrambe le componenti visive all'interno della stessa pagina dell'app
            col_tav1, col_tav2 = st.columns(2)
            
            # --- GESTIONE PARTE INFERIORE ---
            with col_tav1:
                st.markdown("#### 🔽 Parte Inferiore (Approccio)")
                img_path_23inf = os.path.join(cartella_app, "TAVOLA 23 - PARTE INFERIORE.jpg")
                if os.path.exists(img_path_23inf):
                    try:
                        with Image.open(img_path_23inf) as immagine_originale:
                            base_img_23inf = immagine_originale.convert("RGB")
                        width_23inf, height_23inf = base_img_23inf.size
                        draw_23inf = ImageDraw.Draw(base_img_23inf)
                        font_km = load_custom_font(18)
                        font_ub_23inf = load_custom_font(25)
                        BORDO = 2
                        
                        etichette_23inf = [
                            (int(width_23inf * 0.66), int(height_23inf * 0.847), f"KM {km_23inf_lavori}"),
                            (int(width_23inf * 0.66), int(height_23inf * 0.690), f"KM {km_23inf_limite110}"),
                            (int(width_23inf * 0.66), int(height_23inf * 0.490), f"KM {km_23inf_restr}"),
                            (int(width_23inf * 0.66), int(height_23inf * 0.260), f"KM {km_23inf_limite80}"),
                            (int(width_23inf * 0.66), int(height_23inf * 0.2000), f"KM {km_23inf_obbligo}"),
                            (int(width_23inf * 0.66), int(height_23inf * 0.0850), f"KM {km_23inf_sorpasso}")
                        ]
                        for x, y, testo in etichette_23inf:
                            draw_23inf.text((x, y), testo, fill=(255, 0, 0), font=font_km, stroke_width=BORDO, stroke_fill=(255, 255, 255))
                            
                        draw_23inf.multiline_text(
                            (int(width_23inf * 0.68), int(height_23inf * 0.380)), 
                            ubicazione_testo_23, 
                            fill=(255, 0, 0), 
                            font=font_ub_23inf, 
                            align="center", 
                            spacing=6, 
                            stroke_width=BORDO, 
                            stroke_fill=(255, 255, 255)
                        )
                        
                        img_io_23inf = io.BytesIO()
                        base_img_23inf.save(img_io_23inf, format="JPEG", quality=100, subsampling=0)
                        st.image(base_img_23inf, caption="Tavola 23 - Parte Inferiore Compilata", use_container_width=True)
                        st.download_button(label="📥 Scarica JPG Parte Inferiore", data=img_io_23inf.getvalue(), file_name="Tavola_23_Parte_Inferiore.jpg", mime="image/jpeg")
                    except Exception as e:
                        st.error(f"Errore grafica parte inferiore: {e}")
                else:
                    st.error("Immagine 'TAVOLA 23 - PARTE INFERIORE.jpg' non trovata.")
        
            # --- GESTIONE PARTE SUPERIORE ---
            with col_tav2:
                st.markdown("#### 🔼 Parte Superiore (Chiusura)")
                img_path_23 = os.path.join(cartella_app, "TAVOLA 23.jpg")
                if os.path.exists(img_path_23):
                    try:
                        with Image.open(img_path_23) as immagine_originale:
                            base_img_23 = immagine_originale.convert("RGB")
                        width_23, height_23 = base_img_23.size
                        draw_23 = ImageDraw.Draw(base_img_23)
                        font_km = load_custom_font(23)
                        font_ub_23 = load_custom_font(31)
                        BORDO = 2
                        
                        etichette_23 = [
                            (int(width_23 * 0.54), int(height_23 * 0.835), f"KM {km_23_restr1}"),
                            (int(width_23 * 0.54), int(height_23 * 0.727), f"KM {km_23_limite60}"),
                            (int(width_23 * 0.54), int(height_23 * 0.6350), f"KM {km_23_restr2}"),
                            (int(width_23 * 0.54), int(height_23 * 0.528), f"KM {km_23_iniziotestata}"),
                            (int(width_23 * 0.54), int(height_23 * 0.415), f"KM {km_23_finetestata}"),
                            (int(width_23 * 0.54), int(height_23 * 0.215), f"KM {km_23_vialibera}")
                        ]
                        for x, y, testo in etichette_23:
                            draw_23.text((x, y), testo, fill=(255, 0, 0), font=font_km, stroke_width=BORDO, stroke_fill=(255, 255, 255))
                            
                        draw_23.multiline_text(
                            (int(width_23 * 0.64), int(height_23 * 0.2900)), 
                            ubicazione_testo_23, 
                            fill=(255, 0, 0), 
                            font=font_ub_23, 
                            align="center", 
                            spacing=6, 
                            stroke_width=BORDO, 
                            stroke_fill=(255, 255, 255)
                        )
                        
                        img_io_23 = io.BytesIO()
                        base_img_23.save(img_io_23, format="JPEG", quality=100, subsampling=0)
                        st.image(base_img_23, caption="Tavola 23 - Parte Superiore Compilata", use_container_width=True)
                        st.download_button(label="📥 Scarica JPG Parte Superiore", data=img_io_23.getvalue(), file_name="Tavola_23_Parte_Superiore.jpg", mime="image/jpeg")
                    except Exception as e:
                        st.error(f"Errore grafica parte superiore: {e}")
                else:
                    st.error("Immagine 'TAVOLA 23.jpg' non trovata.")
        
        
        # ====================================================================
        # TAVOLA 24 - UNIFICATA (PARTE BASSA + PARTE ALTA)
        # ====================================================================
        elif "Tavola 24" in tavola_selezionata:
            st.subheader("📍 Dati e Progressive Chilometriche - Tavola 24 (Unificata)")
            ubicazione_testo_24 = st.text_area(
                "Ubicazione / Direzione (Tavola 24)", value="DIREZIONE SUD"
            )
        
            st.write(
                "### 🛣️ Progressive Chilometriche (In ordine dal basso verso l'alto)"
            )
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Sezione Inferiore (Parte Bassa)**")
                km_24inf_lavori = st.text_input(
                    "KM Lavori in corso (In basso)", value="805+560"
                )
                km_24inf_limite110 = st.text_input(
                    "KM Limite 110 km/h", value="805+800"
                )
                km_24inf_restr = st.text_input("KM Restringimento", value="806+040")
                km_24inf_limite80 = st.text_input("KM Limite 80 km/h", value="806+280")
                km_24inf_limite60 = st.text_input("KM Limite 60 km/h", value="806+400")
                km_24inf_iniziotestata = st.text_input(
                    "KM Inizio testata", value="806+508"
                )
                km_24inf_finetestata = st.text_input("KM Fine testata", value="806+688")
            with col2:
                st.markdown("**Sezione Superiore (Parte Alta)**")
                km_24_restr1 = st.text_input("KM Restringimento 1", value="806+860")
                km_24_limite60 = st.text_input("KM Limite 60 km/h", value="806+956")
                km_24_restr2 = st.text_input("KM Restringimento 2", value="807+052")
                km_24_iniziotestata = st.text_input(
                    "KM Inizio testata", value="807+148"
                )
                km_24_finetestata = st.text_input("KM Fine testata", value="807+256")
                km_24_vialibera = st.text_input(
                    "KM Via libera / Fine (In alto)", value="807+436"
                )
        
            if st.button("Genera Abaco Segnaletica (Excel) - Tavola 24 Unificata"):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Abaco Tavola 24 Completa"
                ws.merge_cells("A1:E1")
        
                titolo_principale = ws["A1"]
                titolo_principale.value = (
                    "TABELLA RIEPILOGATIVA SEGNALETICA DI CANTIERE - TAVOLA 24"
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
        
                headers = [
                    "SEGNALE",
                    "FIGURA",
                    "DESCRIZIONE",
                    "DETTAGLI",
                    "SEGNALETICA",
                ]
                for col_idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=2, column=col_idx)
                    cell.value = header
                    cell.fill = fill_verde_scuro
                    cell.font = font_header
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True
                    )
                    cell.border = bordo_sottile
        
                # Elenco figure per l'Abaco della Tavola 24
               # Elenco figure corretto per l'Abaco della Tavola 24
                dati_righe_24completa = [
                    (
                        "fig_383.jpg",
                        "FIG. II 383",
                        "PREAVVISO LAVORI",
                        "",
                        "SEGNALI TEMPORANEI",
                    ),
                    (
                        "fig050_lim110.jpg",
                        "FIG. II 050",
                        "LIMITE MASSIMO DI VELOCITÀ 110 KM/H",
                        "",
                        "SEGNALI DI PRESCRIZIONE",
                    ),
                    (
                        "fig_400.jpg",
                        "FIG. II 400",
                        "SEGNALE MOBILE DI PREAVVISO",
                        "",
                        "SEGNALI TEMPORANEI",
                    ),
                    (
                        "fig050_lim80.jpg",
                        "FIG. II 050",
                        "LIMITE MASSIMO DI VELOCITÀ 80 KM/H",
                        "",
                        "SEGNALI DI DIVIETO",
                    ),
                    (
                        "fig_050_lim60.jpg",
                        "FIG. II 050",
                        "LIMITE MASSIMO DI VELOCITÀ 60 KM/H",
                        "",
                        "SEGNALI DI PRESCRIZIONE",
                    ),
                    (
                        "fig_052.jpg",
                        "FIG. II 052",
                        "DIVIETO DI SORPASSO PER GLI AUTOCARRI",
                        "",
                        "SEGNALI DI DIVIETO",
                    ),
                    (
                        "fig_082_b.jpg",
                        "FIG. II 082/b",
                        "PASSAGGIO OBBLIGATORIO A DESTRA",
                        "",
                        "SEGNALETICA DI OBBLIGO",
                    ),
                    (
                        "fig_411a.jpg",
                        "FIG. II 411/a",
                        "SEGNALE DI CORSIA CHIUSA",
                        "",
                        "SEGNALI TEMPORANEI",
                    ),
                    (
                        "fig_048.jpg",
                        "FIG. II 048",
                        "DIVIETO DI SORPASSO",
                        "",
                        "SEGNALI DI DIVIETO",
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
                for i, row_data in enumerate(dati_righe_24completa, start=3):
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
                            temp_path = os.path.join(
                                "/tmp", f"abaco_24completa_{i}.jpg"
                            )
                            img_pil.save(temp_path, "JPEG")
                            img_obj = OpenpyxlImage(temp_path)
                            marker_from = AnchorMarker(
                                col=0, colOff=150000, row=i - 1, rowOff=40000
                            )
                            marker_to = AnchorMarker(
                                col=1, colOff=-150000, row=i, rowOff=-40000
                            )
                            img_obj.anchor = TwoCellAnchor(
                                editAs="twoCell", _from=marker_from, to=marker_to
                            )
                            ws.add_image(img_obj)
                            immagini_inserite += 1
                        except Exception as e:
                            st.warning(
                                f"Errore caricamento immagine per {figura}: {e}"
                            )
                    else:
                        c_segnale.value = "[Img non trovata]"
                        c_segnale.font = Font(name="Arial", size=8, color="FF0000")
                        c_segnale.alignment = Alignment(
                            horizontal="center", vertical="center"
                        )
        
                output = io.BytesIO()
                wb.save(output)
                excel_data = output.getvalue()
                st.success(
                    f"Abaco Excel Unificato generato! (Immagini inserite: {immagini_inserite}/{len(dati_righe_24completa)})"
                )
                st.download_button(
                    label="📥 Scarica Excel Abaco Segnaletica Tavola 24 Completa",
                    data=excel_data,
                    file_name="Abaco_Segnaletica_Tavola_24_Completa.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        
            st.divider()
            st.subheader("🖼️ Anteprima Tavola 24 Personalizzata (Unificata)")
        
            col_tav1, col_tav2 = st.columns(2)
        
            # --- GESTIONE PARTE BASSA ---
            with col_tav1:
                st.markdown("#### 🔽 Parte Bassa")
                img_path_24inf = os.path.join(cartella_app, "TAVOLA 24 PARTE BASSA.jpg")
                if os.path.exists(img_path_24inf):
                    try:
                        with Image.open(img_path_24inf) as immagine_originale:
                            base_img_24inf = immagine_originale.convert("RGB")
                        width_24inf, height_24inf = base_img_24inf.size
                        draw_24inf = ImageDraw.Draw(base_img_24inf)
                        font_km = load_custom_font(25)
                        font_ub_24inf = load_custom_font(33)
                        BORDO = 2
        
                        # Coordinate percentuali adattate a "TAVOLA 24 PARTE BASSA.jpg" (dal basso verso l'alto)
                        etichette_24inf = [
                            (
                                int(width_24inf * 0.65),
                                int(height_24inf * 0.710),
                                f"KM {km_24inf_lavori}",
                            ),
                            (
                                int(width_24inf * 0.65),
                                int(height_24inf * 0.610),
                                f"KM {km_24inf_limite110}",
                            ),
                            (
                                int(width_24inf * 0.65),
                                int(height_24inf * 0.510),
                                f"KM {km_24inf_restr}",
                            ),
                            (
                                int(width_24inf * 0.65),
                                int(height_24inf * 0.408),
                                f"KM {km_24inf_limite80}",
                            ),
                            (
                                int(width_24inf * 0.65),
                                int(height_24inf * 0.3370),
                                f"KM {km_24inf_limite60}",
                            ),
                            (
                                int(width_24inf * 0.65),
                                int(height_24inf * 0.2680),
                                f"KM {km_24inf_iniziotestata}",
                            ),
                            (
                                int(width_24inf * 0.65),
                                int(height_24inf * 0.142),
                                f"KM {km_24inf_finetestata}",
                            ),
                        ]
                        for x, y, testo in etichette_24inf:
                            draw_24inf.text(
                                (x, y),
                                testo,
                                fill=(255, 0, 0),
                                font=font_km,
                                stroke_width=BORDO,
                                stroke_fill=(255, 255, 255),
                            )
        
                        draw_24inf.multiline_text(
                            (int(width_24inf * 0.66), int(height_24inf * 0.060)),
                            ubicazione_testo_24,
                            fill=(255, 0, 0),
                            font=font_ub_24inf,
                            align="center",
                            spacing=6,
                            stroke_width=BORDO,
                            stroke_fill=(255, 255, 255),
                        )
        
                        img_io_24inf = io.BytesIO()
                        base_img_24inf.save(
                            img_io_24inf, format="JPEG", quality=100, subsampling=0
                        )
                        st.image(
                            base_img_24inf,
                            caption="Tavola 24 - Parte Bassa Compilata",
                            use_container_width=True,
                        )
                        st.download_button(
                            label="📥 Scarica JPG Parte Bassa",
                            data=img_io_24inf.getvalue(),
                            file_name="Tavola_24_Parte_Bassa.jpg",
                            mime="image/jpeg",
                        )
                    except Exception as e:
                        st.error(f"Errore grafica parte bassa: {e}")
                else:
                    st.error("Immagine 'TAVOLA 24 PARTE BASSA.jpg' non trovata.")
        
            # --- GESTIONE PARTE ALTA ---
            with col_tav2:
                st.markdown("#### 🔼 Parte Alta")
                img_path_24alta = os.path.join(cartella_app, "TAVOLA 24 PARTE ALTA.jpg")
                if os.path.exists(img_path_24alta):
                    try:
                        with Image.open(img_path_24alta) as immagine_originale:
                            base_img_24alta = immagine_originale.convert("RGB")
                        width_24alta, height_24alta = base_img_24alta.size
                        draw_24alta = ImageDraw.Draw(base_img_24alta)
                        font_km = load_custom_font(28)
                        font_ub_24alta = load_custom_font(35)
                        BORDO = 2
        
                        # Coordinate percentuali adattate a "TAVOLA 24 PARTE ALTA.jpg" (dal basso verso l'alto)
                        etichette_24alta = [
                            (
                                int(width_24alta * 0.52),
                                int(height_24alta * 0.8570),
                                f"KM {km_24_restr1}",
                            ),
                            (
                                int(width_24alta * 0.52),
                                int(height_24alta * 0.786),
                                f"KM {km_24_limite60}",
                            ),
                            (
                                int(width_24alta * 0.52),
                                int(height_24alta * 0.7200),
                                f"KM {km_24_restr2}",
                            ),
                            (
                                int(width_24alta * 0.52),
                                int(height_24alta * 0.633),
                                f"KM {km_24_iniziotestata}",
                            ),
                            (
                                int(width_24alta * 0.52),
                                int(height_24alta * 0.5140),
                                f"KM {km_24_finetestata}",
                            ),
                            (
                                int(width_24alta * 0.52),
                                int(height_24alta * 0.312),
                                f"KM {km_24_vialibera}",
                            ),
                        ]
                        for x, y, testo in etichette_24alta:
                            draw_24alta.text(
                                (x, y),
                                testo,
                                fill=(255, 0, 0),
                                font=font_km,
                                stroke_width=BORDO,
                                stroke_fill=(255, 255, 255),
                            )
        
                        draw_24alta.multiline_text(
                            (int(width_24alta * 0.48), int(height_24alta * 0.215)),
                            ubicazione_testo_24,
                            fill=(255, 0, 0),
                            font=font_ub_24alta,
                            align="center",
                            spacing=6,
                            stroke_width=BORDO,
                            stroke_fill=(255, 255, 255),
                        )
        
                        img_io_24alta = io.BytesIO()
                        base_img_24alta.save(
                            img_io_24alta, format="JPEG", quality=100, subsampling=0
                        )
                        st.image(
                            base_img_24alta,
                            caption="Tavola 24 - Parte Alta Compilata",
                            use_container_width=True,
                        )
                        st.download_button(
                            label="📥 Scarica JPG Parte Alta",
                            data=img_io_24alta.getvalue(),
                            file_name="Tavola_24_Parte_Alta.jpg",
                            mime="image/jpeg",
                        )
                    except Exception as e:
                        st.error(f"Errore grafica parte alta: {e}")
                else:
                    st.error("Immagine 'TAVOLA 24 PARTE ALTA.jpg' non trovata.")
        
        
    

# ====================================================================
# SEZIONE 2: PROGRAMMA LAVORI MENSILE (CREAZIONE DA ZERO)
# ====================================================================
elif sezione_app == "Programma Lavori Mensile":
    st.title("📅 Programma Lavori Mensile")
    st.caption("Compila le lavorazioni, personalizza il sottotitolo e genera il file Excel formattato.")

    mesi_dict = {
        1: "GENNAIO", 2: "FEBBRAIO", 3: "MARZO", 4: "APRILE",
        5: "MAGGIO", 6: "GIUGNO", 7: "LUGLIO", 8: "AGOSTO",
        9: "SETTEMBRE", 10: "OTTOBRE", 11: "NOVEMBRE", 12: "DICEMBRE"
    }

    col1, col2 = st.columns(2)
    with col1:
        selected_month_num = st.selectbox(
            "Seleziona Mese", options=list(mesi_dict.keys()), format_func=lambda x: mesi_dict[x], index=10
        )
    with col2:
        selected_year = st.number_input("Anno", min_value=2024, max_value=2030, value=2028)

    # Campo di testo modificabile per il sottotitolo
    testo_sottotitolo = st.text_input(
        "Modifica Sottotitolo / Descrizione Cantiere:",
        value="S.S. n° 16 “Adriatica”. Lavori di manutenzione straordinaria per la riqualificazione dell’itinerario Bari-Brindisi-Lecce.  Lotto 1 – Stralcio 2 "
    )

    st.markdown("### 📝 Inserisci le Lavorazioni")
    st.info("Puoi aggiungere, modificare o eliminare le righe delle lavorazioni secondo le tue esigenze.")

    default_data = [
        {"Descrizione Lavori": "DEMOLIZIONE INTEGRALE DI STRUTTURE", "Area di Intervento": "OM11 - OM12", "Impresa": "TESORO SRL"},
        {"Descrizione Lavori": "SCAVO A FONDAZIONE", "Area di Intervento": "OM11 - OM12", "Impresa": "TESORO SRL"},
        {"Descrizione Lavori": "CONFERIMENTO A DISCARICA", "Area di Intervento": "OM11 - OM12", "Impresa": "TESORO SRL"},
        {"Descrizione Lavori": "CONGLOMERATO CEMENTIZIO MAGRONE", "Area di Intervento": "OM11 - OM12", "Impresa": "TESORO SRL"},
        {"Descrizione Lavori": "ACCIAIO IN BARRE TONDE B450C", "Area di Intervento": "OM11 - OM12", "Impresa": "TESORO SRL"},
        {"Descrizione Lavori": "CALCESTRUZZI STRUTTURALI", "Area di Intervento": "OM11 - OM12", "Impresa": "TESORO SRL"},
        {"Descrizione Lavori": "CASSEFORMI", "Area di Intervento": "OM11 - OM12", "Impresa": "TESORO SRL"},
    ]

    df_lavori = pd.DataFrame(default_data)
    edited_df = st.data_editor(df_lavori, num_rows="dynamic", use_container_width=True)

    if st.button("🚀 Genera File Excel Programma Lavori"):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = mesi_dict[selected_month_num]

        # Griglie disattivate sullo sfondo per pulizia visiva identica all'originale
        ws.views.sheetView[0].showGridLines = False

        # Font e stili
        font_title = Font(name="Franklin Gothic Medium", size=48, bold=False)
        font_year = Font(name="Franklin Gothic Medium", size=48, bold=True)
        font_subtitle = Font(name="Franklin Gothic Book", size=14, bold=True)
        font_month = Font(name="Franklin Gothic Book", size=11, bold=True)
        font_header = Font(name="Franklin Gothic Book", size=9, bold=True, color="FFFFFF" ) # Testo bianco per risaltare sul fondo scuro
        font_body = Font(name="Franklin Gothic Book", size=9, bold=True)
        font_footer = Font(name="Franklin Gothic Book", size=11, bold=True)

        fill_header_dark = PatternFill(start_color="595959", end_color="595959", fill_type="solid") # Grigio scuro per tutte le intestazioni
        thin_side = Side(style='thin', color='000000')
        border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        # Altezze righe
        ws.row_dimensions[1].height = 4.9
        ws.row_dimensions[2].height = 65.0
        ws.row_dimensions[3].height = 45.0
        ws.row_dimensions[4].height = 22.0
        ws.row_dimensions[5].height = 19.9
        ws.row_dimensions[6].height = 27.75

        # Calcolo giorni del mese selezionato
        num_days = calendar.monthrange(selected_year, selected_month_num)[1]
        last_calendar_col_idx = 4 + num_days
        importo_col_idx = last_calendar_col_idx + 1
        importo_col_letter = get_column_letter(importo_col_idx)

        # 1. Titolo principale in B2 (dimensione 48)
        col_fine_titolo = max(5, last_calendar_col_idx - 4)
        ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=col_fine_titolo)
        ws['B2'] = "PROGRAMMA LAVORI MENSILE"
        ws['B2'].font = font_title
        ws['B2'].alignment = Alignment(horizontal="left", vertical="center")

        # 2. Anno in alto a destra (dimensione 48, unito)
        anno_start_col = col_fine_titolo + 1
        anno_end_col = importo_col_idx
        
        cell_anno = ws.cell(row=2, column=anno_start_col)
        cell_anno.value = selected_year
        cell_anno.font = font_year
        cell_anno.alignment = Alignment(horizontal="right", vertical="center")
        
        ws.merge_cells(start_row=2, start_column=anno_start_col, end_row=2, end_column=anno_end_col)

        # 3. Sottotitolo in B3 (dimensione 14, modificabile)
        ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=importo_col_idx)
        ws['B3'] = testo_sottotitolo
        ws['B3'].font = font_subtitle
        ws['B3'].alignment = Alignment(horizontal="center", vertical="center")

        # 4. Mese centrato in E4 sopra il blocco giorni
        ws.merge_cells(start_row=4, start_column=5, end_row=4, end_column=last_calendar_col_idx)
        ws['E4'] = mesi_dict[selected_month_num]
        ws['E4'].font = font_month
        ws['E4'].alignment = Alignment(horizontal="center", vertical="center")

        # 5. Intestazioni principali con fondo scuro e testo bianco (righe 5 e 6)
        for col_let, col_idx, testo in [('B', 2, "DESCRIZIONE LAVORI"), ('C', 3, "AREA DI INTERVENTO"), ('D', 4, "IMPRESA")]:
            ws.merge_cells(start_row=5, start_column=col_idx, end_row=6, end_column=col_idx)
            cell = ws[f'{col_let}5']
            cell.value = testo
            cell.font = font_header
            cell.fill = fill_header_dark
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws[f'{col_let}5'].border = border_all
            ws[f'{col_let}6'].border = border_all

        # 6. Primo giorno in E5 (fondo scuro) e giorno settimana in E6 (fondo scuro)
        first_date = datetime.date(selected_year, selected_month_num, 1)
        ws['E5'] = first_date
        ws['E5'].number_format = 'd'
        ws['E5'].font = font_header
        ws['E5'].fill = fill_header_dark
        ws['E5'].alignment = Alignment(horizontal="center", vertical="center")
        ws['E5'].border = border_all

        ws['E6'] = '=CHOOSE(WEEKDAY(E5, 2), "Lu", "Ma", "Me", "Gi", "Ve", "Sa", "Do")'
        ws['E6'].font = font_header
        ws['E6'].fill = fill_header_dark
        ws['E6'].alignment = Alignment(horizontal="center", vertical="center")
        ws['E6'].border = border_all

        # Colonne successive dei giorni (tutte con fondo scuro)
        for col_idx in range(1, num_days):
            col_letter = get_column_letter(5 + col_idx)
            prev_col_letter = get_column_letter(4 + col_idx)
            
            cell_date = ws[f'{col_letter}5']
            cell_date.value = f'=IF(MONTH($E$5+{col_idx})=MONTH($E$5), {prev_col_letter}5+1, "")'
            cell_date.number_format = 'd'
            cell_date.font = font_header
            cell_date.fill = fill_header_dark
            cell_date.alignment = Alignment(horizontal="center", vertical="center")
            cell_date.border = border_all

            cell_wd = ws[f'{col_letter}6']
            cell_wd.value = f'=IF({col_letter}5<>"", CHOOSE(WEEKDAY({col_letter}5, 2), "Lu", "Ma", "Me", "Gi", "Ve", "Sa", "Do"), "")'
            cell_wd.font = font_header
            cell_wd.fill = fill_header_dark
            cell_wd.alignment = Alignment(horizontal="center", vertical="center")
            cell_wd.border = border_all

        # 7. Colonna Importo Presunto Netto con fondo scuro (righe 5 e 6)
        ws.merge_cells(start_row=5, start_column=importo_col_idx, end_row=6, end_column=importo_col_idx)
        cell_imp_header = ws.cell(row=5, column=importo_col_idx)
        cell_imp_header.value = "IMPORTO PRESUNTO\nNETTO"
        cell_imp_header.font = font_header
        cell_imp_header.fill = fill_header_dark
        cell_imp_header.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(row=5, column=importo_col_idx).border = border_all
        ws.cell(row=6, column=importo_col_idx).border = border_all

        # 8. Inserimento righe lavorazioni dinamiche (con wrap_text attivo sulla colonna B)
        start_row = 7
        row_heights_default = [48.0, 47.25, 47.25, 48.75, 48.75, 48.75, 48.75]

        for idx, row in edited_df.iterrows():
            r = start_row + idx
            ws.row_dimensions[r].height = row_heights_default[idx % len(row_heights_default)]
            
            c_desc = ws[f'B{r}']
            c_desc.value = row.get("Descrizione Lavori", "")
            c_desc.font = font_body
            c_desc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True) # Testo a capo abilitato
            c_desc.border = border_all
            
            c_area = ws[f'C{r}']
            c_area.value = row.get("Area di Intervento", "")
            c_area.font = font_body
            c_area.alignment = Alignment(horizontal="center", vertical="center")
            c_area.border = border_all
            
            c_imp = ws[f'D{r}']
            c_imp.value = row.get("Impresa", "TESORO SRL")
            c_imp.font = font_body
            c_imp.alignment = Alignment(horizontal="left", vertical="center")
            c_imp.border = border_all

            # Bordi celle calendario e importo per la riga
            for c_idx in range(5, importo_col_idx + 1):
                col_letter = get_column_letter(c_idx)
                cell_cal = ws[f'{col_letter}{r}']
                cell_cal.font = font_body
                cell_cal.border = border_all

        # 9. Righe finali Personale e Mezzi Stimati
        pers_row = start_row + len(edited_df)
        mezzi_row = pers_row + 1

        ws.row_dimensions[pers_row].height = 18.95
        ws.row_dimensions[mezzi_row].height = 18.95

        for r_fin, label_testo in [(pers_row, "PERSONALE STIMATO"), (mezzi_row, "MEZZI STIMATI")]:
            ws[f'B{r_fin}'] = label_testo
            ws[f'B{r_fin}'].font = font_footer
            ws.merge_cells(start_row=r_fin, start_column=2, end_row=r_fin, end_column=4)
            
            for c_idx in range(2, importo_col_idx + 1):
                col_let = get_column_letter(c_idx)
                cell_f = ws[f'{col_let}{r_fin}']
                cell_f.font = font_footer
                cell_f.border = border_all

        # Larghezze colonne ottimizzate
        ws.column_dimensions['B'].width = 27.78
        ws.column_dimensions['C'].width = 19.44
        ws.column_dimensions['D'].width = 13.0
        for col_idx in range(num_days):
            col_letter = get_column_letter(5 + col_idx)
            ws.column_dimensions[col_letter].width = 4.77
        ws.column_dimensions[importo_col_letter].width = 16.0

        # Esportazione in memoria
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("File Excel generato con successo: sottotitolo modificabile, testo a capo in colonna B e intestazioni scure!")
        st.download_button(
            label="📥 Scarica Programma Lavori Excel",
            data=output,
            file_name=f"PROGRAMMA_MENSILE_{mesi_dict[selected_month_num]}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
