import io
import os
import urllib.request

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont


# ==================================================
# CONFIGURAZIONE STREAMLIT
# ==================================================
st.set_page_config(
    page_title="Gestione Cantieri Stradali - Abaco & Tavole",
    layout="wide",
)

st.title("🚧 Gestione Cantieri Stradali e Segnaletica")

st.write(
    "Inserisci le progressive chilometriche: il programma "
    "genererà l'Abaco Excel e la Tavola JPG personalizzata."
)


# ==================================================
# CARICAMENTO FONT: COMPATIBILE CON STREAMLIT CLOUD
# ==================================================
@st.cache_resource
def load_custom_font(size):
    """
    Cerca un font TrueType in grassetto.
    Se non lo trova, prova a scaricarlo.
    Come ultima alternativa usa il font incluso in Pillow.
    """

    possibili_font = [
        "DejaVuSans-Bold.ttf",
        os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "DejaVuSans-Bold.ttf",
        ),
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/arialbd.ttf",
        "/Library/Fonts/Arial Bold.ttf",
    ]

    # 1. Cerca un font già presente
    for percorso in possibili_font:
        if os.path.isfile(percorso):
            try:
                return ImageFont.truetype(percorso, size)
            except (OSError, ValueError):
                continue

    # 2. Prova a scaricare il font in una cartella temporanea
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

    # 3. Ultima alternativa: font Pillow
    # Compatibilità con versioni differenti di Pillow
    try:
        font = ImageFont.load_default(size=size)
    except TypeError:
        font = ImageFont.load_default()

    return font


# ==================================================
# SELEZIONE DELLA TAVOLA
# ==================================================
tavola_selezionata = st.selectbox(
    "Seleziona la Tavola di Cantiere",
    [
        "Tavola 39 - Cantiere mobile carreggiata a due corsie"
    ],
)


if "39" in tavola_selezionata:

    st.subheader(
        "📍 Dati e Progressive Chilometriche - Tavola 39"
    )

    st.markdown(
        "*Direzione Sud, Viadotto Torre Guaceto 1 - "
        "Chiusura corsia di destra*"
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
    # GENERAZIONE FILE EXCEL (ABACO)
    # ==================================================
    if st.button("Genera Abaco Segnaletica (Excel)"):

        dati_abaco = {
            "SEGNALE": [
                "Preavviso Cantiere Mobile",
                "Segnale Mobile di Protezione",
                "Via Libera / Fine Cantiere",
            ],
            "FIGURA": [
                "FIG. 400",
                "FIG. 401",
                "FIG. 070",
            ],
            "DESCRIZIONE": [
                "SEGNALE MOBILE DI PREAVVISO",
                "SEGNALE MOBILE DI PROTEZIONE",
                "VIA LIBERA",
            ],
            "PROGRESSIVA CHILOMETRICA": [
                km_preavviso,
                km_inizio,
                km_fine,
            ],
            "SEGNALETICA": [
                "SEGNALI TEMPORANEI",
                "SEGNALI TEMPORANEI",
                "SEGNALI DI PRESCRIZIONE",
            ],
        }

        df_risultato = pd.DataFrame(dati_abaco)

        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl",
        ) as writer:

            df_risultato.to_excel(
                writer,
                index=False,
                sheet_name="Abaco Tavola 39",
            )

        excel_data = output.getvalue()

        st.success("Abaco Excel generato con successo!")

        st.download_button(
            label="📥 Scarica Excel Abaco Segnaletica",
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

    # Cerca l'immagine nella cartella del programma
    cartella_app = os.path.dirname(os.path.abspath(__file__))
    img_path = os.path.join(cartella_app, "TAVOLA 39.jpg")

    if os.path.exists(img_path):

        try:
            # Carica l'immagine originale
            with Image.open(img_path) as immagine_originale:
                base_img = immagine_originale.convert("RGB")

            width, height = base_img.size

            st.caption(
                f"Immagine originale: {width} × {height} pixel"
            )

            draw = ImageDraw.Draw(base_img)

            # ------------------------------------------
            # DIMENSIONE FONT: 200 PIXEL
            # ------------------------------------------
            DIMENSIONE_FONT = 200
            font_km = load_custom_font(DIMENSIONE_FONT)

            # Contorno bianco per leggibilità
            BORDO = 4

            # ------------------------------------------
            # PROGRESSIVE CHILOMETRICHE
            # ------------------------------------------
            etichette = [
                (
                    int(width * 0.60),
                    int(height * 0.002),
                    f"KM {km_fine}",
                ),
                (
                    int(width * 0.59),
                    int(height * 0.525),
                    f"KM {km_testata}",
                ),
                (
                    int(width * 0.59),
                    int(height * 0.645),
                    f"KM {km_preavviso}",
                ),
                (
                    int(width * 0.59),
                    int(height * 0.905),
                    f"KM {km_inizio}",
                ),
            ]

            # ------------------------------------------
            # DISEGNO DELLE SCRITTE
            # ------------------------------------------
            for x, y, testo in etichette:

                draw.text(
                    (x, y),
                    testo,
                    fill=(255, 0, 0),
                    font=font_km,
                    stroke_width=BORDO,
                    stroke_fill=(255, 255, 255),
                )

            # ------------------------------------------
            # SALVATAGGIO JPG IN MEMORIA
            # ------------------------------------------
            img_io = io.BytesIO()

            base_img.save(
                img_io,
                format="JPEG",
                quality=100,
                subsampling=0,
            )

            jpg_data = img_io.getvalue()

            # ------------------------------------------
            # ANTEPRIMA
            # ------------------------------------------
            st.image(
                base_img,
                caption=(
                    "Tavola 39 - Progressive chilometriche "
                    "ingrandite"
                ),
                use_container_width=True,
            )

            # ------------------------------------------
            # DOWNLOAD JPG
            # ------------------------------------------
            st.download_button(
                label="📥 Scarica Tavola in formato JPG",
                data=jpg_data,
                file_name="Tavola_39_Compilata.jpg",
                mime="image/jpeg",
            )

            st.success(
                "Tavola generata correttamente."
            )

        except Exception as errore:
            st.error(
                "Errore durante la generazione della tavola. "
                "Controlla il file immagine e il font."
            )
            st.exception(errore)

    else:

        st.error(
            "Immagine 'TAVOLA 39.jpg' non trovata. "
            "Carica il file nella stessa cartella di app.py."
        )
