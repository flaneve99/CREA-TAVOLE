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
# CARICAMENTO FONT TRUE TYPE
# ==================================================
@st.cache_resource
def load_custom_font(size):
    font_path = "DejaVuSans-Bold.ttf"

    if not os.path.exists(font_path):
        url = (
            "https://github.com/dejavu-fonts/dejavu-fonts/"
            "raw/master/ttf/DejaVuSans-Bold.ttf"
        )

        try:
            urllib.request.urlretrieve(url, font_path)
        except Exception:
            font_path = ""

    if font_path and os.path.exists(font_path):
        return ImageFont.truetype(font_path, size)

    # Prova altri font di sistema prima di usare quello predefinito
    possibili_font = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        "C:/Windows/Fonts/arialbd.ttf",
        "/Library/Fonts/Arial Bold.ttf",
    ]

    for percorso in possibili_font:
        if os.path.exists(percorso):
            return ImageFont.truetype(percorso, size)

    raise RuntimeError(
        "Font TrueType non disponibile. "
        "Inserisci DejaVuSans-Bold.ttf nella cartella del programma."
    )


# ==================================================
# SELEZIONE TAVOLA
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
    # GENERAZIONE EXCEL
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

    img_path = "TAVOLA 39.jpg"

    if os.path.exists(img_path):

        # Apri l'immagine originale
        base_img = Image.open(img_path).convert("RGB")

        width, height = base_img.size

        st.caption(
            f"Dimensioni immagine originale: {width} × {height} pixel"
        )

        draw = ImageDraw.Draw(base_img)

        # ----------------------------------------------
        # DIMENSIONE FONT: 200 PIXEL
        # ----------------------------------------------
        DIMENSIONE_FONT = 200

        font_km = load_custom_font(DIMENSIONE_FONT)

        # Contorno bianco per rendere le scritte più leggibili
        BORDO = 4

        # ----------------------------------------------
        # POSIZIONI DELLE PROGRESSIVE
        # ----------------------------------------------
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

        # ----------------------------------------------
        # DISEGNA LE SCRITTE
        # ----------------------------------------------
        for x, y, testo in etichette:
            draw.text(
                (x, y),
                testo,
                fill=(255, 0, 0),
                font=font_km,
                stroke_width=BORDO,
                stroke_fill=(255, 255, 255),
            )

        # ----------------------------------------------
        # SALVATAGGIO IN MEMORIA
        # ----------------------------------------------
        img_io = io.BytesIO()

        base_img.save(
            img_io,
            format="JPEG",
            quality=100,
            subsampling=0,
        )

        jpg_data = img_io.getvalue()

        # ----------------------------------------------
        # ANTEPRIMA
        # ----------------------------------------------
        st.image(
            base_img,
            caption="Tavola 39 - Progressive chilometriche ingrandite",
            use_container_width=True,
        )

        # ----------------------------------------------
        # DOWNLOAD JPG
        # ----------------------------------------------
        st.download_button(
            label="📥 Scarica Tavola in formato JPG",
            data=jpg_data,
            file_name="Tavola_39_Compilata.jpg",
            mime="image/jpeg",
        )

        st.success(
            f"Tavola generata. Dimensione font utilizzata: "
            f"{DIMENSIONE_FONT} pixel."
        )

    else:
        st.error(
            f"Immagine '{img_path}' non trovata. "
            "Inserisci il file TAVOLA 39.jpg nella stessa "
            "cartella del file Python."
        )
