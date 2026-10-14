import io
import os
import urllib.request
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

st.set_page_config(
    page_title="Gestione Cantieri Stradali - Abaco & Tavole",
    layout="wide"
)

st.title("🚧 Gestione Cantieri Stradali e Segnaletica")
st.write(
    "Inserisci i dati chilometrici: il sito aggiornerà l'Abaco Excel "
    "e genererà la Tavola in JPG personalizzata!"
)


# --------------------------------------------------
# FUNZIONE CARICAMENTO FONT
# --------------------------------------------------
@st.cache_resource
def load_custom_font(size):
    font_path = "DejaVuSans-Bold.ttf"

    if not os.path.exists(font_path):
        try:
            url = (
                "https://github.com/dejavu-fonts/dejavu-fonts/"
                "raw/master/ttf/DejaVuSans-Bold.ttf"
            )
            urllib.request.urlretrieve(url, font_path)
        except Exception:
            pass

    if os.path.exists(font_path):
        return ImageFont.truetype(font_path, size)

    return ImageFont.load_default()


# --------------------------------------------------
# SELEZIONE TAVOLA
# --------------------------------------------------
tavola_selezionata = st.selectbox(
    "Seleziona la Tavola di Cantiere",
    ["Tavola 39 - Cantiere mobile carreggiata a due corsie"],
)


if "39" in tavola_selezionata:

    st.subheader("📍 Dati e Progressive Chilometriche - Tavola 39")

    st.markdown(
        "*Direzione Sud, Viadotto Torre Guaceto 1 - "
        "Chiusura corsia di destra*"
    )

    col1, col2 = st.columns(2)

    with col1:
        km_inizio = st.text_input(
            "KM Inizio Cantiere (Basso)",
            "35+560"
        )

        km_preavviso = st.text_input(
            "KM Segnale Mobile Preavviso (Centro)",
            "36+060"
        )

    with col2:
        km_fine = st.text_input(
            "KM Fine / Rientro (Alto)",
            "36+600"
        )

        km_testata = st.text_input(
            "KM Inizio Testata / Intermedio",
            "36+110"
        )

        distanza_coni = st.text_input(
            "Interasse Coni / Dettagli",
            "12 m"
        )

    # --------------------------------------------------
    # 1. GENERAZIONE FILE EXCEL (ABACO)
    # --------------------------------------------------
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
            engine="openpyxl"
        ) as writer:

            df_risultato.to_excel(
                writer,
                index=False,
                sheet_name="Abaco Tavola 39"
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

    # --------------------------------------------------
    # 2. GENERAZIONE TAVOLA PERSONALIZZATA
    # --------------------------------------------------
    st.divider()

    st.subheader("🖼️ Anteprima Tavola 39 Personalizzata")

    img_path = "TAVOLA 39.jpg"

    if os.path.exists(img_path):

        base_img = Image.open(img_path).convert("RGB")
        draw = ImageDraw.Draw(base_img)

        width, height = base_img.size

        # ----------------------------------------------
        # DIMENSIONI CARATTERI: 200 PIXEL
        # ----------------------------------------------
        font_grande = load_custom_font(200)
        font_medio = load_custom_font(200)

        # Bordo bianco per aumentare la leggibilità
        bordo = 4

        # ----------------------------------------------
        # 1. KM FINE (IN ALTO)
        # ----------------------------------------------
        draw.text(
            (int(width * 0.60), int(height * 0.002)),
            f"KM {km_fine}",
            fill="red",
            font=font_grande,
            stroke_width=bordo,
            stroke_fill="white",
        )

        # ----------------------------------------------
        # 2. KM TESTATA / INTERMEDIO
        # ----------------------------------------------
        draw.text(
            (int(width * 0.59), int(height * 0.525)),
            f"KM {km_testata}",
            fill="red",
            font=font_medio,
            stroke_width=bordo,
            stroke_fill="white",
        )

        # ----------------------------------------------
        # 3. KM PREAVVISO (AL CENTRO)
        # ----------------------------------------------
        draw.text(
            (int(width * 0.59), int(height * 0.645)),
            f"KM {km_preavviso}",
            fill="red",
            font=font_medio,
            stroke_width=bordo,
            stroke_fill="white",
        )

        # ----------------------------------------------
        # 4. KM INIZIO (IN BASSO)
        # ----------------------------------------------
        draw.text(
            (int(width * 0.59), int(height * 0.905)),
            f"KM {km_inizio}",
            fill="red",
            font=font_medio,
            stroke_width=bordo,
            stroke_fill="white",
        )

        # ----------------------------------------------
        # SALVATAGGIO IMMAGINE IN MEMORIA
        # ----------------------------------------------
        img_io = io.BytesIO()

        base_img.save(
            img_io,
            format="JPEG",
            quality=95
        )

        img_io.seek(0)

        # ----------------------------------------------
        # ANTEPRIMA
        # ----------------------------------------------
        st.image(
            base_img,
            caption="Tavola 39 - Progressive chilometriche",
            use_container_width=True,
        )

        # ----------------------------------------------
        # DOWNLOAD JPG
        # ----------------------------------------------
        st.download_button(
            label="📥 Scarica Tavola in formato JPG",
            data=img_io.getvalue(),
            file_name="Tavola_39_Compilata.jpg",
            mime="image/jpeg",
        )

    else:

        st.warning(
            f"⚠️ L'immagine '{img_path}' non è stata trovata. "
            "Assicurati che sia presente nella stessa cartella "
            "del file Python e che il nome sia esattamente "
            "'TAVOLA 39.jpg'."
        )
