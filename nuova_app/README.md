# Progetto Tavole e Abaco Segnaletica

## Avvio locale
1. Installare Python 3.10 o superiore.
2. Aprire un terminale nella cartella del progetto.
3. Eseguire `pip install -r requirements.txt`.
4. Eseguire `streamlit run app.py`.

## Contenuto
- `app.py`: app Streamlit iniziale.
- `data/tavole.json`: inventario delle 159 pagine del PDF e intestazioni rilevate automaticamente.
- `data/segnali.json`: catalogo dei dati testuali dell'Excel.
- `assets/`: PDF e Excel originali e immagini estratte dal foglio ABACO.

## Limiti dichiarati
Questa versione permette di consultare le pagine del PDF, cercare nel catalogo, comporre manualmente un abaco ed esportarlo in Excel con dati di cantiere. Non associa automaticamente i segnali a ogni tavola: questa mappatura richiede una verifica schema per schema. Le celle `#VALUE!` nel file originale sono conservate come campo vuoto nel catalogo esportato, senza inventare codici sostitutivi.
