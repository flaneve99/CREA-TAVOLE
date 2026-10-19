# config_tavole.py

CONFIG_TAVOLE = {
    "Tavola 16 - Chiusura della corsia di marcia su carreggiata a due corsie": {
        "immagine_sfondo": "TAVOLA 16.jpg",
        "moltiplicatore_font_km": 25,
        "moltiplicatore_font_ub": 30,
        "ubicazione_default": "DIREZIONE SUD\nVIADOTTO TORRE\nGUACETO 1",
        "posizione_ubicazione": {"x_pct": 0.67, "y_pct": 0.27},
        "inputs_km": [
            {"id": "km_fine_16", "label": "KM Fine / Rientro (In alto)", "default": "36+600", "x_pct": 0.34, "y_pct": 0.07},
            {"id": "km_preavviso_700", "label": "KM Preavviso 700m (In basso)", "default": "34+860", "x_pct": 0.49, "y_pct": 0.925},
            {"id": "km_limite_110", "label": "KM Limite 110 km/h", "default": "35+100", "x_pct": 0.49, "y_pct": 0.865},
            {"id": "km_limite_90", "label": "KM Limite 90 km/h", "default": "35+220", "x_pct": 0.49, "y_pct": 0.800},
            {"id": "km_preavviso_250", "label": "KM Preavviso Chiusura Corsia 250m", "default": "35+310", "x_pct": 0.49, "y_pct": 0.735},
            {"id": "km_limite_60", "label": "KM Limite 60 km/h", "default": "35+430", "x_pct": 0.49, "y_pct": 0.675},
            {"id": "km_preavviso_100", "label": "KM Preavviso Chiusura Corsia 100m", "default": "35+490", "x_pct": 0.49, "y_pct": 0.610},
            {"id": "km_inizio_testata", "label": "KM Inizio Testata (Freccia blu)", "default": "35+560", "x_pct": 0.49, "y_pct": 0.505}
        ],
        "abaco_segnali": [
            ("fig_383.jpg", "FIG. II 383", "LAVORI IN CORSO", "", "SEGNALI TEMPORANEI"),
            ("fig_082a.jpg", "FIG. II 082/a", "PASSAGGIO OBBLIGATORIO A SINISTRA", "", "SEGNALI DI OBBLIGO"),
            ("fig_048.jpg", "FIG. II 048", "DIVIETO DI SORPASSO", "", "SEGNALI DI DIVIETO"),
            ("fig_411a.jpg", "FIG. II 411/a", "SEGNALE DI CORSIA CHIUSA (CHIUSURA CORSIA DI DESTRA)", "", "SEGNALI TEMPORANEI"),
            ("fig050_lim110.jpg", "FIG. II 050", "LIMITE DI VELOCITÀ 110 KM/H", "", "SEGNALI DI PRESCRIZIONE"),
            ("fig050_lim90.jpg", "FIG. II 050", "LIMITE DI VELOCITÀ 90 KM/H", "", "SEGNALI DI PRESCRIZIONE"),
            ("fig_050_lim60.jpg", "FIG. II 050", "LIMITE DI VELOCITÀ 60 KM/H", "", "SEGNALI DI PRESCRIZIONE"),
            ("fig_070.jpg", "FIG. II 070", "VIA LIBERA", "", "SEGNALI DI DIVIETO")
        ]
    },
    "Tavola 39 - Cantiere mobile carreggiata a due corsie": {
        "immagine_sfondo": "TAVOLA 39.jpg",
        "moltiplicatore_font_km": 60,
        "moltiplicatore_font_ub": 69,
        "ubicazione_default": "DIREZIONE SUD\nVIADOTTO TORRE\nGUACETO 1",
        "posizione_ubicazione": {"x_pct": 0.65, "y_pct": 0.225},
        "inputs_km": [
            {"id": "km_fine", "label": "KM Fine / Rientro (Alto)", "default": "36+600", "x_pct": 0.45, "y_pct": 0.015},
            {"id": "km_testata", "label": "KM Inizio Testata / Intermedio", "default": "36+110", "x_pct": 0.62, "y_pct": 0.525},
            {"id": "km_preavviso", "label": "KM Segnale Mobile Preavviso (Centro)", "default": "36+060", "x_pct": 0.62, "y_pct": 0.645},
            {"id": "km_inizio", "label": "KM Inizio Cantiere (Basso)", "default": "35+560", "x_pct": 0.62, "y_pct": 0.905}
        ],
        "abaco_segnali": [
            ("fig_399.jpg", "FIG. II 399", "PREAVVISO DI CANTIERE MOBILE", "", "SEGNALI TEMPORANEI"),
            ("fig_383.jpg", "FIG. II 383", "LAVORI IN CORSO", "", "SEGNALI TEMPORANEI"),
            ("fig_070.jpg", "FIG. II 070", "FINE CANTIERE", "", "SEGNALI DI PRESCRIZIONE"),
            ("fig_400.jpg", "FIG. II 400", "SEGNALE MOBILE DI PREAVVISO CON FRECCIA LUMINOSA", "", "SEGNALI TEMPORANEI"),
            ("fig_401.jpg", "FIG. II 401", "VEICOLO DI PROTEZIONE CON PANNELLO", "", "SEGNALI TEMPORANEI")
        ]
    }
}
