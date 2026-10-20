# =============================================================
# SEZIONE 3: PROGRAMMA LAVORI SETTIMANALE
# =============================================================
elif sezione_app == "Programma Lavori Settimanale":
    st.title("📅 Programma Lavori Settimanale")
    st.caption("Configura il programma settimanale di cantiere e genera l'Excel conforme alle specifiche.")
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        data_inizio_settimana = st.date_input("Data Inizio Settimana (Lunedì)", value=datetime.date(2026, 10, 12))
    with col_d2:
        data_fine_settimana = st.date_input("Data Fine Settimana (Domenica)", value=datetime.date(2026, 10, 18))
        
    col_datore, col_appalto = st.columns([1, 2])
    with col_datore:
        nome_datore = st.text_input(
            "Datore / Impresa (Celle B1:B2):",
            value="VALORI S.C.A.R.L."
        )
    with col_appalto:
        intestazione_appalto = st.text_area(
            "Descrizione Appalto (Celle C1:L1):",
            value="S.S. n° 16 “Adriatica”. Lavori di manutenzione straordinaria per la riqualificazione dell’itinerario Bari-Brindisi-Lecce. \nLotto 1 – Stralcio 6 - Cod. SIL: AANOBA00234."
        )
    
    st.markdown("### 📝 Attività della Settimana")
    st.info("Inserisci le lavorazioni, l'area d'intervento, l'impresa, spunta con 'X' i giorni e indica l'importo presunto netto.")
    
    default_settimana = [
        {"DESCRIZIONE LAVORI": "scavo plinto e armatura", "Area Intervento": "Portale tipo B2 km 827+553", "Impresa": "Tesoro Srl", "Lun": "X", "Mar": "", "Mer": "", "Gio": "", "Ven": "", "Sab": "", "Dom": "", "IMPORTO PRESUNTO NETTO": 30000},
        {"DESCRIZIONE LAVORI": "Casseratura e Getto Cls", "Area Intervento": "Portale tipo B2 km 827+553", "Impresa": "Tesoro Srl", "Lun": "", "Mar": "X", "Mer": "", "Gio": "", "Ven": "", "Sab": "", "Dom": "", "IMPORTO PRESUNTO NETTO": None},
        {"DESCRIZIONE LAVORI": "Scasseratura", "Area Intervento": "Portale tipo B2 km 827+553", "Impresa": "Tesoro Srl", "Lun": "", "Mar": "", "Mer": "X", "Gio": "", "Ven": "", "Sab": "", "Dom": "", "IMPORTO PRESUNTO NETTO": None},
    ]
    df_sett = pd.DataFrame(default_settimana)
    edited_df_sett = st.data_editor(df_sett, num_rows="dynamic", use_container_width=True)
    
    st.markdown("### 👷 Risorse stimate giornaliere (Lunedì - Venerdì)")
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        pers_lun = st.text_input("Personale Lun", value="3/5")
        pers_mar = st.text_input("Personale Mar", value="3/5")
        pers_mer = st.text_input("Personale Mer", value="3/5")
        pers_gio = st.text_input("Personale Gio", value="3/5")
        pers_ven = st.text_input("Personale Ven", value="3/5")
    with col_r2:
        mez_lun = st.text_input("Mezzi Lun", value="2/3")
        mez_mar = st.text_input("Mezzi Mar", value="2/3")
        mez_mer = st.text_input("Mezzi Mer", value="2/3")
        mez_gio = st.text_input("Mezzi Gio", value="2/3")
        mez_ven = st.text_input("Mezzi Ven", value="2/3")

    if st.button("🚀 Genera File Excel Programma Lavori Settimanale"):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PROGRAMMA SS 16- L1S2"
        ws.views.sheetView[0].showGridLines = False
        
        font_company = Font(name="Arial", size=10, bold=True)
        font_appalto = Font(name="Arial", size=11, bold=True)
        font_subtitle_week = Font(name="Arial", size=11, bold=True)
        font_header = Font(name="Arial", size=10, bold=True, color="000000")
        font_body = Font(name="Arial", size=10)
        font_footer = Font(name="Arial", size=10, bold=True)
        
        fill_header = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        fill_weekend = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
        thin_side = Side(style='thin', color='000000')
        thick_side = Side(style='medium', color='000000')
        border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        
        # Unione celle B1:B2 e riempimento giallo (Datore di lavoro)
        ws.merge_cells("B1:B2")
        cell_b1 = ws.cell(row=1, column=2, value=nome_datore)
        cell_b1.font = font_company
        cell_b1.fill = fill_yellow
        cell_b1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in range(1, 3):
            ws.cell(row=r, column=2).border = border_all
            
        # Unione celle C1:L1 (Intestazione Appalto)
        ws.merge_cells("C1:L1")
        cell_c1 = ws.cell(row=1, column=3, value=intestazione_appalto)
        cell_c1.font = font_appalto
        cell_c1.fill = fill_yellow
        cell_c1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for c in range(3, 13):
            ws.cell(row=1, column=c).border = border_all
        ws.row_dimensions[1].height = 40
        
        # Unione celle C2:L2 (Periodo Settimana)
        ws.merge_cells("C2:L2")
        str_periodo = f"Programma Lavori Settimanale - dal {data_inizio_settimana.strftime('%d/%m/%Y')} al {data_fine_settimana.strftime('%d/%m/%Y')}"
        cell_c2 = ws.cell(row=2, column=3, value=str_periodo)
        cell_c2.font = font_subtitle_week
        cell_c2.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(3, 13):
            ws.cell(row=2, column=c).border = border_all
        ws.row_dimensions[2].height = 25
        
        giorni_nomi = {0: "Lun", 1: "Mar", 2: "Mer", 3: "Gio", 4: "Ven", 5: "Sab", 6: "Dom"}
        
        # Riga 3: Intestazioni colonne
        headers_base = ["DESCRIZIONE LAVORI", "Area Intervento", "Impresa"]
        ws.row_dimensions[3].height = 30
        
        for c_idx, h_text in enumerate(headers_base, start=2):
            cell = ws.cell(row=3, column=c_idx, value=h_text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_all
            
        # Intestazioni giorni (Sabato e Domenica con sfondo rosso chiaro)
        for i in range(7):
            current_d = data_inizio_settimana + datetime.timedelta(days=i)
            nome_g = giorni_nomi[current_d.weekday()]
            testo_giorno = f"{nome_g} {current_d.strftime('%d/%m')}"
            cell_g = ws.cell(row=3, column=5 + i, value=testo_giorno)
            cell_g.font = font_header
            if current_d.weekday() >= 5:
                cell_g.fill = fill_weekend
            else:
                cell_g.fill = fill_header
            cell_g.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell_g.border = border_all
            
        cell_imp_h = ws.cell(row=3, column=12, value="IMPORTO PRESUNTO NETTO (lavori+sic)")
        cell_imp_h.font = font_header
        cell_imp_h.fill = fill_header
        cell_imp_h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell_imp_h.border = border_all
                
        # Inserimento righe lavorazioni
        start_row_sett = 4
        current_row = start_row_sett
        
        for idx, row in edited_df_sett.iterrows():
            ws.row_dimensions[current_row].height = 22
            
            c_desc = ws.cell(row=current_row, column=2, value=row.get("DESCRIZIONE LAVORI", ""))
            c_desc.font = font_body
            c_desc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c_desc.border = border_all
            
            c_area = ws.cell(row=current_row, column=3, value=row.get("Area Intervento", ""))
            c_area.font = font_body
            c_area.alignment = Alignment(horizontal="center", vertical="center")
            c_area.border = border_all
            
            c_imp = ws.cell(row=current_row, column=4, value=row.get("Impresa", ""))
            c_imp.font = font_body
            c_imp.alignment = Alignment(horizontal="center", vertical="center")
            c_imp.border = border_all
            
            # Giorni (colonne 5 a 11)
            giorni_keys = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]
            for g_i, g_key in enumerate(giorni_keys):
                val_g = row.get(g_key, "")
                cell_g = ws.cell(row=current_row, column=5 + g_i, value=val_g if val_g else "")
                cell_g.alignment = Alignment(horizontal="center", vertical="center")
                cell_g.font = font_body
                if g_i >= 5:
                    cell_g.fill = fill_weekend
                cell_g.border = border_all
                
            imp_val = row.get("IMPORTO PRESUNTO NETTO", None)
            cell_importo = ws.cell(row=current_row, column=12)
            if pd.notnull(imp_val) and imp_val != "":
                cell_importo.value = imp_val
            else:
                cell_importo.value = None
            cell_importo.font = font_body
            cell_importo.alignment = Alignment(horizontal="center", vertical="center")
            cell_importo.number_format = '€ #,##0.00'
            cell_importo.border = border_all
                
            current_row += 1
            
        pers_row = current_row
        mezzi_row = pers_row + 1
        
        # Personale Stimato (Lun-Ven inseriti, Sab-Dom vuoti e rossi)
        ws.row_dimensions[pers_row].height = 20
        ws.merge_cells(start_row=pers_row, start_column=2, end_row=pers_row, end_column=4)
        cell_p_label = ws.cell(row=pers_row, column=2, value="PERSONALE STIMATO")
        cell_p_label.font = font_footer
        cell_p_label.alignment = Alignment(horizontal="center", vertical="center")
        for c_i in range(2, 5):
            ws.cell(row=pers_row, column=c_i).fill = fill_green
            ws.cell(row=pers_row, column=c_i).border = Border(top=thick_side, left=thin_side, right=thin_side, bottom=thin_side)
            
        pers_valori = [pers_lun, pers_mar, pers_mer, pers_gio, pers_ven, "", ""]
        for p_i, p_val in enumerate(pers_valori):
            cell_p = ws.cell(row=pers_row, column=5 + p_i, value=p_val)
            cell_p.alignment = Alignment(horizontal="center", vertical="center")
            cell_p.font = font_body
            if p_i >= 5:
                cell_p.fill = fill_weekend
            cell_p.border = Border(top=thick_side, left=thin_side, right=thin_side, bottom=thin_side)
            
        cell_l_p = ws.cell(row=pers_row, column=12)
        cell_l_p.border = Border(top=thick_side, left=thin_side, right=thin_side, bottom=thin_side)
            
        # Mezzi Stimati (Lun-Ven inseriti, Sab-Dom vuoti e rossi)
        ws.row_dimensions[mezzi_row].height = 20
        ws.merge_cells(start_row=mezzi_row, start_column=2, end_row=mezzi_row, end_column=4)
        cell_m_label = ws.cell(row=mezzi_row, column=2, value="MEZZI STIMATI")
        cell_m_label.font = font_footer
        cell_m_label.alignment = Alignment(horizontal="center", vertical="center")
        for c_i in range(2, 5):
            ws.cell(row=mezzi_row, column=c_i).fill = fill_green
            ws.cell(row=mezzi_row, column=c_i).border = border_all
            
        mezzi_valori = [mez_lun, mez_mar, mez_mer, mez_gio, mez_ven, "", ""]
        for m_i, m_val in enumerate(mezzi_valori):
            cell_m = ws.cell(row=mezzi_row, column=5 + m_i, value=m_val)
            cell_m.alignment = Alignment(horizontal="center", vertical="center")
            cell_m.font = font_body
            if m_i >= 5:
                cell_m.fill = fill_weekend
            cell_m.border = border_all
            
        ws.cell(row=mezzi_row, column=12).border = border_all
            
        # Riga Totale
        tot_row = mezzi_row + 2
        ws.row_dimensions[tot_row].height = 22
        cell_t_label = ws.cell(row=tot_row, column=11, value="TOTALE")
        cell_t_label.font = font_footer
        cell_t_label.alignment = Alignment(horizontal="center", vertical="center")
        
        cell_tot_sett = ws.cell(row=tot_row, column=12)
        cell_tot_sett.value = f"=SUM(L{start_row_sett}:L{pers_row-1})"
        cell_tot_sett.font = font_footer
        cell_tot_sett.alignment = Alignment(horizontal="center", vertical="center")
        cell_tot_sett.number_format = '€ #,##0.00'
        cell_tot_sett.border = border_all
        
        # Applicazione della BORDATURA SPESSA al PERIMETRO ESTERNO della tabella
        max_r_tabella = mezzi_row
        for r_idx in range(1, max_r_tabella + 1):
            for c_idx in range(2, 13):
                cell = ws.cell(row=r_idx, column=c_idx)
                t_top = thick_side if r_idx == 1 else cell.border.top
                t_bottom = thick_side if r_idx == max_r_tabella else cell.border.bottom
                t_left = thick_side if c_idx == 2 else cell.border.left
                t_right = thick_side if c_idx == 12 else cell.border.right
                cell.border = Border(top=t_top, bottom=t_bottom, left=t_left, right=t_right)
        
        # Larghezze colonne dinamiche
        ws.column_dimensions['A'].width = 3
        max_len_desc = max([len(str(r.get("DESCRIZIONE LAVORI", ""))) for r in edited_df_sett.to_dict("records")] + [20])
        ws.column_dimensions['B'].width = max(max_len_desc + 4, 25)
        max_len_area = max([len(str(r.get("Area Intervento", ""))) for r in edited_df_sett.to_dict("records")] + [20])
        ws.column_dimensions['C'].width = max(max_len_area + 4, 28)
        max_len_imp = max([len(str(r.get("Impresa", ""))) for r in edited_df_sett.to_dict("records")] + [10])
        ws.column_dimensions['D'].width = max(max_len_imp + 4, 15)
        
        for col_idx in range(5, 12):
            ws.column_dimensions[get_column_letter(col_idx)].width = 13
        ws.column_dimensions['L'].width = 32
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.success("File Excel Programma Settimanale generato con successo!")
        st.download_button(
            label="📥 Scarica Programma Settimanale Excel",
            data=output,
            file_name=f"PROGRAMMA_SETTIMANALE_{data_inizio_settimana.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
