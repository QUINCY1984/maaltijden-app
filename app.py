import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import io
import re

st.set_page_config(page_title="Maaltijden Converter", layout="centered")
st.title("🍽️ Maaltijden Converter")
st.write("Upload het ruwe Excel-bestand uit het systeem en bekijk de live statistieken voordat je downloadt.")

uploaded_file = st.file_uploader("Kies het ruwe Excel-bestand", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # --- 1. DATUM UIT BESTANDSNAAM HALEN ---
        file_date = ""
        date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', uploaded_file.name)
        if date_match:
            file_date = f"{date_match.group(3)}-{date_match.group(2)}-{date_match.group(1)}"
        else:
            date_match_2 = re.search(r'(\d{2})-(\d{2})-(\d{4})', uploaded_file.name)
            if date_match_2:
                file_date = date_match_2.group(0)

        # --- 2. RUWE DATA INLEZEN ---
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        
        mask = df_raw.apply(lambda row: row.astype(str).str.strip().eq('Kamer').any(), axis=1)
        if not mask.any():
            st.error("Fout: Kan de kolom 'Kamer' niet vinden in dit bestand.")
            st.stop()
        
        header_row_idx = mask.idxmax()
        header_row = df_raw.iloc[header_row_idx].fillna('').astype(str).str.strip()
        
        expected_cols = ['Kamer', 'Gast(en)', 'Boeker', 'Total guests', 'Posted meals', 'Notities (gast)', 'Prijscode', 'MP Code']
        col_indices = []
        
        for expected in expected_cols:
            matches = header_row[header_row == expected]
            if not matches.empty:
                col_indices.append(matches.index[0])
            else:
                partial = header_row[header_row.str.contains(expected, case=False, regex=False)]
                if not partial.empty:
                    col_indices.append(partial.index[0])
                else:
                    st.error(f"Let op: Kolom '{expected}' niet gevonden!")
                    st.stop()

        df = df_raw.iloc[header_row_idx + 1:, col_indices].copy()
        df.columns = expected_cols
        df = df.dropna(how='all')
        
        df = df[df['Kamer'].astype(str).str.match(r'^\d+$')]
        df['Kamer'] = pd.to_numeric(df['Kamer'], errors='coerce')
        df = df.dropna(subset=['Kamer'])
        df['Kamer'] = df['Kamer'].astype(int)
        df['Total guests'] = pd.to_numeric(df['Total guests'], errors='coerce').fillna(0).astype(int)
        df['Posted meals'] = pd.to_numeric(df['Posted meals'], errors='coerce').fillna(0).astype(int)
        
        # Alle kamers onder de 6000, met uitsluiting van de ongeldige 5810
        df = df[(df['Kamer'] < 6000) & (df['Kamer'] != 5810)]
        
        # --- 3. MAALTIJD FILTER & DASHBOARD OP DE WEBSITE ---
        unique_meals = df['MP Code'].dropna().astype(str).unique().tolist()
        if not unique_meals:
            unique_meals = ["Onbekend"]
            
        st.markdown("---")
        selected_meal = st.selectbox("📌 Welke maaltijd wil je verwerken?", unique_meals)
        
        df_filtered = df[df['MP Code'] == selected_meal]
        
        df_1300 = df_filtered[((df_filtered['Kamer'] >= 1000) & (df_filtered['Kamer'] < 2000)) | ((df_filtered['Kamer'] >= 3000) & (df_filtered['Kamer'] < 4000))]
        df_other = df_filtered[~(((df_filtered['Kamer'] >= 1000) & (df_filtered['Kamer'] < 2000)) | ((df_filtered['Kamer'] >= 3000) & (df_filtered['Kamer'] < 4000)))]
        
        # Live dashboard op het scherm
        st.markdown("### 📊 Live Overzicht (Gasten)")
        col1, col2, col3 = st.columns(3)
        col1.metric("Totaal in huis", df_filtered['Total guests'].sum())
        col2.metric("🏨 Het buffet", df_1300['Total guests'].sum())
        col3.metric("🌅 plad'O", df_other['Total guests'].sum())
        st.markdown("---")

        # --- 4. EXCEL GENEREREN ---
        display_cols = ['Kamer', 'Gast(en)', 'Boeker', 'Guests', 'Meals', 'Notities (gast)', 'Prijscode', 'MP Code']
        
        wb = Workbook()
        wb.remove(wb.active)
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        zebra_fill_1 = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        zebra_fill_2 = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")
        pink_fill = PatternFill(start_color="FFD2D2", end_color="FFD2D2", fill_type="solid")
        allergy_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid") # Fel geel voor diëten
        total_font = Font(bold=True)
        total_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Slimme zoekwoorden voor de keuken/bediening
        allergy_keywords = ['allerg', 'gluten', 'lactose', 'vegan', 'vege', 'dieet', 'intoleran', 'halal', 'noten', 'pinda', 'zonder', 'vrij']

        def format_sheet(ws, df_subset):
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0 
            
            ws.page_margins.left = 0.25
            ws.page_margins.right = 0.25
            ws.page_margins.top = 0.75
            ws.page_margins.bottom = 0.75
            ws.page_margins.header = 0.3
            ws.page_margins.footer = 0.3

            ws.merge_cells('A1:F1')
            # Titel automatisch aangepast aan het type maaltijd
            title_cell = ws.cell(row=1, column=1, value=f"Maaltijdlijsten - {selected_meal.upper()}")
            title_cell.font = Font(bold=True, italic=True, underline="single", size=20)
            title_cell.alignment = Alignment(horizontal='center', vertical='center')
            
            ws.merge_cells('G1:H1')
            date_cell = ws.cell(row=1, column=7, value=file_date)
            date_cell.font = Font(size=12)
            date_cell.alignment = Alignment(horizontal='right', vertical='center')
            
            ws.row_dimensions[1].height = 30
            ws.append([]) # Lege rij

            df_subset = df_subset.sort_values(by='Kamer')
            total_g = df_subset['Total guests'].sum()
            total_m = df_subset['Posted meals'].sum()
            df_display = df_subset.copy()
            df_display.columns = display_cols
            
            rows = dataframe_to_rows(df_display, index=False, header=True)
            for r_idx, row in enumerate(rows, 1):
                ws.append(row)
                excel_row = ws.max_row
                
                is_group = (r_idx > 1 and "groep" in str(row[6]).lower())
                current_fill = header_fill if r_idx == 1 else (pink_fill if is_group else (zebra_fill_1 if r_idx % 2 == 0 else zebra_fill_2))
                
                for c_idx, cell in enumerate(ws[excel_row], 1):
                    cell.border = thin_border
                    cell.fill = current_fill
                    
                    if r_idx == 1:
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    else:
                        align_kwargs = {'vertical': 'center'}
                        if c_idx == 1: align_kwargs['horizontal'] = 'left'
                        elif c_idx in [4, 5]: align_kwargs['horizontal'] = 'center'
                        elif c_idx == 6: 
                            align_kwargs['horizontal'] = 'center'
                            align_kwargs['wrap_text'] = True
                        cell.alignment = Alignment(**align_kwargs)

                        # Check op dieetwensen in kolom 'Notities (gast)'
                        if c_idx == 6 and pd.notna(row[5]):
                            notitie_tekst = str(row[5]).lower()
                            if any(keyword in notitie_tekst for keyword in allergy_keywords):
                                cell.fill = allergy_fill
                                cell.font = Font(bold=True) # Maak tekst vetgedrukt voor extra attentie
            
            max_row = ws.max_row + 1
            ws.cell(row=max_row, column=1, value="TOTAAL").font = total_font
            ws.cell(row=max_row, column=4, value=total_g).font = total_font
            ws.cell(row=max_row, column=5, value=total_m).font = total_font
            ws.cell(row=max_row, column=4).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=max_row, column=5).alignment = Alignment(horizontal='center', vertical='center')
            
            for c_idx in range(1, 9):
                cell = ws.cell(row=max_row, column=c_idx)
                if c_idx not in [1, 4, 5]: cell.value = ""
                cell.fill = total_fill
                cell.border = thin_border
            
            widths = {'A': 6.20, 'B': 30, 'C': 30, 'D': 7, 'E': 7, 'F': 51.57, 'G': 30, 'H': 7}
            for col, w in widths.items():
                ws.column_dimensions[col].width = w
            ws.column_dimensions['B'].hidden = True

        format_sheet(wb.create_sheet(title="Het buffet"), df_1300)
        format_sheet(wb.create_sheet(title="plad'O"), df_other)
        
        # --- DERDE TABBLAD: GROEPEN OVERZICHT ---
        ws_groepen = wb.create_sheet(title="Groepen Overzicht")
        ws_groepen.page_setup.orientation = ws_groepen.ORIENTATION_LANDSCAPE
        ws_groepen.page_setup.paperSize = ws_groepen.PAPERSIZE_A4
        ws_groepen.sheet_properties.pageSetUpPr.fitToPage = True
        ws_groepen.page_setup.fitToWidth = 1
        ws_groepen.page_setup.fitToHeight = 0
        ws_groepen.page_margins.left = 0.25
        ws_groepen.page_margins.right = 0.25
        ws_groepen.page_margins.top = 0.75
        ws_groepen.page_margins.bottom = 0.75
        ws_groepen.page_margins.header = 0.3
        ws_groepen.page_margins.footer = 0.3
        
        def add_group_summary_to_sheet(ws, title, df_subset, start_row):
            df_groups = df_subset[df_subset['Prijscode'].astype(str).str.lower().str.contains('groep')]
            if df_groups.empty:
                ws.cell(row=start_row, column=1, value=f"{title} - GEEN GROEPEN").font = Font(bold=True, size=12)
                return start_row + 2
            
            ws.cell(row=start_row, column=1, value=title).font = Font(bold=True, size=12)
            
            headers = ["Boeker", "Guests", "Meals"]
            for col_num, header_title in enumerate(headers, 1):
                cell = ws.cell(row=start_row+1, column=col_num, value=header_title)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            group_summary = df_groups.groupby('Boeker', as_index=False)[['Total guests', 'Posted meals']].sum()
            
            r = start_row + 2
            for idx, grp_row in group_summary.iterrows():
                c1 = ws.cell(row=r, column=1, value=grp_row['Boeker'])
                c2 = ws.cell(row=r, column=2, value=grp_row['Total guests'])
                c3 = ws.cell(row=r, column=3, value=grp_row['Posted meals'])
                
                fill_color = zebra_fill_1 if idx % 2 == 0 else zebra_fill_2
                for cell in [c1, c2, c3]:
                    cell.border = thin_border
                    cell.fill = fill_color
                    if cell.column > 1:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                r += 1
            
            c1_tot = ws.cell(row=r, column=1, value="TOTAAL")
            c2_tot = ws.cell(row=r, column=2, value=group_summary['Total guests'].sum())
            c3_tot = ws.cell(row=r, column=3, value=group_summary['Posted meals'].sum())
            
            for cell in [c1_tot, c2_tot, c3_tot]:
                cell.font = total_font
                cell.fill = total_fill
                cell.border = thin_border
                if cell.column > 1:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            
            return r + 3

        next_row = 1
        next_row = add_group_summary_to_sheet(ws_groepen, f"GROEPEN: HET BUFFET ({selected_meal.upper()})", df_1300, next_row)
        add_group_summary_to_sheet(ws_groepen, f"GROEPEN: PLAD'O ({selected_meal.upper()})", df_other, next_row)
        
        ws_groepen.column_dimensions['A'].width = 40
        ws_groepen.column_dimensions['B'].width = 12
        ws_groepen.column_dimensions['C'].width = 12

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Grote, duidelijke downloadknop
        st.download_button(
            label=f"📥 DOWNLOAD: {selected_meal.upper()} LIJSTEN",
            data=output,
            file_name=f"Processed_{selected_meal}_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Er ging iets mis tijdens het verwerken: {e}")
