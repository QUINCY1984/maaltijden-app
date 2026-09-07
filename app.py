import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import io
import re

st.set_page_config(page_title="Maaltijden Converter", layout="centered")
st.title("🍽️ Flexibele Maaltijden Converter")
st.write("Upload het ruwe Excel-bestand en kies zelf welke maaltijden je op je lijsten wilt combineren.")

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
        
        expected_cols = ['Kamer', 'Gast(en)', 'Boeker', 'Total guests', 'Posted meals', 'Kind 0 - 3', 'Kind 4 - 11', 'Notities (gast)', 'Prijscode', 'MP Code']
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
                    st.error(f"Let op: Kolom '{expected}' niet gevonden in het bestand!")
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
        df['Kind 0 - 3'] = pd.to_numeric(df['Kind 0 - 3'], errors='coerce').fillna(0).astype(int)
        df['Kind 4 - 11'] = pd.to_numeric(df['Kind 4 - 11'], errors='coerce').fillna(0).astype(int)
        
        df = df[(df['Kamer'] < 6000) & (df['Kamer'] != 5810)]
        
        for col in ['Gast(en)', 'Boeker', 'Notities (gast)', 'Prijscode']:
            df[col] = df[col].fillna('')
            
        df['MP_Standaard'] = df['MP Code'].astype(str).str.lower().replace({
            'ontbijt': 'Ontbijt', 'breakfast': 'Ontbijt',
            'lunch': 'Lunch',
            'diner': 'Diner', 'dinner': 'Diner'
        })
        
        # --- 3. DYNAMISCHE KEUZE VAN MAALTIJDEN ---
        st.markdown("---")
        beschikbare_maaltijden = ['Ontbijt', 'Lunch', 'Diner']
        selected_meals = st.multiselect(
            "📌 Welke maaltijden wil je op deze lijst combineren?",
            beschikbare_maaltijden,
            default=beschikbare_maaltijden
        )
        
        if not selected_meals:
            st.warning("Kies minstens één maaltijd om verder te gaan.")
            st.stop()
            
        df = df[df['MP_Standaard'].isin(selected_meals)]
        
        if df.empty:
            st.error("Er is geen data gevonden voor de geselecteerde maaltijd(en).")
            st.stop()

        has_mp_code = len(selected_meals) == 1
        pivot_index = ['Kamer', 'Gast(en)', 'Boeker', 'Total guests', 'Kind 0 - 3', 'Kind 4 - 11', 'Notities (gast)', 'Prijscode']
        if has_mp_code:
            pivot_index.append('MP Code')

        df_pivot = pd.pivot_table(
            df,
            index=pivot_index,
            columns='MP_Standaard',
            values='Posted meals', 
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        
        # Hernoemen naar de strakke lay-out
        df_pivot.rename(columns={'Total guests': 'Guests', 'Kind 0 - 3': 'Kind -3', 'Kind 4 - 11': 'Kind -11'}, inplace=True)
        
        for meal in selected_meals:
            if meal not in df_pivot.columns:
                df_pivot[meal] = 0
                
        df_1300 = df_pivot[((df_pivot['Kamer'] >= 1000) & (df_pivot['Kamer'] < 2000)) | ((df_pivot['Kamer'] >= 3000) & (df_pivot['Kamer'] < 4000))]
        df_other = df_pivot[~(((df_pivot['Kamer'] >= 1000) & (df_pivot['Kamer'] < 2000)) | ((df_pivot['Kamer'] >= 3000) & (df_pivot['Kamer'] < 4000)))]
        
        # --- Live Overzicht Dashboard ---
        st.markdown("### 📊 Live Overzicht")
        cols = st.columns(len(selected_meals) + 1)
        cols[0].metric("👥 Totaal Gasten", df_pivot['Guests'].sum())
        for i, meal in enumerate(selected_meals):
            cols[i+1].metric(f"🍽️ {meal}", df_pivot[meal].sum())
        st.markdown("---")

        # --- 4. DYNAMISCHE EXCEL GENEREREN ---
        # Display kolommen (Met Kind -3 en Kind -11)
        display_cols = ['Kamer', 'Gast(en)', 'Boeker', 'Guests'] + selected_meals + ['Kind -3', 'Kind -11', 'Notities (gast)', 'Prijscode']
        if has_mp_code:
            display_cols.append('MP Code')
            
        meal_indices = list(range(5, 5 + len(selected_meals)))
        kind1_idx = 5 + len(selected_meals)
        kind2_idx = kind1_idx + 1
        notities_idx = kind2_idx + 1
        prijscode_idx = notities_idx + 1
        mp_code_idx = prijscode_idx + 1 if has_mp_code else None
        
        base_total = 145.77
        fixed_width = 6.20 + 30 + 7 + 7 + 7 + 30 
        meal_width = len(selected_meals) * 7
        mp_width = 7 if has_mp_code else 0
        notities_width = base_total - fixed_width - meal_width - mp_width
        
        wb = Workbook()
        wb.remove(wb.active)
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        zebra_fill_1 = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        zebra_fill_2 = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")
        pink_fill = PatternFill(start_color="FFD2D2", end_color="FFD2D2", fill_type="solid")
        allergy_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid") 
        total_font = Font(bold=True)
        total_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        allergy_keywords = ['allerg', 'gluten', 'lactose', 'vegan', 'vege', 'dieet', 'intoleran', 'halal', 'noten', 'pinda', 'zonder', 'vrij']

        def format_sheet(ws, df_subset, is_first_sheet=False):
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

            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=notities_idx)
            titel_tekst = f"Maaltijdlijsten - {', '.join(selected_meals).upper()}"
            title_cell = ws.cell(row=1, column=1, value=titel_tekst)
            title_cell.font = Font(bold=True, italic=True, underline="single", size=20)
            title_cell.alignment = Alignment(horizontal='center', vertical='center')
            
            if has_mp_code:
                ws.merge_cells(start_row=1, start_column=prijscode_idx, end_row=1, end_column=mp_code_idx)
                
            date_cell = ws.cell(row=1, column=prijscode_idx, value=file_date)
            date_cell.font = Font(size=12)
            date_cell.alignment = Alignment(horizontal='right', vertical='center')
            
            ws.row_dimensions[1].height = 35 
            ws.append([]) # Lege rij

            df_subset = df_subset.sort_values(by='Kamer')
            df_display = df_subset[display_cols].copy()
            
            rows = dataframe_to_rows(df_display, index=False, header=True)
            for r_idx, row in enumerate(rows, 1):
                ws.append(row)
                excel_row = ws.max_row
                
                is_group = (r_idx > 1 and "groep" in str(row[prijscode_idx - 1]).lower())
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
                        elif c_idx in [4, kind1_idx, kind2_idx] or c_idx in meal_indices: align_kwargs['horizontal'] = 'center'
                        elif c_idx == notities_idx: 
                            align_kwargs['horizontal'] = 'center'
                            align_kwargs['wrap_text'] = True
                        elif has_mp_code and c_idx == mp_code_idx:
                            align_kwargs['horizontal'] = 'center'
                            
                        cell.alignment = Alignment(**align_kwargs)

                        if c_idx == notities_idx and pd.notna(row[notities_idx - 1]):
                            notitie_tekst = str(row[notities_idx - 1]).lower()
                            if any(keyword in notitie_tekst for keyword in allergy_keywords):
                                cell.fill = allergy_fill
                                cell.font = Font(bold=True)
            
            max_row = ws.max_row + 1
            max_col = mp_code_idx if has_mp_code else prijscode_idx
            
            ws.cell(row=max_row, column=1, value="TOTAAL").font = total_font
            ws.cell(row=max_row, column=4, value=df_subset['Guests'].sum()).font = total_font
            ws.cell(row=max_row, column=4).alignment = Alignment(horizontal='center', vertical='center')
            
            for m_idx in meal_indices:
                meal_name = selected_meals[m_idx - 5]
                meal_tot = df_subset[meal_name].sum()
                ws.cell(row=max_row, column=m_idx, value=meal_tot).font = total_font
                ws.cell(row=max_row, column=m_idx).alignment = Alignment(horizontal='center', vertical='center')
            
            # Totalen voor de kinderen
            ws.cell(row=max_row, column=kind1_idx, value=df_subset['Kind -3'].sum()).font = total_font
            ws.cell(row=max_row, column=kind1_idx).alignment = Alignment(horizontal='center', vertical='center')
            
            ws.cell(row=max_row, column=kind2_idx, value=df_subset['Kind -11'].sum()).font = total_font
            ws.cell(row=max_row, column=kind2_idx).alignment = Alignment(horizontal='center', vertical='center')
            
            for c_idx in range(1, max_col + 1):
                cell = ws.cell(row=max_row, column=c_idx)
                if c_idx not in [1, 4, kind1_idx, kind2_idx] + meal_indices: cell.value = ""
                cell.fill = total_fill
                cell.border = thin_border
            
            col_widths = {1: 6.20, 2: 30, 3: 30, 4: 7}
            for m_idx in meal_indices:
                col_widths[m_idx] = 7
                
            col_widths[kind1_idx] = 7
            col_widths[kind2_idx] = 7
            col_widths[notities_idx] = notities_width
            col_widths[prijscode_idx] = 30
            if has_mp_code:
                col_widths[mp_code_idx] = 7
            
            for col_idx, w in col_widths.items():
                ws.column_dimensions[get_column_letter(col_idx)].width = w
            ws.column_dimensions['B'].hidden = True

        format_sheet(wb.create_sheet(title="Het buffet"), df_1300, is_first_sheet=True)
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
            
            headers = ["Boeker", "Guests"] + selected_meals + ["Kind -3", "Kind -11"]
            for col_num, header_title in enumerate(headers, 1):
                cell = ws.cell(row=start_row+1, column=col_num, value=header_title)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            group_summary = df_groups.groupby('Boeker', as_index=False)[['Guests'] + selected_meals + ['Kind -3', 'Kind -11']].sum()
            
            r = start_row + 2
            for idx, grp_row in group_summary.iterrows():
                ws.cell(row=r, column=1, value=grp_row['Boeker']).border = thin_border
                ws.cell(row=r, column=1).fill = zebra_fill_1 if idx % 2 == 0 else zebra_fill_2
                
                for m_idx, col_name in enumerate(['Guests'] + selected_meals + ['Kind -3', 'Kind -11'], 2):
                    c = ws.cell(row=r, column=m_idx, value=grp_row[col_name])
                    c.border = thin_border
                    c.fill = zebra_fill_1 if idx % 2 == 0 else zebra_fill_2
                    c.alignment = Alignment(horizontal='center', vertical='center')
                r += 1
            
            ws.cell(row=r, column=1, value="TOTAAL").font = total_font
            ws.cell(row=r, column=1).fill = total_fill
            ws.cell(row=r, column=1).border = thin_border
            
            for m_idx, col_name in enumerate(['Guests'] + selected_meals + ['Kind -3', 'Kind -11'], 2):
                c = ws.cell(row=r, column=m_idx, value=group_summary[col_name].sum())
                c.font = total_font
                c.fill = total_fill
                c.border = thin_border
                c.alignment = Alignment(horizontal='center', vertical='center')
            
            return r + 3

        next_row = 1
        titel_toevoeging = " + ".join(selected_meals).upper()
        next_row = add_group_summary_to_sheet(ws_groepen, f"GROEPEN: HET BUFFET ({titel_toevoeging})", df_1300, next_row)
        add_group_summary_to_sheet(ws_groepen, f"GROEPEN: PLAD'O ({titel_toevoeging})", df_other, next_row)
        
        ws_groepen.column_dimensions['A'].width = 40
        for i in range(1 + len(selected_meals) + 2):
            ws_groepen.column_dimensions[get_column_letter(2 + i)].width = 12 

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 DOWNLOAD LIJSTEN",
            data=output,
            file_name=f"Lijsten_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Er ging iets mis tijdens het verwerken: {e}")
