import os
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference

# Constants
STYLE_HEADER = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
STYLE_FONT_WHITE = Font(color='FFFFFF')
CENTRAL_ALIGN = Alignment(horizontal='center', vertical='center')

def clean_level(df):
    df['LEVEL'] = df['LEVEL'].str.replace(r'(?i)^level_?', '', regex=True).astype(int)
    return df.sort_values('LEVEL')

def process_files(start_file, complete_file):
    # Read and clean data
    df_start = pd.read_csv(start_file).rename(columns={'USERS': 'Start Users'})
    df_complete = pd.read_csv(complete_file).rename(columns={'USERS': 'Complete Users'})
    
    df_start = clean_level(df_start)
    df_complete = clean_level(df_complete)

    # Merge data
    merged = pd.merge(
        df_start[['GAME_ID', 'DIFFICULTY', 'LEVEL', 'Start Users']],
        df_complete[['GAME_ID', 'DIFFICULTY', 'LEVEL', 'Complete Users', 
                    'PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPTS_SUM']],
        on=['GAME_ID', 'DIFFICULTY', 'LEVEL'],
        how='outer'
    )

    # Add calculated columns
    merged['Game Play Drop'] = merged['Start Users'] - merged['Complete Users']
    merged['Popup Drop'] = merged['Start Users'] * 0.03
    merged['Total Level Drop'] = merged['Game Play Drop'] + merged['Popup Drop']
    merged['Retention %'] = (merged['Complete Users'] / merged['Start Users'] * 100).round(2)
    
    return merged

def apply_formatting(ws):
    # Freeze header and apply styles
    ws.freeze_panes = 'A2'
    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.fill = STYLE_HEADER
            cell.font = STYLE_FONT_WHITE
            cell.alignment = CENTRAL_ALIGN

    # Set central alignment for all cells
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = CENTRAL_ALIGN

    # Autofit columns
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

def create_charts(ws, max_row):
    # Retention Line Chart
    retention_chart = LineChart()
    data = Reference(ws, min_col=6, min_row=1, max_row=max_row, max_col=6)
    retention_chart.add_data(data, titles_from_data=True)
    retention_chart.title = "Retention % Trend"
    retention_chart.style = 13
    ws.add_chart(retention_chart, "N2")

    # Total Level Drop Bar Chart
    drop_chart = BarChart()
    data = Reference(ws, min_col=8, min_row=1, max_row=max_row, max_col=8)
    drop_chart.add_data(data, titles_from_data=True)
    drop_chart.title = "Total Level Drop"
    ws.add_chart(drop_chart, "N39")

    # Combined Drop Chart
    combined_chart = BarChart()
    data = Reference(ws, min_col=5, min_row=1, max_row=max_row, max_col=7)
    combined_chart.add_data(data, titles_from_data=True)
    combined_chart.title = "Game Play vs Popup Drops"
    ws.add_chart(combined_chart, "N70")

def create_main_tab(wb, all_sheets):
    main_ws = wb.create_sheet("MAIN_TAB", 0)
    headers = [
        "Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
        "Total Level Drop Count", "LEVEL_Start", "USERS_starts",
        "LEVEL_End", "USERS_END", "Link to Sheet"
    ]
    main_ws.append(headers)
    
    for idx, sheet_data in enumerate(all_sheets, 1):
        game_name = sheet_data['name']
        main_ws.append([
            idx,
            game_name,
            f"Count > 3%",
            f"Count > 3%",
            f"Count > 3%",
            f"Level {sheet_data['min_level']}",
            sheet_data['max_start'],
            f"Level {sheet_data['max_level']}",
            sheet_data['min_complete'],
            f'=HYPERLINK("##{game_name}!A1", "View {game_name}")'
        ])
    
    apply_formatting(main_ws)

def process_data(uploaded_files):
    wb = Workbook()
    del wb['Sheet']  # Remove default sheet
    
    all_sheets = []
    for start_file, complete_file in zip(uploaded_files[::2], uploaded_files[1::2]):
        df = process_files(start_file, complete_file)
        game_id = df['GAME_ID'].iloc[0]
        difficulty = df['DIFFICULTY'].iloc[0]
        sheet_name = f"{game_id}_{difficulty}"
        
        # Create new worksheet
        ws = wb.create_sheet(sheet_name)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # Apply formatting and charts
        apply_formatting(ws)
        create_charts(ws, len(df)+1)
        
        # Add backlink to main tab
        ws['A1'] = f'=HYPERLINK("##MAIN_TAB!A1", "Back to Main")'
        
        # Collect sheet metadata
        all_sheets.append({
            'name': sheet_name,
            'min_level': df['LEVEL'].min(),
            'max_level': df['LEVEL'].max(),
            'max_start': df['Start Users'].max(),
            'min_complete': df['Complete Users'].min()
        })
    
    create_main_tab(wb, all_sheets)
    return wb

# Streamlit UI
st.title("🎮 Game Analytics Dashboard")
uploaded_files = st.file_uploader("Upload LEVEL_START and LEVEL_COMPLETE CSVs", 
                                type="csv", accept_multiple_files=True)

if len(uploaded_files) % 2 == 0 and uploaded_files:
    with st.spinner('Processing data...'):
        wb = process_data(uploaded_files)
        wb.save("Consolidated.xlsx")
        
    with open("Consolidated.xlsx", "rb") as f:
        st.download_button("📥 Download Consolidated Report", f, 
                         file_name="Game_Analytics_Report.xlsx",
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.warning("Please upload matching pairs of LEVEL_START and LEVEL_COMPLETE files")
