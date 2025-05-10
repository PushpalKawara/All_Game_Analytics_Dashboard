# ========================== Step 1: Required Imports ========================== #
import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import datetime
import matplotlib.pyplot as plt
from io import BytesIO
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
import tempfile

# ========================== Step 2: Streamlit Config ========================== #
st.set_page_config(page_title="GAME PROGRESSION", layout="wide")
st.title("📊 GAME PROGRESSION Dashboard")

# ========================== Step 3: Core Functions (Updated Retention Formula) ========================== #
def merge_and_calculate(start_df, complete_df):
    merged = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')
    
    max_start_users = merged['START_USERS'].max()
    merged['RETENTION_%'] = (merged['COMPLETE_USERS'] / max_start_users) * 100  # Corrected formula
    
    merged['GAME_PLAY_DROP'] = ((merged['START_USERS'] - merged['COMPLETE_USERS']) / merged['START_USERS']) * 100
    merged['POPUP_DROP'] = ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1)) / merged['COMPLETE_USERS']) * 100
    merged['TOTAL_LEVEL_DROP'] = ((merged['START_USERS'] - merged['START_USERS'].shift(-1)) / merged['START_USERS']) * 100

    return merged.round(2)

# ========================== Step 4: Enhanced Charting Functions ========================== #
def create_charts(df, version, date_selected):
    charts = {}
    df_100 = df[df['LEVEL'] <= 100].copy()

    # Common settings
    xtick_labels = [f"$\\bf{{{x}}}$" if x % 5 == 0 else str(x) for x in range(1, 101)]
    
    # Retention Chart
    fig1, ax1 = plt.subplots(figsize=(15, 7))
    ax1.plot(df_100['LEVEL'], df_100['RETENTION_%'], color='#F57C00', linewidth=2)
    format_chart(ax1, "Retention Chart", version, date_selected, xtick_labels)
    charts['retention'] = fig1

    # Total Drop Chart
    fig2, ax2 = plt.subplots(figsize=(15, 6))
    ax2.bar(df_100['LEVEL'], df_100['TOTAL_LEVEL_DROP'], color='#EF5350')
    format_chart(ax2, "Total Level Drop Chart", version, date_selected, xtick_labels)
    charts['total_drop'] = fig2

    # Combo Drop Chart
    fig3, ax3 = plt.subplots(figsize=(15, 6))
    width = 0.4
    ax3.bar(df_100['LEVEL'] + width/2, df_100['GAME_PLAY_DROP'], width, color='#66BB6A', label='Game Play Drop')
    ax3.bar(df_100['LEVEL'] - width/2, df_100['POPUP_DROP'], width, color='#42A5F5', label='Popup Drop')
    format_chart(ax3, "Game Play & Popup Drop Chart", version, date_selected, xtick_labels)
    ax3.legend()
    charts['combo_drop'] = fig3

    return charts

def format_chart(ax, title, version, date_selected, xtick_labels):
    ax.set_xlim(1, 100)
    ax.set_xticks(np.arange(1, 101, 1))
    ax.set_xticklabels(xtick_labels, fontsize=6)
    ax.set_title(f"{title} | Version {version} | {date_selected.strftime('%d-%m-%Y')}", 
                fontsize=12, fontweight='bold')
    ax.grid(True, linestyle='--', linewidth=0.5)
    ax.tick_params(axis='x', labelsize=6)

# ========================== Step 5: Enhanced Excel Generation ========================== #
def generate_excel_report(processed_data, version, date_selected):
    wb = Workbook()
    wb.remove(wb.active)

    # Create main sheet
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_headers = ["Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
                   "Total Level Drop Count", "LEVEL_Start", "USERS_starts", 
                   "LEVEL_End", "USERS_END", "Link to Sheet"]
    main_sheet.append(main_headers)

    # Define styles
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(bold=True, color="000000")
    cell_alignment = Alignment(horizontal='center', vertical='center')
    highlight_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    highlight_font = Font(color="FF0000")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))

    for idx, (game_name, df) in enumerate(processed_data.items(), start=1):
        sheet = wb.create_sheet(game_name[:31])  # Excel sheet name limit
        
        # Create headers
        headers = [
            'Level', 'Start Users', 'Complete Users', 'Game Play Drop',
            'Popup Drop', 'Total Level Drop', 'Retention %', 'PLAY_TIME_AVG',
            'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM'
        ]
        sheet.append(headers)

        # Format headers
        for col in range(1, len(headers)+1):
            cell = sheet.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = cell_alignment
            cell.border = thin_border

        # Add data rows
        for _, row in df.iterrows():
            sheet.append([
                row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'],
                row['GAME_PLAY_DROP'], row['POPUP_DROP'], row['TOTAL_LEVEL_DROP'],
                row['RETENTION_%'], row.get('PLAY_TIME_AVG', 0),
                row.get('HINT_USED_SUM', 0), row.get('SKIPPED_SUM', 0),
                row.get('ATTEMPT_SUM', 0)
            ])

        # Apply formatting to data cells
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
            for cell in row:
                cell.alignment = cell_alignment
                cell.border = thin_border
                
                # Highlight problematic drops
                if cell.column in [4,5,6] and isinstance(cell.value, (int, float)):
                    if cell.value >= 3 or cell.value < 0:
                        cell.fill = highlight_fill
                        cell.font = highlight_font

        # Add charts as images
        add_charts_to_sheet(sheet, create_charts(df, version, date_selected))

        # Add hyperlink to main sheet
        main_sheet.append([
            idx, game_name,
            df['GAME_PLAY_DROP'].count(),
            df['POPUP_DROP'].count(),
            df['TOTAL_LEVEL_DROP'].count(),
            df['LEVEL'].min(), df['START_USERS'].max(),
            df['LEVEL'].max(), df['COMPLETE_USERS'].iloc[-1],
            f'=HYPERLINK("#{game_name[:31]}!A1","View {game_name}")'
        ])

    # Format main sheet
    for col in range(1, len(main_headers)+1):
        cell = main_sheet.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = cell_alignment

    return wb

def add_charts_to_sheet(sheet, charts):
    """Insert charts as images in Excel sheet"""
    img_size = (480, 280)  # Width, Height in pixels
    
    # Save charts to bytes
    retention_img = BytesIO()
    charts['retention'].savefig(retention_img, format='png', dpi=150, bbox_inches='tight')
    retention_img.seek(0)
    
    total_drop_img = BytesIO()
    charts['total_drop'].savefig(total_drop_img, format='png', dpi=150, bbox_inches='tight')
    total_drop_img.seek(0)
    
    combo_drop_img = BytesIO()
    charts['combo_drop'].savefig(combo_drop_img, format='png', dpi=150, bbox_inches='tight')
    combo_drop_img.seek(0)

    # Insert images into sheet
    img = Image(retention_img)
    img.width, img.height = img_size
    sheet.add_image(img, 'M2')
    
    img = Image(total_drop_img)
    img.width, img.height = img_size
    sheet.add_image(img, 'M35')
    
    img = Image(combo_drop_img)
    img.width, img.height = img_size
    sheet.add_image(img, 'M68')

# ========================== Step 6: Enhanced Streamlit UI ========================== #
def main():
    st.sidebar.header("Upload Files")
    start_files = st.sidebar.file_uploader("LEVEL_START Files", type=["csv", "xlsx"], accept_multiple_files=True)
    complete_files = st.sidebar.file_uploader("LEVEL_COMPLETE Files", type=["csv", "xlsx"], accept_multiple_files=True)

    version = st.sidebar.text_input("Game Version", "1.0.0")
    date_selected = st.sidebar.date_input("Analysis Date", datetime.date.today())

    if start_files and complete_files:
        with st.spinner("Processing files..."):
            processed_data = process_game_data(start_files, complete_files)

            if processed_data:
                # Generate and download Excel report
                wb = generate_excel_report(processed_data, version, date_selected)
                with tempfile.NamedTemporaryFile(delete=False) as tmp:
                    wb.save(tmp.name)
                    tmp.seek(0)
                    excel_data = tmp.read()

                st.download_button(
                    label="📥 Download Full Report",
                    data=excel_data,
                    file_name=f"Game_Analytics_{version}_{date_selected}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Display preview section
                selected_game = st.selectbox("Select Game to Preview", list(processed_data.keys()))
                st.subheader(f"📊 {selected_game} Analysis Preview")
                
                col1, col2 = st.columns([2, 3])
                with col1:
                    st.dataframe(processed_data[selected_game].style
                                .applymap(lambda x: 'background-color: yellow; color: red' 
                                        if isinstance(x, (int, float)) and (x >= 3 or x < 0) 
                                        else '', 
                                        subset=['GAME_PLAY_DROP', 'POPUP_DROP', 'TOTAL_LEVEL_DROP']))
                
                with col2:
                    charts = create_charts(processed_data[selected_game], version, date_selected)
                    tab1, tab2, tab3 = st.tabs(["Retention", "Total Drop", "Combo Drop"])
                    
                    with tab1:
                        st.pyplot(charts['retention'])
                    with tab2:
                        st.pyplot(charts['total_drop'])
                    with tab3:
                        st.pyplot(charts['combo_drop'])

if __name__ == "__main__":
    main()
