# ========================== Step 1: Required Imports ==========================

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
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage
import tempfile

# ========================== Step 2: Streamlit Config ==========================

st.set_page_config(page_title="GAME PROGRESSION", layout="wide")
st.title("📊 GAME PROGRESSION Dashboard")

# ========================== Step 3: Core Functions ==========================

def process_game_data(start_files, complete_files):
    processed_games = {}

    start_map = {os.path.splitext(f.name)[0].upper(): f for f in start_files}
    complete_map = {os.path.splitext(f.name)[0].upper(): f for f in complete_files}

    common_games = set(start_map.keys()) & set(complete_map.keys())

    for game_name in common_games:
        start_df = load_and_clean_file(start_map[game_name], is_start_file=True)
        complete_df = load_and_clean_file(complete_map[game_name], is_start_file=False)

        if start_df is not None and complete_df is not None:
            merged_df = merge_and_calculate(start_df, complete_df)
            processed_games[game_name] = merged_df

    return processed_games

def load_and_clean_file(file_obj, is_start_file=True):
    try:
        df = pd.read_csv(file_obj) if file_obj.name.endswith('.csv') else pd.read_excel(file_obj)
        df.columns = df.columns.str.strip().str.upper()

        level_col = next((col for col in df.columns if 'LEVEL' in col), None)
        if level_col:
            # Handle non-numeric level values safely
            df['LEVEL'] = pd.to_numeric(
                df[level_col].astype(str).str.extract('(\d+)', expand=False),
                errors='coerce'
            ).dropna().astype(int)

        user_col = next((col for col in df.columns if 'USER' in col), None)
        if user_col:
            # Fixed syntax error in rename operation
            new_name = 'START_USERS' if is_start_file else 'COMPLETE_USERS'
            df = df.rename(columns={user_col: new_name})

        if is_start_file:
            df = df[['LEVEL', 'START_USERS']].drop_duplicates().sort_values('LEVEL')
        else:
            keep_cols = ['LEVEL', 'COMPLETE_USERS', 'PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']
            df = df[[col for col in keep_cols if col in df.columns]]

        return df.dropna().sort_values('LEVEL')
    except Exception as e:
        st.error(f"Error processing {file_obj.name}: {str(e)}")
        return None

def merge_and_calculate(start_df, complete_df):
    merged = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')

    # Handle division by zero and NaN values
    with np.errstate(divide='ignore', invalid='ignore'):
        merged['GAME_PLAY_DROP'] = np.where(
            merged['START_USERS'] > 0,
            ((merged['START_USERS'] - merged['COMPLETE_USERS']) / merged['START_USERS']) * 100,
            0
        )

        merged['POPUP_DROP'] = np.where(
            merged['COMPLETE_USERS'] > 0,
            ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1)) / merged['COMPLETE_USERS']) * 100,
            0
        )

        merged['TOTAL_LEVEL_DROP'] = np.where(
            merged['START_USERS'] > 0,
            ((merged['START_USERS'] - merged['START_USERS'].shift(-1)) / merged['START_USERS']) * 100,
            0
        )

    merged['RETENTION_%'] = (merged['START_USERS'] / merged['START_USERS'].max()) * 100

    return merged.round(2).fillna(0)

# ========================== Step 4: Charting Functions ==========================

def create_charts(df, version, date_selected):
    charts = {}
    df_100 = df[df['LEVEL'] <= 100].copy()

    # Retention Chart
    fig1, ax1 = plt.subplots(figsize=(15, 7))
    ax1.plot(df_100['LEVEL'], df_100['RETENTION_%'], color='#F57C00', linewidth=2)
    format_chart(ax1, "Retention Chart", version, date_selected)
    charts['retention'] = fig1

    # Total Drop Chart
    fig2, ax2 = plt.subplots(figsize=(15, 6))
    ax2.bar(df_100['LEVEL'], df_100['TOTAL_LEVEL_DROP'], color='#EF5350')
    format_chart(ax2, "Total Level Drop Chart", version, date_selected)
    charts['total_drop'] = fig2

    # Combo Drop Chart
    fig3, ax3 = plt.subplots(figsize=(15, 6))
    width = 0.4
    ax3.bar(df_100['LEVEL'] + width/2, df_100['GAME_PLAY_DROP'], width, color='#66BB6A', label='Game Play Drop')
    ax3.bar(df_100['LEVEL'] - width/2, df_100['POPUP_DROP'], width, color='#42A5F5', label='Popup Drop')
    format_chart(ax3, "Game Play & Popup Drop Chart", version, date_selected)
    ax3.legend()
    charts['combo_drop'] = fig3

    return charts

def format_chart(ax, title, version, date_selected):
    ax.set_xlim(1, 100)
    ax.set_xticks(np.arange(1, 101, 1))
    ax.set_xticklabels([f"$\bf{{{x}}}$" if x % 5 == 0 else str(x) for x in range(1, 101)], fontsize=6)
    ax.set_title(f"{title} | Version {version} | {date_selected.strftime('%d-%m-%Y')}",
                fontsize=12, fontweight='bold')
    ax.grid(True, linestyle='--', linewidth=0.5)
    ax.tick_params(axis='x', labelsize=6)

# ========================== Step 5: Excel Generation ==========================

def generate_excel_report(processed_data, version, date_selected):
    wb = Workbook()
    wb.remove(wb.active)

    # Main sheet setup
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_headers = [
        "Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
        "Total Level Drop Count", "LEVEL_Start", "USERS_starts", "LEVEL_End", "USERS_END", "Link to Sheet"
    ]
    main_sheet.append(main_headers)

    for idx, (game_name, df) in enumerate(processed_data.items(), start=1):
        sheet_name = game_name[:31].replace(':', '_')  # Sanitize sheet name
        sheet = wb.create_sheet(sheet_name)
        headers = [
            "Level", "Start Users", "Complete Users", "Game Play Drop",
            "Popup Drop", "Total Level Drop", "Retention %", "PLAY_TIME_AVG",
            "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPT_SUM"
        ]
        sheet.append(headers)

        # Add data rows
        for _, row in df.iterrows():
            sheet.append([
                row['LEVEL'],
                row['START_USERS'],
                row['COMPLETE_USERS'],
                row['GAME_PLAY_DROP'],
                row['POPUP_DROP'],
                row['TOTAL_LEVEL_DROP'],
                row['RETENTION_%'],
                row.get('PLAY_TIME_AVG', 0),
                row.get('HINT_USED_SUM', 0),
                row.get('SKIPPED_SUM', 0),
                row.get('ATTEMPT_SUM', 0)
            ])

        # Add charts as images
        charts = create_charts(df, version, date_selected)
        add_charts_to_sheet(sheet, charts)

        # Main sheet data
        main_sheet.append([
            idx,
            game_name,
            df['GAME_PLAY_DROP'].count(),
            df['POPUP_DROP'].count(),
            df['TOTAL_LEVEL_DROP'].count(),
            df['LEVEL'].min(),
            df['START_USERS'].max(),
            df['LEVEL'].max(),
            df['COMPLETE_USERS'].iloc[-1],
            f'=HYPERLINK("#{sheet_name}!A1","Click to view {game_name}")'
        ])

    format_workbook(wb)
    return wb

def add_charts_to_sheet(sheet, charts):
    # Save charts to bytes and insert into Excel
    chart_positions = {
        'retention': ('M1', (15, 7)),
        'total_drop': ('M35', (15, 6)),
        'combo_drop': ('M65', (15, 6))
    }

    for chart_name, (position, size) in chart_positions.items():
        img = BytesIO()
        charts[chart_name].savefig(img, format='png', dpi=150, bbox_inches='tight')
        img.seek(0)

        excel_img = OpenpyxlImage(img)
        # Adjust image dimensions (width, height in pixels)
        excel_img.width = 800
        excel_img.height = 400 if chart_name == 'retention' else 350
        sheet.add_image(excel_img, position)
        plt.close(charts[chart_name])

def format_workbook(wb):
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    center_align = Alignment(horizontal='center', vertical='center')

    for sheet in wb:
        # Header formatting
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        # Column widths and formatting
        for col in sheet.columns:
            max_length = max(
                (len(str(cell.value)) for cell in col
                if cell.value is not None
            ), default=0)
            sheet.column_dimensions[get_column_letter(col[0].column)].width = min(max_length + 2, 50)

            # Highlight drops over 3%
            if col[0].column in [4, 5, 6]:  # D, E, F columns
                for cell in col[1:]:
                    if isinstance(cell.value, (int, float)) and cell.value >= 3:
                        cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE")

# ========================== Step 6: Streamlit UI ==========================

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
                wb = generate_excel_report(processed_data, version, date_selected)

                with tempfile.NamedTemporaryFile() as tmp:
                    wb.save(tmp.name)
                    tmp.seek(0)
                    excel_data = tmp.read()

                st.download_button(
                    label="📥 Download Full Report",
                    data=excel_data,
                    file_name=f"Game_Analytics_{version}_{date_selected.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                selected_game = st.selectbox("Select Game to Preview", list(processed_data.keys()))
                st.dataframe(processed_data[selected_game])

                chart_type = st.selectbox("Select Chart Type", ["Retention", "Total Drop", "Combo Drop"])
                chart_key = f'{chart_type.lower().replace(" ", "_")}'
                fig = create_charts(processed_data[selected_game], version, date_selected)[chart_key]
                st.pyplot(fig)
                plt.close(fig)

# ========================== Entry Point ==========================

if __name__ == "__main__":
    main()
