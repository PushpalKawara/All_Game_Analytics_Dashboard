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
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import tempfile


# ========================== Step 3: Main App ========================== #
st.set_page_config(page_title="GAME PROGRESSION", layout="wide")
st.title("📊 GAME PROGRESSION Dashboard")

# ========================== Step 4: Core Processing Functions ========================== #
def process_game_data(start_files, complete_files):
    processed_games = {}

    # Create file mappings
    start_map = {os.path.splitext(f.name)[0].upper(): f for f in start_files}
    complete_map = {os.path.splitext(f.name)[0].upper(): f for f in complete_files}

    common_games = set(start_map.keys()) & set(complete_map.keys())

    for game_name in common_games:
        # Process start file
        start_df = load_and_clean_file(start_map[game_name], is_start_file=True)
        complete_df = load_and_clean_file(complete_map[game_name], is_start_file=False)

        if start_df is not None and complete_df is not None:
            merged_df = merge_and_calculate(start_df, complete_df)
            processed_games[game_name] = merged_df

    return processed_games

def load_and_clean_file(file_obj, is_start_file=True):
    try:
        # Read file
        if file_obj.name.endswith('.csv'):
            df = pd.read_csv(file_obj)
        else:
            df = pd.read_excel(file_obj)

        # Clean columns
        df.columns = df.columns.str.strip().str.upper()

        # Extract level
        level_col = next((col for col in df.columns if 'LEVEL' in col), None)
        if level_col:
            df['LEVEL'] = df[level_col].astype(str).str.extract('(\d+)').astype(int)

        # Handle user columns
        user_col = next((col for col in df.columns if 'USER' in col), None)
        if user_col:
            df = df.rename(columns={user_col: 'START_USERS' if is_start_file else 'COMPLETE_USERS'})

        # Select relevant columns
        if is_start_file:
            df = df[['LEVEL', 'START_USERS']].drop_duplicates().sort_values('LEVEL')
        else:
            keep_cols = ['LEVEL', 'COMPLETE_USERS', 'PLAY_TIME_AVG',
                        'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']
            df = df[[col for col in keep_cols if col in df.columns]]

        return df.dropna().sort_values('LEVEL')

    except Exception as e:
        st.error(f"Error processing {file_obj.name}: {str(e)}")
        return None

def merge_and_calculate(start_df, complete_df):
    merged = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')

    # Calculate metrics
    merged['GAME_PLAY_DROP'] = ((merged['START_USERS'] - merged['COMPLETE_USERS']) / merged['START_USERS']) * 100
    merged['POPUP_DROP'] = ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1)) / merged['COMPLETE_USERS']) * 100
    merged['TOTAL_LEVEL_DROP'] = ((merged['START_USERS'] - merged['START_USERS'].shift(-1)) / merged['START_USERS']) * 100
    merged['RETENTION_%'] = (merged['START_USERS'] / merged['START_USERS'].max()) * 100

    return merged.round(2)

# ========================== Step 5: Chart Generation ========================== #
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
    ax3.bar(df_100['LEVEL'] + width/2, df_100['GAME_PLAY_DROP'], width, color='#66BB6A')
    ax3.bar(df_100['LEVEL'] - width/2, df_100['POPUP_DROP'], width, color='#42A5F5')
    format_chart(ax3, "Game Play & Popup Drop Chart", version, date_selected)
    charts['combo_drop'] = fig3

    return charts

def format_chart(ax, title, version, date_selected):
    ax.set_xlim(1, 100)
    ax.set_xticks(np.arange(1, 101, 1))
    ax.set_xticklabels([f"$\\bf{{{x}}}$" if x%5==0 else str(x) for x in range(1, 101)], fontsize=6)
    ax.set_title(f"{title} | Version {version} | {date_selected.strftime('%d-%m-%Y')}", fontsize=12, fontweight='bold')
    ax.grid(True, linestyle='--', linewidth=0.5)
    ax.tick_params(axis='x', labelsize=6)

# ========================== Step 6: Excel Generation ========================== #
def generate_excel_report(processed_data, version, date_selected):
    wb = Workbook()
    wb.remove(wb.active)

    # Create Main Tab
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_sheet.append([
        "Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
        "Total Level Drop Count", "LEVEL_Start", "USERS_starts", "LEVEL_End", "USERS_END", "Link to Sheet"
    ])

    # Process each game
    for idx, (game_name, df) in enumerate(processed_data.items(), start=1):
        # Create game sheet
        sheet = wb.create_sheet(game_name)
        sheet.append([
            "Level", "Start Users", "Complete Users", "Game Play Drop",
            "Popup Drop", "Total Level Drop", "Retention %", "PLAY_TIME_AVG",
            "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPT_SUM"
        ])

        # Add data
        for _, row in df.iterrows():
            sheet.append([
                row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'],
                row['GAME_PLAY_DROP'], row['POPUP_DROP'], row['TOTAL_LEVEL_DROP'],
                row['RETENTION_%'], row.get('PLAY_TIME_AVG', 0),
                row.get('HINT_USED_SUM', 0), row.get('SKIPPED_SUM', 0),
                row.get('ATTEMPT_SUM', 0)
            ])

        # Add charts
        charts = create_charts(df, version, date_selected)
        add_charts_to_sheet(sheet, charts)

        # Add main sheet entry
        main_sheet.append([
            idx, game_name,
            df['GAME_PLAY_DROP'].count(),
            df['POPUP_DROP'].count(),
            df['TOTAL_LEVEL_DROP'].count(),
            df['LEVEL'].min(), df['START_USERS'].max(),
            df['LEVEL'].max(), df['COMPLETE_USERS'].iloc[-1],
            f'=HYPERLINK("#{game_name}!A1","Click to view {game_name}")'
        ])

    # Format workbook
    format_workbook(wb)
    return wb

def add_charts_to_sheet(sheet, charts):
    # This function would need actual chart image insertion logic
    # Placeholder for demonstration
    sheet['M1'] = "Retention Chart →"
    sheet['M35'] = "Total Drop Chart →"
    sheet['M65'] = "Combo Drop Chart →"

def format_workbook(wb):
    for sheet in wb:
        # Header formatting
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4F81BD")
            cell.alignment = Alignment(horizontal='center')

        # Auto-fit columns
        for col in sheet.columns:
            max_length = max(len(str(cell.value)) for cell in col)
            sheet.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

        # Conditional formatting
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE")
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                if cell.column_letter in ['D', 'E', 'F'] and isinstance(cell.value, (int, float)):
                    if cell.value >= 3:
                        cell.fill = red_fill

# ========================== Step 7: Streamlit UI ========================== #
def main():
    st.sidebar.header("Upload Folders")
    start_files = st.sidebar.file_uploader("LEVEL_START Files", type=["csv", "xlsx"], accept_multiple_files=True)
    complete_files = st.sidebar.file_uploader("LEVEL_COMPLETE Files", type=["csv", "xlsx"], accept_multiple_files=True)

    version = st.sidebar.text_input("Game Version", "1.0.0")
    date_selected = st.sidebar.date_input("Analysis Date", datetime.date.today())

    if start_files and complete_files:
        with st.spinner("Processing files..."):
            processed_data = process_game_data(start_files, complete_files)

            if processed_data:
                # Generate Excel report
                wb = generate_excel_report(processed_data, version, date_selected)

                # Save to temporary file
                with tempfile.NamedTemporaryFile(delete=False) as tmp:
                    wb.save(tmp.name)
                    tmp.seek(0)
                    excel_data = tmp.read()

                # Download button
                st.download_button(
                    label="📥 Download Full Report",
                    data=excel_data,
                    file_name=f"Game_Analytics_{version}_{date_selected}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Show preview
                selected_game = st.selectbox("Select Game to Preview", list(processed_data.keys()))
                st.dataframe(processed_data[selected_game])

                # Show sample charts
                st.pyplot(create_charts(processed_data[selected_game], version, date_selected)['retention'])

if __name__ == "__main__":
    main()
