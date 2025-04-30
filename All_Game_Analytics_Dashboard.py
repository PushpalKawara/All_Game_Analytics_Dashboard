# ========================== Step 1: Required Imports & Folder Paths ========================== #
import streamlit as st
import pandas as pd
import numpy as np
import os
import re
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
import datetime
import matplotlib.pyplot as plt
from io import BytesIO
import tempfile

st.set_page_config(page_title="Game_Analytics", layout="wide")
st.title("📊 ALL_Game_Analytics Dashboard")

# ========================== Step 2: File Processing Functions ========================== #
def process_files(start_files, complete_files):
    all_data = {}

    # Create mapping of base filenames to file objects
    start_map = {os.path.splitext(f.name)[0]: f for f in start_files}
    complete_map = {os.path.splitext(f.name)[0]: f for f in complete_files}

    common_files = set(start_map.keys()) & set(complete_map.keys())

    for game_name in common_files:
        # Process start file
        start_df = pd.read_csv(start_map[game_name]) if start_map[game_name].name.endswith('.csv') else pd.read_excel(start_map[game_name])
        start_df = clean_start_file(start_df)

        # Process complete file
        complete_df = pd.read_csv(complete_map[game_name]) if complete_map[game_name].name.endswith('.csv') else pd.read_excel(complete_map[game_name])
        complete_df = clean_complete_file(complete_df)

        # Merge and calculate metrics
        merged_df = merge_and_calculate(start_df, complete_df)
        all_data[game_name] = merged_df

    return all_data

def clean_start_file(df):
    df.columns = df.columns.str.strip().str.upper()
    df = df.rename(columns={[col for col in df.columns if 'USER' in col][0]: 'START_USERS'})

    # Extract level numbers
    df['LEVEL'] = df['LEVEL'].astype(str).str.extract('(\d+)').astype(int)
    df = df.sort_values('LEVEL').drop_duplicates('LEVEL')
    return df[['LEVEL', 'START_USERS']]

def clean_complete_file(df):
    df.columns = df.columns.str.strip().str.upper()
    df = df.rename(columns={[col for col in df.columns if 'USER' in col][0]: 'COMPLETE_USERS'})

    # Extract level numbers
    df['LEVEL'] = df['LEVEL'].astype(str).str.extract('(\d+)').astype(int)
    df = df.sort_values('LEVEL').drop_duplicates('LEVEL')

    # Keep additional columns
    keep_cols = ['LEVEL', 'COMPLETE_USERS', 'PLAY_TIME_AVG',
                'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']
    return df[[col for col in keep_cols if col in df.columns]]

def merge_and_calculate(start_df, complete_df):
    merged = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')

    # Calculate metrics
    merged['GAME_PLAY_DROP'] = ((merged['START_USERS'] - merged['COMPLETE_USERS']) / merged['START_USERS']) * 100
    merged['POPUP_DROP'] = ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1)) / merged['COMPLETE_USERS']) * 100
    merged['TOTAL_LEVEL_DROP'] = ((merged['START_USERS'] - merged['START_USERS'].shift(-1)) / merged['START_USERS']) * 100
    merged['RETENTION_%'] = (merged['START_USERS'] / merged['START_USERS'].max()) * 100

    return merged.round(2)

# ========================== Step 3: Excel Generation Functions ========================== #
def create_excel_workbook(all_data, version, date_selected):
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Create Main Tab
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_sheet.append([
        "Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
        "Total Level Drop Count", "LEVEL_Start", "USERS_starts", "LEVEL_End", "USERS_END", "Link to Sheet"
    ])

    # Process each game
    for idx, (game_name, df) in enumerate(all_data.items(), start=1):
        sheet = wb.create_sheet(game_name)
        add_game_sheet_content(sheet, df)
        add_main_sheet_row(main_sheet, idx, game_name, df)

    format_main_sheet(main_sheet)
    return wb

def add_game_sheet_content(sheet, df):
    # Add hyperlink back to main tab
    sheet['A1'] = 'Back to MAIN_TAB'
    sheet['A1'].hyperlink = "#'MAIN_TAB!A1'"

    # Write headers
    headers = [
        "Level", "Start Users", "Complete Users", "Game Play Drop",
        "Popup Drop", "Total Level Drop", "Retention %", "PLAY_TIME_AVG",
        "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPT_SUM"
    ]
    sheet.append(headers)

    # Write data
    for _, row in df.iterrows():
        sheet.append([
            row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'],
            row['GAME_PLAY_DROP'], row['POPUP_DROP'], row['TOTAL_LEVEL_DROP'],
            row['RETENTION_%'], row.get('PLAY_TIME_AVG', 0),
            row.get('HINT_USED_SUM', 0), row.get('SKIPPED_SUM', 0),
            row.get('ATTEMPT_SUM', 0)
        ])

    # Add charts
    add_charts(sheet, df)

def add_main_sheet_row(main_sheet, idx, game_name, df):
    start_users = df['START_USERS'].max()
    end_users = df['COMPLETE_USERS'].iloc[-1]

    main_sheet.append([
        idx, game_name,
        df['GAME_PLAY_DROP'].count(),
        df['POPUP_DROP'].count(),
        df['TOTAL_LEVEL_DROP'].count(),
        df['LEVEL'].min(), start_users,
        df['LEVEL'].max(), end_users,
        f'=HYPERLINK("#{game_name}!A1","Click to view {game_name}")'
    ])

# ========================== Step 4: Formatting Functions ========================== #
def format_main_sheet(sheet):
    # Set column widths
    col_widths = [8, 15, 18, 15, 18, 12, 12, 12, 12, 20]
    for i, width in enumerate(col_widths, 1):
        sheet.column_dimensions[get_column_letter(i)].width = width

    # Apply header styling
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # Freeze header row
    sheet.freeze_panes = 'A2'

def apply_sheet_formatting(sheet):
    # Header formatting
    for cell in sheet[2]:  # Data starts at row 2
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
        cell.alignment = Alignment(horizontal='center')

    # Conditional formatting for drop columns
    red_fill = PatternFill(start_color="FFEEEE", end_color="FFEEEE", fill_type="solid")
    for col in ['D', 'E', 'F']:  # Game Play, Popup, Total Level Drops
        for cell in sheet[col][1:]:  # Skip header
            if cell.value and cell.value >= 3:
                cell.fill = red_fill
                cell.font = Font(color="FF0000")

    # Auto-fit columns
    for col in sheet.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        sheet.column_dimensions[col[0].column_letter].width = max_length + 2

# ========================== Step 5: Streamlit UI ========================== #
def main():
    st.sidebar.header("Upload Folders")
    start_files = st.sidebar.file_uploader("LEVEL_START Folder", type=["csv", "xlsx"], accept_multiple_files=True)
    complete_files = st.sidebar.file_uploader("LEVEL_COMPLETE Folder", type=["csv", "xlsx"], accept_multiple_files=True)

    version = st.sidebar.text_input("Game Version", "1.0.0")
    date_selected = st.sidebar.date_input("Analysis Date", datetime.date.today())

    if start_files and complete_files:
        with st.spinner("Processing files..."):
            all_data = process_files(start_files, complete_files)

            # Generate Excel
            wb = create_excel_workbook(all_data, version, date_selected)

            # Save to temporary file
            with tempfile.NamedTemporaryFile(delete=False) as tmp:
                wb.save(tmp.name)
                tmp.seek(0)
                data = tmp.read()

            # Download button
            st.download_button(
                label="📥 Download Consolidated Report",
                data=data,
                file_name=f"Game_Analytics_{version}_{date_selected}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Show preview
            st.subheader("Preview of Processed Data")
            selected_game = st.selectbox("Select Game to Preview", list(all_data.keys()))
            st.dataframe(all_data[selected_game])

if __name__ == "__main__":
    main()
