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
from openpyxl.styles import Font, Alignment, PatternFill, numbers
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.formatting.rule import CellIsRule
import tempfile

# ========================== Step 2: Streamlit Config ========================== #
st.set_page_config(page_title="GAME PROGRESSION", layout="wide")
st.title("📊 GAME PROGRESSION Dashboard")

# ========================== Step 3: Core Functions ========================== #
def process_game_data(start_files, complete_files):
    processed_games = {}
    start_map = {os.path.splitext(f.name)[0].upper(): f for f in start_files}
    complete_map = {os.path.splitext(f.name)[0].upper(): f for f in complete_files}
    common_games = set(start_map.keys()) & set(complete_map.keys())

    for idx, game_name in enumerate(sorted(common_games), start=1):
        start_df = load_and_clean_file(start_map[game_name], is_start_file=True)
        complete_df = load_and_clean_file(complete_map[game_name], is_start_file=False)

        if start_df is not None and complete_df is not None:
            merged_df = merge_and_calculate(start_df, complete_df)
            display_name = f"ID_{idx}_{game_name}"
            processed_games[display_name] = merged_df

    return processed_games

def load_and_clean_file(file_obj, is_start_file=True):
    try:
        df = pd.read_csv(file_obj) if file_obj.name.endswith('.csv') else pd.read_excel(file_obj)
        df.columns = df.columns.str.strip().str.upper()

        level_col = next((col for col in df.columns if 'LEVEL' in col), None)
        if level_col:
            df['LEVEL'] = df[level_col].astype(str).str.extract(r'(\d+)').astype(int)

        user_col = next((col for col in df.columns if 'USER' in col), None)
        if user_col:
            new_name = 'START_USERS' if is_start_file else 'COMPLETE_USERS'
            df = df.rename(columns={user_col: new_name})

        if is_start_file:
            df = df[['LEVEL', 'START_USERS']].drop_duplicates().sort_values('LEVEL')
        else:
            keep_cols = ['LEVEL', 'COMPLETE_USERS', 'PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']
            df = df.reindex(columns=keep_cols, fill_value=0)

        return df.sort_values('LEVEL')
    except Exception as e:
        st.error(f"Error processing {file_obj.name}: {str(e)}")
        return None

def merge_and_calculate(start_df, complete_df):
    merged = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')
    merged.fillna(0, inplace=True)

    # Calculate percentages as decimals
    merged['GAME_PLAY_DROP'] = (merged['START_USERS'] - merged['COMPLETE_USERS']) / merged['START_USERS'].replace(0, np.nan)
    merged['POPUP_DROP'] = 0.03  # 3% as per requirement
    merged['TOTAL_LEVEL_DROP'] = merged['GAME_PLAY_DROP'] + merged['POPUP_DROP']
    merged['RETENTION_%'] = merged['COMPLETE_USERS'] / merged['START_USERS'].replace(0, np.nan)

    # Handle divisions
    merged.replace([np.inf, -np.inf], np.nan, inplace=True)
    merged.fillna(0, inplace=True)
    return merged

# ========================== Step 4: Excel Generation ========================== #
def generate_excel_report(processed_data, version, date_selected):
    wb = Workbook()
    wb.remove(wb.active)

    # Create MAIN_TAB sheet
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_headers = [
        "Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
        "Total Level Drop Count", "LEVEL_Start", "USERS_starts", "LEVEL_End", "USERS_END", "Link to Sheet"
    ]
    main_sheet.append(main_headers)

    # Process each game
    for idx, (game_name, df) in enumerate(processed_data.items(), start=1):
        sheet = wb.create_sheet(game_name[:31])
        sheet.append([
            'Level', 'Start Users', 'Complete Users', 'Game Play Drop',
            'Popup Drop', 'Total Level Drop', 'Retention %',
            'PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM'
        ])

        # Add data rows
        for _, row in df.iterrows():
            sheet.append([
                row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'],
                row['GAME_PLAY_DROP'], row['POPUP_DROP'], row['TOTAL_LEVEL_DROP'],
                row['RETENTION_%'],
                row['PLAY_TIME_AVG'], row['HINT_USED_SUM'], row['SKIPPED_SUM'], row['ATTEMPT_SUM']
            ])

        # Add backlink to MAIN_TAB
        sheet['A1'] = f'=HYPERLINK("#MAIN_TAB!A1", "Back to Main")'

        # MAIN_TAB data
        game_play_count = (df['GAME_PLAY_DROP'] >= 0.03).sum()
        popup_count = (df['POPUP_DROP'] >= 0.03).sum()
        total_count = (df['TOTAL_LEVEL_DROP'] >= 0.03).sum()
        start_level = df['LEVEL'].min()
        start_users = df.loc[df['LEVEL'] == start_level, 'START_USERS'].iloc[0]
        end_level = df['LEVEL'].max()
        end_users = df.loc[df['LEVEL'] == end_level, 'COMPLETE_USERS'].iloc[0]

        main_sheet.append([
            idx, game_name, game_play_count, popup_count, total_count,
            start_level, start_users, end_level, end_users,
            f'=HYPERLINK("#{game_name[:31]}!A1", "View {game_name}")'
        ])

    format_workbook(wb)
    return wb

def format_workbook(wb):
    # Define styles
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD")
    header_font = Font(bold=True, color="FFFFFF")
    percent_format = '0.00%'
    red_fills = {
        3: PatternFill(start_color="FFC7CE", end_color="FFC7CE"),
        5: PatternFill(start_color="FF6666", end_color="FF6666"),
        10: PatternFill(start_color="FF0000", end_color="FF0000")
    }

    for sheet in wb:
        if sheet.title == 'MAIN_TAB':
            sheet.freeze_panes = 'A2'
            for cell in sheet[1]:
                cell.font = header_font
                cell.fill = header_fill
            continue

        # Format game sheets
        sheet.freeze_panes = 'A2'
        for cell in sheet[1]:
            cell.font = header_font
            cell.fill = header_fill

        # Apply percentage format and conditional formatting
        for col in ['D', 'E', 'F', 'G']:
            col_idx = column_index_from_string(col)
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=col_idx, max_col=col_idx):
                for cell in row:
                    cell.number_format = percent_format

        # Conditional formatting for drops
        for col in ['D', 'E', 'F']:
            col_letter = get_column_letter(column_index_from_string(col))
            range_str = f"{col_letter}2:{col_letter}{sheet.max_row}"
            for threshold in [10, 5, 3]:
                sheet.conditional_formatting.add(range_str, CellIsRule(
                    operator='greaterThanOrEqual',
                    formula=[str(threshold/100)],
                    fill=red_fills[threshold]
                ))

        # Autofit columns
        for column in sheet.columns:
            max_length = max(len(str(cell.value)) for cell in column)
            sheet.column_dimensions[get_column_letter(column[0].column)].width = max_length + 2

# ========================== Step 5: Streamlit UI ========================== #
def main():
    st.sidebar.header("Upload Files")
    start_files = st.sidebar.file_uploader("LEVEL_START Files", type=["csv", "xlsx"], accept_multiple_files=True)
    complete_files = st.sidebar.file_uploader("LEVEL_COMPLETE Files", type=["csv", "xlsx"], accept_multiple_files=True)

    version = st.sidebar.text_input("Game Version", "1.0.0")
    date_selected = st.sidebar.date_input("Analysis Date", datetime.date.today())

    if start_files and complete_files:
        processed_data = process_game_data(start_files, complete_files)
        if processed_data:
            wb = generate_excel_report(processed_data, version, date_selected)
            with tempfile.NamedTemporaryFile(delete=False) as tmp:
                wb.save(tmp.name)
                with open(tmp.name, "rb") as f:
                    st.download_button(
                        label="📥 Download Consolidated Report",
                        data=f,
                        file_name=f"Game_Analytics_{version}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

if __name__ == "__main__":
    main()
