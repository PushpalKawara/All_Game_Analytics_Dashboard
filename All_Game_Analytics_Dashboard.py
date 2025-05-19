import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re

# Streamlit setup
st.set_page_config(page_title="Game Analytics Tool", layout="wide")
st.title("🎮 ALL_GAMES_ANALYZER")

# --------------------- Helper Functions ---------------------
def clean_level(level):
    if pd.isna(level):
        return 0
    return int(re.sub(r'\D', '', str(level)))

def process_files(start_df, complete_df):
    for df in [start_df, complete_df]:
        df['LEVEL'] = df['LEVEL'].apply(clean_level)
        df.sort_values('LEVEL', inplace=True)

    start_df = start_df.rename(columns={'USERS': 'START_USERS'})
    complete_df = complete_df.rename(columns={'USERS': 'COMPLETE_USERS'})

    merge_cols = ['GAME_ID', 'DIFFICULTY', 'LEVEL']
    merged = pd.merge(start_df, complete_df, on=merge_cols, how='outer')

    keep_cols = ['GAME_ID', 'DIFFICULTY', 'LEVEL', 'START_USERS', 'COMPLETE_USERS',
                 'PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPTS_SUM']
    for col in keep_cols:
        if col not in merged:
            merged[col] = 0

    merged = merged[keep_cols]
    merged['GAME_PLAY_DROP'] = ((merged['START_USERS'] - merged['COMPLETE_USERS']) / merged['START_USERS'].replace(0, np.nan)) * 100
    merged['POPUP_DROP'] = ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1)) / merged['COMPLETE_USERS'].replace(0, np.nan)) * 100
    merged['TOTAL_LEVEL_DROP'] = ((merged['START_USERS'] - merged['START_USERS'].shift(-1)) / merged['START_USERS'].replace(0, np.nan)) * 100
    merged['RETENTION_%'] = (merged['START_USERS'] / merged['START_USERS'].max()) * 100

    merged.fillna(0, inplace=True)
    return merged

def create_charts(df, game_name):
    charts = {}

    fig1, ax1 = plt.subplots(figsize=(12, 4))
    ax1.plot(df['LEVEL'], df['RETENTION_%'], color='#4CAF50')
    ax1.set_title(f"{game_name} - RETENTION_%", fontsize=10)
    charts['retention'] = fig1

    fig2, ax2 = plt.subplots(figsize=(12, 4))
    ax2.bar(df['LEVEL'], df['TOTAL_LEVEL_DROP'], color='#F44336')
    ax2.set_title(f"{game_name} - TOTAL_LEVEL_DROP", fontsize=10)
    charts['total_drop'] = fig2

    fig3, ax3 = plt.subplots(figsize=(12, 4))
    width = 0.35
    ax3.bar(df['LEVEL'] - width/2, df['GAME_PLAY_DROP'], width, label='GAME_PLAY_DROP')
    ax3.bar(df['LEVEL'] + width/2, df['POPUP_DROP'], width, label='POPUP_DROP')
    ax3.set_title(f"{game_name} - Drop", fontsize=10)
    ax3.legend()
    charts['combined_drop'] = fig3

    return charts

def add_charts_to_excel(worksheet, charts):
    img_positions = {'retention': 'M2', 'total_drop': 'N32', 'combined_drop': 'N65'}
    for chart_type in img_positions:
        img_data = BytesIO()
        charts[chart_type].savefig(img_data, format='png', dpi=150, bbox_inches='tight')
        img_data.seek(0)
        img = OpenpyxlImage(img_data)
        worksheet.add_image(img, img_positions[chart_type])
        plt.close(charts[chart_type])

def apply_sheet_formatting(sheet):
    sheet.freeze_panes = 'A1'
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    for col in sheet.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        sheet.column_dimensions[get_column_letter(col[0].column)].width = max_len + 2

def apply_conditional_formatting(sheet, num_rows):
    drop_columns = {'D', 'E', 'F'}
    red_scale = {
        '3': PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
        '7': PatternFill(start_color='FF9999', end_color='FF9999', fill_type='solid'),
        '10': PatternFill(start_color='FF6666', end_color='FF6666', fill_type='solid')
    }

    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            if cell.column_letter in drop_columns and isinstance(cell.value, (int, float)):
                val = cell.value
                if val >= 10:
                    cell.fill = red_scale['10']
                elif val >= 7:
                    cell.fill = red_scale['7']
                elif val >= 3:
                    cell.fill = red_scale['3']
                cell.font = Font(color="FFFFFF")

    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

def generate_excel(processed_data):
    wb = Workbook()
    wb.remove(wb.active)
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_headers = ["Index", "Sheet Name", "GAME_PLAY_DROP_Count", "POPUP_DROP_Count",
                    "TOTAL_LEVEL_DROP_Count", "LEVEL_Start", "USERS_starts",
                    "LEVEL_End", "USERS_END", "Link to Sheet"]
    main_sheet.append(main_headers)
    for col in main_sheet[1]:
        col.font = Font(bold=True, color="FFFFFF")
        col.fill = PatternFill("solid", fgColor="4F81BD")

    for idx, (game_id, df) in enumerate(processed_data.items(), start=1):
        sheet_name = f"{game_id}_{df['DIFFICULTY'].iloc[0]}"[:31]
        ws = wb.create_sheet(sheet_name)
        ws['A1'] = '=HYPERLINK("#MAIN_TAB!A1", "Back to Main")'
        ws['A1'].font = Font(color="0000FF", underline="single")

        headers = ["Level", "START_USERS", "COMPLETE_USERS", "GAME_PLAY_DROP", "POPUP_DROP",
                   "TOTAL_LEVEL_DROP", "RETENTION_%", "PLAY_TIME_AVG", "HINT_USED_SUM",
                   "SKIPPED_SUM", "ATTEMPTS_SUM"]
        ws.append(headers)

        for _, row in df.iterrows():
            ws.append([
                row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'],
                row['GAME_PLAY_DROP'], row['POPUP_DROP'], row['TOTAL_LEVEL_DROP'],
                row['RETENTION_%'], row.get('PLAY_TIME_AVG', 0), row.get('HINT_USED_SUM', 0),
                row.get('SKIPPED_SUM', 0), row.get('ATTEMPTS_SUM', 0)
            ])

        charts = create_charts(df, sheet_name)
        add_charts_to_excel(ws, charts)
        apply_sheet_formatting(ws)
        apply_conditional_formatting(ws, df.shape[0])

        main_row = [
            idx, sheet_name,
            sum(df['GAME_PLAY_DROP'] >= 0.03),
            sum(df['POPUP_DROP'] >= 0.03),
            sum(df['TOTAL_LEVEL_DROP'] >= 0.03),
            df['LEVEL'].min(), df['START_USERS'].max(),
            df['LEVEL'].max(), df['COMPLETE_USERS'].iloc[-1],
            f'=HYPERLINK("#{sheet_name}!A1", " Click to analyze")'
        ]
        main_sheet.append(main_row)

    for col in range(1, len(main_headers)+1):
        main_sheet.column_dimensions[get_column_letter(col)].width = 18

    return wb

# --------------------- Streamlit UI ---------------------
def main():
    st.sidebar.header("Upload CSV Files")

    uploaded_start = st.sidebar.file_uploader("Upload START Data CSV", type=["csv"])
    uploaded_complete = st.sidebar.file_uploader("Upload COMPLETE Data CSV", type=["csv"])

    if uploaded_start and uploaded_complete:
        start_df = pd.read_csv(uploaded_start)
        complete_df = pd.read_csv(uploaded_complete)

        if 'GAME_ID' in start_df and 'GAME_ID' in complete_df:
            processed_data = {}
            game_ids = set(start_df['GAME_ID'].unique()) & set(complete_df['GAME_ID'].unique())

            for gid in game_ids:
                start_subset = start_df[start_df['GAME_ID'] == gid]
                complete_subset = complete_df[complete_df['GAME_ID'] == gid]
                processed_df = process_files(start_subset, complete_subset)
                processed_data[gid] = processed_df

            wb = generate_excel(processed_data)
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            st.success("Excel report generated successfully!")
            st.download_button(label="📥 Download Excel Report",
                               data=buffer,
                               file_name="Game_Analytics_Report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("Both CSVs must include 'GAME_ID' column.")

if __name__ == "__main__":
    main()
