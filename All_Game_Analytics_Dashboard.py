import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import re

st.set_page_config(page_title="Game Analytics Tool", layout="wide")
st.title("🎮 ALL GAMES ANALYZER")

# ========== Clean LEVEL ==========
def clean_level(level):
    if pd.isna(level):
        return 0
    return int(re.sub(r'\D', '', str(level)))

# ========== Process Files ==========
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

# ========== Create Charts ==========
def create_charts(df, game_name):
    charts = {}
    fig1, ax1 = plt.subplots()
    ax1.plot(df['LEVEL'], df['RETENTION_%'], color='#4CAF50')
    ax1.set_title(f"{game_name} - RETENTION_%")
    charts['retention'] = fig1

    fig2, ax2 = plt.subplots()
    ax2.bar(df['LEVEL'], df['TOTAL_LEVEL_DROP'], color='#F44336')
    ax2.set_title(f"{game_name} - TOTAL_LEVEL_DROP")
    charts['total_drop'] = fig2

    fig3, ax3 = plt.subplots()
    width = 0.35
    ax3.bar(df['LEVEL'] - width / 2, df['GAME_PLAY_DROP'], width, label='GAME_PLAY_DROP')
    ax3.bar(df['LEVEL'] + width / 2, df['POPUP_DROP'], width, label='POPUP_DROP')
    ax3.set_title(f"{game_name} - DROP")
    ax3.legend()
    charts['combined_drop'] = fig3

    return charts

# ========== Add Charts to Excel ==========
def add_charts_to_excel(worksheet, charts):
    img_positions = {'retention': 'M2', 'total_drop': 'N32', 'combined_drop': 'N65'}

    for name, fig in charts.items():
        img_data = BytesIO()
        fig.savefig(img_data, format='png', bbox_inches='tight')
        img_data.seek(0)
        img = OpenpyxlImage(img_data)
        worksheet.add_image(img, img_positions[name])
        plt.close(fig)

# ========== Sheet Formatting ==========
def apply_sheet_formatting(sheet):
    sheet.freeze_panes = 'A2'
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")

    for col in sheet.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        sheet.column_dimensions[get_column_letter(col[0].column)].width = max_len + 2

# ========== Generate Excel ==========
def generate_excel(processed_data):
    wb = Workbook()
    wb.remove(wb.active)
    main = wb.create_sheet("MAIN_TAB")
    headers = ["Index", "Sheet Name", "GAME_PLAY_DROP_Count", "POPUP_DROP_Count",
               "TOTAL_LEVEL_DROP_Count", "LEVEL_Start", "USERS_starts",
               "LEVEL_End", "USERS_END", "Link to Sheet"]
    main.append(headers)

    for col in main[1]:
        col.font = Font(bold=True, color="FFFFFF")
        col.fill = PatternFill("solid", fgColor="4F81BD")

    for idx, (game_id, df) in enumerate(processed_data.items(), start=1):
        sheet_name = f"{game_id}_{df['DIFFICULTY'].iloc[0]}"[:31]
        ws = wb.create_sheet(sheet_name)

        ws['A1'] = '=HYPERLINK("#MAIN_TAB!A1", "Back to Main")'
        ws['A1'].font = Font(color="0000FF", underline="single")

        headers = ["LEVEL", "START_USERS", "COMPLETE_USERS", "GAME_PLAY_DROP", "POPUP_DROP",
                   "TOTAL_LEVEL_DROP", "RETENTION_%", "PLAY_TIME_AVG", "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPTS_SUM"]
        ws.append(headers)

        for _, row in df.iterrows():
            ws.append([
                row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'], row['GAME_PLAY_DROP'],
                row['POPUP_DROP'], row['TOTAL_LEVEL_DROP'], row['RETENTION_%'],
                row.get('PLAY_TIME_AVG', 0), row.get('HINT_USED_SUM', 0),
                row.get('SKIPPED_SUM', 0), row.get('ATTEMPTS_SUM', 0)
            ])

        charts = create_charts(df, sheet_name)
        add_charts_to_excel(ws, charts)
        apply_sheet_formatting(ws)

        main.append([
            idx, sheet_name,
            sum(df['GAME_PLAY_DROP'] >= 3),
            sum(df['POPUP_DROP'] >= 3),
            sum(df['TOTAL_LEVEL_DROP'] >= 3),
            df['LEVEL'].min(), df['START_USERS'].max(),
            df['LEVEL'].max(), df['COMPLETE_USERS'].iloc[-1],
            f'=HYPERLINK("#{sheet_name}!A1", "Click to analyze")'
        ])

    return wb

# ========== STREAMLIT UI ==========
start_csv = st.file_uploader("Upload START_USERS CSV", type="csv")
complete_csv = st.file_uploader("Upload COMPLETE_USERS CSV", type="csv")

if start_csv and complete_csv:
    start_df = pd.read_csv(start_csv)
    complete_df = pd.read_csv(complete_csv)

    combined_df = process_files(start_df, complete_df)

    grouped = dict(tuple(combined_df.groupby(['GAME_ID'])))

    wb = generate_excel(grouped)

    # Download Excel
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("✅ Excel file generated with charts and analytics.")
    st.download_button("📥 Download Excel File", data=output, file_name="Game_Analytics_Report.xlsx")
