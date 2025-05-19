# game_data_excel_export.py

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import re
import tempfile

st.set_page_config(page_title="Game Analytics Tool", layout="wide")
st.title("🎮 Game Level Data Analyzer")

def clean_level(level):
    if pd.isna(level):
        return 0
    return int(re.sub(r'\D', '', str(level)))

def process_files(start_df, complete_df):
    for df in [start_df, complete_df]:
        df['LEVEL'] = df['LEVEL'].apply(clean_level)
        df.sort_values('LEVEL', inplace=True)

    start_df = start_df.rename(columns={'USERS': 'Start Users'})
    complete_df = complete_df.rename(columns={'USERS': 'Complete Users'})

    merge_cols = ['GAME_ID', 'DIFFICULTY', 'LEVEL']
    merged = pd.merge(start_df, complete_df, on=merge_cols, how='outer', suffixes=('_start', '_complete'))

    keep_cols = ['GAME_ID', 'DIFFICULTY', 'LEVEL', 'Start Users', 'Complete Users',
                 'PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPTS_SUM']
    merged = merged[keep_cols]

    merged['Game Play Drop'] = ((merged['Start Users'] - merged['Complete Users']) / merged['Start Users'].replace(0, np.nan)) * 100
    merged['Popup Drop'] = ((merged['Complete Users'] - merged['Start Users'].shift(-1)) / merged['Complete Users'].replace(0, np.nan)) * 100
    merged['Total Level Drop'] = ((merged['Start Users'] - merged['Start Users'].shift(-1)) / merged['Start Users'].replace(0, np.nan)) * 100
    merged['Retention %'] = (merged['Start Users'] / merged['Start Users'].max()) * 100

    for col in ['Game Play Drop', 'Popup Drop', 'Total Level Drop', 'Retention %', 'PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPTS_SUM']:
        merged[col] = merged[col].apply(lambda x: round(x, 2) if not pd.isna(x) else 0)

    merged.fillna("null", inplace=True)
    return merged

def create_charts(df, game_name):
    charts = {}
    fig1, ax1 = plt.subplots(figsize=(12, 4))
    ax1.plot(df['LEVEL'], df['Retention %'], color='#4CAF50')
    ax1.set_title(f"{game_name} - Retention %", fontsize=10)
    charts['retention'] = fig1

    fig2, ax2 = plt.subplots(figsize=(12, 4))
    ax2.bar(df['LEVEL'], df['Total Level Drop'], color='#F44336')
    ax2.set_title(f"{game_name} - Total Level Drop", fontsize=10)
    charts['total_drop'] = fig2

    fig3, ax3 = plt.subplots(figsize=(12, 4))
    width = 0.35
    ax3.bar(df['LEVEL'] - width/2, df['Game Play Drop'], width, label='Game Play Drop')
    ax3.bar(df['LEVEL'] + width/2, df['Popup Drop'], width, label='Popup Drop')
    ax3.set_title(f"{game_name} - Drop Comparison", fontsize=10)
    ax3.legend()
    charts['combined_drop'] = fig3
    return charts

def add_charts_to_excel(worksheet, charts):
    positions = {'retention': 'M2', 'total_drop': 'N32', 'combined_drop': 'N65'}
    for chart_name in charts:
        img_data = BytesIO()
        charts[chart_name].savefig(img_data, format='png', dpi=150, bbox_inches='tight')
        img_data.seek(0)
        img = OpenpyxlImage(img_data)
        worksheet.add_image(img, positions[chart_name])
        plt.close(charts[chart_name])

def generate_excel(processed_data):
    wb = Workbook()
    wb.remove(wb.active)
    main = wb.create_sheet("MAIN_TAB")
    headers = ["MAIN_TAB"] + [
        "Level", "Start Users", "Complete Users", "Game Play Drop",
        "Popup Drop", "Total Level Drop", "Retention %",
        "PLAY_TIME_AVG", "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPTS_SUM"]

    for idx, (game_id, df) in enumerate(processed_data.items(), 1):
        sheet_name = f"{game_id}_{df['DIFFICULTY'].iloc[0]}"[:31]
        ws = wb.create_sheet(sheet_name)

        # Insert hyperlink in A1
        ws.cell(row=1, column=1, value='=HYPERLINK("#MAIN_TAB!A1", "MAIN_TAB")')
        ws.cell(row=1, column=1).font = Font(color="0000FF", underline="single")

        # Headers start in row 1
        for col_idx, col_name in enumerate(headers[1:], start=2):
            ws.cell(row=1, column=col_idx).value = col_name
            ws.cell(row=1, column=col_idx).font = Font(bold=True)
            ws.cell(row=1, column=col_idx).fill = PatternFill("solid", fgColor="DDDDDD")
            ws.cell(row=1, column=col_idx).alignment = Alignment(horizontal='center', vertical='center')

        # Data starts in row 2
        for row_idx, (_, row) in enumerate(df.iterrows(), start=2):
            values = [
                row['LEVEL'], row['Start Users'], row['Complete Users'], row['Game Play Drop'],
                row['Popup Drop'], row['Total Level Drop'], row['Retention %'],
                row['PLAY_TIME_AVG'], row['HINT_USED_SUM'], row['SKIPPED_SUM'], row['ATTEMPTS_SUM']
            ]
            for col_idx, value in enumerate(values, start=2):
                ws.cell(row=row_idx, column=col_idx).value = value
                ws.cell(row=row_idx, column=col_idx).alignment = Alignment(horizontal='center', vertical='center')

        apply_conditional_formatting(ws, df.shape[0])
        charts = create_charts(df, sheet_name)
        add_charts_to_excel(ws, charts)

        # Auto-adjust column width
        for col in ws.columns:
            max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2
    return wb

def apply_conditional_formatting(sheet, row_count):
    drop_cols = ['D', 'E', 'F']
    fill_map = {
        '3': PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
        '7': PatternFill(start_color='FF9999', end_color='FF9999', fill_type='solid'),
        '10': PatternFill(start_color='FF6666', end_color='FF6666', fill_type='solid')
    }
    for r in sheet.iter_rows(min_row=2, max_row=row_count + 1):
        for cell in r:
            if cell.column_letter in drop_cols and isinstance(cell.value, (int, float)):
                if cell.value >= 10:
                    cell.fill = fill_map['10']
                elif cell.value >= 7:
                    cell.fill = fill_map['7']
                elif cell.value >= 3:
                    cell.fill = fill_map['3']
                cell.font = Font(color="FFFFFF")
            cell.alignment = Alignment(horizontal='center', vertical='center')

def main():
    st.sidebar.header("Upload Files")
    start_file = st.sidebar.file_uploader("LEVEL_START.csv", type="csv")
    complete_file = st.sidebar.file_uploader("LEVEL_COMPLETE.csv", type="csv")

    if start_file and complete_file:
        with st.spinner("Processing data..."):
            try:
                start_df = pd.read_csv(start_file)
                complete_df = pd.read_csv(complete_file)
                merged = process_files(start_df, complete_df)
                grouped = {f"{gid}": g for (gid, _), g in merged.groupby(['GAME_ID', 'DIFFICULTY'])}
                wb = generate_excel(grouped)
                with tempfile.NamedTemporaryFile(delete=False) as tmp:
                    wb.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        excel_bytes = f.read()
                st.success("Processing complete!")
                st.download_button("📥 Download Consolidated Report", excel_bytes, "Game_Analytics_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with st.expander("Preview Processed Data"):
                    st.dataframe(merged.head(20))
            except Exception as e:
                st.error(f"Error processing files: {str(e)}")

if __name__ == "__main__":
    main()
