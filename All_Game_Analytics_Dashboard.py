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
    merged.fillna({'Start Users': 0, 'Complete Users': 0}, inplace=True)
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
    img_positions = {
        'retention': 'M2',
        'total_drop': 'N32',
        'combined_drop': 'N65'
    }

    for chart_type in charts:
        img_data = BytesIO()
        charts[chart_type].savefig(img_data, format='png', dpi=150, bbox_inches='tight')
        img_data.seek(0)
        img = OpenpyxlImage(img_data)
        worksheet.add_image(img, img_positions[chart_type])
        plt.close(charts[chart_type])

def generate_excel(processed_data):
    wb = Workbook()
    wb.remove(wb.active)

    main_sheet = wb.create_sheet("MAIN_TAB")
    main_sheet['A1'] = '=HYPERLINK("#MAIN_TAB!A1", "MAIN_TAB")'
    for col in range(2, 12):
        main_sheet.cell(row=1, column=col).value = ""
    for cell in main_sheet[1]:
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.font = Font(bold=True)

    main_headers = [
        "Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
        "Total Level Drop Count", "LEVEL_Start", "Start Users",
        "LEVEL_End", "USERS_END", "Link to Sheet"
    ]
    for col_index, title in enumerate(main_headers, start=2):
        main_sheet.cell(row=2, column=col_index).value = title
    for cell in main_sheet[2]:
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4F81BD")

    row_ptr = 3
    for idx, (game_id, df) in enumerate(processed_data.items(), start=1):
        sheet_name = f"{game_id}_{df['DIFFICULTY'].iloc[0]}"[:31]
        ws = wb.create_sheet(sheet_name)
        ws['A1'] = '=HYPERLINK("#MAIN_TAB!A1", "MAIN_TAB")'
        ws['A1'].font = Font(color="0000FF", underline="single")

        headers = ["Level", "Start Users", "Complete Users", "Game Play Drop",
                   "Popup Drop", "Total Level Drop", "Retention %",
                   "PLAY_TIME_AVG", "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPTS_SUM"]
        ws.append(headers)

        for _, row in df.iterrows():
            values = [
                row['LEVEL'],
                row['Start Users'] if not pd.isna(row['Start Users']) else 0,
                row['Complete Users'] if not pd.isna(row['Complete Users']) else 0,
                round(row['Game Play Drop'] if not pd.isna(row['Game Play Drop']) else 0, 2),
                round(row['Popup Drop'] if not pd.isna(row['Popup Drop']) else 0, 2),
                round(row['Total Level Drop'] if not pd.isna(row['Total Level Drop']) else 0, 2),
                round(row['Retention %'] if not pd.isna(row['Retention %']) else 0, 2),
                round(row['PLAY_TIME_AVG'] if not pd.isna(row['PLAY_TIME_AVG']) else 0, 2),
                round(row['HINT_USED_SUM'] if not pd.isna(row['HINT_USED_SUM']) else 0, 2),
                round(row['SKIPPED_SUM'] if not pd.isna(row['SKIPPED_SUM']) else 0, 2),
                round(row['ATTEMPTS_SUM'] if not pd.isna(row['ATTEMPTS_SUM']) else 0, 2),
            ]
            ws.append([val if val != "" else "null" for val in values])

        charts = create_charts(df, sheet_name)
        add_charts_to_excel(ws, charts)
        apply_sheet_formatting(ws)
        apply_conditional_formatting(ws, df.shape[0])

        main_sheet.cell(row=row_ptr, column=2).value = idx
        main_sheet.cell(row=row_ptr, column=3).value = sheet_name
        main_sheet.cell(row=row_ptr, column=4).value = sum(df['Game Play Drop'] >= 3)
        main_sheet.cell(row=row_ptr, column=5).value = sum(df['Popup Drop'] >= 3)
        main_sheet.cell(row=row_ptr, column=6).value = sum(df['Total Level Drop'] >= 3)
        main_sheet.cell(row=row_ptr, column=7).value = df['LEVEL'].min()
        main_sheet.cell(row=row_ptr, column=8).value = df['Start Users'].max()
        main_sheet.cell(row=row_ptr, column=9).value = df['LEVEL'].max()
        main_sheet.cell(row=row_ptr, column=10).value = df['Complete Users'].iloc[-1]
        main_sheet.cell(row=row_ptr, column=11).value = f'=HYPERLINK("#{sheet_name}!A1", "View")'

        for cell in main_sheet[row_ptr]:
            cell.alignment = Alignment(horizontal='center', vertical='center')
        row_ptr += 1

    for col in range(1, 12):
        main_sheet.column_dimensions[get_column_letter(col)].width = 18

    return wb

def apply_sheet_formatting(sheet):
    sheet.freeze_panes = 'A1'
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    for col in sheet.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        sheet.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

def apply_conditional_formatting(sheet, num_rows):
    drop_columns = {'D', 'E', 'F'}
    red_scale = {
        '3': PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
        '7': PatternFill(start_color='FF9999', end_color='FF9999', fill_type='solid'),
        '10': PatternFill(start_color='FF6666', end_color='FF6666', fill_type='solid')
    }
    for row in sheet.iter_rows(min_row=3, max_row=num_rows+2):
        for cell in row:
            if cell.column_letter in drop_columns and cell.value is not None:
                value = cell.value
                if value >= 10:
                    cell.fill = red_scale['10']
                elif value >= 7:
                    cell.fill = red_scale['7']
                elif value >= 3:
                    cell.fill = red_scale['3']
                cell.font = Font(color="FFFFFF")

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

                processed_data = {}
                for (game_id, difficulty), group in merged.groupby(['GAME_ID', 'DIFFICULTY']):
                    processed_data[f"{game_id}"] = group

                wb = generate_excel(processed_data)

                with tempfile.NamedTemporaryFile(delete=False) as tmp:
                    wb.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        excel_bytes = f.read()

                st.success("Processing complete!")
                st.download_button(
                    label="📥 Download Consolidated Report",
                    data=excel_bytes,
                    file_name="Game_Analytics_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                with st.expander("Preview Processed Data"):
                    st.dataframe(merged.head(20))

            except Exception as e:
                st.error(f"Error processing files: {str(e)}")

if __name__ == "__main__":
    main()
