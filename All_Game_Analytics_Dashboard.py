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
import tempfile

# Initialize Streamlit app
st.set_page_config(page_title="Game Analytics Tool", layout="wide")
st.title("🎮 Game Level Data Analyzer")

# ======================== DATA PROCESSING FUNCTIONS ========================
def clean_level(level):
    """Extract numeric value from LEVEL column"""
    if pd.isna(level):
        return 0
    return int(re.sub(r'\D', '', str(level)))

def process_files(start_df, complete_df):
    """Process and merge the two dataframes with flexible column name handling."""
    def get_column(df, possible_names):
        """Return the first matching column name from the dataframe."""
        for col in df.columns:
            if col.strip().lower() in [name.lower() for name in possible_names]:
                return col
        return None

    # Flexible column matching
    level_col = get_column(start_df, ['LEVEL', 'TOTALLEVELS', 'STAGE'])
    game_col = get_column(start_df, ['GAME_ID', 'CATEGORY', 'Game_name'])
    diff_col = get_column(start_df, ['DIFFICULTY', 'mode'])

    playtime_col = get_column(complete_df, ['PLAY_TIME_AVG', 'PLAYTIME', 'PLAYTIME_AVG', 'playtime_avg'])
    hint_col = get_column(complete_df, ['HINT_USED_SUM', 'HINT_USED', 'HINT'])
    skipped_col = get_column(complete_df, ['SKIPPED_SUM', 'SKIPPED', 'SKIP'])
    attempts_col = get_column(complete_df, ['ATTEMPTS_SUM', 'ATTEMPTS', 'TRY_COUNT'])

    # Clean LEVELs
    for df in [start_df, complete_df]:
        df[level_col] = df[level_col].apply(clean_level)
        df.sort_values(level_col, inplace=True)

    # Dynamic column renaming
    rename_dict_start = {level_col: 'LEVEL', 'USERS': 'Start Users'}
    if game_col: rename_dict_start[game_col] = 'GAME_ID'
    if diff_col: rename_dict_start[diff_col] = 'DIFFICULTY'
    start_df.rename(columns=rename_dict_start, inplace=True)

    rename_dict_complete = {level_col: 'LEVEL', 'USERS': 'Complete Users'}
    if game_col: rename_dict_complete[game_col] = 'GAME_ID'
    if diff_col: rename_dict_complete[diff_col] = 'DIFFICULTY'
    if playtime_col: rename_dict_complete[playtime_col] = 'PLAY_TIME_AVG'
    if hint_col: rename_dict_complete[hint_col] = 'HINT_USED_SUM'
    if skipped_col: rename_dict_complete[skipped_col] = 'SKIPPED_SUM'
    if attempts_col: rename_dict_complete[attempts_col] = 'ATTEMPTS_SUM'
    complete_df.rename(columns=rename_dict_complete, inplace=True)

    # Dynamic merge columns
    merge_on = ['LEVEL']
    if 'GAME_ID' in start_df.columns: merge_on.append('GAME_ID')
    if 'DIFFICULTY' in start_df.columns: merge_on.append('DIFFICULTY')

    merged = pd.merge(start_df, complete_df, on=merge_on, how='outer', suffixes=('_start', '_complete'))

    # Build dynamic column list
    keep_cols = ['LEVEL', 'Start Users', 'Complete Users']
    if 'GAME_ID' in merged.columns: keep_cols.append('GAME_ID')
    if 'DIFFICULTY' in merged.columns: keep_cols.append('DIFFICULTY')
    if playtime_col: keep_cols.append('PLAY_TIME_AVG')
    if hint_col: keep_cols.append('HINT_USED_SUM')
    if skipped_col: keep_cols.append('SKIPPED_SUM')
    if attempts_col: keep_cols.append('ATTEMPTS_SUM')

    merged = merged[[col for col in keep_cols if col in merged.columns]]

    # Calculate drops with zero handling
    merged['Game Play Drop'] = np.where(
        merged['Start Users'] > 0,
        ((merged['Start Users'] - merged['Complete Users']) / merged['Start Users']) * 100,
        0
    )
    
    # Calculate Popup Drop with forward fill for last level
    merged['Popup Drop'] = np.where(
        merged['Complete Users'] > 0,
        ((merged['Complete Users'] - merged['Start Users'].shift(-1).fillna(0)) / merged['Complete Users']) * 100,
        0
    )

    # Dynamic retention grouping with zero handling
    def calculate_retention(group):
        base_users = group[group['LEVEL'].isin([1, 2])]['Start Users'].max()
        if base_users == 0 or pd.isnull(base_users):
            base_users = group['Start Users'].max() or 1  # Prevent division by zero
        group['Retention %'] = np.where(
            base_users > 0,
            (group['Start Users'] / base_users) * 100,
            0
        )
        return group

    group_cols = []
    if 'GAME_ID' in merged.columns: group_cols.append('GAME_ID')
    if 'DIFFICULTY' in merged.columns: group_cols.append('DIFFICULTY')

    if group_cols:
        merged = merged.groupby(group_cols, group_keys=False).apply(calculate_retention)
    else:
        merged = calculate_retention(merged)

    # Fill all remaining NaNs with 0 for key columns
    merged.fillna({
        'Start Users': 0,
        'Complete Users': 0,
        'Game Play Drop': 0,
        'Popup Drop': 0,
        'Total Level Drop': 0,
        'Retention %': 0
    }, inplace=True)

    # Fill all remaining numeric columns with 0
    numeric_cols = merged.select_dtypes(include=[np.number]).columns
    merged[numeric_cols] = merged[numeric_cols].fillna(0)
  
    merged['Total Level Drop'] = merged['Game Play Drop'] + merged['Popup Drop']
    
    return merged

# ======================== CHART GENERATION ========================
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

# ======================== CHART ADDITION TO EXCEL ========================
def add_charts_to_excel(worksheet, charts):
    img_positions = {'retention': 'M2', 'total_drop': 'N32', 'combined_drop': 'N65'}
    for chart_type, position in img_positions.items():
        img_data = BytesIO()
        charts[chart_type].savefig(img_data, format='png', dpi=150, bbox_inches='tight')
        img_data.seek(0)
        worksheet.add_image(OpenpyxlImage(img_data), position)
        plt.close(charts[chart_type])

# ======================== EXCEL GENERATION ========================
def generate_excel(processed_data):
    wb = Workbook()
    wb.remove(wb.active)
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_headers = ["Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
                    "Total Level Drop Count", "LEVEL_Start", "Start Users",
                    "LEVEL_End", "USERS_END", "Link to Sheet"]
    main_sheet.append(main_headers)

    for col in main_sheet[1]:
        col.font = Font(bold=True, color="FFFFFF")
        col.fill = PatternFill("solid", fgColor="4F81BD")

    for idx, (game_key, df) in enumerate(processed_data.items(), start=1):
        sheet_name = str(game_key)[:31]
        ws = wb.create_sheet(sheet_name)
        headers = ["=HYPERLINK(\"#MAIN_TAB!A1\", \"Back to Main\")", "Start Users", "Complete Users",
                   "Game Play Drop", "Popup Drop", "Total Level Drop", "Retention %",
                   "PLAY_TIME_AVG", "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPTS_SUM"]
        ws.append(headers)

        ws['A1'].font = Font(color="0000FF", underline="single", bold=True, size=14)
        ws['A1'].fill = PatternFill("solid", fgColor="FFFF00")
        ws.column_dimensions['A'].width = 25

        for _, row in df.iterrows():
            values = [
                int(row['LEVEL']),
                int(row.get('Start Users', 0)),
                int(row.get('Complete Users', 0)),
                round(float(row.get('Game Play Drop', 0)), 2),
                round(float(row.get('Popup Drop', 0)), 2),
                round(float(row.get('Total Level Drop', 0)), 2),
                round(float(row.get('Retention %', 0)), 2),
                round(float(row.get('PLAY_TIME_AVG', 0)), 2),
                int(row.get('HINT_USED_SUM', 0)),
                int(row.get('SKIPPED_SUM', 0)),
                int(row.get('ATTEMPTS_SUM', 0)),
            ]
            ws.append(values)

        charts = create_charts(df, sheet_name)
        add_charts_to_excel(ws, charts)
        apply_sheet_formatting(ws)
        apply_conditional_formatting(ws, df.shape[0])

        main_row = [
            idx, sheet_name,
            sum(df['Game Play Drop'] >= (df['Start Users'] * 0.03)),
            sum(df['Popup Drop'] >= (df['Start Users'] * 0.03)),
            sum(df['Total Level Drop'] >= (df['Start Users'] * 0.03)),
            int(df['LEVEL'].min()), int(df['Start Users'].max()),
            int(df['LEVEL'].max()), int(df['Complete Users'].iloc[-1]),
            f'=HYPERLINK("#{sheet_name}!A1", "View")'
        ]
        main_sheet.append(main_row)

    for row in main_sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if cell.row == 1:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="4F81BD")

    column_widths = [8, 25, 20, 18, 20, 12, 15, 12, 15, 15]
    for i, width in enumerate(column_widths, start=1):
        main_sheet.column_dimensions[get_column_letter(i)].width = width

    return wb

def apply_sheet_formatting(sheet):
    sheet.freeze_panes = 'A2'
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    if sheet.title != "MAIN_TAB":
        sheet['A1'].font = Font(color="0000FF", underline="single", bold=True, size=11)
        sheet['A1'].fill = PatternFill("solid", fgColor="FFFF00")
        sheet.column_dimensions['A'].width = 14
    for col in sheet.columns:
        if col[0].column == 1 and sheet.title != "MAIN_TAB":
            continue
        max_length = max(len(str(cell.value)) for cell in col)
        sheet.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

def apply_conditional_formatting(sheet, num_rows):
    red_scale = {
        '3': PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
        '7': PatternFill(start_color='FF9999', end_color='FF9999', fill_type='solid'),
        '10': PatternFill(start_color='FF6666', end_color='FF6666', fill_type='solid')
    }
    for row in sheet.iter_rows(min_row=2, max_row=num_rows+1):
        for cell in row:
            if cell.column_letter in {'D', 'E', 'F'} and cell.value is not None:
                value = cell.value
                if value >= 10: cell.fill = red_scale['10']
                elif value >= 7: cell.fill = red_scale['7']
                elif value >= 3: cell.fill = red_scale['3']
                cell.font = Font(color="FFFFFF")
    for row in sheet.iter_rows():
        for cell in row:
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

                # Dynamic grouping for report
                group_cols = []
                if 'GAME_ID' in merged.columns: group_cols.append('GAME_ID')
                if 'DIFFICULTY' in merged.columns: group_cols.append('DIFFICULTY')

                processed_data = {}
                if group_cols:
                    for group_key, group in merged.groupby(group_cols):
                        key = "_".join(map(str, group_key)) if isinstance(group_key, tuple) else str(group_key)
                        processed_data[key] = group
                else:
                    processed_data['All Data'] = merged

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
