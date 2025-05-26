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
    level_col_start = get_column(start_df, ['LEVEL', 'TOTALLEVELS', 'STAGE'])
    game_col_start = get_column(start_df, ['GAME_ID', 'CATEGORY', 'Game_name'])
    diff_col_start = get_column(start_df, ['DIFFICULTY', 'mode'])

    level_col_comp = get_column(complete_df, ['LEVEL', 'TOTALLEVELS', 'STAGE'])
    game_col_comp = get_column(complete_df, ['GAME_ID', 'CATEGORY', 'Game_name'])
    diff_col_comp = get_column(complete_df, ['DIFFICULTY', 'mode'])

    playtime_col = get_column(complete_df, ['PLAY_TIME_AVG', 'PLAYTIME', 'PLAYTIME_AVG', 'playtime_avg'])
    hint_col = get_column(complete_df, ['HINT_USED_SUM', 'HINT_USED', 'HINT'])
    skipped_col = get_column(complete_df, ['SKIPPED_SUM', 'SKIPPED', 'SKIP'])
    attempts_col = get_column(complete_df, ['ATTEMPTS_SUM', 'ATTEMPTS', 'TRY_COUNT'])

    # Clean LEVELs
    for df in [start_df, complete_df]:
        level_col = get_column(df, ['LEVEL', 'TOTALLEVELS', 'STAGE'])
        if level_col:
            df[level_col] = df[level_col].apply(clean_level)
            df.sort_values(level_col, inplace=True)

    # Dynamically build rename dictionaries
    def build_rename_dict(df, level_col, game_col, diff_col):
        rename_dict = {}
        if level_col:
            rename_dict[level_col] = 'LEVEL'
        if game_col:
            rename_dict[game_col] = 'GAME_ID'
        if diff_col:
            rename_dict[diff_col] = 'DIFFICULTY'
        return rename_dict

    # Rename columns
    start_rename = build_rename_dict(start_df, level_col_start, game_col_start, diff_col_start)
    start_rename['USERS'] = 'Start Users'
    start_df.rename(columns=start_rename, inplace=True)

    comp_rename = build_rename_dict(complete_df, level_col_comp, game_col_comp, diff_col_comp)
    comp_rename['USERS'] = 'Complete Users'
    if playtime_col:
        comp_rename[playtime_col] = 'PLAY_TIME_AVG'
    if hint_col:
        comp_rename[hint_col] = 'HINT_USED_SUM'
    if skipped_col:
        comp_rename[skipped_col] = 'SKIPPED_SUM'
    if attempts_col:
        comp_rename[attempts_col] = 'ATTEMPTS_SUM'
    complete_df.rename(columns=comp_rename, inplace=True)

    # Determine merge columns
    merge_cols = ['LEVEL']
    if 'GAME_ID' in start_df.columns and 'GAME_ID' in complete_df.columns:
        merge_cols.append('GAME_ID')
    if 'DIFFICULTY' in start_df.columns and 'DIFFICULTY' in complete_df.columns:
        merge_cols.append('DIFFICULTY')

    # Merge dataframes
    merged = pd.merge(
        start_df, 
        complete_df, 
        on=merge_cols, 
        how='outer', 
        suffixes=('_start', '_complete')
    )

    # Build dynamic column list
    keep_cols = ['LEVEL', 'Start Users', 'Complete Users']
    if 'GAME_ID' in merged.columns:
        keep_cols.append('GAME_ID')
    if 'DIFFICULTY' in merged.columns:
        keep_cols.append('DIFFICULTY')
    if playtime_col:
        keep_cols.append('PLAY_TIME_AVG')
    if hint_col:
        keep_cols.append('HINT_USED_SUM')
    if skipped_col:
        keep_cols.append('SKIPPED_SUM')
    if attempts_col:
        keep_cols.append('ATTEMPTS_SUM')

    merged = merged[[col for col in keep_cols if col in merged.columns]]

    # Calculate metrics
    merged['Game Play Drop'] = ((merged['Start Users'] - merged['Complete Users']) / 
                               merged['Start Users'].replace(0, np.nan)) * 100
    merged['Popup Drop'] = ((merged['Complete Users'] - merged['Start Users'].shift(-1)) / 
                            merged['Complete Users'].replace(0, np.nan)) * 100

    # Retention calculation
    def calculate_retention(group):
        base_users = group[group['LEVEL'].isin([1, 2])]['Start Users'].max()
        if base_users == 0 or pd.isnull(base_users):
            base_users = group['Start Users'].max()
        group['Retention %'] = (group['Start Users'] / base_users) * 100
        return group

    group_cols = []
    if 'GAME_ID' in merged.columns:
        group_cols.append('GAME_ID')
    if 'DIFFICULTY' in merged.columns:
        group_cols.append('DIFFICULTY')

    if group_cols:
        merged = merged.groupby(group_cols, group_keys=False).apply(calculate_retention)
    else:
        merged = calculate_retention(merged)

    merged.fillna({'Start Users': 0, 'Complete Users': 0}, inplace=True)
    merged['Total Level Drop'] = merged['Game Play Drop'] + merged['Popup Drop']
    return merged

# ======================== CHART GENERATION ========================
def create_charts(df, group_name):
    charts = {}
    
    # Retention Chart
    fig1, ax1 = plt.subplots(figsize=(12, 4))
    ax1.plot(df['LEVEL'], df['Retention %'], color='#4CAF50')
    ax1.set_title(f"{group_name} - Retention %", fontsize=10)
    charts['retention'] = fig1

    # Total Level Drop Chart
    fig2, ax2 = plt.subplots(figsize=(12, 4))
    ax2.bar(df['LEVEL'], df['Total Level Drop'], color='#F44336')
    ax2.set_title(f"{group_name} - Total Level Drop", fontsize=10)
    charts['total_drop'] = fig2

    # Combined Drop Chart
    fig3, ax3 = plt.subplots(figsize=(12, 4))
    width = 0.35
    ax3.bar(df['LEVEL'] - width/2, df['Game Play Drop'], width, label='Game Play Drop')
    ax3.bar(df['LEVEL'] + width/2, df['Popup Drop'], width, label='Popup Drop')
    ax3.set_title(f"{group_name} - Drop Comparison", fontsize=10)
    ax3.legend()
    charts['combined_drop'] = fig3

    return charts

# ======================== EXCEL GENERATION ========================
def generate_excel(processed_data):
    wb = Workbook()
    wb.remove(wb.active)
    
    # Main sheet setup
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_headers = ["Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
                    "Total Level Drop Count", "LEVEL_Start", "Start Users",
                    "LEVEL_End", "USERS_END", "Link to Sheet"]
    main_sheet.append(main_headers)

    # Format main headers
    for col in main_sheet[1]:
        col.font = Font(bold=True, color="FFFFFF")
        col.fill = PatternFill("solid", fgColor="4F81BD")

    # Process each group
    for idx, (group_name, df) in enumerate(processed_data.items(), start=1):
        sheet = wb.create_sheet(str(group_name)[:30])
        
        # Add hyperlink header
        sheet.append(["=HYPERLINK(\"#MAIN_TAB!A1\", \"Back to Main\")", "Start Users", "Complete Users",
                     "Game Play Drop", "Popup Drop", "Total Level Drop", "Retention %",
                     "PLAY_TIME_AVG", "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPTS_SUM"])
        
        # Format hyperlink cell
        sheet['A1'].font = Font(color="0000FF", underline="single", bold=True)
        sheet['A1'].fill = PatternFill("solid", fgColor="FFFF00")
        
        # Add data rows
        for _, row in df.iterrows():
            sheet.append([
                row['LEVEL'],
                row.get('Start Users', 0),
                row.get('Complete Users', 0),
                round(row.get('Game Play Drop', 0), 2),
                round(row.get('Popup Drop', 0), 2),
                round(row.get('Total Level Drop', 0), 2),
                round(row.get('Retention %', 0), 2),
                round(row.get('PLAY_TIME_AVG', 0), 2),
                round(row.get('HINT_USED_SUM', 0), 2),
                round(row.get('SKIPPED_SUM', 0), 2),
                round(row.get('ATTEMPTS_SUM', 0), 2)
            ])
        
        # Add charts
        charts = create_charts(df, group_name)
        add_charts_to_excel(sheet, charts)
        
        # Formatting
        apply_sheet_formatting(sheet)
        apply_conditional_formatting(sheet, df.shape[0])

        # Update main sheet
        main_row = [
            idx, group_name,
            sum(df['Game Play Drop'] >= 3),
            sum(df['Popup Drop'] >= 3),
            sum(df['Total Level Drop'] >= 3),
            df['LEVEL'].min(),
            df['Start Users'].max(),
            df['LEVEL'].max(),
            df['Complete Users'].iloc[-1],
            f'=HYPERLINK("#{group_name}!A1", "View")'
        ]
        main_sheet.append(main_row)

    # Format main sheet
    for row in main_sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

    column_widths = [8, 25, 20, 18, 20, 12, 15, 12, 15, 15]
    for i, width in enumerate(column_widths, start=1):
        main_sheet.column_dimensions[get_column_letter(i)].width = width

    return wb

# ======================== HELPER FUNCTIONS ========================
def add_charts_to_excel(sheet, charts):
    positions = {'retention': 'M2', 'total_drop': 'N32', 'combined_drop': 'N65'}
    for chart_type, pos in positions.items():
        img_data = BytesIO()
        charts[chart_type].savefig(img_data, format='png', dpi=150)
        img_data.seek(0)
        sheet.add_image(OpenpyxlImage(img_data), pos)
        plt.close(charts[chart_type])

def apply_sheet_formatting(sheet):
    sheet.freeze_panes = 'A2'
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
    for column in sheet.columns:
        sheet.column_dimensions[get_column_letter(column[0].column)].width = 15

def apply_conditional_formatting(sheet, num_rows):
    red_scale = {
        3: PatternFill(start_color='FFC7CE'),
        7: PatternFill(start_color='FF9999'),
        10: PatternFill(start_color='FF6666')
    }
    for row in sheet.iter_rows(min_row=2, max_row=num_rows+1):
        for cell in row:
            if cell.column_letter in ['D', 'E', 'F'] and isinstance(cell.value, (int, float)):
                if cell.value >= 10:
                    cell.fill = red_scale[10]
                elif cell.value >= 7:
                    cell.fill = red_scale[7]
                elif cell.value >= 3:
                    cell.fill = red_scale[3]

# ======================== STREAMLIT UI ========================
def main():
    st.sidebar.header("Upload Files")
    start_file = st.sidebar.file_uploader("LEVEL_START.csv", type="csv")
    complete_file = st.sidebar.file_uploader("LEVEL_COMPLETE.csv", type="csv")

    if start_file and complete_file:
        with st.spinner("Processing data..."):
            try:
                # Process data
                start_df = pd.read_csv(start_file)
                complete_df = pd.read_csv(complete_file)
                merged = process_files(start_df, complete_df)

                # Dynamic grouping
                group_cols = []
                if 'GAME_ID' in merged.columns:
                    group_cols.append('GAME_ID')
                if 'DIFFICULTY' in merged.columns:
                    group_cols.append('DIFFICULTY')

                if group_cols:
                    groups = list(merged.groupby(group_cols))
                else:
                    groups = [('All Data', merged)]

                processed_data = {}
                for group_key, group_df in groups:
                    if isinstance(group_key, tuple):
                        group_name = '_'.join(str(k) for k in group_key)
                    else:
                        group_name = str(group_key)
                    processed_data[group_name] = group_df

                # Generate Excel
                wb = generate_excel(processed_data)
                with tempfile.NamedTemporaryFile() as tmp:
                    wb.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        excel_bytes = f.read()

                st.success("Processing complete!")
                st.download_button(
                    label="📥 Download Report",
                    data=excel_bytes,
                    file_name="Game_Analytics.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                with st.expander("Preview Data"):
                    st.dataframe(merged.head())

            except Exception as e:
                st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
