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

    # Rename required columns
    start_df.rename(columns={
        level_col: 'LEVEL',
        game_col: 'GAME_ID',
        diff_col: 'DIFFICULTY',
        'USERS': 'Start Users',
        playtime_col: 'PLAY_TIME_AVG' if playtime_col else None,
        hint_col: 'HINT_USED_SUM' if hint_col else None,
        skipped_col: 'SKIPPED_SUM' if skipped_col else None,
        attempts_col: 'ATTEMPTS_SUM' if attempts_col else None,
    }, inplace=True)

    complete_df.rename(columns={
        level_col: 'LEVEL',
        game_col: 'GAME_ID',
        diff_col: 'DIFFICULTY',
        'USERS': 'Complete Users'
    }, inplace=True)

    # Merge
    merged = pd.merge(start_df, complete_df, on=['GAME_ID', 'DIFFICULTY', 'LEVEL'], how='outer', suffixes=('_start', '_complete'))

    # Build dynamic column list
    keep_cols = ['GAME_ID', 'DIFFICULTY', 'LEVEL', 'Start Users', 'Complete Users']
    if playtime_col: keep_cols.append('PLAY_TIME_AVG')
    if hint_col: keep_cols.append('HINT_USED_SUM')
    if skipped_col: keep_cols.append('SKIPPED_SUM')
    if attempts_col: keep_cols.append('ATTEMPTS_SUM')

    merged = merged[[col for col in keep_cols if col in merged.columns]]

    # Calculate drops and retention
    merged['Game Play Drop'] = ((merged['Start Users'] - merged['Complete Users']) / merged['Start Users'].replace(0, np.nan)) * 100
    merged['Popup Drop'] = ((merged['Complete Users'] - merged['Start Users'].shift(-1)) / merged['Complete Users'].replace(0, np.nan)) * 100
    merged['Total Level Drop'] = ((merged['Start Users'] - merged['Start Users'].shift(-1)) / merged['Start Users'].replace(0, np.nan)) * 100
    merged['Retention %'] = (merged['Start Users'] / merged['Start Users'].max()) * 100

    # Clean NaNs
    merged.fillna({'Start Users': 0, 'Complete Users': 0}, inplace=True)
    return merged


# ======================== CHART GENERATION ========================
def create_charts(df, game_name):
    """Generate matplotlib charts"""
    charts = {}

    # Retention Chart
    fig1, ax1 = plt.subplots(figsize=(12, 4))
    ax1.plot(df['LEVEL'], df['Retention %'], color='#4CAF50')
    ax1.set_title(f"{game_name} - Retention %", fontsize=10)
    charts['retention'] = fig1

    # Total Level Drop Chart
    fig2, ax2 = plt.subplots(figsize=(12, 4))
    ax2.bar(df['LEVEL'], df['Total Level Drop'], color='#F44336')
    ax2.set_title(f"{game_name} - Total Level Drop", fontsize=10)
    charts['total_drop'] = fig2

    # Combined Drop Chart
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
    """Add matplotlib charts to Excel worksheet as images"""
    img_positions = {
        'retention': 'M2',
        'total_drop': 'N32',
        'combined_drop': 'N65'
    }

    for chart_type in ['retention', 'total_drop', 'combined_drop']:
        # Save chart to bytes buffer
        img_data = BytesIO()
        charts[chart_type].savefig(img_data, format='png', dpi=150, bbox_inches='tight')
        img_data.seek(0)

        # Create image object
        img = OpenpyxlImage(img_data)

        # Add image to worksheet
        worksheet.add_image(img, img_positions[chart_type])

        # Close figure to prevent memory leaks
        plt.close(charts[chart_type])

# ======================== EXCEL GENERATION ========================
def generate_excel(processed_data):
    """Create Excel workbook with formatted sheets"""
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Create MAIN_TAB sheet
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_headers = ["Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
                    "Total Level Drop Count", "LEVEL_Start", "Start Users",
                    "LEVEL_End", "USERS_END", "Link to Sheet"]
    main_sheet.append(main_headers)


    for cell in main_sheet[row_ptr]:
            cell.alignment = Alignment(horizontal='center', vertical='center')
    row_ptr += 1



    # row_ptr = 2
    # while row_ptr <= main_sheet.max_row:  # or any row number you want
    #    for cell in main_sheet[row_ptr]:
    #      cell.alignment = Alignment(horizontal='center', vertical='center')
    #    row_ptr += 1



    # Format main sheet headers
    for col in main_sheet[1]:
        col.font = Font(bold=True, color="FFFFFF")
        col.fill = PatternFill("solid", fgColor="4F81BD")

    # Process each game variant
    for idx, (game_id, df) in enumerate(processed_data.items(), start=1):
        sheet_name = f"{game_id}_{df['DIFFICULTY'].iloc[0]}"[:31]
        ws = wb.create_sheet(sheet_name)


        headers = ["=HYPERLINK(\"#MAIN_TAB!A1\", \"Back to Main\")", "Start Users", "Complete Users",
                   "Game Play Drop", "Popup Drop", "Total Level Drop", "Retention %",
                   "PLAY_TIME_AVG", "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPTS_SUM"]

        ws.append(headers)

        # Apply formatting to A1 hyperlink (now embedded in header)
        ws['A1'].font = Font(color="0000FF", underline="single", bold=True)



        # Add data rows
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
            ws.append([val if val != "" else "0" for val in values])

            for cell in main_sheet[row_ptr]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            row_ptr += 1


        # Add charts
        charts = create_charts(df, sheet_name)
        add_charts_to_excel(ws, charts)

        # Formatting
        apply_sheet_formatting(ws)
        apply_conditional_formatting(ws, df.shape[0])

        # Update MAIN_TAB
        main_row = [
            idx, sheet_name,
            sum(df['Game Play Drop'] >= (df['Start Users'] * 0.03)),
            sum(df['Popup Drop'] >= (df['Start Users'] * 0.03)),
            sum(df['Total Level Drop'] >= (df['Start Users'] * 0.03)),
            df['LEVEL'].min(), df['Start Users'].max(),
            df['LEVEL'].max(), df['Complete Users'].iloc[-1],
            f'=HYPERLINK("#{sheet_name}!A1", "View")'
        ]
        main_sheet.append(main_row)

    # Format main sheet
    for col in range(1, len(main_headers)+1):
        main_sheet.column_dimensions[get_column_letter(col)].width = 18

    return wb

# ======================== REMAINING FUNCTIONS AND UI (UNCHANGED) ========================
# [Keep the apply_sheet_formatting, apply_conditional_formatting, and main() functions
# from the previous implementation unchanged]

def apply_sheet_formatting(sheet):
    """Apply consistent formatting to sheets"""
    # Freeze header row
    sheet.freeze_panes = 'A1'

    # Format headers
    for cell in sheet[1]:  # Data headers start at row 1
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")

    # Auto-fit columns
    for col in sheet.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        sheet.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

def apply_conditional_formatting(sheet, num_rows):
    """Apply color scale formatting to drop columns"""
    drop_columns = {'D', 'E', 'F'}  # Game Play Drop, Popup Drop, Total Level Drop

    red_scale = {
        '3': PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
        '7': PatternFill(start_color='FF9999', end_color='FF9999', fill_type='solid'),
        '10': PatternFill(start_color='FF6666', end_color='FF6666', fill_type='solid')
    }

    for row in sheet.iter_rows(min_row=2, max_row=num_rows+1):
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


     # Center alignment for all cells
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

# ======================== STREAMLIT UI ========================
def main():
    st.sidebar.header("Upload Files")
    start_file = st.sidebar.file_uploader("LEVEL_START.csv", type="csv")
    complete_file = st.sidebar.file_uploader("LEVEL_COMPLETE.csv", type="csv")

    if start_file and complete_file:
        with st.spinner("Processing data..."):
            try:
                # Read and process data
                start_df = pd.read_csv(start_file)
                complete_df = pd.read_csv(complete_file)
                merged = process_files(start_df, complete_df)

                # Group by game and difficulty
                processed_data = {}
                for (game_id, difficulty), group in merged.groupby(['GAME_ID', 'DIFFICULTY']):
                    processed_data[f"{game_id}"] = group

                # Generate Excel file
                wb = generate_excel(processed_data)

                # Save to bytes buffer
                with tempfile.NamedTemporaryFile(delete=False) as tmp:
                    wb.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        excel_bytes = f.read()

                # Download button
                st.success("Processing complete!")
                st.download_button(
                    label="📥 Download Consolidated Report",
                    data=excel_bytes,
                    file_name="Game_Analytics_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Show preview
                with st.expander("Preview Processed Data"):
                    st.dataframe(merged.head(20))

            except Exception as e:
                st.error(f"Error processing files: {str(e)}")

if __name__ == "__main__":
    main()
