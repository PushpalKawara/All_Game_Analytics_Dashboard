# ========================== Step 1: Required Imports ==========================

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
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage
import tempfile

# ========================== Step 2: Streamlit Config ==========================

st.set_page_config(page_title="GAME PROGRESSION", layout="wide")
st.title("📊 GAME PROGRESSION Dashboard")

# ========================== Step 3: Core Functions ==========================

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
            df = df.rename(columns={user_col: 'START_USERS' if is_start_file else 'COMPLETE_USERS'})

        if is_start_file:
            df = df[['LEVEL', 'START_USERS']].drop_duplicates().sort_values('LEVEL')
        else:
            keep_cols = ['LEVEL', 'COMPLETE_USERS', 'PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']
            for col in keep_cols:
                if col not in df.columns:
                    df[col] = 0
            df = df[keep_cols]

        return df.dropna().sort_values('LEVEL')
    except Exception as e:
        st.error(f"Error processing {file_obj.name}: {str(e)}")
        return None

def merge_and_calculate(start_df, complete_df):
    merged = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')

    merged['START_USERS'].fillna(0, inplace=True)
    merged['COMPLETE_USERS'].fillna(0, inplace=True)

    merged['GAME_PLAY_DROP'] = ((merged['START_USERS'] - merged['COMPLETE_USERS']) / merged['START_USERS'].replace(0, np.nan))
    merged['POPUP_DROP'] = ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1)) / merged['COMPLETE_USERS'].replace(0, np.nan))
    merged['TOTAL_LEVEL_DROP'] = ((merged['START_USERS'] - merged['START_USERS'].shift(-1)) / merged['START_USERS'].replace(0, np.nan))
    merged['RETENTION_%'] = (merged['START_USERS'] / merged['START_USERS'].max())

    return merged.round(4)

# ========================== Step 4: Charting Functions ==========================

def create_charts(df, version, date_selected):
    charts = {}
    df_100 = df[df['LEVEL'] <= 100].copy()

    fig1, ax1 = plt.subplots(figsize=(15, 7))
    ax1.plot(df_100['LEVEL'], df_100['RETENTION_%'] * 100, color='#F57C00', linewidth=2)
    format_chart(ax1, "Retention Chart", version, date_selected)
    charts['retention'] = fig1

    fig2, ax2 = plt.subplots(figsize=(15, 6))
    ax2.bar(df_100['LEVEL'], df_100['TOTAL_LEVEL_DROP'] * 100, color='#EF5350')
    format_chart(ax2, "Total Level Drop Chart", version, date_selected)
    charts['total_drop'] = fig2

    fig3, ax3 = plt.subplots(figsize=(15, 6))
    width = 0.4
    ax3.bar(df_100['LEVEL'] + width/2, df_100['GAME_PLAY_DROP'] * 100, width, color='#66BB6A')
    ax3.bar(df_100['LEVEL'] - width/2, df_100['POPUP_DROP'] * 100, width, color='#42A5F5')
    format_chart(ax3, "Game Play & Popup Drop Chart", version, date_selected)
    charts['combo_drop'] = fig3

    return charts

def format_chart(ax, title, version, date_selected):
    ax.set_xlim(1, 100)
    ax.set_xticks(np.arange(1, 101, 1))
    ax.set_xticklabels([f"$\bf{{{x}}}$" if x % 5 == 0 else str(x) for x in range(1, 101)], fontsize=6)
    ax.set_title(f"{title} | Version {version} | {date_selected.strftime('%d-%m-%Y')}", fontsize=12, fontweight='bold')
    ax.grid(True, linestyle='--', linewidth=0.5)
    ax.tick_params(axis='x', labelsize=6)

# ========================== Step 5: Excel Generation ==========================

def generate_excel_report(processed_data, version, date_selected):
    wb = Workbook()
    wb.remove(wb.active)

    main_sheet = wb.create_sheet("MAIN_TAB")
    main_sheet.append([
        "Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
        "Total Level Drop Count", "LEVEL_Start", "USERS_starts", "LEVEL_End", "USERS_END", "Link to Sheet"
    ])

    for idx, (game_name, df) in enumerate(processed_data.items(), start=1):
        sheet = wb.create_sheet(game_name[:30])
        sheet.append([
            '=HYPERLINK("#MAIN_TAB!A1", "Back to Locate Sheet")',
            "Level", "Start Users", "Complete Users", "Game Play Drop",
            "Popup Drop", "Total Level Drop", "Retention %", "PLAY_TIME_AVG",
            "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPT_SUM"
        ])

        for _, row in df.iterrows():
            sheet.append([
                f'=HYPERLINK("#MAIN_TAB!A1", "{game_name}")',
                row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'],
                row['GAME_PLAY_DROP'], row['POPUP_DROP'], row['TOTAL_LEVEL_DROP'],
                row['RETENTION_%'], row.get('PLAY_TIME_AVG', 0),
                row.get('HINT_USED_SUM', 0), row.get('SKIPPED_SUM', 0),
                row.get('ATTEMPT_SUM', 0)
            ])

        charts = create_charts(df, version, date_selected)
        add_charts_to_sheet(sheet, charts)

        main_sheet.append([
            idx, game_name,
            (df['GAME_PLAY_DROP'] >= 0.03).sum(),
            (df['POPUP_DROP'] >= 0.03).sum(),
            (df['TOTAL_LEVEL_DROP'] >= 0.03).sum(),
            df['LEVEL'].min(), df['START_USERS'].max(),
            df['LEVEL'].max(), df['COMPLETE_USERS'].iloc[-1],
            f'=HYPERLINK("#{game_name[:30]}!A1","🔍 Analyze {game_name}")'
        ])

    format_excel(wb)
    return wb

def add_charts_to_sheet(sheet, charts):
    row_map = {'retention': 1, 'total_drop': 35, 'combo_drop': 65}
    for name, fig in charts.items():
        buf = BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        img = OpenpyxlImage(buf)
        img.anchor = f"M{row_map[name]}"
        sheet.add_image(img)

def format_excel(wb):
    header_fill = PatternFill("solid", fgColor="2C3E50")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    data_font = Font(size=10)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center')

    red_scale = {
        0.10: PatternFill(start_color="8B0000"),  # Dark Red >=10%
        0.05: PatternFill(start_color="CD5C5C"),  # Medium Red >=5%
        0.03: PatternFill(start_color="FFA07A")   # Light Red >=3%
    }

    for sheet in wb:
        sheet.freeze_panes = 'A2'

        # Format Headers
        for cell in sheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = align_center

        # Format Data
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.font = data_font
                cell.border = border
                cell.alignment = align_center

                # Percentage Formatting
                if cell.column_letter in ['D', 'E', 'F', 'G']:
                    cell.number_format = '0.00%'

                # Conditional Formatting for Drop Columns
                if cell.column_letter in ['D', 'E', 'F'] and isinstance(cell.value, (int, float)):
                    for threshold in [0.10, 0.05, 0.03]:
                        if cell.value >= threshold:
                            cell.fill = red_scale[threshold]
                            break

        # Auto-fit columns based on header text
        for col in sheet.columns:
            header_cell = col[0]
            max_length = len(str(header_cell.value)) if header_cell.value else 0
            adjusted_width = max_length + 2
            sheet.column_dimensions[get_column_letter(header_cell.column)].width = adjusted_width

# ========================== Step 6: Streamlit UI ==========================

def main():
    st.sidebar.header("Upload Files")
    start_files = st.sidebar.file_uploader("LEVEL_START Files", type=["csv", "xlsx"], accept_multiple_files=True)
    complete_files = st.sidebar.file_uploader("LEVEL_COMPLETE Files", type=["csv", "xlsx"], accept_multiple_files=True)

    version = st.sidebar.text_input("Game Version", "1.0.0")
    date_selected = st.sidebar.date_input("Analysis Date", datetime.date.today())

    if start_files and complete_files:
        with st.spinner("Processing files..."):
            processed_data = process_game_data(start_files, complete_files)

            if processed_data:
                wb = generate_excel_report(processed_data, version, date_selected)

                with tempfile.NamedTemporaryFile(delete=False) as tmp:
                    wb.save(tmp.name)
                    tmp.seek(0)
                    excel_data = tmp.read()

                st.download_button(
                    label="📥 Download Full Report",
                    data=excel_data,
                    file_name=f"Game_Analytics_{version}_{date_selected}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                selected_game = st.selectbox("Select Game to Preview", list(processed_data.keys()))
                st.dataframe(processed_data[selected_game])
                st.pyplot(create_charts(processed_data[selected_game], version, date_selected)['retention'])

# ========================== Entry Point ==========================

if __name__ == "__main__":
    main()
