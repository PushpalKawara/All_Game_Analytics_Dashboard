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
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage
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

    merged['GAME_PLAY_DROP'] = ((merged['START_USERS'] - merged['COMPLETE_USERS']) / merged['START_USERS'].replace(0, np.nan)) * 100
    merged['POPUP_DROP'] = ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1)) / merged['COMPLETE_USERS'].replace(0, np.nan)) * 100
    merged['TOTAL_LEVEL_DROP'] = ((merged['START_USERS'] - merged['START_USERS'].shift(-1)) / merged['START_USERS'].replace(0, np.nan)) * 100
    merged['RETENTION_%'] = (merged['START_USERS'] / merged['START_USERS'].max()) * 100

    return merged.round(2)

# ========================== Step 4: Enhanced Charting Functions ========================== #
def create_charts(df, version, date_selected):
    charts = {}
    df_100 = df[df['LEVEL'] <= 100].copy()
    xtick_labels = [f"$\\bf{{{x}}}$" if x % 5 == 0 else str(x) for x in range(1, 101)]

    # Retention Chart
    fig1, ax1 = plt.subplots(figsize=(15, 7))
    ax1.plot(df_100['LEVEL'], df_100['RETENTION_%'], color='#F57C00', linewidth=2)
    format_chart(ax1, "Retention Chart (Levels 1-100)", version, date_selected)
    ax1.set_ylim(0, 110)
    ax1.set_yticks(np.arange(0, 110, 5))
    
    # Add retention percentages below x-axis
    for x, y in zip(df_100['LEVEL'], df_100['RETENTION_%']):
        if not np.isnan(y):
            ax1.text(x, -5, f"{int(y)}", ha='center', va='top', fontsize=7)
    charts['retention'] = fig1

    # Total Drop Chart
    fig2, ax2 = plt.subplots(figsize=(15, 6))
    bars = ax2.bar(df_100['LEVEL'], df_100['TOTAL_LEVEL_DROP'], color='#EF5350')
    format_chart(ax2, "Total Level Drop Chart (Levels 1-100)", version, date_selected)
    ax2.set_ylim(0, max(df_100['TOTAL_LEVEL_DROP'].max(), 10) + 10)
    
    # Add drop percentages below bars
    for bar in bars:
        x = bar.get_x() + bar.get_width() / 2
        y = bar.get_height()
        ax2.text(x, -2, f"{y:.0f}", ha='center', va='top', fontsize=7)
    charts['total_drop'] = fig2

    # Combo Drop Chart
    fig3, ax3 = plt.subplots(figsize=(15, 6))
    width = 0.4
    bars1 = ax3.bar(df_100['LEVEL'] + width/2, df_100['GAME_PLAY_DROP'], width, color='#66BB6A')
    bars2 = ax3.bar(df_100['LEVEL'] - width/2, df_100['POPUP_DROP'], width, color='#42A5F5')
    format_chart(ax3, "Game Play & Popup Drop Chart (Levels 1-100)", version, date_selected)
    max_drop = max(df_100['GAME_PLAY_DROP'].max(), df_100['POPUP_DROP'].max())
    ax3.set_ylim(0, max(max_drop, 10) + 10)
    
    # Add percentages below bars
    for bar in bars1:
        x = bar.get_x() + bar.get_width() / 2
        y = bar.get_height()
        ax3.text(x, -2, f"{y:.0f}", ha='center', va='top', fontsize=7, color='#66BB6A')
    for bar in bars2:
        x = bar.get_x() + bar.get_width() / 2
        y = bar.get_height()
        ax3.text(x, -5, f"{y:.0f}", ha='center', va='top', fontsize=7, color='#42A5F5')
    charts['combo_drop'] = fig3

    return charts

def format_chart(ax, title, version, date_selected):
    ax.set_xlim(1, 100)
    ax.set_xticks(np.arange(1, 101, 1))
    ax.set_xticklabels([f"$\\bf{{{x}}}$" if x % 5 == 0 else str(x) for x in range(1, 101)], fontsize=6)
    ax.set_title(f"{title} | Version {version} | {date_selected.strftime('%d-%m-%Y')}", 
                 fontsize=12, fontweight='bold', pad=20)
    ax.grid(True, linestyle='--', linewidth=0.5)
    ax.tick_params(axis='x', labelsize=6)
    ax.set_xlabel("Level", labelpad=15)
    ax.set_ylabel("% Of Users" if "Retention" in title else "% Of Users Drop", labelpad=15)

# ========================== Step 5: Excel Generation ========================== #
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
            '=HYPERLINK("#MAIN_TAB!A1", "Back to MAIN TAB")',
            "Level", "Start Users", "Complete Users", "Game Play Drop",
            "Popup Drop", "Total Level Drop", "Retention %", "PLAY_TIME_AVG",
            "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPT_SUM"
        ])

        # Hyperlink styling
        hyperlink_cell = sheet['A1']
        hyperlink_cell.font = Font(color="FFFFFF", bold=True)
        hyperlink_cell.fill = PatternFill("solid", fgColor="0000FF")
        hyperlink_cell.alignment = Alignment(horizontal='center', vertical='center')
        hyperlink_cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )

        for _, row in df.iterrows():
            sheet.append([
                f'=HYPERLINK("#MAIN_TAB!A1", "{game_name}")',
                row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'],
                row['GAME_PLAY_DROP'], row['POPUP_DROP'], row['TOTAL_LEVEL_DROP'],
                row['RETENTION_%'], row.get('PLAY_TIME_AVG', 0),
                row.get('HINT_USED_SUM', 0), row.get('SKIPPED_SUM', 0),
                row.get('ATTEMPT_SUM', 0)
            ])

        # Add annotated charts
        charts = create_charts(df, version, date_selected)
        add_charts_to_sheet(sheet, charts)

        # Main sheet data
        game_play_drop_count = (df['GAME_PLAY_DROP'] >= 3).sum()
        popup_drop_count = (df['POPUP_DROP'] >= 3).sum()
        total_level_drop_count = (df['TOTAL_LEVEL_DROP'] >= 3).sum()

        main_row = [
            idx, game_name,
            game_play_drop_count,
            popup_drop_count,
            total_level_drop_count,
            df['LEVEL'].min(), df['START_USERS'].max(),
            df['LEVEL'].max(), df['COMPLETE_USERS'].iloc[-1],
            f'=HYPERLINK("#{sheet.title}!A1", "{game_name}")'
        ]
        main_sheet.append(main_row)

        # Hyperlink styling in main sheet
        hyperlink_cell_main = main_sheet.cell(row=main_sheet.max_row, column=10)
        hyperlink_cell_main.font = Font(color="0000FF", underline="single")

    format_workbook(wb)
    return wb

def add_charts_to_sheet(sheet, charts):
    row_map = {'retention': 1, 'total_drop': 35, 'combo_drop': 65}
    for name, fig in charts.items():
        buf = BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight', dpi=150)
        buf.seek(0)
        img = OpenpyxlImage(buf)
        img.anchor = f"M{row_map[name]}"
        sheet.add_image(img)
        plt.close(fig)  # Close figure to free memory

def format_workbook(wb):
    drop_columns = {"E", "F", "G"}
    thin_border = Border(left=Side(style='thin'), 
                        right=Side(style='thin'), 
                        top=Side(style='thin'), 
                        bottom=Side(style='thin'))

    for sheet in wb:
        sheet.freeze_panes = sheet["B2"] if sheet.title != "MAIN_TAB" else sheet["A2"]

        # Header styling
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4F81BD")
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border

        # Auto-fit and formatting
        for col in sheet.columns:
            max_length = max(
                len(str(cell.value)) if cell.value else 0
                for cell in col
            )
            sheet.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

            for cell in col:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
                if cell.column_letter in drop_columns and isinstance(cell.value, (int, float)) and sheet.title != "MAIN_TAB":
                    value = cell.value
                    if value >= 10:
                        cell.fill = PatternFill("solid", fgColor="7B241C")
                    elif value >= 5:
                        cell.fill = PatternFill("solid", fgColor="C0392B")
                    elif value >= 3:
                        cell.fill = PatternFill("solid", fgColor="F1948A")
                    cell.font = Font(color="FFFFFF")

# ========================== Step 6: Streamlit UI Integration ========================== #
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
                # Generate Excel Report
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

                # Streamlit Display
                selected_game = st.selectbox("Select Game to Preview", list(processed_data.keys()))
                df = processed_data[selected_game]
                
                st.subheader("📈 Retention Chart (Levels 1-100)")
                charts = create_charts(df, version, date_selected)
                st.pyplot(charts['retention'])
                
                st.subheader("📉 Total Drop Chart (Levels 1-100)")
                st.pyplot(charts['total_drop'])
                
                st.subheader("📉 Combo Drop Chart (Levels 1-100)")
                st.pyplot(charts['combo_drop'])

# ========================== Entry Point ========================== #
if __name__ == "__main__":
    main()
