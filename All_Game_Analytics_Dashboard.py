# ========================== IMPORTS & CONFIG ========================== #
import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import datetime
import matplotlib.pyplot as plt
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
import tempfile

st.set_page_config(page_title="GAME ANALYTICS PRO", layout="wide")
st.title("🎮 PROFESSIONAL GAME ANALYTICS DASHBOARD")

# ========================== CONSTANTS ========================== #
LEVEL_ALIASES = ['LEVEL', 'LEVELPLAYED', 'TOTALLEVELPLAYED',
                'TOTALLEVELSPLAYED', 'STAGE', 'CURRENTLEVEL']
USER_PATTERN = r'USER|PLAYER|CLIENT|COUNT'
OPTIONAL_COLS = ['PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']

# ========================== CORE PROCESSING ========================== #
def process_game_data(start_files, complete_files):
    processed = {}
    start_map = {os.path.splitext(f.name)[0].upper(): f for f in start_files}
    complete_map = {os.path.splitext(f.name)[0].upper(): f for f in complete_files}

    for game in set(start_map.keys()) & set(complete_map.keys()):
        try:
            start_df = load_clean_file(start_map[game], is_start=True)
            complete_df = load_clean_file(complete_map[game], is_start=False)

            if start_df is not None and complete_df is not None:
                merged_df = merge_and_calculate(start_df, complete_df)
                processed[game] = merged_df
        except Exception as e:
            st.error(f"Error processing {game}: {str(e)}")
    return processed

def load_clean_file(file_obj, is_start):
    try:
        df = pd.read_csv(file_obj) if file_obj.name.endswith('.csv') else pd.read_excel(file_obj)
        df.columns = df.columns.str.strip().str.upper()

        # Handle level column
        level_col = next((col for col in df.columns if col in LEVEL_ALIASES), None)
        if not level_col:
            st.warning(f"Level column not found in {file_obj.name}")
            return None
        df['LEVEL'] = df[level_col].astype(str).str.extract('(\d+)').astype(int)

        # Handle user column
        user_col = next((col for col in df.columns if re.search(USER_PATTERN, col, re.I)), None)
        if not user_col:
            st.warning(f"User column not found in {file_obj.name}")
            return None
        new_user_col = 'START_USERS' if is_start else 'COMPLETE_USERS'
        df = df.rename(columns={user_col: new_user_col})

        # Handle optional columns
        keep_cols = ['LEVEL', new_user_col]
        if not is_start:
            keep_cols += [col for col in OPTIONAL_COLS if col in df.columns]

        return df[keep_cols].dropna().sort_values('LEVEL')
    except Exception as e:
        st.error(f"Error processing {file_obj.name}: {str(e)}")
        return None

def merge_and_calculate(start_df, complete_df):
    merged = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')

    # Required calculations
    merged['GAME_PLAY_DROP'] = ((merged['START_USERS'] - merged['COMPLETE_USERS']) /
                               merged['START_USERS']).fillna(0) * 100
    merged['POPUP_DROP'] = ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1)) /
                          merged['COMPLETE_USERS']).fillna(0) * 100
    merged['TOTAL_LEVEL_DROP'] = ((merged['START_USERS'] - merged['START_USERS'].shift(-1)) /
                                merged['START_USERS']).fillna(0) * 100
    merged['RETENTION_%'] = (merged['START_USERS'] / merged['START_USERS'].max()).fillna(0) * 100

    # Add missing optional columns
    for col in OPTIONAL_COLS:
        if col not in merged:
            merged[col] = 0

    return merged.round(2)

# ========================== CHARTING SYSTEM ========================== #
def create_charts(df, version, date):
    charts = {}
    df_100 = df[df['LEVEL'] <= 100]

    # Retention Chart
    fig1, ax1 = plt.subplots(figsize=(15,7))
    ax1.plot(df_100['LEVEL'], df_100['RETENTION_%'], color='#F57C00', lw=2)
    format_retention(ax1, version, date)
    charts['retention'] = fig1

    # Total Drop Chart
    fig2, ax2 = plt.subplots(figsize=(15,6))
    bars = ax2.bar(df_100['LEVEL'], df_100['TOTAL_LEVEL_DROP'], color='#EF5350')
    format_total_drop(ax2, bars, version, date)
    charts['total_drop'] = fig2

    # Combo Drop Chart
    fig3, ax3 = plt.subplots(figsize=(15,6))
    width = 0.4
    bar1 = ax3.bar(df_100['LEVEL']+width/2, df_100['GAME_PLAY_DROP'],
                 width, color='#66BB6A', label='Game Play Drop')
    bar2 = ax3.bar(df_100['LEVEL']-width/2, df_100['POPUP_DROP'],
                 width, color='#42A5F5', label='Popup Drop')
    format_combo_drop(ax3, bar1, bar2, version, date)
    charts['combo_drop'] = fig3

    return charts

# ... (remaining functions for formatting and Excel generation continue)

def format_retention(ax, version, date):
    ax.set(xlim=(1,100), ylim=(0,110),
          xticks=np.arange(1,101), yticks=np.arange(0,110,10))
    ax.set_title(f"Retention Chart | v{version} | {date.strftime('%d-%m-%Y')}",
                fontsize=14, pad=20, weight='bold')
    ax.grid(ls='--', alpha=0.7)
    ax.tick_params(labelsize=10)

def format_total_drop(ax, bars, version, date):
    ax.set(xlim=(0.5,100.5), ylim=(0, max([b.get_height() for b in bars]+[15])))
    ax.set_title(f"Total Drops | v{version} | {date.strftime('%d-%m-%Y')}",
                fontsize=14, pad=20, weight='bold')
    ax.grid(ls='--', alpha=0.7)
    for bar in bars:
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.5,
               f"{bar.get_height():.1f}%", ha='center', va='bottom', fontsize=8)

def format_combo_drop(ax, bar1, bar2, version, date):
    max_drop = max(max([b.get_height() for b in bar1], default=0),
                  max([b.get_height() for b in bar2], default=0))
    ax.set(xlim=(0.5,100.5), ylim=(0, max(max_drop+10,15)))
    ax.set_title(f"Drop Comparison | v{version} | {date.strftime('%d-%m-%Y')}",
                fontsize=14, pad=20, weight='bold')
    ax.legend()
    ax.grid(ls='--', alpha=0.7)
    for bars in [bar1, bar2]:
        for bar in bars:
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.5,
                   f"{bar.get_height():.1f}%", ha='center', va='bottom',
                   fontsize=7, rotation=90)

# ========================== EXCEL ENGINE ========================== #
def generate_excel_report(processed_data, version, date):
    wb = Workbook()
    wb.remove(wb.active)

    # Main Sheet
    main = wb.create_sheet("MAIN_DASHBOARD")
    main.append(["ID", "Game Name", "Start Level", "Max Users",
                "End Level", "Retained Users", "Total Drops", "Link"])

    # Game Sheets
    for idx, (game, df) in enumerate(processed_data.items(), 1):
        sheet = wb.create_sheet(game[:30])
        add_game_sheet(sheet, game, df, version, date)
        main.append([
            idx, game,
            df['LEVEL'].min(), df['START_USERS'].max(),
            df['LEVEL'].max(), df['COMPLETE_USERS'].iloc[-1],
            df['TOTAL_LEVEL_DROP'].count(),
            f'=HYPERLINK("#{game[:30]}!A1","🔍 Analyze")'
        ])

    format_excel(wb)
    return wb

def add_game_sheet(sheet, game, df, version, date):
    # Header with single hyperlink
    sheet.append(['=HYPERLINK("#MAIN_DASHBOARD!A1", "🔙 Main Dashboard")'] +
                ["Level", "Start Users", "Complete Users",
                 "Game Drop%", "Popup Drop%", "Total Drop%",
                 "Retention%"] + OPTIONAL_COLS)

    # Data rows
    for _, row in df.iterrows():
        sheet.append([
            "",  # Empty cell instead of hyperlink
            row['LEVEL'],
            row['START_USERS'],
            row['COMPLETE_USERS'],
            row['GAME_PLAY_DROP'],
            row['POPUP_DROP'],
            row['TOTAL_LEVEL_DROP'],
            row['RETENTION_%']
        ] + [row.get(col, 0) for col in OPTIONAL_COLS])

    # Add charts
    charts = create_charts(df, version, date)
    add_charts_to_excel(sheet, charts)

def add_charts_to_excel(sheet, charts):
    def save_chart(fig):
        buf = BytesIO()
        fig.savefig(buf, format='png', dpi=120, bbox_inches='tight')
        plt.close(fig)
        return Image(buf)

    sheet.add_image(save_chart(charts['retention']), 'M1')
    sheet.add_image(save_chart(charts['total_drop']), 'M35')
    sheet.add_image(save_chart(charts['combo_drop']), 'M65')

def format_excel(wb):
    # Style Config
    header_fill = PatternFill("solid", fgColor="2C3E50")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    data_font = Font(size=10)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center')

    # Conditional Formatting
    red_scale = {
        12: PatternFill(start_color="8B0000"),  # Dark Red
        7: PatternFill(start_color="CD5C5C"),   # Medium Red
        3: PatternFill(start_color="FFA07A")    # Light Red
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
                if cell.column_letter in ['D','E','F','G']:
                    cell.number_format = '0.00%'

                # Red Scale Formatting
                if cell.column_letter in ['D','E','F'] and cell.value:
                    try:
                        value = float(cell.value)
                        for threshold in [12, 7, 3]:
                            if value >= threshold:
                                cell.fill = red_scale[threshold]
                                break
                    except:
                        pass

        # Column Widths
        for col in sheet.columns:
            max_len = max(len(str(cell.value)) for cell in col)
            sheet.column_dimensions[get_column_letter(col[0].column)].width = max_len * 1.3 + 2

# ========================== STREAMLIT UI ========================== #
st.subheader("1️⃣ Upload Your Game Data")

start_files = st.file_uploader("Upload START data files (CSV or Excel)", type=['csv', 'xlsx', 'xls'], accept_multiple_files=True)
complete_files = st.file_uploader("Upload COMPLETE data files (CSV or Excel)", type=['csv', 'xlsx', 'xls'], accept_multiple_files=True)

version = st.text_input("Enter Game Version (e.g., 1.0.3):", "1.0.0")
date = st.date_input("Select Date for Report:", datetime.date.today())

if st.button("📊 Process Game Analytics"):
    if not start_files or not complete_files:
        st.warning("Please upload both START and COMPLETE files to continue.")
    else:
        with st.spinner("Processing... Please wait ⏳"):
            try:
                processed_data = process_game_data(start_files, complete_files)
                if not processed_data:
                    st.warning("No matching game files were found or processed.")
                else:
                    st.success("Data processed successfully!")

                    # Display dataframes
                    for game, df in processed_data.items():
                        st.markdown(f"### 📂 Game: `{game}`")
                        st.dataframe(df)

                    # Excel Download
                    wb = generate_excel_report(processed_data, version, date)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        wb.save(tmp.name)
                        st.download_button(
                            label="📥 Download Excel Report",
                            data=open(tmp.name, "rb").read(),
                            file_name=f"GameAnalyticsReport_v{version}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            except Exception as e:
                st.error(f"Unexpected error occurred: {str(e)}")
