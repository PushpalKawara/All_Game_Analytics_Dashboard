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
st.title("🎮 GAME PROGRESSION ANALYTICS PRO")

# ========================== CORE PROCESSING ========================== #
def process_game_data(start_files, complete_files):
    processed = {}
    start_map = {os.path.splitext(f.name)[0].upper(): f for f in start_files}
    complete_map = {os.path.splitext(f.name)[0].upper(): f for f in complete_files}

    for game in set(start_map.keys()) & set(complete_map.keys()):
        start_df = load_clean_file(start_map[game], True)
        complete_df = load_clean_file(complete_map[game], False)

        if start_df is not None and complete_df is not None:
            merged = merge_calculate(start_df, complete_df)
            processed[game] = merged
    return processed

def load_clean_file(file_obj, is_start):
    try:
        df = pd.read_csv(file_obj) if file_obj.name.endswith('.csv') else pd.read_excel(file_obj)
        df.columns = df.columns.str.strip().str.upper()

        if 'LEVEL' in df.columns:
            df['LEVEL'] = df['LEVEL'].astype(str).str.extract('(\d+)').astype(int)

        user_col = next((c for c in df.columns if 'USER' in c), None)
        if user_col:
            df = df.rename(columns={user_col: 'START_USERS' if is_start else 'COMPLETE_USERS'})

        keep_cols = ['LEVEL', 'START_USERS'] if is_start else ['LEVEL', 'COMPLETE_USERS',
                   'PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']
        return df[keep_cols].dropna().sort_values('LEVEL')
    except Exception as e:
        st.error(f"Error processing {file_obj.name}: {str(e)}")
        return None

def merge_calculate(start_df, complete_df):
    merged = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')
    merged['GAME_PLAY_DROP'] = ((merged['START_USERS'] - merged['COMPLETE_USERS'])/merged['START_USERS'])*100
    merged['POPUP_DROP'] = ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1))/merged['COMPLETE_USERS'])*100
    merged['TOTAL_LEVEL_DROP'] = ((merged['START_USERS'] - merged['START_USERS'].shift(-1))/merged['START_USERS'])*100
    merged['RETENTION_%'] = (merged['START_USERS']/merged['START_USERS'].max())*100
    return merged.round(2)

# ========================== ADVANCED CHARTING ========================== #
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

    # Combo Drop Chart (Corrected Implementation)
    fig3, ax3 = plt.subplots(figsize=(15,6))
    width = 0.4

    # Create bars and store references
    bar1 = ax3.bar(df_100['LEVEL'] + width/2,
                 df_100['GAME_PLAY_DROP'],
                 width,
                 color='#66BB6A',
                 label='Game Play Drop')

    bar2 = ax3.bar(df_100['LEVEL'] - width/2,
                 df_100['POPUP_DROP'],
                 width,
                 color='#42A5F5',
                 label='Popup Drop')

     # Pass bars to formatting function
    format_combo_drop(ax3, bar1, bar2, version, date)
    charts['combo_drop'] = fig3


    return charts

def format_retention(ax, version, date):
    ax.set(xlim=(1,100), ylim=(0,110),
          xticks=np.arange(1,101), yticks=np.arange(0,110,5))
    ax.set_xlabel("Level", labelpad=15)
    ax.set_ylabel("% Retention", labelpad=15)
    ax.set_title(f"Retention Chart | v{version} | {date.strftime('%d-%m-%Y')}",
                fontsize=14, pad=20, weight='bold')
    ax.grid(ls='--', alpha=0.7)
    ax.tick_params(labelsize=8)
    plt.tight_layout()

def format_total_drop(ax, bars, version, date):
    ax.set(xlim=(1,100), ylim=(0, max([b.get_height() for b in bars]+[10])+10))
    ax.set_title(f"Total Level Drops | v{version} | {date.strftime('%d-%m-%Y')}",
                fontsize=14, pad=20, weight='bold')
    ax.grid(ls='--', alpha=0.7)
    ax.tick_params(labelsize=8)
    for bar in bars:
        ax.text(bar.get_x()+bar.get_width()/2, -5, f"{bar.get_height():.0f}",
               ha='center', va='top', fontsize=7)
    plt.tight_layout()


def format_combo_drop(ax, bar1, bar2, version, date):
    # Get max value safely from both bar groups
    max_drop = max(
        max([b.get_height() for b in bar1], default=0),
        max([b.get_height() for b in bar2], default=0)
    )

    # Set axis limits with buffer
    ax.set(
        xlim=(0.5, 100.5),
        ylim=(0, max(max_drop + 10, 10)),
        xticks=np.arange(1, 101),
        yticks=np.arange(0, max(max_drop + 15, 15), 5)
    )

    # Format labels and titles
    ax.set_xlabel("Game Level", fontsize=10, labelpad=10)
    ax.set_ylabel("% of Users Dropped", fontsize=10, labelpad=10)
    ax.set_title(
        f"Drop Comparison Analysis | v{version} | {date.strftime('%d-%m-%Y')}",
        fontsize=12,
        pad=15,
        weight='bold'
    )

    # Custom grid and ticks
    ax.grid(True, linestyle='--', alpha=0.6, axis='both')
    ax.tick_params(axis='x', labelsize=8, rotation=0)
    ax.tick_params(axis='y', labelsize=9)

    # Custom x-axis labels
    xtick_labels = [f"L{v}" if v % 5 == 0 else "" for v in range(1, 101)]
    ax.set_xticklabels(xtick_labels)

    # Add legend with enhanced styling
    ax.legend(
        handles=[bar1, bar2],
        loc='upper right',
        fontsize=9,
        frameon=True,
        shadow=True,
        facecolor='white'
    )

    # Add value labels
    for bars in [bar1, bar2]:
        for bar in bars:
            height = bar.get_height()
            if height > 0:  # Only label visible bars
                ax.text(
                    bar.get_x() + bar.get_width()/2,
                    height + 0.5,
                    f'{height:.1f}%',
                    ha='center',
                    va='bottom',
                    fontsize=7,
                    rotation=90
                )

    plt.tight_layout()

# ========================== EXCEL ENGINE ========================== #
def generate_excel(processed_data, version, date):
    wb = Workbook()
    wb.remove(wb.active)

    # Main Sheet Setup
    main = wb.create_sheet("MAIN_DASHBOARD")
    main.append(["ID", "Game Name", "Start Level", "Max Users", "End Level",
                "Retained Users", "Total Drops", "Analysis Link"])

    # Game Sheets
    for idx, (game, df) in enumerate(processed_data.items(), 1):
        sheet = wb.create_sheet(game[:30])
        add_game_data(sheet, game, df)
        add_charts_to_sheet(sheet, create_charts(df, version, date))

        # Main Sheet Entry
        main.append([
            idx, game,
            df['LEVEL'].min(), df['START_USERS'].max(),
            df['LEVEL'].max(), df['COMPLETE_USERS'].iloc[-1],
            df['TOTAL_LEVEL_DROP'].count(),
            f'=HYPERLINK("#{game[:30]}!A1","🔍 Analyze")'
        ])

    format_excel(wb)
    return wb

def add_game_data(sheet, game, df):
    # Headers
    sheet.append(["🔙 Main Dashboard", "Level", "Start Users", "Complete Users",
                "Game Drop%", "Popup Drop%", "Total Drop%", "Retention%",
                "Play Time", "Hints Used", "Skips", "Attempts"])

    # Data Rows
    for _, row in df.iterrows():
        sheet.append([
            f'=HYPERLINK("#MAIN_DASHBOARD!A1","{game}")',
            row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'],
            row['GAME_PLAY_DROP'], row['POPUP_DROP'], row['TOTAL_LEVEL_DROP'],
            row['RETENTION_%'], row.get('PLAY_TIME_AVG',0),
            row.get('HINT_USED_SUM',0), row.get('SKIPPED_SUM',0),
            row.get('ATTEMPT_SUM',0)
        ])

def add_charts_to_sheet(sheet, charts):
    def save_chart(fig):
        buf = BytesIO()
        fig.savefig(buf, format='png', dpi=120, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return Image(buf)

    sheet.add_image(save_chart(charts['retention']), 'M1')
    sheet.add_image(save_chart(charts['total_drop']), 'M35')
    sheet.add_image(save_chart(charts['combo_drop']), 'M65')

def format_excel(wb):
    # Style Config
    header_style = Font(bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill("solid", fgColor="2C3E50")
    data_font = Font(size=10)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center')

    # Conditional Formatting
    red_scale = {
        12: PatternFill(start_color="8B0000", end_color="8B0000"),  # Dark Red
        7: PatternFill(start_color="CD5C5C", end_color="CD5C5C"),   # Medium Red
        3: PatternFill(start_color="FFA07A", end_color="FFA07A")    # Light Red
    }

    for sheet in wb:
        sheet.freeze_panes = 'A2'

        # Headers
        for cell in sheet[1]:
            cell.font = header_style
            cell.fill = header_fill
            cell.border = border
            cell.alignment = align_center

        # Data Formatting
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.font = data_font
                cell.border = border
                cell.alignment = align_center

                # Percentage Formatting
                if cell.column_letter in ['D','E','F','G']:
                    cell.number_format = '0.00%'

                # Red Scale Formatting
                if cell.column_letter in ['D','E','F']:
                    try:
                        value = float(cell.value)
                        for threshold in sorted(red_scale.keys(), reverse=True):
                            if value >= threshold:
                                cell.fill = red_scale[threshold]
                                break
                    except:
                        pass

        # Column Sizing
        for col in sheet.columns:
            sheet.column_dimensions[get_column_letter(col[0].column)].width = \
            max(len(str(cell.value)) * 1.3 for cell in col) + 2

# ========================== STREAMLIT UI ========================== #
def main():
    st.sidebar.header("⚙️ DATA UPLOAD")
    start_files = st.sidebar.file_uploader("LEVEL_START FILES",
                      type=["csv","xlsx"], accept_multiple_files=True)
    complete_files = st.sidebar.file_uploader("LEVEL_COMPLETE FILES",
                      type=["csv","xlsx"], accept_multiple_files=True)

    version = st.sidebar.text_input("VERSION", "1.0.0")
    date = st.sidebar.date_input("REPORT DATE", datetime.date.today())

    if start_files and complete_files:
        with st.spinner("🧠 ANALYZING GAME DATA..."):
            processed = process_game_data(start_files, complete_files)

            if processed:
                # Excel Report
                with tempfile.NamedTemporaryFile(delete=False) as tmp:
                    generate_excel(processed, version, date).save(tmp.name)
                    st.download_button(
                        "💾 DOWNLOAD FULL REPORT",
                        data=open(tmp.name, "rb").read(),
                        file_name=f"Game_Analytics_{version}_{date}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # Preview Section
                game = st.selectbox("SELECT GAME PREVIEW", list(processed.keys()))
                df = processed[game]

                # Formatted Display
                st.dataframe(df.style.format({
                    'GAME_PLAY_DROP': '{:.2f}%',
                    'POPUP_DROP': '{:.2f}%',
                    'TOTAL_LEVEL_DROP': '{:.2f}%',
                    'RETENTION_%': '{:.2f}%'
                }).apply(highlight_drops, axis=None))

                # Interactive Charts
                charts = create_charts(df, version, date)
                st.pyplot(charts['retention'])
                st.pyplot(charts['total_drop'])
                st.pyplot(charts['combo_drop'])

def highlight_drops(val):
    try:
        num = float(val)
        return [
            f"background: #8B0000; color: white" if num >=12 else
            f"background: #CD5C5C; color: white" if num >=7 else
            f"background: #FFA07A" if num >=3 else ""
        ][0]
    except:
        return ""

if __name__ == "__main__":
    main()
