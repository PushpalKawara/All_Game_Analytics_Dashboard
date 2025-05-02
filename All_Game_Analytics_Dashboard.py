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

    for game_name in common_games:
        start_df = load_and_clean_file(start_map[game_name], is_start_file=True)
        complete_df = load_and_clean_file(complete_map[game_name], is_start_file=False)

        if start_df is not None and complete_df is not None:
            merged_df = merge_and_calculate(start_df, complete_df)
            processed_games[game_name] = merged_df

    return processed_games

def load_and_clean_file(file_obj, is_start_file=True):
    try:
        df = pd.read_csv(file_obj) if file_obj.name.endswith('.csv') else pd.read_excel(file_obj)
        df.columns = df.columns.str.strip().str.upper()

        level_col = next((col for col in df.columns if 'LEVEL' in col), None)
        if level_col:
            df['LEVEL'] = df[level_col].astype(str).str.extract('(\d+)').astype(int)

        user_col = next((col for col in df.columns if 'USER' in col), None)
        if user_col:
            df = df.rename(columns={user_col: 'START_USERS' if is_start_file else 'COMPLETE_USERS'})

        if is_start_file:
            df = df[['LEVEL', 'START_USERS']].drop_duplicates().sort_values('LEVEL')
        else:
            keep_cols = ['LEVEL', 'COMPLETE_USERS', 'PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']
            df = df[[col for col in keep_cols if col in df.columns]]

        return df.dropna().sort_values('LEVEL')
    except Exception as e:
        st.error(f"Error processing {file_obj.name}: {str(e)}")
        return None

def merge_and_calculate(start_df, complete_df):
    merged = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')

    merged['GAME_PLAY_DROP'] = ((merged['START_USERS'] - merged['COMPLETE_USERS']) / merged['START_USERS']) * 100
    merged['POPUP_DROP'] = ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1)) / merged['COMPLETE_USERS']) * 100
    merged['TOTAL_LEVEL_DROP'] = ((merged['START_USERS'] - merged['START_USERS'].shift(-1)) / merged['START_USERS']) * 100
    merged['RETENTION_%'] = (merged['START_USERS'] / merged['START_USERS'].max()) * 100

    return merged.round(2)

# ========================== Step 4: Charting Functions ==========================

def create_charts(df, version, date_selected):
    charts = {}
    df_100 = df[df['LEVEL'] <= 100].copy()

    # Retention Chart
    fig1, ax1 = plt.subplots(figsize=(15, 7))
    ax1.plot(df_100['LEVEL'], df_100['RETENTION_%'], color='#F57C00', linewidth=2)
    format_chart(ax1, "Retention Chart", version, date_selected)
    charts['retention'] = fig1

    # Total Drop Chart
    fig2, ax2 = plt.subplots(figsize=(15, 6))
    ax2.bar(df_100['LEVEL'], df_100['TOTAL_LEVEL_DROP'], color='#EF5350')
    format_chart(ax2, "Total Level Drop Chart", version, date_selected)
    charts['total_drop'] = fig2

    # Combo Drop Chart
    fig3, ax3 = plt.subplots(figsize=(15, 6))
    width = 0.4
    ax3.bar(df_100['LEVEL'] + width/2, df_100['GAME_PLAY_DROP'], width, color='#66BB6A', label='Game Play Drop')
    ax3.bar(df_100['LEVEL'] - width/2, df_100['POPUP_DROP'], width, color='#42A5F5', label='Popup Drop')
    format_chart(ax3, "Game Play & Popup Drop Chart", version, date_selected)
    ax3.legend()
    charts['combo_drop'] = fig3

    return charts

def format_chart(ax, title, version, date_selected):
    ax.set_xlim(1, 100)
    ax.set_xticks(np.arange(1, 101, 1))
    ax.set_xticklabels([f"\$\bf{{{x}}}\$" if x % 5 == 0 else str(x) for x in range(1, 101)], fontsize=6)
    ax.set_title(f"{title} | Version {version} | {date_selected.strftime('%d-%m-%Y')}",
                 fontsize=12, fontweight='bold')
    ax.grid(True, linestyle='--', linewidth=0.5)
    ax.tick_params(axis='x', labelsize=6)

# ========================== Step 5: Excel Generation ==========================

def generate_excel_report(processed_data, version, date_selected):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#D9E1F2',
            'border': 1
        })

        cell_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter'
        })

        highlight_format = workbook.add_format({
            'font_color': 'red',
            'bg_color': 'yellow',
            'align': 'center',
            'valign': 'vcenter'
        })

        # Create sheets for each game
        main_tab_data = []
        for idx, (game_name, df) in enumerate(processed_data.items(), start=1):
            sheet_name = game_name[:31]
            df_export = df.copy()
            df_export.columns = [
                "Level", "Start Users", "Complete Users", "Game Play Drop",
                "Popup Drop", "Total Level Drop", "Retention %",
                "PLAY_TIME_AVG", "HINT_USED_SUM", "SKIPPED_SUM", "ATTEMPT_SUM"
            ]
            df_export.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]

            # Apply formats
            for col_num, value in enumerate(df_export.columns):
                worksheet.write(0, col_num, value, header_format)

            for row_num in range(1, len(df_export)+1):
                for col_num in range(len(df_export.columns)):
                    value = df_export.iloc[row_num-1, col_num]
                    col_name = df_export.columns[col_num]

                    if pd.isna(value): value = ""

                    if col_name in ['Game Play Drop', 'Popup Drop', 'Total Level Drop']:
                        if isinstance(value, (int, float)) and value >= 3:
                            worksheet.write(row_num, col_num, value, highlight_format)
                        else:
                            worksheet.write(row_num, col_num, value, cell_format)
                    else:
                        worksheet.write(row_num, col_num, value, cell_format)

            # Set column widths
            for i, col in enumerate(df_export.columns):
                column_len = max(df_export[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_len)

            # Add charts
            charts = create_charts(df, version, date_selected)
            chart_positions = {'retention': 'M2', 'total_drop': 'M35', 'combo_drop': 'M68'}
            for chart_name, chart in charts.items():
                imgdata = BytesIO()
                chart.savefig(imgdata, format='png')
                imgdata.seek(0)
                worksheet.insert_image(chart_positions[chart_name], f"{chart_name}.png", {'image_data': imgdata})
                plt.close(chart)

            # Collect data for MAIN_TAB
            main_tab_data.append([
                idx, game_name,
                df['GAME_PLAY_DROP'].count(),
                df['POPUP_DROP'].count(),
                df['TOTAL_LEVEL_DROP'].count(),
                df['LEVEL'].min(), df['START_USERS'].max(),
                df['LEVEL'].max(), df['COMPLETE_USERS'].iloc[-1],
                f'=HYPERLINK("#{sheet_name}!A1","Click to view {game_name}")'
            ])

        # Create MAIN_TAB
        main_tab_df = pd.DataFrame(main_tab_data, columns=[
            "Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
            "Total Level Drop Count", "LEVEL_Start", "USERS_starts", "LEVEL_End", "USERS_END", "Link to Sheet"
        ])
        main_tab_df.to_excel(writer, sheet_name='MAIN_TAB', index=False)
        worksheet = writer.sheets['MAIN_TAB']

        # Format MAIN_TAB
        for col_num, value in enumerate(main_tab_df.columns):
            worksheet.write(0, col_num, value, header_format)

        for row_num in range(1, len(main_tab_df)+1):
            for col_num in range(len(main_tab_df.columns)):
                value = main_tab_df.iloc[row_num-1, col_num]
                worksheet.write(row_num, col_num, value, cell_format)

        # Set column widths for MAIN_TAB
        for i, col in enumerate(main_tab_df.columns):
            column_len = max(main_tab_df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)

    output.seek(0)
    return output

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
                excel_output = generate_excel_report(processed_data, version, date_selected)

                st.download_button(
                    label="📥 Download Full Report",
                    data=excel_output.getvalue(),
                    file_name=f"Game_Analytics_{version}_{date_selected.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                selected_game = st.selectbox("Select Game to Preview", list(processed_data.keys()))
                st.dataframe(processed_data[selected_game])

                charts = create_charts(processed_data[selected_game], version, date_selected)
                st.pyplot(charts['retention'])
                st.pyplot(charts['total_drop'])
                st.pyplot(charts['combo_drop'])

if __name__ == "__main__":
    main()
