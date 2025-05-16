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
import datetime
import tempfile

# Set page config
st.set_page_config(page_title="Game Level Data Merger", layout="wide")
st.title("📊 Game Level Analytics Tool")

# ========================== Data Processing Functions ==========================
def clean_level(level_val):
    return ''.join(filter(str.isdigit, str(level_val)))

def process_files(start_file, complete_file):
    # Read and clean data
    df_start = pd.read_csv(start_file)
    df_complete = pd.read_csv(complete_file)
    
    for df in [df_start, df_complete]:
        df['LEVEL'] = df['LEVEL'].apply(clean_level)
        df.sort_values('LEVEL', inplace=True)
    
    # Merge datasets
    merge_cols = ['GAME_ID', 'DIFFICULTY', 'LEVEL']
    df_merge = pd.merge(
        df_start.rename(columns={'USERS': 'START_USERS'}),
        df_complete.rename(columns={'USERS': 'COMPLETE_USERS'}),
        on=merge_cols,
        how='outer'
    )
    
    # Calculate metrics
    df_merge['GAME_PLAY_DROP'] = (
        (df_merge['START_USERS'] - df_merge['COMPLETE_USERS']) / 
        df_merge['START_USERS'].replace(0, np.nan)
    ) * 100
    
    df_merge['POPUP_DROP'] = (
        (df_merge['COMPLETE_USERS'] - df_merge['START_USERS'].shift(-1)) / 
        df_merge['COMPLETE_USERS'].replace(0, np.nan)
    ) * 100
    
    df_merge['TOTAL_LEVEL_DROP'] = (
        (df_merge['START_USERS'] - df_merge['START_USERS'].shift(-1)) / 
        df_merge['START_USERS'].replace(0, np.nan)
    ) * 100
    
    df_merge['RETENTION_%'] = (
        df_merge['START_USERS'] / 
        df_merge['START_USERS'].max()
    ) * 100
    
    return df_merge.fillna(0)

# ========================== Charting Functions ==========================
def create_charts(df, version, date_selected):
    charts = {}
    df_100 = df[df['LEVEL'] <= 100].copy()

    # Retention Chart
    fig1, ax1 = plt.subplots(figsize=(15, 7))
    ax1.plot(df_100['LEVEL'], df_100['RETENTION_%'], color='#F57C00', linewidth=2)
    format_chart(ax1, "Retention Chart", version, date_selected)
    charts['retention'] = fig1

    # Total Level Drop Chart
    fig2, ax2 = plt.subplots(figsize=(15, 6))
    ax2.bar(df_100['LEVEL'], df_100['TOTAL_LEVEL_DROP'], color='#EF5350')
    format_chart(ax2, "Total Level Drop Chart", version, date_selected)
    charts['total_drop'] = fig2

    # Combined Drop Chart
    fig3, ax3 = plt.subplots(figsize=(15, 6))
    width = 0.4
    ax3.bar(df_100['LEVEL'] + width/2, df_100['GAME_PLAY_DROP'], width, 
            color='#66BB6A', label='Game Play Drop')
    ax3.bar(df_100['LEVEL'] - width/2, df_100['POPUP_DROP'], width, 
            color='#42A5F5', label='Popup Drop')
    format_chart(ax3, "Game Play & Popup Drop Chart", version, date_selected)
    charts['combo_drop'] = fig3

    return charts

def format_chart(ax, title, version, date_selected):
    ax.set_xlim(1, 100)
    ax.set_xticks(np.arange(1, 101, 1))
    ax.set_xticklabels([f"Lv{x}" if x % 5 == 0 else "" for x in range(1, 101)])
    ax.set_title(f"{title} | Version {version} | {date_selected.strftime('%d-%m-%Y')}", 
                fontsize=12, fontweight='bold')
    ax.grid(True, linestyle='--', linewidth=0.5)
    ax.tick_params(axis='x', rotation=45)

# ========================== Excel Report Generation ==========================
def generate_excel_report(data_dict, version, date_selected):
    wb = Workbook()
    wb.remove(wb.active)
    
    # Create main sheet
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_sheet.append([
        "Index", "Game ID", "Difficulty", "Level Start", "Level End",
        "Start Users", "End Users", "Retention %", "Critical Issues",
        "Game Play Drops", "Popup Drops", "Sheet Link"
    ])
    
    # Process each game variant
    for idx, (game_id, diff), in enumerate(data_dict.keys(), start=1):
        df = data_dict[(game_id, diff)]
        sheet_name = f"{game_id}_{diff}"[:31]
        ws = wb.create_sheet(sheet_name)
        
        # Add data
        ws.append(["Level", "Start Users", "Complete Users", "Game Play Drop", 
                  "Popup Drop", "Total Level Drop", "Retention %"])
        for _, row in df.iterrows():
            ws.append([
                row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'],
                row['GAME_PLAY_DROP'], row['POPUP_DROP'], 
                row['TOTAL_LEVEL_DROP'], row['RETENTION_%']
            ])
        
        # Add charts
        charts = create_charts(df, version, date_selected)
        add_charts_to_sheet(ws, charts)
        
        # Update main sheet
        main_sheet.append([
            idx, game_id, diff,
            df['LEVEL'].min(), df['LEVEL'].max(),
            df['START_USERS'].max(), df['COMPLETE_USERS'].iloc[-1],
            df['RETENTION_%'].iloc[-1],
            (df['GAME_PLAY_DROP'] > 10).sum(),
            (df['GAME_PLAY_DROP'] > 5).sum(),
            (df['POPUP_DROP'] > 5).sum(),
            f'=HYPERLINK("#{sheet_name}!A1", "View")'
        ])
    
    format_workbook(wb)
    return wb

def add_charts_to_sheet(ws, charts):
    img_height = 300
    positions = {
        'retention': 'A10',
        'total_drop': 'A30',
        'combo_drop': 'A50'
    }
    
    for chart_type, fig in charts.items():
        img_file = BytesIO()
        fig.savefig(img_file, format='png', dpi=150, bbox_inches='tight')
        img_file.seek(0)
        img = OpenpyxlImage(img_file)
        img.height = img_height
        ws.add_image(img, positions[chart_type])

def format_workbook(wb):
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    orange_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    
    for ws in wb:
        # Headers
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4F81BD")
        
        # Auto-width columns
        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 2
        
        # Conditional formatting
        if ws.title != "MAIN_TAB":
            for row in ws.iter_rows(min_row=2):
                for cell in row[3:6]:  # Drop columns
                    if cell.value is not None and isinstance(cell.value, (int, float)):
                        if cell.value >= 10:
                            cell.fill = red_fill
                        elif cell.value >= 5:
                            cell.fill = orange_fill

# ========================== Streamlit UI ==========================
def main():
    st.sidebar.header("Upload Files")
    start_file = st.sidebar.file_uploader("LEVEL_START.csv", type="csv")
    complete_file = st.sidebar.file_uploader("LEVEL_COMPLETE.csv", type="csv")
    
    st.sidebar.header("Report Settings")
    version = st.sidebar.text_input("Version", "1.0.0")
    report_date = st.sidebar.date_input("Report Date", datetime.date.today())
    
    if start_file and complete_file:
        df = process_files(start_file, complete_file)
        st.success("✅ Data processed successfully!")
        
        with st.expander("Preview Processed Data"):
            st.dataframe(df.head(50))
        
        # Generate report
        with st.spinner("Generating Excel report..."):
            data_dict = {
                (df['GAME_ID'].iloc[0], df['DIFFICULTY'].iloc[0]): df
            }
            wb = generate_excel_report(data_dict, version, report_date)
            
            # Save to temp file
            with tempfile.NamedTemporaryFile(delete=False) as tmp:
                wb.save(tmp.name)
                tmp.seek(0)
                excel_data = tmp.read()
            
            # Download button
            st.download_button(
                label="📥 Download Full Report",
                data=excel_data,
                file_name=f"Game_Analytics_{version}_{report_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Show charts
        st.header("Key Metrics Visualization")
        charts = create_charts(df, version, report_date)
        
        col1, col2 = st.columns(2)
        with col1:
            st.pyplot(charts['retention'])
            st.pyplot(charts['total_drop'])
        with col2:
            st.pyplot(charts['combo_drop'])

if __name__ == "__main__":
    main()
