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
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
import tempfile
import zipfile
import shutil

# ========================== Step 2: Configuration ========================== #
st.set_page_config(page_title="Game Analytics Pro", layout="wide")
st.title("🎮 Game Analytics Dashboard")

# ========================== Step 3: File Processing Functions ========================== #
def process_uploaded_folder(uploaded_file, folder_type):
    """Process uploaded ZIP folder and return dictionary of DataFrames"""
    temp_dir = tempfile.mkdtemp()
    game_data = {}

    try:
        # Save uploaded file
        upload_path = os.path.join(temp_dir, uploaded_file.name)
        with open(upload_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Extract if ZIP
        if zipfile.is_zipfile(upload_path):
            with zipfile.ZipFile(upload_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            os.remove(upload_path)

        # Process CSV files
        for root, _, files in os.walk(temp_dir):
            for file in files:
                if file.lower().endswith('.csv'):
                    file_path = os.path.join(root, file)
                    try:
                        df = pd.read_csv(file_path)
                        game_name = os.path.splitext(file)[0]
                        df = clean_data(df, folder_type)
                        game_data[game_name] = df
                    except Exception as e:
                        st.warning(f"Error processing {file}: {str(e)}")
        
        return game_data
    finally:
        shutil.rmtree(temp_dir)

def clean_data(df, folder_type):
    """Clean and process raw DataFrame based on folder type"""
    df.columns = df.columns.str.strip().str.upper()
    
    # Extract level number
    if 'LEVEL' in df.columns:
        df['LEVEL'] = df['LEVEL'].astype(str).str.extract(r'(\d+)').astype(int)
    
    # Select and rename columns
    if folder_type == 'start':
        df = df.rename(columns={'USERS': 'START_USERS'})
        return df[['LEVEL', 'START_USERS']].drop_duplicates().sort_values('LEVEL')
    else:
        df = df.rename(columns={'USERS': 'COMPLETE_USERS'})
        keep_cols = ['LEVEL', 'COMPLETE_USERS', 'PLAY_TIME_AVG', 
                    'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']
        return df[[c for c in keep_cols if c in df.columns]].sort_values('LEVEL')

# ========================== Step 4: Data Calculation & Merging ========================== #
def merge_and_calculate(start_df, complete_df):
    """Merge datasets and calculate metrics"""
    merged = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')
    
    # Calculate metrics
    merged['GAME_PLAY_DROP'] = ((merged['START_USERS'] - merged['COMPLETE_USERS']) / merged['START_USERS']).fillna(0) * 100
    merged['POPUP_DROP'] = ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1)) / merged['COMPLETE_USERS']).fillna(0) * 100
    merged['TOTAL_LEVEL_DROP'] = ((merged['START_USERS'] - merged['START_USERS'].shift(-1)) / merged['START_USERS']).fillna(0) * 100
    merged['RETENTION_%'] = (merged['START_USERS'] / merged['START_USERS'].max()) * 100
    
    return merged.round(2)

# ========================== Step 5: Chart Generation ========================== #
def create_game_charts(df, game_name):
    """Generate matplotlib charts for game data"""
    charts = {}
    df = df[df['LEVEL'] <= 100]
    
    # Retention Chart
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(df['LEVEL'], df['RETENTION_%'], color='#FF6B6B', linewidth=2)
    format_chart(ax, f"{game_name} Retention")
    charts['retention'] = fig
    
    # Total Drop Chart
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.bar(df['LEVEL'], df['TOTAL_LEVEL_DROP'], color='#4ECDC4')
    format_chart(ax, f"{game_name} Total Drops")
    charts['total_drop'] = fig
    
    # Combo Drop Chart
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.bar(df['LEVEL'], df['GAME_PLAY_DROP'], width=0.4, label='Game Play', color='#66BB6A')
    ax.bar(df['LEVEL']+0.4, df['POPUP_DROP'], width=0.4, label='Popup', color='#42A5F5')
    format_chart(ax, f"{game_name} Combo Drops")
    ax.legend()
    charts['combo_drop'] = fig
    
    return charts

def format_chart(ax, title):
    """Standard chart formatting"""
    ax.set_xlim(0, 100)
    ax.set_xticks(range(0, 101, 5))
    ax.grid(True, alpha=0.3)
    ax.set_title(title, fontsize=12, pad=15)
    plt.tight_layout()

# ========================== Step 6: Excel Report Generation ========================== #
def generate_excel_report(processed_data, version, date):
    """Create formatted Excel workbook with multiple sheets"""
    wb = Workbook()
    wb.remove(wb.active)
    
    # Create Main Tab
    main_sheet = wb.create_sheet("Main Tab")
    main_sheet.append([
        "Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
        "Total Level Drop Count", "LEVEL_Start", "USERS_starts", "LEVEL_End", 
        "USERS_END", "Link to Sheet"
    ])
    
    # Process each game
    for idx, (game_name, game_df) in enumerate(processed_data.items(), 1):
        # Create game sheet
        sheet = wb.create_sheet(game_name[:31])
        
        # Add headers
        sheet.append([
            "Level", "Start Users", "Complete Users",
            "Game Play Drop", "Popup Drop", "Total Level Drop",
            "Retention %"
        ])
        
        # Add data rows
        for _, row in game_df.iterrows():
            sheet.append([
                row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'],
                row['GAME_PLAY_DROP'], row['POPUP_DROP'], 
                row['TOTAL_LEVEL_DROP'], row['RETENTION_%']
            ])
        
        # Add charts
        add_charts_to_sheet(sheet, game_name, game_df, version, date)
        
        # Add to main sheet
        main_sheet.append([
            idx, game_name,
            sum(game_df['GAME_PLAY_DROP'] >= 3),
            sum(game_df['POPUP_DROP'] >= 3),
            sum(game_df['TOTAL_LEVEL_DROP'] >= 3),
            game_df['LEVEL'].min(), game_df['START_USERS'].max(),
            game_df['LEVEL'].max(), game_df['COMPLETE_USERS'].iloc[-1],
            f'=HYPERLINK("#{game_name[:31]}!A1","View")'
        ])
        
        # Format game sheet
        format_excel_sheet(sheet)
    
    # Format main sheet
    format_excel_sheet(main_sheet, is_main=True)
    return wb

def add_charts_to_sheet(sheet, game_name, df, version, date):
    """Add charts to Excel sheet at specific locations"""
    img_dir = tempfile.mkdtemp()
    
    try:
        # Generate charts
        charts = create_game_charts(df, game_name)
        
        # Save and insert charts
        chart_positions = {
            'retention': ('N2', (600, 400)),
            'total_drop': ('N39', (600, 400)),
            'combo_drop': ('N70', (600, 400))
        }
        
        for chart_type, (position, size) in chart_positions.items():
            img_path = os.path.join(img_dir, f"{chart_type}.png")
            charts[chart_type].savefig(img_path, dpi=300, bbox_inches='tight')
            img = Image(img_path)
            img.width, img.height = size
            sheet.add_image(img, position)
            
    finally:
        shutil.rmtree(img_dir)

def format_excel_sheet(sheet, is_main=False):
    """Apply consistent formatting to Excel sheets"""
    # Headers
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2c3e50")
        cell.alignment = Alignment(horizontal="center")
    
    # Data formatting
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(horizontal="center")
            
            # Conditional formatting for drops
            if not is_main and cell.column_letter in ['D', 'E', 'F']:
                if cell.value and cell.value >= 3:
                    cell.fill = PatternFill("solid", fgColor="FFFF00")
                    cell.font = Font(color="FF0000")
    
    # Auto-fit columns
    for col in sheet.columns:
        max_len = max(len(str(cell.value)) for cell in col)
        sheet.column_dimensions[col[0].column_letter].width = max_len + 2
    
    # Freeze header row
    sheet.freeze_panes = 'A2'

# ========================== Step 7: Streamlit UI ========================== #
def main():
    st.sidebar.header("⚙️ Configuration")
    version = st.sidebar.text_input("Game Version", "1.0.0")
    analysis_date = st.sidebar.date_input("Analysis Date", datetime.date.today())
    
    st.sidebar.header("📁 Data Upload")
    start_zip = st.sidebar.file_uploader("LEVEL_START Folder (ZIP)", type="zip")
    complete_zip = st.sidebar.file_uploader("LEVEL_COMPLETE Folder (ZIP)", type="zip")
    
    if start_zip and complete_zip:
        with st.spinner("Processing game data..."):
            # Process uploaded files
            start_data = process_uploaded_folder(start_zip, 'start')
            complete_data = process_uploaded_folder(complete_zip, 'complete')
            
            if not start_data or not complete_data:
                st.error("Failed to process uploaded files")
                return
            
            # Merge datasets
            processed_games = {}
            common_games = set(start_data.keys()) & set(complete_data.keys())
            
            for game in common_games:
                merged = merge_and_calculate(start_data[game], complete_data[game])
                processed_games[game] = merged
            
            if processed_games:
                # Generate Excel report
                report = generate_excel_report(processed_games, version, analysis_date)
                
                # Download button
                with tempfile.NamedTemporaryFile() as tmp:
                    report.save(tmp.name)
                    st.download_button(
                        label="📥 Download Full Report",
                        data=open(tmp.name, "rb").read(),
                        file_name=f"Game_Analytics_{version}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Show preview
                selected = st.selectbox("Select Game", list(processed_games.keys()))
                st.dataframe(processed_games[selected].head())
                
                # Show sample chart
                charts = create_game_charts(processed_games[selected], selected)
                st.pyplot(charts['retention'])
            else:
                st.warning("No matching game data found")

if __name__ == "__main__":
    main()
