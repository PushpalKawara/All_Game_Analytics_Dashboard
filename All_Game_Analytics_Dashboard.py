import streamlit as st
import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
import tempfile

# ======================== STREAMLIT UI CONFIG ========================
st.set_page_config(page_title="Game Analytics Pro", layout="wide", page_icon="📊")
st.title("🎮 Game Level Analysis Dashboard")
# ... (keep the same CSS styles from previous version)

# ======================== DATA PROCESSING ========================
def standardize_column_names(df):
    """Standardize column names to handle variations"""
    df.columns = (df.columns.str.strip()
                  .str.lower()
                  .str.replace('[^a-z0-9]+', '_', regex=True))
    return df

def clean_level(level):
    """Extract numeric level value with enhanced parsing"""
    try:
        return int(re.sub(r'\D', '', str(level)))
    except ValueError:
        return 0

def validate_columns(df, required_cols, file_type):
    """Validate required columns exist in dataframe"""
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in {file_type} file: {', '.join(missing)}")
    return df

def process_data(start_df, complete_df):
    """Process and merge data with comprehensive validation"""
    # Standardize column names
    start_df = standardize_column_names(start_df)
    complete_df = standardize_column_names(complete_df)

    # Validate required columns
    start_cols = ['game_id', 'difficulty', 'level', 'users']
    complete_cols = ['game_id', 'difficulty', 'level', 'users', 
                    'play_time_avg', 'hint_used_sum', 'skipped_sum', 'attempts_sum']
    
    start_df = validate_columns(start_df, start_cols, "LEVEL_START")
    complete_df = validate_columns(complete_df, complete_cols, "LEVEL_COMPLETE")

    # Data Cleaning
    for df in [start_df, complete_df]:
        df['level'] = df['level'].apply(clean_level)
        df.sort_values(['game_id', 'difficulty', 'level'], inplace=True)
        df.drop_duplicates(['game_id', 'difficulty', 'level'], keep='first', inplace=True)

    # Column Renaming
    start_df = start_df.rename(columns={'users': 'start_users'})
    complete_df = complete_df.rename(columns={
        'users': 'complete_users',
        'play_time_avg': 'play_time_avg',
        'hint_used_sum': 'hint_used_sum',
        'skipped_sum': 'skipped_sum',
        'attempts_sum': 'attempts_sum'
    })

    # Merging with outer join
    merge_keys = ['game_id', 'difficulty', 'level']
    merged_df = pd.merge(
        start_df[merge_keys + ['start_users']],
        complete_df[merge_keys + ['complete_users', 'play_time_avg',
                   'hint_used_sum', 'skipped_sum', 'attempts_sum']],
        on=merge_keys, how='outer', suffixes=('', '_y')
    )
    
    # Clean merged columns
    merged_df = merged_df.loc[:,~merged_df.columns.duplicated()]
    
    # Fill NaN values and convert to appropriate types
    numeric_cols = ['start_users', 'complete_users', 'play_time_avg',
                    'hint_used_sum', 'skipped_sum', 'attempts_sum']
    merged_df[numeric_cols] = merged_df[numeric_cols].fillna(0)
    merged_df[numeric_cols] = merged_df[numeric_cols].apply(pd.to_numeric, errors='coerce')
    
    # Calculate metrics
    merged_df['game_play_drop'] = merged_df['start_users'] - merged_df['complete_users']
    merged_df['popup_drop'] = (merged_df['start_users'] * 0.03).round(2)
    merged_df['total_level_drop'] = merged_df['game_play_drop'] + merged_df['popup_drop']
    merged_df['retention_%'] = np.where(
        merged_df['start_users'] == 0, 0,
        (merged_df['complete_users'] / merged_df['start_users']) * 100
    ).round(2)
    
    return merged_df

# ======================== EXCEL FORMATTING ========================
# ... (keep the same formatting functions from previous version but update column names)

# ======================== CHART GENERATION ========================
# ... (update chart code to use new column names: 'level' -> 'Level', etc.)

# ======================== EXCEL WORKBOOK GENERATION ========================
def generate_workbook(processed_data):
    """Create comprehensive Excel report"""
    wb = Workbook()
    wb.remove(wb.active)
    
    # Create MAIN_TAB
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_headers = [
        "Index", "Game/Difficulty", "High Game Drops", "High Popup Drops",
        "Total Issues", "First Level", "Max Players", "Last Level", 
        "Final Players", "Analysis Link"
    ]
    main_sheet.append(main_headers)
    apply_sheet_formatting(main_sheet)
    
    # Create game sheets
    for idx, (game_key, df) in enumerate(processed_data.items(), start=1):
        sheet_name = f"{game_key}"[:31]
        ws = wb.create_sheet(sheet_name)
        
        # Write headers
        headers = [
            'Level', 'Start Users', 'Complete Users', 'Game Play Drop',
            'Popup Drop', 'Total Level Drop', 'Retention %',
            'Avg Play Time', 'Hints Used', 'Skips', 'Attempts'
        ]
        ws.append(headers)
        
        # Write data
        for _, row in df.iterrows():
            ws.append([
                row['level'], 
                int(row['start_users']),
                int(row['complete_users']),
                float(row['game_play_drop']),
                float(row['popup_drop']),
                float(row['total_level_drop']),
                float(row['retention_%']),
                float(row.get('play_time_avg', 0)),
                int(row.get('hint_used_sum', 0)),
                int(row.get('skipped_sum', 0)),
                int(row.get('attempts_sum', 0))
            ])
        
        # ... (rest of the Excel generation code remains the same)

# ======================== STREAMLIT INTERFACE ========================
def main():
    st.sidebar.header("📁 Data Upload")
    start_file = st.sidebar.file_uploader("LEVEL_START.csv", type="csv")
    complete_file = st.sidebar.file_uploader("LEVEL_COMPLETE.csv", type="csv")
    
    if start_file and complete_file:
        with st.spinner("🔍 Analyzing game data..."):
            try:
                # Read and process data
                start_df = pd.read_csv(start_file)
                complete_df = pd.read_csv(complete_file)
                
                merged_df = process_data(start_df, complete_df)
                
                # ... (rest of the main function remains the same)

if __name__ == "__main__":
    main()
