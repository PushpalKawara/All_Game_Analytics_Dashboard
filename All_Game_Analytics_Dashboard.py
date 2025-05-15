import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import tempfile
import datetime
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set up Streamlit page configuration
st.set_page_config(page_title="Game Level Data Merger", layout="wide")
st.title("📊 Streamlit Tool for Merging and Analyzing Game Level Data")

# ========================== Helper Functions ========================== #
def clean_level(level_val):
    """Extract numeric values from level strings."""
    return ''.join(filter(str.isdigit, str(level_val)))

def format_sheet(workbook, worksheet, df):
    """Apply consistent formatting to Excel sheets."""
    header_format = workbook.add_format({'bold': True})
    worksheet.freeze_panes(1, 0)
    for i, col in enumerate(df.columns):
        worksheet.write(0, i, col, header_format)
        worksheet.set_column(i, i, 18)

def apply_conditional_formatting(worksheet, df, workbook):
    """Add color scales to highlight drop percentages."""
    drop_cols = ['Game Play Drop', 'Popup Drop', 'Total Level Drop']
    for col in drop_cols:
        if col in df.columns:
            col_idx = df.columns.get_loc(col)
            worksheet.conditional_format(1, col_idx, len(df), col_idx, {
                'type': 'cell',
                'criteria': '>=',
                'value': 10,
                'format': workbook.add_format({'bg_color': '#8B0000', 'font_color': '#FFFFFF'})
            })
            # Add other conditional formatting rules...

def process_data(df_start, df_complete):
    """Process and merge start/complete dataframes."""
    # Clean and merge data
    df_start['LEVEL'] = df_start['LEVEL'].apply(clean_level)
    df_complete['LEVEL'] = df_complete['LEVEL'].apply(clean_level)
    
    merged = pd.merge(
        df_start.rename(columns={'USERS': 'Start Users'}),
        df_complete.rename(columns={'USERS': 'Complete Users'}),
        on=['GAME_ID', 'DIFFICULTY', 'LEVEL'],
        how='outer'
    ).sort_values('LEVEL')
    
    # Calculate metrics
    merged['Game Play Drop'] = ((merged['Start Users'] - merged['Complete Users']) / 
                               merged['Start Users'].replace(0, np.nan)) * 100
    merged['Total Level Drop'] = ((merged['Start Users'] - merged['Start Users'].shift(-1)) / 
                                 merged['Start Users'].replace(0, np.nan)) * 100
    merged['Retention %'] = (merged['Start Users'] / merged['Start Users'].max()) * 100
    
    return merged

def create_charts(df, version, date_selected):
    """Generate matplotlib visualizations."""
    charts = {}
    df_100 = df[df['LEVEL'] <= 100].copy()
    
    # Retention chart
    fig, ax = plt.subplots(figsize=(15, 7))
    ax.plot(df_100['LEVEL'], df_100['Retention %'], color='#F57C00')
    # Add chart formatting...
    
    return charts

# ========================== Main Application Logic ========================== #
def main():
    st.sidebar.header("Upload CSV Files")
    level_start = st.sidebar.file_uploader("LEVEL_START.csv", type="csv")
    level_complete = st.sidebar.file_uploader("LEVEL_COMPLETE.csv", type="csv")
    
    version = st.sidebar.text_input("Game Version", "1.0.0")
    analysis_date = st.sidebar.date_input("Analysis Date", datetime.date.today())

    if level_start and level_complete:
        df_start = pd.read_csv(level_start)
        df_complete = pd.read_csv(level_complete)
        
        processed_df = process_data(df_start, df_complete)
        
        # Generate Excel report
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            processed_df.to_excel(writer, index=False, sheet_name='Report')
            workbook = writer.book
            worksheet = writer.sheets['Report']
            apply_conditional_formatting(worksheet, processed_df, workbook)
        
        # Download functionality
        st.download_button(
            label="Download Report",
            data=output.getvalue(),
            file_name=f"game_analysis_{version}_{analysis_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Display preview
        st.dataframe(processed_df)
        st.pyplot(create_charts(processed_df, version, analysis_date)['retention'])

if __name__ == "__main__":
    main()
