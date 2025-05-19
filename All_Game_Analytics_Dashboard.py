import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
import tempfile

# ======================== STREAMLIT UI CONFIG ========================
st.set_page_config(page_title="Game Analytics Suite", layout="wide", page_icon="🎮")
st.title("🎮 Game Level Analytics Dashboard")
st.markdown("""
<style>
    .stDownloadButton button {
        background: linear-gradient(45deg, #4CAF50, #2E7D32);
        color: white !important;
        font-weight: bold;
        border-radius: 8px;
        padding: 12px 24px;
    }
    .sidebar .sidebar-content {
        background: #f5f5f5;
    }
    h1 {
        color: #2E7D32;
    }
</style>
""", unsafe_allow_html=True)

# ======================== DATA PROCESSING FUNCTIONS ========================
def clean_level(level):
    """Extract numeric value from LEVEL column with error handling"""
    try:
        return int(re.sub(r'\D', '', str(level)))
    except:
        return 0

def process_data(start_df, complete_df):
    """Process and merge dataframes with enhanced validation"""
    # Data Cleaning
    for df in [start_df, complete_df]:
        df['LEVEL'] = df['LEVEL'].apply(clean_level)
        df.sort_values(['GAME_ID', 'DIFFICULTY', 'LEVEL'], inplace=True)
        df.drop_duplicates(['GAME_ID', 'DIFFICULTY', 'LEVEL'], keep='first', inplace=True)

    # Column Renaming
    start_df = start_df.rename(columns={'USERS': 'Start Users'})
    complete_df = complete_df.rename(columns={'USERS': 'Complete Users'})

    # Merging with outer join to handle missing data
    merge_keys = ['GAME_ID', 'DIFFICULTY', 'LEVEL']
    merged_df = pd.merge(
        start_df[merge_keys + ['Start Users']],
        complete_df[merge_keys + ['Complete Users', 'PLAY_TIME_AVG', 
                   'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPTS_SUM']],
        on=merge_keys, how='outer'
    )

    # Fill NaN values after merge
    merged_df['Start Users'] = merged_df['Start Users'].fillna(0)
    merged_df['Complete Users'] = merged_df['Complete Users'].fillna(0)
    
    # Calculate metrics
    merged_df['Game Play Drop'] = merged_df['Start Users'] - merged_df['Complete Users']
    merged_df['Popup Drop'] = merged_df['Start Users'] * 0.03
    merged_df['Total Level Drop'] = merged_df['Game Play Drop'] + merged_df['Popup Drop']
    merged_df['Retention %'] = np.where(
        merged_df['Start Users'] == 0, 0,
        (merged_df['Complete Users'] / merged_df['Start Users']) * 100
    )
    
    return merged_df

# ======================== EXCEL FORMATTING ========================
def apply_sheet_formatting(ws):
    """Apply consistent formatting to Excel sheets"""
    # Header formatting
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Cell formatting
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.number_format = '0.00' if isinstance(cell.value, float) else 'General'
            cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Auto-fit columns
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width

def apply_conditional_formatting(ws, last_row):
    """Apply color scale formatting to drop columns"""
    red_scale = {
        3: PatternFill(start_color='FFC7CE', end_color='FFC7CE'),
        7: PatternFill(start_color='FF6666', end_color='FF6666'),
        10: PatternFill(start_color='8B0000', end_color='8B0000')
    }
    
    drop_columns = {'D', 'E', 'F'}  # Game Play Drop, Popup Drop, Total Level Drop
    
    for col in drop_columns:
        col_idx = ord(col) - 64  # Convert column letter to index
        for row in range(2, last_row + 2):
            cell = ws.cell(row=row, column=col_idx)
            if cell.value is not None:
                value = cell.value
                if value >= 10:
                    cell.fill = red_scale[10]
                elif value >= 7:
                    cell.fill = red_scale[7]
                elif value >= 3:
                    cell.fill = red_scale[3]
                cell.font = Font(color="FFFFFF")

# ======================== CHART GENERATION ========================
def create_charts(df, sheet_name):
    """Generate matplotlib charts with consistent styling"""
    charts = {}
    plt.style.use('seaborn')
    
    # Retention Chart
    fig1, ax1 = plt.subplots(figsize=(12, 4))
    ax1.plot(df['LEVEL'], df['Retention %'], color='#4CAF50', marker='o')
    ax1.set_title(f"{sheet_name} - Retention %", fontsize=12)
    ax1.grid(True, alpha=0.3)
    charts['retention'] = fig1
    
    # Drop Comparison Chart
    fig2, ax2 = plt.subplots(figsize=(12, 4))
    ax2.bar(df['LEVEL'] - 0.2, df['Game Play Drop'], 0.4, label='Game Play Drop')
    ax2.bar(df['LEVEL'] + 0.2, df['Popup Drop'], 0.4, label='Popup Drop')
    ax2.set_title(f"{sheet_name} - Drop Comparison", fontsize=12)
    ax2.legend()
    ax2.grid(True, alpha=0.3)
    charts['drops'] = fig2
    
    # Total Level Drop Chart
    fig3, ax3 = plt.subplots(figsize=(12, 4))
    ax3.bar(df['LEVEL'], df['Total Level Drop'], color='#F44336')
    ax3.set_title(f"{sheet_name} - Total Level Drop", fontsize=12)
    ax3.grid(True, alpha=0.3)
    charts['total_drop'] = fig3
    
    return charts

# ======================== EXCEL WORKBOOK GENERATION ========================
def generate_workbook(processed_data):
    """Create formatted Excel workbook with multiple sheets"""
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Create MAIN_TAB sheet
    main_sheet = wb.create_sheet("MAIN_TAB")
    main_headers = [
        "Index", "Sheet Name", "Game Play Drop Count", "Popup Drop Count",
        "Total Level Drop Count", "LEVEL Start", "Max Start Users", 
        "LEVEL End", "End Users", "Link to Sheet"
    ]
    main_sheet.append(main_headers)
    apply_sheet_formatting(main_sheet)

    # Create individual game sheets
    for idx, (game_key, df) in enumerate(processed_data.items(), start=1):
        sheet_name = f"{game_key}"[:31]
        ws = wb.create_sheet(sheet_name)
        
        # Add backlink to MAIN_TAB
        ws['A1'] = f'=HYPERLINK("#MAIN_TAB!A1", "🔙 Main Dashboard")'
        ws['A1'].font = Font(color="0000FF", underline="single")
        
        # Write headers
        headers = [
            'Level', 'Start Users', 'Complete Users', 'Game Play Drop',
            'Popup Drop', 'Total Level Drop', 'Retention %',
            'Play Time Avg', 'Hints Used', 'Skipped', 'Attempts'
        ]
        ws.append(headers)
        
        # Write data
        for _, row in df.iterrows():
            ws.append([
                row['LEVEL'], row['Start Users'], row['Complete Users'],
                row['Game Play Drop'], row['Popup Drop'], row['Total Level Drop'],
                row['Retention %'], row.get('PLAY_TIME_AVG', 0),
                row.get('HINT_USED_SUM', 0), row.get('SKIPPED_SUM', 0), 
                row.get('ATTEMPTS_SUM', 0)
            ])
        
        # Apply formatting
        apply_sheet_formatting(ws)
        apply_conditional_formatting(ws, len(df))
        
        # Add charts
        charts = create_charts(df, sheet_name)
        for chart_name, fig in charts.items():
            img_path = BytesIO()
            fig.savefig(img_path, format='png', dpi=150, bbox_inches='tight')
            img_path.seek(0)
            img = Image(img_path)
            img.anchor = 'M2' if chart_name == 'retention' else 'M25' if chart_name == 'drops' else 'M48'
            ws.add_image(img)
            plt.close(fig)
        
        # Update MAIN_TAB
        main_row = [
            idx, sheet_name,
            sum(df['Game Play Drop'] >= 3),
            sum(df['Popup Drop'] >= 3),
            sum(df['Total Level Drop'] >= 3),
            df['LEVEL'].min(), df['Start Users'].max(),
            df['LEVEL'].max(), df['Complete Users'].iloc[-1],
            f'=HYPERLINK("#{sheet_name}!A1", "🔍 Analyze {sheet_name}")'
        ]
        main_sheet.append(main_row)
    
    return wb

# ======================== STREAMLIT UI ========================
def main():
    st.sidebar.header("📤 Data Upload")
    start_file = st.sidebar.file_uploader("LEVEL_START.csv", type="csv")
    complete_file = st.sidebar.file_uploader("LEVEL_COMPLETE.csv", type="csv")

    if start_file and complete_file:
        with st.spinner("🔍 Analyzing your data..."):
            try:
                # Process data
                start_df = pd.read_csv(start_file)
                complete_df = pd.read_csv(complete_file)
                merged_df = process_data(start_df, complete_df)
                
                # Group data by game and difficulty
                processed_data = {}
                for (game_id, difficulty), group in merged_df.groupby(['GAME_ID', 'DIFFICULTY']):
                    key = f"{game_id}_{difficulty}"
                    processed_data[key] = group
                
                # Generate Excel report
                wb = generate_workbook(processed_data)
                
                # Save to bytes buffer
                with tempfile.NamedTemporaryFile() as tmp:
                    wb.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        excel_data = f.read()
                
                # Show success and download button
                st.success("✅ Analysis completed successfully!")
                st.balloons()
                
                col1, col2 = st.columns([1, 3])
                with col1:
                    st.download_button(
                        label="📥 Download Full Report",
                        data=excel_data,
                        file_name="Game_Analytics_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col2:
                    with st.expander("🔍 Preview Processed Data"):
                        st.dataframe(merged_df.head(10), use_container_width=True)

            except Exception as e:
                st.error(f"❌ Error processing files: {str(e)}")

if __name__ == "__main__":
    main()
