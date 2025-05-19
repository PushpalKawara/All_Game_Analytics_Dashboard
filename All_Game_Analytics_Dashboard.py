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
st.markdown("""
<style>
    .stDownloadButton button {
        background: linear-gradient(45deg, #4CAF50, #2E7D32);
        color: white !important;
        border-radius: 8px;
        padding: 12px 24px;
        transition: all 0.3s;
    }
    .stDownloadButton button:hover {
        transform: scale(1.05);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    .sidebar .sidebar-content {
        background: #f8f9fa;
        padding: 20px;
        border-right: 2px solid #dee2e6;
    }
    h1 {
        color: #2E7D32;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

# ======================== DATA PROCESSING ========================
def clean_level(level):
    """Extract numeric level value with enhanced parsing"""
    try:
        return int(re.sub(r'^(Level|LEVEL)[_ ]?', '', str(level), 10)
    except ValueError:
        return 0

def process_data(start_df, complete_df):
    """Process and merge data with comprehensive validation"""
    # Data Cleaning
    for df in [start_df, complete_df]:
        df['LEVEL'] = df['LEVEL'].apply(clean_level)
        df.sort_values(['GAME_ID', 'DIFFICULTY', 'LEVEL'], inplace=True)
        df.drop_duplicates(['GAME_ID', 'DIFFICULTY', 'LEVEL'], keep='first', inplace=True)

    # Column Renaming
    start_df = start_df.rename(columns={'USERS': 'Start Users'})
    complete_df = complete_df.rename(columns={'USERS': 'Complete Users'})

    # Merging with outer join
    merge_keys = ['GAME_ID', 'DIFFICULTY', 'LEVEL']
    merged_df = pd.merge(
        start_df[merge_keys + ['Start Users']],
        complete_df[merge_keys + ['Complete Users', 'PLAY_TIME_AVG',
                   'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPTS_SUM']],
        on=merge_keys, how='outer', suffixes=('', '_y')
    
    # Clean merged columns
    merged_df = merged_df.loc[:,~merged_df.columns.duplicated()]
    
    # Fill NaN values
    numeric_cols = ['Start Users', 'Complete Users', 'PLAY_TIME_AVG',
                    'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPTS_SUM']
    merged_df[numeric_cols] = merged_df[numeric_cols].fillna(0)
    
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
    """Apply professional formatting to Excel sheets"""
    # Header formatting
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)
    
    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=Side(border_style="thick", color="000000"))

    # Data cell formatting
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00' if abs(cell.value) >= 1000 else '0.00'
    
    # Auto-fit columns
    for col in ws.columns:
        max_length = max(
            (len(str(cell.value)) for cell in col),
            default=0
        )
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width

def apply_conditional_formatting(ws, last_row):
    """Apply three-color scale conditional formatting"""
    red_scale = {
        3: PatternFill(start_color='FFC7CE', end_color='FFC7CE'),
        7: PatternFill(start_color='FF6666', end_color='FF6666'),
        10: PatternFill(start_color='8B0000', end_color='8B0000')
    }
    
    drop_cols = {'D', 'E', 'F'}  # Game Play Drop, Popup Drop, Total Level Drop
    
    for col in drop_cols:
        col_idx = ord(col) - 64
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
                cell.font = Font(color="FFFFFF", bold=True)

# ======================== CHART GENERATION ========================
def create_charts(df, sheet_name):
    """Generate professional matplotlib charts"""
    plt.style.use('seaborn-whitegrid')
    charts = {}
    
    # Retention Chart
    fig1, ax1 = plt.subplots(figsize=(12, 4))
    ax1.plot(df['LEVEL'], df['Retention %'], 
            marker='o', color='#4CAF50', linewidth=2)
    ax1.set_title(f"{sheet_name} Retention Rate", fontsize=14, pad=20)
    ax1.set_xlabel("Level", fontsize=12)
    ax1.set_ylabel("Retention (%)", fontsize=12)
    ax1.grid(True, linestyle='--', alpha=0.7)
    charts['retention'] = fig1
    
    # Drop Analysis Chart
    fig2, ax2 = plt.subplots(figsize=(12, 4))
    width = 0.35
    x = np.arange(len(df['LEVEL']))
    ax2.bar(x - width/2, df['Game Play Drop'], width, label='Game Play Drop')
    ax2.bar(x + width/2, df['Popup Drop'], width, label='Popup Drop')
    ax2.set_title(f"{sheet_name} Drop Analysis", fontsize=14, pad=20)
    ax2.set_xticks(x)
    ax2.set_xticklabels(df['LEVEL'])
    ax2.legend()
    ax2.grid(True, axis='y', linestyle='--', alpha=0.7)
    charts['drops'] = fig2
    
    # Total Level Drop Chart
    fig3, ax3 = plt.subplots(figsize=(12, 4))
    ax3.bar(df['LEVEL'], df['Total Level Drop'], color='#F44336')
    ax3.set_title(f"{sheet_name} Total Level Drop", fontsize=14, pad=20)
    ax3.set_xlabel("Level", fontsize=12)
    ax3.grid(True, axis='y', linestyle='--', alpha=0.7)
    charts['total_drop'] = fig3
    
    return charts

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
        
        # Add navigation
        ws['A1'] = f'=HYPERLINK("#MAIN_TAB!A1", "🏠 Main Dashboard")'
        ws['A1'].font = Font(color="0000FF", underline="single")
        
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
        chart_positions = {
            'retention': 'M2',
            'drops': 'M25',
            'total_drop': 'M48'
        }
        for chart_name, fig in charts.items():
            img_bytes = BytesIO()
            fig.savefig(img_bytes, format='png', dpi=150, bbox_inches='tight')
            img_bytes.seek(0)
            img = Image(img_bytes)
            img.anchor = chart_positions[chart_name]
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

# ======================== STREAMLIT INTERFACE ========================
def main():
    st.sidebar.header("📁 Data Upload")
    start_file = st.sidebar.file_uploader("LEVEL_START.csv", type="csv")
    complete_file = st.sidebar.file_uploader("LEVEL_COMPLETE.csv", type="csv")
    
    if start_file and complete_file:
        with st.spinner("🔍 Analyzing game data..."):
            try:
                # Process data
                start_df = pd.read_csv(start_file)
                complete_df = pd.read_csv(complete_file)
                merged_df = process_data(start_df, complete_df)
                
                # Group data
                processed_data = {}
                for (game_id, difficulty), group in merged_df.groupby(['GAME_ID', 'DIFFICULTY']):
                    key = f"{game_id}_{difficulty}"
                    processed_data[key] = group
                
                # Generate Excel
                wb = generate_workbook(processed_data)
                
                # Save to bytes
                with tempfile.NamedTemporaryFile() as tmp:
                    wb.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        excel_bytes = f.read()
                
                # Show results
                st.success("✅ Analysis completed successfully!")
                st.balloons()
                
                col1, col2 = st.columns([1, 2])
                with col1:
                    st.download_button(
                        label="📥 Download Full Report",
                        data=excel_bytes,
                        file_name="Game_Analysis_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                with col2:
                    with st.expander("📄 Preview Processed Data"):
                        st.dataframe(merged_df.head(8), use_container_width=True)
            
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

if __name__ == "__main__":
    main()
