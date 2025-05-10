# ========================== IMPORTS & CONFIGURATION ========================== #
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

# Configure Streamlit
st.set_page_config(page_title="GAME ANALYTICS PRO", layout="wide")
st.title("🎮 PROFESSIONAL GAME PROGRESSION ANALYTICS")

# ========================== DATA PROCESSING FUNCTIONS ========================== #
def process_game_data(start_files, complete_files):
    """Process and merge game data from start and complete files"""
    processed = {}
    start_map = {os.path.splitext(f.name)[0].upper(): f for f in start_files}
    complete_map = {os.path.splitext(f.name)[0].upper(): f for f in complete_files}
    
    common_games = set(start_map.keys()) & set(complete_map.keys())
    
    for game in common_games:
        try:
            start_df = load_clean_file(start_map[game], is_start=True)
            complete_df = load_clean_file(complete_map[game], is_start=False)
            
            if start_df is not None and complete_df is not None:
                merged_df = merge_and_calculate(start_df, complete_df)
                processed[game] = merged_df
        except Exception as e:
            st.error(f"Error processing {game}: {str(e)}")
    
    return processed

def load_clean_file(file_obj, is_start=True):
    """Load and clean game data files"""
    try:
        # Read file
        df = pd.read_csv(file_obj) if file_obj.name.endswith('.csv') else pd.read_excel(file_obj)
        df.columns = df.columns.str.strip().str.upper()
        
        # Process level column
        if 'LEVEL' in df.columns:
            df['LEVEL'] = df['LEVEL'].astype(str).str.extract('(\d+)').astype(int)
        
        # Rename user column
        user_col = next((col for col in df.columns if 'USER' in col), None)
        if user_col:
            new_name = 'START_USERS' if is_start else 'COMPLETE_USERS'
            df = df.rename(columns={user_col: new_name})
        
        # Select relevant columns
        if is_start:
            return df[['LEVEL', 'START_USERS']].drop_duplicates().sort_values('LEVEL')
        else:
            cols = ['LEVEL', 'COMPLETE_USERS', 'PLAY_TIME_AVG', 
                    'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']
            return df[[col for col in cols if col in df.columns]].dropna()
            
    except Exception as e:
        st.error(f"Error processing {file_obj.name}: {str(e)}")
        return None

def merge_and_calculate(start_df, complete_df):
    """Merge start and complete data and calculate metrics"""
    merged = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')
    
    # Calculate metrics
    merged['GAME_PLAY_DROP'] = ((merged['START_USERS'] - merged['COMPLETE_USERS']) / 
                               merged['START_USERS']) * 100
    merged['POPUP_DROP'] = ((merged['COMPLETE_USERS'] - merged['START_USERS'].shift(-1)) / 
                           merged['COMPLETE_USERS']) * 100
    merged['TOTAL_LEVEL_DROP'] = ((merged['START_USERS'] - merged['START_USERS'].shift(-1)) / 
                                 merged['START_USERS']) * 100
    merged['RETENTION_%'] = (merged['START_USERS'] / merged['START_USERS'].max()) * 100
    
    return merged.round(2)

# ========================== CHARTING FUNCTIONS ========================== #
def create_charts(df, version, date_selected):
    """Create all three analytics charts"""
    charts = {}
    df_100 = df[df['LEVEL'] <= 100].copy()
    
    # Retention Chart
    retention_fig, ax1 = plt.subplots(figsize=(15, 7))
    ax1.plot(df_100['LEVEL'], df_100['RETENTION_%'], 
            color='#F57C00', linewidth=2, label='Retention')
    format_retention_chart(ax1, version, date_selected)
    charts['retention'] = retention_fig
    
    # Total Drop Chart
    drop_fig, ax2 = plt.subplots(figsize=(15, 6))
    bars = ax2.bar(df_100['LEVEL'], df_100['TOTAL_LEVEL_DROP'], color='#EF5350')
    format_total_drop_chart(ax2, bars, version, date_selected)
    charts['total_drop'] = drop_fig
    
    # Combo Drop Chart (Fixed Implementation)
    combo_fig, ax3 = plt.subplots(figsize=(15, 6))
    width = 0.4
    bar1 = ax3.bar(df_100['LEVEL'] + width/2, df_100['GAME_PLAY_DROP'], 
                 width, color='#66BB6A', label='Game Play Drop')
    bar2 = ax3.bar(df_100['LEVEL'] - width/2, df_100['POPUP_DROP'], 
                 width, color='#42A5F5', label='Popup Drop')
    format_combo_drop_chart(ax3, bar1, bar2, version, date_selected)
    charts['combo_drop'] = combo_fig
    
    return charts

def format_retention_chart(ax, version, date):
    """Format retention chart appearance"""
    ax.set(xlim=(1, 100), ylim=(0, 110),
          xticks=np.arange(1, 101, 1), yticks=np.arange(0, 110, 10))
    ax.set_xlabel("Level", labelpad=15, fontsize=12)
    ax.set_ylabel("% Retention", labelpad=15, fontsize=12)
    ax.set_title(f"User Retention | Version {version} | {date.strftime('%d-%m-%Y')}", 
                fontsize=14, pad=20, weight='bold')
    ax.grid(True, linestyle='--', alpha=0.6)
    ax.tick_params(labelsize=10)
    ax.legend(loc='lower left', fontsize=10)
    plt.tight_layout()

def format_total_drop_chart(ax, bars, version, date):
    """Format total drop chart appearance"""
    max_drop = max([bar.get_height() for bar in bars], default=0)
    ax.set(xlim=(0.5, 100.5), ylim=(0, max(max_drop + 10, 15)),
          xticks=np.arange(1, 101, 1))
    ax.set_xlabel("Level", fontsize=12)
    ax.set_ylabel("% Drop Rate", fontsize=12)
    ax.set_title(f"Total Drop Rate | Version {version} | {date.strftime('%d-%m-%Y')}", 
                fontsize=14, pad=20, weight='bold')
    ax.grid(True, linestyle='--', alpha=0.6)
    
    # Add value labels
    for bar in bars:
        height = bar.get_height()
        if height > 0:
            ax.text(bar.get_x() + bar.get_width()/2, height + 0.5,
                   f'{height:.1f}%', ha='center', va='bottom', fontsize=8)
    
    plt.tight_layout()

def format_combo_drop_chart(ax, bar1, bar2, version, date):
    """Format combo drop chart appearance (Fixed Version)"""
    # Calculate max drop from both bar groups
    max_drop = max(
        max([b.get_height() for b in bar1], default=0),
        max([b.get_height() for b in bar2], default=0)
    )
    
    ax.set(xlim=(0.5, 100.5), ylim=(0, max(max_drop + 10, 15)),
          xticks=np.arange(1, 101, 1))
    ax.set_xlabel("Level", fontsize=12)
    ax.set_ylabel("% Drop Rate", fontsize=12)
    ax.set_title(f"Drop Rate Comparison | Version {version} | {date.strftime('%d-%m-%Y')}", 
                fontsize=14, pad=20, weight='bold')
    ax.grid(True, linestyle='--', alpha=0.6)
    
    # Add value labels
    for bars in [bar1, bar2]:
        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax.text(bar.get_x() + bar.get_width()/2, height + 0.5,
                       f'{height:.1f}%', ha='center', va='bottom', 
                       fontsize=7, rotation=90)
    
    ax.legend(loc='upper right', fontsize=10)
    plt.tight_layout()

# ========================== EXCEL REPORT GENERATION ========================== #
def generate_excel_report(processed_data, version, date_selected):
    """Generate Excel report with formatted sheets and charts"""
    wb = Workbook()
    wb.remove(wb.active)
    
    # Create main dashboard sheet
    main_sheet = wb.create_sheet("MAIN_DASHBOARD")
    main_sheet.append([
        "ID", "Game Name", "Start Level", "Max Users", 
        "End Level", "Retained Users", "Total Drops", "Link"
    ])
    
    # Create game-specific sheets
    for idx, (game_name, df) in enumerate(processed_data.items(), start=1):
        sheet = wb.create_sheet(game_name[:30])  # Excel sheet name limit
        
        # Add headers
        sheet.append([
            '=HYPERLINK("#MAIN_DASHBOARD!A1", "🔙 Dashboard")',
            "Level", "Start Users", "Complete Users", 
            "Game Drop%", "Popup Drop%", "Total Drop%", 
            "Retention%", "Play Time", "Hints Used", 
            "Skips", "Attempts"
        ])
        
        # Add data rows
        for _, row in df.iterrows():
            sheet.append([
                f'=HYPERLINK("#MAIN_DASHBOARD!A1", "{game_name}")',
                row['LEVEL'], row['START_USERS'], row['COMPLETE_USERS'],
                row['GAME_PLAY_DROP'], row['POPUP_DROP'], 
                row['TOTAL_LEVEL_DROP'], row['RETENTION_%'],
                row.get('PLAY_TIME_AVG', 0), row.get('HINT_USED_SUM', 0),
                row.get('SKIPPED_SUM', 0), row.get('ATTEMPT_SUM', 0)
            ])
        
        # Add charts
        charts = create_charts(df, version, date_selected)
        add_charts_to_sheet(sheet, charts)
        
        # Add main sheet entry
        main_sheet.append([
            idx, game_name,
            df['LEVEL'].min(), df['START_USERS'].max(),
            df['LEVEL'].max(), df['COMPLETE_USERS'].iloc[-1],
            df['TOTAL_LEVEL_DROP'].count(),
            f'=HYPERLINK("#{game_name[:30]}!A1","🔍 Analyze")'
        ])
    
    format_excel_workbook(wb)
    return wb

def add_charts_to_sheet(sheet, charts):
    """Add matplotlib charts to Excel sheet as images"""
    def save_figure(fig):
        buf = BytesIO()
        fig.savefig(buf, format='png', dpi=120, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return Image(buf)
    
    # Insert charts at specific positions
    sheet.add_image(save_figure(charts['retention']), 'M1')
    sheet.add_image(save_figure(charts['total_drop']), 'M35')
    sheet.add_image(save_figure(charts['combo_drop']), 'M65')

def format_excel_workbook(wb):
    """Apply professional formatting to Excel workbook"""
    # Style definitions
    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill("solid", fgColor="2C3E50")
    data_font = Font(size=10)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal='center', vertical='center')
    
    # Conditional formatting colors
    red_scale = {
        12: PatternFill(start_color="8B0000"),  # Dark red
        7: PatternFill(start_color="CD5C5C"),   # Medium red
        3: PatternFill(start_color="FFA07A")    # Light red
    }

    for sheet in wb:
        # Freeze header row
        sheet.freeze_panes = 'A2'
        
        # Format headers
        for cell in sheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = center_alignment
        
        # Format data cells
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = center_alignment
                
                # Percentage formatting
                if cell.column_letter in ['D', 'E', 'F', 'G']:
                    cell.number_format = '0.00%'
                
                # Red scale conditional formatting
                if cell.column_letter in ['D', 'E', 'F']:
                    try:
                        value = float(cell.value)
                        for threshold in sorted(red_scale.keys(), reverse=True):
                            if value >= threshold:
                                cell.fill = red_scale[threshold]
                                break
                    except:
                        pass
        
        # Auto-adjust column widths
        for col in sheet.columns:
            max_length = max(len(str(cell.value)) for cell in col)
            adjusted_width = (max_length + 2) * 1.2
            sheet.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width

# ========================== STREAMLIT UI ========================== #
def main():
    """Main Streamlit application"""
    st.sidebar.header("📤 DATA UPLOAD")
    start_files = st.sidebar.file_uploader(
        "LEVEL START FILES",
        type=["csv", "xlsx"],
        accept_multiple_files=True
    )
    complete_files = st.sidebar.file_uploader(
        "LEVEL COMPLETE FILES", 
        type=["csv", "xlsx"],
        accept_multiple_files=True
    )
    
    version = st.sidebar.text_input("VERSION", "1.0.0")
    date_selected = st.sidebar.date_input("REPORT DATE", datetime.date.today())
    
    if start_files and complete_files:
        with st.spinner("🔍 ANALYZING GAME DATA..."):
            processed_data = process_game_data(start_files, complete_files)
            
            if processed_data:
                # Generate Excel report
                with tempfile.NamedTemporaryFile(delete=False) as tmp:
                    wb = generate_excel_report(processed_data, version, date_selected)
                    wb.save(tmp.name)
                    
                    st.download_button(
                        label="📥 DOWNLOAD FULL REPORT",
                        data=open(tmp.name, "rb").read(),
                        file_name=f"Game_Analytics_{version}_{date_selected}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Show preview
                selected_game = st.selectbox("SELECT GAME PREVIEW", list(processed_data.keys()))
                df = processed_data[selected_game]
                
                # Display styled dataframe
                st.dataframe(
                    df.style.format({
                        'GAME_PLAY_DROP': '{:.2f}%',
                        'POPUP_DROP': '{:.2f}%',
                        'TOTAL_LEVEL_DROP': '{:.2f}%',
                        'RETENTION_%': '{:.2f}%'
                    }).applymap(highlight_drops)
                )
                
                # Display interactive charts
                charts = create_charts(df, version, date_selected)
                st.pyplot(charts['retention'])
                st.pyplot(charts['total_drop'])
                st.pyplot(charts['combo_drop'])

def highlight_drops(value):
    """Highlight drop percentages in Streamlit dataframe"""
    try:
        num = float(value)
        if num >= 12:
            return 'background-color: #8B0000; color: white'
        elif num >= 7:
            return 'background-color: #CD5C5C; color: white'
        elif num >= 3:
            return 'background-color: #FFA07A'
    except:
        return ''

if __name__ == "__main__":
    main()
