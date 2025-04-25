# ========================== Step 1: Required Imports & Folder Paths ========================== #
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
import re
import datetime
import matplotlib.pyplot as plt
from io import BytesIO
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
import zipfile
import tempfile
import shutil

# Set page config
st.set_page_config(page_title="Game_Analytics", layout="wide")
st.title("📊 ALL_Game_Analytics Dashboard")

# -------------------- HELPER FUNCTIONS -------------------- #
def clean_level(x):
    """Clean level values to extract numeric part"""
    try:
        return int(re.search(r"(\d+)", str(x)).group(1))
    except:
        return None

def process_file(df, file_type):
    """Process and clean either start or complete file"""
    df.columns = df.columns.str.strip().str.upper()

    level_columns = ['LEVEL', 'LEVELPLAYED', 'TOTALLEVELPLAYED', 'TOTALLEVELSPLAYED']
    level_col = next((col for col in df.columns if col in level_columns), None)

    if file_type == 'start':
        user_col = next((col for col in df.columns if 'USER' in col), None)
        cols_to_keep = [level_col, user_col] if level_col and user_col else None
    else:
        user_col = next((col for col in df.columns if 'USER' in col), None)
        additional_columns = ['PLAY_TIME_AVG', 'HINT_USED_SUM', 'RETRY_COUNT_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']
        available_additional_cols = [col for col in additional_columns if col in df.columns]
        cols_to_keep = [level_col, user_col] + available_additional_cols if level_col and user_col else None

    if not cols_to_keep:
        st.error(f"❌ Required columns not found in {file_type} file.")
        return None

    df = df[cols_to_keep]
    df['LEVEL_CLEAN'] = df[level_col].apply(clean_level)
    df.dropna(subset=['LEVEL_CLEAN'], inplace=True)
    df['LEVEL_CLEAN'] = df['LEVEL_CLEAN'].astype(int)
    df.sort_values('LEVEL_CLEAN', inplace=True)

    if file_type == 'start':
        df.rename(columns={user_col: 'Start Users'}, inplace=True)
    else:
        df.rename(columns={user_col: 'Complete Users'}, inplace=True)
        if available_additional_cols:
            df[available_additional_cols] = df[available_additional_cols].round(2)

    return df

def generate_excel(df_export, retention_fig, drop_fig, drop_comb_fig, version, date_selected):
    """Generate Excel file with formatted sheets"""
    output = BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Remove duplicate levels
        df_export = df_export.drop_duplicates(subset='Level', keep='first').reset_index(drop=True)

        # Write main dataframe to Excel
        df_export.to_excel(writer, sheet_name='Summary', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Summary']

        # Header format
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#D9E1F2',
            'border': 1
        })

        # Cell format
        cell_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter'
        })

        # Highlight format for drop rates ≥ 3
        highlight_format = workbook.add_format({
            'font_color': 'red',
            'bg_color': 'yellow',
            'align': 'center',
            'valign': 'vcenter'
        })

        # Apply formats
        for col_num, value in enumerate(df_export.columns):
            worksheet.write(0, col_num, value, header_format)

        # Apply conditional formatting
        for row_num in range(1, len(df_export) + 1):
            for col_num in range(len(df_export.columns)):
                value = df_export.iloc[row_num - 1, col_num]
                col_name = df_export.columns[col_num]

                if pd.isna(value):
                    value = ""

                if isinstance(value, (np.generic, np.bool_)):
                    value = value.item()

                try:
                    if col_name in ['Game Play Drop', 'Popup Drop', 'Total Level Drop'] and isinstance(value, (int, float)) and value >= 3:
                        worksheet.write(row_num, col_num, value, highlight_format)
                    else:
                        worksheet.write(row_num, col_num, value, cell_format)
                except Exception as e:
                    st.warning(f"⚠️ Could not write value at row {row_num} col {col_num}: {e}")

        # Freeze top row
        worksheet.freeze_panes(1, 0)

        # Set column widths
        for i, col in enumerate(df_export.columns):
            column_len = max(df_export[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)

        # Insert charts
        if retention_fig:
            retention_img = BytesIO()
            retention_fig.savefig(retention_img, format='png', dpi=300, bbox_inches='tight')
            retention_img.seek(0)
            worksheet.insert_image('M2', 'retention_chart.png', {'image_data': retention_img})

        if drop_fig:
            drop_img = BytesIO()
            drop_fig.savefig(drop_img, format='png', dpi=300, bbox_inches='tight')
            drop_img.seek(0)
            worksheet.insert_image('M37', 'drop_chart.png', {'image_data': drop_img})

        if drop_comb_fig:
            drop_comb_img = BytesIO()
            drop_comb_fig.savefig(drop_comb_img, format='png', dpi=300, bbox_inches='tight')
            drop_comb_img.seek(0)
            worksheet.insert_image('M67', 'drop_comb_chart.png', {'image_data': drop_comb_img})

        # Create additional sheets with same formatting
        for sheet_name in ['Raw_Data', 'Analysis']:
            workbook.add_worksheet(sheet_name)
            worksheet = writer.sheets[sheet_name]

            # Apply same formatting to new sheets
            for col_num in range(len(df_export.columns)):
                worksheet.write(0, col_num, df_export.columns[col_num], header_format)
                worksheet.set_column(col_num, col_num, 20)

            worksheet.freeze_panes(1, 0)

    output.seek(0)
    return output

def format_final_excel(file_path):
    """Apply final formatting to Excel file"""
    wb = load_workbook(file_path)
    sheets = wb.sheetnames

    # 1. Create Locate Sheet
    if 'LOCATE SHEET' not in sheets:
        locate_sheet = wb.create_sheet(title="LOCATE SHEET", index=0)
    else:
        locate_sheet = wb['LOCATE SHEET']

    locate_sheet['A1'] = "Available Sheets"
    locate_sheet['A1'].font = Font(bold=True)
    locate_sheet['A1'].alignment = Alignment(horizontal='center')
    locate_sheet.column_dimensions['A'].width = 40

    # Add clickable links to all sheets except LOCATE SHEET
    for idx, sheet_name in enumerate([s for s in sheets if s != 'LOCATE SHEET'], start=2):
        locate_sheet[f'A{idx}'] = sheet_name
        locate_sheet[f'A{idx}'].font = Font(color='0000FF', underline='single')
        locate_sheet[f'A{idx}'].hyperlink = f"#{sheet_name}!A1"

    locate_sheet.freeze_panes = 'A2'

    # 2. Apply Formatting to All Sheets
    for sheet_name in sheets:
        if sheet_name == 'LOCATE SHEET':
            continue

        ws = wb[sheet_name]

        # Freeze Row-1
        ws.freeze_panes = 'A2'

        # Add Back to Locate Sheet hyperlink in A1
        ws['A1'] = 'Back to LOCATE SHEET'
        ws['A1'].font = Font(color='0000FF', underline='single')
        ws['A1'].hyperlink = "#'LOCATE SHEET'!A1"

        # Auto-fit Columns
        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 2

        # Header Styling (Row 1)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

        # Center Align all rows
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')

        # Highlight Drop Columns >= 3%
        drop_cols = ["Game Play Drop", "Popup Drop", "Total Level Drop"]
        for drop_col in drop_cols:
            if drop_col in [cell.value for cell in ws[1]]:
                col_idx = [cell.value for cell in ws[1]].index(drop_col) + 1
                for cell in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2):
                    for c in cell:
                        if isinstance(c.value, (int, float)) and c.value >= 3:
                            c.font = Font(color="9C0006")
                            c.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    wb.save(file_path)
    return file_path

def process_folder(uploaded_folder, folder_type):
    """Process uploaded folder containing CSV files and return combined DataFrame"""
    temp_dir = tempfile.mkdtemp()
    all_dfs = []

    try:
        # Save the uploaded file to temp directory
        upload_path = os.path.join(temp_dir, uploaded_folder.name)
        with open(upload_path, "wb") as f:
            f.write(uploaded_folder.getbuffer())

        # Handle ZIP files
        if zipfile.is_zipfile(upload_path):
            with zipfile.ZipFile(upload_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            # Remove the zip file after extraction
            os.remove(upload_path)

        # Walk through the temp directory to find CSV files
        for root, _, files in os.walk(temp_dir):
            for file in files:
                file_lower = file.lower()

                # Only process CSV files
                if file_lower.endswith('.csv'):
                    file_path = os.path.join(root, file)
                    try:
                        # Read CSV file with error handling for different encodings
                        try:
                            df = pd.read_csv(file_path)
                        except UnicodeDecodeError:
                            df = pd.read_csv(file_path, encoding='latin1')

                        # Process the file
                        processed_df = process_file(df, folder_type)
                        if processed_df is not None:
                            all_dfs.append(processed_df)
                    except Exception as e:
                        st.warning(f"Could not process file {file}: {str(e)}")
                        continue

        # Clean up temp directory
        for root, dirs, files in os.walk(temp_dir, topdown=False):
            for name in files:
                os.remove(os.path.join(root, name))
            for name in dirs:
                os.rmdir(os.path.join(root, name))
        os.rmdir(temp_dir)

        if not all_dfs:
            st.error(f"No valid CSV files found in the {folder_type} folder.")
            return None

        # Combine all DataFrames and group by level
        combined_df = pd.concat(all_dfs, ignore_index=True)
        return combined_df.groupby('LEVEL_CLEAN').sum().reset_index()

    except Exception as e:
        st.error(f"Error processing {folder_type} folder: {str(e)}")
        return None

def get_max_min_start_user_info(df_start):
    """Get max/min level and user count from start data"""
    if df_start is None or 'Start Users' not in df_start.columns or 'LEVEL_CLEAN' not in df_start.columns:
        return None

    max_level = df_start['LEVEL_CLEAN'].max()
    max_level_user_count = df_start.loc[df_start['LEVEL_CLEAN'] == max_level, 'Start Users'].values[0]

    min_level = df_start['LEVEL_CLEAN'].min()
    min_level_user_count = df_start.loc[df_start['LEVEL_CLEAN'] == min_level, 'Start Users'].values[0]

    return {
        'max_level': max_level,
        'max_level_user_count': max_level_user_count,
        'min_level': min_level,
        'min_level_user_count': min_level_user_count
    }

def prepare_locate_sheet_data(file_info_list):
    """Prepare data for the locate sheet"""
    locate_data = []

    for idx, file_info in enumerate(file_info_list, start=1):
        locate_data.append({
            'Sr No': idx,
            'File Name': file_info['file_name'],
            'Max Start Level': file_info['max_level'],
            'User Count (Max Level)': file_info['max_level_user_count'],
            'Min Start Level': file_info['min_level'],
            'User Count (Min Level)': file_info['min_level_user_count'],
            'Sheet Name': file_info['sheet_name']
        })

    return pd.DataFrame(locate_data)

# -------------------- MAIN FUNCTION -------------------- #
def main():
    # -------------- FILE UPLOAD SECTION ------------------ #
    st.sidebar.header("Upload Data")
    start_folder = st.sidebar.file_uploader(
        "📂 Upload LEVEL_START Folder (ZIP or folder with CSVs)",
        type=["zip", "csv", "tar", "gz"]
    )
    complete_folder = st.sidebar.file_uploader(
        "📂 Upload LEVEL_COMPLETE Folder (ZIP or folder with CSVs)",
        type=["zip", "csv", "tar", "gz"]
    )

    version = st.sidebar.text_input("📌 Game Version", value="1.0.0")
    date_selected = st.sidebar.date_input("📅 Select Date", value=datetime.date.today())

    if start_folder and complete_folder:
        with st.spinner("Processing files..."):
            # Process folders
            df_start = process_folder(start_folder, 'start')
            df_complete = process_folder(complete_folder, 'complete')

            if df_start is None or df_complete is None:
                st.error("Failed to process one or both folders. Please check that they contain valid CSV files.")
                return

            # Get max/min info for locate sheet
            start_info = get_max_min_start_user_info(df_start)
            complete_info = get_max_min_start_user_info(df_complete)

            if not start_info or not complete_info:
                st.error("Could not extract level information from the data.")
                return

            # ------------ MERGE AND CALCULATE METRICS ------------- #
            df = pd.merge(df_start, df_complete, on='LEVEL_CLEAN', how='outer').sort_values('LEVEL_CLEAN')

            # Calculate metrics
            df['Game Play Drop'] = ((df['Start Users'] - df['Complete Users']) / df['Start Users']) * 100
            df['Popup Drop'] = ((df['Complete Users'] - df['Start Users'].shift(-1)) / df['Complete Users']) * 100
            df['Total Level Drop'] = ((df['Start Users'] - df['Start Users'].shift(-1)) / df['Start Users']) * 100
            max_start_users = df['Start Users'].max()
            df['Retention %'] = (df['Start Users'] / max_start_users) * 100

            metric_cols = ['Game Play Drop', 'Popup Drop', 'Total Level Drop', 'Retention %']
            df[metric_cols] = df[metric_cols].round(2)

            # Get available additional columns from complete data
            additional_cols = [col for col in df.columns if col in ['PLAY_TIME_AVG', 'HINT_USED_SUM',
                                                                  'RETRY_COUNT_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']]

            # ------------ CHARTS ------------ #
            df_100 = df[df['LEVEL_CLEAN'] <= 100]

            # Custom x tick labels
            xtick_labels = []
            for val in np.arange(1, 101, 1):
                if val % 5 == 0:
                    xtick_labels.append(f"$\\bf{{{val}}}$")  # Bold using LaTeX
                else:
                    xtick_labels.append(str(val))

            # ------------ RETENTION CHART ------------ #
            st.subheader("📈 Retention Chart (Levels 1-100)")
            retention_fig, ax = plt.subplots(figsize=(15, 7))

            ax.plot(df_100['LEVEL_CLEAN'], df_100['Retention %'],
                    linestyle='-', color='#F57C00', linewidth=2, label='RETENTION')

            ax.set_xlim(1, 100)
            ax.set_ylim(0, 110)
            ax.set_xticks(np.arange(1, 101, 1))
            ax.set_yticks(np.arange(0, 110, 5))
            ax.set_xlabel("Level", labelpad=15)
            ax.set_ylabel("% Of Users", labelpad=15)
            ax.set_title(f"Retention Chart (Levels 1-100) | Version {version} | Date: {date_selected.strftime('%d-%m-%Y')}",
                         fontsize=12, fontweight='bold')

            ax.set_xticklabels(xtick_labels, fontsize=6)
            ax.tick_params(axis='x', labelsize=6)
            ax.grid(True, linestyle='--', linewidth=0.5)

            # Annotate data points below x-axis
            for x, y in zip(df_100['LEVEL_CLEAN'], df_100['Retention %']):
                ax.text(x, -5, f"{int(y)}", ha='center', va='top', fontsize=7)

            ax.legend(loc='lower left', fontsize=8)
            plt.tight_layout(rect=[0, 0.03, 1, 0.97])
            st.pyplot(retention_fig)

            # ------------ TOTAL DROP CHART ------------ #
            st.subheader("📉 Total Drop Chart (Levels 1-100)")
            drop_fig, ax2 = plt.subplots(figsize=(15, 6))
            bars = ax2.bar(df_100['LEVEL_CLEAN'], df_100['Total Level Drop'], color='#EF5350', label='DROP RATE')

            ax2.set_xlim(1, 100)
            ax2.set_ylim(0, max(df_100['Total Level Drop'].max(), 10) + 10)
            ax2.set_xticks(np.arange(1, 101, 1))
            ax2.set_yticks(np.arange(0, max(df_100['Total Level Drop'].max(), 10) + 11, 5))
            ax2.set_xlabel("Level")
            ax2.set_ylabel("% Of Users Drop")
            ax2.set_title(f"Total Level Drop Chart | Version {version} | Date: {date_selected.strftime('%d-%m-%Y')}",
                          fontsize=12, fontweight='bold')

            ax2.set_xticklabels(xtick_labels, fontsize=6)
            ax2.tick_params(axis='x', labelsize=6)
            ax2.grid(True, linestyle='--', linewidth=0.5)

            # Annotate data points below x-axis
            for bar in bars:
                x = bar.get_x() + bar.get_width() / 2
                y = bar.get_height()
                ax2.text(x, -2, f"{y:.0f}", ha='center', va='top', fontsize=7)

            ax2.legend(loc='upper right', fontsize=8)
            plt.tight_layout()
            st.pyplot(drop_fig)

            # ------------ COMBO DROP CHART ------------ #
            st.subheader("📉 Combo Drop Chart (Levels 1-100)")
            drop_comb_fig, ax3 = plt.subplots(figsize=(15, 6))

            # Plot both drop types
            width = 0.4
            x = df_100['LEVEL_CLEAN']
            ax3.bar(x + width/2, df_100['Game Play Drop'], width, color='#66BB6A', label='Game Play Drop')
            ax3.bar(x - width/2, df_100['Popup Drop'], width, color='#42A5F5', label='Popup Drop')

            ax3.set_xlim(1, 100)
            max_drop = max(df_100['Game Play Drop'].max(), df_100['Popup Drop'].max())
            ax3.set_ylim(0, max(max_drop, 10) + 10)
            ax3.set_xticks(np.arange(1, 101, 1))
            ax3.set_yticks(np.arange(0, max(max_drop, 10) + 11, 5))
            ax3.set_xlabel("Level")
            ax3.set_ylabel("% Of Users Dropped")
            ax3.set_title(f"Game Play & Popup Drop Chart | Version {version} | Date: {date_selected.strftime('%d-%m-%Y')}",
                          fontsize=12, fontweight='bold')

            ax3.set_xticklabels(xtick_labels, fontsize=6)
            ax3.tick_params(axis='x', labelsize=6)
            ax3.grid(True, linestyle='--', linewidth=0.5)
            ax3.legend(loc='upper right', fontsize=8)
            plt.tight_layout()
            st.pyplot(drop_comb_fig)

            # ------------ DOWNLOAD SECTION ------------ #
            st.subheader("⬇️ Download Excel Report")

            # Prepare export dataframe
            export_columns = ['LEVEL_CLEAN', 'Start Users', 'Complete Users',
                             'Game Play Drop', 'Popup Drop', 'Total Level Drop',
                             'Retention %'] + additional_cols

            df_export = df[export_columns].rename(columns={'LEVEL_CLEAN': 'Level'})

            # Display dataframe
            st.dataframe(df_export)

            # Generate and download Excel
            excel_data = generate_excel(df_export, retention_fig, drop_fig, drop_comb_fig)

            # Create download button
            st.download_button(
                label="📥 Download Excel Report",
                data=excel_data,
                file_name=f"GAME_PROGRESSION_Report_{version}_{date_selected.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
