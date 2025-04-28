import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import datetime
import re
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter, quote_sheetname
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.comments import Comment

st.set_page_config(page_title="All GAME PROGRESSION", layout="wide")
st.title("📊 GAME PROGRESSION Dashboard")

# -------------------- FUNCTION TO EXPORT EXCEL -------------------- #
def generate_excel(merged_data, locate_sheet_data, version, date_selected):
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Create Locate Sheet first
        locate_df = pd.DataFrame(locate_sheet_data, columns=["Index", "Sheet Name", "Drop Count", "Link to Sheet"])
        locate_df.to_excel(writer, sheet_name="Locate Sheet", index=False)

        # Write each game's data to separate sheets
        for sheet_name, data in merged_data.items():
            df_export = data['df_export']
            df_export.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)

    # Open the workbook to apply formatting
    book = load_workbook(output)

    # Format Locate Sheet
    locate_sheet = book["Locate Sheet"]
    locate_sheet.freeze_panes = "A2"

    # Header formatting
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center")

    for cell in locate_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    # Format each game sheet
    for sheet_name in book.sheetnames:
        if sheet_name != "Locate Sheet":
            sheet = book[sheet_name]

            # Format headers
            for cell in sheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment

            # Format data cells
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                    # Highlight problematic drops
                    if cell.column_letter in ['D', 'E', 'F'] and isinstance(cell.value, (int, float)) and cell.value >= 3:
                        cell.font = Font(color="FF0000", bold=True)
                        cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

            # Add back to Locate Sheet link
            sheet['A1'].value = f'=HYPERLINK("#Locate Sheet!A1", "Back to Locate Sheet")'
            sheet['A1'].font = Font(color="0000FF", underline="single")
            sheet['A1'].alignment = Alignment(horizontal="center")

            # Adjust column widths
            for column in sheet.columns:
                max_length = max([len(str(cell.value)) if cell.value else 0 for cell in column])
                sheet.column_dimensions[get_column_letter(column[0].column)].width = max_length + 2

    # Save the formatted workbook
    output = BytesIO()
    book.save(output)
    output.seek(0)
    return output

def clean_level(x):
    try:
        return int(re.search(r"(\d+)", str(x)).group(1))
    except:
        return None

def process_file(file_path, version, date_selected):
    try:
        # Load file
        df = pd.read_csv(file_path) if file_path.endswith(".csv") else pd.read_excel(file_path)

        # Clean column names
        df.columns = df.columns.str.strip().str.upper()

        # Identify level and user columns
        level_columns = ['LEVEL', 'LEVELPLAYED', 'TOTALLEVELPLAYED', 'TOTALLEVELSPLAYED']
        level_col = next((col for col in df.columns if col in level_columns), None)
        user_col = next((col for col in df.columns if 'USER' in col), None)

        if not level_col or not user_col:
            st.warning(f"Skipping {file_path} - required columns not found")
            return None, None

        # Clean data
        df = df[[level_col, user_col]]
        df['LEVEL_CLEAN'] = df[level_col].apply(clean_level)
        df.dropna(inplace=True)
        df['LEVEL_CLEAN'] = df['LEVEL_CLEAN'].astype(int)
        df.sort_values('LEVEL_CLEAN', inplace=True)
        df.rename(columns={user_col: 'Users'}, inplace=True)

        # Calculate metrics
        max_users = df['Users'].max()
        df['Retention %'] = (df['Users'] / max_users) * 100
        df['Game Play Drop'] = ((df['Users'] - df['Users'].shift(-1)) / df['Users']) * 100
        df[['Retention %', 'Game Play Drop']] = df[['Retention %', 'Game Play Drop']].round(2)

        # Prepare export dataframe
        df_export = df[['LEVEL_CLEAN', 'Users', 'Retention %', 'Game Play Drop']]
        df_export = df_export.rename(columns={'LEVEL_CLEAN': 'Level'})

        # Create charts
        df_100 = df[df['LEVEL_CLEAN'] <= 100]

        # Retention Chart
        retention_fig, ax = plt.subplots(figsize=(15, 7))
        ax.plot(df_100['LEVEL_CLEAN'], df_100['Retention %'],
                linestyle='-', color='#F57C00', linewidth=2, label='RETENTION')
        ax.set_xlim(1, 100)
        ax.set_ylim(0, 110)
        ax.set_xticks(np.arange(1, 101, 1))
        ax.set_yticks(np.arange(0, 110, 5))
        ax.set_xlabel("Level", labelpad=15)
        ax.set_ylabel("% Of Users", labelpad=15)
        ax.set_title(f"Retention Chart | Version {version} | Date: {date_selected.strftime('%d-%m-%Y')}",
                    fontsize=12, fontweight='bold')
        ax.grid(True, linestyle='--', linewidth=0.5)
        ax.legend(loc='lower left', fontsize=8)
        plt.tight_layout(rect=[0, 0.03, 1, 0.97])

        # Drop Chart
        drop_fig, ax2 = plt.subplots(figsize=(15, 6))
        bars = ax2.bar(df_100['LEVEL_CLEAN'], df_100['Game Play Drop'], color='#EF5350', label='DROP RATE')
        ax2.set_xlim(1, 100)
        ax2.set_ylim(0, max(df_100['Game Play Drop'].max(), 10) + 10)
        ax2.set_xticks(np.arange(1, 101, 1))
        ax2.set_yticks(np.arange(0, max(df_100['Game Play Drop'].max(), 10) + 11, 5))
        ax2.set_xlabel("Level")
        ax2.set_ylabel("% Of Users Drop")
        ax2.set_title(f"Drop Rate Chart | Version {version} | Date: {date_selected.strftime('%d-%m-%Y')}",
                      fontsize=12, fontweight='bold')
        ax2.grid(True, linestyle='--', linewidth=0.5)
        ax2.legend(loc='upper right', fontsize=8)
        plt.tight_layout()

        # Get file name without extension
        sheet_name = os.path.basename(file_path).split(".")[0]

        return {
            'df_export': df_export,
            'retention_fig': retention_fig,
            'drop_fig': drop_fig,
            'sheet_name': sheet_name
        }

    except Exception as e:
        st.error(f"Error processing {file_path}: {str(e)}")
        return None

def main():
    # -------------- FILE UPLOAD SECTION ------------------ #
    st.sidebar.header("Upload Settings")
    version = st.sidebar.text_input("📌 Game Version", value="1.0.0")
    date_selected = st.sidebar.date_input("📅 Select Date", value=datetime.date.today())

    uploaded_files = st.sidebar.file_uploader(
        "📂 Upload CSV Files",
        type=["csv"],
        accept_multiple_files=True
    )

    if uploaded_files:
        merged_data = {}
        locate_sheet_data = []

        for uploaded_file in uploaded_files:
            # Save the uploaded file temporarily
            with open(uploaded_file.name, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Process the file
            result = process_file(uploaded_file.name, version, date_selected)

            if result:
                sheet_name = result['sheet_name']
                merged_data[sheet_name] = {
                    'df_export': result['df_export'],
                    'retention_fig': result['retention_fig'],
                    'drop_fig': result['drop_fig']
                }

                # Add to locate sheet data
                drop_count = len(result['df_export'][result['df_export']['Game Play Drop'] >= 3])
                locate_sheet_data.append([
                    len(locate_sheet_data) + 1,
                    sheet_name,
                    drop_count,
                    f'=HYPERLINK("#{sheet_name}!A1", "Click to view")'
                ])

                # Display the game's data
                st.subheader(f"📊 {sheet_name} Analysis")
                col1, col2 = st.columns(2)

                with col1:
                    st.pyplot(result['retention_fig'])

                with col2:
                    st.pyplot(result['drop_fig'])

                st.dataframe(result['df_export'])

                # Remove the temporary file
                os.remove(uploaded_file.name)

        # Generate and download Excel
        if merged_data:
            st.sidebar.subheader("⬇️ Download Consolidated Report")
            excel_data = generate_excel(merged_data, locate_sheet_data, version, date_selected)

            st.sidebar.download_button(
                label="📥 Download Excel Report",
                data=excel_data,
                file_name=f"ALL_GAMES_PROGRESSION_Report_{version}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
