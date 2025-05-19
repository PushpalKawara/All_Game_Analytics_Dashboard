import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import xlsxwriter

st.set_page_config(page_title="📊 Game Level Data Merger", layout="wide")
st.title("📊 Streamlit Tool for Merging and Analyzing Game Level Data")

# ---- FILE UPLOADS ----
st.sidebar.header("Upload CSV Files")
level_start_file = st.sidebar.file_uploader("Upload LEVEL_START.csv", type="csv")
level_complete_file = st.sidebar.file_uploader("Upload LEVEL_COMPLETE.csv", type="csv")

def clean_level(level_val):
    return ''.join(filter(str.isdigit, str(level_val)))

def format_sheet(workbook, worksheet, df):
    header_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#DCE6F1', 'border': 1
    })
    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    
    worksheet.freeze_panes(1, 0)
    for i, col in enumerate(df.columns):
        worksheet.write(0, i, col, header_format)
        worksheet.set_column(i, i, 18, center_format)

def apply_conditional_formatting(worksheet, df, workbook):
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
            worksheet.conditional_format(1, col_idx, len(df), col_idx, {
                'type': 'cell',
                'criteria': 'between',
                'minimum': 7,
                'maximum': 9.99,
                'format': workbook.add_format({'bg_color': '#CD5C5C', 'font_color': '#FFFFFF'})
            })
            worksheet.conditional_format(1, col_idx, len(df), col_idx, {
                'type': 'cell',
                'criteria': 'between',
                'minimum': 3,
                'maximum': 6.99,
                'format': workbook.add_format({'bg_color': '#FFCCCC', 'font_color': '#FFFFFF'})
            })

# ---- PROCESS FILES ----
if level_start_file and level_complete_file:
    df_start = pd.read_csv(level_start_file)
    df_complete = pd.read_csv(level_complete_file)

    df_start['LEVEL'] = df_start['LEVEL'].apply(clean_level)
    df_complete['LEVEL'] = df_complete['LEVEL'].apply(clean_level)

    df_start = df_start.rename(columns={'USERS': 'Start Users'})
    df_complete = df_complete.rename(columns={'USERS': 'Complete Users'})

    df_start = df_start.sort_values(by='LEVEL')
    df_complete = df_complete.sort_values(by='LEVEL')

    merged_df = pd.merge(df_start, df_complete, on=['GAME_ID', 'DIFFICULTY', 'LEVEL'], how='outer')

    merged_df['Game Play Drop'] = merged_df['Start Users'] - merged_df['Complete Users']
    merged_df['Popup Drop'] = merged_df['Start Users'] * 0.03
    merged_df['Total Level Drop'] = merged_df['Game Play Drop'] + merged_df['Popup Drop']
    merged_df['Retention %'] = (merged_df['Complete Users'] / merged_df['Start Users']) * 100

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        merged_df.to_excel(writer, index=False, sheet_name='MergedData')
        workbook = writer.book
        worksheet = writer.sheets['MergedData']

        format_sheet(workbook, worksheet, merged_df)
        apply_conditional_formatting(worksheet, merged_df, workbook)

        # Later steps: Insert charts, add MAIN_TAB and backlinks
    st.success("✅ Merging complete. You can download the Excel below:")

    st.download_button(
        label="📥 Download Merged Excel",
        data=output.getvalue(),
        file_name="Consolidated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
