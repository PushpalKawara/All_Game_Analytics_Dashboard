import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import xlsxwriter

st.set_page_config(page_title="Game Level Data Merger", layout="wide")
st.title("📊 Streamlit Tool for Merging and Analyzing Game Level Data")

# ---- File Upload ----
st.sidebar.header("Upload CSV Files")
level_start_file = st.sidebar.file_uploader("Upload LEVEL_START.csv", type="csv")
level_complete_file = st.sidebar.file_uploader("Upload LEVEL_COMPLETE.csv", type="csv")

# ---- Helper Functions ----
def clean_level(level_val):
    return ''.join(filter(str.isdigit, str(level_val)))

def format_sheet(workbook, worksheet, df):
    header_format = workbook.add_format({'bold': True})
    worksheet.freeze_panes(1, 0)
    for i, col in enumerate(df.columns):
        worksheet.write(0, i, col, header_format)
        worksheet.set_column(i, i, 18)

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

# ---- Main Logic ----
if level_start_file and level_complete_file:
    df_start = pd.read_csv(level_start_file)
    df_complete = pd.read_csv(level_complete_file)

    df_start['LEVEL'] = df_start['LEVEL'].apply(clean_level)
    df_complete['LEVEL'] = df_complete['LEVEL'].apply(clean_level)

    df_start = df_start.rename(columns={'USERS': 'Start Users'})
    df_complete = df_complete.rename(columns={'USERS': 'Complete Users'})

    merge_cols = ['GAME_ID', 'DIFFICULTY', 'LEVEL']
    df_merge = pd.merge(df_start, df_complete, on=merge_cols, how='outer').sort_values(by='LEVEL')

    df_merge.fillna(0, inplace=True)

    # Calculations
    df_merge['Game Play Drop'] = ((df_merge['Start Users'] - df_merge['Complete Users']) / df_merge['Start Users'].replace(0, np.nan)) * 100
    df_merge['Popup Drop'] = ((df_merge['Complete Users'] - df_merge['Start Users'].shift(-1)) / df_merge['Complete Users'].replace(0, np.nan)) * 100
    df_merge['Total Level Drop'] = ((df_merge['Start Users'] - df_merge['Start Users'].shift(-1)) / df_merge['Start Users'].replace(0, np.nan)) * 100
    df_merge['Retention %'] = (df_merge['Start Users'] / df_merge['Start Users'].max()) * 100

    # Download Button for Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sheet_name = f"{df_merge.iloc[0]['GAME_ID']}_{df_merge.iloc[0]['DIFFICULTY']}"
        df_merge.to_excel(writer, index=False, sheet_name=sheet_name, startrow=1)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        format_sheet(workbook, worksheet, df_merge)
        apply_conditional_formatting(worksheet, df_merge, workbook)

        # Backlink
        worksheet.write('A1', f'=HYPERLINK("#\'MAIN_TAB\'!A1", "Back to Locate Sheet")')

    st.download_button(
        label="📥 Download Processed Excel",
        data=output.getvalue(),
        file_name="Processed_Game_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Display in app
    st.subheader("Preview of Merged and Processed Data")
    st.dataframe(df_merge.head(50))

else:
    st.info("Please upload both LEVEL_START.csv and LEVEL_COMPLETE.csv files.")
