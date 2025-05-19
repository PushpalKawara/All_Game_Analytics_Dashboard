import streamlit as st
import pandas as pd
import os
import io
from io import BytesIO
import numpy as np
import xlsxwriter

st.set_page_config(page_title="Game Level Data Merger", layout="wide")
st.title("📊 Streamlit Tool for Merging and Analyzing Game Level Data")

# ---- FILE UPLOADS ----
st.sidebar.header("Upload CSV Files")
level_start_file = st.sidebar.file_uploader("Upload LEVEL_START.csv", type="csv")
level_complete_file = st.sidebar.file_uploader("Upload LEVEL_COMPLETE.csv", type="csv")

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

    df_start = df_start.sort_values(by='LEVEL')
    df_complete = df_complete.sort_values(by='LEVEL')

    df_start = df_start.rename(columns={'USERS': 'Start Users'})
    df_complete = df_complete.rename(columns={'USERS': 'Complete Users'})

    merge_cols = ['GAME_ID', 'DIFFICULTY', 'LEVEL']
    df_merge = pd.merge(df_start, df_complete, on=merge_cols, how='outer')

    df_merge = df_merge[['LEVEL','GAME_ID', 'DIFFICULTY', 'Start Users', 'Complete Users', 'PLAY_TIME_AVG',
                         'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPTS_SUM']]

    df_merge['Start Users'] = df_merge['Start Users'].fillna(0)
    df_merge['Complete Users'] = df_merge['Complete Users'].fillna(0)
    df_merge['Game Play Drop'] = df_merge['Start Users'] - df_merge['Complete Users']
    df_merge['Popup Drop'] = df_merge['Start Users'] * 0.03
    df_merge['Total Level Drop'] = df_merge['Game Play Drop'] + df_merge['Popup Drop']
    df_merge['Retention %'] = np.where(df_merge['Start Users'] == 0, 0,
                                       (df_merge['Complete Users'] / df_merge['Start Users']) * 100)

    # Prepare summary sheet data
    sheet_name = df_start.iloc[0]['GAME_ID'] + "_" + df_start.iloc[0]['DIFFICULTY']
    summary_data = {
        "Sheet Name": sheet_name,
        "Game Play Drop Count": (df_merge['Game Play Drop'] >= (df_merge['Start Users'] * 0.03)).sum(),
        "Popup Drop Count": (df_merge['Popup Drop'] >= (df_merge['Start Users'] * 0.03)).sum(),
        "Total Level Drop Count": (df_merge['Total Level Drop'] >= (df_merge['Start Users'] * 0.03)).sum(),
        "LEVEL_Start": df_merge['LEVEL'].min(),
        "USERS_starts": df_merge['Start Users'].max(),
        "LEVEL_End": df_merge['LEVEL'].max(),
        "USERS_END": df_merge['Start Users'].iloc[-1],
        "Link": f'=HYPERLINK("#{sheet_name}!A1", "{sheet_name}")'
    }

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Write merged sheet
        df_merge.to_excel(writer, index=False, sheet_name=sheet_name, startrow=1)
        worksheet = writer.sheets[sheet_name]
        format_sheet(workbook, worksheet, df_merge)

        # Backlink
        worksheet.write('A1', f'=HYPERLINK("#\'MAIN_TAB\'!A1", "Back to Locate Sheet")')

        # Charts
        chart1 = workbook.add_chart({'type': 'line'})
        chart1.add_series({'values': f'={sheet_name}!$J$2:$J${len(df_merge)+1}', 'name': 'Retention %'})
        chart1.set_title({'name': 'Retention Line Chart'})
        worksheet.insert_chart('N2', chart1)

        chart2 = workbook.add_chart({'type': 'column'})
        chart2.add_series({'values': f'={sheet_name}!$I$2:$I${len(df_merge)+1}', 'name': 'Total Level Drop'})
        chart2.set_title({'name': 'Total Level Drop'})
        worksheet.insert_chart('N39', chart2)

        chart3 = workbook.add_chart({'type': 'column'})
        chart3.add_series({'values': f'={sheet_name}!$G$2:$G${len(df_merge)+1}', 'name': 'Game Play Drop'})
        chart3.add_series({'values': f'={sheet_name}!$H$2:$H${len(df_merge)+1}', 'name': 'Popup Drop'})
        chart3.set_title({'name': 'Game + Popup Drop'})
        worksheet.insert_chart('N70', chart3)

        # Conditional Formatting
        apply_conditional_formatting(worksheet, df_merge, workbook)

        # MAIN_TAB
        summary_df = pd.DataFrame([summary_data])
        summary_df.to_excel(writer, sheet_name='MAIN_TAB', index=False)
        format_sheet(workbook, writer.sheets['MAIN_TAB'], summary_df)

    st.success("✅ Merging completed successfully!")
    st.download_button(label="📥 Download Consolidated Excel",
                       data=output.getvalue(),
                       file_name="Consolidated.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
