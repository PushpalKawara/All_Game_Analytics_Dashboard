import streamlit as st
import pandas as pd
import numpy as np
import io
import xlsxwriter

st.set_page_config(page_title="Game Progression Report Generator")
st.title("🎮 BrainGames Progression Report Generator")
st.write("Upload your **level_started** and **level_completed** files to generate a full game analytics report with sheet links, drop rates, and retention.")

# Upload
start_file = st.file_uploader("Upload level_started.csv", type="csv")
complete_file = st.file_uploader("Upload level_completed.csv", type="csv")

if start_file and complete_file:
    df_start = pd.read_csv(start_file)
    df_complete = pd.read_csv(complete_file)

    # Expected columns
    expected_cols = ['GAME_ID', 'GAME_MODE', 'Level', 'user_id', 'timestamp']
    if not all(col in df_start.columns for col in expected_cols) or not all(col in df_complete.columns for col in expected_cols):
        st.error("❌ One of the files is missing required columns.")
        st.stop()

    # Process input
    start_grp = df_start.groupby(['GAME_ID', 'GAME_MODE', 'Level'])['user_id'].nunique().reset_index()
    start_grp.rename(columns={'user_id': 'Start Users'}, inplace=True)

    complete_grp = df_complete.groupby(['GAME_ID', 'GAME_MODE', 'Level'])['user_id'].nunique().reset_index()
    complete_grp.rename(columns={'user_id': 'Complete Users'}, inplace=True)

    merged = pd.merge(start_grp, complete_grp, on=['GAME_ID', 'GAME_MODE', 'Level'], how='left')
    merged['Complete Users'] = merged['Complete Users'].fillna(0).astype(int)
    merged.sort_values(by=['GAME_ID', 'GAME_MODE', 'Level'], inplace=True)

    # Drop Rate
    merged['Total Level Drop'] = (merged['Start Users'] - merged['Complete Users']) / merged['Start Users']
    merged['Total Level Drop'] = merged['Total Level Drop'].round(4)

    # Retention %
    merged['Retention %'] = 0.0
    for (game, mode), group in merged.groupby(['GAME_ID', 'GAME_MODE']):
        idx = group.index
        start_values = group['Start Users'].values
        retention = [start_values[i + 1] / start_values[i] if i + 1 < len(start_values) and start_values[i] > 0 else 0 for i in range(len(start_values))]
        retention.append(0)
        merged.loc[idx, 'Retention %'] = retention

    merged['Retention %'] = merged['Retention %'].round(4)

    # Prepare Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Create main sheet with hyperlinks
        summary_sheet = 'LIS'
        summary_df = merged.groupby(['GAME_ID', 'GAME_MODE']).size().reset_index().drop(columns=0)
        summary_df['Link'] = ""

        for idx, row in summary_df.iterrows():
            sheet_name = f"{row['GAME_ID']}_{row['GAME_MODE']}".replace(" ", "_")[:31]
            summary_df.at[idx, 'Link'] = f"=HYPERLINK(\"#{sheet_name}!A1\", \"View\")"

        summary_df.to_excel(writer, sheet_name=summary_sheet, index=False)
        summary_ws = writer.sheets[summary_sheet]

        # Format main sheet
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'})
        for col_num, value in enumerate(summary_df.columns):
            summary_ws.write(0, col_num, value, header_fmt)
        summary_ws.freeze_panes(1, 0)

        # Create each sheet
        for (game, mode), group in merged.groupby(['GAME_ID', 'GAME_MODE']):
            sheet_name = f"{game}_{mode}".replace(" ", "_")[:31]
            group.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]

            # Format headers
            for col_num, value in enumerate(group.columns):
                ws.write(0, col_num, value, header_fmt)

            # Freeze header
            ws.freeze_panes(1, 0)

            # Apply conditional formatting for Drop > 5%
            drop_col_idx = group.columns.get_loc('Total Level Drop')
            ws.conditional_format(1, drop_col_idx, len(group), drop_col_idx, {
                'type': 'cell',
                'criteria': '>',
                'value': 0.05,
                'format': workbook.add_format({'bg_color': 'yellow', 'font_color': 'red'})
            })

            # Add "Back to Main Sheet" link
            ws.write('K1', f'=HYPERLINK("#{summary_sheet}!A1", "← Back to Summary")',
                     workbook.add_format({'bold': True, 'font_color': 'blue'}))

    st.success("✅ Report ready!")
    st.download_button(
        label="📥 Download Game Progression Report",
        data=output.getvalue(),
        file_name="BrainGames_allgameProgression.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
