
import streamlit as st
import pandas as pd
import os
from io import BytesIO
import xlsxwriter

st.set_page_config(page_title="Game Analytics Processor", layout="wide")
st.title("📊 Game Analytics Processor")

uploaded_files = st.file_uploader("Upload CSV pairs (LEVEL_START and LEVEL_COMPLETE for each game)", type=["csv"], accept_multiple_files=True)

def preprocess_and_merge(start_df, complete_df):
    start_df = start_df.rename(columns={"LEVEL": "LEVEL_CLEAN", "USERS": "Start Users"})
    complete_df = complete_df.rename(columns={"LEVEL": "LEVEL_CLEAN", "USERS": "Complete Users"})

    df = pd.merge(start_df[["LEVEL_CLEAN", "Start Users"]], complete_df[["LEVEL_CLEAN", "Complete Users"]], on="LEVEL_CLEAN", how="outer").fillna(0)
    df["LEVEL_CLEAN"] = df["LEVEL_CLEAN"].str.extract("(\d+)").astype(int)

    df = df.sort_values(by="LEVEL_CLEAN").reset_index(drop=True)

    df["Game Play Drop"] = ((df["Start Users"] - df["Complete Users"]) / df["Start Users"].replace(0, pd.NA)) * 100
    df["Popup Drop"] = ((df["Complete Users"] - df["Start Users"].shift(-1)) / df["Complete Users"].replace(0, pd.NA)) * 100
    df["Total Level Drop"] = ((df["Start Users"] - df["Start Users"].shift(-1)) / df["Start Users"].replace(0, pd.NA)) * 100

    level1_users = df[df["LEVEL_CLEAN"] == 1]["Complete Users"].values[0] if (df["LEVEL_CLEAN"] == 1).any() else 0
    level2_users = df[df["LEVEL_CLEAN"] == 2]["Complete Users"].values[0] if (df["LEVEL_CLEAN"] == 2).any() else 0
    max_start_users = max(level1_users, level2_users) if max(level1_users, level2_users) != 0 else 1

    df["Retention %"] = (df["Complete Users"] / max_start_users) * 100
    df[["Game Play Drop", "Popup Drop", "Total Level Drop", "Retention %"]] = df[["Game Play Drop", "Popup Drop", "Total Level Drop", "Retention %"]].round(2)

    return df

if uploaded_files and len(uploaded_files) % 2 == 0:
    games_data = {}
    for i in range(0, len(uploaded_files), 2):
        start_file = uploaded_files[i]
        complete_file = uploaded_files[i+1]

        game_name = Path(start_file.name).stem.replace("_LEVEL_START", "").replace("LEVEL_START", "").strip()

        start_df = pd.read_csv(start_file)
        complete_df = pd.read_csv(complete_file)
        merged_df = preprocess_and_merge(start_df, complete_df)
        games_data[game_name] = merged_df

    if games_data:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            workbook = writer.book
            summary_df = pd.DataFrame()
            for idx, (game, df) in enumerate(games_data.items(), start=2):
                df.to_excel(writer, sheet_name=game, index=False)
                worksheet = writer.sheets[game]
                worksheet.write_url("A1", f"internal:'MAIN_TAB'!A1", string="Go to MAIN_TAB")
                summary_df = pd.concat([summary_df, df.assign(Game=game)])

            summary_df.to_excel(writer, sheet_name="MAIN_TAB", index=False)
            summary_ws = writer.sheets["MAIN_TAB"]
            for col_num, value in enumerate(summary_df.columns.values):
                summary_ws.write(0, col_num, value)

            # Apply conditional formatting
            drop_cols = ["Game Play Drop", "Popup Drop", "Total Level Drop"]
            for col in drop_cols:
                if col in summary_df.columns:
                    col_idx = summary_df.columns.get_loc(col)
                    summary_ws.conditional_format(1, col_idx, len(summary_df), col_idx, {
                        "type": "3_color_scale"
                    })
            # Add charts
            chart = workbook.add_chart({"type": "column"})
            row_count = len(summary_df)
            col_idx = summary_df.columns.get_loc("Game Play Drop")
            chart.add_series({
                "name": "Drop %",
                "categories": f"=MAIN_TAB!$C$2:$C${row_count+1}",
                "values":     f"=MAIN_TAB!${chr(65+col_idx)}$2:${chr(65+col_idx)}${row_count+1}",
            })
            chart.set_title({"name": "Game Play Drop Comparison"})
            summary_ws.insert_chart("L2", chart)

            line_chart = workbook.add_chart({"type": "line"})
            col_idx_ret = summary_df.columns.get_loc("Retention %")
            line_chart.add_series({
                "name": "Retention %",
                "categories": f"=MAIN_TAB!$C$2:$C${row_count+1}",
                "values":     f"=MAIN_TAB!${chr(65+col_idx_ret)}$2:${chr(65+col_idx_ret)}${row_count+1}",
            })
            line_chart.set_title({"name": "Retention Comparison"})
            summary_ws.insert_chart("L18", line_chart)

        st.success("Excel Report Generated with MAIN_TAB and all Games")
        st.download_button("📥 Download Excel Report", data=output.getvalue(), file_name="GameAnalytics_Report.xlsx")
else:
    st.warning("Please upload an even number of CSV files (pairs of LEVEL_START and LEVEL_COMPLETE).")
