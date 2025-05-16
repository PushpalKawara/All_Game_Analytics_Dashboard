# ========================== Step 1: Required Imports ========================== #
import streamlit as st
import pandas as pd
import numpy as np
import os
import datetime
import matplotlib.pyplot as plt
from io import BytesIO
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage

# ========================== Step 2: Streamlit Config ========================== #
st.set_page_config(page_title="GAME PROGRESSION", layout="wide")
st.title("📊 GAME PROGRESSION Dashboard")

# ========================== Step 3: File Upload Section ========================== #
st.sidebar.header("Upload Level Start & Complete Files")
start_files = st.sidebar.file_uploader("Upload START files", type=["csv", "xlsx"], accept_multiple_files=True)
complete_files = st.sidebar.file_uploader("Upload COMPLETE files", type=["csv", "xlsx"], accept_multiple_files=True)
version = st.sidebar.text_input("Game Version", "1.0.0")
date_selected = st.sidebar.date_input("Report Date", datetime.date.today())

# ========================== Step 4: Data Processing ========================== #
def process_game_data(start_files, complete_files):
    processed_games = {}
    start_map = {os.path.splitext(f.name)[0].upper(): f for f in start_files}
    complete_map = {os.path.splitext(f.name)[0].upper(): f for f in complete_files}
    common_games = set(start_map.keys()) & set(complete_map.keys())

    for game_name in sorted(common_games):
        start_df = load_and_clean_file(start_map[game_name], is_start_file=True)
        complete_df = load_and_clean_file(complete_map[game_name], is_start_file=False)
        if start_df is not None and complete_df is not None:
            merged_df = merge_and_calculate(start_df, complete_df)
            processed_games[game_name] = merged_df
    return processed_games


def load_and_clean_file(file_obj, is_start_file=True):
    try:
        df = pd.read_csv(file_obj) if file_obj.name.endswith('.csv') else pd.read_excel(file_obj)
        # Handle MultiIndex columns
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = ['_'.join(map(str, col)).strip().upper() for col in df.columns]
        else:
            df.columns = df.columns.str.strip().str.upper()
        # Extract LEVEL
        level_col = next((c for c in df.columns if 'LEVEL' in c), None)
        if level_col:
            df['LEVEL'] = df[level_col].astype(str).str.extract(r'(\d+)').astype(int)
        # Rename user column
        user_col = next((c for c in df.columns if 'USER' in c), None)
        if user_col:
            df = df.rename(columns={user_col: 'START_USERS' if is_start_file else 'COMPLETE_USERS'})
        # Select and clean columns
        if is_start_file:
            df = df[['LEVEL', 'START_USERS']].drop_duplicates().sort_values('LEVEL')
        else:
            cols = ['LEVEL', 'COMPLETE_USERS', 'PLAY_TIME_AVG', 'HINT_USED_SUM', 'SKIPPED_SUM', 'ATTEMPT_SUM']
            for c in cols:
                df[c] = df.get(c, 0)
            df = df[cols]
        return df.dropna().sort_values('LEVEL')
    except Exception as e:
        st.error(f"Error processing file {file_obj.name}: {e}")
        return None


def merge_and_calculate(start_df, complete_df):
    m = pd.merge(start_df, complete_df, on='LEVEL', how='outer').sort_values('LEVEL')
    m[['START_USERS', 'COMPLETE_USERS']] = m[['START_USERS', 'COMPLETE_USERS']].fillna(0)
    m['GAME_PLAY_DROP'] = ((m['START_USERS'] - m['COMPLETE_USERS']) / m['START_USERS'].replace(0, np.nan) * 100).round(2)
    m['POPUP_DROP'] = ((m['COMPLETE_USERS'] - m['START_USERS'].shift(-1)) / m['COMPLETE_USERS'].replace(0, np.nan) * 100).round(2)
    m['TOTAL_LEVEL_DROP'] = ((m['START_USERS'] - m['START_USERS'].shift(-1)) / m['START_USERS'].replace(0, np.nan) * 100).round(2)
    m['RETENTION_%'] = (m['START_USERS'] / m['START_USERS'].max() * 100).round(2)
    # Convert other numeric cols to int
    for c in m.select_dtypes(include=[np.number]).columns:
        if c not in ['GAME_PLAY_DROP', 'POPUP_DROP', 'TOTAL_LEVEL_DROP', 'RETENTION_%']:
            m[c] = m[c].fillna(0).astype(int)
    return m

# ========================== Step 5: Charting ========================== #
def create_charts(df, version, date_selected):
    charts = {}
    d100 = df[df['LEVEL'] <= 100]
    # Retention
    fig1, ax1 = plt.subplots(figsize=(12, 6))
    ax1.plot(d100['LEVEL'], d100['RETENTION_%'], linewidth=2)
    format_chart(ax1, 'Retention Chart', version, date_selected)
    charts['retention'] = fig1
    # Total Level Drop
    fig2, ax2 = plt.subplots(figsize=(12, 5))
    ax2.bar(d100['LEVEL'], d100['TOTAL_LEVEL_DROP'])
    format_chart(ax2, 'Total Level Drop', version, date_selected)
    charts['total_drop'] = fig2
    # Combo Drop
    fig3, ax3 = plt.subplots(figsize=(12, 5))
    w=0.4
    ax3.bar(d100['LEVEL']-w/2, d100['GAME_PLAY_DROP'], width=w)
    ax3.bar(d100['LEVEL']+w/2, d100['POPUP_DROP'], width=w)
    format_chart(ax3, 'Game & Popup Drop', version, date_selected)
    charts['combo_drop'] = fig3
    return charts


def format_chart(ax, title, version, date_selected):
    ax.set_xlim(1,100)
    ax.set_title(f"{title} | v{version} | {date_selected}", weight='bold')
    ax.grid(True)

# ========================== Step 6: Excel Report ========================== #
def generate_excel_report(data, version, date_selected):
    wb = Workbook()
    wb.remove(wb.active)
    # Main sheet
    ms = wb.create_sheet('MAIN_TAB')
    ms.append(['Index','Game','Play Drop Count','Popup Drop Count','Total Drop Count','Lvl Start','Users Start','Lvl End','Users End','Link'])
    for i,(name,df) in enumerate(data.items(),1):
        sh = wb.create_sheet(name[:30])
        # Header row with Back button
        sh.append(['=HYPERLINK("#MAIN_TAB!A1","Back to MAIN TAB")','Level','Start Users','Complete Users','Game Play Drop','Popup Drop','Total Drop','Retention %','PLAY_TIME_AVG','HINT_USED_SUM','SKIPPED_SUM','ATTEMPT_SUM'])
        # Data rows
        for _,r in df.iterrows():
            sh.append([
                f'=HYPERLINK("#MAIN_TAB!A1","{name}")',
                r['LEVEL'],r['START_USERS'],r['COMPLETE_USERS'],r['GAME_PLAY_DROP'],r['POPUP_DROP'],r['TOTAL_LEVEL_DROP'],r['RETENTION_%'],
                r.get('PLAY_TIME_AVG',0),r.get('HINT_USED_SUM',0),r.get('SKIPPED_SUM',0),r.get('ATTEMPT_SUM',0)
            ])
        # Add chart images
        for key,fig in create_charts(df,version,date_selected).items():
            buf=BytesIO();fig.savefig(buf, bbox_inches='tight');buf.seek(0)
            img=OpenpyxlImage(buf)
            pos={'retention':'M1','total_drop':'M25','combo_drop':'M50'}[key]
            sh.add_image(img,pos)
        # Main tab summary link
        ms.append([i,name,(df['GAME_PLAY_DROP']>=3).sum(),(df['POPUP_DROP']>=3).sum(),(df['TOTAL_LEVEL_DROP']>=3).sum(),df['LEVEL'].min(),df['START_USERS'].max(),df['LEVEL'].max(),df['COMPLETE_USERS'].iloc[-1],f'=HYPERLINK("#{sh.title}!A1","{name}")'])
    return wb

# ========================== Step 7: Build UI & Download ========================== #
if start_files and complete_files:
    processed = process_game_data(start_files, complete_files)
    if processed:
        wb = generate_excel_report(processed, version, date_selected)
        # Write to bytes buffer
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        data = buf.read()
        st.success("✅ Excel report is ready!")
        st.download_button(
            label="📥 Download Excel Report",
            data=data,
            file_name=f"Game_Progression_{version}_{date_selected}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
