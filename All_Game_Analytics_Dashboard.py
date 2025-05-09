import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import CellIsRule
from datetime import datetime

def clean_level(level_str):
    """Extract numerical value from level string"""
    try:
        return int(level_str.split('_')[-1])
    except (ValueError, AttributeError):
        return np.nan

def process_file(filename, file_type):
    """Process either start or complete file"""
    df = pd.read_csv(filename)
    
    # Select and rename columns based on file type
    if file_type == 'start':
        df = df[['GAME_ID', 'DIFFICULTY', 'LEVEL', 'USERS']]
        df.rename(columns={'USERS': 'USERS_START'}, inplace=True)
    elif file_type == 'complete':
        df = df[['GAME_ID', 'DIFFICULTY', 'LEVEL', 'USERS', 
                'PLAY_TIME', 'SKIPPED', 'HINT_USED', 'ATTEMPTS']]
        df.rename(columns={'USERS': 'USERS_COMPLETE'}, inplace=True)
    
    # Clean level numbers
    df['LEVEL_CLEAN'] = df['LEVEL'].apply(clean_level)
    df = df[~df['LEVEL_CLEAN'].isna()]
    df['LEVEL_CLEAN'] = df['LEVEL_CLEAN'].astype(int)
    
    return df.drop(columns=['LEVEL']).rename(columns={'LEVEL_CLEAN': 'LEVEL'})

def create_game_progression_report(started_file, completed_file):
    # Process both files
    df_start = process_file(started_file, 'start')
    df_complete = process_file(completed_file, 'complete')

    # Merge data
    merged = pd.merge(df_start, df_complete,
                     on=['GAME_ID', 'DIFFICULTY', 'LEVEL'],
                     how='outer').fillna(0)

    # Convert numeric columns to appropriate types
    numeric_cols = ['USERS_START', 'USERS_COMPLETE', 
                   'PLAY_TIME', 'SKIPPED', 'HINT_USED', 'ATTEMPTS']
    merged[numeric_cols] = merged[numeric_cols].astype(float)

    # Calculate metrics
    merged['GAME_PLAY_DROP'] = np.where(
        merged['USERS_START'] > 0,
        (merged['USERS_START'] - merged['USERS_COMPLETE']) / merged['USERS_START'],
        0
    )
    merged['RETENTION_%'] = np.where(
        merged['USERS_START'] > 0,
        merged['USERS_COMPLETE'] / merged['USERS_START'],
        0
    )
    merged['POPUP_DROP'] = 0  # Placeholder for actual popup drop calculation
    merged['TOTAL_LEVEL_DROP'] = merged['GAME_PLAY_DROP'] + merged['POPUP_DROP']

    # Format percentages
    percentage_cols = ['GAME_PLAY_DROP', 'POPUP_DROP', 'TOTAL_LEVEL_DROP', 'RETENTION_%']
    merged[percentage_cols] = merged[percentage_cols].applymap(
        lambda x: f"{x:.2%}" if isinstance(x, float) else x
    )

    # Create Excel workbook
    wb = Workbook()
    main_sheet = wb.active
    main_sheet.title = "MAIN_TAB"
    
    # Main sheet headers
    main_headers = ["Index", "Sheet Name", "GAME_ID", "DIFFICULTY", "Link to Sheet"]
    main_sheet.append(main_headers)
    
    # Formatting styles
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    hyperlink_font = Font(color="0000FF", underline="single")
    
    # Create game sheets
    game_groups = merged.groupby(['GAME_ID', 'DIFFICULTY'])
    for idx, ((game_id, difficulty), group) in enumerate(game_groups, start=1):
        # Create sheet
        sheet_name = f"{game_id}_{difficulty}"[:31]
        ws = wb.create_sheet(title=sheet_name)
        
        # Sheet header
        headers = ["BACK TO MAIN_TAB", "GAME_ID", "DIFFICULTY", "Level", 
                  "Start Users", "Complete Users", "Game Play Drop", 
                  "Popup Drop", "Total Level Drop", "Retention %",
                  "PLAY_TIME_AVG", "SKIPPED_SUM", "HINT_USED_SUM", "ATTEMPTS_AVG"]
        ws.append(headers)
        
        # Add data rows
        for _, row in group.iterrows():
            ws.append([
                "",  # Placeholder for back link
                game_id,
                difficulty,
                row['LEVEL'],
                int(row['USERS_START']),
                int(row['USERS_COMPLETE']),
                row['GAME_PLAY_DROP'],
                row['POPUP_DROP'],
                row['TOTAL_LEVEL_DROP'],
                row['RETENTION_%'],
                round(row['PLAY_TIME'], 2),
                int(row['SKIPPED']),
                int(row['HINT_USED']),
                round(row['ATTEMPTS'], 2)
            ])
        
        # Add hyperlink back to main sheet
        for row in ws.iter_rows(min_row=2, max_row=len(group)+1, max_col=1):
            row[0].value = f'=HYPERLINK("#MAIN_TAB!A1", "Back")'
            row[0].font = hyperlink_font
        
        # Formatting
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            
        # Conditional formatting for drops >5%
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00")
        red_font = Font(color="FF0000", bold=True)
        drop_columns = ['G', 'H', 'I']  # Columns G, H, I for drops
        for col in drop_columns:
            ws.conditional_formatting.add(f"{col}2:{col}{len(group)+1}",
                CellIsRule(operator='greaterThan', formula=['0.05'], 
                          stopIfTrue=True, fill=yellow_fill, font=red_font)
            )
        
        # Add to main sheet
        main_sheet.append([
            idx,
            sheet_name,
            game_id,
            difficulty,
            f'=HYPERLINK("#\'{sheet_name}\'!A1", "Click to view")'
        ])
    
    # Format main sheet
    for cell in main_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
    
    # Set column widths
    for col in ['A', 'B', 'C', 'D', 'E']:
        main_sheet.column_dimensions[col].width = 15
    
    # Save file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"BrainGames_allgameProgression-ver.21(6-8 May)_{timestamp}.xlsx"
    wb.save(filename)
    print(f"Report generated: {filename}")

# Usage
create_game_progression_report("level_started.csv", "level_completed.csv")
