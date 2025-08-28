# -*- coding: utf-8 -*-
"""
Attendance List Generator

This script creates attendance lists based on data from Mentimeter (Excel)
and Blackboard (CSV) group exports.

Usage:
    1. Export group lists from Blackboard as CSV.
    2. Export attendance from Mentimeter as Excel.
    3. Run this script and follow the prompts.

Created on Fri Aug 23 13:58:40 2024
Author: krystad

Requirements:
    - pandas
    - numpy
    - easygui
    - openpyxl (for Excel handling)
"""

import sys
import csv
from datetime import datetime
import easygui
import pandas as pd
import numpy as np

# ---------- Helper Functions ----------

def format_date(date_str):
    """
    Convert a string date 'YYYY-MM-DD' to a datetime object.
    """
    return datetime(int(date_str[:4]), int(date_str[5:7]), int(date_str[8:]))

def short_date(dt):
    """
    Convert a datetime object to 'YYYY-MM-DD' string.
    """
    return str(dt).split(" ")[0]

def remove_quotes(s):
    """
    Remove quotes from a string.
    """
    if isinstance(s, str):
        return s.replace('"', '')
    return s

def show_intro():
    """
    Display the introduction dialog and options.
    """
    return easygui.buttonbox(
        'This program creates an attendance list. You need:\n\n'
        '1. An Excel file from Mentimeter\n'
        '2. A CSV file exported from Groups in Blackboard\n\n',
        'Attendance List',
        ('Continue (standard username)', 'Continue (custom username)', 'More info', 'Cancel')
    )

def show_info():
    """
    Display info dialog for exporting Blackboard group lists.
    """
    return easygui.buttonbox(
        'To download the group list from Blackboard (only needed the first time or if the group list has changed):\n\n'
        '1. In your course on Blackboard, go to Groups. Click Export to Excel.\n'
        '2. Save the file as .csv\n'
        '3. Use this file in the program.',
        'More info',
        ('OK', 'Cancel')
    )

# ---------- Main Script ----------

def main():
    """
    Main workflow for generating attendance list.
    """
    intro = show_intro()
    custom_username = intro == 'Continue (custom username)'

    # Exit conditions
    if intro in (None, 'Cancel'):
        sys.exit()
    elif intro == 'More info':
        if show_info() == 'Cancel':
            sys.exit()
        intro = show_intro()
        if intro in (None, 'Cancel'):
            sys.exit()
        custom_username = intro == 'Continue (custom username)'

    # Select files
    mentimeter_file = easygui.fileopenbox('Select Excel file from Mentimeter', 'Attendance', '*.xlsx')
    if mentimeter_file is None:
        sys.exit()
    student_file = easygui.fileopenbox('Select CSV file from Blackboard', 'Student List', '*.csv')
    if student_file is None:
        sys.exit()

    # Read student CSV and clean data
    df_students = pd.read_csv(
        student_file,
        names=("Class", "Username", "No", "FirstName", "LastName"),
        sep=',\s*',
        engine='python'
    )
    df_students = df_students.applymap(remove_quotes)
    df_students = df_students.drop(columns=['No'], errors='ignore')

    # Read Mentimeter Excel and extract attendance sheets
    excel = pd.ExcelFile(mentimeter_file)
    sheet_names = excel.sheet_names
    num_sheets = len(sheet_names)

    first_date, last_date = None, None
    missing_usernames = []
    missing_details = []
    ignore_usernames = []
    username_map = {}

    # Process each attendance sheet (skip first, usually metadata)
    for i in range(1, num_sheets):
        df_attendance = pd.read_excel(mentimeter_file, sheet_name=i)
        date = format_date(df_attendance["Unnamed: 1"][0])
        df_students[date] = np.nan

        # Track first/last date
        if i == 1:
            first_date = date
        if i == num_sheets - 1:
            last_date = date

        date_str = date.strftime("%d.%m")
        for raw_username in df_attendance["Question 1"][7:]:
            username = str(raw_username).lower().split("@")[0].replace(" ", "")
            if username in ignore_usernames:
                continue
            if username in username_map:
                username = username_map[username]

            match_idx = np.where(df_students["Username"] == username)[0]
            if len(match_idx) == 0:
                missing_usernames.append(username)
                missing_details.append(date_str)
                if custom_username:
                    new_username = easygui.enterbox(
                        f'Username {username} not found. Correct username: (Type i to ignore)',
                        'Custom username'
                    )
                    if new_username == "i":
                        ignore_usernames.append(username)
                        continue
                    elif new_username is None:
                        easygui.msgbox('Program exited.')
                        sys.exit()
                    username_map[username] = new_username
                    match_idx = np.where(df_students["Username"] == new_username)[0]
                    if len(match_idx):
                        df_students.loc[match_idx[0], date] = 1
                    else:
                        easygui.msgbox(f'{new_username} not found either')
            elif len(match_idx) == 1:
                df_students.loc[match_idx[0], date] = 1

    # Report missing usernames
    if missing_usernames:
        msg = "\n".join(
            f"{name} ({date})"
            for name, date in zip(missing_usernames, missing_details)
        )
        user_choice = easygui.buttonbox(
            f'These usernames were not found:\n{msg}',
            'Username not found in student list',
            ('OK', 'Cancel')
        )
        if user_choice == "Cancel":
            easygui.msgbox('Program exited.')
            sys.exit()
    else:
        easygui.msgbox('All usernames found in the student list.')

    # Summarize attendance
    df_students.loc[len(df_students)] = df_students.sum(axis=0, numeric_only=True)
    df_students.insert(loc=0, column="TimesPresent", value=np.nan)
    df_students["TimesPresent"] = df_students.sum(axis=1, numeric_only=True)

    # Filter students that attended at least once
    present_mask = df_students["TimesPresent"] > 0
    df_present = df_students[present_mask].reset_index(drop=True)

    # Hide sum row's "TimesPresent"
    if not df_present.empty:
        df_present.loc[len(df_present) - 1, "TimesPresent"] = None

    # Rename date columns to readable format
    df_present = df_present.rename(columns={
        col: col.strftime("%d.%m") for col in df_present.columns if isinstance(col, datetime)
    })
    df_present = df_present.sort_values(by="TimesPresent", ascending=False)

    output_filename = (
        mentimeter_file[:-5] +
        f"_from_{short_date(first_date)}_to_{short_date(last_date)}.xlsx"
    )

    # Prepare missing usernames DataFrame
    df_missing = pd.DataFrame({"Username": missing_usernames, "Date": missing_details})

    # Save to Excel
    try:
        with pd.ExcelWriter(output_filename) as writer:
            df_present.to_excel(writer, sheet_name='Attendance', index=False)
            if not df_missing.empty:
                df_missing.to_excel(writer, sheet_name="Not found", index=False)
    except PermissionError:
        easygui.msgbox('Cannot write to Excel file. You need to close it. Program exited.', 'Attendance List')
        sys.exit()

    easygui.msgbox(f'Attendance saved in file\n\n {output_filename}', 'Attendance List')

if __name__ == "__main__":
    main()
