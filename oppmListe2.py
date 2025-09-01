# -*- coding: utf-8 -*-
"""
Oppmøtelistegenerator

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
        'Dette programmet lager oppmøteliste. Du trenger:\n\n'
        '1. En Excel-fil fra Mentimeter\n'
        '2. En csv-fil fra Grupper på Blackboard\n\n',
        'Oppmøteliste',
        ('Videre (standard brukernavn)', 'Videre (tilpass brukernavn)', 'Mer info', 'Avbryt')
    )

def show_info():
    """
    Display the info dialog for exporting Blackboard group lists.
    """
    return easygui.buttonbox(
        'Hente ned gruppeliste fra Blackboard (kun første gang eller hvis gruppelista er endret):\n\n1. Inne på emnet ditt på Blackboard, gå til Grupper. \n   Trykk Eksporter --> Kun gruppemedlemmer.\n2. Du får en epost "Masseeksport fullført". Lagre fila.\n\nOpprette avstemning i Mentimeter (kun første gang):\n\n1. Gå til Mentimeter www.mentimeter.com/auth/saml/ntnu\n2. Trykk New Menti --> Start from scratch --> Open Ended\n3. Øverst skriver du navn på presentasjonen,\n   f.eks. "Oppmøte Teksam 1FA 2024/25".\n3. Bytt ut «Ask your question here…» med "Oppmøte: Skriv ditt NTNU‐brukernavn".\n\nAvstemning:\n\n1. I timen viser du Menti‐presentasjonen. Skru på QR‐kode.\n   Bruk samme presentasjon hver gang.\n2. Etter at alle har skrevet seg inn, trykk Manage Results --> Reset results.\n   Mentimeter har lagret resultatene, selv om du ikke ser dem.\n3. Finn presentasjonen i Mentimeter og trykk\n   View Results --> Download --> Spreadsheet (XLSX).\n\nDet kan hende studenter skriver brukernavnet feil. I så fall kan du trykke "Tilpass brukernavn".',
        'Mer info',
        ('OK', 'Avbryt')
    )

# ---------- Main Script ----------

def main():
    """
    Main workflow for generating attendance list.
    """
    intro = show_intro()
    custom_username = intro == 'Videre (tilpass brukernavn)'

    # Exit conditions
    if intro in (None, 'Avbryt'):
        sys.exit()
    elif intro == 'Mer info':
        if show_info() == 'Avbryt':
            sys.exit()
        intro = show_intro()
        if intro in (None, 'Avbryt'):
            sys.exit()
        custom_username = intro == 'Videre (tilpass brukernavn)'

    # Select files
    mentimeter_file = easygui.fileopenbox('Velg Excel-fil fra Mentimeter', 'Oppmøte', '*.xlsx')
    if mentimeter_file is None:
        sys.exit()
    student_file = easygui.fileopenbox('Velg csv-fil fra Blackboard', 'Studentliste', '*.csv')
    if student_file is None:
        sys.exit()

    # Read student CSV and clean data
    df_students = pd.read_csv(
        student_file,
        names=("Klasse", "Brukernavn", "Nr", "Fornavn", "Etternavn"),
        sep=',\s*',
        engine='python'
    )
    df_students = df_students.applymap(remove_quotes)
    df_students = df_students.drop(columns=['Nr'], errors='ignore')

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

            match_idx = np.where(df_students["Brukernavn"] == username)[0]
            if len(match_idx) == 0:
                missing_usernames.append(username)
                missing_details.append(date_str)
                if custom_username:
                    new_username = easygui.enterbox(
                        f'Brukernavn {username} finnes ikke. Riktig brukernavn: (Skriv i for å ignorere)',
                        'Tilpass brukernavn'
                    )
                    if new_username == "i":
                        ignore_usernames.append(username)
                        continue
                    elif new_username is None:
                        easygui.msgbox('Programmet avsluttes.')
                        sys.exit()
                    username_map[username] = new_username
                    match_idx = np.where(df_students["Brukernavn"] == new_username)[0]
                    if len(match_idx):
                        df_students.loc[match_idx[0], date] = 1
                    else:
                        easygui.msgbox(f'{new_username} finnes heller ikke')
            elif len(match_idx) == 1:
                df_students.loc[match_idx[0], date] = 1

    # Report missing usernames
    if missing_usernames:
        msg = "\n".join(
            f"{name} ({date})"
            for name, date in zip(missing_usernames, missing_details)
        )
        user_choice = easygui.buttonbox(
            f'Disse brukernavnene finnes ikke:\n{msg}',
            'Brukernavn ikke funnet på klasselista',
            ('OK', 'Avbryt')
        )
        if user_choice == "Avbryt":
            easygui.msgbox('Programmet avsluttes.')
            sys.exit()
    else:
        easygui.msgbox('Alle brukernavn funnet på klasselista.')

    # Summarize attendance
    df_students.loc[len(df_students)] = df_students.sum(axis=0, numeric_only=True)
    df_students.insert(loc=0, column="Ganger", value=np.nan)
    df_students["Ganger"] = df_students.sum(axis=1, numeric_only=True)

    # Filter students that attended at least once
    present_mask = df_students["Ganger"] > 0
    df_present = df_students[present_mask].reset_index(drop=True)

    # Hide sum row's "Ganger"
    if not df_present.empty:
        df_present.loc[len(df_present) - 1, "Ganger"] = None

    # Rename date columns to readable format
    df_present = df_present.rename(columns={
        col: col.strftime("%d.%m") for col in df_present.columns if isinstance(col, datetime)
    })
    df_present = df_present.sort_values(by="Ganger", ascending=False)

    output_filename = (
        mentimeter_file[:-5] +
        f"_fra_{short_date(first_date)}_til_{short_date(last_date)}.xlsx"
    )

    # Prepare missing usernames DataFrame
    df_missing = pd.DataFrame({"Brukernavn": missing_usernames, "Dato": missing_details})

    # Save to Excel
    try:
        with pd.ExcelWriter(output_filename) as writer:
            df_present.to_excel(writer, sheet_name='Oppmøte', index=False)
            if not df_missing.empty:
                df_missing.to_excel(writer, sheet_name="Ikke funnet", index=False)
    except PermissionError:
        easygui.msgbox('Kan ikke skrive til Excel-fila. Du må lukke den. Programmet avsluttes.', 'Oppmøteliste')
        sys.exit()

    easygui.msgbox(f'Oppmøte lagret i fil\n\n {output_filename}', 'Oppmøteliste')

if __name__ == "__main__":
    main()

