import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import numpy as np

def load_csv_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        return pd.read_csv(file_path)
    return None

def convert_and_sort(df, column_name):
    # remove dollar signs and commas, then convert to numeric
    df[column_name] = df[column_name].replace('[$,]', '', regex=True).astype(float)
    sorted_df = df.sort_values(by=column_name, ascending=False)
    return sorted_df[column_name].values, sorted_df.index

def create_matchbooks(our_df, bank_df, ours_is_credit):
    our_column, bank_column = ('Credits', 'Debits') if ours_is_credit else ('Debits', 'Credits')
    our_values, our_indices = convert_and_sort(our_df, our_column)
    bank_values, bank_indices = convert_and_sort(bank_df, bank_column)

    results = []
    our_index = bank_index = 0

    while our_index < len(our_values) and bank_index < len(bank_values):
        our_val = our_values[our_index]
        bank_val = bank_values[bank_index]

        if our_val == bank_val:
            if our_val == 0:
                break
            our_row = our_df.loc[our_indices[our_index]]     # TODO is this robust to identical values ?
            bank_row = bank_df.loc[bank_indices[bank_index]] # ^
            results.append({
                'Our_Date': our_row['Date'],
                'Our_Comment': our_row['Comment'],
                'Our_Value': our_val,
                'Bank_Value': bank_val,
                'Bank_Date': bank_row['Date'],
                'Bank_Description': bank_row['Description'],
                'Match': 'MATCH'
            })
            our_index += 1
            bank_index += 1
        elif our_val > bank_val:
            our_row = our_df.loc[our_indices[our_index]]
            results.append({
                'Our_Date': our_row['Date'],
                'Our_Comment': our_row['Comment'],
                'Our_Value': our_val,
                'Bank_Value': 'XXX',
                'Bank_Date': 'XXX',
                'Bank_Description': 'XXX',
                'Match': 'MISMATCH'
            })
            our_index += 1
        else:
            bank_row = bank_df.loc[bank_indices[bank_index]]
            results.append({
                'Our_Date': 'XXX',
                'Our_Comment': 'XXX',
                'Our_Value': 'XXX',
                'Bank_Value': bank_val,
                'Bank_Date': bank_row['Date'],
                'Bank_Description': bank_row['Description'],
                'Match': 'MISMATCH'
            })
            bank_index += 1

    # Remaining entries from our_column
    while our_index < len(our_values):
        if our_values[our_index] == 0: break
        our_row = our_df.loc[our_indices[our_index]]
        results.append({
            'Our_Date': our_row['Date'],
            'Our_Comment': our_row['Comment'],
            'Our_Value': our_row[our_column],
            'Bank_Value': 'XXX',
            'Bank_Date': 'XXX',
            'Bank_Description': 'XXX',
            'Match': 'MISMATCH'
        })
        our_index += 1

    # Remaining entries from bank_column
    while bank_index < len(bank_values):
        if bank_values[bank_index] == 0: break
        bank_row = bank_df.loc[bank_indices[bank_index]]
        results.append({
            'Our_Date': 'XXX',
            'Our_Comment': 'XXX',
            'Our_Value': 'XXX',
            'Bank_Value': bank_row[bank_column],
            'Bank_Date': bank_row['Date'],
            'Bank_Description': bank_row['Description'],
            'Match': 'MISMATCH'
        })
        bank_index += 1

    return pd.DataFrame(results)

def save_to_excel_with_color(df, filename):
    wb = Workbook()
    ws = wb.active

    green_fill = PatternFill(start_color='00C851', end_color='00C851', fill_type='solid')
    red_fill = PatternFill(start_color='FF4444', end_color='FF4444', fill_type='solid')

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
        ws.append(row)
        if r_idx==0:
            continue
        row_color = green_fill if row[-1] == 'MATCH' else red_fill # TODO
        for cell in ws[r_idx + 1]:
            cell.fill = row_color

    ws.delete_cols(ws.max_column)
    wb.save(filename)
    print("Saved " + filename)

class CSVMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Matcher")

        self.our_csv = None
        self.bank_csv = None

        self.our_file_btn = tk.Button(root, text="Load OUR CSV", command=self.load_our_csv)
        self.our_file_btn.pack()

        self.bank_file_btn = tk.Button(root, text="Load BANK CSV", command=self.load_bank_csv)
        self.bank_file_btn.pack()

        self.match_btn = tk.Button(root, text="Match CSVs", command=self.match_csvs, state=tk.DISABLED)
        self.match_btn.pack()

    def load_our_csv(self):
        self.our_csv = load_csv_file()
        if self.our_csv is not None:
            self.check_files_loaded()

    def load_bank_csv(self):
        self.bank_csv = load_csv_file()
        if self.bank_csv is not None:
            self.check_files_loaded()

    def check_files_loaded(self):
        if self.our_csv is not None and self.bank_csv is not None:
            self.match_btn.config(state=tk.NORMAL)

    def match_csvs(self):
        matched_df = create_matchbooks(self.our_csv, self.bank_csv, True)
        save_to_excel_with_color(matched_df, 'matched_output_ourcredits.xlsx')
        matched_df = create_matchbooks(self.our_csv, self.bank_csv, False)
        save_to_excel_with_color(matched_df, 'matched_output_ourdebits.xlsx')

root = tk.Tk()
app = CSVMatcherApp(root)
root.mainloop()

