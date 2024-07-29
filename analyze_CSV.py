import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

def choose_files():
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
    return file_paths

def process_csv_files(file_paths):
    result_data = []
    for file_path in file_paths:
        with open(file_path, 'r', encoding='utf-8') as file:
            block = []
            recording = False
            #for line in file:
            #    line = line.strip()
            #    if line == "[PANEL_INSP_RESULT]":
            #        recording = True
            #    elif line == "[PANEL_INSP_RESULT_END]":
            #        recording = False
            #        result_data.append(block)
            #        block = []
            #    elif recording:
            #        block.append(line)
            for line in file:
                line = line.strip()
                if line == "[PANEL_INSP_RESULT]":
                    recording = True
                elif line == "[PANEL_INSP_RESULT_END]":
                    recording = False
                    if len(block) > 1:
                        key = block[1]
                        values = block[2:]
                        result_data.append({key: values})
                elif recording:
                    block.append(line)
    return result_data

def save_to_excel(data, output_path):
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)
    
def open_excel(file_path):
    os.system(f'start EXCEL.EXE "{file_path}"')

if __name__ == "__main__":
    csv_files = choose_files()
    if not csv_files:
        print("No files selected.")
    else:
        data = process_csv_files(csv_files)
        if not data:
            print("No valid data found in the selected files.")
        else:
            output_path = "output.xlsx"
            save_to_excel(data, output_path)
            open_excel(output_path)
