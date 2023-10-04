import tkinter as tk
from tkinter import filedialog
import pandas as pd
import json

def convert_excel_to_json(excel_file_path):
    df = pd.read_excel(excel_file_path)
    json_data = df.to_json(orient='records', lines=True)
    return json_data

def process_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        json_data = convert_excel_to_json(file_path)
        save_json(json_data)

def save_json(json_data):
    save_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
    if save_path:
        with open(save_path, 'w', encoding='utf-8') as json_file:
            json_file.write(json_data)
        print("הקובץ הוזן בהצלחה:", save_path)

root = tk.Tk()
root.title("Convert Excel to JSON")

# יצירת Label עבור הטקסט הנוסף
label = tk.Label(root, text="@By Engineering Solution Division - Bynet Data Communication", font=("Helvetica", 10))
label.pack()

button = tk.Button(root, text="Select Excel File", command=process_file)
button.pack()

root.geometry("400x150")  # כאן נגדיל את גודל החלון

root.mainloop()
