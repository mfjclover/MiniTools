import os
import pandas as pd
import openpyxl
import requests
from tkinter import Label, Entry, Button, filedialog, messagebox, StringVar, OptionMenu
from tkinter.ttk import Progressbar
from tkinterdnd2 import DND_FILES, TkinterDnD
import string

# DeepL API Key
API_KEY = ''

# Available languages for translation (DeepL supported languages including Auto-detect for source)
LANGUAGES = {
    'Auto-detect': 'AUTO',
    'Bulgarian': 'BG',
    'Czech': 'CS',
    'Danish': 'DA',
    'German': 'DE',
    'Greek': 'EL',
    'English': 'EN',
    'Spanish': 'ES',
    'Estonian': 'ET',
    'Finnish': 'FI',
    'French': 'FR',
    'Hungarian': 'HU',
    'Indonesian': 'ID',
    'Italian': 'IT',
    'Japanese': 'JA',
    'Korean': 'KO',
    'Lithuanian': 'LT',
    'Latvian': 'LV',
    'Dutch': 'NL',
    'Polish': 'PL',
    'Portuguese': 'PT',
    'Romanian': 'RO',
    'Russian': 'RU',
    'Slovak': 'SK',
    'Slovenian': 'SL',
    'Swedish': 'SV',
    'Turkish': 'TR',
    'Ukrainian': 'UK',
    'Chinese': 'ZH'
}


# Function to translate using DeepL
def translate_deepl(text, api_key, source_lang, target_lang):
    url = "https://api-free.deepl.com/v2/translate"
    params = {
        "auth_key": api_key,
        "text": text,
        "target_lang": target_lang,
    }

    if source_lang != 'AUTO':
        params["source_lang"] = source_lang

    response = requests.post(url, data=params)

    if response.status_code != 200:
        messagebox.showerror("API Error", f"DeepL API returned an error: {response.status_code} - {response.text}")
        return None

    result = response.json()

    if 'translations' not in result:
        messagebox.showerror("Translation Error", f"Error in translation response: {result}")
        return None

    return result["translations"][0]["text"]


# Function to translate the selected Excel column using DeepL
def translate_excel_column_deepl(input_file, column_name, api_key, source_lang, target_lang):
    try:
        wb = openpyxl.load_workbook(input_file)
        ws = wb.active
    except FileNotFoundError:
        messagebox.showerror("Error", f"The file {input_file} was not found")
        return
    except Exception as e:
        messagebox.showerror("Error", f"Error reading the Excel file: {e}")
        return

    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        messagebox.showerror("Error", f"Error reading the Excel file: {e}")
        return

    if column_name not in df.columns:
        messagebox.showerror("Error", f"The column '{column_name}' is not in the Excel file.")
        return

    total_rows = len(df)
    progress = 0

    for index, row in df.iterrows():
        if pd.notnull(row[column_name]):
            text_to_translate = str(row[column_name]).strip()
            if text_to_translate:
                translation = translate_deepl(text_to_translate, api_key, source_lang, target_lang)
                if translation is None:
                    return

                ws.cell(row=index + 2, column=df.columns.get_loc(column_name) + 1, value=translation)

        progress += 1
        progress_bar['value'] = (progress / total_rows) * 100
        root.update_idletasks()

    output_directory = os.path.join(os.path.dirname(input_file), 'translate')
    os.makedirs(output_directory, exist_ok=True)

    file_name, extension = os.path.splitext(os.path.basename(input_file))

    output_file = os.path.join(output_directory, f"{file_name}-translated{extension}")

    wb.save(output_file)
    messagebox.showinfo("Success", f"Translated file saved at: {output_file}")


# Function to select a file
def select_file():
    input_file = filedialog.askopenfilename(title="Select Excel file",
                                            filetypes=[("Excel files", "*.xlsx")])
    file_entry.config(state='normal')
    file_entry.delete(0, 'end')
    file_entry.insert(0, input_file)
    file_entry.config(state='readonly')
    load_columns(input_file)


# Function to select a file via drag-and-drop
def select_file_drop(event):
    input_file = event.data
    input_file = input_file.strip("{}")
    file_entry.config(state='normal')
    file_entry.delete(0, 'end')
    file_entry.insert(0, input_file)
    file_entry.config(state='readonly')
    load_columns(input_file)


# Function to load column names from the Excel file
def load_columns(file_path):
    try:
        df = pd.read_excel(file_path)
        column_options.set('')
        column_menu['menu'].delete(0, 'end')

        for idx, col in enumerate(df.columns):
            column_letter = string.ascii_uppercase[idx]
            column_label = f"{column_letter}: {col}"
            column_menu['menu'].add_command(label=column_label,
                                            command=lambda value=column_label: column_options.set(value))

        column_options.set(f"A: {df.columns[0]}")
    except Exception as e:
        messagebox.showerror("Error", f"Error reading the Excel file: {e}")


# Function to execute translation
def execute_translation():
    input_file = file_entry.get()
    selected_column = column_options.get()

    if ": " not in selected_column:
        messagebox.showerror("Error", "Invalid column format. Please select a valid column.")
        return

    selected_column = selected_column.split(": ")[1]

    source_lang = source_lang_options.get()
    target_lang = target_lang_options.get()

    if not input_file:
        messagebox.showerror("Error", "Please select a file.")
        return
    if not selected_column:
        messagebox.showerror("Error", "Please select a column.")
        return
    if source_lang == target_lang:
        messagebox.showerror("Error", "Source language and target language cannot be the same.")
        return

    translate_excel_column_deepl(input_file, selected_column, API_KEY, LANGUAGES[source_lang], LANGUAGES[target_lang])


# Create the graphical interface
root = TkinterDnD.Tk()
root.title("Excel Column Translator")

# File selection
Label(root, text="Excel file:").grid(row=0, column=0, padx=10, pady=10)
file_entry = Entry(root, width=50, state='readonly')
file_entry.grid(row=0, column=1, padx=10, pady=10)
Button(root, text="Select", command=select_file).grid(row=0, column=2, padx=10, pady=10)

# Dropdown for column selection
column_options = StringVar(root)
column_menu = OptionMenu(root, column_options, '')
Label(root, text="Select Column:").grid(row=1, column=0, padx=10, pady=10)
column_menu.grid(row=1, column=1, padx=10, pady=10)

# Dropdown for source language selection
source_lang_options = StringVar(root)
source_lang_menu = OptionMenu(root, source_lang_options, *LANGUAGES.keys())
Label(root, text="Source Language:").grid(row=2, column=0, padx=10, pady=10)
source_lang_menu.grid(row=2, column=1, padx=10, pady=10)
source_lang_options.set('Auto-detect')  # Default source language set to Auto-detect

# Dropdown for target language selection
target_lang_options = StringVar(root)
target_lang_menu = OptionMenu(root, target_lang_options, *LANGUAGES.keys())
Label(root, text="Target Language:").grid(row=3, column=0, padx=10, pady=10)
target_lang_menu.grid(row=3, column=1, padx=10, pady=10)
target_lang_options.set('French')  # Default target language

# Button to execute translation
Button(root, text="Translate", command=execute_translation).grid(row=4, columnspan=3, pady=10)

# Progress bar
progress_bar = Progressbar(root, orient='horizontal', length=300, mode='determinate')
progress_bar.grid(row=5, columnspan=3, pady=10)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', select_file_drop)

root.mainloop()
