# Excel processing tools

## 1. Excel Column Translator
This project provides a graphical interface for translating columns from Excel files using the DeepL API. It allows the user to select an Excel file, choose a column for translation, and specify the source and target languages. The translated column is then saved in a new file.

### Features
- Drag-and-drop file selection.
- Choose source and target languages for translation.
- Supports multiple languages, including automatic language detection.
- Displays progress during translation.
- Saves translated content in a new Excel file.

### Requirements
- Python 3.x
- Required Libraries:
  - `os`
  - `pandas`
  - `openpyxl`
  - `requests`
  - `tkinter`
  - `tkinterdnd2`
  - `string`

Install the necessary libraries using the following command:
```bash
pip install pandas openpyxl requests tkinter tkinterdnd2
