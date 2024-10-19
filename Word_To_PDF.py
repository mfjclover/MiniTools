import os
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
from tkinterdnd2 import DND_FILES, TkinterDnD

input_path = ""


# Opens a directory dialog and stores the selected path in `input_path`.
def select_directory():
    global input_path
    try:
        input_path = filedialog.askdirectory(title="Select the input directory")
        if input_path:
            directory_label.config(text=f"Selected directory: {input_path}")
            convert_button.config(state=tk.NORMAL)
        else:
            messagebox.showwarning("Warning", "No directory selected.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while selecting the directory: {e}")


# Handles dropped files or folders, updating `input_path` and the UI.
def drop(event):
    global input_path
    try:
        data = root.tk.splitlist(event.data)
        for item in data:
            if os.path.isdir(item):
                input_path = item
                directory_label.config(text=f"Selected directory: {input_path}")
                convert_button.config(state=tk.NORMAL)
            elif os.path.isfile(item) and item.endswith(".docx"):
                input_path = item
                directory_label.config(text=f"Selected file: {input_path}")
                convert_button.config(state=tk.NORMAL)
            else:
                messagebox.showwarning("Warning", "Please drop a valid directory or .docx file.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while handling the dropped file: {e}")


# Updates the progress bar based on the current conversion progress.
def update_progress(current, total):
    try:
        progress["value"] = current
        progress_label.config(text=f"Progress: {current} of {total}")
        root.update_idletasks()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while updating the progress: {e}")


# Converts .docx files to PDF and updates progress and UI with results.
def convert_docs():
    global input_path
    try:
        if os.path.isfile(input_path):
            docx_files = [input_path]
            output_dir = os.path.join(os.path.dirname(input_path), "convertedToPdf")
        else:
            output_dir = os.path.join(input_path, "convertedToPdf")
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            files = os.listdir(input_path)
            docx_files = [os.path.join(input_path, f) for f in files if f.endswith(".docx")]

        if not docx_files:
            messagebox.showwarning("Warning", "No .docx files found to convert.")
            return

        total_files = len(docx_files)
        success_count = 0
        failed_files = []

        progress["maximum"] = total_files

        word = win32.Dispatch("Word.Application")
        word.Visible = False

        current_progress = 0
        for index, docx_file in enumerate(docx_files, start=1):
            output_path = os.path.join(output_dir, os.path.basename(docx_file).replace(".docx", ".pdf"))

            try:
                doc = word.Documents.Open(os.path.abspath(docx_file))
                doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
                doc.Close()
                success_count += 1
            except Exception as e:
                failed_files.append(docx_file)
                messagebox.showerror("Conversion Error", f"Error converting {docx_file}: {e}")

            current_progress += 1
            update_progress(current_progress, total_files)

        word.Quit()

        info_label.config(
            text=f"Total: {total_files}, Converted: {success_count}, Failed: {len(failed_files)}")
        messagebox.showinfo("Conversion complete",
                            f"Successfully converted files: {success_count}\nFailed files: {len(failed_files)}")

        if failed_files:
            print("Files that could not be converted:")
            for file in failed_files:
                print(file)
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred during conversion: {e}")


# Initialize the main window with drag-and-drop support and set its title and size.
root = TkinterDnD.Tk()
root.title("Word to PDF Converter")
root.geometry("400x350")

# Create a label instructing the user to select a folder or drag and drop files.
label = tk.Label(root, text="Select the input directory or drag and drop a file/folder here")
label.pack(pady=10)

# Create a button that allows the user to select a folder.
select_button = tk.Button(root, text="Select folder", command=select_directory)
select_button.pack(pady=5)

# Label to display the selected directory or file
directory_label = tk.Label(root, text="Selected directory or file: None")
directory_label.pack(pady=5)

# Button to trigger the file conversion
convert_button = tk.Button(root, text="Convert", command=convert_docs, state=tk.DISABLED)
convert_button.pack(pady=5)

# Create a progress bar to visually indicate conversion progress.
progress = Progressbar(root, orient=tk.HORIZONTAL, length=300, mode='determinate')
progress.pack(pady=20)

# Label to display progress updates, initially empty.
progress_label = tk.Label(root, text="")
progress_label.pack(pady=5)

# Label to show final conversion information
info_label = tk.Label(root, text="")
info_label.pack(pady=5)

# Enable drag-and-drop functionality for files and bind it to the 'drop' event.
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', drop)

root.mainloop()
