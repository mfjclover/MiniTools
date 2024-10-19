import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

def get_all_images(directory):
    image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp']
    images = set()
    for root, _, files in os.walk(directory):
        for file in files:
            if any(file.lower().endswith(ext) for ext in image_extensions):
                images.add(os.path.join(root, file))
    return images

def get_unique_filename(directory, filename):
    base, ext = os.path.splitext(filename)
    counter = 1
    unique_filename = filename
    while os.path.exists(os.path.join(directory, unique_filename)):
        unique_filename = f"{base}_{counter}{ext}"
        counter += 1
    return unique_filename

def select_directory(label_var):
    directory = filedialog.askdirectory()
    if directory:
        label_var.set(directory)

def select_file(label_var):
    file = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if file:
        label_var.set(file)

def execute_copy():
    image_directory = source_dir.get()
    name_file = name_file_path.get()
    destination_directory = dest_dir.get()

    if not all([image_directory, name_file, destination_directory]):
        messagebox.showwarning("Input error", "Please select all directories and the .txt file.")
        return

    # Create destination directory if it doesn't exist
    os.makedirs(destination_directory, exist_ok=True)

    # Read listed names
    with open(name_file, 'r') as file:
        listed_names = {line.strip() for line in file if line.strip()}

    # Get all images in the directory and subdirectories
    all_images = get_all_images(image_directory)

    # Determine images to copy
    images_to_copy = set()
    for image in all_images:
        image_name = os.path.splitext(os.path.basename(image))[0]
        if image_name in listed_names:
            images_to_copy.add(image)

    # Copy listed images
    copied_count = 0
    for image in images_to_copy:
        source_path = image
        original_filename = os.path.basename(image)
        unique_filename = get_unique_filename(destination_directory, original_filename)
        destination_path = os.path.join(destination_directory, unique_filename)
        try:
            shutil.copy(source_path, destination_path)
            copied_count += 1
        except OSError as e:
            print(f'Error copying {image}: {e}')

    messagebox.showinfo("Result", f'{copied_count} images have been copied.')

# Create the GUI
root = tk.Tk()
root.title("Copy Listed Images")

source_dir = tk.StringVar()
dest_dir = tk.StringVar()
name_file_path = tk.StringVar()

tk.Label(root, text="Image directory:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
tk.Entry(root, textvariable=source_dir, width=50).grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="Select", command=lambda: select_directory(source_dir)).grid(row=0, column=2, padx=10, pady=5)

tk.Label(root, text="Destination directory:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
tk.Entry(root, textvariable=dest_dir, width=50).grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="Select", command=lambda: select_directory(dest_dir)).grid(row=1, column=2, padx=10, pady=5)

tk.Label(root, text="File containing image names:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
tk.Entry(root, textvariable=name_file_path, width=50).grid(row=2, column=1, padx=10, pady=5)
tk.Button(root, text="Select", command=lambda: select_file(name_file_path)).grid(row=2, column=2, padx=10, pady=5)

tk.Button(root, text="Execute", command=execute_copy).grid(row=3, column=1, pady=20)

root.mainloop()
