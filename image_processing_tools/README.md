# Image processing tools
# 1. Move/Copy Images Tool

## Overview
This tool allows you to copy or move images from one directory to another based on a list of filenames provided in a `.txt` file. It supports a variety of image formats and provides a simple graphical interface.

## Features
- Select source and destination directories.
- Use a `.txt` file with image names to determine which images to process.
- Options to either **copy** or **move** images.
- Automatically handles duplicate filenames by creating unique names if necessary.
- Supports image formats: `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`.

## Requirements
- `tkinter` (usually included in Python)
- `shutil` and `os` (standard Python libraries)
