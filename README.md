# PPTX to PDF Automation

This repository contains a **PowerShell script** that automates the conversion of multiple PowerPoint (.pptx) files into PDF format. 

The script is designed to handle **bulk conversions** efficiently, minimizing PowerPoint during the process, showing **live progress**, and giving a **sound alert** when the task is completed.

## Features

- Bulk conversion of PPTX files to PDF
- Minimized PowerPoint window (avoids unnecessary pop-ups)
- Live progress counter showing number of files processed
- Sound alert when conversion is complete
- Automatic creation of output folder if missing
- Simple and easy-to-use PowerShell script

## Usage

1. Place your `.pptx` files in the `input` folder.
2. Update the folder paths in the script (`$inputFolder` and `$outputFolder`) if needed.
3. Open PowerShell (normal mode is fine) and run the script.
4. Wait for conversion to finish — you’ll see progress and hear a beep when done.
5. Converted PDFs are saved in the `output` folder.

## Requirements

- Windows OS
- Microsoft PowerPoint installed
- PowerShell (built-in)


---

This script is perfect for teachers, students, or anyone who works with multiple PowerPoint files and wants a **quick, hassle-free PDF export**.
