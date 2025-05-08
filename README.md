# README - Railway Accidents Report Analysis System

## Overview
This project implements an application designed to analyse railway accident reports from PDF Reports within the ERAIL database using OpenAI's GPT-4 Turbo model. It features:
- Railway-themed GUI
- Excel-based incident data loading and filtering 
- Country-based filtering system  
- Extraction of structured insights from corresponding PDF reports:
    - Accident Cause and Decision Responses

The application reads the source code written, processes it, and executes accordingly.

Developed as part of a dissertation project by **Sri Pradeeptaa Anantha**.

## Features Implemented 
- 🚆 Railway-themed GUI built with `tkinter`
- 📊 Load and filter accident data from Excel
- 📄 Automatically locate and extract text from accident report PDFs
- 🧠 Analyse reports using GPT-4-Turbo to extract:
  - Cause of the accident
  - Decision responses (communications and actions)
- 💾 Save analysis results as structured JSON
- 🔍 Filter data by country

## Requirements 
- **Python 3.x**: The application is built using Python 3. Make sure Python 3.13.x or later is installed on your machine.
- **OpenAI API key**: This is required for accessing GPT-4 Turbo; you must have a valid key from your OpenAI account, or you can use the preconfigured key already present in the application (if available).

## Running the Program

1. **Clone or Download the Repositry**:
  Download the project files from the repository or unzip the provided files.

2. **Install Python**:
  - Make sure that Python 3.13.x (or later) is installed on your system.
  - You can download Python from the official website: [python.org](https://www.python.org/downloads/). 

3. **Install Dependencies**:
  ```bash
  pip install openai pypdf pandas pillow
  ```

4. **Running the application**:
  - Ensure all required files are present (Excel file, PDFs, and logo image)
  - Open a terminal/command prompt and navigate to the project directory 
  - Run the 'python_accident_analyzer_UI.py' file by typing the following command:
    ```
    python python_accident_analyzer_UI.py
    ```
  - Alternatively, you can simply run `python_accident_analyzer_UI.py`.
  - In the GUI or within the source code file itself:
    - Open Excel file (File → Open Excel File) or set `self.excel_path` manually in the code
    - Set the PDF directory (File → Change PDF Directory) or update `self.pdf_directory` in the code
    - Enter your OpenAI API Key (File → Set API Key) or assign it directly to `self.api_key` in the script
  - Select an accident from the list that has a valid Final Report ID and an available PDF (preferably a UK-based incident).
  - Click "Analyse Report" to view insights

## Output
Analysis results are displayed in the UI and saved as `.json` in the `Analysis` subdirectory within your PDF directory.

---

**Note:** 
Make sure the PDF filenames include the "Final Report ID" from the Excel file for matching to work properly.

## License
For academic and research use only.
