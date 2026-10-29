# VTU Marks Extractor

This project extracts student marks from VTU (Visvesvaraya Technological University) PDF result sheets and organizes them into a structured Excel file.

## Features

- Extracts the following information from each PDF:
  - Student Name
  - University Seat Number (USN)
  - Marks for each subject (Internal, External, and Total)
- Supports both command-line and GUI-based operation.
- Allows users to specify the subjects to be extracted via a simple text file.
- Creates a well-formatted Excel file with a two-row header for clarity.
- Handles PDF text extraction using either `pdfplumber` or `PyPDF2`.
- Logs the extraction process to `marks_extraction.log`.

## Project Structure

```
.
├── 6thsemresult/           # Folder for PDF result files
├── extracted_marks.xlsx    # Output Excel file
├── requirements.txt        # Python dependencies
├── Subjects.txt            # List of subject codes
├── VTU_marks_extracter.py  # Command-line script
└── VTU_marks_GUI.py        # GUI script
```

## Installation

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd <repository-folder>
    ```

2.  **Install the required dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
    You may also need to install `tkinter` if it's not already part of your Python installation:
    ```bash
    sudo apt-get install python3-tk  # For Debian/Ubuntu
    # Or use the appropriate package manager for your system
    ```

## Usage

There are two ways to use this tool:

### 1. GUI Version (Recommended)

The GUI version is the most user-friendly and flexible.

1.  **Run the script:**
    ```bash
    python VTU_marks_GUI.py
    ```

2.  **Select the Subjects File:**
    A file dialog will open. Choose the `Subjects.txt` file (or your own custom list) that contains the subject codes you want to extract, with one code per line.

3.  **Select the PDF Folder:**
    Another dialog will appear. Choose the folder that contains all the student PDF result files (e.g., `6thsemresult`).

4.  **Done!**
    The script will process the files and save the output as `extracted_marks.xlsx` in the parent directory of the PDF folder.

### 2. Command-Line Version

The command-line script is less flexible and uses hardcoded paths.

1.  **Prepare the files:**
    - Place all your PDF result files inside the `6thsemresult` folder.
    - Make sure the `Subjects.txt` file is present in the same directory as the script.

2.  **Run the script:**
    ```bash
    python VTU_marks_extracter.py
    ```

3.  **Done!**
    The script will create `extracted_marks.xlsx` in the project's root directory.

## `Subjects.txt` File

This file tells the script which subjects to look for.

- Each subject code should be on a new line.
- You can add comments using the `#` symbol.
- The codes are case-insensitive.

**Example `Subjects.txt`:**
```
# --- Core Subjects ---
BAD601
BDS602
BCD685

# --- Electives ---
BPEK658
BNSK658
```

## Output

The script generates an Excel file named `extracted_marks.xlsx` with the following structure:

|         |         |         | **BAD601** |            |         | **BDS602** |            |         | ... |
|:--------|:--------|:--------|:-----------|:-----------|:--------|:-----------|:-----------|:--------|:----|
|**SerialNo**|**Name**|**USN**| **I**      | **E**      | **T**   | **I**      | **E**      | **T**   | ... |
| 1       | Student A | 1XX...  | 20         | 75         | 95      | 18         | 60         | 78      | ... |
| 2       | Student B | 1XX...  | 19         | 80         | 99      | 17         | 55         | 72      | ... |

- **I**: Internal Marks
- **E**: External Marks
- **T**: Total Marks
