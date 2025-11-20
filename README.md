# Cells - Documentation
## Overview

This is a powerful GUI-based Excel/CSV data editor built using Python's Tkinter library and Openpyxl for robust Excel (xlsx) file handling.

![Cells](https://github.com/zrnge/Cells/blob/main/Cells.png)

## Key Features

* **File Management:** Open/Save/Save As for both `.xlsx` and `.csv` files.
* **Multi-Sheet Support:** Seamlessly switch between sheets in a loaded Excel workbook.
* **Undo/Redo:** Full history tracking for all data modifications.
* **Data Manipulation:** Add/Delete/Move Rows and Columns.
* **Smart Paste:** Paste vertical or horizontal data from the clipboard, with a pre-paste dialog for selecting delimiters (Tab, Comma, Space, Newline) and insertion mode (Overwrite, Insert Before, Insert After, Append).
* **Sorting & Filtering:** Sort data by clicking column headers. Filter data using the search bar (supports keyword or `ColumnName:value1,value2` syntax).
* **Customization:** Dark theme and toggleable grid lines for visual clarity.

## Installation

1.  **Clone the Repository:**
    ```bash
    git clone https://github.com/zrnge/Cells.git
    cd Cells
    ```

2.  **Install Requirements:**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  **Run the application:**
    ```bash
    python cells.py
    ```
2.  **Open File (File > Open):** Select an `.xlsx` or `.csv` file.
3.  **Sheet Selector:** Use the **Sheet:** dropdown in the icon bar to navigate sheets (for `.xlsx` files).
4.  **Editing:** Double-click any cell to edit its value inline.
5.  **Context Menu (Right-Click):** Access advanced features like copy, move, delete, and the smart paste options.

---
