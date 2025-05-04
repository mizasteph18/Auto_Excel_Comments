# Auto_Excel_Comments

# Script to Update Excel Files with Comments and Flags

This Python script is designed to update a main Excel file by adding comments and flags based on a mapping provided in a separate key Excel file. It's specifically built to handle main Excel files with multiple spreadsheets and a particular header structure.

## Functionality

The script performs the following actions:

1.  **Reads Configuration:** It takes several configuration parameters (explained below) to specify file paths, sheet names, header rows, and column names.
2.  **Loads Excel Files:** It opens both the main Excel file (the one to be updated) and the key Excel file (containing the comment and flag information).
3.  **Reads Key Data:** It reads the "code", "comment", and "flag" columns from a specific sheet (default: "AutoComment") in the key Excel file and stores this information in a Python dictionary for efficient lookups. The "code" column in the key file is used as the key in this dictionary.
4.  **Processes Main Excel Sheets:** It iterates through each sheet in the main Excel file.
5.  **Header Check:** For each sheet in the main Excel file, it checks if the header row, starting on the **4th row**, begins with a cell containing the word "comments" (case-insensitive).
6.  **Column Identification:** If the header condition is met, it attempts to identify the column indices for the key column (default: "code"), the comments column (default: "comments"), and the flag column (default: "flag") within that sheet's header row.
7.  **Data Update:** It then iterates through the rows below the header in the current sheet. For each row, it reads the value in the specified key column. If this value exists as a key in the dictionary created from the key file, it retrieves the corresponding "comment" and "flag" and writes them into the respective "comments" and "flag" columns of the current row in the main file.
8.  **Skips Unmatching Sheets:** If a sheet's header on the 4th row does not start with "comments", that sheet is skipped without processing.
9.  **Saves Updated File:** Finally, it saves the modified main Excel file to a specified output file path.

## Specificity and Configuration

This script is tailored with the following specific behaviors and requires careful configuration:

* **Main File Header Row:** The script **assumes that the header row in the sheets to be processed starts on the 4th row** of the main Excel file.
* **"Comments" Trigger:** A sheet in the main Excel file is only processed if the **first cell of the 4th row** contains the word "comments" (case-insensitive). This acts as an identifier for the sheets you want to update.
* **Adjustable Column Names (Main File):** You can configure the names of the key column (`main_key_column_name`), the comments column to be updated (`main_comments_column_name`), and the flag column to be updated (`main_flag_column_name`) in the main Excel file. The default values are "code", "comments", and "flag", respectively.
* **Key File Sheet:** The script specifically looks for a sheet named **"AutoComment"** in the key Excel file to retrieve the mapping data. This sheet name is configurable.
* **Key File Columns:** The script expects the key Excel sheet ("AutoComment") to have columns named **"code"**, **"comment"**, and **"flag"**. These column names are also configurable.
* **Output File:** The updated main Excel data is saved to a new file specified by the `output_file_path`. You can set this to the same as the input `main_file_path` if you want to overwrite the original (use with caution!).

## How to Use

1.  **Save the Script:** Save the Python script (e.g., as `excel_updater.py`).
2.  **Configure Parameters:** Open the script in a text editor and carefully adjust the configuration variables in the `if __name__ == "__main__":` block to match your specific file paths, sheet names, and column headers. Ensure the paths to your main and key Excel files are correct.
3.  **Run the Script:** Open a terminal or command prompt, navigate to the directory where you saved the script, and run it using the command: `python excel_updater.py`
4.  **Check Output:** After the script finishes, the updated Excel file will be saved to the specified output path.

## Important Notes

* Ensure that the key column values in your main Excel file exactly match the "code" values in the "AutoComment" sheet of your key Excel file for the comments and flags to be applied correctly.
* The script is case-sensitive for column header names (except for the initial "comments" check on row 4). Make sure the configured column names exactly match the headers in your Excel files.
* The script requires the `openpyxl` library to be installed. If you haven't already, you can install it using pip: `pip install openpyxl`.
* It's always recommended to back up your original Excel files before running any script that modifies them.
