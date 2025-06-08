# Password Update Macro (UpdatePassword)

## Overview
This VBA macro updates passwords in the `IP_List` sheet of an Excel file (`sample.xlsm`) using a list of unit names (`unit_list.txt`) and a password (`sample.txt`). It searches for unit names in column A (starting from row 2) and updates the corresponding password in column G (row 20 or later).

## Features
- Reads all unit names (e.g., `Kagoshima-DR`, `Hokkaido-DS`) from `unit_list.txt`.
- Reads a single password from `sample.txt`.
- Searches for each unit name in column A of the `IP_List` sheet and updates the corresponding G column if needed.
- Logs updates and errors, displaying them in a message box.
- Saves the Excel file if updates are made.

## File Structure
The following files must be placed in `C:\Users\austr\OneDrive\Desktop\password\`:
- **`sample.xlsm`**: The target Excel file.
  - Sheet name: `IP_List`
  - Column A (row 2 and beyond): Unit names (e.g., `Kagoshima-DR`).
  - Column G (row 20 and beyond): Passwords.
- **`unit_list.txt`**: List of unit names (one per line, e.g., `Kagoshima-DR\nHokkaido-DS\nTokyo-DT`).
- **`sample.txt`**: Password (single line, e.g., `P@ssw0rd123`).

## Prerequisites
- Microsoft Excel installed with macros enabled.
- Read/write permissions for the folder `C:\Users\austr\OneDrive\Desktop\password\`.
- `unit_list.txt` and `sample.txt` in text format (UTF-8 or ANSI encoding).
- Unit names in column A and passwords in column G are treated as strings.
- Updates in column G are restricted to row 20 or later.

## Execution Steps
1. Place `sample.xlsm`, `unit_list.txt`, and `sample.txt` in the specified folder.
2. Open `sample.xlsm` and enable macros.
3. Run the `UpdatePassword` macro from the VBA editor.
4. A start message appears, followed by a result message box showing updates or errors.

## Output
- **Success**: Displays updated rows (e.g., `Column G, row 25 (Kagoshima-DR): oldPass → P@ssw0rd123`). The Excel file is saved.
- **No Updates**: Displays "No updates were made."
- **Errors**: Logs errors (e.g., `Error: Unit name 'Tokyo-DT' not found in column A (row 2 and beyond).`) in the message box.

## Error Handling
Handles the following errors:
- Missing files (`sample.xlsm`, `unit_list.txt`, `sample.txt`).
- Missing `IP_List` sheet.
- Unit name not found in column A.
- Unit name found in a row before 20.
- Empty cell in column G.
- Other errors (displays error number and description).

## Notes
- Empty lines in `unit_list.txt` are skipped.
- `sample.txt` supports only a single password (first line used if multiple).
- Backup functionality (`temp_backup`, `backup`) is defined but not implemented.
- Execution date/time (e.g., 2025/06/08 22:31:00 JST) is not included in logs but can be added.

## Customization
- To support multiple passwords or output logs to a file, modify the macro.
- To add execution date/time to logs, insert:
  ```vba
  updateLog = updateLog & "Update Time: " & Format(Now, "yyyy/mm/dd hh:nn:ss") & vbCrLf
  ```
- To search a different column or sheet, update column numbers (A: 1, G: 7) or `SHEET_NAME`.

## Test Example
**Input**:
- `unit_list.txt`:
  ```
  Kagoshima-DR
  Hokkaido-DS
  Tokyo-DT
  Osaka-DU
  Kyoto-DV
  Fukuoka-DW
  ```
- `sample.txt`: `P@ssw0rd123`

**Output Example**:
```
Results:
Column G, row 25 (Kagoshima-DR): oldPass → P@ssw0rd123
Error: Unit name 'Tokyo-DT' not found in column A (row 2 and beyond).
Column G, row 30 (Kyoto-DV): anotherPass → P@ssw0rd123
```
