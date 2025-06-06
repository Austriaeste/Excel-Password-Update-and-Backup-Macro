# Excel Password Update and Backup Macro

This VBA macro automates password updates and backups for an Excel file. It matches unit names from a text file with a column in an Excel sheet, updates passwords in another column if they differ from a provided password file, and creates backups with a date-based folder structure.

## Features

- **Password Update (`UpdatePassword`)**
  - Reads a unit name from `unit_list.txt`.
  - Searches for the unit name in column D (starting from row 2) of the `IP_List` sheet.
  - Updates the corresponding password in column G (row 20 or later) if it differs from `sample.txt`.
  - Skips empty cells in column G.
  - Logs and displays updated rows in a single `MsgBox` at the end.
- **Backup (`BackupExcel`)**
  - Copies the Excel file to a temporary folder (`temp_backup` in Downloads).
  - Moves it to a final backup folder (`backup/YYYYMMDD`).
  - Cleans up the temporary folder.
- **Form Controls**
  - Two buttons in the `IP_List` sheet: "Password Update" and "Backup".
- **Error Handling**
  - Checks for file/sheet existence, invalid rows, and empty cells.
  - Displays detailed error messages.

## Requirements

- Windows OS with Microsoft Excel (VBA-enabled).
- Excel file must be saved as `.xlsm` (macro-enabled).
- File paths must be accessible (local or network).

## Setup

1. **File Structure**
   - Place all files in the same folder (default: `C:\Users\<YourUser>\OneDrive\Desktop\password\`).
   - Required files:
     - `sample.xlsm`: Excel file with the `IP_List` sheet.
     - `unit_list.txt`: Single-line unit name (e.g., `Unit005`).
     - `sample.txt`: Single-line password (e.g., `xyz789`).

2. **Excel Configuration**
   - Open `sample.xlsm`.
   - Ensure the `IP_List` sheet exists.
   - Column D (from row 2): Unit names (e.g., `Unit001`).
   - Column G (from row 20): Passwords (e.g., `pass001`). Empty cells are skipped.
   - Enable the Developer tab:
     - File > Options > Customize Ribbon > Check "Developer".

3. **VBA Macro**
   - Open VBA Editor (Alt + F11).
   - Insert a new module and paste the code from `PasswordAndBackup.vba`.
   - Save `sample.xlsm` as a macro-enabled workbook (`.xlsm`).

4. **Form Controls**
   - In the `IP_List` sheet, add two buttons:
     - Developer > Insert > Form Controls > Button.
     - Button 1: Assign to `UpdatePassword`, label as "Password Update".
     - Button 2: Assign to `BackupExcel`, label as "Backup".
   - Save the workbook.

## Usage

1. **Password Update**
   - Click the "Password Update" button.
   - The macro:
     - Reads `unit_list.txt` (e.g., `Unit005`).
     - Finds `Unit005` in column D (row 2 onward).
     - Checks the corresponding row in column G (must be ≥ row 20).
     - If G differs from `sample.txt` (e.g., `xyz789`), updates it.
     - Shows a `MsgBox` with updated rows (e.g., `G6 (Unit005): pass005 -> xyz789`) or "No updates."
   - Errors (e.g., missing unit, empty G cell) are displayed.

2. **Backup**
   - Click the "Backup" button.
   - The macro:
     - Copies `sample.xlsm` to `C:\Users\<YourUser>\Downloads\temp_backup`.
     - Moves it to `C:\Users\<YourUser>\OneDrive\Desktop\password\backup\YYYYMMDD`.
     - Deletes the temporary folder.
     - Shows a `MsgBox` with the backup path.

## Test Data

- **sample.xlsm** (`IP_List` sheet):
  ```csv
  D,G
  ,,
  Unit001,pass001
  Unit002,
  Unit003,pass003
  Unit004,
  Unit005,pass005
  Unit006,pass006
  ```

- **unit_list.txt**:
  ```
  Unit005
  ```

- **sample.txt**:
  ```
  xyz789
  ```

- **Expected Output**:
  - Password Update (for `Unit005`):
    - Finds D6 (`Unit005`), checks G6 (`pass005`).
    - Updates G6 to `xyz789`.
    - `MsgBox`: `G6 (Unit005): pass005 -> xyz789`
  - Backup:
    - Saves to `backup\20250606\sample.xlsm`.

## Configuration

- **File Paths**
  - Edit `SOURCE_FOLDER_PATH` in the VBA code for your environment (e.g., `\\Server\Share\password\`).
  - Japanese paths (`デスクトップ`) may need to be `Desktop` in some systems.
- **Sheet Name**
  - Change `SHEET_NAME` if different from `IP_List`.
- **File Names**
  - Update `SOURCE_EXCEL_FILE_NAME`, `SOURCE_TEXT_FILE_NAME`, `UNIT_LIST_FILE_NAME` as needed.

## Notes

- **Single Unit**: `unit_list.txt` contains one unit name. For multiple units, modify the code to loop through lines.
- **Unique Units**: D column unit names are assumed unique. For duplicates, the first match is used.
- **Row Mapping**: D2 corresponds to G2 (adjust if D2->G20, etc.).
- **Empty Cells**: G column empty cells trigger an error. Modify to skip silently if needed.
- **Environment**: Windows only. Network path issues may require admin rights.

## Contributing

Fork, submit issues, or PRs for enhancements (e.g., multiple unit support, file logging).

## License

MIT License. See `LICENSE` file for details.
