# Attendance and Overtime Cell Coloring Tool
Project Description -
This tool is designed to help you quickly identify discrepancies between your employee attendance records and their reported overtime. It's especially useful for checking if employees who were marked as absent or on holiday (HN) still reported overtime hours for that same day.
The tool takes two Excel files as input:
 * Your Attendance Excel File, which contains employee IDs and their daily attendance status.
 * Your Overtime Excel File, which contains employee IDs and their reported overtime hours for various dates.
After comparing the data, it generates a new Excel file based on your Overtime file. In this new file, specific cells will be colored:
 * Red: If an employee was marked as Absent ("A") or Holiday ("HN") in the Attendance file for a specific date, but they still reported more than 0 hours of overtime for that same date in the Overtime file. This highlights potential mistakes or inconsistencies.
 * Green: If an employee was marked as Absent ("A") or Holiday ("HN") in the Attendance file for a specific date, and they reported 0 hours of overtime for that same date in the Overtime file. This confirms consistency for absent employees.
Features
 * Easy File Selection: Uses simple pop-up windows to help you select your Excel files.
 * Automated Comparison: Compares attendance and overtime data automatically.
 * Visual Highlighting: Clearly marks inconsistent cells with colors (Red for issues, Green for consistency).
 * New Output File: Creates a new Excel file with the colored results, leaving your original files untouched.
 * Supports Standard Excel Files: Works with .xlsx and .xls formats.
 * User-Friendly Messages: Provides clear messages during the process and explains any errors.
Prerequisites (Only if running the Python script directly)
If you plan to use the executable .exe file (recommended for non-coders), you do NOT need to install Python or these libraries. These prerequisites are only for those who want to run the Python script directly.
 * How to check if Python is installed:
   * Open your Command Prompt (on Windows, search for "cmd") or Terminal (on Mac/Linux).
   * Type python --version and press Enter.
   * If you see a version number (e.g., Python 3.9.7), Python is installed. If you get an error, you need to install it.
 * How to install Python (if needed):
   * Go to the official Python website: https://www.python.org/downloads/
   * Download the latest stable version of Python 3 (e.g., Python 3.10.x or newer).
   * Run the installer. IMPORTANT: During installation, make sure to check the box that says "Add Python to PATH" (or similar) on the first screen. This is crucial for the tool to work easily.
Installation (Only if running the Python script directly)
Once Python is installed, you need to install a few additional libraries that the tool uses.
 * Open your Command Prompt (Windows) or Terminal (Mac/Linux).
   * Windows: Search for "cmd" in the Start menu.
   * Mac/Linux: Open your Terminal application.
 * Type the following command and press Enter:
   pip install pandas openpyxl

   You should see messages indicating that the packages are being downloaded and installed. Once it's done, you're ready to use the tool!
How to Use the Executable (.exe) File (For Windows Users)
If you have received an executable .exe file of this tool (e.g., attendance_ot_tool.exe), you can run it directly without needing to install Python or any libraries. This is the easiest way for non-coders to use the tool.
 * Download the Executable:
   * Download the .exe file (e.g., attendance_ot_tool.exe) to a folder on your computer. You can create a new folder specifically for this tool.
 * Prepare Your Excel Files:
   * Make sure your "Attendance Excel File" and "Overtime Excel File" are ready.
   * Crucially, ensure your Excel files match the expected format described in the "Expected Excel File Formats" section below. Incorrect formats will cause errors.
   * It's a good idea to place your Excel files in the same folder as the .exe file, or at least remember their exact location.
 * Run the Tool:
   * Navigate to the folder where you saved the .exe file.
   * Double-click on the executable file (e.g., attendance_ot_tool.exe). A black window (Command Prompt) will briefly appear and then pop-up windows will guide you.
 * Select Your Files (Pop-up Windows):
   * A small pop-up window titled "Select Attendance Excel File" will appear. Use this window to navigate to and select your Attendance Excel File. Click "Open".
   * Another pop-up window titled "Select Overtime Excel File" will then appear. Use this to select your Overtime Excel File. Click "Open".
 * Wait for Processing:
   * The black command prompt window will show messages like "Processing attendance data..." and "Processing overtime data...". This means the tool is working. Please be patient, especially with large files.
 * View Results:
   * Once the process is complete, a final pop-up message box will appear, saying "Success!" and providing a summary of how many cells were colored red and green.
   * A new Excel file will be created in the same folder as your "Overtime Excel File". Its name will be similar to your Overtime file, but with _colored added to the end (e.g., if your Overtime file was OvertimeReport.xlsx, the new file will be OvertimeReport_colored.xlsx).
   * Open this new _colored.xlsx file to see the highlighted cells.
 * Exit the Tool:
   * The command prompt window will automatically close after you click "OK" on the success message box.
How to Use the Python Script Directly (For Developers/Advanced Users)
If you prefer to run the Python script directly (or if you are on Mac/Linux), follow these steps:
 * Download the Tool:
   * Download the Python script file (e.g., attendance_ot_tool.py) to a folder on your computer.
 * Prepare Your Excel Files:
   * Make sure your "Attendance Excel File" and "Overtime Excel File" are ready.
   * Crucially, ensure your Excel files match the expected format described in the "Expected Excel File Formats" section below. Incorrect formats will cause errors.
   * It's a good idea to place your Excel files in the same folder as the Python script, or at least remember their exact location.
 * Run the Tool:
   * Open your Command Prompt (Windows) or Terminal (Mac/Linux).
   * Navigate to the folder where you saved the Python script using the cd command. For example, if your script is in C:\Users\YourName\Documents\ExcelTool, you would type:
     cd C:\Users\YourName\Documents\ExcelTool

     and press Enter.
   * Once you are in the correct folder, type the following command and press Enter:
     python attendance_ot_tool.py

     (Replace attendance_ot_tool.py with the actual name of your script if it's different).
 * Select Your Files (Pop-up Windows):
   * A small pop-up window titled "Select Attendance Excel File" will appear. Use this window to navigate to and select your Attendance Excel File. Click "Open".
   * Another pop-up window titled "Select Overtime Excel File" will then appear. Use this to select your Overtime Excel File. Click "Open".
 * Wait for Processing:
   * The black command prompt/terminal window will show messages like "Processing attendance data..." and "Processing overtime data...". This means the tool is working. Please be patient, especially with large files.
 * View Results:
   * Once the process is complete, a final pop-up message box will appear, saying "Success!" and providing a summary of how many cells were colored red and green.
   * A new Excel file will be created in the same folder as your "Overtime Excel File". Its name will be similar to your Overtime file, but with _colored added to the end (e.g., if your Overtime file was OvertimeReport.xlsx, the new file will be OvertimeReport_colored.xlsx).
   * Open this new _colored.xlsx file to see the highlighted cells.
 * Exit the Tool:
   * In the black command prompt/terminal window, you will see "Press Enter to exit...". Simply press the Enter key on your keyboard to close the window.
Expected Excel File Formats
It is critical that your input Excel files follow these structures for the tool to work correctly.
1. Attendance Excel File
This file should contain employee IDs and their attendance status for various dates.
 * Column A (or similar, but the content is most important): Should contain a column with Employee IDs. The tool looks for a column that eventually becomes the "Employee ID" for comparison. Ensure these IDs are unique and consistent with your Overtime file.
 * Columns from 11th column onwards (L, M, N, etc.): These columns are expected to represent dates. The header of these columns should be recognizable as dates (e.g., 1/1/2024, 01-Jan-2024, 2024-01-01).
 * Cell Values for Dates: In the date columns, for each employee, the values should indicate attendance. The tool specifically looks for:
   * Empty cells (blank): Treated as absent.
   * "A" (case-insensitive): Treated as absent.
   * "HN" (case-insensitive): Treated as absent.
   * Any other value (e.g., "P" for present, or a time) will be considered present.
Example Attendance File Structure:
| Employee ID | Name | Department | ... (up to 10 columns) ... | 1/1/2024 | 1/2/2024 | 1/3/2024 | ... |
|---|---|---|---|---|---|---|---|
| 1001 | John | Sales |  | P | A | P |  |
| 1002 | Jane | Marketing |  | P | P | HN |  |
| 1003 | Mike | HR |  | A | P | P |  |
2. Overtime Excel File
This file should contain employee IDs and their reported overtime hours for various dates.
 * Column A (or similar): Must be named exactly Emp_Id (case-sensitive) and contain Employee IDs. These IDs must match the Employee IDs in your Attendance file.
 * Other Columns: The headers of these columns should be recognizable as dates (e.g., 1/1/2024, 01-Jan-2024, 2024-01-01).
 * Cell Values for Dates: These cells should contain the overtime hours reported for that employee on that specific date. Values should be numbers (e.g., 0, 2.5, 8).
Example Overtime File Structure:
| Emp_Id | 1/1/2024 | 1/2/2024 | 1/3/2024 | ... |
|---|---|---|---|---|
| 1001 | 0 | 2.0 | 0 |  |
| 1002 | 0 | 0 | 1.5 |  |
| 1003 | 3.0 | 0 | 0 |  |
Troubleshooting
 * "No file selected. Exiting...": You closed the file selection window without choosing a file. Run the script/executable again.
 * "File not found...": The file path you provided (or selected) is incorrect, or the file has been moved/deleted. Double-check the path and file existence.
 * "The selected file appears to be empty or corrupted.": The Excel file you selected is empty or cannot be read by the tool. Try opening it manually to ensure it's not corrupted.
 * "Required column not found: 'Employee ID'" or "Required column not found: 'Emp_Id'": Your Excel file headers do not match the expected names. Crucially, ensure the Overtime file has a column named Emp_Id (case-sensitive). For the attendance file, ensure there's a column that acts as the employee ID.
 * No colors appear in the output file:
   * Ensure your input files match the "Expected Excel File Formats" exactly.
   * If running the Python script directly, check if Python and the libraries were installed correctly (re-run the pip install command).
   * There might be no discrepancies based on the "Red" and "Green" rules. Check the summary message box for "Red cells" and "Green cells" counts.
 * The script closes immediately after double-clicking (Windows, for Python script): This usually means Python is not correctly added to your system's PATH. Use the .exe file if available, or use Option B (running from Command Prompt) or reinstall Python making sure to check "Add Python to PATH".
If you encounter persistent issues, double-check the "Expected Excel File Formats" section very carefully, as most problems stem from incorrect file structures.
License
This project is open-source and available under the MIT License. You are free to use, modify, and distribute this tool as needed.
