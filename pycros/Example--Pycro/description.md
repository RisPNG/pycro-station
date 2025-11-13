> Batch update Excel A1 to "test complete" (red on black)

This Pycro lets you select multiple Excel workbooks and processes them in one go. For each selected file, it writes "test complete" to cell A1 on the active sheet, sets the text color to red, and fills the cell background black. A live log shows progress and any errors as it runs.

Notes
- Supports multi-select via the Select Files button.
- Uses openpyxl; install via the Requirements button if needed.
- Overwrites only cell A1 on the active sheet and saves the workbook in place.