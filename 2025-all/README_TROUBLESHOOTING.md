# Troubleshooting Guide

## Excel "Update Links" Warning

**Problem:** After generating OJS files, Excel shows "This workbook contains links to other data sources" warning when opening.

**Root Cause:** The VBA project in the template file contains references to external libraries or workbooks that can't be removed by Python.

**Solution:**

### Clean the Template (Recommended)
1. Open `tournament_template.xlsm` in Excel
2. Press `Alt+F11` to open VBA Editor
3. Go to **Tools → References**
4. **Uncheck** any references marked as MISSING or pointing to external files
5. Keep only these standard references:
   - Visual Basic For Applications
   - Microsoft Excel Object Library
   - OLE Automation
   - Microsoft Office Object Library
6. Click OK and save the template
7. Re-run `build-tournament-folders.py`

### Fix Generated OJS Files
If files are already generated:
1. Open the OJS file
2. When prompted about links, click **"Don't Update"**
3. Press `Alt+F11`
4. Go to **Tools → References**
5. Uncheck any MISSING references
6. Close VBA Editor
7. Save the file

The warning will not appear again.
