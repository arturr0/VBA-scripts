# Excel-Word Automation Scripts

![VBA](https://img.shields.io/badge/VBA-Excel/Word-blue)
![Office Automation](https://img.shields.io/badge/Office-Automation-green)

A collection of 11 VBA macros for automating data transfer and formatting between Microsoft Excel and Word.

## Script Overview

### 1. Date-Based Formatting (`date.xlsx`)
- Changes cell font color to red on May 5, 2024
- Usage: Highlight important dates automatically

### 2. Excel to Word with Commas (`from many columns Excel write coma in word.xlsx`)
- Transfers Excel data to Word with comma-separated values
- Processes entire used range
- Preserves row structure with line breaks

### 3. Word to Excel Transformation (`many words to excel.xlsx`)
- Imports comma-separated Word content to Excel
- Splits text at commas into columns
- Maintains paragraph structure in rows

### 4. Range to Word Transfer (`Range to Word.xlsx`)
- Copies specific Excel range to Word
- Preserves table formatting
- Supports custom range selection

### 5. Non-Sequential Word to Excel (`row not sequence.xlsx`)
- Imports Word content to Excel columns
- Right-aligns all data
- Splits at commas into separate columns

### 6. Table to Word (`Table to Word.xlsx`)
- Exports Excel tables to Word with formatting
- Uses `PasteExcelTable` for perfect table transfer
- Polish variable naming (mojDok = myDoc)

### 7. Word to Excel Rows (`to excel row.xlsx`)
- Transforms Word paragraphs to Excel rows
- Splits comma-separated values across columns
- Handles variable-length data

### 8. Comma-Separated Word to Excel Columns (`Word coma to columns Excel.xlsx`)
- Specifically for two-column output
- Requires exactly two comma-separated values per line
- Strict formatting for paired data

### 9. Line-by-Line Word to Excel (`Word Line to column Excel.xlsx`)
- Basic Word-to-Excel transfer
- Each paragraph becomes a row in column A
- Simple vertical arrangement

### 10-11. Multi-Column Comma Export (`write coma many columns Excel.xlsx` and `write coma Word.xlsx`)
- Two similar implementations of multi-column export
- Create comma-separated strings from Excel data
- Transfer to Word with line breaks between rows

## Common Features

- **Error Handling**: All scripts include basic error handling
- **Object Management**: Proper object cleanup with `Nothing`
- **Flexible Ranges**: Most work with used ranges or customizable areas
- **Visibility Control**: Applications can run visibly or invisibly

## Installation

1. Open Excel file containing the macro
2. Enable macros when prompted (requires medium security level)
3. Access macros via Developer tab or assigned buttons

## Usage Instructions

1. **For Excel-to-Word scripts**:
   - Set correct file paths in the code
   - Define your source range (or use UsedRange)
   - Run macro to transfer data

2. **For Word-to-Excel scripts**:
   - Update the Word document path
   - Configure column/row placement as needed
   - Run macro to import data

3. **Special-purpose scripts**:
   - Date-based formatting requires no configuration
   - Table transfer needs table range definition

## Requirements

- Microsoft Office 2010 or later
- VBA enabled
- References to Microsoft Word Object Library (for Word interaction)
- Trust access to the VBA project object model

## Customization Tips

1. **Path Modification**: Change `C:\VB\...` to your actual file paths
2. **Range Adjustment**: Modify `UsedRange` or specific ranges like `A1:C10`
3. **Formatting**: Adjust font colors, alignment in the VBA code
4. **Delimiters**: Change comma (`,`) to other separators if needed

## Troubleshooting

**Common Issues**:
1. "Object library not found" - Enable Word Object Library reference
2. File not found - Verify all paths are correct
3. Data misalignment - Check delimiter consistency in source files

**Solutions**:
1. In VBA Editor: Tools > References > Check "Microsoft Word XX.X Object Library"
2. Use full absolute paths (avoid relative paths)
3. Clean source data before processing

## Security Notes

- Macros will only run with explicit permission
- No external data connections beyond specified files
- All objects are properly released after use

## License

Free to use and modify. Please credit original author when redistributing.

## Support

For assistance, please provide:
1. Excel/Word versions
2. Exact error message
3. Sample of problematic data
