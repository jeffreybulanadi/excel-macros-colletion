# Sample Excel File Instructions

## SplitTextToColumns Sample

### Instructions to create the sample file:

1. Create a new Excel workbook (.xlsx)
2. Enter the following data in column A:

```
A1: EXCEL
A2: MACRO
A3: FUNCTION
A4: VBA
A5: CODE
A6: GITHUB
A7: REPOSITORY
A8: SPLIT
A9: TEXT
A10: COLUMNS
```

3. Save this file as `SplitTextToColumns_Sample.xlsx` in this Samples directory

### Expected Results:

After running the SplitTextToColumns macro on this sample data:

- Row 1: "E" "X" "C" "E" "L" (in cells B1 through F1)
- Row 2: "M" "A" "C" "R" "O" (in cells B2 through F2)
- etc.

### Testing the Macro:

1. Open the sample Excel file
2. Import the SplitTextToColumns.bas module from the Macros directory
3. Run the SplitTextToColumns macro
4. Observe how the text data is split into individual characters across columns