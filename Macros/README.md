# SplitTextToColumns Macro

## Description
This macro splits text in cells from column A into individual characters across columns. Each character of the text will be placed in its own cell to the right of the original cell.

## Primary Purpose
This macro was primarily created for validating Electronic Funds Transfer (EFT) files exported from various banking systems. By splitting text into individual characters, it enables detailed validation of each character position against specific banking format requirements.

### EFT File Validation Applications
- **Australian EFT Requirements**: Validate position-specific characters in Australian banking file formats
- **NACHA Compliance**: Ensure exported files conform to National Automated Clearing House Association requirements
- **ANZ/ANZ Formats**: Verify character positions match ANZ banking specifications
- **Fixed-width Record Analysis**: Easily inspect numeric and text fields in fixed-width financial record formats

## Use Case Examples
- Breaking down product codes or serial numbers into individual components
- Parsing fixed-width text data
- Creating character-by-character analysis of text data
- Validating that each character in an EFT file meets position-specific requirements
- Identifying incorrect character types in banking file formats

## Parameters
- The macro processes cells from A1:A10 by default (can be modified in the code)

## How to Import
1. Open your Excel workbook
2. Press Alt + F11 to open the VBA Editor
3. Right-click on your project in the Project Explorer
4. Select Import File
5. Navigate to this .bas file and click Open

## How to Run
1. Press Alt + F8 to open the Macros dialog
2. Select "SplitTextToColumns"
3. Click Run

## Example
If cell A1 contains "EXCEL", after running the macro:
- B1 will contain "E"
- C1 will contain "X"
- D1 will contain "C"
- E1 will contain "E"
- F1 will contain "L"

## EFT Validation Example
If cell A1 contains a NACHA record line like "101021000019..." (Record Type, Priority Code, etc.):
- B1 will contain "1" (Record Type Code)
- C1 will contain "0" (Priority Code 1st digit)
- D1 will contain "1" (Priority Code 2nd digit)
- E1, F1, etc. will contain subsequent characters

This makes it easy to verify that each position contains the expected value type (numeric, alphabetic, special character) according to the banking format specification.