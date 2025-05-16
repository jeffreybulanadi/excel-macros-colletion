# Excel Macros Collection

This repository contains a collection of useful Excel VBA macros for automation and productivity enhancement.

## Available Macros

### SplitTextToColumns
Splits text in cells from column A into individual characters across columns. This is useful for processing text data that needs to be broken down character by character.

#### Primary Use Case: EFT File Validation
This macro was primarily created for validating the characters and values in Electronic Funds Transfer (EFT) files exported from various systems. It helps ensure compliance with banking standards such as:
- Australian EFT requirements
- NACHA (National Automated Clearing House Association) formats
- ANZ/ANZ banking formats

By splitting text into individual characters, financial professionals can easily validate that each position in an EFT file contains the correct character type and value according to the required specifications.

## How to Use

1. **Enable Developer Tab**:
   - In Excel, go to File > Options > Customize Ribbon
   - Check "Developer" in the right panel and click OK

2. **Import a Macro**:
   - Open Excel and press Alt + F11 to open the VBA editor
   - Right-click on your workbook in the Project Explorer
   - Select Import File and choose the .vba or .bas file

3. **Run a Macro**:
   - In Excel, go to the Developer tab
   - Click on Macros
   - Select the macro you want to run and click Run

## Sample Files

The `Samples` directory contains example Excel files that demonstrate how each macro works.

## Future Plans

This repository will be updated with additional macros to enhance Excel productivity.

## Author

Jeffrey Bulanadi

## Last Updated

May 16, 2025
