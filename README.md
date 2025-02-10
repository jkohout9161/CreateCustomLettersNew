# CreateCustomLettersNew
# UiPath Automation: Excel to Parameterized PDF

## Overview
This UiPath automation processes an Excel file containing patient data and generates parameterized PDF files using the data. Each patient's details are extracted from the Excel file and used to create a customized PDF.

## Prerequisites
Before running the automation, ensure the following:
- **UiPath Studio** and **UiPath Robot** are installed.
- Required dependencies and packages are installed (e.g., `UiPath.Excel.Activities`, `UiPath.PDF.Activities`).
- Input **Excel file** containing patient data is available.
- A **template** for generating PDFs is prepared.

## Input File Format
The input Excel file should have the following structure:

```plaintext
| Patient ID | Name     | Age | Gender | Diagnosis | Doctor     | Date       |
|------------|---------|-----|--------|-----------|------------|------------|
| 001        | John Doe | 45  | Male   | Flu       | Dr. Smith  | 01/10/2024 |
```

## Workflow Steps
1. **Read Excel File:** The automation reads the patient data from the specified Excel file.
2. **Loop Through Records:** Iterates through each row to extract patient details.
3. **Generate PDF:** Uses a predefined template to create a parameterized PDF for each patient.
4. **Save and Store PDFs:** Saves the generated PDFs in the specified output directory.

## Configuration
Modify the following variables in the UiPath workflow as needed:

- `InputExcelPath`: Path to the input Excel file.
- `OutputFolderPath`: Path to the folder where generated PDFs will be stored.
- `TemplateFilePath`: Path to the PDF template file (if applicable).

## Execution
To run the automation:

1. Open **UiPath Studio**.
2. Load the project.
3. Update the required paths in the configuration.
4. Run the workflow.

## Error Handling
- If the **input file is missing** or has incorrect formatting, an error message will be logged.
- If any **row contains missing or incorrect data**, it will be skipped, and a warning will be logged.
- Any **unexpected errors** will be captured and logged in the error log file.

## Output
- A set of **PDF files**, each containing patient-specific details, stored in the designated output directory.

## Contact
For any issues or support, please contact the development team.

