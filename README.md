# Word Document Template Filler

A reusable Python script that automatically fills Word document templates with user-provided data by replacing placeholders.

## Features

- **Interactive Setup**: Prompts for template file path and output directory (no code editing required)
- **Placeholder Replacement**: Replaces placeholders in format `<<PLACEHOLDER_NAME>>` throughout documents
- **Employee Data Collection**: Collects name, login, device ID, and password information
- **Text Formatting**: Formats replaced text as bold with 14pt font size
- **Table Support**: Works with placeholders in both paragraphs and table cells
- **Path Validation**: Checks that template file exists before processing
- **Input Validation**: Ensures all required fields are provided
- **Safe File Naming**: Handles special characters in employee names
- **Error Handling**: Clear error messages with helpful suggestions

## Required Python Library

Install the required library:

```bash
pip install python-docx
```

## Template Format

Your Word document template should contain placeholders:
- `<<EMPLOYEE_NAME>>`
- `<<COMPUTER_NAME>>`
- `<<COMPUTER_LOGIN>>`
- `<<PASSWORD>>`

Or edit the code to create your own custom template variables for your word doc.

## Usage

1. Run the script: `python script_name.py`
2. Enter the full path to your Word template file
3. Enter the output directory path (will be created if it doesn't exist)
4. Provide employee information when prompted
5. The filled document will be saved as `[Employee_Name]_first_day_form.docx`

## Example

```
=== Word Document Template Filler ===

Template File Setup:
Enter the full path to your Word template file: C:\templates\employee_form.docx

Output Directory Setup:
Enter the output directory path: C:\output

Employee Information:
Enter Employee Name (first + last): John Smith
Enter Computer Login (Firstname.Lastname): John.Smith
Enter Device ID: COMP-001
Enter Password: TempPass123

Processing template...
Document filled and saved to C:\output\John_Smith_first_day_form.docx

Success! Your document has been created.
```

## Notes

- The script is fully portable - no hardcoded paths to modify
- Output directory will be created automatically if it doesn't exist
- Employee names with spaces or special characters are safely converted for filenames
