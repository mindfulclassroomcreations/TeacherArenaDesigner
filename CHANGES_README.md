# Worksheet Generator - Changes Summary

## Overview
The worksheet generator has been updated to load curriculum data from an Excel file (`details.xlsx`) instead of hardcoding it in the source code. This makes it easy to create worksheets for any curriculum without modifying the Python code.

## Key Changes

### 1. **Excel-Based Curriculum Loading**
- Added `CurriculumData` class to store curriculum metadata
- Added `load_curriculum_from_excel()` function to read `details.xlsx`
- The function can handle multiple curricula in a single Excel file
- Supports flexible Excel structures with clear section markers

### 2. **Excel File Structure**
The `details.xlsx` file should follow this format:

```
Row 1:  Subject Name -     | [Subject Name]
Row 2:  Grade level -      | [Grade Level]
Row 3:  Curriculum -       | [Curriculum Name]
Row 4+: [Units and Topics]
```

Each unit section has:
- **Merged row**: Unit title (e.g., "MS-PS1: Matter and Its Interactions")
- **Header row**: NO | STANDARD | TITLE | NOTE
- **Data rows**: Topic details with curriculum codes
- **Empty row**: Separator between units

### 3. **New Folder Structure**
The generator now creates folders with this hierarchy:

```
Academy Ready - 03 WORKSHEETS + 30 TASK CARDS - DD Month YYYY - h:mm am/pm/
└── [Grade] - [Curriculum] - [Subject]/
    └── [##. Standard Code - Title]/
        └── [##. Standard Code — Topic Title]/
            ├── 01. True or False Questions Worksheet/
            ├── 02. Multiple Choice Questions Worksheet/
            ├── 03. Short Answer Type Questions Worksheet/
            ├── 04. Task Cards/
            └── PREVIEW PDFs (Do not Upload This)/
```

**Example:**
```
Academy Ready - 03 WORKSHEETS + 30 TASK CARDS - 19 October 2025 - 5:00 pm/
└── Grade 07 - NGSS - Physical Science/
    └── 01. MS-PS1 - Matter and Its Interactions/
        └── 01. MS-PS1-1 — Atomic Models and Structure/
            ├── 01. True or False Questions Worksheet/
            ├── 02. Multiple Choice Questions Worksheet/
            ├── 03. Short Answer Type Questions Worksheet/
            ├── 04. Task Cards/
            └── PREVIEW PDFs (Do not Upload This)/
```

### 4. **Curriculum Context for AI**
- All OpenAI prompts now include curriculum context (Subject, Grade, Curriculum)
- Instructions to AI ensure questions align with specific grade levels and curriculum standards
- Questions are tailored to the exact curriculum specified in the Excel file

### 5. **Multiple Curriculum Support**
- The system can process multiple curricula in one run
- Each curriculum creates its own separate folder structure
- All curricula are organized under the main date-stamped folder

## How to Use

### Step 1: Prepare Your Excel File
Create a file named `details.xlsx` in the same folder as `work_sheets.py` with your curriculum data.

**Column Structure:**
- **Column A (NO)**: Topic number (1, 2, 3, etc.)
- **Column B (STANDARD)**: Curriculum code (e.g., "MS-PS1-1")
- **Column C (TITLE)**: Topic title/description
- **Column D (NOTE)**: Teacher notes or guidance

### Step 2: Set Up Environment
```bash
# Make sure you have required packages
pip install pandas openpyxl reportlab openai backoff pypdf

# Set your OpenAI API key
export OPENAI_API_KEY="your-api-key-here"
```

### Step 3: Run the Generator
```bash
cd "/Users/sankalpa/Desktop/TPT/02_PYTHON CODES/11_Academy Ready/01_WORKSHEETS"
python3 work_sheets.py
```

### Step 4: Find Your Generated Files
Look for a folder named `Academy Ready - 03 WORKSHEETS + 30 TASK CARDS - [DD Month YYYY - h:mm am/pm]` containing all your generated worksheets and task cards.

## Benefits

1. **No Code Changes**: Update curriculum by editing Excel, not Python
2. **Curriculum Alignment**: All questions automatically align with specified standards
3. **Grade-Appropriate**: Questions match the specified grade level
4. **Multiple Curricula**: Process multiple curricula in one run
5. **Clear Organization**: Folder structure clearly shows grade, curriculum, and subject
6. **Date Stamping**: Each run creates a dated folder for version control

## Important Notes

- **Excel File Location**: Must be named `details.xlsx` and placed in the same folder as `work_sheets.py`
- **Standard Codes**: Include curriculum codes (like "MS-PS1-1") in the STANDARD column for proper organization
- **Title Formatting**: Use " — " or " – " to separate standard codes from titles in unit names
- **Empty Rows**: Use empty rows to separate different units in your Excel file
- **Multiple Curricula**: Start each new curriculum section with "Subject Name -" in the first column

## Example Excel Layout

```
| A               | B        | C                              | D                                  |
|-----------------|----------|--------------------------------|------------------------------------|
| Subject Name -  | Physical Science                |                                    |
| Grade level -   | Grade 07                        |                                    |
| Curriculum -    | NGSS                            |                                    |
|                 |          |                                |                                    |
| MS-PS1: Matter and Its Interactions                          |                                    |
| NO              | STANDARD | TITLE                          | NOTE                               |
| 1               | MS-PS1-1 | Atomic Models and Structure    | Explore atomic composition...      |
| 2               | MS-PS1-1 | Periodic Table Patterns        | Understand organization...         |
|                 |          |                                |                                    |
| MS-PS2: Motion and Stability                                 |                                    |
| NO              | STANDARD | TITLE                          | NOTE                               |
| 1               | MS-PS2-1 | Newton's Laws in Action        | Explore Newton's three laws...     |
```

## Troubleshooting

If you encounter errors:

1. **File Not Found**: Ensure `details.xlsx` is in the correct location
2. **Empty Output**: Check that your Excel file has the correct structure
3. **API Errors**: Verify your OPENAI_API_KEY is set correctly
4. **PDF Errors**: Install missing packages: `pip install pypdf` or `pip install PyPDF2`

## Questions or Issues?

Check the console output for detailed error messages and progress updates. The script provides verbose feedback about what it's processing and where files are being saved.
