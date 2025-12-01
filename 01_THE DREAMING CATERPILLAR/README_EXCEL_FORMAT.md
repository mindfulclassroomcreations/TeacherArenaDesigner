# Worksheet PDF Builder - Excel-Based Curriculum Loader

## Overview
This updated version (v10.0) loads curriculum data from an Excel file (`details.xlsx`) instead of hardcoded curriculum data. It creates a comprehensive folder structure with worksheets aligned to your specific grade level and curriculum standards.

## Excel File Format

### Required Structure

Your `details.xlsx` file should have the following structure:

#### Metadata Section (First 3 rows)
- **Row 1**: Subject Name
  - Column A: `Subject Name -`
  - Column B: Your subject (e.g., `Physical Sciences`)

- **Row 2**: Grade Level
  - Column A: `Grade level -`
  - Column B: Your grade (e.g., `Grade 7` or `Grades 9-12`)

- **Row 3**: Curriculum Name
  - Column A: `Curriculum -`
  - Column B: Your curriculum (e.g., `NGSS`, `Common Core`, `NHES`)

#### Unit Sections (Row 4 onwards)

Each unit follows this pattern:

1. **Unit Title Row** (merged across columns A-D)
   - Example: `MS-PS1: Matter and Its Interactions`

2. **Header Row**
   - Column A: `NO`
   - Column B: `STANDARD`
   - Column C: `TITLE`
   - Column D: `NOTE`

3. **Data Rows** (one per topic)
   - Column A: Sequential number (1, 2, 3, ...)
   - Column B: Standard code (e.g., `MS-PS1-1`)
   - Column C: Topic title (e.g., `Developing Models of Matter`)
   - Column D: Teacher note/description

4. **Empty Row** (separator between units)

5. Repeat for next unit...

### Example Excel Layout

```
| A               | B                    | C                              | D                                           |
|-----------------|----------------------|--------------------------------|---------------------------------------------|
| Subject Name -  | Physical Sciences    |                                |                                             |
| Grade level -   | Grade 7              |                                |                                             |
| Curriculum -    | NGSS                 |                                |                                             |
| MS-PS1: Matter and Its Interactions (merged across A-D)                                                      |
| NO              | STANDARD             | TITLE                          | NOTE                                        |
| 1               | MS-PS1-1             | Developing Models of Matter    | Students develop models to describe...      |
| 2               | MS-PS1-2             | Pure Substances and Mixtures   | Analyze and interpret data on properties... |
| 3               | MS-PS1-3             | Chemical Reactions             | Gather and make sense of information...     |
|                 |                      |                                |                                             |
| MS-PS2: Motion and Stability (merged across A-D)                                                             |
| NO              | STANDARD             | TITLE                          | NOTE                                        |
| 1               | MS-PS2-1             | Newton's Third Law             | Apply Newton's Third Law to design...       |
| 2               | MS-PS2-2             | Electric and Magnetic Forces   | Plan an investigation to provide...         |
```

## Output Folder Structure

When you run the code, it creates the following structure:

```
Simply Growth - WORK SHEET (3 ACTIVITIES) - 2025-10-20 03-45 PM/
└── Grade 7 - NGSS - Physical Sciences/
    ├── 01. MS-PS1 - Matter and Its Interactions/
    │   ├── 01. MS-PS1-1 - Developing Models of Matter/
    │   │   ├── 01. Underline Answer Type/
    │   │   │   ├── [Topic] – MCQ Worksheet.pdf
    │   │   │   └── [Topic] – MCQ Answer Sheet.pdf
    │   │   ├── 02. True-False/
    │   │   │   ├── [Topic] – True-False Worksheet.pdf
    │   │   │   └── [Topic] – True-False Answer Sheet.pdf
    │   │   ├── 03. Short Answer/
    │   │   │   ├── [Topic] – Short Answer Worksheet.pdf
    │   │   │   └── [Topic] – Short Answer Answer Sheet.pdf
    │   │   └── 04. Preview PDFs/
    │   │       └── [Topic] – Preview.pdf
    │   ├── 02. MS-PS1-2 - Pure Substances and Mixtures/
    │   │   └── ... (same structure)
    │   └── ...
    └── 02. MS-PS2 - Motion and Stability/
        └── ... (same structure)
```

## Usage

### Basic Usage (Generate all topics)
```bash
python work_sheets.py
```

### List Available Standards
```bash
python work_sheets.py --list
```

### Generate Specific Standards
```bash
# Generate only standard 1
python work_sheets.py --standards 1

# Generate standards 1, 3, 4, and 5
python work_sheets.py --standards 1,3-5
```

### Generate Specific Subtopics
```bash
# Generate subtopics 1, 2, and 3 from all standards
python work_sheets.py --subs 1-3

# Generate subtopics 1 and 3 from standard 2
python work_sheets.py --standards 2 --subs 1,3
```

### Dry Run (Preview without generating)
```bash
python work_sheets.py --dry-run
```

### Custom Excel File
```bash
python work_sheets.py --excel my_curriculum.xlsx
```

## Key Features

1. **Curriculum Alignment**: All questions are aligned to your specific grade level and curriculum standards
2. **Automatic Folder Creation**: Creates timestamped folders with proper naming conventions
3. **Three Question Types**: 
   - Multiple Choice Questions (MCQ)
   - True/False
   - Short Answer
4. **Answer Sheets**: Separate answer sheets with explanations
5. **Preview PDFs**: Combined preview of all three question types
6. **Flexible Filtering**: Generate only what you need

## Requirements

Install required packages:
```bash
pip install pandas openpyxl reportlab openai backoff
```

Set your OpenAI API key:
```bash
export OPENAI_API_KEY='your-api-key-here'
```

## Folder Naming Templates

- **Root Folder**: `Simply Growth - WORK SHEET (3 ACTIVITIES) - [Date & Time]`
- **Main Folder**: `[Grade] - [Curriculum] - [Subject]`
- **Unit Folder**: `[No]. [Unit Title]`
- **Topic Folder**: `[No]. [Standard] - [Title]`

Example: `01. MS-PS1-1 - Developing Models of Matter`

## Tips

1. **Merged Cells**: Make sure unit title rows are properly merged across columns A-D
2. **Empty Rows**: Always include an empty row between units to separate them
3. **Standard Codes**: Include standard codes in Column B for proper labeling
4. **Teacher Notes**: Use Column D to provide context for question generation
5. **Sequential Numbering**: Use sequential numbers in Column A (NO) for each unit

## Troubleshooting

**Error: "Could not find details.xlsx"**
- Make sure the Excel file is in the same directory as `work_sheets.py`

**Error: "No units found in Excel file"**
- Check that you have header rows (NO, STANDARD, TITLE, NOTE)
- Verify empty rows separate units
- Ensure unit titles are in merged cells

**Questions not aligned to grade level**
- Check that Row 2 (Grade level) is correctly filled
- Verify the grade format (e.g., "Grade 7" or "Grades 9-12")

## Sample Excel File

A sample `details.xlsx` file has been created in this directory. You can:
1. Open it in Excel
2. Replace the sample data with your curriculum
3. Maintain the same structure
4. Run the script

Enjoy generating your curriculum-aligned worksheets! 🎉
