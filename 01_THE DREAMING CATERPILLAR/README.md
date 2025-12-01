# 🎓 Simply Growth - Worksheet Generator v10.0

**Excel-Based Curriculum Loader with Automatic Folder Structure**

---

## 📋 Quick Start

1. **Install dependencies**:
   ```bash
   pip install pandas openpyxl reportlab openai backoff
   ```

2. **Set your OpenAI API key**:
   ```bash
   export OPENAI_API_KEY='your-api-key-here'
   ```

3. **Edit the Excel file** (`details.xlsx`) with your curriculum

4. **Generate worksheets**:
   ```bash
   python work_sheets.py
   ```

---

## 📚 Documentation

This package includes comprehensive documentation:

### 📖 [README_EXCEL_FORMAT.md](README_EXCEL_FORMAT.md)
Complete guide to the Excel file format, usage instructions, and troubleshooting.

### 📖 [EXCEL_GUIDE.md](EXCEL_GUIDE.md)
Step-by-step instructions for creating your `details.xlsx` file with visual examples.

### 📖 [FOLDER_STRUCTURE.md](FOLDER_STRUCTURE.md)
Detailed explanation of the output folder structure and naming conventions.

### 📖 [CHANGES_SUMMARY.md](CHANGES_SUMMARY.md)
Summary of changes from v9.1 to v10.0 and migration guide.

---

## 🎯 Key Features

✅ **Excel-Based Curriculum**: No code changes needed - just edit Excel file  
✅ **Automatic Folder Structure**: Creates professional, timestamped folders  
✅ **Curriculum Alignment**: All questions aligned to your grade level and standards  
✅ **Three Activity Types**: MCQ, True/False, and Short Answer worksheets  
✅ **Answer Sheets**: Separate answer sheets with explanations  
✅ **Preview PDFs**: Combined preview of all three types  
✅ **Flexible Filtering**: Generate specific standards or topics  

---

## 📊 Excel File Format (Quick Reference)

### Structure:
```
Row 1: Subject Name - | [Your Subject]
Row 2: Grade level -  | [Your Grade]
Row 3: Curriculum -   | [Your Curriculum]
Row 4: [Unit Title - merged across columns]
Row 5: NO | STANDARD | TITLE | NOTE
Row 6+: Data rows for each topic
Empty Row: Separator between units
```

### Sample `details.xlsx` Included
A sample Excel file is included. You can:
- Open it to see the format
- Replace with your curriculum
- Save and run!

---

## 🚀 Usage Examples

### Generate All Worksheets
```bash
python work_sheets.py
```

### List Available Standards
```bash
python work_sheets.py --list
```
**Output**:
```
Curriculum: NGSS
Subject: Physical Sciences
Grade: Grade 7

Available Standards:
  1: MS-PS1: Matter and Its Interactions (3 subtopics)
     1.1: MS-PS1-1 - Developing Models of Matter
     1.2: MS-PS1-2 - Pure Substances and Mixtures
     1.3: MS-PS1-3 - Chemical Reactions
  ...
```

### Generate Specific Standards
```bash
# Only standard 1
python work_sheets.py --standards 1

# Standards 1, 3, and 5
python work_sheets.py --standards 1,3,5

# Standards 1 through 4
python work_sheets.py --standards 1-4
```

### Generate Specific Topics
```bash
# Topics 1-3 from all standards
python work_sheets.py --subs 1-3

# Topics 1 and 3 from standard 2
python work_sheets.py --standards 2 --subs 1,3
```

### Preview Without Generating
```bash
python work_sheets.py --dry-run
```

### Use Different Excel File
```bash
python work_sheets.py --excel my_curriculum.xlsx
```

---

## 📁 Output Structure

```
Simply Growth - WORK SHEET (3 ACTIVITIES) - 2025-10-20 03-45 PM/
└── Grade 7 - NGSS - Physical Sciences/
    └── 01. MS-PS1 - Matter and Its Interactions/
        └── 01. MS-PS1-1 - Developing Models of Matter/
            ├── 01. Underline Answer Type/
            │   ├── [Topic] – MCQ Worksheet.pdf
            │   └── [Topic] – MCQ Answer Sheet.pdf
            ├── 02. True-False/
            │   ├── [Topic] – True-False Worksheet.pdf
            │   └── [Topic] – True-False Answer Sheet.pdf
            ├── 03. Short Answer/
            │   ├── [Topic] – Short Answer Worksheet.pdf
            │   └── [Topic] – Short Answer Answer Sheet.pdf
            └── 04. Preview PDFs/
                └── [Topic] – Preview.pdf
```

**Each topic generates 7 PDF files** (3 worksheets + 3 answer sheets + 1 preview)

---

## 🎨 Folder Naming Templates

### Root Folder
`Simply Growth - WORK SHEET (3 ACTIVITIES) - [Date & Time]`

### Main Folder
`[Grade] - [Curriculum] - [Subject]`  
Example: `Grade 7 - NGSS - Physical Sciences`

### Unit Folders
`[No]. [Unit Title]`  
Example: `01. MS-PS1 - Matter and Its Interactions`

### Topic Folders
`[No]. [Standard] - [Title]`  
Example: `01. MS-PS1-1 - Developing Models of Matter`

---

## 🔧 Requirements

### Python Packages
```bash
pip install pandas openpyxl reportlab openai backoff
```

### Font Files (Included)
- Fredoka-VariableFont_wdth,wght.ttf
- DejaVuSans.ttf
- NotoSans fonts

### OpenAI API Key
Required for question generation. Set as environment variable:
```bash
export OPENAI_API_KEY='sk-...'
```

---

## 📝 Workflow

1. **Prepare Excel File**
   - Open `details.xlsx`
   - Add your subject, grade, curriculum (rows 1-3)
   - Add units and topics
   - Save

2. **Test Loading**
   ```bash
   python work_sheets.py --list
   ```

3. **Preview Generation**
   ```bash
   python work_sheets.py --dry-run
   ```

4. **Generate One Topic** (test)
   ```bash
   python work_sheets.py --standards 1 --subs 1
   ```

5. **Generate All**
   ```bash
   python work_sheets.py
   ```

---

## ❓ Troubleshooting

### "Could not find details.xlsx"
- Ensure file is in same directory as `work_sheets.py`
- Check filename spelling

### "No units found in Excel file"
- Verify header rows (NO, STANDARD, TITLE, NOTE)
- Check merged cells for unit titles
- Ensure empty rows between units

### "OPENAI_API_KEY not set"
```bash
export OPENAI_API_KEY='your-key-here'
```

### Questions don't match grade level
- Check Row 2 (Grade level) in Excel
- Verify format: "Grade 7" or "Grades 9-12"

---

## 📖 What's Included

| File | Description |
|------|-------------|
| `work_sheets.py` | Main Python script (v10.0) |
| `details.xlsx` | Sample curriculum Excel file |
| `README_EXCEL_FORMAT.md` | Comprehensive format guide |
| `EXCEL_GUIDE.md` | Step-by-step Excel creation guide |
| `FOLDER_STRUCTURE.md` | Output folder structure explanation |
| `CHANGES_SUMMARY.md` | Version changes and migration guide |
| `README.md` | This file |

---

## 🎓 Example Curricula

### Elementary School
```
Grade 3 - Common Core - Mathematics
└── 3.OA - Operations and Algebraic Thinking
    └── 3.OA.1 - Multiplication as Groups
```

### Middle School
```
Grade 7 - NGSS - Physical Sciences
└── MS-PS1 - Matter and Its Interactions
    └── MS-PS1-1 - Developing Models of Matter
```

### High School
```
Grades 9-10 - NGSS - Biology
└── HS-LS1 - From Molecules to Organisms
    └── HS-LS1-1 - Structure and Function
```

---

## 💡 Tips

✅ **Start Small**: Test with one standard first  
✅ **Use Dry-Run**: Preview before generating  
✅ **Detailed Notes**: Better notes = better questions  
✅ **Save Often**: Excel autosave is your friend  
✅ **Organize Units**: Use clear, consistent naming  

---

## 🆘 Need Help?

1. Check the documentation files
2. Review the sample `details.xlsx`
3. Run `--list` to verify Excel loading
4. Use `--dry-run` to preview output

---

## 🎉 You're Ready!

1. ✅ Edit `details.xlsx` with your curriculum
2. ✅ Set your OpenAI API key
3. ✅ Run `python work_sheets.py --list` to verify
4. ✅ Run `python work_sheets.py` to generate

**Happy worksheet creating!** 📚✨

---

## Version History

- **v10.0** (2025-10-20): Excel-based curriculum loader, automatic folder structure
- **v9.1**: Fixed short-answer underscores
- **v9.0**: Initial three-activity worksheet generator

---

Made with ❤️ for educators everywhere
