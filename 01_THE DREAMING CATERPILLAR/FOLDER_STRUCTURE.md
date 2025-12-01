# Output Folder Structure

## Complete Folder Hierarchy

When you run the code, here's exactly what will be created:

```
Simply Growth - WORK SHEET (3 ACTIVITIES) - 2025-10-20 03-45 PM/
│
└── Grade 7 - NGSS - Physical Sciences/
    │
    ├── 01. MS-PS1 - Matter and Its Interactions/
    │   │
    │   ├── 01. MS-PS1-1 - Developing Models of Matter/
    │   │   ├── 01. Underline Answer Type/
    │   │   │   ├── MS-PS1-1 - Developing Models of Matter – MCQ Worksheet.pdf
    │   │   │   └── MS-PS1-1 - Developing Models of Matter – MCQ Answer Sheet.pdf
    │   │   │
    │   │   ├── 02. True-False/
    │   │   │   ├── MS-PS1-1 - Developing Models of Matter – True-False Worksheet.pdf
    │   │   │   └── MS-PS1-1 - Developing Models of Matter – True-False Answer Sheet.pdf
    │   │   │
    │   │   ├── 03. Short Answer/
    │   │   │   ├── MS-PS1-1 - Developing Models of Matter – Short Answer Worksheet.pdf
    │   │   │   └── MS-PS1-1 - Developing Models of Matter – Short Answer Answer Sheet.pdf
    │   │   │
    │   │   └── 04. Preview PDFs/
    │   │       └── MS-PS1-1 - Developing Models of Matter – Preview.pdf
    │   │
    │   ├── 02. MS-PS1-2 - Pure Substances and Mixtures/
    │   │   ├── 01. Underline Answer Type/
    │   │   ├── 02. True-False/
    │   │   ├── 03. Short Answer/
    │   │   └── 04. Preview PDFs/
    │   │
    │   └── 03. MS-PS1-3 - Chemical Reactions/
    │       ├── 01. Underline Answer Type/
    │       ├── 02. True-False/
    │       ├── 03. Short Answer/
    │       └── 04. Preview PDFs/
    │
    └── 02. MS-PS2 - Motion and Stability - Forces and Interactions/
        │
        ├── 01. MS-PS2-1 - Newtons Third Law/
        │   ├── 01. Underline Answer Type/
        │   ├── 02. True-False/
        │   ├── 03. Short Answer/
        │   └── 04. Preview PDFs/
        │
        └── 02. MS-PS2-2 - Electric and Magnetic Forces/
            ├── 01. Underline Answer Type/
            ├── 02. True-False/
            ├── 03. Short Answer/
            └── 04. Preview PDFs/
```

## Folder Naming Templates

### Level 1: Root Folder
**Template**: `Simply Growth - WORK SHEET (3 ACTIVITIES) - [Date & Time]`

**Example**: `Simply Growth - WORK SHEET (3 ACTIVITIES) - 2025-10-20 03-45 PM`

**Components**:
- Fixed prefix: `Simply Growth - WORK SHEET (3 ACTIVITIES)`
- Date format: `YYYY-MM-DD`
- Time format: `HH-MM AM/PM`

### Level 2: Main Curriculum Folder
**Template**: `[Grade] - [Curriculum] - [Subject]`

**Example**: `Grade 7 - NGSS - Physical Sciences`

**Components**:
- Grade from Excel Row 2
- Curriculum from Excel Row 3
- Subject from Excel Row 1

### Level 3: Unit Folders
**Template**: `[No]. [Unit Title]`

**Example**: `01. MS-PS1 - Matter and Its Interactions`

**Components**:
- 2-digit number (01, 02, 03...)
- Full unit title from Excel (merged row)

### Level 4: Topic Folders
**Template**: `[No]. [Standard] - [Title]`

**Example**: `01. MS-PS1-1 - Developing Models of Matter`

**Components**:
- 2-digit number (01, 02, 03...)
- Standard code from Column B
- Topic title from Column C

### Level 5: Activity Type Folders
**Fixed Names**:
1. `01. Underline Answer Type`
2. `02. True-False`
3. `03. Short Answer`
4. `04. Preview PDFs`

## File Naming

### MCQ Files
- Worksheet: `[Topic] – MCQ Worksheet.pdf`
- Answer Sheet: `[Topic] – MCQ Answer Sheet.pdf`

### True-False Files
- Worksheet: `[Topic] – True-False Worksheet.pdf`
- Answer Sheet: `[Topic] – True-False Answer Sheet.pdf`

### Short Answer Files
- Worksheet: `[Topic] – Short Answer Worksheet.pdf`
- Answer Sheet: `[Topic] – Short Answer Answer Sheet.pdf`

### Preview File
- Combined: `[Topic] – Preview.pdf`

## Real-World Examples

### Example 1: Elementary Math
```
Simply Growth - WORK SHEET (3 ACTIVITIES) - 2025-10-20 02-30 PM/
└── Grade 3 - Common Core - Mathematics/
    └── 01. 3.OA - Operations and Algebraic Thinking/
        └── 01. 3.OA.1 - Multiplication as Groups/
            ├── 01. Underline Answer Type/
            ├── 02. True-False/
            ├── 03. Short Answer/
            └── 04. Preview PDFs/
```

### Example 2: High School Biology
```
Simply Growth - WORK SHEET (3 ACTIVITIES) - 2025-10-20 04-15 PM/
└── Grades 9-10 - NGSS - Biology/
    └── 01. HS-LS1 - From Molecules to Organisms/
        └── 01. HS-LS1-1 - Structure and Function/
            ├── 01. Underline Answer Type/
            ├── 02. True-False/
            ├── 03. Short Answer/
            └── 04. Preview PDFs/
```

### Example 3: Middle School Health
```
Simply Growth - WORK SHEET (3 ACTIVITIES) - 2025-10-20 09-00 AM/
└── Grade 8 - NHES - Health Education/
    └── 01. Standard 1 - Core Concepts/
        └── 01. NHES.1.1 - Body Systems/
            ├── 01. Underline Answer Type/
            ├── 02. True-False/
            ├── 03. Short Answer/
            └── 04. Preview PDFs/
```

## Total Files Per Topic

Each topic generates **7 PDF files**:
1. MCQ Worksheet
2. MCQ Answer Sheet
3. True-False Worksheet
4. True-False Answer Sheet
5. Short Answer Worksheet
6. Short Answer Answer Sheet
7. Combined Preview

**Example**: If you have 5 topics, you'll get 5 × 7 = **35 PDF files**

## Folder Count

For the sample curriculum (2 units, 5 topics):
- 1 Root folder
- 1 Main curriculum folder
- 2 Unit folders
- 5 Topic folders
- 20 Activity type folders (5 topics × 4 types)
- **Total: 29 folders + 35 PDF files**

## Benefits of This Structure

✅ **Clear Organization**: Easy to navigate  
✅ **Professional Naming**: Consistent and descriptive  
✅ **Curriculum-Aligned**: Folders match standards  
✅ **Easy Distribution**: Share entire units or individual topics  
✅ **Timestamp**: Never overwrite previous generations  
✅ **Scalable**: Works for any curriculum size  

## Navigation Tips

### To find a specific topic:
1. Open the dated root folder
2. Navigate to your grade/curriculum/subject folder
3. Find the unit folder (by number or name)
4. Open the topic folder
5. Choose the activity type you need

### To share:
- **Share everything**: Zip the root folder
- **Share one unit**: Zip a unit folder
- **Share one topic**: Zip a topic folder
- **Share one type**: Copy files from activity type folder

Enjoy your well-organized worksheets! 📁
