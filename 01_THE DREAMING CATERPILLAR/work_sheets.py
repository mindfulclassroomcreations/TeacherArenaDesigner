#!/usr/bin/env python3

"""

Worksheet PDF Builder – True/False • TF+Explanation • Short-Answer • Open-Ended • Scenario-Based  (v11.0)

───────────────────────────────────────────────────────────────

Change from v10.0 → v11.0

─────────────────────

• Replaced "01. Underline Answer Type" (MCQ) with "1. True, False Type Questions" (25 per PDF)
• Replaced "02. True-False" with "2. True-False Type Questions with Explanation" (20 per PDF)
• Renamed "03. Short Answer" to "3. Short Answer Type Questions" (25 per PDF)
• Added "4. Open-Ended Questions" (20 per PDF)
• Added "5. Scenario-Based Questions" (10 per PDF)
• Preview now combines all 5 sections; preview folder renamed to "06. Preview PDFs"
"""

import os, re, json, random, pathlib, sys, time, argparse
from datetime import datetime
from typing import List, Tuple, Dict

from reportlab.lib.pagesizes import letter
from reportlab.platypus      import (SimpleDocTemplate, BaseDocTemplate, Paragraph, Spacer,
                                    KeepTogether, PageBreak, PageTemplate, Frame, NextPageTemplate,
                                    Table, TableStyle)
from reportlab.lib.styles    import getSampleStyleSheet, ParagraphStyle
from reportlab.lib           import colors
from reportlab.pdfbase       import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from xml.sax.saxutils import escape as xml_escape
from reportlab.lib.utils    import ImageReader
import openai, backoff
import pandas as pd
import io
import shutil
from reportlab.pdfgen import canvas as pdfcanvas
# Optional libraries for generating JPG thumbnails from PDFs
try:
    import fitz  # type: ignore  # PyMuPDF (optional)
    HAS_FITZ = True
except Exception:
    HAS_FITZ = False
try:
    from pdf2image import convert_from_path  # type: ignore  # requires poppler (optional)
    HAS_PDF2IMG = True
except Exception:
    HAS_PDF2IMG = False


# ───────────────────────  CONFIG  ────────────────────────────────
openai.api_key = openai.api_key or os.getenv("OPENAI_API_KEY")
OFFLINE_MODE = False

def ensure_api_key():
    """Exit with a helpful message if the API key is missing.

    Deferred so that non‑generation commands like --list or --dry-run still work
    without requiring a key.
    """
    global OFFLINE_MODE
    if not openai.api_key:
        OFFLINE_MODE = True
        print("⚠️  OPENAI_API_KEY not set. Running in OFFLINE mode with placeholder content.")
MODEL = "gpt-4o-mini"

FONT_PATHS = [
   "Fredoka-VariableFont_wdth.ttf",
   r"C:\Users\DELL\Desktop\2. BIT\0. Fonts\Fredoka-Regular.ttf",
   "DejaVuSans.ttf"
]
def register_font():
    pass
def _try_register_font(candidates):
   search_dirs = [
       pathlib.Path('.'),
       pathlib.Path(__file__).parent,
       pathlib.Path(__file__).parent / 'fonts',
       pathlib.Path.home() / 'Library' / 'Fonts',           # macOS user
       pathlib.Path('/Library/Fonts'),                      # macOS system
       pathlib.Path('/System/Library/Fonts'),               # macOS system
       pathlib.Path('C:/Windows/Fonts'),                    # Windows
   ]
   for name in candidates:
       for d in search_dirs:
           p = d / name
           try:
               if p.exists():
                   fam = pathlib.Path(name).stem
                   pdfmetrics.registerFont(TTFont(fam, str(p)))
                   return fam
           except Exception:
               continue
   # Fallback to built-in
   return 'Helvetica'

# Requested fonts
FONT_LATO_REG   = _try_register_font(["Lato-Regular.ttf", "Lato-Regular.ttf"])  # duplicate for iteration
FONT_LATO_LIGHT = _try_register_font(["Lato-Light.ttf"]) 
FONT_RALEWAY_SB = _try_register_font(["Raleway-SemiBold.ttf", "Raleway-SemiBold.ttf"]) 
FONT_DEJAVU     = _try_register_font(["DejaVuSans.ttf", "DejaVu Sans.ttf"]) 

# Previous default for items we should not change (e.g., underscore lines, page numbers)
def _register_default_font():
   candidates = [
       "Fredoka-VariableFont_wdth.ttf",
       r"C:\\Users\\DELL\\Desktop\\2. BIT\\0. Fonts\\Fredoka-Regular.ttf",
       "DejaVuSans.ttf"
   ]
   return _try_register_font(candidates)

FONT = _register_default_font()

# Special character set for DejaVuSans switching
SPECIAL_CHARS = set([
    # Arrows
    '→','←','↑','↓','↔','⇒','⇐','↕','⇔',
    # Math symbols
    '±','≈','≠','<','>','≤','≥','×','÷','∑','∏','∫','√','∞','∆','π','•','°',
    # Logic symbols
    'ℓ','ℵ','∀','∃','∅','∈','∉','⊆','⊇','⊂','⊃',
    # Superscripts
    '¹','²','³','⁴','⁵','⁶','⁷','⁸','⁹','⁰','ⁿ',
    # Subscripts
    '₀','₁','₂','₃','₄','₅','₆','₇','₈','₉',
])

def wrap_special(text: str, base_font: str) -> str:
   """Return text with base_font applied, but any special characters wrapped with DejaVuSans.
   Keeps underscore-only lines unchanged.
   Escapes XML special chars in the content.
   """
   if not isinstance(text, str):
       text = str(text)
   # Support explicit <sub>...</sub> and <super>...</super> tags by converting them
   # to Unicode subscript/superscript characters before further processing.
   def _convert_sub_super(s: str) -> str:
       # replace <sub>...</sub>
       def _sub(m):
           inner = m.group(1)
           try:
               return inner.translate(SUB_MAP)
           except Exception:
               return inner
       # replace <super>...</super>
       SUP_DIGITS = { '0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹' }
       def _sup(m):
           inner = m.group(1)
           out = []
           for ch in inner:
               if ch.isdigit():
                   out.append(SUP_DIGITS.get(ch, ch))
               elif ch in SUP_MAP:
                   out.append(SUP_MAP[ch])
               else:
                   out.append(ch)
           return ''.join(out)
       s = re.sub(r'<sub>(.*?)</sub>', _sub, s, flags=re.IGNORECASE|re.DOTALL)
       s = re.sub(r'<super>(.*?)</super>', _sup, s, flags=re.IGNORECASE|re.DOTALL)
       return s
   text = _convert_sub_super(text)
   # Do not alter pure underscore lines
   s = text.strip()
   if s and set(s) == {'_'}:
       return text
   out = []
   for ch in text:
       if ch in SPECIAL_CHARS:
           out.append(f"<font name='{FONT_DEJAVU}'>{xml_escape(ch)}</font>")
       else:
           out.append(xml_escape(ch))
   inner = ''.join(out)
   return f"<font name='{base_font}'>{inner}</font>"
FONT = _register_default_font()

# Global curriculum data (loaded from Excel)
CURRICULUM_DATA = None


# ───────────────────  CURRICULUM LOADER  ─────────────────────────────────
class CurriculumData:
    """Stores curriculum metadata and topics from Excel"""
    def __init__(self, subject: str, grade: str, curriculum: str):
        self.subject_name = subject
        self.grade_level = grade
        self.curriculum_name = curriculum
        self.units: List[Tuple[int, str, List[Tuple[int, str, str]]]] = []
    
    def get_prompt_context(self) -> str:
        """Returns context string for OpenAI prompts"""
        return (
            f"Subject: {self.subject_name}\n"
            f"Grade Level: {self.grade_level}\n"
            f"Curriculum: {self.curriculum_name}\n"
            f"IMPORTANT: All questions MUST align with {self.grade_level} standards and stay within the {self.curriculum_name} curriculum. Do NOT create questions outside these standards."
        )

def load_curriculum_from_excel(excel_path: str = "details.xlsx") -> List[CurriculumData]:
    """
    Load multiple curricula from Excel file.
    
    Expected Excel structure (can repeat for multiple curricula):
    - Row X: Subject Name - [value]
    - Row X+1: Grade level - [value]
    - Row X+2: Curriculum - [value]
    - Row X+3+: Unit sections for this curriculum
    - ... (next curriculum starts with Subject Name - again)
    
    Each unit section has:
    - Merged row: Unit title (e.g., "MS-PS1: Matter and Its Interactions")
    - Header row: NO, STANDARD, TITLE, NOTE
    - Data rows: Topic details (Column A=NO, B=STANDARD, C=TITLE, D=NOTE)
    - Empty row: Separator between units
    
    Returns a list of CurriculumData objects (one per curriculum in the file).
    """
    try:
        # Read entire Excel file without headers to parse structure
        df_raw = pd.read_excel(excel_path, header=None)
        print(f"📊 Excel file loaded: {len(df_raw)} rows, {len(df_raw.columns)} columns\n")
        
        all_curricula = []
        
        # Find all curriculum section markers (rows starting with "Subject Name -")
        curriculum_start_rows = []
        for idx in range(len(df_raw)):
            first_col = str(df_raw.iloc[idx, 0]).strip() if pd.notna(df_raw.iloc[idx, 0]) else ''
            if first_col.lower().startswith('subject name'):
                curriculum_start_rows.append(idx)
        
        if not curriculum_start_rows:
            print("❌ No 'Subject Name -' markers found.")
            sys.exit(1)
        
        print(f"📚 Found {len(curriculum_start_rows)} curriculum section(s)\n")
        
        # Process each curriculum section
        for curr_idx, start_row in enumerate(curriculum_start_rows):
            # Determine end row (either next curriculum start or end of file)
            end_row = curriculum_start_rows[curr_idx + 1] if curr_idx + 1 < len(curriculum_start_rows) else len(df_raw)
            
            print(f"{'='*60}")
            print(f"Processing Curriculum #{curr_idx + 1} (rows {start_row + 1} to {end_row})")
            print(f"{'='*60}")
            
            # Extract metadata from first 3 rows of this section
            # Check if data is in column 0 (legacy format) or column 1 (new format with labels)
            first_cell = str(df_raw.iloc[start_row, 0]).strip() if pd.notna(df_raw.iloc[start_row, 0]) else ""
            
            if first_cell.lower().startswith('subject name'):
                # New format: Labels in column 0, data in column 1
                subject_name = str(df_raw.iloc[start_row, 1]).strip() if pd.notna(df_raw.iloc[start_row, 1]) else "Unknown Subject"
                grade_level = str(df_raw.iloc[start_row + 1, 1]).strip() if pd.notna(df_raw.iloc[start_row + 1, 1]) else "Unknown Grade"
                curriculum_name = str(df_raw.iloc[start_row + 2, 1]).strip() if pd.notna(df_raw.iloc[start_row + 2, 1]) else "Unknown Curriculum"
            else:
                # Legacy format: Data in column 0 with prefix
                subject_name = first_cell
                subject_name = re.sub(r'^Subject Name\s*-\s*', '', subject_name, flags=re.IGNORECASE).strip()
                
                grade_level = str(df_raw.iloc[start_row + 1, 0]).strip() if pd.notna(df_raw.iloc[start_row + 1, 0]) else "Unknown Grade"
                grade_level = re.sub(r'^Grade level\s*-\s*', '', grade_level, flags=re.IGNORECASE).strip()
                
                curriculum_name = str(df_raw.iloc[start_row + 2, 0]).strip() if pd.notna(df_raw.iloc[start_row + 2, 0]) else "Unknown Curriculum"
                curriculum_name = re.sub(r'^Curriculum\s*-\s*', '', curriculum_name, flags=re.IGNORECASE).strip()
            
            print(f"📚 Subject: {subject_name}")
            print(f"🎓 Grade Level: {grade_level}")
            print(f"📖 Curriculum: {curriculum_name}")
            
            # Create curriculum data object
            curriculum_data = CurriculumData(subject_name, grade_level, curriculum_name)
            
            # Parse units starting from row start_row + 3
            unit_index = 0
            current_unit_title = None
            current_subtopics = []
            in_data_section = False
            
            for idx in range(start_row + 3, end_row):
                row = df_raw.iloc[idx]
                
                # Get first column value
                first_col = str(row[0]).strip() if pd.notna(row[0]) else ''
                
                # Stop if we encounter another "Subject Name -" marker
                if first_col.lower().startswith('subject name'):
                    break
                
                # Check if this is a header row (contains NO, STANDARD, TITLE, NOTE)
                row_values = [str(val).strip().upper() for val in row.values if pd.notna(val)]
                is_header = any(h in row_values for h in ['NO', 'TITLE', 'STANDARD', 'NOTE'])
                
                if is_header:
                    # Found header row - next rows will be data
                    in_data_section = True
                    print(f"📋 Found header at row {idx + 1}")
                    continue
                
                # Check if row is empty (signals end of current unit)
                row_empty = all(pd.isna(val) or str(val).strip() == '' for val in row.values)
                
                if row_empty and in_data_section:
                    # End of current unit - save it
                    if current_unit_title and current_subtopics:
                        unit_index += 1
                        curriculum_data.units.append((unit_index, current_unit_title, current_subtopics))
                        print(f"✅ Unit {unit_index}: {current_unit_title} ({len(current_subtopics)} topics)")
                        current_subtopics = []
                    in_data_section = False
                    continue
                
                if not in_data_section:
                    # We're in the title section (merged rows) - this is a unit title
                    if first_col and len(first_col) > 3:
                        current_unit_title = first_col
                        print(f"📖 Unit Title: {current_unit_title}")
                else:
                    # We're in data section - parse topic data
                    # Expected columns: NO (A), STANDARD (B), TITLE (C), NOTE (D)
                    if len(row) >= 3:
                        no_val = row[0]
                        standard_val = row[1]
                        title_val = row[2]
                        note_val = row[3] if len(row) > 3 else ''
                        
                        # Convert NO to integer
                        try:
                            sub_idx = int(float(no_val)) if pd.notna(no_val) else len(current_subtopics) + 1
                        except (ValueError, TypeError):
                            sub_idx = len(current_subtopics) + 1
                        
                        # Get title
                        title = str(title_val).strip() if pd.notna(title_val) else ''
                        
                        # Get standard and prepend if available
                        if pd.notna(standard_val):
                            standard = str(standard_val).strip()
                            if standard and standard != '':
                                title = f"{standard} - {title}"
                        
                        # Get note
                        note = str(note_val).strip() if pd.notna(note_val) else ''
                        
                        # Add if valid
                        if title and title != 'nan' and len(title) > 2:
                            current_subtopics.append((sub_idx, title, note))

                        # If NOTE column contains 'MINI BUNDLE', this marks the end of the current unit
                        try:
                            note_lower = str(note).strip().lower()
                        except Exception:
                            note_lower = ''
                        if in_data_section and note_lower and 'mini bundle' in note_lower:
                            if current_unit_title and current_subtopics:
                                unit_index += 1
                                curriculum_data.units.append((unit_index, current_unit_title, current_subtopics))
                                print(f"✅ Unit {unit_index}: {current_unit_title} ({len(current_subtopics)} topics) [ended by MINI BUNDLE]")
                                current_subtopics = []
                                current_unit_title = None
                                in_data_section = False
                            continue
            
            # Don't forget the last unit if section doesn't end with empty row
            if current_unit_title and current_subtopics:
                unit_index += 1
                curriculum_data.units.append((unit_index, current_unit_title, current_subtopics))
                print(f"✅ Unit {unit_index}: {current_unit_title} ({len(current_subtopics)} topics)")
            
            if curriculum_data.units:
                total_topics = sum(len(topics) for _, _, topics in curriculum_data.units)
                print(f"✅ Curriculum loaded: {len(curriculum_data.units)} units, {total_topics} topics\n")
                all_curricula.append(curriculum_data)
            else:
                print(f"⚠️  Warning: No units found in curriculum section #{curr_idx + 1}\n")
        
        if not all_curricula:
            raise ValueError("No valid curriculum sections found in Excel file")
        
        print(f"{'='*60}")
        print(f"🎉 Successfully loaded {len(all_curricula)} curriculum(s)")
        print(f"{'='*60}\n")
        
        return all_curricula
    
    except FileNotFoundError:
        print(f"❌ Error: Could not find '{excel_path}'")
        print(f"   Please ensure the file exists in the same directory as this script.")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error loading curriculum from Excel: {e}")
        print(f"   Error type: {type(e).__name__}")
        import traceback
        traceback.print_exc()
        print("\n📝 Expected Excel structure:")
        print("   Row 1: Subject Name - [value] (or just the value)")
        print("   Row 2: Grade level - [value]")
        print("   Row 3: Curriculum - [value]")
        print("   Row 4+: Unit sections")
        print("   Each unit:")
        print("     - Merged row: Unit title")
        print("     - Header row: NO, STANDARD, TITLE, NOTE")
        print("     - Data rows: Topic details")
        print("     - Empty row: Separator")
        sys.exit(1)



# ─────────────────────  UTILITIES  ───────────────────────────────
SUB_MAP = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
SUP_MAP = {"+": "⁺", "-": "⁻"}
DIGIT_RUN  = re.compile(r'([A-Za-z\)])(\d+)')
ION_CHARGE = re.compile(r'([A-Za-z₀-₉]+)([+-])$')
STRIP_BOX  = str.maketrans({"■": ""})

def clean(txt: str) -> str:
   txt = str(txt)
   txt = txt.translate(STRIP_BOX)
   txt = DIGIT_RUN.sub(lambda m: m.group(1) + m.group(2).translate(SUB_MAP), txt)
   txt = ION_CHARGE.sub(lambda m: m.group(1) + SUP_MAP[m.group(2)], txt)
   return txt

INVALID = re.compile(r'[<>:"/\\|?*]')
safe_name = lambda s: re.sub(r"\s{2,}", " ", INVALID.sub("", s)).strip()

def strip_title_prefix(s: str) -> str:
    """Remove a leading '1 - ' (with optional spaces) from titles displayed in PDFs."""
    return re.sub(r'^\s*1\s*-\s*', '', s).strip()

def body_font(chars: int) -> int:
   if chars < 235:   return 12
   if chars <= 250:  return 11
   if chars <= 325:  return 10
   return 0


# ───────────────────  REPORTLAB STYLES  ──────────────────────────
def get_styles():
    st = getSampleStyleSheet()
    # Titles
    st.add(ParagraphStyle('Doc',  fontName=FONT_RALEWAY_SB, fontSize=16, leading=20, alignment=1, spaceAfter=12, bold=True))
    st.add(ParagraphStyle('AnsH', fontName=FONT_RALEWAY_SB, fontSize=15, leading=20, alignment=1, spaceAfter=20, bold=True))
    # Body
    st.add(ParagraphStyle('Q',    fontName=FONT_LATO_REG,   fontSize=12, leading=15, leftIndent=10, spaceAfter=6))
    st.add(ParagraphStyle('Opt',  fontName=FONT_LATO_REG,   fontSize=12, leading=15, leftIndent=25, spaceAfter=2))
    st.add(ParagraphStyle('Blue', parent=st['Opt'],         textColor=colors.blue))
    st.add(ParagraphStyle('Red',  parent=st['Opt'],         textColor=colors.red))
    st.add(ParagraphStyle('RedU', parent=st['Opt'],         textColor=colors.red, underline=True))
    st.add(ParagraphStyle('Expl', fontName=FONT_LATO_LIGHT, fontSize=10, leading=13, leftIndent=25, spaceAfter=10, textColor=colors.black))
    st.add(ParagraphStyle('TF',   fontName=FONT_LATO_REG,   fontSize=12, leading=15, leftIndent=10, spaceAfter=4))
    st.add(ParagraphStyle('Instr',fontName=FONT_LATO_REG,   fontSize=11, leading=14, leftIndent=10, spaceAfter=8, textColor=colors.black))
    # Keep underscore line font unchanged
    st.add(ParagraphStyle('Line', fontName=FONT,            fontSize=12, leading=14, leftIndent=20, textColor=colors.grey))
    st.add(ParagraphStyle('Note', fontName=FONT_LATO_REG,   fontSize=11, leading=14, textColor=colors.red, alignment=0, spaceAfter=10))
    return st
ST = get_styles()


# ───────────────────  PAGE DECORATION  ───────────────────────────
PAGE_W, PAGE_H = letter
def draw_frame(c):
    # Border removed per user request. This function intentionally does nothing now.
    return


# Page template files directory and paths
TEMPLATES_DIR = pathlib.Path(__file__).parent / "01_Page Templates"
QUESTION_FIRST_IMG       = TEMPLATES_DIR / "01. 1st Question page.jpg"
ANSWER_FIRST_IMG         = TEMPLATES_DIR / "02. 1st Answer page.jpg"
QA_OTHER_IMG             = TEMPLATES_DIR / "03. All Questions,Answers Page.jpg"
PREVIEW_FIRST_IMG        = TEMPLATES_DIR / "04. 1st Preview Page.jpg"
PREVIEW_OTHER_IMG        = TEMPLATES_DIR / "05. Other Preview Pages.jpg"


def _draw_image_if_exists(c, img_path: pathlib.Path):
   """Helper: draw image filling the full page if the path exists. Fails silently."""
   try:
       if img_path.exists():
           img = ImageReader(str(img_path))
           c.drawImage(img, 0, 0, width=PAGE_W, height=PAGE_H, mask='auto')
   except Exception:
       pass


def draw_template_image(c, img_path: pathlib.Path):
    """Draw a specific full-page template image (if present)."""
    _draw_image_if_exists(c, img_path)

def q_first(c, d):
    draw_template_image(c, QUESTION_FIRST_IMG)
    draw_frame(c)
    c.setFont(FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))

def a_first(c, d):
    draw_template_image(c, ANSWER_FIRST_IMG)
    draw_frame(c)
    c.setFont(FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))

def qa_other(c, d):
    draw_template_image(c, QA_OTHER_IMG)
    draw_frame(c)
    c.setFont(FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))

def preview_first_onpage(c, d):
    draw_template_image(c, PREVIEW_FIRST_IMG)
    draw_frame(c)
    c.setFont(FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))

def preview_other_onpage(c, d):
    draw_template_image(c, PREVIEW_OTHER_IMG)
    draw_frame(c)
    c.setFont(FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))


# ───────────────────  OPENAI HELPERS  ────────────────────────────
@backoff.on_exception(backoff.expo, openai.RateLimitError, max_time=180)
def _ask(prompt: str) -> str:
   return openai.chat.completions.create(
       model=MODEL,
       messages=[{"role": "user", "content": prompt}],
       temperature=0.7
   ).choices[0].message.content

def extract_json(md: str) -> list:
   """Extract JSON array from markdown/text response."""
   if not md:
       return []
   
   # Remove markdown code blocks if present
   if '```' in md:
       # Extract content between code fences
       parts = md.split('```')
       if len(parts) >= 3:
           md = parts[1]
           # Remove language identifier (json, python, etc.)
           if md.strip().startswith(('json', 'JSON')):
               md = md[4:].strip()
   
   # Find JSON array
   match = re.search(r'\[.*\]', md, re.DOTALL)
   if not match:
       return []
   
   try:
       result = json.loads(match.group())
       # Ensure it's a list
       if isinstance(result, list):
           return result
       else:
           return []
   except json.JSONDecodeError as e:
       print(f"⚠️  JSON decode error: {e}")
       return []

def get_json(prompt: str) -> list:
    """Send prompt to OpenAI and parse JSON response."""
    if 'OFFLINE_MODE' in globals() and OFFLINE_MODE:
        return []
    try:
        response = _ask(prompt)
        result = extract_json(response)
        if not result:
            # Persist the raw response for debugging
            try:
                with open(pathlib.Path(__file__).parent / 'openai_last_responses.log', 'a', encoding='utf-8') as fh:
                    fh.write('\n--- PROMPT ---\n')
                    fh.write(prompt + '\n')
                    fh.write('\n--- RESPONSE ---\n')
                    fh.write(response + '\n')
            except Exception:
                pass
            print(f"⚠️  No valid JSON found in response (raw response saved to openai_last_responses.log)")
        return result
    except Exception as e:
        print(f"⚠️  Error getting JSON: {e}")
        return []


# ───────────────────  PROMPTS  ───────────────────────────────────
def p_mcq(topic, note, n, curriculum_context):
    """Generate MCQ prompt with curriculum alignment"""
    return (
        f"{curriculum_context}\n\n"
        f"Write EXACTLY {n} higher-order MCQs for: {topic}. "
        f"Ensure each question aligns with the grade level and curriculum specified above and the provided teacher note.\n"
        f"Teacher note: {note}\n"
        'Return JSON list: {"q":"","correct":"","distractors":["","",""],"explanation":""}\n'
        "≤325 chars total per item. Randomise answer order."
    )

def p_tf(topic, note, n, curriculum_context):
    """Generate True/False prompt with curriculum alignment"""
    return (
        f"{curriculum_context}\n\n"
        f"Write EXACTLY {n} higher-order True/False statements for: {topic}. "
        f"Ensure alignment with the grade level and curriculum specified above and the teacher note.\n"
        f"Teacher note: {note}\n"
        'Return JSON list: {"statement":"","answer":true/false,"explanation":""}\n'
        "If answer is false, give ≤15-word explanation, else \"\". ≤325 chars item."
    )

def p_sa(topic, note, n, curriculum_context):
    """Generate short-answer prompt with curriculum alignment"""
    return (
        f"{curriculum_context}\n\n"
        f"Write EXACTLY {n} higher-order short-answer Qs for: {topic}. "
        f"Ensure each question aligns with the grade level and curriculum specified above and the teacher note.\n"
        f"Teacher note: {note}\n"
        'Return JSON list: {"q":"","answer":""}\n'
        "Question ≤250 chars, answer ≤25 words."
    )

def p_tf_with_expl(topic, note, n, curriculum_context):
    """Generate True/False WITH explanation for all items"""
    return (
        f"{curriculum_context}\n\n"
        f"Write EXACTLY {n} True/False statements for: {topic}. After each, provide a concise explanation of why it is True or False. "
        f"Ensure alignment with the grade level, curriculum, and teacher note.\n"
        f"Teacher note: {note}\n"
        'Return JSON list: {"statement":"","answer":true/false,"explanation":""}\n'
        "Each explanation ≤25 words. ≤325 chars per item."
    )

def p_open(topic, note, n, curriculum_context):
    """Generate open-ended questions"""
    return (
        f"{curriculum_context}\n\n"
        f"Write EXACTLY {n} open-ended science questions for: {topic}. Questions should encourage explanation, examples, and reasoning; no single right answer. "
        f"Keep them age-appropriate per the curriculum and teacher note.\n"
        f"Teacher note: {note}\n"
        'Return JSON list: {"q":"","answer":""}\n'
        "Question ≤200 chars, sample answer ≤40 words."
    )

def p_scenario(topic, note, n, curriculum_context):
    """Generate scenario-based questions"""
    return (
        f"{curriculum_context}\n\n"
        f"Write EXACTLY {n} short real-life scenarios for: {topic}. For each, ask ONE question that applies the concept to the scenario. "
        f"Ensure authentic contexts and alignment with teacher note.\n"
        f"Teacher note: {note}\n"
        'Return JSON list: {"q":"","answer":""}\n'
        "Scenario+question ≤275 chars total, sample answer ≤40 words."
    )


# ───────────────────  BUILDERS  ──────────────────────────────────
def build_mcq(topic, note, curriculum_context):
    """Build 30 MCQ cards with curriculum alignment"""
    deck = []
    max_attempts = 15
    attempts = 0
    
    while len(deck) < 30 and attempts < max_attempts:
        attempts += 1
        items = get_json(p_mcq(topic, note, 30-len(deck), curriculum_context))
        
        if not items:
            print(f"   ⚠️  Attempt {attempts}: No items returned, retrying...")
            continue
        
        for itm in items:
            try:
                # Validate item structure
                if not isinstance(itm, dict):
                    continue
                if "q" not in itm or "correct" not in itm or "distractors" not in itm:
                    continue
                
                q  = clean(itm["q"])
                ok = clean(itm["correct"])
                ds = [clean(d) for d in itm["distractors"]]
                exp= clean(itm.get("explanation", ""))
                
                if not q or not ok or len(ds) != 3:
                    continue
                
                opts = ds + [ok]
                random.shuffle(opts)
                if body_font(len(q)+sum(len(o) for o in opts)) == 0:
                    continue
                deck.append((q, opts, "ABCD"[opts.index(ok)], exp))
                if len(deck) == 30:
                    break
            except (KeyError, TypeError, IndexError) as e:
                print(f"   ⚠️  Skipping malformed item: {e}")
                continue
    
    if len(deck) < 30:
        if len(deck) == 0:
            raise RuntimeError(f"Could not build any MCQ cards after {attempts} attempts")
        else:
            print(f"⚠️  Warning: Only built {len(deck)}/30 MCQ cards after {attempts} attempts — continuing with partial deck")
    return deck

def bool_to_str(val):
    if isinstance(val, bool): return "True" if val else "False"
    if isinstance(val, str):  return val.strip().capitalize()
    return ""

def build_tf(topic, note, curriculum_context, target: int = 30):
    """Build True/False cards with curriculum alignment (target count configurable)"""
    deck = []
    max_attempts = 15
    attempts = 0
    
    while len(deck) < target and attempts < max_attempts:
        attempts += 1
        items = get_json(p_tf(topic, note, target-len(deck), curriculum_context))
        
        if not items:
            print(f"   ⚠️  Attempt {attempts}: No items returned, retrying...")
            continue
        
        for itm in items:
            try:
                if not isinstance(itm, dict):
                    continue
                if "statement" not in itm or "answer" not in itm:
                    continue
                
                stmt = clean(itm["statement"])
                ans  = bool_to_str(itm["answer"])
                exp  = clean(itm.get("explanation", ""))
                
                if not stmt or ans not in ("True", "False"):
                    continue
                if body_font(len(stmt)+5) == 0:
                    continue
                deck.append((stmt, ans, exp))
                if len(deck) == target:
                    break
            except (KeyError, TypeError) as e:
                print(f"   ⚠️  Skipping malformed item: {e}")
                continue
    
    if len(deck) < target:
        if len(deck) == 0:
            raise RuntimeError(f"Could not build any True/False cards after {attempts} attempts")
        else:
            print(f"⚠️  Warning: Only built {len(deck)}/{target} True/False cards after {attempts} attempts — continuing with partial deck")
    return deck

def build_sa(topic, note, curriculum_context, target: int = 30):
    """Build Short Answer cards with curriculum alignment (target count configurable)"""
    deck = []
    max_attempts = 15
    attempts = 0
    
    while len(deck) < target and attempts < max_attempts:
        attempts += 1
        items = get_json(p_sa(topic, note, target-len(deck), curriculum_context))
        
        if not items:
            print(f"   ⚠️  Attempt {attempts}: No items returned, retrying...")
            continue
        
        for itm in items:
            try:
                if not isinstance(itm, dict):
                    continue
                if "q" not in itm or "answer" not in itm:
                    continue
                
                q   = clean(itm["q"])
                ans = clean(itm["answer"])
                
                if not q or not ans:
                    continue
                if body_font(len(q)+len(ans)) == 0:
                    continue
                deck.append((q, ans))
                if len(deck) == target:
                    break
            except (KeyError, TypeError) as e:
                print(f"   ⚠️  Skipping malformed item: {e}")
                continue
    
    if len(deck) < target:
        if len(deck) == 0:
            raise RuntimeError(f"Could not build any Short Answer cards after {attempts} attempts")
        else:
            print(f"⚠️  Warning: Only built {len(deck)}/{target} Short Answer cards after {attempts} attempts — continuing with partial deck")
    return deck

def build_tf_expl(topic, note, curriculum_context, target: int = 20):
    """Build True/False with explanation cards (every item includes explanation)."""
    deck = []
    max_attempts = 15
    attempts = 0

    while len(deck) < target and attempts < max_attempts:
        attempts += 1
        items = get_json(p_tf_with_expl(topic, note, target - len(deck), curriculum_context))
        if not items:
            print(f"   ⚠️  Attempt {attempts}: No items returned, retrying...")
            continue
        for itm in items:
            try:
                if not isinstance(itm, dict):
                    continue
                if "statement" not in itm or "answer" not in itm:
                    continue
                stmt = clean(itm["statement"])
                ans  = bool_to_str(itm["answer"])
                exp  = clean(itm.get("explanation", ""))
                if not stmt or ans not in ("True", "False"):
                    continue
                if body_font(len(stmt)+5) == 0:
                    continue
                # For this deck, explanation should always be present; if not, keep empty but continue
                deck.append((stmt, ans, exp))
                if len(deck) == target:
                    break
            except (KeyError, TypeError) as e:
                print(f"   ⚠️  Skipping malformed item: {e}")
                continue
    if len(deck) < target:
        if len(deck) == 0:
            raise RuntimeError(f"Could not build any TF-with-explanation cards after {attempts} attempts")
        else:
            print(f"⚠️  Warning: Only built {len(deck)}/{target} TF-with-explanation cards after {attempts} attempts — continuing with partial deck")
    return deck

def build_open(topic, note, curriculum_context, target: int = 20):
    deck = []
    max_attempts = 15
    attempts = 0
    while len(deck) < target and attempts < max_attempts:
        attempts += 1
        items = get_json(p_open(topic, note, target - len(deck), curriculum_context))
        if not items:
            print(f"   ⚠️  Attempt {attempts}: No items returned, retrying...")
            continue
        for itm in items:
            try:
                if not isinstance(itm, dict) or "q" not in itm or "answer" not in itm:
                    continue
                q   = clean(itm["q"])
                ans = clean(itm["answer"])
                if not q or not ans:
                    continue
                deck.append((q, ans))
                if len(deck) == target:
                    break
            except (KeyError, TypeError) as e:
                print(f"   ⚠️  Skipping malformed item: {e}")
                continue
    if len(deck) < target:
        if len(deck) == 0:
            raise RuntimeError(f"Could not build any Open-Ended cards after {attempts} attempts")
        else:
            print(f"⚠️  Warning: Only built {len(deck)}/{target} Open-Ended cards after {attempts} attempts — continuing with partial deck")
    return deck

def build_scenario(topic, note, curriculum_context, target: int = 10):
    deck = []
    max_attempts = 15
    attempts = 0
    while len(deck) < target and attempts < max_attempts:
        attempts += 1
        items = get_json(p_scenario(topic, note, target - len(deck), curriculum_context))
        if not items:
            print(f"   ⚠️  Attempt {attempts}: No items returned, retrying...")
            continue
        for itm in items:
            try:
                if not isinstance(itm, dict) or "q" not in itm or "answer" not in itm:
                    continue
                q   = clean(itm["q"])
                ans = clean(itm["answer"])
                if not q or not ans:
                    continue
                deck.append((q, ans))
                if len(deck) == target:
                    break
            except (KeyError, TypeError) as e:
                print(f"   ⚠️  Skipping malformed item: {e}")
                continue
    if len(deck) < target:
        if len(deck) == 0:
            raise RuntimeError(f"Could not build any Scenario-Based cards after {attempts} attempts")
        else:
            print(f"⚠️  Warning: Only built {len(deck)}/{target} Scenario-Based cards after {attempts} attempts — continuing with partial deck")
    return deck


# ───────────────────  PDF HELPERS  ───────────────────────────────
def doc(path):
   return SimpleDocTemplate(str(path), pagesize=letter,
                            leftMargin=35, rightMargin=35,
                            topMargin=50,  bottomMargin=40)

def sa_lines():  # TWO lines, 85 underscores each
   line = "_" * 85
   return [Paragraph(line, ST["Line"]) for _ in range(2)]

def lines_n(n: int):
    line = "_" * 85
    return [Paragraph(line, ST["Line"]) for _ in range(max(1, n))]


# ───────────────────  THUMBNAILS (PDF → JPG)  ───────────────────
def render_first_page_to_jpg(pdf_path: pathlib.Path, jpg_path: pathlib.Path, dpi: int = 150) -> bool:
   """Render the first page of a PDF to a JPG file. Returns True on success."""
   try:
       if HAS_FITZ:
           doc = fitz.open(str(pdf_path))
           if doc.page_count == 0:
               return False
           page = doc.load_page(0)
           pix = page.get_pixmap(dpi=dpi)
           pix.save(str(jpg_path))
           doc.close()
           return True
       elif HAS_PDF2IMG:
           pages = convert_from_path(str(pdf_path), first_page=1, last_page=1, dpi=dpi)
           if pages:
               pages[0].save(str(jpg_path), 'JPEG')
               return True
           return False
       else:
           print("⚠️  Cannot create thumbnails: install 'pymupdf' (fitz) or 'pdf2image' + poppler.")
           return False
   except Exception as e:
       print(f"⚠️  Failed to render thumbnail for {pdf_path.name}: {e}")
       return False


# ───────────────────  WORKSHEET MAKERS  ──────────────────────────
def make_mcq(ws_pdf, ans_pdf, main, sub, deck):
    sub_t = strip_title_prefix(sub)
    story=[Paragraph(f"<b>{wrap_special(sub_t, FONT_RALEWAY_SB)}</b>", ST["Doc"]), Spacer(1,10)]
    for i,(q,opts,_,_) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"])] + \
             [Paragraph(f"{l}. {wrap_special(t, FONT_LATO_REG)}", ST["Opt"]) for l,t in zip("ABCD",opts)] + \
             [Spacer(1,12)]
        story.append(KeepTogether(blk))
    doc(ws_pdf).build(story, onFirstPage=q_first, onLaterPages=qa_other)

    hdr=f"{sub_t} - Answer Sheet"
    story=[Paragraph(wrap_special(hdr, FONT_RALEWAY_SB), ST["AnsH"]), Spacer(1,10)]
    for i,(q,opts,ans,exp) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"])]
        for l,t in zip("ABCD",opts):
            style = ST["Red"] if l==ans else ST["Opt"]
            blk.append(Paragraph(f"{l}. {wrap_special(t, FONT_LATO_REG)}", style))
        blk.append(Paragraph(wrap_special(f"Here’s Why : {exp}", FONT_LATO_LIGHT), ST["Expl"]))
        blk.append(Spacer(1,8))
        story.append(KeepTogether(blk))
    doc(ans_pdf).build(story, onFirstPage=a_first, onLaterPages=qa_other)

def make_tf(ws_pdf, ans_pdf, main, sub, deck):
    sub_t = strip_title_prefix(sub)
    story=[Paragraph(f"<b>{wrap_special(sub_t + ' – True/False', FONT_RALEWAY_SB)}</b>", ST["Doc"]), Spacer(1,10)]
    for i,(stmt,_,_) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(stmt, FONT_LATO_REG)}", ST["Q"]) ]
        # One-line options: "a. True" and "b. False"
        opt_row = Table([[Paragraph("a. True", ST["Opt"]), Paragraph("b. False", ST["Opt"]) ]], colWidths=None, hAlign='LEFT')
        opt_row.setStyle(TableStyle([
            ('LEFTPADDING', (0,0), (-1,-1), 25),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0, colors.white),
        ]))
        blk += [opt_row, Spacer(1,12)]
        story.append(KeepTogether(blk))
    doc(ws_pdf).build(story, onFirstPage=q_first, onLaterPages=qa_other)

    hdr=f"{sub_t} - Answer Sheet"
    story=[Paragraph(wrap_special(hdr, FONT_RALEWAY_SB), ST["AnsH"]), Spacer(1,10)]
    for i,(stmt,ans,exp) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(stmt, FONT_LATO_REG)}", ST["Q"])]
        # One-line options on Answer Sheet with red underline for the correct answer
        styleA = ST["RedU"] if ans=="True" else ST["Opt"]
        styleB = ST["RedU"] if ans=="False" else ST["Opt"]
        opt_row = Table([[Paragraph("a. True", styleA), Paragraph("b. False", styleB)]], colWidths=None, hAlign='LEFT')
        opt_row.setStyle(TableStyle([
            ('LEFTPADDING', (0,0), (-1,-1), 25),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0, colors.white),
        ]))
        blk.append(opt_row)
        if ans=="False" and exp:
            blk.append(Paragraph(wrap_special(f"Here’s Why : {exp}", FONT_LATO_LIGHT), ST["Expl"]))
        blk.append(Spacer(1,8))
        story.append(KeepTogether(blk))
    doc(ans_pdf).build(story, onFirstPage=a_first, onLaterPages=qa_other)

def make_sa(ws_pdf, ans_pdf, main, sub, deck):
    sub_t = strip_title_prefix(sub)
    story=[Paragraph(f"<b>{wrap_special(sub_t + ' – Short Answer', FONT_RALEWAY_SB)}</b>", ST["Doc"]), Spacer(1,10)]
    for i,(q,_) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"]) ] + sa_lines() + [Spacer(1,12)]
        story.append(KeepTogether(blk))
    doc(ws_pdf).build(story, onFirstPage=q_first, onLaterPages=qa_other)

    hdr=f"{sub_t} - Answer Sheet"
    story=[Paragraph(wrap_special(hdr, FONT_RALEWAY_SB), ST["AnsH"]), Spacer(1,10)]
    for i,(q,ans) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"]),
             Paragraph(f"<font color='red'>{wrap_special(ans, FONT_LATO_REG)}</font>", ST["Opt"]),
             Spacer(1,8)]
        story.append(KeepTogether(blk))
    doc(ans_pdf).build(story, onFirstPage=a_first, onLaterPages=qa_other)

def make_tf_with_expl(ws_pdf, ans_pdf, main, sub, deck):
    sub_t = strip_title_prefix(sub)
    story=[Paragraph(f"<b>{wrap_special(sub_t + ' – True/False with Explanation', FONT_RALEWAY_SB)}</b>", ST["Doc"]), Spacer(1,10)]
    instr = "Read each statement carefully. Mark it as True (a) or False (b), then explain your answer."
    story.append(Paragraph(wrap_special(instr, FONT_LATO_REG), ST["Instr"]))
    for i,(stmt,_,_) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(stmt, FONT_LATO_REG)}", ST["Q"]) ]
        opt_row = Table([[Paragraph("a. True", ST["Opt"]), Paragraph("b. False", ST["Opt"]) ]], colWidths=None, hAlign='LEFT')
        opt_row.setStyle(TableStyle([
            ('LEFTPADDING', (0,0), (-1,-1), 25),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0, colors.white),
        ]))
        blk += [opt_row]
        blk += lines_n(2) + [Spacer(1,12)]
        story.append(KeepTogether(blk))
    doc(ws_pdf).build(story, onFirstPage=q_first, onLaterPages=qa_other)

    hdr=f"{sub_t} - Answer Sheet"
    story=[Paragraph(wrap_special(hdr, FONT_RALEWAY_SB), ST["AnsH"]), Spacer(1,10)]
    for i,(stmt,ans,exp) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(stmt, FONT_LATO_REG)}", ST["Q"]) ]
        styleA = ST["RedU"] if ans=="True" else ST["Opt"]
        styleB = ST["RedU"] if ans=="False" else ST["Opt"]
        opt_row = Table([[Paragraph("a. True", styleA), Paragraph("b. False", styleB)]], colWidths=None, hAlign='LEFT')
        opt_row.setStyle(TableStyle([
            ('LEFTPADDING', (0,0), (-1,-1), 25),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0, colors.white),
        ]))
        blk.append(opt_row)
        if exp:
            blk.append(Paragraph(wrap_special(f"Here’s Why : {exp}", FONT_LATO_LIGHT), ST["Expl"]))
        blk.append(Spacer(1,8))
        story.append(KeepTogether(blk))
    doc(ans_pdf).build(story, onFirstPage=a_first, onLaterPages=qa_other)

def make_open(ws_pdf, ans_pdf, main, sub, deck):
    sub_t = strip_title_prefix(sub)
    story=[Paragraph(f"<b>{wrap_special(sub_t + ' – Open-Ended Questions', FONT_RALEWAY_SB)}</b>", ST["Doc"]), Spacer(1,10)]
    for i,(q,_) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"])]+ lines_n(3) + [Spacer(1,12)]
        story.append(KeepTogether(blk))
    doc(ws_pdf).build(story, onFirstPage=q_first, onLaterPages=qa_other)

    hdr=f"{sub_t} - Sample Answers"
    story=[Paragraph(wrap_special(hdr, FONT_RALEWAY_SB), ST["AnsH"]), Spacer(1,10)]
    for i,(q,ans) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"]),
             Paragraph(f"<font color='red'>{wrap_special(ans, FONT_LATO_REG)}</font>", ST["Opt"]),
             Spacer(1,8)]
        story.append(KeepTogether(blk))
    doc(ans_pdf).build(story, onFirstPage=a_first, onLaterPages=qa_other)

def make_scenario(ws_pdf, ans_pdf, main, sub, deck):
    sub_t = strip_title_prefix(sub)
    story=[Paragraph(f"<b>{wrap_special(sub_t + ' – Scenario-Based Questions', FONT_RALEWAY_SB)}</b>", ST["Doc"]), Spacer(1,10)]
    for i,(q,_) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"])]+ lines_n(4) + [Spacer(1,12)]
        story.append(KeepTogether(blk))
    doc(ws_pdf).build(story, onFirstPage=q_first, onLaterPages=qa_other)

    hdr=f"{sub_t} - Sample Answers"
    story=[Paragraph(wrap_special(hdr, FONT_RALEWAY_SB), ST["AnsH"]), Spacer(1,10)]
    for i,(q,ans) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"]),
             Paragraph(f"<font color='red'>{wrap_special(ans, FONT_LATO_REG)}</font>", ST["Opt"]),
             Spacer(1,8)]
        story.append(KeepTogether(blk))
    doc(ans_pdf).build(story, onFirstPage=a_first, onLaterPages=qa_other)


# ───────────────────  PREVIEW PDF  ───────────────────────────────
def make_preview(preview_pdf, main, tf_basic, tf_expl, sa, openq, scen):
    # We'll build the preview to a temporary PDF first, then rasterize it
    # into a final image-based PDF so text cannot be copied.
    temp_pdf = pathlib.Path(str(preview_pdf) + '.tmp')
    # Build with BaseDocTemplate to switch page backgrounds
    frame = Frame(35, 40, PAGE_W - 35 - 35, PAGE_H - 50 - 40, id='normal')
    pdf = BaseDocTemplate(str(temp_pdf), pagesize=letter)
    pdf.addPageTemplates([
        PageTemplate(id='PreviewFirst', frames=[frame], onPage=preview_first_onpage, autoNextPageTemplate='PreviewOther'),
        PageTemplate(id='PreviewOther', frames=[frame], onPage=preview_other_onpage),
    ])

    story = []
    # First page (04): emit a PageBreak to create a blank page using the default template
    story.append(PageBreak())

    # Helper renderers mirroring individual PDFs
    def add_tf_questions(deck):
        for i,(stmt,_,_) in enumerate(deck,1):
            blk=[Paragraph(f"{i}. {wrap_special(stmt, FONT_LATO_REG)}", ST["Q"]) ]
            opt_row = Table([[Paragraph("a. True", ST["Opt"]), Paragraph("b. False", ST["Opt"]) ]], colWidths=None, hAlign='LEFT')
            opt_row.setStyle(TableStyle([
                ('LEFTPADDING', (0,0), (-1,-1), 25),
                ('RIGHTPADDING', (0,0), (-1,-1), 0),
                ('TOPPADDING', (0,0), (-1,-1), 0),
                ('BOTTOMPADDING', (0,0), (-1,-1), 0),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0, colors.white),
            ]))
            blk += [opt_row, Spacer(1,12)]
            story.append(KeepTogether(blk))

    def add_tf_answers(deck):
        hdr=f"True/False - Answer Sheet"
        story.append(Paragraph(wrap_special(hdr, FONT_RALEWAY_SB), ST["AnsH"]))
        story.append(Spacer(1,10))
        for i,(stmt,ans,exp) in enumerate(deck,1):
            for_line = lambda opt: f"{opt}. {'True' if opt=='A' else 'False'}"
            blk=[Paragraph(f"{i}. {wrap_special(stmt, FONT_LATO_REG)}", ST["Q"]) ]
            blk.append(Paragraph(for_line('A'), ST["RedU"] if ans=="True"  else ST["Opt"]))
            blk.append(Paragraph(for_line('B'), ST["RedU"] if ans=="False" else ST["Opt"]))
            if exp:
                blk.append(Paragraph(wrap_special(f"Here’s Why : {exp}", FONT_LATO_LIGHT), ST["Expl"]))
            blk.append(Spacer(1,8))
            story.append(KeepTogether(blk))

    def add_tf_expl_questions(deck):
        for i,(stmt,_,_) in enumerate(deck,1):
            blk=[Paragraph(f"{i}. {wrap_special(stmt, FONT_LATO_REG)}", ST["Q"]) ]
            opt_row = Table([[Paragraph("a. True", ST["Opt"]), Paragraph("b. False", ST["Opt"]) ]], colWidths=None, hAlign='LEFT')
            opt_row.setStyle(TableStyle([
                ('LEFTPADDING', (0,0), (-1,-1), 25),
                ('RIGHTPADDING', (0,0), (-1,-1), 0),
                ('TOPPADDING', (0,0), (-1,-1), 0),
                ('BOTTOMPADDING', (0,0), (-1,-1), 0),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0, colors.white),
            ]))
            blk += [opt_row]
            blk += lines_n(2) + [Spacer(1,12)]
            story.append(KeepTogether(blk))

    def add_tf_expl_answers(deck):
        hdr=f"True/False with Explanation - Answer Sheet"
        story.append(Paragraph(wrap_special(hdr, FONT_RALEWAY_SB), ST["AnsH"]))
        story.append(Spacer(1,10))
        for i,(stmt,ans,exp) in enumerate(deck,1):
            for_line = lambda opt: f"{opt}. {'True' if opt=='A' else 'False'}"
            blk=[Paragraph(f"{i}. {wrap_special(stmt, FONT_LATO_REG)}", ST["Q"]) ]
            blk.append(Paragraph(for_line('A'), ST["RedU"] if ans=="True"  else ST["Opt"]))
            blk.append(Paragraph(for_line('B'), ST["RedU"] if ans=="False" else ST["Opt"]))
            if exp:
                blk.append(Paragraph(wrap_special(f"Here’s Why : {exp}", FONT_LATO_LIGHT), ST["Expl"]))
            blk.append(Spacer(1,8))
            story.append(KeepTogether(blk))

    def add_sa_questions(deck):
        for i,(q,_) in enumerate(deck,1):
            blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"]) ] + sa_lines() + [Spacer(1,12)]
            story.append(KeepTogether(blk))

    def add_sa_answers(deck):
        hdr=f"Short Answer - Answer Sheet"
        story.append(Paragraph(wrap_special(hdr, FONT_RALEWAY_SB), ST["AnsH"]))
        story.append(Spacer(1,10))
        for i,(q,ans) in enumerate(deck,1):
            blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"]) ,
                 Paragraph(f"<font color='red'>{wrap_special(ans, FONT_LATO_REG)}</font>", ST["Opt"]),
                 Spacer(1,8)]
            story.append(KeepTogether(blk))

    def add_open_questions(deck):
        for i,(q,_) in enumerate(deck,1):
            blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"]) ] + lines_n(3) + [Spacer(1,12)]
            story.append(KeepTogether(blk))

    def add_open_answers(deck):
        hdr=f"Open-Ended - Sample Answers"
        story.append(Paragraph(wrap_special(hdr, FONT_RALEWAY_SB), ST["AnsH"]))
        story.append(Spacer(1,10))
        for i,(q,ans) in enumerate(deck,1):
            blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"]) ,
                 Paragraph(f"<font color='red'>{wrap_special(ans, FONT_LATO_REG)}</font>", ST["Opt"]),
                 Spacer(1,8)]
            story.append(KeepTogether(blk))

    def add_scen_questions(deck):
        for i,(q,_) in enumerate(deck,1):
            blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"]) ] + lines_n(4) + [Spacer(1,12)]
            story.append(KeepTogether(blk))

    def add_scen_answers(deck):
        hdr=f"Scenario-Based - Sample Answers"
        story.append(Paragraph(wrap_special(hdr, FONT_RALEWAY_SB), ST["AnsH"]))
        story.append(Spacer(1,10))
        for i,(q,ans) in enumerate(deck,1):
            blk=[Paragraph(f"{i}. {wrap_special(q, FONT_LATO_REG)}", ST["Q"]) ,
                 Paragraph(f"<font color='red'>{wrap_special(ans, FONT_LATO_REG)}</font>", ST["Opt"]),
                 Spacer(1,8)]
            story.append(KeepTogether(blk))

    # Now compose the full preview content: Q then Answers for each type
    # True/False
    story.append(Paragraph(wrap_special("True/False – Questions", FONT_RALEWAY_SB), ST["AnsH"]))
    add_tf_questions(tf_basic)
    story.append(PageBreak())
    add_tf_answers(tf_basic)

    # TF with Explanation
    story.append(PageBreak())
    story.append(Paragraph(wrap_special("True/False with Explanation – Questions", FONT_RALEWAY_SB), ST["AnsH"]))
    instr = "Read each statement carefully. Mark it as True (a) or False (b), then explain your answer."
    story.append(Paragraph(wrap_special(instr, FONT_LATO_REG), ST["Instr"]))
    add_tf_expl_questions(tf_expl)
    story.append(PageBreak())
    add_tf_expl_answers(tf_expl)

    # Short Answer
    story.append(PageBreak())
    story.append(Paragraph(wrap_special("Short Answer – Questions", FONT_RALEWAY_SB), ST["AnsH"]))
    add_sa_questions(sa)
    story.append(PageBreak())
    add_sa_answers(sa)

    # Open-Ended
    story.append(PageBreak())
    story.append(Paragraph(wrap_special("Open-Ended – Questions", FONT_RALEWAY_SB), ST["AnsH"]))
    add_open_questions(openq)
    story.append(PageBreak())
    add_open_answers(openq)

    # Scenario-Based
    story.append(PageBreak())
    story.append(Paragraph(wrap_special("Scenario-Based – Questions", FONT_RALEWAY_SB), ST["AnsH"]))
    add_scen_questions(scen)
    story.append(PageBreak())
    add_scen_answers(scen)

    pdf.build(story)

    # Rasterize the generated temporary PDF into images and write the final PDF
    def _rasterize_pdf_to_image_pdf(src: pathlib.Path, dst: pathlib.Path, dpi: int = 150):
        """Render each page of src PDF to an image and create a new PDF dst with those images as full pages.
        Uses PyMuPDF (fitz) if available, else pdf2image. If neither available, copies src to dst as fallback.
        """
        try:
            if 'fitz' in globals() and HAS_FITZ:
                doc = fitz.open(str(src))
                c = pdfcanvas.Canvas(str(dst), pagesize=letter)
                for p in doc:
                    mat = fitz.Matrix(dpi/72.0, dpi/72.0)
                    pix = p.get_pixmap(matrix=mat, alpha=False)
                    img_bytes = pix.tobytes('png')
                    img = ImageReader(io.BytesIO(img_bytes))
                    c.drawImage(img, 0, 0, width=PAGE_W, height=PAGE_H, mask='auto')
                    c.showPage()
                c.save()
                doc.close()
                return True
            elif 'convert_from_path' in globals() and HAS_PDF2IMG:
                pages = convert_from_path(str(src), dpi=dpi)
                c = pdfcanvas.Canvas(str(dst), pagesize=letter)
                for pil in pages:
                    img = ImageReader(pil)
                    c.drawImage(img, 0, 0, width=PAGE_W, height=PAGE_H, mask='auto')
                    c.showPage()
                c.save()
                return True
            else:
                # No raster library available — fallback to copying (text will remain selectable).
                shutil.copyfile(str(src), str(dst))
                return False
        except Exception as e:
            print(f"⚠️  Failed to rasterize preview PDF: {e}")
            try:
                shutil.copyfile(str(src), str(dst))
            except Exception:
                pass
            return False

    # Attempt rasterization and remove temp file afterwards
    try:
        _rasterize_pdf_to_image_pdf(temp_pdf, pathlib.Path(preview_pdf), dpi=150)
    finally:
        try:
            temp_pdf.unlink()
        except Exception:
            pass


# ───────────────────  MAIN  ──────────────────────────────────────
def parse_num_list(spec: str) -> List[int]:
    """Parse a comma separated list of ints and ranges like '1,3-5,8'."""
    out: List[int] = []
    for part in (spec or "").split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            a,b = part.split('-',1)
            if a.isdigit() and b.isdigit():
                out.extend(range(int(a), int(b)+1))
        elif part.isdigit():
            out.append(int(part))
    return sorted(set(out))

def list_standards():
    """List available standards from loaded curriculum"""
    global CURRICULUM_DATA
    if not CURRICULUM_DATA:
        print("❌ No curriculum loaded. Please ensure details.xlsx exists.")
        return
    
    print(f"\n{'='*70}")
    print(f"Curriculum: {CURRICULUM_DATA.curriculum_name}")
    print(f"Subject: {CURRICULUM_DATA.subject_name}")
    print(f"Grade: {CURRICULUM_DATA.grade_level}")
    print(f"{'='*70}\n")
    print("Available Standards (unit_index – title):")
    for m_idx, main_title, subs in CURRICULUM_DATA.units:
        print(f"  {m_idx}: {main_title}  ({len(subs)} subtopics)")
        for s_idx, sub_title, _ in subs:
            print(f"     {m_idx}.{s_idx}: {sub_title}")

def build_selected_for_curriculum(standards: List[int], subs_filter: List[int], dry_run: bool, main_folder: pathlib.Path):
    """Build worksheets for selected standards in a specific curriculum"""
    global CURRICULUM_DATA
    if not CURRICULUM_DATA:
        print("❌ No curriculum loaded. Please ensure details.xlsx exists.")
        return
    
    # Get curriculum context for OpenAI
    curriculum_context = CURRICULUM_DATA.get_prompt_context()
    
    tasks = []
    for m_idx, main_title, subs in CURRICULUM_DATA.units:
        if standards and m_idx not in standards:
            continue
        for s_idx, sub_title, note in subs:
            if subs_filter and s_idx not in subs_filter:
                continue
            tasks.append((m_idx, main_title, s_idx, sub_title, note))
    
    if dry_run:
        print(f"🛈 Dry-run: would generate {len(tasks)} topic(s) for this curriculum.")
        for m_idx, main_title, s_idx, sub_title, _ in tasks[:25]:
            print(f"  {m_idx}.{s_idx} – {sub_title} (in {main_title})")
        if len(tasks) > 25:
            print(f"  … {len(tasks)-25} more")
        return

    if not tasks:
        print("⚠  No matching standards/subtopics for this curriculum.")
        return

    ensure_api_key()
    
    start = time.time()
    # Track created Main Category (unit) folders to avoid duplicates for the same title
    # SIMPLIFIED: Key strictly by normalized main_title (unit title) only.
    # This guarantees a single 3rd-level folder per unit regardless of subtopic content/formatting.
    unit_folder_map: Dict[str, pathlib.Path] = {}
    for m_idx, main_title, s_idx, sub_title, note in tasks:
        # Build/lookup the 3rd-level folder strictly by unit title (main_title)
        unit_key = safe_name(main_title.strip())
        if unit_key in unit_folder_map:
            unit_dir = unit_folder_map[unit_key]
        else:
            unit_name = f"{m_idx:02d}. {main_title.strip()}"
            unit_dir = main_folder / safe_name(unit_name)
            unit_dir.mkdir(exist_ok=True)
            unit_folder_map[unit_key] = unit_dir

        # 4th level: Subcategory (Performance Expectation) folder
        topic_dir = unit_dir / safe_name(f"{s_idx:02d}. {sub_title}")
        topic_dir.mkdir(exist_ok=True)

        tf_basic_dir = topic_dir / "1. True, False Type Questions";                       tf_basic_dir.mkdir(exist_ok=True)
        tf_expl_dir  = topic_dir / "2. True-False Type Questions with Explanation";       tf_expl_dir.mkdir(exist_ok=True)
        sa_dir       = topic_dir / "3. Short Answer Type Questions";                      sa_dir.mkdir(exist_ok=True)
        open_dir     = topic_dir / "4. Open-Ended Questions";                              open_dir.mkdir(exist_ok=True)
        scen_dir     = topic_dir / "5. Scenario-Based Questions";                          scen_dir.mkdir(exist_ok=True)
        prev_dir     = topic_dir / "6. Preview PDFs (Do not include this Folder in Archive.zip file)"; prev_dir.mkdir(exist_ok=True)

        print(f"🔧  Building sets for {m_idx}.{s_idx} – {sub_title} …")
        # Build decks with new target sizes (fall back to placeholders if generation fails)
        try:
            tf_basic = build_tf(sub_title, note, curriculum_context, target=25)
        except Exception as e:
            print(f"   ⚠️  TF generation failed: {e}. Using placeholder items.")
            tf_basic = [(f"Sample statement for {sub_title}", "True", "")] * 5
        try:
            tf_expl  = build_tf_expl(sub_title, note, curriculum_context, target=25)
        except Exception as e:
            print(f"   ⚠️  TF with Explanation generation failed: {e}. Using placeholder items.")
            tf_expl = [(f"Sample statement for {sub_title}", "True", "Because it aligns with the concept.")] * 5
        try:
            sa       = build_sa(sub_title, note, curriculum_context, target=20)
        except Exception as e:
            print(f"   ⚠️  Short Answer generation failed: {e}. Using placeholder items.")
            sa = [(f"Sample short-answer question for {sub_title}", "A concise model answer.")] * 4
        try:
            openq    = build_open(sub_title, note, curriculum_context, target=20)
        except Exception as e:
            print(f"   ⚠️  Open-Ended generation failed: {e}. Using placeholder items.")
            openq = [(f"Sample open-ended question for {sub_title}", "A thoughtful sample response.")] * 4
        try:
            scen     = build_scenario(sub_title, note, curriculum_context, target=10)
        except Exception as e:
            print(f"   ⚠️  Scenario-Based generation failed: {e}. Using placeholder items.")
            scen = [(f"Sample scenario for {sub_title} – What would you predict?", "A plausible explanation.")] * 2

        # Remove leading "1 - " from file base names if present, then sanitize
        base_raw = re.sub(r"^\s*1\s*-\s*", "", sub_title)
        base = safe_name(base_raw)
        # Prepare PDF paths
        tf_ws_pdf   = tf_basic_dir / f"{base} – True-False Worksheet.pdf"
        tf_ans_pdf  = tf_basic_dir / f"{base} – True-False Answer Sheet.pdf"
        tfe_ws_pdf  = tf_expl_dir  / f"{base} – True-False with Explanation Worksheet.pdf"
        tfe_ans_pdf = tf_expl_dir  / f"{base} – True-False with Explanation Answer Sheet.pdf"
        sa_ws_pdf   = sa_dir       / f"{base} – Short Answer Worksheet.pdf"
        sa_ans_pdf  = sa_dir       / f"{base} – Short Answer Answer Sheet.pdf"
        op_ws_pdf   = open_dir     / f"{base} – Open-Ended Worksheet.pdf"
        op_ans_pdf  = open_dir     / f"{base} – Open-Ended Sample Answers.pdf"
        sc_ws_pdf   = scen_dir     / f"{base} – Scenario-Based Worksheet.pdf"
        sc_ans_pdf  = scen_dir     / f"{base} – Scenario-Based Sample Answers.pdf"

        # Generate PDFs
        make_tf         (tf_ws_pdf,  tf_ans_pdf,  main_title, sub_title, tf_basic)
        make_tf_with_expl(tfe_ws_pdf, tfe_ans_pdf, main_title, sub_title, tf_expl)
        make_sa         (sa_ws_pdf,  sa_ans_pdf,  main_title, sub_title, sa)
        make_open       (op_ws_pdf,  op_ans_pdf,  main_title, sub_title, openq)
        make_scenario   (sc_ws_pdf,  sc_ans_pdf,  main_title, sub_title, scen)

        make_preview(prev_dir / f"{base} – Preview.pdf",
                     main_title, tf_basic, tf_expl, sa, openq, scen)
        print("   ✔  Done")
    
    print(f"\n  ✅ Completed {len(tasks)} topic(s) for this curriculum in {time.time()-start:.1f}s")

def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Worksheet PDF Builder (Excel-based curriculum loader)")
    p.add_argument('--excel','-E', default="details.xlsx", help="Path to Excel file with curriculum data (default: details.xlsx)")
    p.add_argument('--standards','-S', help="Comma list / ranges of standard unit indices (e.g. 1,3-5). Default: all.")
    p.add_argument('--subs','-T', help="Comma list / ranges of subtopic indices within selected standards (e.g. 1,4-6). Default: all.")
    p.add_argument('--list', action='store_true', help="List available standards & subtopics then exit.")
    p.add_argument('--dry-run', action='store_true', help="Show what would be generated without calling the API.")
    return p.parse_args(argv)

def main(argv=None):
    global CURRICULUM_DATA
    
    args = parse_args(argv)
    
    # Load curricula from Excel (can be multiple)
    script_dir = pathlib.Path(__file__).parent
    excel_file = script_dir / args.excel
    
    print(f"📖 Loading curricula from {excel_file}\n")
    all_curricula = load_curriculum_from_excel(str(excel_file))
    
    if args.list:
        # List all curricula
        for curr_idx, curriculum in enumerate(all_curricula, 1):
            CURRICULUM_DATA = curriculum
            print(f"\n{'='*70}")
            print(f"CURRICULUM #{curr_idx}")
            print(f"{'='*70}")
            list_standards()
        return
    
    # Create main root folder with required naming: "The Dreaming Caterpillar - DATE - TIME(eg 09.00 AM)"
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%I.%M %p")
    root_folder_name = f"The Dreaming Caterpillar  - {date_str} - {time_str}"
    root_folder = pathlib.Path(root_folder_name)
    root_folder.mkdir(exist_ok=True)
    print(f"📦 Created root folder: {root_folder}\n")
    
    overall_start = time.time()
    
    # Process each curriculum separately
    for curr_idx, curriculum in enumerate(all_curricula, 1):
        CURRICULUM_DATA = curriculum
        
        print(f"\n{'#'*70}")
        print(f"# Processing Curriculum {curr_idx}/{len(all_curricula)}")
        print(f"# {CURRICULUM_DATA.subject_name} - {CURRICULUM_DATA.grade_level}")
        print(f"{'#'*70}\n")
        
        # 2nd folder: curriculum folder: "Grade - Curriculum name - Subject name"
        main_folder_name = f"{safe_name(CURRICULUM_DATA.grade_level)} - {safe_name(CURRICULUM_DATA.curriculum_name)} - {safe_name(CURRICULUM_DATA.subject_name)}"
        main_folder = root_folder / main_folder_name
        main_folder.mkdir(exist_ok=True)
        print(f"📁 Main curriculum folder: {main_folder}\n")
        
        standards = parse_num_list(args.standards) if args.standards else []
        subs_filter = parse_num_list(args.subs) if args.subs else []
        
        # Build selected topics for this curriculum
        build_selected_for_curriculum(standards, subs_filter, args.dry_run, main_folder)
    
    print(f"\n{'='*70}")
    print(f"🎉  ALL CURRICULA COMPLETED in {time.time()-overall_start:.1f}s")
    print(f"📦  All files saved in: {root_folder.absolute()}")
    print(f"{'='*70}")


# ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
   random.seed()
   main()
