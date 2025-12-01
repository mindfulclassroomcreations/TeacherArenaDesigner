#!/usr/bin/env python3

"""

Worksheet PDF Builder – MCQ • True/False • Short-Answer  (v9.1)

───────────────────────────────────────────────────────────────
 
Change from v9 → v9.1

─────────────────────

• Short-Answer worksheets and preview now show TWO lines of
 98 consecutive underscores:

     ____________________________________________________________________________________

 (no extra characters, no hair-spaces).
"""

import os, re, json, random, pathlib, sys, time, argparse, string
from typing import List, Tuple
from datetime import datetime

import pandas as pd

from reportlab.lib.pagesizes import letter
from reportlab.platypus      import (SimpleDocTemplate, Paragraph, Spacer,
                                    KeepTogether, PageBreak, Flowable,
                                    BaseDocTemplate, Frame, PageTemplate, NextPageTemplate)
from reportlab.lib.styles    import getSampleStyleSheet, ParagraphStyle
from reportlab.lib           import colors
from reportlab.pdfbase       import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen        import canvas
import openai, backoff, tempfile
try:
    # Prefer pypdf if available
    from pypdf import PdfMerger, PdfReader, PdfWriter
except Exception:
    try:
        # Newer PyPDF2 versions expose similar names
        from PyPDF2 import PdfMerger, PdfReader, PdfWriter
    except Exception:
        PdfMerger = PdfReader = PdfWriter = None


# ───────────────────────  CONFIG  ────────────────────────────────
openai.api_key = openai.api_key or os.getenv("OPENAI_API_KEY")

def ensure_api_key():
    """Exit with a helpful message if the API key is missing.

    Deferred so that non‑generation commands like --list or --dry-run still work
    without requiring a key.
    """
    if not openai.api_key:
         sys.exit("❌  Set OPENAI_API_KEY environment variable (OPENAI_API_KEY).")
MODEL = "gpt-4o-mini"

FONT_PATHS = [
    "DejaVuSans.ttf",
    "NotoSans-VariableFont_wdth,wght.ttf",
    "Fredoka-VariableFont_wdth.ttf",
    r"C:\Users\DELL\Desktop\2. BIT\0. Fonts\Fredoka-Regular.ttf",
]

def register_fonts():
   # Title Font: Comic Neue Bold for main titles and section headers
   title_candidates = [
       "ComicNeue-Bold.ttf",
       "Comic Neue Bold.ttf",
       "/Library/Fonts/ComicNeue-Bold.ttf",
       "/System/Library/Fonts/ComicNeue-Bold.ttf",
       "CenturyGothic.ttf",
       "Century Gothic.ttf",
       "/Library/Fonts/Century Gothic.ttf",
       "/System/Library/Fonts/Century Gothic.ttf",
   ]
   
   # Question & Answer Font: Nunito Regular
   body_candidates = [
      # Prefer DejaVu or Noto for robust Unicode sub/superscript rendering
      "DejaVuSans.ttf",
      "NotoSans-VariableFont_wdth,wght.ttf",
      "Nunito-Regular.ttf",
      "Nunito.ttf",
      "/Library/Fonts/Nunito-Regular.ttf",
      "/System/Library/Fonts/Nunito-Regular.ttf",
      "Verdana.ttf",
      "/Library/Fonts/Verdana.ttf",
      "/System/Library/Fonts/Verdana.ttf",
   ]
   
   # Explanation Font: Nunito Light Italic
   explanation_candidates = [
       "Nunito-LightItalic.ttf",
       "Nunito-Light-Italic.ttf",
       "/Library/Fonts/Nunito-LightItalic.ttf",
       "/System/Library/Fonts/Nunito-LightItalic.ttf",
       "Nunito-Italic.ttf",
       "/Library/Fonts/Nunito-Italic.ttf",
   ]

   title = None
   body = None
   explanation = None

   # Try title candidates then fallbacks
   for p in title_candidates + FONT_PATHS:
       try:
           fam = pathlib.Path(p).stem
           pdfmetrics.registerFont(TTFont(fam, p))
           title = fam
           break
       except Exception:
           pass

   # Try body candidates then fallbacks
   for p in body_candidates + FONT_PATHS:
       try:
           fam = pathlib.Path(p).stem
           pdfmetrics.registerFont(TTFont(fam, p))
           body = fam
           break
       except Exception:
           pass
   
   # Try explanation candidates then fall back to body italic
   for p in explanation_candidates:
       try:
           fam = pathlib.Path(p).stem
           pdfmetrics.registerFont(TTFont(fam, p))
           explanation = fam
           break
       except Exception:
           pass

   if not body:
       # Try some common system locations for DejaVu or Noto before falling back
       extra_paths = [
           "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
           "/usr/local/share/fonts/DejaVuSans.ttf",
           "/Library/Fonts/DejaVuSans.ttf",
           "/usr/share/fonts/truetype/noto/NotoSans-VariableFont_wdth,wght.ttf",
           "/Library/Fonts/NotoSans-VariableFont_wdth,wght.ttf",
       ]
       for p in extra_paths:
           try:
               if pathlib.Path(p).exists():
                   fam = pathlib.Path(p).stem
                   pdfmetrics.registerFont(TTFont(fam, p))
                   body = fam
                   break
           except Exception:
               pass
       if not body:
           body = "Helvetica"
   if not title:
       title = body
   if not explanation:
       explanation = body  # Fallback to body font if explanation font not found
   
   return title, body, explanation

TITLE_FONT, BODY_FONT, EXPL_FONT = register_fonts()

# Curriculum constants
CURRICULUM_NAME = "NGSS - Middle School Physical Sciences"
# Grade level
GRADE_LEVELS = "6,7,8"


# ─────────────────────  UTILITIES  ───────────────────────────────
SUB_MAP = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
SUP_MAP = {"+": "⁺", "-": "⁻"}
DIGIT_RUN  = re.compile(r'([A-Za-z\)])(\d+)')
ION_CHARGE = re.compile(r'([A-Za-z₀-₉]+)([+-])$')
STRIP_BOX  = str.maketrans({
    "■": "",
    "�": "",
    "□": "",
    "▯": "",
})

def clean(txt: str) -> str:
   txt = str(txt)
   txt = txt.translate(STRIP_BOX)
   txt = DIGIT_RUN.sub(lambda m: m.group(1) + m.group(2).translate(SUB_MAP), txt)
   txt = ION_CHARGE.sub(lambda m: m.group(1) + SUP_MAP[m.group(2)], txt)
   return txt


def strip_curriculum_code(s: str) -> str:
    """Remove leading curriculum code like 'MS-LS1-1 – ' or similar from a title.
    If no dash separator is found, return original string.
    """
    if not s or not isinstance(s, str):
        return s
    for sep in ('–', '—', '-'):
        if sep in s:
            left, right = s.split(sep, 1)
            # only strip when left side looks like a code (contains letters/digits and is reasonably short)
            if re.search(r'[A-Za-z0-9]', left) and len(left) < 40:
                return right.strip()
    return s

INVALID = re.compile(r'[<>:"/\\|?*]')
safe_name = lambda s: re.sub(r"\s{2,}", " ", INVALID.sub("", s)).strip()

def body_font(chars: int) -> int:
   if chars < 235:   return 12
   if chars <= 250:  return 11
   if chars <= 325:  return 10
   return 0

# Task card character cap
MAX_CARD_CHARS = 325


# ───────────────────  REPORTLAB STYLES  ──────────────────────────
def get_styles():
    st = getSampleStyleSheet()
    def add(n, **kw):
        # Use BODY_FONT as the default fontName unless overridden
        font = kw.pop('fontName', BODY_FONT)
        st.add(ParagraphStyle(n, fontName=font, **kw))
    add('Doc',   fontName=TITLE_FONT, fontSize=16, leading=20, alignment=1, spaceAfter=12, bold=True)
    add('Q',     fontSize=12, leading=15, leftIndent=10, spaceAfter=6)
    add('Opt',   fontSize=12, leading=15, leftIndent=25, spaceAfter=2)
    add('Blue',  parent=st['Opt'], textColor=colors.blue)
    add('Expl',  fontName=EXPL_FONT, fontSize=10, leading=13, leftIndent=25, spaceAfter=10, textColor=colors.grey)
    add('AnsH',  fontName=TITLE_FONT, fontSize=15, leading=20, alignment=1, spaceAfter=20, bold=True)
    add('TF',    fontSize=12, leading=15, leftIndent=10, spaceAfter=4)
    add('Line',  fontSize=12, leading=14, leftIndent=20, textColor=colors.grey)
    add('Note',  fontSize=11, leading=14, textColor=colors.red, alignment=0, spaceAfter=10)
    # Preview title styles (center-aligned, blue, bold, Century Gothic if available)
    add('PreviewTitleMain', fontName=TITLE_FONT, fontSize=24, leading=28, alignment=1, textColor=colors.blue)
    add('PreviewTitleSub',  fontName=TITLE_FONT, fontSize=16, leading=20, alignment=1, textColor=colors.blue)
    return st
ST = get_styles()


# ───────────────────  PAGE DECORATION  ───────────────────────────
PAGE_W, PAGE_H = letter
# Standard page margins used across the document and for the border
M_LEFT = 35
M_RIGHT = 35
M_TOP = 50
M_BOTTOM = 40

# How far the border is inset from the absolute page edge (very small)
BORDER_INSET = 6

# Task card corner radius (points). 1pt ≈ 1px at 72dpi; choose 10pt (~10px) in the requested 8–12px range.
CARD_CORNER_RADIUS = 10

def draw_frame(c):
    """Draw a rectangular border close to the absolute page edge (small inset)."""
    x = BORDER_INSET
    y = BORDER_INSET
    width = PAGE_W - 2 * BORDER_INSET
    height = PAGE_H - 2 * BORDER_INSET
    c.setStrokeColor(colors.black)
    c.setLineWidth(2)
    c.rect(x, y, width, height)

def first(c, d):
    draw_frame(c)
    # Place Name and Section inside the page border (top-left area)
    c.setFont(BODY_FONT, 12)
    name_x = BORDER_INSET + 10
    name_y = PAGE_H - BORDER_INSET - 26 
    c.drawString(name_x, name_y, "Name: ...........................................................          Date : ........................")
    c.setFont(BODY_FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))

def first_page_no_watermark(c, d):
    # First page with border but no watermark
    draw_frame(c)
    # Add notice text in red above the page number
    c.saveState()
    c.setFillColor(colors.red)
    # Use the registered title font for the notice (fall back will be handled by register_fonts)
    c.setFont(TITLE_FONT, 18)

    notice_text = (
        "This is a combined preview only. After purchase, you will receive separate question PDFs and "
        "separate answer PDFs with explanations."
    )

    # Wrap text to fit within inner border width (inside the drawn frame)
    max_width = PAGE_W - 2 * BORDER_INSET - 40
    words = notice_text.split()
    lines = []
    current_line = []
    for word in words:
        test_line = ' '.join(current_line + [word])
        if pdfmetrics.stringWidth(test_line, TITLE_FONT, 18) <= max_width:
            current_line.append(word)
        else:
            if current_line:
                lines.append(' '.join(current_line))
            current_line = [word]
    if current_line:
        lines.append(' '.join(current_line))

    # Position text above page number (page number is at y=25)
    line_height = 22
    start_y = 25 + 15 + (len(lines) * line_height)
    y = start_y
    for line in lines:
        c.drawCentredString(PAGE_W/2, y, line)
        y -= line_height
    c.restoreState()

    # Place Name and Section inside the page border (top-left area)
    # Note: first preview page intentionally omits the Name/Date overlay
    # (other pages use the later() handler which draws the Name/Date)

    # Draw page number at bottom
    c.setFont(BODY_FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))

def later(c, d):
    draw_frame(c)
    # Place Name and Section inside the page border (top-left area)
    c.setFont(BODY_FONT, 12)
    name_x = BORDER_INSET + 10
    name_y = PAGE_H - BORDER_INSET - 26
    c.drawString(name_x, name_y, "Name: ...........................................................          Date : ........................")
    c.setFont(BODY_FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))

def preview_page(c, d):
    later(c, d)
    c.saveState()
    c.setFont(BODY_FONT, 130)
    c.setFillColor(colors.lightgrey)
    c.translate(PAGE_W/2, PAGE_H/2)
    c.rotate(45)
    c.drawCentredString(0, 0, "PREVIEW")
    c.restoreState()


def preview_page_no_watermark(c, d):
    # Border without watermark - for title pages
    draw_frame(c)
    # Place Name and Section inside the page border (top-left area)
    c.setFont(BODY_FONT, 12)
    name_x = BORDER_INSET + 10
    name_y = PAGE_H - BORDER_INSET - 26
    c.drawString(name_x, name_y, "Name: ...........................................................          Date : ........................")
    c.setFont(BODY_FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))


def preview_page_plain(c, d):
    # No frame, no page number, no watermark — truly plain for title pages
    # Intentionally empty: TitlePage draws its own white background and text.
    return


class TitlePage(Flowable):
    """A full-page title sheet with centered multi-line text with different fonts and sizes."""
    def __init__(self, line_specs: List[Tuple[str, str, int]]):
        """
        line_specs: List of tuples (text, fontName, fontSize)
        Example: [("WORKSHEET 01", "Helvetica-Bold", 26), ("True or False Questions", "Helvetica-Bold", 22)]
        """
        super().__init__()
        self.line_specs = line_specs

    def wrap(self, availWidth, availHeight):
        self._availWidth = availWidth
        self._availHeight = availHeight
        return availWidth, availHeight

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFillColor(colors.black)
        
        # Calculate total height needed for all lines
        line_heights = []
        for text, font_name, font_size in self.line_specs:
            try:
                ascent = pdfmetrics.getAscent(font_name)
                descent = abs(pdfmetrics.getDescent(font_name))
                ascent_pt = ascent * font_size / 1000.0
                descent_pt = descent * font_size / 1000.0
                line_h = ascent_pt + descent_pt
            except Exception:
                line_h = font_size * 1.2
            line_heights.append((line_h, descent_pt if 'descent_pt' in locals() else font_size * 0.2))
        
        # Calculate spacing between lines (proportional to font size)
        spacings = []
        for i in range(len(self.line_specs) - 1):
            avg_size = (self.line_specs[i][2] + self.line_specs[i+1][2]) / 2
            spacings.append(avg_size * 0.8)  # 80% of average font size as spacing
        
        total_h = sum(h for h, d in line_heights) + sum(spacings)
        
        # Start position to center the whole block vertically
        start_y = PAGE_H/2 + (total_h / 2)
        
        y = start_y
        for i, (text, font_name, font_size) in enumerate(self.line_specs):
            c.setFont(font_name, font_size)
            line_h, descent = line_heights[i]
            
            # Center text horizontally
            try:
                w = pdfmetrics.stringWidth(text, font_name, font_size)
            except Exception:
                w = font_size * len(text) * 0.5
            
            x = PAGE_W/2 - (w / 2)
            c.drawString(x, y - line_h + descent, text)
            
            # Move down for next line
            y -= line_h
            if i < len(spacings):
                y -= spacings[i]
        
        c.restoreState()


# Small Flowable to render an underlined answer with a blue underline
class UnderlinedAnswer(Flowable):
    def __init__(self, text, width=400, fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=0):
        super().__init__()
        self.text = text
        self.width = width
        self.fontName = fontName
        self.fontSize = fontSize
        self.lineColor = lineColor
        self.leftIndent = leftIndent

    def wrap(self, availWidth, availHeight):
        # height: text height + small gap + line thickness
        return availWidth, self.fontSize + 6

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFont(self.fontName, self.fontSize)
        c.setFillColor(colors.black)
        # honor left indent so the underlined answer lines up with Paragraphs using ST['Opt']
        x = getattr(self, 'leftIndent', 0) or 0
        c.drawString(x, 4, self.text)
        # draw underline in requested color
        text_width = pdfmetrics.stringWidth(self.text, self.fontName, self.fontSize)
        c.setStrokeColor(self.lineColor)
        c.setLineWidth(1.5)
        c.line(x, 2, x + text_width + 2, 2)
        c.restoreState()


# ───────────────────  OPENAI HELPERS  ────────────────────────────
@backoff.on_exception(backoff.expo, openai.RateLimitError, max_time=180)
def _ask(prompt: str) -> str:
   return openai.chat.completions.create(
       model=MODEL,
       messages=[{"role": "user", "content": prompt}],
       temperature=0.7
   ).choices[0].message.content

def extract_json(md: str) -> list:
   match = re.search(r"\[.*\]", md, re.S)
   return json.loads(match.group()) if match else []

def get_json(prompt: str) -> list:
   try:
       response = _ask(prompt)
       result = extract_json(response)
       if not result:
           print(f"⚠️  Warning: No JSON found in API response")
       return result
   except json.JSONDecodeError as e:
       print(f"⚠️  JSON decode error: {e}")
       return []
   except Exception as e:
       print(f"⚠️  API error: {e}")
       return []


# ───────────────────  PROMPTS  ───────────────────────────────────
def p_mcq(topic, note, n):
   # Day 2 - Knowledge Builder: Multiple Choice Review
   # Cognitive Depth: Concept recognition & key details
   # Bloom's Level: Understand / Apply
   return (
       f"Write EXACTLY {n} MCQs for Day 2 - Knowledge Builder: Multiple Choice Review on: {topic}. "
       f"Focus on concept recognition and key details. Target Bloom's levels: Understand/Apply. "
       f"Ensure each question aligns with the {CURRICULUM_NAME} ({GRADE_LEVELS}) and the provided teacher note.\n"
       f"Teacher note: {note}\n"
       'Return JSON list: {"q":"","correct":"","distractors":["","",""],"explanation":""}\n'
       "≤325 chars total per item. Randomise answer order."
   )


def build_task_cards(topic, note, n=30):
    """Generate n multiple-choice cards suitable for task-cards.
    Returns list of tuples: (title, q, opts, correct_letter, explanation, num).
    After 5 failed attempts the builder will request only the remaining items using
    the short/simple MCQ prompt so previously-collected items are preserved.
    """
    deck = []
    attempts = 0
    max_attempts = 60
    # diagnostics
    skip_too_long = 0
    skip_too_small = 0
    skip_insufficient_opts = 0
    skip_missing_correct = 0
    skip_other = 0

    # Fallback strategy:
    # - attempts <= simple_switch: use full prompt for remaining items
    # - simple_switch < attempts <= single_item_switch: use simple prompt requesting remaining items
    # - attempts > single_item_switch: request single simple items repeatedly to try fill slots one-by-one
    simple_switch = 5
    single_item_switch = 18
    while len(deck) < n and attempts < max_attempts:
        attempts += 1
        print(f"   Task card attempt {attempts}, have {len(deck)}/{n} cards...")
        if attempts > single_item_switch:
            prompt = p_mcq_simple(topic, note, 1)
        elif attempts > simple_switch:
            rem = n - len(deck)
            prompt = p_mcq_simple(topic, note, rem)
        else:
            prompt = p_mcq(topic, note, n-len(deck))

        items = get_json(prompt)
        if not items:
            print(f"   ⚠️  No items returned from API on attempt {attempts}")
            continue

        short_title = strip_curriculum_code(topic).upper()
        for itm in items:
            if not isinstance(itm, dict):
                continue
            q   = clean(itm.get("q", ""))
            ok  = clean(itm.get("correct", ""))
            ds  = [clean(d) for d in itm.get("distractors", [])][:3]
            exp = clean(itm.get("explanation", ""))
            opts = ds + [ok]

            # Ensure 4 options
            if len(opts) < 4:
                skip_insufficient_opts += 1
                continue
            random.shuffle(opts)
            if ok not in opts:
                skip_missing_correct += 1
                continue

            total_chars = len(q) + sum(len(o) for o in opts)
            if total_chars > MAX_CARD_CHARS:
                skip_too_long += 1
                continue
            if body_font(total_chars) == 0:
                skip_too_small += 1
                continue

            num = len(deck) + 1
            title = short_title
            deck.append((title, q, opts, "ABCD"[opts.index(ok)], exp, num))
            if len(deck) == n:
                break

    if len(deck) < n:
        print(f"   ⚠️  Warning: Only generated {len(deck)}/{n} task cards after {attempts} attempts")
        print(f"     Skipped: too_long={skip_too_long}, too_small_font={skip_too_small}, insufficient_opts={skip_insufficient_opts}, missing_correct={skip_missing_correct}, other={skip_other}")
    return deck

def p_tf(topic, note, n):
   # Day 1 - Concept Check: True or False
   # Cognitive Depth: Basic recall & misconception check
   # Bloom's Level: Remember / Understand
   return (
       f"Write EXACTLY {n} True/False statements for Day 1 - Concept Check on: {topic}. "
       f"Focus on basic recall and misconception checks. Target Bloom's levels: Remember/Understand. "
       f"Ensure alignment with the {CURRICULUM_NAME} ({GRADE_LEVELS}) and the teacher note.\n"
       f"Teacher note: {note}\n"
       'Return JSON list: {"statement":"","answer":true/false,"explanation":""}\n'
       "If answer is false, give ≤15-word explanation, else \"\". ≤325 chars item."
   )

def p_sa(topic, note, n):
   # Day 4 - Critical Thinking: Short Response Questions
   # Cognitive Depth: Explain, connect, apply reasoning
   # Bloom's Level: Apply / Analyze
   return (
       f"Write EXACTLY {n} short-answer questions for Day 4 - Critical Thinking: Short Response on: {topic}. "
       f"Focus on explaining, connecting concepts, and applying reasoning. Target Bloom's levels: Apply/Analyze. "
       f"Ensure each question aligns with the {CURRICULUM_NAME} ({GRADE_LEVELS}) and the teacher note.\n"
       f"Teacher note: {note}\n"
       'Return JSON list: {"q":"","answer":""}\n'
       "Question ≤250 chars, answer ≤25 words."
   )


def p_mcq_simple(topic, note, n):
    """A simplified MCQ prompt used as a fallback to quickly fill remaining slots.
    Requests short stems and concise options so the model returns brief items.
    """
    return (
        f"Write EXACTLY {n} SHORT and SIMPLE MCQs on: {topic}. "
        f"Use concise stems (≤100 chars) and short options (≤40 chars). Keep language direct and avoid multi-part scenarios. "
        f"Return JSON list: {{\"q\":\"\",\"correct\":\"\",\"distractors\":[\"\",\"\",\"\"],\"explanation\":\"\"}}. Randomise order."
    )


def p_tf_simple(topic, note, n):
    """Simplified True/False prompt for fallback cases."""
    return (
        f"Write EXACTLY {n} SHORT True/False statements on: {topic}. "
        f"Make each statement direct and concise (≤120 chars). Return JSON list: {{\"statement\":\"\",\"answer\":true/false,\"explanation\":\"\"}}."
    )


def p_sa_simple(topic, note, n):
    """Simplified Short-Answer prompt for fallback cases."""
    return (
        f"Write EXACTLY {n} SHORT-answer questions on: {topic}. "
        f"Keep questions concise (≤120 chars) and targeted. Return JSON list: {{\"q\":\"\",\"answer\":\"\"}}."
    )


# ───────────────────  BUILDERS  ──────────────────────────────────
def build_mcq(topic, note, n=25):
    deck = []
    max_attempts = 20  # allow more retries for larger batches
    attempts = 0
    simple_switch = 5
    single_item_switch = 12
    while len(deck) < n and attempts < max_attempts:
        attempts += 1
        remaining = n - len(deck)
        print(f"   MCQ attempt {attempts}, have {len(deck)}/{n} questions...")
        if attempts > single_item_switch:
            prompt = p_mcq_simple(topic, note, 1)
        elif attempts > simple_switch:
            prompt = p_mcq_simple(topic, note, remaining)
        else:
            prompt = p_mcq(topic, note, remaining)

        items = get_json(prompt)
        if not items:
            print(f"   ⚠️  No items returned from API on attempt {attempts}")
            continue
        for itm in items:
            if not isinstance(itm, dict):
                continue
            q   = clean(itm.get("q", ""))
            ok  = clean(itm.get("correct", ""))
            ds  = [clean(d) for d in itm.get("distractors", [])]
            exp = clean(itm.get("explanation", ""))
            opts = ds + [ok]
            random.shuffle(opts)
            if body_font(len(q) + sum(len(o) for o in opts)) == 0:
                continue
            if ok not in opts:
                continue
            deck.append((q, opts, "ABCD"[opts.index(ok)], exp))
            if len(deck) == n:
                break
    if len(deck) < n:
        print(f"   ⚠️  Warning: Only generated {len(deck)}/{n} MCQ questions after {attempts} attempts")
    return deck

def p_fill(topic, note, n):
    # Day 3 - Vocabulary Practice: Fill-in-the-Blanks
    # Cognitive Depth: Vocabulary recall & key terms
    # Bloom's Level: Remember / Understand
    return (
        f"Write EXACTLY {n} fill-in-the-blank questions for Day 3 - Vocabulary Practice on: {topic}. "
        f"Focus on vocabulary recall and key terms. Target Bloom's levels: Remember/Understand. "
        f"Ensure each question aligns with the {CURRICULUM_NAME} ({GRADE_LEVELS}) and the teacher note.\n"
        f"Teacher note: {note}\n"
        "Each question MUST include a blank shown as '____' in the stem. "
        'Return JSON list: {"q":"","answer":"","distractors":["","",""],"explanation":""}\n'
        "≤325 chars total per item. Provide a single-word or short-phrase answer and three plausible distractors."
    )

def p_essay(topic, note, n):
    return (
        f"Write EXACTLY {n} essay-style prompts for: {topic}. "
        f"Ensure each prompt aligns with the {CURRICULUM_NAME} ({GRADE_LEVELS}) and the teacher note.\n"
        f"Teacher note: {note}\n"
        'Return JSON list: {"prompt":"","guidance":"","sample_points":""}\n'
        "Prompt ≤400 chars. Guidance ≤60 words. Provide 3–5 sample marking points in a single string."
    )

def bool_to_str(val):
   if isinstance(val, bool): return "True" if val else "False"
   if isinstance(val, str):  return val.strip().capitalize()
   return ""

def build_tf(topic, note, n=25):
    deck = []
    max_attempts = 20
    attempts = 0
    simple_switch = 5
    single_item_switch = 12
    while len(deck) < n and attempts < max_attempts:
        attempts += 1
        remaining = n - len(deck)
        print(f"   T/F attempt {attempts}, have {len(deck)}/{n} questions...")
        if attempts > single_item_switch:
            prompt = p_tf_simple(topic, note, 1)
        elif attempts > simple_switch:
            prompt = p_tf_simple(topic, note, remaining)
        else:
            prompt = p_tf(topic, note, remaining)
        items = get_json(prompt)
        if not items:
            print(f"   ⚠️  No items returned from API on attempt {attempts}")
            continue
        for itm in items:
            if not isinstance(itm, dict):
                continue
            stmt = clean(itm.get("statement", ""))
            ans  = bool_to_str(itm.get("answer", ""))
            exp  = clean(itm.get("explanation", ""))
            if ans not in ("True", "False"):
                continue
            if body_font(len(stmt)+5) == 0:
                continue
            deck.append((stmt, ans, exp))
            if len(deck) == n:
                break
    if len(deck) < n:
        print(f"   ⚠️  Warning: Only generated {len(deck)}/{n} T/F questions after {attempts} attempts")
    return deck

def build_sa(topic, note, n=25):
    deck = []
    max_attempts = 20
    attempts = 0
    simple_switch = 5
    single_item_switch = 12
    while len(deck) < n and attempts < max_attempts:
        attempts += 1
        remaining = n - len(deck)
        print(f"   Short Answer attempt {attempts}, have {len(deck)}/{n} questions...")
        if attempts > single_item_switch:
            prompt = p_sa_simple(topic, note, 1)
        elif attempts > simple_switch:
            prompt = p_sa_simple(topic, note, remaining)
        else:
            prompt = p_sa(topic, note, remaining)
        items = get_json(prompt)
        if not items:
            print(f"   ⚠️  No items returned from API on attempt {attempts}")
            continue
        for itm in items:
            if not isinstance(itm, dict):
                continue
            q   = clean(itm.get("q", ""))
            ans = clean(itm.get("answer", ""))
            if body_font(len(q)+len(ans)) == 0:
                continue
            deck.append((q, ans))
            if len(deck) == n:
                break
    if len(deck) < n:
        print(f"   ⚠️  Warning: Only generated {len(deck)}/{n} Short Answer questions after {attempts} attempts")
    return deck

def build_fill(topic, note):
    deck = []
    max_attempts = 10
    attempts = 0
    while len(deck) < 10 and attempts < max_attempts:
        attempts += 1
        print(f"   Fill-in-Blank attempt {attempts}, have {len(deck)}/10 questions...")
        items = get_json(p_fill(topic, note, 10-len(deck)))
        if not items:
            print(f"   ⚠️  No items returned from API on attempt {attempts}")
            continue
        for itm in items:
            if not isinstance(itm, dict):
                continue
            q   = clean(itm.get("q", ""))
            ans = clean(itm.get("answer", ""))
            ds  = [clean(d) for d in itm.get("distractors", [])][:3]
            exp = clean(itm.get("explanation", ""))
            opts = ds + [ans]
            if len(opts) < 4:
                # skip incomplete items
                continue
            # Ensure the question stem contains an explicit blank '____'
            if '____' not in q:
                if ans and ans.lower() in q.lower():
                    # replace first occurrence of the answer with a blank (preserve spacing)
                    pattern = re.compile(re.escape(ans), re.I)
                    q = pattern.sub('____', q, count=1)
                # else: Don't add ____ at the end - let AI provide proper format
            random.shuffle(opts)
            if body_font(len(q)+sum(len(o) for o in opts)) == 0:
                continue
            deck.append((q, opts, "ABCD"[opts.index(ans)], exp))
            if len(deck) == 10:
                break
    if len(deck) < 10:
        print(f"   ⚠️  Warning: Only generated {len(deck)}/10 Fill-in-Blank questions after {attempts} attempts")
    return deck


def p_scenario(topic, note, n):
    return (
        f"Write EXACTLY {n} scenario-based short-answer questions for: {topic}. "
        f"Each item should present a short realistic scenario and ask a focused question students can answer in 1–2 sentences. "
        f"Align with the {CURRICULUM_NAME} ({GRADE_LEVELS}). Teacher note: {note}\n"
        'Return JSON list: {"q":"","answer":"","explanation":""}\n'
        "Question ≤300 chars; answer ≤40 words."
    )


def build_scenario(topic, note):
    deck = []
    max_attempts = 10
    attempts = 0
    while len(deck) < 10 and attempts < max_attempts:
        attempts += 1
        print(f"   Scenario attempt {attempts}, have {len(deck)}/10 questions...")
        items = get_json(p_scenario(topic, note, 10-len(deck)))
        if not items:
            print(f"   ⚠️  No items returned from API on attempt {attempts}")
            continue
        for itm in items:
            if not isinstance(itm, dict):
                continue
            q   = clean(itm.get("q", ""))
            ans = clean(itm.get("answer", ""))
            exp = clean(itm.get("explanation", ""))
            if body_font(len(q)+len(ans)) == 0:
                continue
            deck.append((q, ans, exp))
            if len(deck) == 10:
                break
    if len(deck) < 10:
        print(f"   ⚠️  Warning: Only generated {len(deck)}/10 Scenario questions after {attempts} attempts")
    return deck

def build_essay(topic, note):
    deck = []
    while len(deck) < 10:
        for itm in get_json(p_essay(topic, note, 10-len(deck))):
            if not isinstance(itm, dict):
                continue
            prompt = clean(itm.get("prompt", ""))
            guide  = clean(itm.get("guidance", ""))
            points = clean(itm.get("sample_points", ""))
            if body_font(len(prompt)+len(guide)) == 0:
                continue
            deck.append((prompt, guide, points))
            if len(deck) == 10:
                break
    return deck


# ───────────────────  MASTER REVIEW GENERATION  ──────────────────
def build_master_review(topic, note):
    """Generate Day 5 Master Review combining all question types (5 each).
    Returns: (tf_deck, mcq_deck, fill_deck, sa_deck)
    Each deck contains 5 questions.
    """
    print(f"   Generating Master Review (5 T/F, 5 MCQ, 5 Fill, 5 SA)...")
    
    # Generate 5 True/False questions
    tf_deck = []
    max_attempts = 5
    attempts = 0
    while len(tf_deck) < 5 and attempts < max_attempts:
        attempts += 1
        items = get_json(p_tf(topic, note, 5-len(tf_deck)))
        for itm in items:
            if not isinstance(itm, dict):
                continue
            stmt = clean(itm.get("statement", ""))
            ans  = bool_to_str(itm.get("answer", ""))
            exp  = clean(itm.get("explanation", ""))
            if ans not in ("True", "False"):
                continue
            if body_font(len(stmt)+5) == 0:
                continue
            tf_deck.append((stmt, ans, exp))
            if len(tf_deck) == 5:
                break
    
    # Generate 5 Multiple Choice questions
    mcq_deck = []
    attempts = 0
    while len(mcq_deck) < 5 and attempts < max_attempts:
        attempts += 1
        items = get_json(p_mcq(topic, note, 5-len(mcq_deck)))
        for itm in items:
            if not isinstance(itm, dict):
                continue
            q   = clean(itm.get("q", ""))
            ok  = clean(itm.get("correct", ""))
            ds  = [clean(d) for d in itm.get("distractors", [])]
            exp = clean(itm.get("explanation", ""))
            opts = ds + [ok]
            random.shuffle(opts)
            if body_font(len(q) + sum(len(o) for o in opts)) == 0:
                continue
            if ok not in opts:
                continue
            mcq_deck.append((q, opts, "ABCD"[opts.index(ok)], exp))
            if len(mcq_deck) == 5:
                break
    
    # Generate 5 Fill-in-Blank questions
    fill_deck = []
    attempts = 0
    while len(fill_deck) < 5 and attempts < max_attempts:
        attempts += 1
        items = get_json(p_fill(topic, note, 5-len(fill_deck)))
        for itm in items:
            if not isinstance(itm, dict):
                continue
            q   = clean(itm.get("q", ""))
            ans = clean(itm.get("answer", ""))
            ds  = [clean(d) for d in itm.get("distractors", [])][:3]
            exp = clean(itm.get("explanation", ""))
            opts = ds + [ans]
            if len(opts) < 4:
                continue
            if '____' not in q:
                if ans and ans.lower() in q.lower():
                    pattern = re.compile(re.escape(ans), re.I)
                    q = pattern.sub('____', q, count=1)
                # else: Don't add ____ at the end - let AI provide proper format
            random.shuffle(opts)
            if body_font(len(q)+sum(len(o) for o in opts)) == 0:
                continue
            fill_deck.append((q, opts, "ABCD"[opts.index(ans)], exp))
            if len(fill_deck) == 5:
                break
    
    # Generate 5 Short Answer questions
    sa_deck = []
    attempts = 0
    while len(sa_deck) < 5 and attempts < max_attempts:
        attempts += 1
        items = get_json(p_sa(topic, note, 5-len(sa_deck)))
        for itm in items:
            if not isinstance(itm, dict):
                continue
            q   = clean(itm.get("q", ""))
            ans = clean(itm.get("answer", ""))
            if body_font(len(q)+len(ans)) == 0:
                continue
            sa_deck.append((q, ans))
            if len(sa_deck) == 5:
                break
    
    # Normalize all decks to exactly 5 items
    tf_deck = normalize_deck(tf_deck, expected=5)
    mcq_deck = normalize_deck(mcq_deck, expected=5)
    fill_deck = normalize_deck(fill_deck, expected=5)
    sa_deck = normalize_deck(sa_deck, expected=5)
    
    return tf_deck, mcq_deck, fill_deck, sa_deck


# ───────────────────  PDF HELPERS  ───────────────────────────────
def doc(path):
   return SimpleDocTemplate(str(path), pagesize=letter,
                            leftMargin=35, rightMargin=35,
                            topMargin=50,  bottomMargin=40)

def sa_lines():  # TWO lines, 85 underscores each
    line = "_" * 85
    return [Paragraph(line, ST["Line"]) for _ in range(2)]


def normalize_deck(deck, expected=10):
    """Trim or pad a deck to the expected length. If deck is shorter, pad with placeholders.
    Return a new list of length exactly expected.
    """
    if not isinstance(deck, list):
        return deck
    if len(deck) > expected:
        return deck[:expected]
    # If too short, pad with empty items depending on item shape
    out = list(deck)
    # determine shape
    if not out:
        # If caller provided an empty deck, return a padded list of empty placeholders
        # so downstream PDF builders still produce files and do not skip titles.
        # We'll assume a common MCQ-like shape (q, opts, letter, exp) when unknown.
        # Individual callers may replace these with more suitable placeholders if needed.
        placeholder = ("", ["", "", "", ""], "A", "")
        return [placeholder for _ in range(expected)]
    sample = out[0]
    while len(out) < expected:
        if isinstance(sample, tuple) and len(sample) == 4:
            out.append(("", ["", "", "", ""], "A", ""))
        elif isinstance(sample, tuple) and len(sample) == 3:
            out.append(("", "A", ""))
        elif isinstance(sample, tuple) and len(sample) == 2:
            out.append(("", ""))
        else:
            out.append(sample)
    return out


def pad_task_cards(deck, expected=30):
    """Ensure task card deck has exactly `expected` items. Task cards use a 6-tuple shape:
    (title, q, opts, letter, exp, num). Pad with empty placeholders where necessary.
    """
    out = list(deck) if isinstance(deck, list) else []
    # Create placeholder template
    while len(out) < expected:
        num = len(out) + 1
        out.append(("", "", ["", "", "", ""], "A", "", num))
    if len(out) > expected:
        return out[:expected]
    return out


# ───────────────────  WORKSHEET MAKERS  ──────────────────────────
def make_mcq(ws_pdf, ans_pdf, main, sub, deck):
    # Reduced spacing (previously two spacers totaling 18pt). Now ~9pt before title.
    story=[Spacer(1,6), Paragraph(f"<b>{sub}</b>", ST["Doc"]), Spacer(1,8)]
    for i,(q,opts,_,_) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {q}", ST["Q"])] + \
             [Paragraph(f"{l}. {t}", ST["Opt"]) for l,t in zip("ABCD",opts)] + \
             [Spacer(1,12)]
        story.append(KeepTogether(blk))
    doc(ws_pdf).build(story, onFirstPage=first, onLaterPages=later)

    hdr=f"{sub} - Answer Sheet"
    story=[Paragraph(hdr, ST["AnsH"]), Spacer(1,10)]
    opt_indent = getattr(ST["Opt"], 'leftIndent', 0)
    for i,(q,opts,ans,exp) in enumerate(deck,1):
        story.append(Paragraph(f"{i}. {q}", ST["Q"]))
        # Render each option; show each option once. For the correct option, render as an underlined line
        for l,t in zip("ABCD",opts):
            text = f"{l}. {t}"
            if l == ans:
                # Put the underlined answer aligned with ST['Opt']
                story.append(UnderlinedAnswer(text, fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=opt_indent))
            else:
                story.append(Paragraph(text, ST["Opt"]))
        if exp:
            story.append(Paragraph(f"Explanation: {exp}", ST["Expl"]))
        story.append(Spacer(1,8))
    doc(ans_pdf).build(story, onFirstPage=first, onLaterPages=later)

def make_tf(ws_pdf, ans_pdf, main, sub, deck):
    story = [Spacer(1,6), Paragraph(f"<b>{sub} – True/False</b>", ST["Doc"]), Spacer(1,8)]
    for i, (stmt, _, _) in enumerate(deck, 1):
        blk = [Paragraph(f"{i}. {stmt}", ST["Q"]),
               Paragraph("A. True", ST["Opt"]),
               Paragraph("B. False", ST["Opt"]),
               Spacer(1,12)]
        story.append(KeepTogether(blk))
    doc(ws_pdf).build(story, onFirstPage=first, onLaterPages=later)

    hdr = f"{sub} - Answer Sheet"
    story = [Paragraph(hdr, ST["AnsH"]), Spacer(1,10)]
    opt_indent = getattr(ST["Opt"], 'leftIndent', 0)
    for i, (stmt, ans, exp) in enumerate(deck, 1):
        story.append(Paragraph(f"{i}. {stmt}", ST["Q"]))
        a_text = "A. True"
        b_text = "B. False"
        if ans == "True":
            story.append(UnderlinedAnswer(a_text, fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=opt_indent))
            story.append(Paragraph(b_text, ST["Opt"]))
        else:
            story.append(Paragraph(a_text, ST["Opt"]))
            story.append(UnderlinedAnswer(b_text, fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=opt_indent))
        if ans == "False" and exp:
            story.append(Paragraph(f"Explanation: {exp}", ST["Expl"]))
        story.append(Spacer(1,8))
    doc(ans_pdf).build(story, onFirstPage=first, onLaterPages=later)

def make_sa(ws_pdf, ans_pdf, main, sub, deck):
    story = [Spacer(1,6), Paragraph(f"<b>{sub} – Short Answer</b>", ST["Doc"]), Spacer(1,8)]
    for i, (q, _) in enumerate(deck, 1):
        blk = [Paragraph(f"{i}. {q}", ST["Q"])] + sa_lines() + [Spacer(1,12)]
        story.append(KeepTogether(blk))
    doc(ws_pdf).build(story, onFirstPage=first, onLaterPages=later)

    hdr = f"{sub} - Answer Sheet"
    story = [Paragraph(hdr, ST["AnsH"]), Spacer(1,10)]
    for i, (q, ans) in enumerate(deck, 1):
        story.append(Paragraph(f"{i}. {q}", ST["Q"]))
        # For short-answer answer sheets, show the model answer in blue text (no underline)
        story.append(Paragraph(ans, ST["Blue"]))
        story.append(Spacer(1,8))
    doc(ans_pdf).build(story, onFirstPage=first, onLaterPages=later)
def make_fill(ws_pdf, ans_pdf, main, sub, deck):
    # Fill-in-the-blank worksheet (render as 4-option MCQ) with reduced top spacing
    story = [Spacer(1,6), Paragraph(f"<b>{sub} – Fill in the Blank</b>", ST["Doc"]), Spacer(1,8)]
    for i, (q, opts, _, _) in enumerate(deck, 1):
        blk = [Paragraph(f"{i}. {q}", ST["Q"])] + [Paragraph(f"{l}. {t}", ST["Opt"]) for l,t in zip("ABCD", opts)] + [Spacer(1,12)]
        story.append(KeepTogether(blk))
    doc(ws_pdf).build(story, onFirstPage=first, onLaterPages=later)

    # Answer sheet with explanation
    hdr = f"{sub} - Answer Sheet"
    story = [Paragraph(hdr, ST["AnsH"]), Spacer(1,10)]
    opt_indent = getattr(ST["Opt"], 'leftIndent', 0)
    for i, (q, opts, ans, exp) in enumerate(deck, 1):
        story.append(Paragraph(f"{i}. {q}", ST["Q"]))
        for l, t in zip("ABCD", opts):
            text = f"{l}. {t}"
            if l == ans:
                story.append(UnderlinedAnswer(text, fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=opt_indent))
            else:
                story.append(Paragraph(text, ST["Opt"]))
        if exp:
            story.append(Paragraph(f"Explanation: {exp}", ST["Expl"]))
        story.append(Spacer(1,8))
    doc(ans_pdf).build(story, onFirstPage=first, onLaterPages=later)


def make_essay(ws_pdf, ans_pdf, main, sub, deck):
    # Deprecated: Essay generation replaced by scenario-based short answer. Keep for backward compatibility but do nothing.
    return


def make_scenario(ws_pdf, ans_pdf, main, sub, deck):
    # Scenario-based short answer worksheet (10 items)
    story=[Paragraph(f"<b>{sub} – Scenario-Based Short Answer</b>", ST["Doc"]), Spacer(1,10)]
    for i,(q,_,_) in enumerate(deck,1):
        blk=[Paragraph(f"{i}. {q}", ST["Q"])] + sa_lines() + [Spacer(1,12)]
        story.append(KeepTogether(blk))
    doc(ws_pdf).build(story, onFirstPage=first, onLaterPages=later)

    # Answer sheet with explanation
    hdr=f"{sub} - Answer Sheet (model answers)"
    story=[Paragraph(hdr, ST["AnsH"]), Spacer(1,10)]
    for i,(q,ans,exp) in enumerate(deck,1):
        story.append(Paragraph(f"{i}. {q}", ST["Q"]))
        story.append(UnderlinedAnswer(ans, fontName=BODY_FONT, fontSize=12, lineColor=colors.blue))
        if exp:
            story.append(Paragraph(f"Explanation: {exp}", ST["Expl"]))
        story.append(Spacer(1,8))
        
        # ...keep entries sequentially
    doc(ans_pdf).build(story, onFirstPage=first, onLaterPages=later)


def make_task_cards_pdf(cards, out_pdf, answer_pdf, main, sub, preview=False):
    """Render task cards (2x2 grid per page) and create a separate answer sheet.
    cards: list of (title, q, opts, correct_letter, explanation)
    """
    # Card geometry: US Letter, margins 0.5in => 36pt
    margin = 36
    gutter = 12
    usable_w = PAGE_W - 2 * margin
    usable_h = PAGE_H - 2 * margin
    cols = 2
    rows = 3
    card_w = (usable_w - gutter * (cols - 1)) / cols
    card_h = (usable_h - gutter * (rows - 1)) / rows

    c = canvas.Canvas(str(out_pdf), pagesize=letter)
    c.setTitle(f"{sub} - Task Cards")

    # Draw cards
    per_page = cols * rows
    for idx, card in enumerate(cards):
        page_idx = idx // per_page
        pos_in_page = idx % per_page
        if pos_in_page == 0:
            # start of a new page (including first). If not the first, finish prior page first
            if idx != 0:
                c.showPage()
            # If rendering a preview copy, draw the PREVIEW watermark big & light on each page
            if preview:
                c.saveState()
                c.setFont(BODY_FONT, 130)
                c.setFillColor(colors.lightgrey)
                c.translate(PAGE_W/2, PAGE_H/2)
                c.rotate(45)
                c.drawCentredString(0, 0, "PREVIEW")
                c.restoreState()
        # Draw page-level dashed cut guides (once per page)
        if pos_in_page == 0:
            c.setStrokeColor(colors.black)
            c.setDash(3,2)
            # vertical separator between cols
            gx = margin + card_w + (gutter / 2)
            c.line(gx, margin + 6, gx, PAGE_H - margin - 6)
            # horizontal separators between rows (mid-gutter positions)
            for r in range(rows - 1):
                gy = PAGE_H - margin - (r+1) * card_h - r * gutter - (gutter / 2)
                c.line(margin + 6, gy, PAGE_W - margin - 6, gy)
            c.setDash()
        col = pos_in_page % cols
        row = pos_in_page // cols
        x = margin + col * (card_w + gutter)
        # rows counted from top
        y = PAGE_H - margin - (row + 1) * card_h - row * gutter + (card_h - card_h)

        # card may be (title,q,opts,ans,exp,num) or legacy 5-tuple
        if len(card) == 6:
            title, q, opts, ans, exp, num = card
        else:
            title, q, opts, ans, exp = card
            num = None
        # Header area: No filled black box per new spec. Draw title text in black bold and
        # add a horizontal divider line to separate the title bar from the main content.
        hdr_h = 0.8 * 72  # 0.8 inches
        hdr_y = y + card_h - hdr_h
        # Title text (black, bold) - prefer a bold variant of TITLE_FONT if registered
        c.setFillColor(colors.black)
        # Try several fallbacks to ensure a bold font is used
        bold_font_candidates = [f"{TITLE_FONT}-Bold", f"{TITLE_FONT} Bold", "Helvetica-Bold", "Times-Bold", "Courier-Bold", TITLE_FONT, BODY_FONT]
        for bf in bold_font_candidates:
            try:
                pdfmetrics.getFont(bf)
                c.setFont(bf, 12)
                break
            except Exception:
                try:
                    # Some font registrations accept the family name; try setting anyway
                    c.setFont(bf, 12)
                    break
                except Exception:
                    continue

        title_text = title.upper()
        # Wrap title into up to 3 lines (same logic as before)
        hdr_lines = []
        words = title_text.split()
        cur = ''
        for w in words:
            test = (cur + ' ' + w).strip()
            if pdfmetrics.stringWidth(test, TITLE_FONT, 11) <= (card_w - 16):
                cur = test
            else:
                hdr_lines.append(cur)
                cur = w
        if cur:
            hdr_lines.append(cur)
        hdr_lines = hdr_lines[:3]
        start_y = hdr_y + hdr_h - (hdr_h / 2) + (6 * (len(hdr_lines)-1))
        ty = start_y
        # Draw lines in bold by using the title font (already set)
        for ln in hdr_lines:
            c.drawCentredString(x + card_w/2, ty, ln)
            ty -= 12

        # (Divider removed from task cards per request.)

        # Draw card border with rounded corners
        c.setStrokeColor(colors.black)
        c.setLineWidth(1)
        try:
            c.roundRect(x, y, card_w, card_h, CARD_CORNER_RADIUS)
        except Exception:
            # Fallback to square corners if roundRect is not available in this ReportLab version
            c.rect(x, y, card_w, card_h)

    # Body: question and options
        body_x = x + 8
        body_y = y + card_h - hdr_h - 18
        c.setFillColor(colors.black)
        # choose font size based on length
        total_chars = len(q) + sum(len(o) for o in opts)
        fsize = body_font(total_chars)
        if fsize == 0:
            # skip drawing this card (shouldn't happen since generation checks length)
            continue
        c.setFont(BODY_FONT, fsize)
        # Prepare question text: if num provided, prefix number before the question
        if num is not None:
            q_text = f"{num:02d}. {q}"
        else:
            q_text = q
        # Draw question (naive wrapping)
        max_width = card_w - 16
        lines = []
        for part in q_text.split('\n'):
            words = part.split()
            cur = ''
            for w in words:
                test = (cur + ' ' + w).strip()
                if pdfmetrics.stringWidth(test, BODY_FONT, fsize) <= max_width:
                    cur = test
                else:
                    lines.append(cur)
                    cur = w
            if cur:
                lines.append(cur)
        ty = body_y
        for ln in lines:
            c.drawString(body_x, ty, ln)
            ty -= fsize + 2
        ty -= fsize  # blank line
        # Options
        for label, opt in zip('ABCD', opts):
            text = f"{label}. {opt}"
            # wrap as needed
            cur = ''
            words = text.split()
            for w in words:
                test = (cur + ' ' + w).strip()
                if pdfmetrics.stringWidth(test, BODY_FONT, fsize) <= max_width:
                    cur = test
                else:
                    c.drawString(body_x, ty, cur)
                    ty -= fsize + 2
                    cur = w
            if cur:
                c.drawString(body_x, ty, cur)
                ty -= fsize + 2

        # (Removed per-card dashed cut guides here to avoid cut guides drawing inside task cards.)

    c.showPage()
    c.save()

    # Answer sheet
    story = [Paragraph(f"{sub} - Task Cards Answer Sheet", ST['AnsH']), Spacer(1,12)]
    for i, card in enumerate(cards, 1):
        # card might be 6-tuple
        if len(card) == 6:
            title, q, opts, ans, exp, num = card
        else:
            title, q, opts, ans, exp = card
            num = i
        # show question stem for teacher reference (numbered)
        story.append(Paragraph(f"{num:02d}. {q}", ST['Q']))
        # find option text for the correct letter and print without question number
        letter_index = ord(ans) - ord('A') if isinstance(ans, str) and ans in 'ABCD' else 0
        correct_text = opts[letter_index]
        story.append(Paragraph(f"{ans}. {correct_text}", ST['Blue']))
        story.append(Spacer(1,6))
    # If this is a preview version, use the preview watermark page handler for the answer sheet
    if preview:
        doc(answer_pdf).build(story, onFirstPage=preview_page, onLaterPages=preview_page)
    else:
        doc(answer_pdf).build(story, onFirstPage=first, onLaterPages=later)


def make_task_cards_intro_page(path, main, sub, count=30):
    """Create a single preview page that announces the Task Cards section.
    The page contains two centered lines:
      30 Task Cards
      Includes Answer Key with Explanations
    Built without the large PREVIEW watermark (so it reads cleanly inside the preview).
    """
    # Use the existing TitlePage Flowable for consistent centering and font handling
    line_specs = [ (f"{count} Task Cards", TITLE_FONT, 36), ("Includes Answer Key with Explanations", BODY_FONT, 16) ]
    tp = TitlePage(line_specs)
    docp = SimpleDocTemplate(str(path), pagesize=letter, leftMargin=35, rightMargin=35, topMargin=50, bottomMargin=40)
    docp.build([tp], onFirstPage=preview_page_no_watermark, onLaterPages=preview_page_no_watermark)




def make_master_review(ws_pdf, ans_pdf, main, sub, tf_deck, mcq_deck, fill_deck, sa_deck):
    """Generate Day 5 Master Review combining all question types."""
    # Worksheet
    story = [Spacer(1,6), Paragraph(f"<b>{sub} – Master Review</b>", ST["Doc"]), Spacer(1,8)]
    
    # Part A: True/False (5 questions)
    story.append(Paragraph("<b>Part A – True or False</b>", ST["AnsH"]))
    for i, (stmt, _, _) in enumerate(tf_deck, 1):
        blk = [Paragraph(f"{i}. {stmt}", ST["Q"]),
               Paragraph("A. True", ST["Opt"]),
               Paragraph("B. False", ST["Opt"]),
               Spacer(1,12)]
        story.append(KeepTogether(blk))
    
    story.append(Spacer(1, 20))
    
    # Part B: Multiple Choice (5 questions)
    story.append(Paragraph("<b>Part B – Multiple Choice</b>", ST["AnsH"]))
    for i, (q, opts, _, _) in enumerate(mcq_deck, 6):  # Continue numbering from 6
        blk = [Paragraph(f"{i}. {q}", ST["Q"])] + \
              [Paragraph(f"{l}. {t}", ST["Opt"]) for l, t in zip("ABCD", opts)] + \
              [Spacer(1,12)]
        story.append(KeepTogether(blk))
    
    story.append(Spacer(1, 20))
    
    # Part C: Fill-in-the-Blank (5 questions)
    story.append(Paragraph("<b>Part C – Fill-in-the-Blank</b>", ST["AnsH"]))
    for i, (q, opts, _, _) in enumerate(fill_deck, 11):  # Continue from 11
        blk = [Paragraph(f"{i}. {q}", ST["Q"])] + \
              [Paragraph(f"{l}. {t}", ST["Opt"]) for l, t in zip("ABCD", opts)] + \
              [Spacer(1,12)]
        story.append(KeepTogether(blk))
    
    story.append(Spacer(1, 20))
    
    # Part D: Short Answer (5 questions)
    story.append(Paragraph("<b>Part D – Short Answer</b>", ST["AnsH"]))
    for i, (q, _) in enumerate(sa_deck, 16):  # Continue from 16
        blk = [Paragraph(f"{i}. {q}", ST["Q"])] + sa_lines() + [Spacer(1,12)]
        story.append(KeepTogether(blk))
    
    doc(ws_pdf).build(story, onFirstPage=first, onLaterPages=later)
    
    # Answer Sheet
    hdr = f"{sub} - Answer Sheet"
    story = [Paragraph(hdr, ST["AnsH"]), Spacer(1,10)]
    opt_indent = getattr(ST["Opt"], 'leftIndent', 0)
    
    # Part A Answers
    story.append(Paragraph("<b>Part A – True or False</b>", ST["AnsH"]))
    for i, (stmt, ans, exp) in enumerate(tf_deck, 1):
        blk = [Paragraph(f"{i}. {stmt}", ST["Q"])]
        if ans == "True":
            blk.append(UnderlinedAnswer("A. True", fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=opt_indent))
            blk.append(Paragraph("B. False", ST["Opt"]))
        else:
            blk.append(Paragraph("A. True", ST["Opt"]))
            blk.append(UnderlinedAnswer("B. False", fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=opt_indent))
        if ans == "False" and exp:
            blk.append(Paragraph(f"Explanation: {exp}", ST["Expl"]))
        blk.append(Spacer(1,8))
        story.append(KeepTogether(blk))
    
    story.append(Spacer(1, 15))
    
    # Part B Answers
    story.append(Paragraph("<b>Part B – Multiple Choice</b>", ST["AnsH"]))
    for i, (q, opts, ans, exp) in enumerate(mcq_deck, 6):
        blk = [Paragraph(f"{i}. {q}", ST["Q"])]
        for l, t in zip("ABCD", opts):
            text = f"{l}. {t}"
            if l == ans:
                blk.append(UnderlinedAnswer(text, fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=opt_indent))
            else:
                blk.append(Paragraph(text, ST["Opt"]))
        blk.append(Paragraph(f"Explanation: {exp}", ST["Expl"]))
        blk.append(Spacer(1,8))
        story.append(KeepTogether(blk))
    
    story.append(Spacer(1, 15))
    
    # Part C Answers
    story.append(Paragraph("<b>Part C – Fill-in-the-Blank</b>", ST["AnsH"]))
    for i, (q, opts, ans, exp) in enumerate(fill_deck, 11):
        blk = [Paragraph(f"{i}. {q}", ST["Q"])]
        for l, t in zip("ABCD", opts):
            text = f"{l}. {t}"
            if l == ans:
                blk.append(UnderlinedAnswer(text, fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=opt_indent))
            else:
                blk.append(Paragraph(text, ST["Opt"]))
        if exp:
            blk.append(Paragraph(f"Explanation: {exp}", ST["Expl"]))
        blk.append(Spacer(1,8))
        story.append(KeepTogether(blk))
    
    story.append(Spacer(1, 15))
    
    # Part D Answers
    story.append(Paragraph("<b>Part D – Short Answer</b>", ST["AnsH"]))
    for i, (q, ans) in enumerate(sa_deck, 16):
        story.append(Paragraph(f"{i}. {q}", ST["Q"]))
        story.append(Paragraph(ans, ST["Blue"]))
        story.append(Spacer(1,8))
    
    doc(ans_pdf).build(story, onFirstPage=first, onLaterPages=later)


# ───────────────────  PREVIEW PDF  ───────────────────────────────
def make_preview(preview_pdf, main, mcq, tf, sa):
   notice = (
       "Here, in the preview, all 3 sections are combined, but you will "
       "receive 3 separate print-ready PDFs:<br/><br/>"
       "1. 30 MCQs<br/>"
       "2. 30 True/False questions<br/>"
       "3. 30 Short Answer type questions."
   )
   story = [
       Paragraph(notice, ST["Note"]),
       Paragraph(f"<b>{main} – COMBINED PREVIEW</b>", ST["Doc"]),
       Spacer(1, 12),

       Paragraph("Section A – MCQ", ST["AnsH"])
   ]

   for i,(q,opts,_,_) in enumerate(mcq,1):
       blk=[Paragraph(f"{i}. {q}", ST["Q"])]
       blk += [Paragraph(f"{l}. {t}", ST["Opt"]) for l,t in zip("ABCD",opts)]
       blk.append(Spacer(1,12))
       story.append(KeepTogether(blk))

   story.append(PageBreak())
   story.append(Paragraph("Section B – True/False", ST["AnsH"]))

   for i,(stmt,_,_) in enumerate(tf,1):
       blk=[Paragraph(f"{i}. {stmt}", ST["Q"]),
            Paragraph("A. True",  ST["Opt"]),
            Paragraph("B. False", ST["Opt"]),
            Spacer(1,12)]
       story.append(KeepTogether(blk))

   story.append(PageBreak())
   story.append(Paragraph("Section C – Short Answer", ST["AnsH"]))

   for i,(q,_) in enumerate(sa,1):
       blk=[Paragraph(f"{i}. {q}", ST["Q"])] + sa_lines() + [Spacer(1,12)]
       story.append(KeepTogether(blk))

   doc(preview_pdf).build(story, onFirstPage=preview_page, onLaterPages=preview_page)


def make_full_preview(preview_path, main, tf_d, mcq_d, fill_d, sa_d, mr_tf=None, mr_mcq=None, mr_fill=None, mr_sa=None, scenario_d=None):
    # Build with two templates: Framed (with watermark) + NoWatermark (border without watermark)
    doc = BaseDocTemplate(
        str(preview_path),
        pagesize=letter,
        leftMargin=35, rightMargin=35, topMargin=50, bottomMargin=40
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin,
                  doc.width, doc.height, id='normal')

    first_framed = PageTemplate(id='FirstFramed', frames=[frame], onPage=first_page_no_watermark)
    framed = PageTemplate(id='Framed', frames=[frame], onPage=preview_page)
    no_watermark = PageTemplate(id='NoWatermark', frames=[frame], onPage=preview_page_no_watermark)

    doc.addPageTemplates([first_framed, framed, no_watermark])

    # Use first page template without watermark for the very first page
    story = [NextPageTemplate('FirstFramed')]
    # Uppercase main title per requirement
    story.append(Paragraph(f"<b>{main.upper()} – COMBINED PREVIEW</b>", ST["Doc"]))
    story.append(Spacer(1,12))
    story.append(NextPageTemplate('Framed'))  # Switch to regular framed after first page

    # Helper: push a centered preview title block (vertically and horizontally centered on the page)
    def push_title(day_title: str, label: str, question_count: int = 10):
        # Calculate vertical centering: page height ~11in=792pt; usable height ~692pt after margins
        # Three lines: 24pt + 16pt + 16pt = 56pt content + ~48pt leading = ~104pt total block height
        # Center: (692 - 104)/2 ≈ 294pt top spacer
        story.append(Spacer(1, 294))
        story.append(Paragraph(f"<b>{day_title}</b>", ST["PreviewTitleMain"]))
        story.append(Paragraph(f"<b>{question_count} {label}</b>", ST["PreviewTitleSub"]))
        story.append(Paragraph("<b>Includes Answer Key with Explanations</b>", ST["PreviewTitleSub"]))
        story.append(PageBreak())

    # Section 01: True/False worksheet then its answer sheet (label updated)
    push_title("True or False WorkSheet", "True or False Questions", 25)
    story.append(Paragraph("True or False WorkSheet    ", ST["AnsH"]))
    for i,(stmt,_,_) in enumerate(tf_d,1):
        blk=[Paragraph(f"{i}. {stmt}", ST["Q"]),
             Paragraph("A. True", ST["Opt"]),
             Paragraph("B. False", ST["Opt"]),
             Spacer(1,12)]
        story.append(KeepTogether(blk))
    story.append(PageBreak())

    story.append(Paragraph("True or False - Answer Sheet    ", ST["AnsH"]))
    opt_indent = getattr(ST["Opt"], 'leftIndent', 0)
    for i,(stmt,ans,exp) in enumerate(tf_d,1):
        blk=[Paragraph(f"{i}. {stmt}", ST["Q"])]
        # True/False: place underlined correct option aligned with ST['Opt']
        if ans == "True":
            blk.append(UnderlinedAnswer("A. True", fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=opt_indent))
            blk.append(Paragraph("B. False", ST["Opt"]))
        else:
            blk.append(Paragraph("A. True", ST["Opt"]))
            blk.append(UnderlinedAnswer("B. False", fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=opt_indent))
        if ans=="False" and exp:
            blk.append(Paragraph(f"Explanation: {exp}", ST["Expl"]))
        blk.append(Spacer(1,8))
        story.append(KeepTogether(blk))
    story.append(PageBreak())

    # Section 02: MCQ
    push_title("Multiple Choice Questions WorkSheet", "Multiple Choice Type Questions", 25)
    story.append(Paragraph("Multiple Choice Questions WorkSheet    ", ST["AnsH"]))
    for i,(q,opts,_,_) in enumerate(mcq_d,1):
        blk=[Paragraph(f"{i}. {q}", ST["Q"])]
        blk += [Paragraph(f"{l}. {t}", ST["Opt"]) for l,t in zip("ABCD",opts)]
        blk.append(Spacer(1,12))
        story.append(KeepTogether(blk))
    story.append(PageBreak())

    story.append(Paragraph("Multiple Choice - Answer Sheet    ", ST["AnsH"]))
    opt_indent = getattr(ST["Opt"], 'leftIndent', 0)
    for i,(q,opts,ans,exp) in enumerate(mcq_d,1):
        blk=[Paragraph(f"{i}. {q}", ST["Q"])]
        for l,t in zip("ABCD",opts):
            text = f"{l}. {t}"
            if l == ans:
                blk.append(UnderlinedAnswer(text, fontName=BODY_FONT, fontSize=12, lineColor=colors.blue, leftIndent=opt_indent))
            else:
                blk.append(Paragraph(text, ST["Opt"]))
        blk.append(Paragraph(f"Explanation: {exp}", ST["Expl"]))
        blk.append(Spacer(1,8))
        story.append(KeepTogether(blk))
    story.append(PageBreak())

    # Section 03: Fill-in-the-Blank removed per new requirements

    # Section 04: Short Answer
    push_title("Short Answer WorkSheet", "Short Answer Type Questions", 25)
    story.append(Paragraph("Short Answer WorkSheet    ", ST["AnsH"]))
    for i,(q,_) in enumerate(sa_d,1):
        blk=[Paragraph(f"{i}. {q}", ST["Q"]) ] + sa_lines() + [Spacer(1,12)]
        story.append(KeepTogether(blk))
    story.append(PageBreak())

    story.append(Paragraph("Short Answer - Answer Sheet    ", ST["AnsH"]))
    for i,(q,ans) in enumerate(sa_d,1):
        story.append(Paragraph(f"{i}. {q}", ST["Q"]))
        # In preview, show the short-answer model answer in blue text (no underline)
        story.append(Paragraph(ans, ST["Blue"]))
        story.append(Spacer(1,8))
    story.append(PageBreak())

    # Scenario-based section removed from preview

    # Build with Framed as default
    doc.build(story)


# ───────────────────  CURRICULUM LOADER  ─────────────────────────────────
class CurriculumData:
    """Stores curriculum metadata and topics"""
    def __init__(self, subject: str, grade: str, curriculum: str):
        self.subject_name = subject
        self.grade_level = grade
        self.curriculum_name = curriculum
        self.units: List[Tuple[int, str, List[Tuple[int, str, str]]]] = []
    
    def get_prompt_context(self) -> str:
        """Returns context string for OpenAI prompts"""
        return (f"Subject: {self.subject_name}\n"
                f"Grade Level: {self.grade_level}\n"
                f"Curriculum: {self.curriculum_name}\n"
                f"IMPORTANT: All questions MUST align with {self.grade_level} standards "
                f"and stay within the {self.curriculum_name} curriculum. "
                f"Do NOT create questions outside these standards.")

def load_curriculum_from_excel(excel_path: str = "details.xlsx") -> List[CurriculumData]:
    """
    Load multiple curricula from Excel file with flexible structure.
    
    Expected Excel structure (can repeat for multiple curricula):
    - Row X: Subject Name - (marker for new curriculum section)
    - Row X+1: Grade level -
    - Row X+2: Curriculum -
    - Row X+3+: Headers and data sections for each unit
    
    Each unit section has:
    - Merged row: Unit subtitle (e.g., "MS-PS1: Matter and Its Interactions")
    - Header row: NO, STANDARD, TITLE, NOTE
    - Data rows: Topic details
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
            print("❌ No 'Subject Name -' markers found. Using legacy single-curriculum format.")
            curriculum_start_rows = [0]  # Assume single curriculum starting at row 0
        
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
                
                # Stop if we encounter another "Subject Name -" marker (shouldn't happen but safety check)
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
                    # Expected columns: NO, STANDARD, TITLE, NOTE
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
                                title = f"{standard} — {title}"
                        
                        # Get note
                        note = str(note_val).strip() if pd.notna(note_val) else ''
                        
                        # Add if valid
                        if title and title != 'nan' and len(title) > 2:
                            current_subtopics.append((sub_idx, title, note))
            
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
        print("   Row 1: Subject Name (e.g., 'Physical Sciences')")
        print("   Row 2: Grade level (e.g., 'Grade 7' or 'Middle School')")
        print("   Row 3: Curriculum (e.g., 'NGSS' or 'Common Core')")
        print("   Row 4+: Unit sections")
        print("   Each unit:")
        print("     - Merged row: Unit title")
        print("     - Header row: NO, STANDARD, TITLE, NOTE")
        print("     - Data rows: Topic details")
        print("     - Empty row: Separator")
        sys.exit(1)

# Initialize curriculum variable (will be loaded from Excel in main())
CURRICULUM_DATA: CurriculumData = None


# ───────────────────  CURRICULUM (Legacy - now loaded from Excel)  ─────────────────────────────────
# Format: (unit_index, unit_title, [ (sub_index, sub_topic_title, teacher_note), … ])
# This is now loaded from details.xlsx instead of being hardcoded
CURRICULUM: List[Tuple[int, str, List[Tuple[int, str, str]]]] = []


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
   print("Available Standards (unit_index – title):")
   for m_idx, main_title, subs in CURRICULUM:
       print(f"  {m_idx}: {main_title}  ({len(subs)} subtopics)")
       for s_idx, sub_title, _ in subs:
           print(f"     {m_idx}.{s_idx}: {sub_title}")

def build_selected(standards: List[int], subs_filter: List[int], dry_run: bool):
    tasks = []
    for m_idx, main_title, subs in CURRICULUM:
        if standards and m_idx not in standards:
            continue
        for s_idx, sub_title, note in subs:
            if subs_filter and s_idx not in subs_filter:
                continue
            tasks.append((m_idx, main_title, s_idx, sub_title, note))

    if dry_run:
        print(f"🛈 Dry-run: would generate {len(tasks)} topic(s).")
        for m_idx, main_title, s_idx, sub_title, _ in tasks[:25]:
            print(f"  {m_idx}.{s_idx} – {sub_title} (in {main_title})")
        if len(tasks) > 25:
            print(f"  … {len(tasks)-25} more")
        return

    if not tasks:
        print("⚠  No matching standards/subtopics.")
        return

    ensure_api_key()
    start = time.time()
    for m_idx, main_title, s_idx, sub_title, note in tasks:
        main_dir = pathlib.Path(f"{m_idx:02d}. {safe_name(main_title)}")
        main_dir.mkdir(exist_ok=True)
        topic_dir = main_dir / f"{s_idx:02d}. {safe_name(sub_title)}"
        topic_dir.mkdir(exist_ok=True)

        # Create folders with requested exact numbered names (swapped per latest instruction)
        tf_dir    = topic_dir / "01. True or False Questions Worksheet"
        tf_dir.mkdir(exist_ok=True)
        mcq_dir   = topic_dir / "02. Multiple Choice Questions Worksheet"
        mcq_dir.mkdir(exist_ok=True)
        # Fill-in-the-Blank (Day 3) removed — do not create
        sa_dir    = topic_dir / "03. Short Answer Type Questions Worksheet"
        sa_dir.mkdir(exist_ok=True)
        # Master Review (Day 5) removed — do not create
        # New folder for task cards
        tc_dir    = topic_dir / "04. Task Cards"
        tc_dir.mkdir(exist_ok=True)
        prev_dir  = topic_dir / "PREVIEW PDFs (Do not Upload This)"
        prev_dir.mkdir(exist_ok=True)

        print(f"🔧  Building sets for {m_idx}.{s_idx} – {sub_title} …")
        # Request slightly larger batches from the API to allow filtering and quality checks
        # AI request sizes:
        #   Multiple Choice Questions: 30
        #   True/False Questions: 30
        #   Short Answer Questions: 30
        #   Task Cards: 35
        mcq = build_mcq(sub_title, note, n=30)
        tf  = build_tf (sub_title, note, n=30)
        sa  = build_sa (sub_title, note, n=30)

        # Ensure each deck has exactly the expected items (trim or pad as needed)
        mcq = normalize_deck(mcq, expected=25)
        tf  = normalize_deck(tf, expected=25)
        sa  = normalize_deck(sa, expected=25)
        scenario = []  # Scenario section removed

        # Generate Task Cards (request 35 to allow filtering, but include 30 in PDFs)
        task_cards = build_task_cards(sub_title, note, n=35)
        # Ensure task cards deck has exactly 30 items (trim or pad)
        task_cards = pad_task_cards(task_cards, expected=30)

        # ------------------
        # Ensure correct answers are randomly distributed across options
        # for all multiple-choice items (MCQ worksheet + Task Cards).
        # This avoids placing the correct answer always in the same position.
        # ------------------
        def redistribute_correct_positions(mcq_deck, tc_deck):
            """mcq_deck: list of (q, opts, letter, exp)
               tc_deck: list of (title, q, opts, letter, exp, num)
               Returns updated (mcq_deck, tc_deck) with options re-ordered and letters updated.
            """
            # Combine counts
            total = len(mcq_deck) + len(tc_deck)
            if total == 0:
                return mcq_deck, tc_deck

            # Build a balanced pool of target letters A-D with near-equal counts
            base = total // 4
            rem = total % 4
            letters = []
            for i, ch in enumerate('ABCD'):
                cnt = base + (1 if i < rem else 0)
                letters.extend([ch] * cnt)
            random.shuffle(letters)

            # Helper to rotate options so correct_text lands at desired index
            def rotate_to(opts, correct_text, desired_index):
                try:
                    cur = opts.index(correct_text)
                except ValueError:
                    # If correct_text not found (shouldn't happen), leave as-is
                    return opts
                # Number of left rotations to align cur -> desired_index
                shift = (desired_index - cur) % len(opts)
                if shift == 0:
                    return opts
                return opts[shift:] + opts[:shift]

            # Apply to mcq_deck then tc_deck
            idx = 0
            new_mcq = []
            for q, opts, letter, exp in mcq_deck:
                correct_text = opts[ord(letter) - ord('A')] if isinstance(letter, str) and letter in 'ABCD' else opts[0]
                desired_letter = letters[idx]
                desired_index = ord(desired_letter) - ord('A')
                new_opts = rotate_to(list(opts), correct_text, desired_index)
                new_letter = "ABCD"[new_opts.index(correct_text)]
                new_mcq.append((q, new_opts, new_letter, exp))
                idx += 1

            new_tc = []
            for title, q, opts, letter, exp, num in tc_deck:
                correct_text = opts[ord(letter) - ord('A')] if isinstance(letter, str) and letter in 'ABCD' else opts[0]
                desired_letter = letters[idx]
                desired_index = ord(desired_letter) - ord('A')
                new_opts = rotate_to(list(opts), correct_text, desired_index)
                new_letter = "ABCD"[new_opts.index(correct_text)]
                new_tc.append((title, q, new_opts, new_letter, exp, num))
                idx += 1

            return new_mcq, new_tc

        mcq, task_cards = redistribute_correct_positions(mcq, task_cards)

        # Use stripped title (no curriculum code) for display and filenames
        display_sub = strip_curriculum_code(sub_title)
        base = safe_name(display_sub)
        make_mcq (mcq_dir / f"{base} – Multiple Choice Worksheet.pdf",
            mcq_dir / f"{base} – Multiple Choice Answer Sheet.pdf",
            main_title, display_sub, mcq)
        make_tf  (tf_dir  / f"{base} – True or False Worksheet.pdf",
            tf_dir  / f"{base} – True or False Answer Sheet.pdf",
            main_title, display_sub, tf)
        make_sa  (sa_dir  / f"{base} – Short Answer Worksheet.pdf",
            sa_dir  / f"{base} – Short Answer Answer Sheet.pdf",
            main_title, display_sub, sa)

        # Task cards outputs
        tc_out = tc_dir / f"{base} – Task Cards.pdf"
        tc_ans = tc_dir / f"{base} – Task Cards Answer Sheet.pdf"

        # Create canonical (clean) Task Cards files in task-cards folder
        make_task_cards_pdf(task_cards, tc_out, tc_ans, main_title, display_sub, preview=False)

        # Create preview pieces in temporary files and merge into a single final preview file
        final_preview = prev_dir / f"{base} – Preview with Task Cards.pdf"

        try:
            if PdfMerger is None:
                raise RuntimeError("PDF merger not available")

            # Temporary files for main preview, intro, and watermarked copies
            tmp_preview = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            tmp_intro = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            tmp_tc_preview = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            tmp_tc_ans_preview = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            tmp_preview.close(); tmp_intro.close(); tmp_tc_preview.close(); tmp_tc_ans_preview.close()

            # Build main preview into a temp file (do not leave Preview.pdf in PREVIEW folder)
            make_full_preview(tmp_preview.name, main_title, tf, mcq, None, sa)

            # Build intro page into a temp file
            make_task_cards_intro_page(tmp_intro.name, main_title, display_sub, count=30)

            # Build watermarked task cards into temp files (these are not written in Task Cards folder)
            make_task_cards_pdf(task_cards, tmp_tc_preview.name, tmp_tc_ans_preview.name, main_title, display_sub, preview=True)

            # Merge temp files into the single final preview in PREVIEW folder
            merger = PdfMerger()
            merger.append(tmp_preview.name)
            merger.append(tmp_intro.name)
            merger.append(tmp_tc_preview.name)
            merger.append(tmp_tc_ans_preview.name)
            merger.write(str(final_preview))
            merger.close()

        except Exception as e:
            print(f"⚠️  Could not produce merged preview with task cards: {e}")
            # Fallback: write main preview alone into the PREVIEW folder so there's at least something
            try:
                make_full_preview(final_preview, main_title, tf, mcq, None, sa)
            except Exception:
                pass
        finally:
            # Clean up temporary files if they exist
            for p in (locals().get('tmp_preview'), locals().get('tmp_intro'), locals().get('tmp_tc_preview'), locals().get('tmp_tc_ans_preview')):
                try:
                    if p:
                        os.unlink(p.name)
                except Exception:
                    pass

        print("   ✔  Done")

    print(f"\n🎉  Completed {len(tasks)} topic(s) in {time.time()-start:.1f}s")

def parse_args(argv=None):
   p = argparse.ArgumentParser(description="Worksheet PDF Builder (filtered by NHES standards and subtopics)")
   p.add_argument('--standards','-S', help="Comma list / ranges of standard unit indices (e.g. 1,3-5). Default: all.")
   p.add_argument('--subs','-T', help="Comma list / ranges of subtopic indices within selected standards (e.g. 1,4-6). Default: all.")
   p.add_argument('--list', action='store_true', help="List available standards & subtopics then exit.")
   p.add_argument('--dry-run', action='store_true', help="Show what would be generated without calling the API.")
   return p.parse_args(argv)

def main(argv=None):
    global CURRICULUM, CURRICULUM_DATA
    
    # Load curricula from Excel file
    script_dir = pathlib.Path(__file__).parent
    excel_file = script_dir / "details.xlsx"
    
    print(f"📖 Loading curricula from {excel_file}\n")
    all_curricula = load_curriculum_from_excel(str(excel_file))
    print(f"✅ Loaded {len(all_curricula)} curriculum(s)\n")
    
    overall_start = time.time()
    
    # Create main folder with current date & 12-hour time, e.g., "19 October 2025 - 5:00 pm"
    now = datetime.now()
    day = now.day  # no leading zero
    month = now.strftime("%B")
    year = now.year
    hour = now.strftime("%I").lstrip("0") or "0"  # 12-hour without leading zero
    minute = now.strftime("%M")
    ampm = now.strftime("%p").lower()
    formatted_dt = f"{day} {month} {year} - {hour}:{minute} {ampm}"
    root_folder = pathlib.Path(f"Academy Ready - 03 WORKSHEETS + 30 TASK CARDS - {formatted_dt}")
    root_folder.mkdir(exist_ok=True)
    print(f"📦 Created root folder: {root_folder}\n")
    
    # Process each curriculum separately
    for curriculum_idx, CURRICULUM_DATA in enumerate(all_curricula, 1):
        print(f"\n{'#'*70}")
        print(f"# Processing Curriculum {curriculum_idx}/{len(all_curricula)}")
        print(f"# {CURRICULUM_DATA.subject_name} - {CURRICULUM_DATA.grade_level}")
        print(f"{'#'*70}\n")
        
        # Get curriculum context for OpenAI prompts
        curriculum_context = CURRICULUM_DATA.get_prompt_context()
        print(f"🎯 Curriculum Context:\n{curriculum_context}\n")
        
        t0 = time.time()
        
        # Update global CURRICULUM variable for compatibility with existing functions
        CURRICULUM = CURRICULUM_DATA.units
        
        # Create main output directory with format: "Grade - Curriculum - Subject"
        main_folder_name = f"{safe_name(CURRICULUM_DATA.grade_level)} - {safe_name(CURRICULUM_DATA.curriculum_name)} - {safe_name(CURRICULUM_DATA.subject_name)}"
        main_dir = root_folder / main_folder_name
        main_dir.mkdir(exist_ok=True)
        print(f"📁 Output directory: {main_dir}\n")
        
        ensure_api_key()
        
        for m_i, m_t, subs in CURRICULUM_DATA.units:
            # Create unit directory with number and full standard code (e.g., "01. MS-PS1 - Matter and Its Interactions")
            # Extract just the standard code part if title contains " - "
            if " - " in m_t or " – " in m_t:
                # Split by either hyphen type
                parts = re.split(r'\s+[-–]\s+', m_t, 1)
                if len(parts) == 2:
                    standard_code = parts[0].strip()
                    standard_title = parts[1].strip()
                    unit_folder_name = f"{m_i:02d}. {standard_code} - {standard_title}"
                else:
                    unit_folder_name = f"{m_i:02d}. {m_t}"
            else:
                unit_folder_name = f"{m_i:02d}. {m_t}"
            
            unit_dir = main_dir / safe_name(unit_folder_name)
            unit_dir.mkdir(exist_ok=True)

            for sub in subs:
                s_i, s_t = sub[:2]
                note = sub[2] if len(sub) > 2 else ""

                # Create subfolder with number and full title (including curriculum code)
                sub_dir = unit_dir / safe_name(f"{s_i:02d}. {s_t}")
                sub_dir.mkdir(exist_ok=True)
                
                # Create worksheet folders
                tf_dir    = sub_dir / "01. True or False Questions Worksheet"
                tf_dir.mkdir(exist_ok=True)
                mcq_dir   = sub_dir / "02. Multiple Choice Questions Worksheet"
                mcq_dir.mkdir(exist_ok=True)
                sa_dir    = sub_dir / "03. Short Answer Type Questions Worksheet"
                sa_dir.mkdir(exist_ok=True)
                tc_dir    = sub_dir / "04. Task Cards"
                tc_dir.mkdir(exist_ok=True)
                prev_dir  = sub_dir / "PREVIEW PDFs (Do not Upload This)"
                prev_dir.mkdir(exist_ok=True)

                print(f"🔍  Generating {m_i}.{s_i} {strip_curriculum_code(s_t)}")
                
                # Build questions with curriculum context
                mcq = build_mcq(s_t, note, n=30)
                tf  = build_tf (s_t, note, n=30)
                sa  = build_sa (s_t, note, n=30)
                
                # Normalize decks
                mcq = normalize_deck(mcq, expected=25)
                tf  = normalize_deck(tf, expected=25)
                sa  = normalize_deck(sa, expected=25)
                
                # Generate Task Cards
                task_cards = build_task_cards(s_t, note, n=35)
                task_cards = pad_task_cards(task_cards, expected=30)
                
                # Redistribute correct answer positions
                def redistribute_correct_positions(mcq_deck, tc_deck):
                    total = len(mcq_deck) + len(tc_deck)
                    if total == 0:
                        return mcq_deck, tc_deck

                    base = total // 4
                    rem = total % 4
                    letters = []
                    for i, ch in enumerate('ABCD'):
                        cnt = base + (1 if i < rem else 0)
                        letters.extend([ch] * cnt)
                    random.shuffle(letters)

                    def rotate_to(opts, correct_text, desired_index):
                        try:
                            cur = opts.index(correct_text)
                        except ValueError:
                            return opts
                        shift = (desired_index - cur) % len(opts)
                        if shift == 0:
                            return opts
                        return opts[shift:] + opts[:shift]

                    idx = 0
                    new_mcq = []
                    for q, opts, letter, exp in mcq_deck:
                        correct_text = opts[ord(letter) - ord('A')] if isinstance(letter, str) and letter in 'ABCD' else opts[0]
                        desired_letter = letters[idx]
                        desired_index = ord(desired_letter) - ord('A')
                        new_opts = rotate_to(list(opts), correct_text, desired_index)
                        new_letter = "ABCD"[new_opts.index(correct_text)]
                        new_mcq.append((q, new_opts, new_letter, exp))
                        idx += 1

                    new_tc = []
                    for title, q, opts, letter, exp, num in tc_deck:
                        correct_text = opts[ord(letter) - ord('A')] if isinstance(letter, str) and letter in 'ABCD' else opts[0]
                        desired_letter = letters[idx]
                        desired_index = ord(desired_letter) - ord('A')
                        new_opts = rotate_to(list(opts), correct_text, desired_index)
                        new_letter = "ABCD"[new_opts.index(correct_text)]
                        new_tc.append((title, q, new_opts, new_letter, exp, num))
                        idx += 1

                    return new_mcq, new_tc

                mcq, task_cards = redistribute_correct_positions(mcq, task_cards)

                # Use stripped title for display and filenames
                display_sub = strip_curriculum_code(s_t)
                base = safe_name(display_sub)
                
                make_mcq(mcq_dir / f"{base} – Multiple Choice Worksheet.pdf",
                        mcq_dir / f"{base} – Multiple Choice Answer Sheet.pdf",
                        m_t, display_sub, mcq)
                make_tf(tf_dir / f"{base} – True or False Worksheet.pdf",
                       tf_dir / f"{base} – True or False Answer Sheet.pdf",
                       m_t, display_sub, tf)
                make_sa(sa_dir / f"{base} – Short Answer Worksheet.pdf",
                       sa_dir / f"{base} – Short Answer Answer Sheet.pdf",
                       m_t, display_sub, sa)

                # Task cards
                tc_out = tc_dir / f"{base} – Task Cards.pdf"
                tc_ans = tc_dir / f"{base} – Task Cards Answer Sheet.pdf"
                make_task_cards_pdf(task_cards, tc_out, tc_ans, m_t, display_sub, preview=False)

                # Create preview
                final_preview = prev_dir / f"{base} – Preview with Task Cards.pdf"

                try:
                    if PdfMerger is None:
                        raise RuntimeError("PDF merger not available")

                    tmp_preview = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                    tmp_intro = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                    tmp_tc_preview = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                    tmp_tc_ans_preview = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                    tmp_preview.close(); tmp_intro.close(); tmp_tc_preview.close(); tmp_tc_ans_preview.close()

                    make_full_preview(tmp_preview.name, m_t, tf, mcq, None, sa)
                    make_task_cards_intro_page(tmp_intro.name, m_t, display_sub, count=30)
                    make_task_cards_pdf(task_cards, tmp_tc_preview.name, tmp_tc_ans_preview.name, m_t, display_sub, preview=True)

                    merger = PdfMerger()
                    merger.append(tmp_preview.name)
                    merger.append(tmp_intro.name)
                    merger.append(tmp_tc_preview.name)
                    merger.append(tmp_tc_ans_preview.name)
                    merger.write(str(final_preview))
                    merger.close()

                except Exception as e:
                    print(f"⚠️  Could not produce merged preview: {e}")
                    try:
                        make_full_preview(final_preview, m_t, tf, mcq, None, sa)
                    except Exception:
                        pass
                finally:
                    for p in (locals().get('tmp_preview'), locals().get('tmp_intro'), 
                             locals().get('tmp_tc_preview'), locals().get('tmp_tc_ans_preview')):
                        try:
                            if p:
                                os.unlink(p.name)
                        except Exception:
                            pass

                print(f"✅  Saved → {sub_dir}")

        print(f"\n✅ Curriculum {curriculum_idx} completed in {time.time()-t0:.1f}s")
        print(f"📦  Files saved in: {main_dir.absolute()}")

    print(f"\n{'='*70}")
    print(f"🎉  ALL CURRICULA COMPLETED in {time.time()-overall_start:.1f}s")
    print(f"📦  All files saved in: {root_folder.absolute()}")
    print(f"{'='*70}")


# ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
   random.seed()
   main()
