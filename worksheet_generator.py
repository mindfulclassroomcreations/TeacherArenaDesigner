#!/usr/bin/env python3

"""
Worksheet PDF Builder – MCQ • True/False • Short-Answer  (v9.1)
Refactored for Web Application
"""

import os, re, json, random, pathlib, sys, time, string, shutil
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
    from pypdf import PdfMerger, PdfReader, PdfWriter
except Exception:
    try:
        from PyPDF2 import PdfMerger, PdfReader, PdfWriter
    except Exception:
        PdfMerger = PdfReader = PdfWriter = None

# ───────────────────────  CONFIG  ────────────────────────────────
MODEL = "gpt-4o-mini"

FONT_PATHS = [
    "DejaVuSans.ttf",
    "NotoSans-VariableFont_wdth,wght.ttf",
    "Fredoka-VariableFont_wdth.ttf",
]

def register_fonts(font_dir="."):
   # Title Font: Comic Neue Bold for main titles and section headers
   title_candidates = [
       "ComicNeue-Bold.ttf",
       "Comic Neue Bold.ttf",
       "CenturyGothic.ttf",
       "Century Gothic.ttf",
   ]
   
   # Question & Answer Font: Nunito Regular
   body_candidates = [
      "DejaVuSans.ttf",
      "NotoSans-VariableFont_wdth,wght.ttf",
      "Nunito-Regular.ttf",
      "Nunito.ttf",
      "Verdana.ttf",
   ]
   
   # Explanation Font: Nunito Light Italic
   explanation_candidates = [
       "Nunito-LightItalic.ttf",
       "Nunito-Light-Italic.ttf",
       "Nunito-Italic.ttf",
   ]

   title = None
   body = None
   explanation = None

   # Helper to register font from directory
   def try_register(candidates):
       for p in candidates:
           # Check in current dir and font_dir
           paths_to_check = [p, os.path.join(font_dir, p)]
           for path in paths_to_check:
               if os.path.exists(path):
                   try:
                       fam = pathlib.Path(path).stem
                       pdfmetrics.registerFont(TTFont(fam, path))
                       return fam
                   except Exception:
                       pass
       return None

   title = try_register(title_candidates)
   body = try_register(body_candidates)
   explanation = try_register(explanation_candidates)

   if not body:
       body = "Helvetica"
   if not title:
       title = body
   if not explanation:
       explanation = body
   
   return title, body, explanation

# Initialize fonts globally (will be re-initialized in generate_worksheets if needed)
TITLE_FONT, BODY_FONT, EXPL_FONT = "Helvetica", "Helvetica", "Helvetica"

# Curriculum constants
CURRICULUM_NAME = "NGSS - Middle School Physical Sciences"
GRADE_LEVELS = "6,7,8"


# ─────────────────────  UTILITIES  ───────────────────────────────
SUB_MAP = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
SUP_MAP = {"+": "⁺", "-": "⁻"}
DIGIT_RUN  = re.compile(r'([A-Za-z\)])(\d+)')
ION_CHARGE = re.compile(r'([A-Za-z₀-₉]+)([+-])$')
STRIP_BOX  = str.maketrans({
    "■": "",
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
    if not s or not isinstance(s, str):
        return s
    for sep in ('–', '—', '-'):
        if sep in s:
            left, right = s.split(sep, 1)
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

MAX_CARD_CHARS = 325


# ───────────────────  REPORTLAB STYLES  ──────────────────────────
ST = None # Initialized later

def get_styles():
    st = getSampleStyleSheet()
    def add(n, **kw):
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
    add('PreviewTitleMain', fontName=TITLE_FONT, fontSize=24, leading=28, alignment=1, textColor=colors.blue)
    add('PreviewTitleSub',  fontName=TITLE_FONT, fontSize=16, leading=20, alignment=1, textColor=colors.blue)
    return st


# ───────────────────  PAGE DECORATION  ───────────────────────────
PAGE_W, PAGE_H = letter
M_LEFT = 35
M_RIGHT = 35
M_TOP = 50
M_BOTTOM = 40
BORDER_INSET = 6
CARD_CORNER_RADIUS = 10

def draw_frame(c):
    x = BORDER_INSET
    y = BORDER_INSET
    width = PAGE_W - 2 * BORDER_INSET
    height = PAGE_H - 2 * BORDER_INSET
    c.setStrokeColor(colors.black)
    c.setLineWidth(2)
    c.rect(x, y, width, height)

def first(c, d):
    draw_frame(c)
    c.setFont(BODY_FONT, 12)
    name_x = BORDER_INSET + 10
    name_y = PAGE_H - BORDER_INSET - 26 
    c.drawString(name_x, name_y, "Name: ...........................................................          Date : ........................")
    c.setFont(BODY_FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))

def first_page_no_watermark(c, d):
    draw_frame(c)
    c.saveState()
    c.setFillColor(colors.red)
    c.setFont(TITLE_FONT, 18)

    notice_text = (
        "This is a combined preview only. After purchase, you will receive separate question PDFs and "
        "separate answer PDFs with explanations."
    )

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

    line_height = 22
    start_y = 25 + 15 + (len(lines) * line_height)
    y = start_y
    for line in lines:
        c.drawCentredString(PAGE_W/2, y, line)
        y -= line_height
    c.restoreState()

    c.setFont(BODY_FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))

def later(c, d):
    draw_frame(c)
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
    draw_frame(c)
    c.setFont(BODY_FONT, 12)
    name_x = BORDER_INSET + 10
    name_y = PAGE_H - BORDER_INSET - 26
    c.drawString(name_x, name_y, "Name: ...........................................................          Date : ........................")
    c.setFont(BODY_FONT, 10)
    c.drawCentredString(PAGE_W/2, 25, str(d.page))

class TitlePage(Flowable):
    def __init__(self, line_specs: List[Tuple[str, str, int]]):
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
        
        spacings = []
        for i in range(len(self.line_specs) - 1):
            avg_size = (self.line_specs[i][2] + self.line_specs[i+1][2]) / 2
            spacings.append(avg_size * 0.8)
        
        total_h = sum(h for h, d in line_heights) + sum(spacings)
        start_y = PAGE_H/2 + (total_h / 2)
        
        y = start_y
        for i, (text, font_name, font_size) in enumerate(self.line_specs):
            c.setFont(font_name, font_size)
            line_h, descent = line_heights[i]
            try:
                w = pdfmetrics.stringWidth(text, font_name, font_size)
            except Exception:
                w = font_size * len(text) * 0.5
            x = PAGE_W/2 - (w / 2)
            c.drawString(x, y - line_h + descent, text)
            y -= line_h
            if i < len(spacings):
                y -= spacings[i]
        c.restoreState()

class UnderlinedAnswer(Flowable):
    def __init__(self, text, width=400, fontName=None, fontSize=12, lineColor=colors.blue, leftIndent=0):
        super().__init__()
        self.text = text
        self.width = width
        self.fontName = fontName or BODY_FONT
        self.fontSize = fontSize
        self.lineColor = lineColor
        self.leftIndent = leftIndent

    def wrap(self, availWidth, availHeight):
        return availWidth, self.fontSize + 6

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFont(self.fontName, self.fontSize)
        c.setFillColor(colors.black)
        x = getattr(self, 'leftIndent', 0) or 0
        c.drawString(x, 4, self.text)
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
   return (
       f"Write EXACTLY {n} MCQs for Day 2 - Knowledge Builder: Multiple Choice Review on: {topic}. "
       f"Focus on concept recognition and key details. Target Bloom's levels: Understand/Apply. "
       f"Ensure each question aligns with the {CURRICULUM_NAME} ({GRADE_LEVELS}) and the provided teacher note.\n"
       f"Teacher note: {note}\n"
       'Return JSON list: {"q":"","correct":"","distractors":["","",""],"explanation":""}\n'
       "≤325 chars total per item. Randomise answer order."
   )

def build_task_cards(topic, note, n=30):
    deck = []
    attempts = 0
    max_attempts = 60
    simple_switch = 5
    single_item_switch = 18
    while len(deck) < n and attempts < max_attempts:
        attempts += 1
        if attempts > single_item_switch:
            prompt = p_mcq_simple(topic, note, 1)
        elif attempts > simple_switch:
            rem = n - len(deck)
            prompt = p_mcq_simple(topic, note, rem)
        else:
            prompt = p_mcq(topic, note, n-len(deck))

        items = get_json(prompt)
        if not items: continue

        short_title = strip_curriculum_code(topic).upper()
        for itm in items:
            if not isinstance(itm, dict): continue
            q   = clean(itm.get("q", ""))
            ok  = clean(itm.get("correct", ""))
            ds  = [clean(d) for d in itm.get("distractors", [])][:3]
            exp = clean(itm.get("explanation", ""))
            opts = ds + [ok]
            if len(opts) < 4: continue
            random.shuffle(opts)
            if ok not in opts: continue
            total_chars = len(q) + sum(len(o) for o in opts)
            if total_chars > MAX_CARD_CHARS: continue
            if body_font(total_chars) == 0: continue
            num = len(deck) + 1
            title = short_title
            deck.append((title, q, opts, "ABCD"[opts.index(ok)], exp, num))
            if len(deck) == n: break
    return deck

def p_tf(topic, note, n):
   return (
       f"Write EXACTLY {n} True/False statements for Day 1 - Concept Check on: {topic}. "
       f"Focus on basic recall and misconception checks. Target Bloom's levels: Remember/Understand. "
       f"Ensure alignment with the {CURRICULUM_NAME} ({GRADE_LEVELS}) and the teacher note.\n"
       f"Teacher note: {note}\n"
       'Return JSON list: {"statement":"","answer":true/false,"explanation":""}\n'
       "If answer is false, give ≤15-word explanation, else \"\". ≤325 chars item."
   )

def p_sa(topic, note, n):
   return (
       f"Write EXACTLY {n} short-answer questions for Day 4 - Critical Thinking: Short Response on: {topic}. "
       f"Focus on explaining, connecting concepts, and applying reasoning. Target Bloom's levels: Apply/Analyze. "
       f"Ensure each question aligns with the {CURRICULUM_NAME} ({GRADE_LEVELS}) and the teacher note.\n"
       f"Teacher note: {note}\n"
       'Return JSON list: {"q":"","answer":""}\n'
       "Question ≤250 chars, answer ≤25 words."
   )

def p_mcq_simple(topic, note, n):
    return (
        f"Write EXACTLY {n} SHORT and SIMPLE MCQs on: {topic}. "
        f"Use concise stems (≤100 chars) and short options (≤40 chars). Keep language direct and avoid multi-part scenarios. "
        f"Return JSON list: {{\"q\":\"\",\"correct\":\"\",\"distractors\":[\"\",\"\",\"\"],\"explanation\":\"\"}}. Randomise order."
    )

def p_tf_simple(topic, note, n):
    return (
        f"Write EXACTLY {n} SHORT True/False statements on: {topic}. "
        f"Make each statement direct and concise (≤120 chars). Return JSON list: {{\"statement\":\"\",\"answer\":true/false,\"explanation\":\"\"}}."
    )

def p_sa_simple(topic, note, n):
    return (
        f"Write EXACTLY {n} SHORT-answer questions on: {topic}. "
        f"Keep questions concise (≤120 chars) and targeted. Return JSON list: {{\"q\":\"\",\"answer\":\"\"}}."
    )

def build_mcq(topic, note, n=25):
    deck = []
    max_attempts = 20
    attempts = 0
    simple_switch = 5
    single_item_switch = 12
    while len(deck) < n and attempts < max_attempts:
        attempts += 1
        remaining = n - len(deck)
        if attempts > single_item_switch:
            prompt = p_mcq_simple(topic, note, 1)
        elif attempts > simple_switch:
            prompt = p_mcq_simple(topic, note, remaining)
        else:
            prompt = p_mcq(topic, note, remaining)

        items = get_json(prompt)
        if not items: continue
        for itm in items:
            if not isinstance(itm, dict): continue
            q   = clean(itm.get("q", ""))
            ok  = clean(itm.get("correct", ""))
            ds  = [clean(d) for d in itm.get("distractors", [])]
            exp = clean(itm.get("explanation", ""))
            opts = ds + [ok]
            random.shuffle(opts)
            if body_font(len(q) + sum(len(o) for o in opts)) == 0: continue
            if ok not in opts: continue
            deck.append((q, opts, "ABCD"[opts.index(ok)], exp))
            if len(deck) == n: break
    return deck

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
        if attempts > single_item_switch:
            prompt = p_tf_simple(topic, note, 1)
        elif attempts > simple_switch:
            prompt = p_tf_simple(topic, note, remaining)
        else:
            prompt = p_tf(topic, note, remaining)
        items = get_json(prompt)
        if not items: continue
        for itm in items:
            if not isinstance(itm, dict): continue
            stmt = clean(itm.get("statement", ""))
            ans  = bool_to_str(itm.get("answer", ""))
            exp  = clean(itm.get("explanation", ""))
            if ans not in ("True", "False"): continue
            if body_font(len(stmt)+5) == 0: continue
            deck.append((stmt, ans, exp))
            if len(deck) == n: break
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
        if attempts > single_item_switch:
            prompt = p_sa_simple(topic, note, 1)
        elif attempts > simple_switch:
            prompt = p_sa_simple(topic, note, remaining)
        else:
            prompt = p_sa(topic, note, remaining)
        items = get_json(prompt)
        if not items: continue
        for itm in items:
            if not isinstance(itm, dict): continue
            q   = clean(itm.get("q", ""))
            ans = clean(itm.get("answer", ""))
            if body_font(len(q)+len(ans)) == 0: continue
            deck.append((q, ans))
            if len(deck) == n: break
    return deck

# ───────────────────  PDF HELPERS  ───────────────────────────────
def doc(path):
   return SimpleDocTemplate(str(path), pagesize=letter,
                            leftMargin=35, rightMargin=35,
                            topMargin=50,  bottomMargin=40)

def sa_lines():
    line = "_" * 85
    return [Paragraph(line, ST["Line"]) for _ in range(2)]

def normalize_deck(deck, expected=10):
    if not isinstance(deck, list): return deck
    if len(deck) > expected: return deck[:expected]
    out = list(deck)
    if not out:
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
    out = list(deck) if isinstance(deck, list) else []
    while len(out) < expected:
        num = len(out) + 1
        out.append(("", "", ["", "", "", ""], "A", "", num))
    if len(out) > expected:
        return out[:expected]
    return out

# ───────────────────  WORKSHEET MAKERS  ──────────────────────────
def make_mcq(ws_pdf, ans_pdf, main, sub, deck):
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
        for l,t in zip("ABCD",opts):
            text = f"{l}. {t}"
            if l == ans:
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
        story.append(Paragraph(ans, ST["Blue"]))
        story.append(Spacer(1,8))
    doc(ans_pdf).build(story, onFirstPage=first, onLaterPages=later)

def make_task_cards_pdf(cards, out_pdf, answer_pdf, main, sub, preview=False):
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

    per_page = cols * rows
    for idx, card in enumerate(cards):
        page_idx = idx // per_page
        pos_in_page = idx % per_page
        if pos_in_page == 0:
            if idx != 0:
                c.showPage()
            if preview:
                c.saveState()
                c.setFont(BODY_FONT, 130)
                c.setFillColor(colors.lightgrey)
                c.translate(PAGE_W/2, PAGE_H/2)
                c.rotate(45)
                c.drawCentredString(0, 0, "PREVIEW")
                c.restoreState()
        if pos_in_page == 0:
            c.setStrokeColor(colors.black)
            c.setDash(3,2)
            gx = margin + card_w + (gutter / 2)
            c.line(gx, margin + 6, gx, PAGE_H - margin - 6)
            for r in range(rows - 1):
                gy = PAGE_H - margin - (r+1) * card_h - r * gutter - (gutter / 2)
                c.line(margin + 6, gy, PAGE_W - margin - 6, gy)
            c.setDash()
        col = pos_in_page % cols
        row = pos_in_page // cols
        x = margin + col * (card_w + gutter)
        y = PAGE_H - margin - (row + 1) * card_h - row * gutter + (card_h - card_h)

        if len(card) == 6:
            title, q, opts, ans, exp, num = card
        else:
            title, q, opts, ans, exp = card
            num = None
        
        hdr_h = 0.8 * 72
        hdr_y = y + card_h - hdr_h
        c.setFillColor(colors.black)
        bold_font_candidates = [f"{TITLE_FONT}-Bold", f"{TITLE_FONT} Bold", "Helvetica-Bold", TITLE_FONT, BODY_FONT]
        for bf in bold_font_candidates:
            try:
                pdfmetrics.getFont(bf)
                c.setFont(bf, 12)
                break
            except Exception:
                try:
                    c.setFont(bf, 12)
                    break
                except Exception:
                    continue

        title_text = title.upper()
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
        for ln in hdr_lines:
            c.drawCentredString(x + card_w/2, ty, ln)
            ty -= 12

        c.setStrokeColor(colors.black)
        c.setLineWidth(1)
        try:
            c.roundRect(x, y, card_w, card_h, CARD_CORNER_RADIUS)
        except Exception:
            c.rect(x, y, card_w, card_h)

        body_x = x + 8
        body_y = y + card_h - hdr_h - 18
        c.setFillColor(colors.black)
        total_chars = len(q) + sum(len(o) for o in opts)
        fsize = body_font(total_chars)
        if fsize == 0: continue
        c.setFont(BODY_FONT, fsize)
        if num is not None:
            q_text = f"{num:02d}. {q}"
        else:
            q_text = q
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
        ty -= fsize
        for label, opt in zip('ABCD', opts):
            text = f"{label}. {opt}"
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

    c.showPage()
    c.save()

    story = [Paragraph(f"{sub} - Task Cards Answer Sheet", ST['AnsH']), Spacer(1,12)]
    for i, card in enumerate(cards, 1):
        if len(card) == 6:
            title, q, opts, ans, exp, num = card
        else:
            title, q, opts, ans, exp = card
            num = i
        story.append(Paragraph(f"{num:02d}. {q}", ST['Q']))
        letter_index = ord(ans) - ord('A') if isinstance(ans, str) and ans in 'ABCD' else 0
        correct_text = opts[letter_index]
        story.append(Paragraph(f"{ans}. {correct_text}", ST['Blue']))
        story.append(Spacer(1,6))
    if preview:
        doc(answer_pdf).build(story, onFirstPage=preview_page, onLaterPages=preview_page)
    else:
        doc(answer_pdf).build(story, onFirstPage=first, onLaterPages=later)

def make_task_cards_intro_page(path, main, sub, count=30):
    line_specs = [ (f"{count} Task Cards", TITLE_FONT, 36), ("Includes Answer Key with Explanations", BODY_FONT, 16) ]
    tp = TitlePage(line_specs)
    docp = SimpleDocTemplate(str(path), pagesize=letter, leftMargin=35, rightMargin=35, topMargin=50, bottomMargin=40)
    docp.build([tp], onFirstPage=preview_page_no_watermark, onLaterPages=preview_page_no_watermark)

def make_full_preview(preview_path, main, tf_d, mcq_d, fill_d, sa_d, mr_tf=None, mr_mcq=None, mr_fill=None, mr_sa=None, scenario_d=None):
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

    story = [NextPageTemplate('FirstFramed')]
    story.append(Paragraph(f"<b>{main.upper()} – COMBINED PREVIEW</b>", ST["Doc"]))
    story.append(Spacer(1,12))
    story.append(NextPageTemplate('Framed'))

    def push_title(day_title: str, label: str, question_count: int = 10):
        story.append(Spacer(1, 294))
        story.append(Paragraph(f"<b>{day_title}</b>", ST["PreviewTitleMain"]))
        story.append(Paragraph(f"<b>{question_count} {label}</b>", ST["PreviewTitleSub"]))
        story.append(Paragraph("<b>Includes Answer Key with Explanations</b>", ST["PreviewTitleSub"]))
        story.append(PageBreak())

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

    push_title("Short Answer WorkSheet", "Short Answer Type Questions", 25)
    story.append(Paragraph("Short Answer WorkSheet    ", ST["AnsH"]))
    for i,(q,_) in enumerate(sa_d,1):
        blk=[Paragraph(f"{i}. {q}", ST["Q"]) ] + sa_lines() + [Spacer(1,12)]
        story.append(KeepTogether(blk))
    story.append(PageBreak())

    story.append(Paragraph("Short Answer - Answer Sheet    ", ST["AnsH"]))
    for i,(q,ans) in enumerate(sa_d,1):
        story.append(Paragraph(f"{i}. {q}", ST["Q"]))
        story.append(Paragraph(ans, ST["Blue"]))
        story.append(Spacer(1,8))
    story.append(PageBreak())

    doc.build(story)


# ───────────────────  CURRICULUM LOADER  ─────────────────────────────────
class CurriculumData:
    def __init__(self, subject: str, grade: str, curriculum: str):
        self.subject_name = subject
        self.grade_level = grade
        self.curriculum_name = curriculum
        self.units: List[Tuple[int, str, List[Tuple[int, str, str]]]] = []
    
    def get_prompt_context(self) -> str:
        return (f"Subject: {self.subject_name}\n"
                f"Grade Level: {self.grade_level}\n"
                f"Curriculum: {self.curriculum_name}\n"
                f"IMPORTANT: All questions MUST align with {self.grade_level} standards "
                f"and stay within the {self.curriculum_name} curriculum. "
                f"Do NOT create questions outside these standards.")

def load_curriculum_from_excel(excel_path: str) -> List[CurriculumData]:
    try:
        df_raw = pd.read_excel(excel_path, header=None)
        all_curricula = []
        curriculum_start_rows = []
        for idx in range(len(df_raw)):
            first_col = str(df_raw.iloc[idx, 0]).strip() if pd.notna(df_raw.iloc[idx, 0]) else ''
            if first_col.lower().startswith('subject name'):
                curriculum_start_rows.append(idx)
        
        if not curriculum_start_rows:
            curriculum_start_rows = [0]
        
        for curr_idx, start_row in enumerate(curriculum_start_rows):
            end_row = curriculum_start_rows[curr_idx + 1] if curr_idx + 1 < len(curriculum_start_rows) else len(df_raw)
            first_cell = str(df_raw.iloc[start_row, 0]).strip() if pd.notna(df_raw.iloc[start_row, 0]) else ""
            
            if first_cell.lower().startswith('subject name'):
                subject_name = str(df_raw.iloc[start_row, 1]).strip() if pd.notna(df_raw.iloc[start_row, 1]) else "Unknown Subject"
                grade_level = str(df_raw.iloc[start_row + 1, 1]).strip() if pd.notna(df_raw.iloc[start_row + 1, 1]) else "Unknown Grade"
                curriculum_name = str(df_raw.iloc[start_row + 2, 1]).strip() if pd.notna(df_raw.iloc[start_row + 2, 1]) else "Unknown Curriculum"
            else:
                subject_name = re.sub(r'^Subject Name\s*-\s*', '', first_cell, flags=re.IGNORECASE).strip()
                grade_level = re.sub(r'^Grade level\s*-\s*', '', str(df_raw.iloc[start_row + 1, 0]).strip(), flags=re.IGNORECASE).strip()
                curriculum_name = re.sub(r'^Curriculum\s*-\s*', '', str(df_raw.iloc[start_row + 2, 0]).strip(), flags=re.IGNORECASE).strip()
            
            curriculum_data = CurriculumData(subject_name, grade_level, curriculum_name)
            unit_index = 0
            current_unit_title = None
            current_subtopics = []
            in_data_section = False
            
            for idx in range(start_row + 3, end_row):
                row = df_raw.iloc[idx]
                first_col = str(row[0]).strip() if pd.notna(row[0]) else ''
                if first_col.lower().startswith('subject name'): break
                row_values = [str(val).strip().upper() for val in row.values if pd.notna(val)]
                is_header = any(h in row_values for h in ['NO', 'TITLE', 'STANDARD', 'NOTE'])
                
                if is_header:
                    in_data_section = True
                    continue
                
                row_empty = all(pd.isna(val) or str(val).strip() == '' for val in row.values)
                if row_empty and in_data_section:
                    if current_unit_title and current_subtopics:
                        unit_index += 1
                        curriculum_data.units.append((unit_index, current_unit_title, current_subtopics))
                        current_subtopics = []
                    in_data_section = False
                    continue
                
                if not in_data_section:
                    if first_col and len(first_col) > 3:
                        current_unit_title = first_col
                else:
                    if len(row) >= 3:
                        no_val = row[0]
                        standard_val = row[1]
                        title_val = row[2]
                        note_val = row[3] if len(row) > 3 else ''
                        try:
                            sub_idx = int(float(no_val)) if pd.notna(no_val) else len(current_subtopics) + 1
                        except (ValueError, TypeError):
                            sub_idx = len(current_subtopics) + 1
                        title = str(title_val).strip() if pd.notna(title_val) else ''
                        if pd.notna(standard_val):
                            standard = str(standard_val).strip()
                            if standard and standard != '':
                                title = f"{standard} — {title}"
                        note = str(note_val).strip() if pd.notna(note_val) else ''
                        if title and title != 'nan' and len(title) > 2:
                            current_subtopics.append((sub_idx, title, note))
            
            if current_unit_title and current_subtopics:
                unit_index += 1
                curriculum_data.units.append((unit_index, current_unit_title, current_subtopics))
            
            if curriculum_data.units:
                all_curricula.append(curriculum_data)
        
        if not all_curricula:
            raise ValueError("No valid curriculum sections found in Excel file")
        return all_curricula
    except Exception as e:
        raise RuntimeError(f"Error loading curriculum: {e}")

# ───────────────────  MAIN GENERATION FUNCTION  ──────────────────────────────────────
def generate_worksheets(excel_path: str, output_dir: str, api_key: str):
    """
    Main entry point for generating worksheets.
    Yields progress updates and results.
    """
    global openai, CURRICULUM_NAME, GRADE_LEVELS, TITLE_FONT, BODY_FONT, EXPL_FONT, ST
    
    openai.api_key = api_key
    if not openai.api_key:
        raise ValueError("OpenAI API Key is required")

    # Initialize fonts and styles
    script_dir = os.path.dirname(os.path.abspath(__file__))
    TITLE_FONT, BODY_FONT, EXPL_FONT = register_fonts(font_dir=script_dir)
    ST = get_styles()

    yield {'type': 'progress', 'message': 'Loading curriculum...'}
    all_curricula = load_curriculum_from_excel(excel_path)
    
    root_folder = pathlib.Path(output_dir)
    root_folder.mkdir(parents=True, exist_ok=True)
    
    total_units = sum(len(c.units) for c in all_curricula)
    total_subtopics = sum(len(subs) for c in all_curricula for _, _, subs in c.units)
    
    yield {'type': 'progress', 'message': f'Found {total_units} units with {total_subtopics} subtopics.'}

    processed_count = 0

    for curriculum_idx, CURRICULUM_DATA in enumerate(all_curricula, 1):
        CURRICULUM_NAME = CURRICULUM_DATA.curriculum_name
        GRADE_LEVELS = CURRICULUM_DATA.grade_level
        
        main_folder_name = f"{safe_name(CURRICULUM_DATA.grade_level)} - {safe_name(CURRICULUM_DATA.curriculum_name)} - {safe_name(CURRICULUM_DATA.subject_name)}"
        main_dir = root_folder / main_folder_name
        main_dir.mkdir(exist_ok=True)
        
        for m_i, m_t, subs in CURRICULUM_DATA.units:
            if " - " in m_t or " – " in m_t:
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

                yield {'type': 'progress', 'message': f'Generating: {s_t}'}

                sub_dir = unit_dir / safe_name(f"{s_i:02d}. {s_t}")
                sub_dir.mkdir(exist_ok=True)
                
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

                mcq = build_mcq(s_t, note, n=30)
                tf  = build_tf (s_t, note, n=30)
                sa  = build_sa (s_t, note, n=30)
                
                mcq = normalize_deck(mcq, expected=25)
                tf  = normalize_deck(tf, expected=25)
                sa  = normalize_deck(sa, expected=25)
                
                task_cards = build_task_cards(s_t, note, n=35)
                task_cards = pad_task_cards(task_cards, expected=30)
                
                def redistribute_correct_positions(mcq_deck, tc_deck):
                    total = len(mcq_deck) + len(tc_deck)
                    if total == 0: return mcq_deck, tc_deck
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
                        if shift == 0: return opts
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

                tc_out = tc_dir / f"{base} – Task Cards.pdf"
                tc_ans = tc_dir / f"{base} – Task Cards Answer Sheet.pdf"
                make_task_cards_pdf(task_cards, tc_out, tc_ans, m_t, display_sub, preview=False)

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
                            if p: os.unlink(p.name)
                        except Exception:
                            pass
                
                processed_count += 1
                yield {
                    'type': 'result',
                    'topic': s_t,
                    'path': str(sub_dir),
                    'progress': f"{processed_count}/{total_subtopics}"
                }

    yield {'type': 'complete', 'path': str(root_folder)}
