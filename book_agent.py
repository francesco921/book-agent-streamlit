# ==========================================
# 🧠 COSA FA (spiegato semplice)
# ------------------------------------------
# Questo programma crea un libro a partire dal tuo TOC:
# - carichi il TOC
# - controlli/correggi il TOC (anche con AI)
# - confermi
# - genera i testi (in blocchi da max 500 parole a chiamata)
# - scarichi DOCX/PDF con formattazione.
# ==========================================

# ==========================================
# 📦 IMPORT: prendo gli attrezzi che servono
# ==========================================
import io
import os
import re
import math
from dataclasses import dataclass, field
from typing import List, Optional

import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PyPDF2 import PdfReader

# PDF (per impaginare il libro)
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.units import cm

# Rilevazione lingua (se presente, altrimenti user selects)
HAS_LANGID = False
try:
    import langid  # pip install langid
    HAS_LANGID = True
except Exception:
    HAS_LANGID = False

# OpenAI (per scrivere i testi veri, se c’è la chiave)
OPENAI_OK = False
try:
    from openai import OpenAI
    if os.getenv("OPENAI_API_KEY"):
        openai_client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        OPENAI_OK = True
except Exception:
    OPENAI_OK = False

# ==========================================
# 🔧 COSTANTI (regole semplici e chiare)
# ==========================================
MAX_SUBGEN_WORDS = 500           # mai più di 500 parole per singola generazione
MIN_SECTION_WORDS_USEFUL = 250   # meno di così una sezione non “respira”
MAX_SECTION_WORDS_SOFT = 1500    # sopra questa soglia segnaliamo “troppo lunga”

# Mappa lingue “umane” → etichette per prompt
LANG_LABELS = {
    "auto": "auto",
    "it": "Italian",
    "en": "English",
    "es": "Spanish",
    "fr": "French",
}

# Formati pagina “da libro”
PAGE_SIZES = {
    "A4": A4,
    "Letter": letter,
    "6x9": (6 * 72, 9 * 72),           # pollici → punti tipografici
    "8.5x11": (8.5 * 72, 11 * 72),
}

# Scelte font (DOCX + PDF verranno armonizzate più avanti)
FONT_CHOICES = ["Times New Roman", "Roboto", "Comfortaa"]

# Tono di voce
TONE_CHOICES = ["Scientifico", "Colloquiale", "Narrativo"]

# ==========================================
# 🧱 SCATOLE DOVE METTO I DATI (modelli)
# ==========================================
@dataclass
class Section:
    title: str
    target_words: int = 0
    blocks: int = 0
    texts: List[str] = field(default_factory=list)  # testi finali dei sottoblocchi (concatenati a vista)

@dataclass
class Chapter:
    title: str
    sections: List[Section] = field(default_factory=list)
    target_words: int = 0
    blocks: int = 0

@dataclass
class BookPlan:
    title: str
    subtitle: str
    author: str
    total_words: int
    block_size: int
    chapters: List[Chapter] = field(default_factory=list)
    language_code: str = "auto"          # "auto", "it", "en", ...
    tone: str = "Colloquiale"            # uno tra TONE_CHOICES
    brief: str = ""                      # descrizione breve che guida lo stile
    pdf_page: str = "6x9"                # formato libro
    font_name: str = "Times New Roman"   # font preferito

# ==========================================
# 🖥️ IMPOSTAZIONI BASE DELLA PAGINA
# ==========================================
st.set_page_config(page_title="Book Agent — Generatore Libro", page_icon="📘", layout="wide")
st.title("📘 Book Agent — Generatore Libro")
st.caption("Carica TOC → rivedi/approva → genera → scarica DOCX/PDF")

# Inizializzo “memoria” (serve dopo)
for key, default in {
    "chapters": None,
    "allocation_done": False,
    "detected_lang": "auto",
    "confirmed_toc_text": None,
    "generated_plan": None,
    "docx_bytes": None,
    "pdf_bytes": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default
# ==========================================
# 📂 BLOCK 2 — TOC UPLOAD AND REVIEW
# ------------------------------------------
# User uploads TOC (DOCX or PDF), the app extracts headings,
# detects language (or lets user choose), and allows manual
# or AI refinement. Proceed only after final confirmation.
# ==========================================

st.subheader("📄 Step 1 — Upload your TOC")

uploaded_file = st.file_uploader(
    "Upload a file that contains the table of contents (DOCX or PDF)",
    type=["docx", "pdf"]
)

def extract_toc_from_docx(file):
    """Extract headings and subheadings from DOCX in a simple way."""
    doc = Document(file)
    toc_lines = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt and not txt.isdigit() and len(txt) > 2:
            toc_lines.append(txt)
    return "\n".join(toc_lines)

def extract_toc_from_pdf(file):
    """Extract readable lines from the first pages of a PDF."""
    reader = PdfReader(file)
    lines = []
    for page in reader.pages[:3]:
        text = page.extract_text()
        if text:
            for line in text.splitlines():
                line = line.strip()
                if 2 < len(line) < 120:
                    lines.append(line)
    return "\n".join(lines)

# Localized chapter keyword to enforce in AI refinement
def _chapter_word(lang_code: str) -> str:
    mapping = {"it": "Capitolo", "en": "Chapter", "es": "Capítulo", "fr": "Chapitre"}
    return mapping.get(lang_code, "Chapter")

if uploaded_file:
    # --- TOC extraction ---
    with st.spinner("Reading the TOC..."):
        if uploaded_file.name.endswith(".docx"):
            toc_text = extract_toc_from_docx(uploaded_file)
        else:
            toc_text = extract_toc_from_pdf(uploaded_file)

    # --- Language detection ---
    detected = "auto"
    if HAS_LANGID:
        try:
            detected = langid.classify(toc_text[:500])[0]
        except Exception:
            detected = "auto"
    st.session_state.detected_lang = detected

    # --- Show captured TOC ---
    st.success(f"Detected language: **{detected.upper()}**")

    # Initialize widget state if needed
    if "toc_text_editable" not in st.session_state:
        st.session_state["toc_text_editable"] = toc_text

    # Render editable TOC area (binds to session_state)
    st.text_area("Captured TOC:", key="toc_text_editable", height=300)

    # --- Action buttons ---
    col1, col2 = st.columns([1, 1])
    with col1:
        confirm_toc = st.button("✅ Confirm this TOC")
    with col2:
        refine_toc = st.button("🧠 Refine TOC with AI")

    # --- AI refinement ---
    if refine_toc and OPENAI_OK:
        lang_code = st.session_state.get("detected_lang", "en")
        chap_word = _chapter_word(lang_code if lang_code in ["it", "en", "es", "fr"] else "en")

        with st.spinner("Generating an improved TOC..."):
            prompt_refine = (
                "You are a professional non-fiction editor.\n"
                "Task: Clean up and balance the table of contents provided below.\n"
                f"- Normalize main headings as '{chap_word} 1', '{chap_word} 2', ...\n"
                "- Convert subsections to numeric scheme like 1.1, 1.2, 2.1, 2.2.\n"
                "- Ensure consistent casing and concise phrasing.\n"
                "- Keep meaning but improve clarity.\n"
                "- Output only the cleaned list, one heading per line.\n\n"
                f"Original TOC:\n{st.session_state['toc_text_editable']}"
            )
            resp = openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You refine and standardize book tables of contents."},
                    {"role": "user", "content": prompt_refine},
                ],
                temperature=0.4,
                max_tokens=800,
            )
            new_toc = (resp.choices[0].message.content or "").strip()
            st.session_state["toc_text_editable"] = new_toc
            st.info("AI proposal applied. You can still edit before confirming.")
            st.rerun()

    # --- Final confirmation ---
    if confirm_toc:
        st.session_state.confirmed_toc_text = st.session_state.toc_text_editable
        st.success("✅ TOC confirmed. You can proceed to allocation.")

# ==========================================
# 🧮 BLOCK 3 — WORD ALLOCATION & 500-WORD BLOCKING
# ------------------------------------------
# Take the confirmed TOC, collect book metadata, validate
# total words, and split into chapters/sections/blocks (≤500).
# ==========================================

st.subheader("🧩 Step 2 — Book data & allocation")

if st.session_state.confirmed_toc_text:

    # --- helpers ---

    def _strip_leading_markers(s: str) -> str:
        """Remove leading '-', '*', '1.', '1)', '1.1 ' etc., and extra spaces."""
        s = s.strip()
        s = re.sub(r"^[\-\*\u2022]+\s*", "", s)                 # bullets
        s = re.sub(r"^\d+[\.\)]\s*", "", s)                    # 1.  / 1)
        s = re.sub(r"^\d+\.\d+\s*", "", s)                     # 1.1
        s = re.sub(r"^\.\d+\s*", "", s)                        # .1  (fix user issue)
        return s.strip()

    def parse_confirmed_toc(toc_text: str) -> List[Chapter]:
        """
        Line-by-line TOC parser:
        - Plain line → Chapter
        - Lines starting with '-', '*', '1.'/'1)', '1.1', or indented → Section of current chapter
        - If a chapter has no sections, create 'Section 1'
        - Normalize subsection numbering to 1.1, 1.2, ... per chapter
        """
        lines = [ln.rstrip() for ln in toc_text.splitlines() if ln.strip()]
        chapters: List[Chapter] = []
        current: Optional[Chapter] = None
        chap_idx = 0

        def _is_section_line(ln: str) -> bool:
            if ln.startswith(("-", "*")):
                return True
            if re.match(r"^\s+\S", ln):                         # indentation
                return True
            if re.match(r"^\d+[\.\)]\s+", ln):                  # 1.  or 1)
                return True
            if re.match(r"^\d+\.\d+\s+", ln):                   # 1.1
                return True
            if re.match(r"^\.\d+\s+", ln):                      # .1  (user case)
                return True
            return False

        for raw in lines:
            ln = raw.strip()

            # Heuristic: if obviously a section marker → section
            is_section = _is_section_line(ln)

            # If not obviously a section and we have no current chapter, open a chapter
            if not is_section and current is None:
                chap_idx += 1
                current = Chapter(title=_strip_leading_markers(ln))
                chapters.append(current)
                continue

            # If not obviously a section but we do have a current chapter,
            # treat short, title-case lines as chapter starts; else fallback to section
            if not is_section:
                looks_like_chapter = (
                    len(ln.split()) <= 12 and ln[:1].islower() is False and not re.search(r"[.:;\-]$", ln)
                )
                if looks_like_chapter:
                    chap_idx += 1
                    current = Chapter(title=_strip_leading_markers(ln))
                    chapters.append(current)
                    continue
                else:
                    is_section = True

            # Section branch
            if current is None:
                chap_idx += 1
                current = Chapter(title=f"Chapter {chap_idx}")
                chapters.append(current)

            clean_title = _strip_leading_markers(ln)
            current.sections.append(Section(title=clean_title))

        # Ensure each chapter has at least one section
        for ch in chapters:
            if not ch.sections:
                ch.sections.append(Section(title="Section 1"))

        # Normalize subsection numbering per chapter: 1.1, 1.2, ...
        for ci, ch in enumerate(chapters, start=1):
            for si, sec in enumerate(ch.sections, start=1):
                # Prefix only if not already properly numbered
                if not re.match(rf"^{ci}\.{si}\s+", sec.title):
                    base = _strip_leading_markers(sec.title)
                    sec.title = f"{ci}.{si} {base}"
        return chapters

    def allocate_words(chapters: List[Chapter], total_words: int, block_size_limit: int = 500) -> List[Chapter]:
        """
        Even split:
        - distribute words across chapters
        - then across sections
        - blocks = ceil(words_per_section / 500), at least 1
        """
        n_ch = max(len(chapters), 1)
        base = total_words // n_ch
        rem = total_words % n_ch

        for i, ch in enumerate(chapters):
            ch.target_words = base + (1 if i < rem else 0)
            n_sec = max(len(ch.sections), 1)
            sec_base = ch.target_words // n_sec
            sec_rem = ch.target_words % n_sec
            ch_total_blocks = 0
            for j, sec in enumerate(ch.sections):
                sec.target_words = sec_base + (1 if j < sec_rem else 0)
                sec.blocks = max(1, math.ceil(sec.target_words / block_size_limit))
                ch_total_blocks += sec.blocks
            ch.blocks = ch_total_blocks
        return chapters

    # --- UI: language, tone, brief, formats, font ---
    lang_code = st.selectbox(
        "Generation language",
        ["auto", "it", "en", "es", "fr"],
        index=["auto", "it", "en", "es", "fr"].index(
            st.session_state.detected_lang if st.session_state.detected_lang in ["it", "en", "es", "fr"] else "auto"
        ),
        help="Leave 'auto' to use the detected language from the TOC, or force a specific language."
    )

    TONE_CHOICES_EN = ["Scientific", "Conversational", "Narrative"]
    tone = st.selectbox("Tone of voice", TONE_CHOICES_EN, index=TONE_CHOICES_EN.index("Conversational"))
    brief = st.text_area(
        "Brief (what the model should optimize for)",
        height=120,
        placeholder="Example: beginner audience, practical style, real examples, avoid heavy jargon..."
    )

    # LIMIT FORMATS to only 6x9 and 8.5x11
    pdf_page = st.selectbox("Page size", ["6x9", "8.5x11"], index=0)

    font_name = st.selectbox("Primary font", FONT_CHOICES, index=0)

    # --- Book metadata form ---
    with st.form("book_info_form"):
        title = st.text_input("Book title")
        subtitle = st.text_input("Subtitle")
        author = st.text_input("Author")
        total_words = st.number_input("Total target words", min_value=500, step=500, value=20000)
        submitted_meta = st.form_submit_button("Save book data")

    if submitted_meta:
        st.success("Book data saved.")

    # --- Minimal validation: enough words for sections? ---
    temp_chapters = parse_confirmed_toc(st.session_state.confirmed_toc_text)
    total_sections = sum(len(ch.sections) for ch in temp_chapters)
    min_needed = total_sections * MIN_SECTION_WORDS_USEFUL  # 250 words/section minimum

    if total_words < min_needed:
        st.error(
            f"With {total_sections} sections, you need at least {min_needed} words "
            f"(250 per section). Increase the total or reduce sections in the TOC."
        )
        st.stop()

    # --- Allocation and on-screen preview ---
    chapters_alloc = allocate_words(temp_chapters, total_words, block_size_limit=MAX_SUBGEN_WORDS)

    plan_preview = BookPlan(
        title=title or "Title",
        subtitle=subtitle or "",
        author=author or "",
        total_words=total_words,
        block_size=MAX_SUBGEN_WORDS,
        chapters=chapters_alloc,
        language_code=lang_code,
        tone=tone,
        brief=brief.strip(),
        pdf_page=pdf_page,
        font_name=font_name,
    )

    st.session_state.generated_plan = plan_preview
    st.session_state.chapters = chapters_alloc
    st.session_state.allocation_done = True

    st.success("✅ Allocation ready. Here is the split:")
    for i, ch in enumerate(chapters_alloc, start=1):
        with st.expander(f"Chapter {i}: {ch.title} — {ch.target_words} words — {ch.blocks} total blocks"):
            for j, sec in enumerate(ch.sections, start=1):
                st.write(f"• Section {j}: {sec.title} — {sec.target_words} words — {sec.blocks} blocks (≤500 words each)")

    st.info("When satisfied, proceed to content generation.")
else:
    st.warning("Please confirm the TOC in Step 1 first.")
# ----- SAFETY HELPERS: define if missing -----
if 'generate_block_text' not in globals():
    def _effective_language_label(plan: BookPlan) -> str:
        code = plan.language_code
        if code == "auto":
            det = st.session_state.get("detected_lang", "en")
            code = det if det in LANG_LABELS else "en"
        return LANG_LABELS.get(code, "English")

    def _tone_instruction(tone: str) -> str:
        t = tone.lower()
        if t.startswith("scien"): return "Use a precise, rigorous, evidence-based tone."
        if t.startswith("narr"):  return "Use a narrative, evocative tone with smooth transitions."
        return "Use a clear, friendly, practical tone."

    def _generate_subchunk(prompt_sys: str, prompt_user: str) -> str:
        if not OPENAI_OK:
            return " ".join(["[placeholder text]"] * 50)
        try:
            resp = openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "system", "content": prompt_sys},
                          {"role": "user", "content": prompt_user}],
                temperature=0.7,
            )
            return (resp.choices[0].message.content or "").strip()
        except Exception as e:
            return f"[generation error] {e}"

    def generate_block_text(plan: BookPlan, ch_title: str, sec_title: str,
                            target_words: int, prev_summary: str = "",
                            is_last_block: bool = False) -> str:
        lang = _effective_language_label(plan)
        tone_ins = _tone_instruction(plan.tone)
        n_sub = max(1, math.ceil(target_words / MAX_SUBGEN_WORDS))
        words_per_sub = min(math.ceil(target_words / n_sub), MAX_SUBGEN_WORDS)

        sys = (
            "You are an expert non-fiction writer. "
            f"Write in {lang}. {tone_ins} Avoid repetition. "
            "Do not restate the titles inside the text. Use continuous prose."
        )
        parts = []
        for idx in range(n_sub):
            role_note = "Start the section naturally." if idx == 0 else "Continue smoothly from the previous text."
            if idx == n_sub - 1 and is_last_block:
                role_note += " Conclude the section with a natural closing."
            guidance = []
            if plan.brief: guidance.append(f"Brief to follow: {plan.brief}")
            if prev_summary: guidance.append(f"Previous context summary: {prev_summary}")
            user = (
                f"Book title: {plan.title}\nSubtitle: {plan.subtitle}\nAuthor: {plan.author}\n"
                f"Chapter: {ch_title}\nSection: {sec_title}\n"
                f"Goal words for this part: ~{words_per_sub} (hard limit per request: 500)\n{role_note}\n"
                + ("\n".join(guidance) if guidance else "")
            )
            subtext = _generate_subchunk(sys, user)
            parts.append(subtext.strip())
        return " ".join(p for p in parts if p).strip()

if 'generate_all_sections' not in globals():
    def generate_all_sections(plan: BookPlan):
        total_blocks = sum(sec.blocks for ch in plan.chapters for sec in ch.sections)
        done = 0
        bar = st.progress(0, text="Writing in progress...")
        prev_summary = ""
        for ch in plan.chapters:
            for sec in ch.sections:
                sec.texts = []
                block_target = max(1, math.ceil(sec.target_words / max(1, sec.blocks)))
                for b in range(sec.blocks):
                    is_last = (b == sec.blocks - 1)
                    text = generate_block_text(
                        plan=plan,
                        ch_title=ch.title,
                        sec_title=sec.title,
                        target_words=block_target,
                        prev_summary=prev_summary,
                        is_last_block=is_last
                    )
                    sec.texts.append(text)
                    words = re.split(r"\s+", text.strip())
                    prev_summary = " ".join(words[:60]) + " ... " + " ".join(words[-40:]) if len(words) > 120 else text[:800]
                    done += 1
                    bar.progress(done / total_blocks, text=f"Blocks completed: {done}/{total_blocks}")
        bar.empty()
        st.success("✅ Content generation completed.")

# ==========================================
# ✍️ BLOCK 4 — GENERATION + DOCX/PDF EXPORT
# ------------------------------------------
# Generate content (≤500 words per API call), show preview,
# and export DOCX/PDF with proper formatting, ToC, and options.
# ==========================================

st.subheader("🖋️ Step 3 — Content generation & export")

# ---------- helpers already defined above are reused ----------
# _effective_language_label(), _tone_instruction(), _count_words(),
# _generate_subchunk(), generate_block_text(), generate_all_sections()

# Extra imports needed for this block
from docx.shared import Cm  # margins
from docx.oxml.ns import qn
from docx.enum.section import WD_SECTION_START
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY

# Safe filename helper: Title_Subtitle
def _safe_filename(plan: BookPlan) -> str:
    base = f"{plan.title.strip()}_{plan.subtitle.strip()}" if plan.subtitle.strip() else plan.title.strip()
    base = re.sub(r"[^\w\-]+", "_", base).strip("_")
    base = re.sub(r"_+", "_", base)
    return base or "book"

def _add_docx_toc(doc):
    """
    Insert a Word ToC field. Word will update fields at open (or via F9).
    """
    # Place a heading for TOC
    p = doc.add_paragraph()
    run = p.add_run("Table of Contents")
    run.bold = True
    run.font.size = Pt(16)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Field code: TOC \o "1-3" \h \z \u
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as _qn
    p = doc.add_paragraph()
    fld = OxmlElement('w:fldSimple')
    fld.set(_qn('w:instr'), r'TOC \o "1-3" \h \z \u')
    p._p.append(fld)

def build_docx(plan: BookPlan, include_toc: bool = False, include_copyright: bool = False) -> bytes:
    """
    DOCX:
      - Title page: centered title/subtitle; author centered lower on page.
      - Optional: Copyright/Disclaimer page.
      - Optional: Table of Contents (Word updates page numbers at open).
      - Body: headings left; normal text JUSTIFY.
      - Page size: 6x9 or 8.5x11; margins 2.54 cm if 8.5x11.
      - Font: set Normal style to chosen font.
    """
    doc = Document()

    # --- page setup (first section) ---
    sec = doc.sections[0]
    if plan.pdf_page == "6x9":
        sec.page_width, sec.page_height = Inches(6), Inches(9)
    else:  # "8.5x11"
        sec.page_width, sec.page_height = Inches(8.5), Inches(11)
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(2.54)

    # --- Title page (centered) ---
    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(plan.title)
    run.bold = True
    run.font.size = Pt(26)

    # Subtitle
    if plan.subtitle.strip():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(plan.subtitle)
        run.bold = False
        run.font.size = Pt(16)

    # Push author lower on the page
    for _ in range(12):
        doc.add_paragraph("")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(plan.author)
    run.bold = False
    run.font.size = Pt(12)

    doc.add_page_break()

    # --- Optional: Copyright/Disclaimer page ---
    if include_copyright:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("Copyright & Disclaimer")
        r.bold = True
        r.font.size = Pt(14)

        doc.add_paragraph("")
        copy_txt = (
            f"© {plan.author}. All rights reserved.\n\n"
            "No part of this publication may be reproduced, distributed, or transmitted in any form or by any means, "
            "including photocopying, recording, or other electronic or mechanical methods, without the prior written "
            "permission of the publisher, except in the case of brief quotations embodied in critical reviews.\n\n"
            "Disclaimer: The information in this book is provided for educational purposes only and does not constitute "
            "professional advice. Always consult a qualified professional for your specific situation."
        )
        para = doc.add_paragraph(copy_txt)
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        doc.add_page_break()

    # --- Optional: Table of Contents (field) ---
    if include_toc:
        _add_docx_toc(doc)
        doc.add_page_break()

    # --- Normal style & font ---
    style = doc.styles["Normal"]
    style.font.name = plan.font_name
    try:
        style._element.rPr.rFonts.set(qn('w:eastAsia'), plan.font_name)  # type: ignore
    except Exception:
        pass

    # --- Body content ---
    def _docx_heading(text, size_pt, bold=True):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(text)
        r.bold = bold
        r.font.size = Pt(size_pt)

    for ch in plan.chapters:
        _docx_heading(ch.title, 18, bold=True)
        for sec_obj in ch.sections:
            _docx_heading(sec_obj.title, 14, bold=False)
            for text in sec_obj.texts:
                para = doc.add_paragraph(text)
                para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        doc.add_page_break()

    # ensure Normal paragraphs justify (safety)
    for paragraph in doc.paragraphs:
        if paragraph.style.name == "Normal":
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# PDF font mapping
PDF_FONT_MAP = {
    "Times New Roman": "Times-Roman",
    "Roboto": "Helvetica",
    "Comfortaa": "Courier"
}

# Custom DocTemplate to collect headings for ToC
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
class _TocDocTemplate(SimpleDocTemplate):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.heading_entries = []

    def afterFlowable(self, flowable):
        if isinstance(flowable, Paragraph):
            style_name = getattr(flowable.style, "name", "")
            if style_name in ("H1", "H2"):
                level = 0 if style_name == "H1" else 1
                text = flowable.getPlainText()
                page = self.canv.getPageNumber()
                self.notify('TOCEntry', (level, text, page))

def build_pdf(plan: BookPlan, include_toc: bool = False, include_copyright: bool = False) -> bytes:
    """
    PDF:
      - Page size: 6x9 or 8.5x11
      - Margins: 2.54 cm if 8.5x11, else 2 cm
      - Title page centered; author lower
      - Optional: Copyright/Disclaimer page
      - Optional: Table of Contents with page numbers (Platypus ToC)
    """
    pagesize = PAGE_SIZES.get(plan.pdf_page, PAGE_SIZES["6x9"])
    buf = io.BytesIO()
    if plan.pdf_page == "8.5x11":
        lm = rm = tm = bm = 2.54 * cm
    else:
        lm = rm = tm = bm = 2 * cm

    doc = _TocDocTemplate(buf, pagesize=pagesize, leftMargin=lm, rightMargin=rm, topMargin=tm, bottomMargin=bm)
    styles = getSampleStyleSheet()
    base_font = PDF_FONT_MAP.get(plan.font_name, "Times-Roman")

    # Styles
    H1 = ParagraphStyle("H1", parent=styles["Heading1"], fontName=base_font, alignment=TA_LEFT, spaceBefore=12, spaceAfter=6)
    H2 = ParagraphStyle("H2", parent=styles["Heading2"], fontName=base_font, alignment=TA_LEFT, spaceBefore=6, spaceAfter=4)
    Body = ParagraphStyle("Body", parent=styles["BodyText"], fontName=base_font, alignment=TA_JUSTIFY, leading=14)
    TitleC = ParagraphStyle("TitleC", parent=styles["Title"], fontName=base_font, alignment=TA_CENTER, spaceAfter=12)
    SubC = ParagraphStyle("SubC", parent=styles["BodyText"], fontName=base_font, alignment=TA_CENTER, spaceAfter=24)
    AuthorC = ParagraphStyle("AuthorC", parent=styles["BodyText"], fontName=base_font, alignment=TA_CENTER, spaceBefore=12)

    story = []

    # Title page
    story.append(Spacer(1, 40))
    story.append(Paragraph(plan.title, TitleC))
    if plan.subtitle.strip():
        story.append(Paragraph(plan.subtitle, SubC))
    story.append(Spacer(1, pagesize[1] * 0.55))
    story.append(Paragraph(plan.author, AuthorC))
    story.append(PageBreak())

    # Optional: Copyright/Disclaimer
    if include_copyright:
        story.append(Paragraph("Copyright & Disclaimer", ParagraphStyle("CPH", parent=styles["Heading2"], fontName=base_font, alignment=TA_CENTER)))
        story.append(Spacer(1, 12))
        copy_txt = (
            f"© {plan.author}. All rights reserved.<br/><br/>"
            "No part of this publication may be reproduced, distributed, or transmitted in any form or by any means, "
            "including photocopying, recording, or other electronic or mechanical methods, without the prior written "
            "permission of the publisher, except in the case of brief quotations embodied in critical reviews.<br/><br/>"
            "Disclaimer: The information in this book is provided for educational purposes only and does not constitute "
            "professional advice. Always consult a qualified professional for your specific situation."
        )
        story.append(Paragraph(copy_txt, Body))
        story.append(PageBreak())

    # Optional: Table of Contents (ReportLab ToC)
    if include_toc:
        toc = TableOfContents()
        toc.levelStyles = [
            ParagraphStyle(fontName=base_font, name='TOCHeading1', leftIndent=20, firstLineIndent=-10, spaceBefore=6, leading=12),
            ParagraphStyle(fontName=base_font, name='TOCHeading2', leftIndent=36, firstLineIndent=-10, spaceBefore=4, leading=12),
        ]
        story.append(Paragraph("Table of Contents", ParagraphStyle("TOCTitle", parent=styles["Heading1"], fontName=base_font, alignment=TA_CENTER)))
        story.append(Spacer(1, 12))
        story.append(toc)
        story.append(PageBreak())

    # Content
    for ch in plan.chapters:
        story.append(Paragraph(ch.title, H1))
        for sec_obj in ch.sections:
            story.append(Paragraph(sec_obj.title, H2))
            for text in sec_obj.texts:
                story.append(Paragraph(text, Body))
                story.append(Spacer(1, 8))
        story.append(PageBreak())

    doc.build(story)
    return buf.getvalue()

# ----- UI: generate + preview + downloads + options -----
if st.session_state.allocation_done and st.session_state.generated_plan:
    plan: BookPlan = st.session_state.generated_plan

    # Options
    c1, c2 = st.columns(2)
    with c1:
        opt_toc = st.checkbox("Include Table of Contents with page numbers", value=True)
    with c2:
        opt_copyright = st.checkbox("Include Copyright & Disclaimer page", value=False)

    if st.button("🚀 CONFIRM AND GENERATE CONTENT", type="primary", use_container_width=True):
        generate_all_sections(plan)
        try:
            st.session_state["docx_bytes"] = build_docx(plan, include_toc=opt_toc, include_copyright=opt_copyright)
            st.session_state["pdf_bytes"] = build_pdf(plan, include_toc=opt_toc, include_copyright=opt_copyright)
        except Exception as e:
            st.error(f"Export error: {e}")

    # Preview (snippets)
    if any(sec.texts for ch in plan.chapters for sec in ch.sections):
        st.subheader("👁️ Preview (snippets)")
        max_preview = 3
        shown = 0
        for i, ch in enumerate(plan.chapters, start=1):
            for j, sec in enumerate(ch.sections, start=1):
                if sec.texts:
                    st.markdown(f"**Chapter {i} — {ch.title}**  \n*Section {j} — {sec.title}*")
                    st.write((sec.texts[0][:1200] + "…") if len(sec.texts[0]) > 1200 else sec.texts[0])
                    st.divider()
                    shown += 1
                    if shown >= max_preview:
                        break
            if shown >= max_preview:
                break

    # Downloads with Title_Subtitle filenames
    if st.session_state.get("docx_bytes") and st.session_state.get("pdf_bytes"):
        st.subheader("📥 Download your book")
        fname = _safe_filename(plan)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="Download DOCX",
                data=st.session_state["docx_bytes"],
                file_name=f"{fname}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        with c2:
            st.download_button(
                label="Download PDF",
                data=st.session_state["pdf_bytes"],
                file_name=f"{fname}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
else:
    st.info("Complete the previous steps to generate and download the book.")

