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

    # ---------- PENDING PATCH (MUST RUN BEFORE WIDGET RENDERING) ----------
    # If AI refinement set a pending value, apply it now (before widget creation)
    if "toc_text_editable_pending" in st.session_state:
        st.session_state["toc_text_editable"] = st.session_state["toc_text_editable_pending"]
        del st.session_state["toc_text_editable_pending"]

    # Initialize widget state if first time
    if "toc_text_editable" not in st.session_state:
        st.session_state["toc_text_editable"] = toc_text
    # ----------------------------------------------------------------------

    # --- Show captured TOC ---
    st.success(f"Detected language: **{detected.upper()}**")
    st.text_area("Captured TOC:", key="toc_text_editable", height=300)

    # --- Action buttons ---
    col1, col2 = st.columns([1, 1])
    with col1:
        confirm_toc = st.button("✅ Confirm this TOC")
    with col2:
        refine_toc = st.button("🧠 Refine TOC with AI")

    # --- AI refinement (use PENDING to avoid Streamlit widget-key assignment error) ---
    if refine_toc and OPENAI_OK:
        lang_code = st.session_state.get("detected_lang", "en")
        chap_word = _chapter_word(lang_code if lang_code in ["it", "en", "es", "fr"] else "en")

        with st.spinner("Generating an improved TOC..."):
            prompt_refine = (
                "You are a professional non-fiction editor.\n"
                "Task: Clean up and balance the table of contents provided below.\n"
                f"- Normalize main headings as '{chap_word} 1', '{chap_word} 2', ...\n"
                "- Convert subsections to a numeric scheme like 1.1, 1.2, 2.1, 2.2.\n"
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

            # Save to PENDING, then rerun so we apply it BEFORE widget render
            st.session_state["toc_text_editable_pending"] = new_toc
            st.info("AI proposal applied. Reloading the editor...")
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

    # ---------- helpers ----------

    def _strip_leading_markers(s: str) -> str:
        """
        Remove any leading numbering/bullets:
        - '-', '*', '•'
        - '1.' / '1)' / '1 -' / '1.' with spaces
        - '1.1', '1.1)', '1.1 -', '1.1.' (also '1.1.1'...)
        - '.1' (edge case)
        - extra spaces
        """
        s = s.strip()
        # bullets
        s = re.sub(r"^[\-\*\u2022]+\s*", "", s)
        # numeric sequences like 1. , 1) , 1.1 , 1.1.1) , 2.3. , 3) :
        s = re.sub(r"^\d+(?:\.\d+){0,3}[\.\)\-:]?\s*", "", s)
        # leading '.1 '
        s = re.sub(r"^\.\d+\s*", "", s)
        return s.strip()

    def parse_confirmed_toc(toc_text: str) -> List[Chapter]:
        """
        - Plain line → Chapter
        - Bullet/indent/numbered line → Section of current chapter
        - Ensure every chapter has at least one section
        - Normalize section numbering to '1.1', '1.2', ... (no double prefixes)
        """
        lines = [ln.rstrip() for ln in toc_text.splitlines() if ln.strip()]
        chapters: List[Chapter] = []
        current: Optional[Chapter] = None
        chap_idx = 0

        def _is_section_line(ln: str) -> bool:
            if ln.startswith(("-", "*")): return True
            if re.match(r"^\s+\S", ln): return True                       # indentation
            if re.match(r"^\d+[\.\)]\s+", ln): return True                # 1.  / 1)
            if re.match(r"^\d+(?:\.\d+){1,3}[\.\)]?\s+", ln): return True # 1.1 / 1.1.1)
            if re.match(r"^\.\d+\s+", ln): return True                    # .1
            return False

        for raw in lines:
            ln = raw.strip()
            is_section = _is_section_line(ln)

            if not is_section and current is None:
                chap_idx += 1
                current = Chapter(title=_strip_leading_markers(ln))
                chapters.append(current)
                continue

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

            if current is None:
                chap_idx += 1
                current = Chapter(title=f"Chapter {chap_idx}")
                chapters.append(current)

            clean_title = _strip_leading_markers(ln)
            current.sections.append(Section(title=clean_title))

        # ensure at least one section per chapter
        for ch in chapters:
            if not ch.sections:
                ch.sections.append(Section(title="Section 1"))

        # Normalize numbering to 'ci.si Title' without duplicating existing prefixes
        for ci, ch in enumerate(chapters, start=1):
            for si, sec in enumerate(ch.sections, start=1):
                if re.match(rf"^{ci}\.{si}\s+", sec.title):
                    continue  # already correct
                base = re.sub(r"^\d+(?:\.\d+){1,3}[\.\)]?\s*", "", sec.title).strip()
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

    # ---------- UI (EN) ----------
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

# ==========================================
# ✍️ BLOCK 4 — CONTENT GENERATION & EXPORT
# ------------------------------------------
# Generate content (≤500 words/request), preview snippets,
# and export DOCX/PDF with proper styles, ToC, and options.
# ==========================================

st.subheader("🖋️ Step 3 — Content generation & export")

# ---------- EXTRA IMPORTS ----------
from docx.shared import Cm
from docx.oxml.ns import qn
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY

# ---------- GENERATION HELPERS ----------
def _effective_language_label(plan: BookPlan) -> str:
    code = plan.language_code
    if code == "auto":
        det = st.session_state.get("detected_lang", "en")
        code = det if det in LANG_LABELS else "en"
    return LANG_LABELS.get(code, "English")

def _tone_instruction(tone: str) -> str:
    t = (tone or "").lower()
    if t.startswith("scien"): return "Use a precise, rigorous, evidence-based tone."
    if t.startswith("narr"):  return "Use a narrative, evocative tone with smooth transitions."
    return "Use a clear, friendly, and practical tone."

def _generate_subchunk(prompt_sys: str, prompt_user: str) -> str:
    if not OPENAI_OK:
        return " ".join(["[placeholder text]"] * 50)
    try:
        resp = openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": prompt_sys},
                {"role": "user", "content": prompt_user},
            ],
            temperature=0.7,
        )
        return (resp.choices[0].message.content or "").strip()
    except Exception as e:
        return f"[generation error] {e}"

def generate_block_text(plan: BookPlan, ch_title: str, sec_title: str, target_words: int,
                        prev_summary: str = "", is_last_block: bool = False) -> str:
    lang = _effective_language_label(plan)
    tone_ins = _tone_instruction(plan.tone)
    n_sub = max(1, math.ceil(target_words / MAX_SUBGEN_WORDS))
    words_per_sub = min(math.ceil(target_words / n_sub), MAX_SUBGEN_WORDS)
    sys = (
        "You are an expert non-fiction writer. "
        f"Write in {lang}. {tone_ins} Avoid repetition. "
        "Do not restate book/chapter/section titles. "
        "Write continuous prose (no lists unless necessary)."
    )
    parts = []
    for idx in range(n_sub):
        note = "Start naturally." if idx == 0 else "Continue smoothly."
        if idx == n_sub - 1 and is_last_block:
            note += " Conclude naturally."
        context = []
        if plan.brief: context.append(f"Brief: {plan.brief}")
        if prev_summary: context.append(f"Previous context: {prev_summary}")
        user = (
            f"Book title: {plan.title}\nSubtitle: {plan.subtitle}\nAuthor: {plan.author}\n"
            f"Chapter: {ch_title}\nSection: {sec_title}\nTarget: ~{words_per_sub} words\n"
            f"{note}\n" + ("\n".join(context) if context else "")
        )
        txt = _generate_subchunk(sys, user)
        parts.append(txt.strip())
    return " ".join(parts).strip()

def generate_all_sections(plan: BookPlan):
    total_blocks = sum(sec.blocks for ch in plan.chapters for sec in ch.sections)
    if total_blocks <= 0:
        st.warning("No blocks to generate. Check your allocation.")
        return
    bar = st.progress(0, text="Writing in progress...")
    done, prev_summary = 0, ""
    for ch in plan.chapters:
        for sec in ch.sections:
            sec.texts = []
            block_target = max(1, math.ceil(sec.target_words / max(1, sec.blocks)))
            for b in range(sec.blocks):
                text = generate_block_text(plan, ch.title, sec.title,
                                           target_words=block_target,
                                           prev_summary=prev_summary,
                                           is_last_block=(b == sec.blocks - 1))
                sec.texts.append(text)
                words = re.split(r"\s+", text.strip())
                prev_summary = (
                    " ".join(words[:60]) + " ... " + " ".join(words[-40:])
                    if len(words) > 120 else text[:800]
                )
                done += 1
                bar.progress(done / total_blocks, text=f"Blocks completed: {done}/{total_blocks}")
    bar.empty()
    st.success("✅ Content generation completed.")

# ---------- EXPORT HELPERS ----------
PDF_FONT_MAP = {"Times New Roman": "Times-Roman", "Roboto": "Helvetica", "Comfortaa": "Courier"}

def _safe_filename(plan: BookPlan) -> str:
    base = f"{plan.title.strip()}_{plan.subtitle.strip()}" if plan.subtitle.strip() else plan.title.strip()
    base = re.sub(r"[^\w\-]+", "_", base).strip("_")
    return re.sub(r"_+", "_", base) or "book"

def _add_docx_toc(doc):
    p = doc.add_paragraph()
    run = p.add_run("Table of Contents")
    run.bold = True
    run.font.size = Pt(16)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as _qn
    p = doc.add_paragraph()
    fld = OxmlElement("w:fldSimple")
    fld.set(_qn("w:instr"), r'TOC \o "1-3" \h \z \u')
    p._p.append(fld)

def build_docx(plan: BookPlan, include_toc=True, include_copyright=False) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    if plan.pdf_page == "8.5x11":
        sec.page_width, sec.page_height = Inches(8.5), Inches(11)
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(2.54)
    else:
        sec.page_width, sec.page_height = Inches(6), Inches(9)
    # title page
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(plan.title); r.bold = True; r.font.size = Pt(26)
    if plan.subtitle.strip():
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(plan.subtitle); r.font.size = Pt(16)
    for _ in range(12): doc.add_paragraph("")
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(plan.author); r.font.size = Pt(12)
    doc.add_page_break()
    # copyright
    if include_copyright:
        p = doc.add_paragraph("Copyright & Disclaimer"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        para = doc.add_paragraph(
            f"© {plan.author}. All rights reserved.\n\n"
            "No part of this publication may be reproduced or distributed without permission.\n\n"
            "Disclaimer: Educational use only; not professional advice."
        ); para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        doc.add_page_break()
    # toc
    if include_toc:
        _add_docx_toc(doc); doc.add_page_break()
    # font
    style = doc.styles["Normal"]; style.font.name = plan.font_name
    try: style._element.rPr.rFonts.set(qn("w:eastAsia"), plan.font_name)
    except Exception: pass
    # content
    for ch in plan.chapters:
        p = doc.add_paragraph(ch.title); p.style = doc.styles["Heading 1"]
        for sec in ch.sections:
            p = doc.add_paragraph(sec.title); p.style = doc.styles["Heading 2"]
            for text in sec.texts:
                para = doc.add_paragraph(text)
                para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        doc.add_page_break()
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

# ---------- PDF BUILD ----------
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
class _TocDocTemplate(SimpleDocTemplate):
    def afterFlowable(self, f):
        if isinstance(f, Paragraph):
            nm = getattr(f.style, "name", "")
            if nm in ("H1", "H2"):
                lvl = 0 if nm == "H1" else 1
                self.notify("TOCEntry", (lvl, f.getPlainText(), self.canv.getPageNumber()))

def build_pdf(plan: BookPlan, include_toc=True, include_copyright=False) -> bytes:
    pagesize = PAGE_SIZES.get(plan.pdf_page, PAGE_SIZES["6x9"])
    buf = io.BytesIO()
    m = 2.54 * cm if plan.pdf_page == "8.5x11" else 2 * cm
    doc = _TocDocTemplate(buf, pagesize=pagesize, leftMargin=m, rightMargin=m, topMargin=m, bottomMargin=m)
    styles = getSampleStyleSheet(); fnt = PDF_FONT_MAP.get(plan.font_name, "Times-Roman")
    H1 = ParagraphStyle("H1", parent=styles["Heading1"], fontName=fnt, alignment=TA_LEFT)
    H2 = ParagraphStyle("H2", parent=styles["Heading2"], fontName=fnt, alignment=TA_LEFT)
    Body = ParagraphStyle("Body", parent=styles["BodyText"], fontName=fnt, alignment=TA_JUSTIFY)
    TitleC = ParagraphStyle("TitleC", parent=styles["Title"], fontName=fnt, alignment=TA_CENTER)
    SubC = ParagraphStyle("SubC", parent=styles["BodyText"], fontName=fnt, alignment=TA_CENTER)
    story = []
    # title
    story += [Spacer(1, 40), Paragraph(plan.title, TitleC)]
    if plan.subtitle.strip(): story.append(Paragraph(plan.subtitle, SubC))
    story += [Spacer(1, pagesize[1]*0.55), Paragraph(plan.author, SubC), PageBreak()]
    # copyright
    if include_copyright:
        story += [Paragraph("Copyright & Disclaimer", H2),
                  Paragraph(f"© {plan.author}. Educational use only; not professional advice.", Body),
                  PageBreak()]
    # toc
    if include_toc:
        toc = TableOfContents()
        toc.levelStyles = [ParagraphStyle(fontName=fnt, name="TOC1", leftIndent=20, firstLineIndent=-10),
                           ParagraphStyle(fontName=fnt, name="TOC2", leftIndent=36, firstLineIndent=-10)]
        story += [Paragraph("Table of Contents", H1), Spacer(1,12), toc, PageBreak()]
    # content
    for ch in plan.chapters:
        story.append(Paragraph(ch.title, H1))
        for sec in ch.sections:
            story.append(Paragraph(sec.title, H2))
            for text in sec.texts:
                story.append(Paragraph(text, Body))
                story.append(Spacer(1, 8))
        story.append(PageBreak())
    doc.build(story); return buf.getvalue()

# ---------- UI ----------
if st.session_state.allocation_done and st.session_state.generated_plan:
    plan: BookPlan = st.session_state.generated_plan
    c1, c2 = st.columns(2)
    with c1: opt_toc = st.checkbox("Include Table of Contents", value=True)
    with c2: opt_copy = st.checkbox("Include Copyright page", value=False)
    if st.button("🚀 CONFIRM AND GENERATE CONTENT", type="primary", use_container_width=True):
        generate_all_sections(plan)
        try:
            st.session_state["docx_bytes"] = build_docx(plan, opt_toc, opt_copy)
            st.session_state["pdf_bytes"] = build_pdf(plan, opt_toc, opt_copy)
        except Exception as e:
            st.error(f"Export error: {e}")
    if any(sec.texts for ch in plan.chapters for sec in ch.sections):
        st.subheader("👁️ Preview (snippets)")
        shown = 0
        for ch in plan.chapters:
            for sec in ch.sections:
                if sec.texts:
                    st.markdown(f"**{ch.title}** — *{sec.title}*")
                    st.write((sec.texts[0][:1200]+"…") if len(sec.texts[0])>1200 else sec.texts[0])
                    st.divider()
                    shown += 1
                    if shown >= 3: break
            if shown >= 3: break
    if st.session_state.get("docx_bytes") and st.session_state.get("pdf_bytes"):
        fname = _safe_filename(plan)
        c1, c2 = st.columns(2)
        with c1: st.download_button("Download DOCX", st.session_state["docx_bytes"], f"{fname}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        with c2: st.download_button("Download PDF", st.session_state["pdf_bytes"], f"{fname}.pdf",
                                    mime="application/pdf", use_container_width=True)
else:
    st.info("Complete the previous steps to generate and download the book.")
