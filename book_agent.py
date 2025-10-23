import io
import math
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn
from PyPDF2 import PdfReader
from reportlab.lib.pagesizes import A4, letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm


# ============ Data models ============

@dataclass
class Section:
    title: str
    target_words: int = 0
    blocks: int = 0


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
    total_words: int
    block_size: int
    chapters: List[Chapter] = field(default_factory=list)


# ============ TOC extraction utilities ============

H1_PATTERNS = [
    r"^\s*(?:chapter|capitolo)\s+\d+[:.\-\s]+(.+)$",
    r"^\s*\d+\.\s+(.+)$",
    r"^\s*[IVXLC]+\.\s+(.+)$",
]
H2_PATTERNS = [
    r"^\s*(?:section|sezione)\s+\d+(?:\.\d+)?[:.\-\s]+(.+)$",
    r"^\s*\d+\.\d+\.\s+(.+)$",
    r"^\s*\d+\)\s+(.+)$",
]

def _match_first(text: str, patterns: List[str]) -> Optional[str]:
    for pat in patterns:
        m = re.match(pat, text, flags=re.IGNORECASE)
        if m:
            g = m.group(1).strip()
            if g:
                return g
    return None

def normalize_heading(text: str) -> str:
    t = re.sub(r"\s+", " ", text or "").strip()
    # Rimuovi puntini di riempimento e numeri pagina tipici degli indici PDF
    t = re.sub(r"\.{2,}\s*\d+$", "", t).strip()
    return t

def guess_is_heading(line: str) -> bool:
    # Heuristics: line with few words and Title Case or ALL CAPS often indicates heading
    clean = normalize_heading(line)
    if not clean:
        return False
    words = clean.split()
    if len(words) <= 12:
        if clean.isupper():
            return True
        titlecase_ratio = sum(1 for w in words if w[:1].isupper()) / max(len(words), 1)
        if titlecase_ratio >= 0.6:
            return True
    # Numeric or roman numerals prefix
    if re.match(r"^\s*(?:\d+|[IVXLC]+)[\.\)\s]", clean):
        return True
    return False

def extract_toc_from_docx(file_bytes: bytes) -> List[Chapter]:
    doc = Document(io.BytesIO(file_bytes))
    chapters: List[Chapter] = []
    current_ch: Optional[Chapter] = None

    for p in doc.paragraphs:
        text = normalize_heading(p.text)
        if not text:
            continue

        style_name = (getattr(p.style, "name", "") or "").lower()

        if "heading 1" in style_name or _match_first(text, H1_PATTERNS) or guess_is_heading(text):
            # If style says H1, prefer raw text; otherwise, use matched group if present
            h = text
            matched = _match_first(text, H1_PATTERNS)
            if matched:
                h = matched
            current_ch = Chapter(title=h)
            chapters.append(current_ch)
            continue

        if current_ch:
            if "heading 2" in style_name or _match_first(text, H2_PATTERNS):
                h = text
                matched = _match_first(text, H2_PATTERNS)
                if matched:
                    h = matched
                current_ch.sections.append(Section(title=h))
    # Guarantee at least one section per chapter
    for ch in chapters:
        if not ch.sections:
            ch.sections.append(Section(title="Sezione 1"))
    return chapters

def extract_toc_from_pdf(file_bytes: bytes) -> List[Chapter]:
    reader = PdfReader(io.BytesIO(file_bytes))
    lines: List[str] = []
    for page in reader.pages:
        try:
            txt = page.extract_text() or ""
        except Exception:
            txt = ""
        for raw in txt.splitlines():
            s = normalize_heading(raw)
            if s:
                lines.append(s)

    chapters: List[Chapter] = []
    current_ch: Optional[Chapter] = None

    for ln in lines:
        # Try H1 patterns
        h1 = _match_first(ln, H1_PATTERNS)
        h2 = _match_first(ln, H2_PATTERNS)

        if h1:
            current_ch = Chapter(title=h1)
            chapters.append(current_ch)
            continue

        if h2 and current_ch:
            current_ch.sections.append(Section(title=h2))
            continue

        # Fallback heuristic if no explicit patterns
        if guess_is_heading(ln):
            if not current_ch:
                current_ch = Chapter(title=ln)
                chapters.append(current_ch)
            else:
                # If chapter exists and last action was chapter without sections, treat as section
                current_ch.sections.append(Section(title=ln))

    # Guarantee at least one section per chapter
    for ch in chapters:
        if not ch.sections:
            ch.sections.append(Section(title="Sezione 1"))
    return chapters


# ============ Allocation ============

def allocate_words(chapters: List[Chapter], total_words: int, block_size: int) -> List[Chapter]:
    if total_words <= 0:
        total_words = 1
    if block_size <= 0:
        block_size = 500

    n_ch = max(len(chapters), 1)
    base_per_ch = total_words // n_ch
    remainder = total_words % n_ch

    for idx, ch in enumerate(chapters):
        ch_words = base_per_ch + (1 if idx < remainder else 0)
        ch.target_words = ch_words
        ch.blocks = math.ceil(ch_words / block_size)

        n_sec = max(len(ch.sections), 1)
        sec_base = ch_words // n_sec
        sec_rem = ch_words % n_sec

        for j, sec in enumerate(ch.sections):
            sec_words = sec_base + (1 if j < sec_rem else 0)
            sec.target_words = sec_words
            sec.blocks = math.ceil(sec_words / block_size)

    return chapters


# ============ DOCX builder ============

def _set_section_margins(doc: Document, top_cm=2.0, bottom_cm=2.0, left_cm=2.0, right_cm=2.0):
    for section in doc.sections:
        section.top_margin = CmSafe(top_cm)
        section.bottom_margin = CmSafe(bottom_cm)
        section.left_margin = CmSafe(left_cm)
        section.right_margin = CmSafe(right_cm)

def CmSafe(x: float):
    # Helper to avoid import issues if user wants to tweak margins
    return Inches(x / 2.54)

def add_styled_heading(p, bold=True, size=16, align_center=False):
    run = p.add_run()
    run.bold = bold
    run.font.size = Pt(size)
    if align_center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return run

def create_docx(plan: BookPlan) -> bytes:
    doc = Document()

    # Margins
    _set_section_margins(doc, 2.0, 2.0, 2.0, 2.0)

    # Title page
    p_title = doc.add_paragraph()
    run_t = add_styled_heading(p_title, bold=True, size=24, align_center=True)
    run_t.text = plan.title

    if plan.subtitle.strip():
        p_sub = doc.add_paragraph()
        run_s = add_styled_heading(p_sub, bold=False, size=14, align_center=True)
        run_s.text = plan.subtitle

    # Spacer
    doc.add_paragraph()

    # Metadata paragraph
    meta = doc.add_paragraph()
    meta_run = meta.add_run(f"Totale parole: {plan.total_words} — Dimensione blocchi: {plan.block_size}")
    meta_run.font.size = Pt(10)

    # Content
    for ch in plan.chapters:
        # Chapter heading
        p = doc.add_paragraph()
        run = add_styled_heading(p, bold=True, size=18, align_center=False)
        run.text = ch.title

        # Chapter meta
        cm = doc.add_paragraph()
        cm_run = cm.add_run(f"Parole assegnate: {ch.target_words}  Blocchi: {ch.blocks}")
        cm_run.font.size = Pt(10)

        for sec in ch.sections:
            sp = doc.add_paragraph()
            srun = add_styled_heading(sp, bold=False, size=14, align_center=False)
            srun.text = sec.title

            sm = doc.add_paragraph()
            sm_run = sm.add_run(f"Parole assegnate: {sec.target_words}  Blocchi: {sec.blocks}")
            sm_run.font.size = Pt(10)

            # Insert block placeholders
            for b in range(1, sec.blocks + 1):
                bp = doc.add_paragraph()
                br = bp.add_run(f"[{sec.title}] Blocco {b} di {sec.blocks} — target {plan.block_size} parole")
                br.italic = True
                br.font.size = Pt(10)
                # Add a spacer paragraph for content writing
                doc.add_paragraph()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============ PDF builder ============

def create_pdf(plan: BookPlan, pagesize: str = "A4") -> bytes:
    if str(pagesize).lower() == "letter":
        psize = letter
    else:
        psize = A4

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=psize)
    width, height = psize

    left = 2.0 * cm
    top = height - 2.0 * cm
    line_h = 14

    def write_line(text: str, y: int, bold: bool = False, size: int = 12):
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        c.drawString(left, y, text)

    # Title page
    y = top - 40
    write_line(plan.title, y, bold=True, size=20)
    y -= 24
    if plan.subtitle.strip():
        write_line(plan.subtitle, y, bold=False, size=12)
        y -= 12

    y -= 20
    write_line(f"Totale parole: {plan.total_words}  Dimensione blocchi: {plan.block_size}", y, size=10)
    c.showPage()

    # Content pages
    y = top
    for ch in plan.chapters:
        # Chapter
        if y < 80:
            c.showPage()
            y = top
        write_line(ch.title, y, bold=True, size=16)
        y -= line_h
        write_line(f"Parole assegnate: {ch.target_words}  Blocchi: {ch.blocks}", y, size=9)
        y -= line_h

        for sec in ch.sections:
            if y < 80:
                c.showPage()
                y = top
            write_line(sec.title, y, bold=False, size=12)
            y -= line_h
            write_line(f"Parole assegnate: {sec.target_words}  Blocchi: {sec.blocks}", y, size=9)
            y -= line_h

            for b in range(1, sec.blocks + 1):
                if y < 80:
                    c.showPage()
                    y = top
                write_line(f"[{sec.title}] Blocco {b} di {sec.blocks}  target {plan.block_size} parole", y, size=9)
                y -= line_h
                # Space for writing
                y -= line_h * 2

    c.showPage()
    c.save()
    return buf.getvalue()


# ============ Streamlit App ============

st.set_page_config(page_title="Book Agent — Splitter 500", page_icon="📚", layout="wide")

st.title("Book Agent — Generatore blocchi 500 parole")
st.caption("Upload TOC in DOCX o PDF, allocazione parole e split automatico senza conferme. Esporta DOCX e PDF.")

with st.sidebar:
    st.header("Parametri")
    input_title = st.text_input("Titolo libro", value="")
    input_subtitle = st.text_input("Sottotitolo", value="")

    total_words = st.number_input("Totale parole target", min_value=1, step=500, value=20000)
    block_size = st.number_input("Dimensione blocchi", min_value=1, step=50, value=500)
    pdf_pagesize = st.selectbox("Formato PDF", options=["A4", "Letter"], index=0)

    st.markdown("---")
    st.caption("Suggerimento: i blocchi sono calcolati con arrotondamento per eccesso per coprire l’obiettivo di parole.")

st.subheader("Carica il tuo TOC")
uploaded = st.file_uploader("Seleziona un file DOCX o PDF contenente i titoli di capitoli e sezioni", type=["docx", "pdf"])

placeholder_report = st.empty()
colA, colB = st.columns(2)

if uploaded is not None:
    fname = uploaded.name.lower()
    data = uploaded.read()

    try:
        if fname.endswith(".docx"):
            chapters = extract_toc_from_docx(data)
        else:
            chapters = extract_toc_from_pdf(data)

        if not chapters:
            st.error("Nessun capitolo riconosciuto. Verifica che il file contenga un indice leggibile oppure usa stili Heading 1 e Heading 2 in DOCX.")
        else:
            chapters = allocate_words(chapters, total_words, block_size)
            plan = BookPlan(
                title=input_title.strip() or "Titolo",
                subtitle=input_subtitle.strip(),
                total_words=total_words,
                block_size=block_size,
                chapters=chapters,
            )

            # Report sintetico a schermo
            with placeholder_report.container():
                st.success("TOC estratto e allocato correttamente.")
                st.write(f"Capitoli: {len(plan.chapters)}  |  Totale parole: {plan.total_words}  |  Blocchi da {plan.block_size}")

                for idx, ch in enumerate(plan.chapters, start=1):
                    with st.expander(f"Capitolo {idx}: {ch.title}  —  parole {ch.target_words}  —  blocchi {ch.blocks}"):
                        for j, sec in enumerate(ch.sections, start=1):
                            st.write(f"Sezione {j}: {sec.title}  |  parole {sec.target_words}  |  blocchi {sec.blocks}")

            # Generazione file senza chiedere conferme
            docx_bytes = create_docx(plan)
            pdf_bytes = create_pdf(plan, pagesize=pdf_pagesize)

            with colA:
                st.download_button(
                    label="Scarica DOCX",
                    data=docx_bytes,
                    file_name="book_plan_blocks.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            with colB:
                st.download_button(
                    label="Scarica PDF",
                    data=pdf_bytes,
                    file_name="book_plan_blocks.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

    except Exception as e:
        st.error(f"Errore durante l’elaborazione: {e}")

else:
    st.info("Carica un file DOCX o PDF con l’indice del libro. Per DOCX usa Heading 1 per i capitoli e Heading 2 per le sezioni.")
