import io
import math
import os
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import streamlit as st

# Word/docx
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# PDF
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.units import cm

# PDF parsing
from PyPDF2 import PdfReader

# Optional LLM
OPENAI_OK = False
try:
    from openai import OpenAI
    if os.getenv("OPENAI_API_KEY"):
        OPENAI_OK = True
except Exception:
    OPENAI_OK = False


# ===================== Data models =====================

@dataclass
class Section:
    title: str
    target_words: int = 0
    blocks: int = 0
    texts: List[str] = field(default_factory=list)  # testo per blocco


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
    brief: str = ""


# ===================== TOC extraction =====================

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
    t = re.sub(r"\.{2,}\s*\d+$", "", t).strip()
    return t

def guess_is_heading(line: str) -> bool:
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
            h = _match_first(text, H1_PATTERNS) or text
            current_ch = Chapter(title=h)
            chapters.append(current_ch)
            continue

        if current_ch:
            if "heading 2" in style_name or _match_first(text, H2_PATTERNS):
                h = _match_first(text, H2_PATTERNS) or text
                current_ch.sections.append(Section(title=h))

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
        h1 = _match_first(ln, H1_PATTERNS)
        h2 = _match_first(ln, H2_PATTERNS)

        if h1:
            current_ch = Chapter(title=h1)
            chapters.append(current_ch)
            continue
        if h2 and current_ch:
            current_ch.sections.append(Section(title=h2))
            continue

        if guess_is_heading(ln):
            if not current_ch:
                current_ch = Chapter(title=ln)
                chapters.append(current_ch)
            else:
                current_ch.sections.append(Section(title=ln))

    for ch in chapters:
        if not ch.sections:
            ch.sections.append(Section(title="Sezione 1"))
    return chapters


# ===================== Allocation =====================

def allocate_words(chapters: List[Chapter], total_words: int, block_size: int) -> List[Chapter]:
    total_words = max(total_words, 1)
    block_size = max(block_size, 1)

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
            sec.blocks = max(1, math.ceil(sec_words / block_size))
            sec.texts = [""] * sec.blocks

    return chapters


# ===================== Generation with OpenAI =====================

def gen_block_text_openai(client: "OpenAI", model: str, lang: str, plan: BookPlan,
                          chapter_title: str, section_title: str,
                          block_idx: int, block_count: int, block_size: int) -> str:
    sys = (
        "Sei un autore professionista. Scrivi contenuti scorrevoli, informativi e concreti. "
        "Non ripetere il titolo del libro, né il sottotitolo, né i titoli capitolo/sezione a inizio blocco. "
        "Mantieni tono chiaro e diretto. Evita preamboli inutili. Nessuna lista se non strettamente necessario."
    )
    user = (
        f"Lingua: {lang}\n"
        f"Libro: {plan.title}\n"
        f"Sottotitolo: {plan.subtitle}\n"
        f"Brief sintetico: {plan.brief or 'Nessun brief aggiuntivo'}\n"
        f"Capitolo: {chapter_title}\n"
        f"Sezione: {section_title}\n"
        f"Blocco: {block_idx+1} su {block_count}\n"
        f"Obiettivo parole: ~{block_size}\n\n"
        "Scrivi il testo del blocco in prosa continua, senza inserire il titolo della sezione, "
        "senza meta-commenti, senza istruzioni. Concludi il blocco in modo naturale, "
        "lasciando continuità al successivo se non è l’ultimo."
    )
    # Compatibile con openai>=2.x
    resp = client.chat.completions.create(
        model=model,
        temperature=0.7,
        messages=[
            {"role": "system", "content": sys},
            {"role": "user", "content": user},
        ]
    )
    return resp.choices[0].message.content.strip()


def generate_all_texts(plan: BookPlan, language: str, model_name: str = "gpt-4o-mini"):
    if not OPENAI_OK:
        return False  # fallback ai segnaposto

    client = OpenAI()
    total_blocks = sum(sec.blocks for ch in plan.chapters for sec in ch.sections)
    progress = st.progress(0, text="Generazione contenuti in corso...")
    done = 0

    for ch in plan.chapters:
        for sec in ch.sections:
            for b in range(sec.blocks):
                try:
                    txt = gen_block_text_openai(
                        client=client,
                        model=model_name,
                        lang=language,
                        plan=plan,
                        chapter_title=ch.title,
                        section_title=sec.title,
                        block_idx=b,
                        block_count=sec.blocks,
                        block_size=plan.block_size,
                    )
                except Exception as e:
                    txt = f"[SEGNAPOSTO] Impossibile generare il testo: {e}. Scrivi qui ~{plan.block_size} parole sul tema."

                sec.texts[b] = txt
                done += 1
                progress.progress(done / total_blocks, text=f"Blocchi generati: {done}/{total_blocks}")

    progress.empty()
    return True


# ===================== Builders =====================

def CmSafe(x: float):
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

    # title page
    p_title = doc.add_paragraph()
    run_t = add_styled_heading(p_title, bold=True, size=24, align_center=True)
    run_t.text = plan.title

    if plan.subtitle.strip():
        p_sub = doc.add_paragraph()
        run_s = add_styled_heading(p_sub, bold=False, size=14, align_center=True)
        run_s.text = plan.subtitle

    doc.add_paragraph().add_run(f"Totale parole: {plan.total_words}  Dimensione blocchi: {plan.block_size}").font.size = Pt(10)
    doc.add_page_break()

    # content
    for ch in plan.chapters:
        p = doc.add_paragraph()
        add_styled_heading(p, bold=True, size=18).text = ch.title

        for sec in ch.sections:
            sp = doc.add_paragraph()
            add_styled_heading(sp, bold=False, size=14).text = sec.title

            # concatena i blocchi della sezione
            for idx, text in enumerate(sec.texts):
                # se non abbiamo testo generato, metti segnaposto
                if not text.strip():
                    text = f"[SEGNAPOSTO] {sec.title} — Blocco {idx+1} di {len(sec.texts)}. Scrivi ~{plan.block_size} parole."
                doc.add_paragraph(text)

        doc.add_page_break()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def create_pdf(plan: BookPlan, pagesize: str = "A4") -> bytes:
    psize = letter if str(pagesize).lower() == "letter" else A4
    buf = io.BytesIO()

    doc = SimpleDocTemplate(buf, pagesize=psize,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    H1 = styles["Heading1"]
    H2 = styles["Heading2"]
    Body = styles["BodyText"]

    flow = []
    flow.append(Paragraph(plan.title, H1))
    if plan.subtitle.strip():
        flow.append(Paragraph(plan.subtitle, Body))
    flow.append(Paragraph(f"Totale parole: {plan.total_words}  Dimensione blocchi: {plan.block_size}", Body))
    flow.append(Spacer(1, 12))
    flow.append(PageBreak())

    for ch in plan.chapters:
        flow.append(Paragraph(ch.title, H1))
        for sec in ch.sections:
            flow.append(Paragraph(sec.title, H2))
            for idx, text in enumerate(sec.texts):
                if not text.strip():
                    text = f"[SEGNAPOSTO] {sec.title} — Blocco {idx+1} di {len(sec.texts)}. Scrivi ~{plan.block_size} parole."
                flow.append(Paragraph(text, Body))
                flow.append(Spacer(1, 8))
        flow.append(PageBreak())

    doc.build(flow)
    return buf.getvalue()


# ===================== Streamlit App =====================

st.set_page_config(page_title="Book Agent — Generatore 500", page_icon="📚", layout="wide")
st.title("Book Agent — Generatore blocchi 500 parole")
st.caption("Upload TOC in DOCX o PDF, allocazione parole, generazione testo automatica per blocchi, export DOCX e PDF.")

with st.sidebar:
    st.header("Parametri")
    input_title = st.text_input("Titolo libro", value="")
    input_subtitle = st.text_input("Sottotitolo", value="")
    total_words = st.number_input("Totale parole target", min_value=1, step=500, value=20000)
    block_size = st.number_input("Dimensione blocchi", min_value=1, step=50, value=500)
    pdf_pagesize = st.selectbox("Formato PDF", options=["A4", "Letter"], index=0)
    language = st.selectbox("Lingua di scrittura", ["Italiano", "English"], index=0)
    brief = st.text_area("Brief opzionale del libro", placeholder="Target, stile, promesse, struttura desiderata...")
    auto_generate = st.checkbox("Genera contenuti con OpenAI", value=True if OPENAI_OK else False,
                                help="Richiede OPENAI_API_KEY configurata")

st.subheader("Carica il tuo TOC")
uploaded = st.file_uploader("Seleziona un file DOCX o PDF con capitoli (H1) e sezioni (H2)", type=["docx", "pdf"])

placeholder_report = st.empty()
colA, colB = st.columns(2)

if uploaded is not None:
    fname = uploaded.name.lower()
    data = uploaded.read()

    try:
        chapters = extract_toc_from_docx(data) if fname.endswith(".docx") else extract_toc_from_pdf(data)
        if not chapters:
            st.error("Nessun capitolo riconosciuto. Per DOCX usa Heading 1 per capitoli e Heading 2 per sezioni.")
        else:
            chapters = allocate_words(chapters, total_words, block_size)
            plan = BookPlan(
                title=input_title.strip() or "Titolo",
                subtitle=input_subtitle.strip(),
                total_words=total_words,
                block_size=block_size,
                chapters=chapters,
                brief=brief.strip(),
            )

            # Generazione contenuti
            if auto_generate:
                ok = generate_all_texts(plan, language=("Italiano" if language=="Italiano" else "English"))
                if not ok:
                    st.warning("OPENAI_API_KEY non configurata. Verranno inseriti segnaposto nei blocchi.")
            else:
                st.info("Generazione LLM disattivata. Verranno inseriti segnaposto.")

            # Report
            with placeholder_report.container():
                st.success("TOC estratto e allocato correttamente.")
                st.write(f"Capitoli: {len(plan.chapters)}  |  Totale parole: {plan.total_words}  |  Blocchi da {plan.block_size}")
                for idx, ch in enumerate(plan.chapters, start=1):
                    with st.expander(f"Capitolo {idx}: {ch.title}"):
                        for j, sec in enumerate(ch.sections, start=1):
                            st.write(f"Sezione {j}: {sec.title}  |  blocchi {sec.blocks}")

            # Export
            docx_bytes = create_docx(plan)
            pdf_bytes = create_pdf(plan, pagesize=pdf_pagesize)

            with colA:
                st.download_button(
                    label="Scarica DOCX",
                    data=docx_bytes,
                    file_name="book_generated.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            with colB:
                st.download_button(
                    label="Scarica PDF",
                    data=pdf_bytes,
                    file_name="book_generated.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

    except Exception as e:
        st.error(f"Errore durante l’elaborazione: {e}")

else:
    st.info("Carica un DOCX o PDF con l’indice. Consigliato DOCX con Heading 1 e 2.")
