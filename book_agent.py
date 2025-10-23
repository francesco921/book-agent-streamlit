# ==========================================
# 🧠 COSA FA QUESTO PROGRAMMA
# ------------------------------------------
# Ti aiuta a costruire un libro partendo dal suo indice (TOC).
# 1) Carichi un file (DOCX o PDF) con capitoli e sezioni
# 2) Scrivi titolo, sottotitolo, autore e le parole totali
# 3) Premi “GENERA ALLOCAZIONE” e vedi i blocchi da 500 parole
# 4) Premi “CONFERMA E GENERA” e lui scrive i testi vero-VERI
# 5) Alla fine puoi SCARICARE il libro in DOCX e in PDF
# ==========================================

# ==========================================
# 📦 PRENDO GLI ATTREZZI CHE MI SERVONO
# ==========================================
import io
import os
import re
import math
from dataclasses import dataclass, field
from typing import List

import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PyPDF2 import PdfReader

# Per fare il PDF
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.units import cm

# Provo a usare OpenAI (serve OPENAI_API_KEY)
OPENAI_OK = False
try:
    from openai import OpenAI
    if os.getenv("OPENAI_API_KEY"):
        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        OPENAI_OK = True
except Exception:
    OPENAI_OK = False


# ==========================================
# 🧱 LE SCATOLE DOVE METTO I DATI DEL LIBRO
# ==========================================
@dataclass
class Section:
    title: str
    target_words: int = 0
    blocks: int = 0
    texts: List[str] = field(default_factory=list)  # i testi dei blocchi

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


# ==========================================
# 🧹 PULISCO E TROVO I TITOLI NEL FILE
# ==========================================
def normalize_heading(text: str) -> str:
    t = re.sub(r"\s+", " ", text or "").strip()
    t = re.sub(r"\.{2,}\s*\d+$", "", t).strip()
    return t

def guess_is_heading(line: str) -> bool:
    clean = normalize_heading(line)
    if not clean:
        return False
    words = clean.split()
    if len(words) <= 10 and clean.isupper():
        return True
    if re.match(r"^\s*(?:\d+|[IVXLC]+)[\.\)\s]", clean):
        return True
    return False

def extract_toc_from_docx(file_bytes: bytes) -> List[Chapter]:
    doc = Document(io.BytesIO(file_bytes))
    chapters, current = [], None
    for p in doc.paragraphs:
        text = normalize_heading(p.text)
        if not text:
            continue
        style = (getattr(p.style, "name", "") or "").lower()
        if "heading 1" in style or guess_is_heading(text):
            current = Chapter(title=text)
            chapters.append(current)
        elif "heading 2" in style and current:
            current.sections.append(Section(title=text))
    for ch in chapters:
        if not ch.sections:
            ch.sections.append(Section(title="Sezione 1"))
    return chapters

def extract_toc_from_pdf(file_bytes: bytes) -> List[Chapter]:
    reader = PdfReader(io.BytesIO(file_bytes))
    lines = []
    for p in reader.pages:
        txt = p.extract_text() or ""
        for l in txt.splitlines():
            t = normalize_heading(l)
            if t:
                lines.append(t)
    chapters, current = [], None
    for l in lines:
        if guess_is_heading(l):
            if not current:
                current = Chapter(title=l)
                chapters.append(current)
            else:
                current.sections.append(Section(title=l))
    for ch in chapters:
        if not ch.sections:
            ch.sections.append(Section(title="Sezione 1"))
    return chapters


# ==========================================
# ✏️ DIVIDO LE PAROLE E I BLOCCHI
# ==========================================
def allocate_words(chapters: List[Chapter], total_words: int, block_size: int) -> List[Chapter]:
    total_words = max(total_words, 1)
    block_size = max(block_size, 1)
    n_ch = max(len(chapters), 1)
    base = total_words // n_ch
    rem = total_words % n_ch

    for i, ch in enumerate(chapters):
        ch.target_words = base + (1 if i < rem else 0)
        ch.blocks = max(1, math.ceil(ch.target_words / block_size))
        n_sec = max(len(ch.sections), 1)
        base_s = ch.target_words // n_sec
        rem_s = ch.target_words % n_sec
        for j, sec in enumerate(ch.sections):
            sec.target_words = base_s + (1 if j < rem_s else 0)
            sec.blocks = max(1, math.ceil(sec.target_words / block_size))
            sec.texts = [""] * sec.blocks
    return chapters


# ==========================================
# 🧠 SCRIVO DAVVERO I TESTI (CON AI O SEGNAPOSTO)
# ==========================================
def generate_text_block(plan: BookPlan, ch: Chapter, sec: Section, block_index: int) -> str:
    """Scrive ~500 parole. Se manca la chiave OpenAI, usa un segnaposto utile."""
    if not OPENAI_OK:
        return (
            f"[SEGNAPOSTO] Sezione: {sec.title} — Blocco {block_index+1}/{sec.blocks}. "
            f"Scrivi circa {plan.block_size} parole con esempi, spiegazioni pratiche e finale naturale."
        )

    prompt = (
        f"Scrivi circa {plan.block_size} parole per un libro intitolato '{plan.title}' "
        f"(autore: {plan.author}, sottotitolo: {plan.subtitle}). "
        f"Capitolo: '{ch.title}'. Sezione: '{sec.title}'. "
        f"Vai dritto al punto, tono chiaro e concreto, niente meta-commenti, niente titoli ripetuti."
    )
    try:
        res = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Sei uno scrittore esperto. Scrivi testi chiari, interessanti e naturali."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7
        )
        return (res.choices[0].message.content or "").strip()
    except Exception as e:
        return f"[ERRORE AI] {e}"


def generate_all(plan: BookPlan):
    """Crea tutti i blocchi del libro, uno ad uno, con una barra che avanza."""
    total_blocks = sum(sec.blocks for ch in plan.chapters for sec in ch.sections)
    done = 0
    bar = st.progress(0, text="Sto scrivendo i blocchi...")
    for ch in plan.chapters:
        for sec in ch.sections:
            for b in range(sec.blocks):
                text = generate_text_block(plan, ch, sec, b)
                sec.texts[b] = text
                done += 1
                bar.progress(done / total_blocks, text=f"Blocchi completati: {done}/{total_blocks}")
    bar.empty()
    st.success("✅ Tutti i blocchi sono stati scritti!")


# ==========================================
# 🧾 COSTRUISCO IL FILE WORD (DOCX)
# ==========================================
def _add_heading(par, text, size, bold, align_center=False):
    run = par.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if align_center:
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER

def build_docx(plan: BookPlan) -> bytes:
    doc = Document()

    # Copertina
    p = doc.add_paragraph()
    _add_heading(p, plan.title, 24, True, align_center=True)
    if plan.subtitle.strip():
        p = doc.add_paragraph()
        _add_heading(p, plan.subtitle, 14, False, align_center=True)
    p = doc.add_paragraph()
    _add_heading(p, f"Autore: {plan.author}", 10, False, align_center=True)
    p = doc.add_paragraph()
    _add_heading(p, f"Totale parole: {plan.total_words} — Blocchi da {plan.block_size}", 10, False, align_center=True)
    doc.add_page_break()

    # Contenuti
    for ch in plan.chapters:
        p = doc.add_paragraph()
        _add_heading(p, ch.title, 18, True, align_center=False)

        for sec in ch.sections:
            p = doc.add_paragraph()
            _add_heading(p, sec.title, 14, False, align_center=False)

            for idx, text in enumerate(sec.texts, start=1):
                if not (text or "").strip():
                    text = f"[SEGNAPOSTO] {sec.title} — Blocco {idx} — scrivi ~{plan.block_size} parole."
                doc.add_paragraph(text)

        doc.add_page_break()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ==========================================
# 🧾 COSTRUISCO IL FILE PDF (PAGINE VERE)
# ==========================================
def build_pdf(plan: BookPlan, pagesize_name="A4") -> bytes:
    pagesize = A4 if str(pagesize_name).lower() == "a4" else letter
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=pagesize,
        leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm
    )
    styles = getSampleStyleSheet()
    H1, H2, Body = styles["Heading1"], styles["Heading2"], styles["BodyText"]

    flow = []
    # Copertina
    flow.append(Paragraph(plan.title, H1))
    if plan.subtitle.strip():
        flow.append(Paragraph(plan.subtitle, Body))
    flow.append(Paragraph(f"Autore: {plan.author}", Body))
    flow.append(Paragraph(f"Totale parole: {plan.total_words} — Blocchi da {plan.block_size}", Body))
    flow.append(Spacer(1, 16))
    flow.append(PageBreak())

    # Contenuti
    for ch in plan.chapters:
        flow.append(Paragraph(ch.title, H1))
        for sec in ch.sections:
            flow.append(Paragraph(sec.title, H2))
            for idx, text in enumerate(sec.texts, start=1):
                if not (text or "").strip():
                    text = f"[SEGNAPOSTO] {sec.title} — Blocco {idx} — scrivi ~{plan.block_size} parole."
                flow.append(Paragraph(text, Body))
                flow.append(Spacer(1, 8))
        flow.append(PageBreak())

    doc.build(flow)
    return buf.getvalue()


# ==========================================
# 🖥️ LA PAGINA CHE USI (CON I BOTTONI)
# ==========================================
st.set_page_config(page_title="Book Agent - Generatore", page_icon="📘", layout="wide")
st.title("📘 Book Agent — Generatore blocchi da 500 parole")
st.caption("Carica TOC → genera l’allocazione → conferma → scrivi testi → scarica DOCX/PDF.")

with st.sidebar:
    st.header("📋 Dati del libro")
    title = st.text_input("Titolo")
    subtitle = st.text_input("Sottotitolo")
    author = st.text_input("Autore")
    total_words = st.number_input("Totale parole", min_value=5000, step=500, value=20000)
    block_size = 500
    pdf_size = st.selectbox("Formato PDF", options=["A4", "Letter"], index=0)
    st.info(f"I blocchi sono sempre da {block_size} parole.")
    if not OPENAI_OK:
        st.warning("OPENAI_API_KEY non trovata: scriverò segnaposto invece di testo vero.")

uploaded = st.file_uploader("Carica il tuo TOC (DOCX o PDF)", type=["docx", "pdf"])

# memorie temporanee
if "chapters" not in st.session_state:
    st.session_state["chapters"] = None
if "allocation_done" not in st.session_state:
    st.session_state["allocation_done"] = False
if "generated_plan" not in st.session_state:
    st.session_state["generated_plan"] = None
if "docx_bytes" not in st.session_state:
    st.session_state["docx_bytes"] = None
if "pdf_bytes" not in st.session_state:
    st.session_state["pdf_bytes"] = None

# --- STEP 1: genera allocazione ---
if uploaded and title and subtitle and author and total_words:
    if st.button("GENERA ALLOCAZIONE", type="primary", use_container_width=True):
        data = uploaded.read()
        fname = uploaded.name.lower()
        chapters = extract_toc_from_docx(data) if fname.endswith(".docx") else extract_toc_from_pdf(data)
        if not chapters:
            st.error("Non ho trovato capitoli. Usa Heading 1 e Heading 2 nel DOCX.")
        else:
            chapters = allocate_words(chapters, total_words, block_size)
            st.session_state["chapters"] = chapters
            st.session_state["allocation_done"] = True
            st.session_state["generated_plan"] = None
            st.session_state["docx_bytes"] = None
            st.session_state["pdf_bytes"] = None

# --- STEP 2: mostra struttura e bottone “Conferma e Genera” ---
if st.session_state["allocation_done"] and st.session_state["chapters"]:
    chapters = st.session_state["chapters"]
    st.subheader("📚 Struttura del libro divisa in blocchi")
    for i, ch in enumerate(chapters, start=1):
        with st.expander(f"Capitolo {i}: {ch.title} — {ch.target_words} parole — {ch.blocks} blocchi"):
            for j, sec in enumerate(ch.sections, start=1):
                st.write(f"• {sec.title} — {sec.target_words} parole — {sec.blocks} blocchi")

    if st.button("CONFERMA E GENERA", type="primary", use_container_width=True):
        plan = BookPlan(
            title=title,
            subtitle=subtitle,
            author=author,
            total_words=total_words,
            block_size=block_size,
            chapters=chapters
        )
        generate_all(plan)
        st.session_state["generated_plan"] = plan

        # Costruisco subito i file da scaricare
        try:
            st.session_state["docx_bytes"] = build_docx(plan)
            st.session_state["pdf_bytes"] = build_pdf(plan, pagesize_name=pdf_size)
        except Exception as e:
            st.error(f"Errore nel creare i file: {e}")

# --- STEP 3: dopo la generazione compaiono i bottoni di download ---
if st.session_state.get("generated_plan") and st.session_state.get("docx_bytes") and st.session_state.get("pdf_bytes"):
    st.subheader("📥 Scarica il tuo libro")
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="Scarica DOCX",
            data=st.session_state["docx_bytes"],
            file_name="libro_generato.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    with col2:
        st.download_button(
            label="Scarica PDF",
            data=st.session_state["pdf_bytes"],
            file_name="libro_generato.pdf",
            mime="application/pdf",
            use_container_width=True
        )
else:
    st.info("Per scaricare, prima genera i testi e poi torneranno i bottoni qui sotto.")
