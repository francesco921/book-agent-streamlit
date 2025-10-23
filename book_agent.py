import io
import math
import re
import os
from dataclasses import dataclass, field
from typing import List, Optional

import streamlit as st
from docx import Document
from PyPDF2 import PdfReader

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.units import cm

# ============= DATA MODELS ============= #

@dataclass
class Section:
    title: str
    target_words: int = 0
    blocks: int = 0
    texts: List[str] = field(default_factory=list)

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

# ============= HELPERS ============= #

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

def allocate_words(chapters: List[Chapter], total_words: int, block_size: int) -> List[Chapter]:
    total_words = max(total_words, 1)
    block_size = max(block_size, 1)
    n_ch = max(len(chapters), 1)
    base = total_words // n_ch
    rem = total_words % n_ch

    for i, ch in enumerate(chapters):
        ch.target_words = base + (1 if i < rem else 0)
        ch.blocks = math.ceil(ch.target_words / block_size)
        n_sec = max(len(ch.sections), 1)
        base_s = ch.target_words // n_sec
        rem_s = ch.target_words % n_sec
        for j, sec in enumerate(ch.sections):
            sec.target_words = base_s + (1 if j < rem_s else 0)
            sec.blocks = math.ceil(sec.target_words / block_size)
    return chapters

# ============= STREAMLIT UI ============= #

st.set_page_config(page_title="Book Agent - Generatore", page_icon="📘", layout="wide")
st.title("📘 Book Agent — Generatore Blocco 500 Parole")

st.markdown("Flusso in due fasi: caricamento e allocazione prima, generazione dopo.")

# --- Sidebar Input ---
with st.sidebar:
    st.header("Dati libro")
    title = st.text_input("Titolo")
    subtitle = st.text_input("Sottotitolo")
    author = st.text_input("Autore")
    total_words = st.number_input("Totale parole", min_value=1000, step=500, value=20000)
    block_size = 500  # fisso
    st.info(f"I blocchi sono di default da {block_size} parole.")

# --- Upload TOC ---
uploaded = st.file_uploader("Carica TOC (DOCX o PDF)", type=["docx", "pdf"])

# Stato intermedio
if "chapters" not in st.session_state:
    st.session_state["chapters"] = None
if "allocation_done" not in st.session_state:
    st.session_state["allocation_done"] = False

# Step 1: Genera allocazione
if uploaded and title and subtitle and author and total_words:
    if st.button("GENERA ALLOCAZIONE", type="primary", use_container_width=True):
        data = uploaded.read()
        fname = uploaded.name.lower()
        chapters = extract_toc_from_docx(data) if fname.endswith(".docx") else extract_toc_from_pdf(data)
        if not chapters:
            st.error("Nessun capitolo riconosciuto. Usa Heading 1 e Heading 2 nel DOCX.")
        else:
            chapters = allocate_words(chapters, total_words, block_size)
            st.session_state["chapters"] = chapters
            st.session_state["allocation_done"] = True

# Step 2: Visualizza allocazione
if st.session_state["allocation_done"] and st.session_state["chapters"]:
    chapters = st.session_state["chapters"]
    st.subheader("📖 Struttura allocata")
    for i, ch in enumerate(chapters, start=1):
        with st.expander(f"Capitolo {i}: {ch.title} — parole {ch.target_words} — blocchi {ch.blocks}"):
            for j, sec in enumerate(ch.sections, start=1):
                st.write(f"• {sec.title} — {sec.target_words} parole — {sec.blocks} blocchi")

    if st.button("CONFERMA E GENERA", type="primary", use_container_width=True):
        st.info("⚙️ Generazione contenuti in corso... (questa parte verrà implementata nella prossima versione)")
        st.success("Allocazione confermata. Generazione pronta.")
else:
    st.warning("Carica TOC e compila tutti i campi prima di generare l’allocazione.")
