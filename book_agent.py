# ==========================================
# 🧠 COSA FA QUESTO PROGRAMMA
# ------------------------------------------
# Ti aiuta a costruire un libro partendo dal suo indice (TOC).
# 1. Carichi un file con capitoli e sezioni (DOCX o PDF)
# 2. Inserisci titolo, sottotitolo, autore e parole totali
# 3. Premi “GENERA ALLOCAZIONE” e vedi come vengono divisi i blocchi
# 4. Premi “CONFERMA E GENERA” e lui scrive davvero i testi (con OpenAI)
# ==========================================

# ==========================================
# 📦 IMPORTO LE COSE CHE SERVONO
# ==========================================
import io
import math
import re
import os
from dataclasses import dataclass, field
from typing import List

import streamlit as st
from docx import Document
from PyPDF2 import PdfReader

# provo a importare OpenAI
try:
    from openai import OpenAI
    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
    OPENAI_OK = True
except Exception:
    OPENAI_OK = False


# ==========================================
# 🧱 COSTRUISCO LE “SCATOLE” PER I DATI
# ==========================================

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


# ==========================================
# 🧹 COME PULISCO E RICONOSCO I TITOLI
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
# ✏️ COME DIVIDO LE PAROLE E I BLOCCHI
# ==========================================
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


# ==========================================
# 🧠 COME GENERO IL TESTO VERO (con OpenAI)
# ==========================================
def generate_text_block(plan: BookPlan, ch: Chapter, sec: Section, block_index: int) -> str:
    """Chiedo all’AI di scrivere un blocco di testo vero"""
    if not OPENAI_OK:
        return f"[SEGNAPOSTO] {sec.title} — Blocco {block_index+1} (500 parole)."

    prompt = (
        f"Scrivi circa 500 parole per un libro intitolato '{plan.title}' (autore: {plan.author}). "
        f"Il capitolo è '{ch.title}', la sezione è '{sec.title}'. "
        f"Non fare introduzioni generiche, scrivi subito contenuto concreto e fluido."
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
        return res.choices[0].message.content.strip()
    except Exception as e:
        return f"[ERRORE AI] {e}"


def generate_all(plan: BookPlan):
    """Crea tutti i blocchi del libro"""
    total_blocks = sum(sec.blocks for ch in plan.chapters for sec in ch.sections)
    done = 0
    bar = st.progress(0, text="Sto scrivendo i blocchi...")
    for ch in plan.chapters:
        for sec in ch.sections:
            sec.texts = []
            for b in range(sec.blocks):
                text = generate_text_block(plan, ch, sec, b)
                sec.texts.append(text)
                done += 1
                bar.progress(done / total_blocks, text=f"Blocchi completati: {done}/{total_blocks}")
    bar.empty()
    st.success("✅ Tutti i blocchi generati con successo!")


# ==========================================
# 🖥️ LA PARTE GRAFICA (QUELLA CHE USI)
# ==========================================
st.set_page_config(page_title="Book Agent - Generatore", page_icon="📘", layout="wide")
st.title("📘 Book Agent — Generatore blocchi da 500 parole")
st.caption("Carica il TOC, genera l’allocazione e poi i testi reali.")

with st.sidebar:
    st.header("📋 Dati del libro")
    title = st.text_input("Titolo")
    subtitle = st.text_input("Sottotitolo")
    author = st.text_input("Autore")
    total_words = st.number_input("Totale parole", min_value=1000, step=500, value=20000)
    block_size = 500
    st.info(f"I blocchi sono sempre da {block_size} parole.")

uploaded = st.file_uploader("Carica il tuo TOC (DOCX o PDF)", type=["docx", "pdf"])

# Creo “memoria” temporanea
if "chapters" not in st.session_state:
    st.session_state["chapters"] = None
if "allocation_done" not in st.session_state:
    st.session_state["allocation_done"] = False

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
            title=title, subtitle=subtitle, author=author,
            total_words=total_words, block_size=block_size, chapters=chapters
        )
        generate_all(plan)

else:
    st.warning("Carica il TOC e inserisci tutti i dati prima di cliccare 'GENERA ALLOCAZIONE'.")
