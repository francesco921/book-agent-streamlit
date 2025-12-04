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
from docx.oxml.ns import qn
from PyPDF2 import PdfReader

# PDF (per impaginare il libro)
from reportlab.lib.pagesizes import A4, letter  # (non usati in UI, ma safe to keep)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY

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

# Formati pagina “da libro” (limitati a 6x9 e 8.5x11 come richiesto)
PAGE_SIZES = {
    "6x9": (6 * 72, 9 * 72),           # pollici → punti tipografici
    "8.5x11": (8.5 * 72, 11 * 72),
}

# Scelte font (DOCX + PDF verranno armonizzate più avanti)
FONT_CHOICES = ["Times New Roman", "Roboto", "Comfortaa"]

# Tono di voce (UI inglese usa lista dedicata)
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
    tone: str = "Conversational"         # UI inglese
    brief: str = ""                      # descrizione breve che guida lo stile
    pdf_page: str = "6x9"                # formato libro
    font_name: str = "Times New Roman"   # font preferito

# ==========================================
# 🖥️ IMPOSTAZIONI BASE DELLA PAGINA (UI in inglese)
# ==========================================
st.set_page_config(page_title="Book Agent — Book Generator", page_icon="📘", layout="wide")
st.title("📘 Book Agent — Book Generator")
st.caption("Upload TOC → review/approve → generate → download DOCX/PDF")

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
# - Upload DOCX / PDF / TXT (opzionale)
# - Oppure incolla direttamente il TOC
# - Refine TOC con AI (usa chiave OpenAI se presente)
# - Conferma TOC per passare allo Step 2
# ==========================================

st.subheader("📄 Step 1 — Upload or paste your TOC")

# Applica eventuale TOC "pending" PRIMA di creare il widget
if "toc_text_editable_pending" in st.session_state:
    st.session_state["toc_text_editable"] = st.session_state["toc_text_editable_pending"]
    del st.session_state["toc_text_editable_pending"]

# Inizializza stato
if "toc_text_editable" not in st.session_state:
    st.session_state["toc_text_editable"] = ""
if "detected_lang" not in st.session_state:
    st.session_state["detected_lang"] = "auto"
if "_last_uploaded_name" not in st.session_state:
    st.session_state["_last_uploaded_name"] = None

uploaded_file = st.file_uploader(
    "Upload a file that contains the table of contents (DOCX, PDF or TXT) (optional)",
    type=["docx", "pdf", "txt"]
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

def extract_toc_from_txt(file):
    """Extract text from TXT file, line by line."""
    content = file.read().decode("utf-8", errors="ignore")
    lines = [ln for ln in content.splitlines() if ln.strip()]
    return "\n".join(lines)

# Localized chapter keyword to enforce in AI refinement
def _chapter_word(lang_code: str) -> str:
    mapping = {"it": "Capitolo", "en": "Chapter", "es": "Capítulo", "fr": "Chapitre"}
    return mapping.get(lang_code, "Chapter")

# Se viene caricato un file, leggo il TOC ma NON sovrascrivo sempre:
# solo se è un file nuovo (nome diverso)
if uploaded_file:
    filename = uploaded_file.name.lower()

    with st.spinner("Reading the TOC..."):
        if filename.endswith(".docx"):
            toc_text = extract_toc_from_docx(uploaded_file)
        elif filename.endswith(".pdf"):
            toc_text = extract_toc_from_pdf(uploaded_file)
        else:
            toc_text = extract_toc_from_txt(uploaded_file)

    # Language detection solo se c'è testo
    detected = "auto"
    if toc_text.strip() and HAS_LANGID:
        try:
            detected = langid.classify(toc_text[:500])[0]
        except Exception:
            detected = "auto"
    st.session_state["detected_lang"] = detected

    # Se è un nuovo file, popolo la text_area con il contenuto estratto
    if st.session_state["_last_uploaded_name"] != uploaded_file.name:
        st.session_state["toc_text_editable"] = toc_text
        st.session_state["_last_uploaded_name"] = uploaded_file.name

    st.success(f"Detected language: **{detected.upper()}**")
else:
    # Nessun file: mostro eventuale lingua rilevata in passato
    detected = st.session_state.get("detected_lang", "auto")
    if detected != "auto":
        st.info(f"Detected language (from previous run): **{detected.upper()}**")

# Text area SEMPRE visibile per incolla/manual edit
st.text_area(
    "Captured / pasted TOC:",
    key="toc_text_editable",
    height=300,
    help="You can paste your TOC directly here, or edit what has been extracted from the uploaded file."
)

# Bottoni azione
col1, col2 = st.columns([1, 1])
with col1:
    confirm_toc = st.button("✅ Confirm this TOC")
with col2:
    refine_toc = st.button("🧠 Refine TOC with AI")

# AI refinement
if refine_toc:
    if not OPENAI_OK:
        st.error("OpenAI API key not configured. Cannot refine TOC with AI.")
    elif not st.session_state["toc_text_editable"].strip():
        st.error("TOC is empty. Paste or upload a TOC before asking AI refinement.")
    else:
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

        # Applico via PENDING + rerun per evitare conflitti con widget
        st.session_state["toc_text_editable_pending"] = new_toc
        st.info("AI proposal applied. Reloading the editor...")
        st.rerun()

# Conferma TOC
if confirm_toc:
    if not st.session_state["toc_text_editable"].strip():
        st.error("TOC is empty. Please upload or paste a valid TOC before confirming.")
    else:
        st.session_state.confirmed_toc_text = st.session_state.toc_text_editable
        st.success("✅ TOC confirmed. You can proceed to allocation.")
# ==========================================
# 🧮 BLOCK 3 — WORD ALLOCATION & 500-WORD BLOCKING
# ------------------------------------------
# Dal TOC confermato:
# - Riconosce TOC / PARTE / CAPITOLO
# - Identifica i sottocapitoli (foglie) e ne pulisce il titolo
# - Legge [parole] o (blocchi) alla fine riga
# - Se mancano allocazioni, chiede parole per OGNI sottocapitolo mancante
# - Calcola blocchi da max 500 parole
# - Riscrive il TOC normalizzato nell'area di testo
# ==========================================

# ---------- HELPERS ----------

def _strip_leading_markers(s: str) -> str:
    """
    Rimuove:
    - bullet iniziali
    - numerazioni tipo 1., 1.1, 1) ecc.
    - prefissi testuali tipo 'SOTTOCAPITOLO', 'Subchapter', 'Section' + eventuale numero.
    """
    s = s.strip()
    # bullet
    s = re.sub(r"^[\-\*\u2022]+\s*", "", s)
    # numeri tipo 1. 1.1 1.1.1 etc
    s = re.sub(r"^\d+(?:\.\d+){0,3}[\.\)\-:]?\s*", "", s)
    # forme tipo ".1 "
    s = re.sub(r"^\.\d+\s*", "", s)
    # prefissi testuali di sottocapitolo / sezione
    s = re.sub(
        r"^(sottocapitolo|sottcapitolo|sotto capitolo|subchapter|sub-chapter|section)\s*\d*\s*[:\-\.\)]*\s*",
        "",
        s,
        flags=re.IGNORECASE,
    )
    return s.strip()

def _is_toc_label(ln: str) -> bool:
    """Riconosce TOC / INDEX / INDICE ecc."""
    s = ln.strip().lower()
    return (
        s in {"toc", "t.o.c.", "index", "indice", "table of contents"} or
        "table of contents" in s
    )

def _is_part_label(ln: str) -> bool:
    """Riconosce PART / PARTE (multilingua basilare)."""
    s = ln.strip().lower()
    return bool(re.match(r"^(part|parte|parte|partie)\b", s))

def _is_chapter_keyword_line(ln: str) -> bool:
    """
    Riconosce CAPITOLO / CHAPTER solo come parola intera
    all'inizio della riga (non SOTTOCAPITOLO).
    """
    s = ln.strip().lower()
    if re.match(r"^(chapter|capitolo)\b", s):
        return True
    if re.match(r"^(chapter|capitolo)\s+\d+\b", s):
        return True
    return False

def _parse_allocation_from_title(raw: str):
    """
    Estrae titolo + parole/blocchi da fine riga.

    Formati supportati (alla fine):
    - '... [800]'
    - '... [800 words]'
    - '... [800 parole]'
    - '... (2 blocchi)'
    - '... (3 blocks)'

    Regola:
    - se contiene 'block|blocks|blocchi|blocco' => blocchi
    - altrimenti il numero è parole
    """
    text = raw.rstrip()
    words = 0
    blocks = 0

    m = re.search(r'[\[\(]([^\]\)]*)[\]\)]\s*$', text)
    if m:
        inner = m.group(1).strip()
        text = text[:m.start()].rstrip()
        num_match = re.search(r"(\d+)", inner)
        if num_match:
            n = int(num_match.group(1))
            if re.search(r"(block|blocks|blocchi|blocco)", inner, re.IGNORECASE):
                blocks = n
            else:
                words = n

    return text.strip(), words, blocks

def _looks_like_chapter_without_keyword(ln: str) -> bool:
    """
    Euristica per capitoli senza parola CAPITOLO/CHAPTER.
    NON deve riconoscere sottocapitoli (contengono 'sott', 'sub', 'section').
    """
    s = ln.strip().lower()

    # se contiene indicatori di sottosezione → non è capitolo
    if "sott" in s or "sub" in s or "section" in s:
        return False

    if len(ln.split()) > 12:
        return False
    if not ln[:1].isalpha():
        return False
    if ln[:1].islower():
        return False
    if re.search(r"[.:;]$", ln):
        return False
    return True

def parse_confirmed_toc(toc_text: str):
    """
    Parsing del TOC confermato.

    Ritorna:
    - chapters: List[Chapter]
    - chapter_parts: List[Optional[str]]  (titolo PARTE associato a ogni capitolo)
    """
    lines = [ln.rstrip() for ln in toc_text.splitlines() if ln.strip()]

    chapters: List[Chapter] = []
    chapter_parts: List[Optional[str]] = []

    current_chapter: Optional[Chapter] = None
    current_part: Optional[str] = None
    chap_idx = 0

    for raw in lines:
        ln = raw.strip()
        if not ln:
            continue

        # 1) TOC
        if _is_toc_label(ln):
            continue

        # 2) PART / PARTE: solo struttura, ma la teniamo per i capitoli successivi
        if _is_part_label(ln):
            current_part = ln.strip()
            continue

        # 3) CAPITOLO esplicito
        if _is_chapter_keyword_line(ln):
            chap_idx += 1
            title_clean, _, _ = _parse_allocation_from_title(ln)
            title_clean = _strip_leading_markers(title_clean)
            current_chapter = Chapter(title=title_clean)
            chapters.append(current_chapter)
            chapter_parts.append(current_part)
            continue

        # 4) CAPITOLO euristico (senza keyword)
        if current_chapter is None:
            if _looks_like_chapter_without_keyword(ln):
                chap_idx += 1
                title_clean, _, _ = _parse_allocation_from_title(ln)
                title_clean = _strip_leading_markers(title_clean)
                current_chapter = Chapter(title=title_clean)
                chapters.append(current_chapter)
                chapter_parts.append(current_part)
                continue
            else:
                # fallback: crea capitolo generico
                chap_idx += 1
                current_chapter = Chapter(title=f"Chapter {chap_idx}")
                chapters.append(current_chapter)
                chapter_parts.append(current_part)

        # 5) Tutto il resto = sottocapitolo foglia
        title_clean, words, blocks = _parse_allocation_from_title(ln)
        title_clean = _strip_leading_markers(title_clean)

        if not title_clean:
            continue

        sec = Section(title=title_clean, target_words=words, blocks=blocks)
        current_chapter.sections.append(sec)

    # safety: almeno 1 sezione per capitolo
    for ch in chapters:
        if not ch.sections:
            ch.sections.append(Section(title="Section 1"))

    return chapters, chapter_parts

def finalize_allocation_from_toc(chapters: List[Chapter]):
    """
    - Normalizza parole/blocchi per ogni sezione:
      * se blocks e non words: words = blocks * MAX_SUBGEN_WORDS
      * se words e non blocks: blocks = ceil(words / MAX_SUBGEN_WORDS)
    - Ritorna:
      * chapters aggiornati
      * total_words
      * lista sezioni ancora senza allocazione (dovrebbero essere 0
        se lo UI ha raccolto tutto correttamente)
    """
    all_secs: List[Section] = [sec for ch in chapters for sec in ch.sections]

    # normalizza dove ho dati
    for sec in all_secs:
        if sec.blocks and not sec.target_words:
            sec.target_words = sec.blocks * MAX_SUBGEN_WORDS
        elif sec.target_words and not sec.blocks:
            sec.blocks = max(1, math.ceil(sec.target_words / MAX_SUBGEN_WORDS))

    # sezioni ancora senza allocazione
    missing = [s for s in all_secs if s.target_words <= 0 or s.blocks <= 0]

    total_words = sum(sec.target_words for sec in all_secs)

    # aggrega per capitolo
    for ch in chapters:
        ch.target_words = sum(s.target_words for s in ch.sections)
        ch.blocks = sum(s.blocks for s in ch.sections)

    return chapters, total_words, missing

def rebuild_toc_from_plan(chapters: List[Chapter], chapter_parts: List[Optional[str]]) -> str:
    """
    Ricostruisce un TOC lineare normalizzato:

    TOC
    PARTE 1 ...
    CAPITOLO ...
    Titolo sottocapitolo [parole]
    """
    lines = []
    lines.append("TOC")

    last_part = None
    for ch, part in zip(chapters, chapter_parts):
        # nuova PARTE
        if part and part != last_part:
            lines.append(part.strip())
            last_part = part

        # titolo capitolo (così com'è stato pulito)
        if ch.title.strip():
            lines.append(ch.title.strip())

        # sottocapitoli con parole
        for sec in ch.sections:
            lines.append(f"{sec.title.strip()} [{sec.target_words}]")

    return "\n".join(lines)

# ---------- UI STEP 2 ----------

st.subheader("🧩 Step 2 — Book data & allocation")

if st.session_state.confirmed_toc_text:

    # 1) Parsing TOC con struttura (PARTE + CAPITOLO + SOTTOCAPITOLO)
    temp_chapters, chapter_parts = parse_confirmed_toc(st.session_state.confirmed_toc_text)

    # individua sezioni senza parole/blocchi
    all_secs = [sec for ch in temp_chapters for sec in ch.sections]
    missing_initial = [s for s in all_secs if s.target_words <= 0 and s.blocks <= 0]
    needs_per_section_input = len(missing_initial) > 0

    # 2) UI: lingua, tono, brief, formato, font
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

    pdf_page = st.selectbox("Page size", ["6x9", "8.5x11"], index=0)
    font_name = st.selectbox("Primary font", FONT_CHOICES, index=0)

    # 3) Form metadati libro + parole per sezioni senza allocazione
    missing_specs = []

    with st.form("book_info_form"):
        title = st.text_input("Book title")
        subtitle = st.text_input("Subtitle")
        author = st.text_input("Author")

        if needs_per_section_input:
            st.warning(
                f"There are {len(missing_initial)} sections without allocation "
                f"(no words or blocks). Please specify words for each."
            )
            for idx, sec in enumerate(missing_initial):
                v = st.number_input(
                    f"Words for section: '{sec.title}'",
                    min_value=MIN_SECTION_WORDS_USEFUL,
                    step=100,
                    value=MIN_SECTION_WORDS_USEFUL,
                    key=f"missing_words_{idx}",
                )
                missing_specs.append((sec, v))

        submitted_meta = st.form_submit_button("Save book data & compute allocation")

    if submitted_meta:
        # assegna le parole inserite per ogni sezione mancante
        for sec, v in missing_specs:
            sec.target_words = int(v)
            sec.blocks = max(1, math.ceil(sec.target_words / MAX_SUBGEN_WORDS))

        # 4) Finalizza allocazione
        chapters_alloc, total_words, missing_after = finalize_allocation_from_toc(temp_chapters)

        if missing_after:
            st.error("Some sections are still missing allocation. Check your TOC or the per-section words.")
            st.stop()

        # 5) Costruisce il piano di libro
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

               # 6) Prepara il TOC normalizzato per la text_area (usando PENDING)
        new_toc = rebuild_toc_from_plan(chapters_alloc, chapter_parts)
        st.session_state["toc_text_editable_pending"] = new_toc
        st.session_state["confirmed_toc_text"] = new_toc

        st.success(f"✅ Allocation ready. Total words: ~{total_words}. TOC has been normalized.")

        # Forzo un rerun così Step 1 applica il valore pending
        st.rerun()



        # 7) Preview allocazione
        for i, ch in enumerate(chapters_alloc, start=1):
            with st.expander(f"Chapter {i}: {ch.title} — {ch.target_words} words — {ch.blocks} total blocks"):
                for j, sec in enumerate(ch.sections, start=1):
                    st.write(
                        f"• Section {j}: {sec.title} — {sec.target_words} words — "
                        f"{sec.blocks} blocks (≤{MAX_SUBGEN_WORDS} words each)"
                    )

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
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = cm_to_Cm = Cm(2.54)
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(2.54)
    else:
        sec.page_width, sec.page_height = Inches(6), Inches(9)

    # Title page
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(plan.title); r.bold = True; r.font.size = Pt(26)
    if plan.subtitle.strip():
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(plan.subtitle); r.font.size = Pt(16)
    for _ in range(12): doc.add_paragraph("")
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(plan.author); r.font.size = Pt(12)
    doc.add_page_break()

    # Copyright/Disclaimer
    if include_copyright:
        p = doc.add_paragraph("Copyright & Disclaimer"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        para = doc.add_paragraph(
            f"© {plan.author}. All rights reserved.\n\n"
            "No part of this publication may be reproduced, distributed, or transmitted in any form or by any means, "
            "including photocopying, recording, or other electronic or mechanical methods, without the prior written "
            "permission of the publisher, except in the case of brief quotations embodied in critical reviews.\n\n"
            "Disclaimer: The information in this book is provided for educational purposes only and does not constitute "
            "professional advice. Always consult a qualified professional for your specific situation."
        ); para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        doc.add_page_break()

    # ToC
    if include_toc:
        _add_docx_toc(doc); doc.add_page_break()

    # Font base
    style = doc.styles["Normal"]; style.font.name = plan.font_name
    try: style._element.rPr.rFonts.set(qn("w:eastAsia"), plan.font_name)
    except Exception: pass

    # Content — Heading 1/2 reali (necessari per il ToC)
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
    H1 = ParagraphStyle("H1", parent=styles["Heading1"], fontName=fnt, alignment=TA_LEFT, spaceBefore=12, spaceAfter=6)
    H2 = ParagraphStyle("H2", parent=styles["Heading2"], fontName=fnt, alignment=TA_LEFT, spaceBefore=6, spaceAfter=4)
    Body = ParagraphStyle("Body", parent=styles["BodyText"], fontName=fnt, alignment=TA_JUSTIFY, leading=14)
    TitleC = ParagraphStyle("TitleC", parent=styles["Title"], fontName=fnt, alignment=TA_CENTER, spaceAfter=12)
    SubC = ParagraphStyle("SubC", parent=styles["BodyText"], fontName=fnt, alignment=TA_CENTER, spaceAfter=24)

    story = []
    # Title page
    story += [Spacer(1, 40), Paragraph(plan.title, TitleC)]
    if plan.subtitle.strip(): story.append(Paragraph(plan.subtitle, SubC))
    story += [Spacer(1, pagesize[1]*0.55), Paragraph(plan.author, SubC), PageBreak()]

    # Copyright/Disclaimer
    if include_copyright:
        story += [
            Paragraph("Copyright & Disclaimer", H2),
            Paragraph(
                f"© {plan.author}. All rights reserved.<br/><br/>"
                "No part of this publication may be reproduced, distributed, or transmitted in any form or by any means, "
                "including photocopying, recording, or other electronic or mechanical methods, without the prior written "
                "permission of the publisher, except in the case of brief quotations embodied in critical reviews.<br/><br/>"
                "Disclaimer: The information in this book is provided for educational purposes only and does not constitute "
                "professional advice. Always consult a qualified professional for your specific situation.", Body
            ),
            PageBreak()
        ]

    # ToC
    if include_toc:
        toc = TableOfContents()
        toc.levelStyles = [
            ParagraphStyle(fontName=fnt, name="TOC1", leftIndent=20, firstLineIndent=-10, spaceBefore=6, leading=12),
            ParagraphStyle(fontName=fnt, name="TOC2", leftIndent=36, firstLineIndent=-10, spaceBefore=4, leading=12),
        ]
        story += [Paragraph("Table of Contents", H1), Spacer(1, 12), toc, PageBreak()]

    # Content
    for ch in plan.chapters:
        story.append(Paragraph(ch.title, H1))
        for sec in ch.sections:
            story.append(Paragraph(sec.title, H2))
            for text in sec.texts:
                story.append(Paragraph(text, Body))
                story.append(Spacer(1, 8))
        story.append(PageBreak())

    doc.build(story); return buf.getvalue()

# ----- UI: generate + preview + downloads + options -----
if st.session_state.allocation_done and st.session_state.generated_plan:
    plan: BookPlan = st.session_state.generated_plan

    c1, c2 = st.columns(2)
    with c1:
        opt_toc = st.checkbox("Include Table of Contents", value=True)
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
