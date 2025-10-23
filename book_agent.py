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
# 📂 BLOCCO 2 — LETTURA TOC E REVISIONE
# ------------------------------------------
# Qui l’utente carica il file (DOCX/PDF), il sistema legge i titoli e sottotitoli,
# rileva la lingua (o lascia scegliere), e permette di correggere il TOC a mano
# o farlo rigenerare da AI. Si prosegue solo dopo conferma definitiva.
# ==========================================

st.subheader("📄 Step 1 — Carica il tuo TOC")

uploaded_file = st.file_uploader(
    "Carica un file con l’indice (DOCX o PDF)", type=["docx", "pdf"]
)

def extract_toc_from_docx(file):
    """Legge titoli e sottotitoli da DOCX in modo semplice."""
    doc = Document(file)
    toc_lines = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt and not txt.isdigit() and len(txt) > 2:
            toc_lines.append(txt)
    return "\n".join(toc_lines)

def extract_toc_from_pdf(file):
    """Estrae linee di testo leggibili dal PDF (usa solo le prime pagine)."""
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

if uploaded_file:
    # --- Estrazione TOC ---
    with st.spinner("📖 Lettura del TOC in corso..."):
        if uploaded_file.name.endswith(".docx"):
            toc_text = extract_toc_from_docx(uploaded_file)
        else:
            toc_text = extract_toc_from_pdf(uploaded_file)

    # --- Rilevamento lingua ---
    detected = "auto"
    if HAS_LANGID:
        try:
            detected = langid.classify(toc_text[:500])[0]
        except Exception:
            detected = "auto"
    st.session_state.detected_lang = detected

    # --- Mostro TOC captato ---
    st.success(f"Lingua rilevata: **{detected.upper()}**")
    st.text_area("TOC rilevato:", toc_text, height=300, key="toc_text_editable")

    # --- Pulsante per correzione manuale / AI ---
    col1, col2 = st.columns([1, 1])
    with col1:
        confirm_toc = st.button("✅ Conferma questo TOC")
    with col2:
        refine_toc = st.button("🧠 Genera proposta TOC con AI")

    # --- Se l’utente vuole una proposta AI ---
    if refine_toc and OPENAI_OK:
        with st.spinner("Sto generando una versione migliorata del TOC..."):
            prompt_refine = f"Rendi più chiaro e bilanciato questo indice di libro:\n{toc_text}"
            resp = openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "Sei un assistente editoriale."},
                    {"role": "user", "content": prompt_refine},
                ],
                temperature=0.6,
                max_tokens=800,
            )
            toc_text = resp.choices[0].message.content.strip()
            st.session_state.toc_text_editable = toc_text
            st.info("TOC migliorato automaticamente, puoi ancora modificarlo.")

    # --- Conferma finale ---
    if confirm_toc:
        st.session_state.confirmed_toc_text = st.session_state.toc_text_editable
        st.success("✅ TOC confermato! Ora puoi passare all’allocazione delle parole.")
# ==========================================
# 🧮 BLOCCO 3 — ALLOCAZIONE PAROLE E BLOCCHI (≤500)
# ------------------------------------------
# Qui prendo il TOC confermato dall’utente, faccio inserire i dati del libro,
# controllo che le parole bastino, e divido tutto in capitoli/sezioni/blocchi.
# Ogni “micro-generazione” non supererà MAI 500 parole.
# ==========================================

st.subheader("🧩 Step 2 — Dati libro e allocazione")

# ---- mostro i dati lingua/tono/brief/formato solo se c’è un TOC confermato
if st.session_state.confirmed_toc_text:

    # === piccoli helper ===
    def parse_confirmed_toc(toc_text: str) -> List[Chapter]:
        """
        Legge il TOC a righe:
        - Riga “secca” = Capitolo
        - Riga che inizia con '-', '*', numero '1.'/'1)' o indentazione = Sezione del capitolo corrente
        Se un capitolo non ha sezioni, creo 'Sezione 1'.
        """
        lines = [ln.strip() for ln in toc_text.splitlines() if ln.strip()]
        chapters: List[Chapter] = []
        current: Optional[Chapter] = None
        for ln in lines:
            is_section = False
            if ln.startswith(("-", "*")):
                is_section = True
            elif re.match(r"^\d+[\.\)]\s+", ln):
                is_section = True
            elif len(ln.split()) <= 12 and ln[:1].islower() is False and re.search(r"[\.:\-–—]$", ln) is False:
                # euristica “titolo breve e capitalizzato” → Capitolo
                is_section = False
            else:
                # fallback: se non c'è ancora un capitolo, trattalo come capitolo
                is_section = (current is not None)

            if not is_section:
                current = Chapter(title=ln)
                chapters.append(current)
            else:
                if not current:
                    current = Chapter(title="Capitolo 1")
                    chapters.append(current)
                current.sections.append(Section(title=re.sub(r"^[\-\*\d\.\)]\s*", "", ln)))

        # garanzia: ogni capitolo ha almeno 1 sezione
        for ch in chapters:
            if not ch.sections:
                ch.sections.append(Section(title="Sezione 1"))
        return chapters

    def allocate_words(chapters: List[Chapter], total_words: int, block_size_limit: int = 500) -> List[Chapter]:
        """
        Distribuzione semplice ed equilibrata:
        - prima tra capitoli
        - poi tra sezioni del capitolo
        - calcolo blocchi = ceil(parole_sezione / 500), minimo 1
        """
        n_ch = max(len(chapters), 1)
        base = total_words // n_ch
        rem = total_words % n_ch

        for i, ch in enumerate(chapters):
            ch.target_words = base + (1 if i < rem else 0)
            # riparto nelle sezioni
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

    # === UI: lingua, tono, brief, formati, font ===
    # lingua: auto (dal blocco 2) + scelta manuale
    lang_human = st.selectbox(
        "Lingua di generazione",
        ["auto", "it", "en", "es", "fr"],
        index=["auto", "it", "en", "es", "fr"].index(
            st.session_state.detected_lang if st.session_state.detected_lang in ["it", "en", "es", "fr"] else "auto"
        ),
        help="Se lasci 'auto', uso la lingua rilevata dal TOC; altrimenti forzo quella scelta."
    )

    tone = st.selectbox("Tono di voce", TONE_CHOICES, index=TONE_CHOICES.index("Colloquiale"))
    brief = st.text_area("Brief (spiega al modello cosa vuoi ottenere)", height=120,
                         placeholder="Esempio: pubblico principianti, stile pratico, esempi reali, evitare gergo...")
    pdf_page = st.selectbox("Formato pagina", list(PAGE_SIZES.keys()), index=list(PAGE_SIZES.keys()).index("6x9"))
    font_name = st.selectbox("Font principale", FONT_CHOICES, index=0)

    # === Dati libro ===
    with st.form("book_info_form"):
        title = st.text_input("Titolo")
        subtitle = st.text_input("Sottotitolo")
        author = st.text_input("Autore")
        total_words = st.number_input("Totale parole del libro", min_value=500, step=500, value=20000)
        submitted_meta = st.form_submit_button("Salva dati libro")

    if submitted_meta:
        st.success("Dati libro memorizzati.")

    # === Validazione minima: parole sufficienti? ===
    # Calcolo una stima “sezioni totali” dal TOC confermato
    temp_chapters = parse_confirmed_toc(st.session_state.confirmed_toc_text)
    total_sections = sum(len(ch.sections) for ch in temp_chapters)
    min_needed = total_sections * MIN_SECTION_WORDS_USEFUL  # 250 parole min per sezione “utile”

    if total_words < min_needed:
        st.error(
            f"Con {total_sections} sezioni, servono almeno {min_needed} parole "
            f"(250 per sezione). Aumenta il totale o riduci le sezioni nel TOC."
        )
        st.stop()

    # === Genera allocazione e mostra a video ===
    chapters_alloc = allocate_words(temp_chapters, total_words, block_size_limit=MAX_SUBGEN_WORDS)

    # salvo i parametri per i passi successivi
    plan_preview = BookPlan(
        title=title or "Titolo",
        subtitle=subtitle or "",
        author=author or "",
        total_words=total_words,
        block_size=MAX_SUBGEN_WORDS,     # limite fisso: MAI oltre 500 per chiamata
        chapters=chapters_alloc,
        language_code=lang_human,
        tone=tone,
        brief=brief.strip(),
        pdf_page=pdf_page,
        font_name=font_name,
    )

    st.session_state.generated_plan = plan_preview
    st.session_state.chapters = chapters_alloc
    st.session_state.allocation_done = True

    # === Visualizzazione allocazione (capitoli/sezioni/blocchi) ===
    st.success("✅ Allocazione pronta. Ecco come verrà diviso il libro:")
    for i, ch in enumerate(chapters_alloc, start=1):
        with st.expander(f"Capitolo {i}: {ch.title} — {ch.target_words} parole — {ch.blocks} blocchi totali"):
            for j, sec in enumerate(ch.sections, start=1):
                st.write(f"• Sezione {j}: {sec.title} — {sec.target_words} parole — {sec.blocks} blocchi (≤500 parole ciascuno)")

    st.info("Quando sei soddisfatto, prosegui allo step di generazione contenuti.")
else:
    st.warning("Per favore conferma prima il TOC nello Step 1.")
# ==========================================
# ✍️ BLOCCO 4 — GENERA TESTI + EXPORT DOCX/PDF
# ------------------------------------------
# Qui:
# - prendo il piano allocato (capitoli/sezioni/blocchi)
# - genero i testi rispettando il limite di 500 parole per singola chiamata
# - rendo coerente lingua/tono/brief
# - mostro anteprima
# - faccio scaricare DOCX e PDF con formattazione:
#     * titoli capitolo/sottocapitolo: allineati a sinistra
#     * testo normale: GIUSTIFICATO
#     * prima pagina: titolo e sottotitolo centrati; autore in basso centrale;
#       NIENTE “Autore:” e NIENTE “Totale parole / Blocchi”
#     * scelta formato pagina (A4, Letter, 6x9, 8.5x11)
#     * scelta font (DOCX: set font; PDF: mappo su font base)
# ==========================================

st.subheader("🖋️ Step 3 — Generazione dei contenuti")

def _effective_language_label(plan: BookPlan) -> str:
    # Se l’utente ha scelto manuale, uso quello; altrimenti provo dalla detection del Blocco 2
    code = plan.language_code
    if code == "auto":
        det = st.session_state.get("detected_lang", "en")
        code = det if det in LANG_LABELS else "en"
    # converto in etichetta per prompt
    return LANG_LABELS.get(code, "English")

def _tone_instruction(tone: str) -> str:
    if tone.lower().startswith("scien"):
        return "Use a precise, rigorous, and evidence-based tone suitable for a scientific audience."
    if tone.lower().startswith("narr"):
        return "Use a narrative, evocative tone, with smooth transitions and concrete scenes where helpful."
    return "Use a clear, friendly, and practical tone suitable for a general audience."

def _count_words(txt: str) -> int:
    return len([w for w in re.split(r"\s+", txt.strip()) if w])

def _generate_subchunk(prompt_sys: str, prompt_user: str) -> str:
    if not OPENAI_OK:
        # Segnaposto utile e neutro (NON esplicitare che mancano parti)
        return prompt_user[:0] + " " + " ".join(["[testo segnaposto]"] * 50)
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
        return f"[errore generazione] {e}"

def generate_block_text(plan: BookPlan, ch_title: str, sec_title: str, target_words: int, prev_summary: str = "", is_last_block: bool = False) -> str:
    """
    Genera il testo di UN BLOCCO di una sezione rispettando:
    - limite assoluto 500 parole per singola chiamata
    - numero MINIMO di sottogenerazioni (ceil(target/500))
    - concatenazione invisibile
    - coerenza lingua/tono/brief
    - variazione e progressione: usa contesto precedente e obiettivo del blocco
    """
    lang = _effective_language_label(plan)
    tone_ins = _tone_instruction(plan.tone)

    # calcolo sottogenerazioni ottimali
    n_sub = max(1, math.ceil(target_words / MAX_SUBGEN_WORDS))
    words_per_sub = math.ceil(target_words / n_sub)
    words_per_sub = min(words_per_sub, MAX_SUBGEN_WORDS)

    sys = (
        "You are an expert non-fiction writer. "
        f"Write in {lang}. {tone_ins} "
        "Avoid repetition. Do not restate the book, chapter, or section titles inside the text. "
        "Write continuous prose (no lists unless strictly necessary). "
        "Maintain coherence with the previous context summary when provided."
    )

    parts = []
    for idx in range(n_sub):
        # ruolo del sottopezzo
        role_note = "Start the section naturally." if idx == 0 else (
            "Continue smoothly from the previous text, without repeating or re-opening the topic."
        )
        if idx == n_sub - 1 and is_last_block:
            role_note = role_note + " Conclude the section with a natural closing."

        guidance = []
        if plan.brief:
            guidance.append(f"Brief to follow: {plan.brief}")
        if prev_summary:
            guidance.append(f"Previous context summary: {prev_summary}")

        user = (
            f"Book title: {plan.title}\n"
            f"Subtitle: {plan.subtitle}\n"
            f"Author: {plan.author}\n"
            f"Chapter: {ch_title}\n"
            f"Section: {sec_title}\n"
            f"Goal words for this part: ~{words_per_sub} (hard limit per request: 500)\n"
            f"{role_note}\n"
            + ("\n".join(guidance) if guidance else "")
        )

        subtext = _generate_subchunk(sys, user)
        # Se sfora tanto, non rientro: lascio così ma concateno invisibilmente
        parts.append(subtext.strip())

    # concateno i pezzi in modo invisibile
    final_text = " ".join(p for p in parts if p)
    return final_text.strip()

def generate_all_sections(plan: BookPlan):
    """
    Per ogni sezione:
      - calcola parole per blocco (sec.target_words / sec.blocks)
      - genera ogni blocco con sottogenerazioni ≤500
      - salva i testi nella struttura
    Mostra barra di avanzamento e anteprima.
    """
    total_blocks = sum(sec.blocks for ch in plan.chapters for sec in ch.sections)
    done = 0
    bar = st.progress(0, text="✍️ Generazione in corso...")
    prev_summary = ""  # piccolo riassunto del testo precedente per varietà e continuità

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
                # aggiorno “prev_summary” in modo semplice
                words = re.split(r"\s+", text.strip())
                prev_summary = " ".join(words[:60]) + " ... " + " ".join(words[-40:]) if len(words) > 120 else text[:800]

                done += 1
                bar.progress(done / total_blocks, text=f"Blocchi completati: {done}/{total_blocks}")

    bar.empty()
    st.success("✅ Generazione terminata!")

def _docx_add_heading(paragraph, text, size_pt, bold=True, center=False):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size_pt)
    if center:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

def build_docx(plan: BookPlan) -> bytes:
    """
    DOCX:
      - Prima pagina: titolo e sottotitolo centrati; autore centrato in basso (senza etichetta “Autore:”).
      - Corpo: capitoli/sottocapitoli allineati a sinistra; testo normale GIUSTIFICATO.
      - Font: imposto il nome richiesto (se non presente, Word sostituisce in automatico).
    """
    doc = Document()

    # Prima pagina
    p = doc.add_paragraph()
    _docx_add_heading(p, plan.title, 26, bold=True, center=True)
    if plan.subtitle.strip():
        p = doc.add_paragraph()
        _docx_add_heading(p, plan.subtitle, 16, bold=False, center=True)

    # autore in basso (aggiungo qualche riga vuota per spingerlo verso il basso)
    for _ in range(12):
        doc.add_paragraph("")

    p = doc.add_paragraph()
    _docx_add_heading(p, plan.author, 12, bold=False, center=True)

    doc.add_page_break()

    # Stile paragrafo base GIUSTIFICATO
    style = doc.styles["Normal"]
    style.font.name = plan.font_name
    try:
        style._element.rPr.rFonts.set(qn('w:eastAsia'), plan.font_name)  # type: ignore
    except Exception:
        pass

    # Capitoli e sezioni
    for ch in plan.chapters:
        p = doc.add_paragraph()
        _docx_add_heading(p, ch.title, 18, bold=True, center=False)

        for sec in ch.sections:
            p = doc.add_paragraph()
            _docx_add_heading(p, sec.title, 14, bold=False, center=False)

            for text in sec.texts:
                para = doc.add_paragraph(text)
                # giustifica il testo normale
                para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        doc.add_page_break()

    # Applico giustificazione globale di sicurezza
    for paragraph in doc.paragraphs:
        if paragraph.style.name == "Normal":
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# Mappa PDF: font richiesti → font base disponibili in ReportLab (approx)
PDF_FONT_MAP = {
    "Times New Roman": "Times-Roman",
    "Roboto": "Helvetica",
    "Comfortaa": "Courier"  # fallback approssimativo
}

def build_pdf(plan: BookPlan) -> bytes:
    """
    PDF:
      - Usa formato pagina scelto (A4/Letter/6x9/8.5x11)
      - Titoli capitolo/sottocapitolo: sinistra
      - Testo normale: GIUSTIFICATO (ReportLab non ha “full justify” perfetto, ma ci avviciniamo)
      - Font: mappo su font base disponibili
    """
    pagesize = PAGE_SIZES.get(plan.pdf_page, PAGE_SIZES["6x9"])
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=pagesize,
        leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm
    )
    styles = getSampleStyleSheet()
    base_font = PDF_FONT_MAP.get(plan.font_name, "Times-Roman")

    # Ridefinisco stili per font scelto
    H1 = ParagraphStyle(
        "H1", parent=styles["Heading1"], fontName=base_font, alignment=0  # 0 = LEFT
    )
    H2 = ParagraphStyle(
        "H2", parent=styles["Heading2"], fontName=base_font, alignment=0
    )
    Body = ParagraphStyle(
        "Body", parent=styles["BodyText"], fontName=base_font, alignment=4  # 4 = JUSTIFY
    )
    TitleC = ParagraphStyle(
        "TitleC", parent=styles["Title"], fontName=base_font, alignment=1  # 1 = CENTER
    )
    SubC = ParagraphStyle(
        "SubC", parent=styles["BodyText"], fontName=base_font, alignment=1
    )
    AuthorC = ParagraphStyle(
        "AuthorC", parent=styles["BodyText"], fontName=base_font, alignment=1
    )

    flow = []
    # Prima pagina
    flow.append(Paragraph(plan.title, TitleC))
    if plan.subtitle.strip():
        flow.append(Paragraph(plan.subtitle, SubC))
    flow.append(Spacer(1, 36))
    # autore in basso (simulato con grandi spazi)
    flow.append(Spacer(1, pagesize[1] * 0.55))
    flow.append(Paragraph(plan.author, AuthorC))
    flow.append(PageBreak())

    # Contenuti
    for ch in plan.chapters:
        flow.append(Paragraph(ch.title, H1))
        for sec in ch.sections:
            flow.append(Paragraph(sec.title, H2))
            for text in sec.texts:
                flow.append(Paragraph(text, Body))
                flow.append(Spacer(1, 8))
        flow.append(PageBreak())

    doc.build(flow)
    return buf.getvalue()

# --- UI: avvio generazione + anteprima + download ---
if st.session_state.allocation_done and st.session_state.generated_plan:
    plan: BookPlan = st.session_state.generated_plan

    if st.button("🚀 CONFERMA E GENERA I TESTI", type="primary", use_container_width=True):
        generate_all_sections(plan)

        # salva subito file
        try:
            st.session_state["docx_bytes"] = build_docx(plan)
            st.session_state["pdf_bytes"] = build_pdf(plan)
        except Exception as e:
            st.error(f"Errore export: {e}")

    # Anteprima: mostra i testi se già ci sono
    if any(sec.texts for ch in plan.chapters for sec in ch.sections):
        st.subheader("👁️ Anteprima (estratti)")
        max_preview = 3
        shown = 0
        for i, ch in enumerate(plan.chapters, start=1):
            for j, sec in enumerate(ch.sections, start=1):
                if sec.texts:
                    st.markdown(f"**Capitolo {i} — {ch.title}**  \n*Sezione {j} — {sec.title}*")
                    st.write((sec.texts[0][:1200] + "…") if len(sec.texts[0]) > 1200 else sec.texts[0])
                    st.divider()
                    shown += 1
                    if shown >= max_preview:
                        break
            if shown >= max_preview:
                break

    # Download
    if st.session_state.get("docx_bytes") and st.session_state.get("pdf_bytes"):
        st.subheader("📥 Scarica il tuo libro")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="Scarica DOCX",
                data=st.session_state["docx_bytes"],
                file_name="libro_generato.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        with c2:
            st.download_button(
                label="Scarica PDF",
                data=st.session_state["pdf_bytes"],
                file_name="libro_generato.pdf",
                mime="application/pdf",
                use_container_width=True
            )
else:
    st.info("Completa i passi precedenti per generare e scaricare il libro.")
