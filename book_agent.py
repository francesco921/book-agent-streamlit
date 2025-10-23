# book_agent.py
# Requisiti: pip install openai python-docx reportlab pydantic
# Ambiente: esporta OPENAI_API_KEY
# Nota: la TOC in DOCX si aggiorna in Word con Aggiorna campo. Nel PDF creiamo una TOC semplice.

import os
import math
import json
from dataclasses import dataclass, asdict
from typing import List, Dict, Optional, Tuple

from pydantic import BaseModel, Field, validator
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

try:
    from openai import OpenAI
    _OPENAI_AVAILABLE = True
except Exception:
    _OPENAI_AVAILABLE = False

# -------------- Config --------------
MODEL = "gpt-4.1-mini"
BLOCK_WORDS = 500
DEFAULT_DISCLAIMER = (
    "Disclaimer: il contenuto di questo libro ha scopo informativo. "
    "Non costituisce consulenza professionale. Verificare sempre le informazioni "
    "critiche e rivolgersi a un professionista qualificato quando necessario."
)
COPYRIGHT_TEMPLATE = "© {year} {author}. Tutti i diritti riservati."

# -------------- Data models --------------
class Section(BaseModel):
    id: str
    title: str
    words: int
    blocks: int

class Chapter(BaseModel):
    number: int
    title: str
    words: int
    sections: List[Section] = Field(default_factory=list)

class Project(BaseModel):
    title: str
    subtitle: str
    author: str
    total_words: int
    chapters: int
    use_sections: bool
    sections_per_chapter: int = 0
    toc: List[Chapter] = Field(default_factory=list)
    block_size: int = BLOCK_WORDS
    tone: str = "professionale, chiaro, concreto"
    style_guide: str = "corpo 11 pt, Calibri, giustificato; titoli coerenti; esempi pratici"

    @validator("sections_per_chapter")
    def check_sections(cls, v, values):
        if values.get("use_sections") and v < 1:
            raise ValueError("sections_per_chapter deve essere maggiore di zero se use_sections è True")
        return v

# -------------- LLM helper --------------
def _client() -> OpenAI:
    if not _OPENAI_AVAILABLE:
        raise RuntimeError("Libreria openai non disponibile")
    key = os.environ.get("OPENAI_API_KEY")
    if not key:
        raise RuntimeError("OPENAI_API_KEY non impostata")
    return OpenAI()

def llm(system: str, user: str, temperature: float = 0.2, max_tokens: int = 1800) -> str:
    client = _client()
    rsp = client.chat.completions.create(
        model=MODEL,
        temperature=temperature,
        max_tokens=max_tokens,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    return rsp.choices[0].message.content.strip()

# -------------- TOC generation and allocation --------------
def allocate_words(total: int, n: int) -> List[int]:
    base = total // n
    rem = total % n
    alloc = [base] * n
    for i in range(rem):
        alloc[i] += 1
    return alloc

def build_toc(project: Project) -> Project:
    # Genera titoli capitolo via LLM, poi ripartisce parole e sezioni
    system = "Sei un editor e outliner di saggistica. Fornisci titoli efficaci, chiari e non clickbait."
    prompt = f"""
Libro: {project.title}
Sottotitolo: {project.subtitle}
Autore: {project.author}
Obiettivo: sommario di {project.chapters} capitoli.
Vincoli: titoli concisi max 8 parole, nessun numero all'inizio. Niente punti finali.
Fornisci JSON con array "chapters": [{{"title": "..."}}] lungo {project.chapters}.
"""
    raw = llm(system, prompt, temperature=0.3, max_tokens=800)
    try:
        data = json.loads(raw)
        ch_titles = [c["title"] for c in data["chapters"]]
        if len(ch_titles) != project.chapters:
            raise ValueError("numero capitoli non coerente")
    except Exception:
        # fallback deterministico
        ch_titles = [f"Capitolo di tema: {i+1}" for i in range(project.chapters)]

    per_chapter = allocate_words(project.total_words, project.chapters)
    toc: List[Chapter] = []
    for idx, words in enumerate(per_chapter):
        chap = Chapter(number=idx + 1, title=ch_titles[idx], words=words, sections=[])
        if project.use_sections:
            # ripartizione in sezioni per capitolo
            per_section = allocate_words(words, project.sections_per_chapter)
            sections = []
            # titoli sezioni via LLM
            system_s = "Sei un editor. Fornisci titoli di sezioni concreti e non sensazionalistici."
            prompt_s = f"""
Capitolo: {chap.title}
Genera {project.sections_per_chapter} titoli di sezione, concisi.
JSON: {{"sections": [{{"title": "..."}}]}}
"""
            raw_s = llm(system_s, prompt_s, temperature=0.3, max_tokens=600)
            try:
                data_s = json.loads(raw_s)
                sec_titles = [s["title"] for s in data_s["sections"]]
                if len(sec_titles) != project.sections_per_chapter:
                    raise ValueError
            except Exception:
                sec_titles = [f"Sezione {j+1}" for j in range(project.sections_per_chapter)]
            for j, w in enumerate(per_section):
                blocks = max(1, round(w / project.block_size))
                sections.append(Section(
                    id=f"{chap.number}.{j+1}",
                    title=sec_titles[j],
                    words=w,
                    blocks=blocks,
                ))
            chap.sections = sections
        else:
            # nessuna sezione, si generano blocchi a livello capitolo con una sola sezione logica
            blocks = max(1, round(words / project.block_size))
            chap.sections = [Section(id=f"{chap.number}.1", title="Contenuto", words=words, blocks=blocks)]
        toc.append(chap)

    project.toc = toc
    return project

def summarize_toc(project: Project) -> str:
    lines = []
    for ch in project.toc:
        lines.append(f"Capitolo {ch.number} - {ch.title}  [{ch.words} parole]")
        for s in ch.sections:
            lines.append(f"  Sezione {s.id} - {s.title}  [{s.words} parole, {s.blocks} blocchi da ~{project.block_size}]")
    return "\n".join(lines)

# -------------- Block prompts and writing --------------
def write_block(project: Project, ch: Chapter, sec: Section, block_index: int) -> str:
    # block_index inizia da 1
    target_words = max(120, round(sec.words / sec.blocks))
    system = "Sei un writer professionale. Stile chiaro, denso di utilita, senza ridondanze. Evita claim medici o legali."
    user = f"""
Libro: {project.title}
Sottotitolo: {project.subtitle}
Autore: {project.author}
Tono: {project.tone}
Style guide: {project.style_guide}

Stai scrivendo un blocco di circa {target_words} parole.
Capitolo {ch.number}: {ch.title}
Sezione {sec.id}: {sec.title}
Blocco: {block_index} di {sec.blocks}

Istruzioni di struttura
1. Micro hook iniziale orientato al beneficio.
2. Sviluppo con un esempio o micro caso concreto.
3. Piccola checklist di 3 punti finali con verbi di azione.

Scrivi solo il testo definitivo, niente etichette o prefissi elenco.
"""
    text = llm(system, user, temperature=0.35, max_tokens=min(1200, target_words + 200))
    return text

# -------------- DOCX export --------------
def export_docx(project: Project, blocks_map: Dict[str, List[str]], out_path: str,
                disclaimer: str = DEFAULT_DISCLAIMER) -> str:
    doc = Document()

    # stile base
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # front matter
    p_title = doc.add_paragraph(project.title)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_title.runs[0]
    r.bold = True
    r.font.size = Pt(20)

    p_sub = doc.add_paragraph(project.subtitle)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_auth = doc.add_paragraph(f"Di {project.author}")
    p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_page_break()

    # copyright e disclaimer
    from datetime import datetime
    year = datetime.now().year
    doc.add_paragraph(COPYRIGHT_TEMPLATE.format(year=year, author=project.author))
    doc.add_paragraph(disclaimer)
    doc.add_page_break()

    # TOC semplice come elenco. In Word la TOC formattata si può aggiungere e aggiornare.
    toc_head = doc.add_paragraph("Indice")
    toc_head.runs[0].bold = True
    for ch in project.toc:
        doc.add_paragraph(f"Capitolo {ch.number}. {ch.title}")
        for s in ch.sections:
            doc.add_paragraph(f"  {s.id} {s.title}")
    doc.add_page_break()

    # contenuti
    for ch in project.toc:
        h1 = doc.add_paragraph(f"Capitolo {ch.number}. {ch.title}")
        h1.runs[0].bold = True
        h1.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for s in ch.sections:
            h2 = doc.add_paragraph(f"{s.id} {s.title}")
            h2.runs[0].bold = True

            blocks = blocks_map.get(s.id, [])
            for b in blocks:
                for para in b.split("\n\n"):
                    p = doc.add_paragraph(para.strip())
                    p.paragraph_format.first_line_indent = Cm(0.5)

    doc.save(out_path)
    return out_path

# -------------- PDF export --------------
def export_pdf(project: Project, blocks_map: Dict[str, List[str]], out_path: str,
               disclaimer: str = DEFAULT_DISCLAIMER) -> str:
    c = canvas.Canvas(out_path, pagesize=A4)
    width, height = A4
    left = 2.2 * cm
    top = height - 2.5 * cm
    line_height = 14

    def draw_paragraph(text: str, start_y: float) -> float:
        y = start_y
        for line in text.split("\n"):
            wrapped = wrap_text(line, 90)
            for w in wrapped:
                if y < 2.5 * cm:
                    c.showPage()
                    y = top
                c.drawString(left, y, w)
                y -= line_height
        return y

    # front
    c.setFont("Times-Roman", 20)
    c.drawCentredString(width / 2, top, project.title)
    c.setFont("Times-Roman", 14)
    c.drawCentredString(width / 2, top - 30, project.subtitle)
    c.drawCentredString(width / 2, top - 55, f"Di {project.author}")
    c.showPage()

    # disclaimer
    c.setFont("Times-Roman", 11)
    y = top
    y = draw_paragraph(COPYRIGHT_TEMPLATE.format(year=get_year(), author=project.author), y)
    y -= line_height
    y = draw_paragraph(disclaimer, y)
    c.showPage()

    # indice semplice
    y = top
    c.setFont("Times-Roman", 12)
    c.drawString(left, y, "Indice")
    y -= line_height * 2
    for ch in project.toc:
        y = draw_paragraph(f"Capitolo {ch.number}. {ch.title}", y)
        for s in ch.sections:
            y = draw_paragraph(f"  {s.id} {s.title}", y)
    c.showPage()

    # contenuti
    for ch in project.toc:
        y = top
        c.setFont("Times-Roman", 14)
        c.drawCentredString(width / 2, y, f"Capitolo {ch.number}. {ch.title}")
        y -= line_height * 2
        c.setFont("Times-Roman", 11)
        for s in ch.sections:
            if y < 3.5 * cm:
                c.showPage()
                y = top
            c.setFont("Times-Roman", 12)
            c.drawString(left, y, f"{s.id} {s.title}")
            y -= line_height * 1.5
            c.setFont("Times-Roman", 11)
            for b in blocks_map.get(s.id, []):
                for para in b.split("\n\n"):
                    y = draw_paragraph(para.strip(), y)
                    y -= line_height * 0.5
            y -= line_height

    c.save()
    return out_path

def wrap_text(text: str, width: int) -> List[str]:
    words = text.split()
    lines = []
    cur = []
    cur_len = 0
    for w in words:
        if cur_len + len(w) + (1 if cur else 0) <= width:
            cur.append(w)
            cur_len += len(w) + (1 if cur_len > 0 else 0)
        else:
            lines.append(" ".join(cur))
            cur = [w]
            cur_len = len(w)
    if cur:
        lines.append(" ".join(cur))
    return lines

def get_year() -> int:
    from datetime import datetime
    return datetime.now().year

# -------------- Interactive CLI --------------
def ask_int(prompt_text: str, minimum: int = 1) -> int:
    while True:
        try:
            v = int(input(prompt_text).strip())
            if v >= minimum:
                return v
        except Exception:
            pass
        print(f"Inserisci un intero maggiore o uguale a {minimum}.")

def yesno(prompt_text: str) -> bool:
    v = input(prompt_text + " [y/n]: ").strip().lower()
    return v in ["y", "yes", "s", "si"]

def cli():
    print("Configurazione progetto")
    title = input("Titolo: ").strip()
    subtitle = input("Sottotitolo: ").strip()
    author = input("Autore: ").strip()
    total_words = ask_int("Numero parole totali: ", minimum=5000)
    chapters = ask_int("Numero capitoli: ", minimum=3)
    use_sections = yesno("Vuoi dividere i capitoli in sezioni")
    sections_per = 0
    if use_sections:
        sections_per = ask_int("Quante sezioni per capitolo: ", minimum=1)

    project = Project(
        title=title,
        subtitle=subtitle,
        author=author,
        total_words=total_words,
        chapters=chapters,
        use_sections=use_sections,
        sections_per_chapter=sections_per,
    )

    # TOC loop
    while True:
        project = build_toc(project)
        print("\nProposta TOC con allocazione parole")
        print(summarize_toc(project))
        if yesno("Confermi il TOC proposto"):
            break
        action = input("Digita 'm' per modificare manualmente, 'r' per rigenerare: ").strip().lower()
        if action == "m":
            # modifica titoli o parole per capitolo
            for ch in project.toc:
                new_t = input(f"Titolo cap {ch.number} [{ch.title}] lascia vuoto per mantenere: ").strip()
                if new_t:
                    ch.title = new_t
                new_w = input(f"Parole cap {ch.number} [{ch.words}] lascia vuoto per mantenere: ").strip()
                if new_w.isdigit():
                    ch.words = int(new_w)
                # ricalcolo sezioni e blocchi
                if project.use_sections:
                    per_section = allocate_words(ch.words, project.sections_per_chapter)
                    for j, s in enumerate(ch.sections):
                        s.words = per_section[j]
                        s.blocks = max(1, round(s.words / project.block_size))
                else:
                    ch.sections[0].words = ch.words
                    ch.sections[0].blocks = max(1, round(ch.words / project.block_size))
        else:
            # rigenera
            continue

    # scrittura blocco per blocco
    blocks_map: Dict[str, List[str]] = {}
    print("\nScrittura contenuti a blocchi. Verrà richiesto di procedere blocco per blocco.")
    for ch in project.toc:
        for s in ch.sections:
            blocks_map.setdefault(s.id, [])
            for bidx in range(1, s.blocks + 1):
                if not yesno(f"Generare blocco {bidx}/{s.blocks} per sezione {s.id}"):
                    print("Interruzione richiesta. Si procede a export con quanto disponibile.")
                    return finalize(project, blocks_map)
                text = write_block(project, ch, s, bidx)
                print("\n--- Testo generato ---\n")
                print(text[:1000] + ("..." if len(text) > 1000 else ""))
                if yesno("Accetti questo blocco"):
                    blocks_map[s.id].append(text)
                else:
                    if yesno("Rigenero il blocco"):
                        text = write_block(project, ch, s, bidx)
                        print("\n--- Nuova versione ---\n")
                        print(text[:1000] + ("..." if len(text) > 1000 else ""))
                        if yesno("Accetti questa versione"):
                            blocks_map[s.id].append(text)
                    else:
                        print("Blocco saltato.")

    finalize(project, blocks_map)

def finalize(project: Project, blocks_map: Dict[str, List[str]]):
    os.makedirs("output", exist_ok=True)
    docx_path = f"output/{sanitize(project.title)}.docx"
    pdf_path = f"output/{sanitize(project.title)}.pdf"
    export_docx(project, blocks_map, docx_path)
    export_pdf(project, blocks_map, pdf_path)
    # salva anche JSON progetto
    with open(f"output/{sanitize(project.title)}.json", "w", encoding="utf-8") as f:
        json.dump(project.dict(), f, ensure_ascii=False, indent=2)
    print("\nExport completato")
    print(f"DOCX: {docx_path}")
    print(f"PDF:  {pdf_path}")
    print(f"JSON: output/{sanitize(project.title)}.json")

def sanitize(name: str) -> str:
    keep = "".join(c for c in name if c.isalnum() or c in " ._-")
    return keep.strip().replace(" ", "_")

if __name__ == "__main__":
    cli()
