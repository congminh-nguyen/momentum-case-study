#!/usr/bin/env python3
"""
Rebuild the Momentum / GBF case package into a clean, workstream-oriented
structure. Reading material is delivered as PDF, fill-in templates as Word,
data and trackers as Excel, and the board deck as PowerPoint.

Final structure:
  Participants/
    0_Start_Here/            (PDF)
    1_Client_Data_Room/      (PDF + Excel datasets)
    2_Workstream_A_Impact_Strategy/          (PDF + Word templates)
    3_Workstream_B_Operations_and_Analytics/ (PDF + Word templates)
    4_Shared_Toolkit/        (PDF + Excel toolkit + PowerPoint + Word templates)
  Facilitators/
    PDF/  Word/              (staff only)
"""

import shutil
import subprocess
import sys
import unicodedata
from pathlib import Path

from docx import Document
from docx.document import Document as _DocClass
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from fpdf import FPDF

ROOT = Path(__file__).resolve().parents[1]
PART = ROOT / "Participants"
FAC = ROOT / "Facilitators"
TOOLS = ROOT / "tools"
CACHE = TOOLS / ".cache"
VENV_PY = ROOT / ".venv" / "bin" / "python"

sys.path.insert(0, str(TOOLS))


# ---------------------------------------------------------------------------
# Text sanitisation for the core PDF font (Helvetica is Latin-1 only)
# ---------------------------------------------------------------------------

def _sanitize(text: str) -> str:
    if not text:
        return ""
    repl = {
        "\u2014": "-", "\u2013": "-", "\u2018": "'", "\u2019": "'",
        "\u201c": '"', "\u201d": '"', "\u2022": "-", "\u2192": "->",
        "\u2265": ">=", "\u2264": "<=", "\u2026": "...",
    }
    for k, v in repl.items():
        text = text.replace(k, v)
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")


# ---------------------------------------------------------------------------
# PDF renderer
# ---------------------------------------------------------------------------

class PdfExport(FPDF):
    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 8, "Momentum Programme  |  Scaling with Integrity  |  "
                  f"Page {self.page_no()}", align="C")

    def add_title(self, text):
        self.ln(1)
        self.set_font("Helvetica", "B", 17)
        self.set_text_color(27, 42, 74)
        self.multi_cell(0, 8, _sanitize(text))
        self.set_draw_color(0, 122, 135)
        self.set_line_width(0.6)
        y = self.get_y() + 1
        self.line(self.l_margin, y, self.w - self.r_margin, y)
        self.ln(4)

    def add_band(self, text):
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(0, 122, 135)
        self.multi_cell(0, 5, _sanitize(text))
        self.ln(1)

    def add_heading(self, text):
        self.ln(1)
        self.set_font("Helvetica", "B", 12)
        self.set_text_color(27, 42, 74)
        self.multi_cell(0, 6, _sanitize(text))
        self.ln(0.5)

    def add_para(self, text, italic=False):
        if not text.strip():
            self.ln(2)
            return
        self.set_font("Helvetica", "I" if italic else "", 10)
        self.set_text_color(45, 45, 45)
        self.multi_cell(0, 5, _sanitize(text))
        self.ln(1)

    def add_list_item(self, text, prefix="-  "):
        self.set_font("Helvetica", "", 10)
        self.set_text_color(45, 45, 45)
        x = self.l_margin + 4
        self.set_x(x)
        self.multi_cell(self.w - self.r_margin - x, 5, _sanitize(prefix + text))
        self.ln(0.5)

    def _break_word(self, word, w):
        """Split a single token that is wider than the column into pieces."""
        pieces, cur = [], ""
        for ch in word:
            if self.get_string_width(cur + ch) <= w - 2 or not cur:
                cur += ch
            else:
                pieces.append(cur)
                cur = ch
        if cur:
            pieces.append(cur)
        return pieces or [word]

    def _wrap(self, text, w):
        text = _sanitize(text)
        words = text.split()
        lines, cur = [], ""
        for word in words:
            if self.get_string_width(word) > w - 2:
                if cur:
                    lines.append(cur)
                    cur = ""
                pieces = self._break_word(word, w)
                lines.extend(pieces[:-1])
                cur = pieces[-1]
                continue
            trial = (cur + " " + word).strip()
            if self.get_string_width(trial) <= w - 2 or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = word
        if cur:
            lines.append(cur)
        return lines or [""]

    def add_grid(self, rows, header=True, widths=None):
        if not rows:
            return
        cols = max(len(r) for r in rows)
        epw = self.w - self.l_margin - self.r_margin
        if widths and len(widths) == cols:
            tot = sum(widths)
            col_w = [epw * x / tot for x in widths]
        else:
            col_w = [epw / cols] * cols
        line_h = 4.6
        for ri, row in enumerate(rows):
            cells = [str(c) for c in row] + [""] * (cols - len(row))
            is_header = header and ri == 0
            self.set_font("Helvetica", "B" if is_header else "", 8.5)
            wrapped = [self._wrap(cells[ci], col_w[ci]) for ci in range(cols)]
            nlines = max(len(w) for w in wrapped)
            row_h = line_h * nlines + 1.5
            if self.get_y() + row_h > self.page_break_trigger:
                self.add_page()
            x0, y0 = self.l_margin, self.get_y()
            if is_header:
                self.set_fill_color(27, 42, 74)
                self.set_text_color(255, 255, 255)
            elif ri % 2 == 0:
                self.set_fill_color(242, 244, 247)
                self.set_text_color(40, 40, 40)
            else:
                self.set_fill_color(255, 255, 255)
                self.set_text_color(40, 40, 40)
            for ci in range(cols):
                x = x0 + sum(col_w[:ci])
                self.rect(x, y0, col_w[ci], row_h, style="FD")
                ty = y0 + 0.8
                for ln in wrapped[ci]:
                    self.set_xy(x + 1, ty)
                    self.cell(col_w[ci] - 2, line_h, ln)
                    ty += line_h
            self.set_xy(x0, y0 + row_h)
        self.set_text_color(45, 45, 45)
        self.ln(2)


def _iter_blocks(doc):
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def _para_kind(para):
    style = para.style.name if para.style else ""
    if "Heading 1" in style:
        return "heading"
    if "Heading" in style:
        return "subheading"
    if "List Bullet" in style:
        return "bullet"
    if "List Number" in style:
        return "number"
    runs = [r for r in para.runs if r.text.strip()]
    if runs:
        r0 = runs[0]
        size = r0.font.size.pt if r0.font.size else None
        if r0.bold and size and size >= 20:
            return "title"
        if r0.bold and size and size >= 13:
            return "heading"
        if r0.font.color is not None and r0.font.color.type is not None:
            try:
                if r0.font.color.rgb is not None and str(r0.font.color.rgb) == "007A87":
                    return "band"
            except Exception:
                pass
    return "body"


def _row_is_header(row):
    seen = False
    for cell in row.cells:
        text = cell.text.strip()
        if not text:
            continue
        seen = True
        bold = any(run.bold for p in cell.paragraphs for run in p.runs)
        if not bold:
            return False
    return seen


def docx_to_pdf(docx_path: Path, pdf_path: Path):
    doc = Document(str(docx_path))
    pdf = PdfExport()
    pdf.set_auto_page_break(auto=True, margin=16)
    pdf.set_margins(18, 16, 18)
    pdf.add_page()
    number_counter = 0
    prev_kind = None
    for block in _iter_blocks(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                continue
            kind = _para_kind(block)
            if kind != "number":
                number_counter = 0
            if kind == "title":
                pdf.add_title(text)
            elif kind == "band":
                pdf.add_band(text)
            elif kind == "heading":
                pdf.add_heading(text)
            elif kind == "subheading":
                pdf.add_heading(text)
            elif kind == "bullet":
                pdf.add_list_item(text)
            elif kind == "number":
                number_counter += 1
                pdf.add_list_item(text, prefix=f"{number_counter}.  ")
            else:
                italic = any(r.italic for r in block.runs if r.text.strip())
                pdf.add_para(text, italic=italic)
            prev_kind = kind
        else:  # Table
            rows = [[c.text.strip() for c in r.cells] for r in block.rows]
            rows = [r for r in rows if any(cell for cell in r)]
            if not rows:
                continue
            header = _row_is_header(block.rows[0])
            widths = []
            for c in block.rows[0].cells:
                widths.append(c.width.inches if c.width else None)
            if any(w is None or w <= 0 for w in widths):
                widths = None
            pdf.add_grid(rows, header=header, widths=widths)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    pdf.output(str(pdf_path))


# ---------------------------------------------------------------------------
# Excel: datasets workbook
# ---------------------------------------------------------------------------

def csv_to_excel(raw_dir: Path, out_xlsx: Path):
    import csv
    wb = Workbook()
    wb.remove(wb.active)
    header_fill = PatternFill("solid", fgColor="007A87")
    header_font = Font(name="Calibri", bold=True, color="FFFFFF")

    ws = wb.create_sheet("DATA_DICTIONARY")
    ws["A1"] = "GBF datasets - field notes and known data-quality issues"
    ws["A1"].font = Font(bold=True, size=13, color="1B2A4A")
    ws["A2"] = "See the Workstream B brief (DB-1). These issues are intentional; find and handle them."
    ws["A2"].font = Font(italic=True, color="666666")
    ws["A4"], ws["B4"] = "Field / area", "What to watch for"
    for c in ("A4", "B4"):
        ws[c].font = header_font
        ws[c].fill = header_fill
    notes = [
        ("beneficiary_id", "Duplicate IDs exist by design; de-duplicate carefully"),
        ("placed_90d", "Coded as Y / Yes / 1 / N / No / 0 / blank - normalise before counting"),
        ("province", "'HCMC' and 'Ho Chi Minh City' both appear - standardise"),
        ("date fields", "Mixed formats (dd/mm/yyyy and yyyy-mm-dd)"),
        ("age", "At least one impossible value - investigate, do not delete silently"),
        ("wage_vnd_monthly", "At least one obvious typo/outlier"),
        ("hub = DN (Dong Nai)", "Underperforms; Q3 costs appear double-entered in hub_costs_2025"),
        ("volunteer_hours", "Hours logged against an inactive mentor - exclude for DB-4"),
        ("employment_type", "Formal vs Gig - the placement rate depends on which you count"),
        ("sample size", "This export (~520) is smaller than the 2,847 in D-01; state this limitation"),
    ]
    for i, (a, b) in enumerate(notes, 5):
        ws.cell(row=i, column=1, value=a)
        ws.cell(row=i, column=2, value=b)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 70

    for csv_file in sorted(raw_dir.glob("*.csv")):
        sheet = wb.create_sheet(csv_file.stem[:31])
        with open(csv_file, newline="", encoding="utf-8") as f:
            for ri, row in enumerate(csv.reader(f), 1):
                for ci, val in enumerate(row, 1):
                    cell = sheet.cell(row=ri, column=ci, value=val)
                    if ri == 1:
                        cell.font = header_font
                        cell.fill = header_fill
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = 16
        sheet.freeze_panes = "A2"
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_xlsx)
    print(f"  Excel: {out_xlsx.relative_to(ROOT)}")


# ---------------------------------------------------------------------------
# PowerPoint board deck template
# ---------------------------------------------------------------------------

def build_ppt(out_path: Path):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    NAVY = RGBColor(0x1B, 0x2A, 0x4A)
    TEAL = RGBColor(0x00, 0x7A, 0x87)
    GREY = RGBColor(0x66, 0x66, 0x66)

    def title_slide():
        s = prs.slides.add_slide(blank)
        box = s.shapes.add_textbox(Inches(0.9), Inches(2.4), Inches(11.5), Inches(2.6))
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = "GBF Strategy 2026-2028"
        p.font.size = Pt(40); p.font.bold = True; p.font.color.rgb = NAVY
        p2 = tf.add_paragraph()
        p2.text = "[Your one-sentence recommendation goes here]"
        p2.font.size = Pt(20); p2.font.color.rgb = TEAL
        p3 = tf.add_paragraph()
        p3.text = "Momentum Programme  |  Generation Bridge Foundation  |  13 August 2026"
        p3.font.size = Pt(13); p3.font.color.rgb = GREY

    def content_slide(headline, evidence, insight, ask, notes=""):
        s = prs.slides.add_slide(blank)
        hb = s.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12.1), Inches(1.3))
        htf = hb.text_frame; htf.word_wrap = True
        hp = htf.paragraphs[0]
        hp.text = headline
        hp.font.size = Pt(26); hp.font.bold = True; hp.font.color.rgb = NAVY
        bb = s.shapes.add_textbox(Inches(0.6), Inches(1.9), Inches(12.1), Inches(4.8))
        btf = bb.text_frame; btf.word_wrap = True
        for label, val in [("Evidence", evidence), ("Insight", insight),
                           ("What we recommend / ask", ask)]:
            lp = btf.add_paragraph()
            lr = lp.add_run(); lr.text = label; lr.font.bold = True
            lr.font.size = Pt(15); lr.font.color.rgb = TEAL
            vp = btf.add_paragraph()
            vr = vp.add_run(); vr.text = val
            vr.font.size = Pt(15); vr.font.color.rgb = RGBColor(0x30, 0x30, 0x30)
            btf.add_paragraph()
        if notes:
            s.notes_slide.notes_text_frame.text = notes

    title_slide()
    content_slide(
        "GBF should [action] - because [reason]",
        "[Lead stat + source, e.g. waitlist 2,400; placement 63% formal]",
        "[The single most important 'so what']",
        "[The specific decision you want from the Board]",
        "Open with the answer. You have 60 seconds of Board attention.")
    content_slide(
        "Demand has outpaced funding",
        "Waitlist ~2,400 (D-01); income grew 3% while need grew far faster (D-12)",
        "There is moral urgency to grow - but scaling a broken model repeats Dong Nai",
        "[Frame the central trade-off]")
    content_slide(
        "Young people drop out when transport costs bite",
        "BN-02 and ST-05; survey barrier data; Dong Nai internship drop-out 38% (D-02)",
        "A modest transport stipend may lift completion more than any new hub",
        "[Recommend a transport pilot? Quantify it]")
    content_slide(
        "The headline placement rate flatters us",
        "71% broad vs ~63% formal (ST-02); definition footnote in D-01",
        "Honest dual reporting protects the VPBank relationship, not endangers it",
        "[Recommend a reporting standard]")
    content_slide(
        "We weighed [X] against [Y] and [Z]",
        "Prioritisation matrix; funding scenarios (DB-5)",
        "[The decisive trade-off between the options]",
        "[Name your recommended option]")
    content_slide(
        "A 90-day plan to start well",
        "EP-02 and ST-05 on employer quality; D-11 on staff workload",
        "Early, visible wins build Board and donor confidence",
        "0-30 days: [ ...]   31-60: [ ...]   61-90: [ ...]")
    content_slide(
        "The decision we need today",
        "[Recap the one-line recommendation]",
        "[Why now]",
        "1) Approve direction  2) Authorise pilot budget  3) Mandate honest M&E")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    print(f"  PowerPoint: {out_path.relative_to(ROOT)}")


# ---------------------------------------------------------------------------
# Facilitator materials
# ---------------------------------------------------------------------------

def build_facilitator_docs():
    from brand import (set_doc_defaults, add_cover, add_heading, add_body,
                       add_bullet, add_number, add_table, add_quote)

    fac_word = FAC / "Word"
    fac_word.mkdir(parents=True, exist_ok=True)

    def newdoc(title, subtitle):
        d = Document()
        set_doc_defaults(d)
        add_cover(d, title, subtitle, classification="FACILITATOR ONLY - do not distribute",
                  doc_ref="Facilitator pack")
        return d

    # Facilitator Guide
    d = newdoc("Facilitator Guide", "How to run the six-week Momentum engagement")
    add_heading(d, "Before you start", 1)
    add_body(d, "Distribute only the Participants/ folder. Never share the Facilitators/ folder. "
             "Confirm each team has five or six members and assign one Buddy per team.")
    add_heading(d, "Kickoff (Day 1) - what to say", 1)
    add_body(d, "Set the tone in the first ten minutes:")
    add_bullet(d, "You are consultants, not students. The Board is paying for judgement.")
    add_bullet(d, "There is no single right answer. We assess your reasoning and evidence.")
    add_bullet(d, "Everything you need is in the pack. Do not go looking outside it.")
    add_bullet(d, "The two workstreams must audit each other. That tension is deliberate.")
    add_heading(d, "Week-by-week facilitation", 1)
    add_table(d, ["Week", "Watch for", "Nudge if..."], [
        ("1", "Vague problem statements; skipping the credibility check",
         "they treat every source as equally reliable"),
        ("2", "Data cleaning done silently; empathy maps with no citations",
         "they delete outliers without recording why"),
        ("3", "Options that are not genuinely distinct",
         "all three options are really the same idea"),
        ("4", "Recommendations that ignore validation",
         "nothing changed after they read the transcripts"),
        ("5", "Data-dump slides; burying bad news",
         "Dong Nai or the 63% formal rate is missing"),
        ("6", "Shallow reflection; blame in peer review",
         "the debrief avoids the ethical trade-offs"),
    ], widths=[0.6, 3.0, 3.0])
    add_heading(d, "Escalation triggers", 1)
    add_bullet(d, "Team conflict unresolved for more than 48 hours.")
    add_bullet(d, "Any safeguarding or ethical red flag in how a team behaves.")
    add_bullet(d, "A Buddy giving answers rather than asking questions.")
    d.save(fac_word / "01_Facilitator_Guide.docx")

    # Buddy Guide
    d = newdoc("Buddy Guide", "Progressive hints, not answers")
    add_body(d, "Your job is to ask better questions, not to solve the case. Use the three-level "
             "hint ladder below. Only move to the next level if the team is genuinely stuck.")
    add_heading(d, "The hint ladder", 1)
    for topic, l1, l2, l3 in [
        ("Problem framing",
         "What exactly is the Board asking you to decide?",
         "Is your problem statement about programmes, geography, funding, or all three?",
         "Point them to the four factions in the handbook and ask which trade-off matters most."),
        ("Data credibility",
         "Do all your sources deserve equal trust?",
         "Compare what D-01 says with what D-02 says - which is candid?",
         "Ask them to build the credibility matrix before believing any single number."),
        ("Placement rate",
         "What does 'placed' actually mean here?",
         "Whose definition are you using - D-01's or ST-02's?",
         "Have them compute both the broad and the formal rate and report both."),
        ("Recommendations",
         "What is the evidence for that claim?",
         "Which transcript or dataset backs this up?",
         "Ask them to map every recommendation to a source in the validation protocol."),
    ]:
        add_heading(d, topic, 2)
        add_number(d, "Level 1: " + l1)
        add_number(d, "Level 2: " + l2)
        add_number(d, "Level 3: " + l3)
    add_heading(d, "Check-in schedule", 1)
    add_body(d, "Fridays of Weeks 1-5. Fifteen minutes. Review project-management discipline "
             "(logs, minutes, RACI) and whether claims are cited. Do not review answers for correctness.")
    add_heading(d, "Common mistakes to expect", 1)
    add_bullet(d, "Accepting the 71% placement rate at face value.")
    add_bullet(d, "Recommending national scale without fixing Dong Nai.")
    add_bullet(d, "Treating the board email (D-08) as fact.")
    add_bullet(d, "Claiming attendance causes placement.")
    add_bullet(d, "Hiding the disability placement gap.")
    d.save(fac_word / "02_Buddy_Guide.docx")

    # Assessment Rubric
    d = newdoc("Assessment Rubric", "100 points - descriptors for each level")
    add_body(d, "Score each criterion against the descriptors. Weight as shown.")
    for crit, pts, exc, good, sat, poor in [
        ("Problem understanding", 15,
         "Precise, MECE issue tree; clear trade-offs identified",
         "Solid framing; minor gaps",
         "Generic framing; some overlap",
         "Misreads the problem"),
        ("Research quality", 10,
         "Sources triangulated; credibility assessed; contradictions resolved",
         "Most sources used well",
         "Uses sources uncritically",
         "Ignores or misuses sources"),
        ("Design thinking", 15,
         "Evidence-based empathy and journey maps; ideation tied to pain points",
         "Competent maps and ideation",
         "Mechanical, thin on evidence",
         "Superficial or missing"),
        ("Data analysis", 20,
         "Clean data, honest definitions, correct interpretation, limitations stated",
         "Mostly sound; small errors",
         "Basic analysis; weak interpretation",
         "Errors or misleading charts"),
        ("Recommendations", 20,
         "Distinct options; clear, feasible, evidence-based choice; ethics engaged",
         "Reasonable, mostly supported",
         "Vague or weakly supported",
         "Unsupported or infeasible"),
        ("Presentation", 10,
         "Answer-first storyline; every slide earns its place",
         "Clear and mostly disciplined",
         "Some data-dumping",
         "Disorganised or decorative"),
        ("Project management", 5,
         "Logs, RACI, version control all maintained",
         "Mostly maintained",
         "Patchy",
         "Absent"),
        ("Teamwork", 5,
         "Genuine cross-audit; shared ownership",
         "Good collaboration",
         "Uneven contribution",
         "Free-riding evident"),
    ]:
        add_heading(d, f"{crit} ({pts} points)", 2)
        add_table(d, ["Excellent", "Good", "Satisfactory", "Poor"],
                  [(exc, good, sat, poor)], widths=[1.7, 1.6, 1.6, 1.5])
    d.save(fac_word / "03_Assessment_Rubric.docx")

    # Solution Paths
    d = newdoc("Solution Paths", "CONFIDENTIAL - there is no single correct answer")
    add_body(d, "Reward reasoning, not a particular destination. The following are strong, "
             "defensible paths. Weak answers pick one slogan and ignore the evidence against it.")
    for name, thesis, strengths, risks in [
        ("Path A - Depth before scale",
         "Pause expansion, fix Dong Nai and employer quality, negotiate a smaller quality-focused VPBank grant.",
         "Honest about D-02; addresses EP-02 and ST-05; sustainable.",
         "May leave money on the table; must convince VPBank (see DN-01)."),
        ("Path B - Conditional scale (most common strong answer)",
         "Accept VPBank but gate expansion behind quality milestones; fix Dong Nai first, then Long An.",
         "Balances DN-01 and BD-01; keeps the grant; stages risk.",
         "Execution-heavy; needs credible M&E to hold the line."),
        ("Path C - Employer-led pivot",
         "Certify and audit employers; place only with vetted partners; make quality the brand.",
         "Directly answers EP-02, ST-05, BN-02; differentiates from YouthWorks.",
         "Slower growth; needs employer buy-in."),
        ("Path D - Digital hybrid",
         "Grow Pathway Digital for reach, keep Skills Forward for depth.",
         "Lower unit cost (D-12); serves waitlist faster.",
         "Digital placement is weaker; equity risk (BN-01 has no laptop); only valid with caveats."),
    ]:
        add_heading(d, name, 2)
        add_body(d, "Thesis: " + thesis)
        add_body(d, "Strengths: " + strengths)
        add_body(d, "Risks to probe: " + risks)
    add_heading(d, "Red lines (mark down if crossed)", 1)
    add_bullet(d, "Recommending 60-day reporting to look better (fails ED-02).")
    add_bullet(d, "Replacing caseworkers with volunteers (contradicts VR-01).")
    add_bullet(d, "National scale with no fix for Dong Nai.")
    d.save(fac_word / "04_Solution_Paths.docx")

    # Data Task Solutions
    d = newdoc("Data Task Solutions", "CONFIDENTIAL - expected findings and pitfalls")
    add_heading(d, "DB-1 Data quality - what they should find", 1)
    add_bullet(d, "Duplicate beneficiary IDs (built in) - de-duplicate before counting.")
    add_bullet(d, "placed_90d coded as Y/Yes/1/N/No/0/blank - normalise.")
    add_bullet(d, "Mixed date formats across tables.")
    add_bullet(d, "A wage outlier (obvious typo) on one record.")
    add_bullet(d, "Dong Nai Q3 costs double-entered in hub_costs_2025.")
    add_bullet(d, "Hours logged against an inactive mentor in volunteer_hours.")
    add_bullet(d, "'HCMC' vs 'Ho Chi Minh City' inconsistency.")
    add_heading(d, "DB-2/DB-3 expected patterns", 1)
    add_bullet(d, "Dong Nai placement clearly below HCMC hubs.")
    add_bullet(d, "Formal-only rate roughly 8 points below the broad rate (aligns with ST-02).")
    add_bullet(d, "Higher attendance associates with placement - but confounded by transport/motivation.")
    add_bullet(d, "Disability placement materially below overall (aligns with ST-03).")
    add_heading(d, "DB-4 / DB-5", 1)
    add_bullet(d, "Mentor hours weakly associate with completion; do not overclaim causation.")
    add_bullet(d, "Aggressive growth struggles to hold 65% honestly without cherry-picking - the ED-01 tension.")
    d.save(fac_word / "05_Data_Task_Solutions.docx")

    # Debrief + FAQ
    d = newdoc("Debrief Guide & FAQ", "Run a 90-minute debrief; answer common questions")
    add_heading(d, "Debrief flow (90 minutes)", 1)
    add_number(d, "Team reflection (15 min): what surprised you?")
    add_number(d, "Evidence round (20 min): each team names the fact that changed their mind.")
    add_number(d, "Ethics round (20 min): how did you resolve ED-01 and ED-02?")
    add_number(d, "Reveal (20 min): show the multiple valid paths; there was no single answer.")
    add_number(d, "Skills forward (15 min): one thing each person will practise next.")
    add_heading(d, "Frequently asked questions", 1)
    add_body(d, "Can we collect more data or interview real NGO staff?", bold=True)
    add_body(d, "No. Everything required is in the pack. Managing incomplete information is part "
             "of the exercise; state your assumptions instead.")
    add_body(d, "Which placement rate is correct?", bold=True)
    add_body(d, "Both are 'correct' under their definitions. The learning point is to state the "
             "definition and report honestly (D-01 footnote vs ST-02).")
    add_body(d, "Is there a right recommendation?", bold=True)
    add_body(d, "No. Several paths are defensible (see Solution Paths). Judgement and evidence win.")
    d.save(fac_word / "06_Debrief_Guide_and_FAQ.docx")

    # Convert all facilitator docs to PDF
    fac_pdf = FAC / "PDF"
    for docx in sorted(fac_word.glob("*.docx")):
        docx_to_pdf(docx, fac_pdf / docx.with_suffix(".pdf").name)
        print(f"  Facilitator: {docx.stem}")


# ---------------------------------------------------------------------------
# Root quick-start PDFs
# ---------------------------------------------------------------------------

def build_root_pdfs():
    pdf = PdfExport()
    pdf.set_margins(18, 16, 18)
    pdf.add_page()
    pdf.add_title("START HERE")
    pdf.add_band("Momentum Programme  |  Generation Bridge Foundation engagement")
    for line in [
        "Programme runs Monday 13 July to Friday 21 August 2026.",
        "Board presentation: Thursday 13 August 2026, 09:00 ICT.",
        "All final deliverables: Friday 14 August 2026, 17:00 ICT.",
        "",
    ]:
        pdf.add_para(line)
    pdf.add_heading("Read in this order")
    for i, line in enumerate([
        "0_Start_Here/00_Programme_Handbook.pdf",
        "0_Start_Here/01_Data_Room_Index.pdf",
        "0_Start_Here/02_Deadlines_and_Milestones.pdf",
        "2_Workstream_A_Impact_Strategy/WSA_Start_Here.pdf",
        "3_Workstream_B_Operations_and_Analytics/WSB_Start_Here.pdf",
        "4_Shared_Toolkit/GBF_Consulting_Toolkit.xlsx (Master Timeline tab)",
    ], 1):
        pdf.add_list_item(line, prefix=f"{i}.  ")
    pdf.add_heading("The rules")
    for line in [
        "Use only the materials in this package - no external research.",
        "Cite every claim (document ref, transcript ref, or dataset).",
        "Read as PDF, fill in as Word, analyse in Excel, present in PowerPoint.",
    ]:
        pdf.add_list_item(line)
    pdf.output(str(ROOT / "START_HERE.pdf"))
    print("  Root: START_HERE.pdf")

    pdf = PdfExport()
    pdf.set_margins(18, 16, 18)
    pdf.add_page()
    pdf.add_title("Package Map")
    pdf.add_heading("Participants/  (distribute this folder)")
    pdf.add_grid([
        ["Folder", "Contents", "Formats"],
        ["0_Start_Here", "Handbook, data room index, deadlines", "PDF"],
        ["1_Client_Data_Room", "GBF documents, interview transcripts, datasets", "PDF, Excel"],
        ["2_Workstream_A_Impact_Strategy", "Start here, brief, playbook, templates A1-A7", "PDF, Word"],
        ["3_Workstream_B_Operations_and_Analytics", "Start here, brief, dashboard, templates B1-B5", "PDF, Word, Excel"],
        ["4_Shared_Toolkit", "Toolkit workbook, presentation guide, board deck, shared templates", "Excel, PDF, PowerPoint, Word"],
    ], widths=[2.2, 3.4, 1.6])
    pdf.add_heading("Facilitators/  (staff only - do not share)")
    pdf.add_para("Facilitator guide, buddy guide, assessment rubric, solution paths, data-task "
                 "solutions, debrief guide and FAQ - in PDF and Word.")
    pdf.output(str(ROOT / "PACKAGE_MAP.pdf"))
    print("  Root: PACKAGE_MAP.pdf")


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def clean():
    for target in (PART, FAC):
        if target.exists():
            shutil.rmtree(target)
    # Remove any stray legacy numbered folders / old readmes
    for p in ROOT.iterdir():
        if p.is_dir() and p.name[:2].isdigit() and p.name not in ("Participants", "Facilitators"):
            shutil.rmtree(p)
    for stray in ("README.md",):
        f = ROOT / stray
        if f.exists():
            f.unlink()


def main():
    print("=== Clean ===")
    clean()

    print("=== Step 1: generate data ===")
    subprocess.run([str(VENV_PY), str(TOOLS / "generate_data.py")], check=True)

    print("=== Step 2: build documents (staging) ===")
    import content
    manifest = content.build_all()

    print("=== Step 3: place documents (PDF read-only / Word templates) ===")
    for item in manifest:
        src = content.STAGING / f"{item['name']}.docx"
        dest_dir = PART / item["dest"]
        if item["fmt"] == "pdf":
            docx_to_pdf(src, dest_dir / f"{item['name']}.pdf")
        else:
            tdir = dest_dir / "Templates"
            tdir.mkdir(parents=True, exist_ok=True)
            shutil.copy(src, tdir / f"{item['name']}.docx")
    print(f"  placed {len(manifest)} documents")

    print("=== Step 4: Excel workbooks ===")
    import generate_excel as ge
    ge.OUT = PART / "4_Shared_Toolkit"
    ge.OUT.mkdir(parents=True, exist_ok=True)
    ge.main()
    csv_to_excel(CACHE / "raw", PART / "1_Client_Data_Room" / "GBF_Datasets.xlsx")

    print("=== Step 5: PowerPoint ===")
    build_ppt(PART / "4_Shared_Toolkit" / "GBF_Board_Deck_Template.pptx")

    print("=== Step 6: Facilitator materials ===")
    build_facilitator_docs()

    print("=== Step 7: Root quick-start PDFs ===")
    build_root_pdfs()

    print("\nDone. Structure:")
    for d in sorted(PART.iterdir()):
        print(f"  Participants/{d.name}/")
    print("  Facilitators/PDF, Facilitators/Word")


if __name__ == "__main__":
    main()
