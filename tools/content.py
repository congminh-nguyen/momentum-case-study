#!/usr/bin/env python3
"""
Realistic document content for the Momentum / GBF case study.

Each builder writes a .docx into the staging folder and is registered in
MANIFEST with its final destination folder and output format:
  fmt = "pdf"  -> read-only reference material (converted to PDF)
  fmt = "docx" -> fill-in template (delivered as Word)
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt

from brand import (
    set_doc_defaults, add_cover, add_heading, add_body, add_bullet, add_number,
    add_table, add_meta_table, add_quote, add_signature_block, add_page_break,
)

ROOT = Path(__file__).resolve().parents[1]
STAGING = ROOT / "tools" / ".cache" / "word"
STAGING.mkdir(parents=True, exist_ok=True)

MANIFEST = []  # list of dicts: {name, dest, fmt}


def _doc(title, subtitle="", classification="Participant material", ref="", issued=""):
    doc = Document()
    set_doc_defaults(doc)
    add_cover(doc, title, subtitle, classification, ref, issued)
    return doc


def _save(doc, name, dest, fmt):
    path = STAGING / f"{name}.docx"
    doc.save(path)
    MANIFEST.append({"name": name, "dest": dest, "fmt": fmt})
    print(f"  built: {name} -> {dest} ({fmt})")
    return path


# ===========================================================================
# 0. START HERE
# ===========================================================================

def build_programme_handbook():
    doc = _doc(
        "Programme Handbook & Engagement Brief",
        "Everything your team needs to run the Generation Bridge Foundation engagement",
        ref="MOM-2026-HB-01", issued="Monday 13 July 2026",
    )

    add_heading(doc, "1. Welcome to the engagement", 1)
    add_body(doc, "Your team has been retained as external strategy consultants by the Board "
             "of Generation Bridge Foundation (GBF), a Vietnamese non-profit working on youth "
             "unemployment. Over the next six weeks you will diagnose a genuine strategic "
             "problem, test your thinking against incomplete evidence, and present a "
             "recommendation to the Board.")
    add_body(doc, "This is a simulated but realistic engagement. There is no single correct "
             "answer. You will be assessed on the quality of your thinking, the honesty of your "
             "evidence, the discipline of your teamwork, and the clarity of your communication.")
    add_body(doc, "Everything you need is contained in this package. You must not use external "
             "research, the internet, or any information beyond the materials provided. "
             "Operational datasets are held by Workflow B only. Where evidence is incomplete, "
             "say so; managing ambiguity is part of the exercise.", bold=True)

    add_heading(doc, "2. The client at a glance", 1)
    add_table(doc, ["Item", "Detail"], [
        ("Organisation", "Generation Bridge Foundation (Cau Noi The He)"),
        ("Founded", "2014, Ho Chi Minh City"),
        ("Mission", "Connect 16-24 year-olds who are not in education, employment or training (NEET) to dignified work"),
        ("2025 budget", "VND 18.2 billion"),
        ("Young people served in 2025", "2,847"),
        ("Current waitlist (Jan 2026)", "approximately 2,400"),
        ("Staff", "67 full-time, 12 part-time, ~340 active volunteers"),
        ("Delivery hubs", "Ho Chi Minh City (2), Binh Duong (1), Dong Nai (1), Long An (pilot)"),
    ], widths=[2.4, 4.2])

    add_heading(doc, "3. The question the Board has asked you", 1)
    add_quote(doc, "What should GBF's strategy be for 2026 to 2028, and how should we allocate "
              "our limited resources to maximise dignified youth employment outcomes?",
              "Engagement letter from the GBF Board, 15 January 2026")
    add_body(doc, "The Board is divided. Read the client data room carefully: senior staff, the "
             "largest donor and individual board members do not agree with one another. Some of "
             "what you are told will not survive contact with the operational record held by "
             "Workflow B.")

    add_heading(doc, "4. How your team is organised: two workflows", 1)
    add_body(doc, "Every team of five or six splits into two consulting workflows. The split is "
             "deliberate. Each workflow holds incomplete information. You must talk to each other "
             "to converge. Read both engagement briefs.")
    add_table(doc, ["Workflow", "Owns", "Folder"], [
        ("A - Impact Strategy", "Problem framing, stakeholder synthesis, the questions worth asking, "
         "strategic options, and the Board narrative",
         "2_Workflow_A_Impact_Strategy"),
        ("B - Operations & Analytics", "The operational datasets, data dictionary, cleaning, "
         "analysis, and plain-language translation of findings",
         "3_Workflow_B_Operations_and_Analytics"),
    ], widths=[1.7, 3.0, 2.0])
    add_heading(doc, "The analysis request protocol", 2)
    add_body(doc, "Workflow A does not receive the operational datasets and must not ask "
             "facilitators for the file. When A needs numbers, A files an Analysis Request. "
             "Workflow B answers with a Findings Memo in plain language, plus any charts B "
             "chooses to share. B must not hand over the raw workbook.")
    add_body(doc, "Before final options, the team must complete at least three Analysis Requests "
             "with Findings Memos returned. Every quantitative claim in A's deliverables must "
             "cite a Request ID or Findings Memo. The Cross-Workflow Handoff Checklist in the "
             "Shared Toolkit is the gate before submission.", bold=True)

    add_heading(doc, "5. What you will produce", 1)
    add_table(doc, ["Deliverable", "Owner", "Format", "Due"], [
        ("Project Charter & RACI", "Whole team", "Word -> PDF; Excel", "Tue 14 Jul"),
        ("Problem statement & issue tree", "Workflow A", "Word -> PDF", "Wed 15 Jul"),
        ("Stakeholder map", "Workflow A", "Word -> PDF", "Thu 16 Jul"),
        ("Information credibility matrix", "Whole team", "Word -> PDF", "Thu 16 Jul"),
        ("Analysis Requests (min. 3 across Weeks 2-4)", "Workflow A", "Word -> PDF", "Ongoing"),
        ("Data quality report (internal to B)", "Workflow B", "Word -> PDF", "Tue 21 Jul"),
        ("Findings Memos (responses to requests)", "Workflow B", "Word -> PDF", "Ongoing"),
        ("Empathy & journey maps", "Workflow A", "Word -> PDF", "Wed 22 Jul"),
        ("Performance dashboard (B working file)", "Workflow B", "Excel", "Fri 24 Jul"),
        ("Strategic options (cite Request IDs)", "Workflow A", "Word -> PDF", "Thu 30 Jul"),
        ("Evidence-check memo", "Workflow B", "Word -> PDF", "Fri 31 Jul"),
        ("Trade-off reflection", "Workflow A", "Word -> PDF", "Wed 5 Aug"),
        ("Funding scenarios", "Workflow B", "Excel", "Thu 6 Aug"),
        ("Validation protocol & revision log", "Whole team", "Word -> PDF", "Thu 6 Aug"),
        ("Final report", "Whole team", "Word -> PDF", "Fri 14 Aug"),
        ("Board presentation", "Whole team", "PowerPoint", "Thu 13 Aug"),
        ("Peer evaluation", "Individual", "Word -> PDF", "Mon 17 Aug"),
        ("Reflection journal", "Individual", "Word -> PDF", "Tue 18 Aug"),
    ], widths=[3.0, 1.4, 1.3, 1.0])
    add_body(doc, "The complete task-by-task timeline is in GBF_Consulting_Toolkit.xlsx, "
             "tab 'Master Timeline'. Track requests on the 'Request Log' tab.")

    add_heading(doc, "6. Programme calendar", 1)
    add_table(doc, ["Week", "Dates (2026)", "Focus"], [
        ("1", "13-17 July", "Frame the problem; set up the request protocol"),
        ("2", "20-24 July", "B cleans data; A files first requests; map beneficiary voices"),
        ("3", "27-31 July", "Findings exchange; options drafted; evidence-check"),
        ("4", "3-7 August", "Validate against interviews; model funding; revise"),
        ("5", "10-14 August", "Finalise report and deck; present to the Board"),
        ("6", "17-21 August", "Peer evaluation, reflection, debrief"),
    ], widths=[0.8, 1.8, 4.0])

    add_heading(doc, "7. How you will be assessed", 1)
    add_body(doc, "A 100-point rubric applies. The weighting tells you where to invest effort.")
    add_table(doc, ["Criterion", "Points"], [
        ("Problem understanding", "15"), ("Research quality", "10"),
        ("Design thinking", "15"), ("Data analysis and translation", "20"),
        ("Recommendations", "20"), ("Presentation", "10"),
        ("Project management", "5"), ("Teamwork and cross-workflow exchange", "5"),
    ], widths=[3.6, 1.2])
    add_body(doc, "Strong teams ask precise questions, translate findings honestly, and own the "
             "costs of their preferred recommendation. Numbers without judgement, or judgement "
             "without numbers, both fall short.")

    add_heading(doc, "8. Ground rules", 1)
    add_number(doc, "Use only the materials in this package. Do not invent facts or seek outside data.")
    add_number(doc, "Workflow A must not open or request the raw operational datasets. Route all "
               "quantitative needs through Analysis Requests to Workflow B.")
    add_number(doc, "Workflow B must not share the raw workbook with Workflow A. Share Findings "
               "Memos and curated charts only.")
    add_number(doc, "Every claim must cite its source: a document (D-02), a transcript (ST-01), "
               "or a Request / Findings ID (R-03 / FM-03).")
    add_number(doc, "State the definitions you use, especially for the placement rate.")
    add_number(doc, "Deliverables are read as PDF, filled in as Word, analysed in Excel, and "
               "presented in PowerPoint.")
    add_number(doc, "Your Buddy is a facilitator, not a consultant. They will ask questions and "
               "offer graduated hints, but they will not give you the answer or the data file.")

    add_heading(doc, "9. Package contents", 1)
    add_table(doc, ["Folder", "What is inside"], [
        ("0_Start_Here", "This handbook, the data room index, the deadline summary"),
        ("1_Client_Data_Room", "GBF documents (D-01 to D-12) and interview transcripts (no datasets)"),
        ("2_Workflow_A_Impact_Strategy", "Workflow A brief, playbook, A templates"),
        ("3_Workflow_B_Operations_and_Analytics", "Workflow B brief, data dictionary, datasets, dashboard, B templates"),
        ("4_Shared_Toolkit", "Toolkit workbook (incl. Request Log), guides, board deck, shared templates"),
    ], widths=[2.9, 3.6])

    _save(doc, "00_Programme_Handbook", "0_Start_Here", "pdf")


def build_data_room_index():
    doc = _doc("Data Room Index", "A guide to documents and transcripts in the shared client room",
               ref="MOM-2026-DRI-01", issued="Monday 13 July 2026")
    add_body(doc, "This index lists everything in the shared client data room. The reliability "
             "column is our honest assessment; form your own view in the information credibility "
             "matrix. Operational datasets are not in this folder. They are held by Workflow B "
             "under the analysis request protocol.")

    add_heading(doc, "Client documents", 1)
    add_table(doc, ["Ref", "Document", "Reliability", "Watch for"], [
        ("D-01", "Annual Report 2024 (public)", "Medium",
         "Placement rate uses a generous definition"),
        ("D-02", "Programme Review 2025 (internal memo)", "High",
         "Candid; contradicts parts of D-01"),
        ("D-03", "VPBank Foundation grant agreement (excerpt)", "Medium",
         "Conditions are demanding; targets are ambitious"),
        ("D-08", "Board email thread", "Low to medium",
         "Strong opinions; politically motivated"),
        ("D-10", "Competitor brief: YouthWorks", "Medium",
         "May overstate or understate the threat"),
        ("D-11", "Staff engagement survey 2025", "High",
         "Reliable on staff sentiment and workload"),
        ("D-12", "Financial statements 2023-2025 (audited)", "High",
         "Trustworthy; use for all cost figures"),
    ], widths=[0.6, 3.0, 1.2, 1.9])

    add_heading(doc, "Interview transcripts", 1)
    add_body(doc, "The full transcripts are in '20_Interview_Transcripts'. Fourteen "
             "conversations were recorded in January 2026 by the engagement team.")
    add_table(doc, ["Ref", "Interviewee", "Role"], [
        ("ST-01", "Hung Vo", "Programmes Director"),
        ("ST-02", "Thao Le", "Operations Director"),
        ("ST-03", "Linh Nguyen", "Monitoring & Evaluation Officer"),
        ("ST-04", "Anh Duc Pham", "Development Director"),
        ("ST-05", "Minh Tran", "Field Officer, Dong Nai hub"),
        ("BN-01", "Mai", "Programme graduate, Ho Chi Minh City"),
        ("BN-02", "Tuan", "Withdrew from programme, Dong Nai"),
        ("BN-03", "Lan", "Graduate with a disability, Ho Chi Minh City"),
        ("BN-04", "Hang", "On the waitlist, Binh Duong"),
        ("DN-01", "Ms. Huong", "Programme Officer, VPBank Foundation"),
        ("EP-01", "Mr. Khang", "HR Manager, gold-tier employer partner"),
        ("EP-02", "Ms. Trang", "Employer partner (rated poorly by youth)"),
        ("VR-01", "Chi Hoa", "Volunteer mentor, three years' service"),
        ("BD-01", "Mr. T.N.", "Member of the GBF Board"),
    ], widths=[0.7, 2.0, 3.9])

    add_heading(doc, "Operational datasets (Workflow B only)", 1)
    add_body(doc, "The monitoring export and data dictionary sit in Workflow B's folder. "
             "Workflow A does not receive these files. If you need a number, a comparison, or a "
             "test of a hypothesis, file an Analysis Request. Headline figures in D-01 and D-12 "
             "are public context; they are not a substitute for querying the operational record.")

    _save(doc, "01_Data_Room_Index", "0_Start_Here", "pdf")


def build_deadlines():
    doc = _doc("Deadlines & Milestones", "Every submission, with owner and time (Indochina Time)",
               ref="MOM-2026-DL-01", issued="Monday 13 July 2026")
    add_body(doc, "All times are 17:00 Indochina Time (ICT, UTC+7) unless stated otherwise. "
             "Submit through the programme portal and to your Buddy. The authoritative, "
             "task-level schedule is the 'Master Timeline' tab of the Shared Toolkit workbook.")

    for week, dates, rows in [
        ("Week 1", "13-17 July", [
            ("Tue 14 Jul", "Project Charter and RACI", "Whole team"),
            ("Wed 15 Jul", "Problem statement and issue tree", "Workflow A"),
            ("Thu 16 Jul", "Stakeholder map", "Workflow A"),
            ("Thu 16 Jul", "Information credibility matrix", "Whole team"),
            ("Fri 17 Jul", "Agree request protocol; Buddy check-in 1", "Whole team"),
        ]),
        ("Week 2", "20-24 July", [
            ("Mon 20 Jul", "First Analysis Request filed (R-01)", "Workflow A"),
            ("Tue 21 Jul", "Data quality report (B internal)", "Workflow B"),
            ("Wed 22 Jul", "Empathy and journey maps; interview synthesis", "Workflow A"),
            ("Thu 23 Jul", "Findings Memo FM-01 returned", "Workflow B"),
            ("Fri 24 Jul", "Dashboard v1; Buddy check-in 2", "Workflow B"),
        ]),
        ("Week 3", "27-31 July", [
            ("Mon 27 Jul", "Analysis Requests R-02 and R-03 filed", "Workflow A"),
            ("Wed 29 Jul", "Findings Memos FM-02/FM-03; analysis summary", "Workflow B"),
            ("Thu 30 Jul", "Strategic options draft (cite Request IDs)", "Workflow A"),
            ("Fri 31 Jul", "Evidence-check memo; Buddy check-in 3", "Workflow B"),
        ]),
        ("Week 4", "3-7 August", [
            ("Wed 5 Aug", "Validation protocol; trade-off reflection", "Workflow A / All"),
            ("Thu 6 Aug", "Funding scenarios; recommendation revision log", "B / All"),
            ("Fri 7 Aug", "Executive summary (draft); Buddy check-in 4", "Workflow A"),
        ]),
        ("Week 5", "10-14 August", [
            ("Tue 11 Aug", "Final report (draft for handoff)", "Whole team"),
            ("Wed 12 Aug 12:00", "Signed Cross-Workflow Handoff Checklist", "Whole team"),
            ("Thu 13 Aug 09:00", "BOARD PRESENTATION", "Whole team"),
            ("Fri 14 Aug", "All final deliverables", "Whole team"),
        ]),
        ("Week 6", "17-21 August", [
            ("Mon 17 Aug", "Peer evaluation (confidential)", "Individual"),
            ("Tue 18 Aug", "Reflection journal", "Individual"),
            ("Thu 20 Aug", "Lessons-learned retrospective", "Whole team"),
        ]),
    ]:
        add_heading(doc, f"{week} ({dates})", 2)
        add_table(doc, ["Due", "Deliverable", "Owner"], rows, widths=[1.7, 3.5, 1.3])

    _save(doc, "02_Deadlines_and_Milestones", "0_Start_Here", "pdf")


# ===========================================================================
# 1. CLIENT DATA ROOM
# ===========================================================================

def build_wsa_start():
    doc = _doc(
        "Workflow A - Start Here",
        "Impact Strategy: what is in this folder and what you must deliver",
        classification="Workflow A",
        ref="MOM-2026-WSA-00",
        issued="Monday 13 July 2026",
    )
    add_body(doc, "You are on Workflow A (Impact Strategy). Read this page first, then the "
             "engagement brief and design-thinking playbook. You set direction and frame the "
             "questions. You do not have access to the operational datasets. Do not ask "
             "facilitators for the file. Route every quantitative need through an Analysis "
             "Request to Workflow B.")
    add_heading(doc, "Files in this folder", 1)
    add_table(doc, ["File", "Purpose"], [
        ("WSA_Engagement_Brief.pdf", "Your mandate, deliverables and constraints"),
        ("WSA_Design_Thinking_Playbook.pdf", "Run sheets for workshops and activities"),
        ("Templates/", "Word files to complete and export to PDF before submission"),
    ], widths=[2.8, 3.6])
    add_body(doc, "Analysis Request and Findings Memo templates live in 4_Shared_Toolkit/Templates/. "
             "Log every request on the Request Log tab of GBF_Consulting_Toolkit.xlsx.")
    add_heading(doc, "Your deliverables", 1)
    add_table(doc, ["Ref", "Deliverable", "Template", "Due (17:00 ICT)"], [
        ("A1", "Problem statement and issue tree", "Problem_Statement_and_Issue_Tree.docx", "Wed 15 Jul"),
        ("A1b", "Stakeholder map", "Stakeholder_Map.docx", "Thu 16 Jul"),
        ("A2", "Interview synthesis", "Interview_Synthesis.docx", "Wed 22 Jul"),
        ("A3", "Empathy and journey maps", "Empathy_and_Journey_Map.docx", "Wed 22 Jul"),
        ("A4", "Strategic options and recommendation", "Strategic_Options.docx", "Thu 30 Jul"),
        ("A5", "Trade-off reflection", "Tradeoff_Reflection.docx", "Wed 5 Aug"),
        ("A6", "Validation protocol", "Validation_Protocol.docx", "Wed 5 Aug"),
        ("A7", "Recommendation revision log", "Recommendation_Revision_Log.docx", "Thu 6 Aug"),
    ], widths=[0.5, 2.4, 2.6, 1.1])
    add_heading(doc, "Before you submit anything", 1)
    add_number(doc, "Every insight from transcripts cites ST-, BN-, DN-, EP- or a client document.")
    add_number(doc, "Every number cites a Findings Memo / Request ID from Workflow B.")
    add_number(doc, "At least three Analysis Requests have Findings Memos returned.")
    add_number(doc, "Workflow B has signed Section B of the Cross-Workflow Handoff Checklist.")
    _save(doc, "WSA_Start_Here", "2_Workflow_A_Impact_Strategy", "pdf")


def build_wsb_start():
    doc = _doc(
        "Workflow B - Start Here",
        "Operations and Analytics: what is in this folder and what you must deliver",
        classification="Workflow B",
        ref="MOM-2026-WSB-00",
        issued="Monday 13 July 2026",
    )
    add_body(doc, "You are on Workflow B (Operations and Analytics). You alone hold the "
             "operational datasets and the data dictionary. Your job is to clean, analyse, and "
             "translate findings so Workflow A can decide. Do not share the raw workbook with A. "
             "Answer their Analysis Requests with Findings Memos in plain language.")
    add_heading(doc, "Files in this folder", 1)
    add_table(doc, ["File", "Purpose"], [
        ("WSB_Engagement_Brief.pdf", "Your mandate, deliverables and constraints"),
        ("WSB_Data_Dictionary.pdf", "Schema and field notes for the export you hold"),
        ("WSB_Data_Analysis_Workbook.pdf", "Guidance for analysis tasks DB-1 to DB-5"),
        ("GBF_Datasets.xlsx", "Operational export (B only - do not share raw)"),
        ("GBF_Performance_Dashboard.xlsx", "Build your working dashboard here"),
        ("Templates/", "Word files to complete and export to PDF before submission"),
    ], widths=[2.8, 3.6])
    add_heading(doc, "Your deliverables", 1)
    add_table(doc, ["Ref", "Deliverable", "Where", "Due (17:00 ICT)"], [
        ("B1", "Data quality report (internal)", "Templates/Data_Quality_Report.docx", "Tue 21 Jul"),
        ("B2", "Performance dashboard (working file)", "GBF_Performance_Dashboard.xlsx", "Fri 24 Jul"),
        ("B3", "Analysis summary + Findings Memos", "Templates/Analysis_Summary.docx + Findings_Memo", "Wed 29 Jul"),
        ("B4", "Evidence-check memo on A's options", "Templates/Evidence_Check_Memo.docx", "Fri 31 Jul"),
        ("B5", "Funding scenarios", "Shared Toolkit / Funding Scenarios tab", "Thu 6 Aug"),
    ], widths=[0.5, 2.2, 2.4, 1.3])
    add_heading(doc, "Before you submit anything", 1)
    add_number(doc, "State the placement definition you used and why.")
    add_number(doc, "Document every cleaning decision in the Assumption Log.")
    add_number(doc, "Findings Memos use plain language; charts are curated, not a dump of pivots.")
    add_number(doc, "Do not claim causation from correlation on attendance or mentoring.")
    add_number(doc, "Workflow A has signed Section A of the Cross-Workflow Handoff Checklist.")
    _save(doc, "WSB_Start_Here", "3_Workflow_B_Operations_and_Analytics", "pdf")


def build_d01_annual_report():
    doc = _doc("Annual Report 2024", "Generation Bridge Foundation (public edition, excerpt)",
               classification="Client document D-01", ref="D-01", issued="Published April 2025")

    add_meta_table(doc, [
        ("Organisation", "Generation Bridge Foundation (Cau Noi The He)"),
        ("Registered office", "47 Nguyen Van Troi, Phu Nhuan District, Ho Chi Minh City"),
        ("Registration", "Decision 124/QD-SLDTBXH, 2017"),
        ("Reporting period", "1 January to 31 December 2024"),
        ("Publication", "April 2025"),
    ])
    add_heading(doc, "A message from our Executive Director", 1)
    add_body(doc, "In a year of rising need, Generation Bridge Foundation stood beside more young "
             "people than ever before. Across four provinces, 2,104 young men and women completed "
             "one of our programmes, and I am proud to report that 71 per cent of them were in "
             "paid work within ninety days - comfortably ahead of our 65 per cent target.")
    add_body(doc, "None of this is possible without our partners. I want to thank the VPBank "
             "Foundation, whose multi-year commitment underpins our flagship Skills Forward "
             "programme, and the hundreds of volunteer mentors who give their evenings and "
             "weekends to walk alongside our young people.")
    add_body(doc, "Demand for our services continues to outstrip what we can offer. As I write, "
             "more than two thousand young people are waiting for a place. Meeting that need, "
             "without compromising the quality that makes our model work, is the central "
             "challenge of the years ahead.")
    add_quote(doc, "We do not simply train young people and hope. We connect them to real "
              "employers, real internships, and real jobs.", "Dr. Lan Phuong Nguyen, Executive Director")

    add_heading(doc, "Our year in numbers", 1)
    add_table(doc, ["Indicator", "2023", "2024"], [
        ("Young people completing a programme", "1,760", "2,104"),
        ("Job placement rate within 90 days", "68%", "71%"),
        ("Active employer partners", "104", "127"),
        ("Active volunteer mentors", "290", "340"),
        ("Provinces served", "3", "4"),
    ], widths=[3.2, 1.5, 1.5])
    add_body(doc, "Placement is defined as any paid employment of at least fifteen hours per "
             "week within ninety days of completing a programme. This includes flexible and "
             "platform-based work such as ride-hailing and online selling.", italic=True, size=10)

    add_heading(doc, "Our programmes", 1)
    add_body(doc, "Skills Forward, our flagship, combines eight weeks of employability training "
             "with an eight-week paid internship at a partner employer. Pathway Digital offers a "
             "twelve-week online course in e-commerce and digital administration, followed by a "
             "four-week in-person lab. Bridge Mentors pairs young people with trained volunteer "
             "mentors for six months of one-to-one support.")

    add_heading(doc, "Looking ahead", 1)
    add_body(doc, "In 2025 and beyond we intend to deepen our employer relationships, strengthen "
             "our monitoring, and explore how we can responsibly reach the thousands of young "
             "people still waiting for a place.")
    _save(doc, "D01_Annual_Report_2024", "1_Client_Data_Room", "pdf")


def build_d02_programme_review():
    doc = _doc("Programme Review 2025", "Internal memorandum - not for external circulation",
               classification="Client document D-02 (internal)", ref="D-02")
    add_meta_table(doc, [
        ("To", "Dr. Lan Phuong Nguyen, Executive Director"),
        ("From", "Hung Vo, Programmes Director"),
        ("Date", "8 December 2025"),
        ("Subject", "Honest review of 2025 delivery ahead of the Board strategy session"),
        ("Classification", "Internal - candid; please do not forward"),
    ])
    add_body(doc, "Lan, ahead of the January strategy session you asked me for a frank picture "
             "rather than the version we put in front of donors. Here it is.")

    add_heading(doc, "1. Dong Nai is in trouble", 1)
    add_body(doc, "The Dong Nai hub is our weakest performer and the gap is widening. Internship "
             "drop-out reached 38 per cent last quarter against an organisation-wide average of "
             "22 per cent. Two causes stand out. First, transport: young people are commuting 25 "
             "to 40 kilometres to industrial-zone placements and paying VND 150,000 to 200,000 a "
             "month in bus fares that we do not subsidise. Second, employer behaviour: several "
             "partners treat interns as free labour, with long hours and no clear duties. I have "
             "not shared the full Dong Nai numbers with VPBank.")

    add_heading(doc, "2. Pathway Digital is being reported too generously", 1)
    add_body(doc, "The completion rate we reported to the donor was 68 per cent. The true "
             "figure, measured as completers divided by those who actually started, is 52 per "
             "cent. The digital team used enrolments as the denominator. We need to correct this "
             "before it becomes a reputational problem.")

    add_heading(doc, "3. The Long An pilot has not worked", 1)
    add_body(doc, "Drop-out in the Long An pilot is around 60 per cent. The catchment is thinner "
             "than we assumed and we do not have enough local employers. I would not expand there "
             "until the model is fixed elsewhere.")

    add_heading(doc, "4. Our people are stretched", 1)
    add_body(doc, "Field staff turnover was 34 per cent this year. The team is doing heroic work, "
             "but it is not sustainable, and every departure costs us employer relationships that "
             "took years to build.")

    add_heading(doc, "Recommendation", 1)
    add_body(doc, "Before we promise anyone growth, we should fix Dong Nai, correct our "
             "reporting, and invest in the quality of our employer partnerships. Scaling a "
             "broken internship experience multiplies harm for the young people least able to "
             "absorb it. I know this is not the story Development wants to tell. It is the true one.")
    add_body(doc, "Hung")
    _save(doc, "D02_Programme_Review_Internal_Memo", "1_Client_Data_Room", "pdf")


def build_d03_grant_agreement():
    doc = _doc("Conditional Grant Agreement", "VPBank Foundation and Generation Bridge Foundation (excerpt)",
               classification="Client document D-03", ref="D-03", issued="Draft, dated 26 September 2025")
    add_body(doc, "This excerpt reproduces the key commercial terms of the conditional renewal "
             "offered by the VPBank Foundation. Legal boilerplate has been omitted.")
    add_meta_table(doc, [
        ("Grantor", "VPBank Foundation"),
        ("Grantee", "Generation Bridge Foundation"),
        ("Grant value", "VND 8,000,000,000 (eight billion dong)"),
        ("Term", "1 January 2026 to 31 December 2028 (three years)"),
        ("Disbursement", "Annual tranches, released on achievement of conditions"),
    ])

    add_heading(doc, "Clause 3 - Conditions of disbursement", 1)
    add_body(doc, "Each annual tranche is conditional on the Grantee demonstrating, to the "
             "reasonable satisfaction of the Grantor, that:")
    add_number(doc, "The job placement rate is at least 65 per cent, measured as the proportion "
               "of programme completers in paid employment of fifteen hours per week or more, "
               "ninety days after completion;")
    add_number(doc, "The Grantee is on a credible trajectory to serve a cumulative total of "
               "4,000 young people by 31 December 2028; and")
    add_number(doc, "The Grantee maintains a monitoring and evaluation framework acceptable to "
               "the Grantor, with quarterly reporting.")

    add_heading(doc, "Clause 4 - Measurement", 1)
    add_body(doc, "The ninety-day window is the contractual standard. The Grantor will accept "
             "flexible and platform-based employment where it is stable and provides at least "
             "fifteen hours of work per week, but requests that formal, contracted employment be "
             "reported separately. Early indications at sixty days may be submitted as "
             "supplementary information only and will not be treated as the primary measure.")

    add_heading(doc, "Clause 6 - Use of funds", 1)
    add_body(doc, "Funds are provided for the expansion of the Grantee's employment programmes, "
             "including the establishment of new delivery capacity. The Grantor supports, but "
             "does not require, expansion into new provinces, provided that quality standards are "
             "maintained.")
    add_body(doc, "Note appended by the GBF Development team: 'This is the largest single "
             "commitment in our history. We should accept.'", italic=True, size=10)
    _save(doc, "D03_VPBank_Grant_Agreement", "1_Client_Data_Room", "pdf")


def build_d08_email_thread():
    doc = _doc("Board Correspondence", "Email thread forwarded to the engagement team",
               classification="Client document D-08 (unverified)", ref="D-08")
    add_body(doc, "The following thread was forwarded to the engagement team by a member of "
             "staff who asked not to be named. It has not been verified. Treat it as an "
             "indication of Board sentiment, not as fact.", italic=True, size=10)

    def email(frm, to, date, subject, body_paras):
        add_meta_table(doc, [("From", frm), ("To", to), ("Date", date), ("Subject", subject)])
        for p in body_paras:
            add_body(doc, p)
        doc.add_paragraph()

    email("T.N. (Board member)", "GBF Board distribution list", "Mon 5 Jan 2026, 21:14",
          "We cannot afford to be timid",
          ["Colleagues, I will say what others are thinking. Lan is too cautious. VPBank is "
           "offering us eight billion dong over three years. If we do not scale, YouthWorks will "
           "eat our lunch in Binh Duong within a year.",
           "I have spoken informally to contacts at the provincial Department of Labour. They "
           "will back us if we commit to Long An. This is our moment. Let us not waste it.",
           "T.N."])
    email("Dr. Lan Phuong Nguyen", "T.N.; GBF Board distribution list", "Tue 6 Jan 2026, 08:02",
          "RE: We cannot afford to be timid",
          ["Thank you, T.N. I share the ambition. My concern is that our Dong Nai hub is not "
           "yet delivering, and scaling a model that is not working simply multiplies the "
           "problem. I have asked for an external review before we commit. I hope the Board will "
           "support that.",
           "Lan"])
    email("Anh Duc Pham (Development Director)", "GBF Board distribution list", "Tue 6 Jan 2026, 09:40",
          "RE: We cannot afford to be timid",
          ["For what it is worth, donors fund growth stories, not pauses. A 'stop and fix' "
           "narrative will cost us momentum with VPBank and others. We should commit publicly and "
           "solve the operational issues in parallel.",
           "Anh Duc"])
    _save(doc, "D08_Board_Email_Thread", "1_Client_Data_Room", "pdf")


def build_d10_competitor():
    doc = _doc("Competitor Brief: YouthWorks", "Prepared by the GBF Development team",
               classification="Client document D-10", ref="D-10", issued="November 2025")
    add_heading(doc, "Summary", 1)
    add_body(doc, "YouthWorks is a government-backed youth training initiative that launched in "
             "Binh Duong in November 2025. It offers a free six-week training course with an "
             "official government certificate. It has strong public funding and a recognised "
             "brand.")
    add_heading(doc, "How it compares to GBF", 1)
    add_table(doc, ["Feature", "YouthWorks", "GBF Skills Forward"], [
        ("Cost to young person", "Free", "Free"),
        ("Duration", "6 weeks", "16 weeks"),
        ("Paid internship", "No", "Yes (8 weeks)"),
        ("Government certificate", "Yes", "No"),
        ("Published placement rate", "Not disclosed", "71% (2024)"),
    ], widths=[2.3, 2.1, 2.1])
    add_heading(doc, "Assessment", 1)
    add_body(doc, "YouthWorks' key weakness is the absence of an internship, which is the part "
             "of our model most associated with real job outcomes. However, 'free and fast and "
             "certified' is attractive, and our field staff estimate that up to 40 per cent of "
             "our current waitlist would switch to YouthWorks if it had the capacity to take them.")
    add_body(doc, "The Development Director's view is that the threat is overstated because "
             "YouthWorks cannot place young people into jobs. Field staff in the border districts "
             "disagree. The truth is probably somewhere in between and worth testing against the "
             "data.", italic=True, size=10)
    _save(doc, "D10_Competitor_Brief_YouthWorks", "1_Client_Data_Room", "pdf")


def build_d11_staff_survey():
    doc = _doc("Staff Engagement Survey 2025", "Summary of results prepared by Human Resources",
               classification="Client document D-11", ref="D-11", issued="November 2025")
    add_heading(doc, "Method", 1)
    add_body(doc, "The annual anonymous survey was open for two weeks in October 2025. Fifty-two "
             "of sixty-seven full-time staff responded, a response rate of 78 per cent.")
    add_heading(doc, "Headline results", 1)
    add_table(doc, ["Statement", "Agree or strongly agree"], [
        ("My workload is sustainable", "39%"),
        ("I have the tools and data I need to do my job well", "44%"),
        ("I would recommend GBF as a place to work", "63%"),
        ("Our monitoring data is accurate", "31%"),
        ("Beneficiaries would benefit most from transport support", "72%"),
    ], widths=[4.2, 2.2])
    add_body(doc, "Field staff turnover during 2025 was 34 per cent, against a sector benchmark "
             "of approximately 22 per cent.")
    add_heading(doc, "Selected verbatim comments", 1)
    add_quote(doc, "I love the mission but I cannot keep working sixty-hour weeks. Something has to give.")
    add_quote(doc, "We are told to visit every employer monthly. With my caseload that is simply impossible.")
    add_quote(doc, "The data system is a mess. I do not trust our own placement numbers.")
    _save(doc, "D11_Staff_Engagement_Survey_2025", "1_Client_Data_Room", "pdf")


def build_d12_financials():
    doc = _doc("Financial Statements 2023-2025", "Audited summary prepared by Finance",
               classification="Client document D-12 (audited)", ref="D-12")
    add_heading(doc, "Statement of income by source", 1)
    add_body(doc, "All figures in VND billion.")
    add_table(doc, ["Source", "2023", "2024", "2025"], [
        ("Corporate foundations and CSR (incl. VPBank)", "6.9", "7.2", "7.6"),
        ("International NGOs and bilateral", "5.4", "5.1", "5.1"),
        ("Individual philanthropy", "1.9", "2.1", "2.2"),
        ("Government co-funding", "1.3", "1.5", "2.0"),
        ("Social enterprise surplus", "1.3", "1.7", "1.3"),
        ("Total income", "16.8", "17.6", "18.2"),
    ], widths=[3.4, 1.0, 1.0, 1.0])
    add_body(doc, "Of 2025 income, 78 per cent was restricted to specific programmes or "
             "provinces and 22 per cent was flexible core funding. The VPBank Foundation was the "
             "single largest funder at approximately 14 per cent of total income.")

    add_heading(doc, "Unit costs", 1)
    add_table(doc, ["Programme", "Cost per completer (VND)"], [
        ("Skills Forward (incl. staff time and internship support)", "12,800,000"),
        ("Pathway Digital", "6,200,000"),
    ], widths=[4.2, 2.2])

    add_heading(doc, "Note on the social enterprise", 1)
    add_body(doc, "GBF Staffing Solutions, established in 2019, places graduates into logistics "
             "and hospitality roles and earns a placement fee from employers. Its surplus "
             "cross-subsidises about 7 per cent of the budget. The auditors note the potential "
             "conflict of interest where the same organisation both advises young people and "
             "earns fees from placing them.")
    _save(doc, "D12_Financial_Statements_2023-2025", "1_Client_Data_Room", "pdf")


def build_interviews():
    doc = _doc("Interview Transcripts", "Field research conducted January 2026",
               classification="Client data room", ref="Transcripts ST/BN/DN/EP/VR/BD")
    add_body(doc, "The following transcripts have been lightly edited for length and clarity. "
             "Speaker names are shown in bold. The interviewer is a member of the engagement "
             "team. Cite these conversations by their reference code, for example (ST-01).",
             italic=True, size=10)

    def transcript(ref, name, role, exchanges):
        add_heading(doc, f"{ref}  -  {name}, {role}", 2)
        for speaker, line in exchanges:
            p = doc.add_paragraph()
            sr = p.add_run(f"{speaker}: ")
            sr.bold = True
            sr.font.size = Pt(10.5)
            br = p.add_run(line)
            br.font.size = Pt(10.5)
        doc.add_paragraph()

    transcript("ST-01", "Hung Vo", "Programmes Director", [
        ("Interviewer", "What worries you most right now?"),
        ("Hung", "Honestly, we are succeeding ourselves into a crisis. There are 2,400 young "
         "people on the waitlist. Families call us in tears. And my field team in Dong Nai is "
         "working sixty-hour weeks because we will not hire until the Board decides on a strategy."),
        ("Interviewer", "Tell me about Dong Nai."),
        ("Hung", "Transport, mostly. The young people commute a long way to industrial zones and "
         "pay 150 to 200 thousand dong a month for buses that we do not cover. Internship "
         "drop-out hit 38 per cent last quarter. Anh Duc wanted to report 'completion' instead of "
         "'internship completion' to make it look better. I refused."),
        ("Interviewer", "Should GBF expand to Long An?"),
        ("Hung", "Not until Dong Nai works. The Long An pilot lost about 60 per cent of its "
         "young people. There is a board member pushing Long An for reasons that have more to do "
         "with his own contacts than with strategy."),
        ("Interviewer", "And Pathway Digital?"),
        ("Hung", "It is cheaper per head, yes. But the real completion rate is 52 per cent, not "
         "the 68 we reported. The digital team counted enrolments, not people who actually "
         "finished. I have asked them to correct it."),
        ("Interviewer", "If you could get one thing from this review?"),
        ("Hung", "Permission to pause expansion and fix quality. Even if a couple of bad "
         "employers complain to the Board."),
    ])

    transcript("ST-02", "Thao Le", "Operations Director", [
        ("Interviewer", "Is the 71 per cent placement rate accurate?"),
        ("Thao", "It depends what you mean by placed. If you include gig work at fifteen hours a "
         "week, yes. If you mean a formal contract, it is closer to 63 per cent. VPBank knows we "
         "use the broader definition."),
        ("Interviewer", "What does it cost to serve one young person?"),
        ("Thao", "Skills Forward is about 12.8 million dong per completer including staff time. "
         "Pathway Digital is about 6.2 million, but its placement rate is lower."),
        ("Interviewer", "Can GBF reach 4,000 young people by 2028?"),
        ("Thao", "Only with the VPBank money and roughly 25 more staff. And recruitment is hard "
         "- we lose people to factory HR jobs and to YouthWorks. Turnover was 34 per cent."),
        ("Interviewer", "The Executive Director has mentioned an intake chatbot."),
        ("Thao", "I think it is a distraction. Our young people need a human voice on the phone, "
         "not a chatbot that cannot understand a Mekong Delta accent."),
    ])

    transcript("ST-03", "Linh Nguyen", "Monitoring & Evaluation Officer", [
        ("Interviewer", "How reliable is your data?"),
        ("Linh", "Frankly, it is a mess. We have duplicate beneficiary IDs, dates in three "
         "different formats, and field staff who mark someone as placed after a single "
         "ride-hailing shift. There are four of us covering 2,800 young people."),
        ("Interviewer", "Anything specific we should check?"),
        ("Linh", "Look at the Dong Nai costs for the third quarter - I think a month was entered "
         "twice. And there is at least one wage figure that is obviously a typo."),
        ("Interviewer", "What about young people with disabilities?"),
        ("Linh", "We report six per cent of our beneficiaries have a disability. The population "
         "figure is nearer nine. And their placement rate is about 49 per cent against 71 overall. "
         "Donors never see that breakdown."),
    ])

    transcript("ST-04", "Anh Duc Pham", "Development Director", [
        ("Interviewer", "How serious is the YouthWorks threat?"),
        ("Anh Duc", "Overstated. They have no internship. Our outcomes are better. Yes, they are "
         "free and we take longer, but young people who want a real job come to us."),
        ("Interviewer", "And the VPBank offer?"),
        ("Anh Duc", "It is the opportunity of a decade. Eight billion dong if we hit 65 per cent "
         "and scale. We should commit publicly. Hesitation kills donor confidence."),
        ("Interviewer", "Some staff have raised reporting placement at sixty days."),
        ("Anh Duc", "That is industry practice. When an outcome is probable, you can measure it "
         "early. M&E is being precious about it."),
    ])

    transcript("ST-05", "Minh Tran", "Field Officer, Dong Nai", [
        ("Interviewer", "Describe your caseload."),
        ("Minh", "There are four of us for about 120 young people. Policy says we visit every "
         "employer monthly. That is simply not possible."),
        ("Interviewer", "You mentioned employer problems."),
        ("Minh", "Two employers on our partner list exploit the interns - ten-hour days, no real "
         "duties. I reported them months ago. Nothing happened."),
        ("Interviewer", "Is YouthWorks affecting you?"),
        ("Minh", "Last month they took about fifteen young people off our waitlist in the border "
         "districts. And I will tell you something we modelled quietly: a transport stipend of "
         "100 thousand dong a month would roughly halve our drop-out."),
    ])

    transcript("BN-01", "Mai", "Programme graduate, Ho Chi Minh City", [
        ("Interviewer", "How did you find the programme?"),
        ("Mai", "I waited five months for a place and almost gave up. The training was good - I "
         "learned Excel and customer service. The internship was far from home, though, about "
         "170 thousand dong a month in bus fares that I paid myself."),
        ("Interviewer", "Did you get a job?"),
        ("Mai", "Yes, at the same logistics company, a real contract, 5.8 million dong a month. "
         "I was lucky. My friend in the Dong Nai group quit because the employer made them work "
         "ten-hour days for nothing."),
        ("Interviewer", "Would Pathway Digital have suited you?"),
        ("Mai", "No. I do not have a laptop and my phone is old. Skills Forward was right for me."),
    ])

    transcript("BN-02", "Tuan", "Withdrew from the programme, Dong Nai", [
        ("Interviewer", "Why did you leave?"),
        ("Tuan", "The training was fine. The internship was miserable - no contract, ten-hour "
         "days, and the boss told us to be grateful. My mother is ill and I could not afford the "
         "buses and lose time at her coffee stall."),
        ("Interviewer", "What would have kept you in?"),
        ("Tuan", "Employers who respect us, and some help with transport. GBF never checked the "
         "workplace before they sent us. People in meetings talk about numbers. Nobody asked me "
         "what the bus cost or what the boss said when I asked for a contract. I still want a "
         "factory job with proper insurance."),
    ])

    transcript("BN-03", "Lan", "Graduate with a disability, Ho Chi Minh City", [
        ("Interviewer", "Was the programme accessible?"),
        ("Lan", "The building had no ramp at first; they moved the class downstairs after I "
         "complained. My mentor was kind but did not know much about disability employment rights."),
        ("Interviewer", "How is the job search?"),
        ("Lan", "Still looking after four months. Employers say they will call and they do not. "
         "GBF counts me as 'in progress'. I feel like a statistic."),
    ])

    transcript("BN-04", "Hang", "On the waitlist, Binh Duong", [
        ("Interviewer", "How long have you been waiting?"),
        ("Hang", "Seven months. YouthWorks offered me a six-week course, but there was no "
         "internship, so I stayed on the GBF list. I will take whichever one gets me a real job "
         "with social insurance."),
    ])

    transcript("DN-01", "Ms. Huong", "Programme Officer, VPBank Foundation", [
        ("Interviewer", "What do you need to see to renew?"),
        ("Huong", "Three things: a placement rate of 65 per cent at ninety days, a credible path "
         "to 4,000 young people by 2028, and a monitoring framework we can trust."),
        ("Interviewer", "Does gig work count?"),
        ("Huong", "If it is stable and at least fifteen hours a week, yes. But we would prefer "
         "formal employment reported separately."),
        ("Interviewer", "What if GBF asked to slow down and fix quality first?"),
        ("Huong", "We would listen. We could consider a smaller, quality-focused grant. But the "
         "Board would need to bring us a credible theory of change, not simply 'give us time'. "
         "Other organisations are competing for this money."),
        ("Interviewer", "And sixty-day reporting?"),
        ("Huong", "Ninety days is the contract. Early figures are supplementary, nothing more."),
    ])

    transcript("EP-01", "Mr. Khang", "HR Manager, gold-tier employer partner", [
        ("Interviewer", "How do the GBF interns perform?"),
        ("Khang", "Mixed. The best are excellent - we hired twelve last year. The weakest lack "
         "basic punctuality. GBF prepares them on soft skills but not on factory discipline."),
        ("Interviewer", "Would you take more?"),
        ("Khang", "If they were better pre-screened. We get too many who want an office job, not "
         "logistics. And, frankly, it is awkward that GBF Staffing Solutions then charges us a "
         "fee to hire them permanently."),
    ])

    transcript("EP-02", "Ms. Trang", "Operations manager, logistics subcontractor (anonymous)", [
        ("Interviewer", "How does your firm use GBF interns?"),
        ("Trang", "Look, I will be straight with you. We are short-staffed in peak season. "
         "The interns help with sorting and packing. We pay the legal minimum stipend but "
         "the work is repetitive and the hours are long. GBF visited once at the start and "
         "we have not seen a caseworker since. I am not proud of it, but if they stopped "
         "sending people we would manage."),
        ("Interviewer", "Would you take more interns?"),
        ("Trang", "Only if someone actually checked that we were treating them properly."),
    ])

    transcript("VR-01", "Chi Hoa", "Volunteer mentor, three years", [
        ("Interviewer", "Is the mentoring model sustainable?"),
        ("Hoa", "Volunteers burn out. I mentor four young people, which is too many. We get a "
         "two-hour orientation and then we are on our own."),
        ("Interviewer", "Could volunteers replace paid field staff?"),
        ("Hoa", "No. Absolutely not. I have dealt with domestic violence, debt, and a young "
         "person at risk of suicide. That needs a professional. Volunteers support the work; we "
         "cannot be the work."),
    ])

    transcript("BD-01", "Mr. T.N.", "Member of the GBF Board", [
        ("Interviewer", "Where do you stand on strategy?"),
        ("T.N.", "We must accept the VPBank money and scale to Long An. I have political support "
         "lined up. If we hesitate, we lose momentum and we lose donors. The 'pause and fix' talk "
         "worries me."),
    ])

    _save(doc, "20_Interview_Transcripts", "1_Client_Data_Room", "pdf")


# ===========================================================================
# 2. WORKSTREAM A
# ===========================================================================

def build_wsa_brief():
    doc = _doc("Workflow A - Impact Strategy", "Engagement brief and requirements",
               classification="Participant material - Workflow A", ref="MOM-2026-WSA-01")
    add_heading(doc, "Purpose of this workflow", 1)
    add_body(doc, "Workflow A owns direction. You make sense of what stakeholders say, frame "
             "the problem, decide which questions matter, develop strategic options, and build "
             "the story the Board will hear. Workflow B holds the operational record. You cannot "
             "look the numbers up yourself; you must ask.")

    add_heading(doc, "Hard constraint on data", 1)
    add_body(doc, "You do not have access to GBF_Datasets.xlsx or the data dictionary. Do not "
             "ask facilitators for those files. When you need a quantitative answer, complete an "
             "Analysis Request (Shared Toolkit). Working out what to ask is part of the exercise. "
             "Vague requests get vague answers.", bold=True)

    add_heading(doc, "Your responsibilities", 1)
    add_bullet(doc, "Synthesise interview transcripts into clear, cited insights.")
    add_bullet(doc, "Frame the problem with a MECE issue tree and testable hypotheses.")
    add_bullet(doc, "File Analysis Requests that turn hypotheses into concrete asks for B.")
    add_bullet(doc, "Ground recommendations in specific voices from the transcripts.")
    add_bullet(doc, "Develop two or three strategic options and a single recommendation.")
    add_bullet(doc, "Sit with the costs of your preferred choice (trade-off reflection).")
    add_bullet(doc, "Own the executive summary and the narrative of the board presentation.")

    add_heading(doc, "Deliverables, owners and deadlines", 1)
    add_table(doc, ["#", "Deliverable", "Template", "Due"], [
        ("A1", "Problem statement and issue tree", "Templates/Problem_Statement_and_Issue_Tree.docx", "Wed 15 Jul"),
        ("A1b", "Stakeholder map", "Templates/Stakeholder_Map.docx", "Thu 16 Jul"),
        ("A2", "Interview synthesis", "Templates/Interview_Synthesis.docx", "Wed 22 Jul"),
        ("A3", "Empathy and journey maps", "Templates/Empathy_and_Journey_Map.docx", "Wed 22 Jul"),
        ("R", "Analysis Requests (minimum three)", "4_Shared_Toolkit/Templates/Analysis_Request.docx", "Weeks 2-4"),
        ("A4", "Strategic options", "Templates/Strategic_Options.docx", "Thu 30 Jul"),
        ("A5", "Trade-off reflection", "Templates/Tradeoff_Reflection.docx", "Wed 5 Aug"),
        ("A6", "Validation protocol (with WF-B)", "Templates/Validation_Protocol.docx", "Wed 5 Aug"),
        ("A7", "Recommendation revision log", "Templates/Recommendation_Revision_Log.docx", "Thu 6 Aug"),
    ], widths=[0.5, 2.4, 2.8, 0.9])

    add_heading(doc, "Detailed requirements", 1)
    add_body(doc, "A1 - Problem statement and issue tree.", bold=True)
    add_bullet(doc, "State the Board's question in one precise sentence of your own.")
    add_bullet(doc, "Build an issue tree that is mutually exclusive and collectively exhaustive.")
    add_bullet(doc, "Write at least four hypotheses. Each hypothesis should imply a question you "
               "could ask Workflow B.")

    add_body(doc, "Analysis Requests.", bold=True)
    add_bullet(doc, "Before final options, complete at least three requests with Findings Memos returned.")
    add_bullet(doc, "Each request states: the decision it informs, the hypothesis, what a useful "
               "answer would look like, and urgency. Do not prescribe field names you have not been given.")
    add_bullet(doc, "Cite Request IDs (R-01, R-02...) on every quantitative claim in A4 and later.")

    add_body(doc, "A2 - Interview synthesis.", bold=True)
    add_bullet(doc, "Summarise by theme, not by person. Flag contradictions between sources.")
    add_bullet(doc, "Every insight cites at least one transcript reference.")

    add_body(doc, "A3 - Empathy and journey maps.", bold=True)
    add_bullet(doc, "Choose one beneficiary archetype grounded in BN-01 to BN-04.")
    add_bullet(doc, "Map the journey from awareness to a job at ninety days. Attach a quote or "
               "a Findings Memo cite to each major pain point.")

    add_body(doc, "A4 - Strategic options.", bold=True)
    add_bullet(doc, "Present two or three genuinely different options.")
    add_bullet(doc, "For each: rationale, cost implication (from B), main risks, who wins and loses.")
    add_bullet(doc, "Recommend one. State what would change your mind.")

    add_body(doc, "A5 - Trade-off reflection.", bold=True)
    add_bullet(doc, "Answer the open prompts in the template honestly. There is no checklist of "
               "named dilemmas to tick off.")
    add_bullet(doc, "Make clear what your recommendation costs - and for whom.")

    add_heading(doc, "How Workflow B will challenge you", 1)
    add_body(doc, "Before you submit, Workflow B will check that every number cites a Request or "
             "Findings ID, that options are financially realistic, and that you have not ignored "
             "uncomfortable findings. You cannot submit without their signature on the handoff "
             "checklist.")
    _save(doc, "WSA_Engagement_Brief", "2_Workflow_A_Impact_Strategy", "pdf")


def build_wsa_playbook():
    doc = _doc("Design-Thinking Playbook", "Facilitated activities for Workflow A",
               classification="Participant material - Workflow A", ref="MOM-2026-WSA-02")
    add_body(doc, "This playbook gives you a run sheet for each design-thinking activity. Every "
             "activity must connect to evidence. Design thinking interprets the evidence; it does "
             "not replace it. When you need a number, write an Analysis Request rather than guessing.")

    for title, when, dur, steps, output in [
        ("Stakeholder mapping", "Week 1", "90 minutes",
         ["Brainstorm every party affected by the decision, one per sticky note.",
          "Cluster them, then plot on a power/interest grid.",
          "Star the five you most need to understand and record why."],
         "A stakeholder map and a shortlist for deeper study."),
        ("Problem framing", "Week 1", "120 minutes",
         ["Each member writes 'the real problem is...' independently.",
          "Share, cluster, and agree a single problem statement.",
          "Build the issue tree and generate hypotheses that imply asks for Workflow B."],
         "A problem statement, issue tree and hypotheses (deliverable A1)."),
        ("First Analysis Request", "Week 2", "45 minutes",
         ["Pick the hypothesis that would most change your recommendation if false.",
          "Write R-01: decision question, hypothesis, what a useful answer looks like.",
          "Hand it to Workflow B and log it on the Request Log."],
         "R-01 filed; waiting for Findings Memo FM-01."),
        ("Empathy mapping", "Week 2", "90 minutes",
         ["Choose one beneficiary archetype grounded in BN-01 to BN-04.",
          "Fill the four quadrants - says, thinks, does, feels - using cited quotes only.",
          "List what you still do not know, and which of those gaps need a request to B."],
         "An evidence-based empathy map (part of deliverable A3)."),
        ("Journey mapping", "Week 2", "120 minutes",
         ["Map the stages from awareness to ninety-day employment.",
          "Mark the emotional highs and lows and the pain points.",
          "Attach a quote or a Findings Memo cite to each pain point."],
         "A current-state journey map with prioritised pain points."),
        ("Ideation - How Might We and Crazy Eights", "Week 3", "120 minutes",
         ["Turn each priority pain point into a 'How might we...' question.",
          "Generate ideas quickly and without judgement, then sketch eight variations of the best.",
          "Dot-vote; check surviving ideas against what the transcripts actually say."],
         "A shortlist of concepts to develop into options."),
        ("Prioritisation with Findings", "Week 3", "60 minutes",
         ["Plot options on an impact-versus-feasibility grid.",
          "Ask Workflow B to challenge placements using Findings Memos, not raw tables.",
          "Record the assumptions behind every placement."],
         "A prioritisation matrix feeding deliverable A4."),
    ]:
        add_heading(doc, title, 2)
        add_meta_table(doc, [("When", when), ("Duration", dur)])
        add_body(doc, "Run sheet:", bold=True, space_after=2)
        for s in steps:
            add_number(doc, s)
        add_body(doc, f"Output: {output}", italic=True, size=10)
    _save(doc, "WSA_Design_Thinking_Playbook", "2_Workflow_A_Impact_Strategy", "pdf")


# ===========================================================================
# 3. WORKSTREAM B
# ===========================================================================

def build_wsb_brief():
    doc = _doc("Workflow B - Operations & Analytics", "Engagement brief and requirements",
               classification="Participant material - Workflow B", ref="MOM-2026-WSB-01")
    add_heading(doc, "Purpose of this workflow", 1)
    add_body(doc, "Workflow B owns the operational evidence. You hold the datasets and the data "
             "dictionary. You clean, analyse, and translate findings into language Workflow A can "
             "use. Success is not a dump of pivots; it is a clear answer to a question A actually asked.")

    add_heading(doc, "Hard constraint on sharing", 1)
    add_body(doc, "Do not share GBF_Datasets.xlsx or the data dictionary with Workflow A. "
             "Respond to Analysis Requests with Findings Memos and curated charts or summary "
             "tables you choose to attach. If A asks for 'the file', refuse and ask what decision "
             "the number needs to support.", bold=True)

    add_heading(doc, "Your responsibilities", 1)
    add_bullet(doc, "Own and apply the data dictionary; profile and clean the export.")
    add_bullet(doc, "Maintain the Request Log; answer Analysis Requests on time.")
    add_bullet(doc, "Translate findings into plain language with caveats and definitions.")
    add_bullet(doc, "Build a working dashboard (B only) and pull curated charts into Findings Memos.")
    add_bullet(doc, "Model funding scenarios behind any recommendation A is considering.")
    add_bullet(doc, "Give an honest assessment of placement under clear definitions.")
    add_bullet(doc, "Maintain the assumption log, risk register and version control.")

    add_heading(doc, "Deliverables, owners and deadlines", 1)
    add_table(doc, ["#", "Deliverable", "Format", "Due"], [
        ("B1", "Data quality report (internal to B)", "Templates/Data_Quality_Report.docx -> PDF", "Tue 21 Jul"),
        ("FM", "Findings Memos (responses to R-01+)", "4_Shared_Toolkit/Templates/Findings_Memo.docx", "Ongoing"),
        ("B2", "Performance dashboard (working file)", "GBF_Performance_Dashboard.xlsx", "Fri 24 Jul"),
        ("B3", "Analysis summary", "Templates/Analysis_Summary.docx -> PDF", "Wed 29 Jul"),
        ("B4", "Evidence-check memo on A's options", "Templates/Evidence_Check_Memo.docx -> PDF", "Fri 31 Jul"),
        ("B5", "Funding scenarios", "Toolkit workbook, 'Funding Scenarios' tab", "Thu 6 Aug"),
    ], widths=[0.5, 2.5, 2.7, 0.9])

    add_heading(doc, "The five analysis tasks", 1)
    add_body(doc, "Complete these as your technical backbone. Prefer to package results as "
             "responses to A's requests where the ask fits. Interpretation matters more than "
             "technical sophistication.")

    add_body(doc, "DB-1 - Data quality audit.", bold=True)
    add_bullet(doc, "Profile every table for missing values, duplicates and format problems.")
    add_bullet(doc, "Use WSB_Data_Dictionary.pdf as your field guide. Investigate anomalies; "
               "do not delete silently.")
    add_bullet(doc, "Record every cleaning decision in the assumption log. Share with A only the "
               "material issues that affect a decision (via Findings Memo), not the full cleaning log.")

    add_body(doc, "DB-2 - Performance dashboard.", bold=True)
    add_bullet(doc, "Use pivot tables to show placement by hub, programme and gender.")
    add_bullet(doc, "Flag any rate below 65 per cent. Title charts with the insight, not the topic.")
    add_bullet(doc, "Export only curated views into Findings Memos when A asks.")

    add_body(doc, "DB-3 - Attendance and outcomes.", bold=True)
    add_bullet(doc, "Join attendance to outcomes; compare means for placed vs not placed.")
    add_bullet(doc, "State that association is not causation; name at least two confounders.")

    add_body(doc, "DB-4 - Mentoring and completion.", bold=True)
    add_bullet(doc, "Total mentor hours per young person, excluding inactive mentors.")
    add_bullet(doc, "Compare completers with those who withdrew. Do not overclaim.")

    add_body(doc, "DB-5 - Funding scenarios.", bold=True)
    add_bullet(doc, "Model hold steady, moderate growth (+30%), aggressive growth (+80%).")
    add_bullet(doc, "Use D-12 unit costs and D-03 VPBank terms.")
    add_bullet(doc, "Judge whether each scenario can meet the 65% condition under an honest "
               "placement definition.")

    add_heading(doc, "How Workflow A will challenge you", 1)
    add_body(doc, "Before you submit, Workflow A will check that you answered their questions "
             "in plain language, that you did not share the raw file, that definitions and "
             "caveats are stated, and that unanswered requests are logged. You cannot submit "
             "without their signature on the handoff checklist.")
    _save(doc, "WSB_Engagement_Brief", "3_Workflow_B_Operations_and_Analytics", "pdf")


def build_wsb_data_dictionary():
    doc = _doc("Data Dictionary", "Operational export held by Workflow B",
               classification="Workflow B only - do not share with Workflow A",
               ref="MOM-2026-WSB-DD")
    add_body(doc, "This dictionary describes the eight tables in GBF_Datasets.xlsx. It is for "
             "Workflow B only. Do not forward this file or the workbook to Workflow A.")

    add_heading(doc, "Tables", 1)
    add_table(doc, ["Table", "Approx. rows", "Purpose"], [
        ("beneficiary_registry", "520", "Master list of young people in the export"),
        ("employment_outcomes", "460", "Employment status ~90 days after completion"),
        ("programme_attendance", "3,700", "Session-level attendance"),
        ("survey_outcomes_2025", "350", "End-of-programme satisfaction survey"),
        ("volunteer_hours", "1,050", "Mentor hours logged"),
        ("hub_costs_2025", "60+", "Monthly operating costs by hub"),
        ("funding_by_source", "48", "Quarterly income by funder"),
        ("geographic_waitlist", "24", "Waitlist by district (Jan 2026 snapshot)"),
    ], widths=[2.2, 1.2, 3.0])

    add_heading(doc, "Key fields", 1)
    add_table(doc, ["Field", "Notes"], [
        ("beneficiary_id", "Primary join key across tables; check for duplicates"),
        ("placed_90d", "Placement flag; coding may be inconsistent across rows"),
        ("employment_type", "Formal vs Gig - definition choice changes the rate"),
        ("province / hub", "Geography labels may not be standardised"),
        ("date fields", "Expect mixed formats"),
        ("wage_vnd_monthly", "Investigate outliers before averaging"),
        ("disability_status", "Missing values exist; treat carefully"),
        ("amount_vnd_m", "Funding amounts; watch for non-numeric cells"),
    ], widths=[2.0, 4.4])

    add_heading(doc, "Known data-quality risks", 1)
    add_body(doc, "The export contains genuine quality problems. Expect duplicates, mixed "
             "coding, format issues, at least one impossible age, at least one wage typo, "
             "possible double-entered costs, and mentor hours logged against inactive status. "
             "Find them, document your handling, and decide what Workflow A needs to know.")

    add_heading(doc, "Sample vs organisation totals", 1)
    add_body(doc, "This export (~520 beneficiaries) is smaller than the 2,847 figure in D-01. "
             "State that limitation whenever you report rates. Do not silently extrapolate.")
    _save(doc, "WSB_Data_Dictionary", "3_Workflow_B_Operations_and_Analytics", "pdf")


def build_wsb_workbook():
    doc = _doc("Data Analysis Workbook", "Step-by-step guidance for Workflow B",
               classification="Participant material - Workflow B", ref="MOM-2026-WSB-02")
    add_body(doc, "This workbook explains how to complete each analysis task in Excel. Package "
             "results as Findings Memos when answering Workflow A's requests. You do not need "
             "to write code. If you use Python or R, the same documentation standards apply.")

    add_heading(doc, "Concepts in plain English", 1)
    add_table(doc, ["Term", "What it means"], [
        ("Pivot table", "A table that groups and summarises your data automatically"),
        ("XLOOKUP", "Finds a value in one table and returns a matching value from another"),
        ("Placement rate", "Placed completers divided by total completers - define both parts"),
        ("Correlation", "Two things move together; it does not prove one causes the other"),
        ("Confounder", "A hidden factor that affects both things you are comparing"),
        ("Selection bias", "Those who attend more may already differ in ways that affect outcomes"),
    ], widths=[1.8, 4.6])

    add_heading(doc, "From request to Findings Memo", 1)
    add_number(doc, "Read A's request. Restate the decision question in your own words.")
    add_number(doc, "Decide which tables and fields answer it. Clean only what you need first.")
    add_number(doc, "Compute the result. Note definitions and sample limitations.")
    add_number(doc, "Write the Findings Memo in plain language. Attach at most two curated charts.")
    add_number(doc, "Update the Request Log status to 'Returned'.")

    add_heading(doc, "Worked approach to DB-3 (attendance and outcomes)", 1)
    add_number(doc, "In programme_attendance, create a helper column that is 1 when attended "
               "and 0 when not, normalising mixed Y/N/1/0 coding.")
    add_number(doc, "Build a pivot for each young person's attendance rate.")
    add_number(doc, "Use XLOOKUP to bring in placement from employment_outcomes.")
    add_number(doc, "Compare average attendance of those placed with those not placed.")
    add_number(doc, "Interpret the gap in two or three sentences and name confounders "
               "(for example transport and motivation).")

    add_heading(doc, "Reporting standards", 1)
    add_bullet(doc, "State the placement definition you used and why.")
    add_bullet(doc, "Never delete an outlier silently - investigate and record the decision.")
    add_bullet(doc, "Give every chart a title that states the finding.")
    add_bullet(doc, "Note the sample-size limitation versus the 2,847 in D-01.")
    add_bullet(doc, "Never attach the raw workbook to a Findings Memo.")
    _save(doc, "WSB_Data_Analysis_Workbook", "3_Workflow_B_Operations_and_Analytics", "pdf")


# ===========================================================================
# 4. SHARED TOOLKIT (guides + templates)
# ===========================================================================

def build_storytelling_guide():
    doc = _doc("Presentation & Storytelling Guide", "How to present to the GBF Board",
               classification="Participant material - shared", ref="MOM-2026-ST-01")
    add_heading(doc, "The one rule", 1)
    add_body(doc, "Every slide answers a single question, and its title states the answer. Build "
             "each slide as: headline (the answer), then evidence, then the insight, then the "
             "implication or recommendation.")
    add_heading(doc, "What to avoid", 1)
    add_table(doc, ["Instead of this", "Do this"], [
        ("A slide titled 'Data analysis'", "A slide titled 'Dong Nai trails every hub on placement'"),
        ("Twelve numbers on one chart", "One chart making one point"),
        ("A decorative stock photo", "White space and one piece of evidence"),
        ("Eight bullet points of recommendations", "One recommendation per slide"),
        ("Ending on 'thank you'", "Ending on the decision you need from the Board"),
    ], widths=[3.2, 3.2])
    add_heading(doc, "A ten-slide structure", 1)
    for i, (h, e) in enumerate([
        ("Your recommendation in one sentence", "The single most important message"),
        ("Demand has outpaced funding", "Waitlist and budget trend"),
        ("Why young people drop out", "A beneficiary quote plus the transport data"),
        ("The real placement picture", "Formal versus broad definition"),
        ("Where performance differs", "Hub or programme comparison"),
        ("The options we weighed", "Two or three, with trade-offs"),
        ("Our recommendation", "The choice and its theory of change"),
        ("How it would be delivered", "A ninety-day roadmap"),
        ("The costs we are accepting", "What your recommendation gives up, and for whom"),
        ("The decision we need today", "Two or three specific asks"),
    ], 1):
        add_bullet(doc, f"Slide {i}: {h} - {e}")
    add_body(doc, "Keep detailed methodology and full financials in an appendix. Cite Request IDs "
             "on any slide that quotes an operational number.")
    _save(doc, "Presentation_and_Storytelling_Guide", "4_Shared_Toolkit", "pdf")


def build_roles_guide():
    doc = _doc("Team Roles & Ways of Working", "Shared reference for the whole team",
               classification="Participant material - shared", ref="MOM-2026-TR-01")
    add_heading(doc, "Suggested roles", 1)
    add_table(doc, ["Role", "Count", "Owns"], [
        ("Engagement Lead", "1", "Timeline, board narrative, final integration, Request Log health"),
        ("Workflow A Lead", "1", "Issue tree, options, trade-off reflection, Analysis Requests"),
        ("Research Lead", "1", "Interview synthesis, credibility matrix, validation"),
        ("Workflow B Lead", "1", "Data dictionary, cleaning, Findings Memos, dashboard"),
        ("Operations Analyst", "1", "Funding scenarios, cost analysis"),
        ("Project & Quality Lead", "1", "Minutes, version control, handoff checklist"),
    ], widths=[1.9, 0.7, 3.8])
    add_body(doc, "In a team of five, combine the Research Lead with the Workflow A Lead.")
    add_heading(doc, "Preventing free-riding", 1)
    add_bullet(doc, "Every deliverable has exactly one accountable owner in the RACI.")
    add_bullet(doc, "Each member speaks at the board presentation.")
    add_bullet(doc, "A must file requests; B must answer them - neither can do the other's job.")
    add_bullet(doc, "The time tracker is reviewed every Friday; peer evaluation is confidential.")
    add_heading(doc, "Decision-making", 1)
    add_body(doc, "Decide by consensus where you can. Where you cannot, the Engagement Lead "
             "decides and the disagreement is recorded in the decision log. Cross-workflow "
             "handoff sign-off is required before anything leaves the team.")
    _save(doc, "Team_Roles_and_Ways_of_Working", "4_Shared_Toolkit", "pdf")


# ---- Templates (delivered as Word) ----

def _template(name, dest, title, builder):
    doc = _doc(title, "Template - complete and export to PDF before submission",
               classification="Participant template")
    builder(doc)
    _save(doc, name, dest, "docx")


def build_templates():
    # Shared
    def charter(doc):
        add_heading(doc, "1. Engagement summary", 2)
        add_meta_table(doc, [
            ("Team name", "________________________________"),
            ("Client", "Generation Bridge Foundation"),
            ("Board question", "What should GBF's strategy be for 2026-2028?"),
            ("Board presentation", "Thursday 13 August 2026, 09:00 ICT"),
        ])
        add_heading(doc, "2. Team members and roles", 2)
        add_table(doc, ["Name", "Role", "Workflow"],
                  [("", "Engagement Lead", ""), ("", "Workflow A Lead", "A"),
                   ("", "Workflow B Lead", "B"), ("", "Research Lead", "A"),
                   ("", "Operations Analyst", "B"), ("", "Project & Quality Lead", "")],
                  widths=[2.6, 2.4, 1.4])
        add_heading(doc, "3. Scope", 2)
        add_bullet(doc, "In scope: strategy recommendation, analysis-request protocol, validation "
                   "against transcripts, board presentation and final report.")
        add_bullet(doc, "Out of scope: legal review, full financial audit, implementation.")
        add_bullet(doc, "Data rule: A does not hold raw datasets; B does not share the raw workbook.")
        add_heading(doc, "4. Ways of working", 2)
        add_table(doc, ["Item", "Agreement"], [
            ("Meeting cadence", ""), ("Communication channel", ""),
            ("Decision rule", "Consensus; escalate to Engagement Lead; log disagreements"),
            ("Analysis requests", "Minimum three with Findings Memos before final options"),
            ("Handoff deadline", "48 hours before any submission"),
        ], widths=[2.2, 4.2])
        add_heading(doc, "5. Sign-off", 2)
        add_body(doc, "All members sign before Tuesday 14 July 2026, 17:00 ICT.")
        for _ in range(3):
            add_signature_block(doc, "________________________", "Team member")
    _template("Project_Charter", "4_Shared_Toolkit", "Project Charter", charter)

    def handoff(doc):
        add_body(doc, "No deliverable may be submitted until both workflows have completed and "
                 "signed their sections. The Engagement Lead signs only if the Request Log shows "
                 "genuine exchange (not a single dump at the end).")
        add_heading(doc, "Section A - Workflow A reviews Workflow B", 2)
        for item in [
            "Findings Memos answer the question asked, in plain language.",
            "The raw workbook and data dictionary were not shared with A.",
            "Placement definition and sample limitations are stated.",
            "Caveats are honest; uncomfortable findings are not hidden.",
            "Unanswered or delayed requests are logged with reasons.",
            "No causation is claimed from correlation.",
            "Charts attached to memos are curated, not a pivot dump.",
        ]:
            add_bullet(doc, "[  ]  " + item)
        add_signature_block(doc, "________________________", "Workflow A reviewer")
        add_heading(doc, "Section B - Workflow B reviews Workflow A", 2)
        for item in [
            "Every quantitative claim cites a Request ID or Findings Memo ID.",
            "At least three Analysis Requests have Findings Memos returned.",
            "Strategic options are financially feasible against D-12 and D-03.",
            "Options changed after at least one B challenge (see revision log).",
            "Journey-map pain points link to recommendations and/or Findings cites.",
            "Trade-off reflection names who gains and who loses.",
            "Implementation roadmap has a realistic timeline and cost.",
        ]:
            add_bullet(doc, "[  ]  " + item)
        add_signature_block(doc, "________________________", "Workflow B reviewer")
        add_heading(doc, "Section C - Joint final sign-off", 2)
        add_body(doc, "Request Log shows minimum three completed R/FM pairs: [  ] Yes")
        add_signature_block(doc, "________________________", "Engagement Lead")
    _template("Cross_Workflow_Handoff_Checklist", "4_Shared_Toolkit",
              "Cross-Workflow Handoff Checklist", handoff)

    def analysis_request(doc):
        add_body(doc, "Workflow A completes one form per ask. Do not prescribe field names you "
                 "have not been given. Working out what to ask is part of the exercise.")
        add_meta_table(doc, [
            ("Request ID (R-xx)", ""),
            ("Date filed", ""),
            ("Filed by", ""),
            ("Urgency (routine / needed for options / blocking)", ""),
        ])
        add_heading(doc, "1. Decision this informs", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "2. Hypothesis or claim to test", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "3. What a useful answer would look like", 2)
        add_body(doc, "(For example: a comparison across hubs; a rate under a stated definition; "
                 "a cost range. Be specific about the decision, not the spreadsheet.)")
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "4. Qualitative sources that prompted this ask", 2)
        add_body(doc, "(Transcript or document refs)")
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "5. Workflow B acknowledgement", 2)
        add_meta_table(doc, [("Received by", ""), ("Target return date", ""),
                             ("Findings Memo ID (FM-xx)", "")])
    _template("Analysis_Request", "4_Shared_Toolkit", "Analysis Request", analysis_request)

    def findings_memo(doc):
        add_body(doc, "Workflow B completes one memo per Analysis Request. Write for a non-analyst. "
                 "Do not attach the raw workbook.")
        add_meta_table(doc, [
            ("Findings Memo ID (FM-xx)", ""),
            ("Responds to Request ID", ""),
            ("Date returned", ""),
            ("Author", ""),
        ])
        add_heading(doc, "1. Question restated", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "2. Method (brief)", 2)
        add_body(doc, "Tables used, cleaning choices that affect this answer, sample note:")
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "3. Finding in plain language", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "4. Definitions used", 2)
        add_body(doc, "(Especially placement / completion)")
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "5. Caveats and what this does not prove", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "6. Attachments", 2)
        add_body(doc, "List curated charts or summary tables only (max two recommended):")
        add_body(doc, "________________________________________________________________")
    _template("Findings_Memo", "4_Shared_Toolkit", "Findings Memo", findings_memo)

    def minutes(doc):
        add_meta_table(doc, [("Meeting", ""), ("Date", ""), ("Time", ""),
                             ("Attendees", ""), ("Absent", ""), ("Note-taker", "")])
        add_heading(doc, "Decisions", 2)
        add_table(doc, ["Decision", "Rationale", "Owner"], [("", "", ""), ("", "", "")],
                  widths=[2.6, 2.6, 1.2])
        add_heading(doc, "Actions", 2)
        add_table(doc, ["Action", "Owner", "Due", "Status"],
                  [("", "", "", ""), ("", "", "", "")], widths=[2.8, 1.4, 1.2, 1.0])
        add_heading(doc, "Open Analysis Requests", 2)
        add_table(doc, ["Request ID", "Status", "Owner"], [("", "", ""), ("", "", "")],
                  widths=[1.4, 2.6, 2.4])
        add_body(doc, "Minutes are circulated within 24 hours; corrections raised within 48 hours.")
    _template("Meeting_Minutes", "4_Shared_Toolkit", "Meeting Minutes", minutes)

    def execsum(doc):
        add_heading(doc, "Key message", 2)
        add_body(doc, "[One sentence: the single most important thing the Board should know.]")
        add_heading(doc, "Context", 2)
        add_body(doc, "[Three or four sentences on why this decision matters now.]")
        add_heading(doc, "Recommendation", 2)
        add_number(doc, "Primary recommendation:")
        add_number(doc, "Supporting actions:")
        add_number(doc, "What GBF should stop doing:")
        add_heading(doc, "Evidence", 2)
        add_table(doc, ["Finding", "Source (D-/ST-/R-/FM-)", "Implication"],
                  [("", "", ""), ("", "", "")], widths=[2.4, 1.8, 2.2])
        add_heading(doc, "Ninety-day roadmap", 2)
        add_table(doc, ["Phase", "Actions", "Cost estimate"],
                  [("0-30 days", "", ""), ("31-60 days", "", ""), ("61-90 days", "", "")],
                  widths=[1.4, 3.4, 1.6])
        add_heading(doc, "What we still do not know", 2)
        add_body(doc, "[Honest limitations - this builds credibility.]")
    _template("Executive_Summary", "4_Shared_Toolkit", "Executive Summary", execsum)

    def finalrep(doc):
        add_body(doc, "Recommended length 15-25 pages excluding the appendix. Use this structure.")
        for sec in [
            "Executive summary", "Engagement approach (including request protocol)",
            "Problem definition", "Situation analysis", "Stakeholder insights",
            "Findings from Analysis Requests", "Strategic options", "Recommendation",
            "Implementation roadmap", "Trade-offs accepted", "Risks and mitigations",
            "Limitations and further work", "Appendices (Request Log extract, key Findings Memos)",
        ]:
            add_number(doc, sec)
        add_heading(doc, "Standards", 2)
        add_bullet(doc, "Every claim carries a source reference (document, transcript, or R/FM ID).")
        add_bullet(doc, "The placement definition is stated wherever the rate is used.")
        add_bullet(doc, "Charts are sourced and dated.")
    _template("Final_Report", "4_Shared_Toolkit", "Final Report", finalrep)

    def peer(doc):
        add_body(doc, "Confidential. Submit individually by Monday 17 August 2026, 17:00 ICT. "
                 "Facilitators use aggregate results; individual ratings remain confidential.")
        add_body(doc, "Rate each teammate from 1 (poor) to 5 (excellent).")
        add_table(doc, ["Team member", "Contribution", "Quality", "Communication",
                        "Reliability", "Collaboration"],
                  [("", "", "", "", "", ""), ("", "", "", "", "", ""),
                   ("", "", "", "", "", ""), ("", "", "", "", "", "")],
                  widths=[1.7, 1.0, 0.9, 1.2, 1.0, 1.1])
        add_heading(doc, "Open questions", 2)
        add_body(doc, "1. Who contributed most to the team's success, and why?")
        add_body(doc, "2. Did the Analysis Request / Findings Memo exchange work in practice?")
        add_body(doc, "3. Anything the facilitators should know (optional)?")
    _template("Peer_Evaluation", "4_Shared_Toolkit", "Peer Evaluation (Confidential)", peer)

    def reflect(doc):
        add_body(doc, "Individual and private. Submit by Tuesday 18 August 2026, 17:00 ICT.")
        for q in ["What did you learn about asking for evidence rather than assuming it?",
                  "What was hardest about working across two workflows with incomplete information?",
                  "Describe one trade-off your team faced and what you decided.",
                  "What did you contribute, and what would you do differently?",
                  "One skill you want to develop next:"]:
            add_heading(doc, q, 2)
            add_body(doc, "________________________________________________________________")
            add_body(doc, "________________________________________________________________")
    _template("Reflection_Journal", "4_Shared_Toolkit", "Reflection Journal (Individual)", reflect)

    # Workflow A templates
    def issue_tree(doc):
        add_heading(doc, "Problem statement", 2)
        add_body(doc, "Restate the Board's question in one precise sentence of your own:")
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "Issue tree (MECE)", 2)
        add_body(doc, "Complete the branches below. Sub-questions at each level must be "
                 "mutually exclusive and collectively exhaustive.")
        add_table(doc, ["Branch", "Sub-questions", "Testable hypothesis"], [
            ("Programmes", "", ""), ("Geography", "", ""), ("Funding", "", ""),
            ("Operations", "", ""), ("Measurement", "", "")], widths=[1.4, 2.6, 2.4])
        add_heading(doc, "Hypotheses to test (each should imply an Analysis Request)", 2)
        for i in range(1, 5):
            add_body(doc, f"H{i}: ________________________________________________")
            add_body(doc, f"    Possible ask to Workflow B: ________________________")
    _template("Problem_Statement_and_Issue_Tree", "2_Workflow_A_Impact_Strategy",
              "Problem Statement and Issue Tree", issue_tree)

    def stakeholder(doc):
        add_body(doc, "Plot stakeholders on a power/interest grid. Star the five you must "
                 "understand most deeply. Cite at least one source per starred stakeholder.")
        add_table(doc, ["Stakeholder", "Power (1-5)", "Interest (1-5)", "Position on strategy", "Source"],
                  [("", "", "", "", "")] * 8, widths=[1.8, 1.0, 1.0, 2.0, 1.6])
    _template("Stakeholder_Map", "2_Workflow_A_Impact_Strategy", "Stakeholder Map", stakeholder)

    def credibility(doc):
        add_body(doc, "Rate each shared source. Justify your rating in one sentence. "
                 "Operational datasets are held by Workflow B; rate Findings Memos once received.")
        add_table(doc, ["Source", "Reliability (H/M/L)", "Why", "Key fact you trust or doubt"],
                  [("D-01 Annual Report", "", "", ""), ("D-02 Internal memo", "", "", ""),
                   ("D-03 VPBank agreement", "", "", ""), ("D-08 Board emails", "", "", ""),
                   ("D-10 YouthWorks brief", "", "", ""), ("D-11 Staff survey", "", "", ""),
                   ("D-12 Financials", "", "", ""), ("Interview transcripts", "", "", ""),
                   ("Findings Memos (from B)", "", "", "")],
                  widths=[1.6, 1.2, 2.0, 2.6])
    _template("Information_Credibility_Matrix", "4_Shared_Toolkit",
              "Information Credibility Matrix", credibility)

    def synth(doc):
        add_body(doc, "Summarise by theme, not by person. Every insight cites a transcript.")
        add_table(doc, ["Theme", "Insight", "Sources", "Contradicts?"],
                  [("", "", "", "")] * 5, widths=[1.5, 2.6, 1.2, 1.1])
        add_heading(doc, "Contradictions to resolve", 2)
        add_body(doc, "For example: is YouthWorks a serious threat? Compare ST-04, ST-05 and D-10.")
        add_heading(doc, "Questions this raises for Workflow B", 2)
        add_body(doc, "________________________________________________________________")
    _template("Interview_Synthesis", "2_Workflow_A_Impact_Strategy",
              "Interview Synthesis", synth)

    def empathy(doc):
        add_heading(doc, "Chosen archetype", 2)
        add_body(doc, "Ground this in a specific voice from the transcripts (cite BN-xx).")
        add_heading(doc, "Empathy map (cite quotes)", 2)
        add_table(doc, ["Says", "Thinks"], [("", "")], widths=[3.2, 3.2])
        add_table(doc, ["Does", "Feels"], [("", "")], widths=[3.2, 3.2])
        add_heading(doc, "Journey map", 2)
        add_table(doc, ["Stage", "What happens", "Emotion", "Pain point", "Evidence (quote or FM-)"], [
            ("Awareness", "", "", "", ""), ("Waitlist", "", "", "", ""),
            ("Training", "", "", "", ""), ("Internship", "", "", "", ""),
            ("Job search", "", "", "", ""), ("90-day outcome", "", "", "", "")],
            widths=[1.1, 1.4, 0.9, 1.2, 1.8])
    _template("Empathy_and_Journey_Map", "2_Workflow_A_Impact_Strategy",
              "Empathy & Journey Map", empathy)

    def options(doc):
        add_body(doc, "Present two or three genuinely different options. Cite Request / Findings "
                 "IDs for every quantitative claim.")
        for letter in ["A", "B", "C"]:
            add_heading(doc, f"Option {letter}", 2)
            add_meta_table(doc, [("Description", ""), ("Rationale", ""),
                                 ("Cost implication (cite FM-/D-12)", ""),
                                 ("Main risks", ""), ("Who wins / who loses", ""),
                                 ("Evidence cites (R-/FM-/ST-/D-)", "")])
        add_heading(doc, "Recommendation", 2)
        add_body(doc, "We recommend Option ____ because ____________________________.")
        add_body(doc, "We would change our mind if ____________________________.")
    _template("Strategic_Options", "2_Workflow_A_Impact_Strategy",
              "Strategic Options", options)

    def tradeoff(doc):
        add_body(doc, "Complete after you have a preferred recommendation. Answer in your own "
                 "words. There is no checklist of named dilemmas to tick.")
        add_heading(doc, "1. Who gains under your recommendation?", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "2. Who loses, waits longer, or is deprioritised?", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "3. What would you need to believe for the opposite choice to be right?", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "4. What cost does your preferred path accept, and why is that acceptable?", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "5. Which transcript voice would be hardest to explain this to - and how "
                    "would you explain it?", 2)
        add_body(doc, "________________________________________________________________")
    _template("Tradeoff_Reflection", "2_Workflow_A_Impact_Strategy",
              "Trade-off Reflection", tradeoff)

    def validation(doc):
        add_body(doc, "Use interview transcripts and Findings Memos. Map each draft recommendation "
                 "to evidence. Minimum six citations: at least two beneficiary (BN-), two staff "
                 "(ST-), one donor (DN-) and one employer (EP-). Include at least two R-/FM- cites.")
        add_table(doc, ["Recommendation", "Source", "What it says", "Supports / challenges", "Action"],
                  [("", "", "", "", "")] * 6, widths=[1.5, 1.0, 1.6, 1.3, 1.0])
    _template("Validation_Protocol", "2_Workflow_A_Impact_Strategy",
              "Validation Protocol", validation)

    def revlog(doc):
        add_body(doc, "Record how validation and Findings Memos changed your thinking. At least "
                 "one substantive revision after a Workflow B challenge is expected.")
        add_table(doc, ["Original element", "What challenge revealed", "Revised?", "New position"],
                  [("", "", "", "")] * 5, widths=[1.8, 2.2, 0.9, 1.5])
    _template("Recommendation_Revision_Log", "2_Workflow_A_Impact_Strategy",
              "Recommendation Revision Log", revlog)

    # Workflow B templates
    def dqr(doc):
        add_body(doc, "Internal to Workflow B. Two pages maximum. Share with A only material "
                 "issues that affect a decision, via Findings Memo.")
        add_heading(doc, "Issues identified", 2)
        add_table(doc, ["#", "Table", "Issue", "Severity", "Action taken"],
                  [("", "", "", "", "")] * 8, widths=[0.4, 1.6, 2.0, 1.0, 1.4])
        add_heading(doc, "Cleaning decisions and rationale", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "Placement definition used", 2)
        add_body(doc, "State which definition you adopted and why.")
        add_heading(doc, "What Workflow A needs to know (for FM translation)", 2)
        add_body(doc, "________________________________________________________________")
    _template("Data_Quality_Report", "3_Workflow_B_Operations_and_Analytics",
              "Data Quality Report", dqr)

    def analysis(doc):
        add_heading(doc, "Headline findings", 2)
        add_table(doc, ["Finding", "Linked Request / FM", "Caveat"], [("", "", "")] * 4,
                  widths=[2.4, 1.8, 2.2])
        add_heading(doc, "Attendance and outcomes (DB-3)", 2)
        add_body(doc, "State the association and at least two confounders. Correlation is not causation.")
        add_heading(doc, "Mentoring and completion (DB-4)", 2)
        add_body(doc, "________________________________________________________________")
    _template("Analysis_Summary", "3_Workflow_B_Operations_and_Analytics",
              "Analysis Summary", analysis)

    def evcheck(doc):
        add_body(doc, "Workflow B checks each Workflow A option against Findings Memos and D-12/D-03.")
        add_table(doc, ["Option / claim", "Supported (cite FM-)?", "Financially feasible?", "Concern"],
                  [("", "", "", "")] * 5, widths=[2.0, 1.6, 1.4, 1.4])
    _template("Evidence_Check_Memo", "3_Workflow_B_Operations_and_Analytics",
              "Evidence-Check Memo", evcheck)


def build_all():
    print("Building realistic documents...")
    build_programme_handbook()
    build_data_room_index()
    build_deadlines()
    build_d01_annual_report()
    build_d02_programme_review()
    build_d03_grant_agreement()
    build_d08_email_thread()
    build_d10_competitor()
    build_d11_staff_survey()
    build_d12_financials()
    build_interviews()
    build_wsa_start()
    build_wsb_start()
    build_wsa_brief()
    build_wsa_playbook()
    build_wsb_brief()
    build_wsb_data_dictionary()
    build_wsb_workbook()
    build_storytelling_guide()
    build_roles_guide()
    build_templates()
    return MANIFEST


if __name__ == "__main__":
    build_all()
