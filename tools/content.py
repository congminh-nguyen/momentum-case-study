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
             "problem, test your thinking against evidence, weigh difficult ethical trade-offs, "
             "and present a recommendation to the Board.")
    add_body(doc, "This is a simulated but realistic engagement. There is no single correct "
             "answer. You will be assessed on the quality of your thinking, the honesty of your "
             "evidence, the discipline of your teamwork, and the clarity of your communication.")
    add_body(doc, "Everything you need is contained in this package. You must not use external "
             "research, the internet, or any information beyond the documents, transcripts and "
             "datasets provided. Where the evidence is incomplete, say so; managing ambiguity is "
             "part of the exercise.", bold=True)

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
             "largest donor and individual board members do not agree with one another, and some "
             "of what you are told will not survive contact with the data.")

    add_heading(doc, "4. How your team is organised: two workstreams", 1)
    add_body(doc, "Every team of five or six splits into two consulting workstreams. Each "
             "workstream has its own detailed engagement brief in its own folder. Read both.")
    add_table(doc, ["Workstream", "Owns", "Folder"], [
        ("A - Impact Strategy", "Stakeholders, design thinking, strategic options, ethics, the recommendation narrative",
         "2_Workstream_A_Impact_Strategy"),
        ("B - Operations & Analytics", "Data cleaning, dashboard, financial and funding analysis, impact measurement",
         "3_Workstream_B_Operations_and_Analytics"),
    ], widths=[1.9, 3.4, 2.4])
    add_body(doc, "The two workstreams are not silos. Each must formally audit the other before "
             "anything is submitted (see the Cross-Workstream Audit Checklist in the Shared "
             "Toolkit). Neither workstream may submit a deliverable without the other's signature.",
             bold=True)

    add_heading(doc, "5. What you will produce", 1)
    add_table(doc, ["Deliverable", "Owner", "Format", "Due"], [
        ("Project Charter & RACI", "Whole team", "Word -> PDF; Excel", "Tue 14 Jul"),
        ("Problem statement & issue tree", "Workstream A", "Word -> PDF", "Wed 15 Jul"),
        ("Stakeholder map", "Workstream A", "Word -> PDF", "Thu 16 Jul"),
        ("Information credibility matrix", "Whole team", "Word -> PDF", "Thu 16 Jul"),
        ("Data quality report", "Workstream B", "Word -> PDF", "Tue 21 Jul"),
        ("Empathy & journey maps", "Workstream A", "Word -> PDF", "Wed 22 Jul"),
        ("Performance dashboard", "Workstream B", "Excel", "Fri 24 Jul"),
        ("Strategic options", "Workstream A", "Word -> PDF", "Thu 30 Jul"),
        ("Funding scenarios", "Workstream B", "Excel", "Thu 6 Aug"),
        ("Validation protocol & revision log", "Whole team", "Word -> PDF", "Thu 6 Aug"),
        ("Final report", "Whole team", "Word -> PDF", "Fri 14 Aug"),
        ("Board presentation", "Whole team", "PowerPoint", "Thu 13 Aug"),
        ("Peer evaluation", "Individual", "Word -> PDF", "Mon 17 Aug"),
        ("Reflection journal", "Individual", "Word -> PDF", "Tue 18 Aug"),
    ], widths=[2.9, 1.5, 1.4, 1.0])
    add_body(doc, "The complete task-by-task timeline, with owners and times, is in the "
             "Shared Toolkit workbook (GBF_Consulting_Toolkit.xlsx, tab 'Master Timeline').")

    add_heading(doc, "6. Programme calendar", 1)
    add_table(doc, ["Week", "Dates (2026)", "Focus"], [
        ("1", "13-17 July", "Frame the problem, set up the team, assess the evidence"),
        ("2", "20-24 July", "Clean the data, map the beneficiary, build the first dashboard"),
        ("3", "27-31 July", "Analyse, generate options, cross-check evidence"),
        ("4", "3-7 August", "Validate against interviews, model funding, revise"),
        ("5", "10-14 August", "Finalise report and deck, present to the Board"),
        ("6", "17-21 August", "Peer evaluation, reflection, debrief"),
    ], widths=[0.8, 1.8, 4.0])

    add_heading(doc, "7. How you will be assessed", 1)
    add_body(doc, "A 100-point rubric applies. The weighting tells you where to invest effort.")
    add_table(doc, ["Criterion", "Points"], [
        ("Problem understanding", "15"), ("Research quality", "10"),
        ("Design thinking", "15"), ("Data analysis", "20"),
        ("Recommendations", "20"), ("Presentation", "10"),
        ("Project management", "5"), ("Teamwork", "5"),
    ], widths=[3.6, 1.2])
    add_body(doc, "Note that data analysis is 20 points, not the majority. Strong teams use data "
             "to support judgement; they do not let it dominate the engagement.")

    add_heading(doc, "8. Ethical dilemmas you must engage with", 1)
    add_body(doc, "Workstream A must address at least three of the following in the ethical "
             "analysis template. There are no perfect answers; we assess how you reason.")
    add_table(doc, ["Ref", "Dilemma"], [
        ("ED-01", "Should GBF prioritise easier-to-place youth to meet the VPBank placement target?"),
        ("ED-02", "Should the organisation report placement at 60 days instead of 90?"),
        ("ED-03", "Can trained volunteers replace paid field staff?"),
        ("ED-04", "Should GBF pilot an AI chatbot for waitlist intake?"),
        ("ED-05", "Should marginal funds go to prevention or intensive NEET support?"),
        ("ED-06", "Should graduates be steered toward placements that earn GBF Staffing Solutions a fee?"),
    ], widths=[0.7, 5.7])

    add_heading(doc, "9. Ground rules", 1)
    add_number(doc, "Use only the materials in this package. Do not invent facts or seek outside data.")
    add_number(doc, "Every claim in a deliverable must cite its source: a document reference "
               "(for example D-02), a transcript reference (for example ST-01), or a named dataset.")
    add_number(doc, "State the definitions you use, especially for the placement rate.")
    add_number(doc, "Deliverables are read as PDF, filled in as Word, analysed in Excel, and "
               "presented in PowerPoint. No other formats.")
    add_number(doc, "Your Buddy is a facilitator, not a consultant. They will ask questions and "
               "offer graduated hints, but they will not give you the answer.")

    add_heading(doc, "10. Package contents", 1)
    add_table(doc, ["Folder", "What is inside"], [
        ("0_Start_Here", "This handbook, the data room index, the deadline summary"),
        ("1_Client_Data_Room", "GBF documents (D-01 to D-12), interview transcripts, the datasets workbook"),
        ("2_Workstream_A_Impact_Strategy", "Workstream A brief, design-thinking playbook, Workstream A templates"),
        ("3_Workstream_B_Operations_and_Analytics", "Workstream B brief, data analysis workbook, Workstream B templates"),
        ("4_Shared_Toolkit", "Consulting toolkit workbook, presentation guide, board deck, shared templates"),
    ], widths=[2.9, 3.6])

    _save(doc, "00_Programme_Handbook", "0_Start_Here", "pdf")


def build_data_room_index():
    doc = _doc("Data Room Index", "A guide to every document, transcript and dataset in the package",
               ref="MOM-2026-DRI-01", issued="Monday 13 July 2026")
    add_body(doc, "This index lists everything in the client data room and where to find it. "
             "The reliability column is our honest assessment; you should form your own view and "
             "record it in your information credibility matrix.")

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

    add_heading(doc, "Datasets", 1)
    add_body(doc, "All eight tables are in a single Excel workbook, GBF_Datasets.xlsx, one tab "
             "per table, with a data dictionary on the first tab. The data is a real export from "
             "GBF's monitoring system and contains genuine quality problems that you are expected "
             "to find and handle.")
    add_table(doc, ["Table", "Rows (approx.)", "Purpose"], [
        ("beneficiary_registry", "520", "Master list of young people served"),
        ("employment_outcomes", "460", "Employment status 90 days after completion"),
        ("programme_attendance", "3,700", "Session-level attendance"),
        ("survey_outcomes_2025", "350", "End-of-programme satisfaction survey"),
        ("volunteer_hours", "1,050", "Mentor hours logged"),
        ("hub_costs_2025", "36", "Monthly operating costs by hub"),
        ("funding_by_source", "48", "Quarterly income by funder"),
        ("geographic_waitlist", "24", "Waitlist by district"),
    ], widths=[2.4, 1.4, 2.8])
    add_body(doc, "A note on scale: the export covers the 2024-2025 cohorts and is smaller than "
             "the 2,847 figure quoted in the Annual Report. Treat it as a representative sample "
             "and be explicit about that limitation in your analysis.")

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
            ("Wed 15 Jul", "Problem statement and issue tree", "Workstream A"),
            ("Thu 16 Jul", "Stakeholder map", "Workstream A"),
            ("Thu 16 Jul", "Information credibility matrix", "Whole team"),
            ("Fri 17 Jul", "Populated Gantt; Buddy check-in 1", "Whole team"),
        ]),
        ("Week 2", "20-24 July", [
            ("Tue 21 Jul", "Data quality report", "Workstream B"),
            ("Wed 22 Jul", "Empathy and journey maps", "Workstream A"),
            ("Wed 22 Jul", "Interview synthesis", "Workstream A"),
            ("Fri 24 Jul", "Dashboard v1; Buddy check-in 2", "Workstream B"),
        ]),
        ("Week 3", "27-31 July", [
            ("Wed 29 Jul", "Dashboard v2 and analysis summary", "Workstream B"),
            ("Thu 30 Jul", "Strategic options (draft)", "Workstream A"),
            ("Fri 31 Jul", "Evidence-check memo; Buddy check-in 3", "Workstream B"),
        ]),
        ("Week 4", "3-7 August", [
            ("Wed 5 Aug", "Validation protocol (completed)", "Whole team"),
            ("Thu 6 Aug", "Funding scenarios", "Workstream B"),
            ("Thu 6 Aug", "Recommendation revision log", "Whole team"),
            ("Fri 7 Aug", "Executive summary (draft); Buddy check-in 4", "Workstream A"),
        ]),
        ("Week 5", "10-14 August", [
            ("Tue 11 Aug", "Final report (draft for audit)", "Whole team"),
            ("Wed 12 Aug 12:00", "Signed cross-workstream audit checklist", "Whole team"),
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
        "Workstream A - Start Here",
        "Impact Strategy: what is in this folder and what you must deliver",
        classification="Workstream A",
        ref="MOM-2026-WSA-00",
        issued="Monday 13 July 2026",
    )
    add_body(doc, "You are on Workstream A (Impact Strategy). Read this page first, then the "
             "engagement brief and design-thinking playbook in this folder. Workstream B handles "
             "the data; you own the story, the options and the ethics. You cannot submit anything "
             "without Workstream B's signed audit on the cross-workstream checklist.")
    add_heading(doc, "Files in this folder", 1)
    add_table(doc, ["File", "Purpose"], [
        ("WSA_Engagement_Brief.pdf", "Your mandate, deliverables and requirements"),
        ("WSA_Design_Thinking_Playbook.pdf", "Run sheets for workshops and activities"),
        ("Templates/", "Word files to complete and export to PDF before submission"),
    ], widths=[2.8, 3.6])
    add_heading(doc, "Your deliverables", 1)
    add_table(doc, ["Ref", "Deliverable", "Template", "Due (17:00 ICT)"], [
        ("A1", "Problem statement and issue tree", "Problem_Statement_and_Issue_Tree.docx", "Wed 15 Jul"),
        ("A1b", "Stakeholder map", "Stakeholder_Map.docx", "Thu 16 Jul"),
        ("A2", "Interview synthesis", "Interview_Synthesis.docx", "Wed 22 Jul"),
        ("A3", "Empathy and journey maps", "Empathy_and_Journey_Map.docx", "Wed 22 Jul"),
        ("A4", "Strategic options and recommendation", "Strategic_Options.docx", "Thu 30 Jul"),
        ("A5", "Ethical analysis (minimum three dilemmas)", "Ethical_Analysis.docx", "Wed 5 Aug"),
        ("A6", "Validation protocol", "Validation_Protocol.docx", "Wed 5 Aug"),
        ("A7", "Recommendation revision log", "Recommendation_Revision_Log.docx", "Thu 6 Aug"),
    ], widths=[0.5, 2.4, 2.6, 1.1])
    add_heading(doc, "Before you submit anything", 1)
    add_number(doc, "Every insight cites a transcript (ST-, BN-, DN-, EP-) or a client document (D-01 to D-12).")
    add_number(doc, "Workstream B has signed Section B of the Cross-Workstream Audit Checklist.")
    add_number(doc, "You have exported the Word template to PDF and filed it in the team drive.")
    _save(doc, "WSA_Start_Here", "2_Workstream_A_Impact_Strategy", "pdf")


def build_wsb_start():
    doc = _doc(
        "Workstream B - Start Here",
        "Operations and Analytics: what is in this folder and what you must deliver",
        classification="Workstream B",
        ref="MOM-2026-WSB-00",
        issued="Monday 13 July 2026",
    )
    add_body(doc, "You are on Workstream B (Operations and Analytics). Read this page first, then "
             "the engagement brief and data analysis workbook. You turn the messy export into "
             "findings the team can trust. Workstream A owns the recommendation; you own the "
             "evidence base and keep the numbers honest.")
    add_heading(doc, "Files in this folder", 1)
    add_table(doc, ["File", "Purpose"], [
        ("WSB_Engagement_Brief.pdf", "Your mandate, deliverables and requirements"),
        ("WSB_Data_Analysis_Workbook.pdf", "Step-by-step guidance for tasks DB-1 to DB-5"),
        ("GBF_Performance_Dashboard.xlsx", "Build your dashboard and analyses here"),
        ("Templates/", "Word files to complete and export to PDF before submission"),
    ], widths=[2.8, 3.6])
    add_body(doc, "The raw data tables are in the Client Data Room: "
             "1_Client_Data_Room/GBF_Datasets.xlsx (one tab per table).")
    add_heading(doc, "Your deliverables", 1)
    add_table(doc, ["Ref", "Deliverable", "Where", "Due (17:00 ICT)"], [
        ("B1", "Data quality report", "Templates/Data_Quality_Report.docx", "Tue 21 Jul"),
        ("B2", "Performance dashboard", "GBF_Performance_Dashboard.xlsx", "Fri 24 Jul"),
        ("B3", "Analysis summary", "Templates/Analysis_Summary.docx", "Wed 29 Jul"),
        ("B4", "Evidence-check memo on WS-A options", "Templates/Evidence_Check_Memo.docx", "Fri 31 Jul"),
        ("B5", "Funding scenarios", "Shared Toolkit / Funding Scenarios tab", "Thu 6 Aug"),
    ], widths=[0.5, 2.2, 2.4, 1.3])
    add_heading(doc, "Before you submit anything", 1)
    add_number(doc, "State the placement definition you used and why (compare D-01 footnote with ST-02).")
    add_number(doc, "Document every cleaning decision in the Assumption Log (Shared Toolkit workbook).")
    add_number(doc, "Do not claim causation from correlation on attendance or mentoring.")
    add_number(doc, "Workstream A has signed Section A of the Cross-Workstream Audit Checklist.")
    _save(doc, "WSB_Start_Here", "3_Workstream_B_Operations_and_Analytics", "pdf")


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
             "reporting, and invest in the quality of our employer partnerships. I know this is "
             "not the story Development wants to tell. It is the true one.")
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
         "workplace before they sent us. I still want a factory job with proper insurance."),
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
    doc = _doc("Workstream A - Impact Strategy", "Engagement brief and requirements",
               classification="Participant material - Workstream A", ref="MOM-2026-WSA-01")
    add_heading(doc, "Purpose of this workstream", 1)
    add_body(doc, "Workstream A owns the human and strategic side of the engagement. You make "
             "sense of what stakeholders need, translate that into strategic options, weigh the "
             "ethics, and build the story the Board will hear. Workstream B supplies the numbers; "
             "you supply the meaning, and you challenge each other.")

    add_heading(doc, "Your responsibilities", 1)
    add_bullet(doc, "Synthesise the interview transcripts into clear, cited insights.")
    add_bullet(doc, "Frame the problem with a MECE issue tree and testable hypotheses.")
    add_bullet(doc, "Run the design-thinking activities (empathy map, journey map, ideation).")
    add_bullet(doc, "Develop two or three strategic options and a single recommendation.")
    add_bullet(doc, "Lead the ethical analysis.")
    add_bullet(doc, "Own the executive summary and the narrative of the board presentation.")

    add_heading(doc, "Deliverables, owners and deadlines", 1)
    add_table(doc, ["#", "Deliverable", "Template", "Due"], [
        ("A1", "Problem statement and issue tree", "Templates/Problem_Statement_and_Issue_Tree.docx", "Wed 15 Jul"),
        ("A1b", "Stakeholder map", "Templates/Stakeholder_Map.docx", "Thu 16 Jul"),
        ("A2", "Interview synthesis", "Templates/Interview_Synthesis.docx", "Wed 22 Jul"),
        ("A3", "Empathy and journey maps", "Templates/Empathy_and_Journey_Map.docx", "Wed 22 Jul"),
        ("A4", "Strategic options", "Templates/Strategic_Options.docx", "Thu 30 Jul"),
        ("A5", "Ethical analysis", "Templates/Ethical_Analysis.docx", "Wed 5 Aug"),
        ("A6", "Validation protocol (with WS-B)", "Templates/Validation_Protocol.docx", "Wed 5 Aug"),
        ("A7", "Recommendation revision log", "Templates/Recommendation_Revision_Log.docx", "Thu 6 Aug"),
    ], widths=[0.5, 2.6, 2.6, 0.9])

    add_heading(doc, "Detailed requirements", 1)
    add_body(doc, "A1 - Problem statement and issue tree.", bold=True)
    add_bullet(doc, "State the Board's question in one precise sentence of your own.")
    add_bullet(doc, "Build an issue tree that is mutually exclusive and collectively exhaustive, "
               "with branches covering programmes, geography, funding, operations and measurement.")
    add_bullet(doc, "Write at least four hypotheses you can test against the data room.")

    add_body(doc, "A1b - Stakeholder map.", bold=True)
    add_bullet(doc, "Plot every party affected by the decision on a power/interest grid.")
    add_bullet(doc, "Star the five stakeholders you must understand most deeply and cite at least "
               "one source per starred stakeholder.")

    add_body(doc, "A2 - Interview synthesis.", bold=True)
    add_bullet(doc, "Summarise findings by theme, not by person. Do not simply quote at length.")
    add_bullet(doc, "Every insight must cite at least one transcript reference.")
    add_bullet(doc, "Flag where sources contradict one another (for example ST-04 versus ST-05 "
               "and D-10 on the YouthWorks threat).")

    add_body(doc, "A3 - Empathy and journey maps.", bold=True)
    add_bullet(doc, "Choose one beneficiary archetype (for example the Dong Nai intern who "
               "withdraws) and build an evidence-based empathy map using BN-01 to BN-04.")
    add_bullet(doc, "Map the journey from awareness to a job at ninety days and mark the pain "
               "points, quantifying them with data where you can.")

    add_body(doc, "A4 - Strategic options.", bold=True)
    add_bullet(doc, "Present two or three genuinely different options, described so that a "
               "reasonable person could choose any of them.")
    add_bullet(doc, "For each option give the rationale, the cost implication, the main risks, "
               "and who wins and loses.")
    add_bullet(doc, "Recommend one, and be explicit about what would change your mind.")

    add_body(doc, "A5 - Ethical analysis.", bold=True)
    add_bullet(doc, "Address at least three of the six ethical dilemmas set out in the handbook.")
    add_bullet(doc, "For each, state the tension, the stakeholders affected, your position, and "
               "the residual risk you are accepting.")

    add_heading(doc, "How Workstream B will audit you", 1)
    add_body(doc, "Before you submit, Workstream B will check that every claim you make is "
             "supported by data or a cited source, that your options are financially realistic, "
             "and that your ethical analysis engages with the numbers rather than avoiding them. "
             "You cannot submit without their signature on the audit checklist.")
    _save(doc, "WSA_Engagement_Brief", "2_Workstream_A_Impact_Strategy", "pdf")


def build_wsa_playbook():
    doc = _doc("Design-Thinking Playbook", "Facilitated activities for Workstream A",
               classification="Participant material - Workstream A", ref="MOM-2026-WSA-02")
    add_body(doc, "This playbook gives you a run sheet for each design-thinking activity. Every "
             "activity must connect to evidence in the data room; design thinking is a way to "
             "interpret the evidence, not a substitute for it.")

    for title, when, dur, steps, output in [
        ("Stakeholder mapping", "Week 1", "90 minutes",
         ["Brainstorm every party affected by the decision, one per sticky note.",
          "Cluster them, then plot on a power/interest grid.",
          "Star the five you most need to understand and record why."],
         "A stakeholder map and a shortlist for deeper study."),
        ("Problem framing", "Week 1", "120 minutes",
         ["Each member writes 'the real problem is...' independently.",
          "Share, cluster, and agree a single problem statement.",
          "Build the issue tree and generate hypotheses."],
         "A problem statement, issue tree and hypotheses (deliverable A1)."),
        ("Empathy mapping", "Week 2", "90 minutes",
         ["Choose one beneficiary archetype grounded in BN-01 to BN-04.",
          "Fill the four quadrants - says, thinks, does, feels - using cited quotes only.",
          "List what you still do not know."],
         "An evidence-based empathy map (part of deliverable A3)."),
        ("Journey mapping", "Week 2", "120 minutes",
         ["Map the stages from awareness to ninety-day employment.",
          "Mark the emotional highs and lows and the pain points.",
          "Attach a data point or quote to each pain point."],
         "A current-state journey map with prioritised pain points."),
        ("Ideation - How Might We and Crazy Eights", "Week 3", "120 minutes",
         ["Turn each priority pain point into a 'How might we...' question.",
          "Generate ideas quickly and without judgement, then sketch eight variations of the best.",
          "Dot-vote and check each surviving idea against GBF's values."],
         "A shortlist of concepts to develop into options."),
        ("Prioritisation", "Week 3", "60 minutes",
         ["Plot options on an impact-versus-feasibility grid.",
          "Let Workstream B challenge the placement of each with data.",
          "Record the assumptions behind every placement."],
         "A prioritisation matrix feeding deliverable A4."),
    ]:
        add_heading(doc, title, 2)
        add_meta_table(doc, [("When", when), ("Duration", dur)])
        add_body(doc, "Run sheet:", bold=True, space_after=2)
        for s in steps:
            add_number(doc, s)
        add_body(doc, f"Output: {output}", italic=True, size=10)
    _save(doc, "WSA_Design_Thinking_Playbook", "2_Workstream_A_Impact_Strategy", "pdf")


# ===========================================================================
# 3. WORKSTREAM B
# ===========================================================================

def build_wsb_brief():
    doc = _doc("Workstream B - Operations & Analytics", "Engagement brief and requirements",
               classification="Participant material - Workstream B", ref="MOM-2026-WSB-01")
    add_heading(doc, "Purpose of this workstream", 1)
    add_body(doc, "Workstream B owns the evidence base. You turn a messy data export into "
             "trustworthy findings, model the money, and give the team an honest picture of what "
             "is working and what is not. You also keep the team disciplined: assumptions, risks "
             "and version control run through you.")

    add_heading(doc, "Your responsibilities", 1)
    add_bullet(doc, "Profile and clean the datasets, documenting every decision.")
    add_bullet(doc, "Build a clear performance dashboard.")
    add_bullet(doc, "Analyse the relationships between attendance, mentoring and outcomes.")
    add_bullet(doc, "Model the funding scenarios behind any recommendation.")
    add_bullet(doc, "Give an honest assessment of the true placement rate.")
    add_bullet(doc, "Maintain the assumption log, risk register and version control.")

    add_heading(doc, "Deliverables, owners and deadlines", 1)
    add_table(doc, ["#", "Deliverable", "Format", "Due"], [
        ("B1", "Data quality report", "Templates/Data_Quality_Report.docx -> PDF", "Tue 21 Jul"),
        ("B2", "Performance dashboard", "GBF_Performance_Dashboard.xlsx", "Fri 24 Jul"),
        ("B3", "Analysis summary", "Templates/Analysis_Summary.docx -> PDF", "Wed 29 Jul"),
        ("B4", "Evidence-check memo on WS-A options", "Templates/Evidence_Check_Memo.docx -> PDF", "Fri 31 Jul"),
        ("B5", "Funding scenarios", "Toolkit workbook, 'Funding Scenarios' tab", "Thu 6 Aug"),
    ], widths=[0.5, 2.5, 2.7, 0.9])

    add_heading(doc, "The five analysis tasks", 1)
    add_body(doc, "Data analysis is worth 20 points. Interpretation matters more than technical "
             "sophistication; explain every finding in plain English and state its limitations.")

    add_body(doc, "DB-1 - Data quality audit.", bold=True)
    add_bullet(doc, "Profile every table for missing values, duplicates and format problems.")
    add_bullet(doc, "You should expect to find, among other things: duplicate beneficiary IDs; "
               "dates in more than one format; the placement field coded inconsistently (Y, Yes, "
               "1, N, No, 0, blank); at least one impossible age; at least one wage typo; a "
               "double-entered month of Dong Nai costs; and hours logged against an inactive mentor.")
    add_bullet(doc, "Record every cleaning decision and its rationale in the assumption log.")

    add_body(doc, "DB-2 - Performance dashboard.", bold=True)
    add_bullet(doc, "Use pivot tables to show the placement rate by hub, by programme and by gender.")
    add_bullet(doc, "Apply conditional formatting to flag any rate below 65 per cent.")
    add_bullet(doc, "Produce three charts, each with a title that states the insight, not the topic.")

    add_body(doc, "DB-3 - Attendance and outcomes.", bold=True)
    add_bullet(doc, "Join attendance to outcomes on beneficiary_id using XLOOKUP.")
    add_bullet(doc, "Compare mean attendance for those placed and not placed.")
    add_bullet(doc, "State clearly that association is not causation, and name at least two "
               "confounding factors (for example motivation and transport).")

    add_body(doc, "DB-4 - Mentoring and completion.", bold=True)
    add_bullet(doc, "Use SUMIFS to total mentor hours per young person, excluding inactive mentors.")
    add_bullet(doc, "Compare completers with those who withdrew.")

    add_body(doc, "DB-5 - Funding scenarios.", bold=True)
    add_bullet(doc, "Model three scenarios: hold steady, moderate growth (+30 per cent), and "
               "aggressive growth (+80 per cent).")
    add_bullet(doc, "Use the audited unit costs from D-12 and the VPBank terms from D-03.")
    add_bullet(doc, "Judge whether each scenario can meet the 65 per cent condition honestly, "
               "without selecting easier-to-place young people.")

    add_heading(doc, "How Workstream A will audit you", 1)
    add_body(doc, "Before you submit, Workstream A will check that your cleaning choices do not "
             "quietly hide uncomfortable truths, that your chart titles match what the interviews "
             "say, and that you have not claimed causation from correlation. You cannot submit "
             "without their signature on the audit checklist.")
    _save(doc, "WSB_Engagement_Brief", "3_Workstream_B_Operations_and_Analytics", "pdf")


def build_wsb_workbook():
    doc = _doc("Data Analysis Workbook", "Step-by-step guidance for Workstream B",
               classification="Participant material - Workstream B", ref="MOM-2026-WSB-02")
    add_body(doc, "This workbook explains, in plain language, how to complete each analysis task "
             "in Excel. You do not need to write any code. If your team chooses to use Python or "
             "R, the same standards of documentation and interpretation apply.")

    add_heading(doc, "Concepts in plain English", 1)
    add_table(doc, ["Term", "What it means"], [
        ("Pivot table", "A table that groups and summarises your data automatically"),
        ("XLOOKUP", "Finds a value in one table and returns a matching value from another"),
        ("Placement rate", "Placed completers divided by total completers - define both parts"),
        ("Correlation", "Two things move together; it does not prove one causes the other"),
        ("Confounder", "A hidden factor that affects both things you are comparing"),
        ("Selection bias", "Those who attend more may already differ in ways that affect outcomes"),
    ], widths=[1.8, 4.6])

    add_heading(doc, "Worked approach to DB-3 (attendance and outcomes)", 1)
    add_number(doc, "In programme_attendance, create a helper column that is 1 when the young "
               "person attended and 0 when they did not, normalising the mixed Y/N/1/0 coding.")
    add_number(doc, "Build a pivot table to get each young person's attendance rate: sessions "
               "attended divided by sessions offered.")
    add_number(doc, "Use XLOOKUP to bring in the placement outcome from employment_outcomes.")
    add_number(doc, "Compare the average attendance of those placed with those not placed.")
    add_number(doc, "Write two or three sentences interpreting the gap, and name the confounders. "
               "For example: those facing transport barriers may miss both sessions and "
               "internships, so lower attendance may be a symptom, not a cause.")

    add_heading(doc, "Reporting standards", 1)
    add_bullet(doc, "State the placement definition you used and why.")
    add_bullet(doc, "Never delete an outlier silently - investigate it and record the decision.")
    add_bullet(doc, "Give every chart a title that states the finding.")
    add_bullet(doc, "Note the sample-size limitation: this export is smaller than the full 2,847.")
    _save(doc, "WSB_Data_Analysis_Workbook", "3_Workstream_B_Operations_and_Analytics", "pdf")


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
        ("The ethics we are not ducking", "One dilemma, resolved honestly"),
        ("The decision we need today", "Two or three specific asks"),
    ], 1):
        add_bullet(doc, f"Slide {i}: {h} - {e}")
    add_body(doc, "Keep detailed data, methodology and full financials in an appendix.")
    _save(doc, "Presentation_and_Storytelling_Guide", "4_Shared_Toolkit", "pdf")


def build_roles_guide():
    doc = _doc("Team Roles & Ways of Working", "Shared reference for the whole team",
               classification="Participant material - shared", ref="MOM-2026-TR-01")
    add_heading(doc, "Suggested roles", 1)
    add_table(doc, ["Role", "Count", "Owns"], [
        ("Engagement Lead", "1", "Timeline, board narrative, final integration"),
        ("Workstream A Lead", "1", "Issue tree, options, ethics, empathy and journey maps"),
        ("Research Lead", "1", "Interview synthesis, credibility matrix, validation"),
        ("Workstream B Lead", "1", "Data cleaning, dashboard, analysis"),
        ("Operations Analyst", "1", "Funding scenarios, cost analysis"),
        ("Project & Quality Lead", "1", "Minutes, version control, audit coordination"),
    ], widths=[1.9, 0.7, 3.8])
    add_body(doc, "In a team of five, combine the Research Lead with the Workstream A Lead.")
    add_heading(doc, "Preventing free-riding", 1)
    add_bullet(doc, "Every deliverable has exactly one accountable owner in the RACI.")
    add_bullet(doc, "Each member speaks at the board presentation.")
    add_bullet(doc, "Everyone touches both qualitative and quantitative work.")
    add_bullet(doc, "The time tracker is reviewed every Friday; peer evaluation is confidential.")
    add_heading(doc, "Decision-making", 1)
    add_body(doc, "Decide by consensus where you can. Where you cannot, the Engagement Lead "
             "decides and the disagreement is recorded in the decision log. Cross-workstream "
             "sign-off is required before anything leaves the team.")
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
        add_table(doc, ["Name", "Role", "Workstream"],
                  [("", "Engagement Lead", ""), ("", "Workstream A Lead", "A"),
                   ("", "Workstream B Lead", "B"), ("", "Research Lead", "A"),
                   ("", "Operations Analyst", "B"), ("", "Project & Quality Lead", "")],
                  widths=[2.6, 2.4, 1.4])
        add_heading(doc, "3. Scope", 2)
        add_bullet(doc, "In scope: strategy recommendation, data analysis, validation against "
                   "transcripts, ethical analysis, board presentation and final report.")
        add_bullet(doc, "Out of scope: legal review, full financial audit, implementation.")
        add_heading(doc, "4. Ways of working", 2)
        add_table(doc, ["Item", "Agreement"], [
            ("Meeting cadence", ""), ("Communication channel", ""),
            ("Decision rule", "Consensus; escalate to Engagement Lead; log disagreements"),
            ("Cross-audit deadline", "48 hours before any submission"),
        ], widths=[2.2, 4.2])
        add_heading(doc, "5. Sign-off", 2)
        add_body(doc, "All members sign before Tuesday 14 July 2026, 17:00 ICT.")
        for _ in range(3):
            add_signature_block(doc, "________________________", "Team member")
    _template("Project_Charter", "4_Shared_Toolkit", "Project Charter", charter)

    def audit(doc):
        add_body(doc, "No deliverable may be submitted until the other workstream has completed "
                 "and signed its section of this checklist.")
        add_heading(doc, "Section A - Workstream A audits Workstream B", 2)
        for item in ["Data cleaning choices are documented in the assumption log.",
                     "The placement definition is stated explicitly.",
                     "Charts are titled honestly and axes are not misleading.",
                     "Outliers were investigated, not silently removed.",
                     "Findings are triangulated with interview evidence.",
                     "No causation is claimed from correlation.",
                     "Disability and gender breakdowns are included where relevant.",
                     "Uncomfortable findings (for example Dong Nai) are not hidden."]:
            add_bullet(doc, "[  ]  " + item)
        add_signature_block(doc, "________________________", "Workstream A auditor")
        add_heading(doc, "Section B - Workstream B audits Workstream A", 2)
        for item in ["Every recommendation claim cites data or a transcript reference.",
                     "Strategic options are financially feasible against D-12 and D-03.",
                     "Ethical trade-offs are quantified where possible.",
                     "The stakeholder map reflects the interviews.",
                     "Journey-map pain points are linked to recommendations.",
                     "There are no contradictions with the cleaned data.",
                     "The implementation roadmap has a realistic timeline and cost.",
                     "Validation feedback has been incorporated (revision log complete)."]:
            add_bullet(doc, "[  ]  " + item)
        add_signature_block(doc, "________________________", "Workstream B auditor")
        add_heading(doc, "Section C - Joint final sign-off", 2)
        add_signature_block(doc, "________________________", "Engagement Lead")
    _template("Cross_Workstream_Audit_Checklist", "4_Shared_Toolkit",
              "Cross-Workstream Audit Checklist", audit)

    def minutes(doc):
        add_meta_table(doc, [("Meeting", ""), ("Date", ""), ("Time", ""),
                             ("Attendees", ""), ("Absent", ""), ("Note-taker", "")])
        add_heading(doc, "Decisions", 2)
        add_table(doc, ["Decision", "Rationale", "Owner"], [("", "", ""), ("", "", "")],
                  widths=[2.6, 2.6, 1.2])
        add_heading(doc, "Actions", 2)
        add_table(doc, ["Action", "Owner", "Due", "Status"],
                  [("", "", "", ""), ("", "", "", "")], widths=[2.8, 1.4, 1.2, 1.0])
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
        add_table(doc, ["Finding", "Source", "Implication"], [("", "", ""), ("", "", "")],
                  widths=[2.4, 1.6, 2.4])
        add_heading(doc, "Ninety-day roadmap", 2)
        add_table(doc, ["Phase", "Actions", "Cost estimate"],
                  [("0-30 days", "", ""), ("31-60 days", "", ""), ("61-90 days", "", "")],
                  widths=[1.4, 3.4, 1.6])
        add_heading(doc, "What we still do not know", 2)
        add_body(doc, "[Honest limitations - this builds credibility.]")
    _template("Executive_Summary", "4_Shared_Toolkit", "Executive Summary", execsum)

    def finalrep(doc):
        add_body(doc, "Recommended length 15-25 pages excluding the appendix. Use this structure.")
        for n, sec in enumerate([
            "Executive summary", "Engagement approach", "Problem definition",
            "Situation analysis", "Stakeholder insights", "Data analysis and findings",
            "Strategic options", "Recommendation", "Implementation roadmap",
            "Ethical analysis", "Risks and mitigations", "Limitations and further work",
            "Appendices"], 1):
            add_number(doc, sec)
        add_heading(doc, "Standards", 2)
        add_bullet(doc, "Every claim carries a source reference.")
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
        add_body(doc, "2. Did the cross-workstream audit work in practice?")
        add_body(doc, "3. Anything the facilitators should know (optional)?")
    _template("Peer_Evaluation", "4_Shared_Toolkit", "Peer Evaluation (Confidential)", peer)

    def reflect(doc):
        add_body(doc, "Individual and private. Submit by Tuesday 18 August 2026, 17:00 ICT.")
        for q in ["What did you learn about structured problem solving?",
                  "What was hardest about working with imperfect data?",
                  "Describe one ethical trade-off your team faced and what you decided.",
                  "What did you contribute, and what would you do differently?",
                  "One skill you want to develop next:"]:
            add_heading(doc, q, 2)
            add_body(doc, "________________________________________________________________")
            add_body(doc, "________________________________________________________________")
    _template("Reflection_Journal", "4_Shared_Toolkit", "Reflection Journal (Individual)", reflect)

    # Workstream A templates
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
        add_heading(doc, "Hypotheses to test", 2)
        for i in range(1, 5):
            add_body(doc, f"H{i}: ________________________________________________")
    _template("Problem_Statement_and_Issue_Tree", "2_Workstream_A_Impact_Strategy",
              "Problem Statement and Issue Tree", issue_tree)

    def stakeholder(doc):
        add_body(doc, "Plot stakeholders on a power/interest grid. Star the five you must "
                 "understand most deeply. Cite at least one source per starred stakeholder.")
        add_table(doc, ["Stakeholder", "Power (1-5)", "Interest (1-5)", "Position on strategy", "Source"],
                  [("", "", "", "", "")] * 8, widths=[1.8, 1.0, 1.0, 2.0, 1.6])
    _template("Stakeholder_Map", "2_Workstream_A_Impact_Strategy", "Stakeholder Map", stakeholder)

    def credibility(doc):
        add_body(doc, "Rate each source in the data room. Justify your rating in one sentence.")
        add_table(doc, ["Source", "Reliability (H/M/L)", "Why", "Key fact you trust or doubt"],
                  [("D-01 Annual Report", "", "", ""), ("D-02 Internal memo", "", "", ""),
                   ("D-03 VPBank agreement", "", "", ""), ("D-08 Board emails", "", "", ""),
                   ("D-10 YouthWorks brief", "", "", ""), ("D-11 Staff survey", "", "", ""),
                   ("D-12 Financials", "", "", ""), ("Interview transcripts", "", "", "")],
                  widths=[1.6, 1.2, 2.0, 2.6])
    _template("Information_Credibility_Matrix", "4_Shared_Toolkit",
              "Information Credibility Matrix", credibility)

    def synth(doc):
        add_body(doc, "Summarise by theme, not by person. Every insight cites a transcript.")
        add_table(doc, ["Theme", "Insight", "Sources", "Contradicts?"],
                  [("", "", "", "")] * 5, widths=[1.5, 2.6, 1.2, 1.1])
        add_heading(doc, "Contradictions to resolve", 2)
        add_body(doc, "For example: is YouthWorks a serious threat? Compare ST-04, ST-05 and D-10.")
    _template("Interview_Synthesis", "2_Workstream_A_Impact_Strategy",
              "Interview Synthesis", synth)

    def empathy(doc):
        add_heading(doc, "Chosen archetype", 2)
        add_body(doc, "For example: a young person in Dong Nai who withdraws during the internship.")
        add_heading(doc, "Empathy map (cite quotes)", 2)
        add_table(doc, ["Says", "Thinks"], [("", "")], widths=[3.2, 3.2])
        add_table(doc, ["Does", "Feels"], [("", "")], widths=[3.2, 3.2])
        add_heading(doc, "Journey map", 2)
        add_table(doc, ["Stage", "What happens", "Emotion", "Pain point", "Evidence"], [
            ("Awareness", "", "", "", ""), ("Waitlist", "", "", "", ""),
            ("Training", "", "", "", ""), ("Internship", "", "", "", ""),
            ("Job search", "", "", "", ""), ("90-day outcome", "", "", "", "")],
            widths=[1.2, 1.6, 1.0, 1.3, 1.3])
    _template("Empathy_and_Journey_Map", "2_Workstream_A_Impact_Strategy",
              "Empathy & Journey Map", empathy)

    def options(doc):
        add_body(doc, "Present two or three genuinely different options. Complete one block each.")
        for letter in ["A", "B", "C"]:
            add_heading(doc, f"Option {letter}", 2)
            add_meta_table(doc, [("Description", ""), ("Rationale", ""),
                                 ("Cost implication", ""), ("Main risks", ""),
                                 ("Who wins / who loses", "")])
        add_heading(doc, "Recommendation", 2)
        add_body(doc, "We recommend Option ____ because ____________________________.")
        add_body(doc, "We would change our mind if ____________________________.")
    _template("Strategic_Options", "2_Workstream_A_Impact_Strategy",
              "Strategic Options", options)

    def ethics(doc):
        add_body(doc, "Address at least three dilemmas. Complete one block each.")
        for ed in ["ED-01 Easiest versus most in need", "ED-02 Metric integrity (60 vs 90 days)",
                   "ED-03 Volunteers versus paid staff", "ED-04 AI intake chatbot",
                   "ED-05 Prevention versus treatment", "ED-06 Social-enterprise conflict"]:
            add_heading(doc, ed, 2)
            add_meta_table(doc, [("The tension", ""), ("Stakeholders affected", ""),
                                 ("Our position", ""), ("Residual risk accepted", "")])
    _template("Ethical_Analysis", "2_Workstream_A_Impact_Strategy", "Ethical Analysis", ethics)

    def validation(doc):
        add_body(doc, "Use only the interview transcripts. Map each draft recommendation to "
                 "evidence. Minimum six citations: at least two beneficiary (BN-), two staff "
                 "(ST-), one donor (DN-) and one employer (EP-).")
        add_table(doc, ["Recommendation", "Transcript", "What it says", "Supports / challenges", "Action"],
                  [("", "", "", "", "")] * 6, widths=[1.5, 1.0, 1.6, 1.3, 1.0])
    _template("Validation_Protocol", "2_Workstream_A_Impact_Strategy",
              "Validation Protocol", validation)

    def revlog(doc):
        add_body(doc, "Record how validation changed your thinking. At least one substantive "
                 "revision is expected.")
        add_table(doc, ["Original element", "What validation revealed", "Revised?", "New position"],
                  [("", "", "", "")] * 5, widths=[1.8, 2.2, 0.9, 1.5])
    _template("Recommendation_Revision_Log", "2_Workstream_A_Impact_Strategy",
              "Recommendation Revision Log", revlog)

    # Workstream B templates
    def dqr(doc):
        add_body(doc, "Two pages maximum. Report what you found and what you did about it.")
        add_heading(doc, "Issues identified", 2)
        add_table(doc, ["#", "Table", "Issue", "Severity", "Action taken"],
                  [("", "", "", "", "")] * 8, widths=[0.4, 1.6, 2.0, 1.0, 1.4])
        add_heading(doc, "Cleaning decisions and rationale", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "Placement definition used", 2)
        add_body(doc, "State which definition you adopted and why.")
    _template("Data_Quality_Report", "3_Workstream_B_Operations_and_Analytics",
              "Data Quality Report", dqr)

    def analysis(doc):
        add_heading(doc, "Headline findings", 2)
        add_table(doc, ["Finding", "Evidence", "Caveat"], [("", "", "")] * 4,
                  widths=[2.4, 2.4, 1.6])
        add_heading(doc, "Attendance and outcomes (DB-3)", 2)
        add_body(doc, "State the association and at least two confounders. Correlation is not causation.")
        add_heading(doc, "Mentoring and completion (DB-4)", 2)
        add_body(doc, "________________________________________________________________")
    _template("Analysis_Summary", "3_Workstream_B_Operations_and_Analytics",
              "Analysis Summary", analysis)

    def evcheck(doc):
        add_body(doc, "Workstream B checks each Workstream A option against the evidence.")
        add_table(doc, ["Option / claim", "Supported by data?", "Financially feasible?", "Concern"],
                  [("", "", "", "")] * 5, widths=[2.0, 1.6, 1.6, 1.2])
    _template("Evidence_Check_Memo", "3_Workstream_B_Operations_and_Analytics",
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
    build_wsb_workbook()
    build_storytelling_guide()
    build_roles_guide()
    build_templates()
    return MANIFEST


if __name__ == "__main__":
    build_all()
