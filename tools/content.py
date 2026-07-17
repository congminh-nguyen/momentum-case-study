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
    add_table, add_meta_table, add_quote, add_signature_block, add_online_confirm,
    add_page_break,
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
        "Programme Guideline",
        "Handbook, weekly deliverables, roles, submission rules and ground rules in one place",
        ref="MOM-2026-GL-01", issued="Sunday 26 July 2026",
    )

    add_heading(doc, "1. Welcome to the engagement", 1)
    add_body(doc, "Your team has been retained as external strategy consultants by the Board "
             "of Generation Bridge Foundation (GBF), a Vietnamese non-profit working on youth "
             "unemployment. The engagement runs from Sunday 26 July to Sunday 20 September 2026 "
             "(eight working weeks, Monday to Friday). You will diagnose a genuine strategic "
             "problem, test your thinking against incomplete evidence, and present a "
             "recommendation to the Board.")
    add_body(doc, "This is a simulated but realistic engagement. There is no single correct "
             "answer. You will be assessed on the quality of your thinking, the honesty of your "
             "evidence, the discipline of your teamwork, and the clarity of your communication.")
    add_body(doc, "Generation Bridge Foundation, VPBank Foundation (as portrayed here), named "
             "staff, beneficiaries and employers are fictional constructs created for the "
             "Momentum Impact Case. They do not represent real organisations or people. Do not "
             "import outside data, news or personal knowledge into deliverables or inferences.",
             bold=True)

    add_heading(doc, "2. The client at a glance", 1)
    add_table(doc, ["Item", "Detail"], [
        ("Organisation", "Generation Bridge Foundation (Cau Noi The He) - fictional case client"),
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
    add_heading(doc, "Suggested roles (assign in Week 1)", 2)
    add_table(doc, ["Role", "Count", "Owns"], [
        ("Engagement Lead", "1", "Timeline, board narrative, final integration, Request Log health"),
        ("Workflow A Lead", "1", "Issue tree, options, trade-off reflection, Analysis Requests"),
        ("Research Lead", "1", "Interview synthesis, credibility matrix, validation"),
        ("Workflow B Lead", "1", "Data dictionary, cleaning, Findings Memos, dashboard"),
        ("Operations Analyst", "1", "Funding scenarios, cost analysis"),
        ("Project & Quality Lead", "1", "Minutes, version control, handoff checklist, file naming"),
    ], widths=[1.9, 0.7, 3.8])
    add_body(doc, "In a team of five, combine Research Lead with Workflow A Lead. Full ways of "
             "working (decision rules, free-riding prevention) are also in "
             "4_Shared_Toolkit/Team_Roles_and_Ways_of_Working.pdf.")

    add_heading(doc, "The analysis request protocol", 2)
    add_body(doc, "Workflow A does not receive the operational datasets and must not ask "
             "facilitators for the file. When A needs numbers, A files an Analysis Request. "
             "Workflow B answers with a Findings Memo in plain language, plus any charts B "
             "chooses to share. B must not hand over the raw workbook.")
    add_body(doc, "Before final options, the team must complete at least three Analysis Requests "
             "with Findings Memos returned. Every quantitative claim in A's deliverables must "
             "cite a Request ID or Findings Memo. The Cross-Workflow Handoff Checklist is the "
             "gate before submission (online tick confirmation - no wet signature required).",
             bold=True)

    add_heading(doc, "5. What you will produce - by week", 1)
    add_body(doc, "Programme opens Sunday 26 July 2026 and closes Sunday 20 September 2026. "
             "Working weeks run Monday to Friday. Weekly submissions are due Friday 17:00 ICT "
             "unless a Board or handoff time is stated. Day-by-day detail lives only in "
             "GBF_Consulting_Toolkit.xlsx (Master Timeline).")
    add_table(doc, ["Week", "Dates (2026)", "Submit by end of week"], [
        ("1", "27-31 July",
         "Kickoff; roles confirmed; Project Charter & RACI; Request Log set up; start data-room reading"),
        ("2", "3-7 August",
         "A1 issue tree; A1b stakeholder map; information credibility matrix"),
        ("3", "10-14 August",
         "R-01 filed + FM-01 returned; B1 data quality (B internal); A2 interview synthesis; "
         "A3 empathy & journey maps; B2 dashboard v1"),
        ("4", "17-21 August",
         "R-02 and R-03 + FM-02/FM-03; B3 analysis summary; A4 strategic options (cite R/FM IDs); "
         "B4 evidence-check memo"),
        ("5", "24-28 August",
         "A5 trade-off reflection; A6 validation; B5 funding scenarios; A7 revision log; "
         "executive summary draft"),
        ("6", "31 Aug-4 Sep",
         "Final report draft for handoff; online handoff checklist confirmed"),
        ("7", "7-11 September",
         "BOARD PRESENTATION Thu 10 Sep 09:00; all final deliverables Fri 11 Sep 17:00"),
        ("8", "14-18 September",
         "Peer evaluation (Mon); reflection journal (Tue); lessons-learned; programme closes Sun 20 Sep"),
    ], widths=[0.7, 1.5, 4.2])
    add_body(doc, "Track Analysis Requests on the Request Log tab. Minimum three completed "
             "R/FM pairs before final options.")

    add_heading(doc, "6. How you will be assessed", 1)
    add_body(doc, "A 100-point rubric applies. The table below is the expected scale for planning "
             "your effort. Facilitators may refine descriptors or band cut-offs; any change will "
             "be announced. Treat these weights as indicative, not a frozen contract.",
             italic=True)
    add_table(doc, ["Criterion", "Points (expected)"], [
        ("Problem understanding", "15"), ("Research quality", "10"),
        ("Design thinking", "15"), ("Data analysis and translation", "20"),
        ("Recommendations", "20"), ("Presentation", "10"),
        ("Project management", "5"), ("Teamwork and cross-workflow exchange", "5"),
    ], widths=[3.6, 1.8])
    add_body(doc, "Presentation Day also applies a panel split (Buddy / Core Team / Guest Board). "
             "See Facilitators/Presentation Day Plan for weights. Strong teams ask precise "
             "questions, translate findings honestly, and own the costs of their preferred "
             "recommendation.")

    add_heading(doc, "7. Citation standard", 1)
    add_body(doc, "Every claim must be trackable for markers. Do not cite a document code alone.")
    add_table(doc, ["Source type", "Required citation form", "Example"], [
        ("Client document", "Ref + section or page heading", "D-02 §1 Dong Nai; D-01 'Our year in numbers'"),
        ("Interview transcript", "Transcript code + speaker", "BN-02 (Tuan); ST-05 (Minh)"),
        ("Findings / requests", "Request or Findings ID", "R-02; FM-03"),
        ("Dataset (B only, via memo)", "Table + field + definition used", "employment_outcomes.placed_90d (formal)"),
    ], widths=[1.6, 2.4, 2.4])

    add_heading(doc, "8. How to get the package and how to submit", 1)
    add_heading(doc, "Download from GitHub", 2)
    add_number(doc, "Open the programme repository (link from your Buddy).")
    add_number(doc, "Click Code, then Download ZIP - one click downloads the full Participants pack "
               "(and Facilitators only if you are staff).")
    add_number(doc, "Or clone with git if your team prefers version control.")
    add_heading(doc, "Shared workspace", 2)
    add_body(doc, "Each team works in the Google Drive / Google Workspace folder created by the "
             "programme (or Microsoft 365 if your cohort is assigned that stack). Buddies and "
             "Core Team only review files in that folder and via the submission Form.")
    add_heading(doc, "File naming (mandatory)", 2)
    add_body(doc, "Use: TeamShortName_Owner_DeliverableCode_Version.ext")
    add_table(doc, ["Example", "Meaning"], [
        ("Tiger_All_Charter_v1.pdf", "Whole-team Project Charter, version 1"),
        ("Tiger_WSA_A1_IssueTree_v2.pdf", "Workflow A deliverable A1"),
        ("Tiger_WSA_R01_Request_v1.pdf", "Analysis Request R-01 (filed by A)"),
        ("Tiger_WSB_FM01_Findings_v1.pdf", "Findings Memo FM-01 (returned by B)"),
        ("Tiger_All_BoardDeck_vFinal.pptx", "Final board deck"),
    ], widths=[3.2, 3.2])
    add_heading(doc, "A_B_Exchange folder (mandatory for assessment)", 2)
    add_body(doc, "Create a folder in your team Drive named A_B_Exchange. Every official "
             "cross-workflow communication must be saved there as PDF using the templates below. "
             "Markers and Buddies assess teamwork from this folder - not from chat history.")
    add_table(doc, ["What", "Template", "Who files"], [
        ("Analysis Request (R-xx)", "4_Shared_Toolkit/Templates/Analysis_Request.docx", "Workflow A"),
        ("Findings Memo (FM-xx)", "4_Shared_Toolkit/Templates/Findings_Memo.docx", "Workflow B"),
    ], widths=[2.2, 3.0, 1.4])
    add_number(doc, "Export each completed Word template to PDF before filing.")
    add_number(doc, "Name files with R-xx / FM-xx codes (examples above).")
    add_number(doc, "Update the Request Log tab whenever a file is added.")
    add_number(doc, "Informal chat is fine for logistics only. Decisions, asks and answers that "
               "affect the recommendation must appear in A_B_Exchange.")
    add_number(doc, "Minimum three completed R/FM pairs in A_B_Exchange before final options.")
    add_heading(doc, "Submission channel", 2)
    add_number(doc, "Upload the named file to your team Drive folder.")
    add_number(doc, "Submit the Drive link (or attach if under size limit) via the programme "
               "Google Form for that week - your Buddy will share the Form URL at kickoff.")
    add_number(doc, "Email is backup only if the Form is down; always CC your Buddy.")
    add_number(doc, "When submitting options or the final report, include the link to A_B_Exchange "
               "so markers can see the request trail.")

    add_heading(doc, "9. Ground rules", 1)
    add_number(doc, "Fictional case only: do not treat GBF or named individuals as real; do not "
               "bring outside data, news or personal research into deliverables or inferences.")
    add_number(doc, "Use only the materials in this package.")
    add_number(doc, "Workflow A must not open or request the raw operational datasets. Route all "
               "quantitative needs through Analysis Requests to Workflow B.")
    add_number(doc, "Workflow B must not share the raw workbook with Workflow A. Share Findings "
               "Memos and curated charts only.")
    add_number(doc, "Official A-B asks and answers live only in A_B_Exchange/ using the Request "
               "and Findings Memo templates. Chat is not the assessment record.")
    add_number(doc, "Cite every claim to the citation standard in Section 7.")
    add_number(doc, "State the definitions you use, especially for the placement rate.")
    add_number(doc, "Fill templates in Word using Nunito or Lexend (Vietnamese-safe). Present in "
               "PowerPoint with the same fonts. Export read-only finals as PDF.")
    add_number(doc, "Handoff confirmation is an online tick + typed name/date - no wet signature.")
    add_number(doc, "Your Buddy is a facilitator, not a consultant. They will not give you the "
               "answer or the data file.")

    add_heading(doc, "10. Package contents", 1)
    add_table(doc, ["Folder", "What is inside"], [
        ("0_Start_Here", "This Programme Guideline; data room index"),
        ("1_Client_Data_Room", "GBF documents (D-01 to D-12) and interview transcripts (no datasets)"),
        ("2_Workflow_A_Impact_Strategy", "Workflow A brief, playbook, A templates"),
        ("3_Workflow_B_Operations_and_Analytics", "Workflow B brief, data dictionary, datasets, dashboard, B templates"),
        ("4_Shared_Toolkit", "Toolkit workbook (Request Log), guides, board deck, shared templates"),
    ], widths=[2.9, 3.6])

    _save(doc, "00_Programme_Guideline", "0_Start_Here", "pdf")


def build_data_room_index():
    doc = _doc("Data Room Index", "A guide to documents and transcripts in the shared client room",
               ref="MOM-2026-DRI-01", issued="Sunday 26 July 2026")
    add_body(doc, "This index lists what is in the shared client data room. In a real engagement, "
             "clients hand you mixed-quality material and expect you to decide what to trust. "
             "We do not pre-rate reliability for you. That is your job in the Information "
             "Credibility Matrix (Week 2). Operational datasets are not in this folder; they "
             "are held by Workflow B under the analysis request protocol.")

    add_heading(doc, "How to question what you read", 1)
    add_body(doc, "Before you lean on any claim, ask:")
    add_number(doc, "Who produced this, for whom, and with what incentive to look good or look candid?")
    add_number(doc, "Are headline metrics defined? Would another definition change the story?")
    add_number(doc, "Does another document or interview say something different about the same topic?")
    add_number(doc, "Is this a measured fact, a management opinion, or a stakeholder preference?")
    add_number(doc, "What would you need Workflow B to check in the operational record before you "
               "treat a number as settled?")
    add_body(doc, "Capture your judgements in the credibility matrix. Public reports, internal "
             "memos, emails, surveys, audited accounts and interviews are not automatically "
             "equal. Disagreement between sources is normal; resolving it is part of the work.")

    add_heading(doc, "How to cite these sources", 1)
    add_body(doc, "Markers will not hunt through an entire PDF for your claim. Cite document "
             "reference plus section heading or page topic, not the code alone. Transcripts: cite "
             "code and speaker. Examples: (D-02, §1 Dong Nai); (D-01, 'Our year in numbers'); "
             "(BN-02, Tuan).")

    add_heading(doc, "Client documents", 1)
    add_table(doc, ["Ref", "Document", "Type / audience"], [
        ("D-01", "Annual Report 2024", "Public report (donors, partners, public)"),
        ("D-02", "Programme Review 2025", "Internal memo (senior staff)"),
        ("D-03", "VPBank Foundation grant agreement (excerpt)", "Funder contract excerpt"),
        ("D-08", "Board email thread", "Informal Board correspondence"),
        ("D-10", "Competitor brief: YouthWorks", "Internal competitive note"),
        ("D-11", "Staff engagement survey 2025", "Anonymous staff survey summary"),
        ("D-12", "Financial statements 2023-2025", "Audited accounts"),
    ], widths=[0.6, 3.2, 2.8])

    add_heading(doc, "Interview transcripts", 1)
    add_body(doc, "The full transcripts are in '20_Interview_Transcripts'. Fourteen "
             "conversations were recorded in January 2026 by the engagement team. Cite as "
             "CODE (Speaker), e.g. ST-01 (Hung). Treat each voice as one perspective, not "
             "as organisational truth.")
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
        ("EP-02", "Ms. Trang", "Employer partner"),
        ("VR-01", "Chi Hoa", "Volunteer mentor, three years' service"),
        ("BD-01", "Mr. T.N.", "Member of the GBF Board"),
    ], widths=[0.7, 2.0, 3.9])

    add_heading(doc, "Operational datasets (Workflow B only)", 1)
    add_body(doc, "The monitoring export and data dictionary sit in Workflow B's folder. "
             "Workflow A does not receive these files. If you need a number, a comparison, or a "
             "test of a hypothesis, file an Analysis Request. Figures that appear in shared "
             "documents are starting points for questions, not a substitute for querying the "
             "operational record.")

    _save(doc, "01_Data_Room_Index", "0_Start_Here", "pdf")


def build_deadlines():
    """Deprecated: weekly deadlines live in 00_Programme_Guideline. Kept as no-op for safety."""
    return


# ===========================================================================
# 1. CLIENT DATA ROOM
# ===========================================================================

def build_wsa_start():
    doc = _doc(
        "Workflow A - Start Here",
        "Impact Strategy: what is in this folder and what you must deliver",
        classification="Workflow A",
        ref="MOM-2026-WSA-00",
        issued="Sunday 26 July 2026",
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
             "Log every request on the Request Log tab of GBF_Consulting_Toolkit.xlsx. Save every "
             "R-xx and FM-xx PDF in your team Drive folder A_B_Exchange/ - that folder is what "
             "markers assess for cross-workflow exchange.")
    add_heading(doc, "Your deliverables", 1)
    add_table(doc, ["Ref", "Deliverable", "Template", "Due (17:00 ICT)"], [
        ("A1", "Problem statement and issue tree", "Problem_Statement_and_Issue_Tree.docx", "Wed 5 Aug"),
        ("A1b", "Stakeholder map", "Stakeholder_Map.docx", "Thu 6 Aug"),
        ("A2", "Interview synthesis", "Interview_Synthesis.docx", "Wed 12 Aug"),
        ("A3", "Empathy and journey maps", "Empathy_and_Journey_Map.docx", "Wed 12 Aug"),
        ("A4", "Strategic options and recommendation", "Strategic_Options.docx", "Thu 20 Aug"),
        ("A5", "Trade-off reflection", "Tradeoff_Reflection.docx", "Wed 26 Aug"),
        ("A6", "Validation protocol", "Validation_Protocol.docx", "Wed 26 Aug"),
        ("A7", "Recommendation revision log", "Recommendation_Revision_Log.docx", "Thu 27 Aug"),
    ], widths=[0.5, 2.4, 2.6, 1.1])
    add_heading(doc, "Before you submit anything", 1)
    add_number(doc, "Every insight from transcripts cites CODE (Speaker), e.g. BN-02 (Tuan).")
    add_number(doc, "Every document cite includes section/heading, e.g. D-02 §1 Dong Nai.")
    add_number(doc, "Every number cites a Findings Memo / Request ID from Workflow B.")
    add_number(doc, "At least three Analysis Requests have Findings Memos returned "
               "(PDFs in A_B_Exchange/).")
    add_number(doc, "Workflow B has completed the online tick on Section B of the "
               "Cross-Workflow Handoff Checklist (typed name + date - no wet signature).")
    _save(doc, "WSA_Start_Here", "2_Workflow_A_Impact_Strategy", "pdf")


def build_wsb_start():
    doc = _doc(
        "Workflow B - Start Here",
        "Operations and Analytics: what is in this folder and what you must deliver",
        classification="Workflow B",
        ref="MOM-2026-WSB-00",
        issued="Sunday 26 July 2026",
    )
    add_body(doc, "You are on Workflow B (Operations and Analytics). You alone hold the "
             "operational datasets and the data dictionary. You also have full access to the "
             "shared client data room (1_Client_Data_Room/) and the Programme Guideline - read "
             "those like any consultant would. Your job is to clean, analyse, and translate "
             "findings so Workflow A can decide. Do not share the raw workbook with A. Answer "
             "their Analysis Requests with Findings Memos in plain language.")
    add_heading(doc, "What you can open", 1)
    add_table(doc, ["Location", "Access"], [
        ("0_Start_Here/", "Yes - Guideline and Data Room Index"),
        ("1_Client_Data_Room/", "Yes - documents D-01 to D-12 and all interview transcripts"),
        ("3_Workflow_B_.../ (this folder)", "Yes - datasets, dictionary, dashboard, B templates"),
        ("4_Shared_Toolkit/", "Yes - Request Log, A_B_Exchange templates, board deck"),
        ("2_Workflow_A_.../ Templates", "Read if useful; do not own A's narrative deliverables"),
        ("GBF_Datasets.xlsx", "B only - never share the raw file with A"),
    ], widths=[3.2, 3.2])
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
        ("B1", "Data quality report (internal)", "Templates/Data_Quality_Report.docx", "Tue 11 Aug"),
        ("B2", "Performance dashboard (working file)", "GBF_Performance_Dashboard.xlsx", "Fri 14 Aug"),
        ("B3", "Analysis summary + Findings Memos", "Templates/Analysis_Summary.docx + Findings_Memo", "Wed 19 Aug"),
        ("B4", "Evidence-check memo on A's options", "Templates/Evidence_Check_Memo.docx", "Fri 21 Aug"),
        ("B5", "Funding scenarios", "Shared Toolkit / Funding Scenarios tab", "Thu 27 Aug"),
    ], widths=[0.5, 2.2, 2.4, 1.3])
    add_heading(doc, "Before you submit anything", 1)
    add_number(doc, "State the placement definition you used and why.")
    add_number(doc, "Document every cleaning decision in the Assumption Log.")
    add_number(doc, "Findings Memos use plain language; charts are curated, not a dump of pivots.")
    add_number(doc, "Every Findings Memo PDF is saved in A_B_Exchange/ and logged on the Request Log.")
    add_number(doc, "Do not claim causation from correlation on attendance or mentoring.")
    add_number(doc, "Workflow A has completed the online tick on Section A of the "
               "Cross-Workflow Handoff Checklist (typed name + date - no wet signature).")
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

    add_heading(doc, "How the exchange with B works", 1)
    add_body(doc, "B is not idle while you frame. In Weeks 1-3 they profile and clean the export "
             "and build a working dashboard so they can answer you quickly. Your job is to file "
             "precise requests once hypotheses exist (from Week 3). Do not delay R-01 until "
             "options are finished - B needs early asks to stress-test your thinking.")

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
        ("A1", "Problem statement and issue tree", "Templates/Problem_Statement_and_Issue_Tree.docx", "Wed 5 Aug"),
        ("A1b", "Stakeholder map", "Templates/Stakeholder_Map.docx", "Thu 6 Aug"),
        ("A2", "Interview synthesis", "Templates/Interview_Synthesis.docx", "Wed 12 Aug"),
        ("A3", "Empathy and journey maps", "Templates/Empathy_and_Journey_Map.docx", "Wed 12 Aug"),
        ("R", "Analysis Requests (minimum three)", "4_Shared_Toolkit/Templates/Analysis_Request.docx", "Weeks 3-5"),
        ("A4", "Strategic options", "Templates/Strategic_Options.docx", "Thu 20 Aug"),
        ("A5", "Trade-off reflection", "Templates/Tradeoff_Reflection.docx", "Wed 26 Aug"),
        ("A6", "Validation protocol (with WF-B)", "Templates/Validation_Protocol.docx", "Wed 26 Aug"),
        ("A7", "Recommendation revision log", "Templates/Recommendation_Revision_Log.docx", "Thu 27 Aug"),
    ], widths=[0.5, 2.4, 2.8, 0.9])

    add_heading(doc, "Detailed requirements", 1)
    add_body(doc, "A1 - Problem statement and issue tree.", bold=True)
    add_bullet(doc, "State the Board's question in one precise sentence of your own.")
    add_bullet(doc, "Build an issue tree that is mutually exclusive and collectively exhaustive.")
    add_bullet(doc, "Write at least four hypotheses. Each hypothesis should imply a question you "
               "could ask Workflow B.")

    add_body(doc, "Analysis Requests.", bold=True)
    add_bullet(doc, "Before final options, complete at least three requests with Findings Memos returned.")
    add_bullet(doc, "File R-01 in Week 3 as soon as A1 hypotheses exist - do not wait for a polished deck.")
    add_bullet(doc, "Each request states: the decision it informs, the hypothesis, what a useful "
               "answer would look like, and urgency. Do not prescribe field names you have not been given.")
    add_bullet(doc, "Ask what would change your mind - not everything that looks interesting. "
               "Use the official Analysis Request template; chat messages are not the record.")
    add_bullet(doc, "Cite Request IDs (R-01, R-02...) on every quantitative claim in A4 and later.")
    add_bullet(doc, "Save every filed request and returned memo as PDF in your team Drive folder "
               "A_B_Exchange/ (see Programme Guideline). Markers assess the exchange from that folder.")

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
             "uncomfortable findings. You cannot submit without their online confirmation on the "
             "handoff checklist.")
    _save(doc, "WSA_Engagement_Brief", "2_Workflow_A_Impact_Strategy", "pdf")


def build_wsa_playbook():
    doc = _doc("Design-Thinking Playbook", "Facilitated activities for Workflow A",
               classification="Participant material - Workflow A", ref="MOM-2026-WSA-02")
    add_body(doc, "This playbook gives you a run sheet for each design-thinking activity. Every "
             "activity must connect to evidence. Design thinking interprets the evidence; it does "
             "not replace it. When you need a number, write an Analysis Request rather than guessing.")

    for title, when, dur, steps, output in [
        ("Stakeholder mapping", "Week 2", "90 minutes",
         ["Brainstorm every party affected by the decision, one per sticky note.",
          "Cluster them, then plot on a power/interest grid.",
          "Star the five you most need to understand and record why."],
         "A stakeholder map and a shortlist for deeper study."),
        ("Problem framing", "Week 2", "120 minutes",
         ["Each member writes 'the real problem is...' independently.",
          "Share, cluster, and agree a single problem statement.",
          "Build the issue tree and generate hypotheses that imply asks for Workflow B."],
         "A problem statement, issue tree and hypotheses (deliverable A1)."),
        ("First Analysis Request", "Week 3", "45 minutes",
         ["Pick the hypothesis that would most change your recommendation if false.",
          "Look back at interviews: which claims about how the programme works still lack numbers?",
          "Complete the official Analysis Request template (R-01) and save the PDF to A_B_Exchange/.",
          "Log it on the Request Log and hand it to Workflow B."],
         "R-01 filed in A_B_Exchange/; waiting for Findings Memo FM-01."),
        ("Empathy mapping", "Week 3", "90 minutes",
         ["Choose one beneficiary archetype grounded in BN-01 to BN-04.",
          "Fill the four quadrants - says, thinks, does, feels - using cited quotes only.",
          "List what you still do not know, and which of those gaps need a request to B."],
         "An evidence-based empathy map (part of deliverable A3)."),
        ("Journey mapping", "Week 3", "120 minutes",
         ["Map the stages from awareness to ninety-day employment.",
          "Mark the emotional highs and lows and the pain points.",
          "Attach a quote or a Findings Memo cite to each pain point."],
         "A current-state journey map with prioritised pain points."),
        ("Ideation - How Might We and Crazy Eights", "Week 4", "120 minutes",
         ["Turn each priority pain point into a 'How might we...' question.",
          "Generate ideas quickly and without judgement, then sketch eight variations of the best.",
          "Dot-vote; check surviving ideas against what the transcripts actually say."],
         "A shortlist of concepts to develop into options."),
        ("Prioritisation with Findings", "Week 4", "60 minutes",
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
             "use. Success is not a dump of pivots; it is a clear answer to a question A actually asked "
             "- delivered fast because you prepared the data before the ask arrived.")

    add_heading(doc, "Hard constraint on sharing", 1)
    add_body(doc, "Do not share GBF_Datasets.xlsx or the data dictionary with Workflow A. "
             "Respond to Analysis Requests with Findings Memos and curated charts or summary "
             "tables you choose to attach. If A asks for 'the file', refuse and ask what decision "
             "the number needs to support.", bold=True)

    add_heading(doc, "Two modes of work (do not wait idle)", 1)
    add_body(doc, "You do not need A's permission to understand the data. You do need a request "
             "before you push decision numbers into A's narrative.")
    add_table(doc, ["Mode", "When", "What you do"], [
        ("Foundation (proactive)", "Weeks 1-3",
         "Read the dictionary; profile and clean (B1 / DB-1); build a working dashboard (B2 / DB-2). "
         "This is your readiness work while A frames the problem."),
        ("Service (reactive)", "From Week 3",
         "Answer Analysis Requests with Findings Memos. Package only what the request needs."),
        ("Challenge (joint)", "Weeks 4-6",
         "Evidence-check A's options (B4); run funding scenarios (B5) once options exist."),
    ], widths=[1.6, 1.2, 3.6])
    add_body(doc, "If no request has arrived by mid-Week 3, nudge A for R-01. Do not invent a "
             "strategy for them - ask what decision is stuck.")

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
        ("B1", "Data quality report (internal to B)", "Templates/Data_Quality_Report.docx -> PDF", "Tue 11 Aug"),
        ("FM", "Findings Memos (responses to R-01+)", "4_Shared_Toolkit/Templates/Findings_Memo.docx", "From Week 3"),
        ("B2", "Performance dashboard (working file)", "GBF_Performance_Dashboard.xlsx", "Fri 14 Aug"),
        ("B3", "Analysis summary (tied to Requests answered)", "Templates/Analysis_Summary.docx -> PDF", "Wed 19 Aug"),
        ("B4", "Evidence-check memo on A's options", "Templates/Evidence_Check_Memo.docx -> PDF", "Fri 21 Aug"),
        ("B5", "Funding scenarios", "Toolkit workbook, 'Funding Scenarios' tab", "Thu 27 Aug"),
    ], widths=[0.5, 2.5, 2.7, 0.9])

    add_heading(doc, "Analysis backlog", 1)
    add_body(doc, "DB-1 and DB-2 are mandatory foundation. Everything else follows the request: "
             "when A asks a decision question, decide which tables answer it, analyse, and return "
             "a Findings Memo. If you are unsure how to approach a request, ask your Buddy for "
             "coaching questions - not for the answer.")

    add_body(doc, "DB-1 - Data quality audit (foundation).", bold=True)
    add_bullet(doc, "Profile every table for missing values, duplicates and format problems.")
    add_bullet(doc, "Use WSB_Data_Dictionary.pdf as your field guide. Investigate anomalies; "
               "do not delete silently.")
    add_bullet(doc, "Record every cleaning decision in the assumption log. Share with A only the "
               "material issues that affect a decision (via Findings Memo), not the full cleaning log.")

    add_body(doc, "DB-2 - Performance dashboard (foundation).", bold=True)
    add_bullet(doc, "Use pivot tables to show placement by hub, programme and gender.")
    add_bullet(doc, "Flag any rate that looks material under the definition you state.")
    add_bullet(doc, "Keep this as a B working file. Export only curated views into Findings Memos "
               "when A asks.")

    add_body(doc, "On-request analysis (no fixed recipe).", bold=True)
    add_bullet(doc, "When A files an Analysis Request in A_B_Exchange/, respond with a Findings Memo "
               "in the same folder using the official template.")
    add_bullet(doc, "Your job: restate their decision question, choose tables from the dictionary, "
               "compute carefully, state definitions and caveats.")
    add_bullet(doc, "Do not invent a full analysis plan before a request arrives. Do not share "
               "raw tables. If stuck on method, ask Buddy - they will nudge, not solve.")
    add_bullet(doc, "Chat acknowledgements are fine; the assessable answer is the FM-xx PDF.")

    add_body(doc, "Funding scenarios (after draft options).", bold=True)
    add_bullet(doc, "Run once A has sketched options or asks what growth costs (deliverable B5).")
    add_bullet(doc, "Use audited costs (D-12) and funder terms (D-03); state every assumption.")
    add_bullet(doc, "Judge whether each scenario can meet placement conditions under an honest "
               "definition you state.")

    add_heading(doc, "How Workflow A will challenge you", 1)
    add_body(doc, "Before you submit, Workflow A will check that you answered their questions "
             "in plain language, that you did not share the raw file, that definitions and "
             "caveats are stated, and that unanswered requests are logged. You cannot submit "
             "without their online confirmation on the handoff checklist.")
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
        ("beneficiary_id", "Primary join key across tables"),
        ("placed_90d", "Placement flag used in outcome reporting"),
        ("employment_type", "Categories of employment recorded at follow-up"),
        ("province / hub", "Geography of delivery or residence"),
        ("date fields", "Enrolment, completion and outcome dates"),
        ("wage_vnd_monthly", "Self-reported or recorded monthly wage"),
        ("disability_status", "Disability flag where captured"),
        ("amount_vnd_m", "Funding amounts (often in millions VND)"),
    ], widths=[2.0, 4.4])

    add_heading(doc, "Treat the export like a real monitoring dump", 1)
    add_body(doc, "Operational exports are rarely clean. Profile every table before you analyse. "
             "Document what you find, how you handled it, and what Workflow A needs to know "
             "before relying on a rate. Do not assume definitions match the language in public "
             "reports. If a figure in a shared document and a figure you calculate disagree, "
             "investigate - do not silently pick the nicer one.")

    add_heading(doc, "Sample vs organisation totals", 1)
    add_body(doc, "This export covers a subset of people served, not necessarily the full "
             "organisation total quoted in public materials. State that limitation whenever "
             "you report rates. Do not silently extrapolate.")
    _save(doc, "WSB_Data_Dictionary", "3_Workflow_B_Operations_and_Analytics", "pdf")


def build_wsb_workbook():
    doc = _doc("Data Analysis Workbook", "Foundation methods and request handling for Workflow B",
               classification="Participant material - Workflow B", ref="MOM-2026-WSB-02")
    add_body(doc, "This workbook covers foundation work (quality + dashboard) and how to turn an "
             "Analysis Request into a Findings Memo. It does not prescribe every deep-dive method. "
             "When A asks something new, design the analysis yourself; ask your Buddy for coaching "
             "questions if you are stuck. You do not need to write code.")

    add_heading(doc, "Suggested order", 1)
    add_number(doc, "Weeks 1-3: dictionary, profile, clean (DB-1 / B1), then dashboard pivots (DB-2 / B2).")
    add_number(doc, "From Week 3: answer each Analysis Request with a Findings Memo before starting "
               "optional deep-dives.")
    add_number(doc, "Funding scenarios (B5) once A has draft options or asks what growth costs.")

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
    add_body(doc, "If the request is vague ('send us everything on mentoring'), send it back for a "
             "decision question. If the method is unclear, ask Buddy: 'Which tables would you look "
             "at first?' - not 'What is the answer?'")

    add_heading(doc, "Foundation: dashboard pivots (DB-2)", 1)
    add_number(doc, "After cleaning, build placement rates by hub, programme and gender.")
    add_number(doc, "State the placement definition on every chart.")
    add_number(doc, "Keep the workbook internal; only curated views leave via Findings Memos.")

    add_heading(doc, "Reporting standards", 1)
    add_bullet(doc, "State the placement definition you used and why.")
    add_bullet(doc, "Never delete an outlier silently - investigate and record the decision.")
    add_bullet(doc, "Give every chart a title that states the finding.")
    add_bullet(doc, "Note sample-size limitations versus organisation-wide totals in public materials.")
    add_bullet(doc, "Never claim causation from a simple association.")
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
    add_body(doc, "Role assignments also appear in 00_Programme_Guideline.pdf (Section 4) so "
             "you can set them in Week 1 without opening this file. This page adds ways of "
             "working and anti-free-riding rules.")
    add_heading(doc, "Suggested roles", 1)
    add_table(doc, ["Role", "Count", "Owns"], [
        ("Engagement Lead", "1", "Timeline, board narrative, final integration, Request Log health"),
        ("Workflow A Lead", "1", "Issue tree, options, trade-off reflection, Analysis Requests"),
        ("Research Lead", "1", "Interview synthesis, credibility matrix, validation"),
        ("Workflow B Lead", "1", "Data dictionary, cleaning, Findings Memos, dashboard"),
        ("Operations Analyst", "1", "Funding scenarios, cost analysis"),
        ("Project & Quality Lead", "1", "Minutes, version control, handoff checklist, file naming"),
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
             "handoff confirmation (online tick) is required before anything leaves the team.")
    add_heading(doc, "Fonts for Vietnamese", 1)
    add_body(doc, "Use Nunito (body) or Lexend (headings) in Word and PowerPoint so Vietnamese "
             "text does not break. Lora is acceptable for long-form narrative if preferred.")
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
            ("Board presentation", "Thursday 10 September 2026, 09:00 ICT"),
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
        add_heading(doc, "5. Online confirmation", 2)
        add_body(doc, "No wet signature required. Each member ticks and types name + date "
                 "(Drive comment or sheet tick is fine).")
        for _ in range(3):
            add_online_confirm(doc, "Team member")
    _template("Project_Charter", "4_Shared_Toolkit", "Project Charter", charter)

    def handoff(doc):
        add_body(doc, "No deliverable may be submitted until both workflows have completed and "
                 "confirmed their sections. Use online ticks (typed name + date) - no wet "
                 "signature. The Engagement Lead confirms only if A_B_Exchange/ and the Request "
                 "Log show genuine exchange (not a single dump at the end).")
        add_heading(doc, "Section A - Workflow A reviews Workflow B", 2)
        for item in [
            "Findings Memos answer the question asked, in plain language.",
            "Each FM-xx PDF is saved in A_B_Exchange/ and logged on the Request Log.",
            "The raw workbook and data dictionary were not shared with A.",
            "Placement definition and sample limitations are stated.",
            "Caveats are honest; uncomfortable findings are not hidden.",
            "Unanswered or delayed requests are logged with reasons.",
            "No causation is claimed from correlation.",
            "Charts attached to memos are curated, not a pivot dump.",
        ]:
            add_bullet(doc, "[  ]  " + item)
        add_online_confirm(doc, "Workflow A reviewer")
        add_heading(doc, "Section B - Workflow B reviews Workflow A", 2)
        for item in [
            "Every quantitative claim cites a Request ID or Findings Memo ID.",
            "Document cites include section/heading (not code alone).",
            "At least three Analysis Request PDFs are in A_B_Exchange/ with matching Findings Memos.",
            "Strategic options are financially feasible against D-12 and D-03.",
            "Options changed after at least one B challenge (see revision log).",
            "Journey-map pain points link to recommendations and/or Findings cites.",
            "Trade-off reflection names who gains and who loses.",
            "Implementation roadmap has a realistic timeline and cost.",
        ]:
            add_bullet(doc, "[  ]  " + item)
        add_online_confirm(doc, "Workflow B reviewer")
        add_heading(doc, "Section C - Joint final confirmation", 2)
        add_body(doc, "A_B_Exchange/ and Request Log show minimum three completed R/FM pairs: [  ] Yes")
        add_body(doc, "Drive link to A_B_Exchange/: ________________________________")
        add_online_confirm(doc, "Engagement Lead")
    _template("Cross_Workflow_Handoff_Checklist", "4_Shared_Toolkit",
              "Cross-Workflow Handoff Checklist", handoff)

    def analysis_request(doc):
        add_body(doc, "OFFICIAL CHANNEL - Workflow A. Complete one form per ask. Export to PDF and "
                 "save in your team Drive folder A_B_Exchange/ using naming "
                 "Team_WSA_Rxx_Request_vN.pdf. Also log the request on the Request Log tab. "
                 "Chat or verbal asks do not count for assessment.")
        add_body(doc, "Do not prescribe field names you have not been given. Working out what to "
                 "ask is part of the exercise.")
        add_meta_table(doc, [
            ("Request ID (R-xx)", ""),
            ("Date filed", ""),
            ("Filed by (name + role)", ""),
            ("Urgency (routine / needed for options / blocking)", ""),
            ("Saved to A_B_Exchange? (Y/N)", ""),
        ])
        add_heading(doc, "1. Decision this informs", 2)
        add_body(doc, "What choice would this answer change?")
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "2. Hypothesis or claim to test", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "3. What a useful answer would look like", 2)
        add_body(doc, "(For example: a comparison across hubs; a rate under a stated definition; "
                 "a cost range. Be specific about the decision, not the spreadsheet.)")
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "4. Qualitative sources that prompted this ask", 2)
        add_body(doc, "(Transcript or document refs with section/speaker)")
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "5. What would change your mind", 2)
        add_body(doc, "If the operational record showed X instead of Y, we would ________________")
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "6. Workflow B acknowledgement", 2)
        add_meta_table(doc, [("Received by", ""), ("Target return date", ""),
                             ("Findings Memo ID (FM-xx)", "")])
    _template("Analysis_Request", "4_Shared_Toolkit", "Analysis Request", analysis_request)

    def findings_memo(doc):
        add_body(doc, "OFFICIAL CHANNEL - Workflow B. Complete one memo per Analysis Request. "
                 "Export to PDF and save in A_B_Exchange/ as Team_WSB_FMxx_Findings_vN.pdf. "
                 "Update the Request Log to Returned. Do not attach the raw workbook. Write for "
                 "a non-analyst.")
        add_meta_table(doc, [
            ("Findings Memo ID (FM-xx)", ""),
            ("Responds to Request ID", ""),
            ("Date returned", ""),
            ("Author (name + role)", ""),
            ("Saved to A_B_Exchange? (Y/N)", ""),
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
        add_heading(doc, "6. Implication for A's decision (optional one sentence)", 2)
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "7. Attachments", 2)
        add_body(doc, "List curated charts or summary tables only (max two recommended). "
                 "File names of any charts also saved in A_B_Exchange/:")
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
        add_body(doc, "Confidential. Submit individually by Monday 14 September 2026, 17:00 ICT. "
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
        add_body(doc, "Individual and private. Submit by Tuesday 15 September 2026, 17:00 ICT.")
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
        add_table(doc, ["Branch", "Sub-questions", "Testable hypothesis", "Sources"], [
            ("Programmes", "", "", ""), ("Geography", "", "", ""), ("Funding", "", "", ""),
            ("Operations", "", "", ""), ("Measurement", "", "", "")], widths=[1.1, 1.8, 1.8, 1.7])
        add_body(doc, "Sources column: cite transcript CODE (Speaker) and/or document Ref + "
                 "section (e.g. D-02 §1; BN-02 Tuan). Do not leave Sources blank for any "
                 "hypothesis you intend to test.", italic=True, size=10)
        add_heading(doc, "Hypotheses to test (each should imply an Analysis Request)", 2)
        for i in range(1, 5):
            add_body(doc, f"H{i}: ________________________________________________")
            add_body(doc, f"    Sources: ________________________________________")
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
        add_body(doc, "You rate each shared source - the data room does not do it for you. "
                 "For each row, write why you trust or doubt it (author, audience, incentive, "
                 "definitions, contradictions with other sources). Operational datasets are held "
                 "by Workflow B; rate Findings Memos once received.")
        add_table(doc, ["Source", "Reliability (H/M/L)", "Why", "Key fact you trust or doubt"],
                  [("D-01 Annual Report", "", "", ""), ("D-02 Internal memo", "", "", ""),
                   ("D-03 VPBank agreement", "", "", ""), ("D-08 Board emails", "", "", ""),
                   ("D-10 YouthWorks brief", "", "", ""), ("D-11 Staff survey", "", "", ""),
                   ("D-12 Financials", "", "", ""), ("Interview transcripts", "", "", ""),
                   ("Findings Memos (from B)", "", "", "")],
                  widths=[1.6, 1.2, 2.0, 2.6])
        add_body(doc, "Prompt questions (optional): Who benefits if this source is believed? "
                 "What definition sits behind any rate? What would change your rating?")
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
        add_heading(doc, "Deep-dives completed (only if A requested them)", 2)
        add_body(doc, "Summarise any on-request analyses (e.g. placement drivers, attendance, "
                 "mentoring, completion). For each: what A asked, what you found, definitions used, "
                 "and whether you are claiming association or causation.")
        add_body(doc, "________________________________________________________________")
        add_heading(doc, "Open questions for further requests", 2)
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
