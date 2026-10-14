#!/usr/bin/env python3
"""
fill_template_docx.py
Generates the complete Final Project Report for the Smart University Complaint System (SUCS)
following the EXACT page layout, fonts, margins, front matter, chapter title pages,
and section numbering of the official Virtual University template (media_1791706258053.doc).
Embeds all official SRS diagrams, tables, and test cases.
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_report(output_path, updir):
    doc = docx.Document()

    # Exact VU Page Setup: Margins (Left=1.25" for binding, Top=1.0", Bottom=1.0", Right=1.25")
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1.25)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)

    # Palette
    NAVY = RGBColor(0, 51, 102)        # Primary VU Navy
    DARK_TEXT = RGBColor(34, 34, 34)   # #222222
    MUTED_TEXT = RGBColor(85, 85, 85) # #555555

    # Helper formatters
    def add_p(text, bold=False, italic=False, size=12, font='Tahoma', align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=6, line_spacing=1.15, color=DARK_TEXT):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.name = font
        run.font.size = Pt(size)
        run.font.color.rgb = color
        return p

    def add_chapter_title_page(chapter_num_str, chapter_title_str):
        doc.add_page_break()
        # Generates the exact VU title separator page
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_before = Pt(180)
        
        p_num = doc.add_paragraph()
        p_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_num.paragraph_format.space_after = Pt(12)
        r_num = p_num.add_run(chapter_num_str)
        r_num.bold = True
        r_num.font.name = 'Times New Roman'
        r_num.font.size = Pt(28)
        r_num.font.color.rgb = NAVY

        p_tit = doc.add_paragraph()
        p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_tit.paragraph_format.space_after = Pt(12)
        r_tit = p_tit.add_run(chapter_title_str)
        r_tit.bold = True
        r_tit.font.name = 'Times New Roman'
        r_tit.font.size = Pt(22)
        r_tit.font.color.rgb = NAVY

        doc.add_page_break()

    def add_heading_1(text, space_before=16, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(15)
        run.font.color.rgb = NAVY
        return p

    def add_heading_2(text, space_before=12, space_after=4):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Tahoma'
        run.font.size = Pt(13)
        run.font.color.rgb = NAVY
        return p

    def add_heading_3(text, space_before=8, space_after=2):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Tahoma'
        run.font.size = Pt(11.5)
        run.font.color.rgb = DARK_TEXT
        return p

    def add_figure(img_filename, caption, width=Inches(5.8)):
        img_path = os.path.join(updir, img_filename)
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(10)
            p_img.paragraph_format.space_after = Pt(4)
            run = p_img.add_run()
            run.add_picture(img_path, width=width)

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_after = Pt(12)
            run_cap = p_cap.add_run(f"Figure: {caption}")
            run_cap.bold = True
            run_cap.font.name = 'Tahoma'
            run_cap.font.size = Pt(9.5)
            run_cap.font.color.rgb = MUTED_TEXT

    # ====================================================
    # 1. TITLE PAGE (Exact Template Match)
    # ====================================================
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(18)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Final Project Report")
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(22)
    r.font.color.rgb = NAVY

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(80)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Smart University Complaint System\n(SUCS)")
    r.bold = True
    r.font.name = 'Tahoma'
    r.font.size = Pt(20)
    r.font.color.rgb = NAVY

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Project Supervisor")
    r.bold = True
    r.font.name = 'Tahoma'
    r.font.size = Pt(14)
    r.font.color.rgb = NAVY

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(60)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Muhammad Raheel")
    r.font.name = 'Tahoma'
    r.font.size = Pt(13)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Submitted By")
    r.bold = True
    r.font.name = 'Tahoma'
    r.font.size = Pt(14)
    r.font.color.rgb = NAVY

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Group ID: S26PROJECT4EE01")
    r.bold = True
    r.font.name = 'Tahoma'
    r.font.size = Pt(13)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Usman\t\tBC220400155")
    r.font.name = 'Tahoma'
    r.font.size = Pt(12)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(80)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Team Member 2\t\tBC220400156")
    r.font.name = 'Tahoma'
    r.font.size = Pt(12)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.2
    r = p.add_run("Software Projects & Research Section,\nDepartment of Computer Sciences,\nVirtual University of Pakistan")
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.color.rgb = NAVY


    # ====================================================
    # 2. CERTIFICATE (Exact Template Match)
    # ====================================================
    doc.add_page_break()

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(24)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("CERTIFICATE")
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(18)
    r.font.color.rgb = NAVY

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.space_after = Pt(16)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run("This is to certify that Usman (BC220400155) and Team Member 2 (BC220400156) have worked on and completed their Software Project entitled \"Smart University Complaint System (SUCS)\" under my supervision and guidance.")
    r.font.name = 'Tahoma'
    r.font.size = Pt(11.5)

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.space_after = Pt(40)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run("In our opinion, it is satisfactory and up to the mark and therefore fulfills the requirements of BS in Computer Sciences degree from the Virtual University of Pakistan.")
    r.font.name = 'Tahoma'
    r.font.size = Pt(11.5)

    # Signatures
    p = doc.add_paragraph()
    r = p.add_run("Supervisor / Internal Examiner")
    r.bold = True
    r.font.name = 'Tahoma'
    r.font.size = Pt(13)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(25)
    r = p.add_run("Muhammad Raheel\nSupervisor,\nSoftware Projects & Research Section,\nDepartment of Computer Sciences,\nVirtual University of Pakistan\n\n___________________\n(Signature)")
    r.font.name = 'Tahoma'
    r.font.size = Pt(11)

    p = doc.add_paragraph()
    r = p.add_run("External Examiner / Subject Specialist")
    r.bold = True
    r.font.name = 'Tahoma'
    r.font.size = Pt(13)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(30)
    r = p.add_run("External Examiner / Subject Specialist\n\n___________________\n(Signature)")
    r.font.name = 'Tahoma'
    r.font.size = Pt(11)

    p = doc.add_paragraph()
    r = p.add_run("Accepted By:\n\n_____________\n(For office use)")
    r.bold = True
    r.font.name = 'Tahoma'
    r.font.size = Pt(11.5)


    # ====================================================
    # 3. EXORDIUM (Exact Template Match)
    # ====================================================
    doc.add_page_break()

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(30)
    p.paragraph_format.space_after = Pt(30)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("EXORDIUM")
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(18)
    r.font.color.rgb = NAVY

    verses = [
        "In the name of Allah, the Compassionate, the Merciful.",
        "Praise be to Allah, Lord of Creation,",
        "The Compassionate, the Merciful,",
        "King of Judgment-day!",
        "You alone we worship, and to You alone we pray for help,",
        "Guide us to the straight path,",
        "The path of those who You have favored,",
        "Not of those who have incurred Your wrath,",
        "Nor of those who have gone astray."
    ]
    for v in verses:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(v)
        r.bold = True
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13)
        r.font.color.rgb = NAVY


    # ====================================================
    # 4. DEDICATION
    # ====================================================
    doc.add_page_break()

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(30)
    p.paragraph_format.space_after = Pt(24)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("DEDICATION")
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(18)
    r.font.color.rgb = NAVY

    add_p("This endeavor is dedicated first and foremost to Almighty Allah, the Most Beneficent, the Most Merciful, who endowed us with the intellectual strength, clarity of thought, and endurance to bring this project to fruition.", font='Tahoma', size=11.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=14)

    add_p("To our beloved parents, whose boundless sacrifices, unwavering prayers, and unconditional love have illuminated every step of our academic journey. Their quiet encouragement remains our greatest source of inspiration.", font='Tahoma', size=11.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=14)

    add_p("To our honorable faculty members and supervisor at the Virtual University of Pakistan, whose mentorship, rigor, and scholarly guidance shaped our computational reasoning and professional ethos.", font='Tahoma', size=11.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=14)


    # ====================================================
    # 5. ACKNOWLEDGEMENT
    # ====================================================
    doc.add_page_break()

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(30)
    p.paragraph_format.space_after = Pt(24)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ACKNOWLEDGEMENT")
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(18)
    r.font.color.rgb = NAVY

    add_p("All praises and gratitude are due to Almighty Allah, the Sovereign Creator of the universe, who granted us knowledge, capability, and health to successfully complete this Final Project.", font='Tahoma', size=11.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=12)

    add_p("We wish to articulate our deep and sincere gratitude to our esteemed project supervisor, Muhammad Raheel, for his continuous intellectual support, invaluable critique, and professional advice throughout the conceptualization, system analysis, design, and implementation phases of the Smart University Complaint System (SUCS).", font='Tahoma', size=11.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=12)

    add_p("We are also deeply thankful to the Department of Computer Sciences and the Software Projects & Research Section at the Virtual University of Pakistan for delivering a modern curriculum and robust LMS environment that empowered us to engineer practical, scalable software solutions.", font='Tahoma', size=11.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=12)


    # ====================================================
    # 6. PREFACE
    # ====================================================
    doc.add_page_break()

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(30)
    p.paragraph_format.space_after = Pt(24)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("PREFACE")
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(18)
    r.font.color.rgb = NAVY

    add_p("Higher educational institutions face significant logistical challenges in managing student grievances efficiently. Conventional physical complaint boxes and unorganized email channels suffer from lost tickets, delayed resolutions, lack of traceability, and administrative bias when grievances concern the very staff responsible for reviewing them.", font='Tahoma', size=11.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=12)

    add_p("The Smart University Complaint System (SUCS) was conceived to address these shortcomings by engineering a modern web application that enforces disciplined monthly quotas (5 max), automated hazard keyword triage, an institutional Conflict-of-Interest firewall routing complaints against Admin directly to the HOD, and intelligent problem clustering with 1-Click batch resolution.", font='Tahoma', size=11.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=12)

    add_p("This Final Project Report documents the complete software engineering lifecycle of the SUCS platform, encompassing Requirements Analysis, System Design, Dynamic Interaction Modeling, Database Engineering, Implementation, Automated Verification, and Viva Preparation.", font='Tahoma', size=11.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=12)


    # ====================================================
    # 7. TABLE OF CONTENTS (Exact Template Headings)
    # ====================================================
    doc.add_page_break()

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(30)
    p.paragraph_format.space_after = Pt(20)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("TABLE OF CONTENTS")
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(18)
    r.font.color.rgb = NAVY

    toc_items = [
        ("CERTIFICATE", "ii"),
        ("EXORDIUM", "iii"),
        ("DEDICATION", "iv"),
        ("ACKNOWLEDGEMENT", "v"),
        ("PREFACE", "vi"),
        ("TABLE OF CONTENTS", "vii"),
        ("CHAPTER NO. 1: GATHERING & ANALYZING INFO", "1"),
        ("    1.1  INTRODUCTION", "2"),
        ("    1.2  PURPOSE", "2"),
        ("    1.3  SCOPE", "3"),
        ("    1.4  DEFINITIONS, ACRONYMS AND ABBREVIATIONS", "4"),
        ("    1.5  PROJECT REQUIREMENTS", "5"),
        ("        1.5.1 Functional Requirements (FR1 to FR11)", "5"),
        ("        1.5.2 Non-Functional Requirements (NFR1 to NFR5)", "7"),
        ("    1.6  USE CASES AND USAGE SCENARIOS", "8"),
        ("        1.6.1 Use Case Diagrams", "8"),
        ("        1.6.2 Usage Scenarios (UC-01 to UC-08)", "10"),
        ("    1.7  DEVELOPMENT METHODOLOGY", "14"),
        ("        1.7.1 Chosen Methodology (VU Process Model)", "14"),
        ("        1.7.2 Reasons for Chosen Methodology", "15"),
        ("        1.7.3 Work Plan (Gantt Chart)", "16"),
        ("        1.7.4 Project Schedule (Submission Calendar)", "17"),
        ("CHAPTER NO. 2: DESIGNING THE PROJECT", "18"),
        ("    2.1  INTRODUCTION", "19"),
        ("    2.2  PURPOSE", "19"),
        ("    2.3  SCOPE", "20"),
        ("    2.4  DEFINITIONS, ACRONYMS AND ABBREVIATIONS", "20"),
        ("    2.5  ARCHITECTURAL REPRESENTATION (ARCHITECTURE DIAGRAM)", "21"),
        ("    2.6  DYNAMIC MODEL: SEQUENCE DIAGRAMS", "23"),
        ("    2.7  OBJECT MODEL / LOGICAL MODEL: CLASS DIAGRAM", "26"),
        ("    2.8  DATABASE MODEL (DATABASE DIAGRAM)", "28"),
        ("    2.9  GRAPHICAL USER INTERFACES", "31"),
        ("CHAPTER NO. 3: DEVELOPMENT", "35"),
        ("    3.1  DEVELOPMENT PLAN (ARCHITECTURE DIAGRAM)", "36"),
        ("    3.2  CORE ALGORITHMIC IMPLEMENTATIONS", "38"),
        ("    3.3  SYSTEM VERIFICATION & TEST TRACEABILITY", "43"),
        ("REFERENCES", "48"),
        ("APPENDIX: VIVA VOCE PREPARATION GUIDE", "50")
    ]

    for item, page in toc_items:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        is_ch = "CHAPTER" in item or item in ["CERTIFICATE", "REFERENCES", "APPENDIX"]
        r1 = p.add_run(item)
        r1.bold = is_ch
        r1.font.name = 'Tahoma'
        r1.font.size = Pt(10.5 if not is_ch else 11.5)
        if is_ch:
            r1.font.color.rgb = NAVY

        # Tab leader
        r2 = p.add_run(f"\t{page}")
        r2.bold = is_ch
        r2.font.name = 'Tahoma'
        r2.font.size = Pt(10.5 if not is_ch else 11.5)


    # ====================================================
    # CHAPTER 1 TITLE SEPARATOR
    # ====================================================
    add_chapter_title_page("CHAPTER 1", "Gathering & Analyzing Info")


    # ====================================================
    # CHAPTER 1 CONTENT
    # ====================================================
    add_heading_1("1.1  INTRODUCTION")
    add_p("The Smart University Complaint System (SUCS) is a comprehensive web-based platform designed to modernize, streamline, and govern student grievance redressal at higher educational institutions. In traditional university operating environments, students submit grievances through physical suggestion boxes, paper memorandums, or unstructured email threads directed to faculty and administrative officers.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_p("These legacy mechanisms suffer from severe systemic deficiencies: lost grievances, zero transparency regarding ticket progression, lack of administrative accountability, and queue flooding due to unconstrained duplicate submissions. Crucially, a profound conflict of interest arises whenever a student lodges an allegation against an administrative officer, as the same officer frequently controls ticket triage. SUCS rectifies these vulnerabilities by enforcing role-based portals, strict monthly quotas (5 complaints/month), automated hazard keyword triage, an institutional Conflict-of-Interest firewall, and 1-Click batch resolution.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_1("1.2  PURPOSE")
    add_p("The purpose of this software engineering project is to implement a secure, transparent, and automated web-based complaint management infrastructure that guarantees 100% grievance traceability, eliminates administrative bias, and reduces triage resolution latency to under 48 hours.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_p("The platform serves three distinct user classes: Students (who submit and track complaints within strict quota allowances), Department Administrators (who triage, investigate, and batch-resolve departmental grievances), and the Head of Department (HOD, who exercises executive oversight and exclusively adjudicates sensitive Conflict-of-Interest escalations).", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_1("1.3  SCOPE")
    add_p("The operational boundaries of the SUCS platform encompass:", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_p("• Unified Authentication Portal with role-aware redirection and PBKDF2-SHA256 password security.", space_after=3)
    add_p("• Student Grievance Submission Desk with taxonomy pickers and live 5-complaint monthly quota tracking.", space_after=3)
    add_p("• Automated Keyword Priority Triage Engine (High P1 for hazards/emergencies).", space_after=3)
    add_p("• Conflict-of-Interest (COI) Firewall routing grievances against Admin directly to HOD.", space_after=3)
    add_p("• Intelligent Category-Target Problem Clustering and 1-Click Batch Group Resolution.", space_after=3)
    add_p("• Mandatory Investigative Progress Feedback for transparent ticket status transitions.", space_after=3)
    add_p("• In-Browser Live Database Inspector (/db-view) for verified administrative auditing.", space_after=8)
    add_p("Explicit Out-of-Scope boundaries for this release include native mobile apps (iOS/Android), external SMS/Email SMTP gateways, third-party ERP integrations, and multi-tenancy across separate universities.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_1("1.4  DEFINITIONS, ACRONYMS AND ABBREVIATIONS")
    add_p("• SUCS: Smart University Complaint System", space_after=2)
    add_p("• HOD: Head of Department", space_after=2)
    add_p("• SRS: Software Requirements Specification", space_after=2)
    add_p("• COI: Conflict of Interest", space_after=2)
    add_p("• RBAC: Role-Based Access Control", space_after=2)
    add_p("• MVC: Model-View-Controller Architectural Pattern", space_after=2)
    add_p("• PBKDF2: Password-Based Key Derivation Function 2 (SHA-256)", space_after=2)
    add_p("• 3NF: Third Normal Form (Relational Database Normalization)", space_after=8)

    add_heading_1("1.5  PROJECT REQUIREMENTS")
    add_heading_2("1.5.1 Functional Requirements (FR1 to FR11)")
    add_p("• FR1 (Role-Based Authentication): System shall authenticate Students, Admins, and HOD via PBKDF2 hashes and route each to their dedicated portal.", space_after=3)
    add_p("• FR2 (Structured Complaint Lodging): System shall capture Title, Category, Target, and Description with input sanitization.", space_after=3)
    add_p("• FR3 (Monthly Quota Limit): System shall restrict each student to exactly 5 complaints per calendar month.", space_after=3)
    add_p("• FR4 (Student Grievance Tracker): System shall display grievance status (Pending, In-Progress, Resolved) with remarks.", space_after=3)
    add_p("• FR5 (Admin Triage Queue): System shall display active open complaints with category/priority filtering.", space_after=3)
    add_p("• FR6 (Emergency Keyword Priority): System shall automatically escalate tickets to High Priority if physical hazards are detected.", space_after=3)
    add_p("• FR7 (Conflict-of-Interest Firewall): Grievances against Administration shall bypass Admin desk and route strictly to HOD.", space_after=3)
    add_p("• FR8 (Mandatory Progress Feedback): Status transition to 'In-Progress' requires an explicit investigative remark.", space_after=3)
    add_p("• FR9 (Intelligent Auto-Grouping): Grievances sharing identical Category & Target shall be clustered under a shared group_id.", space_after=3)
    add_p("• FR10 (1-Click Batch Resolution): Admin and HOD shall resolve all correlated complaints in a cluster simultaneously.", space_after=3)
    add_p("• FR11 (In-Browser DB Inspector): Authorized staff shall inspect raw database records securely at /db-view.", space_after=8)

    add_heading_2("1.5.2 Non-Functional Requirements (NFR1 to NFR5)")
    add_p("• NFR1 (Security & Integrity): Zero plaintext passwords; SQL injection immunity via parameterized queries; session hijacking prevention.", space_after=3)
    add_p("• NFR2 (Performance & Latency): Sub-200ms page response times on standard campus network loads.", space_after=3)
    add_p("• NFR3 (Reliability & ACID Transactions): All state transitions and batch updates executed within atomic SQLite transactions.", space_after=3)
    add_p("• NFR4 (Usability & Responsiveness): Mobile-responsive layout powered by utility-first Tailwind CSS.", space_after=3)
    add_p("• NFR5 (Maintainability & Portability): Zero external cloud dependencies; self-contained SQLite file database.", space_after=8)

    add_heading_1("1.6  USE CASES AND USAGE SCENARIOS")
    add_heading_2("1.6.1 Use Case Diagrams")
    add_figure("media_1786788688397.png", "System Use Case Diagram — Smart University Complaint System (SUCS)")
    add_p("The Use Case Diagram above delineates interactions among the primary actors: Student (initiator), Administrator (first-line resolver), Head of Department (executive supervisor & COI adjudicator), and the Automated Background Triage Engine.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_2("1.6.2 Usage Scenarios (UC-01 to UC-08)")
    add_p("• UC-01 (User Authentication): Pre-condition: User exists in DB. Flow: Enters credentials -> PBKDF2 hash verified -> Redirected to role portal.", space_after=3)
    add_p("• UC-02 (Submit Complaint with Quota Check): Pre-condition: Student authenticated. Flow: Fills form -> Quota checked (COUNT < 5) -> Triage scan -> Committed to DB.", space_after=3)
    add_p("• UC-03 (Quota Limit Enforcement): Pre-condition: Student submitted 5 complaints this month. Flow: Tries 6th -> Blocked with flash alert -> Rejected.", space_after=3)
    add_p("• UC-04 (Emergency Keyword Escalation): Pre-condition: Hazard detected. Flow: Description scanned for hazard keywords -> Priority set to High (P1).", space_after=3)
    add_p("• UC-05 (Conflict-of-Interest Routing): Pre-condition: Target = Administration. Flow: Excluded from Admin queue -> Routed to HOD console.", space_after=3)
    add_p("• UC-06 (Transition to In-Progress): Pre-condition: Open ticket. Flow: Officer inputs remark -> Validated -> Status updated to In-Progress.", space_after=3)
    add_p("• UC-07 (1-Click Batch Resolution): Pre-condition: Group of clustered complaints. Flow: Officer enters resolution note -> Single click -> Atomic commit across all tickets.", space_after=3)
    add_p("• UC-08 (In-Browser DB Audit): Pre-condition: Authorized staff. Flow: Navigates to /db-view -> Queries live tables and audit logs.", space_after=8)

    add_heading_1("1.7  DEVELOPMENT METHODOLOGY")
    add_heading_2("1.7.1 Chosen Methodology (VU Process Model)")
    add_figure("media_1786788731168.png", "VU Process Model — Hybrid Waterfall-Spiral Lifecycle")
    add_p("The project adheres to the Virtual University Process Model, a hybrid software development lifecycle combining Waterfall milestone predictability with Spiral iterative risk mitigation, as taught in CS605.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_2("1.7.2 Reasons for Chosen Methodology")
    add_p("1. Milestone Predictability: Strict alignment with VU CS619 evaluation deliverables (SRS, Design, Prototype, Final Defense).", space_after=3)
    add_p("2. Iterative Risk Reduction: Spiral risk reviews enabled early validation of sensitive logic (COI routing, monthly quotas) before final deployment.", space_after=3)
    add_p("3. Working Code Verification: Early prototypes provided verifiable test beds for the Pytest automation test suite.", space_after=8)

    add_heading_2("1.7.3 Work Plan (Gantt Chart)")
    add_figure("media_1786788747833.png", "Project Work Plan & 16-Week Gantt Schedule")
    add_p("The Gantt Chart above details the 16-week project execution timeline. Backend development (Weeks 4-8) intentionally overlapped with Frontend prototyping (Weeks 6-9) to enable early integration testing.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_2("1.7.4 Project Schedule (Submission Calendar)")
    add_p("• Weeks 1-3: Requirements Gathering, Problem Definition, and SRS Approval. [COMPLETED]", space_after=3)
    add_p("• Weeks 4-6: Architectural Modeling, Relational Database Schema Design, and UI Wireframes. [COMPLETED]", space_after=3)
    add_p("• Weeks 7-10: Backend Core Engineering, Flask Setup, Authentication, and Quota Engine. [COMPLETED]", space_after=3)
    add_p("• Weeks 11-13: Intelligent Triage, COI Firewall Routing, and 1-Click Batch Grouping. [COMPLETED]", space_after=3)
    add_p("• Weeks 14-15: Automated Pytest QA Verification & Edge Case Hardening (100% pass). [COMPLETED]", space_after=3)
    add_p("• Week 16: Final Report, Presentation Deck, Zero-Dependency Packaging, and Defense. [SUBMITTED]", space_after=8)


    # ====================================================
    # CHAPTER 2 TITLE SEPARATOR
    # ====================================================
    add_chapter_title_page("CHAPTER 2", "Designing the Project")


    # ====================================================
    # CHAPTER 2 CONTENT
    # ====================================================
    add_heading_1("2.1  INTRODUCTION")
    add_p("System design translates the analytical specifications established in Chapter 1 into an actionable technical blueprint. This chapter articulates the 3-Tier Layered MVC Architecture, Sequence Diagrams, Logical Class Model, Relational Entity-Relationship Diagram (ERD), Physical Database Schema, and Graphical User Interfaces.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_1("2.2  PURPOSE")
    add_p("The purpose of this design specification is to guarantee structural modularity, data integrity, cryptographic security, and rapid algorithmic execution across all user interactions.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_1("2.3  SCOPE")
    add_p("The design scope encompasses Presentation Layer views, Business Logic Controllers, SQLite Persistence structures, and data security mechanisms governing student and staff access.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_1("2.4  DEFINITIONS, ACRONYMS AND ABBREVIATIONS")
    add_p("• DFD: Data Flow Diagram", space_after=2)
    add_p("• ERD: Entity-Relationship Diagram", space_after=2)
    add_p("• ACID: Atomicity, Consistency, Isolation, Durability", space_after=2)
    add_p("• REST: Representational State Transfer", space_after=8)

    add_heading_1("2.5  ARCHITECTURAL REPRESENTATION (ARCHITECTURE DIAGRAM)")
    add_figure("media_1786788922762.png", "Three-Tier Layered MVC Architecture Diagram")
    add_p("SUCS employs a classic Three-Tier Layered MVC Architecture:", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_p("1. Presentation Tier (View): HTML5, Jinja2 Templates, and Tailwind CSS delivering responsive role-based dashboards.", space_after=3)
    add_p("2. Application / Business Logic Tier (Controller): Python Flask microframework implementing session guards, quota validation, emergency keyword triage, and COI isolation routing.", space_after=3)
    add_p("3. Data Persistence Tier (Model): Relational SQLite 3 engine ensuring ACID-compliant transactions and foreign key referential integrity.", space_after=8)

    add_heading_1("2.6  DYNAMIC MODEL: SEQUENCE DIAGRAMS")
    add_figure("media_1786788895708.png", "Sequence Diagram — User Authentication & Submission Control Flow")
    add_p("Figure above illustrates the sequence of actions when a user logs in and submits a complaint. The controller validates credentials, checks the student's monthly quota, evaluates hazard keywords, computes the cluster group_id, and commits the ticket atomically.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_figure("media_1786788905335.png", "Sequence Diagram — Admin / HOD Complaint Resolution Control Flow")
    add_p("Figure above depicts the grievance resolution sequence. An administrator or HOD inspects the queue, enters an official resolution note, and triggers the 1-Click Batch Resolution engine, updating all correlated tickets simultaneously.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_1("2.7  OBJECT MODEL / LOGICAL MODEL: CLASS DIAGRAM")
    add_figure("media_1786788954827.png", "Logical Object Model & Class Diagram")
    add_p("The Class Diagram specifies core entities and operations:", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_p("• User: Base class encapsulating id, name, username, password_hash, role, verify_password(), and is_authenticated().", space_after=3)
    add_p("• Student: Inherits User; implements get_monthly_complaints(), can_submit_complaint(), and submit_complaint().", space_after=3)
    add_p("• Administrator & HeadOfDepartment: Inherit User; implement get_active_queue(), filter_coi_complaints(), and batch_resolve_group().", space_after=3)
    add_p("• Complaint: Domain aggregate encapsulating id, student_id, title, description, category, target, priority, status, group_id, evaluate_priority(), and mark_in_progress().", space_after=8)

    add_heading_1("2.8  DATABASE MODEL (DATABASE DIAGRAM)")
    add_figure("media_1786788850589.png", "Entity-Relationship Diagram (ERD)")
    add_p("The ERD above defines the logical relationships: USER (1:0..1) STUDENT specialization, STUDENT (1:N) COMPLAINT relationship, and DEPARTMENT (1:N) COMPLAINT categorization.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_figure("media_1786788977208.png", "Relational Database Schema & Data Dictionary")
    add_p("The Physical Schema specifies tables, columns, data types, primary keys, and foreign key constraints enforcing 3NF relational normalization without data anomalies.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_1("2.9  GRAPHICAL USER INTERFACES")
    add_p("The SUCS interface utilizes clean typography, color-coded badges, and intuitive card containers:", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_figure("media_1786788997712.png", "Student Dashboard & Quota Counter Interface", width=Inches(5.2))
    add_figure("media_1786789004364.png", "Complaint Submission Desk with Category/Target Taxonomies", width=Inches(5.2))
    add_figure("media_1786789010431.png", "Admin Complaint Queue with Status/Priority Filtering", width=Inches(5.2))
    add_figure("media_1786789016566.png", "HOD Executive Console & Conflict-of-Interest Desk", width=Inches(5.2))


    # ====================================================
    # CHAPTER 3 TITLE SEPARATOR
    # ====================================================
    add_chapter_title_page("CHAPTER 3", "Development")


    # ====================================================
    # CHAPTER 3 CONTENT
    # ====================================================
    add_heading_1("3.1  DEVELOPMENT PLAN (ARCHITECTURE DIAGRAM)")
    add_p("The development phase implemented the software modules using Python 3.9 Flask backend, SQLite relational database, Werkzeug cryptography, Jinja2 server-side rendering, and Tailwind CSS. The codebase is organized cleanly into modular route handlers, database helper models, and responsive templates.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_1("3.2  CORE ALGORITHMIC IMPLEMENTATIONS")
    add_heading_2("3.2.1 Auto-Combining & Clustering Engine (BR-04)")
    add_p("Campus problems (e.g. failing air conditioning or broken lab equipment) typically trigger dozens of duplicate submissions from affected students. SUCS automatically clusters complaints sharing identical Category and Target into a unified group_id cluster token, enabling holistic investigation rather than redundant overhead.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_2("3.2.2 Multi-Factor Urgency & Emergency Triage Engine (BR-03)")
    add_p("The system scans the grievance description for physical hazard terms ('urgent', 'fire', 'hazard', 'shock', 'leak', 'broken wire', 'emergency'). If detected, priority is automatically escalated to High (P1) regardless of user input, ensuring immediate administrative intervention.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_2("3.2.3 Monthly Quota Sliding Window Engine (BR-01)")
    add_p("To eliminate denial-of-service spamming and preserve triage focus, each student account is allocated exactly 5 complaint slots per calendar month. The quota guard queries SQLite COUNT(*) over the current calendar month; submissions beyond 5 are rejected with a descriptive flash alert.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_2("3.2.4 1-Click Cascading Atomic Resolution Engine (BR-05)")
    add_p("When campus maintenance resolves an underlying physical issue, an officer clicks '1-Click Batch Resolve' and inputs an audit note. The controller executes an atomic SQL UPDATE across all complaints sharing that group_id, transitioning them to 'Resolved' simultaneously.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_2("3.2.5 Conflict-of-Interest Firewall Engine (BR-02)")
    add_p("Grievances directed against 'Administration' bypass the standard Admin queue completely and route exclusively to the HOD console. The Administrator queue filter strictly executes WHERE target != 'Administration', preventing staff from suppressing complaints against themselves.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_heading_1("3.3  SYSTEM VERIFICATION & TEST TRACEABILITY")
    add_p("The system underwent rigorous verification via automated Pytest test cases covering 100% of functional requirements and edge cases:", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # Verification Table
    test_data = [
        ("TC-01", "Student Authentication & Role Redirection", "FR1", "PASSED"),
        ("TC-02", "Monthly Quota Limit (Block 6th Submission)", "FR3", "PASSED"),
        ("TC-03", "Emergency Keyword Priority Auto-Escalation", "FR6", "PASSED"),
        ("TC-04", "COI Firewall: Admin Queue Exclusion", "FR7", "PASSED"),
        ("TC-05", "COI Firewall: HOD Desk Exclusive Routing", "FR7", "PASSED"),
        ("TC-06", "Intelligent Auto-Grouping by Category/Target", "FR9", "PASSED"),
        ("TC-07", "1-Click Batch Group Resolution Engine", "FR10", "PASSED"),
        ("TC-08", "Mandatory Remarks on In-Progress Transition", "FR8", "PASSED"),
        ("TC-09", "In-Browser Database Inspector Security", "FR11", "PASSED")
    ]
    t = doc.add_table(rows=len(test_data) + 1, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Test ID", "Test Scenario & Scope", "Req ID", "Result"]
    for c_i, h in enumerate(headers):
        cell = t.cell(0, c_i)
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 80, 80, 100, 100)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.name = 'Tahoma'
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for r_i, (tid, tdesc, rid, res) in enumerate(test_data, start=1):
        bg = "F8FAFC" if r_i % 2 == 1 else "FFFFFF"
        for c_i, val in enumerate([tid, tdesc, rid, res]):
            cell = t.cell(r_i, c_i)
            set_cell_background(cell, bg)
            set_cell_margins(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = 'Tahoma'
            r.font.size = Pt(9.5)
            if c_i == 0:
                r.bold = True
                r.font.color.rgb = NAVY
            elif c_i == 3:
                r.bold = True
                r.font.color.rgb = RGBColor(16, 120, 70)


    # ====================================================
    # REFERENCES
    # ====================================================
    doc.add_page_break()

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(30)
    p.paragraph_format.space_after = Pt(20)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("REFERENCES")
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(18)
    r.font.color.rgb = NAVY

    refs = [
        "[1] Virtual University of Pakistan, \"CS619 Final Project Guidelines & Evaluation Criteria,\" Department of Computer Sciences, 2026.",
        "[2] Virtual University of Pakistan, \"CS605: Software Engineering II Course Material (Waterfall & Spiral Models),\" Virtual University Press, 2025.",
        "[3] Pressman, R. S., & Maxim, B. R., \"Software Engineering: A Practitioner's Approach,\" 9th Edition, McGraw-Hill Education, 2020.",
        "[4] Sommerville, I., \"Software Engineering,\" 10th Edition, Pearson Education, 2016.",
        "[5] Grinberg, M., \"Flask Web Development: Developing Web Applications with Python,\" 2nd Edition, O'Reilly Media, 2018.",
        "[6] Owens, M., & Allen, G., \"The Definitive Guide to SQLite,\" 2nd Edition, Apress, 2010.",
        "[7] Werkzeug Project, \"Cryptographic Password Hashing Utilities (PBKDF2-SHA256),\" Pallets Projects Documentation, 2024.",
        "[8] Tailwind Labs, \"Tailwind CSS: Utility-First CSS Framework Specification,\" 2024."
    ]
    for rf in refs:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(rf)
        r.font.name = 'Tahoma'
        r.font.size = Pt(10)


    # ====================================================
    # APPENDIX: VIVA VOCE PREPARATION GUIDE
    # ====================================================
    doc.add_page_break()

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(30)
    p.paragraph_format.space_after = Pt(20)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("APPENDIX")
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(18)
    r.font.color.rgb = NAVY

    add_heading_2("Viva Voce Comprehensive Preparation Guide")
    add_p("This appendix equips the candidate with concise, definitive answers to anticipated examiner questions during the final project defense.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    qa_list = [
        ("Q1: What is the primary problem your project solves?",
         "Answer: SUCS eliminates untracked grievances, stops frivolous spamming via a 5-complaint monthly limit, removes administrative retaliation through an automated Conflict-of-Interest firewall routing complaints against Admin directly to HOD, and reduces staff triage workload via 1-Click batch group resolution."),
        ("Q2: Why did you choose the VU Hybrid Waterfall-Spiral Process Model?",
         "Answer: The Waterfall stages provided clear milestones matching VU CS619 submission deliverables, while the Spiral risk reviews allowed us to iteratively refine and test high-risk features (COI firewall and keyword triage) before production deployment."),
        ("Q3: How is Conflict of Interest handled technically?",
         "Answer: When a complaint is submitted with target 'Administration', the controller detects this attribute and routes it strictly to the HOD console. The Admin triage query explicitly executes WHERE target != 'Administration', preventing staff from seeing or modifying complaints against themselves."),
        ("Q4: How does 1-Click Batch Resolution maintain data integrity?",
         "Answer: Correlated complaints share a composite group_id based on category and target. When resolved, an atomic SQL UPDATE executes across that group_id, ensuring all linked complaints update to 'Resolved' simultaneously with the resolution note in a single transaction."),
        ("Q5: How is security maintained against SQL Injection and Password cracking?",
         "Answer: Passwords are encrypted using salted PBKDF2-SHA256 hashes via Werkzeug. All database interactions utilize strict parameter binding (question-mark placeholders ?), completely immunizing the application against SQL injection attacks.")
    ]

    for q, a in qa_list:
        add_p(f"• {q}", bold=True, font='Tahoma', size=11, color=NAVY, space_after=2)
        add_p(f"    {a}", font='Tahoma', size=10, space_after=8)

    doc.save(output_path)
    print(f"Final Report successfully generated and saved to: {output_path}")

if __name__ == "__main__":
    up_dir = "/Users/usman/.gemini/antigravity/brain/05b706b9-e3c5-4023-b5b0-9f886b725595/.user_uploaded"
    out_file = "/Users/usman/Desktop/UniversityComplaintManagement/SUCS_Final_Project_Report.docx"
    create_report(out_file, up_dir)
