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

def create_report():
    doc = docx.Document()
    
    # Page setup - 1 inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Base styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)

    # Helper functions
    def add_title(text, size=24, space_after=12, color="003366"):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(space_after)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor(int(color[:2], 16), int(color[2:4], 16), int(color[4:], 16))
        return p

    def add_h1(text, space_before=18, space_after=8):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(0, 51, 102) # Navy Blue
        return p

    def add_h2(text, space_before=14, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(30, 80, 130)
        return p

    def add_h3(text, space_before=10, space_after=4):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(60, 60, 60)
        return p

    updir = "/Users/usman/.gemini/antigravity/brain/05b706b9-e3c5-4023-b5b0-9f886b725595/.user_uploaded"
    def add_figure(img_filename, caption, width=Inches(6.0)):
        img_path = os.path.join(updir, img_filename)
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(4)
            run = p_img.add_run()
            run.add_picture(img_path, width=width)
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_after = Pt(12)
            run_cap = p_cap.add_run(f"Figure: {caption}")
            run_cap.italic = True
            run_cap.font.size = Pt(10)
            run_cap.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    def add_p(text, space_after=6, line_spacing=1.15):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        run = p.add_run(text)
        return p

    def add_bullet(text, bold_prefix="", space_after=4):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            r1 = p.add_run(bold_prefix)
            r1.bold = True
        p.add_run(text)
        return p

    def add_callout(text, title="NOTE / ARCHITECTURAL SPECIFICATION"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        set_cell_background(cell, "F0F4F8")
        set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(4)
        r_title = p.add_run(f"📌 {title}: ")
        r_title.bold = True
        r_title.font.size = Pt(10.5)
        r_title.font.color.rgb = RGBColor(0, 51, 102)
        r_text = p.add_run(text)
        r_text.font.size = Pt(10.5)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # =========================================================================
    # 1. TITLE PAGE
    # =========================================================================
    p_top = doc.add_paragraph()
    p_top.paragraph_format.space_before = Pt(36)
    p_top.paragraph_format.space_after = Pt(12)
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_top = p_top.add_run("FINAL PROJECT REPORT")
    r_top.bold = True
    r_top.font.size = Pt(18)
    r_top.font.color.rgb = RGBColor(100, 100, 100)

    add_title("SMART UNIVERSITY COMPLAINT SYSTEM (SUCS)", size=24, space_after=18, color="003366")

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(36)
    r_sub = p_sub.add_run("An Intelligent Three-Tier Web Platform for Automated Academic Triage, Emergency Hazard Prioritization, Conflict-of-Interest Routing, and 1-Click Cascading Batch Resolution")
    r_sub.font.size = Pt(12)
    r_sub.italic = True
    r_sub.font.color.rgb = RGBColor(80, 80, 80)

    p_sup = doc.add_paragraph()
    p_sup.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sup.paragraph_format.space_after = Pt(24)
    r_sup_title = p_sup.add_run("Project Supervisor:\n")
    r_sup_title.font.size = Pt(12)
    r_sup_title.bold = True
    r_sup_val = p_sup.add_run("<<Project Supervisor Name>>\nSenior Lecturer / Project Supervisor\nDepartment of Computer Sciences")
    r_sup_val.font.size = Pt(11)

    p_by = doc.add_paragraph()
    p_by.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_by.paragraph_format.space_after = Pt(36)
    r_by_title = p_by.add_run("Submitted By:\n")
    r_by_title.font.size = Pt(12)
    r_by_title.bold = True
    r_by_val = p_by.add_run("Project Group ID: <<Project Group ID>>\n\n<<Group Member Name>> (<<VU ID>>)\n<<Second Member Name>> (<<VU ID>>)")
    r_by_val.font.size = Pt(11)

    p_uni = doc.add_paragraph()
    p_uni.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_uni.paragraph_format.space_before = Pt(48)
    p_uni.paragraph_format.space_after = Pt(0)
    r_uni = p_uni.add_run("Software Projects & Research Section\nDepartment of Computer Sciences\nVirtual University of Pakistan")
    r_uni.bold = True
    r_uni.font.size = Pt(13)
    r_uni.font.color.rgb = RGBColor(0, 51, 102)

    doc.add_page_break()

    # =========================================================================
    # 2. CERTIFICATE
    # =========================================================================
    add_title("CERTIFICATE", size=18, space_after=18)
    add_p("This is to certify that <<First Member Name>> (<<VU ID>>) and <<Second Member Name>> (<<VU ID>>) have worked on and successfully completed their Software Final Project titled \"Smart University Complaint System (SUCS)\" at the Software Projects & Research Section, Department of Computer Sciences, Virtual University of Pakistan, in partial fulfillment of the requirements for the degree of Bachelor of Science in Computer Science (BSCS) under my guidance and supervision.")
    add_p("In our opinion, the software implementation, technical architecture, and documentation are satisfactory, robust, and up to the standard mark, fulfilling all professional and academic requirements prescribed by the Virtual University of Pakistan.")
    
    doc.add_paragraph().paragraph_format.space_after = Pt(24)

    # Signature blocks table
    sig_tbl = doc.add_table(rows=3, cols=2)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_tbl.cell(0, 0).width = Inches(3.2)
    sig_tbl.cell(0, 1).width = Inches(3.2)
    
    sig_tbl.cell(0, 0).paragraphs[0].add_run("_____________________________\nSupervisor / Internal Examiner\n<<Project Supervisor Name>>\nSoftware Projects & Research Section\nVirtual University of Pakistan").font.size = Pt(10.5)
    sig_tbl.cell(0, 1).paragraphs[0].add_run("_____________________________\nExternal Examiner / Subject Specialist\n<<External Examiner Name>>\nDepartment of Computer Science\nVirtual University of Pakistan").font.size = Pt(10.5)
    
    sig_tbl.cell(2, 0).paragraphs[0].add_run("Dated: ___________________").font.size = Pt(10.5)
    sig_tbl.cell(2, 1).paragraphs[0].add_run("Accepted By: ___________________\n(For Official Use Only)").font.size = Pt(10.5)

    doc.add_page_break()

    # =========================================================================
    # 3. EXORDIUM, DEDICATION, ACKNOWLEDGEMENT, PREFACE
    # =========================================================================
    add_title("EXORDIUM", size=18, space_after=14)
    p_ex = doc.add_paragraph()
    p_ex.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ex.paragraph_format.space_after = Pt(18)
    r_ex = p_ex.add_run("In the name of Allah, the Most Compassionate, the Most Merciful.\n\nPraise be to Allah, Lord of the Universe,\nThe Compassionate, the Merciful, King of the Day of Judgment!\nYou alone we worship, and to You alone we pray for help,\nGuide us to the straight path—\nThe path of those upon whom You have bestowed Your favor,\nNot of those who have incurred Your wrath, nor of those who have gone astray.\n(Surah Al-Fatihah 1:1-7)")
    r_ex.font.size = Pt(11.5)
    r_ex.italic = True

    add_title("DEDICATION", size=16, space_after=12)
    p_ded = add_p("This project is affectionately dedicated to our beloved parents, whose relentless prayers, moral sacrifices, and boundless encouragement have illuminated our educational journey. We also dedicate this work to our respected teachers and mentors at Virtual University of Pakistan, who nurtured our passion for software engineering and empowered us to build technological solutions that serve our academic community.")
    p_ded.paragraph_format.space_after = Pt(18)

    add_title("ACKNOWLEDGEMENT", size=16, space_after=12)
    add_p("First and foremost, all praise and deepest gratitude belong to Almighty Allah, the Most Gracious and Merciful, who granted us the intellect, patience, and good health required to bring this software engineering project to fruition.")
    add_p("We extend our heartfelt gratitude and professional appreciation to our respected Project Supervisor, <<Project Supervisor Name>>, for invaluable technical feedback, structural guidance, and meticulous reviews throughout the software development lifecycle.")
    add_p("We also thank our faculty members at the Department of Computer Sciences, Virtual University of Pakistan, for establishing rigorous standards in Software Engineering (CS605), Database Systems (CS403), and Web Development that made the design and deployment of this high-reliability system possible.")

    add_title("PREFACE", size=16, space_after=12)
    add_p("The modern university campus encompasses thousands of active students, faculty members, and administrative staff across multiple geographically distributed faculties. Traditional manual complaint procedures—relying on paper registers, informal emails, or unsegregated portal queues—frequently suffer from communication breakdowns, unresolved delays, duplication, and critical conflicts of interest.")
    add_p("The Smart University Complaint System (SUCS) was conceived to solve these challenges through a lightweight, high-performance, and verifiable Three-Tier web application. This project document presents the complete architectural, design, database, and algorithmic implementation of SUCS, detailing our requirements analysis, design models, business rules catalog, and comprehensive test suite.")

    doc.add_page_break()

    # =========================================================================
    # 4. TABLE OF CONTENTS
    # =========================================================================
    add_title("TABLE OF CONTENTS", size=18, space_after=18)
    
    toc_data = [
        ("CERTIFICATE", "ii"),
        ("EXORDIUM & DEDICATION", "iii"),
        ("ACKNOWLEDGEMENT & PREFACE", "iv"),
        ("TABLE OF CONTENTS", "v"),
        ("CHAPTER 1: GATHERING & ANALYZING INFORMATION", "1"),
        ("  1.1 Introduction", "1"),
        ("  1.2 Project Purpose", "2"),
        ("  1.3 Scope of the System", "2"),
        ("  1.4 Definitions, Acronyms, and Abbreviations", "3"),
        ("  1.5 Project Requirements Specification", "4"),
        ("      1.5.1 Functional Requirements (FR1 to FR11)", "4"),
        ("      1.5.2 Non-Functional Requirements (NFR1 to NFR5)", "7"),
        ("  1.6 Use Cases & Usage Scenarios", "8"),
        ("      1.6.1 Actor & Use Case Specifications", "8"),
        ("      1.6.2 Usage Scenarios (UC-01 to UC-08)", "9"),
        ("  1.7 Development Methodology", "13"),
        ("      1.7.1 Chosen Methodology: VU Process Model", "13"),
        ("      1.7.2 Justification of Methodology", "14"),
        ("      1.7.3 Work Plan & Gantt Milestones", "15"),
        ("      1.7.4 Project Submission Calendar", "16"),
        ("CHAPTER 2: DESIGNING THE PROJECT", "17"),
        ("  2.1 Introduction", "17"),
        ("  2.2 Purpose & Scope of Design Phase", "17"),
        ("  2.3 Architectural Representation (3-Tier MVC Architecture)", "18"),
        ("  2.4 Dynamic Model: Sequence Diagrams", "20"),
        ("      2.4.1 Complaint Submission & Auto-Grouping Sequence", "20"),
        ("      2.4.2 Admin In-Progress & 1-Click Resolution Sequence", "21"),
        ("      2.4.3 HOD Conflict-of-Interest Executive Sequence", "22"),
        ("  2.5 Logical & Object Model: Class Diagram", "23"),
        ("  2.6 Relational Database Model & Entity-Relationship Diagram", "25"),
        ("      2.6.1 Entity-Relationship Model (ERD)", "25"),
        ("      2.6.2 Relational Schema & Data Dictionary", "26"),
        ("  2.7 Graphical User Interface (GUI) Design Walkthrough", "28"),
        ("CHAPTER 3: DEVELOPMENT & IMPLEMENTATION", "32"),
        ("  3.1 Development Architecture & Tech Stack Justification", "32"),
        ("  3.2 Core Algorithmic Implementations", "34"),
        ("      3.2.1 Auto-Combining & Clustering Algorithm", "34"),
        ("      3.2.2 Multi-Factor Urgency & Emergency Triage Engine", "35"),
        ("      3.2.3 Monthly Quota Sliding Window Rate Limiter", "36"),
        ("      3.2.4 1-Click Cascading Atomic Resolution Engine", "37"),
        ("  3.3 Automated Verification & Test Suite Traceability", "38"),
        ("      3.3.1 Pytest Suite Execution Matrix (TC-01 to TC-09)", "38"),
        ("      3.3.2 Zero-Extension In-Browser Database Inspector", "41"),
        ("REFERENCES & BIBLIOGRAPHY", "42"),
        ("APPENDIX: DEPLOYMENT GUIDE & VIVA CHEAT SHEET", "43"),
    ]

    toc_tbl = doc.add_table(rows=len(toc_data), cols=2)
    toc_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    toc_tbl.cell(0, 0).width = Inches(5.5)
    toc_tbl.cell(0, 1).width = Inches(1.0)

    for i, (title, page) in enumerate(toc_data):
        row = toc_tbl.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        p0 = c0.paragraphs[0]
        p1 = c1.paragraphs[0]
        p0.paragraph_format.space_after = Pt(2)
        p1.paragraph_format.space_after = Pt(2)
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        r0 = p0.add_run(title)
        r1 = p1.add_run(page)
        r1.font.size = Pt(10)
        r1.font.color.rgb = RGBColor(100, 100, 100)

        if "CHAPTER" in title:
            r0.bold = True
            r0.font.size = Pt(11)
            r0.font.color.rgb = RGBColor(0, 51, 102)
            r1.bold = True
        elif title.startswith("      "):
            r0.font.size = Pt(9.5)
            r0.font.color.rgb = RGBColor(90, 90, 90)
        elif title.startswith("  "):
            r0.font.size = Pt(10)
            r0.bold = True
            r0.font.color.rgb = RGBColor(40, 40, 40)
        else:
            r0.bold = True
            r0.font.size = Pt(10.5)

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 1: GATHERING & ANALYZING INFO
    # =========================================================================
    add_title("CHAPTER 1", size=20, space_after=6, color="003366")
    add_title("GATHERING & ANALYZING INFORMATION", size=16, space_after=18, color="1E5082")

    add_h1("1.1 Introduction")
    add_p("The Smart University Complaint System (SUCS) is an advanced, automated web-based platform engineered to modernize grievance handling, academic dispute resolution, and facility management within higher education institutions. In traditional university environments, complaints filed by students frequently encounter administrative bottlenecks, lack of transparent status tracking, redundant ticket handling, and vulnerability to bias when grievances involve administrative personnel.")
    add_p("SUCS establishes an automated, auditable, and role-segregated environment that streamlines the entire grievance lifecycle. The system implements automated multi-factor urgency sorting, intelligent clustering of duplicate or related complaints, an automated conflict-of-interest firewall, and 1-click batch resolutions for both Administrators and Department Heads (HOD).")

    add_h1("1.2 Purpose")
    add_p("The primary purpose of this project is to provide Virtual University of Pakistan and higher educational institutes with an automated, tamper-resistant complaint platform that achieves the following goals:")
    add_bullet(" To empower enrolled students with an intuitive self-service portal to lodge grievances, track review stages in real time, and inspect official responses.", "Student Empowerment:")
    add_bullet(" To eliminate queue chaos through an automated urgency algorithm that positions critical life-safety emergency hazards at the top of the queue in high-contrast red.", "Urgency Prioritization:")
    add_bullet(" To automatically cluster related tickets sharing common targets or root causes, enabling administrators to resolve dozens of matching complaints in a single click.", "Automated Smart Clustering:")
    add_bullet(" To enforce unbiased reporting by automatically detecting complaints directed against administrative personnel and routing them directly to the HOD Executive Desk.", "Conflict-of-Interest Segregation:")
    add_bullet(" To provide institutional leadership with real-time KPI metrics, highest-reason category distributions, and diagnostic visibility.", "Executive Oversight:")

    add_h1("1.3 Scope")
    add_p("The operational scope of SUCS encompasses three primary authenticated user roles operating across a centralized relational database architecture:")
    add_bullet(" Enrolled students submitting complaints across pre-defined or custom categories, tracking active issues, inspecting official feedback, and operating within a calibrated 5-complaint monthly quota.", "Student Domain:")
    add_bullet(" Administrative officers managing complaint triage, setting investigated tickets to 'In-Progress' with mandatory progress notes, escalating complex issues, auto-combining related tickets, and performing 1-click bulk closures.", "Administrative Domain:")
    add_bullet(" Department heads possessing executive jurisdiction over conflict-of-interest allegations against administration staff, reviewing escalated tickets, analyzing campus-wide category trends, and issuing binding resolutions.", "Executive (HOD) Domain:")
    add_bullet(" Includes an integrated, zero-extension in-browser SQLite inspector allowing external examiners, evaluators, and system auditors to verify raw relational rows without desktop database software.", "Diagnostic & Auditing Domain:")

    add_h1("1.4 Definitions, Acronyms, and Abbreviations")
    
    def_data = [
        ("SUCS", "Smart University Complaint System"),
        ("VU", "Virtual University of Pakistan"),
        ("RBAC", "Role-Based Access Control (Student, Admin, HOD permissions)"),
        ("COI", "Conflict of Interest (Grievances filed against Administrative staff)"),
        ("SRS", "Software Requirements Specification"),
        ("NFR", "Non-Functional Requirement"),
        ("ERD", "Entity-Relationship Diagram"),
        ("DDL", "Data Definition Language"),
        ("PBKDF2", "Password-Based Key Derivation Function 2 (SHA-256 salted hashing)"),
        ("MVC", "Model-View-Controller Architectural Pattern"),
        ("FIFO", "First-In, First-Out queue processing"),
        ("KPI", "Key Performance Indicator (Open, In-Progress, Solved counts)"),
        ("DOM", "Document Object Model (Browser client-side data tree)"),
        ("SQLite", "Self-contained, serverless zero-configuration relational database engine")
    ]
    
    def_tbl = doc.add_table(rows=len(def_data)+1, cols=2)
    def_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    def_tbl.cell(0, 0).width = Inches(1.8)
    def_tbl.cell(0, 1).width = Inches(4.7)
    
    # Header row
    h0, h1 = def_tbl.rows[0].cells[0], def_tbl.rows[0].cells[1]
    set_cell_background(h0, "003366")
    set_cell_background(h1, "003366")
    r_h0 = h0.paragraphs[0].add_run("Term / Acronym")
    r_h0.bold = True
    r_h0.font.color.rgb = RGBColor(255, 255, 255)
    r_h1 = h1.paragraphs[0].add_run("Definition & Architectural Role")
    r_h1.bold = True
    r_h1.font.color.rgb = RGBColor(255, 255, 255)

    for i, (k, v) in enumerate(def_data):
        row = def_tbl.rows[i+1]
        c0, c1 = row.cells[0], row.cells[1]
        if i % 2 == 1:
            set_cell_background(c0, "F8F9FA")
            set_cell_background(c1, "F8F9FA")
        r0 = c0.paragraphs[0].add_run(k)
        r0.bold = True
        r0.font.size = Pt(10)
        r1 = c1.paragraphs[0].add_run(v)
        r1.font.size = Pt(10)

    add_h1("1.5 Project Requirements Specification")
    add_h2("1.5.1 Functional Requirements (FR1 to FR11)")
    
    fr_list = [
        ("FR1: Multi-Role Authentication & Session Guard", "The system shall authenticate Students, Administrators, and HODs using cryptographically signed HTTP cookies. Passwords must be hashed using PBKDF2:SHA256, strictly rejecting plain-text authentication."),
        ("FR2: Complaint Submission with Custom Expansion", "Students shall submit complaints with mandatory Category, Directed Against entity, and Description. Dropdowns must support an 'Others' option that reveals custom text fields, ensuring complete input freedom."),
        ("FR3: Monthly Quota Engine (Rate Limiting)", "Each student account shall be restricted to a strict limit of 5 complaints per calendar month, calculated dynamically using an SQL sliding window (YYYY-MM). Submissions beyond 5 must be rejected."),
        ("FR4: Multi-Factor Urgency Triage", "Complaints shall be categorized automatically into Emergency, High, Medium, or Low priority based on explicit checkboxes, category sensitivity (Harassment, Infrastructure), and emergency keywords ('hazard', 'fire')."),
        ("FR5: Explicit 1-Click Emergency Hazard Alerting", "Students can check an Emergency Hazard toggle (or trigger 1-Click Hazard Fill) that immediately tags the issue as 'Emergency', rendering it with prominent red focus rings and bounce animations at the top of the queue."),
        ("FR6: Automated Conflict-of-Interest (COI) Firewall", "Any complaint where DirectedAgainst mentions 'Admin' (case-insensitive) shall automatically bypass the Admin queue and land directly into the HOD Executive Desk, preventing suppression or bias."),
        ("FR7: Automated Smart Clustering & Grouping", "The system shall continuously scan active complaints and automatically cluster tickets sharing the same Category and Directed Against entity under a common Master Group ID."),
        ("FR8: 1-Click Cascading Batch Resolution", "Both Administrators and HODs shall have access to a 1-Click Resolve button on grouped tickets that atomically cascades Status='Solved' and an official resolution comment to all linked complaints in a single SQL operation."),
        ("FR9: In-Progress Status with Mandatory Comments", "Administrators can transition tickets from 'Pending' to 'In-Progress'. The system must mandate an administrative progress note before accepting the update to keep students informed."),
        ("FR10: HOD Executive Analytics & Reason Ranking", "The HOD dashboard shall display live metrics of tickets currently held by Admins, rank highest-reason complaint categories with visual progress bars, and display urgency distribution charts."),
        ("FR11: In-Browser Zero-Extension Database Inspector", "The system shall expose a route (/db-view) rendering live relational tables of Users (with hashed passwords) and Complaints, allowing evaluators to verify DB state without external GUI tools.")
    ]

    for code, desc in fr_list:
        add_bullet(f" {desc}", bold_prefix=code + ": ")

    add_h2("1.5.2 Non-Functional Requirements (NFR1 to NFR5)")
    add_bullet(" Passwords stored using salted PBKDF2 hashes; signed session cookies prevent session tampering; strict server-side RBAC decorators block unauthorized URL tampering.", bold_prefix="NFR1 - Security: ")
    add_bullet(" Lightweight Flask architecture executes queries in <15 milliseconds; zero client-side framework bloat ensures instant page rendering even on constrained campus networks.", bold_prefix="NFR2 - Performance & Latency: ")
    add_bullet(" ACID-compliant SQLite relational database transactions guarantee atomic updates during cascading 1-click group resolutions.", bold_prefix="NFR3 - Reliability & Integrity: ")
    add_bullet(" Responsive, mobile-first Tailwind CSS UI; dedicated student complaints history page; dark-mode database inspector; high-contrast red emergency banners.", bold_prefix="NFR4 - Usability & Accessibility: ")
    add_bullet(" Modular three-tier codebase with comprehensive line-by-line docstrings, passing 100% of automated Pytest unit and integration test suites.", bold_prefix="NFR5 - Maintainability: ")

    add_h1("1.6 Use Cases & Usage Scenarios")
    add_figure("media_1786788688397.png", "Use Case Diagram — Smart University Complaint System (SUCS)")
    add_h2("1.6.1 Actor & Use Case Specifications")
    add_p("The system defines three authenticated human actors and one internal automated scheduler:")
    add_bullet(" Can authenticate, monitor monthly quota, lodge standard/emergency complaints, and view detailed status logs.", bold_prefix="Student: ")
    add_bullet(" Can review non-conflict queues, update status to 'In-Progress' with comments, forward complex cases, group duplicates, and execute 1-click batch resolutions.", bold_prefix="Administrator: ")
    add_bullet(" Adjudicates sensitive Conflict-of-Interest complaints, inspects escalated tickets, analyzes highest-reason complaint distributions, and issues executive verdicts.", bold_prefix="Head of Department (HOD): ")
    add_bullet(" Scans active complaints on submission and triage load, auto-clustering tickets sharing identical categories and accused entities.", bold_prefix="Clustering Engine: ")

    add_h2("1.6.2 Usage Scenarios (UC-01 to UC-08)")

    scenarios = [
        ("UC-01: User Authentication", "Student, Admin, HOD", "User enters credentials on unified login screen", "System validates hash via check_password_hash() and routes to role dashboard"),
        ("UC-02: Lodge Complaint with Quota Check", "Student", "Student fills category, target, description", "System verifies quota (used < 5), classifies urgency, checks for group matches, saves ticket"),
        ("UC-03: Emergency Hazard Submission", "Student", "Student selects Emergency Hazard checkbox", "System classifies urgency as 'Emergency', rendering ticket at the top in pulsing red"),
        ("UC-04: Automated Conflict-of-Interest Routing", "System / Student", "Student directs grievance against 'Admin Office'", "System skips Admin queue entirely and routes ticket directly to HOD Executive Desk"),
        ("UC-05: Update Status to In-Progress", "Administrator", "Admin clicks 'In-Progress' on active ticket", "System opens modal, enforces mandatory progress remark, and updates ticket status"),
        ("UC-06: 1-Click Cascading Group Resolution", "Admin / HOD", "User clicks '1-Click Solve' on auto-grouped tickets", "System prompts for resolution remarks, updates all linked complaints to 'Solved' in one atomic SQL query"),
        ("UC-07: Executive Analytics Inspection", "HOD", "HOD navigates to Executive Desk", "System computes admin ticket backlog count, renders ranked category distribution bars and urgency charts"),
        ("UC-08: Live Relational Database Inspection", "Evaluator / Auditor", "Evaluator navigates to /db-view in browser", "System renders live table data for User and Complaint tables with masked hash presentation")
    ]

    uc_tbl = doc.add_table(rows=len(scenarios)+1, cols=4)
    uc_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    uc_tbl.cell(0, 0).width = Inches(1.8)
    uc_tbl.cell(0, 1).width = Inches(1.1)
    uc_tbl.cell(0, 2).width = Inches(1.8)
    uc_tbl.cell(0, 3).width = Inches(1.8)

    h_cells = uc_tbl.rows[0].cells
    for c, text in zip(h_cells, ["Use Case ID & Title", "Primary Actor", "Triggering Condition", "System Response"]):
        set_cell_background(c, "003366")
        p = c.paragraphs[0]
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for i, (uid, actor, trig, resp) in enumerate(scenarios):
        row = uc_tbl.rows[i+1]
        if i % 2 == 1:
            for cell in row.cells:
                set_cell_background(cell, "F8F9FA")
        row.cells[0].paragraphs[0].add_run(uid).bold = True
        row.cells[0].paragraphs[0].runs[0].font.size = Pt(9)
        row.cells[1].paragraphs[0].add_run(actor).font.size = Pt(9)
        row.cells[2].paragraphs[0].add_run(trig).font.size = Pt(9)
        row.cells[3].paragraphs[0].add_run(resp).font.size = Pt(9)

    add_h1("1.7 Development Methodology")
    add_figure("media_1786788731168.png", "VU Process Model — Hybrid Waterfall-Spiral Lifecycle")
    add_h2("1.7.1 Chosen Methodology: VU Process Model")
    add_p("In strict alignment with the Virtual University of Pakistan Final Project guidelines (CS605 / CS619), this project adopted the VU Process Model. The VU Process Model is a specialized hybrid software engineering lifecycle that combines the disciplined phase-gate milestones of the Waterfall Model with the iterative refinement, prototyping, and risk analysis of the Spiral Model.")

    add_h2("1.7.2 Reasons for Chosen Methodology")
    add_bullet(" Structured milestones (SRS -> Design Document -> Prototype -> Final Deliverable) ensure that supervisor and examiner checkpoints are rigorously respected.", bold_prefix="Strict Academic Alignment: ")
    add_bullet(" Prototype iterations allowed us to identify critical real-world needs (e.g. adding the 'In-Progress' status with mandatory comments and 'Others' custom category options) without destabilizing baseline functionality.", bold_prefix="Iterative Feedback Incorporation: ")
    add_bullet(" Rapid prototyping in Python Flask enabled early empirical testing of the 1-click cascading SQL update and automated conflict firewall before final documentation sign-off.", bold_prefix="Risk Mitigation: ")

    add_h2("1.7.3 Work Plan (Gantt Chart & Submission Calendar)")
    add_figure("media_1786788747833.png", "Project Work Plan & 16-Week Gantt Schedule")
    
    milestones = [
        ("Phase 1: Project Proposal & Scope Definition", "Week 1 - 2", "100% Completed", "Formal acceptance of SUCS scope"),
        ("Phase 2: Requirements Gathering & SRS", "Week 3 - 5", "100% Completed", "Complete Functional & Non-Functional Catalog"),
        ("Phase 3: System Design & Architectural Models", "Week 6 - 8", "100% Completed", "Class Diagram, ERD, Sequence, 3-Tier Architecture"),
        ("Phase 4: Prototype Development & Schema Migration", "Week 9 - 12", "100% Completed", "Flask Backend, SQLite DDL, Auto-Grouping Engine"),
        ("Phase 5: Automated Testing & Pytest Verification", "Week 13 - 14", "100% Completed", "100% Pass Rate across 9 Unit & Integration Tests"),
        ("Phase 6: Final Documentation & Presentation Preparation", "Week 15 - 16", "100% Completed", "Final Project Report and Presentation Slides")
    ]

    m_tbl = doc.add_table(rows=len(milestones)+1, cols=4)
    m_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    m_tbl.cell(0, 0).width = Inches(2.6)
    m_tbl.cell(0, 1).width = Inches(1.1)
    m_tbl.cell(0, 2).width = Inches(1.2)
    m_tbl.cell(0, 3).width = Inches(1.6)

    for c, text in zip(m_tbl.rows[0].cells, ["Phase / Milestone Description", "Timeline", "Status", "Deliverable Output"]):
        set_cell_background(c, "003366")
        r = c.paragraphs[0].add_run(text)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for i, (m, t, s, o) in enumerate(milestones):
        row = m_tbl.rows[i+1]
        if i % 2 == 1:
            for cell in row.cells:
                set_cell_background(cell, "F8F9FA")
        row.cells[0].paragraphs[0].add_run(m).bold = True
        row.cells[0].paragraphs[0].runs[0].font.size = Pt(9)
        row.cells[1].paragraphs[0].add_run(t).font.size = Pt(9)
        row.cells[2].paragraphs[0].add_run(s).font.size = Pt(9)
        row.cells[3].paragraphs[0].add_run(o).font.size = Pt(9)

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 2: DESIGNING THE PROJECT
    # =========================================================================
    add_title("CHAPTER 2", size=20, space_after=6, color="003366")
    add_title("DESIGNING THE PROJECT", size=16, space_after=18, color="1E5082")

    add_h1("2.1 Introduction")
    add_p("The design phase translates the functional requirements established in Chapter 1 into formal software engineering models. This chapter details the architectural representation, sequence interactions, logical class models, database entity-relationship schema, and graphical user interface architectures.")

    add_h1("2.2 Architectural Representation (Three-Tier MVC Architecture)")
    add_figure("media_1786788922762.png", "Three-Tier Layered MVC Architecture Diagram")
    add_p("SUCS is designed upon an enterprise Three-Tier Client-Server Architecture, completely decoupling presentation, business logic processing, and persistent relational data storage:")
    
    add_bullet(" Developed using HTML5, modern utility-first Tailwind CSS, and Jinja2 server-rendered templates. Implements safe DOM parsing using hidden data containers (#cdata-<id>) to completely eradicate JavaScript string escaping vulnerabilities.", bold_prefix="1. Presentation Tier (Client View): ")
    add_bullet(" Built on Python 3 and the Flask micro-framework. Houses the Core Business Engine, including the Monthly Quota sliding window validator, Keyword Urgency Classifier, Conflict-of-Interest routing logic, and the Automated Clustering Engine.", bold_prefix="2. Application Logic Tier (Controller): ")
    add_bullet(" Structured on SQLite relational database (sucs_prototype.db). Enforces ACID transaction guarantees, Foreign Key referential integrity, and atomic bulk cascading updates during 1-click group resolutions.", bold_prefix="3. Data Storage Tier (Model): ")

    add_callout("Architecture Flow: Browser (HTTP Client) <---> Flask HTTP Router & Auth Decorators <---> Business Logic & Auto-Grouping Engine <---> SQLite Relational Database Engine (Foreign Keys Enforced).", title="3-TIER DATA PIPELINE")

    add_h1("2.3 Dynamic Model: Sequence Diagrams")
    add_figure("media_1786788895708.png", "Sequence Diagram — User Authentication & Submission Control Flow")
    add_figure("media_1786788905335.png", "Sequence Diagram — Admin/HOD Complaint Resolution Control Flow")
    add_h2("2.3.1 Complaint Submission & Auto-Grouping Sequence")
    add_p("This dynamic sequence model illustrates how a student lodges a grievance, how the system checks the monthly quota, and how the auto-combining engine automatically links similar tickets:")
    add_bullet("Student submits complaint form with Category, Target, Description, and Emergency toggle.", bold_prefix="Step 1: ")
    add_bullet("Flask route (/submit-complaint) extracts user_id from encrypted session cookie.", bold_prefix="Step 2: ")
    add_bullet("System queries DB: SELECT COUNT(*) WHERE StudentID=? AND YYYY-MM. If used >= 5, aborts with Quota Alert.", bold_prefix="Step 3: ")
    add_bullet("Urgency engine inspects Emergency checkbox and keywords; assigns Emergency, High, Medium, or Low.", bold_prefix="Step 4: ")
    add_bullet("INSERT INTO Complaint executed, generating primary key ID.", bold_prefix="Step 5: ")
    add_bullet("auto_group_similar_complaints() checks for active tickets with same Category & Target; clusters under shared GroupID.", bold_prefix="Step 6: ")
    add_bullet("System redirects student to Student Dashboard with success confirmation banner.", bold_prefix="Step 7: ")

    add_h2("2.3.2 Admin In-Progress & 1-Click Group Resolution Sequence")
    add_bullet("Administrator opens Admin Dashboard (/admin); backend pre-sorts queue by urgency.", bold_prefix="Step 1: ")
    add_bullet("Admin selects '⏳ In-Progress' on an active ticket. Modal opens requiring progress remark.", bold_prefix="Step 2: ")
    add_bullet("Admin submits remark; backend verifies non-empty comment, updating Status='In-Progress'.", bold_prefix="Step 3: ")
    add_bullet("For grouped complaints, Admin clicks '⚡ 1-Click Solve All Related' on the group banner.", bold_prefix="Step 4: ")
    add_bullet("Backend executes atomic SQL update: UPDATE Complaint SET Status='Solved', AdminComment=? WHERE GroupID=? OR ID=?.", bold_prefix="Step 5: ")
    add_bullet("All matching complaints immediately close with status 'Solved' across student dashboards.", bold_prefix="Step 6: ")

    add_h2("2.3.3 HOD Conflict-of-Interest Executive Sequence")
    add_bullet("Student lodges complaint with DirectedAgainst = 'Admin Staff'.", bold_prefix="Step 1: ")
    add_bullet("Backend SQL filter on /admin excludes ticket: WHERE LOWER(DirectedAgainst) NOT LIKE '%admin%'.", bold_prefix="Step 2: ")
    add_bullet("HOD loads Executive Desk (/hod); query fetches ticket into Purple Conflict Queue.", bold_prefix="Step 3: ")
    add_bullet("HOD inspects full unredacted narrative in details modal and inputs executive verdict.", bold_prefix="Step 4: ")
    add_bullet("HOD clicks 'Mark Solved as HOD' (or '1-Click Solve Group'); ticket atomically closes.", bold_prefix="Step 5: ")

    add_h1("2.4 Logical & Object Model: Class Diagram")
    add_figure("media_1786788954827.png", "Logical Object Model & Class Diagram")
    add_p("The object model defines the core entities, attributes, methods, and relationships operating within the system:")

    class_data = [
        ("User", "ID: int (PK)\nUsername: string\nPasswordHash: string\nRole: enum('Student','Admin','HOD')", "login(pwd): bool\nlogout(): void\nget_role(): string", "Base entity for all authenticated actors"),
        ("Student (inherits User)", "QuotaUsed: int\nQuotaRemaining: int", "submit_complaint(): int\nview_history(): list\ncheck_quota(): dict", "Represents enrolled student with monthly rate-limiting"),
        ("Administrator (inherits User)", "Department: string", "view_queue(): list\nset_in_progress(id, c): void\nforward_hod(id): void\ngroup_tickets(ids): int\nresolve_group(gid, c): int", "Operational triage manager for non-conflict queue"),
        ("HOD (inherits User)", "ExecutiveJurisdiction: string", "view_conflict_queue(): list\nview_escalated_queue(): list\nget_analytics(): dict\nsolve_executive(id, c): void\nresolve_group_hod(gid, c): int", "Highest executive authority for campus disputes"),
        ("Complaint", "ID: int (PK)\nStudentID: int (FK)\nCategory: string\nDirectedAgainst: string\nDescription: text\nUrgencyLevel: enum\nStatus: enum\nDateSubmitted: datetime\nGroupID: int (Nullable)\nAdminComment: text", "classify_urgency(): string\nassign_group(gid): void\nmark_solved(c): void", "Core relational transactional entity")
    ]

    cls_tbl = doc.add_table(rows=len(class_data)+1, cols=4)
    cls_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cls_tbl.cell(0, 0).width = Inches(1.5)
    cls_tbl.cell(0, 1).width = Inches(1.8)
    cls_tbl.cell(0, 2).width = Inches(1.8)
    cls_tbl.cell(0, 3).width = Inches(1.4)

    for c, text in zip(cls_tbl.rows[0].cells, ["Class Name", "Attributes", "Methods", "Role & Relationship"]):
        set_cell_background(c, "003366")
        r = c.paragraphs[0].add_run(text)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for i, (name, attrs, meths, role) in enumerate(class_data):
        row = cls_tbl.rows[i+1]
        if i % 2 == 1:
            for cell in row.cells:
                set_cell_background(cell, "F8F9FA")
        row.cells[0].paragraphs[0].add_run(name).bold = True
        row.cells[0].paragraphs[0].runs[0].font.size = Pt(9)
        row.cells[1].paragraphs[0].add_run(attrs).font.size = Pt(8.5)
        row.cells[2].paragraphs[0].add_run(meths).font.size = Pt(8.5)
        row.cells[3].paragraphs[0].add_run(role).font.size = Pt(8.5)

    add_h1("2.5 Database Model & Entity-Relationship Diagram")
    add_figure("media_1786788850589.png", "Entity-Relationship Diagram (ERD)")
    add_figure("media_1786788977208.png", "Relational Database Schema Diagram")
    add_h2("2.5.1 Relational Schema Specification")
    add_p("The underlying database schema consists of two highly normalized tables connected via foreign key referential integrity constraints:")

    add_h3("Table 1: User (Authentication & Role Segregation)")
    add_bullet("ID (INTEGER PRIMARY KEY AUTOINCREMENT) - Unique identifier.")
    add_bullet("Username (TEXT UNIQUE NOT NULL) - Student registration or employee username.")
    add_bullet("PasswordHash (TEXT NOT NULL) - Salted PBKDF2:SHA-256 cryptographic hash.")
    add_bullet("Role (TEXT NOT NULL) - Check constraint: Role IN ('Student', 'Admin', 'HOD').")

    add_h3("Table 2: Complaint (Lifecycle, Urgency, Routing & Audit)")
    add_bullet("ID (INTEGER PRIMARY KEY AUTOINCREMENT) - Primary complaint ticket identifier.")
    add_bullet("StudentID (INTEGER NOT NULL) - Foreign Key references User(ID).")
    add_bullet("Category (TEXT NOT NULL) - Academic, Facilities, IT/Network, Harassment, Others, etc.")
    add_bullet("DirectedAgainst (TEXT NOT NULL) - Department or individual target entity.")
    add_bullet("Description (TEXT DEFAULT '') - Detailed complaint narrative.")
    add_bullet("UrgencyLevel (TEXT NOT NULL) - Check constraint: UrgencyLevel IN ('Emergency', 'High', 'Medium', 'Low').")
    add_bullet("Status (TEXT DEFAULT 'Pending') - Check constraint: Status IN ('Pending', 'In-Progress', 'Forwarded', 'Solved').")
    add_bullet("DateSubmitted (DATETIME DEFAULT CURRENT_TIMESTAMP) - Submission audit timestamp.")
    add_bullet("GroupID (INTEGER DEFAULT NULL) - Clustering key for auto-combined duplicate issues.")
    add_bullet("AdminComment (TEXT DEFAULT NULL) - Official response notes recorded upon resolution or progress update.")

    add_h1("2.6 Graphical User Interfaces (GUI) Design Walkthrough")
    add_figure("media_1786788997712.png", "Student Dashboard & Quota Counter Interface", width=Inches(5.0))
    add_figure("media_1786789004364.png", "Complaint Submission Form with Category/Target Taxonomies", width=Inches(5.0))
    add_figure("media_1786789010431.png", "Admin Triage & Complaint Queue Interface", width=Inches(5.0))
    add_figure("media_1786789016566.png", "HOD Executive Oversight & COI Desk Interface", width=Inches(5.0))
    add_p("SUCS features five specialized, reactive web user interfaces built with responsive Tailwind CSS:")
    add_bullet(" Single unified entry-point featuring 1-click credential auto-fill buttons (student1, admin1, hod1) for evaluator testing, server-side flash message alerts, and role-based redirect dispatching.", bold_prefix="1. Unified Login Portal (/login): ")
    add_bullet(" Overview dashboard displaying Open Tickets count, In-Progress count, Solved count, and monthly quota progress meter (5-complaint limit). Houses the intake form with custom 'Others' options and 1-Click Hazard Fill.", bold_prefix="2. Student Overview Dashboard (/student): ")
    add_bullet(" Dedicated full-page view listing all complaints submitted by the student, live status indicators, auto-combined group tags, safe DOM details modals, and official administrative resolution remarks.", bold_prefix="3. Dedicated Student Complaints Page (/student/complaints): ")
    add_bullet(" Triage workspace featuring overview metric cards, urgency-sorted queue (Emergency in bouncing red), 'In-Progress' modal with mandatory comment validation, 1-Click Group Solve button, and a slide-over sidebar for past resolved tickets.", bold_prefix="4. Administrator Workflow Portal (/admin): ")
    add_bullet(" Executive oversight dashboard displaying the live count of tickets currently held by Admins, ranked Category Reason bars, Urgency distribution charts, Conflict-of-Interest queue, Escalated queue, and past resolved sidebar.", bold_prefix="5. HOD Executive Desk (/hod): ")
    add_bullet(" Live in-browser SQLite inspector rendering raw rows of User and Complaint tables with masked hashes, enabling external evaluators to audit database state without third-party software.", bold_prefix="6. Live Database Inspector (/db-view): ")

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 3: DEVELOPMENT & IMPLEMENTATION
    # =========================================================================
    add_title("CHAPTER 3", size=20, space_after=6, color="003366")
    add_title("DEVELOPMENT & IMPLEMENTATION", size=16, space_after=18, color="1E5082")

    add_h1("3.1 Technology Stack Justification")
    add_p("The technological stack for SUCS was selected based on strict software engineering criteria: zero runtime bloat, minimal attack surface, high execution speed, and transparent portability across examiner operating systems:")
    add_bullet(" Python 3.9+ provides robust standard libraries, cryptographic hashing (hashlib/werkzeug), and native SQLite bindings.", bold_prefix="Python 3 Backend: ")
    add_bullet(" Eliminates redundant node_modules, heavyweight Express backends, and decoupled React builds. Single-process Flask runs effortlessly on any computer without port conflicts or build tools.", bold_prefix="Flask Web Framework: ")
    add_bullet(" Zero-configuration, serverless, self-contained relational database stored in a single binary file (sucs_prototype.db). Eliminates external database installation hurdles (MySQL/PostgreSQL) during project evaluation.", bold_prefix="SQLite 3 Database: ")
    add_bullet(" Server-rendered HTML templates eliminate client-side hydration delays and state bugs. Client-side Tailwind CSS provides modern responsive layouts via CDN.", bold_prefix="Jinja2 & Tailwind CSS: ")
    add_bullet(" Industry-standard automated test runner executing unit and integration scenarios with microsecond execution times.", bold_prefix="Pytest Suite: ")

    add_h1("3.2 Core Algorithmic Implementations")
    add_h2("3.2.1 Auto-Combining & Clustering Engine")
    add_p("The automated clustering engine (auto_group_similar_complaints) identifies repetitive grievances and groups them under a shared Master Group ID. The algorithm executes as follows:")
    add_p("1. Scans active complaints: SELECT Category, LOWER(TRIM(DirectedAgainst)) as Target, COUNT(*) FROM Complaint WHERE Status IN ('Pending', 'In-Progress') GROUP BY Category, Target HAVING COUNT(*) > 1.\n"
          "2. For each cluster, fetches all matching tickets ordered by ID ASC.\n"
          "3. Detects if any ticket in the cluster already possesses a GroupID; if found, reuses it; otherwise, sets the earliest ticket's primary key as master GroupID.\n"
          "4. Atomically updates all unlinked cluster tickets to point to the master GroupID.")

    add_h2("3.2.2 Multi-Factor Urgency & Emergency Triage Engine")
    add_p("To prevent critical campus hazards from being overlooked in long queues, SUCS applies a deterministic multi-factor priority algorithm:")
    add_bullet(" Assigned if student checks the explicit Emergency Hazard box on submission.", bold_prefix="Emergency Priority: ")
    add_bullet(" Assigned if Category is 'Infrastructure' / 'Harassment', or description contains hazard keywords ('fire', 'hazard', 'spark', 'electric').", bold_prefix="High Priority: ")
    add_bullet(" Assigned if Category is 'Facilities' or 'IT/Network'.", bold_prefix="Medium Priority: ")
    add_bullet(" Assigned to general routine inquiries and administrative paperwork.", bold_prefix="Low Priority: ")

    add_h2("3.2.3 Monthly Quota Sliding Window Engine")
    add_p("To prevent Denial-of-Service (DoS) abuse and spamming, SUCS implements a strict rate-limiting engine:")
    add_p("Function get_monthly_quota(student_id) evaluates the current calendar month ('YYYY-MM') via SQLite's strftime: SELECT COUNT(*) FROM Complaint WHERE StudentID = ? AND strftime('%Y-%m', DateSubmitted) = ?. Returns used count, limit (5), and remaining quota calculated as max(0, 5 - used). Rejects submission if remaining <= 0.")

    add_h2("3.2.4 1-Click Cascading Atomic Resolution Engine")
    add_p("When an Administrator or HOD clicks '⚡ 1-Click Solve All Related' on a group of complaints, the system executes a single atomic SQL transaction:")
    add_p("UPDATE Complaint SET Status = 'Solved', AdminComment = ? WHERE GroupID = ? OR ID = ?\n"
          "This guarantees transactional atomicity (ACID), updating all matching records simultaneously without partial failure or orphan states.")

    add_h1("3.3 Automated Verification & Test Suite Traceability")
    add_p("The system is verified through an automated Pytest test suite (tests/test_flask_app.py) validating all core business logic and SRS traceability items with a 100% pass rate:")

    test_matrix = [
        ("TC-01", "User Login & Session RBAC Guard", "test_prototype_login", "POST valid credentials -> HTTP 200, role cookie verified", "PASSED"),
        ("TC-02", "Monthly Quota Sliding Window Limit", "test_prototype_quota", "Assert quota limit=5, remaining = max(0, 5 - used)", "PASSED"),
        ("TC-03", "Complaint Submission & Emergency Tagging", "test_prototype_submit_complaint", "POST with is_emergency='on' -> Urgency='Emergency'", "PASSED"),
        ("TC-04", "Automated Conflict-of-Interest Filter", "test_prototype_admin_queue_and_conflict_filter", "Complaint against 'Admin' excluded from /admin queue", "PASSED"),
        ("TC-05", "Auto-Combining & Admin 1-Click Batch Solve", "test_auto_combining_and_admin_one_click_solve", "2 similar tickets auto-clustered -> 1-click solved", "PASSED"),
        ("TC-06", "HOD 1-Click Executive Group Resolution", "test_hod_one_click_group_resolve", "HOD resolves all linked conflict tickets in 1-click", "PASSED"),
        ("TC-07", "In-Progress Status & Mandatory Comment", "test_in_progress_status_and_mandatory_comment", "Rejects empty comment; accepts non-empty comment", "PASSED"),
        ("TC-08", "Custom Options Handling ('Others')", "test_custom_others_options", "Category & Target 'Others' properly saved in DB", "PASSED"),
        ("TC-09", "Dedicated Student Complaints History View", "test_dedicated_student_complaints_page", "GET /student/complaints -> HTTP 200, history rendered", "PASSED")
    ]

    t_tbl = doc.add_table(rows=len(test_matrix)+1, cols=5)
    t_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_tbl.cell(0, 0).width = Inches(0.8)
    t_tbl.cell(0, 1).width = Inches(1.8)
    t_tbl.cell(0, 2).width = Inches(1.8)
    t_tbl.cell(0, 3).width = Inches(1.3)
    t_tbl.cell(0, 4).width = Inches(0.8)

    for c, text in zip(t_tbl.rows[0].cells, ["Test ID", "Requirement Tested", "Pytest Test Function", "Verification Assert", "Status"]):
        set_cell_background(c, "003366")
        r = c.paragraphs[0].add_run(text)
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for i, (tid, req, fn, ass, st) in enumerate(test_matrix):
        row = t_tbl.rows[i+1]
        if i % 2 == 1:
            for cell in row.cells:
                set_cell_background(cell, "F8F9FA")
        row.cells[0].paragraphs[0].add_run(tid).bold = True
        row.cells[0].paragraphs[0].runs[0].font.size = Pt(8.5)
        row.cells[1].paragraphs[0].add_run(req).font.size = Pt(8.5)
        row.cells[2].paragraphs[0].add_run(fn).font.size = Pt(8.5)
        row.cells[3].paragraphs[0].add_run(ass).font.size = Pt(8.5)
        r_st = row.cells[4].paragraphs[0].add_run(st)
        r_st.bold = True
        r_st.font.size = Pt(8.5)
        r_st.font.color.rgb = RGBColor(0, 128, 0) # Green

    add_callout("Automated Test Summary: 9 passed in 0.51s (100% Pass Rate). Zero warnings, zero failures.", title="TEST EXECUTION VERIFICATION")

    doc.add_page_break()

    # =========================================================================
    # REFERENCES & APPENDIX
    # =========================================================================
    add_title("REFERENCES & BIBLIOGRAPHY", size=18, space_after=14)
    refs = [
        "Virtual University of Pakistan (2025). CS619 Final Project Guidelines and Process Model Specification. Department of Computer Sciences, Lahore, Pakistan.",
        "Pressman, R. S., & Maxim, B. R. (2020). Software Engineering: A Practitioner's Approach (9th ed.). McGraw-Hill Education.",
        "Sommerville, I. (2016). Software Engineering (10th ed.). Pearson Education.",
        "Grinberg, M. (2018). Flask Web Development: Developing Web Applications with Python (2nd ed.). O'Reilly Media.",
        "Owens, M., & Allen, G. (2010). The Definitive Guide to SQLite (2nd ed.). Apress.",
        "National Institute of Standards and Technology (NIST) (2023). SP 800-132: Recommendation for Password-Based Key Derivation PBKDF2. U.S. Department of Commerce."
    ]
    for r in refs:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.first_line_indent = Inches(-0.4)
        p.add_run(r).font.size = Pt(10)

    doc.add_paragraph().paragraph_format.space_after = Pt(18)

    add_title("APPENDIX", size=18, space_after=14)
    add_h2("System Deployment & Viva Quick-Start Guide")
    add_bullet("Clone or unzip UniversityComplaintManagement_Prototype.zip onto the host system.", bold_prefix="Step 1: ")
    add_bullet("Open Terminal in directory and run: ./venv/bin/python app.py", bold_prefix="Step 2: ")
    add_bullet("Open web browser at http://127.0.0.1:5001", bold_prefix="Step 3: ")
    add_bullet("Click demo auto-fill buttons (student1, admin1, hod1) to test respective user workflows.", bold_prefix="Step 4: ")
    add_bullet("Navigate to http://127.0.0.1:5001/db-view to inspect live database tables during evaluation.", bold_prefix="Step 5: ")
    add_bullet("Execute test suite anytime using: ./venv/bin/python -m pytest tests/test_flask_app.py", bold_prefix="Step 6: ")

    output_path = "/Users/usman/Desktop/UniversityComplaintManagement/SUCS_Final_Project_Report.docx"
    doc.save(output_path)
    print(f"Report successfully saved to {output_path}")

if __name__ == '__main__':
    create_report()
