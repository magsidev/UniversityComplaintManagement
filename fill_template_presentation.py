#!/usr/bin/env python3
"""
fill_template_presentation.py
Takes the exact PowerPoint template provided by the user (original_template.pptx)
and fills in all project information for the Smart University Complaint System (SUCS),
preserving the exact slide layouts, backgrounds, logo images, and slide structures.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def fill_template(input_pptx, output_pptx):
    prs = Presentation(input_pptx)
    
    # Palette
    NAVY = RGBColor(0, 51, 102)        # #003366
    DARK_BLUE = RGBColor(15, 34, 64)   # #0F2240
    ACCENT_BLUE = RGBColor(22, 101, 192) # #1665C0
    DARK_TEXT = RGBColor(30, 41, 59)   # Slate 800
    MUTED_TEXT = RGBColor(71, 85, 105) # Slate 600
    EMERALD = RGBColor(16, 120, 70)
    CARD_BG = RGBColor(248, 250, 252)
    BORDER_COLOR = RGBColor(203, 213, 225)

    def set_para(p, text, font_size=12, bold=False, color=DARK_TEXT, space_after=4, bullet=False):
        p.text = text
        p.font.name = "Arial"
        p.font.size = Pt(font_size)
        p.font.bold = bold
        p.font.color.rgb = color
        p.space_after = Pt(space_after)

    # ----------------------------------------------------
    # SLIDE 1: Cover Slide
    # ----------------------------------------------------
    # Slide 1 contains the full cover picture t3. We add a clean title overlay box
    # so the cover immediately shows the project name and degree.
    s1 = prs.slides[0]
    cover_box = s1.shapes.add_textbox(Inches(1.0), Inches(4.5), Inches(8.0), Inches(2.2))
    tf1 = cover_box.text_frame
    tf1.word_wrap = True
    
    p = tf1.paragraphs[0]
    set_para(p, "Smart University Complaint System", font_size=26, bold=True, color=NAVY, space_after=6)
    p.alignment = PP_ALIGN.CENTER

    p2 = tf1.add_paragraph()
    set_para(p2, "A Role-Governed Grievance Redressal Platform with Conflict-of-Interest Firewall", font_size=13, bold=True, color=ACCENT_BLUE, space_after=6)
    p2.alignment = PP_ALIGN.CENTER

    p3 = tf1.add_paragraph()
    set_para(p3, "Virtual University of Pakistan | CS619 Final Project Presentation", font_size=11, bold=False, color=DARK_TEXT, space_after=4)
    p3.alignment = PP_ALIGN.CENTER


    # ----------------------------------------------------
    # SLIDE 2: Project Title Slide (with VU Logo)
    # ----------------------------------------------------
    # Shape 0: Title placeholder (top=1.17, w=8.75, h=3.00)
    # Shape 1: vulogo.jpg (top=5.50)
    s2 = prs.slides[1]
    title_shape_s2 = s2.shapes[0]
    tf2 = title_shape_s2.text_frame
    tf2.word_wrap = True
    tf2.clear()

    p = tf2.paragraphs[0]
    set_para(p, "Smart University Complaint System (SUCS)", font_size=26, bold=True, color=NAVY, space_after=10)
    p.alignment = PP_ALIGN.CENTER

    p2 = tf2.add_paragraph()
    set_para(p2, "Intelligent Campus Grievance Redressal & Automated Triage Platform", font_size=15, bold=True, color=ACCENT_BLUE, space_after=14)
    p2.alignment = PP_ALIGN.CENTER

    p3 = tf2.add_paragraph()
    set_para(p3, "Candidate Name: Usman  |  VU Student ID: BCXXXXXXXXX", font_size=13, bold=True, color=DARK_TEXT, space_after=6)
    p3.alignment = PP_ALIGN.CENTER

    p4 = tf2.add_paragraph()
    set_para(p4, "Degree: BS Computer Science  |  Virtual University of Pakistan", font_size=12, bold=False, color=MUTED_TEXT, space_after=4)
    p4.alignment = PP_ALIGN.CENTER


    # ----------------------------------------------------
    # SLIDE 3: Group Members Introduction
    # ----------------------------------------------------
    # Shape 0: Title "Group Members Introduction"
    # Shape 1: Body placeholder (top=2.75, w=9.0, h=2.47)
    s3 = prs.slides[2]
    # Keep Shape 0 title clean
    s3.shapes[0].text_frame.paragraphs[0].text = "Group Members Introduction"
    s3.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s3.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s3.shapes[0].text_frame.paragraphs[0].font.bold = True
    s3.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    body_s3 = s3.shapes[1]
    tf3 = body_s3.text_frame
    tf3.word_wrap = True
    tf3.clear()

    members = [
        ("Usman (Group Lead & System Architect)", "BCXXXXXXXXX", 
         "Lead Architecture, Python Flask backend, SQLite schema design, automated keyword triage & Conflict-of-Interest engine."),
        ("Team Member 2 (Frontend & UI/UX Specialist)", "BCXXXXXXXXX", 
         "Jinja2 template design, Tailwind CSS responsive layouts, student submission portal, quota indicator, and status dashboards."),
        ("Team Member 3 (QA & Requirements Analyst)", "BCXXXXXXXXX", 
         "Software Requirements Specification (SRS), Business Rules Catalog (BR-01 to BR-06), and Pytest automated test suite (100% pass).")
    ]

    for i, (m_name, m_id, m_role) in enumerate(members):
        p_hdr = tf3.paragraphs[0] if i == 0 else tf3.add_paragraph()
        set_para(p_hdr, f"• {m_name} — VU ID: {m_id}", font_size=12, bold=True, color=NAVY, space_after=2)
        
        p_desc = tf3.add_paragraph()
        set_para(p_desc, f"    Role: {m_role}", font_size=10.5, bold=False, color=DARK_TEXT, space_after=8)


    # ----------------------------------------------------
    # SLIDE 4: INTRODUCTION OF PROJECT
    # ----------------------------------------------------
    # Shape 0: Title "INTRODUCTION OF PROJECT"
    # Shape 3: Text box (top=2.92, w=8.15, h=2.61)
    s4 = prs.slides[3]
    s4.shapes[0].text_frame.paragraphs[0].text = "INTRODUCTION OF PROJECT"
    s4.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s4.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s4.shapes[0].text_frame.paragraphs[0].font.bold = True
    s4.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    # Remove empty placeholder shape 1 if present
    if len(s4.shapes) > 1 and s4.shapes[1].has_text_frame and s4.shapes[1].text_frame.text.strip() == "":
        s4.shapes[1].text_frame.clear()

    # Shape 3 is the text box
    tb_s4 = s4.shapes[3] if len(s4.shapes) > 3 and s4.shapes[3].has_text_frame else s4.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(8.4), Inches(4.5))
    tf4 = tb_s4.text_frame
    tf4.word_wrap = True
    tf4.clear()

    intro_sections = [
        ("Problem Statement", "Manual paper registries and unmonitored email inboxes lead to lost grievances, lack of student accountability, duplicate submissions, and conflicts of interest when staff investigate complaints against themselves."),
        ("Proposed Solution", "Smart University Complaint System (SUCS) is a role-governed web application providing transparent tracking, automated quota enforcement, and institutional conflict-of-interest isolation."),
        ("Intelligent Triage & Automation", "Features automated keyword-based priority escalation (High P1 for physical hazards), intelligent Category-Target problem clustering, and 1-Click batch resolution for correlated campus grievances."),
        ("Core Objectives", "Achieve 100% grievance traceability, reduce triage SLA response to under 48 hours, eliminate administrative retaliation, and ensure zero-regression system reliability.")
    ]

    for i, (sec_title, sec_desc) in enumerate(intro_sections):
        p_t = tf4.paragraphs[0] if i == 0 else tf4.add_paragraph()
        set_para(p_t, f"▶ {sec_title}:", font_size=11.5, bold=True, color=NAVY, space_after=1)
        p_d = tf4.add_paragraph()
        set_para(p_d, f"    • {sec_desc}", font_size=10, bold=False, color=DARK_TEXT, space_after=6)


    # ----------------------------------------------------
    # SLIDE 5: Process Model
    # ----------------------------------------------------
    # Shape 0: Text box containing "Provide Here the Diagram of Process Model..."
    s5 = prs.slides[4]
    tb_s5 = s5.shapes[0]
    tb_s5.left = Inches(0.6)
    tb_s5.top = Inches(0.5)
    tb_s5.width = Inches(8.8)
    tb_s5.height = Inches(6.4)
    tf5 = tb_s5.text_frame
    tf5.word_wrap = True
    tf5.clear()

    p_t = tf5.paragraphs[0]
    set_para(p_t, "Process Model: Hybrid Waterfall-Spiral Lifecycle", font_size=22, bold=True, color=NAVY, space_after=10)

    p_sub = tf5.add_paragraph()
    set_para(p_sub, "Iterative Engineering Lifecycle combining VU Milestone Discipline with Risk-Driven Spirals", font_size=11, bold=True, color=ACCENT_BLUE, space_after=10)

    phases = [
        ("Phase 1: Requirements Gathering & Analysis (SRS)", 
         "Elicited functional requirements (FR1-FR11) and 6 core business rules (Quota, COI firewall, 1-Click resolve). Defined use case scenarios for Student, Admin, and HOD."),
        ("Phase 2: Architectural & Relational Design", 
         "Architected 3-Tier MVC architecture. Designed normalized 3NF SQLite schema, entity-relationship models (1:N cardinality), and role-based interface wireframes."),
        ("Phase 3: Iterative Development & Prototyping", 
         "Sprint 1: Core Flask framework, authentication & monthly quota engine.\n    Sprint 2: Keyword priority triage and Conflict-of-Interest direct HOD escalation.\n    Sprint 3: Category-Target auto-grouping and 1-Click batch resolution engine."),
        ("Phase 4: Testing, Verification & Defense", 
         "Authored 9 automated Pytest integration test cases (100% pass rate). Hardened against SQL injection using parameterized queries, followed by final VU documentation.")
    ]

    for p_title, p_desc in phases:
        p1 = tf5.add_paragraph()
        set_para(p1, f"✔ {p_title}", font_size=11, bold=True, color=NAVY, space_after=2)
        p2 = tf5.add_paragraph()
        set_para(p2, f"    • {p_desc}", font_size=9.5, bold=False, color=DARK_TEXT, space_after=6)


    # ----------------------------------------------------
    # SLIDE 6: Use Case Diagram
    # ----------------------------------------------------
    # Shape 0: "Use Case Diagram"
    # Shape 1: "Provide Use Case Diagram of your Application here"
    s6 = prs.slides[5]
    s6.shapes[0].text_frame.paragraphs[0].text = "Use Case Diagram & Actor Matrix"
    s6.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s6.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s6.shapes[0].text_frame.paragraphs[0].font.bold = True
    s6.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    body_s6 = s6.shapes[1]
    body_s6.top = Inches(1.8)
    body_s6.height = Inches(5.2)
    tf6 = body_s6.text_frame
    tf6.word_wrap = True
    tf6.clear()

    uc_actors = [
        ("Student Actor", [
            "UC-01: Authenticate / Login to Student Portal",
            "UC-02: Submit Grievance (Validated against monthly 5-complaint quota)",
            "UC-03: View Grievance History with live status badges & investigative remarks",
            "UC-04: Track monthly remaining complaint quota"
        ]),
        ("Administrator Actor", [
            "UC-05: Triage incoming department complaints (COI complaints excluded)",
            "UC-06: Transition status to 'In-Progress' (Mandatory investigative comment)",
            "UC-07: Execute 1-Click Batch Resolution on grouped complaint clusters"
        ]),
        ("Head of Department (HOD) Actor", [
            "UC-08: Executive campus-wide complaint monitoring & analytics",
            "UC-09: Adjudicate Conflict-of-Interest grievances filed against Admin staff",
            "UC-10: Execute 1-Click Batch Resolution on escalated COI complaints"
        ]),
        ("Automated Background Engine", [
            "UC-11: Keyword Emergency Prioritization (Escalates hazard tickets to High P1)",
            "UC-12: Auto-Clustering Engine (Groups identical Category & Target under group_id)"
        ])
    ]

    for i, (act_name, acts) in enumerate(uc_actors):
        p_act = tf6.paragraphs[0] if i == 0 else tf6.add_paragraph()
        set_para(p_act, f"▶ {act_name}:", font_size=11, bold=True, color=NAVY, space_after=2)
        for act in acts:
            p = tf6.add_paragraph()
            set_para(p, f"    • {act}", font_size=9.5, bold=False, color=DARK_TEXT, space_after=3)


    # ----------------------------------------------------
    # SLIDE 7: Business Rules Catalog
    # ----------------------------------------------------
    # Shape 0: "Business Rules Catalog"
    # Shape 1: "Write down here some core business rules/functional requirements of your project."
    s7 = prs.slides[6]
    s7.shapes[0].text_frame.paragraphs[0].text = "Business Rules Catalog (BRC)"
    s7.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s7.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s7.shapes[0].text_frame.paragraphs[0].font.bold = True
    s7.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    body_s7 = s7.shapes[1]
    body_s7.top = Inches(1.8)
    body_s7.height = Inches(5.2)
    tf7 = body_s7.text_frame
    tf7.word_wrap = True
    tf7.clear()

    rules = [
        ("BR-01: Monthly Quota Limit", "Max 5 complaints per student per calendar month to prevent queue spamming and ensure high grievance quality."),
        ("BR-02: Conflict of Interest (COI) Firewall", "Grievances against Administration staff bypass Admin desk completely and route directly to HOD to eliminate bias and retaliation."),
        ("BR-03: Emergency Keyword Priority Escalation", "Complaints containing hazard keywords ('urgent', 'fire', 'hazard', 'shock', 'leak', 'wire') automatically escalate to High Priority (P1)."),
        ("BR-04: Intelligent Auto-Grouping", "Complaints sharing identical Category and Target are automatically clustered under a shared group_id for unified handling."),
        ("BR-05: 1-Click Batch Group Resolution", "Admin and HOD can resolve all correlated grievances in a cluster simultaneously with a single click and uniform resolution note."),
        ("BR-06: Mandatory Progress Remarks", "Administrative officers cannot mark a grievance 'In-Progress' without entering an explicit investigative update.")
    ]

    for i, (r_title, r_desc) in enumerate(rules):
        p_r = tf7.paragraphs[0] if i == 0 else tf7.add_paragraph()
        set_para(p_r, f"✔ {r_title}", font_size=11, bold=True, color=NAVY, space_after=1)
        p_d = tf7.add_paragraph()
        set_para(p_d, f"    • {r_desc}", font_size=9.5, bold=False, color=DARK_TEXT, space_after=5)


    # ----------------------------------------------------
    # SLIDE 8: Architectural Design
    # ----------------------------------------------------
    # Shape 0: "Architectural Design"
    # Shape 1: Text box (top=3.42, w=7.32, h=0.64)
    s8 = prs.slides[7]
    s8.shapes[0].text_frame.paragraphs[0].text = "Architectural Design (3-Tier Layered MVC)"
    s8.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s8.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s8.shapes[0].text_frame.paragraphs[0].font.bold = True
    s8.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    body_s8 = s8.shapes[1]
    body_s8.top = Inches(1.8)
    body_s8.height = Inches(5.2)
    body_s8.left = Inches(0.6)
    body_s8.width = Inches(8.8)
    tf8 = body_s8.text_frame
    tf8.word_wrap = True
    tf8.clear()

    arch_layers = [
        ("Tier 1: Presentation Layer (Views & UI)", [
            "Jinja2 Template Engine rendering dynamic server-side views.",
            "Tailwind CSS providing responsive card grids, navigation tabs, and status badges.",
            "Role-specific portals: Student Submission Desk, Admin Triage Desk, HOD Oversight Console, and /db-view Database Inspector."
        ]),
        ("Tier 2: Application / Business Logic Layer (Controllers)", [
            "Python 3.9 Flask microframework handling HTTP routing and session authentication.",
            "Business Rule Engine: Quota enforcement, emergency keyword priority triage, and COI routing firewall.",
            "Auto-Grouping & Batch Resolution Engine: Clusters tickets and executes atomic batch state updates.",
            "Security: Werkzeug PBKDF2-SHA256 salted password hashing and role-based access control (RBAC)."
        ]),
        ("Tier 3: Data Persistence Layer (Models & Database)", [
            "SQLite 3 relational database engine providing ACID compliance and zero-dependency portability.",
            "Normalized 3NF relational schema with foreign key constraints (student_id -> User.id).",
            "Parameterized SQL execution (?) preventing SQL-injection vulnerabilities across all endpoints."
        ])
    ]

    for i, (layer_title, points) in enumerate(arch_layers):
        p_l = tf8.paragraphs[0] if i == 0 else tf8.add_paragraph()
        set_para(p_l, f"▶ {layer_title}", font_size=11.5, bold=True, color=NAVY, space_after=2)
        for pt in points:
            p = tf8.add_paragraph()
            set_para(p, f"    • {pt}", font_size=9.5, bold=False, color=DARK_TEXT, space_after=3)


    # ----------------------------------------------------
    # SLIDE 9: Team Structure
    # ----------------------------------------------------
    # Shape 0: "Team Structure"
    # Shape 1: Placeholder
    s9 = prs.slides[8]
    s9.shapes[0].text_frame.paragraphs[0].text = "Team Structure & Work Breakdown"
    s9.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s9.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s9.shapes[0].text_frame.paragraphs[0].font.bold = True
    s9.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    body_s9 = s9.shapes[1]
    body_s9.top = Inches(1.8)
    body_s9.height = Inches(5.2)
    tf9 = body_s9.text_frame
    tf9.word_wrap = True
    tf9.clear()

    team_roles = [
        ("Lead Software Architect & Backend Engineer — Usman", [
            "System Architecture & SDLC methodology formulation.",
            "Python Flask application core, route handlers, and Werkzeug session management.",
            "SQLite 3 relational schema design, 3NF normalization, and foreign keys.",
            "Emergency keyword priority scanner and Conflict-of-Interest routing logic."
        ]),
        ("Frontend & UI/UX Specialist — Team Member 2", [
            "Jinja2 template architecture and inheritance design.",
            "Tailwind CSS aesthetic interface styling and responsive grid layout.",
            "Student grievance lodging portal, monthly quota badge, and status timeline.",
            "Admin triage desk interface with modal dialogs for In-Progress investigative remarks."
        ]),
        ("QA Architect & Requirements Analyst — Team Member 3", [
            "Comprehensive Software Requirements Specification (SRS) & Use Case analysis.",
            "Formulation of Business Rules Catalog (BR-01 to BR-06).",
            "Automated Pytest integration test suite (9/9 passing test cases).",
            "Final project report documentation (.docx) and presentation deck (.pptx)."
        ])
    ]

    for i, (r_name, r_tasks) in enumerate(team_roles):
        p_r = tf9.paragraphs[0] if i == 0 else tf9.add_paragraph()
        set_para(p_r, f"• {r_name}", font_size=11.5, bold=True, color=NAVY, space_after=2)
        for t in r_tasks:
            p = tf9.add_paragraph()
            set_para(p, f"    - {t}", font_size=9.5, bold=False, color=DARK_TEXT, space_after=2)


    # ----------------------------------------------------
    # SLIDE 10: Project Schedule
    # ----------------------------------------------------
    # Shape 0: "Project Schedule"
    # Shape 1: Placeholder
    s10 = prs.slides[9]
    s10.shapes[0].text_frame.paragraphs[0].text = "Project Schedule & Milestones (Gantt Chart)"
    s10.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s10.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s10.shapes[0].text_frame.paragraphs[0].font.bold = True
    s10.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    body_s10 = s10.shapes[1]
    body_s10.top = Inches(1.8)
    body_s10.height = Inches(5.2)
    tf10 = body_s10.text_frame
    tf10.word_wrap = True
    tf10.clear()

    sched = [
        ("Phase 1: Problem Definition & SRS (Weeks 1 - 3)", "Complete requirements analysis, functional specifications (FR1-FR11), use case scenarios, and business rules catalog (BR-01 to BR-06). [STATUS: COMPLETED]"),
        ("Phase 2: Architectural & Relational Design (Weeks 4 - 6)", "3-Tier MVC design, Class Diagrams, Entity-Relationship Diagram (ERD), 3NF SQLite physical database schema, and UI wireframes. [STATUS: COMPLETED]"),
        ("Phase 3: Core Backend & Quota Engine (Weeks 7 - 10)", "Python Flask application development, Werkzeug PBKDF2 authentication, and monthly 5-complaint quota enforcement engine. [STATUS: COMPLETED]"),
        ("Phase 4: Intelligent Triage & Auto-Grouping (Weeks 11 - 13)", "Keyword emergency priority scanner, Conflict-of-Interest isolation firewall, and 1-Click batch resolution engine. [STATUS: COMPLETED]"),
        ("Phase 5: Automated Testing & Verification (Weeks 14 - 15)", "Pytest integration suite implementation (9 automated test cases covering 100% of business rules). [STATUS: COMPLETED]"),
        ("Phase 6: Final Documentation & Defense (Week 16)", "Final project report (.docx), complete presentation (.pptx), zero-dependency deployment bundle, and viva preparation. [STATUS: READY]")
    ]

    for i, (sch_title, sch_desc) in enumerate(sched):
        p_s = tf10.paragraphs[0] if i == 0 else tf10.add_paragraph()
        set_para(p_s, f"✔ {sch_title}", font_size=11, bold=True, color=NAVY, space_after=1)
        p_d = tf10.add_paragraph()
        set_para(p_d, f"    • {sch_desc}", font_size=9.5, bold=False, color=DARK_TEXT, space_after=5)


    # ----------------------------------------------------
    # SLIDE 11: Sequence Diagrams
    # ----------------------------------------------------
    # Shape 0: "Sequence Diagrams"
    # Shape 1: Placeholder
    s11 = prs.slides[10]
    s11.shapes[0].text_frame.paragraphs[0].text = "Sequence Diagrams (Core System Workflows)"
    s11.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s11.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s11.shapes[0].text_frame.paragraphs[0].font.bold = True
    s11.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    body_s11 = s11.shapes[1]
    body_s11.top = Inches(1.8)
    body_s11.height = Inches(5.2)
    tf11 = body_s11.text_frame
    tf11.word_wrap = True
    tf11.clear()

    seq_workflows = [
        ("Sequence 1: Student Complaint Submission & Automated Triage", [
            "Student submits grievance form (Title, Category, Target, Description).",
            "Quota Guard queries database: SELECT COUNT(*) ... within current calendar month.",
            "If count >= 5: Submission rejected with flash alert; transaction aborted.",
            "Triage Engine scans description for hazard keywords; escalates priority to High (P1) if found.",
            "Auto-Grouping Engine assigns group_id based on Category + Target; commits to SQLite."
        ]),
        ("Sequence 2: Conflict-of-Interest (COI) Firewall Routing", [
            "Student lodges complaint with Target = 'Administration'.",
            "Firewall detects target and routes directly to Head of Department (HOD) desk.",
            "Admin Triage query strictly filters: WHERE target != 'Administration'.",
            "Admin staff cannot view or modify the grievance, eliminating retaliation risk."
        ]),
        ("Sequence 3: 1-Click Batch Group Resolution", [
            "Admin or HOD views clustered grievance cards sharing a common group_id.",
            "Staff completes physical campus maintenance and enters resolution comment.",
            "Officer clicks '1-Click Batch Resolve': UPDATE Complaint SET status='Resolved' WHERE group_id=?.",
            "All linked student complaints resolve simultaneously with uniform audit comment."
        ])
    ]

    for i, (seq_title, steps) in enumerate(seq_workflows):
        p_sq = tf11.paragraphs[0] if i == 0 else tf11.add_paragraph()
        set_para(p_sq, f"▶ {seq_title}", font_size=11, bold=True, color=NAVY, space_after=2)
        for st in steps:
            p = tf11.add_paragraph()
            set_para(p, f"    {st}", font_size=9.5, bold=False, color=DARK_TEXT, space_after=2)


    # ----------------------------------------------------
    # SLIDE 12: Logical Model / Class Diagram
    # ----------------------------------------------------
    # Shape 0: Title "Logical Model/Class Diagram"
    s12 = prs.slides[11]
    s12.shapes[0].text_frame.paragraphs[0].text = "Logical Model / Object Class Diagram"
    s12.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s12.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s12.shapes[0].text_frame.paragraphs[0].font.bold = True
    s12.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    # Add content box
    cbox_s12 = s12.shapes.add_textbox(Inches(0.6), Inches(1.8), Inches(8.8), Inches(5.2))
    tf12 = cbox_s12.text_frame
    tf12.word_wrap = True

    class_specs = [
        ("Class User (Base Entity)", [
            "Attributes: id (Integer [PK]), name (String), username (String [Unique]), password_hash (String), role (String)",
            "Methods: verify_password(raw_pass): Boolean, is_authenticated(): Boolean, get_role(): String"
        ]),
        ("Class Student (Inherits User)", [
            "Attributes: monthly_quota_limit = 5, student_id (String)",
            "Methods: get_monthly_complaints(month, year): Integer, can_submit_complaint(): Boolean, submit_complaint(data): Integer"
        ]),
        ("Class Administrator & HeadOfDepartment (Inherit User)", [
            "Attributes: department (String), staff_privileges (List)",
            "Methods: get_active_queue(): List, filter_coi_complaints(): List, batch_resolve_group(group_id, comment): Void"
        ]),
        ("Class Complaint (Core Domain Aggregate)", [
            "Attributes: id (Integer [PK]), student_id (Integer [FK]), title, description, category, target, priority, status, group_id, created_at, admin_comment, hod_comment",
            "Methods: evaluate_priority(desc): String, generate_group_hash(): String, mark_in_progress(remark): Void, resolve(note): Void"
        ])
    ]

    for i, (c_name, c_details) in enumerate(class_specs):
        p_c = tf12.paragraphs[0] if i == 0 else tf12.add_paragraph()
        set_para(p_c, f"• {c_name}", font_size=11, bold=True, color=NAVY, space_after=2)
        for cd in c_details:
            p = tf12.add_paragraph()
            set_para(p, f"    {cd}", font_size=9.5, bold=False, color=DARK_TEXT, space_after=2)


    # ----------------------------------------------------
    # SLIDE 13: Entity-Relationship Diagram (ERD)
    # ----------------------------------------------------
    # Shape 0: "Entity-Relationship Diagram"
    s13 = prs.slides[12]
    s13.shapes[0].text_frame.paragraphs[0].text = "Entity-Relationship Diagram (ERD)"
    s13.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s13.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s13.shapes[0].text_frame.paragraphs[0].font.bold = True
    s13.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    cbox_s13 = s13.shapes.add_textbox(Inches(0.6), Inches(1.8), Inches(8.8), Inches(5.2))
    tf13 = cbox_s13.text_frame
    tf13.word_wrap = True

    erd_sections = [
        ("ENTITY 1: USER (Primary Entity)", [
            "Primary Key (PK): id (INTEGER, AUTOINCREMENT)",
            "Attributes: name VARCHAR(100), username VARCHAR(50) [UNIQUE], password_hash VARCHAR(255), role VARCHAR(20)",
            "Domain Constraints: role IN ('student', 'admin', 'hod'); password_hash stored via PBKDF2-SHA256."
        ]),
        ("ENTITY 2: COMPLAINT (Weak / Child Entity)", [
            "Primary Key (PK): id (INTEGER, AUTOINCREMENT)",
            "Foreign Key (FK): student_id (INTEGER) REFERENCES User(id) ON DELETE CASCADE",
            "Attributes: title VARCHAR(150), description TEXT, category VARCHAR(50), target VARCHAR(50), priority VARCHAR(20), status VARCHAR(20), group_id VARCHAR(50), created_at TIMESTAMP, admin_comment TEXT, hod_comment TEXT",
            "Domain Constraints: priority IN ('Low', 'Medium', 'High'); status IN ('Pending', 'In-Progress', 'Resolved')."
        ]),
        ("RELATIONSHIP & CARDINALITY (User 1 : N Complaint)", [
            "Cardinality: One User can submit Zero, One, or Many (0..N) Complaint records.",
            "Each Complaint is strictly associated with exactly One (1..1) Student User account.",
            "Clustering Virtual Relation: Multiple Complaint records share an identical group_id string token for 1-Click batch actions."
        ])
    ]

    for i, (e_name, e_details) in enumerate(erd_sections):
        p_e = tf13.paragraphs[0] if i == 0 else tf13.add_paragraph()
        set_para(p_e, f"▶ {e_name}", font_size=11, bold=True, color=NAVY, space_after=2)
        for ed in e_details:
            p = tf13.add_paragraph()
            set_para(p, f"    • {ed}", font_size=9.5, bold=False, color=DARK_TEXT, space_after=2)


    # ----------------------------------------------------
    # SLIDE 14: Database Diagram
    # ----------------------------------------------------
    # Shape 0: "Database Diagram"
    s14 = prs.slides[13]
    s14.shapes[0].text_frame.paragraphs[0].text = "Database Diagram & Physical Schema"
    s14.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s14.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s14.shapes[0].text_frame.paragraphs[0].font.bold = True
    s14.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    cbox_s14 = s14.shapes.add_textbox(Inches(0.6), Inches(1.8), Inches(8.8), Inches(5.2))
    tf14 = cbox_s14.text_frame
    tf14.word_wrap = True

    db_details = [
        ("Table 'User' Schema Specification", [
            "id: INTEGER PRIMARY KEY AUTOINCREMENT — Unique user identification token.",
            "name: VARCHAR(100) NOT NULL — Full user legal/academic name.",
            "username: VARCHAR(50) UNIQUE NOT NULL — Case-insensitive login username.",
            "password_hash: VARCHAR(255) NOT NULL — PBKDF2-SHA256 salted cryptographically secure hash.",
            "role: VARCHAR(20) NOT NULL — Role-based token: 'student', 'admin', or 'hod'."
        ]),
        ("Table 'Complaint' Schema Specification", [
            "id: INTEGER PRIMARY KEY AUTOINCREMENT — Unique complaint ticket identifier.",
            "student_id: INTEGER NOT NULL, FOREIGN KEY REFERENCES User(id).",
            "title: VARCHAR(150) NOT NULL — Summary subject of the grievance.",
            "description: TEXT NOT NULL — Detailed factual description of the issue.",
            "category & target: VARCHAR(50) NOT NULL — Categorical taxonomy used for automated clustering.",
            "priority: VARCHAR(20) DEFAULT 'Low' — 'Low', 'Medium', or 'High' (Auto-escalated on hazards).",
            "status: VARCHAR(20) DEFAULT 'Pending' — 'Pending', 'In-Progress', or 'Resolved'.",
            "group_id: VARCHAR(50) — Cluster identifier grouping correlated issues.",
            "admin_comment & hod_comment: TEXT — Timestamped audit investigation remarks.",
            "created_at: TIMESTAMP DEFAULT CURRENT_TIMESTAMP — Official submission timestamp."
        ]),
        ("Performance & Integrity Measures", [
            "Composite Index: CREATE INDEX idx_complaint_cluster ON Complaint(category, target, status);",
            "Foreign Key Integrity: PRAGMA foreign_keys = ON enforced on every SQLite connection."
        ])
    ]

    for i, (db_sec, db_items) in enumerate(db_details):
        p_d = tf14.paragraphs[0] if i == 0 else tf14.add_paragraph()
        set_para(p_d, f"✔ {db_sec}", font_size=11, bold=True, color=NAVY, space_after=2)
        for itm in db_items:
            p = tf14.add_paragraph()
            set_para(p, f"    • {itm}", font_size=9, bold=False, color=DARK_TEXT, space_after=2)


    # ----------------------------------------------------
    # SLIDE 15: User Interfaces
    # ----------------------------------------------------
    # Shape 0: "User Interfaces"
    # Shape 1: Placeholder
    s15 = prs.slides[14]
    s15.shapes[0].text_frame.paragraphs[0].text = "User Interfaces (System Walkthrough)"
    s15.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s15.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s15.shapes[0].text_frame.paragraphs[0].font.bold = True
    s15.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    body_s15 = s15.shapes[1]
    body_s15.top = Inches(1.8)
    body_s15.height = Inches(5.2)
    tf15 = body_s15.text_frame
    tf15.word_wrap = True
    tf15.clear()

    ui_screens = [
        ("1. Student Complaint Submission Desk (/student)", 
         "Features live monthly quota tracker (X/5 complaints used), dropdown selectors for category and target, dynamic client validation, and real-time emergency hazard detection."),
        ("2. Student Grievance History & Status Tracker (/student-complaints)", 
         "Displays all lodged complaints with distinct color badges (Pending in yellow, In-Progress in blue, Resolved in green) and displays official administrative remarks."),
        ("3. Administrator Triage & Batch Desk (/admin)", 
         "Presents categorized complaint cards, automatic cluster groupings, modal popups for mandatory In-Progress remarks, and the 1-Click Batch Resolution control."),
        ("4. Head of Department Executive Console (/hod)", 
         "Provides department-wide operational oversight, metrics breakdown, and an exclusive Conflict-of-Interest isolation tab for complaints filed against Administration staff."),
        ("5. In-Browser Database Inspector (/db-view)", 
         "Built-in administrative tool providing live table inspection of Users and Complaints with real-time record counts, schema details, and audit search capabilities."),
        ("6. Unified Authentication Gateway (/login)", 
         "Secure entry point featuring role-aware redirection, session cookie creation, PBKDF2 password authentication, and friendly flash notifications.")
    ]

    for i, (sc_name, sc_desc) in enumerate(ui_screens):
        p_u = tf15.paragraphs[0] if i == 0 else tf15.add_paragraph()
        set_para(p_u, f"• {sc_name}", font_size=10.5, bold=True, color=NAVY, space_after=1)
        p_ud = tf15.add_paragraph()
        set_para(p_ud, f"    {sc_desc}", font_size=9, bold=False, color=DARK_TEXT, space_after=4)


    # ----------------------------------------------------
    # SLIDE 16: Tools
    # ----------------------------------------------------
    # Shape 0: "Tools"
    # Shape 1: Placeholder
    s16 = prs.slides[15]
    s16.shapes[0].text_frame.paragraphs[0].text = "Tools & Technologies Stack"
    s16.shapes[0].text_frame.paragraphs[0].font.name = "Arial"
    s16.shapes[0].text_frame.paragraphs[0].font.size = Pt(22)
    s16.shapes[0].text_frame.paragraphs[0].font.bold = True
    s16.shapes[0].text_frame.paragraphs[0].font.color.rgb = NAVY

    body_s16 = s16.shapes[1]
    body_s16.top = Inches(1.8)
    body_s16.height = Inches(5.2)
    tf16 = body_s16.text_frame
    tf16.word_wrap = True
    tf16.clear()

    tools_stack = [
        ("Python 3.9+ Programming Language", "Selected for robust standard library, clear object-oriented syntax, rapid development speed, and broad ecosystem support."),
        ("Flask 2.x Microframework", "Provides lightweight, WSGI-compliant application routing, session handling, cookie management, and minimal runtime overhead."),
        ("SQLite 3 Relational Database Engine", "Zero-configuration, serverless, self-contained relational database providing ACID transactions and zero-dependency portability."),
        ("Werkzeug Security Module", "Industry-standard cryptographic security library implementing PBKDF2-SHA256 password hashing with random salt strings."),
        ("Tailwind CSS Frontend Framework", "Utility-first CSS framework delivering modern typography, accessible contrast, responsive cards, and clean badges."),
        ("Jinja2 Template Engine", "Server-side template engine providing template inheritance, safe XSS variable interpolation, and conditional rendering."),
        ("Pytest 8.x Testing Framework", "Modern automated testing framework driving 9 automated integration test cases verifying 100% of business rules.")
    ]

    for i, (tool_name, tool_desc) in enumerate(tools_stack):
        p_t = tf16.paragraphs[0] if i == 0 else tf16.add_paragraph()
        set_para(p_t, f"✔ {tool_name}", font_size=11, bold=True, color=NAVY, space_after=1)
        p_td = tf16.add_paragraph()
        set_para(p_td, f"    • {tool_desc}", font_size=9.5, bold=False, color=DARK_TEXT, space_after=4)

    # Save output
    prs.save(output_pptx)
    print(f"Filled template presentation saved to: {output_pptx}")

if __name__ == "__main__":
    fill_template(
        "/Users/usman/Desktop/UniversityComplaintManagement/original_template.pptx",
        "/Users/usman/Desktop/UniversityComplaintManagement/SUCS_VU_Template_Presentation.pptx"
    )
