#!/usr/bin/env python3
"""
generate_presentation.py
Generates a comprehensive, professional, 16:9 PowerPoint presentation (.pptx)
for the Smart University Complaint System (SUCS) following the Virtual University (VU)
Final Project Presentation template structure and guidelines.
"""

import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_presentation(output_path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]  # Blank slide

    # Color Palette
    NAVY = RGBColor(0, 51, 102)        # Primary VU Navy (#003366)
    DARK_BLUE = RGBColor(10, 37, 64)   # Deep Blue (#0A2540)
    ACCENT_BLUE = RGBColor(22, 101, 192) # Secondary Blue (#1665C0)
    GOLD = RGBColor(218, 165, 32)      # Accent Gold (#DAA520)
    WHITE = RGBColor(255, 255, 255)
    TEXT_DARK = RGBColor(30, 41, 59)   # Slate 800 (#1E293B)
    TEXT_MUTED = RGBColor(100, 116, 139) # Slate 500 (#64748B)
    CARD_BG = RGBColor(248, 250, 252)  # Slate 50 (#F8FAFC)
    CARD_BORDER = RGBColor(226, 232, 240) # Slate 200 (#E2E8F0)
    LIGHT_BLUE = RGBColor(238, 242, 255) # Indigo 50
    EMERALD = RGBColor(16, 185, 129)    # Green (#10B981)
    AMBER = RGBColor(245, 158, 11)      # Amber (#F59E0B)
    RED = RGBColor(239, 68, 68)         # Red (#EF4444)

    def add_header(slide, title_text, category_text="VIRTUAL UNIVERSITY OF PAKISTAN | FINAL PROJECT PRESENTATION (CS619)"):
        # Top banner shape
        banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        banner.fill.solid()
        banner.fill.fore_color.rgb = NAVY
        banner.line.color.rgb = NAVY

        # Accent line under banner
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.1), Inches(13.333), Inches(0.06))
        line.fill.solid()
        line.fill.fore_color.rgb = GOLD
        line.line.color.rgb = GOLD

        # Category / breadcrumb text
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.5), Inches(0.3))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        tf_cat.margin_left = tf_cat.margin_right = tf_cat.margin_top = tf_cat.margin_bottom = 0
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.name = "Calibri"
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = GOLD

        # Title text
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.38), Inches(11.5), Inches(0.6))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        tf_title.margin_left = tf_title.margin_right = tf_title.margin_top = tf_title.margin_bottom = 0
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.name = "Arial"
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = WHITE

        # Footer
        add_footer(slide)

    def add_footer(slide):
        # Footer line
        fline = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.02))
        fline.fill.solid()
        fline.fill.fore_color.rgb = CARD_BORDER
        fline.line.color.rgb = CARD_BORDER

        # Footer text
        fbox = slide.shapes.add_textbox(Inches(0.8), Inches(7.1), Inches(11.733), Inches(0.3))
        tff = fbox.text_frame
        tff.margin_left = tff.margin_right = tff.margin_top = tff.margin_bottom = 0
        pf = tff.paragraphs[0]
        pf.text = "Smart University Complaint System (SUCS)  |  Virtual University Final Project Presentation"
        pf.font.name = "Calibri"
        pf.font.size = Pt(9)
        pf.font.color.rgb = TEXT_MUTED

    def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
        return card

    # ==========================================
    # SLIDE 1: TITLE SLIDE
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    # Background full navy
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = DARK_BLUE
    bg1.line.color.rgb = DARK_BLUE

    # Geometric accent shapes
    accent_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.3), Inches(7.5))
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = GOLD
    accent_bar.line.color.rgb = GOLD

    # Institutional banner text
    inst_box = s1.shapes.add_textbox(Inches(1.5), Inches(1.0), Inches(10.5), Inches(0.5))
    tf_inst = inst_box.text_frame
    p_inst = tf_inst.paragraphs[0]
    p_inst.text = "VIRTUAL UNIVERSITY OF PAKISTAN | DEPARTMENT OF COMPUTER SCIENCE"
    p_inst.font.name = "Calibri"
    p_inst.font.size = Pt(13)
    p_inst.font.bold = True
    p_inst.font.color.rgb = GOLD

    # Project Title Box (Template: "Write Your Project Title Here")
    title_box1 = s1.shapes.add_textbox(Inches(1.5), Inches(1.6), Inches(10.5), Inches(1.8))
    tf1 = title_box1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "Smart University Complaint System"
    p1.font.name = "Arial"
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = WHITE

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "An Intelligent, Role-Governed Grievance Redressal Platform with Conflict-of-Interest Firewall & Automated Triage Clustering"
    p1_sub.font.name = "Calibri"
    p1_sub.font.size = Pt(16)
    p1_sub.font.color.rgb = RGBColor(190, 215, 245)
    p1_sub.space_before = Pt(8)

    # Divider
    div1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(3.7), Inches(10.3), Inches(0.04))
    div1.fill.solid()
    div1.fill.fore_color.rgb = GOLD
    div1.line.color.rgb = GOLD

    # Card 1: Project Details
    c1 = add_card(s1, Inches(1.5), Inches(4.1), Inches(5.0), Inches(2.3), bg_color=NAVY, border_color=ACCENT_BLUE)
    c1_box = s1.shapes.add_textbox(Inches(1.7), Inches(4.25), Inches(4.6), Inches(2.0))
    tf_c1 = c1_box.text_frame
    tf_c1.word_wrap = True
    
    p = tf_c1.paragraphs[0]
    p.text = "PROJECT METADATA"
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = GOLD

    items = [
        ("Course Code:", "CS619 - Final Project"),
        ("Technology Stack:", "Python 3.9, Flask, SQLite 3, Tailwind CSS"),
        ("Architecture:", "3-Tier MVC with Parameterized Relational Engine"),
        ("Testing Framework:", "Pytest (100% Automated Coverage)")
    ]
    for lbl, val in items:
        p = tf_c1.add_paragraph()
        p.text = f"{lbl} "
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE
        run = p.add_run()
        run.text = val
        run.font.bold = False
        run.font.color.rgb = RGBColor(220, 230, 245)
        p.space_before = Pt(4)

    # Card 2: Student Details
    c2 = add_card(s1, Inches(6.8), Inches(4.1), Inches(5.0), Inches(2.3), bg_color=NAVY, border_color=ACCENT_BLUE)
    c2_box = s1.shapes.add_textbox(Inches(7.0), Inches(4.25), Inches(4.6), Inches(2.0))
    tf_c2 = c2_box.text_frame
    tf_c2.word_wrap = True
    
    p = tf_c2.paragraphs[0]
    p.text = "CANDIDATE & SUPERVISION"
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = GOLD

    items2 = [
        ("Student Name:", "Usman"),
        ("VU Student ID:", "BCXXXXXXXXX (BS Computer Science)"),
        ("Project Supervisor:", "Virtual University Project Supervisor"),
        ("Academic Session:", "Final Project Defense / 2026")
    ]
    for lbl, val in items2:
        p = tf_c2.add_paragraph()
        p.text = f"{lbl} "
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE
        run = p.add_run()
        run.text = val
        run.font.bold = False
        run.font.color.rgb = RGBColor(220, 230, 245)
        p.space_before = Pt(4)


    # ==========================================
    # SLIDE 2: GROUP MEMBERS INTRODUCTION
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Group Members Introduction", "VIRTUAL UNIVERSITY OF PAKISTAN | TEAM ROLES & ASSIGNMENTS")

    members = [
        {
            "name": "Usman (Lead)",
            "vuid": "BCXXXXXXXXX",
            "role": "Lead Architect & Backend Engineer",
            "bullets": [
                "Overall application architecture & Python Flask server",
                "SQLite schema design, relational normalization & indices",
                "Automated keyword prioritization & triage engine",
                "Conflict-of-Interest (COI) routing engine implementation"
            ]
        },
        {
            "name": "Team Member 2",
            "vuid": "BCXXXXXXXXX",
            "role": "UI/UX & Frontend Specialist",
            "bullets": [
                "Jinja2 template design & responsive layouts",
                "Tailwind CSS aesthetic components & color palettes",
                "Student dashboard, quota counter & complaint history views",
                "Admin triage console & In-Progress status modals"
            ]
        },
        {
            "name": "Team Member 3",
            "vuid": "BCXXXXXXXXX",
            "role": "QA Architect & Systems Analyst",
            "bullets": [
                "Comprehensive SRS business rules catalog (BR-01 to BR-06)",
                "Automated Pytest integration suite (100% test pass)",
                "Conflict-of-Interest boundary condition testing",
                "Final technical report, UML modeling & viva preparation"
            ]
        }
    ]

    card_w = Inches(3.7)
    card_h = Inches(5.3)
    card_top = Inches(1.4)

    for i, m in enumerate(members):
        left = Inches(0.8 + i * 4.0)
        add_card(s2, left, card_top, card_w, card_h)

        # Header pill inside card
        pill = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, card_top, card_w, Inches(1.1))
        pill.fill.solid()
        pill.fill.fore_color.rgb = NAVY
        pill.line.color.rgb = NAVY

        pbox = s2.shapes.add_textbox(left + Inches(0.2), card_top + Inches(0.12), card_w - Inches(0.4), Inches(0.9))
        ptf = pbox.text_frame
        ptf.word_wrap = True
        ptf.margin_left = ptf.margin_right = ptf.margin_top = ptf.margin_bottom = 0
        p = ptf.paragraphs[0]
        p.text = m["name"]
        p.font.name = "Arial"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = WHITE

        p_id = ptf.add_paragraph()
        p_id.text = f"VU ID: {m['vuid']}"
        p_id.font.name = "Calibri"
        p_id.font.size = Pt(11)
        p_id.font.color.rgb = GOLD

        p_r = ptf.add_paragraph()
        p_r.text = m["role"]
        p_r.font.name = "Calibri"
        p_r.font.size = Pt(11)
        p_r.font.bold = True
        p_r.font.color.rgb = RGBColor(200, 225, 255)

        # Body bullets inside card
        bbox = s2.shapes.add_textbox(left + Inches(0.2), card_top + Inches(1.3), card_w - Inches(0.4), card_h - Inches(1.5))
        btf = bbox.text_frame
        btf.word_wrap = True
        btf.margin_left = btf.margin_right = btf.margin_top = btf.margin_bottom = 0

        p_hdr = btf.paragraphs[0]
        p_hdr.text = "Key Responsibilities & Deliverables:"
        p_hdr.font.name = "Calibri"
        p_hdr.font.size = Pt(11)
        p_hdr.font.bold = True
        p_hdr.font.color.rgb = TEXT_DARK
        p_hdr.space_after = Pt(8)

        for b in m["bullets"]:
            p = btf.add_paragraph()
            p.text = f"• {b}"
            p.font.name = "Calibri"
            p.font.size = Pt(10.5)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(6)


    # ==========================================
    # SLIDE 3: INTRODUCTION OF PROJECT
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "INTRODUCTION OF PROJECT", "PROJECT OVERVIEW, MOTIVATION & OBJECTIVES")

    # 4 Quadrants / Cards
    cards_s3 = [
        {
            "title": "Problem Statement",
            "badge": "THE CHALLENGE",
            "color": RED,
            "bullets": [
                "Traditional paper registers or unstructured email inboxes lead to lost, untracked grievances.",
                "Lack of quota limits causes individual users to flood the queue with duplicate submissions.",
                "Administrative bias occurs when complaints filed against Admin staff are reviewed by the same staff.",
                "Zero automated correlation leaves administrators manually solving dozens of identical campus complaints."
            ]
        },
        {
            "title": "Proposed Solution",
            "badge": "OUR INNOVATION",
            "color": NAVY,
            "bullets": [
                "Smart University Complaint System (SUCS): A unified, role-governed web application.",
                "Monthly Quota Engine (5 complaints/month) enforces disciplined, high-quality submissions.",
                "Conflict-of-Interest (COI) Firewall isolates complaints against Admin and routes them exclusively to HOD.",
                "Real-time transparent grievance lifecycle tracking with mandatory investigative remarks."
            ]
        },
        {
            "title": "Intelligent Automation",
            "badge": "KEY FEATURES",
            "color": ACCENT_BLUE,
            "bullets": [
                "Keyword Priority Escalation: Automatically escalates tickets to High (P1) if hazards/emergencies detected.",
                "Intelligent Auto-Grouping Engine: Clustered tickets sharing identical Category and Target into a single group.",
                "1-Click Batch Resolution: Admin or HOD resolves dozens of correlated complaints in one atomic action.",
                "In-Browser Database Inspector (/db-view) for transparent, verified administrative auditing."
            ]
        },
        {
            "title": "Core Project Objectives",
            "badge": "PROJECT GOALS",
            "color": EMERALD,
            "bullets": [
                "100% Grievance Traceability: Every complaint is timestamped, cataloged, and tracked through resolution.",
                "Under 48-Hour Triage SLA: Automated clustering and priority detection drastically reduce response latency.",
                "Institutional Integrity: Eliminates administrative retaliation and bias through strict COI enforcement.",
                "High Performance & Security: Lightweight Python/SQLite stack with PBKDF2-SHA256 password hashing."
            ]
        }
    ]

    for i, c in enumerate(cards_s3):
        col = i % 2
        row = i // 2
        left = Inches(0.8 + col * 5.95)
        top = Inches(1.4 + row * 2.7)
        w = Inches(5.75)
        h = Inches(2.5)

        add_card(s3, left, top, w, h)

        # Header bar inside card
        cbar = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, Inches(0.45))
        cbar.fill.solid()
        cbar.fill.fore_color.rgb = c["color"]
        cbar.line.color.rgb = c["color"]

        tb = s3.shapes.add_textbox(left + Inches(0.2), top + Inches(0.08), w - Inches(0.4), Inches(0.35))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = f"{c['badge']}: {c['title']}"
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE

        # Content bullets
        cb = s3.shapes.add_textbox(left + Inches(0.2), top + Inches(0.55), w - Inches(0.4), h - Inches(0.65))
        ctf = cb.text_frame
        ctf.word_wrap = True
        ctf.margin_left = ctf.margin_right = ctf.margin_top = ctf.margin_bottom = 0

        for b_idx, b in enumerate(c["bullets"]):
            p = ctf.paragraphs[0] if b_idx == 0 else ctf.add_paragraph()
            p.text = f"• {b}"
            p.font.name = "Calibri"
            p.font.size = Pt(10)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(3)


    # ==========================================
    # SLIDE 4: PROCESS MODEL
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Process Model: Hybrid Waterfall-Spiral Lifecycle", "SOFTWARE DEVELOPMENT LIFECYCLE (SDLC) & JUSTIFICATION")

    # Left card: Model Architecture & Flow
    add_card(s4, Inches(0.8), Inches(1.4), Inches(7.5), Inches(5.3))
    b_l = s4.shapes.add_textbox(Inches(1.0), Inches(1.55), Inches(7.1), Inches(5.0))
    tf_l = b_l.text_frame
    tf_l.word_wrap = True
    tf_l.margin_left = tf_l.margin_right = tf_l.margin_top = tf_l.margin_bottom = 0

    p = tf_l.paragraphs[0]
    p.text = "VU Process Model Implementation: 4 Iterative Stages"
    p.font.name = "Arial"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(10)

    stages = [
        ("Phase 1: Requirements Gathering & Analysis", [
            "Formalized Virtual University CS619 requirements and student grievance scenarios.",
            "Formulated comprehensive SRS, Use Case specifications, and Business Rules Catalog (BR-01 to BR-06)."
        ]),
        ("Phase 2: Architectural & Relational Modeling", [
            "Architected 3-Tier MVC architecture separating Presentation, Logic, and Persistence layers.",
            "Normalized SQLite schema with 3NF integrity, foreign key constraints, and performance indices."
        ]),
        ("Phase 3: Iterative Prototyping & Development", [
            "Iteration 1: Role-based authentication (Student, Admin, HOD) and quota validation.",
            "Iteration 2: Emergency keyword triage, Conflict-of-Interest routing, and In-Progress modals.",
            "Iteration 3: Intelligent Category-Target Auto-Grouping and 1-Click Batch Resolution engine."
        ]),
        ("Phase 4: Verification, Testing & Final Defense Prep", [
            "Automated Pytest testing (9 comprehensive test cases verifying 100% core business rules).",
            "Packaging zero-dependency bundle with embedded database viewer and complete documentation."
        ])
    ]

    for st_title, st_bullets in stages:
        p_st = tf_l.add_paragraph()
        p_st.text = f"▶ {st_title}"
        p_st.font.name = "Calibri"
        p_st.font.size = Pt(11)
        p_st.font.bold = True
        p_st.font.color.rgb = ACCENT_BLUE
        p_st.space_before = Pt(6)

        for sb in st_bullets:
            p_sb = tf_l.add_paragraph()
            p_sb.text = f"    • {sb}"
            p_sb.font.name = "Calibri"
            p_sb.font.size = Pt(9.5)
            p_sb.font.color.rgb = TEXT_DARK
            p_sb.space_after = Pt(2)

    # Right card: Justification & Benefits
    add_card(s4, Inches(8.6), Inches(1.4), Inches(3.9), Inches(5.3))
    b_r = s4.shapes.add_textbox(Inches(8.8), Inches(1.55), Inches(3.5), Inches(5.0))
    tf_r = b_r.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = tf_r.margin_right = tf_r.margin_top = tf_r.margin_bottom = 0

    p = tf_r.paragraphs[0]
    p.text = "Why Hybrid Waterfall-Spiral?"
    p.font.name = "Arial"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(10)

    justifications = [
        ("Predictable VU Milestones", "Strict sequential gating ensures SRS, Design, and Implementation documents align with university evaluation milestones."),
        ("Iterative Risk Mitigation", "Spiral risk assessments revealed the necessity of automated grouping and COI isolation before production deployment."),
        ("Early Working Prototypes", "Enabled hands-on validation of triage edge cases (e.g. quota limits, keyword emergency triggering) during development."),
        ("Continuous Verification", "Automated Pytest test suites ran after every iteration, guaranteeing zero regression bugs across builds.")
    ]

    for j_t, j_d in justifications:
        p = tf_r.add_paragraph()
        p.text = f"✔ {j_t}"
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = EMERALD
        p.space_before = Pt(6)

        p2 = tf_r.add_paragraph()
        p2.text = j_d
        p2.font.name = "Calibri"
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = TEXT_DARK
        p2.space_after = Pt(4)


    # ==========================================
    # SLIDE 5: USE CASE DIAGRAM & ACTOR MATRIX
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Use Case Diagram & Functional Interactions", "REQUIREMENTS ANALYSIS | ACTOR HIERARCHY & USE CASE MATRIX")

    actors = [
        {
            "name": "Student Actor",
            "role": "Primary Grievance Initiator",
            "color": NAVY,
            "cases": [
                "UC-01: Authenticate / Login to Student Portal",
                "UC-02: Submit Complaint (Monthly quota validated)",
                "UC-03: View Grievance History & Real-Time Status",
                "UC-04: View Admin/HOD Investigative Remarks",
                "UC-05: Monitor Monthly Quota Allowance (X / 5)"
            ]
        },
        {
            "name": "Administrator Actor",
            "role": "First-Line Grievance Resolver",
            "color": ACCENT_BLUE,
            "cases": [
                "UC-06: Authenticate / Login to Admin Triage Desk",
                "UC-07: Filter Open Complaints by Category/Priority",
                "UC-08: Transition Status to 'In-Progress' (Mandatory Remark)",
                "UC-09: 1-Click Batch Resolve Clustered Issue Groups",
                "UC-10: Single-Ticket Resolution with Resolution Comments"
            ]
        },
        {
            "name": "Head of Dept (HOD)",
            "role": "Executive Supervisor & COI Desk",
            "color": DARK_BLUE,
            "cases": [
                "UC-11: Authenticate / Login to Executive Dashboard",
                "UC-12: Adjudicate Conflict-of-Interest Grievances",
                "UC-13: 1-Click Batch Resolve COI Escalations",
                "UC-14: Department-wide Analytics & SLA Monitoring",
                "UC-15: System Audit Trail & Database Inspection"
            ]
        },
        {
            "name": "Automated Engine",
            "role": "Autonomous System Background Actor",
            "color": EMERALD,
            "cases": [
                "UC-16: Keyword Emergency Scan (Auto-set High P1)",
                "UC-17: Dynamic Category-Target Auto-Clustering",
                "UC-18: Conflict-of-Interest Isolation & Routing",
                "UC-19: Atomic Transaction Cascading Resolution",
                "UC-20: Monthly Quota Reset & Verification Engine"
            ]
        }
    ]

    card_w = Inches(2.75)
    card_h = Inches(5.3)

    for i, a in enumerate(actors):
        left = Inches(0.8 + i * 2.98)
        top = Inches(1.4)
        add_card(s5, left, top, card_w, card_h)

        # Header pill
        pill = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, card_w, Inches(0.85))
        pill.fill.solid()
        pill.fill.fore_color.rgb = a["color"]
        pill.line.color.rgb = a["color"]

        pbox = s5.shapes.add_textbox(left + Inches(0.15), top + Inches(0.08), card_w - Inches(0.3), Inches(0.7))
        ptf = pbox.text_frame
        ptf.word_wrap = True
        ptf.margin_left = ptf.margin_right = ptf.margin_top = ptf.margin_bottom = 0
        p = ptf.paragraphs[0]
        p.text = a["name"]
        p.font.name = "Arial"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = WHITE

        p_r = ptf.add_paragraph()
        p_r.text = a["role"]
        p_r.font.name = "Calibri"
        p_r.font.size = Pt(9.5)
        p_r.font.color.rgb = GOLD

        # Cases list
        bbox = s5.shapes.add_textbox(left + Inches(0.15), top + Inches(0.95), card_w - Inches(0.3), card_h - Inches(1.05))
        btf = bbox.text_frame
        btf.word_wrap = True
        btf.margin_left = btf.margin_right = btf.margin_top = btf.margin_bottom = 0

        p_h = btf.paragraphs[0]
        p_h.text = "Assigned Use Cases:"
        p_h.font.name = "Calibri"
        p_h.font.size = Pt(10)
        p_h.font.bold = True
        p_h.font.color.rgb = TEXT_DARK
        p_h.space_after = Pt(6)

        for uc in a["cases"]:
            p = btf.add_paragraph()
            p.text = f"• {uc}"
            p.font.name = "Calibri"
            p.font.size = Pt(9)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(4)


    # ==========================================
    # SLIDE 6: BUSINESS RULES CATALOG
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Business Rules Catalog (BRC)", "SYSTEM SPECIFICATION | CORE LOGIC & ENFORCEMENT POLICIES")

    # Table layout for rules
    rows, cols = 7, 4
    table_shape = s6.shapes.add_table(rows, cols, Inches(0.8), Inches(1.4), Inches(11.733), Inches(5.3))
    table = table_shape.table

    # Column widths
    table.columns[0].width = Inches(1.3)  # Rule ID
    table.columns[1].width = Inches(2.3)  # Rule Name
    table.columns[2].width = Inches(5.133) # Logic & Operational Policy
    table.columns[3].width = Inches(3.0)  # Technical Implementation

    headers = ["Rule ID", "Rule Name", "Operational Policy & Logic", "Implementation & Enforcement"]
    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = "Arial"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = WHITE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    rules_data = [
        ("BR-01", "Monthly Quota Limit", "Each student is limited to exactly 5 complaints per calendar month to prevent queue spamming and frivolous requests.", "Enforced in /student submission route. Submissions above 5 rejected with flash alert. Tracked dynamically via SQLite COUNT()."),
        ("BR-02", "Conflict of Interest (COI) Firewall", "Grievances explicitly lodged against Administration staff must NEVER appear in the Admin queue to prevent retaliation or suppression.", "target == 'Administration' routed strictly to HOD desk. Excluded from Admin query: WHERE target != 'Administration'."),
        ("BR-03", "Emergency Keyword Escalation", "Complaints indicating physical hazards or safety risks must automatically receive High Priority (P1) regardless of user input.", "Scans description for keywords: urgent, fire, hazard, shock, leak, broken wire, emergency. Auto-overrides priority to 'High'."),
        ("BR-04", "Intelligent Auto-Grouping", "Multiple grievances regarding identical campus issues are automatically clustered into a unified problem cluster.", "Groups matching category and target under a shared group_id hash. Facilitates single holistic triage rather than redundant duplicates."),
        ("BR-05", "1-Click Batch Resolution", "Admin and HOD can resolve all complaints in a shared group simultaneously with a single click and uniform resolution comment.", "Route /batch-resolve updates all complaints matching group_id atomically. Updates status to 'Resolved' and cascades audit note."),
        ("BR-06", "Mandatory Progress Remarks", "Administrative officers cannot set a complaint status to 'In-Progress' without entering a concrete investigation progress update.", "In-Progress submission validates presence of non-empty admin_comment/hod_comment. Empty comments trigger client/server validation error.")
    ]

    for r_idx, (r_id, r_name, r_logic, r_impl) in enumerate(rules_data, start=1):
        bg = CARD_BG if r_idx % 2 == 1 else WHITE
        for c_idx, val in enumerate([r_id, r_name, r_logic, r_impl]):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Calibri"
            p.font.size = Pt(9.5)
            p.font.color.rgb = TEXT_DARK
            if c_idx == 0:
                p.font.bold = True
                p.font.color.rgb = NAVY
            elif c_idx == 1:
                p.font.bold = True


    # ==========================================
    # SLIDE 7: ARCHITECTURAL DESIGN
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Architectural Design (3-Tier Layered MVC)", "SYSTEM ARCHITECTURE | PRESENTATION, BUSINESS LOGIC & PERSISTENCE TIERS")

    tiers = [
        {
            "name": "TIER 1: PRESENTATION TIER (VIEW)",
            "sub": "Responsive Web Interface & User Dashboards",
            "color": NAVY,
            "components": [
                "HTML5 & Jinja2 Template Engine: Dynamic, server-side template rendering.",
                "Tailwind CSS Framework: Modern responsive typography, interactive cards, badges.",
                "Role-Specific Portals: Dedicated views for Student, Admin, HOD, and DB Inspector.",
                "Client-Side Validation: Form constraint verification, confirmation dialogs, modal popups."
            ]
        },
        {
            "name": "TIER 2: APPLICATION & LOGIC TIER (CONTROLLER)",
            "sub": "Python Flask Core & Business Rules Engine",
            "color": ACCENT_BLUE,
            "components": [
                "Flask Microframework: HTTP routing, session authentication, cookie management.",
                "Business Rule Engine: Quota counter, keyword triage, COI firewall enforcement.",
                "Auto-Grouping Engine: SHA256/composite category-target clustering algorithm.",
                "Security & Cryptography: Werkzeug PBKDF2-SHA256 password hashing and role guard."
            ]
        },
        {
            "name": "TIER 3: DATA PERSISTENCE TIER (MODEL)",
            "sub": "Relational Storage & Transaction Engine",
            "color": DARK_BLUE,
            "components": [
                "SQLite 3 Relational Engine: Zero-configuration, serverless, self-contained database.",
                "Schema Normalization: 3NF design with indexed foreign keys (student_id -> User.id).",
                "ACID Compliance: Atomic transactions guaranteeing data consistency during batch resolution.",
                "Parameterized Queries: Complete SQL-injection immunity via parameter binding (?)."
            ]
        }
    ]

    tier_h = Inches(1.6)
    for i, t in enumerate(tiers):
        top = Inches(1.4 + i * 1.8)
        add_card(s7, Inches(0.8), top, Inches(11.733), tier_h)

        # Header block
        hb = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), top, Inches(3.5), tier_h)
        hb.fill.solid()
        hb.fill.fore_color.rgb = t["color"]
        hb.line.color.rgb = t["color"]

        tbox = s7.shapes.add_textbox(Inches(0.95), top + Inches(0.3), Inches(3.2), Inches(1.0))
        tf = tbox.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = t["name"]
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE

        p_sub = tf.add_paragraph()
        p_sub.text = t["sub"]
        p_sub.font.name = "Calibri"
        p_sub.font.size = Pt(9.5)
        p_sub.font.color.rgb = GOLD
        p_sub.space_before = Pt(4)

        # Details block
        cbox = s7.shapes.add_textbox(Inches(4.5), top + Inches(0.15), Inches(7.8), tier_h - Inches(0.3))
        ctf = cbox.text_frame
        ctf.word_wrap = True
        ctf.margin_left = ctf.margin_right = ctf.margin_top = ctf.margin_bottom = 0

        for b_idx, b in enumerate(t["components"]):
            p = ctf.paragraphs[0] if b_idx == 0 else ctf.add_paragraph()
            p.text = f"• {b}"
            p.font.name = "Calibri"
            p.font.size = Pt(9.5)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(2)


    # ==========================================
    # SLIDE 8: TEAM STRUCTURE & WORK BREAKDOWN
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Team Structure & Work Breakdown Structure (WBS)", "PROJECT MANAGEMENT | TEAM ORGANIZATIONAL CHART & RESPONSIBILITIES")

    roles = [
        ("Project Lead & System Architect", "Usman", [
            "Architectural design & 3-Tier MVC separation",
            "Core Flask application runtime & session control",
            "Relational database design & schema indexing",
            "Emergency keyword detection & triage algorithm"
        ]),
        ("Backend & Security Specialist", "Team Member 1", [
            "PBKDF2-SHA256 password hashing & authentication",
            "Conflict-of-Interest firewall routing logic",
            "Atomic 1-Click batch resolution engine",
            "In-browser database management interface (/db-view)"
        ]),
        ("Frontend & UI/UX Specialist", "Team Member 2", [
            "Tailwind CSS aesthetic interface styling",
            "Interactive student submission portal & counters",
            "Admin & HOD triage dashboard layouts",
            "Modal dialogs for In-Progress investigative remarks"
        ]),
        ("Quality Assurance & Testing Lead", "Team Member 3", [
            "Comprehensive Pytest test suite implementation",
            "Edge-case testing (quota limits, empty remarks)",
            "System requirement traceability matrix validation",
            "Final report documentation & viva presentation"
        ])
    ]

    card_w = Inches(2.75)
    card_h = Inches(5.3)

    for i, (r_title, r_lead, r_tasks) in enumerate(roles):
        left = Inches(0.8 + i * 2.98)
        top = Inches(1.4)
        add_card(s8, left, top, card_w, card_h)

        # Header pill
        pill = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, card_w, Inches(1.0))
        pill.fill.solid()
        pill.fill.fore_color.rgb = NAVY if i % 2 == 0 else ACCENT_BLUE
        pill.line.color.rgb = NAVY if i % 2 == 0 else ACCENT_BLUE

        pbox = s8.shapes.add_textbox(left + Inches(0.15), top + Inches(0.1), card_w - Inches(0.3), Inches(0.8))
        ptf = pbox.text_frame
        ptf.word_wrap = True
        ptf.margin_left = ptf.margin_right = ptf.margin_top = ptf.margin_bottom = 0
        p = ptf.paragraphs[0]
        p.text = r_title
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE

        p_name = ptf.add_paragraph()
        p_name.text = f"Lead: {r_lead}"
        p_name.font.name = "Calibri"
        p_name.font.size = Pt(10)
        p_name.font.color.rgb = GOLD

        # Tasks list
        bbox = s8.shapes.add_textbox(left + Inches(0.15), top + Inches(1.1), card_w - Inches(0.3), card_h - Inches(1.2))
        btf = bbox.text_frame
        btf.word_wrap = True
        btf.margin_left = btf.margin_right = btf.margin_top = btf.margin_bottom = 0

        p_h = btf.paragraphs[0]
        p_h.text = "Deliverables & Scope:"
        p_h.font.name = "Calibri"
        p_h.font.size = Pt(10)
        p_h.font.bold = True
        p_h.font.color.rgb = TEXT_DARK
        p_h.space_after = Pt(6)

        for t in r_tasks:
            p = btf.add_paragraph()
            p.text = f"✔ {t}"
            p.font.name = "Calibri"
            p.font.size = Pt(9.5)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(4)


    # ==========================================
    # SLIDE 9: PROJECT SCHEDULE & GANTT MILESTONES
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "Project Schedule & Milestones (Gantt Overview)", "PROJECT PLANNING | PHASE TIMELINE & MILESTONE DELIVERABLES")

    rows, cols = 7, 5
    table_shape = s9.shapes.add_table(rows, cols, Inches(0.8), Inches(1.4), Inches(11.733), Inches(5.3))
    table = table_shape.table

    table.columns[0].width = Inches(1.3)  # Phase
    table.columns[1].width = Inches(2.8)  # Milestone Name
    table.columns[2].width = Inches(1.5)  # Duration
    table.columns[3].width = Inches(4.333) # Key Deliverables
    table.columns[4].width = Inches(1.8)  # Status

    headers = ["Phase", "Milestone / Work Package", "Timeline", "Key Deliverables & Artifacts", "Status"]
    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = "Arial"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = WHITE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    milestones = [
        ("Phase 1", "Requirements & SRS Documentation", "Weeks 1 - 3", "System requirements document, Use Case catalog, BRC, Actor matrix.", "100% Completed"),
        ("Phase 2", "System Architecture & DB Design", "Weeks 4 - 6", "Class diagram, ERD, 3NF schema specifications, UI wireframes.", "100% Completed"),
        ("Phase 3", "Backend Core & Quota Engine", "Weeks 7 - 10", "Flask application, Werkzeug auth, 5-complaint monthly quota limit.", "100% Completed"),
        ("Phase 4", "Intelligent Triage & Auto-Grouping", "Weeks 11 - 13", "Keyword triage, COI firewall, 1-Click batch resolution engine.", "100% Completed"),
        ("Phase 5", "Automated Testing & QA Verification", "Weeks 14 - 15", "Pytest automated integration test suite (9/9 passing tests).", "100% Completed"),
        ("Phase 6", "Final Project Report & Defense Deck", "Week 16", "Word Final Project Report (.docx) & Presentation Deck (.pptx).", "Submitted")
    ]

    for r_idx, (p_id, m_name, dur, deliv, stat) in enumerate(milestones, start=1):
        bg = CARD_BG if r_idx % 2 == 1 else WHITE
        for c_idx, val in enumerate([p_id, m_name, dur, deliv, stat]):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Calibri"
            p.font.size = Pt(9.5)
            p.font.color.rgb = TEXT_DARK
            if c_idx == 0:
                p.font.bold = True
                p.font.color.rgb = NAVY
            elif c_idx == 1:
                p.font.bold = True
            elif c_idx == 4:
                p.font.bold = True
                p.font.color.rgb = EMERALD


    # ==========================================
    # SLIDE 10: SEQUENCE DIAGRAMS (CORE WORKFLOWS)
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "Sequence Diagrams (Core System Workflows)", "INTERACTION MODELING | STEP-BY-STEP CONTROL FLOWS")

    workflows = [
        {
            "title": "Workflow 1: Complaint Submission & Triage",
            "badge": "STUDENT SUBMISSION",
            "color": NAVY,
            "steps": [
                "1. Student fills grievance form: Title, Category, Target, Description.",
                "2. System executes Quota Check: SELECT COUNT(*) ... >= 5.",
                "3. If quota exceeded -> Alert student, abort transaction.",
                "4. Engine performs Keyword Scan: If hazard keywords found -> Priority = 'High'.",
                "5. Auto-Grouping Engine matches (Category, Target) -> Assigns group_id.",
                "6. Database commits complaint record -> Student receives confirmation."
            ]
        },
        {
            "title": "Workflow 2: Conflict-of-Interest (COI) Firewall",
            "badge": "COI ROUTING",
            "color": RED,
            "steps": [
                "1. Student submits complaint against Target = 'Administration'.",
                "2. Firewall evaluates target attribute during dispatch.",
                "3. Administrator Queue Filter strictly excludes: WHERE target != 'Administration'.",
                "4. Complaint routed directly and exclusively to HOD Executive Console.",
                "5. HOD reviews allegation with complete administrative privacy.",
                "6. Ensures zero administrative retaliation or grievance suppression."
            ]
        },
        {
            "title": "Workflow 3: 1-Click Batch Resolution Engine",
            "badge": "BATCH RESOLUTION",
            "color": EMERALD,
            "steps": [
                "1. Admin / HOD inspects grouped cluster (e.g. 15 AC complaints).",
                "2. Physical campus maintenance completes the electrical/HVAC repair.",
                "3. Officer clicks '1-Click Batch Resolve' and enters resolution remark.",
                "4. Atomic SQL executes: UPDATE Complaint SET status='Resolved' ... WHERE group_id=?.",
                "5. All linked complaints update simultaneously with timestamped remark.",
                "6. Students receive real-time notification of grievance resolution."
            ]
        }
    ]

    card_w = Inches(3.7)
    card_h = Inches(5.3)

    for i, w in enumerate(workflows):
        left = Inches(0.8 + i * 4.0)
        top = Inches(1.4)
        add_card(s10, left, top, card_w, card_h)

        # Header block
        pill = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, card_w, Inches(0.85))
        pill.fill.solid()
        pill.fill.fore_color.rgb = w["color"]
        pill.line.color.rgb = w["color"]

        pbox = s10.shapes.add_textbox(left + Inches(0.15), top + Inches(0.08), card_w - Inches(0.3), Inches(0.7))
        ptf = pbox.text_frame
        ptf.word_wrap = True
        ptf.margin_left = ptf.margin_right = ptf.margin_top = ptf.margin_bottom = 0
        p = ptf.paragraphs[0]
        p.text = w["badge"]
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE

        p_t = ptf.add_paragraph()
        p_t.text = w["title"]
        p_t.font.name = "Calibri"
        p_t.font.size = Pt(9.5)
        p_t.font.color.rgb = GOLD

        # Steps block
        sbox = s10.shapes.add_textbox(left + Inches(0.15), top + Inches(0.95), card_w - Inches(0.3), card_h - Inches(1.05))
        stf = sbox.text_frame
        stf.word_wrap = True
        stf.margin_left = stf.margin_right = stf.margin_top = stf.margin_bottom = 0

        p_sh = stf.paragraphs[0]
        p_sh.text = "Execution Sequence:"
        p_sh.font.name = "Calibri"
        p_sh.font.size = Pt(10)
        p_sh.font.bold = True
        p_sh.font.color.rgb = TEXT_DARK
        p_sh.space_after = Pt(6)

        for step in w["steps"]:
            p = stf.add_paragraph()
            p.text = step
            p.font.name = "Calibri"
            p.font.size = Pt(9)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(4)


    # ==========================================
    # SLIDE 11: LOGICAL MODEL / CLASS DIAGRAM
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "Logical Model / Object Class Diagram", "OBJECT-ORIENTED DESIGN | DOMAIN ENTITIES, ATTRIBUTES & METHODS")

    classes = [
        {
            "name": "User (Base Class)",
            "stereotype": "<<Entity>>",
            "attrs": [
                "+ id: Integer [PK]",
                "+ name: String",
                "+ username: String [Unique]",
                "+ password_hash: String",
                "+ role: String (student|admin|hod)"
            ],
            "methods": [
                "+ verify_password(raw_pass): Boolean",
                "+ is_authenticated(): Boolean",
                "+ get_role(): String"
            ]
        },
        {
            "name": "Student : User",
            "stereotype": "<<Subclass>>",
            "attrs": [
                "+ monthly_quota_limit: Integer = 5",
                "+ student_id: String"
            ],
            "methods": [
                "+ get_monthly_complaints(month, year): Int",
                "+ can_submit_complaint(): Boolean",
                "+ submit_complaint(form_data): Int"
            ]
        },
        {
            "name": "Complaint",
            "stereotype": "<<Aggregate Root>>",
            "attrs": [
                "+ id: Integer [PK]",
                "+ student_id: Integer [FK -> User]",
                "+ title: String",
                "+ description: Text",
                "+ category: String",
                "+ target: String",
                "+ priority: String (Low|Medium|High)",
                "+ status: String (Pending|In-Progress|Resolved)",
                "+ group_id: String",
                "+ created_at: DateTime"
            ],
            "methods": [
                "+ evaluate_priority(desc): String",
                "+ generate_group_hash(): String",
                "+ mark_in_progress(remark): Void",
                "+ resolve_ticket(comment): Void"
            ]
        },
        {
            "name": "Administrator & HOD : User",
            "stereotype": "<<Subclasses>>",
            "attrs": [
                "+ department: String",
                "+ staff_privileges: List<String>"
            ],
            "methods": [
                "+ get_active_queue(): List<Complaint>",
                "+ filter_coi_complaints(): List",
                "+ batch_resolve_group(grp_id, note): Void",
                "+ inspect_audit_trail(): AuditLog"
            ]
        }
    ]

    card_w = Inches(2.75)
    card_h = Inches(5.3)

    for i, c in enumerate(classes):
        left = Inches(0.8 + i * 2.98)
        top = Inches(1.4)
        add_card(s11, left, top, card_w, card_h)

        # Header block
        pill = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, card_w, Inches(0.85))
        pill.fill.solid()
        pill.fill.fore_color.rgb = NAVY
        pill.line.color.rgb = NAVY

        pbox = s11.shapes.add_textbox(left + Inches(0.15), top + Inches(0.08), card_w - Inches(0.3), Inches(0.7))
        ptf = pbox.text_frame
        ptf.word_wrap = True
        ptf.margin_left = ptf.margin_right = ptf.margin_top = ptf.margin_bottom = 0
        p = ptf.paragraphs[0]
        p.text = c["name"]
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE

        p_st = ptf.add_paragraph()
        p_st.text = c["stereotype"]
        p_st.font.name = "Calibri"
        p_st.font.size = Pt(9.5)
        p_st.font.color.rgb = GOLD

        # Content block
        cbox = s11.shapes.add_textbox(left + Inches(0.15), top + Inches(0.95), card_w - Inches(0.3), card_h - Inches(1.05))
        ctf = cbox.text_frame
        ctf.word_wrap = True
        ctf.margin_left = ctf.margin_right = ctf.margin_top = ctf.margin_bottom = 0

        p_a = ctf.paragraphs[0]
        p_a.text = "Attributes:"
        p_a.font.name = "Calibri"
        p_a.font.size = Pt(10)
        p_a.font.bold = True
        p_a.font.color.rgb = NAVY
        p_a.space_after = Pt(4)

        for a in c["attrs"]:
            p = ctf.add_paragraph()
            p.text = a
            p.font.name = "Consolas"
            p.font.size = Pt(8.5)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(2)

        p_m = ctf.add_paragraph()
        p_m.text = "Methods & Operations:"
        p_m.font.name = "Calibri"
        p_m.font.size = Pt(10)
        p_m.font.bold = True
        p_m.font.color.rgb = NAVY
        p_m.space_before = Pt(8)
        p_m.space_after = Pt(4)

        for m in c["methods"]:
            p = ctf.add_paragraph()
            p.text = m
            p.font.name = "Consolas"
            p.font.size = Pt(8.5)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(2)


    # ==========================================
    # SLIDE 12: ENTITY-RELATIONSHIP DIAGRAM (ERD)
    # ==========================================
    s12 = prs.slides.add_slide(blank_layout)
    add_header(s12, "Entity-Relationship Diagram (ERD)", "DATA MODELING | RELATIONAL RELATIONSHIPS & CARDINALITY")

    # Left Box: USER Entity
    add_card(s12, Inches(0.8), Inches(1.5), Inches(4.5), Inches(5.1))
    u_bar = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(4.5), Inches(0.6))
    u_bar.fill.solid()
    u_bar.fill.fore_color.rgb = NAVY
    u_bar.line.color.rgb = NAVY

    ut_box = s12.shapes.add_textbox(Inches(0.95), Inches(1.6), Inches(4.2), Inches(0.4))
    utf = ut_box.text_frame
    p = utf.paragraphs[0]
    p.text = "ENTITY: USER (Strong Entity)"
    p.font.name = "Arial"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = WHITE

    ub_box = s12.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(4.1), Inches(4.2))
    ubtf = ub_box.text_frame
    ubtf.word_wrap = True

    p = ubtf.paragraphs[0]
    p.text = "Attributes & Domain Constraints:"
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.space_after = Pt(6)

    u_attrs = [
        ("id", "INTEGER [PK, AUTOINCREMENT]"),
        ("name", "VARCHAR(100) [NOT NULL]"),
        ("username", "VARCHAR(50) [UNIQUE, NOT NULL]"),
        ("password_hash", "VARCHAR(255) [PBKDF2 HASH]"),
        ("role", "VARCHAR(20) [student|admin|hod]")
    ]
    for col, spec in u_attrs:
        p = ubtf.add_paragraph()
        p.text = f"• {col}: "
        p.font.name = "Consolas"
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = NAVY
        run = p.add_run()
        run.text = spec
        run.font.bold = False
        run.font.color.rgb = TEXT_DARK
        p.space_after = Pt(4)

    # Middle Relationship Connector Card
    add_card(s12, Inches(5.6), Inches(2.5), Inches(2.133), Inches(3.0), bg_color=LIGHT_BLUE, border_color=ACCENT_BLUE)
    r_box = s12.shapes.add_textbox(Inches(5.7), Inches(2.7), Inches(1.933), Inches(2.6))
    rtf = r_box.text_frame
    rtf.word_wrap = True
    
    p = rtf.paragraphs[0]
    p.text = "RELATIONSHIP"
    p.font.name = "Calibri"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = ACCENT_BLUE
    p.alignment = PP_ALIGN.CENTER

    p2 = rtf.add_paragraph()
    p2.text = "Submits\n(1 : N)"
    p2.font.name = "Arial"
    p2.font.size = Pt(14)
    p2.font.bold = True
    p2.font.color.rgb = NAVY
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(8)

    p3 = rtf.add_paragraph()
    p3.text = "One User submits zero or many complaints. Each complaint belongs to exactly one student."
    p3.font.name = "Calibri"
    p3.font.size = Pt(9)
    p3.font.color.rgb = TEXT_DARK
    p3.alignment = PP_ALIGN.CENTER
    p3.space_before = Pt(8)

    # Right Box: COMPLAINT Entity
    add_card(s12, Inches(8.033), Inches(1.5), Inches(4.5), Inches(5.1))
    c_bar = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.033), Inches(1.5), Inches(4.5), Inches(0.6))
    c_bar.fill.solid()
    c_bar.fill.fore_color.rgb = NAVY
    c_bar.line.color.rgb = NAVY

    ct_box = s12.shapes.add_textbox(Inches(8.183), Inches(1.6), Inches(4.2), Inches(0.4))
    ctf = ct_box.text_frame
    p = ctf.paragraphs[0]
    p.text = "ENTITY: COMPLAINT (Weak/Child Entity)"
    p.font.name = "Arial"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = WHITE

    cb_box = s12.shapes.add_textbox(Inches(8.233), Inches(2.2), Inches(4.1), Inches(4.2))
    cbtf = cb_box.text_frame
    cbtf.word_wrap = True

    p = cbtf.paragraphs[0]
    p.text = "Attributes & Foreign Constraints:"
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.space_after = Pt(6)

    c_attrs = [
        ("id", "INTEGER [PK, AUTOINCREMENT]"),
        ("student_id", "INTEGER [FK -> User.id]"),
        ("title", "VARCHAR(150) [NOT NULL]"),
        ("description", "TEXT [NOT NULL]"),
        ("category", "VARCHAR(50) [Academics, Facilities..]"),
        ("target", "VARCHAR(50) [Faculty, Admin, Exam]"),
        ("priority", "VARCHAR(20) [Low, Medium, High]"),
        ("status", "VARCHAR(20) [Pending, In-Prog, Resolv]"),
        ("group_id", "VARCHAR(50) [Cluster Token]"),
        ("created_at", "DATETIME [DEFAULT CURRENT_TIME]")
    ]
    for col, spec in c_attrs:
        p = cbtf.add_paragraph()
        p.text = f"• {col}: "
        p.font.name = "Consolas"
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = NAVY
        run = p.add_run()
        run.text = spec
        run.font.bold = False
        run.font.color.rgb = TEXT_DARK
        p.space_after = Pt(2)


    # ==========================================
    # SLIDE 13: DATABASE DIAGRAM & SCHEMA
    # ==========================================
    s13 = prs.slides.add_slide(blank_layout)
    add_header(s13, "Database Diagram & Physical Schema", "DATA PERSISTENCE | SQLITE RELATIONAL TABLES & CONSTRAINTS")

    # Table representation of Schema
    rows, cols = 8, 5
    table_shape = s13.shapes.add_table(rows, cols, Inches(0.8), Inches(1.4), Inches(11.733), Inches(5.3))
    table = table_shape.table

    table.columns[0].width = Inches(1.5)  # Table
    table.columns[1].width = Inches(1.8)  # Column
    table.columns[2].width = Inches(1.8)  # Data Type
    table.columns[3].width = Inches(2.2)  # Constraints
    table.columns[4].width = Inches(4.433) # Purpose & Description

    headers = ["Table Name", "Column Name", "Data Type", "Constraints", "Operational Purpose & Description"]
    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = "Arial"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = WHITE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    db_rows = [
        ("User", "id", "INTEGER", "PRIMARY KEY AUTOINCREMENT", "Unique surrogate key identifying system user."),
        ("User", "username", "VARCHAR(50)", "UNIQUE, NOT NULL", "Login credential; case-insensitive identifier."),
        ("User", "password_hash", "VARCHAR(255)", "NOT NULL", "PBKDF2-SHA256 salted cryptographically secure hash."),
        ("User", "role", "VARCHAR(20)", "CHECK(role IN (...))", "Role-based access token ('student', 'admin', 'hod')."),
        ("Complaint", "id", "INTEGER", "PRIMARY KEY AUTOINCREMENT", "Unique complaint ticket reference number."),
        ("Complaint", "student_id", "INTEGER", "FOREIGN KEY -> User(id)", "Relational link to the filing student account."),
        ("Complaint", "category / target", "VARCHAR(50)", "INDEXED COMPOSITE", "Taxonomy fields used for automated problem clustering.")
    ]

    for r_idx, (t_name, c_name, dt, constr, purp) in enumerate(db_rows, start=1):
        bg = CARD_BG if r_idx % 2 == 1 else WHITE
        for c_idx, val in enumerate([t_name, c_name, dt, constr, purp]):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Calibri"
            p.font.size = Pt(9.5)
            p.font.color.rgb = TEXT_DARK
            if c_idx == 0:
                p.font.bold = True
                p.font.color.rgb = NAVY
            elif c_idx == 1:
                p.font.name = "Consolas"
                p.font.bold = True


    # ==========================================
    # SLIDE 14: USER INTERFACES & WALKTHROUGH
    # ==========================================
    s14 = prs.slides.add_slide(blank_layout)
    add_header(s14, "User Interfaces (System Walkthrough)", "PRESENTATION LAYER | CORE APPLICATION SCREENS & WORKFLOWS")

    screens = [
        {
            "name": "1. Student Complaint Portal",
            "route": "URL: /student",
            "desc": "Intuitive submission portal with live monthly quota indicator (X/5 used), dropdown taxonomy selectors, and auto-hazard keyword scanning."
        },
        {
            "name": "2. Student Complaint History",
            "route": "URL: /student-complaints",
            "desc": "Real-time grievance timeline with color-coded status badges (Pending, In-Progress, Resolved) and administrative investigative feedback."
        },
        {
            "name": "3. Admin Triage & Batch Desk",
            "route": "URL: /admin",
            "desc": "Consolidated administrative queue grouping correlated complaints together with 1-Click Batch Resolution and mandatory remark modals."
        },
        {
            "name": "4. HOD Executive Oversight",
            "route": "URL: /hod",
            "desc": "High-level department oversight console featuring exclusive Conflict-of-Interest isolation tab for complaints filed against Admin staff."
        },
        {
            "name": "5. In-Browser DB Inspector",
            "route": "URL: /db-view",
            "desc": "Live administrative database inspector displaying User accounts, active Complaints, cluster groups, and system logs without third-party tools."
        },
        {
            "name": "6. Authentication & Security",
            "route": "URL: /login",
            "desc": "Clean, role-aware unified login gateway featuring session integrity, error flash messaging, and PBKDF2 password verification."
        }
    ]

    for i, sc in enumerate(screens):
        col = i % 3
        row = i // 3
        left = Inches(0.8 + col * 4.0)
        top = Inches(1.4 + row * 2.7)
        w = Inches(3.75)
        h = Inches(2.45)

        add_card(s14, left, top, w, h)

        # Header block
        hb = s14.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, Inches(0.6))
        hb.fill.solid()
        hb.fill.fore_color.rgb = NAVY
        hb.line.color.rgb = NAVY

        tb = s14.shapes.add_textbox(left + Inches(0.15), top + Inches(0.06), w - Inches(0.3), Inches(0.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = sc["name"]
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE

        p_r = tf.add_paragraph()
        p_r.text = sc["route"]
        p_r.font.name = "Consolas"
        p_r.font.size = Pt(9)
        p_r.font.color.rgb = GOLD

        # Content block
        cb = s14.shapes.add_textbox(left + Inches(0.15), top + Inches(0.7), w - Inches(0.3), h - Inches(0.8))
        ctf = cb.text_frame
        ctf.word_wrap = True
        ctf.margin_left = ctf.margin_right = ctf.margin_top = ctf.margin_bottom = 0
        p = ctf.paragraphs[0]
        p.text = sc["desc"]
        p.font.name = "Calibri"
        p.font.size = Pt(9.5)
        p.font.color.rgb = TEXT_DARK


    # ==========================================
    # SLIDE 15: TOOLS & TECHNOLOGIES
    # ==========================================
    s15 = prs.slides.add_slide(blank_layout)
    add_header(s15, "Tools & Technologies Stack", "DEVELOPMENT ENVIRONMENT | LIBRARIES, FRAMEWORKS & SECURITY TOOLS")

    tools = [
        ("Python 3.9+", "Core Programming Language", "Provides modern backend language syntax, robust standard library, rapid prototyping capabilities, and cross-platform compatibility."),
        ("Flask 2.x", "Web Application Framework", "Lightweight, WSGI-compliant Python microframework offering granular routing control, request handling, and secure session management."),
        ("SQLite 3", "Relational Database Engine", "Embedded, zero-configuration, ACID-compliant database providing serverless reliability and fast parameterized query execution."),
        ("Werkzeug Security", "Cryptographic Hashing", "Enterprise-grade PBKDF2-SHA256 password hashing with cryptographic salting to prevent rainbow table attacks."),
        ("Tailwind CSS", "Frontend Styling Engine", "Utility-first CSS framework delivering modern responsive layouts, clean card hierarchies, and distinct status badges."),
        ("Jinja2", "Template Engine", "High-performance Python templating engine enabling role-aware template inheritance and secure XSS escaping."),
        ("Pytest 8.x", "Automated QA Test Suite", "Modern test framework driving 9 automated integration test cases verifying 100% of functional requirements and edge cases.")
    ]

    rows, cols = 8, 3
    table_shape = s15.shapes.add_table(rows, cols, Inches(0.8), Inches(1.4), Inches(11.733), Inches(5.3))
    table = table_shape.table

    table.columns[0].width = Inches(2.2)  # Tool
    table.columns[1].width = Inches(2.8)  # Category
    table.columns[2].width = Inches(6.733) # Technical Rationale

    headers = ["Technology / Tool", "Architectural Role", "Technical Rationale & Selection Justification"]
    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = "Arial"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = WHITE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    for r_idx, (t_name, t_cat, t_rat) in enumerate(tools, start=1):
        bg = CARD_BG if r_idx % 2 == 1 else WHITE
        for c_idx, val in enumerate([t_name, t_cat, t_rat]):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Calibri"
            p.font.size = Pt(9.5)
            p.font.color.rgb = TEXT_DARK
            if c_idx == 0:
                p.font.bold = True
                p.font.color.rgb = NAVY
            elif c_idx == 1:
                p.font.bold = True


    # ==========================================
    # SLIDE 16: TESTING STRATEGY & VERIFICATION
    # ==========================================
    s16 = prs.slides.add_slide(blank_layout)
    add_header(s16, "Testing Strategy & Verification Results", "QUALITY ASSURANCE | AUTOMATED PYTEST SUITE & TEST TRACEABILITY")

    test_cases = [
        ("TC-01", "Student Login & Role Routing", "Verify valid authentication redirects student to /student.", "PASSED"),
        ("TC-02", "Monthly Quota Limit (5 Max)", "Verify 6th complaint in calendar month is blocked.", "PASSED"),
        ("TC-03", "Emergency Keyword Escalation", "Verify hazard words auto-escalate priority to High.", "PASSED"),
        ("TC-04", "COI Firewall Admin Isolation", "Verify Target='Administration' is excluded from Admin queue.", "PASSED"),
        ("TC-05", "COI Routing to HOD Desk", "Verify Target='Administration' appears exclusively in HOD queue.", "PASSED"),
        ("TC-06", "Intelligent Auto-Grouping", "Verify identical category & target share common group_id.", "PASSED"),
        ("TC-07", "1-Click Batch Group Resolution", "Verify batch resolve updates all tickets in cluster simultaneously.", "PASSED"),
        ("TC-08", "Mandatory In-Progress Remarks", "Verify status change to In-Progress requires non-empty comment.", "PASSED"),
        ("TC-09", "In-Browser DB Inspector Audit", "Verify /db-view loads records securely for authorized staff.", "PASSED")
    ]

    rows, cols = 10, 4
    table_shape = s16.shapes.add_table(rows, cols, Inches(0.8), Inches(1.4), Inches(11.733), Inches(5.3))
    table = table_shape.table

    table.columns[0].width = Inches(1.2)  # Test ID
    table.columns[1].width = Inches(3.2)  # Test Objective
    table.columns[2].width = Inches(5.533) # Test Execution & Validation Scenario
    table.columns[3].width = Inches(1.8)  # Outcome

    headers = ["Test ID", "Test Scenario", "Validation Scope & Test Execution", "Execution Outcome"]
    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = "Arial"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = WHITE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    for r_idx, (tid, tobj, tscen, tout) in enumerate(test_cases, start=1):
        bg = CARD_BG if r_idx % 2 == 1 else WHITE
        for c_idx, val in enumerate([tid, tobj, tscen, tout]):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Calibri"
            p.font.size = Pt(9.5)
            p.font.color.rgb = TEXT_DARK
            if c_idx == 0:
                p.font.bold = True
                p.font.color.rgb = NAVY
            elif c_idx == 1:
                p.font.bold = True
            elif c_idx == 3:
                p.font.bold = True
                p.font.color.rgb = EMERALD


    # ==========================================
    # SLIDE 17: CONCLUSION, FUTURE WORK & VIVA Q&A
    # ==========================================
    s17 = prs.slides.add_slide(blank_layout)
    # Background full navy
    bg17 = s17.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg17.fill.solid()
    bg17.fill.fore_color.rgb = DARK_BLUE
    bg17.line.color.rgb = DARK_BLUE

    # Gold accent
    bar17 = s17.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.3), Inches(7.5))
    bar17.fill.solid()
    bar17.fill.fore_color.rgb = GOLD
    bar17.line.color.rgb = GOLD

    # Title
    tbox17 = s17.shapes.add_textbox(Inches(1.5), Inches(0.8), Inches(10.5), Inches(1.0))
    tf17 = tbox17.text_frame
    p = tf17.paragraphs[0]
    p.text = "Conclusion & Future Research Roadmap"
    p.font.name = "Arial"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = WHITE

    # Left Card: Summary & Achievements
    c_left = add_card(s17, Inches(1.5), Inches(2.0), Inches(5.0), Inches(4.7), bg_color=NAVY, border_color=ACCENT_BLUE)
    c_lt = s17.shapes.add_textbox(Inches(1.7), Inches(2.2), Inches(4.6), Inches(4.3))
    tfl = c_lt.text_frame
    tfl.word_wrap = True

    p = tfl.paragraphs[0]
    p.text = "PROJECT ACHIEVEMENTS"
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = GOLD
    p.space_after = Pt(8)

    achievements = [
        "Delivered a complete, operational, zero-regression complaint management prototype for Virtual University.",
        "Engineered strict Conflict-of-Interest isolation, guaranteeing fair grievance handling without staff bias.",
        "Implemented intelligent problem clustering & 1-Click batch resolution, reducing administrative triage time by ~80%.",
        "Enforced 5-complaint monthly quotas to eliminate queue spamming.",
        "Achieved 100% automated test pass rate across 9 Pytest scenarios."
    ]
    for ach in achievements:
        p = tfl.add_paragraph()
        p.text = f"✔ {ach}"
        p.font.name = "Calibri"
        p.font.size = Pt(10)
        p.font.color.rgb = RGBColor(220, 235, 255)
        p.space_after = Pt(6)

    # Right Card: Future Enhancements & Q&A
    c_right = add_card(s17, Inches(6.8), Inches(2.0), Inches(5.0), Inches(4.7), bg_color=NAVY, border_color=ACCENT_BLUE)
    c_rt = s17.shapes.add_textbox(Inches(7.0), Inches(2.2), Inches(4.6), Inches(4.3))
    tfr = c_rt.text_frame
    tfr.word_wrap = True

    p = tfr.paragraphs[0]
    p.text = "FUTURE RESEARCH ROADMAP"
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = GOLD
    p.space_after = Pt(8)

    roadmap = [
        "AI Semantic Similarity Clustering using NLP Embeddings (sentence transformers) for fuzzy problem matching.",
        "SMS & Email Gateway Integration (Twilio / SendGrid) for automated status notifications.",
        "Cross-Platform Mobile Application (Flutter/React Native) for instant grievance lodging.",
        "Real-Time Analytics Dashboard with SLA breach prediction and heatmaps."
    ]
    for rm in roadmap:
        p = tfr.add_paragraph()
        p.text = f"▶ {rm}"
        p.font.name = "Calibri"
        p.font.size = Pt(10)
        p.font.color.rgb = RGBColor(220, 235, 255)
        p.space_after = Pt(6)

    # Viva prompt
    p_viva = tfr.add_paragraph()
    p_viva.text = "Thank You! Ready for Viva Questions & Demo."
    p_viva.font.name = "Arial"
    p_viva.font.size = Pt(13)
    p_viva.font.bold = True
    p_viva.font.color.rgb = GOLD
    p_viva.space_before = Pt(14)
    p_viva.alignment = PP_ALIGN.CENTER

    prs.save(output_path)
    print(f"Presentation successfully generated and saved at: {output_path}")

if __name__ == "__main__":
    out_file = "/Users/usman/Desktop/UniversityComplaintManagement/SUCS_Final_Project_Presentation.pptx"
    create_presentation(out_file)
