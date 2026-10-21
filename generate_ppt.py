import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
# Set 16:9 Widescreen dimensions
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Color Palette
NAVY_DARK = RGBColor(15, 23, 42)      # #0f172a
NAVY_PRIMARY = RGBColor(30, 58, 138)  # #1e3a8a
INDIGO_ACCENT = RGBColor(79, 70, 229) # #4f46e5
EMERALD_GREEN = RGBColor(5, 150, 105) # #059669
CARD_BG = RGBColor(248, 250, 252)     # #f8fafc
CARD_BORDER = RGBColor(203, 213, 225) # #cbd5e1
TEXT_DARK = RGBColor(30, 41, 59)      # #1e293b
TEXT_MUTED = RGBColor(100, 116, 139)  # #64748b
WHITE = RGBColor(255, 255, 255)
ACCENT_AMBER = RGBColor(217, 119, 6)  # #d97706
ACCENT_BLUE_LIGHT = RGBColor(238, 242, 255) # #eef2ff

blank_layout = prs.slide_layouts[6]

def add_header(slide, title_text, category_text="FOUNDATIONS OF ARTIFICIAL INTELLIGENCE"):
    cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.35))
    tf_cat = cat_box.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.name = "Calibri"
    p_cat.font.size = Pt(11)
    p_cat.font.bold = True
    p_cat.font.color.rgb = INDIGO_ACCENT

    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(11.733), Inches(0.75))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = "Georgia"
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = NAVY_PRIMARY

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.45), Inches(11.733), Inches(0.03))
    line.fill.solid()
    line.fill.fore_color.rgb = INDIGO_ACCENT
    line.line.color.rgb = INDIGO_ACCENT

def add_footer(slide, current_slide, total_slides=19):
    footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.35))
    tf = footer_box.text_frame
    p = tf.paragraphs[0]
    p.text = f"Rajalakshmi Engineering College  |  Dept. of Computer Science & Engineering  |  Project Review 2026  —  Slide {current_slide} of {total_slides}"
    p.font.name = "Calibri"
    p.font.size = Pt(9.5)
    p.font.color.rgb = TEXT_MUTED

def add_card(slide, left, top, width, height, title, items, bg_color=CARD_BG, border_color=CARD_BORDER, title_color=NAVY_PRIMARY):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.5)

    tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), height - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True

    p_title = tf.paragraphs[0]
    p_title.text = title
    p_title.font.name = "Georgia"
    p_title.font.size = Pt(15)
    p_title.font.bold = True
    p_title.font.color.rgb = title_color
    p_title.space_after = Pt(8)

    for item in items:
        p = tf.add_paragraph()
        p.text = item
        p.font.name = "Calibri"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(4)

# =============================================================
# SLIDE 1: TITLE SLIDE (Dark Navy Hero Theme)
# =============================================================
slide1 = prs.slides.add_slide(blank_layout)
bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
bg1.fill.solid()
bg1.fill.fore_color.rgb = NAVY_DARK
bg1.line.color.rgb = NAVY_DARK

# Header Badge
tb_b = slide1.shapes.add_textbox(Inches(1.0), Inches(0.7), Inches(11.333), Inches(0.4))
p_b = tb_b.text_frame.paragraphs[0]
p_b.text = "RAJALAKSHMI ENGINEERING COLLEGE (AUTONOMOUS) | DEPARTMENT OF CSE"
p_b.font.name = "Calibri"
p_b.font.size = Pt(12)
p_b.font.bold = True
p_b.font.color.rgb = RGBColor(129, 140, 248) # Indigo 400

# Sub-badge
tb_sub = slide1.shapes.add_textbox(Inches(1.0), Inches(1.1), Inches(11.333), Inches(0.4))
p_s = tb_sub.text_frame.paragraphs[0]
p_s.text = "FOUNDATIONS OF ARTIFICIAL INTELLIGENCE  •  PROJECT REVIEW PRESENTATION"
p_s.font.name = "Calibri"
p_s.font.size = Pt(13)
p_s.font.bold = True
p_s.font.color.rgb = RGBColor(52, 211, 153) # Emerald 400

# Main Project Title
tb_t = slide1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(2.0))
tf_t = tb_t.text_frame
tf_t.word_wrap = True
p_t = tf_t.paragraphs[0]
p_t.text = "TEAMFORGE: Intelligent Student Team Formation Platform"
p_t.font.name = "Georgia"
p_t.font.size = Pt(32)
p_t.font.bold = True
p_t.font.color.rgb = WHITE

p_t2 = tf_t.add_paragraph()
p_t2.text = "Using Constraint Satisfaction and Heuristic Local Search"
p_t2.font.name = "Georgia"
p_t2.font.size = Pt(22)
p_t2.font.color.rgb = RGBColor(224, 231, 255) # Indigo 100
p_t2.space_before = Pt(6)

# Divider
div1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(4.0), Inches(11.333), Inches(0.03))
div1.fill.solid()
div1.fill.fore_color.rgb = INDIGO_ACCENT
div1.line.color.rgb = INDIGO_ACCENT

# Details Box (Left: Students, Right: Supervisor & Date)
tb_stud = slide1.shapes.add_textbox(Inches(1.0), Inches(4.3), Inches(5.5), Inches(2.5))
tf_stud = tb_stud.text_frame
p_st_h = tf_stud.paragraphs[0]
p_st_h.text = "PROJECT CANDIDATES:"
p_st_h.font.name = "Calibri"
p_st_h.font.size = Pt(13)
p_st_h.font.bold = True
p_st_h.font.color.rgb = RGBColor(129, 140, 248)

p_st1 = tf_stud.add_paragraph()
p_st1.text = "• TARUN ANTONY M (Reg No: 2116240701556)"
p_st1.font.name = "Calibri"
p_st1.font.size = Pt(14)
p_st1.font.bold = True
p_st1.font.color.rgb = WHITE
p_st1.space_before = Pt(4)

p_st2 = tf_stud.add_paragraph()
p_st2.text = "• THANIKAIVEL S   (Reg No: 2116240701562)"
p_st2.font.name = "Calibri"
p_st2.font.size = Pt(14)
p_st2.font.bold = True
p_st2.font.color.rgb = WHITE
p_st2.space_before = Pt(4)

p_st3 = tf_stud.add_paragraph()
p_st3.text = "Degree: B.E. Computer Science and Engineering"
p_st3.font.name = "Calibri"
p_st3.font.size = Pt(12)
p_st3.font.color.rgb = RGBColor(148, 163, 184)
p_st3.space_before = Pt(6)

tb_sup = slide1.shapes.add_textbox(Inches(7.0), Inches(4.3), Inches(5.333), Inches(2.5))
tf_sup = tb_sup.text_frame
p_su_h = tf_sup.paragraphs[0]
p_su_h.text = "PROJECT SUPERVISOR & INSTITUTION:"
p_su_h.font.name = "Calibri"
p_su_h.font.size = Pt(13)
p_su_h.font.bold = True
p_su_h.font.color.rgb = RGBColor(129, 140, 248)

p_su1 = tf_sup.add_paragraph()
p_su1.text = "Mrs. B. Deepa, M.E., (Ph.D.)"
p_su1.font.name = "Calibri"
p_su1.font.size = Pt(14)
p_su1.font.bold = True
p_su1.font.color.rgb = WHITE
p_su1.space_before = Pt(4)

p_su2 = tf_sup.add_paragraph()
p_su2.text = "Assistant Professor (Selection Grade), Dept. of CSE\nRajalakshmi Engineering College, Thandalam, Chennai\nAcademic Year: 2026–2027  |  Review Date: November 2026"
p_su2.font.name = "Calibri"
p_su2.font.size = Pt(12)
p_su2.font.color.rgb = RGBColor(148, 163, 184)
p_su2.space_before = Pt(2)


# =============================================================
# SLIDE 2: MOTIVATION & CURRENT PAIN POINTS
# =============================================================
slide2 = prs.slides.add_slide(blank_layout)
add_header(slide2, "1. Background & Motivation: Why TeamForge?", "PROJECT CONTEXT")
add_footer(slide2, 2)

add_card(slide2, Inches(0.8), Inches(1.8), Inches(3.6), Inches(4.8), "High-Stakes Events", [
    "• Collegiate Hackathons (Smart India Hackathon, internal college hackathons) require multi-disciplinary skills within 24-36 hrs.",
    "• Final Year Capstone Projects dictate graduation outcomes and project quality.",
    "• Diverse skill sets needed: UI/UX, Backend, Database, AI/ML, Testing.",
    "• Poorly structured teams lead to project abandonment or failure."
])

add_card(slide2, Inches(4.8), Inches(1.8), Inches(3.6), Inches(4.8), "The Breakdown of Ad-Hoc Teams", [
    "• Social Clique Bias: Friends group together regardless of overlapping or missing technical competencies.",
    "• Severe Skill Gaps: Teams with 4 frontend designers and zero database engineers, or vice versa.",
    "• Schedule Bottlenecks: Uncoordinated commitments lead to inactive contributors.",
    "• Isolated Solitary Talent: Introverted students or juniors with niche skills are systematically excluded."
], title_color=RGBColor(185, 28, 28))

add_card(slide2, Inches(8.8), Inches(1.8), Inches(3.7), Inches(4.8), "The Classical AI Solution", [
    "• Replaces guesswork with mathematically optimal Constraint Satisfaction Problem (CSP).",
    "• Pure Classical AI: Zero external black-box models, zero paid AI APIs, zero hallucinations.",
    "• 100% Deterministic & Explainable audit trail for faculty and student leaders.",
    "• Integrated Collaborative Workspace: Kanban board, sessions, and zero-retention ephemeral chat."
], title_color=EMERALD_GREEN)


# =============================================================
# SLIDE 3: PROBLEM STATEMENT & COMBINATORIAL COMPLEXITY
# =============================================================
slide3 = prs.slides.add_slide(blank_layout)
add_header(slide3, "2. Problem Statement & Mathematical Formulation", "PROBLEM DEFINITION")
add_footer(slide3, 3)

add_card(slide3, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "The Core Problem Statement", [
    "Given a cohort of N registered college students possessing heterogeneous technical skill proficiencies, academic seniority levels, and availability schedules:",
    "",
    "• Objective: Construct an optimal partition of students into cohesive project teams of target size K.",
    "• Hard Feasibility: Every team must satisfy strict capacity limits, pairwise member uniqueness, and collegiate event registration.",
    "• Multi-Objective Balance: Must balance technical skill coverage, role specialization, senior-junior pairing, and sprint availability.",
    "• Explainability Requirement: System must provide an explainable decision log justifying candidate selection."
])

add_card(slide3, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8), "The Combinatorial Explosion Challenge", [
    "• The number of possible team combinations across N candidates for team size K follows the binomial coefficient:",
    "       C(N, K) = N! / [ K! * (N - K)! ]",
    "",
    "• For a typical collegiate cohort (N = 60 students, K = 4 members):",
    "       C(60, 4) = 487,635 possible subsets per team!",
    "• For K = 6 members:",
    "       C(60, 6) = 50,063,860 possible subsets!",
    "",
    "• Brute-Force Exhaustive Enumeration is NP-hard and computationally prohibitive for interactive web applications.",
    "• Requires Classical AI Search Heuristics (MRV + Forward Checking + Hill Climbing) to prune the search space to sub-millisecond execution."
])


# =============================================================
# SLIDE 4: PROJECT OBJECTIVES
# =============================================================
slide4 = prs.slides.add_slide(blank_layout)
add_header(slide4, "3. Key Objectives & Engineering Deliverables", "PROJECT OBJECTIVES")
add_footer(slide4, 4)

objs = [
    ("Objective 1: Classical CSP Modeling", ["Formalize student team formation as a discrete Constraint Satisfaction Problem (CSP) with explicit Variables, Domains, and Hard Constraints."]),
    ("Objective 2: Intelligent Search Space Pruning", ["Implement Backtracking Search augmented by Minimum Remaining Values (MRV) heuristic and Forward Checking (FC) to eliminate dead branches."]),
    ("Objective 3: Heuristic Multi-Objective Local Search", ["Develop a Hill Climbing local search procedure with 1-opt candidate swaps guided by a 5-factor multi-objective scoring function."]),
    ("Objective 4: Explainable Multi-Proposal Generator", ["Generate 3 distinct team proposals (Balanced Alpha, Skill-Dense Beta, Agile Gamma) with transparent step-by-step decision trace logs."]),
    ("Objective 5: Full-Stack Web Architecture & Privacy", ["Engineer a high-throughput, decoupled platform (PHP 8.2 + MySQL 8.0) featuring Kanban boards, standups, and zero-storage ephemeral chat."])
]

for idx, (title, items) in enumerate(objs):
    y = Inches(1.7 + (idx * 1.0))
    add_card(slide4, Inches(0.8), y, Inches(11.733), Inches(0.85), title, items)


# =============================================================
# SLIDE 5: LITERATURE SURVEY & COMPETITIVE MATRIX
# =============================================================
slide5 = prs.slides.add_slide(blank_layout)
add_header(slide5, "4. Literature Review & Research Gaps", "STATE OF THE ART")
add_footer(slide5, 5)

# Table for comparison
table_shape = slide5.shapes.add_table(5, 5, Inches(0.8), Inches(1.8), Inches(11.733), Inches(3.2))
tbl = table_shape.table

tbl.columns[0].width = Inches(2.5)
tbl.columns[1].width = Inches(2.3)
tbl.columns[2].width = Inches(2.3)
tbl.columns[3].width = Inches(2.3)
tbl.columns[4].width = Inches(2.333)

headers = ["Platform / Approach", "Matching Paradigm", "Constraint Guarantees", "AI Explainability", "Data Privacy & Chat"]
for i, h in enumerate(headers):
    cell = tbl.cell(0, i)
    cell.text = h
    cell.fill.solid()
    cell.fill.fore_color.rgb = NAVY_PRIMARY
    p = cell.text_frame.paragraphs[0]
    p.font.name = "Georgia"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

matrix_data = [
    ("Devpost / Hackathon Portals", "Manual bulletin board / Self-selection", "None (Manual checking)", "Zero AI", "External Discord/Slack links"),
    ("Commercial HR AI Tools", "Opaque Deep Learning / Embeddings", "Probabilistic / Weak", "Black-box (Unexplainable)", "Permanent Cloud Data Retention"),
    ("Academic CLI Scripts", "Stand-alone scripts (ILP / GA)", "Strict (Math formulations)", "High (Mathematical rules)", "No integrated communication"),
    ("TeamForge (Proposed)", "Hybrid CSP (MRV+FC) + Hill Climbing", "100% Deterministic & Hard", "Full Step-by-Step Trace Log", "Zero-Retention Ephemeral Buffer")
]

for row_idx, rdata in enumerate(matrix_data):
    for col_idx, text in enumerate(rdata):
        cell = tbl.cell(row_idx + 1, col_idx)
        cell.text = text
        cell.fill.solid()
        if row_idx == 3: # Highlight proposed TeamForge
            cell.fill.fore_color.rgb = RGBColor(238, 242, 255)
        else:
            cell.fill.fore_color.rgb = WHITE
        p = cell.text_frame.paragraphs[0]
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        if row_idx == 3:
            p.font.bold = True
            p.font.color.rgb = INDIGO_ACCENT
        else:
            p.font.color.rgb = TEXT_DARK
        p.alignment = PP_ALIGN.CENTER if col_idx > 0 else PP_ALIGN.LEFT

add_card(slide5, Inches(0.8), Inches(5.3), Inches(11.733), Inches(1.5), "Identified Research Gap Addressed by TeamForge", [
    "• Current academic tools lack end-to-end integration: they produce matching outputs in offline CSV files without project boards, session scheduling, or privacy-first communication.",
    "• TeamForge bridges classical AI rigor (CSP + Local Search) directly into a production-grade collegiate web platform."
], bg_color=RGBColor(240, 253, 244), border_color=RGBColor(187, 247, 208), title_color=EMERALD_GREEN)


# =============================================================
# SLIDE 6: PROPOSED SYSTEM ARCHITECTURE (Figure 3.1)
# =============================================================
slide6 = prs.slides.add_slide(blank_layout)
add_header(slide6, "5. Proposed System Architecture: Decoupled Multi-Tier", "SYSTEM DESIGN")
add_footer(slide6, 6)

# Embed Architecture Diagram Image
img_arch = r"c:\AI & Web project\assets\report_assets\fig_3_1_architecture.png"
if os.path.exists(img_arch):
    slide6.shapes.add_picture(img_arch, Inches(0.8), Inches(1.7), Inches(6.8), Inches(4.8))

add_card(slide6, Inches(7.8), Inches(1.7), Inches(4.7), Inches(4.8), "Architectural Highlights", [
    "1. Presentation Layer (Client):",
    "  • Standards-compliant HTML5, Vanilla CSS3 Design Tokens, Bootstrap 5.3, ES6+ JS.",
    "  • Communicates strictly via asynchronous JSON fetch APIs.",
    "",
    "2. Application & AI Engine Layer (Server):",
    "  • Pure PHP 8.2+ backend with zero external AI API costs.",
    "  • Role-Based Access Control (RBAC): Student, Team Leader, Faculty Admin.",
    "  • Deterministic CSP Engine + Hill Climbing Optimizer.",
    "",
    "3. Data Layer (Dual-Storage Strategy):",
    "  • Persistent State: MySQL 8.0 normalized 3NF schema.",
    "  • Ephemeral State: Rolling in-memory chat buffer with zero persistent database storage."
])


# =============================================================
# SLIDE 7: CLASSICAL AI FORMULATION (CSP ENGINE)
# =============================================================
slide7 = prs.slides.add_slide(blank_layout)
add_header(slide7, "6. Classical AI Formulation: The CSP Framework", "AI ENGINE SPECIFICATIONS")
add_footer(slide7, 7)

add_card(slide7, Inches(0.8), Inches(1.8), Inches(3.6), Inches(4.8), "Variables (X)", [
    "• Defined as the set of unfilled functional role slots within the target team:",
    "    X = { V_1, V_2, ..., V_k }",
    "",
    "• Where k is the required team size (e.g. k = 4):",
    "  - V_1: Team Leader / Systems Lead",
    "  - V_2: Frontend / UI Engineer",
    "  - V_3: Backend / Database Architect",
    "  - V_4: Classical AI & Algo Specialist"
])

add_card(slide7, Inches(4.8), Inches(1.8), Inches(3.6), Inches(4.8), "Domains (D)", [
    "• For each role variable V_i, the candidate domain D(V_i) is filtered dynamically:",
    "    D(V_i) subset of { Candidate Students }",
    "",
    "• Filtering Criteria:",
    "  1. Technical competence matching required role keywords.",
    "  2. Valid registration in the target collegiate hackathon.",
    "  3. Pinned member designation (locked by leader)."
])

add_card(slide7, Inches(8.8), Inches(1.8), Inches(3.7), Inches(4.8), "Hard Constraints (C)", [
    "Must be 100% satisfied for valid assignment:",
    "",
    "1. Team Capacity:",
    "   Exactly k members assigned.",
    "",
    "2. Pairwise Uniqueness (AllDiff):",
    "   forall i != j, V_i != V_j",
    "   (No student occupies 2 slots).",
    "",
    "3. Event Enrollment Constraint:",
    "   All candidates must be registered.",
    "",
    "4. Active Bandwidth Constraint:",
    "   Active team commitments < limit."
], title_color=RGBColor(185, 28, 28))


# =============================================================
# SLIDE 8: SEARCH HEURISTICS: MRV & FORWARD CHECKING
# =============================================================
slide8 = prs.slides.add_slide(blank_layout)
add_header(slide8, "7. Combinatorial Search Optimization: MRV & Forward Checking", "AI ENGINE SPECIFICATIONS")
add_footer(slide8, 8)

add_card(slide8, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Minimum Remaining Values (MRV) Heuristic", [
    "• Often called the 'Most Constrained Variable' or 'Fail-First' Heuristic.",
    "",
    "• Algorithmic Mechanism:",
    "  - At each step in recursive backtracking, inspect all unassigned role variables.",
    "  - Select variable V_chosen with minimum remaining candidates: |D(V_chosen)|.",
    "",
    "• Why It Works in Team Formation:",
    "  - If a team requires a rare skill (e.g. classical graph algorithm specialist with only 2 candidates), assigning it first immediately exposes whether a valid team is possible.",
    "  - Dramatically shrinks the branching factor at the top of the search tree."
], title_color=INDIGO_ACCENT)

add_card(slide8, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8), "Forward Checking (FC) Constraint Propagation", [
    "• Algorithmic Mechanism:",
    "  - When candidate C is tentatively assigned to role V_i:",
    "  - Immediately prune C from the domains of all other unassigned variables:",
    "        D(V_j) <- D(V_j) \\ { C }   forall j != i",
    "",
    "• Domain Wipeout Early Detection:",
    "  - If pruning causes any unassigned variable's domain to become empty (|D(V_j)| = 0):",
    "  - The search initiates an IMMEDIATE backtrack!",
    "",
    "• Prevents the algorithm from wasting cycles traversing hopeless branches that can never form a complete team."
], title_color=EMERALD_GREEN)


# =============================================================
# SLIDE 9: HEURISTIC LOCAL SEARCH (HILL CLIMBING)
# =============================================================
slide9 = prs.slides.add_slide(blank_layout)
add_header(slide9, "8. Heuristic Local Search & Multi-Objective Optimization", "HEURISTIC SEARCH")
add_footer(slide9, 9)

add_card(slide9, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Hill Climbing Local Search (1-Opt Swaps)", [
    "• Systematic search discovers a feasible team, but local search optimizes its qualitative synergy.",
    "",
    "• Iterative Procedure:",
    "  1. Start with initial feasible team assignment A_0 (from CSP / Greedy baseline).",
    "  2. Explore 1-Opt Neighborhood: Systematically evaluate swapping each unpinned team member with an unused candidate in the pool.",
    "  3. Retain any neighbor swap that strictly improves fitness: S(A_neighbor) > S(A_current) + delta.",
    "  4. Terminate when no 1-opt swap improves score (Local Optimum Plateau achieved).",
    "",
    "• Maximum iterations capped at 25 (empirically converges in 8-12 iterations)."
], title_color=INDIGO_ACCENT)

add_card(slide9, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8), "Multi-Objective Fitness Function S(Team)", [
    "• Composite evaluation across 5 critical dimensions:",
    "  S(Team) = W_skill * SkillCov + W_role * RoleCov +",
    "             W_exp * ExpBal + W_avail * AvailScore +",
    "             W_interest * InterestAlign",
    "",
    "• Normalized Weights (Sum = 1.0):",
    "  - W_skill = 0.35 (Mandatory project skill coverage)",
    "  - W_role  = 0.25 (Specialized role role fulfillment)",
    "  - W_exp   = 0.15 (Senior-junior pairing & project history)",
    "  - W_avail = 0.15 (Active weekly commitment & bandwidth)",
    "  - W_interest = 0.10 (Hackathon track interest alignment)",
    "",
    "• Produces a 0 to 100% normalized team harmony score."
], title_color=EMERALD_GREEN)


# =============================================================
# SLIDE 10: MULTI-PROPOSAL GENERATION STRATEGY
# =============================================================
slide10 = prs.slides.add_slide(blank_layout)
add_header(slide10, "9. Explainable Multi-Proposal Generation Strategy", "ALGORITHMIC OUTPUTS")
add_footer(slide10, 10)

add_card(slide10, Inches(0.8), Inches(1.8), Inches(3.6), Inches(4.8), "Proposal Alpha\n(Balanced Optimal)", [
    "• Holistic Harmony Focus",
    "• Algorithm: CSP (MRV+FC) followed by Hill Climbing local search.",
    "• Default Weightings:",
    "  Skill 35%, Role 25%, Exp 15%, Avail 15%, Interest 10%.",
    "",
    "• Empirical Score: 94.2%",
    "• Advantages: Strong multi-dimensional balance; zero single points of failure.",
    "• Ideal For: Standard collegiate hackathons & semester capstone projects."
], title_color=NAVY_PRIMARY)

add_card(slide10, Inches(4.8), Inches(1.8), Inches(3.6), Inches(4.8), "Proposal Beta\n(Skill-Dense Specialists)", [
    "• Deep Technical Specialization",
    "• Algorithm: Re-weighted local search emphasizing hard technical proficiencies.",
    "• Weightings:",
    "  Skill 55%, Role 20%, Exp 10%, Avail 10%, Interest 5%.",
    "",
    "• Empirical Score: 89.8%",
    "• Advantages: 100% skill coverage across complex frameworks & algorithms.",
    "• Ideal For: Advanced AI, blockchain, or competitive coding challenges."
], title_color=RGBColor(185, 28, 28))

add_card(slide10, Inches(8.8), Inches(1.8), Inches(3.7), Inches(4.8), "Proposal Gamma\n(Agile & High Bandwidth)", [
    "• Execution Velocity Focus",
    "• Algorithm: Greedy assignment prioritizing high available hours & open schedules.",
    "• Weightings:",
    "  Avail 35%, Skill 25%, Role 20%, Exp 15%, Interest 5%.",
    "",
    "• Empirical Score: 91.0%",
    "• Advantages: 100% member availability score; zero scheduling conflicts.",
    "• Ideal For: Rapid 24-hr sprint hackathons where communication speed is critical."
], title_color=EMERALD_GREEN)


# =============================================================
# SLIDE 11: PLATFORM CORE MODULES (1 to 6)
# =============================================================
slide11 = prs.slides.add_slide(blank_layout)
add_header(slide11, "10. Functional System Modules Overview", "SYSTEM IMPLEMENTATION")
add_footer(slide11, 11)

modules = [
    ("Module 1: Student & Skill Profiling", ["3NF relational schema storing verified academic data, 1-5 skill proficiency ratings, years of experience, and past verified capstone projects."]),
    ("Module 2: Collegiate Event Manager", ["Faculty creation of collegiate hackathons, symposiums, and project expos with timeline controls, min/max team constraints, and eligibility rules."]),
    ("Module 3 & 4: Classical AI Engine", ["Native PHP 8.2 implementation of CSP, MRV Heuristic, Forward Checking, Hill Climbing, and transparent decision trace audit logging."]),
    ("Module 5: Project Kanban Board", ["Agile task tracking with drag-and-drop transitions (Not Started, In Progress, Completed), member assignments, category badges, and deadlines."]),
    ("Module 6: Privacy Ephemeral Chat", ["Real-time rolling in-memory team chat with ZERO persistent database storage, ensuring strict student privacy and ephemeral communication."]),
    ("Module 7: Session Scheduler & RBAC", ["Standup meetings with structured agendas and attendance logs, reinforced by enterprise Role-Based Access Control and CSRF validation."])
]

for idx, (title, items) in enumerate(modules):
    col = idx % 2
    row = idx // 2
    x = Inches(0.8 + (col * 5.9))
    y = Inches(1.8 + (row * 1.6))
    add_card(slide11, x, y, Inches(5.7), Inches(1.45), title, items)


# =============================================================
# SLIDE 12: EXPERIMENTAL SETUP & EVALUATION ENVIRONMENT
# =============================================================
slide12 = prs.slides.add_slide(blank_layout)
add_header(slide12, "11. Experimental Setup & Benchmarking Environment", "EXPERIMENTAL EVALUATION")
add_footer(slide12, 12)

# Specs Table
t_setup = slide12.shapes.add_table(7, 2, Inches(0.8), Inches(1.8), Inches(6.0), Inches(4.8))
stbl = t_setup.table
stbl.columns[0].width = Inches(2.2)
stbl.columns[1].width = Inches(3.8)

specs = [
    ("Component", "Technical Specification"),
    ("Operating System", "Microsoft Windows 11 Enterprise (64-bit)"),
    ("Web Server & Runtime", "Apache 2.4.58 / PHP 8.2.12 (OPcache Enabled)"),
    ("Database Management", "MySQL 8.0.35 / MariaDB (InnoDB, utf8mb4)"),
    ("Client-Side Stack", "HTML5, Vanilla CSS3 tokens, Bootstrap 5.3, ES6+ JS"),
    ("Processor & RAM", "Intel Core i7-12700H @ 2.30 GHz (14 Cores), 16GB DDR5"),
    ("Pre-Seeded Test Cohort", "60 Students, 15 Master Skills, 4 Active Hackathons")
]

for row_idx, (k, v) in enumerate(specs):
    c0 = stbl.cell(row_idx, 0)
    c1 = stbl.cell(row_idx, 1)
    c0.text = k
    c1.text = v
    c0.fill.solid()
    c1.fill.solid()
    if row_idx == 0:
        c0.fill.fore_color.rgb = NAVY_PRIMARY
        c1.fill.fore_color.rgb = NAVY_PRIMARY
        c0.text_frame.paragraphs[0].font.bold = True
        c1.text_frame.paragraphs[0].font.bold = True
        c0.text_frame.paragraphs[0].font.color.rgb = WHITE
        c1.text_frame.paragraphs[0].font.color.rgb = WHITE
    else:
        c0.fill.fore_color.rgb = WHITE
        c1.fill.fore_color.rgb = WHITE
        c0.text_frame.paragraphs[0].font.bold = True
        c0.text_frame.paragraphs[0].font.color.rgb = NAVY_PRIMARY
        c1.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK
        c0.text_frame.paragraphs[0].font.size = Pt(11)
        c1.text_frame.paragraphs[0].font.size = Pt(11)

add_card(slide12, Inches(7.2), Inches(1.8), Inches(5.333), Inches(4.8), "Evaluation Methodology", [
    "• Benchmark Objective:",
    "  Quantify search space reduction, execution latency, and convergence speed compared to standard algorithms.",
    "",
    "• Key Performance Indicators (KPIs):",
    "  1. Search Tree Nodes Explored: Measure of algorithmic pruning efficiency.",
    "  2. Execution Latency (ms): Server-side solve time for real-time web responsiveness.",
    "  3. Local Search Iterations: Steps required for Hill Climbing convergence.",
    "  4. Multi-Objective Fitness Score (%): Quality of resulting student team synergy."
])


# =============================================================
# SLIDE 13: EMPIRICAL RESULTS: SEARCH PRUNING (Figure 5.3)
# =============================================================
slide13 = prs.slides.add_slide(blank_layout)
add_header(slide13, "12. Results: Search Tree Pruning Efficiency", "PERFORMANCE EVALUATION")
add_footer(slide13, 13)

# Embed Search Pruning Image
img_prune = r"c:\AI & Web project\assets\report_assets\fig_5_3_search_pruning.png"
if os.path.exists(img_prune):
    slide13.shapes.add_picture(img_prune, Inches(0.8), Inches(1.7), Inches(6.4), Inches(4.8))

add_card(slide13, Inches(7.5), Inches(1.7), Inches(5.0), Inches(4.8), "Empirical Pruning Analysis", [
    "• Massive Search Space Compression:",
    "  - For team size K = 6 in a cohort of 60 students:",
    "  - Pure Chronological Backtracking explored: 11,800 nodes.",
    "  - Backtracking + MRV + Forward Checking explored: only 75 nodes!",
    "",
    "• 98.4% Reduction in Search Complexity:",
    "  - Minimum Remaining Values heuristic selects the most constrained role first, cutting wide branches early.",
    "  - Forward checking identifies domain wipeouts before searching dead subtrees.",
    "",
    "• Execution Latency: Sub-45 milliseconds on standard Apache/PHP server, enabling seamless instant recommendations in the web UI."
], title_color=EMERALD_GREEN)


# =============================================================
# SLIDE 14: EMPIRICAL RESULTS: CONVERGENCE & RADAR (Figs 5.1 & 5.2)
# =============================================================
slide14 = prs.slides.add_slide(blank_layout)
add_header(slide14, "13. Results: Local Search Convergence & Proposal Trade-Offs", "PERFORMANCE EVALUATION")
add_footer(slide14, 14)

img_conv = r"c:\AI & Web project\assets\report_assets\fig_5_1_hill_climbing_convergence.png"
if os.path.exists(img_conv):
    slide14.shapes.add_picture(img_conv, Inches(0.8), Inches(1.7), Inches(5.6), Inches(3.6))

img_radar = r"c:\AI & Web project\assets\report_assets\fig_5_2_radar_proposals.png"
if os.path.exists(img_radar):
    slide14.shapes.add_picture(img_radar, Inches(6.8), Inches(1.7), Inches(5.7), Inches(3.6))

add_card(slide14, Inches(0.8), Inches(5.5), Inches(11.733), Inches(1.4), "Summary of Empirical Findings", [
    "• Convergence Dynamics (Left): Hill Climbing local search elevates the greedy baseline score from 68.4% to 90.2% in 5 iterations, smoothly plateauing at 95.5% by iteration 11.",
    "• Multi-Objective Balance (Right): Proposal Alpha achieves 94.2% overall balance, while Proposal Beta provides 100% skill coverage and Proposal Gamma achieves 100% student availability."
], bg_color=CARD_BG, border_color=INDIGO_ACCENT)


# =============================================================
# SLIDE 15: SECURITY & PRIVACY IMPLEMENTATIONS
# =============================================================
slide15 = prs.slides.add_slide(blank_layout)
add_header(slide15, "14. Privacy-Preserving Architecture & Web Security", "SECURITY & COMPLIANCE")
add_footer(slide15, 15)

add_card(slide15, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Zero-Retention Ephemeral Team Chat", [
    "• Critical Privacy Innovation:",
    "  - Unlike Slack, MS Teams, or Discord which permanently store messages in cloud databases:",
    "  - TeamForge chat messages are deliberately NOT stored in MySQL.",
    "",
    "• Rolling In-Memory Buffer:",
    "  - Messages reside strictly in an ephemeral rolling server buffer that auto-expires.",
    "  - Eliminates risks of long-term message surveillance and leaks.",
    "",
    "• Authorization Wall:",
    "  - Access strictly enforced via isUserMemberOf($teamId, $userId). Non-members cannot intercept messages."
], title_color=INDIGO_ACCENT)

add_card(slide15, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8), "Enterprise Web Security Measures", [
    "1. SQL Injection Prevention:",
    "   100% of queries use PDO Parameterized Statements with prepared execution.",
    "",
    "2. CSRF Mitigation:",
    "   Cryptographic anti-CSRF tokens validated on all state-changing POST endpoints.",
    "",
    "3. Session Hardening:",
    "   session_regenerate_id(true) executed on login to prevent session fixation.",
    "",
    "4. XSS Escaping:",
    "   All dynamic inputs escaped using htmlspecialchars(..., ENT_QUOTES, 'UTF-8').",
    "",
    "5. Password Security:",
    "   Bcrypt one-way cryptographic hashing (password_hash() & password_verify())."
], title_color=EMERALD_GREEN)


# =============================================================
# SLIDE 16: SYSTEM DEMONSTRATION & USER INTERFACE
# =============================================================
slide16 = prs.slides.add_slide(blank_layout)
add_header(slide16, "15. Live Platform Demonstration Highlights", "USER INTERFACE")
add_footer(slide16, 16)

ui_cards = [
    ("1. Role-Aware Student Dashboard", ["Personalized greeting, live team formation metrics, active event registrations, skill tags, and quick-access navigation."]),
    ("2. Student Directory & Search", ["Instant multi-criteria search filtering candidates by department, academic year, specific skill proficiency, and availability status."]),
    ("3. Classical AI Generator Console", ["Interactive solver interface where leaders select required skills, tune objective weights, and inspect live execution trace logs."]),
    ("4. Alternative Proposal Cards", ["Visual radar chart breakdown and comparative cards for Proposal Alpha (Balanced), Beta (Specialist), and Gamma (Agile)."]),
    ("5. Project Kanban Task Board", ["Full sprint management with drag-and-drop transitions across Not Started, In Progress, and Completed columns."]),
    ("6. Ephemeral Chat & Standups", ["Zero-retention real-time messaging channel paired with session scheduling for standups and mentor evaluations."])
]

for idx, (title, items) in enumerate(ui_cards):
    col = idx % 2
    row = idx // 2
    x = Inches(0.8 + (col * 5.9))
    y = Inches(1.8 + (row * 1.6))
    add_card(slide16, x, y, Inches(5.7), Inches(1.45), title, items)


# =============================================================
# SLIDE 17: CONCLUSION & FUTURE ENHANCEMENTS
# =============================================================
slide17 = prs.slides.add_slide(blank_layout)
add_header(slide17, "16. Conclusion & Future Enhancements", "CONCLUSION & ROADMAP")
add_footer(slide17, 17)

add_card(slide17, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), "Project Summary & Achievements", [
    "• Rigorous Classical AI: Proved that complex student team formation can be solved effectively using deterministic CSP and Hill Climbing without black-box ML models.",
    "• Computational Superiority: Achieved 98.4% search tree node pruning via MRV and forward checking, executing in sub-45ms.",
    "• Practical Student Impact: Replaces nepotism and social clique bias with balanced, high-synergy multidisciplinary project teams.",
    "• Production-Ready Platform: Complete full-stack web ecosystem with normalized MySQL schema, Kanban boards, and privacy-first ephemeral messaging."
], title_color=NAVY_PRIMARY)

add_card(slide17, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8), "Future Research & Roadmap", [
    "1. Stochastic Metaheuristics:",
    "   Integrate Simulated Annealing with geometric cooling to escape local optima in extremely sparse candidate pools.",
    "",
    "2. Learning Management System (LMS) Integration:",
    "   Connect with Moodle, Canvas, and Blackboard via LTI standards for automated course roster sync.",
    "",
    "3. Cognitive Synergy & Belbin Roles:",
    "   Incorporate optional personality synergy surveys (Belbin Team Roles) into the multi-objective fitness function.",
    "",
    "4. Real-Time WebRTC Media:",
    "   Deploy ephemeral peer-to-peer audio/video standup meetings via WebSockets."
], title_color=EMERALD_GREEN)


# =============================================================
# SLIDE 18: KEY REFERENCES
# =============================================================
slide18 = prs.slides.add_slide(blank_layout)
add_header(slide18, "17. Key Academic References", "BIBLIOGRAPHY")
add_footer(slide18, 18)

refs = [
    ("Russell, S. and Norvig, P. (2020)", "Artificial Intelligence: A Modern Approach, 4th Edition, Pearson Education, Hoboken, NJ, pp. 180-213. [Core text for CSP, MRV, and Forward Checking]"),
    ("Dechter, R. (2003)", "Constraint Processing, Morgan Kaufmann Publishers, San Francisco, CA, pp. 73-124. [Theoretical foundations of constraint propagation and backtracking]"),
    ("Lappas, T., Liu, K. and Terzi, E. (2009)", "'Finding a Team of Experts in Social Networks', Proceedings of the 15th ACM SIGKDD International Conference on Knowledge Discovery and Data Mining (KDD '09), pp. 467-476. [NP-hardness proof of team formation]"),
    ("Aarts, E. and Korst, J. (1989)", "Simulated Annealing and Boltzmann Machines: A Stochastic Approach to Combinatorial Optimization, John Wiley & Sons, New York, pp. 45-88. [Local search principles]"),
    ("Li, C. T., Shan, M. K. and Lin, S. D. (2015)", "'On Team Formation with Skill Competence and Collaboration History in Social Networks', IEEE TKDE, Vol. 27, No. 8, pp. 2201-2213. [Skill competence balance modeling]")
]

tb_ref = slide18.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.733), Inches(4.8))
tf_ref = tb_ref.text_frame
tf_ref.word_wrap = True

for idx, (auth, details) in enumerate(refs):
    p = tf_ref.paragraphs[0] if idx == 0 else tf_ref.add_paragraph()
    r1 = p.add_run()
    r1.text = f"[{idx+1}] {auth}: "
    r1.font.name = "Georgia"
    r1.font.size = Pt(12)
    r1.font.bold = True
    r1.font.color.rgb = NAVY_PRIMARY

    r2 = p.add_run()
    r2.text = details
    r2.font.name = "Calibri"
    r2.font.size = Pt(12)
    r2.font.color.rgb = TEXT_DARK
    p.space_after = Pt(12)


# =============================================================
# SLIDE 19: THANK YOU & Q&A DEFENSE SLIDE (Dark Navy Hero Theme)
# =============================================================
slide19 = prs.slides.add_slide(blank_layout)
bg19 = slide19.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
bg19.fill.solid()
bg19.fill.fore_color.rgb = NAVY_DARK
bg19.line.color.rgb = NAVY_DARK

tb_ty = slide19.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.333), Inches(1.5))
p_ty = tb_ty.text_frame.paragraphs[0]
p_ty.text = "THANK YOU"
p_ty.font.name = "Georgia"
p_ty.font.size = Pt(44)
p_ty.font.bold = True
p_ty.font.color.rgb = WHITE
p_ty.alignment = PP_ALIGN.CENTER

tb_qa = slide19.shapes.add_textbox(Inches(1.0), Inches(2.7), Inches(11.333), Inches(1.0))
p_qa = tb_qa.text_frame.paragraphs[0]
p_qa.text = "Questions & Project Defense Discussion"
p_qa.font.name = "Georgia"
p_qa.font.size = Pt(24)
p_qa.font.color.rgb = RGBColor(129, 140, 248) # Indigo 400
p_qa.alignment = PP_ALIGN.CENTER

div19 = slide19.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.0), Inches(3.6), Inches(5.333), Inches(0.03))
div19.fill.solid()
div19.fill.fore_color.rgb = INDIGO_ACCENT
div19.line.color.rgb = INDIGO_ACCENT

tb_ack = slide19.shapes.add_textbox(Inches(1.0), Inches(4.0), Inches(11.333), Inches(2.8))
tf_ack = tb_ack.text_frame
p_ack1 = tf_ack.paragraphs[0]
p_ack1.text = "PROJECT TEAM: Tarun Antony M (2116240701556)  •  Thanikaivel S (2116240701562)"
p_ack1.font.name = "Calibri"
p_ack1.font.size = Pt(14)
p_ack1.font.bold = True
p_ack1.font.color.rgb = WHITE
p_ack1.alignment = PP_ALIGN.CENTER

p_ack2 = tf_ack.add_paragraph()
p_ack2.text = "SUPERVISOR: Mrs. B. Deepa, Assistant Professor (SG), Dept. of CSE"
p_ack2.font.name = "Calibri"
p_ack2.font.size = Pt(14)
p_ack2.font.bold = True
p_ack2.font.color.rgb = RGBColor(52, 211, 153) # Emerald 400
p_ack2.alignment = PP_ALIGN.CENTER
p_ack2.space_before = Pt(8)

p_ack3 = tf_ack.add_paragraph()
p_ack3.text = "Rajalakshmi Engineering College (Autonomous), Thandalam, Chennai\nFoundations of Artificial Intelligence  |  November 2026"
p_ack3.font.name = "Calibri"
p_ack3.font.size = Pt(12)
p_ack3.font.color.rgb = RGBColor(148, 163, 184)
p_ack3.alignment = PP_ALIGN.CENTER
p_ack3.space_before = Pt(8)

out_pptx = r"c:\AI & Web project\TeamForge_Project_Review.pptx"
prs.save(out_pptx)
print("Project Review PPTX successfully compiled and saved at:", out_pptx)
