"""
Generate a professional Word resume (.docx) tailored for PhD Finance application at FAU.
Run: python3 generate_resume.py
Output: Kateryna_Tsekhmayster_Resume.docx
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── Color constants ──────────────────────────────────────────────────────────
NAVY   = RGBColor(0x1B, 0x2A, 0x4A)   # section headings
GOLD   = RGBColor(0xB8, 0x86, 0x0B)   # accent / rule
BLACK  = RGBColor(0x1A, 0x1A, 0x1A)   # body text
GRAY   = RGBColor(0x55, 0x55, 0x55)   # secondary text

FONT_NAME = "Calibri"


# ── Helpers ──────────────────────────────────────────────────────────────────

def set_font(run, size_pt, bold=False, italic=False, color=BLACK, name=FONT_NAME):
    run.font.name       = name
    run.font.size       = Pt(size_pt)
    run.font.bold       = bold
    run.font.italic     = italic
    run.font.color.rgb  = color


def add_paragraph(doc, text="", style="Normal", space_before=0, space_after=3, keep_together=False):
    p = doc.add_paragraph(style=style)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    if keep_together:
        p.paragraph_format.keep_together = True
    if text:
        p.add_run(text)
    return p


def add_bottom_border(paragraph, color_hex="1B2A4A", size=6):
    """Add a bottom border line under a paragraph."""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    str(size))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color_hex)
    pBdr.append(bottom)
    pPr.append(pBdr)


def section_heading(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(2)
    r = p.add_run(title.upper())
    set_font(r, 11, bold=True, color=NAVY)
    add_bottom_border(p, "1B2A4A", size=8)
    return p


def bullet_run(doc, label, value="", label_bold=True, value_italic=False, size=10.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.left_indent  = Inches(0.2)
    rl = p.add_run(label)
    set_font(rl, size, bold=label_bold, color=BLACK)
    if value:
        rv = p.add_run(value)
        set_font(rv, size, italic=value_italic, color=BLACK)
    return p


def indented_para(doc, text, size=10, color=GRAY, left=0.2, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent  = Inches(left)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(space_after)
    r = p.add_run(text)
    set_font(r, size, color=color)
    return p


def set_margins(doc, top=1, bottom=1, left=1, right=1):
    for section in doc.sections:
        section.top_margin    = Inches(top)
        section.bottom_margin = Inches(bottom)
        section.left_margin   = Inches(left)
        section.right_margin  = Inches(right)


# ── Document builder ─────────────────────────────────────────────────────────

def build_resume():
    doc = Document()
    set_margins(doc, top=0.75, bottom=0.75, left=1.0, right=1.0)

    # ── NAME ──────────────────────────────────────────────────────────────────
    name_p = doc.add_paragraph()
    name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_p.paragraph_format.space_before = Pt(0)
    name_p.paragraph_format.space_after  = Pt(2)
    name_r = name_p.add_run("Kateryna Tsekhmayster")
    set_font(name_r, 24, bold=True, color=NAVY)

    # ── TITLE ─────────────────────────────────────────────────────────────────
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after  = Pt(2)
    title_r = title_p.add_run(
        "Dual Master's Candidate in Economics & Political Science  |  Quantum & Emerging Technology Research"
    )
    set_font(title_r, 10.5, italic=True, color=GRAY)

    # ── CONTACT ───────────────────────────────────────────────────────────────
    contact_p = doc.add_paragraph()
    contact_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact_p.paragraph_format.space_before = Pt(0)
    contact_p.paragraph_format.space_after  = Pt(6)
    contact_info = (
        "ktsekhmayste2022@fau.edu  ·  (561) 379-9646  ·  "
        "linkedin.com/in/kateryna-tsekhmayster  ·  ssrn.com/author=5467415"
    )
    contact_r = contact_p.add_run(contact_info)
    set_font(contact_r, 9.5, color=GRAY)

    # ── RESEARCH OBJECTIVE ────────────────────────────────────────────────────
    section_heading(doc, "Research Objective")
    obj_p = doc.add_paragraph()
    obj_p.paragraph_format.space_before = Pt(3)
    obj_p.paragraph_format.space_after  = Pt(4)
    obj_r = obj_p.add_run(
        "Seeking admission to the PhD in Finance program at Florida Atlantic University to advance research at the "
        "intersection of quantum computing and financial economics. My work develops quantitative frameworks—"
        "including the Quantum Labor Market Readiness Index (QLRI)—that apply quantum algorithms and "
        "computational modeling to emerging-market financial systems, workforce development, and geopolitical "
        "risk. I aim to bridge theoretical economics with cutting-edge quantum and AI methods to address complex "
        "problems in international finance and development."
    )
    set_font(obj_r, 10.5, color=BLACK)

    # ── EDUCATION ─────────────────────────────────────────────────────────────
    section_heading(doc, "Education")

    education = [
        {
            "degree":  "Master of Science, Economics  |  Master of Arts, Political Science",
            "school":  "Florida Atlantic University, Boca Raton, FL",
            "dates":   "In Progress — Expected May 2026",
            "details": [],
        },
        {
            "degree":  "Bachelor of Arts, Political Science and International Relations",
            "school":  "Wilkes Honors College, Florida Atlantic University",
            "dates":   "May 2024",
            "details": [
                "Magna Cum Laude",
                "Outstanding Thesis Award",
                "Undergraduate Research Certificate",
            ],
        },
        {
            "degree":  "Associate of Arts, Political Science and Government",
            "school":  "",
            "dates":   "May 2022",
            "details": [],
        },
    ]

    for edu in education:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after  = Pt(1)
        deg_r = p.add_run(edu["degree"])
        set_font(deg_r, 10.5, bold=True, color=BLACK)
        if edu["school"]:
            p.add_run("  ·  ")
            school_r = p.add_run(edu["school"])
            set_font(school_r, 10.5, color=GRAY)

        dates_p = doc.add_paragraph()
        dates_p.paragraph_format.space_before = Pt(0)
        dates_p.paragraph_format.space_after  = Pt(1)
        dates_r = dates_p.add_run(edu["dates"])
        set_font(dates_r, 10, italic=True, color=GRAY)

        for detail in edu["details"]:
            dp = doc.add_paragraph(style="List Bullet")
            dp.paragraph_format.space_before = Pt(0)
            dp.paragraph_format.space_after  = Pt(0)
            dp.paragraph_format.left_indent  = Inches(0.2)
            dr = dp.add_run(detail)
            set_font(dr, 10, color=BLACK)

    # ── RESEARCH & PUBLICATIONS ───────────────────────────────────────────────
    section_heading(doc, "Research & Publications")

    publications = [
        {
            "title":    "Quantum Labor Market Development in Emerging Countries",
            "date":     "December 2025",
            "note":     "Presented at FAU Graduate Research Day",
            "abstract": (
                "Explores the growing gap between quantum technology progress and workforce readiness. "
                "Introduces the Quantum Labor Market Readiness Index (QLRI)—a composite model assessing "
                "nations’ capabilities across digital infrastructure, human capital, innovation, and "
                "governance—to guide policymakers and industries in expanding quantum education and "
                "workforce development, especially within emerging economies."
            ),
            "keywords": "Quantum Computing, Labor Market, Development Economics, Economic Growth, Market Development",
        },
        {
            "title":    "Freedom as Development: Economics of the Belarusian Brain Drain and Remittances",
            "date":     "June 2025",
            "ssrn":     "DOI: 10.2139/ssrn.5467415",
            "abstract": (
                "Explores the case of Belarus, where governmental influence of Russia drives economic decisions "
                "due to dependency and lack of freedom, ultimately discouraging development. Applies Sen’s "
                "‘development as freedom’ framework to analyze skilled emigration and remittance flows."
            ),
            "keywords": "Economic Growth, Development, Brain Drain, Migration, Belarus, Soviet Union",
        },
        {
            "title":    "Brazilian Oil, Petrobras, and US-China Rivalry",
            "date":     "December 2025",
            "abstract": (
                "Examines how Petrobras’ strategic position influences Brazil’s political and economic "
                "alignment amid intensifying great-power competition between the United States and China in "
                "Latin American energy markets."
            ),
            "keywords": "Oil, Energy, Brazil, United States, China, Commodities, Geopolitics",
        },
        {
            "title":    "Blockchain of International Relations: Universal Cosmopolitan Protocol",
            "date":     "May 2025",
            "ssrn":     "DOI: 10.2139/ssrn.5272407",
            "abstract": (
                "Presents a blockchain architecture for conflict management in international relations. Each "
                "block contains information about nation-state actions and global responses, creating an "
                "immutable record for the United Nations to analyze adversary motives and streamline "
                "peace-establishment processes via smart contracts based on deontological frameworks."
            ),
            "keywords": "Blockchain, Smart Contracts, United Nations, International Relations, Perpetual Peace",
        },
        {
            "title":    "The Georgian Turn to Russia: A Warning for Ukraine",
            "date":     "April 2025",
            "abstract": (
                "Analyzes historical, philosophical, and social contexts of the Russo-Georgian War and the "
                "Ukrainian War to identify patterns of Russian power projection that culminate in military "
                "conflict. The Georgian democratic movement’s sharp turn toward Russia serves as a "
                "cautionary case for Ukraine’s Euro-Atlantic integration trajectory."
            ),
            "keywords": "Georgian War, Ukrainian War, Russian Influence, NATO, EU Membership",
        },
        {
            "title":    "The Tactical Formulation of the U.S. Constitution to an Oligarchic End",
            "date":     "May 2025",
            "abstract": (
                "Analyzes the US Constitution from the perspectives of Aristotle, Hobbes, and Locke, arguing "
                "that the Constitution’s emphasis on property protection created pathways for oligarchical "
                "influence through Supreme Court interpretations favoring corporate interests."
            ),
            "keywords": "Oligarchy, US Constitution, Political Philosophy, Constitutional Law",
        },
    ]

    for pub in publications:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after  = Pt(1)
        title_r = p.add_run(pub["title"])
        set_font(title_r, 10.5, bold=True, color=BLACK)

        meta_parts = [pub["date"]]
        if pub.get("ssrn"):
            meta_parts.append(pub["ssrn"])
        if pub.get("note"):
            meta_parts.append(pub["note"])
        meta_p = doc.add_paragraph()
        meta_p.paragraph_format.space_before = Pt(0)
        meta_p.paragraph_format.space_after  = Pt(1)
        meta_r = meta_p.add_run("  |  ".join(meta_parts))
        set_font(meta_r, 10, italic=True, color=GOLD if pub.get("note") else GRAY)

        abs_p = doc.add_paragraph()
        abs_p.paragraph_format.space_before = Pt(0)
        abs_p.paragraph_format.space_after  = Pt(1)
        abs_p.paragraph_format.left_indent  = Inches(0.15)
        abs_r = abs_p.add_run(pub["abstract"])
        set_font(abs_r, 10, color=GRAY)

        kw_p = doc.add_paragraph()
        kw_p.paragraph_format.space_before = Pt(0)
        kw_p.paragraph_format.space_after  = Pt(2)
        kw_p.paragraph_format.left_indent  = Inches(0.15)
        kw_label = kw_p.add_run("Keywords: ")
        set_font(kw_label, 9.5, bold=True, color=NAVY)
        kw_val = kw_p.add_run(pub["keywords"])
        set_font(kw_val, 9.5, italic=True, color=GRAY)

    # ── PROFESSIONAL EXPERIENCE ───────────────────────────────────────────────
    section_heading(doc, "Professional Experience")

    experience = [
        {
            "role":  "Public Relations Manager",
            "dates": "Feb 2024 – Present",
            "desc":  (
                "Direct marketing strategy development and oversee multi-platform campaigns, ensuring "
                "alignment with organizational objectives and consistent brand messaging across channels."
            ),
        },
        {
            "role":  "Communication Marketing Manager",
            "dates": "Oct 2023 – Feb 2024",
            "desc":  (
                "Developed integrated marketing strategies and managed campaigns across digital and "
                "traditional platforms, coordinating cross-functional teams to drive engagement."
            ),
        },
        {
            "role":  "Mayoral Campaign Assistant",
            "dates": "Jan 2023 – Apr 2023",
            "desc":  (
                "Coordinated research, event organization, and community outreach for a successful "
                "mayoral campaign; produced policy briefings and stakeholder communications."
            ),
        },
        {
            "role":  "Assistant International Commercial Communications Specialist",
            "dates": "Jun 2022 – Mar 2023",
            "desc":  (
                "Assisted in developing international communication strategies aligned with global "
                "branding; supported the creation of multilingual marketing materials and market-entry analyses."
            ),
        },
    ]

    for exp in experience:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after  = Pt(1)
        role_r = p.add_run(exp["role"])
        set_font(role_r, 10.5, bold=True, color=BLACK)
        sep_r = p.add_run("    ")
        set_font(sep_r, 10.5)
        dates_r = p.add_run(exp["dates"])
        set_font(dates_r, 10, italic=True, color=GRAY)

        desc_p = doc.add_paragraph()
        desc_p.paragraph_format.space_before = Pt(0)
        desc_p.paragraph_format.space_after  = Pt(2)
        desc_p.paragraph_format.left_indent  = Inches(0.15)
        desc_r = desc_p.add_run(exp["desc"])
        set_font(desc_r, 10.5, color=BLACK)

    # ── TECHNICAL SKILLS ─────────────────────────────────────────────────────
    section_heading(doc, "Technical Skills")

    skills = [
        "Statistical Analysis & Economic Data Analysis",
        "Quantum Computing Concepts & QLRI Modeling",
        "Blockchain Architecture & Smart Contracts",
        "Political Science & International Relations Research Methods",
        "Project Management & Event Coordination",
        "Digital Marketing & Public Relations",
        "Electronic Database Research",
        "Legislation Drafting & Policy Analysis",
        "Natural Language Processing (applied)",
        "Historical & Comparative Analysis",
    ]

    for skill in skills:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        p.paragraph_format.left_indent  = Inches(0.2)
        r = p.add_run(skill)
        set_font(r, 10.5, color=BLACK)

    # ── LANGUAGES ─────────────────────────────────────────────────────────────
    section_heading(doc, "Languages")

    languages = [
        ("English",   "Native"),
        ("Ukrainian", "Fluent"),
        ("Russian",   "Fluent"),
        ("French",    "Intermediate"),
    ]

    lang_p = doc.add_paragraph()
    lang_p.paragraph_format.space_before = Pt(3)
    lang_p.paragraph_format.space_after  = Pt(4)
    for i, (lang, level) in enumerate(languages):
        sep = "     " if i > 0 else ""
        lang_r = lang_p.add_run(f"{sep}{lang}: ")
        set_font(lang_r, 10.5, bold=True, color=NAVY)
        level_r = lang_p.add_run(level)
        set_font(level_r, 10.5, color=BLACK)

    # ── HONORS & AWARDS ───────────────────────────────────────────────────────
    section_heading(doc, "Honors & Awards")

    awards = [
        "Outstanding Thesis Award — Wilkes Honors College, FAU (May 2024)",
        "Magna Cum Laude — Wilkes Honors College, FAU (May 2024)",
        "Undergraduate Research Certificate — Wilkes Honors College, FAU (May 2024)",
        "Phi Sigma Alpha Honor Society — FAU (January 2025)",
        "Model UN Best Delegate Awards (2022, 2023, 2024, 2025)",
        "National Honor Society Member (2018–2020)",
        "French Honor Society Officer (2018–2020)",
    ]

    for award in awards:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        p.paragraph_format.left_indent  = Inches(0.2)
        r = p.add_run(award)
        set_font(r, 10.5, color=BLACK)

    # ── SAVE ──────────────────────────────────────────────────────────────────
    output_path = "Kateryna_Tsekhmayster_Resume.docx"
    doc.save(output_path)
    print(f"Resume saved to: {output_path}")


if __name__ == "__main__":
    build_resume()
