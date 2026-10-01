"""Editable Word copy of the one-page Data Engineer resume."""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt, RGBColor, Twips

OUT = Path(__file__).resolve().parent / "Bannusha_Shaik_Data_Engineer.docx"
INK = RGBColor(0x11, 0x11, 0x11)
CONTENT_W = Inches(7.51)


def set_run_font(run, size, bold=False, italic=False):
    run.bold = bold
    run.italic = italic
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.color.rgb = INK
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), "Calibri")


def tighten(paragraph, before=0, after=0, exact=None):
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY if exact else WD_LINE_SPACING.SINGLE
    if exact:
        fmt.line_spacing = Pt(exact)


def add_text(paragraph, text, size=10, bold=False, italic=False):
    run = paragraph.add_run(text)
    set_run_font(run, size, bold, italic)
    return run


def add_link(paragraph, text, url, size=10, bold=False, italic=False):
    part = paragraph.part
    rel_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)
    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), "Calibri")
    rPr.append(rFonts)
    if bold:
        b = OxmlElement("w:b")
        rPr.append(b)
    if italic:
        i = OxmlElement("w:i")
        rPr.append(i)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size * 2)))
    rPr.append(szCs)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "111111")
    rPr.append(color)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    run.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def underline(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "222222")
    pBdr.append(bottom)
    pPr.append(pBdr)


def zero_cell_margins(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for edge in ("top", "left", "bottom", "right"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), "0")
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def hide_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "nil")
        borders.append(node)
    tblPr.append(borders)


def pair(doc, left_parts, right_text, before=3):
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    hide_borders(table)
    table.columns[0].width = Inches(5.15)
    table.columns[1].width = Inches(2.36)
    left, right = table.rows[0].cells
    left.width = Inches(5.15)
    right.width = Inches(2.36)
    zero_cell_margins(left)
    zero_cell_margins(right)
    lp = left.paragraphs[0]
    tighten(lp, before=before, after=0, exact=13)
    for kind, value, extra in left_parts:
        if kind == "text":
            add_text(lp, value, 10, bold=True)
        elif kind == "link":
            add_link(lp, value, extra, size=10, bold=True)
    rp = right.paragraphs[0]
    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tighten(rp, before=before, after=0, exact=13)
    if right_text:
        add_text(rp, right_text, 10)
    return table


def tools(doc, text):
    p = doc.add_paragraph()
    tighten(p, before=0, after=0, exact=12)
    add_text(p, text, 10, italic=True)
    return p


def bullet(doc, text):
    p = doc.add_paragraph()
    tighten(p, before=0, after=0, exact=12.4)
    p.paragraph_format.left_indent = Inches(0.16)
    p.paragraph_format.first_line_indent = Inches(-0.12)
    add_text(p, "•  " + text, 10)
    return p


def section(doc, title):
    p = doc.add_paragraph()
    tighten(p, before=7, after=1, exact=14)
    add_text(p, title, 12, bold=True)
    underline(p)
    return p


def main():
    doc = Document()
    section_ = doc.sections[0]
    section_.page_width = Emu(7560310)   # A4
    section_.page_height = Emu(10692130)
    section_.left_margin = Inches(0.38)
    section_.right_margin = Inches(0.38)
    section_.top_margin = Inches(0.28)
    section_.bottom_margin = Inches(0.26)
    section_.header_distance = Inches(0.1)
    section_.footer_distance = Inches(0.1)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)
    normal.font.color.rgb = INK

    name = doc.add_paragraph()
    name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tighten(name, before=0, after=0, exact=28)
    add_text(name, "BANNUSHA SHAIK", 25, bold=True)

    role = doc.add_paragraph()
    role.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tighten(role, before=1, after=0, exact=14)
    add_text(role, "Data Engineer  |  Python · SQL · ETL · Data Modeling · PostgreSQL", 12, bold=True)

    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tighten(contact, before=1, after=0, exact=12)
    add_text(contact, "Bangalore, India  ·  +91 8074671779  ·  ", 10)
    add_link(contact, "bannushashaik85@gmail.com", "mailto:bannushashaik85@gmail.com", 10)
    add_text(contact, "  ·  ", 10)
    add_link(contact, "linkedin.com/in/bannushashaik400", "https://linkedin.com/in/bannushashaik400", 10)
    add_text(contact, "  ·  ", 10)
    add_link(contact, "bannusha.com", "https://bannusha.com", 10)

    section(doc, "SUMMARY")
    summary = doc.add_paragraph()
    summary.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    tighten(summary, before=2, after=0, exact=12.4)
    add_text(
        summary,
        "Data Engineer for operational models, pipelines, and the tables a team can query. "
        "Fortune 500 intern at SKF, Bangalore: paper 5S audits are now an 11-screen Power Apps "
        "and SharePoint system factory teams use, with Power BI score-cards that flag a result "
        "against a qualifying bar. Independently shipped a snapshot pipeline of 100 creators and "
        "about 3,500 videos, and an English-to-SQL service locked to the pasted schema. A city "
        "warehouse is in progress.",
        10,
    )

    section(doc, "PROFESSIONAL EXPERIENCE")
    pair(doc, [("text", "Data Engineer Intern, SKF (Fortune 500 Industrial)", None)], "May 2026 – Present | Bangalore")
    tools(doc, "SQL · data modeling · ETL · data pipelines · schema design · data quality · Power BI")
    bullet(doc, "Replaced paper 5S plant audits with an 11-screen Power Apps and SharePoint system. Factory teams use it for scores, photos, and actions, not files in email.")
    bullet(doc, "Set the grain as one audit across five pillars. Write-back is identity-aware and zone-scoped, so a submitted audit is a row plant leads can review.")
    bullet(doc, "Built the score-card and Power BI views: per-pillar totals, pass/fail against the qualifying bar, and the weakest pillar flagged for leadership.")

    pair(doc, [("text", "GenAI Data Analytics Job Simulation, Tata Group (Forage)", None)], "2024")
    bullet(doc, "Worked delinquency risk on structured financial datasets and wrote the next action a stakeholder could take, tied to the fields in the data. A job simulation, not employment.")

    pair(doc, [("text", "Technical Support, IEEE Student Branch, PES College of Engineering", None)], "Aug 2023 – Sep 2024")
    bullet(doc, "College club. Setup and troubleshooting for technical sessions so the room was ready before the talk started.")

    section(doc, "PROJECTS")
    pair(doc, [
        ("link", "Creator Lab", "https://know-stats.streamlit.app"),
        ("text", "  |  snapshot pipeline and scoring table", None),
    ], "")
    tools(doc, "Python · SQLite · YouTube Data API · pandas · scikit-learn · Streamlit · 100 creators · about 3,500 videos")
    bullet(doc, "Pulled the YouTube Data API into SQLite. Grain is one snapshot of one video, across 100 Indian creators (50 tech, 50 fashion) and about 3,500 videos.")
    bullet(doc, "Engagement is (likes + comments) / views. Outlier is views divided by that creator's own median, so a small channel is not ranked against a large one.")
    bullet(doc, "Streamlit scores a draft and returns a view range plus the nearest videos already stored. pandas and scikit-learn sit on the feature table.")

    pair(doc, [
        ("link", "Schema-Locked SQL", "https://github.com/bannushaxddd/GETYOQUERY"),
        ("text", "  |  English to SQL across 8 dialects", None),
    ], "")
    tools(doc, "Node.js · PostgreSQL · parameterized execution · JWT · 8 SQL dialects")
    bullet(doc, "English to SQL for PostgreSQL, MySQL, SQLite, SQL Server, BigQuery, Snowflake, Oracle, and DuckDB. A pasted CREATE TABLE is the allowlist, so generation cannot invent a column.")
    bullet(doc, "Query history is stored in PostgreSQL. Execution is parameterized, and the API sits behind JWT, so a generated statement cannot invent a column or concatenate input into SQL.")

    pair(doc, [("text", "Bengaluru Urban Warehouse  |  city feeds into one model", None)], "In progress")
    tools(doc, "Python · PostgreSQL · PostGIS · weather, air, and mobility")
    bullet(doc, "Not shipped. Three public feeds — weather, air, and mobility — that do not share a schema, a clock, or a location. Each raw payload is kept.")
    bullet(doc, "Invalid rows go to a reject table with the rule that failed. Ward, station, and date are dimensions in PostgreSQL and PostGIS. An unknown location is surrogate key −1, not a dropped row.")

    section(doc, "TECHNICAL SKILLS")
    for label, rest in (
        ("Data engineering: ", "ETL, data pipelines, data modeling, API ingestion, schema design, snapshot tables, data quality"),
        ("SQL & databases: ", "SQL, PostgreSQL, MySQL, SQLite, joins, window functions, parameterized queries"),
        ("Python: ", "pandas, NumPy, API clients, scikit-learn"),
        ("Serve & plant: ", "Power BI, Streamlit, Grafana, Docker, Git, Power Apps, SharePoint, Microsoft 365"),
    ):
        p = doc.add_paragraph()
        tighten(p, before=0, after=0, exact=12.4)
        add_text(p, label, 10, bold=True)
        add_text(p, rest, 10)

    section(doc, "EDUCATION & CERTIFICATIONS")
    pair(doc, [("text", "B.Tech, Artificial Intelligence & Machine Learning, PES College of Engineering", None)], "2023 – Expected 2027", before=2)
    certs = doc.add_paragraph()
    tighten(certs, before=0, after=0, exact=12)
    add_link(certs, "Machine Learning Specialization, Coursera", "https://www.coursera.org/account/accomplishments/specialization/F97RXO11QI47", 10, italic=True)
    add_text(certs, " · IBM Generative AI Prompt Engineering · ", 10, italic=True)
    add_link(certs, "Tata Group GenAI Data Analytics (Forage)", "https://www.theforage.com/simulations/tata/data-analytics-t3zr", 10, italic=True)
    add_text(certs, ", 2024", 10, italic=True)

    # Keep tables from adding extra paragraph spacing after them.
    for table in doc.tables:
        tbl = table._tbl
        tblPr = tbl.tblPr
        spacing = OxmlElement("w:tblCellSpacing")
        spacing.set(qn("w:w"), "0")
        spacing.set(qn("w:type"), "dxa")
        tblPr.append(spacing)

    doc.core_properties.title = "Bannusha Shaik — Data Engineer"
    doc.core_properties.author = "Bannusha Shaik"
    doc.core_properties.subject = "Data Engineer resume"
    doc.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
