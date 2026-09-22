"""One-page ATS Data Engineer resume. Calibri, single column, standard headings."""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Twips

OUT = Path(__file__).resolve().parent / "Bannusha_Shaik_Data_Engineer.docx"


def set_run_font(run, name="Calibri", size=11, bold=False, color=None):
    run.bold = bold
    run.font.name = name
    run.font.size = Pt(size)
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:eastAsia"), name)
    rFonts.set(qn("w:cs"), name)
    if color:
        run.font.color.rgb = RGBColor(*color)


def set_spacing(p, before=0, after=0, line=230):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line / 240.0
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE


def add_bottom_border(p, size="12", color="000000"):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def heading(doc, text):
    p = doc.add_paragraph()
    set_spacing(p, before=7, after=3, line=200)
    add_bottom_border(p, "8", "000000")
    r = p.add_run(text.upper())
    set_run_font(r, size=11, bold=True)
    return p


def bullet(doc, text, num_id=1):
    p = doc.add_paragraph(style="List Bullet")
    set_spacing(p, before=0, after=1, line=210)
    p.clear()
    r = p.add_run(text)
    set_run_font(r, size=10.5)
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.15)
    return p


def job_line(doc, left, right):
    p = doc.add_paragraph()
    set_spacing(p, before=6, after=0, line=220)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(7.3), WD_TAB_ALIGNMENT.RIGHT)
    r = p.add_run(left)
    set_run_font(r, size=11, bold=True)
    r2 = p.add_run("\t" + right)
    set_run_font(r2, size=10.5)


def loc_line(doc, text):
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=2, line=200)
    r = p.add_run(text)
    set_run_font(r, size=10)
    r.italic = True


def skill_line(doc, label, rest):
    p = doc.add_paragraph()
    set_spacing(p, before=1, after=1, line=220)
    r = p.add_run(label)
    set_run_font(r, size=10.5, bold=True)
    r2 = p.add_run(rest)
    set_run_font(r2, size=10.5)


def main():
    doc = Document()
    for s in doc.sections:
        s.page_width = Inches(8.5)
        s.page_height = Inches(11)
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.6)
        s.right_margin = Inches(0.6)

    name = doc.add_paragraph()
    name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(name, before=0, after=2, line=220)
    r = name.add_run("BANNUSHA SHAIK")
    set_run_font(r, size=18, bold=True)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(sub, before=0, after=2, line=200)
    r = sub.add_run("Data Engineer  |  AI & Machine Learning  |  Python  ·  SQL  ·  ETL  ·  Data Modeling")
    set_run_font(r, size=11)

    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(contact, before=0, after=0, line=200)
    r = contact.add_run(
        "Bangalore, India  ·  +91 8074671779  ·  bannushashaik85@gmail.com"
    )
    set_run_font(r, size=9.5)
    contact2 = doc.add_paragraph()
    contact2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(contact2, before=0, after=6, line=200)
    add_bottom_border(contact2, "12", "000000")
    r = contact2.add_run(
        "linkedin.com/in/bannushashaik400  ·  github.com/bannushaxddd  ·  bannusha.com"
    )
    set_run_font(r, size=9.5)

    heading(doc, "Summary")
    sm = doc.add_paragraph()
    set_spacing(sm, before=3, after=2, line=216)
    r = sm.add_run(
        "Data Engineer with a B.Tech in Artificial Intelligence & Machine Learning and hands-on "
        "work building pipelines, data models, and the applications that sit on them. At SKF "
        "(Fortune 500), turned paper factory audits into a production data system: SharePoint-backed "
        "rows, zone-scoped write-back, and score-card views plant leads use. Independently built a "
        "YouTube snapshot pipeline (100 creators, ~3,500 videos) and a schema-locked SQL generator "
        "across 8 dialects with PostgreSQL history. Skilled in Python, SQL, PostgreSQL, data modeling, "
        "APIs, and ETL. Seeking a Data Engineer role."
    )
    set_run_font(r, size=10.5)

    heading(doc, "Professional Experience")
    job_line(
        doc,
        "Data Engineer Intern, SKF (Fortune 500 Industrial Technology)",
        "May 2026 – Present  |  Bangalore",
    )
    loc_line(doc, "Microsoft Power Platform  ·  SharePoint  ·  Office 365  ·  Power BI  ·  factory operations")
    for t in [
        "Replaced paper 5S factory audits with a production 11-screen Power Apps system on the SKF Office 365 tenant; factory teams capture scores, photo evidence, and actions as structured records.",
        "Modeled operational data (audit drafts, zone lists) with identity-aware, zone-scoped write-back into SharePoint so an audit is a queryable row, not a file in email.",
        "Defined grain as one audit across five 5S pillars and served score-card and compliance-trend views: per-pillar totals, pass/fail against a qualifying bar, weakest pillar for plant leads.",
    ]:
        bullet(doc, t)

    job_line(doc, "GenAI Data Analytics Job Simulation, Tata Group (Forage)", "2024")
    bullet(
        doc,
        "Modeled delinquency risk on structured financial datasets and wrote the next action a stakeholder could take. Simulation, labeled as such.",
    )

    heading(doc, "Projects")
    job_line(
        doc,
        "Indian Creator Lab  |  YouTube Data API, SQLite, pandas, scikit-learn, Streamlit",
        "Live",
    )
    loc_line(doc, "know-stats.streamlit.app  ·  github.com/bannushaxddd/indian-creator-lab")
    for t in [
        "Ingested the YouTube Data API into SQLite for 100 Indian creators (~3,500 videos) using a video key and a snapshot grain so views, likes, and comments are a time series, not a one-time scrape.",
        "Defined engagement as (likes + comments) / views and outlier as views / that creator’s median, then served a Streamlit app that scores a draft title and cut against the table and returns a view range plus nearest videos.",
    ]:
        bullet(doc, t)

    job_line(doc, "GETYOQUERY  |  SQL, PostgreSQL, 8 dialects, schema-locked generation", "")
    loc_line(doc, "github.com/bannushaxddd/GETYOQUERY")
    bullet(
        doc,
        "Built an English-to-SQL service across 8 dialects (PostgreSQL, MySQL, SQLite, SQL Server, BigQuery, Snowflake, Oracle, DuckDB) that accepts CREATE TABLE as a hard constraint, stores query history in PostgreSQL, and runs parameterized queries so generated SQL cannot invent column names.",
    )

    job_line(doc, "Pipeline observability  |  Prometheus, Grafana, Loki, Alertmanager, Docker", "")
    loc_line(doc, "github.com/bannushaxddd/prometheus-grafana-stack")
    bullet(
        doc,
        "Stood up a Docker Compose metrics and log stack: Prometheus, Loki, Alertmanager, database exporters, custom application /metrics, and Grafana boards for system, app, and database health.",
    )

    heading(doc, "Technical Skills")
    skill_line(doc, "Data engineering: ", "ETL, data modeling, snapshot pipelines, API ingestion, schema design, data quality constraints")
    skill_line(doc, "SQL & databases: ", "SQL, PostgreSQL, MySQL, SQLite, joins, grain, parameterized queries, query history")
    skill_line(doc, "Python: ", "Python, pandas, NumPy, API clients, feature tables, scikit-learn")
    skill_line(doc, "Serve & operate: ", "Streamlit, Power BI, Grafana, Docker, Redis, Git")
    skill_line(doc, "Also: ", "Power Apps, SharePoint, Office 365, Excel")

    heading(doc, "Education & Certifications")
    job_line(doc, "B.Tech, Artificial Intelligence & Machine Learning, PES College of Engineering", "Bangalore")
    p = doc.add_paragraph()
    set_spacing(p, before=2, after=0, line=220)
    r = p.add_run(
        "Certifications: Generative AI Prompt Engineering Basics, IBM  ·  GenAI Powered Data Analytics, Tata Group (Forage)"
    )
    set_run_font(r, size=10.5)

    doc.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
