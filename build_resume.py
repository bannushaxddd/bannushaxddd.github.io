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


def set_spacing(p, before=0, after=0, line=240):
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
    set_spacing(p, before=5, after=1, line=200)
    add_bottom_border(p, "8", "000000")
    r = p.add_run(text.upper())
    set_run_font(r, size=11, bold=True)
    return p


def bullet(doc, text, num_id=1):
    p = doc.add_paragraph(style="List Bullet")
    set_spacing(p, before=0, after=0, line=208)
    p.clear()
    r = p.add_run(text)
    set_run_font(r, size=10.5)
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.15)
    return p


def job_line(doc, left, right):
    p = doc.add_paragraph()
    set_spacing(p, before=4, after=0, line=210)
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
    set_spacing(p, before=0, after=0, line=210)
    r = p.add_run(label)
    set_run_font(r, size=10.5, bold=True)
    r2 = p.add_run(rest)
    set_run_font(r2, size=10.5)


def main():
    doc = Document()
    for s in doc.sections:
        s.page_width = Inches(8.5)
        s.page_height = Inches(11)
        s.top_margin = Inches(0.38)
        s.bottom_margin = Inches(0.32)
        s.left_margin = Inches(0.55)
        s.right_margin = Inches(0.55)

    name = doc.add_paragraph()
    name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(name, before=0, after=2, line=220)
    r = name.add_run("BANNUSHA SHAIK")
    set_run_font(r, size=18, bold=True)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(sub, before=0, after=2, line=200)
    r = sub.add_run("Data Engineer  |  Python  ·  SQL  ·  ETL  ·  Data Modeling  ·  PostgreSQL")
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
    set_spacing(sm, before=2, after=1, line=210)
    r = sm.add_run(
        "Data Engineer building ETL pipelines, operational data models, and the tables downstream "
        "teams query. At SKF (Fortune 500), replaced paper factory audits with an 11-screen production "
        "system: each audit is 1 queryable row across 5 pillars, with zone-scoped write-back and "
        "score-card views plant leads use. Skills: Python, SQL, PostgreSQL, pandas, data modeling, "
        "API ingestion, ETL, and Power BI. Seeking a Data Engineer role."
    )
    set_run_font(r, size=10.5)

    heading(doc, "Professional Experience")
    job_line(
        doc,
        "Data Engineer Intern, SKF (Fortune 500 Industrial Technology)",
        "May 2026 – Present  |  Bangalore",
    )
    for t in [
        "Replaced paper 5S factory audits with an 11-screen production Power Apps system on the SKF Office 365 tenant; factory teams now capture scores, photo evidence, and actions as structured records instead of paper forms.",
        "Modeled audit drafts and zone lists with identity-aware, zone-scoped write-back into SharePoint, so each submitted audit is 1 queryable row rather than a file in email.",
        "Defined the grain as 1 audit across 5 pillars and served score-card views to plant leads: per-pillar totals, pass/fail against a qualifying bar, and the weakest pillar on every audit.",
    ]:
        bullet(doc, t)

    job_line(doc, "GenAI Data Analytics Job Simulation, Tata Group (Forage)", "2024")
    bullet(
        doc,
        "Modeled delinquency risk on structured financial datasets and wrote the next action a stakeholder could take.",
    )

    job_line(doc, "Technical Support, IEEE Student Branch, PES College of Engineering", "Aug 2023 – Sep 2024")
    bullet(
        doc,
        "College club role. Supported technical sessions for the IEEE student branch: setup, troubleshooting, and keeping events running for speakers and attendees.",
    )

    heading(doc, "Projects")
    job_line(doc, "Creator Lab  |  Python, SQLite, YouTube Data API, Streamlit", "Live")
    loc_line(doc, "know-stats.streamlit.app  ·  github.com/bannushaxddd/indian-creator-lab")
    for t in [
        "Ingested the YouTube Data API v3 into SQLite for 100 Indian creators (50 tech, 50 fashion and beauty) and about 3,500 videos. Grain is one snapshot of one video, so views, likes, and comments can be collected again instead of overwritten.",
        "Defined engagement as (likes + comments) / views, because public YouTube does not expose shares. Defined outlier as this video's views divided by that creator's own median, so a small channel is not ranked against a large one.",
        "Built a feature table for duration buckets, Shorts versus long-form, title hooks, and publish hour, then shipped a public Streamlit app that scores a draft title and cut and returns a view range, an engagement rate, and the nearest videos already in the table.",
    ]:
        bullet(doc, t)

    job_line(doc, "NammaPulse  |  Python, PostgreSQL, PostGIS", "In progress")
    for t in [
        "Designing a Bengaluru warehouse for weather, air, and mobility feeds that do not share a schema, a clock, or a location model. Each pull lands as an untouched raw payload before any cleaning.",
        "Invalid rows — a missing timestamp, a missing location, or a value outside a declared range — go to a reject table with the rule that failed. Ward, station, and date are dimensions in PostgreSQL and PostGIS. An unknown location is surrogate key -1, not a dropped row.",
    ]:
        bullet(doc, t)

    job_line(doc, "GETYOQUERY  |  SQL, PostgreSQL, 8 dialects", "")
    loc_line(doc, "github.com/bannushaxddd/GETYOQUERY")
    bullet(
        doc,
        "Built an English-to-SQL service across 8 dialects: PostgreSQL, MySQL, SQLite, SQL Server, BigQuery, Snowflake, Oracle, and DuckDB. A pasted CREATE TABLE is the allowlist, so generation may use only those table and column names.",
        "Stored query history in PostgreSQL, put JWT on the API, and ran execution with parameterized queries so a generated statement cannot invent a column or concatenate user input into SQL.",
    )

    job_line(doc, "Pipeline observability  |  Prometheus, Grafana, Loki, Docker", "")
    loc_line(doc, "github.com/bannushaxddd/prometheus-grafana-stack")
    bullet(
        doc,
        "Stood up a Docker Compose stack a pipeline needs once it leaves a laptop: Prometheus for metrics, Loki for logs, Alertmanager for routing, and Grafana boards for system, application, and database health.",
        "Added MySQL and Mongo exporters plus a demo API that exposes custom /metrics, so rows, failures, and latency are scraped numbers rather than something checked by opening a log file.",
    )



    heading(doc, "Education")
    job_line(
        doc,
        "B.Tech, Artificial Intelligence & Machine Learning, PES College of Engineering",
        "2023 – Expected 2027",
    )
    loc_line(doc, "Bangalore, India")
    p = doc.add_paragraph()
    set_spacing(p, before=2, after=0, line=220)
    r = p.add_run("Certifications: ")
    set_run_font(r, size=10.5, bold=True)
    r2 = p.add_run(
        "Generative AI: Prompt Engineering Basics, IBM  ·  GenAI Powered Data Analytics, Tata Group (Forage), 2024"
    )
    set_run_font(r2, size=10.5)

    heading(doc, "Technical Skills")
    skill_line(doc, "Data engineering: ", "ETL, data pipelines, data modeling, API ingestion, schema design, snapshot tables, data quality")
    skill_line(doc, "SQL & databases: ", "SQL, PostgreSQL, MySQL, SQLite, joins, window functions, parameterized queries")
    skill_line(doc, "Python: ", "Python, pandas, NumPy, API clients, scikit-learn")
    skill_line(doc, "BI & tools: ", "Power BI, Streamlit, Grafana, Docker, Git, Power Apps, SharePoint, Office 365")

    doc.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
