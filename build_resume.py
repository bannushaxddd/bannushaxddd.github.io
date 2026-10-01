"""Dense one-page Data Engineer resume in the Bannusha_Shaik.pdf layout."""
from pathlib import Path

from reportlab.lib.colors import HexColor, black
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

pdfmetrics.registerFont(TTFont("Cal", r"C:\Windows\Fonts\calibri.ttf"))
pdfmetrics.registerFont(TTFont("CalB", r"C:\Windows\Fonts\calibrib.ttf"))
pdfmetrics.registerFont(TTFont("CalI", r"C:\Windows\Fonts\calibrii.ttf"))

OUT = Path(__file__).resolve().parent / "resume.pdf"
INK = HexColor("#111111")
RULE = HexColor("#222222")


def styles():
    return {
        "name": ParagraphStyle("name", fontName="CalB", fontSize=16, leading=18, alignment=TA_CENTER, textColor=INK),
        "title": ParagraphStyle("title", fontName="Cal", fontSize=9.5, leading=12, alignment=TA_CENTER, textColor=INK),
        "contact": ParagraphStyle("contact", fontName="Cal", fontSize=8.4, leading=11, alignment=TA_CENTER, textColor=INK),
        "sec": ParagraphStyle("sec", fontName="CalB", fontSize=10, leading=12, textColor=INK, spaceBefore=7, spaceAfter=1),
        "job": ParagraphStyle("job", fontName="CalB", fontSize=9.4, leading=11.5, textColor=INK),
        "date": ParagraphStyle("date", fontName="Cal", fontSize=9, leading=11.5, alignment=TA_RIGHT, textColor=INK),
        "tools": ParagraphStyle("tools", fontName="CalI", fontSize=8.3, leading=10.4, textColor=INK),
        "body": ParagraphStyle("body", fontName="Cal", fontSize=8.7, leading=10.8, textColor=INK, alignment=TA_JUSTIFY),
        "bullet": ParagraphStyle("bullet", fontName="Cal", fontSize=8.7, leading=10.8, textColor=INK, leftIndent=10),
        "skill": ParagraphStyle("skill", fontName="Cal", fontSize=8.7, leading=10.8, textColor=INK),
    }


def rule():
    line = Table([[""]], colWidths=[7.5 * inch])
    line.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (-1, -1), 0.7, RULE),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    return line


def section(s, title):
    return [Paragraph(title, s["sec"]), rule()]


def row(s, left, right):
    t = Table(
        [[Paragraph(left, s["job"]), Paragraph(right, s["date"])]],
        colWidths=[5.35 * inch, 2.15 * inch],
    )
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return t


def bullet(s, text):
    return Paragraph("•  " + text, s["bullet"])


def main():
    s = styles()
    story = [
        Paragraph("BANNUSHA SHAIK", s["name"]),
        Paragraph("Data Engineer &nbsp;|&nbsp; Python · SQL · ETL · Data Modeling · PostgreSQL", s["title"]),
        Paragraph("Bangalore, India · +91 8074671779 · bannushashaik85@gmail.com · linkedin.com/in/bannushashaik400 · bannusha.com", s["contact"]),
    ]

    story += section(s, "SUMMARY")
    story.append(Paragraph(
        "Data Engineer for operational data models, pipelines, and the tables a team can query. Fortune 500 intern at SKF: Microsoft 365, SharePoint, Power Apps, and Power BI score-cards that flag a result against a qualifying bar. Built snapshot pipelines and schema-locked SQL so a number is defended against a baseline, not a raw total. Seeking a Data Engineer role.",
        s["body"],
    ))

    story += section(s, "PROFESSIONAL EXPERIENCE")
    story.append(row(s, "Data Engineer Intern, SKF (Fortune 500 Industrial)", "May 2026 – Present | Bangalore"))
    story.append(Paragraph("Microsoft 365 · SharePoint · Power Apps · Power BI · operational data models", s["tools"]))
    story.append(bullet(s, "Own the operational data routine on the SKF tenant: replaced paper 5S plant audits with an 11-screen Power Apps and SharePoint system factory teams use for scores, photos, and actions."))
    story.append(bullet(s, "Modeled the grain as one audit across five pillars, with identity-aware, zone-scoped write-back, so every submitted audit is a row plant leads can review."))
    story.append(bullet(s, "Served score-card and compliance-trend views: per-pillar totals, pass/fail against a qualifying bar, and the weakest pillar flagged for leadership."))

    story.append(row(s, "GenAI Data Analytics Job Simulation, Tata Group (Forage)", "2024"))
    story.append(bullet(s, "Modeled delinquency risk on structured financial datasets and wrote the next action a stakeholder could take. A job simulation."))

    story.append(row(s, "Technical Support, IEEE Student Branch, PES College of Engineering", "Aug 2023 – Sep 2024"))
    story.append(bullet(s, "College club role. Supported technical sessions: setup, troubleshooting, and keeping events running."))

    story += section(s, "PROJECTS")
    story.append(row(s, "Creator Lab &nbsp;|&nbsp; snapshot pipeline and scoring table", "know-stats.streamlit.app"))
    story.append(Paragraph("Python · SQLite · YouTube Data API · pandas · Streamlit · 100 creators · about 3,500 videos", s["tools"]))
    story.append(bullet(s, "Ingested the YouTube Data API into SQLite. Grain is one snapshot of one video, so views, likes, and comments are a time series rather than a one-time scrape."))
    story.append(bullet(s, "Defined engagement as (likes + comments) / views and outlier as views versus that creator's own median, so a small channel is not ranked against a large one."))
    story.append(bullet(s, "Shipped a public app that scores a draft title and cut against the table and returns a view range plus the nearest videos already stored."))

    story.append(row(s, "Schema-Locked SQL &nbsp;|&nbsp; English to SQL across 8 dialects", "GitHub"))
    story.append(Paragraph("SQL · PostgreSQL · Node.js · parameterized queries · JWT", s["tools"]))
    story.append(bullet(s, "Built an English-to-SQL service for PostgreSQL, MySQL, SQLite, SQL Server, BigQuery, Snowflake, Oracle, and DuckDB. A pasted CREATE TABLE is the allowlist, so generation cannot invent a column. Query history is stored in PostgreSQL."))

    story.append(row(s, "Bengaluru Urban Warehouse &nbsp;|&nbsp; city feeds into one model", "In progress"))
    story.append(Paragraph("Python · PostgreSQL · PostGIS · weather, air, and mobility", s["tools"]))
    story.append(bullet(s, "Designing a warehouse for three public city feeds that do not share a schema, a clock, or a location. Each pull lands untouched. Invalid rows go to a reject table. Ward, station, and date are dimensions. An unknown location is surrogate key −1."))

    story += section(s, "TECHNICAL SKILLS")
    story.append(Paragraph("<b>Data engineering:</b> ETL, data pipelines, data modeling, API ingestion, schema design, snapshot tables, data quality", s["skill"]))
    story.append(Paragraph("<b>SQL &amp; databases:</b> SQL, PostgreSQL, MySQL, SQLite, joins, window functions, parameterized queries", s["skill"]))
    story.append(Paragraph("<b>Python:</b> pandas, NumPy, API clients, scikit-learn", s["skill"]))
    story.append(Paragraph("<b>Tools:</b> Power BI, Streamlit, Grafana, Docker, Git, Power Apps, SharePoint, Microsoft 365", s["skill"]))

    story += section(s, "EDUCATION &amp; CERTIFICATIONS")
    story.append(row(s, "B.Tech, Artificial Intelligence &amp; Machine Learning, PES College of Engineering", "2023 – Expected 2027"))
    story.append(Paragraph("Machine Learning Specialization, Coursera · IBM Generative AI Prompt Engineering · Tata Group GenAI Data Analytics (Forage), 2024", s["tools"]))

    doc = SimpleDocTemplate(
        str(OUT),
        pagesize=letter,
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.38 * inch,
        bottomMargin=0.32 * inch,
        title="Bannusha Shaik — Data Engineer",
        author="Bannusha Shaik",
    )
    doc.build(story)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
