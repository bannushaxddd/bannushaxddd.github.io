"""Dense one-page Data Engineer resume in the Bannusha_Shaik.pdf layout."""
from pathlib import Path

from reportlab.lib.colors import HexColor, black
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.lib.pagesizes import A4
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
    # Type scale measured from Yassine_Erradouani_Data_Engineer.pdf (A4, Charter).
    # Calibri is the installed face; sizes match his: name 25, role/sections 12, body 10.
    return {
        "name": ParagraphStyle("name", fontName="CalB", fontSize=25, leading=27, alignment=TA_CENTER, textColor=INK, spaceAfter=0),
        "title": ParagraphStyle("title", fontName="CalB", fontSize=12, leading=14, alignment=TA_CENTER, textColor=INK, spaceBefore=1),
        "contact": ParagraphStyle("contact", fontName="Cal", fontSize=10, leading=12, alignment=TA_CENTER, textColor=INK, spaceBefore=1),
        "sec": ParagraphStyle("sec", fontName="CalB", fontSize=12, leading=14, textColor=INK, spaceBefore=6, spaceAfter=0),
        "job": ParagraphStyle("job", fontName="CalB", fontSize=10, leading=12, textColor=INK),
        "date": ParagraphStyle("date", fontName="Cal", fontSize=10, leading=12, alignment=TA_RIGHT, textColor=INK),
        "tools": ParagraphStyle("tools", fontName="CalI", fontSize=10, leading=12, textColor=INK),
        "body": ParagraphStyle("body", fontName="Cal", fontSize=10, leading=12.2, textColor=INK, alignment=TA_JUSTIFY),
        "bullet": ParagraphStyle("bullet", fontName="Cal", fontSize=10, leading=12.2, textColor=INK, leftIndent=11),
        "skill": ParagraphStyle("skill", fontName="Cal", fontSize=10, leading=12.2, textColor=INK),
    }


CONTENT_W = 7.49 * inch  # A4 minus 0.38in margins


def rule():
    line = Table([[""]], colWidths=[CONTENT_W])
    line.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (-1, -1), 0.6, RULE),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
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
        colWidths=[5.05 * inch, 2.44 * inch],
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
        "Data Engineer for operational models, pipelines, and the tables a team can query. Fortune 500 intern at SKF, Bangalore: paper 5S audits are now an 11-screen Power Apps and SharePoint system factory teams use, with Power BI score-cards that flag a result against a qualifying bar. Independently shipped a snapshot pipeline of 100 creators and about 3,500 videos, and an English-to-SQL service locked to the pasted schema. A city warehouse is in progress.",
        s["body"],
    ))

    story += section(s, "PROFESSIONAL EXPERIENCE")
    story.append(row(s, "Data Engineer Intern, SKF (Fortune 500 Industrial)", "May 2026 – Present | Bangalore"))
    story.append(Paragraph("Microsoft 365 · SharePoint · Power Apps · Power BI · operational data models", s["tools"]))
    story.append(bullet(s, "Replaced paper 5S plant audits with an 11-screen Power Apps and SharePoint system. Factory teams use it for scores, photos, and actions, not files in email."))
    story.append(bullet(s, "Set the grain as one audit across five pillars. Write-back is identity-aware and zone-scoped, so a submitted audit is a row plant leads can review."))
    story.append(bullet(s, "Built the score-card and Power BI views: per-pillar totals, pass/fail against the qualifying bar, and the weakest pillar flagged for leadership."))

    story.append(row(s, "GenAI Data Analytics Job Simulation, Tata Group (Forage)", "2024"))
    story.append(bullet(s, "Worked delinquency risk on structured financial datasets and wrote the next action a stakeholder could take, tied to the fields in the data. A job simulation, not employment."))

    story.append(row(s, "Technical Support, IEEE Student Branch, PES College of Engineering", "Aug 2023 – Sep 2024"))
    story.append(bullet(s, "College club. Setup and troubleshooting for technical sessions so the room was ready before the talk started."))

    story += section(s, "PROJECTS")
    story.append(row(s, "Creator Lab &nbsp;|&nbsp; snapshot pipeline and scoring table", "know-stats.streamlit.app"))
    story.append(Paragraph("Python · SQLite · YouTube Data API · pandas · scikit-learn · Streamlit · 100 creators · about 3,500 videos", s["tools"]))
    story.append(bullet(s, "Pulled the YouTube Data API into SQLite. Grain is one snapshot of one video, across 100 Indian creators (50 tech, 50 fashion) and about 3,500 videos."))
    story.append(bullet(s, "Engagement is (likes + comments) / views. Outlier is views divided by that creator's own median, so a small channel is not ranked against a large one."))
    story.append(bullet(s, "Streamlit scores a draft and returns a view range plus the nearest videos already stored. pandas and scikit-learn sit on the feature table."))

    story.append(row(s, "Schema-Locked SQL &nbsp;|&nbsp; English to SQL across 8 dialects", "github.com/bannushaxddd/GETYOQUERY"))
    story.append(Paragraph("Node.js · PostgreSQL · parameterized execution · JWT · 8 SQL dialects", s["tools"]))
    story.append(bullet(s, "English to SQL for PostgreSQL, MySQL, SQLite, SQL Server, BigQuery, Snowflake, Oracle, and DuckDB. A pasted CREATE TABLE is the allowlist, so generation cannot invent a column."))
    story.append(bullet(s, "Query history is stored in PostgreSQL. Execution is parameterized, and the API sits behind JWT, so a generated statement cannot invent a column or concatenate input into SQL."))

    story.append(row(s, "Bengaluru Urban Warehouse &nbsp;|&nbsp; city feeds into one model", "In progress"))
    story.append(Paragraph("Python · PostgreSQL · PostGIS · weather, air, and mobility", s["tools"]))
    story.append(bullet(s, "Not shipped. Three public feeds — weather, air, and mobility — that do not share a schema, a clock, or a location. Each raw payload is kept."))
    story.append(bullet(s, "Invalid rows go to a reject table with the rule that failed. Ward, station, and date are dimensions in PostgreSQL and PostGIS. An unknown location is surrogate key −1, not a dropped row."))

    story += section(s, "TECHNICAL SKILLS")
    story.append(Paragraph("<b>Data engineering:</b> ETL, data pipelines, data modeling, API ingestion, schema design, snapshot tables, data quality", s["skill"]))
    story.append(Paragraph("<b>SQL &amp; databases:</b> SQL, PostgreSQL, MySQL, SQLite, joins, window functions, parameterized queries", s["skill"]))
    story.append(Paragraph("<b>Python:</b> pandas, NumPy, API clients, scikit-learn", s["skill"]))
    story.append(Paragraph("<b>Serve &amp; plant:</b> Power BI, Streamlit, Grafana, Docker, Git, Power Apps, SharePoint, Microsoft 365", s["skill"]))

    story += section(s, "EDUCATION &amp; CERTIFICATIONS")
    story.append(row(s, "B.Tech, Artificial Intelligence &amp; Machine Learning, PES College of Engineering", "2023 – Expected 2027"))
    story.append(Paragraph("Machine Learning Specialization, Coursera · IBM Generative AI Prompt Engineering · Tata Group GenAI Data Analytics (Forage), 2024", s["tools"]))

    doc = SimpleDocTemplate(
        str(OUT),
        pagesize=A4,
        leftMargin=0.38 * inch,
        rightMargin=0.38 * inch,
        topMargin=0.26 * inch,
        bottomMargin=0.26 * inch,
        title="Bannusha Shaik — Data Engineer",
        author="Bannusha Shaik",
    )
    doc.build(story)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
