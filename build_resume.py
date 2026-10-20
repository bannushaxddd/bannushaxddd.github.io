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

FONT_DIR = Path("/System/Library/Fonts/Supplemental")
pdfmetrics.registerFont(TTFont("Cal", str(FONT_DIR / "Arial.ttf")))
pdfmetrics.registerFont(TTFont("CalB", str(FONT_DIR / "Arial Bold.ttf")))
pdfmetrics.registerFont(TTFont("CalI", str(FONT_DIR / "Arial Italic.ttf")))

OUT = Path(__file__).resolve().parent / "resume.pdf"
INK = HexColor("#111111")
RULE = HexColor("#222222")
LINK = "#111111"


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


def href(label, url):
    return f'<link href="{url}" color="{LINK}"><u>{label}</u></link>'


def main():
    s = styles()
    story = [
        Paragraph("BANNUSHA SHAIK", s["name"]),
        Paragraph("Data Engineer &nbsp;|&nbsp; Power BI · SharePoint · Power Apps · Data Modeling", s["title"]),
        Paragraph(
            "Bangalore, India · +91 8074671779 · "
            + href("bannushashaik85@gmail.com", "mailto:bannushashaik85@gmail.com")
            + " · "
            + href("linkedin.com/in/bannushashaik400", "https://linkedin.com/in/bannushashaik400")
            + " · "
            + href("bannusha.com", "https://bannusha.com"),
            s["contact"],
        ),
    ]

    story += section(s, "SUMMARY")
    story.append(Paragraph(
        "Data Engineer intern at SKF, with a B.Tech in AI and Machine Learning. Stack: SQL, Python, pandas, Power BI, DAX, SharePoint, Power Apps, Snowflake, Databricks, Microsoft Fabric, dbt, Apache Airflow, Kafka, Delta Lake, and Apache Iceberg. At SKF I ship semantic models for project lead time and scrap value, and an 11-screen audit app.",
        s["body"],
    ))

    story += section(s, "PROFESSIONAL EXPERIENCE")
    story.append(row(s, "Data Engineer Intern, SKF (Fortune 500 Industrial)", "May 2026 – Present | Bangalore"))
    story.append(Paragraph("Power BI · SharePoint · Power Apps · Excel · DAX · data modeling · Office 365", s["tools"]))
    story.append(bullet(s, "Shipped an industrial project tracker: one row per activity, average beside P90, and on-time by task category, product family, and assignee."))
    story.append(bullet(s, "Shipped a finance scrap model on the raw posting workbook, not the year-to-date summary. Grain is one line. A visual ships only if the file can support it."))
    story.append(bullet(s, "Replaced paper 5S audits with an 11-screen Power Apps and SharePoint system. Grain is one audit across five pillars, with identity-aware, zone-scoped write-back and pillar scorecards."))

    story.append(row(s, "GenAI Data Analytics Job Simulation, Tata Group (Forage)", "2024"))
    story.append(bullet(s, "Worked delinquency risk on structured financial datasets and wrote the next action a stakeholder could take, tied to the fields in the data. A job simulation, not employment."))

    story.append(row(s, "Technical Support, IEEE Student Branch, PES College of Engineering", "Aug 2023 – Sep 2024"))
    story.append(bullet(s, "College club. Setup and troubleshooting for technical sessions so the room was ready before the talk started."))

    story += section(s, "PROJECTS")
    story.append(row(s, "Elgin-GTCI Industrial Project Tracker", "SKF · Power BI"))
    story.append(Paragraph("Semantic model · task category · product family · priority · assignee", s["tools"]))
    story.append(bullet(s, "Operating report for the industrial project queue. Average and P90 sit side by side so one long job cannot hide inside the mean."))
    story.append(bullet(s, "Filters for input date, product family, task category, priority, and assignee keep a plant review on one page."))

    story.append(row(s, "Finance Scrap Analysis", "SKF · SharePoint Excel"))
    story.append(Paragraph("SharePoint workbook · one row per posting · reason, material, item, day", s["tools"]))
    story.append(bullet(s, "Finance model on the raw posting file, not a summary workbook. Cards, daily scrap, reasons, material mix, and scrap by item."))
    story.append(bullet(s, "A posting-day view replaces a month chart the file cannot support, so the report stays aligned to the source."))

    story.append(row(s, "5S Audit Management", "SKF · Power Apps"))
    story.append(Paragraph("SharePoint lists · 11 screens · identity-aware write-back · Power BI scorecards", s["tools"]))
    story.append(bullet(s, "Paper checklists became queryable rows: scores, photo evidence, and corrective actions across five pillars."))
    story.append(bullet(s, "Drafts stay out of the trend. Write-back is limited to the zones an identity is allowed to touch. Scorecards flag the weakest pillar."))

    story += section(s, "TECHNICAL SKILLS")
    story.append(Paragraph("<b>Data engineering:</b> data modeling, grain, measures, SharePoint lists, Excel sources, data quality", s["skill"]))
    story.append(Paragraph("<b>BI &amp; apps:</b> Power BI, DAX, semantic models, Power Apps, scorecards, Office 365", s["skill"]))
    story.append(Paragraph("<b>SQL &amp; Python:</b> PostgreSQL, MySQL, SQLite, pandas, Snowflake, Databricks, Microsoft Fabric, dbt, Apache Airflow, Kafka, Delta Lake, Apache Iceberg", s["skill"]))
    story.append(Paragraph("<b>Plant metrics:</b> lead time, on-time completion, scrap value, reason codes, 5S scores", s["skill"]))

    story += section(s, "EDUCATION &amp; CERTIFICATIONS")
    story.append(row(s, "B.Tech, Artificial Intelligence &amp; Machine Learning, PES College of Engineering", "2023 – Expected 2027"))
    story.append(Paragraph(
        href("Machine Learning Specialization, Coursera", "https://www.coursera.org/account/accomplishments/specialization/F97RXO11QI47")
        + " · IBM Generative AI Prompt Engineering"
        + " · "
        + href("Tata Group GenAI Data Analytics (Forage)", "https://www.theforage.com/simulations/tata/data-analytics-t3zr")
        + ", 2024",
        s["tools"],
    ))

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
