"""One-page resume in the Yassine layout: serif header, photo, ruled sections."""
from pathlib import Path

from PIL import Image
from reportlab.lib.colors import Color, HexColor, black, white
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    Image as RLImage,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "resume.pdf"
PHOTO = ROOT / "photo.jpg"
CROP = ROOT / "_photo_crop.jpg"
BLUE = HexColor("#1A5276")
RULE = HexColor("#222222")
MUTED = HexColor("#333333")


def crop_photo():
    im = Image.open(PHOTO).convert("RGB")
    w, h = im.size
    side_w, side_h = 820, 980
    cx, cy = int(w * 0.56), int(h * 0.28)
    left = max(0, min(w - side_w, cx - side_w // 2))
    top = max(0, min(h - side_h, cy - int(side_h * 0.38)))
    im.crop((left, top, left + side_w, top + side_h)).save(CROP, quality=90)


def styles():
    return {
        "name": ParagraphStyle("name", fontName="Times-Bold", fontSize=22, leading=24, alignment=TA_CENTER, textColor=black),
        "role": ParagraphStyle("role", fontName="Times-Roman", fontSize=12, leading=14, alignment=TA_CENTER, textColor=black),
        "contact": ParagraphStyle("contact", fontName="Times-Roman", fontSize=9, leading=12, alignment=TA_CENTER, textColor=MUTED),
        "section": ParagraphStyle("section", fontName="Times-Bold", fontSize=11, leading=12, textColor=black, spaceBefore=5, spaceAfter=0),
        "job": ParagraphStyle("job", fontName="Times-Bold", fontSize=10, leading=12, textColor=BLUE),
        "date": ParagraphStyle("date", fontName="Times-Roman", fontSize=9.5, leading=12, alignment=TA_RIGHT, textColor=black),
        "sub": ParagraphStyle("sub", fontName="Times-Italic", fontSize=9, leading=11, textColor=MUTED),
        "body": ParagraphStyle("body", fontName="Times-Roman", fontSize=9, leading=11, textColor=black, alignment=TA_JUSTIFY),
        "bullet": ParagraphStyle("bullet", fontName="Times-Roman", fontSize=9, leading=11, textColor=black, leftIndent=11, bulletIndent=0),
        "tech": ParagraphStyle("tech", fontName="Times-Roman", fontSize=9, leading=11.2, textColor=black),
    }


def rule():
    line = Table([[""]], colWidths=[7.5 * inch])
    line.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 0.6, RULE),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))
    return line


def section(s, title):
    return [Paragraph(title, s["section"]), rule()]


def job_row(s, left, right):
    t = Table(
        [[Paragraph(left, s["job"]), Paragraph(right, s["date"])]],
        colWidths=[5.5 * inch, 2.0 * inch],
    )
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return t


def bullet(s, text):
    return Paragraph("•  " + text, s["bullet"])


def header(s):
    photo = RLImage(str(CROP), width=0.92 * inch, height=1.1 * inch)
    text = [
        Paragraph("Bannusha SHAIK", s["name"]),
        Paragraph("Data Engineer", s["role"]),
        Spacer(1, 3),
        Paragraph("Bangalore &nbsp;&nbsp;|&nbsp;&nbsp; +91 8074671779 &nbsp;&nbsp;|&nbsp;&nbsp; bannushashaik85@gmail.com", s["contact"]),
        Paragraph("GitHub &nbsp;&nbsp;|&nbsp;&nbsp; LinkedIn &nbsp;&nbsp;|&nbsp;&nbsp; bannusha.com", s["contact"]),
    ]
    block = Table([[text, photo]], colWidths=[6.4 * inch, 1.1 * inch])
    block.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return block


def main():
    crop_photo()
    s = styles()
    story = [header(s), Spacer(1, 6)]

    story += section(s, "Summary")
    story.append(Paragraph(
        "Data Engineer for operational data models, pipelines, and the tables a team can query. Fortune 500 intern at SKF: Microsoft 365 (Excel, SharePoint, Forms), Power Apps, Power BI, and daily data-accuracy routines that flag a score against a qualifying bar. Built snapshot pipelines and schema-locked SQL so a number is defended against a baseline, not a raw total. Seeking a Data Engineer role.",
        s["body"],
    ))

    story += section(s, "Professional Experience")
    story.append(job_row(s, "SKF — Data Engineer Intern", "May 2026 – Present"))
    story.append(Paragraph("Fortune 500 industrial · Bangalore · Microsoft 365, SharePoint, Power Apps, Power BI", s["sub"]))
    story.append(bullet(s, "Own the operational data routine on the SKF tenant: replaced paper 5S plant audits with an 11-screen Power Apps and SharePoint system factory teams use for scores, photos, and actions."))
    story.append(bullet(s, "Modeled the grain as one audit across five pillars, with identity-aware, zone-scoped write-back, so every submitted audit is a row plant leads can review."))
    story.append(bullet(s, "Served score-card and compliance-trend views: per-pillar totals, pass/fail against a qualifying bar, and the weakest pillar flagged for leadership."))
    story.append(bullet(s, "Support daily and weekly reporting to plant operations: summarize issues and keep the capture rules aligned so the table stays accurate."))

    story.append(job_row(s, "Tata Group (Forage) — GenAI Data Analytics", "2024"))
    story.append(bullet(s, "Modeled delinquency risk on structured financial datasets and wrote the next action a stakeholder could take. A job simulation."))

    story.append(job_row(s, "IEEE Student Branch, PES — Technical Support", "Aug 2023 – Sep 2024"))
    story.append(bullet(s, "College club role. Supported technical sessions: setup, troubleshooting, and keeping events running."))

    story += section(s, "Key Projects")
    story.append(job_row(s, "Creator Lab", "Live"))
    story.append(Paragraph("know-stats.streamlit.app", s["sub"]))
    story.append(bullet(s, "Ingested the YouTube Data API into SQLite for 100 Indian creators and about 3,500 videos. Grain is one snapshot of one video, so views, likes, and comments are a time series rather than a one-time scrape."))
    story.append(bullet(s, "Defined engagement as (likes + comments) / views and outlier as views versus that creator's own median, then shipped a public app that scores a draft and returns a view range plus the nearest videos in the table."))
    story.append(Paragraph("<b>Technologies:</b> Python, SQLite, YouTube Data API, pandas, Streamlit", s["tech"]))

    story.append(job_row(s, "Bengaluru Civic Warehouse", "In progress"))
    story.append(Paragraph("Weather, air, and mobility · one city model", s["sub"]))
    story.append(bullet(s, "Designing a warehouse for three public city feeds that do not share a schema, a clock, or a location. Each pull lands as an untouched raw payload before any cleaning."))
    story.append(bullet(s, "Invalid rows go to a reject table with the rule that failed. Ward, station, and date are dimensions in PostgreSQL and PostGIS. An unknown location is surrogate key −1, not a dropped row."))
    story.append(Paragraph("<b>Technologies:</b> Python, PostgreSQL, PostGIS", s["tech"]))

    story += section(s, "Technical Skills")
    story.append(Paragraph("<b>Data engineering:</b> ETL, data pipelines, data modeling, API ingestion, schema design, snapshot tables, data quality", s["tech"]))
    story.append(Paragraph("<b>SQL &amp; databases:</b> SQL, PostgreSQL, MySQL, SQLite, joins, window functions, parameterized queries", s["tech"]))
    story.append(Paragraph("<b>Python:</b> Python, pandas, NumPy, API clients, scikit-learn", s["tech"]))
    story.append(Paragraph("<b>Tools:</b> Power BI, Streamlit, Grafana, Docker, Git, Power Apps, SharePoint", s["tech"]))

    story += section(s, "Education")
    story.append(job_row(s, "B.Tech, Artificial Intelligence &amp; Machine Learning", "2023 – Expected 2027"))
    story.append(Paragraph("PES College of Engineering, Bangalore", s["sub"]))

    story += section(s, "Certifications")
    story.append(job_row(s, "Machine Learning Specialization", "Coursera"))
    story.append(Paragraph("coursera.org/account/accomplishments/specialization/F97RXO11QI47", s["sub"]))
    story.append(Paragraph("Generative AI: Prompt Engineering Basics, IBM &nbsp;·&nbsp; GenAI Powered Data Analytics, Tata Group (Forage), 2024", s["tech"]))

    doc = SimpleDocTemplate(
        str(OUT),
        pagesize=letter,
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.32 * inch,
        bottomMargin=0.26 * inch,
        title="Bannusha Shaik — Data Engineer",
        author="Bannusha Shaik",
    )
    doc.build(story)
    if CROP.exists():
        CROP.unlink()
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
