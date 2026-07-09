#!/usr/bin/env python3
"""Generates HOW_TO_USE.pdf — the instructional guide for the Live Scheduling
Tracker. Rerun after workflow changes:  python3 make_guide.py"""
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, PageBreak, KeepTogether)

APP_URL = "https://andrewgh99.github.io/Live-Scheduling-Tracker/"
REPO_URL = "https://github.com/AndrewGH99/Live-Scheduling-Tracker"

# palette (matches the app)
NAVY = colors.HexColor("#1F4E78")
BLUE = colors.HexColor("#2E75B6")
GREEN_BG = colors.HexColor("#C6EFCE")
YELLOW_BG = colors.HexColor("#FFEB9C")
ORANGE_BG = colors.HexColor("#FFCC99")
RED_BG = colors.HexColor("#FFC7CE")
BLACK_BG = colors.HexColor("#808080")
LIGHT = colors.HexColor("#EEF3F9")
GRID = colors.HexColor("#BBBBBB")

ss = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=ss["Title"], fontSize=20, textColor=NAVY,
                    alignment=0, spaceAfter=2)
SUB = ParagraphStyle("SUB", parent=ss["Normal"], fontSize=10.5,
                     textColor=colors.HexColor("#555555"), spaceAfter=10)
H2 = ParagraphStyle("H2", parent=ss["Heading2"], fontSize=13, textColor=NAVY,
                    spaceBefore=14, spaceAfter=6)
BODY = ParagraphStyle("BODY", parent=ss["Normal"], fontSize=10, leading=14)
STEP = ParagraphStyle("STEP", parent=BODY, leftIndent=2)
NOTE = ParagraphStyle("NOTE", parent=BODY, fontSize=9,
                      textColor=colors.HexColor("#555555"))


def tstyle(header_rows=1):
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, header_rows - 1), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, header_rows - 1), colors.white),
        ("FONTNAME", (0, 0), (-1, header_rows - 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, GRID),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ])


def P(text, style=BODY):
    return Paragraph(text, style)


story = []

# ---------- Page 1: what & how ----------
story.append(P("Live Scheduling Tracker — How to Use", H1))
story.append(P("GH Logistics · Memphis RPDC · tour coverage &amp; weekly reporting", SUB))

story.append(P(
    f'<b>Open the tracker:</b> <a href="{APP_URL}" color="#2E75B6"><u>{APP_URL}</u></a>'
    ' &nbsp;—&nbsp; works on any phone or computer, nothing to install.', BODY))
story.append(Spacer(1, 8))

story.append(P("What it's for", H2))
story.append(P(
    "Confirm FTC coverage across the <b>3 USPS tours × 7 days</b>, see exactly where "
    "staffing is short against the SOP station-priority model, then <b>finalize the week "
    "and export ONE standalone report</b>. The goal is a finished, shareable weekly "
    "report — not a tracker you must keep open forever.", BODY))

story.append(P("The 5-step weekly process", H2))
steps = [
    ("1 — Load the roster",
     "On the <b>Roster</b> tab: connect the shared <b>Google Sheet</b> (paste its link once, "
     "then hit <b>Refresh</b> whenever someone is hired or quits), drag in an "
     "<b>INPUT_TEMPLATE.xlsx</b>, or edit cells directly. Each person's day cells are their "
     "confirmed availability — fastest entry is the dropdown: "
     "<b>Off / Tour 1 / Tour 2 / Tour 3 / Avail.</b>"),
    ("2 — Review coverage",
     "On <b>Board &amp; Gaps</b>, read the 3-tours × 7-days grid. Colors show how many of "
     "the 4 stations are manned. The <b>Gaps table</b> lists every slot below 4 bodies, "
     "worst first — that is your hiring / shift-move punch list."),
    ("3 — Finalize each day",
     "Click a day chip once its coverage is locked (or <b>Finalize all 7</b>). The counter "
     "tracks progress to 7/7 — you walk the week deliberately, day by day, instead of "
     "person by person."),
    ("4 — (Optional) Monitor live",
     "During a shift, the <b>Live Now</b> tab auto-detects the current tour. Toggle who is "
     "physically on; anyone scheduled right now but not toggled on is flagged <b>GAP</b> in red."),
    ("5 — Export the report",
     "Click <b>Export Weekly Report</b>. You get one self-contained HTML file — open, print, "
     "email, or post it. It needs nothing installed and freezes the finalized week as a snapshot."),
]
rows = [[P(f"<b>{t}</b>", STEP), P(d, STEP)] for t, d in steps]
tbl = Table(rows, colWidths=[1.55 * inch, 5.15 * inch])
st = tstyle(header_rows=0)
st.add("ROWBACKGROUNDS", (0, 0), (-1, -1), [LIGHT, colors.white])
tbl.setStyle(st)
story.append(tbl)

story.append(P("Where things live", H2))
story.append(P(
    f'App &amp; files: <a href="{REPO_URL}" color="#2E75B6"><u>{REPO_URL}</u></a><br/>'
    "The <b>Google Sheet</b> (maintained by HR) is the single source of truth for who is on "
    "the roster. The <b>exported report</b> is the shared record of the week. Edits and "
    "finalized days save automatically in your browser.", BODY))

story.append(PageBreak())

# ---------- Page 2: reference ----------
story.append(P("Reference — tours, stations, colors", H1))
story.append(P("Why the board shows what it shows", SUB))

story.append(P("The 3 USPS tours", H2))
t = Table([["Tour", "Window", "Kind"],
           ["Tour 1", "11:00 PM – 6:00 AM", "overnight"],
           ["Tour 2", "6:00 AM – 2:00 PM", "day"],
           ["Tour 3", "2:00 PM – 10:00 PM", "evening"]],
          colWidths=[1.2 * inch, 2.2 * inch, 3.3 * inch])
t.setStyle(tstyle())
story.append(t)
story.append(Spacer(1, 4))
story.append(P("A person counts toward a tour when their hours cover at least <b>half</b> of it. "
               "“Avail.” (the 24/7 lead) counts toward every tour.", NOTE))

story.append(P("The 4 stations — filled in this priority order (SOP)", H2))
stations = [
    ["#1 · Inside/Outside Floor Lead",
     "ALWAYS staffed 24/7/365. Interior floor + exterior docks, first responder, driver "
     "badging / PPE / check-in. The one station that can never be empty."],
    ["#2 · Outside Docks",
     "Continuous dock patrol, hourly chock + trailer audits, meets every arriving truck, "
     "strap collection, driver-compliance fines, trucks unhooked & sealed before departure."],
    ["#3 · Trailer Yard / Inbound",
     "Inbound trailer inspection (four-photo process), running yard inventory/audit, "
     "out-of-service flags, daily technician check, trailer moves."],
    ["#4 · Overflow / Support",
     "Activated only when a 4th body is available. Backs up #2 & #3, carrier hunting, "
     "trailer decaling, owns the 3 PM / 3 AM reactivated-trailer list."],
]
t = Table([[P(f"<b>{a}</b>", STEP), P(b, STEP)] for a, b in stations],
          colWidths=[2.1 * inch, 4.6 * inch])
st = tstyle(header_rows=0)
st.add("ROWBACKGROUNDS", (0, 0), (-1, -1), [LIGHT, colors.white])
t.setStyle(st)
story.append(t)

story.append(P("What the colors mean (bodies on a tour → stations manned)", H2))
color_rows = [
    ["Color", "Bodies", "What it means"],
    ["GREEN", "4+", "All 4 stations manned — full operation with Overflow/surge capacity."],
    ["YELLOW", "3", "#1 · #2 · #3 covered; Overflow (#4) down. Core ops fine, no surge."],
    ["ORANGE", "2", "#1 · #2 only. Trailer Yard/Inbound unmanned — inbound inspection at risk."],
    ["RED", "1", "Floor Lead only; docks unmanned. Critical single point of failure."],
    ["BLACK", "0", "No one on the tour — breaches the 24/7 Station #1 mandate."],
]
t = Table([[P(f"<b>{r[0]}</b>", STEP) if i else r[0], r[1], P(r[2], STEP) if i else r[2]]
           for i, r in enumerate(color_rows)],
          colWidths=[0.9 * inch, 0.8 * inch, 5.0 * inch])
st = tstyle()
for i, bg in [(1, GREEN_BG), (2, YELLOW_BG), (3, ORANGE_BG), (4, RED_BG), (5, BLACK_BG)]:
    st.add("BACKGROUND", (0, i), (0, i), bg)
st.add("TEXTCOLOR", (0, 5), (0, 5), colors.white)
st.add("ALIGN", (1, 0), (1, -1), "CENTER")
t.setStyle(st)
story.append(t)
story.append(Spacer(1, 4))
story.append(P("Baseline = <b>4 bodies per tour</b> (all four stations covered). The Gaps table "
               "lists every tour × day below baseline, worst first.", NOTE))

fmt = Table([
    ["You type / pick", "Meaning"],
    ["Off  (or leave blank)", "Not working that day"],
    ["Tour 1 (11p–6a) / Tour 2 (6a–2p) / Tour 3 (2p–10p)", "Works that full tour — pick from the dropdown"],
    ["Avail.", "24/7 lead — counts toward every tour"],
    ["9a–6p · 10p–8a · 6:30a–3p", "Exact hours (overnight handled automatically)"],
    ["10–11p / 8–9a", "Split shift — separate segments with a slash"],
], colWidths=[3.1 * inch, 3.6 * inch])
fmt.setStyle(tstyle())
story.append(KeepTogether([
    P("Filling in hours — accepted formats", H2),
    fmt,
    Spacer(1, 10),
    P("<b>Questions or roster changes:</b> contact the scheduler / Accountability "
      "Manager. HR updates the shared Google Sheet; everyone else just hits Refresh.", NOTE),
]))

doc = SimpleDocTemplate("HOW_TO_USE.pdf", pagesize=letter,
                        leftMargin=0.7 * inch, rightMargin=0.7 * inch,
                        topMargin=0.65 * inch, bottomMargin=0.65 * inch,
                        title="Live Scheduling Tracker — How to Use",
                        author="GH Logistics")
doc.build(story)
print("✓ HOW_TO_USE.pdf generated")
