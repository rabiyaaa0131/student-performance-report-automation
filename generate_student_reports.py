from pathlib import Path
import re
import sys
from datetime import datetime

import pandas as pd
from PIL import Image

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import mm


# ============================================================
# PROJECT PATHS
# ============================================================

BASE = Path(__file__).resolve().parent

ASSETS = BASE / "assets"
INPUT = BASE / "Input_Data"
INDIVIDUAL_OUTPUT = BASE / "Individual_Reports"
CONSOLIDATED_OUTPUT = BASE / "Consolidated_Reports"

for folder in (INPUT, INDIVIDUAL_OUTPUT, CONSOLIDATED_OUTPUT):
    folder.mkdir(exist_ok=True)


# ============================================================
# INPUT EXCEL FILE
# ============================================================
# If watcher.py sends an Excel file path, use that file.
# Otherwise use the default sample Excel file.

if len(sys.argv) > 1:
    EXCEL_FILE = Path(sys.argv[1])
else:
    EXCEL_FILE = INPUT / "Robotics_50_Students_Sample.xlsx"


# ============================================================
# ASSETS
# ============================================================

SCHOOL_LOGO = ASSETS / "school_logo.jpg"
IKKASHIN_LOGO = ASSETS / "ikkashin_logo.jpg"
ROBOT_IMAGE = ASSETS / "robot_image.jpg"
ROBOT_CROP = ASSETS / "_robot_crop.jpg"


# ============================================================
# STYLES
# ============================================================

styles = getSampleStyleSheet()


def S(name, parent, **kw):
    styles.add(
        ParagraphStyle(
            name=name,
            parent=styles[parent],
            **kw
        )
    )


S(
    "T",
    "Title",
    fontName="Helvetica-Bold",
    fontSize=17,
    leading=19,
    alignment=TA_CENTER,
    textColor=colors.HexColor("#164B67")
)

S(
    "ST",
    "Normal",
    fontSize=8.2,
    leading=10,
    alignment=TA_CENTER,
    textColor=colors.HexColor("#61727B")
)

S(
    "H",
    "Heading2",
    fontName="Helvetica-Bold",
    fontSize=9,
    leading=10,
    textColor=colors.white
)

S(
    "L",
    "Normal",
    fontName="Helvetica-Bold",
    fontSize=6.7,
    leading=8,
    textColor=colors.HexColor("#65757C")
)

S(
    "V",
    "Normal",
    fontSize=8.4,
    leading=10,
    textColor=colors.HexColor("#263238")
)

S(
    "B",
    "BodyText",
    fontSize=7.9,
    leading=10.2,
    textColor=colors.HexColor("#303A40")
)

S(
    "KN",
    "Normal",
    fontName="Helvetica-Bold",
    fontSize=14,
    leading=15,
    alignment=TA_CENTER,
    textColor=colors.HexColor("#164B67")
)

S(
    "KL",
    "Normal",
    fontName="Helvetica-Bold",
    fontSize=6.3,
    leading=7.5,
    alignment=TA_CENTER,
    textColor=colors.HexColor("#65757C")
)

S(
    "F",
    "Normal",
    fontName="Helvetica-Bold",
    fontSize=8.5,
    leading=10,
    alignment=TA_CENTER,
    textColor=colors.white
)


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def safe_filename(s):
    return re.sub(
        r'[\\/:*?"<>|]+',
        "_",
        re.sub(r"\s+", "_", str(s).strip())
    )


def val(row, col, default=""):
    x = row.get(col, default)
    return default if pd.isna(x) else x


def get_report_month(file_path):
    """
    Detect month from Excel filename.

    Example:
    october data.xlsx
    -> October 2026

    If no month is found in the filename,
    current month/year is used.
    """

    filename = file_path.stem.lower()

    months = [
        "january",
        "february",
        "march",
        "april",
        "may",
        "june",
        "july",
        "august",
        "september",
        "october",
        "november",
        "december",
    ]

    for month in months:
        if month in filename:
            return f"{month.capitalize()} {datetime.now().year}"

    return datetime.now().strftime("%B %Y")


def prepare_robot():
    if ROBOT_IMAGE.exists() and not ROBOT_CROP.exists():

        im = Image.open(ROBOT_IMAGE).convert("RGB")

        w, h = im.size

        im.crop(
            (
                int(w * 0.08),
                int(h * 0.02),
                int(w * 0.92),
                int(h * 0.98)
            )
        ).save(
            ROBOT_CROP,
            quality=96
        )


# ============================================================
# INDIVIDUAL STUDENT REPORT
# ============================================================

def make_individual_report(row):

    name = str(
        val(row, "Student Name", "Student")
    ).strip()

    parent = str(
        val(row, "Parent / Guardian", "")
    ).strip()

    grade = str(
        val(row, "Grade", "")
    ).strip()

    section = str(
        val(row, "Section", "")
    ).strip()

    phone = str(
        val(row, "Parent Phone", "")
    ).strip()

    marks = float(
        val(row, "Robotics Marks", 0)
    )

    attendance = float(
        val(row, "Attendance %", 0)
    )

    activities = int(
        float(
            val(row, "Activities", 0)
        )
    )

    present = max(
        0,
        min(
            14,
            round(attendance / 100 * 14)
        )
    )

    absent = 14 - present

    achievement = (
        "Creative Explorer"
        if marks >= 80
        else "Active Learner"
    )

    path = (
        INDIVIDUAL_OUTPUT
        / f"{safe_filename(name)}_Robotics_Report.pdf"
    )

    doc = SimpleDocTemplate(
        str(path),
        pagesize=A4,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=9 * mm,
        bottomMargin=9 * mm
    )

    story = []

    # --------------------------------------------------------
    # TOP RIBBON
    # --------------------------------------------------------

    top = Table(
        [["", "", "", "", "", "", "", ""]],
        colWidths=[22.75 * mm] * 8,
        rowHeights=[3.5 * mm]
    )

    top.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    colors.HexColor("#164B67")
                ),
                (
                    "BACKGROUND",
                    (1, 0),
                    (1, 0),
                    colors.HexColor("#35A7C9")
                ),
                (
                    "BACKGROUND",
                    (3, 0),
                    (3, 0),
                    colors.HexColor("#F2B84B")
                ),
                (
                    "BACKGROUND",
                    (5, 0),
                    (5, 0),
                    colors.HexColor("#D95D8A")
                ),
                (
                    "BACKGROUND",
                    (7, 0),
                    (7, 0),
                    colors.HexColor("#65B98A")
                ),
            ]
        )
    )

    story += [
        top,
        Spacer(1, 2.5 * mm)
    ]

    # --------------------------------------------------------
    # HEADER
    # --------------------------------------------------------

    school = RLImage(
        str(SCHOOL_LOGO)
    )

    school._restrictSize(
        32 * mm,
        30 * mm
    )

    school.hAlign = "CENTER"

    ikk = RLImage(
        str(IKKASHIN_LOGO),
        width=30 * mm,
        height=18 * mm
    )

    robot = RLImage(
        str(
            ROBOT_CROP
            if ROBOT_CROP.exists()
            else ROBOT_IMAGE
        ),
        width=27 * mm,
        height=30 * mm
    )

    session = ParagraphStyle(
        "Session",
        parent=styles["ST"],
        fontSize=7.5,
        textColor=colors.HexColor("#8A5A00")
    )

    head = Table(
        [[
            school,

            [
                Paragraph(
                    "STEM LEARNING OUTCOMES",
                    styles["T"]
                ),

                Paragraph(
                    "ROBOTICS • STUDENT PERFORMANCE REPORT",
                    styles["ST"]
                ),

                Paragraph(
                    "Academic Session 2025–2026",
                    session
                )
            ],

            [ikk, robot]
        ]],
        colWidths=[
            35 * mm,
            110 * mm,
            35 * mm
        ]
    )

    head.setStyle(
        TableStyle(
            [
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),

                (
                    "ALIGN",
                    (0, 0),
                    (0, 0),
                    "CENTER"
                ),

                (
                    "ALIGN",
                    (2, 0),
                    (2, 0),
                    "RIGHT"
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    0
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    0
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    0
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    0
                ),
            ]
        )
    )

    story += [
        head,
        Spacer(1, 2 * mm)
    ]

    # --------------------------------------------------------
    # SECTION BAR
    # --------------------------------------------------------

    def bar(title, bg, accent):

        t = Table(
            [
                [
                    Paragraph(
                        title,
                        styles["H"]
                    ),
                    ""
                ]
            ],
            colWidths=[
                52 * mm,
                128 * mm
            ],
            rowHeights=[
                6.5 * mm
            ]
        )

        t.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, -1),
                        colors.HexColor(bg)
                    ),

                    (
                        "LINEBELOW",
                        (0, 0),
                        (0, 0),
                        2,
                        colors.HexColor(accent)
                    ),

                    (
                        "VALIGN",
                        (0, 0),
                        (-1, -1),
                        "MIDDLE"
                    ),

                    (
                        "LEFTPADDING",
                        (0, 0),
                        (-1, -1),
                        7
                    ),
                ]
            )
        )

        story.append(t)

    # --------------------------------------------------------
    # STUDENT PROFILE
    # --------------------------------------------------------

    bar(
        "STUDENT PROFILE",
        "#2E8FB3",
        "#F2B84B"
    )

    st = Table(
        [
            [
                Paragraph("STUDENT", styles["L"]),
                Paragraph(name, styles["V"]),

                Paragraph(
                    "PARENT / GUARDIAN",
                    styles["L"]
                ),

                Paragraph(
                    parent,
                    styles["V"]
                )
            ],

            [
                Paragraph("GRADE", styles["L"]),
                Paragraph(grade, styles["V"]),

                Paragraph("SECTION", styles["L"]),
                Paragraph(section, styles["V"])
            ],

            [
                Paragraph(
                    "PARENT PHONE",
                    styles["L"]
                ),

                Paragraph(
                    phone,
                    styles["V"]
                ),

                Paragraph(
                    "SUBJECT",
                    styles["L"]
                ),

                Paragraph(
                    "Robotics",
                    styles["V"]
                )
            ]
        ],
        colWidths=[
            28 * mm,
            52 * mm,
            36 * mm,
            64 * mm
        ]
    )

    st.setStyle(
        TableStyle(
            [
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.55,
                    colors.HexColor("#C6D7DE")
                ),

                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.35,
                    colors.HexColor("#DCE7EB")
                ),

                (
                    "BACKGROUND",
                    (0, 0),
                    (0, -1),
                    colors.HexColor("#EFF8FC")
                ),

                (
                    "BACKGROUND",
                    (2, 0),
                    (2, -1),
                    colors.HexColor("#EFF8FC")
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    6
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    6
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),
            ]
        )
    )

    story += [
        st,
        Spacer(1, 3 * mm)
    ]

    # --------------------------------------------------------
    # PERFORMANCE SNAPSHOT
    # --------------------------------------------------------

    bar(
        "PERFORMANCE SNAPSHOT",
        "#3C5B9A",
        "#F2B84B"
    )

    k = Table(
        [
            [
                Paragraph(
                    f"{marks:g} / 100",
                    styles["KN"]
                ),

                Paragraph(
                    f"{attendance:g}%",
                    styles["KN"]
                ),

                Paragraph(
                    str(activities),
                    styles["KN"]
                ),

                Paragraph(
                    f"{present} / 14",
                    styles["KN"]
                )
            ],

            [
                Paragraph(
                    "ROBOTICS SCORE",
                    styles["KL"]
                ),

                Paragraph(
                    "ATTENDANCE",
                    styles["KL"]
                ),

                Paragraph(
                    "ACTIVITIES",
                    styles["KL"]
                ),

                Paragraph(
                    "PRESENT DAYS",
                    styles["KL"]
                )
            ]
        ],
        colWidths=[
            45 * mm
        ] * 4,
        rowHeights=[
            13 * mm,
            7 * mm
        ]
    )

    k.setStyle(
        TableStyle(
            [
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.6,
                    colors.HexColor("#C4D3DB")
                ),

                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.4,
                    colors.HexColor("#D8E3E8")
                ),

                (
                    "BACKGROUND",
                    (0, 0),
                    (0, -1),
                    colors.HexColor("#E5F5FA")
                ),

                (
                    "BACKGROUND",
                    (1, 0),
                    (1, -1),
                    colors.HexColor("#EDE9FB")
                ),

                (
                    "BACKGROUND",
                    (2, 0),
                    (2, -1),
                    colors.HexColor("#FFF4D9")
                ),

                (
                    "BACKGROUND",
                    (3, 0),
                    (3, -1),
                    colors.HexColor("#E6F5EC")
                ),
            ]
        )
    )

    story += [
        k,
        Spacer(1, 3 * mm)
    ]

    # --------------------------------------------------------
    # ACHIEVEMENT + ATTENDANCE
    # --------------------------------------------------------

    aa = Table(
        [[
            Paragraph(
                f"<b>ACHIEVEMENT</b><br/>"
                f"<font size=10 color='#8A5A00'>"
                f"{achievement}"
                f"</font>",
                styles["B"]
            ),

            Paragraph(
                f"<b>ATTENDANCE</b><br/>"
                f"Present {present} • "
                f"Absent {absent} • "
                f"Total 14",
                styles["B"]
            )
        ]],
        colWidths=[
            90 * mm,
            90 * mm
        ]
    )

    aa.setStyle(
        TableStyle(
            [
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.55,
                    colors.HexColor("#C7D5DB")
                ),

                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.35,
                    colors.HexColor("#DCE5E9")
                ),

                (
                    "BACKGROUND",
                    (0, 0),
                    (0, 0),
                    colors.HexColor("#FFF2C9")
                ),

                (
                    "BACKGROUND",
                    (1, 0),
                    (1, 0),
                    colors.HexColor("#DFF2FA")
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    8
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    8
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    6
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    6
                ),
            ]
        )
    )

    story += [
        aa,
        Spacer(1, 3 * mm)
    ]

    # --------------------------------------------------------
    # LEARNING HIGHLIGHTS
    # --------------------------------------------------------

    left = Table(
        [
            [
                Paragraph(
                    "LEARNING HIGHLIGHTS",
                    styles["H"]
                )
            ],

            [
                Paragraph(
                    "✓ Basic robotics concepts<br/>"
                    "✓ Component identification<br/>"
                    "✓ Hands-on model building<br/>"
                    "✓ Team participation<br/>"
                    "✓ Practical problem solving",
                    styles["B"]
                )
            ]
        ],
        colWidths=[
            88 * mm
        ],
        rowHeights=[
            7 * mm,
            35 * mm
        ]
    )

    # --------------------------------------------------------
    # TEACHER OBSERVATION
    # --------------------------------------------------------

    right = Table(
        [
            [
                Paragraph(
                    "TEACHER OBSERVATION",
                    styles["H"]
                )
            ],

            [
                Paragraph(
                    f"{name} participates well in practical "
                    f"activities and follows instructions "
                    f"carefully. The student is developing "
                    f"confidence in building and testing "
                    f"simple robotic models and works "
                    f"positively with classmates.",
                    styles["B"]
                )
            ]
        ],
        colWidths=[
            88 * mm
        ],
        rowHeights=[
            7 * mm,
            35 * mm
        ]
    )

    for t, bg in [
        (left, "#FFF4D9"),
        (right, "#E5F5FA")
    ]:

        t.setStyle(
            TableStyle(
                [
                    (
                        "BOX",
                        (0, 0),
                        (-1, -1),
                        0.55,
                        colors.HexColor("#C7D5DB")
                    ),

                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.HexColor("#2E8FB3")
                    ),

                    (
                        "BACKGROUND",
                        (0, 1),
                        (-1, 1),
                        colors.HexColor(bg)
                    ),

                    (
                        "VALIGN",
                        (0, 0),
                        (-1, -1),
                        "TOP"
                    ),

                    (
                        "LEFTPADDING",
                        (0, 0),
                        (-1, -1),
                        8
                    ),

                    (
                        "RIGHTPADDING",
                        (0, 0),
                        (-1, -1),
                        8
                    ),

                    (
                        "TOPPADDING",
                        (0, 0),
                        (-1, -1),
                        5
                    ),

                    (
                        "BOTTOMPADDING",
                        (0, 0),
                        (-1, -1),
                        5
                    ),
                ]
            )
        )

    main = Table(
        [[left, right]],
        colWidths=[
            90 * mm,
            90 * mm
        ]
    )

    main.setStyle(
        TableStyle(
            [
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    0
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    0
                )
            ]
        )
    )

    story += [
        main,
        Spacer(1, 3 * mm)
    ]

    # --------------------------------------------------------
    # BUILD EXPLORE INNOVATE
    # --------------------------------------------------------

    q = Table(
        [[
            Paragraph(
                "BUILD • EXPLORE • INNOVATE",
                ParagraphStyle(
                    "Q",
                    parent=styles["H"],
                    alignment=TA_CENTER,
                    fontSize=10
                )
            )
        ]],
        colWidths=[
            180 * mm
        ],
        rowHeights=[
            9 * mm
        ]
    )

    q.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    colors.HexColor("#D95D8A")
                )
            ]
        )
    )

    story += [
        q,
        Spacer(1, 3 * mm)
    ]

    # --------------------------------------------------------
    # AUTHORIZATION
    # --------------------------------------------------------

    bar(
        "AUTHORIZATION",
        "#3C5B9A",
        "#D95D8A"
    )

    sig = Table(
        [
            [
                Paragraph(
                    "CLASS TEACHER SIGNATURE",
                    styles["L"]
                ),

                Paragraph(
                    "PRINCIPAL SIGNATURE",
                    styles["L"]
                )
            ],

            [
                "",
                ""
            ]
        ],
        colWidths=[
            90 * mm,
            90 * mm
        ],
        rowHeights=[
            7 * mm,
            15 * mm
        ]
    )

    sig.setStyle(
        TableStyle(
            [
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.55,
                    colors.HexColor("#C7D5DB")
                ),

                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.35,
                    colors.HexColor("#DCE5E9")
                ),

                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor("#EEF3FA")
                ),

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    7
                )
            ]
        )
    )

    story += [
        sig,
        Spacer(1, 3 * mm)
    ]

    # --------------------------------------------------------
    # FOOTER
    # --------------------------------------------------------

    f = Table(
        [[
            Paragraph(
                "✦ Your Growth, Our Priority ✦",
                styles["F"]
            )
        ]],
        colWidths=[
            180 * mm
        ],
        rowHeights=[
            8 * mm
        ]
    )

    f.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    colors.HexColor("#164B67")
                )
            ]
        )
    )

    story.append(f)

    doc.build(story)

    return path


# ============================================================
# CONSOLIDATED MONTHLY REPORT
# ============================================================

def make_consolidated_report(df, report_month=None):

    if report_month is None:
        report_month = datetime.now().strftime("%B %Y")

    marks = pd.to_numeric(
        df["Robotics Marks"],
        errors="coerce"
    ).fillna(0)

    attendance = pd.to_numeric(
        df["Attendance %"],
        errors="coerce"
    ).fillna(0)

    activities = pd.to_numeric(
        df["Activities"],
        errors="coerce"
    ).fillna(0)

    total = len(df)

    avg_marks = marks.mean()
    avg_att = attendance.mean()
    avg_act = activities.mean()

    highest = marks.max()
    lowest = marks.min()

    top_student = str(
        df.loc[
            marks.idxmax(),
            "Student Name"
        ]
    )

    above_80 = int(
        (marks >= 80).sum()
    )

    path = (
        CONSOLIDATED_OUTPUT
        / f"{safe_filename(report_month)}_Consolidated_Report.pdf"
    )

    doc = SimpleDocTemplate(
        str(path),
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm
    )

    body = []

    title = ParagraphStyle(
        "CT",
        parent=styles["T"],
        fontSize=20,
        leading=23
    )

    subtitle = ParagraphStyle(
        "CS",
        parent=styles["ST"],
        fontSize=9,
        leading=12
    )

    body += [
        Paragraph(
            "ROBOTICS • MONTHLY PERFORMANCE ANALYSIS",
            title
        ),

        Spacer(1, 2 * mm),

        Paragraph(
            f"{report_month} | Student Dataset",
            subtitle
        ),

        Spacer(1, 7 * mm)
    ]

    # --------------------------------------------------------
    # KPI CARDS
    # --------------------------------------------------------

    cards = Table(
        [
            [
                Paragraph(
                    str(total),
                    styles["KN"]
                ),

                Paragraph(
                    f"{avg_marks:.1f}",
                    styles["KN"]
                ),

                Paragraph(
                    f"{avg_att:.1f}%",
                    styles["KN"]
                ),

                Paragraph(
                    f"{avg_act:.1f}",
                    styles["KN"]
                )
            ],

            [
                Paragraph(
                    "TOTAL STUDENTS",
                    styles["KL"]
                ),

                Paragraph(
                    "AVG ROBOTICS MARKS",
                    styles["KL"]
                ),

                Paragraph(
                    "AVG ATTENDANCE",
                    styles["KL"]
                ),

                Paragraph(
                    "AVG ACTIVITIES",
                    styles["KL"]
                )
            ]
        ],
        colWidths=[
            45 * mm
        ] * 4,
        rowHeights=[
            14 * mm,
            8 * mm
        ]
    )

    cards.setStyle(
        TableStyle(
            [
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.6,
                    colors.HexColor("#C4D3DB")
                ),

                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.4,
                    colors.HexColor("#D8E3E8")
                ),

                (
                    "BACKGROUND",
                    (0, 0),
                    (0, -1),
                    colors.HexColor("#E5F5FA")
                ),

                (
                    "BACKGROUND",
                    (1, 0),
                    (1, -1),
                    colors.HexColor("#EDE9FB")
                ),

                (
                    "BACKGROUND",
                    (2, 0),
                    (2, -1),
                    colors.HexColor("#FFF4D9")
                ),

                (
                    "BACKGROUND",
                    (3, 0),
                    (3, -1),
                    colors.HexColor("#E6F5EC")
                )
            ]
        )
    )

    body += [
        cards,
        Spacer(1, 7 * mm)
    ]

    # --------------------------------------------------------
    # KEY PERFORMANCE FINDINGS
    # --------------------------------------------------------

    body.append(
        Paragraph(
            "KEY PERFORMANCE FINDINGS",
            ParagraphStyle(
                "BH",
                parent=styles["H"],
                fontSize=10,
                textColor=colors.white
            )
        )
    )

    findings = [
        f"Highest Robotics score: {highest:g}/100 — {top_student}.",

        f"Lowest Robotics score: {lowest:g}/100.",

        f"Students scoring 80 or above: {above_80} of {total}.",

        f"Average Robotics score: {avg_marks:.1f}/100.",

        f"Average attendance: {avg_att:.1f}%.",

        f"Average recorded activities: {avg_act:.1f}."
    ]

    ft = Table(
        [
            [
                Paragraph(
                    "• " + x,
                    styles["B"]
                )
            ]
            for x in findings
        ],
        colWidths=[
            180 * mm
        ]
    )

    ft.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    colors.HexColor("#F5FAFC")
                ),

                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.HexColor("#C7D5DB")
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    8
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    8
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                )
            ]
        )
    )

    body += [
        ft,
        Spacer(1, 7 * mm)
    ]

    # --------------------------------------------------------
    # GRADE / SECTION SUMMARY
    # --------------------------------------------------------

    body.append(
        Paragraph(
            "GRADE / SECTION SUMMARY",
            ParagraphStyle(
                "BH2",
                parent=styles["H"],
                fontSize=10,
                textColor=colors.white
            )
        )
    )

    group = df.copy()

    group["_marks"] = marks
    group["_attendance"] = attendance
    group["_activities"] = activities

    summary = (
        group
        .groupby(
            ["Grade", "Section"],
            dropna=False
        )
        .agg(
            Students=("Student Name", "count"),
            Avg_Marks=("_marks", "mean"),
            Avg_Attendance=(
                "_attendance",
                "mean"
            ),
            Avg_Activities=(
                "_activities",
                "mean"
            )
        )
        .reset_index()
    )

    data = [
        [
            Paragraph("GRADE", styles["KL"]),
            Paragraph("SECTION", styles["KL"]),
            Paragraph("STUDENTS", styles["KL"]),
            Paragraph("AVG MARKS", styles["KL"]),
            Paragraph("AVG ATT.", styles["KL"]),
            Paragraph("AVG ACTIVITIES", styles["KL"])
        ]
    ]

    for _, r in summary.iterrows():

        data.append(
            [
                str(r["Grade"]),
                str(r["Section"]),
                str(int(r["Students"])),
                f'{r["Avg_Marks"]:.1f}',
                f'{r["Avg_Attendance"]:.1f}%',
                f'{r["Avg_Activities"]:.1f}'
            ]
        )

    st = Table(
        data,
        colWidths=[
            25 * mm,
            30 * mm,
            27 * mm,
            31 * mm,
            31 * mm,
            36 * mm
        ]
    )

    st.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor("#2E8FB3")
                ),

                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white
                ),

                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.HexColor("#C7D5DB")
                ),

                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.35,
                    colors.HexColor("#DCE5E9")
                ),

                (
                    "ALIGN",
                    (2, 1),
                    (-1, -1),
                    "CENTER"
                ),

                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [
                        colors.white,
                        colors.HexColor("#F5FAFC")
                    ]
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                )
            ]
        )
    )

    body += [
        st,
        Spacer(1, 7 * mm)
    ]

    # --------------------------------------------------------
    # REPORT PURPOSE
    # --------------------------------------------------------

    body.append(
        Paragraph(
            "REPORT PURPOSE",
            ParagraphStyle(
                "BH3",
                parent=styles["H"],
                fontSize=10,
                textColor=colors.white
            )
        )
    )

    body.append(
        Paragraph(
            "This consolidated report provides a monthly "
            "view of Robotics student performance using "
            "the supplied student dataset. It can be "
            "shared with the school as a summary of "
            "academic performance, attendance and "
            "activity participation.",
            styles["B"]
        )
    )

    body += [
        Spacer(1, 8 * mm)
    ]

    # --------------------------------------------------------
    # FOOTER
    # --------------------------------------------------------

    footer = Table(
        [
            [
                Paragraph(
                    "Your Growth, Our Priority",
                    styles["F"]
                )
            ]
        ],
        colWidths=[
            180 * mm
        ],
        rowHeights=[
            9 * mm
        ]
    )

    footer.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    colors.HexColor("#164B67")
                )
            ]
        )
    )

    body.append(footer)

    doc.build(body)

    return path


# ============================================================
# MAIN
# ============================================================

def main():

    prepare_robot()

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Input Excel not found: {EXCEL_FILE}"
        )

    print(
        f"Processing Excel file: {EXCEL_FILE.name}"
    )

    df = pd.read_excel(
        EXCEL_FILE
    )

    required = [
        "Student Name",
        "Parent / Guardian",
        "Grade",
        "Section",
        "Parent Phone",
        "Robotics Marks",
        "Attendance %",
        "Activities"
    ]

    missing = [
        c
        for c in required
        if c not in df.columns
    ]

    if missing:

        raise ValueError(
            "Missing Excel columns: "
            + ", ".join(missing)
        )

    # --------------------------------------------------------
    # INDIVIDUAL REPORTS
    # --------------------------------------------------------

    for _, row in df.iterrows():

        make_individual_report(row)

    # --------------------------------------------------------
    # MONTHLY CONSOLIDATED REPORT
    # --------------------------------------------------------

    report_month = get_report_month(
        EXCEL_FILE
    )

    consolidated = make_consolidated_report(
        df,
        report_month
    )

    print(
        f"Generated {len(df)} individual reports in: "
        f"{INDIVIDUAL_OUTPUT}"
    )

    print(
        f"Generated consolidated report: "
        f"{consolidated}"
    )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()