from pathlib import Path
import re
import pandas as pd
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import mm

BASE = Path(__file__).resolve().parent
ASSETS = BASE / 'assets'
OUTPUT = BASE / 'Student_PDFs'
OUTPUT.mkdir(exist_ok=True)
EXCEL_FILE = BASE / 'Robotics_50_Students_Sample.xlsx'

SCHOOL_LOGO = ASSETS / 'school_logo.jpg'
IKKASHIN_LOGO = ASSETS / 'ikkashin_logo.jpg'
ROBOT_IMAGE = ASSETS / 'robot_image.jpg'

# Crop the uploaded robot once so it fits the same header area as the sample PDF.
ROBOT_CROP = ASSETS / '_robot_crop.jpg'
if not ROBOT_CROP.exists():
    im = Image.open(ROBOT_IMAGE).convert('RGB')
    w, h = im.size
    im.crop((int(w*.08), int(h*.02), int(w*.92), int(h*.98))).save(ROBOT_CROP, quality=96)

styles = getSampleStyleSheet()
def S(name, parent, **kw):
    styles.add(ParagraphStyle(name=name, parent=styles[parent], **kw))
S('T', 'Title', fontName='Helvetica-Bold', fontSize=17, leading=19, alignment=TA_CENTER, textColor=colors.HexColor('#164B67'))
S('ST', 'Normal', fontSize=8.2, leading=10, alignment=TA_CENTER, textColor=colors.HexColor('#61727B'))
S('H', 'Heading2', fontName='Helvetica-Bold', fontSize=9, leading=10, textColor=colors.white)
S('L', 'Normal', fontName='Helvetica-Bold', fontSize=6.7, leading=8, textColor=colors.HexColor('#65757C'))
S('V', 'Normal', fontSize=8.4, leading=10, textColor=colors.HexColor('#263238'))
S('B', 'BodyText', fontSize=7.9, leading=10.2, textColor=colors.HexColor('#303A40'))
S('KN', 'Normal', fontName='Helvetica-Bold', fontSize=14, leading=15, alignment=TA_CENTER, textColor=colors.HexColor('#164B67'))
S('KL', 'Normal', fontName='Helvetica-Bold', fontSize=6.3, leading=7.5, alignment=TA_CENTER, textColor=colors.HexColor('#65757C'))
S('F', 'Normal', fontName='Helvetica-Bold', fontSize=8.5, leading=10, alignment=TA_CENTER, textColor=colors.white)

def safe_filename(s):
    return re.sub(r'[\\/:*?"<>|]+', '_', re.sub(r'\s+', '_', str(s).strip()))

def val(row, col, default=''):
    x = row.get(col, default)
    return default if pd.isna(x) else x

def make_report(row):
    name = str(val(row,'Student Name','Student')).strip()
    parent = str(val(row,'Parent / Guardian','')).strip()
    grade = str(val(row,'Grade','')).strip()
    section = str(val(row,'Section','')).strip()
    phone = str(val(row,'Parent Phone','')).strip()
    marks = float(val(row,'Robotics Marks',0))
    attendance = float(val(row,'Attendance %',0))
    activities = int(float(val(row,'Activities',0)))
    present = max(0, min(14, round(attendance/100*14)))
    absent = 14 - present
    achievement = 'Creative Explorer' if marks >= 80 else 'Active Learner'

    path = OUTPUT / f'{safe_filename(name)}_Robotics_Report.pdf'
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=9*mm, bottomMargin=9*mm)
    story = []

    top = Table([['','','','','','','','']], colWidths=[22.75*mm]*8, rowHeights=[3.5*mm])
    top.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#164B67')),
        ('BACKGROUND',(1,0),(1,0),colors.HexColor('#35A7C9')),
        ('BACKGROUND',(3,0),(3,0),colors.HexColor('#F2B84B')),
        ('BACKGROUND',(5,0),(5,0),colors.HexColor('#D95D8A')),
        ('BACKGROUND',(7,0),(7,0),colors.HexColor('#65B98A')),
    ]))
    story += [top, Spacer(1,2.5*mm)]

    school = RLImage(str(SCHOOL_LOGO), width=22*mm, height=22*mm)
    ikk = RLImage(str(IKKASHIN_LOGO), width=30*mm, height=18*mm)
    robot = RLImage(str(ROBOT_CROP), width=27*mm, height=30*mm)
    session = ParagraphStyle('Session', parent=styles['ST'], fontSize=7.5, textColor=colors.HexColor('#8A5A00'))
    head = Table([[
        school,
        [Paragraph('STEM LEARNING OUTCOMES', styles['T']), Paragraph('ROBOTICS • STUDENT PERFORMANCE REPORT', styles['ST']), Paragraph('Academic Session 2025–2026', session)],
        [ikk, robot]
    ]], colWidths=[27*mm,118*mm,35*mm])
    head.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),('ALIGN',(0,0),(0,0),'LEFT'),('ALIGN',(2,0),(2,0),'RIGHT'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),0)]))
    story += [head, Spacer(1,2*mm)]

    def bar(title, bg, accent):
        t = Table([[Paragraph(title, styles['H']), '']], colWidths=[52*mm,128*mm], rowHeights=[6.5*mm])
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),colors.HexColor(bg)),('LINEBELOW',(0,0),(0,0),2,colors.HexColor(accent)),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),7)]))
        story.append(t)

    bar('STUDENT PROFILE','#2E8FB3','#F2B84B')
    st = Table([
        [Paragraph('STUDENT',styles['L']),Paragraph(name,styles['V']),Paragraph('PARENT / GUARDIAN',styles['L']),Paragraph(parent,styles['V'])],
        [Paragraph('GRADE',styles['L']),Paragraph(grade,styles['V']),Paragraph('SECTION',styles['L']),Paragraph(section,styles['V'])],
        [Paragraph('PARENT PHONE',styles['L']),Paragraph(phone,styles['V']),Paragraph('SUBJECT',styles['L']),Paragraph('Robotics',styles['V'])]
    ], colWidths=[28*mm,52*mm,36*mm,64*mm])
    st.setStyle(TableStyle([('BOX',(0,0),(-1,-1),.55,colors.HexColor('#C6D7DE')),('INNERGRID',(0,0),(-1,-1),.35,colors.HexColor('#DCE7EB')),('BACKGROUND',(0,0),(0,-1),colors.HexColor('#EFF8FC')),('BACKGROUND',(2,0),(2,-1),colors.HexColor('#EFF8FC')),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
    story += [st, Spacer(1,3*mm)]

    bar('PERFORMANCE SNAPSHOT','#3C5B9A','#F2B84B')
    k = Table([
        [Paragraph(f'{marks:g} / 100',styles['KN']),Paragraph(f'{attendance:g}%',styles['KN']),Paragraph(str(activities),styles['KN']),Paragraph(f'{present} / 14',styles['KN'])],
        [Paragraph('ROBOTICS SCORE',styles['KL']),Paragraph('ATTENDANCE',styles['KL']),Paragraph('ACTIVITIES',styles['KL']),Paragraph('PRESENT DAYS',styles['KL'])]
    ], colWidths=[45*mm]*4, rowHeights=[13*mm,7*mm])
    k.setStyle(TableStyle([('BOX',(0,0),(-1,-1),.6,colors.HexColor('#C4D3DB')),('INNERGRID',(0,0),(-1,-1),.4,colors.HexColor('#D8E3E8')),('BACKGROUND',(0,0),(0,-1),colors.HexColor('#E5F5FA')),('BACKGROUND',(1,0),(1,-1),colors.HexColor('#EDE9FB')),('BACKGROUND',(2,0),(2,-1),colors.HexColor('#FFF4D9')),('BACKGROUND',(3,0),(3,-1),colors.HexColor('#E6F5EC'))]))
    story += [k, Spacer(1,3*mm)]

    aa = Table([[
        Paragraph(f"<b>🏆 ACHIEVEMENT</b><br/><font size=10 color='#8A5A00'>{achievement}</font>",styles['B']),
        Paragraph(f'<b>◉ ATTENDANCE</b><br/>Present {present} • Absent {absent} • Total 14',styles['B'])
    ]], colWidths=[90*mm,90*mm])
    aa.setStyle(TableStyle([('BOX',(0,0),(-1,-1),.55,colors.HexColor('#C7D5DB')),('INNERGRID',(0,0),(-1,-1),.35,colors.HexColor('#DCE5E9')),('BACKGROUND',(0,0),(0,0),colors.HexColor('#FFF2C9')),('BACKGROUND',(1,0),(1,0),colors.HexColor('#DFF2FA')),('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
    story += [aa, Spacer(1,3*mm)]

    left = Table([[Paragraph('LEARNING HIGHLIGHTS',styles['H'])],[Paragraph('✓ Basic robotics concepts<br/>✓ Component identification<br/>✓ Hands-on model building<br/>✓ Team participation<br/>✓ Practical problem solving',styles['B'])]], colWidths=[88*mm], rowHeights=[7*mm,35*mm])
    right = Table([[Paragraph('TEACHER OBSERVATION',styles['H'])],[Paragraph(f'{name} participates well in practical activities and follows instructions carefully. The student is developing confidence in building and testing simple robotic models and works positively with classmates.',styles['B'])]], colWidths=[88*mm], rowHeights=[7*mm,35*mm])
    for t,bg in [(left,'#FFF4D9'),(right,'#E5F5FA')]:
        t.setStyle(TableStyle([('BOX',(0,0),(-1,-1),.55,colors.HexColor('#C7D5DB')),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#2E8FB3')),('BACKGROUND',(0,1),(-1,1),colors.HexColor(bg)),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
    main = Table([[left,right]], colWidths=[90*mm,90*mm])
    main.setStyle(TableStyle([('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0)]))
    story += [main, Spacer(1,3*mm)]

    q = Table([[Paragraph('⚙ BUILD • EXPLORE • INNOVATE ⚙',ParagraphStyle('Q',parent=styles['H'],alignment=TA_CENTER,fontSize=10))]], colWidths=[180*mm], rowHeights=[9*mm])
    q.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#D95D8A'))]))
    story += [q, Spacer(1,3*mm)]

    bar('AUTHORIZATION','#3C5B9A','#D95D8A')
    sig = Table([[Paragraph('CLASS TEACHER SIGNATURE',styles['L']),Paragraph('PRINCIPAL SIGNATURE',styles['L'])],['','']], colWidths=[90*mm,90*mm], rowHeights=[7*mm,15*mm])
    sig.setStyle(TableStyle([('BOX',(0,0),(-1,-1),.55,colors.HexColor('#C7D5DB')),('INNERGRID',(0,0),(-1,-1),.35,colors.HexColor('#DCE5E9')),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#EEF3FA')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),7)]))
    story += [sig, Spacer(1,3*mm)]

    f = Table([[Paragraph('✦ Your Growth, Our Priority ✦',styles['F'])]], colWidths=[180*mm], rowHeights=[8*mm])
    f.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#164B67'))]))
    story.append(f)
    doc.build(story)

# Required Excel columns:
# Student Name, Parent / Guardian, Grade, Section, Parent Phone,
# Robotics Marks, Attendance %, Activities

df = pd.read_excel(EXCEL_FILE)
required = ['Student Name','Parent / Guardian','Grade','Section','Parent Phone','Robotics Marks','Attendance %','Activities']
missing = [c for c in required if c not in df.columns]
if missing:
    raise ValueError('Missing Excel columns: ' + ', '.join(missing))
for _, row in df.iterrows():
    make_report(row)
print(f'Generated {len(df)} PDF reports in: {OUTPUT}')
