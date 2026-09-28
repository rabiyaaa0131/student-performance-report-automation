VS CODE STEPS

1. Open this folder in VS Code.
2. Terminal > New Terminal.
3. Create virtual environment:
   py -m venv .venv
4. Activate it:
   .venv\Scripts\activate
5. Install packages:
   pip install -r requirements.txt
6. Run:
   python generate_student_reports.py
7. Open Student_PDFs. One PDF is created for every Excel row.

For your real Excel, keep these column names exactly:
Student Name
Parent / Guardian
Grade
Section
Parent Phone
Robotics Marks
Attendance %
Activities

The exact Social Baluni logo, IKKASHIN logo and uploaded robot image are already in assets.
