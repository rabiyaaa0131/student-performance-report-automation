# Student Performance Report Automation

## Overview

This project automates the generation of Robotics student performance reports using Python.

The system reads student data from Excel files and generates:

- Individual student performance reports
- One consolidated monthly performance report

## How It Works

1. Excel files are placed in the `Input_Data` folder.
2. The automation identifies the next unprocessed monthly data file.
3. Python reads and validates the student data.
4. Individual PDF reports are generated for each student.
5. A consolidated monthly PDF report is generated.
6. Processed files are recorded so the same month is not processed again.

## Project Structure

```text
student-performance-report-automation/
│
├── generate_student_reports.py
├── watcher.py
├── processed_files.json
├── requirements.txt
├── README.md
│
├── assets/
│   ├── school_logo.jpg
│   ├── ikkashin_logo.jpg
│   └── robot_image.jpg
│
├── Input_Data/
├── Individual_Reports/
└── Consolidated_Reports/
