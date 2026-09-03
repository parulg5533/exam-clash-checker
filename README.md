# Exam Clash Checker & Timetable Auditor 🎓

An automated, high-performance examination clash detection and scheduling auditor built for university course registrations. 

Designed to process student-course bipartite registration data and cross-examine master examination timetables to guarantee that **no student is double-booked**.

---

## Key Features

- **Conflict Graph & Bipartite Mapping**: Fast precomputed mutual exclusivity matrix mapping every pair of courses sharing enrolled students.
- **Timetable Sanity & Clash Audit**:
  - Direct slot clashes (multiple exams scheduled for a student at the exact same time).
  - Multi-slot duplicate courses.
  - Same-day exam fatigue warnings (students with both morning and evening exams).
  - Unscheduled course detection.
- **Conflict-Free Slot Recommender**: Given any course code, automatically identifies all conflict-free exam slots where it can be scheduled or moved.
- **Student Schedule Lookup**: Look up any student by roll number to inspect their personalized datesheet and clash status.
- **Interactive Web Dashboard**:
  - Live course search and conflict inspector.
  - Interactive **Slot Placement Simulator** to test moving courses in real-time.
  - Visual 18-slot timetable matrix with clash alerts.
- **CLI & Excel Reporting**:
  - Beautiful formatted terminal interface powered by `rich`.
  - 1-click export to a multi-tab Excel audit report.

---

## Project Structure

```
├── exam_clash_checker.py    # Core algorithmic engine & graph data structures
├── cli.py                   # Terminal interface (audit, check, safe-slots, student, export)
├── app.py                   # Local web application server (zero external web framework required)
├── web/
│   └── index.html           # Modern glassmorphism web dashboard UI
├── test_checker.py          # Automated unit test suite
├── requirements.txt         # Dependencies
├── Autumn 2026 Registration Data as on 31-8-2026.xlsx  # Registration dataset
├── Tentative TimeTable Autumn 2026 IITP v1.xlsx        # Timetable schedule
└── Exam_Clash_Audit_Report.xlsx                        # Generated audit report
```

---

## Installation & Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/parulg5533/exam-clash-checker.git
   cd exam-clash-checker
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run unit tests**:
   ```bash
   python test_checker.py
   ```

---

## Command Line Interface (CLI)

```bash
# 1. Run a full timetable clash audit
python cli.py audit

# 2. Check conflicts for a specific course
python cli.py check HS4118
python cli.py check CE3101

# 3. Find conflict-free slots to safely schedule a course
python cli.py safe-slots HS4118

# 4. View an individual student's datesheet and clash status
python cli.py student 2301MM04

# 5. Export comprehensive audit report to Excel
python cli.py export Exam_Clash_Audit_Report.xlsx
```

---

## Interactive Web Dashboard

Launch the local web dashboard:
```bash
python app.py
```
Open **[http://localhost:8050](http://localhost:8050)** in your browser.

- **Course Clash Inspector**: Search any of the 266 courses and inspect all conflicting courses and student counts.
- **Slot Placement Simulator**: Select any exam slot from the dropdown to test moving a course and immediately see if clashes occur.
- **Timetable Matrix Grid**: Visual grid of all slots across 9 days with color-coded clash tags.
- **Student Schedule Lookup**: Search any student roll number to view their personalized datesheet.
- **Export Report**: Download the Excel spreadsheet directly from the dashboard.

---

## License
MIT License
