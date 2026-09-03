# IIT Patna - Exam Clash Checker & Timetable Auditor 🎓

An automated, high-performance examination clash detection, mutual exclusivity mapping, and schedule auditing system built for collegiate university course registrations.

Designed to process bipartite student-course registration records and cross-examine master examination schedules to guarantee that **no student is double-booked**. Featuring an interactive **Draft Timetable Builder** with live slot student headcount tracking, real-time clash auditing, and color-coded multi-sheet Excel exports.

Styled with the official **IIT Patna Academic Portal Design System** (inspired by [IIT Patna ACC Portal](https://acc.iitp.ac.in/)), featuring official IIT Patna Deep Navy (`#0B1E3F`), Royal Blue (`#1D4ED8`), crisp institutional typography (`Plus Jakarta Sans`, `Inter`), and high-contrast academic data cards.

---

## 🏛️ Official Academic Portal UI

The web dashboard is styled to match IIT Patna's official academic design language:
- **Header Bar**: Deep Institutional Navy (`#0B1E3F`) with Royal Blue accent borders (`#1D4ED8`), official badge, and live Google Sheet status.
- **Canvas & Cards**: Clean slate background (`#F8FAFC`) with crisp elevated cards (`#FFFFFF`), subtle slate borders (`#E2E8F0`), and soft academic shadows.
- **Typography**: Dual-font hierarchy using Google Fonts **Plus Jakarta Sans** for headings, **Inter** for UI elements, and **JetBrains Mono** for course codes and student roll numbers.
- **Color-Coded Status Badges**:
  - `Morning Session`: Sky Blue (`#0369A1` / `#E0F2FE`)
  - `Evening Session`: Purple (`#7E22CE` / `#F3E8FF`)
  - `Clashes & Conflicts`: Crimson Red (`#DC2626` / `#FEF2F2`)
  - `Draft Moves`: Emerald Green (`#059669` / `#ECFDF5`)
  - `Conflict-Free Safe Slots`: Mint Green (`#166534` / `#F0FDF4`)

---

## ⚡ Key Capabilities

### 1. Conflict Graph & Mutual Exclusivity Matrix
- Fast precomputed bipartite mapping between **3,512 students** and **266 distinct courses**.
- O(1) pairwise lookup of shared enrollments between any two courses.

### 2. Comprehensive Timetable Clash Audit
- **Direct Slot Clashes**: Detects whether any student has two or more exams in the exact same exam slot.
- **Multi-Slot Duplicate Assignments**: Flags courses mistakenly scheduled in multiple slots.
- **Exam Fatigue Warnings**: Flags students having both Morning and Evening exams on the same calendar day.
- **Unscheduled Course Detection**: Identifies registered courses that have not been assigned an examination slot.

### 3. Interactive Draft Timetable Builder ("What-If" Planner)
- **Unlimited Incremental Moves**: Academic administrators can move any course to another slot, then move another course, creating a complete **Draft Timetable**.
- **Live Slot Student Headcount Tracking**: Calculates the exact unique seating headcount in every slot in the draft (e.g. `👥 486 stds (+66)`).
- **Instant Move Preview**: Reassigner dropdown displays predicted headcount changes and real-time conflict status before applying changes.
- **1-Click Move from Inspector**: Inspect any course in the Course Clash Inspector and click `+ Move to Draft` next to any safe slot.
- **Changelog & Undo / Revert Controls**: Complete changelog table with per-course `Revert` buttons and a 1-click `Undo Move` button.

### 4. Color-Coded Multi-Sheet Excel Export (`.xlsx`)
Export the final draft timetable directly to an executive-ready Excel workbook formatted with `openpyxl`:
- **Sheet 1: Draft Timetable**: Full 18-slot matrix with **Slot Student Headcounts**, net load change, scheduled courses, and clash status. **Moved courses are highlighted in soft mint green (`#D1FAE5`)** with bold text and tagged `* (MOVED)`. Any remaining clashes are highlighted in soft red (`#FEE2E2`).
- **Sheet 2: Moved Courses (Changelog)**: Audit table of all course movements with before/after slots, enrolled counts, and target headcount.
- **Sheet 3: Slot Headcounts & Capacity**: Capacity load comparison (Original vs. Draft Headcount, load delta $+/-$, and moved courses per slot).
- **Sheet 4: Draft Clash Audit**: Complete list of remaining clashes, or a celebratory green banner if the draft is 100% clash-free.

### 5. Live Google Sheets Synchronization
- Directly connected to the master university timetable Google Sheet via Google Visualization API (`gviz/tq`).
- 1-click `Sync Live Sheet` button to fetch real-time updates made by scheduling officers.

### 6. Student Schedule Lookup
- Search any student roll number (e.g. `2301MM04`, `2301CS01`) to view their personalized exam datesheet, scheduled timings, and conflict alerts.

---

## 📁 Repository Structure

```
├── exam_clash_checker.py    # Core clash engine, conflict graph, & Excel generator
├── app.py                   # Local HTTP server (zero external web framework needed)
├── cli.py                   # Terminal CLI interface (rich-powered)
├── api/
│   └── index.py             # Vercel serverless function handler (/api/export-draft)
├── web/
│   └── index.html           # Official IIT Patna Academic Theme web dashboard
├── public/
│   ├── index.html           # Production mirror of web dashboard
│   └── data.json            # Precompiled student enrollment & course database
├── vercel.json              # Vercel serverless deployment configuration
├── test_checker.py          # Automated unit test suite
├── requirements.txt         # Dependencies (pandas, openpyxl, rich)
├── Autumn 2026 Registration Data as on 31-8-2026.xlsx  # Registration dataset
└── Tentative TimeTable Autumn 2026 IITP v1.xlsx        # Timetable schedule
```

---

## 🚀 Quickstart & Installation

### 1. Clone & Install Dependencies
```bash
git clone https://github.com/parulg5533/exam-clash-checker.git
cd exam-clash-checker
pip install -r requirements.txt
```

### 2. Run Test Suite
```bash
python test_checker.py
```

### 3. Launch Local Web Dashboard
```bash
python app.py
```
Open **[http://localhost:8050](http://localhost:8050)** in your browser.

---

## 💻 Command Line Interface (CLI)

The CLI provides terminal-based auditing powered by `rich`:

```bash
# 1. Run a comprehensive timetable clash audit
python cli.py audit

# 2. Check conflicts for a specific course
python cli.py check HS4118
python cli.py check CE3101

# 3. Find conflict-free exam slots to safely schedule a course
python cli.py safe-slots HS4118

# 4. View an individual student's datesheet and clash status
python cli.py student 2301MM04

# 5. Export comprehensive audit report to Excel
python cli.py export Exam_Clash_Audit_Report.xlsx
```

---

## 🌐 Deployment (Vercel)

The application is configured for deployment on **Vercel**:
- Static web assets are served from `public/`.
- Serverless backend endpoints (`/api/export-draft` and `/api/audit`) are powered by Python 3.12 via `api/index.py`.
- Configured via `vercel.json`.

---

## 📄 License
MIT License - Open for academic scheduling and institutional use.
