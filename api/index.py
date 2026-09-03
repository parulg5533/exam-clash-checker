"""
Exam Clash Checker - Vercel Serverless API Handler
===================================================
Exports handler for Vercel deployment at /api.
"""

import os
import sys
import json
import io
import urllib.parse
from http.server import BaseHTTPRequestHandler

# Add root directory to sys.path so exam_clash_checker can be imported
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from exam_clash_checker import ExamClashChecker

# Initialize global instance
CHECKER = ExamClashChecker()


class handler(BaseHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed.query)

        # Normalize path from Vercel rewrite (?path=$1) or direct URL
        raw_path = query.get("path", [""])[0]
        if raw_path:
            if not raw_path.startswith("/api/"):
                path = "/api/" + raw_path.lstrip("/")
            else:
                path = raw_path
        else:
            path = parsed.path

        # 1. API: Summary and audit metrics
        if path == "/api/summary":
            audit = CHECKER.audit_timetable()
            data = {
                "total_students": len(CHECKER.student_to_courses),
                "total_courses": len(CHECKER.course_to_students),
                "total_slots": audit.total_slots,
                "scheduled_courses": audit.total_courses_in_timetable,
                "total_clashes": len(audit.clashes),
                "clashes": [
                    {
                        "date": cl.date_str,
                        "slot": cl.slot_name,
                        "course_a": cl.course_a,
                        "course_b": cl.course_b,
                        "overlap_count": cl.overlap_count,
                        "students": cl.students,
                    }
                    for cl in audit.clashes
                ],
                "duplicate_courses": audit.duplicate_courses,
                "unscheduled_courses": audit.unscheduled_courses,
            }
            self.send_json(data)
            return

        # 2. API: All courses list
        if path == "/api/courses":
            courses = []
            for code in sorted(CHECKER.course_to_students.keys()):
                info = CHECKER.course_info[code]
                courses.append({
                    "code": code,
                    "name": info.name,
                    "enrolled": info.enrolled_count,
                })
            self.send_json(courses)
            return

        # 3. API: Single course details & mutual exclusivity
        if path.startswith("/api/course/"):
            course_code = urllib.parse.unquote(path.split("/api/course/")[1]).strip().upper()
            if course_code not in CHECKER.course_to_students:
                self.send_json({"error": f"Course '{course_code}' not found"}, status=404)
                return

            info = CHECKER.course_info[course_code]
            conflicts = CHECKER.get_course_conflicts(course_code)
            safe_slots = CHECKER.find_safe_slots(course_code)
            current_slots = CHECKER.course_to_slots.get(course_code, [])

            slot_status = []
            for s in current_slots:
                chk = CHECKER.check_course_in_slot(course_code, s["id"])
                slot_status.append({
                    "id": s["id"],
                    "date": s["date"],
                    "session": s["session"],
                    "full_name": s["full_name"],
                    "has_clash": chk["has_clash"],
                    "clashing_courses": [c["clashing_course"] for c in chk["clashes"]],
                })

            data = {
                "code": course_code,
                "name": info.name,
                "enrolled": info.enrolled_count,
                "scheduled_slots": slot_status,
                "conflicts": conflicts,
                "safe_slots": safe_slots,
            }
            self.send_json(data)
            return

        # 4. API: Check slot simulation
        if path.startswith("/api/check-slot"):
            course = query.get("course", [""])[0].strip().upper()
            slot_id = query.get("slot", [""])[0].strip()
            if not course or not slot_id:
                self.send_json({"error": "Missing 'course' or 'slot' parameter"}, status=400)
                return
            res = CHECKER.check_course_in_slot(course, slot_id)
            self.send_json(res)
            return

        # 5. API: Full timetable matrix
        if path == "/api/timetable":
            audit = CHECKER.audit_timetable()
            clash_course_set = set()
            for cl in audit.clashes:
                clash_course_set.add(cl.course_a)
                clash_course_set.add(cl.course_b)

            slots_data = []
            for s in CHECKER.slots:
                slot_clashes = [c for c in s["courses"] if c in clash_course_set]
                slots_data.append({
                    "id": s["id"],
                    "date": s["date"],
                    "session": s["session"],
                    "slot_name": s["slot_name"],
                    "full_name": s["full_name"],
                    "courses": s["courses"],
                    "has_clash": any(
                        cl.date_str == s["date"] and cl.slot_name == s["slot_name"]
                        for cl in audit.clashes
                    ),
                    "clashing_courses": slot_clashes,
                })
            self.send_json(slots_data)
            return

        # 6. API: Student schedule lookup
        if path.startswith("/api/student/"):
            roll = urllib.parse.unquote(path.split("/api/student/")[1]).strip().upper()
            res = CHECKER.get_student_schedule(roll)
            self.send_json(res)
            return

        # 7. API: Excel Export
        if path == "/api/export":
            buffer = io.BytesIO()
            CHECKER.export_audit_to_excel(buffer)
            buffer.seek(0)
            excel_bytes = buffer.read()

            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=Exam_Clash_Audit_Report.xlsx")
            self.send_header("Content-Length", str(len(excel_bytes)))
            self.end_headers()
            self.wfile.write(excel_bytes)
            return

        # 404
        self.send_json({"error": "Endpoint not found", "path": path}, status=404)

    def send_json(self, data, status: int = 200):
        body = json.dumps(data, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


# Export for Vercel
app = handler
application = handler
