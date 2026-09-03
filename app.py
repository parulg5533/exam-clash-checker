"""
Exam Clash Checker - Web Application & Serverless Handler
=========================================================
Exports top-level "app", "application", and "handler" for Vercel deployment,
while remaining runnable locally with "python app.py".
"""

import os
import sys
import json
import io
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
from exam_clash_checker import ExamClashChecker

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Initialize global clash checker instance
CHECKER = ExamClashChecker()


class handler(BaseHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed.query)
        path = (
            query.get("path", [""])[0]
            or self.headers.get("x-forwarded-uri", "").split("?")[0]
            or self.headers.get("x-matched-path", "").split("?")[0]
            or parsed.path
        ).strip()

        if path in ("/api/export-draft", "api/export-draft"):
            try:
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length) if content_length > 0 else b"{}"
                payload = json.loads(body.decode("utf-8")) if body else {}
                draft_slots = payload.get("draft_slots", [])
                moves_log = payload.get("moves", [])

                buffer = io.BytesIO()
                CHECKER.export_draft_to_excel(draft_slots, moves_log, buffer)
                buffer.seek(0)
                excel_bytes = buffer.read()

                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", "attachment; filename=Draft_Exam_Timetable_Report.xlsx")
                self.send_header("Content-Length", str(len(excel_bytes)))
                self.end_headers()
                self.wfile.write(excel_bytes)
                return
            except Exception as e:
                self.send_json({"error": f"Failed to export draft: {str(e)}"}, status=500)
                return

        self.send_json({"error": "Endpoint not found", "requested_path": path}, status=404)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed.query)

        # Robust path resolution across Vercel serverless rewrites and local HTTP server
        path = (
            query.get("path", [""])[0]
            or self.headers.get("x-forwarded-uri", "").split("?")[0]
            or self.headers.get("x-matched-path", "").split("?")[0]
            or parsed.path
        ).strip()

        # Static file: Main Web UI (Root, index.html, or Vercel default function paths)
        if path in ["", "/", "/index.html", "/app", "/app.py", "/web", "/web/index.html"]:
            html_candidates = [
                os.path.join(BASE_DIR, "public", "index.html"),
                os.path.join(BASE_DIR, "web", "index.html"),
            ]
            for html_file in html_candidates:
                if os.path.exists(html_file):
                    self.serve_static_file(html_file, "text/html; charset=utf-8")
                    return
            self.send_json({"error": "UI index.html not found on server"}, status=500)
            return

        # API: Summary and audit metrics
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

        # API: All courses list
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

        # API: Single course details & mutual exclusivity
        if path.startswith("/api/course/"):
            course_code = urllib.parse.unquote(path.split("/api/course/")[1]).strip().upper()
            if course_code not in CHECKER.course_to_students:
                self.send_json({"error": f"Course '{course_code}' not found"}, status=404)
                return

            info = CHECKER.course_info[course_code]
            conflicts = CHECKER.get_course_conflicts(course_code)
            safe_slots = CHECKER.find_safe_slots(course_code)
            current_slots = CHECKER.course_to_slots.get(course_code, [])

            # Check clashes in current slots
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

        # API: Check slot simulation
        if path.startswith("/api/check-slot"):
            course = query.get("course", [""])[0].strip().upper()
            slot_id = query.get("slot", [""])[0].strip()
            if not course or not slot_id:
                self.send_json({"error": "Missing 'course' or 'slot' parameter"}, status=400)
                return
            res = CHECKER.check_course_in_slot(course, slot_id)
            self.send_json(res)
            return

        # API: Full timetable matrix
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

        # API: Student schedule lookup
        if path.startswith("/api/student/"):
            roll = urllib.parse.unquote(path.split("/api/student/")[1]).strip().upper()
            res = CHECKER.get_student_schedule(roll)
            self.send_json(res)
            return

        # API: In-memory Excel Export (compatible with Vercel serverless read-only filesystem)
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

        # 404 handler with diagnostic path info
        self.send_json({"error": "Endpoint not found", "requested_path": path}, status=404)

    def serve_static_file(self, filepath: str, content_type: str):
        if not os.path.exists(filepath):
            self.send_json({"error": "File not found", "path": filepath}, status=404)
            return
        with open(filepath, "rb") as f:
            content = f.read()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def send_json(self, data, status: int = 200):
        body = json.dumps(data, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


# Export top-level variables expected by Vercel Python runtime
app = handler
application = handler
ClashCheckerRequestHandler = handler


def run_server(port: int = 8050):
    server_address = ("", port)
    try:
        httpd = HTTPServer(server_address, handler)
        print(f"\n===========================================================")
        print(f"  Exam Clash Checker Dashboard is live!")
        print(f"  URL: http://localhost:{port}")
        print(f"  Press Ctrl+C to stop the server")
        print(f"===========================================================\n")
        httpd.serve_forever()
    except OSError as e:
        if port < 8060:
            run_server(port + 1)
        else:
            raise e


if __name__ == "__main__":
    port = 8050
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            pass
    run_server(port)
