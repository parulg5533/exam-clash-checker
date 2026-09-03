"""
Exam Clash Checker - Command Line Interface (CLI)
=================================================
Usage:
    python cli.py audit
    python cli.py check <COURSE_CODE>
    python cli.py safe-slots <COURSE_CODE>
    python cli.py test-slot <COURSE_CODE> <SLOT_NUMBER>
    python cli.py student <ROLL_NO>
    python cli.py export [OUTPUT_PATH.xlsx]
"""

import sys
import argparse
import io

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from exam_clash_checker import ExamClashChecker

try:
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel
    from rich.text import Text
    USE_RICH = True
    console = Console(force_terminal=True)
except ImportError:
    USE_RICH = False
    console = None


def print_banner():
    msg = """
+------------------------------------------------------------------+
|                   EXAM CLASH CHECKER & AUDITOR                   |
|          IIT Patna - Autumn 2026 Examination Scheduling          |
+------------------------------------------------------------------+
"""
    if USE_RICH:
        console.print(f"[bold cyan]{msg}[/bold cyan]")
    else:
        print(msg)


def handle_audit(checker: ExamClashChecker, export: bool = False):
    audit = checker.audit_timetable()

    if USE_RICH:
        table = Table(title="Examination Timetable Audit Summary", style="cyan")
        table.add_column("Metric", style="bold white")
        table.add_column("Value", style="bold yellow")

        table.add_row("Total Registered Students", str(len(checker.student_to_courses)))
        table.add_row("Total Registered Courses", str(len(checker.course_to_students)))
        table.add_row("Total Exam Slots", str(audit.total_slots))
        table.add_row("Courses Scheduled in Timetable", str(audit.total_courses_in_timetable))
        table.add_row("Direct Slot Clashes Found", f"[bold red]{len(audit.clashes)}[/bold red]" if audit.clashes else "[bold green]0[/bold green]")
        table.add_row("Students Double-Booked in Slot", f"[bold red]{len(audit.students_with_clashes)}[/bold red]" if audit.students_with_clashes else "[bold green]0[/bold green]")
        table.add_row("Courses in Multiple Slots", str(len(audit.duplicate_courses)))
        table.add_row("Unscheduled Courses", str(len(audit.unscheduled_courses)))
        table.add_row("Same-Day 2-Exam Warnings", str(len(audit.same_day_warnings)))
        console.print(table)

        if audit.clashes:
            console.print("\n[bold red]⚠️ DIRECT SLOT CLASHES DETECTED:[/bold red]")
            c_table = Table(style="red")
            c_table.add_column("Date", style="cyan")
            c_table.add_column("Slot", style="white")
            c_table.add_column("Course 1", style="bold magenta")
            c_table.add_column("Course 2", style="bold magenta")
            c_table.add_column("Overlapping Students", style="yellow")
            c_table.add_column("Roll Numbers", style="green")

            for cl in audit.clashes:
                c_table.add_row(
                    cl.date_str,
                    cl.slot_name,
                    cl.course_a,
                    cl.course_b,
                    str(cl.overlap_count),
                    ", ".join(cl.students)
                )
            console.print(c_table)

        if audit.duplicate_courses:
            console.print("\n[bold yellow]ℹ️ COURSES SCHEDULED IN MULTIPLE SLOTS:[/bold yellow]")
            for c, sls in audit.duplicate_courses.items():
                console.print(f"  • [bold cyan]{c}[/bold cyan]: { ' | '.join(sls) }")

    else:
        print("--- AUDIT SUMMARY ---")
        print(f"Total Registered Students: {len(checker.student_to_courses)}")
        print(f"Total Registered Courses: {len(checker.course_to_students)}")
        print(f"Direct Slot Clashes Found: {len(audit.clashes)}")
        for cl in audit.clashes:
            print(f"  [CLASH] {cl.date_str} {cl.slot_name}: {cl.course_a} & {cl.course_b} share {cl.overlap_count} students ({', '.join(cl.students)})")
        if audit.duplicate_courses:
            print("Courses in multiple slots:")
            for c, sls in audit.duplicate_courses.items():
                print(f"  {c}: {' | '.join(sls)}")

    if export:
        out_path = "Exam_Clash_Audit_Report.xlsx"
        checker.export_audit_to_excel(out_path)
        if USE_RICH:
            console.print(f"\n[bold green]✓ Full audit report exported to:[/bold green] [bold underline]{out_path}[/bold underline]")
        else:
            print(f"Audit report exported to: {out_path}")


def handle_check(checker: ExamClashChecker, course_code: str):
    course_code = course_code.strip().upper()
    if course_code not in checker.course_to_students:
        print(f"Error: Course '{course_code}' not found in registration data.")
        return

    c_info = checker.course_info[course_code]
    conflicts = checker.get_course_conflicts(course_code)
    slots = checker.course_to_slots.get(course_code, [])

    if USE_RICH:
        console.print(Panel(
            f"[bold white]Course Code:[/bold white] [bold cyan]{course_code}[/bold cyan]\n"
            f"[bold white]Course Name:[/bold white] {c_info.name}\n"
            f"[bold white]Total Enrolled Students:[/bold white] [bold yellow]{c_info.enrolled_count}[/bold yellow]\n"
            f"[bold white]Currently Scheduled in:[/bold white] " + (", ".join(s['full_name'] for s in slots) if slots else "[red]Unscheduled[/red]"),
            title=f"Course Details: {course_code}",
            style="cyan"
        ))

        # Check current slot clashes
        for s in slots:
            chk = checker.check_course_in_slot(course_code, s["id"])
            if chk["has_clash"]:
                console.print(f"[bold red]❌ CLASH DETECTED in slot '{s['full_name']}'![/bold red]")
                for cl in chk["clashes"]:
                    console.print(f"   Clashes with [bold magenta]{cl['clashing_course']}[/bold magenta] ({cl['course_name']}) - [yellow]{cl['overlap_count']} students[/yellow]: {', '.join(cl['students'])}")
            else:
                console.print(f"[bold green]✓ No clashes in current slot '{s['full_name']}'[/bold green]")

        # Conflict list
        if conflicts:
            table = Table(title=f"Mutually Exclusive Courses for {course_code} (Cannot share exam slot)", style="magenta")
            table.add_column("Conflicting Course", style="bold cyan")
            table.add_column("Course Name", style="white")
            table.add_column("Overlapping Students", style="yellow")
            table.add_column("Sample Roll Numbers", style="green")

            for cf in conflicts[:15]:
                table.add_row(
                    cf["course_code"],
                    cf["course_name"],
                    str(cf["overlap_count"]),
                    ", ".join(cf["overlapping_students"][:5]) + ("..." if len(cf["overlapping_students"]) > 5 else "")
                )
            console.print(table)
            if len(conflicts) > 15:
                console.print(f"[italic]... and {len(conflicts) - 15} more conflicting courses.[/italic]")
        else:
            console.print("[bold green]This course has no student overlaps with any other course![/bold green]")
    else:
        print(f"Course: {course_code} - {c_info.name}")
        print(f"Enrolled: {c_info.enrolled_count}")
        print(f"Total conflicting courses: {len(conflicts)}")


def handle_safe_slots(checker: ExamClashChecker, course_code: str):
    course_code = course_code.strip().upper()
    if course_code not in checker.course_to_students:
        print(f"Error: Course '{course_code}' not found in registration data.")
        return

    res = checker.find_safe_slots(course_code)
    if USE_RICH:
        console.print(f"\n[bold]Slot Availability Analysis for [cyan]{course_code}[/cyan] ({res['course_name']}):[/bold]")
        console.print(f"[bold green]Safe Slots:[/bold green] {res['safe_slots_count']} / {res['total_slots']}")
        console.print(f"[bold red]Clashing Slots:[/bold red] {res['clashing_slots_count']} / {res['total_slots']}\n")

        table = Table(title=f"Available Conflict-Free Exam Slots for {course_code}", style="green")
        table.add_column("#", style="dim")
        table.add_column("Slot ID", style="cyan")
        table.add_column("Date", style="white")
        table.add_column("Session", style="bold yellow")
        table.add_column("Other Courses in Slot", style="magenta")

        for idx, s in enumerate(res["safe_slots"], start=1):
            table.add_row(
                str(idx),
                s["slot_id"],
                s["date"],
                s["session"],
                f"{s['current_course_count']} courses scheduled"
            )
        console.print(table)
    else:
        print(f"Safe slots count: {res['safe_slots_count']}")
        for s in res["safe_slots"]:
            print(f"  [SAFE] {s['full_name']}")


def handle_student(checker: ExamClashChecker, roll_no: str):
    roll_no = roll_no.strip().upper()
    res = checker.get_student_schedule(roll_no)
    if "error" in res:
        print(res["error"])
        return

    if USE_RICH:
        console.print(Panel(
            f"[bold white]Roll Number:[/bold white] [bold cyan]{roll_no}[/bold cyan]\n"
            f"[bold white]Enrolled Courses ({res['total_enrolled']}):[/bold white] {', '.join(res['enrolled_courses'])}\n"
            f"[bold white]Clash Status:[/bold white] " + ("[bold red]❌ DOUBLE-BOOKED![/bold red]" if res["has_clashes"] else "[bold green]✓ Clear (No clashes)[/bold green]"),
            title=f"Student Exam Schedule: {roll_no}",
            style="cyan"
        ))

        table = Table(title=f"Exam Datesheet for {roll_no}", style="blue")
        table.add_column("Course Code", style="bold cyan")
        table.add_column("Course Name", style="white")
        table.add_column("Date", style="yellow")
        table.add_column("Session", style="magenta")
        table.add_column("Status", style="green")

        for item in res["schedule"]:
            table.add_row(
                item["course_code"],
                item["course_name"],
                item.get("date") or "-",
                item.get("session") or "-",
                item["status"]
            )
        console.print(table)

        if res["has_clashes"]:
            console.print("\n[bold red]⚠️ CLASH DETAILS:[/bold red]")
            for cl in res["clashes"]:
                console.print(f"  • {cl['message']}")

        if res["same_day_exams"]:
            console.print("\n[bold yellow]⚠️ SAME-DAY EXAMS (Morning + Evening):[/bold yellow]")
            for sd in res["same_day_exams"]:
                console.print(f"  • Date {sd['date']}: Morning ({', '.join(sd['morning'])}) & Evening ({', '.join(sd['evening'])})")
    else:
        print(f"Student: {roll_no}, Enrolled: {res['enrolled_courses']}")
        print(f"Has clashes: {res['has_clashes']}")


def main():
    parser = argparse.ArgumentParser(description="Exam Clash Checker & Auditor")
    subparsers = parser.add_subparsers(dest="command", help="Commands")

    # audit
    p_audit = subparsers.add_parser("audit", help="Run full timetable audit")
    p_audit.add_argument("--export", action="store_true", help="Also export report to Excel")

    # check
    p_check = subparsers.add_parser("check", help="Check conflicts for a specific course")
    p_check.add_argument("course_code", help="Course code (e.g. CE3101, HS4118)")

    # safe-slots
    p_safe = subparsers.add_parser("safe-slots", help="Find all conflict-free slots for a course")
    p_safe.add_argument("course_code", help="Course code (e.g. HS4118)")

    # student
    p_stu = subparsers.add_parser("student", help="View student's exam datesheet and clashes")
    p_stu.add_argument("roll_no", help="Student roll number (e.g. 2301MM04)")

    # export
    p_exp = subparsers.add_parser("export", help="Export audit report to Excel file")
    p_exp.add_argument("output", nargs="?", default="Exam_Clash_Audit_Report.xlsx", help="Output filepath")

    args = parser.parse_args()
    print_banner()

    checker = ExamClashChecker()

    if args.command == "audit":
        handle_audit(checker, export=args.export)
    elif args.command == "check":
        handle_check(checker, args.course_code)
    elif args.command == "safe-slots":
        handle_safe_slots(checker, args.course_code)
    elif args.command == "student":
        handle_student(checker, args.roll_no)
    elif args.command == "export":
        out = checker.export_audit_to_excel(args.output)
        if USE_RICH:
            console.print(f"[bold green]✓ Successfully exported audit report to: {out}[/bold green]")
        else:
            print(f"Successfully exported audit report to: {out}")
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
