"""
Unit and functional tests for ExamClashChecker
"""

import os
import unittest
from exam_clash_checker import ExamClashChecker


class TestExamClashChecker(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.checker = ExamClashChecker(
            reg_data_path="Autumn 2026 Registration Data as on 31-8-2026.xlsx",
            timetable_path="Tentative TimeTable Autumn 2026 IITP v1.xlsx"
        )

    def test_registration_data_loaded(self):
        self.assertEqual(len(self.checker.student_to_courses), 3512)
        self.assertEqual(len(self.checker.course_to_students), 266)
        self.assertIn("HS4118", self.checker.course_info)
        self.assertIn("HS4119", self.checker.course_info)

    def test_known_conflict_detection(self):
        # HS4118 and HS4119 share 2 students: 2301MM04, 2301ME42
        overlap = self.checker.conflict_graph["HS4118"].get("HS4119", set())
        self.assertEqual(len(overlap), 2)
        self.assertIn("2301MM04", overlap)
        self.assertIn("2301ME42", overlap)

    def test_timetable_audit(self):
        audit = self.checker.audit_timetable()
        self.assertEqual(audit.total_slots, 18)
        self.assertGreaterEqual(len(audit.clashes), 1)
        
        # Verify the specific clash
        clash_pairs = [(c.course_a, c.course_b) for c in audit.clashes]
        self.assertTrue(
            ("HS4118", "HS4119") in clash_pairs or ("HS4119", "HS4118") in clash_pairs
        )
        # Duplicate scheduled courses
        self.assertIn("CH2104", audit.duplicate_courses)
        self.assertIn("CB2104", audit.duplicate_courses)

    def test_find_safe_slots(self):
        # For HS4118, safe slots should be found
        res = self.checker.find_safe_slots("HS4118")
        self.assertGreater(res["safe_slots_count"], 0)
        self.assertGreater(res["clashing_slots_count"], 0)
        # The slot with HS4119 should be among clashing slots
        clashing_slot_names = [s["full_name"] for s in res["clashing_slots"]]
        self.assertTrue(any("2026-09-26 Evening" in name for name in clashing_slot_names))

    def test_student_schedule(self):
        # Student 2301MM04 is enrolled in both HS4118 and HS4119
        res = self.checker.get_student_schedule("2301MM04")
        self.assertTrue(res["has_clashes"])
        self.assertEqual(len(res["clashes"]), 1)
        self.assertIn("HS4118", res["clashes"][0]["courses"])
        self.assertIn("HS4119", res["clashes"][0]["courses"])

    def test_excel_export(self):
        out_file = "test_audit_output.xlsx"
        try:
            self.checker.export_audit_to_excel(out_file)
            self.assertTrue(os.path.exists(out_file))
            self.assertGreater(os.path.getsize(out_file), 1000)
        finally:
            if os.path.exists(out_file):
                os.remove(out_file)


if __name__ == "__main__":
    unittest.main()
