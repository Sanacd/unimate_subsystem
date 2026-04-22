import io
import json
import os
import sys
import unittest


ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
FLASK_DIR = os.path.join(ROOT_DIR, "flask")
if FLASK_DIR not in sys.path:
    sys.path.insert(0, FLASK_DIR)
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from deploy_lite import app
from course_normalizer import normalize_course_record, normalize_credit_hours, normalize_prerequisites
from studyplan_analyzer import (
    _parse_study_plan_text,
    analyze_transcript_and_study_plan_data,
    build_study_plan_audit_workbook,
    detect_study_plan_file_type,
)


TRANSCRIPT_DATA = {
    "student": {
        "student_name": "John Doe",
        "student_id": "12345",
        "program": "BS Test Program",
        "gpa_final": 2.4,
    },
    "courses": [
        {"course_code": "CS101", "course_name": "Intro Programming", "credit_hours": 3.0, "grade": "A", "points": 12.0, "status": "completed", "term_taken": "Fall 2024", "notes": "", "source": "transcript", "prerequisites": [], "year_no": None, "semester_no": None},
        {"course_code": "CS102", "course_name": "Data Structures", "credit_hours": 3.0, "grade": "IP", "points": 0.0, "status": "in_progress", "term_taken": "Spring 2025", "notes": "", "source": "transcript", "prerequisites": [], "year_no": None, "semester_no": None},
        {"course_code": "CS103", "course_name": "Algorithms", "credit_hours": 3.0, "grade": "F", "points": 0.0, "status": "failed", "term_taken": "Spring 2025", "notes": "", "source": "transcript", "prerequisites": [], "year_no": None, "semester_no": None},
    ],
    "gpa_table": [{"Academic Year": "2024-2025", "Semester": "Spring", "Cumulative GPA": 2.4}],
}


PLAN_DATA = {
    "program_name": "BS Test Program",
    "catalog_year": "2025",
    "courses": [
        {"course_code": "CS101", "course_name": "Intro Programming", "credit_hours": 3.0, "category": "Department / Required", "year_no": 1, "semester_no": 1, "status": "", "grade": "", "term_taken": "", "prerequisites": [], "notes": "", "source": "study_plan"},
        {"course_code": "CS102", "course_name": "Data Structures", "credit_hours": 3.0, "category": "Department / Required", "year_no": 1, "semester_no": 1, "status": "", "grade": "", "term_taken": "", "prerequisites": ["CS101"], "notes": "", "source": "study_plan"},
        {"course_code": "CS103", "course_name": "Algorithms", "credit_hours": 3.0, "category": "Department / Required", "year_no": 1, "semester_no": 1, "status": "", "grade": "", "term_taken": "", "prerequisites": ["CS101"], "notes": "", "source": "study_plan"},
        {"course_code": "CS104", "course_name": "Operating Systems", "credit_hours": 3.0, "category": "Department / Required", "year_no": 1, "semester_no": 2, "status": "", "grade": "", "term_taken": "", "prerequisites": ["CS103"], "notes": "", "source": "study_plan"},
    ],
    "slot_rules": {},
    "source_type": "test",
}


class StudyPlanAnalyzerTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()

    def test_missing_transcript_file_returns_400(self):
        study_plan_bytes = io.BytesIO(json.dumps({"courses": []}).encode("utf-8"))
        response = self.client.post(
            "/api/analyze-study-plan",
            data={"study_plan": (study_plan_bytes, "plan.json")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("Transcript file is required", response.get_json()["error"])

    def test_missing_study_plan_file_returns_400(self):
        transcript_bytes = io.BytesIO(b"placeholder transcript")
        response = self.client.post(
            "/api/analyze-study-plan",
            data={"transcript": (transcript_bytes, "transcript.txt")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("Study plan file is required", response.get_json()["error"])

    def test_empty_transcript_result_raises_validation_error(self):
        with self.assertRaisesRegex(ValueError, "No transcript courses could be extracted"):
            analyze_transcript_and_study_plan_data({"student": {}, "courses": []}, PLAN_DATA)

    def test_study_plan_file_type_detection(self):
        self.assertEqual(detect_study_plan_file_type("plan.pdf"), "pdf")
        self.assertEqual(detect_study_plan_file_type("plan.docx"), "docx")
        self.assertEqual(detect_study_plan_file_type("plan.txt"), "txt")
        self.assertEqual(detect_study_plan_file_type("plan.csv"), "csv")
        self.assertEqual(detect_study_plan_file_type("plan.xlsx"), "xlsx")
        self.assertEqual(detect_study_plan_file_type("plan.png"), "image")

    def test_credit_hour_normalization_rejects_poisoned_numbers(self):
        self.assertEqual(normalize_credit_hours("CS101 Intro Programming 3"), 3.0)
        self.assertIsNone(normalize_credit_hours("Credits 3000"))

    def test_prerequisite_normalization(self):
        self.assertEqual(normalize_prerequisites("Prerequisite: CS101, MATH102"), ["CS101", "MATH102"])

    def test_malformed_row_is_skipped_by_normalizer(self):
        record = normalize_course_record(
            {"course_name": "This row accidentally absorbed prerequisite text and many unrelated tokens " * 5},
            source="study_plan_test",
            allow_name_only=True,
        )
        self.assertIsNone(record)

    def test_text_study_plan_parser_best_effort(self):
        text = """
        BS Test Program Study Plan 2025
        First Year
        First Semester
        CS101 Intro Programming 3
        CS102 Data Structures 3 Prerequisite: CS101
        Second Semester
        CS103 Algorithms 3 Prerequisite: CS102
        """
        plan = _parse_study_plan_text(text, "txt")
        rows = {row["course_code"]: row for row in plan["courses"]}
        self.assertEqual(plan["source_type"], "txt")
        self.assertEqual(rows["CS101"]["year_no"], 1)
        self.assertEqual(rows["CS101"]["semester_no"], 1)
        self.assertEqual(rows["CS102"]["prerequisites"], ["CS101"])
        self.assertEqual(rows["CS102"]["credit_hours"], 3.0)
        self.assertEqual(rows["CS103"]["semester_no"], 2)

    def test_status_assignment_and_summary(self):
        result = analyze_transcript_and_study_plan_data(TRANSCRIPT_DATA, PLAN_DATA)

        preview = {row["course_code"]: row for row in result.preview_rows}
        self.assertEqual(preview["CS101"]["status"], "completed")
        self.assertEqual(preview["CS102"]["status"], "in_progress")
        self.assertEqual(preview["CS103"]["status"], "failed")
        self.assertEqual(preview["CS104"]["status"], "blocked")
        self.assertEqual(result.summary["credits_completed"], 3.0)
        self.assertEqual(result.summary["credits_in_progress"], 3.0)
        self.assertEqual(result.summary["credits_remaining"], 9.0)
        self.assertEqual(preview["CS104"]["credit_hours"], 3.0)
        self.assertTrue(any(item["course_code"] == "CS103" for item in result.eligible_next_semester))

    def test_excel_generation_success(self):
        result = analyze_transcript_and_study_plan_data(TRANSCRIPT_DATA, PLAN_DATA)
        workbook = build_study_plan_audit_workbook(result, result.preview_rows)
        self.assertEqual(
            workbook.sheetnames,
            ["Study Plan Audit", "Summary", "Eligible Next Semester", "Personalized Advice"],
        )


if __name__ == "__main__":
    unittest.main()
