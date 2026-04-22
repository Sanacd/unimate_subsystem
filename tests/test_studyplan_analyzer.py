import io
import json
import os
import sys
import unittest
from unittest.mock import patch


ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
FLASK_DIR = os.path.join(ROOT_DIR, "flask")
if FLASK_DIR not in sys.path:
    sys.path.insert(0, FLASK_DIR)
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from deploy_lite import app
from studyplan_analyzer import (
    analyze_study_plan,
    analyze_transcript_and_study_plan,
    analyze_transcript_and_study_plan_data,
    detect_file_type,
)


TRANSCRIPT_DATA = {
    "student": {
        "student_name": "John Doe",
        "student_id": "12345",
        "program": "BS Test Program",
        "gpa_final": 3.1,
    },
    "courses": [
        {
            "course_code": "CS101",
            "course_name": "Intro Programming",
            "credit_hours": 3,
            "grade": "A",
            "status": "completed",
            "term_taken": "Fall 2024",
            "notes": "",
            "points": 12.0,
        },
        {
            "course_code": "CS102",
            "course_name": "Data Structures",
            "credit_hours": 3,
            "grade": "IP",
            "status": "in_progress",
            "term_taken": "Spring 2025",
            "notes": "",
            "points": 0.0,
        },
    ],
}


PLAN_DATA = {
    "program_name": "BS Test Program",
    "catalog_year": "2025",
    "courses": [
        {
            "course_code": "CS101",
            "course_name": "Intro Programming",
            "credit_hours": 3,
            "prerequisites": [],
            "year_no": 1,
            "semester_no": 1,
            "category": "Required",
            "notes": "",
        },
        {
            "course_code": "CS102",
            "course_name": "Data Structures",
            "credit_hours": 3,
            "prerequisites": ["CS101"],
            "year_no": 1,
            "semester_no": 2,
            "category": "Required",
            "notes": "",
        },
    ],
    "slot_rules": {},
    "source_type": "json",
}


MERGED_ROWS = [
    {
        "course_code": "CS101",
        "course_name": "Intro Programming",
        "credit_hours": 3,
        "year_no": 1,
        "semester_no": 1,
        "category": "Required",
        "prerequisites": [],
        "status": "completed",
        "blocked_by_prerequisite": False,
        "matched": True,
        "match_type": "exact_code",
        "transcript_course_code": "CS101",
        "transcript_course_name": "Intro Programming",
        "transcript_credit_hours": 3,
        "grade": "A",
        "term_taken": "Fall 2024",
        "notes": "",
    },
    {
        "course_code": "CS102",
        "course_name": "Data Structures",
        "credit_hours": 3,
        "year_no": 1,
        "semester_no": 2,
        "category": "Required",
        "prerequisites": ["CS101"],
        "status": "in_progress",
        "blocked_by_prerequisite": False,
        "matched": True,
        "match_type": "exact_code",
        "transcript_course_code": "CS102",
        "transcript_course_name": "Data Structures",
        "transcript_credit_hours": 3,
        "grade": "IP",
        "term_taken": "Spring 2025",
        "notes": "",
    },
]


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

    def test_detect_file_type_current_names(self):
        self.assertEqual(detect_file_type("plan.pdf"), "pdf")
        self.assertEqual(detect_file_type("plan.docx"), "docx")
        self.assertEqual(detect_file_type("plan.txt"), "txt")
        self.assertEqual(detect_file_type("plan.json"), "json")
        self.assertEqual(detect_file_type("plan.csv"), "csv")
        self.assertEqual(detect_file_type("plan.xlsx"), "xlsx")

    def test_analyze_study_plan_json_result_shape(self):
        json_payload = {
            "program_name": "BS Test Program",
            "catalog_year": "2025",
            "courses": [
                {
                    "course_code": "CS101",
                    "course_name": "Intro Programming",
                    "credit_hours": 3,
                    "prerequisites": [],
                    "year_no": 1,
                    "semester_no": 1,
                    "category": "Required",
                }
            ],
            "slot_rules": {},
        }
        with patch("studyplan_analyzer.detect_file_type", return_value="json"), patch(
            "studyplan_analyzer.read_json_file", return_value=json_payload
        ):
            result = analyze_study_plan("dummy.json", program_hint="BS Test Program")

        self.assertEqual(result["program_name"], "BS Test Program")
        self.assertEqual(result["catalog_year"], "2025")
        self.assertEqual(result["source_type"], "json")
        self.assertEqual(len(result["courses"]), 1)
        self.assertEqual(result["courses"][0]["course_code"], "CS101")
        self.assertEqual(result["courses"][0]["credit_hours"], 3)

    def test_analyze_transcript_and_study_plan_data_result_shape(self):
        with patch("studyplan_analyzer.GEMINI_API_KEY", "test-key"), patch(
            "studyplan_analyzer._match_study_plan_courses_with_model", return_value=MERGED_ROWS
        ):
            result = analyze_transcript_and_study_plan_data(TRANSCRIPT_DATA, PLAN_DATA)

        self.assertEqual(result["comparison_engine"], "gemini")
        self.assertEqual(result["student"]["student_name"], "John Doe")
        self.assertEqual(result["study_plan_meta"]["program_name"], "BS Test Program")
        self.assertEqual(len(result["merged_rows"]), 2)
        self.assertEqual(result["preview_rows"][0]["course_code"], "CS101")
        self.assertIn("summary", result)
        self.assertIn("advice", result)

    def test_analyze_transcript_and_study_plan_data_generates_excel_paths(self):
        with patch("studyplan_analyzer.GEMINI_API_KEY", "test-key"), patch(
            "studyplan_analyzer._match_study_plan_courses_with_model", return_value=MERGED_ROWS
        ), patch(
            "studyplan_analyzer._build_audit_workbook",
            return_value=os.path.join("flask", "uploads", "audit.xlsx"),
        ) as mock_audit, patch(
            "studyplan_analyzer.build_structured_study_plan_workbook",
            return_value=os.path.join("flask", "uploads", "structured.xlsx"),
        ) as mock_structured:
            result = analyze_transcript_and_study_plan_data(
                TRANSCRIPT_DATA,
                PLAN_DATA,
                output_dir=os.path.join("flask", "uploads"),
            )

        self.assertEqual(result["excel_filename"], "audit.xlsx")
        self.assertEqual(result["structured_excel_filename"], "structured.xlsx")
        mock_audit.assert_called_once()
        mock_structured.assert_called_once()

    def test_full_pipeline_orchestration_uses_current_functions(self):
        expected_result = {
            "student": TRANSCRIPT_DATA["student"],
            "study_plan_meta": {"program_name": "BS Test Program", "catalog_year": "2025", "source_type": "json"},
            "comparison_engine": "gemini",
            "summary": {"total_courses": 2},
            "advice": ["Test advice"],
            "merged_rows": MERGED_ROWS,
            "preview_rows": MERGED_ROWS,
            "excel_path": "",
            "excel_filename": "",
            "structured_excel_path": "",
            "structured_excel_filename": "",
        }
        with patch("pdf_extractor.extract_transcript_data", return_value=TRANSCRIPT_DATA), patch(
            "studyplan_analyzer.analyze_study_plan", return_value=PLAN_DATA
        ), patch(
            "studyplan_analyzer.analyze_transcript_and_study_plan_data", return_value=expected_result.copy()
        ), patch(
            "studyplan_analyzer._build_audit_workbook", return_value=os.path.join("flask", "uploads", "audit.xlsx")
        ), patch(
            "studyplan_analyzer.build_structured_study_plan_workbook",
            return_value=os.path.join("flask", "uploads", "structured.xlsx"),
        ):
            result = analyze_transcript_and_study_plan(
                transcript_path="dummy_transcript.pdf",
                study_plan_path="dummy_plan.pdf",
                output_dir=os.path.join("flask", "uploads"),
            )

        self.assertEqual(result["student"]["student_name"], "John Doe")
        self.assertTrue(result["excel_filename"].startswith("study_plan_audit_"))
        self.assertTrue(result["excel_filename"].endswith(".xlsx"))
        self.assertTrue(result["structured_excel_filename"].startswith("structured_study_plan_"))
        self.assertTrue(result["structured_excel_filename"].endswith(".xlsx"))
        self.assertEqual(len(result["merged_rows"]), 2)


if __name__ == "__main__":
    unittest.main()
