from __future__ import annotations

import json
import os
import re
import uuid
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from functools import lru_cache
from typing import Any
from xml.etree import ElementTree as ET

import pandas as pd
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from course_normalizer import (
    normalize_course_code,
    normalize_course_name,
    normalize_course_name_key,
    normalize_course_record,
    normalize_credit_hours,
    normalize_integer,
    normalize_prerequisites,
    record_match_key,
)
from pdf_extractor import extract_text, extract_transcript_data


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
YEAR_WORDS = {"first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5}
STATUS_LABELS = {
    "completed": "Completed",
    "in_progress": "In Progress",
    "not_taken": "Not Taken",
    "failed": "Failed / Retake Required",
    "blocked": "Blocked by prerequisite",
    "exempt": "Exempt / Transfer",
}
STATUS_FILLS = {
    "completed": "C6EFCE",
    "in_progress": "FFF2CC",
    "not_taken": "D9D9D9",
    "failed": "F4CCCC",
    "blocked": "FCE5CD",
    "exempt": "D9EAD3",
}
TERM_ORDER = {"spring": 1, "summer": 2, "fall": 3}
COURSE_CODE_RE = re.compile(r"[A-Z]{2,6}\s*[-]?\s*(?:\d{3}[A-Z]?|X{2,4})", re.I)


@dataclass
class AnalysisArtifacts:
    student: dict[str, Any]
    summary: dict[str, Any]
    eligible_next_semester: list[dict[str, Any]]
    advice: list[str]
    preview_rows: list[dict[str, Any]]
    excel_path: str


@lru_cache(maxsize=1)
def load_catalog_study_plans() -> dict[str, Any]:
    with open(os.path.join(BASE_DIR, "university_studyplans.json"), "r", encoding="utf-8") as fh:
        data = json.load(fh)
    return data.get("programs", data)


@lru_cache(maxsize=1)
def load_academic_policy() -> dict[str, Any]:
    with open(os.path.join(BASE_DIR, "academic_policy.json"), "r", encoding="utf-8") as fh:
        return json.load(fh)


def normalize_term(term: Any) -> str:
    return re.sub(r"\s+", " ", str(term or "").strip())


def classify_transcript_status(grade: Any, points: Any) -> str:
    grade_text = str(grade or "").strip().upper()
    if grade_text in {"WAIVED", "WAIVE", "TR", "TRANSFER"}:
        return "exempt"
    if grade_text in {"IP", "I", "IN PROGRESS", "INPROGRESS"}:
        return "in_progress"
    if grade_text in {"F", "FA", "NF"}:
        return "failed"
    if grade_text in {"W", "WF", "WP", "WITHDRAWN"}:
        return "not_taken"
    try:
        if float(points or 0) > 0:
            return "completed"
    except (TypeError, ValueError):
        pass
    if grade_text and grade_text not in {"0", "0.0", "0.00"}:
        return "completed"
    return "in_progress"


def _term_sort_key(term: str) -> tuple[int, int, str]:
    match = re.search(r"(20\d{2})", term or "")
    year = int(match.group(1)) if match else 0
    lowered = (term or "").lower()
    season_rank = 0
    for name, rank in TERM_ORDER.items():
        if name in lowered:
            season_rank = rank
            break
    return (year, season_rank, lowered)


def _status_rank(status: str) -> int:
    return {
        "completed": 5,
        "exempt": 4,
        "in_progress": 3,
        "failed": 2,
        "not_taken": 1,
    }.get(status, 0)


def _attempt_sort_key(attempt: dict[str, Any]) -> tuple[int, int, int, str]:
    year, season, lowered = _term_sort_key(attempt.get("term_taken", ""))
    return (_status_rank(attempt.get("status", "")), year, season, lowered)


def _catalog_program_from_hint(program_hint: str | None) -> tuple[str | None, dict[str, Any] | None]:
    if not program_hint:
        return None, None

    normalized_hint = normalize_course_name_key(program_hint)
    programs = load_catalog_study_plans()

    for key, value in programs.items():
        if normalize_course_name_key(key) == normalized_hint:
            return key, value
    for key, value in programs.items():
        if normalized_hint and normalized_hint in normalize_course_name_key(key):
            return key, value
    return None, None


def _infer_year_semester_from_level(level_value: Any) -> tuple[int | None, int | None]:
    level_num = normalize_integer(level_value)
    if level_num is None:
        return None, None
    year_no = ((level_num - 1) // 2) + 1
    semester_no = 1 if level_num % 2 else 2
    return year_no, semester_no


def detect_study_plan_file_type(file_path: str) -> str:
    ext = os.path.splitext(file_path)[1].lower()
    mapping = {
        ".pdf": "pdf",
        ".docx": "docx",
        ".txt": "txt",
        ".json": "json",
        ".csv": "csv",
        ".xlsx": "xlsx",
        ".xls": "xlsx",
    }
    if ext in mapping:
        return mapping[ext]
    if ext in IMAGE_EXTENSIONS:
        return "image"
    return "unknown"


def _extract_docx_text(file_path: str) -> str:
    paragraphs: list[str] = []
    with zipfile.ZipFile(file_path) as docx_zip:
        xml_bytes = docx_zip.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for paragraph in root.findall(".//w:p", namespaces):
        parts = [node.text for node in paragraph.findall(".//w:t", namespaces) if node.text]
        line = "".join(parts).strip()
        if line:
            paragraphs.append(line)
    return "\n".join(paragraphs)


def _extract_plain_text(file_path: str) -> str:
    with open(file_path, "r", encoding="utf-8", errors="ignore") as fh:
        return fh.read()


def _parse_year_semester_from_text(line: str) -> tuple[int | None, int | None]:
    lowered = line.lower()
    level_match = re.search(r"\blevel\s*(\d+)\b", lowered)
    if level_match:
        return _infer_year_semester_from_level(level_match.group(1))

    year_no = None
    semester_no = None

    year_match = re.search(r"\b(first|second|third|fourth|fifth)\s+year\b", lowered)
    if year_match:
        year_no = YEAR_WORDS.get(year_match.group(1))

    sem_match = re.search(r"\b(first|second)\s+semester\b", lowered)
    if sem_match:
        semester_no = 1 if sem_match.group(1) == "first" else 2

    numeric_year_match = re.search(r"\byear\s*(\d+)\b", lowered)
    if numeric_year_match:
        year_no = int(numeric_year_match.group(1))

    numeric_sem_match = re.search(r"\bsemester\s*(\d+)\b", lowered)
    if numeric_sem_match:
        semester_no = int(numeric_sem_match.group(1))

    return year_no, semester_no


def _extract_program_name(text: str, program_hint: str | None = None) -> str:
    if program_hint:
        return program_hint
    for line in text.splitlines()[:20]:
        stripped = line.strip()
        lowered = stripped.lower()
        if "study plan" in lowered or "curriculum" in lowered:
            return stripped
        if any(token in lowered for token in ("bachelor", "bs ", "program", "major")) and len(stripped) < 120:
            return stripped
    return ""


def _extract_catalog_year(text: str) -> str:
    match = re.search(r"(20\d{2}(?:\s*[-/]\s*20\d{2})?)", text)
    return match.group(1).strip() if match else ""


def _normalize_table_headers(values: list[Any]) -> list[str]:
    headers = []
    for item in values:
        text = re.sub(r"[^a-z0-9]+", "_", str(item or "").strip().lower()).strip("_")
        headers.append(text or "col")
    return headers


def _row_to_structured_dict(headers: list[str], values: list[Any]) -> dict[str, Any]:
    row = {}
    for index, value in enumerate(values):
        key = headers[index] if index < len(headers) else f"col_{index}"
        row[key] = value
    return row


def _looks_like_header_row(values: list[Any]) -> bool:
    joined = " ".join(str(item or "").lower() for item in values)
    return "course" in joined and ("credit" in joined or "hour" in joined or "code" in joined)


def _parse_pdf_table_rows(file_path: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                cleaned = [[(cell or "").strip() for cell in raw_row] for raw_row in table if any((cell or "").strip() for cell in raw_row)]
                if not cleaned:
                    continue
                header_values = cleaned[0] if _looks_like_header_row(cleaned[0]) else None
                headers = _normalize_table_headers(header_values or [f"col_{i}" for i in range(len(cleaned[0]))])
                data_rows = cleaned[1:] if header_values else cleaned
                for data_row in data_rows:
                    row_dict = _row_to_structured_dict(headers, data_row)
                    rows.append(
                        {
                            "course_code": row_dict.get("course_code") or row_dict.get("code") or row_dict.get("course") or row_dict.get("col_0"),
                            "course_name": row_dict.get("course_name") or row_dict.get("course_title") or row_dict.get("title") or row_dict.get("col_1"),
                            "credit_hours": row_dict.get("credit_hours") or row_dict.get("credits") or row_dict.get("hours") or row_dict.get("col_2"),
                            "prerequisites": row_dict.get("prerequisite") or row_dict.get("prerequisites"),
                            "category": row_dict.get("category") or row_dict.get("type") or row_dict.get("requirement"),
                            "notes": "Imported from PDF table",
                        }
                    )
    return rows


def _parse_course_line(line: str, year_no: int | None, semester_no: int | None) -> dict[str, Any] | None:
    compact = re.sub(r"\s+", " ", line).strip(" |-")
    if not compact:
        return None

    code_match = COURSE_CODE_RE.search(compact)
    if not code_match:
        return None

    code = normalize_course_code(code_match.group(0))
    remainder = compact[code_match.end():].strip(" -|")
    if not remainder:
        return None

    prereq_match = re.search(r"\b(?:pre[- ]?req(?:uisite)?s?)\b[:\-]?\s*(.+)$", remainder, re.I)
    prereq_text = prereq_match.group(1) if prereq_match else ""
    working_text = remainder[:prereq_match.start()].strip(" -|,") if prereq_match else remainder

    credit_hours = normalize_credit_hours(working_text)
    if credit_hours is not None:
        credit_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:credit|credits|cr|hrs|hours)?\s*$", working_text, re.I)
        title_part = working_text[:credit_match.start()].strip(" -|,") if credit_match else working_text
    else:
        title_part = working_text

    return normalize_course_record(
        {
            "course_code": code,
            "course_name": title_part,
            "credit_hours": credit_hours,
            "prerequisites": prereq_text,
            "year_no": year_no,
            "semester_no": semester_no,
            "notes": "Best-effort text extraction",
        },
        source="study_plan_text",
    )


def _parse_study_plan_text(text: str, source_type: str, program_hint: str | None = None) -> dict[str, Any]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    courses = []
    current_year = None
    current_semester = None

    for line in lines:
        year_no, semester_no = _parse_year_semester_from_text(line)
        if year_no is not None:
            current_year = year_no
        if semester_no is not None:
            current_semester = semester_no

        record = _parse_course_line(line, current_year, current_semester)
        if record:
            courses.append(record)

    return {
        "program_name": _extract_program_name(text, program_hint=program_hint),
        "catalog_year": _extract_catalog_year(text),
        "courses": courses,
        "slot_rules": {},
        "source_type": source_type,
        "ocr_ready": source_type == "image",
    }


def _normalize_plan_rows(
    rows: list[dict[str, Any]],
    *,
    source: str,
    level: Any = None,
    allow_name_only: bool = True,
) -> list[dict[str, Any]]:
    normalized = []
    default_year, default_semester = _infer_year_semester_from_level(level)
    for row in rows:
        record = normalize_course_record(
            row,
            source=source,
            default_year=default_year,
            default_semester=default_semester,
            allow_name_only=allow_name_only,
        )
        if record:
            normalized.append(record)
    return normalized


def _normalize_program_structure(program_name: str, program_data: dict[str, Any]) -> dict[str, Any]:
    courses = []
    for level, rows in (program_data.get("levels") or {}).items():
        courses.extend(_normalize_plan_rows(rows, source="study_plan_catalog", level=level, allow_name_only=True))

    slot_rules = {}
    for slot, rule in (program_data.get("slot_rules") or {}).items():
        slot_rules[normalize_course_code(slot)] = {
            "prefixes": [normalize_course_code(item) for item in rule.get("prefixes", []) if item],
            "specific": [normalize_course_code(item) for item in rule.get("specific", []) if item],
        }

    return {
        "program_name": program_name,
        "catalog_year": str(program_data.get("version") or program_data.get("catalog_year") or "").strip(),
        "courses": courses,
        "slot_rules": slot_rules,
        "source_type": "catalog",
    }


def _merge_plan_sources(*course_lists: list[dict[str, Any]]) -> list[dict[str, Any]]:
    merged: list[dict[str, Any]] = []
    seen = set()
    for course_list in course_lists:
        for record in course_list:
            key = record_match_key(record)
            if key in seen:
                continue
            if not record.get("course_code") and not record.get("course_name"):
                continue
            seen.add(key)
            merged.append(record)
    return merged


def extract_study_plan_data(file_path: str, program_hint: str | None = None) -> dict[str, Any]:
    file_type = detect_study_plan_file_type(file_path)

    if file_type == "pdf":
        table_rows = _parse_pdf_table_rows(file_path)
        normalized_table_rows = _normalize_plan_rows(table_rows, source="study_plan_pdf_table", allow_name_only=False)
        text_rows = _parse_study_plan_text(extract_text(file_path), "pdf", program_hint=program_hint)
        plan = {
            "program_name": text_rows.get("program_name") or program_hint or "",
            "catalog_year": text_rows.get("catalog_year") or "",
            "courses": _merge_plan_sources(normalized_table_rows, text_rows.get("courses", [])),
            "slot_rules": {},
            "source_type": "pdf",
        }
    elif file_type == "docx":
        plan = _parse_study_plan_text(_extract_docx_text(file_path), "docx", program_hint=program_hint)
    elif file_type == "txt":
        plan = _parse_study_plan_text(_extract_plain_text(file_path), "txt", program_hint=program_hint)
    elif file_type == "xlsx":
        excel_book = pd.read_excel(file_path, sheet_name=None)
        structured_rows = []
        for sheet_name, df in excel_book.items():
            for record in _normalize_plan_rows(df.fillna("").to_dict("records"), source=f"study_plan_xlsx:{sheet_name}", allow_name_only=True):
                if not record.get("notes"):
                    record["notes"] = f"Imported from sheet: {sheet_name}"
                structured_rows.append(record)
        plan = {
            "program_name": program_hint or "",
            "catalog_year": "",
            "courses": structured_rows,
            "slot_rules": {},
            "source_type": "xlsx",
        }
    elif file_type == "csv":
        df = pd.read_csv(file_path)
        plan = {
            "program_name": program_hint or "",
            "catalog_year": "",
            "courses": _normalize_plan_rows(df.fillna("").to_dict("records"), source="study_plan_csv", allow_name_only=True),
            "slot_rules": {},
            "source_type": "csv",
        }
    elif file_type == "image":
        plan = {
            "program_name": program_hint or "",
            "catalog_year": "",
            "courses": [],
            "slot_rules": {},
            "source_type": "image",
            "ocr_ready": True,
            "notes": "Image study plan uploaded. OCR is not enabled in the current deployment.",
        }
    else:
        with open(file_path, "r", encoding="utf-8") as fh:
            raw = json.load(fh)

        if isinstance(raw, dict) and "programs" in raw:
            programs = raw["programs"]
            selected_name, selected_program = _catalog_program_from_hint(program_hint)
            if selected_program is None and programs:
                first_key = next(iter(programs.keys()))
                selected_name, selected_program = first_key, programs[first_key]
            plan = _normalize_program_structure(selected_name or (program_hint or ""), selected_program or {})
        elif isinstance(raw, dict) and "levels" in raw:
            plan = _normalize_program_structure(program_hint or raw.get("program_name") or raw.get("name") or "", raw)
        elif isinstance(raw, list):
            plan = {
                "program_name": program_hint or "",
                "catalog_year": "",
                "courses": _normalize_plan_rows(raw, source="study_plan_json_rows", allow_name_only=True),
                "slot_rules": {},
                "source_type": "json_rows",
            }
        else:
            plan = {
                "program_name": raw.get("program_name") or program_hint or "",
                "catalog_year": str(raw.get("catalog_year") or raw.get("version") or "").strip(),
                "courses": _normalize_plan_rows(raw.get("courses", []), source="study_plan_json", allow_name_only=True),
                "slot_rules": {},
                "source_type": "json",
            }

    if not plan.get("courses"):
        selected_name, selected_program = _catalog_program_from_hint(program_hint)
        if selected_program is not None:
            fallback = _normalize_program_structure(selected_name or "", selected_program)
            fallback["source_type"] = f"{plan.get('source_type', 'unknown')}+catalog_fallback"
            if plan.get("ocr_ready"):
                fallback["notes"] = plan.get("notes", "")
            return fallback

    return plan


def _build_transcript_lookup(transcript_data: dict[str, Any]) -> dict[str, Any]:
    by_code: dict[str, list[dict[str, Any]]] = defaultdict(list)
    by_name: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for course in transcript_data.get("courses", []):
        code_key, name_key = record_match_key(course)
        if code_key:
            by_code[code_key].append(course)
        if name_key:
            by_name[name_key].append(course)

    best_attempts: dict[str, dict[str, Any]] = {}
    failed_codes = set()
    repeated_codes = set()

    for code, attempts in by_code.items():
        attempts_sorted = sorted(attempts, key=_attempt_sort_key, reverse=True)
        best_attempts[code] = attempts_sorted[0]
        if len(attempts) > 1:
            repeated_codes.add(code)
        if any(attempt.get("status") == "failed" for attempt in attempts):
            failed_codes.add(code)

    for name, attempts in by_name.items():
        attempts_sorted = sorted(attempts, key=_attempt_sort_key, reverse=True)
        best_attempts.setdefault(name, attempts_sorted[0])

    return {
        "best_attempts": best_attempts,
        "failed_codes": failed_codes,
        "repeated_codes": repeated_codes,
        "completed_codes": {
            normalize_course_code(item.get("course_code"))
            for item in transcript_data.get("courses", [])
            if item.get("status") in {"completed", "exempt"}
        },
        "in_progress_codes": {
            normalize_course_code(item.get("course_code"))
            for item in transcript_data.get("courses", [])
            if item.get("status") == "in_progress"
        },
    }


def _candidate_matches_for_slot(slot_code: str, slot_rule: dict[str, Any], transcript_courses: list[dict[str, Any]]) -> list[dict[str, Any]]:
    specifics = set(slot_rule.get("specific", []))
    prefixes = tuple(slot_rule.get("prefixes", []))
    candidates = []
    for course in transcript_courses:
        course_code = normalize_course_code(course.get("course_code"))
        if course_code == slot_code:
            continue
        if course_code in specifics or (prefixes and any(course_code.startswith(prefix) for prefix in prefixes)):
            candidates.append(course)
    return sorted(candidates, key=_attempt_sort_key, reverse=True)


def _prereq_codes_satisfied(prerequisites: list[str], satisfied_codes: set[str]) -> tuple[bool, list[str]]:
    missing = [item for item in prerequisites if normalize_course_code(item) not in satisfied_codes]
    return len(missing) == 0, missing


def _canonical_merged_row(plan_course: dict[str, Any], transcript_course: dict[str, Any] | None = None) -> dict[str, Any]:
    row = dict(plan_course)
    row["status"] = "not_taken"
    row["grade"] = ""
    row["term_taken"] = ""
    row["source"] = "merged"
    if transcript_course:
        row["status"] = transcript_course.get("status") or "not_taken"
        row["grade"] = str(transcript_course.get("grade") or "").strip()
        row["term_taken"] = normalize_term(transcript_course.get("term_taken"))
        if row.get("credit_hours") is None:
            row["credit_hours"] = transcript_course.get("credit_hours")
    return row


def _match_study_plan_courses(plan_data: dict[str, Any], transcript_data: dict[str, Any]) -> list[dict[str, Any]]:
    lookup = _build_transcript_lookup(transcript_data)
    slot_rules = plan_data.get("slot_rules", {})
    satisfied_for_prereqs = set(lookup["completed_codes"]) | set(lookup["in_progress_codes"])
    used_slot_matches: set[str] = set()
    matched_rows: list[dict[str, Any]] = []

    for course in plan_data.get("courses", []):
        code_key, name_key = record_match_key(course)
        transcript_match = lookup["best_attempts"].get(code_key) if code_key else None
        if transcript_match is None and not code_key and name_key:
            transcript_match = lookup["best_attempts"].get(name_key)

        notes: list[str] = []
        if transcript_match is None and code_key and "X" in code_key:
            candidates = _candidate_matches_for_slot(code_key, slot_rules.get(code_key, {}), transcript_data.get("courses", []))
            for candidate in candidates:
                candidate_code = normalize_course_code(candidate.get("course_code"))
                if candidate_code in used_slot_matches:
                    continue
                transcript_match = candidate
                used_slot_matches.add(candidate_code)
                notes.append(f"Matched elective slot with {candidate_code}")
                break

        row = _canonical_merged_row(course, transcript_match)

        if transcript_match:
            matched_code = normalize_course_code(transcript_match.get("course_code"))
            if matched_code in lookup["repeated_codes"]:
                notes.append("Repeated course detected in transcript history")
            if matched_code in lookup["failed_codes"] and row["status"] in {"completed", "exempt"}:
                notes.append("Completed after at least one failed attempt")

        prereq_ok, missing_prereqs = _prereq_codes_satisfied(row.get("prerequisites", []), satisfied_for_prereqs)
        if row["status"] == "not_taken" and missing_prereqs:
            row["status"] = "blocked"
            notes.append("Missing prerequisites: " + ", ".join(missing_prereqs))
        elif row["status"] == "failed" and prereq_ok:
            notes.append("Eligible for retake when offered")

        row["notes"] = "; ".join(filter(None, [row.get("notes"), *notes])).strip("; ")
        matched_rows.append(row)

    return matched_rows


def _compute_summary(rows: list[dict[str, Any]], transcript_data: dict[str, Any]) -> dict[str, Any]:
    credits_required = round(sum(float(row.get("credit_hours") or 0) for row in rows), 2)
    credits_completed = round(sum(float(row.get("credit_hours") or 0) for row in rows if row.get("status") in {"completed", "exempt"}), 2)
    credits_in_progress = round(sum(float(row.get("credit_hours") or 0) for row in rows if row.get("status") == "in_progress"), 2)
    credits_remaining = round(sum(float(row.get("credit_hours") or 0) for row in rows if row.get("status") not in {"completed", "exempt"}), 2)
    completion_percentage = round((credits_completed / credits_required) * 100, 2) if credits_required else 0.0
    return {
        "credits_required": credits_required,
        "credits_completed": credits_completed,
        "credits_in_progress": credits_in_progress,
        "credits_remaining": credits_remaining,
        "completion_percentage": completion_percentage,
        "courses_completed": sum(1 for row in rows if row.get("status") in {"completed", "exempt"}),
        "courses_in_progress": sum(1 for row in rows if row.get("status") == "in_progress"),
        "courses_failed": sum(1 for row in rows if row.get("status") == "failed"),
        "courses_blocked": sum(1 for row in rows if row.get("status") == "blocked"),
        "courses_remaining": sum(1 for row in rows if row.get("status") in {"not_taken", "blocked", "failed"}),
        "transcript_course_count": len(transcript_data.get("courses", [])),
    }


def _eligible_next_semester(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    eligible = []
    for row in rows:
        if row.get("status") not in {"not_taken", "failed"}:
            continue
        if row.get("status") == "not_taken" and row.get("prerequisites"):
            continue
        note = "Retake required" if row.get("status") == "failed" else "Prerequisites satisfied"
        eligible.append(
            {
                "course_code": row.get("course_code"),
                "course_name": row.get("course_name"),
                "credit_hours": row.get("credit_hours"),
                "year_no": row.get("year_no"),
                "semester_no": row.get("semester_no"),
                "note": note,
            }
        )
    eligible.sort(key=lambda item: ((item.get("year_no") or 99), (item.get("semester_no") or 99), item.get("course_code") or ""))
    return eligible


def _generate_advice(rows: list[dict[str, Any]], summary: dict[str, Any], transcript_data: dict[str, Any], student: dict[str, Any]) -> list[str]:
    policy = load_academic_policy()
    advice = []

    failed = [row for row in rows if row.get("status") == "failed"]
    blocked = [row for row in rows if row.get("status") == "blocked"]
    in_progress = [row for row in rows if row.get("status") == "in_progress"]
    eligible = _eligible_next_semester(rows)

    if failed:
        advice.append(f"Prioritize retaking {', '.join(row['course_code'] for row in failed[:4])} because those credits are still outstanding.")
    if blocked:
        advice.append(f"Resolve prerequisite chains first. Currently blocked: {', '.join(row['course_code'] for row in blocked[:4])}.")
    if eligible:
        advice.append(f"Next-semester eligible options include {', '.join(item['course_code'] for item in eligible[:6])}.")
    if in_progress:
        advice.append(f"You have {len(in_progress)} course(s) in progress. Completing them will open more registration options.")

    gpa_value = student.get("gpa_final")
    min_gpa = float(policy.get("graduation_requirements", {}).get("minimum_gpa", 2.0))
    if isinstance(gpa_value, (int, float)) and gpa_value < min_gpa:
        advice.append(f"Your cumulative GPA is {gpa_value:.2f}, below the graduation minimum of {min_gpa:.2f}.")

    if summary["completion_percentage"] >= 85:
        advice.append("You are close to completion; verify capstone, internship, and final electives early.")
    elif summary["completion_percentage"] < 40:
        advice.append("Focus on core foundational courses first to avoid future prerequisite bottlenecks.")

    if not advice:
        advice.append("Your plan is on track. Continue with the earliest unlocked required courses next.")

    return advice


def build_student_snapshot(transcript_data: dict[str, Any], plan_data: dict[str, Any]) -> dict[str, Any]:
    student = dict(transcript_data.get("student", {}))
    program = student.get("program") or plan_data.get("program_name") or ""
    student["program"] = program
    student["study_plan_program"] = plan_data.get("program_name") or program
    student["study_plan_catalog"] = plan_data.get("catalog_year") or ""
    return student


def build_study_plan_audit_workbook(artifacts: AnalysisArtifacts, rows: list[dict[str, Any]]) -> Workbook:
    wb = Workbook()
    ws_audit = wb.active
    ws_audit.title = "Study Plan Audit"
    ws_summary = wb.create_sheet("Summary")
    ws_eligible = wb.create_sheet("Eligible Next Semester")
    ws_advice = wb.create_sheet("Personalized Advice")

    ws_audit.append(["Year", "Semester", "Course Code", "Course Name", "Credits", "Category", "Prerequisites", "Status", "Grade", "Term Taken", "Notes"])
    for cell in ws_audit[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAD3")

    for row in rows:
        ws_audit.append(
            [
                row.get("year_no"),
                row.get("semester_no"),
                row.get("course_code"),
                row.get("course_name"),
                row.get("credit_hours"),
                row.get("category"),
                ", ".join(row.get("prerequisites", [])),
                STATUS_LABELS.get(row.get("status"), row.get("status")),
                row.get("grade"),
                row.get("term_taken"),
                row.get("notes"),
            ]
        )
        ws_audit.cell(row=ws_audit.max_row, column=8).fill = PatternFill("solid", fgColor=STATUS_FILLS.get(row.get("status"), "FFFFFF"))

    ws_summary.append(["Field", "Value"])
    ws_summary["A1"].font = ws_summary["B1"].font = Font(bold=True)
    for key, value in [
        ("Student Name", artifacts.student.get("student_name") or ""),
        ("Student ID", artifacts.student.get("student_id") or ""),
        ("Program", artifacts.student.get("program") or ""),
        ("Study Plan Catalog", artifacts.student.get("study_plan_catalog") or ""),
        ("Credits Required", artifacts.summary.get("credits_required")),
        ("Credits Completed", artifacts.summary.get("credits_completed")),
        ("Credits In Progress", artifacts.summary.get("credits_in_progress")),
        ("Credits Remaining", artifacts.summary.get("credits_remaining")),
        ("Completion Percentage", artifacts.summary.get("completion_percentage")),
    ]:
        ws_summary.append([key, value])

    ws_eligible.append(["Course Code", "Course Name", "Credits", "Year", "Semester", "Note"])
    for cell in ws_eligible[1]:
        cell.font = Font(bold=True)
    for item in artifacts.eligible_next_semester:
        ws_eligible.append(
            [
                item.get("course_code"),
                item.get("course_name"),
                item.get("credit_hours"),
                item.get("year_no"),
                item.get("semester_no"),
                item.get("note"),
            ]
        )

    ws_advice.append(["Advice"])
    ws_advice["A1"].font = Font(bold=True)
    for item in artifacts.advice:
        ws_advice.append([item])

    for ws in wb.worksheets:
        for column in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = min(max(max_length + 2, 12), 50)

    return wb


def export_study_plan_audit_excel(artifacts: AnalysisArtifacts, rows: list[dict[str, Any]], output_dir: str) -> str:
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"study_plan_audit_{uuid.uuid4().hex}.xlsx")
    build_study_plan_audit_workbook(artifacts, rows).save(output_path)
    return output_path


def analyze_transcript_and_study_plan_data(
    transcript_data: dict[str, Any],
    plan_data: dict[str, Any],
    output_dir: str | None = None,
) -> AnalysisArtifacts:
    if not transcript_data.get("courses"):
        raise ValueError("No transcript courses could be extracted from the uploaded transcript.")
    if not plan_data.get("courses"):
        raise ValueError("No study plan courses could be extracted from the uploaded study plan.")

    student = build_student_snapshot(transcript_data, plan_data)
    merged_rows = _match_study_plan_courses(plan_data, transcript_data)
    summary = _compute_summary(merged_rows, transcript_data)
    advice = _generate_advice(merged_rows, summary, transcript_data, transcript_data.get("student", {}))
    eligible = _eligible_next_semester(merged_rows)
    preview_rows = merged_rows[:10]

    temp = AnalysisArtifacts(
        student=student,
        summary=summary,
        eligible_next_semester=eligible,
        advice=advice,
        preview_rows=preview_rows,
        excel_path="",
    )
    excel_path = export_study_plan_audit_excel(temp, merged_rows, output_dir) if output_dir else ""

    return AnalysisArtifacts(
        student=student,
        summary=summary,
        eligible_next_semester=eligible,
        advice=advice,
        preview_rows=preview_rows,
        excel_path=excel_path,
    )


def analyze_transcript_and_study_plan(transcript_path: str, study_plan_path: str, output_dir: str) -> AnalysisArtifacts:
    transcript_data = extract_transcript_data(transcript_path)
    program_hint = transcript_data.get("student", {}).get("program") or transcript_data.get("major_guess")
    plan_data = extract_study_plan_data(study_plan_path, program_hint=program_hint)
    return analyze_transcript_and_study_plan_data(transcript_data, plan_data, output_dir=output_dir)
