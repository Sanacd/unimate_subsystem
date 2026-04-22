from __future__ import annotations

import json
import os
import re
import uuid
import zipfile
from xml.etree import ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass
from functools import lru_cache
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from pdf_extractor import extract_text, extract_transcript_data


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

STATUS_LABELS = {
    "completed": "Completed",
    "in_progress": "In Progress",
    "not_taken": "Not Taken / Incomplete",
    "failed": "Failed / Retake Required",
    "blocked": "Blocked by prerequisite",
}

STATUS_FILLS = {
    "completed": "C6EFCE",
    "in_progress": "FFF2CC",
    "not_taken": "D9D9D9",
    "failed": "F4CCCC",
    "blocked": "FCE5CD",
}

TERM_ORDER = {"spring": 1, "summer": 2, "fall": 3}

COURSE_CODE_RE = re.compile(r"[A-Z]{2,6}\s*[-]?\s*(?:\d{3}[A-Z]?|X{2,4})", re.I)
YEAR_WORDS = {"first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}


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


def normalize_course_code(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def normalize_course_name(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()


def normalize_term(term: Any) -> str:
    return re.sub(r"\s+", " ", str(term or "").strip())


def classify_transcript_status(grade: Any, points: Any) -> str:
    grade_text = str(grade or "").strip().upper()

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
        "completed": 4,
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

    normalized_hint = normalize_course_name(program_hint)
    programs = load_catalog_study_plans()

    for key, value in programs.items():
        if normalize_course_name(key) == normalized_hint:
            return key, value
    for key, value in programs.items():
        if normalized_hint and normalized_hint in normalize_course_name(key):
            return key, value
    return None, None


def parse_prerequisites(value: Any) -> list[str]:
    text = str(value or "").strip()
    if not text or text in {"-", "None", "N/A"}:
        return []

    matches = [normalize_course_code(match.group(0)) for match in COURSE_CODE_RE.finditer(text)]
    deduped = []
    seen = set()
    for item in matches:
        if item and item not in seen:
            seen.add(item)
            deduped.append(item)
    return deduped


def _guess_category(raw: dict[str, Any]) -> str:
    category = str(raw.get("category") or "").strip()
    if category:
        return category

    parts = [
        str(raw.get("type") or "").strip(),
        str(raw.get("requirement") or "").strip(),
    ]
    category = " / ".join(part for part in parts if part)
    return category or "Uncategorized"


def _infer_year_semester_from_level(level_value: Any) -> tuple[int | None, int | None]:
    try:
        level_num = int(level_value)
    except (TypeError, ValueError):
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
        chunks = []
        for node in paragraph.findall(".//w:t", namespaces):
            if node.text:
                chunks.append(node.text)
        text = "".join(chunks).strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs)


def _extract_plain_text(file_path: str) -> str:
    with open(file_path, "r", encoding="utf-8", errors="ignore") as fh:
        return fh.read()


def _parse_year_semester_from_text(line: str) -> tuple[int | None, int | None]:
    lowered = line.lower()
    year_no = None
    semester_no = None

    level_match = re.search(r"\blevel\s*(\d+)\b", lowered)
    if level_match:
        return _infer_year_semester_from_level(level_match.group(1))

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


def _infer_category_from_line(line: str) -> str:
    lowered = line.lower()
    if "elective" in lowered:
        return "Elective"
    if "university" in lowered or "institution" in lowered:
        return "Institution"
    if "college" in lowered:
        return "College"
    if "department" in lowered:
        return "Department"
    return "Uncategorized"


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
    prereqs = []
    working_text = remainder
    if prereq_match:
        prereqs = parse_prerequisites(prereq_match.group(1))
        working_text = remainder[:prereq_match.start()].strip(" -|,")

    credits = 0.0
    credit_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:credit|credits|cr|hrs|hours)?\s*$", working_text, re.I)
    if credit_match:
        try:
            credits = float(credit_match.group(1))
        except (TypeError, ValueError):
            credits = 0.0
        title_part = working_text[:credit_match.start()].strip(" -|,")
    else:
        title_part = working_text

    title_part = re.sub(r"\b(?:category|type)\b[:\-]?\s*[A-Za-z /&]+$", "", title_part, flags=re.I).strip(" -|,")
    if not title_part:
        title_part = code

    return {
        "course_code": code,
        "course_name": title_part,
        "credits": credits,
        "category": _infer_category_from_line(compact),
        "year_no": year_no,
        "semester_no": semester_no,
        "prerequisites": prereqs,
        "notes": "Best-effort text extraction",
    }


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
        parsed = _parse_course_line(line, current_year, current_semester)
        if parsed is None:
            continue
        courses.append(parsed)

    return {
        "program_name": _extract_program_name(text, program_hint=program_hint),
        "catalog_year": _extract_catalog_year(text),
        "courses": courses,
        "slot_rules": {},
        "source_type": source_type,
        "ocr_ready": source_type == "image",
    }


def _normalize_plan_rows(rows: list[dict[str, Any]], level: Any = None) -> list[dict[str, Any]]:
    normalized = []
    default_year, default_semester = _infer_year_semester_from_level(level)

    for row in rows:
        code = normalize_course_code(
            row.get("course_code")
            or row.get("code")
            or row.get("Course Code")
            or row.get("course")
        )
        name = (
            row.get("course_name")
            or row.get("title")
            or row.get("Course Name")
            or row.get("name")
            or ""
        )
        if not code and not name:
            continue

        credits_raw = (
            row.get("credits")
            or row.get("credit_hours")
            or row.get("Credit Hours")
            or row.get("hours")
            or 0
        )
        try:
            credits = float(credits_raw or 0)
        except (TypeError, ValueError):
            credits = 0.0

        year_no = row.get("year_no")
        semester_no = row.get("semester_no")
        try:
            year_no = int(year_no) if year_no not in (None, "") else default_year
        except (TypeError, ValueError):
            year_no = default_year
        try:
            semester_no = int(semester_no) if semester_no not in (None, "") else default_semester
        except (TypeError, ValueError):
            semester_no = default_semester

        prerequisite_text = row.get("prerequisites")
        if prerequisite_text is None:
            prerequisite_text = row.get("prerequisite") or row.get("Prerequisites") or ""

        normalized.append(
            {
                "course_code": code,
                "course_name": str(name).strip(),
                "credits": credits,
                "category": _guess_category(row),
                "year_no": year_no,
                "semester_no": semester_no,
                "status": "",
                "grade": "",
                "term_taken": "",
                "prerequisites": parse_prerequisites(prerequisite_text),
                "notes": str(row.get("notes") or "").strip(),
            }
        )

    return normalized


def _normalize_program_structure(program_name: str, program_data: dict[str, Any]) -> dict[str, Any]:
    courses = []
    levels = program_data.get("levels", {})
    for level, rows in levels.items():
        courses.extend(_normalize_plan_rows(rows, level=level))

    slot_rules = {}
    raw_slot_rules = program_data.get("slot_rules", {})
    for slot, rule in raw_slot_rules.items():
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


def extract_study_plan_data(file_path: str, program_hint: str | None = None) -> dict[str, Any]:
    file_type = detect_study_plan_file_type(file_path)

    if file_type == "pdf":
        text = extract_text(file_path)
        plan = _parse_study_plan_text(text, "pdf", program_hint=program_hint)
    elif file_type == "docx":
        text = _extract_docx_text(file_path)
        plan = _parse_study_plan_text(text, "docx", program_hint=program_hint)
    elif file_type == "txt":
        text = _extract_plain_text(file_path)
        plan = _parse_study_plan_text(text, "txt", program_hint=program_hint)
    elif file_type == "xlsx":
        excel_book = pd.read_excel(file_path, sheet_name=None)
        structured_rows = []
        for sheet_name, df in excel_book.items():
            sheet_rows = df.fillna("").to_dict("records")
            normalized_rows = _normalize_plan_rows(sheet_rows)
            for row in normalized_rows:
                if not row.get("notes"):
                    row["notes"] = f"Imported from sheet: {sheet_name}"
            structured_rows.extend(normalized_rows)
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
            "courses": _normalize_plan_rows(df.to_dict("records")),
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
            "notes": "Image study plan uploaded. OCR is not required in the current deployment, so extraction is deferred unless a catalog fallback is available.",
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
            plan = _normalize_program_structure(
                program_hint or raw.get("program_name") or raw.get("name") or "",
                raw,
            )
        elif isinstance(raw, list):
            plan = {
                "program_name": program_hint or "",
                "catalog_year": "",
                "courses": _normalize_plan_rows(raw),
                "slot_rules": {},
                "source_type": "json_rows",
            }
        else:
            plan = {
                "program_name": raw.get("program_name") or program_hint or "",
                "catalog_year": str(raw.get("catalog_year") or raw.get("version") or "").strip(),
                "courses": _normalize_plan_rows(raw.get("courses", [])),
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


def _make_course_key(course: dict[str, Any]) -> tuple[str, str]:
    return normalize_course_code(course.get("course_code")), normalize_course_name(course.get("course_name"))


def _build_transcript_lookup(transcript_data: dict[str, Any]) -> dict[str, dict[str, Any]]:
    by_code: dict[str, list[dict[str, Any]]] = defaultdict(list)
    by_name: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for course in transcript_data.get("courses", []):
        by_code[normalize_course_code(course.get("course_code"))].append(course)
        by_name[normalize_course_name(course.get("course_name"))].append(course)

    best_attempts = {}
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
            if item.get("status") == "completed"
        },
        "in_progress_codes": {
            normalize_course_code(item.get("course_code"))
            for item in transcript_data.get("courses", [])
            if item.get("status") == "in_progress"
        },
    }


def _candidate_matches_for_slot(
    slot_code: str,
    slot_rule: dict[str, Any],
    transcript_courses: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    candidates = []
    specifics = set(slot_rule.get("specific", []))
    prefixes = tuple(slot_rule.get("prefixes", []))

    for course in transcript_courses:
        course_code = normalize_course_code(course.get("course_code"))
        if course_code == slot_code:
            continue
        if course_code in specifics or (prefixes and any(course_code.startswith(prefix) for prefix in prefixes)):
            candidates.append(course)

    return sorted(candidates, key=_attempt_sort_key, reverse=True)


def _prereq_codes_satisfied(prerequisites: list[str], satisfied_codes: set[str]) -> tuple[bool, list[str]]:
    missing = [item for item in prerequisites if normalize_course_code(item) not in satisfied_codes]
    return (len(missing) == 0, missing)


def _match_study_plan_courses(plan_data: dict[str, Any], transcript_data: dict[str, Any]) -> list[dict[str, Any]]:
    lookup = _build_transcript_lookup(transcript_data)
    slot_rules = plan_data.get("slot_rules", {})
    used_slot_matches: set[str] = set()
    satisfied_for_prereqs = set(lookup["completed_codes"]) | set(lookup["in_progress_codes"])
    transcript_courses = transcript_data.get("courses", [])
    matched_rows = []

    for course in plan_data.get("courses", []):
        row = dict(course)
        course_code = normalize_course_code(row.get("course_code"))
        course_name = normalize_course_name(row.get("course_name"))
        matched_attempt = lookup["best_attempts"].get(course_code) or lookup["best_attempts"].get(course_name)
        notes: list[str] = []

        if matched_attempt is None and "X" in course_code:
            candidates = _candidate_matches_for_slot(course_code, slot_rules.get(course_code, {}), transcript_courses)
            for candidate in candidates:
                candidate_code = normalize_course_code(candidate.get("course_code"))
                if candidate_code in used_slot_matches:
                    continue
                matched_attempt = candidate
                used_slot_matches.add(candidate_code)
                notes.append(f"Matched elective slot with {candidate_code}")
                break

        status = "not_taken"
        grade = ""
        term_taken = ""

        if matched_attempt:
            status = matched_attempt.get("status") or "not_taken"
            grade = str(matched_attempt.get("grade") or "").strip()
            term_taken = normalize_term(matched_attempt.get("term_taken"))

            matched_code = normalize_course_code(matched_attempt.get("course_code"))
            if matched_code in lookup["repeated_codes"]:
                notes.append("Repeated course detected in transcript history")
            if matched_code in lookup["failed_codes"] and status == "completed":
                notes.append("Completed after at least one failed attempt")

        prereq_ok, missing_prereqs = _prereq_codes_satisfied(row.get("prerequisites", []), satisfied_for_prereqs)
        if status == "not_taken" and missing_prereqs:
            status = "blocked"
            notes.append("Missing prerequisites: " + ", ".join(missing_prereqs))
        elif status == "failed" and prereq_ok:
            notes.append("Eligible for retake when offered")

        row["status"] = status
        row["grade"] = grade
        row["term_taken"] = term_taken
        row["notes"] = "; ".join(filter(None, [row.get("notes"), *notes])).strip("; ")
        matched_rows.append(row)

    return matched_rows


def _compute_summary(rows: list[dict[str, Any]], transcript_data: dict[str, Any]) -> dict[str, Any]:
    credits_required = round(sum(float(row.get("credits") or 0) for row in rows), 2)
    credits_completed = round(sum(float(row.get("credits") or 0) for row in rows if row.get("status") == "completed"), 2)
    credits_in_progress = round(sum(float(row.get("credits") or 0) for row in rows if row.get("status") == "in_progress"), 2)
    credits_remaining = round(sum(float(row.get("credits") or 0) for row in rows if row.get("status") != "completed"), 2)
    completion_percentage = round((credits_completed / credits_required) * 100, 2) if credits_required else 0.0

    return {
        "credits_required": credits_required,
        "credits_completed": credits_completed,
        "credits_in_progress": credits_in_progress,
        "credits_remaining": credits_remaining,
        "completion_percentage": completion_percentage,
        "courses_completed": sum(1 for row in rows if row.get("status") == "completed"),
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
        prereqs = row.get("prerequisites", [])
        note = "Retake required" if row.get("status") == "failed" else "Prerequisites satisfied"
        if prereqs:
            note = f"{note}: {', '.join(prereqs)}"
        eligible.append(
            {
                "course_code": row.get("course_code"),
                "course_name": row.get("course_name"),
                "credits": row.get("credits"),
                "year_no": row.get("year_no"),
                "semester_no": row.get("semester_no"),
                "note": note,
            }
        )

    eligible.sort(key=lambda item: ((item.get("year_no") or 99), (item.get("semester_no") or 99), item.get("course_code") or ""))
    return eligible


def _generate_advice(
    rows: list[dict[str, Any]],
    summary: dict[str, Any],
    transcript_data: dict[str, Any],
    student: dict[str, Any],
) -> list[str]:
    policy = load_academic_policy()
    advice = []

    failed = [row for row in rows if row.get("status") == "failed"]
    blocked = [row for row in rows if row.get("status") == "blocked"]
    in_progress = [row for row in rows if row.get("status") == "in_progress"]
    eligible = _eligible_next_semester(rows)

    if failed:
        advice.append(
            f"Prioritize retaking {', '.join(row['course_code'] for row in failed[:4])} because failed courses still count toward remaining credits."
        )
    if blocked:
        advice.append(
            f"Unblock {len(blocked)} course(s) by finishing prerequisite chains first, starting with {', '.join(row['course_code'] for row in blocked[:4])}."
        )
    if eligible:
        advice.append(
            f"You can plan next-semester registration around {', '.join(item['course_code'] for item in eligible[:6])} based on satisfied prerequisites."
        )
    if in_progress:
        advice.append(
            f"Protect your current load: {len(in_progress)} course(s) are in progress and will unlock more options once completed."
        )

    gpa_value = student.get("gpa_final")
    min_gpa = float(policy.get("graduation_requirements", {}).get("minimum_gpa", 2.0))
    if isinstance(gpa_value, (int, float)) and gpa_value < min_gpa:
        advice.append(
            f"Your cumulative GPA is {gpa_value:.2f}, below the graduation minimum of {min_gpa:.2f}; use retakes and lighter prerequisite sequencing to recover."
        )

    if summary["completion_percentage"] >= 85:
        advice.append("You are close to program completion; verify capstone, internship/co-op, and elective slots early so graduation is not delayed.")
    elif summary["completion_percentage"] < 40:
        advice.append("You are still in the early program phase; focus on foundational prerequisite courses to maximize future scheduling flexibility.")

    if not advice:
        advice.append("Your plan is broadly on track. Keep finishing in-progress courses and prioritize the earliest unlocked required courses next.")

    return advice


def build_student_snapshot(transcript_data: dict[str, Any], plan_data: dict[str, Any]) -> dict[str, Any]:
    student = dict(transcript_data.get("student", {}))
    program = student.get("program") or plan_data.get("program_name") or ""
    student["program"] = program
    student["study_plan_program"] = plan_data.get("program_name") or program
    student["study_plan_catalog"] = plan_data.get("catalog_year") or ""
    return student


def build_study_plan_audit_workbook(
    artifacts: AnalysisArtifacts,
    rows: list[dict[str, Any]],
) -> Workbook:
    wb = Workbook()
    ws_audit = wb.active
    ws_audit.title = "Study Plan Audit"
    ws_summary = wb.create_sheet("Summary")
    ws_eligible = wb.create_sheet("Eligible Next Semester")
    ws_advice = wb.create_sheet("Personalized Advice")

    audit_headers = [
        "Year",
        "Semester",
        "Course Code",
        "Course Name",
        "Credits",
        "Category",
        "Prerequisites",
        "Status",
        "Grade",
        "Term Taken",
        "Notes",
    ]
    ws_audit.append(audit_headers)

    for header_cell in ws_audit[1]:
        header_cell.font = Font(bold=True)
        header_cell.fill = PatternFill("solid", fgColor="D9EAD3")

    for row in rows:
        ws_audit.append(
            [
                row.get("year_no"),
                row.get("semester_no"),
                row.get("course_code"),
                row.get("course_name"),
                row.get("credits"),
                row.get("category"),
                ", ".join(row.get("prerequisites", [])),
                STATUS_LABELS.get(row.get("status"), row.get("status")),
                row.get("grade"),
                row.get("term_taken"),
                row.get("notes"),
            ]
        )
        status_cell = ws_audit.cell(row=ws_audit.max_row, column=8)
        status_key = row.get("status")
        status_cell.fill = PatternFill("solid", fgColor=STATUS_FILLS.get(status_key, "FFFFFF"))

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
                item.get("credits"),
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


def export_study_plan_audit_excel(
    artifacts: AnalysisArtifacts,
    rows: list[dict[str, Any]],
    output_dir: str,
) -> str:
    os.makedirs(output_dir, exist_ok=True)
    filename = f"study_plan_audit_{uuid.uuid4().hex}.xlsx"
    output_path = os.path.join(output_dir, filename)
    wb = build_study_plan_audit_workbook(artifacts, rows)
    wb.save(output_path)
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
    rows = _match_study_plan_courses(plan_data, transcript_data)
    summary = _compute_summary(rows, transcript_data)
    advice = _generate_advice(rows, summary, transcript_data, transcript_data.get("student", {}))
    eligible = _eligible_next_semester(rows)
    preview_rows = rows[:10]

    temp_artifacts = AnalysisArtifacts(
        student=student,
        summary=summary,
        eligible_next_semester=eligible,
        advice=advice,
        preview_rows=preview_rows,
        excel_path="",
    )
    excel_path = ""
    if output_dir:
        excel_path = export_study_plan_audit_excel(temp_artifacts, rows, output_dir)

    return AnalysisArtifacts(
        student=student,
        summary=summary,
        eligible_next_semester=eligible,
        advice=advice,
        preview_rows=preview_rows,
        excel_path=excel_path,
    )


def analyze_transcript_and_study_plan(
    transcript_path: str,
    study_plan_path: str,
    output_dir: str,
) -> AnalysisArtifacts:
    transcript_data = extract_transcript_data(transcript_path)
    program_hint = transcript_data.get("student", {}).get("program") or transcript_data.get("major_guess")
    plan_data = extract_study_plan_data(study_plan_path, program_hint=program_hint)
    return analyze_transcript_and_study_plan_data(transcript_data, plan_data, output_dir=output_dir)
