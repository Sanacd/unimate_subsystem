from __future__ import annotations

import re
from typing import Any


COURSE_CODE_RE = re.compile(r"[A-Z]{2,6}\s*[-]?\s*(?:\d{3}[A-Z]?|X{2,4})", re.I)
MAX_REASONABLE_CREDIT_HOURS = 12.0


def normalize_course_code(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def normalize_course_name(value: Any) -> str:
    text = str(value or "")
    text = re.sub(r"\b(?:pre[- ]?req(?:uisite)?s?|co[- ]?req(?:uisite)?s?)\b.*$", "", text, flags=re.I)
    text = re.sub(r"\b(?:category|type|requirement)\b[:\-]?\s*[A-Za-z /&]+$", "", text, flags=re.I)
    text = re.sub(r"\s+", " ", text).strip(" |-,;")
    return text


def normalize_course_name_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", normalize_course_name(value).lower()).strip()


def normalize_term(term: Any) -> str:
    return re.sub(r"\s+", " ", str(term or "").strip())


def normalize_integer(value: Any) -> int | None:
    if value in (None, ""):
        return None
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        return None


def normalize_credit_hours(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        if 0 < number <= MAX_REASONABLE_CREDIT_HOURS:
            return number
        return None

    text = str(value).strip()
    if not text:
        return None

    matches = re.findall(r"\d+(?:\.\d+)?", text)
    candidates: list[float] = []
    for item in matches:
        try:
            number = float(item)
        except ValueError:
            continue
        if 0 < number <= MAX_REASONABLE_CREDIT_HOURS:
            candidates.append(number)

    if not candidates:
        return None

    return candidates[-1]


def normalize_prerequisites(value: Any) -> list[str]:
    if value in (None, "", "-", "None", "N/A"):
        return []

    if isinstance(value, list):
        items = value
    else:
        items = [value]

    output: list[str] = []
    seen = set()
    for item in items:
        text = str(item or "")
        for match in COURSE_CODE_RE.finditer(text):
            code = normalize_course_code(match.group(0))
            if code and code not in seen:
                seen.add(code)
                output.append(code)
    return output


def normalize_course_record(
    raw: dict[str, Any],
    *,
    source: str,
    default_year: int | None = None,
    default_semester: int | None = None,
    allow_name_only: bool = False,
) -> dict[str, Any] | None:
    code = normalize_course_code(
        raw.get("course_code")
        or raw.get("code")
        or raw.get("Course Code")
        or raw.get("course")
    )
    raw_name_value = (
        raw.get("course_name")
        or raw.get("title")
        or raw.get("Course Name")
        or raw.get("Course Title")
        or raw.get("name")
        or ""
    )
    course_name = normalize_course_name(raw_name_value)

    if not code and not (allow_name_only and course_name):
        return None
    if not course_name and not code:
        return None
    if len(course_name) > 160:
        return None

    prerequisites = normalize_prerequisites(
        raw.get("prerequisites")
        if raw.get("prerequisites") is not None
        else raw.get("prerequisite") or raw.get("Prerequisites")
    )
    credit_hours = normalize_credit_hours(
        raw.get("credit_hours")
        if raw.get("credit_hours") is not None
        else raw.get("credits")
        if raw.get("credits") is not None
        else raw.get("Credit Hours")
        if raw.get("Credit Hours") is not None
        else raw.get("hours")
    )

    category = str(
        raw.get("category")
        or " / ".join(
            part for part in [str(raw.get("type") or "").strip(), str(raw.get("requirement") or "").strip()] if part
        )
        or ""
    ).strip()

    record = {
        "course_code": code,
        "course_name": course_name,
        "credit_hours": credit_hours,
        "prerequisites": prerequisites,
        "year_no": normalize_integer(raw.get("year_no")) or default_year,
        "semester_no": normalize_integer(raw.get("semester_no")) or default_semester,
        "category": category,
        "status": str(raw.get("status") or "").strip(),
        "grade": str(raw.get("grade") or "").strip(),
        "term_taken": normalize_term(raw.get("term_taken")),
        "source": source,
        "notes": str(raw.get("notes") or "").strip(),
    }

    if not record["course_code"] and allow_name_only:
        if record["credit_hours"] is None and len(record["course_name"].split()) > 8:
            return None
        if record["credit_hours"] is None and len(str(raw_name_value or "")) > 80:
            return None

    if not record["course_code"] and not record["course_name"]:
        return None
    return record


def record_match_key(record: dict[str, Any]) -> tuple[str, str]:
    return normalize_course_code(record.get("course_code")), normalize_course_name_key(record.get("course_name"))
