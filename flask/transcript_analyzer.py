import json
import os
import sys

from pdf_extractor import extract_transcript_data


def transform_transcript_for_comparison(raw: dict) -> dict:
    courses = raw.get("courses", [])

    return {
        "student": raw.get("student", {}),
        "courses": [
            {
                "course_code": c.get("course_code", ""),
                "course_name": c.get("course_name", ""),
                "credit_hours": c.get("credit_hours"),
                "grade": c.get("grade", ""),
                "status": c.get("status", ""),
                "term_taken": c.get("term_taken", ""),
                "notes": c.get("notes", ""),
                "points": c.get("points", 0),
            }
            for c in courses
        ],
        "gpa_table": raw.get("gpa_table", []),
        "source_type": "transcript_parser"
    }


def save_transcript_json(input_path: str, output_path: str | None = None) -> dict:
    raw = extract_transcript_data(input_path)
    result = transform_transcript_for_comparison(raw)

    if output_path is None:
        base, _ = os.path.splitext(input_path)
        output_path = base + "_transcript.json"

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"✅ Transcript JSON saved to: {output_path}")
    return result


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python transcript_analyzer.py <transcript_file> [output_json_path]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    save_transcript_json(input_file, output_file)