import json
import os
import re
import sys
import types
import uuid
from typing import Any

import pandas as pd
from flask import Flask, after_this_request, jsonify, request, send_file, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LEGACY_DIR = os.path.join(BASE_DIR, "flask")
if LEGACY_DIR not in sys.path:
    sys.path.insert(0, LEGACY_DIR)

IS_BUNDLED = hasattr(sys, "_MEIPASS")
if IS_BUNDLED:
    appdata = os.getenv("LOCALAPPDATA") or os.path.expanduser("~\\AppData\\Local")
    UPLOAD_FOLDER = os.path.join(appdata, "UniMate", "uploads")
else:
    UPLOAD_FOLDER = os.path.join(LEGACY_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def _levels_to_year_sem(levels_dict: dict) -> dict:
    mapping = {
        "1": ("First Year", "First Semester"),
        "2": ("First Year", "Second Semester"),
        "3": ("Second Year", "First Semester"),
        "4": ("Second Year", "Second Semester"),
        "5": ("Third Year", "First Semester"),
        "6": ("Third Year", "Second Semester"),
        "7": ("Fourth Year", "First Semester"),
        "8": ("Fourth Year", "Second Semester"),
    }
    levels_dict = {str(k): v for k, v in levels_dict.items()}

    structured = {}
    for lvl_key, (year, sem) in mapping.items():
        structured.setdefault(year, {})
        structured[year][sem] = levels_dict.get(lvl_key, [])
    return structured


def _pick_program(programs: dict, major_name: str, df_transcript) -> tuple[str, dict]:
    def norm_title(text: str) -> str:
        return " ".join(str(text or "").lower().strip().split())

    major_norm = norm_title(major_name)
    if major_norm:
        for key in programs.keys():
            if norm_title(key) == major_norm:
                return key, programs[key]

    if major_norm:
        for key in programs.keys():
            if major_norm in norm_title(key):
                return key, programs[key]

    try:
        student_codes = {
            "".join(str(code).upper().split())
            for code in df_transcript["Course Code"].dropna().tolist()
        }
    except Exception:
        student_codes = set()

    best_key, best_hit = None, -1
    for key, plan in programs.items():
        try:
            plan_codes = {
                "".join(course["course_code"].upper().split())
                for level in plan.get("levels", {}).values()
                for course in level
            }
        except Exception:
            continue

        hit = len(student_codes & plan_codes)
        if hit > best_hit:
            best_hit, best_key = hit, key

    if best_key:
        return best_key, programs[best_key]

    default_key = next(iter(programs.keys()))
    return default_key, programs[default_key]


def get_alert_message(facts: dict[str, Any]) -> str | None:
    try:
        gpa = float(facts.get("gpa_ug", 0) or 0)
    except (ValueError, TypeError):
        gpa = 0.0

    try:
        progress = float(facts.get("progress", 0) or 0)
    except (ValueError, TypeError):
        progress = 0.0

    if gpa < 2.0:
        return f"Your GPA is {gpa:.2f}, which is below 2.0. Please meet your academic advisor for a recovery plan."
    if gpa < 2.5:
        return f"Your GPA is {gpa:.2f}. Consider retaking low-grade courses to improve your standing."
    if gpa >= 3.75:
        return f"Excellent work! Your GPA is {gpa:.2f}. You are eligible for the Honor List this semester."

    if 85 <= progress < 100:
        return "You are close to graduation. Make sure you have completed all required electives and co-op or internship."
    if progress == 100:
        return "Congratulations! You have completed all your degree requirements."
    return None


# Provide the helper symbols agents_runtime expects without importing flask/app.py.
app_shim = types.ModuleType("app")
app_shim.UPLOAD_FOLDER = UPLOAD_FOLDER
app_shim._pick_program = _pick_program
app_shim._levels_to_year_sem = _levels_to_year_sem
sys.modules["app"] = app_shim

from agents_runtime import StudentState, load_student_state, save_student_state, ui_agent_handle_upload
from shared_tools import build_chat_fallback, generate_llm_response
from studyplan_analyzer import analyze_transcript_and_study_plan

with open(os.path.join(LEGACY_DIR, "academic_policy.json"), "r", encoding="utf-8") as f:
    ACADEMIC_POLICY = json.load(f)

app = Flask(__name__, static_folder=None)

FRONTEND_ORIGIN = os.environ.get("FRONTEND_ORIGIN", "").strip()
COOKIE_SECURE = os.environ.get("COOKIE_SECURE", "false").lower() == "true"
COOKIE_SAMESITE = os.environ.get("COOKIE_SAMESITE", "Lax")


def _allowed_origin() -> str:
    if FRONTEND_ORIGIN:
        return FRONTEND_ORIGIN
    origin = request.headers.get("Origin", "").strip()
    return origin or "*"


@app.after_request
def add_cors_headers(response):
    origin = _allowed_origin()
    response.headers["Access-Control-Allow-Origin"] = origin
    response.headers["Vary"] = "Origin"
    response.headers["Access-Control-Allow-Credentials"] = "true"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return response


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


def _build_upload_response(file_storage):
    path = os.path.join(UPLOAD_FOLDER, file_storage.filename)
    file_storage.save(path)

    session_id, student, plan, ui_note = ui_agent_handle_upload(
        path,
        user_id="local_user",
        reasoning_mode=request.args.get("mode", "react+reflexion"),
        hitl=False,
    )

    student_state = StudentState(
        session_id=session_id,
        name=getattr(student, "name", None),
        major=getattr(plan, "major", None),
        gpa_prep=plan.summary.get("gpa_prep"),
        gpa_undergrad=plan.summary.get("gpa_undergrad") or plan.summary.get("gpa_final"),
        credits_prep=plan.summary.get("credit_prep"),
        credits_undergrad=plan.summary.get("credit_ug"),
        credits_in_progress=plan.summary.get("credit_inprogress"),
        progress_percent=plan.summary.get("progress_percent"),
        completed_courses=plan.summary.get("completed_courses", []),
        in_progress_courses=plan.summary.get("in_progress_courses", []),
        remaining_courses=plan.summary.get("remaining_courses", []),
    )
    save_student_state(student_state)

    try:
        gpa_final = float(plan.summary.get("gpa_undergrad") or plan.summary.get("gpa_final") or 0)
    except Exception:
        gpa_final = 0.0

    try:
        progress = float(plan.summary.get("progress_percent") or 0)
    except Exception:
        progress = 0.0

    alert_type = None
    alert_message = None
    standing = ""

    rules = ACADEMIC_POLICY.get("gpa_rules", {})
    honors = ACADEMIC_POLICY.get("honors", {})
    grad_req = ACADEMIC_POLICY.get("graduation_requirements", {})

    if "fail" in rules and gpa_final <= rules["fail"].get("max", 1.99):
        alert_type = "warning"
        standing = "Academic Warning"
        alert_message = f"GPA {gpa_final:.2f}: Below {rules['pass']['min']:.2f}. You are at risk of probation."
    elif "pass" in rules and rules["pass"]["min"] <= gpa_final <= rules["pass"]["max"]:
        alert_type = "encourage"
        standing = "Pass Standing"
        alert_message = f"GPA {gpa_final:.2f}: Minimum passing range. Focus on steady improvement."
    elif "good" in rules and rules["good"]["min"] <= gpa_final <= rules["good"]["max"]:
        alert_type = "encourage"
        standing = "Good Standing"
        alert_message = f"GPA {gpa_final:.2f}: Good academic standing. Keep pushing for higher distinction."
    elif "very_good" in rules and rules["very_good"]["min"] <= gpa_final <= rules["very_good"]["max"]:
        alert_type = "encourage"
        standing = "Very Good Standing"
        alert_message = f"GPA {gpa_final:.2f}: Excellent consistency. You're close to First-Class Honors."
    elif "excellent" in rules and rules["excellent"]["min"] <= gpa_final <= rules["excellent"]["max"]:
        alert_type = "excellent"
        standing = "First-Class Honor"
        alert_message = f"GPA {gpa_final:.2f}: Outstanding performance. You qualify for First-Class Honors."

    if alert_message and gpa_final < grad_req.get("minimum_gpa", 2.0):
        alert_message += " GPA below graduation threshold."
    elif alert_message and gpa_final >= honors.get("first_class", {}).get("min", 999):
        alert_message += " Eligible for First-Class Honors."
    elif alert_message and gpa_final >= honors.get("second_class", {}).get("min", 999):
        alert_message += " Eligible for Second-Class Honors."

    if alert_message and 85 <= progress < 100:
        alert_message += " You are close to graduation. Ensure electives and co-op are completed."
    elif alert_message and progress == 100:
        alert_message += " Congratulations! You have completed all degree requirements."

    if not alert_message:
        alert_type = "info"
        alert_message = f"GPA {gpa_final:.2f}: Keep striving for excellence!"

    excel_url = f"/download-report/{os.path.basename(plan.excel_path)}" if plan.excel_path else None

    response = jsonify(
        {
            "session_id": session_id,
            "major": plan.major,
            "excel_path": excel_url,
            "ui_summary": ui_note,
            "progress_percent": plan.summary.get("progress_percent"),
            "remaining_courses": plan.summary.get("remaining_courses", []),
            "gpa": gpa_final,
            "alert_type": alert_type,
            "alert": alert_message,
        }
    )
    response.set_cookie(
        "session_id",
        session_id,
        httponly=True,
        samesite=COOKIE_SAMESITE,
        secure=COOKIE_SECURE,
    )
    return response


def _save_uploaded_file(file_storage, label: str) -> str:
    ext = os.path.splitext(file_storage.filename or "")[1] or ".bin"
    safe_name = f"{label}_{uuid.uuid4().hex}{ext}"
    path = os.path.join(UPLOAD_FOLDER, safe_name)
    file_storage.save(path)
    return path


@app.route("/api/upload-transcript", methods=["POST", "OPTIONS"])
def api_upload_transcript():
    if request.method == "OPTIONS":
        return ("", 204)

    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file uploaded"}), 400
    return _build_upload_response(file)


@app.route("/api/analyze-study-plan", methods=["POST", "OPTIONS"])
def api_analyze_study_plan():
    if request.method == "OPTIONS":
        return ("", 204)

    transcript_file = request.files.get("transcript")
    study_plan_file = request.files.get("study_plan")

    if not transcript_file:
        return jsonify({"success": False, "error": "Transcript file is required."}), 400
    if not study_plan_file:
        return jsonify({"success": False, "error": "Study plan file is required."}), 400

    transcript_path = _save_uploaded_file(transcript_file, "transcript")
    study_plan_path = _save_uploaded_file(study_plan_file, "study_plan")

    try:
        artifacts = analyze_transcript_and_study_plan(
            transcript_path=transcript_path,
            study_plan_path=study_plan_path,
            output_dir=UPLOAD_FOLDER,
        )
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"success": False, "error": f"Study plan analysis failed: {exc}"}), 500

    return jsonify(
        {
            "success": True,
            "student": artifacts.student,
            "summary": artifacts.summary,
            "eligible_next_semester": artifacts.eligible_next_semester,
            "advice": artifacts.advice,
            "excel_file": f"/download-report/{os.path.basename(artifacts.excel_path)}",
            "preview_rows": artifacts.preview_rows,
        }
    )


@app.route("/agents/upload", methods=["POST"])
def agents_upload():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file uploaded"}), 400
    return _build_upload_response(file)


@app.route("/api/chat", methods=["POST", "OPTIONS"])
def api_chat():
    if request.method == "OPTIONS":
        return ("", 204)

    data = request.get_json(force=True) or {}
    message = (data.get("message") or "").strip()
    session_id = (data.get("session_id") or request.cookies.get("session_id") or "").strip()

    if not message:
        return jsonify({"response": "Please enter a question."}), 400
    if not session_id:
        return jsonify({"response": "Please upload your transcript first."}), 400

    st = load_student_state(session_id)
    if not st:
        return jsonify({"response": "Please upload your transcript first."}), 404

    def clean(courses):
        ignore = {"AIXXX", "C3SXXX", "GHALXXX", "GIASXXX", "GSOSXXX"}
        return [course for course in (courses or []) if course not in ignore]

    facts = {
        "name": st.name or "Student",
        "major": st.major or "Not detected",
        "gpa_prep": st.gpa_prep,
        "gpa_ug": st.gpa_undergrad,
        "progress": st.progress_percent,
        "completed": clean(st.completed_courses),
        "remaining": clean(st.remaining_courses),
        "in_progress": clean(st.in_progress_courses),
    }

    is_ar = bool(re.search(r"[\u0600-\u06FF]", message))
    lang = "Arabic" if is_ar else "English"

    prompt = f"""
You are UniMate, the official academic advisor of Prince Muqrin University.

Student Data:
{json.dumps(facts, ensure_ascii=False, indent=2)}

User question:
"{message}"

Rules:
- Respond only in {lang}
- Use ONLY the above student data
- Do NOT invent course names
- If asked about remaining courses, list them
- If asked GPA, use stored values
- If question not academic, give supportive response
- Be concise, friendly, professional

Answer:
"""

    reply = generate_llm_response(prompt, fallback_text=build_chat_fallback(facts, message)).strip()
    if "gpa" in message.lower():
        alert_msg = get_alert_message(facts)
        if alert_msg:
            reply += f"\n\n{alert_msg}"

    response = jsonify({"response": reply, "session_id": session_id})
    response.set_cookie(
        "session_id",
        session_id,
        httponly=True,
        samesite=COOKIE_SAMESITE,
        secure=COOKIE_SECURE,
    )
    return response


@app.route("/download-report/<filename>")
def download_report(filename):
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(file_path):
        return "File not found", 404

    @after_this_request
    def cleanup(response):
        return response

    return send_file(
        file_path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/uploads/<path:filename>")
def serve_upload(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
