import os
import sys
import re
import json
from flask import request, jsonify

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LEGACY_DIR = os.path.join(BASE_DIR, "flask")
if LEGACY_DIR not in sys.path:
    sys.path.insert(0, LEGACY_DIR)

import app as legacy_app_module  # original untouched Flask app
from agents_runtime import load_student_state

app = legacy_app_module.app

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


@app.route("/api/upload-transcript", methods=["POST", "OPTIONS"])
def api_upload_transcript():
    if request.method == "OPTIONS":
        return ("", 204)
    return legacy_app_module.agents_upload()


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
        return jsonify({"response": "📎 Please upload your transcript first."}), 400

    st = load_student_state(session_id)
    if not st:
        return jsonify({"response": "📎 Please upload your transcript first."}), 404

    def clean(courses):
        ignore = {"AIXXX", "C3SXXX", "GHALXXX", "GIASXXX", "GSOSXXX"}
        return [c for c in (courses or []) if c not in ignore]

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

    prompt = f'''
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
'''

    reply = legacy_app_module.generate_llm_response(prompt).strip()
    if "gpa" in message.lower():
        alert_msg = legacy_app_module.get_alert_message(facts)
        if alert_msg:
            reply += f"\n\n{alert_msg}"

    response = jsonify({
        "response": reply,
        "session_id": session_id,
    })
    response.set_cookie(
        "session_id",
        session_id,
        httponly=True,
        samesite=COOKIE_SAMESITE,
        secure=COOKIE_SECURE,
    )
    return response


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
