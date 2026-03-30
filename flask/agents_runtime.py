# -*- coding: utf-8 -*-
"""
agents_runtime.py — Verbose Dev Mode (C)
✅ Compatible with app.py (2025-11-01)
✅ GPA summary injected correctly into compare_transcript_with_plan via extracted_summary_df
✅ Explicit debug prints for GPA table and key steps
"""

from __future__ import annotations

import os
import re
import json
import time
import uuid
import sqlite3
import pathlib
from dataclasses import dataclass, asdict
from typing import Any, Dict, List, Optional, Tuple, Callable

# ---- Third-party ----
import pandas as pd  # (fixed typo: panadas -> pandas)

# ---- Local tools (must exist in backend) ----
from shared_tools import (
    compare_transcript_with_plan,
    generate_structured_study_plan_excel,
    generate_llm_response,
)

from pdf_extractor import (
    extract_text,
    clean_text,
    parse_student_info_v3,
    parse_courses_with_multiline_fix,
    smart_fix_titles,            # not used here but kept for parity with app
    split_by_category_v4,
    extract_gpa_summary_v3,
    create_summary,
    split_academic_blocks,
)

from dataclasses import dataclass, asdict, field
from typing import List, Optional

# =========================================
# 🧠 Student Session Schema (Unified)
# =========================================
from dataclasses import dataclass, field
from typing import List, Optional

@dataclass
class StudentState:
    """Persistent normalized student session (used for chat, memory, and UI)."""
    session_id: Optional[str] = None
    name: Optional[str] = None
    major: Optional[str] = None

    # GPA & Credits
    gpa_prep: Optional[float] = None
    gpa_undergrad: Optional[float] = None
    credits_prep: Optional[float] = None
    credits_undergrad: Optional[float] = None
    credits_in_progress: Optional[float] = None

    # Academic progress
    progress_percent: Optional[float] = None

    # Course lists
    completed_courses: List[str] = field(default_factory=list)
    in_progress_courses: List[str] = field(default_factory=list)
    remaining_courses: List[str] = field(default_factory=list)

    # Optional meta
    timestamp: Optional[float] = None

    # -----------------------------------
    # Conversion helpers
    # -----------------------------------
    @classmethod
    def from_db(cls, session_id: str, state_dict: dict):
        """Load student state from _get_state() dictionary."""
        return cls(
            session_id=session_id,
            name=state_dict.get("student_name"),
            major=state_dict.get("major"),
            gpa_prep=state_dict.get("gpa_prep"),
            gpa_undergrad=state_dict.get("gpa_undergrad"),
            credits_prep=state_dict.get("credits_prep"),
            credits_undergrad=state_dict.get("credits_undergrad"),
            credits_in_progress=state_dict.get("credits_in_progress"),
            progress_percent=state_dict.get("progress_percent"),
            completed_courses=state_dict.get("completed_courses", []),
            in_progress_courses=state_dict.get("in_progress_courses", []),
            remaining_courses=state_dict.get("remaining_courses", []),
        )

    def to_dict(self):
        """Convert dataclass to dictionary for JSON or saving."""
        return asdict(self)


# ---- Load study plans ----
_STUDY_PLANS_PATH = pathlib.Path(__file__).with_name("university_studyplans.json")
with open(_STUDY_PLANS_PATH, "r", encoding="utf-8") as _fp:
    _SP_RAW = json.load(_fp)
STUDY_PLANS: Dict[str, Any] = _SP_RAW.get("programs", _SP_RAW)  # supports both schemas

# ---- DB path ----
DB_PATH = os.environ.get("AGENTS_DB", "agent_runtime.db")


# =========================================
# DB / State Helpers
# =========================================
def _db():
    """Open DB and ensure tables exist."""
    conn = sqlite3.connect(DB_PATH)

    conn.execute("""CREATE TABLE IF NOT EXISTS sessions(
        id TEXT PRIMARY KEY,
        created_at REAL,
        user_id TEXT,
        reasoning_mode TEXT,
        hitl_enabled INTEGER,
        mode TEXT
    )""")

    conn.execute("""CREATE TABLE IF NOT EXISTS events(
        id TEXT PRIMARY KEY,
        session_id TEXT,
        ts REAL,
        agent TEXT,
        type TEXT,
        payload_json TEXT
    )""")

    conn.execute("""CREATE TABLE IF NOT EXISTS state(
        id TEXT PRIMARY KEY,
        session_id TEXT,
        key TEXT,
        value_json TEXT
    )""")

    conn.execute("""CREATE TABLE IF NOT EXISTS artifacts(
        id TEXT PRIMARY KEY,
        session_id TEXT,
        kind TEXT,
        path TEXT,
        meta_json TEXT
    )""")

    conn.execute("""CREATE TABLE IF NOT EXISTS hitl_queue(
        id TEXT PRIMARY KEY,
        session_id TEXT,
        agent TEXT,
        action TEXT,
        payload_json TEXT,
        status TEXT,
        created_at REAL
    )""")

    return conn


def log_event(session_id: str, agent: str, etype: str, payload: Dict[str, Any]):
    conn = _db()
    conn.execute(
        "INSERT INTO events VALUES(?,?,?,?,?,?)",
        (str(uuid.uuid4()), session_id, time.time(), agent, etype, json.dumps(payload, ensure_ascii=False)),
    )
    conn.commit()
    conn.close()


def set_state(session_id: str, key: str, value: Dict[str, Any]):
    conn = _db()
    conn.execute(
        "INSERT OR REPLACE INTO state VALUES(?,?,?,?)",
        (f"{session_id}:{key}", session_id, key, json.dumps(value, ensure_ascii=False)),
    )
    conn.commit()
    conn.close()


def get_state(session_id: str, key: str) -> Optional[Dict[str, Any]]:
    conn = _db()
    cur = conn.execute("SELECT value_json FROM state WHERE session_id=? AND key=?", (session_id, key))
    row = cur.fetchone()
    conn.close()
    return json.loads(row[0]) if row else None


def _save_state(session_id: str, key: str, value: Any):
    """Save any Python object safely as JSON."""
    conn = _db()
    # Serialize once only (avoid nested JSON)
    if isinstance(value, (dict, list)):
        value_json = json.dumps(value, ensure_ascii=False)
    else:
        value_json = json.dumps(str(value), ensure_ascii=False)
    conn.execute(
        "INSERT OR REPLACE INTO state (id, session_id, key, value_json) VALUES (?, ?, ?, ?)",
        (f"{session_id}-{key}", session_id, key, value_json),
    )
    conn.commit()
    conn.close()


def _get_state(session_id: str) -> Dict[str, Any]:
    """Return all key/value pairs as real Python objects."""
    conn = _db()
    cur = conn.execute("SELECT key, value_json FROM state WHERE session_id = ?", (session_id,))
    data = {}
    for key, value_json in cur.fetchall():
        try:
            val = json.loads(value_json)
        except json.JSONDecodeError:
            val = value_json
        # If value was double-encoded, decode again
        if isinstance(val, str) and val.startswith("[") and val.endswith("]"):
            try:
                val = json.loads(val)
            except Exception:
                pass
        data[key] = val
    conn.close()
    return data



# =========================================
# 🧩 Student Session Helpers
# =========================================

def load_student_state(session_id: str) -> Optional[StudentState]:
    """Return a StudentState object for the given session_id."""
    from agents_runtime import _get_state
    try:
        data = _get_state(session_id)
        if not data:
            return None
        return StudentState.from_db(session_id, data)
    except Exception as e:
        print(f"[load_student_state] Error: {e}")
        return None


def save_student_state(student: StudentState):
    """Persist the StudentState object into DB."""
    from agents_runtime import _save_state
    if not student.session_id:
        raise ValueError("StudentState must have a session_id")
    for k, v in student.to_dict().items():
        if k != "session_id":
            _save_state(student.session_id, k, v)

# =========================================
# Data Schemas
# =========================================
@dataclass
class TranscriptState:
    """Temporary structure for transcript analyzer output (used before session normalization)."""
    info: Dict[str, Any]
    courses_df: List[Dict[str, Any]]    # df_all.to_dict(orient="records")
    gpa_table: List[Dict[str, Any]]     # Not strictly required but handy for debugging
    derived: Dict[str, Any]             # {"summary_df": [...]} used to feed GPA table to compare function
    major_guess: Optional[str] = None


@dataclass
class PlanState:
    """Study plan comparison + Excel export state."""
    major: str
    summary: Dict[str, Any]
    excel_path: Optional[str] = None
    pending: Optional[str] = None   # kept for HITL flows


# =========================================
# ReAct Wrapper (auto, verbose)
# =========================================
def _react(agent_name: str, objective: str,
           think: Callable[[], Dict[str, Any]],
           act: Callable[[Dict[str, Any]], Dict[str, Any]],
           session_id: str,
           allow_reflexion: bool = False,
           max_retries: int = 1) -> Dict[str, Any]:

    print(f"\n==== [{agent_name}] Objective: {objective} ====")
    log_event(session_id, agent_name, "thought_summary", {"objective": objective, "phase": "start"})

    plan = think()
    print(f"[{agent_name}] Plan:", plan)
    log_event(session_id, agent_name, "thought_summary", {"plan": plan})

    observation = act(plan)
    print(f"[{agent_name}] Observation keys:", list(observation.keys()))
    log_event(session_id, agent_name, "observation", {"result_keys": list(observation.keys())})

    if allow_reflexion:
        critique = generate_llm_response(
            f"Critique this result for objective='{objective}'. "
            f"Be concise. If fix needed, say 'RETRY' then one fix tip:\n{json.dumps(observation)[:2500]}"
        )
        print(f"[{agent_name}] Critique:", critique)
        log_event(session_id, agent_name, "critique", {"text": critique})

        if "RETRY" in (critique or "").upper() and max_retries > 0:
            fix = generate_llm_response("Propose one precise fix step only.")
            print(f"[{agent_name}] Reflexion fix tip:", fix)
            log_event(session_id, agent_name, "thought_summary", {"reflexion_fix": fix})

            observation = act({**plan, "fix": fix})
            print(f"[{agent_name}] Retried. Observation keys:", list(observation.keys()))
            log_event(session_id, agent_name, "observation", {"result_keys": list(observation.keys()), "retry": True})

    print(f"==== [{agent_name}] Done ====\n")
    return observation


# =========================================
# Agents
# =========================================
def transcript_analyzer_agent(session_id: str, pdf_path: str, reasoning_mode: str = "react") -> TranscriptState:

    """Parses PDF → DataFrames → summary_df (for GPA metrics)."""
    def think():
        return {"steps": ["load_pdf", "parse_courses", "split_blocks", "calc_gpa", "classify", "guess_major"]}

    def act(_plan):
        # 1) Raw extract
        raw = extract_text(pdf_path)
        text = clean_text(raw)

        # 2) Basic info
        info = parse_student_info_v3(text)
        # Clean potential noisy name suffixes (mirrors app.py)
        if "Name" in info:
            info["Name"] = re.split(
                r"University|Admission|Date|Of Prince|Muqrin",
                info["Name"], flags=re.IGNORECASE
            )[0].strip(" ,;:-")

        # 3) Courses
        df_all = parse_courses_with_multiline_fix(text)

        # 4) Split blocks for GPA extraction
        prep_text, ug_text = split_academic_blocks(text)
        df_gpa_prep = extract_gpa_summary_v3(prep_text)
        df_gpa_ug = extract_gpa_summary_v3(ug_text)

        # 5) Categories
        df_prep, df_ug, df_inprog, df_waived = split_by_category_v4(df_all)

        # 6) Summary dataframe (this is what compare_transcript_with_plan expects via extracted_summary_df)
        df_summary = create_summary(
            info, df_all, df_prep, df_ug, df_inprog, df_waived, df_gpa_prep, df_gpa_ug, full_text=text
        )

        # 7) Major guess by overlap
        student_codes = {str(r["Course Code"]).replace(" ", "").upper() for r in df_all.to_dict("records")}
        best_major, best_hit = None, -1
        for major, plan in STUDY_PLANS.items():
            plan_codes = {
                c["course_code"].replace(" ", "").upper()
                for lvl in plan.get("levels", {}).values()
                for c in lvl
            }
            hit = len(student_codes & plan_codes)
            if hit > best_hit:
                best_major, best_hit = major, hit

        print("TranscriptAnalyzer → Name:", info.get("Name", "Unknown"))
        print("TranscriptAnalyzer → Major guess:", best_major)
        print("TranscriptAnalyzer → Courses parsed:", len(df_all))

        return {
            "info": info,
            "courses_df": df_all.to_dict("records"),
            "gpa_table": df_gpa_ug.to_dict("records") if hasattr(df_gpa_ug, "to_dict") else [],
            "derived": {
                "summary_df": df_summary.to_dict("records")  # ✅ this feeds GPA metrics to compare
            },
            "major_guess": best_major,
        }

    allow_reflexion = reasoning_mode.endswith("+reflexion")
    obs = _react("TranscriptAnalyzer", "Parse and structure transcript", think, act, session_id,
                 allow_reflexion=allow_reflexion)
    st = TranscriptState(**obs)
    set_state(session_id, "student_state", asdict(st))
    return st


def study_plan_advisor_agent(session_id: str, student: TranscriptState, chosen_major: Optional[str] = None,
                             reasoning_mode: str = "react", export_excel: bool = False, hitl: bool = False) -> PlanState:
    """Compares transcript vs plan, computes totals & GPA block, optionally exports Excel."""
    def think():
        return {"steps": ["pick_major", "build_df", "inject_gpa_table", "compare", "compute_totals", "maybe_export_excel"]}

    def act(_plan):
        # Pick major
        major = chosen_major or (student.major_guess or "BS in Artificial Intelligence and Data Science")

        # Build DataFrame back from student.courses_df
        df = pd.DataFrame(student.courses_df)

        # Build extracted_summary_df back (this is crucial for GPA metrics)
        extracted_summary_df = pd.DataFrame(student.derived.get("summary_df", []))

        # 🔍 DEV DEBUG: print GPA table head (as you requested)
        try:
            print("DEBUG GPA TABLE →", extracted_summary_df.head())
        except Exception:
            print("DEBUG GPA TABLE → <unavailable>")

        # Compare transcript vs. plan (imported from shared_tools)
        summary = compare_transcript_with_plan(
            major,
            df,
            {"programs": STUDY_PLANS},
            extracted_summary_df=extracted_summary_df  # ✅ feeds GPA to Excel later
        )

        # Totals for UI/Excel
        completed = summary.get("completed_courses", []) or []
        in_progress = summary.get("in_progress_courses", []) or []
        remaining = summary.get("remaining_courses", []) or []

        summary["totals"] = {
            "required": len(completed) + len(in_progress) + len(remaining),
            "completed": len(completed),
            "in_progress": len(in_progress),
            "remaining": len(remaining),
        }

        # GPA nested structure for Excel
        summary["gpa"] = {
            "prep": summary.get("gpa_prep", "—"),
            "undergrad": summary.get("gpa_final", "—"),
        }

        print("Advisor → GPA block:", summary["gpa"])
        print("Advisor → Totals:", summary["totals"])

        # Optional Excel export
        excel_path = None
        if export_excel:
            # Import these from app.py to ensure single-source of truth
            from app import _pick_program, _levels_to_year_sem, UPLOAD_FOLDER

            program_key, program_obj = _pick_program(STUDY_PLANS, summary.get("major") or major, df)
            if not program_obj:
                raise ValueError("No matching study plan found for Excel export.")

            structured = _levels_to_year_sem(program_obj.get("levels", {}))

            excel_path = generate_structured_study_plan_excel(
                student.info,
                summary,
                structured,
                output_path=os.path.join(UPLOAD_FOLDER, f"Student_Summary_{session_id}.xlsx"),
            )
            print("Advisor → Excel saved:", excel_path)

        return {"major": major, "summary": summary, "excel_path": excel_path}

    allow_reflexion = reasoning_mode.endswith("+reflexion")
    obs = _react("StudyPlanAdvisor", "Compare transcript vs. plan and advise", think, act, session_id,
                 allow_reflexion=allow_reflexion)
    ps = PlanState(**obs)

    # Save plan_state for reuse
    save_dict = asdict(ps)
    save_dict.pop("pending", None)
    set_state(session_id, "plan_state", save_dict)

    # Track created artifact
    if ps.excel_path:
        conn = _db()
        conn.execute(
            "INSERT INTO artifacts VALUES(?,?,?,?,?)",
            (str(uuid.uuid4()), session_id, "excel", ps.excel_path, json.dumps({"major": ps.major}, ensure_ascii=False)),
        )
        conn.commit()
        conn.close()

    return ps


# =========================================
# Orchestrator (UI-facing upload handler)
# =========================================
def ui_agent_handle_upload(pdf_path: str, user_id: str, reasoning_mode: str = "react",
                           hitl: bool = False, orchestrator_mode: str = "sequential"):
    """Entry point for /agents/upload (called from app.py)."""
    session_id = str(uuid.uuid4())

    conn = _db()
    conn.execute(
        "INSERT INTO sessions VALUES(?,?,?,?,?,?)",
        (session_id, time.time(), user_id, reasoning_mode, 1 if hitl else 0, orchestrator_mode),
    )
    conn.commit()
    conn.close()

    log_event(session_id, "UI", "user_msg", {"upload_pdf": os.path.basename(pdf_path)})

    # 1) Analyze transcript
    student = transcript_analyzer_agent(session_id, pdf_path, reasoning_mode=reasoning_mode)

    # 2) Advise + Excel export
    plan = study_plan_advisor_agent(session_id, student, reasoning_mode=reasoning_mode,
                                    export_excel=True, hitl=hitl)

    # 3) Build short UI note
    summary_note = generate_llm_response(
        f"Summarize in 5 concise bullet points the student's status using this JSON:\n{json.dumps(plan.summary)[:2500]}"
    )
    log_event(session_id, "UI", "thought_summary", {"ui_summary": summary_note})

    # 4) Normalize summary (for chat + memory)
    summary = plan.summary

    completed = summary.get("completed_courses", []) or []
    in_progress = summary.get("in_progress_courses", []) or []
    remaining = summary.get("remaining_courses", []) or []

    summary["totals"] = {
        "required": len(completed) + len(in_progress) + len(remaining),
        "completed": len(completed),
        "in_progress": len(in_progress),
        "remaining": len(remaining),
    }

    if "gpa" not in summary:
        summary["gpa"] = {}
    summary["gpa"]["prep"] = summary.get("gpa_prep", summary["gpa"].get("prep", "—"))
    summary["gpa"]["undergrad"] = summary.get("gpa_final", summary["gpa"].get("undergrad", "—"))

    # 5️⃣ Save normalized memory for /chat use
    def save(k, v):
        _save_state(session_id, k, v)

    # ✅ Normalized GPA and credits
    normalized_gpa = {
        "prep": summary.get("gpa_prep"),
        "undergrad": summary.get("gpa_undergrad") or summary.get("gpa_final")
    }
    normalized_credits = {
        "prep": summary.get("credit_prep"),
        "undergrad": summary.get("credit_ug"),
        "in_progress": summary.get("credit_inprogress")
    }

    save("major", summary.get("major"))
    save("gpa_prep", normalized_gpa["prep"])
    save("gpa_undergrad", normalized_gpa["undergrad"])
    save("gpa", normalized_gpa)  # store nested dictionary for structured retrieval
    save("totals", summary.get("totals"))

    save("credits_prep", normalized_credits["prep"])
    save("credits_undergrad", normalized_credits["undergrad"])
    save("credits_in_progress", normalized_credits["in_progress"])

    save("progress_percent", summary.get("progress_percent"))
    save("completed_courses", completed)
    save("in_progress_courses", in_progress)
    save("remaining_courses", remaining)

    save("total_courses_count", len(completed)+len(in_progress)+len(remaining))



    print("UI → session_id:", session_id)
    print("UI → major:", summary.get("major"))
    print("UI → GPA (prep, undergrad):", summary["gpa"]["prep"], summary["gpa"]["undergrad"])
    print("UI → Excel path:", plan.excel_path)

    return session_id, student, plan, summary_note

# =========================================
# Chat Advisor Agent (Memory-based Q&A)
# =========================================
def chat_advisor_agent(session_id: str, message: str) -> Dict[str, Any]:
    """Retrieve structured memory from DB and respond precisely."""
    mem = _get_state(session_id)
    if not mem:
        return {"response": "📎 Please upload your transcript (PDF only) first."}

    # normalize values
    major = mem.get("major", "Unknown major")
    completed = mem.get("completed_courses") or []
    in_progress = mem.get("in_progress_courses") or []
    remaining = mem.get("remaining_courses") or []
    progress = round(float(mem.get("progress_percent") or 0), 2)
    gpa_prep = mem.get("gpa_prep") or "N/A"
    gpa_final = mem.get("gpa_undergrad") or mem.get("gpa_final") or "N/A"

    msg = (message or "").lower().strip()
    if "gpa" in msg:
        reply = f"🎓 Preparatory GPA: {gpa_prep}\n📘 Undergraduate GPA: {gpa_final}"
    elif "how many" in msg or "finished" in msg or "completed" in msg:
        total = len(completed) + len(in_progress) + len(remaining)
        reply = f"✅ You have completed **{len(completed)}** courses out of **{total}** total."
    elif "remaining" in msg or "left" in msg:
        reply = f"📚 Remaining courses ({len(remaining)}): {', '.join(remaining)}"
    elif "progress" in msg or "%" in msg:
        reply = f"📈 Your overall program progress is {progress}%."
    elif "register" in msg or "next" in msg:
        suggestion = remaining[:2] if remaining else []
        reply = f"🗓️ Recommended next courses: {', '.join(suggestion) if suggestion else 'All courses completed 🎉'}"
    else:
        reply = (
            f"Major: {major}\n"
            f"Progress: {progress}%\n"
            f"GPA (Prep/UG): {gpa_prep}/{gpa_final}\n"
            f"Completed: {len(completed)}, In progress: {len(in_progress)}, Remaining: {len(remaining)}"
        )

    log_event(session_id, "ChatAdvisor", "reply", {"text": reply})
    print(f"[CHAT DEBUG] completed={len(completed)}, in_progress={len(in_progress)}, remaining={len(remaining)}")

    return {"response": reply}

# =========================================
# HITL helper (no-op unless you wire it)
# =========================================
def request_hitl_action(session_id: str, agent: str, action: str, payload: Dict[str, Any]) -> str:
    conn = _db()
    item_id = str(uuid.uuid4())
    ts = time.time()
    conn.execute(
        "INSERT INTO hitl_queue VALUES(?,?,?,?,?,?,?)",
        (item_id, session_id, agent, action, json.dumps(payload, ensure_ascii=False), "pending", ts),
    )
    conn.commit()
    conn.close()
    return item_id
