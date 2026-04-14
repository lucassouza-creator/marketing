"""Extract comprehensive course data from LearnWorlds API."""

from datetime import datetime, timezone
from .client import LearnWorldsClient


def get_course(client: LearnWorldsClient, course_id: str) -> dict:
    data = client.get(f"/v2/courses/{course_id}")
    return data.get("data", data)


def get_enrolled_users(client: LearnWorldsClient, course_id: str) -> list[dict]:
    users = list(client.get_paginated(f"/v2/courses/{course_id}/users"))
    print(f"[extractor] {len(users)} enrolled users")
    return users


def get_course_reviews(client: LearnWorldsClient, course_id: str) -> list[dict]:
    """Try multiple endpoints for course reviews/satisfaction."""
    for path in (f"/v2/courses/{course_id}/reviews", f"/v2/courses/{course_id}/ratings"):
        try:
            items = list(client.get_paginated(path))
            print(f"[extractor] {len(items)} reviews via {path}")
            return items
        except Exception:
            pass
    print("[extractor] Reviews endpoint not available")
    return []


def _parse_unit_type(unit: dict) -> str:
    t = unit.get("type", "").lower()
    if t in ("exam", "quiz", "assessment"):
        return "quiz"
    if t in ("video", "vimeo", "youtube"):
        return "video"
    if t in ("pdf", "document", "iframe"):
        return "document"
    return t or "other"


def _parse_sections_from_progress(progress: dict) -> tuple[list[dict], dict]:
    """
    Extract per-section data and unit-level stats from a user progress response.
    Returns (section_records, unit_stats) where unit_stats has:
      quiz_units_total, quiz_units_completed, quiz_avg_score,
      video_units_total, video_units_completed
    """
    section_records = []
    quiz_scores = []
    quiz_total = quiz_done = video_total = video_done = 0

    for section in progress.get("sections", []):
        units = section.get("units", [])
        completed_units = sum(1 for u in units if u.get("completed", False))
        section_quiz_scores = []

        for u in units:
            utype = _parse_unit_type(u)
            completed = u.get("completed", False)
            score = u.get("score")

            if utype == "quiz":
                quiz_total += 1
                if completed:
                    quiz_done += 1
                if score is not None:
                    quiz_scores.append(score)
                    section_quiz_scores.append(score)
            elif utype == "video":
                video_total += 1
                if completed:
                    video_done += 1

        section_records.append({
            "section_id": section.get("id") or section.get("sectionId", ""),
            "section_title": section.get("title", "Seção sem título"),
            "total_units": len(units),
            "completed_units": completed_units,
            "section_completion_pct": (
                round(completed_units / len(units) * 100, 1) if units else 0.0
            ),
            "section_quiz_avg": (
                round(sum(section_quiz_scores) / len(section_quiz_scores), 1)
                if section_quiz_scores else None
            ),
        })

    unit_stats = {
        "quiz_units_total": quiz_total,
        "quiz_units_completed": quiz_done,
        "quiz_avg_score": round(sum(quiz_scores) / len(quiz_scores), 1) if quiz_scores else None,
        "video_units_total": video_total,
        "video_units_completed": video_done,
    }
    return section_records, unit_stats


def _days_between(date_str: str) -> int | None:
    """Return days from date_str to now, or None if unparseable."""
    if not date_str:
        return None
    try:
        dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        return (datetime.now(timezone.utc) - dt).days
    except Exception:
        return None


def collect_full_course_data(client: LearnWorldsClient, course_id: str) -> dict:
    """
    Comprehensive data collection:
    - Course metadata
    - All enrolled users with detailed progress
    - Per-user per-section breakdown
    - Course reviews/satisfaction
    """
    print(f"\n[extractor] Collecting full data for course: {course_id}")

    course = get_course(client, course_id)
    users = get_enrolled_users(client, course_id)
    reviews = get_course_reviews(client, course_id)

    progress_records = []
    section_records = []

    for i, user in enumerate(users, 1):
        uid = user.get("id") or user.get("userId")
        if not uid:
            continue

        email = user.get("email", "")
        print(f"[extractor] [{i}/{len(users)}] {email or uid}")

        # Fetch detailed progress
        try:
            resp = client.get(f"/v2/users/{uid}/course-progress/{course_id}")
            progress = resp.get("data", resp)
        except Exception as e:
            print(f"  ↳ progress unavailable: {e}")
            progress = {}

        enrolled_at = user.get("enrolledAt") or user.get("created_at", "")
        last_activity = (
            progress.get("lastActivityAt")
            or progress.get("updatedAt")
            or user.get("lastActivityAt", "")
        )

        completion_pct = float(progress.get("completionPercentage", 0) or 0)
        completed_units = int(progress.get("completedUnits", 0) or 0)
        total_units = int(progress.get("totalUnits", 0) or 0)
        time_s = int(progress.get("totalTimeSpentInSeconds", 0) or 0)
        status = progress.get("status", "not_started") or "not_started"

        # Per-section and unit-level details
        user_sections, unit_stats = _parse_sections_from_progress(progress)

        # Fall back to aggregate score if per-unit quiz data not available
        quiz_avg = unit_stats["quiz_avg_score"] or progress.get("score")

        record = {
            "user_id": uid,
            "email": email,
            "username": user.get("username") or user.get("name") or "",
            "enrolled_at": enrolled_at,
            "last_activity": last_activity,
            "days_since_enrollment": _days_between(enrolled_at),
            "days_since_last_access": _days_between(last_activity),
            "completion_percentage": completion_pct,
            "completed_lessons": completed_units,
            "total_lessons": total_units,
            "time_spent_seconds": time_s,
            "time_spent_minutes": round(time_s / 60, 1),
            "time_spent_hours": round(time_s / 3600, 2),
            "status": status,
            "quiz_score": quiz_avg,
            "quiz_units_total": unit_stats["quiz_units_total"],
            "quiz_units_completed": unit_stats["quiz_units_completed"],
            "video_units_completed": unit_stats["video_units_completed"],
            "certificate_issued": bool(progress.get("certificateIssued", False)),
        }
        progress_records.append(record)

        for sec in user_sections:
            section_records.append({"user_id": uid, "email": email, **sec})

    review_records = [
        {
            "user_id": r.get("userId", ""),
            "rating": r.get("rating") or r.get("score"),
            "comment": r.get("comment") or r.get("body", ""),
            "created_at": r.get("createdAt") or r.get("created_at", ""),
        }
        for r in reviews
    ]

    return {
        "course": course,
        "users": users,
        "progress_records": progress_records,
        "section_records": section_records,
        "reviews": review_records,
    }


# Backward-compatible helpers used by run_course_report.py
def list_courses(client: LearnWorldsClient) -> list[dict]:
    courses = list(client.get_paginated("/v2/courses"))
    print(f"[extractor] Found {len(courses)} courses")
    return courses


def collect_course_data(client: LearnWorldsClient, course_id: str) -> dict:
    """Alias for backward compatibility."""
    return collect_full_course_data(client, course_id)
