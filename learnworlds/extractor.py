"""Extract course data from LearnWorlds API."""

from .client import LearnWorldsClient


def list_courses(client: LearnWorldsClient) -> list[dict]:
    """Return all courses available in the school."""
    courses = list(client.get_paginated("/v2/courses"))
    print(f"[extractor] Found {len(courses)} courses")
    return courses


def get_course(client: LearnWorldsClient, course_id: str) -> dict:
    """Return metadata for a specific course."""
    data = client.get(f"/v2/courses/{course_id}")
    return data.get("data", data)


def get_enrolled_users(client: LearnWorldsClient, course_id: str) -> list[dict]:
    """Return all users enrolled in a course."""
    users = list(client.get_paginated(f"/v2/courses/{course_id}/users"))
    print(f"[extractor] Found {len(users)} enrolled users")
    return users


def get_user_progress(
    client: LearnWorldsClient, user_id: str, course_id: str
) -> dict:
    """Return a single user's progress in a course."""
    try:
        data = client.get(f"/v2/users/{user_id}/course-progress/{course_id}")
        return data.get("data", data)
    except Exception as e:
        print(f"[extractor] Warning: could not fetch progress for user {user_id}: {e}")
        return {}


def collect_course_data(client: LearnWorldsClient, course_id: str) -> dict:
    """
    Collect all relevant data for a course:
    - Course metadata
    - Enrolled users
    - Per-user progress
    Returns a dict with keys: course, users, progress_records
    """
    print(f"[extractor] Collecting data for course {course_id} ...")

    course = get_course(client, course_id)
    users = get_enrolled_users(client, course_id)

    progress_records = []
    for i, user in enumerate(users, 1):
        uid = user.get("id") or user.get("userId")
        if not uid:
            continue
        print(f"[extractor] Fetching progress {i}/{len(users)} (user {uid})")
        progress = get_user_progress(client, uid, course_id)
        progress_records.append(
            {
                "user_id": uid,
                "email": user.get("email", ""),
                "username": user.get("username", ""),
                "enrolled_at": user.get("enrolledAt", ""),
                "completion_percentage": progress.get("completionPercentage", 0),
                "completed_units": progress.get("completedUnits", 0),
                "total_units": progress.get("totalUnits", 0),
                "time_spent_seconds": progress.get("totalTimeSpentInSeconds", 0),
                "last_activity": progress.get("lastActivityAt", ""),
                "status": progress.get("status", "unknown"),
                "score": progress.get("score", None),
                "certificate_issued": progress.get("certificateIssued", False),
            }
        )

    return {
        "course": course,
        "users": users,
        "progress_records": progress_records,
    }
