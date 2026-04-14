"""Generate course analysis reports from extracted LearnWorlds data."""

import os
import json
from datetime import datetime
import pandas as pd


def _output_dir() -> str:
    path = os.path.join(os.path.dirname(__file__), "..", "reports")
    os.makedirs(path, exist_ok=True)
    return os.path.abspath(path)


def build_dataframe(progress_records: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(progress_records)
    if df.empty:
        return df
    df["time_spent_minutes"] = (df["time_spent_seconds"] / 60).round(1)
    df["completion_percentage"] = pd.to_numeric(
        df["completion_percentage"], errors="coerce"
    ).fillna(0)
    return df


def compute_summary(course: dict, df: pd.DataFrame) -> dict:
    """Compute key metrics for the course report."""
    total = len(df)
    if total == 0:
        return {"error": "No enrolled users found"}

    completed = df[df["completion_percentage"] >= 100]
    in_progress = df[
        (df["completion_percentage"] > 0) & (df["completion_percentage"] < 100)
    ]
    not_started = df[df["completion_percentage"] == 0]

    return {
        "course_id": course.get("id", ""),
        "course_title": course.get("title", ""),
        "report_generated_at": datetime.now().isoformat(),
        "total_enrollments": total,
        "completed": len(completed),
        "in_progress": len(in_progress),
        "not_started": len(not_started),
        "completion_rate_pct": round(len(completed) / total * 100, 1),
        "avg_completion_pct": round(df["completion_percentage"].mean(), 1),
        "median_completion_pct": round(df["completion_percentage"].median(), 1),
        "avg_time_spent_minutes": round(df["time_spent_minutes"].mean(), 1),
        "total_certificates_issued": int(df["certificate_issued"].sum()),
    }


def save_csv(df: pd.DataFrame, course_id: str) -> str:
    path = os.path.join(_output_dir(), f"course_{course_id}_users.csv")
    df.to_csv(path, index=False)
    print(f"[reporter] CSV saved → {path}")
    return path


def save_summary_json(summary: dict, course_id: str) -> str:
    path = os.path.join(_output_dir(), f"course_{course_id}_summary.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"[reporter] Summary JSON saved → {path}")
    return path


def print_summary(summary: dict):
    print("\n" + "=" * 55)
    print(f"  COURSE REPORT: {summary.get('course_title', '')}")
    print("=" * 55)
    print(f"  ID:                    {summary.get('course_id')}")
    print(f"  Generated at:          {summary.get('report_generated_at')}")
    print("-" * 55)
    print(f"  Total enrollments:     {summary.get('total_enrollments')}")
    print(f"  Completed:             {summary.get('completed')}  ({summary.get('completion_rate_pct')}%)")
    print(f"  In progress:           {summary.get('in_progress')}")
    print(f"  Not started:           {summary.get('not_started')}")
    print("-" * 55)
    print(f"  Avg completion:        {summary.get('avg_completion_pct')}%")
    print(f"  Median completion:     {summary.get('median_completion_pct')}%")
    print(f"  Avg time spent:        {summary.get('avg_time_spent_minutes')} min")
    print(f"  Certificates issued:   {summary.get('total_certificates_issued')}")
    print("=" * 55 + "\n")


def generate_report(course_data: dict) -> dict:
    """
    Full pipeline: build DataFrame → compute summary → save files → print.
    Returns the summary dict.
    """
    course = course_data["course"]
    progress_records = course_data["progress_records"]
    course_id = course.get("id", "unknown")

    df = build_dataframe(progress_records)
    summary = compute_summary(course, df)

    if not df.empty:
        save_csv(df, course_id)
    save_summary_json(summary, course_id)
    print_summary(summary)

    return summary
