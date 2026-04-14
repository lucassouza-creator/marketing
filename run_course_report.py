#!/usr/bin/env python3
"""
LearnWorlds Course Report Generator
------------------------------------
Usage:
    python run_course_report.py                  # list all courses
    python run_course_report.py <course_id>      # generate report for a course
"""

import sys
from learnworlds.client import LearnWorldsClient
from learnworlds.extractor import list_courses, collect_course_data
from learnworlds.reporter import generate_report


def cmd_list_courses(client: LearnWorldsClient):
    print("\nAvailable courses:\n")
    courses = list_courses(client)
    if not courses:
        print("  (no courses found)")
        return
    for c in courses:
        cid = c.get("id", "?")
        title = c.get("title", "Untitled")
        status = c.get("status", "")
        print(f"  [{cid}]  {title}  ({status})")
    print(
        f"\nRun with a course ID to generate its report:\n"
        f"  python run_course_report.py <course_id>\n"
    )


def cmd_course_report(client: LearnWorldsClient, course_id: str):
    course_data = collect_course_data(client, course_id)
    generate_report(course_data)


if __name__ == "__main__":
    client = LearnWorldsClient()

    if len(sys.argv) < 2:
        cmd_list_courses(client)
    else:
        cmd_course_report(client, sys.argv[1])
