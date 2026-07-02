"""Orchestrator: connection check -> story fetch -> HTML report.

Usage:
  python run_all.py            all three steps
  python run_all.py --check    connection check only
  python run_all.py --fetch    story fetch only
  python run_all.py --report   rebuild the HTML from existing output JSONs
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import settings
from src import build_report, check_connection, fetch_story


def main():
    ap = argparse.ArgumentParser(description="BCS Athena data story")
    ap.add_argument("--check", action="store_true", help="connection check only")
    ap.add_argument("--fetch", action="store_true", help="story fetch only")
    ap.add_argument("--report", action="store_true", help="rebuild HTML only")
    args = ap.parse_args()

    run_all = not (args.check or args.fetch or args.report)

    if run_all or args.check:
        rc = check_connection.main()
        if rc == 1:  # credentials / Athena unreachable: no point continuing
            return rc
        if rc == 2:
            print("\nSome tables had issues; continuing (failures will show in the report).\n")
        if args.check and not run_all:
            return rc

    if run_all or args.fetch:
        fetch_story.main()  # per-query failures are recorded, not fatal

    if run_all or args.report:
        rc = build_report.main()
        if rc == 0:
            print(f"\nOpen the story:  {settings.REPORT_HTML}")
        return rc

    return 0


if __name__ == "__main__":
    sys.exit(main())
