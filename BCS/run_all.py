"""Orchestrator for the BCS data story.

API mode (needs AWS network access from this machine):
  python run_all.py            connection check -> fetch -> report
  python run_all.py --check    connection check only
  python run_all.py --fetch    story fetch only
  python run_all.py --report   rebuild the HTML from existing output JSONs

Manual mode (run queries in the Athena console, download CSVs into
output/csv/ - see README "Manual mode"):
  python run_all.py --import                     import the CSVs
  python run_all.py --import --report --digest   import, build HTML + digest.md

Steps import lazily, so manual mode works even without boto3 installed.
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import settings


def main():
    ap = argparse.ArgumentParser(description="BCS Athena data story")
    ap.add_argument("--check", action="store_true", help="connection check only (API mode)")
    ap.add_argument("--fetch", action="store_true", help="story fetch only (API mode)")
    ap.add_argument("--import", dest="import_csvs", action="store_true",
                    help="import console CSV downloads from output/csv/ (manual mode)")
    ap.add_argument("--report", action="store_true", help="rebuild the HTML report")
    ap.add_argument("--digest", action="store_true",
                    help="write output/digest.md - every result as compact markdown")
    args = ap.parse_args()

    run_all = not (args.check or args.fetch or args.import_csvs or args.report
                   or args.digest)

    if run_all or args.check:
        from src import check_connection
        rc = check_connection.main()
        if rc == 1:  # credentials / Athena unreachable: no point continuing
            print("\nNo AWS path from this machine? Use manual mode: run the sql/ "
                  "files in the Athena console, download the CSVs into output\\csv\\, "
                  "then:  python run_all.py --import --report")
            return rc
        if rc == 2:
            print("\nSome tables had issues; continuing (failures will show in the report).\n")
        if args.check and not run_all:
            return rc

    if run_all or args.fetch:
        from src import fetch_story
        fetch_story.main()  # per-query failures are recorded, not fatal

    if args.import_csvs:
        from src import import_csv
        import_csv.main()

    rc = 0
    if run_all or args.report:
        from src import build_report
        rc = build_report.main()
        if rc == 0:
            print(f"\nOpen the story:  {settings.REPORT_HTML}")

    if run_all or args.digest:
        from src import build_digest
        rc = build_digest.main() or rc

    return rc


if __name__ == "__main__":
    sys.exit(main())
