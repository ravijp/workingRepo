"""Step 1 — run every query registered in sql/manifest.json and save the
results to output/story_data.json.

The manifest binds each .sql file to its tier, report title, question, and
chart type. Every query is aggregate-only, anchored to the newest data in
each table (robust to stale copies), and wrapped so one failure never stops
the run. Entries with fallback_file retry automatically on failure.
"""

import json
import os
import sys
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import settings
from src import athena


def load_manifest():
    with open(settings.MANIFEST, encoding="utf-8") as f:
        manifest = json.load(f)
    return manifest["queries"]


def read_sql(filename):
    with open(os.path.join(settings.SQL_DIR, filename), encoding="utf-8") as f:
        return f.read()


def main():
    settings.ensure_output_dir()
    queries = load_manifest()
    story = {
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "region": settings.AWS_REGION,
        "workgroup": settings.ATHENA_WORKGROUP,
        "stats": [],
    }
    n = len(queries)
    failures = 0

    for i, spec in enumerate(queries, 1):
        print(f"[{i:>2}/{n}] {spec['id']}: {spec['title']} ...", flush=True)
        entry = dict(spec)
        try:
            sql = read_sql(spec["file"])
        except OSError as exc:
            entry.update(status="failed", error=f"cannot read {spec['file']}: {exc}")
            story["stats"].append(entry)
            failures += 1
            print(f"        FAILED  cannot read sql file: {exc}")
            continue

        entry["sql"] = sql
        try:
            res = athena.run(sql)
            entry.update(status="ok", used_fallback=False, **res)
            print(f"        ok  {res['row_count']} row(s), {res['elapsed_s']}s, {res['scanned_mb']} MB scanned")
        except athena.AthenaError as exc:
            fb = spec.get("fallback_file")
            if fb:
                print(f"        primary failed ({str(exc)[:90]}); trying fallback {fb} ...")
                try:
                    fb_sql = read_sql(fb)
                    res = athena.run(fb_sql)
                    entry.update(status="ok", used_fallback=True, sql=fb_sql, **res)
                    print(f"        ok (fallback)  {res['row_count']} row(s)")
                except (OSError, athena.AthenaError) as exc2:
                    entry.update(status="failed", error=str(exc2))
                    failures += 1
                    print(f"        FAILED  {exc2}")
            else:
                entry.update(status="failed", error=str(exc))
                failures += 1
                print(f"        FAILED  {exc}")
        except Exception as exc:  # noqa: BLE001
            hint = athena.friendly_credential_error(exc)
            entry.update(status="failed", error=hint or str(exc))
            failures += 1
            print(f"        FAILED  {hint or exc}")
            if hint:  # credentials died mid-run; stop early
                story["stats"].append(entry)
                break
        story["stats"].append(entry)

    with open(settings.STORY_JSON, "w", encoding="utf-8") as f:
        json.dump(story, f, indent=1, default=str)

    ok = len(story["stats"]) - failures
    print(f"\nStory fetch done: {ok}/{n} stats ok  ->  {settings.STORY_JSON}")
    return 0 if failures == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
