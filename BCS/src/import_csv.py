"""Manual-mode step — import Athena console CSV downloads into the story.

Workflow: paste a sql/ file into the Athena query editor, run it, click
"Download results (.csv)", and drop the file into output/csv/ AS-IS (no
renaming needed). Then:

    python run_all.py --import --report

Matching: every manifest entry declares its exact output header (columns).
A CSV whose header matches one query's columns is assigned to it. A file
named <id>.csv wins outright. Unrecognized headers are listed, never guessed.
Imports MERGE into output/story_data.json: re-importing one CSV never loses
the others, and a newer download of the same query replaces the older one.

Stdlib only - works even where boto3 is not installed.
"""

import csv
import json
import os
import sys
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import settings


def load_manifest():
    with open(settings.MANIFEST, encoding="utf-8") as f:
        return json.load(f)["queries"]


def read_sql(filename):
    try:
        with open(os.path.join(settings.SQL_DIR, filename), encoding="utf-8") as f:
            return f.read()
    except OSError:
        return ""


def _signature(cols):
    return tuple(c.strip().lower() for c in cols)


def _convert(value):
    if value is None:
        return None
    s = value.strip()
    if s == "":
        return None
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        pass
    return s


def read_csv_file(path):
    """Return (columns, rows-as-dicts) with numeric strings converted."""
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        try:
            header = next(reader)
        except StopIteration:
            return [], []
        columns = [c.strip() for c in header]
        rows = []
        for raw in reader:
            raw += [None] * (len(columns) - len(raw))
            rows.append({col: _convert(raw[j]) for j, col in enumerate(columns)})
            if len(rows) >= settings.MAX_RESULT_ROWS:
                break
    return columns, rows


def load_previous():
    if not os.path.exists(settings.STORY_JSON):
        return {}
    try:
        with open(settings.STORY_JSON, encoding="utf-8") as f:
            doc = json.load(f)
        return {s["id"]: s for s in doc.get("stats", []) if "id" in s}
    except (OSError, ValueError, KeyError):
        return {}


def main(argv=None):  # argv accepted for symmetry with the other steps
    settings.ensure_output_dir()
    csv_dir = settings.CSV_DIR
    os.makedirs(csv_dir, exist_ok=True)

    queries = load_manifest()
    by_id = {q["id"]: q for q in queries}
    by_sig = {}
    for q in queries:
        sig = _signature(q.get("columns", []))
        if sig:
            by_sig[sig] = q

    files = sorted(
        (f for f in os.listdir(csv_dir) if f.lower().endswith(".csv")),
        key=lambda f: os.path.getmtime(os.path.join(csv_dir, f)),
    )
    if not files:
        print(f"No CSV files found in {csv_dir}")
        print("Paste a sql/ file into the Athena editor, run it, download the "
              "results CSV, and drop it in that folder.")
        return 1

    previous = load_previous()
    fresh = {}   # id -> entry (later files with same id overwrite earlier ones)
    matched, unmatched = [], []

    for fname in files:
        path = os.path.join(csv_dir, fname)
        try:
            columns, rows = read_csv_file(path)
        except (OSError, csv.Error) as exc:
            unmatched.append((fname, f"unreadable: {exc}"))
            continue

        stem = os.path.splitext(fname)[0].lower()
        spec = by_id.get(stem) or by_sig.get(_signature(columns))
        if spec is None:
            unmatched.append((fname, f"header not recognized: {columns[:6]}"))
            continue

        entry = dict(spec)
        entry.update(
            status="ok",
            source_csv=fname,
            imported_at=datetime.now(timezone.utc).isoformat(timespec="seconds"),
            sql=read_sql(spec["file"]),
            columns=columns,
            rows=rows,
            row_count=len(rows),
        )
        if spec["id"] in fresh:
            matched.append((fname, f"{spec['id']} (replaces earlier file this pass)"))
        else:
            matched.append((fname, spec["id"]))
        fresh[spec["id"]] = entry

    # Merge in manifest order: fresh > previous.
    stats = []
    for q in queries:
        if q["id"] in fresh:
            stats.append(fresh[q["id"]])
        elif q["id"] in previous:
            stats.append(previous[q["id"]])
    story = {
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "mode": "manual-csv",
        "stats": stats,
    }
    with open(settings.STORY_JSON, "w", encoding="utf-8") as f:
        json.dump(story, f, indent=1, default=str)

    done_ids = {s["id"] for s in stats if s.get("status") == "ok"}
    pending = [q for q in queries if q["id"] not in done_ids]

    print(f"Imported {len(fresh)} quer{'y' if len(fresh) == 1 else 'ies'} "
          f"from {len(files)} file(s):")
    for fname, target in matched:
        print(f"  {fname}  ->  {target}")
    for fname, why in unmatched:
        print(f"  [unmatched] {fname}: {why}")
    print(f"\nStory now holds {len(done_ids)} of {len(queries)} queries "
          f"->  {settings.STORY_JSON}")
    if pending:
        print(f"Still pending ({len(pending)}):")
        for q in pending:
            print(f"  {q['id']:<28} paste sql\\{q['file']}")
    print("\nBuild the report with:  python run_all.py --report")
    return 0


if __name__ == "__main__":
    sys.exit(main())
