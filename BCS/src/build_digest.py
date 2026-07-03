"""Step — render output/story_data.json into ONE compact markdown file:
output/digest.md.

Purpose: results must leave the environment without screenshotting a long
HTML page. The digest holds EVERY imported number as tight markdown tables,
core queries first, so a whole run exits in two or three screenshots (or one
copy-paste). No SQL, no explainers, no charts — just the numbers.
"""

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import settings
from src.build_report import TIER_INTRO, TIER_ORDER, fmt, pivot

MAX_DIGEST_ROWS = 500  # far above any query's real output; safety net only


def cell(v, col):
    # fmt() gives thousands separators, % on pct columns, dates for yyyymmdd ints
    return str(fmt(v, col)).replace("|", "\\|")


def table(columns, rows):
    lines = ["| " + " | ".join(str(c) for c in columns) + " |",
             "|" + "|".join(" ---: " if i else " --- " for i in range(len(columns))) + "|"]
    for r in rows[:MAX_DIGEST_ROWS]:
        lines.append("| " + " | ".join(cell(r.get(c), c) for c in columns) + " |")
    if len(rows) > MAX_DIGEST_ROWS:
        lines.append(f"| ... | (+{len(rows) - MAX_DIGEST_ROWS} more rows in the CSV) |")
    return lines


def main():
    if not os.path.exists(settings.STORY_JSON):
        print(f"No {settings.STORY_JSON} found — run --fetch or --import first.")
        return 1
    with open(settings.STORY_JSON, encoding="utf-8") as f:
        story = json.load(f)
    try:
        with open(settings.MANIFEST, encoding="utf-8") as f:
            manifest = json.load(f)["queries"]
    except (OSError, ValueError, KeyError):
        manifest = []

    stats_by_id = {s.get("id"): s for s in story.get("stats", [])}
    ok_n = sum(1 for s in stats_by_id.values() if s.get("status") == "ok")
    total = len(manifest) or len(stats_by_id)

    out = [f"# BCS story digest — {story.get('generated_at', '?')} — "
           f"{ok_n} of {total} queries ({story.get('mode', '?')})", ""]
    pending, failed = [], []

    for tier in TIER_ORDER:
        title = TIER_INTRO.get(tier, (f"Tier {tier}", ""))[0]
        section = []
        # core first, then context — same lead order as the report
        tier_qs = [q for q in manifest if q.get("tier") == tier]
        for q in sorted(tier_qs, key=lambda q: q.get("story", "core") != "core"):
            s = stats_by_id.get(q["id"])
            if s is None:
                pending.append(q["id"])
                continue
            if s.get("status") != "ok":
                failed.append(f"{q['id']}: {s.get('error', 'unknown error')}")
                continue
            tag = " *(context)*" if q.get("story", "core") == "context" else ""
            section.append(f"### {q['id']} — {q.get('title', '')}{tag}")
            rows, cols = s.get("rows", []), s.get("columns", [])
            if not rows:
                section.append("(ran, no rows)")
            elif q.get("render") == "line" and q.get("series_col"):
                # same wide pivot as the HTML: complete AND screenshot-compact
                wide, series = pivot(rows, q["x_col"], q["series_col"], q["value_col"])
                section.extend(table([q["x_col"]] + series, wide))
            else:
                section.extend(table(cols, rows))
            section.append("")
        if section:
            out.append(f"## Tier {tier} — {title}")
            out.append("")
            out.extend(section)

    if failed:
        out.append("## Failed")
        out.extend(f"- {f}" for f in failed)
        out.append("")
    if pending:
        out.append(f"## Pending ({len(pending)})")
        out.append(", ".join(pending))
        out.append("")

    settings.ensure_output_dir()
    with open(settings.DIGEST_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(out))
    print(f"Digest written  ->  {settings.DIGEST_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
