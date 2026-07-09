"""Act-ordered story report — render the January 2025 story acts with their
query cards interleaved: output/uc2_story.html.

Companion to build_report.py (the 88-query atlas), not a replacement. The act
map lives in sql/act_map.json: ordered acts, each with narrative paragraphs
and the query ids whose cards belong under it. Cards render from the same
output/story_data.json the atlas uses, so imports feed both reports and the
story report builds up as CSVs land (missing queries render as pending cards
naming the file to paste).

Reuses build_report's render machinery (render_stat, render_pending, CSS,
load_explains) — one card looks identical in both reports.
"""

import html
import json
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import settings
from src.build_report import CSS, load_explains, render_pending, render_stat

ACT_MAP = os.path.join(settings.SQL_DIR, "act_map.json")
STORY_REPORT_HTML = os.path.join(settings.OUTPUT_DIR, "uc2_story.html")


def esc(x):
    return html.escape("" if x is None else str(x))


def main():
    if not os.path.exists(ACT_MAP):
        print(f"No act map found at {ACT_MAP}")
        return 1
    with open(ACT_MAP, encoding="utf-8") as f:
        act_map = json.load(f)

    story = {"stats": []}
    if os.path.exists(settings.STORY_JSON):
        with open(settings.STORY_JSON, encoding="utf-8") as f:
            story = json.load(f)
    try:
        with open(settings.MANIFEST, encoding="utf-8") as f:
            manifest = {q["id"]: q for q in json.load(f)["queries"]}
    except (OSError, ValueError, KeyError):
        manifest = {}

    stats_by_id = {s.get("id"): s for s in story.get("stats", [])}
    explains = load_explains()

    wanted = [qid for act in act_map.get("acts", []) for qid in act.get("queries", [])]
    have = sum(1 for qid in wanted if stats_by_id.get(qid, {}).get("status") == "ok")

    body = ['<div class="viz-root"><header>']
    body.append(f"<h1>{esc(act_map.get('title', 'Story report'))}</h1>")
    body.append(
        f"<p>{esc(act_map.get('subtitle', ''))} {have} of {len(wanted)} story cards "
        f"have results. Last import {esc(story.get('generated_at', 'never'))} "
        f"· report built {date.today().isoformat()}.</p></header>"
    )
    nav = ["<nav>"]
    for act in act_map.get("acts", []):
        nav.append(f'<a href="#{esc(act["id"])}">{esc(act["title"])}</a>')
    nav.append("</nav><main>")
    body.append("".join(nav))

    for act in act_map.get("acts", []):
        body.append(f'<h2 id="{esc(act["id"])}">{esc(act["title"])}</h2>')
        for para in act.get("narrative", []):
            body.append(f'<p class="tier-intro">{esc(para)}</p>')
        for qid in act.get("queries", []):
            s = stats_by_id.get(qid)
            q = manifest.get(qid)
            exp = explains.get(qid, "")
            if s and s.get("status") == "ok":
                body.append(render_stat(s, exp))
            elif s:
                body.append(render_stat(s, exp))  # failed: card shows the error
            elif q:
                body.append(render_pending(q, exp))
            else:
                body.append(
                    f'<section class="card pending"><h3>{esc(qid)}</h3>'
                    f'<p class="failed">Query id not found in sql/manifest.json — '
                    f"fix the act map or add the query.</p></section>"
                )
        body.append("")

    body.append(
        "</main><footer>Act-ordered view of the same imported results as "
        "bcs_story.html (the atlas). Generated locally from Athena aggregates — "
        "this file contains data extracts; keep it inside the environment; never "
        "commit the output folder.</footer></div>"
    )

    html_doc = (
        "<!doctype html><html lang='en'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width, initial-scale=1'>"
        f"<title>{esc(act_map.get('title', 'Story report'))}</title>"
        f"<style>{CSS}</style></head><body>" + "".join(body) + "</body></html>"
    )
    settings.ensure_output_dir()
    with open(STORY_REPORT_HTML, "w", encoding="utf-8") as f:
        f.write(html_doc)
    print(f"Story report written  ->  {STORY_REPORT_HTML}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
