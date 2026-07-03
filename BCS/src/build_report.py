"""Step 2 — render output/story_data.json (+ connection check) into one
self-contained HTML file: output/bcs_story.html.

No CDN, no JS dependencies: charts are inline SVG generated here, styled
with CSS custom properties so the page follows the OS light/dark theme.
Every chart carries direct value labels and a data table underneath.
"""

import html
import json
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import settings

TIER_INTRO = {
    1: ("The shape of the data",
        "Row counts, time coverage, freshness and key fill rates. This tier says "
        "what the three tables actually contain before any interpretation."),
    2: ("The operational picture",
        "One table at a time: call volumes and mix, handle times, abandons and "
        "transfers, sentiment, conversation length, the delinquency ladder and "
        "the charge-off trend."),
    3: ("The cross-table story",
        "The tables joined up: whether calls resolve to accounts, how much of the "
        "call flow the transcripts cover, who calls repeatedly, how delinquent the "
        "callers are, and what customers say about paying."),
    4: ("Approach validation",
        "The checks that pin the sizing approach down: where calls sit on the DQ "
        "ladder, how a new-delinquency vintage actually rolls or cures, caller vs "
        "non-caller outcomes, dollars at risk vs dollars entering DQ1, payment "
        "after a call, the provision-stage shape, the inbound/outbound mix, and "
        "the re-age signal."),
}


def esc(x):
    return html.escape("" if x is None else str(x))


def fmt(v, col=""):
    if v is None:
        return "–"
    if isinstance(v, bool):
        return "yes" if v else "no"
    if isinstance(v, (int, float)):
        if col.startswith("pct") or col.endswith(("_pct", "_rate")):
            return f"{v:,.1f}%"
        if isinstance(v, int) or float(v).is_integer():
            return f"{int(v):,}"
        return f"{v:,.2f}"
    return str(v)


def nice_label(col):
    return str(col).replace("_", " ")


# ------------------------------------------------------------------ SVG bits
def svg_bars(rows, label_col, value_col, width=680):
    vals = [(r.get(label_col), r.get(value_col)) for r in rows
            if isinstance(r.get(value_col), (int, float))]
    if not vals:
        return ""
    mx = max(v for _, v in vals) or 1
    bar_h, gap, label_w = 22, 8, 210
    h = len(vals) * (bar_h + gap) + 6
    out = [f'<svg viewBox="0 0 {width} {h}" role="img" class="chart">']
    for i, (label, v) in enumerate(vals):
        y = i * (bar_h + gap) + 3
        w = max(2, (width - label_w - 90) * v / mx)
        out.append(
            f'<text x="{label_w - 8}" y="{y + bar_h - 6}" text-anchor="end" class="lbl">{esc(str(label)[:34])}</text>'
            f'<rect x="{label_w}" y="{y}" width="{w:.1f}" height="{bar_h}" rx="4" class="s1">'
            f'<title>{esc(label)}: {fmt(v, value_col)}</title></rect>'
            f'<text x="{label_w + w + 6:.1f}" y="{y + bar_h - 6}" class="val">{fmt(v, value_col)}</text>'
        )
    out.append("</svg>")
    return "".join(out)


def svg_line(rows, x_col, y_cols, width=680, height=240):
    xs = [r.get(x_col) for r in rows]
    if len(xs) < 2:
        return ""
    pad_l, pad_r, pad_t, pad_b = 56, 130, 14, 30
    plot_w, plot_h = width - pad_l - pad_r, height - pad_t - pad_b
    all_vals = [r.get(c) for r in rows for c in y_cols if isinstance(r.get(c), (int, float))]
    if not all_vals:
        return ""
    lo, hi = min(all_vals + [0]), max(all_vals)
    if hi == lo:
        hi = lo + 1
    span = (hi - lo) * 1.08 or 1

    def sx(i):
        return pad_l + plot_w * i / (len(xs) - 1)

    def sy(v):
        return pad_t + plot_h * (1 - (v - lo) / span)

    out = [f'<svg viewBox="0 0 {width} {height}" role="img" class="chart">']
    for g in range(5):
        gv = lo + (hi - lo) * 1.08 * g / 4
        gy = sy(gv)
        out.append(f'<line x1="{pad_l}" y1="{gy:.1f}" x2="{width - pad_r}" y2="{gy:.1f}" class="grid"/>')
        out.append(f'<text x="{pad_l - 6}" y="{gy + 4:.1f}" text-anchor="end" class="lbl">{fmt(round(gv, 2))}</text>')
    for i in (0, len(xs) // 2, len(xs) - 1):
        out.append(f'<text x="{sx(i):.1f}" y="{height - 8}" text-anchor="middle" class="lbl">{esc(xs[i])}</text>')
    for si, col in enumerate(y_cols[:4]):
        pts, markers = [], []
        for i, r in enumerate(rows):
            v = r.get(col)
            if not isinstance(v, (int, float)):
                continue
            x, y = sx(i), sy(v)
            pts.append(f"{x:.1f},{y:.1f}")
            markers.append(
                f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" class="s{si + 1}">'
                f"<title>{esc(xs[i])} — {esc(nice_label(col))}: {fmt(v, col)}</title></circle>"
            )
        if not pts:
            continue
        out.append(f'<polyline points="{" ".join(pts)}" fill="none" stroke-width="2" class="s{si + 1}-stroke"/>')
        out.extend(markers)
        lx, ly = pts[-1].split(",")
        out.append(
            f'<text x="{float(lx) + 8:.1f}" y="{float(ly) + 4:.1f}" class="slbl s{si + 1}-fill">'
            f"{esc(nice_label(col)[:20])}</text>"
        )
    out.append("</svg>")
    return "".join(out)


def kpi_tiles(row):
    tiles = []
    for col, v in row.items():
        tiles.append(
            f'<div class="tile"><div class="tile-v">{fmt(v, col)}</div>'
            f'<div class="tile-l">{esc(nice_label(col))}</div></div>'
        )
    return f'<div class="tiles">{"".join(tiles)}</div>'


def data_table(columns, rows, open_by_default=False):
    head = "".join(f"<th>{esc(nice_label(c))}</th>" for c in columns)
    body = "".join(
        "<tr>" + "".join(f"<td>{fmt(r.get(c), c)}</td>" for c in columns) + "</tr>"
        for r in rows
    )
    tbl = f'<div class="tblwrap"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'
    if open_by_default:
        return tbl
    return f"<details><summary>Data table ({len(rows)} rows)</summary>{tbl}</details>"


def pivot(rows, x_col, series_col, value_col):
    """(x, series, value) rows -> wide rows keyed on x, one column per series."""
    xs, series = [], {}
    for r in rows:
        x = r.get(x_col)
        if x not in xs:
            xs.append(x)
        series.setdefault(r.get(series_col), {})[x] = r.get(value_col)
    wide = [{x_col: x, **{s: series[s].get(x) for s in series}} for x in xs]
    return wide, list(series.keys())


def render_stat(stat):
    status = stat.get("status")
    card = [f'<section class="card" id="{esc(stat["id"])}">']
    card.append(f"<h3>{esc(stat.get('title'))}</h3><p class='q'>{esc(stat.get('question'))}</p>")

    if status != "ok":
        card.append(
            f'<p class="failed">✖ Query failed — {esc(stat.get("error", "unknown error"))}. '
            f"See the appendix for the SQL.</p></section>"
        )
        return "".join(card)

    rows, cols = stat.get("rows", []), stat.get("columns", [])
    render = stat.get("render", "table")
    if not rows:
        card.append('<p class="failed">Query ran but returned no rows.</p></section>')
        return "".join(card)

    if render == "kpis":
        card.append(kpi_tiles(rows[0]))
        if len(rows) > 1:
            card.append(data_table(cols, rows))
    elif render == "bars":
        label_col = stat.get("label_col", cols[0])
        value_col = stat.get("value_col", cols[1] if len(cols) > 1 else cols[0])
        card.append(svg_bars(rows, label_col, value_col))
        card.append(data_table(cols, rows))
    elif render == "line":
        if stat.get("series_col"):
            wide, series_names = pivot(rows, stat["x_col"], stat["series_col"], stat["value_col"])
            card.append(svg_line(wide, stat["x_col"], series_names[:4]))
            card.append(data_table([stat["x_col"]] + series_names, wide))
        else:
            card.append(svg_line(rows, stat["x_col"], stat.get("y_cols", cols[1:])))
            card.append(data_table(cols, rows))
    else:
        card.append(data_table(cols, rows, open_by_default=True))

    if stat.get("source_csv"):
        meta = f'imported from {stat["source_csv"]} at {stat.get("imported_at", "?")}'
    else:
        meta = f'{stat.get("elapsed_s", "?")}s · {stat.get("scanned_mb", "?")} MB scanned'
    if stat.get("used_fallback"):
        meta += " · fallback query used"
    card.append(f'<p class="meta">{esc(meta)}</p></section>')
    return "".join(card)


def render_pending(q):
    return (
        f'<section class="card pending" id="{esc(q["id"])}">'
        f"<h3>{esc(q.get('title'))}</h3><p class='q'>{esc(q.get('question'))}</p>"
        f'<p class="todo">Not run yet — paste <code>sql\\{esc(q["file"])}</code> into the '
        f"Athena query editor, download the results CSV into <code>output\\csv\\</code>, "
        f"then run <code>python run_all.py --import --report</code>.</p></section>"
    )


def connection_section(conn, manual_mode=False):
    if not conn:
        if manual_mode:
            return ("<section class='card'><h3>Connection</h3><p class='q'>"
                    "Manual mode: results were run in the Athena console and imported "
                    "from CSV downloads. No live AWS connection was used from this "
                    "machine.</p></section>")
        return ("<section class='card'><h3>Connection check</h3><p class='q'>"
                "No connection check found (run src/check_connection.py).</p></section>")
    parts = ['<section class="card"><h3>Connection check</h3>']
    for c in conn.get("checks", []):
        st = c.get("status", "?")
        badge = {"ok": "badge-ok", "warn": "badge-warn"}.get(st, "badge-fail")
        name = c.get("check")
        detail = ""
        if name == "sts":
            detail = f"account {esc(c.get('account_masked', ''))}"
        elif name == "databases":
            detail = f"{len(c.get('visible', []))} visible" + (
                f", missing {esc(c.get('missing'))}" if c.get("missing") else ""
            )
        elif name and name.startswith("table:"):
            detail = f"{c.get('column_count', '?')} columns" if st == "ok" else esc(c.get("error", ""))[:140]
        parts.append(
            f'<div class="chk"><span class="badge {badge}">{esc(st)}</span> '
            f"<strong>{esc(name)}</strong> <span class='meta'>{detail}</span></div>"
        )
    parts.append("</section>")
    return "".join(parts)


def appendix(stats):
    rows = []
    for s in stats:
        st = s.get("status", "?")
        badge = "badge-ok" if st == "ok" else "badge-fail"
        err = f'<p class="failed">{esc(s.get("error", ""))}</p>' if st != "ok" else ""
        rows.append(
            f'<details class="appx"><summary><span class="badge {badge}">{esc(st)}</span> '
            f'<code>{esc(s["id"])}</code> — {esc(s.get("title", ""))} '
            f'<span class="meta">({s.get("elapsed_s", "-")}s, {s.get("scanned_mb", "-")} MB)</span></summary>'
            f"{err}<pre>{esc(s.get('sql', ''))}</pre></details>"
        )
    return "".join(rows)


CSS = """
:root { color-scheme: light dark; }
.viz-root {
  --surface-1:#fcfcfb; --page:#f9f9f7; --ink-1:#0b0b0b; --ink-2:#52514e; --muted:#898781;
  --grid:#e1e0d9; --border:rgba(11,11,11,.10);
  --s1:#2a78d6; --s2:#1baf7a; --s3:#eda100; --s4:#e34948;
  --ok:#0ca30c; --fail:#d03b3b; --warn:#c98500;
}
@media (prefers-color-scheme: dark) { .viz-root {
  --surface-1:#1a1a19; --page:#0d0d0d; --ink-1:#ffffff; --ink-2:#c3c2b7; --muted:#898781;
  --grid:#2c2c2a; --border:rgba(255,255,255,.10);
  --s1:#3987e5; --s2:#199e70; --s3:#c98500; --s4:#e66767;
}}
* { box-sizing:border-box; margin:0; }
body { font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif; }
.viz-root { background:var(--page); color:var(--ink-1); min-height:100vh; padding:0 0 60px; }
header { padding:34px 24px 20px; max-width:1020px; margin:0 auto; }
header h1 { font-size:26px; } header p { color:var(--ink-2); margin-top:6px; max-width:75ch; }
nav { max-width:1020px; margin:0 auto 8px; padding:0 24px; display:flex; gap:14px; flex-wrap:wrap; }
nav a { color:var(--s1); text-decoration:none; font-size:14px; } nav a:hover { text-decoration:underline; }
main { max-width:1020px; margin:0 auto; padding:0 24px; }
h2 { font-size:20px; margin:34px 0 4px; }
.tier-intro { color:var(--ink-2); margin-bottom:14px; max-width:75ch; }
.card { background:var(--surface-1); border:1px solid var(--border); border-radius:10px;
        padding:18px 20px 12px; margin:14px 0; }
.card h3 { font-size:16px; } .q { color:var(--ink-2); font-size:13.5px; margin:4px 0 12px; max-width:75ch; }
.meta { color:var(--muted); font-size:12px; margin-top:8px; }
.failed { color:var(--fail); font-size:13.5px; }
.pending { border-style:dashed; }
.todo { color:var(--ink-2); font-size:13.5px; }
.todo code { font-size:12.5px; }
.tiles { display:flex; flex-wrap:wrap; gap:12px; margin:8px 0; }
.tile { border:1px solid var(--border); border-radius:8px; padding:10px 16px; min-width:130px; }
.tile-v { font-size:22px; font-weight:600; } .tile-l { font-size:12px; color:var(--ink-2); }
.chart { width:100%; height:auto; display:block; margin:6px 0; }
.chart .lbl { font-size:11px; fill:var(--muted); }
.chart .val { font-size:11px; fill:var(--ink-2); font-variant-numeric:tabular-nums; }
.chart .slbl { font-size:11px; font-weight:600; }
.chart .grid { stroke:var(--grid); stroke-width:1; }
.chart .s1 { fill:var(--s1); } .chart .s2 { fill:var(--s2); }
.chart .s3 { fill:var(--s3); } .chart .s4 { fill:var(--s4); }
.chart .s1-stroke { stroke:var(--s1); } .chart .s2-stroke { stroke:var(--s2); }
.chart .s3-stroke { stroke:var(--s3); } .chart .s4-stroke { stroke:var(--s4); }
.chart .s1-fill { fill:var(--s1); } .chart .s2-fill { fill:var(--s2); }
.chart .s3-fill { fill:var(--s3); } .chart .s4-fill { fill:var(--s4); }
.tblwrap { overflow-x:auto; margin:8px 0; }
table { border-collapse:collapse; font-size:13px; min-width:50%; }
th, td { padding:5px 12px; text-align:right; border-bottom:1px solid var(--grid); white-space:nowrap; }
th:first-child, td:first-child { text-align:left; }
th { color:var(--ink-2); font-weight:600; } td { font-variant-numeric:tabular-nums; }
details { margin:8px 0; } summary { cursor:pointer; color:var(--ink-2); font-size:13px; }
.badge { display:inline-block; font-size:11px; font-weight:700; padding:1px 8px; border-radius:9px;
         color:#fff; text-transform:uppercase; }
.badge-ok { background:var(--ok); } .badge-fail { background:var(--fail); } .badge-warn { background:var(--warn); }
.chk { padding:4px 0; font-size:14px; }
.appx pre { background:var(--page); border:1px solid var(--border); border-radius:8px;
            padding:12px; overflow-x:auto; font-size:12px; line-height:1.4; margin-top:8px; }
.appx code { font-size:13px; }
footer { max-width:1020px; margin:30px auto 0; padding:0 24px; color:var(--muted); font-size:12.5px; }
"""


def main():
    if not os.path.exists(settings.STORY_JSON):
        print(f"No {settings.STORY_JSON} found — run a fetch (--fetch) or a CSV "
              f"import (--import) first.")
        return 1
    with open(settings.STORY_JSON, encoding="utf-8") as f:
        story = json.load(f)
    conn = None
    if os.path.exists(settings.CONNECTION_JSON):
        with open(settings.CONNECTION_JSON, encoding="utf-8") as f:
            conn = json.load(f)

    try:
        with open(settings.MANIFEST, encoding="utf-8") as f:
            manifest = json.load(f)["queries"]
    except (OSError, ValueError, KeyError):
        manifest = []

    stats = story.get("stats", [])
    stats_by_id = {s.get("id"): s for s in stats}
    manual_mode = story.get("mode") == "manual-csv"
    ok_n = sum(1 for s in stats if s.get("status") == "ok")
    total = len(manifest) or len(stats)

    body = ['<div class="viz-root"><header>']
    body.append("<h1>BCS data story — accounts × calls × transcripts</h1>")
    mode_note = ("results run in the Athena console and imported from CSV downloads"
                 if manual_mode else
                 "windows anchored to each table's newest data")
    body.append(
        f"<p>A tiered walk through three tables: the card-account master, the contact-centre "
        f"call log, and the call transcripts. {ok_n} of {total} queries have results. "
        f"Last update {esc(story.get('generated_at', ''))} ({mode_note}) "
        f"· report built {date.today().isoformat()}.</p></header>"
    )
    body.append(
        '<nav><a href="#conn">Connection</a><a href="#tier1">1 · Shape of the data</a>'
        '<a href="#tier2">2 · Operational picture</a><a href="#tier3">3 · Cross-table story</a>'
        '<a href="#tier4">4 · Approach validation</a>'
        '<a href="#appendix">Appendix: queries</a></nav><main>'
    )

    body.append('<h2 id="conn">Connection</h2>')
    body.append(connection_section(conn, manual_mode))

    for tier in (1, 2, 3, 4):
        title, intro = TIER_INTRO[tier]
        body.append(f'<h2 id="tier{tier}">Tier {tier} — {title}</h2><p class="tier-intro">{esc(intro)}</p>')
        if manifest:
            # Manifest-driven: every registered query gets a card, run or not.
            for q in manifest:
                if q.get("tier") != tier:
                    continue
                s = stats_by_id.get(q["id"])
                body.append(render_stat(s) if s else render_pending(q))
        else:
            for s in stats:
                if s.get("tier") == tier:
                    body.append(render_stat(s))

    body.append('<h2 id="appendix">Appendix — every query, verbatim</h2>')
    body.append(appendix(stats))
    body.append("</main><footer>Generated locally from Athena aggregates. This file contains "
                "data extracts — keep it inside the environment; never commit the output folder."
                "</footer></div>")

    html_doc = (
        "<!doctype html><html lang='en'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width, initial-scale=1'>"
        f"<title>BCS data story</title><style>{CSS}</style></head><body>"
        + "".join(body) + "</body></html>"
    )
    with open(settings.REPORT_HTML, "w", encoding="utf-8") as f:
        f.write(html_doc)
    print(f"Report written  ->  {settings.REPORT_HTML}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
