"""
WAB HTML Story Generator
Standalone script that reads the three generated analysis workbooks and builds
simple multi-page HTML for internal discussion.

Optional input:
- WAB_Ops_UseCases_2026-03-18.xlsx

Dependencies: pandas, openpyxl
"""

from __future__ import annotations

import html
import re
from pathlib import Path
from typing import Callable

import pandas as pd

# ---------------------------------------------------------------------
# EDIT THESE PATHS
# ---------------------------------------------------------------------
INTERNAL_EXTRACT_XLSX = r"C:\Users\YourName\Desktop\wab_output\wab_internal_extract.xlsx"
CASES_DEEP_DIVE_XLSX = r"C:\Users\YourName\Desktop\wab_output\wab_cases_deep_dive.xlsx"
ENTITY_DEEP_DIVE_XLSX = r"C:\Users\YourName\Desktop\wab_output\wab_entity_deep_dive.xlsx"
USECASE_XLSX = r"C:\Users\YourName\Desktop\WAB_Ops_UseCases_2026-03-18.xlsx"  # optional
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_html_story"
SITE_TITLE = "WAB HOA Operations Data Story"
RUN_LABEL = "Phase 1 Internal Discussion Draft"


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------
def ensure_dir(path: str | Path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def norm(s: object) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).strip().lower())


def clean(s: object) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).replace("\r", " ").replace("\n", " ")).strip()


def trunc(s: object, n: int = 140) -> str:
    text = clean(s)
    return text if len(text) <= n else text[:n].rstrip() + "..."


def fmt(v: object) -> str:
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return f"{v:,}"
    if isinstance(v, float):
        if abs(v) >= 1000 and float(v).is_integer():
            return f"{int(v):,}"
        return f"{v:,.1f}"
    return clean(v)


def find_col(df: pd.DataFrame, *cands: str) -> str | None:
    lookup = {norm(c): c for c in df.columns}
    for cand in cands:
        n = norm(cand)
        if n in lookup:
            return lookup[n]
    for cand in cands:
        n = norm(cand)
        for k, v in lookup.items():
            if n in k or k in n:
                return v
    return None


class Book:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.exists = self.path.is_file()
        self.sheet_names: list[str] = []
        self._xls = None
        if self.exists:
            try:
                self._xls = pd.ExcelFile(self.path, engine="openpyxl")
                self.sheet_names = list(self._xls.sheet_names)
            except Exception:
                self.exists = False

    def get(self, sheet: str) -> pd.DataFrame:
        if not self.exists or self._xls is None or sheet not in self.sheet_names:
            return pd.DataFrame()
        try:
            return pd.read_excel(self._xls, sheet_name=sheet)
        except Exception:
            return pd.DataFrame()


class Ctx:
    def __init__(self):
        self.internal = Book(INTERNAL_EXTRACT_XLSX)
        self.cases = Book(CASES_DEEP_DIVE_XLSX)
        self.entity = Book(ENTITY_DEEP_DIVE_XLSX)
        self.usecase = Book(USECASE_XLSX)
        self.outdir = ensure_dir(OUTPUT_DIR)


# ---------------------------------------------------------------------
# HTML fragments
# ---------------------------------------------------------------------
def metric(label: str, value: object, note: str = "") -> str:
    note_html = f'<div class="metric-note">{html.escape(note)}</div>' if note else ""
    return f'<div class="metric"><div class="metric-label">{html.escape(label)}</div><div class="metric-value">{html.escape(fmt(value))}</div>{note_html}</div>'


def table_html(df: pd.DataFrame, max_rows: int = 20, title: str = "", trunc_chars: int = 140) -> str:
    if df is None or df.empty:
        return f'<div class="empty">No data available{f" for {html.escape(title)}" if title else ""}.</div>'
    d = df.copy().head(max_rows)
    for c in d.columns:
        d[c] = d[c].map(lambda x: trunc(fmt(x), trunc_chars))
    out = []
    if title:
        out.append(f"<h3>{html.escape(title)}</h3>")
    out.append('<div class="table-wrap"><table><thead><tr>')
    out.extend(f"<th>{html.escape(str(c))}</th>" for c in d.columns)
    out.append('</tr></thead><tbody>')
    for _, row in d.iterrows():
        out.append('<tr>')
        out.extend(f"<td>{html.escape(str(v))}</td>" for v in row.tolist())
        out.append('</tr>')
    out.append('</tbody></table></div>')
    return ''.join(out)


def section(title: str, body: str) -> str:
    return f'<section class="section"><h2>{html.escape(title)}</h2>{body}</section>'


def bullets(items: list[str]) -> str:
    return '<ul>' + ''.join(f'<li>{html.escape(x)}</li>' for x in items) + '</ul>'


def kv(items: list[tuple[str, object]]) -> str:
    rows = ''.join(f'<tr><th>{html.escape(k)}</th><td>{html.escape(fmt(v))}</td></tr>' for k, v in items)
    return f'<div class="table-wrap"><table class="kv">{rows}</table></div>'


def page_template(title: str, takeaway: str, body: str, slug: str, nav: list[tuple[str, str]]) -> str:
    nav_html = '<nav class="nav">' + ''.join(
        f'<a class="{"active" if s == slug else ""}" href="{s}.html">{html.escape(label)}</a>'
        for s, label in nav
    ) + '</nav>'
    return f'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)} | {html.escape(SITE_TITLE)}</title>
<style>
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;margin:0;background:#f8fafc;color:#1e293b;line-height:1.45}}
.top{{background:#0f172a;color:#fff;border-bottom:4px solid #2563eb}}
.shell{{max-width:1300px;margin:0 auto;padding:24px 28px 48px}}
.top .shell{{padding:18px 28px}}
h1{{margin:0;font-size:28px}} .sub{{margin-top:4px;font-size:14px;opacity:.85}}
.nav{{display:flex;flex-wrap:wrap;gap:8px;margin:18px 0 26px}} .nav a{{text-decoration:none;color:#0f172a;background:#fff;border:1px solid #cbd5e1;border-radius:8px;padding:8px 12px;font-size:14px}} .nav a.active{{background:#dbeafe;border-color:#60a5fa;font-weight:600}}
.hero,.section,.metric{{background:#fff;border:1px solid #e2e8f0;border-radius:14px}}
.hero{{padding:18px 20px;margin-bottom:20px}} .hero h2{{margin:0 0 8px;font-size:24px}} .hero p{{margin:0;font-size:16px;color:#334155}}
.metrics{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin:18px 0 8px}} .metric{{padding:14px 16px}} .metric-label{{color:#475569;font-size:13px;margin-bottom:8px}} .metric-value{{font-size:24px;font-weight:700;color:#0f172a}} .metric-note{{margin-top:6px;font-size:12px;color:#64748b}}
.section{{padding:18px 20px;margin-bottom:18px}} .section h2{{margin:0 0 12px;font-size:20px}} .section h3{{margin:18px 0 10px;font-size:16px}}
.table-wrap{{overflow-x:auto;margin-top:10px}} table{{width:100%;border-collapse:collapse;font-size:13px}} th,td{{border:1px solid #e2e8f0;padding:8px 10px;text-align:left;vertical-align:top}} th{{background:#f8fafc;font-weight:700}} table.kv th{{width:280px}}
ul{{margin:8px 0 0 18px;padding:0}} li{{margin:6px 0}} .empty{{color:#64748b;font-style:italic;padding:6px 0}}
</style>
</head>
<body>
<div class="top"><div class="shell"><h1>{html.escape(SITE_TITLE)}</h1><div class="sub">{html.escape(RUN_LABEL)}</div></div></div>
<div class="shell">{nav_html}<div class="hero"><h2>{html.escape(title)}</h2><p>{html.escape(takeaway)}</p></div>{body}</div>
</body>
</html>'''


# ---------------------------------------------------------------------
# Data helpers
# ---------------------------------------------------------------------
def story_lookup(df: pd.DataFrame, category: str, metric_name: str):
    if df.empty:
        return None
    c1 = find_col(df, "Category")
    c2 = find_col(df, "Metric")
    c3 = find_col(df, "Value")
    if not (c1 and c2 and c3):
        return None
    mask = df[c1].astype(str).str.strip().str.lower().eq(category.lower()) & df[c2].astype(str).str.strip().str.lower().eq(metric_name.lower())
    return df.loc[mask, c3].iloc[0] if mask.any() else None


def usecase_tables(book: Book) -> tuple[pd.DataFrame, pd.DataFrame]:
    return book.get("Top 20 v2"), book.get("Expanded Longlist v2")


# ---------------------------------------------------------------------
# Page builders
# ---------------------------------------------------------------------
def build_executive(ctx: Ctx) -> str:
    e11 = ctx.entity.get("E11_StoryNumbers")
    d15 = ctx.cases.get("D15_GenAI_Evidence")
    join = ctx.internal.get("4_JoinScorecard")
    metrics = []
    if not e11.empty:
        metrics.extend([
            metric("PMCs", story_lookup(e11, "PMC Universe", "Total")),
            metric("HOAs", story_lookup(e11, "HOA Universe", "Total")),
            metric("Client Cases", story_lookup(e11, "Cases (3mo)", "Client")),
            metric("Client Unresolved", story_lookup(e11, "Cases (3mo)", "Client unresolved")),
            metric("Emails (1 day)", story_lookup(e11, "Emails (1day)", "Total")),
        ])
    return ''.join([
        '<div class="metrics">' + ''.join(metrics) + '</div>' if metrics else '',
        section("Top Findings", bullets([
            "Cases provide the strongest internal signal: 3 months of operational history with rich subject, timing, and workload detail.",
            "The main business story is operational friction, backlog growth, and where that friction sits by subject and by client.",
            "Email text is usable for GenAI after preprocessing. Activity Subject is the strongest case-side text field.",
            "The deposit book is concentrated, so pilot prioritization should focus on high-value relationships.",
            "Only a subset of use cases is strongly supported by the current 4-file scope; others remain future-state ideas.",
        ])),
        section("GenAI Opportunity Snapshot", table_html(d15, max_rows=12)),
        section("Join Confidence Snapshot", table_html(join, max_rows=8)),
    ])


def build_scope(ctx: Ctx) -> str:
    summaries = []
    for label, df in [
        ("PMCs", ctx.internal.get("1A_PMC_Vitals")),
        ("HOAs", ctx.internal.get("1B_HOA_Vitals")),
        ("Cases", ctx.internal.get("1C_Case_Vitals")),
        ("Emails", ctx.internal.get("1D_Email_Vitals")),
    ]:
        summaries.append({"File": label, "Status": "Loaded" if not df.empty else "Missing", "Usable Columns": len(df) if not df.empty else "", "Sheet": f"{label} vitals"})
    return ''.join([
        section("File Coverage Summary", table_html(pd.DataFrame(summaries), max_rows=10)),
        section("Date Coverage", table_html(ctx.internal.get("2_DateCoverage"), max_rows=30)),
        section("Join Reliability", table_html(ctx.internal.get("4_JoinScorecard"), max_rows=12)),
        section("Entity Completeness", table_html(ctx.entity.get("E10_Completeness"), max_rows=20)),
        section("Caveats and Reliability Rules", bullets([
            "Cases reflect a 3-month operating window, not a long-run annual history.",
            "Emails reflect a 1-day sample and should be used for text and communication diagnostics, not trend analysis.",
            "PMCs and HOAs are current-state extracts, not longitudinal entity history.",
            "Case -> PMC is the main analytical path. Direct Case -> HOA linkage is weak and should not anchor the story.",
            "NAICS should not be treated as a primary segmentation lens.",
            "Some entity and deposit anomalies remain and should be treated as internal-discussion caveats, not executive-grade facts.",
        ])),
    ])


def build_population(ctx: Ctx) -> str:
    return ''.join([
        section("Population Split", '<p>The operational and GenAI story should be anchored on client cases, with internal/system cases handled separately.</p>' + table_html(ctx.cases.get("D01_PopulationSplit"), max_rows=30)),
        section("Interpretation", bullets([
            "Client-only metrics should drive backlog, SLA, and subject prioritization.",
            "Internal/system cases are still analytically useful, but mainly for understanding automation patterns and operational noise.",
            "Any future KPI reporting should make the population definition explicit.",
        ])),
    ])


def build_weekly(ctx: Ctx) -> str:
    return ''.join([
        section("Weekly Client Case Trend", table_html(ctx.cases.get("D02_ClientWeekly"), max_rows=20)),
        section("What This Means", bullets([
            "Weekly inflow is steady enough to support operational analysis and targeted intervention.",
            "The key question is not whether cases are being worked; it is whether the queue is being fully cleared.",
            "The backlog proxy and still-open pattern are the core leverage signals for GenAI productivity use cases.",
        ])),
    ])


def build_subjects(ctx: Ctx) -> str:
    return ''.join([
        section("Subject-Level Friction Profile", table_html(ctx.cases.get("D03_SubjectDeep"), max_rows=18)),
        section("Origin by Subject", table_html(ctx.cases.get("D10_OriginXSubject"), max_rows=18)),
        section("Interpretation Framework", bullets([
            "Fast and repetitive: good candidates for rules, routing, and limited copilot support.",
            "Fast median with a fat tail: best candidates for escalation prediction, summarization, and missing-info detection.",
            "Structurally slow: process redesign candidates first, AI support second.",
        ])),
    ])


def build_sla(ctx: Ctx) -> str:
    return ''.join([
        section("SLA Breach Profile", table_html(ctx.cases.get("D06_SLA_Breach"), max_rows=20)),
        section("Backlog Aging Detail", table_html(ctx.cases.get("D07_BacklogDetail"), max_rows=25)),
        section("Interpretation", '<p>Subjects with high breach rates and large aged unresolved populations are the clearest evidence of operational pain. These are the best places to test escalation alerts, proactive nudging, and workflow support.</p>'),
    ])


def build_workload(ctx: Ctx) -> str:
    return ''.join([
        section("Day-of-Week Pattern", table_html(ctx.cases.get("D04_DayOfWeek"), max_rows=10)),
        section("Hourly Pattern", table_html(ctx.cases.get("D05_HourlyPattern"), max_rows=24)),
        section("Owner and Pod Workload", table_html(ctx.cases.get("D09_OwnerWorkload"), max_rows=20)),
        section("Retouch / Rework Signal", table_html(ctx.cases.get("D08_Retouch"), max_rows=15)),
    ])


def build_email(ctx: Ctx) -> str:
    text_stats = ctx.internal.get("11_TextFieldStats")
    return ''.join([
        section("Email Overview", table_html(ctx.cases.get("D11_EmailOverview"), max_rows=20)),
        section("Email by Case Subject", table_html(ctx.cases.get("D12_EmailCaseSubjects"), max_rows=20)),
        section("Email Burden per Case", table_html(ctx.cases.get("D13_EmailBurden"), max_rows=20)),
        section("Text Field Stats", table_html(text_stats, max_rows=20)),
        section("Stripped Email / Text Samples", table_html(ctx.cases.get("D14_EmailTextSamples"), max_rows=22, trunc_chars=180)),
        section("Interpretation", bullets([
            "Email text is a stronger GenAI substrate than raw case Description.",
            "HTML stripping and security-banner removal are prerequisites, not optional enhancements.",
            "Activity Subject is the strongest case-side text field for operational copilot use cases.",
        ])),
    ])


def build_deposits(ctx: Ctx) -> str:
    return ''.join([
        section("Deposit Concentration", table_html(ctx.entity.get("E01_DepositConcentr"), max_rows=25)),
        section("Top PMCs", table_html(ctx.entity.get("E02_TopPMCs"), max_rows=25)),
        section("Interpretation", '<p>The deposit book should shape how friction is interpreted. A small number of clients carry a disproportionate share of economic value, so the same operational issue does not have the same strategic importance across all PMCs.</p>'),
    ])


def build_friction_value(ctx: Ctx) -> str:
    return ''.join([
        section("Friction vs Value", table_html(ctx.entity.get("E03_FrictionValue"), max_rows=30)),
        section("Hierarchy Depth", table_html(ctx.entity.get("E07_HierarchyDepth"), max_rows=20)),
        section("Interpretation", bullets([
            "High-value / high-friction PMCs are the clearest pilot targets.",
            "Low-value / high-friction PMCs may indicate operational drag or relationship-review candidates.",
            "HOA-per-PMC concentration helps explain where operational complexity compounds.",
        ])),
    ])


def build_geo_rm(ctx: Ctx) -> str:
    return ''.join([
        section("State Profile", table_html(ctx.entity.get("E04_StateProfile"), max_rows=20)),
        section("RM Coverage", table_html(ctx.entity.get("E05_RM_Coverage"), max_rows=20)),
        section("Platform Mix", table_html(ctx.entity.get("E08_PlatformMix"), max_rows=20)),
        section("Pod Geography", table_html(ctx.entity.get("E09_PodGeography"), max_rows=20)),
    ])


def build_usecases(ctx: Ctx) -> str:
    top20, longlist = usecase_tables(ctx.usecase)
    parts = [section("Evidence from Current Data", table_html(ctx.cases.get("D15_GenAI_Evidence"), max_rows=20))]
    parts.append(section("How to Read the Opportunity Map", bullets([
        "Strongly evidenced now: directly supported by the current 4-file scope.",
        "Partially evidenced: directionally supported but still missing history, labels, or process context.",
        "Not evidenced by these files: important ideas, but they depend on other systems or documents.",
    ])))
    if not top20.empty:
        parts.append(section("Top 20 Use Cases", table_html(top20, max_rows=20)))
    else:
        parts.append(section("Use Case Mapping", bullets([
            "If the optional use case workbook is not available, use this page as a direct evidence-to-opportunity map from D15_GenAI_Evidence.",
            "Strongest current candidates remain classification/routing, summarization, missing-info detection, escalation prediction, and RM briefing.",
        ])))
    if not longlist.empty:
        parts.append(section("Expanded Longlist Snapshot", table_html(longlist.head(20), max_rows=20)))
    return ''.join(parts)


def build_appendix_tables(ctx: Ctx) -> str:
    tables = [
        ("Internal Extract - PMC Concentration", ctx.internal.get("8_PMC_Concentration")),
        ("Internal Extract - NAICS Diagnostic", ctx.internal.get("9_NAICS_Diagnostic")),
        ("Entity Story Numbers", ctx.entity.get("E11_StoryNumbers")),
        ("Entity Completeness", ctx.entity.get("E10_Completeness")),
    ]
    return ''.join(section(title, table_html(df, max_rows=30)) for title, df in tables)


def build_appendix_methods(ctx: Ctx) -> str:
    return ''.join([
        section("Methods and Definitions", bullets([
            "The story assumes the three analysis workbooks were generated successfully from the VDI-side Python modules.",
            "HTML pages use workbook sheets as the system of record for the narrative.",
            "Missing sheets are skipped gracefully; absence of a section usually means the corresponding workbook output was unavailable.",
            "Case-heavy pages should be interpreted on the client-only population whenever the source sheet makes that distinction.",
            "Email insights are descriptive of a one-day communication sample and should not be generalized into long-run time-series claims.",
        ])),
        section("Expected Input Workbooks", kv([
            ("Internal extract workbook", INTERNAL_EXTRACT_XLSX),
            ("Cases deep-dive workbook", CASES_DEEP_DIVE_XLSX),
            ("Entity deep-dive workbook", ENTITY_DEEP_DIVE_XLSX),
            ("Use case workbook (optional)", USECASE_XLSX),
            ("Output directory", OUTPUT_DIR),
        ])),
    ])


# ---------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------
def write_page(ctx: Ctx, slug: str, title: str, takeaway: str, nav: list[tuple[str, str]], builder: Callable[[Ctx], str]) -> None:
    html_text = page_template(title, takeaway, builder(ctx), slug, nav)
    (ctx.outdir / f"{slug}.html").write_text(html_text, encoding="utf-8")


def main() -> None:
    ctx = Ctx()
    pages: list[tuple[str, str, str, Callable[[Ctx], str]]] = [
        ("executive_summary", "Executive Summary", "A full-data, internal discussion view of what the four files support and where GenAI opportunities are strongest.", build_executive),
        ("scope_reliability", "Data Scope, Caveats, and Join Reliability", "This page defines what the data is, what it is not, and how much trust to place in the main join paths.", build_scope),
        ("case_population", "Case Population and Operating Baseline", "Client-only cases should anchor the operational story; internal/system cases distort the raw view.", build_population),
        ("weekly_backlog", "Weekly Volume and Backlog", "The operational problem is not lack of activity; it is persistent backlog accumulation against otherwise steady inflow.", build_weekly),
        ("subject_friction", "Subject-Level Friction", "Subjects split into fast/repetitive, fat-tail, and structurally slow groups; GenAI should target the right group.", build_subjects),
        ("sla_aging", "SLA Breach and Aging", "Aged unresolved work and subject-specific breach patterns are the clearest pain signals in the dataset.", build_sla),
        ("workload_capacity", "Workload Patterns and Capacity Signals", "When work arrives and who carries it helps explain where service strain accumulates.", build_workload),
        ("email_text", "Email and Text Feasibility", "Email text is the richest current language asset, but preprocessing is required before it becomes a GenAI input.", build_email),
        ("deposit_clients", "Deposit and Client Concentration", "Economic concentration changes how operational friction should be prioritized across the client base.", build_deposits),
        ("friction_value", "Friction vs Value Prioritization", "Not all friction deserves the same response; the key is where friction intersects relationship value.", build_friction_value),
        ("geography_rm_platform", "Geography, RM Coverage, and Platform Signals", "Geography, relationship coverage, and platform mix provide context for where pilots should land first.", build_geo_rm),
        ("genai_opportunity_map", "GenAI Opportunity Map", "Current data strongly supports some use cases, partially supports others, and leaves some outside the current evidence boundary.", build_usecases),
        ("appendix_tables", "Appendix: Detailed Tables", "Supporting tables for readers who want more of the underlying output without leaving the HTML pack.", build_appendix_tables),
        ("appendix_methods", "Appendix: Methods and Definitions", "Definitions, input assumptions, and generator-level expectations.", build_appendix_methods),
    ]
    nav = [(slug, label) for slug, label, _, _ in pages]
    for slug, label, takeaway, builder in pages:
        write_page(ctx, slug, label, takeaway, nav, builder)
    (ctx.outdir / "index.html").write_text(f'<!doctype html><html><head><meta http-equiv="refresh" content="0; url={pages[0][0]}.html"></head><body><a href="{pages[0][0]}.html">Open story</a></body></html>', encoding="utf-8")
    print(f"HTML story written to: {ctx.outdir}")


if __name__ == "__main__":
    main()
