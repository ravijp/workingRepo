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


def para(text: str) -> str:
    return f"<p>{html.escape(text)}</p>"


def bullets(items: list[str]) -> str:
    return '<ul>' + ''.join(f'<li>{html.escape(x)}</li>' for x in items) + '</ul>'


def kv(items: list[tuple[str, object]]) -> str:
    rows = ''.join(f'<tr><th>{html.escape(k)}</th><td>{html.escape(fmt(v))}</td></tr>' for k, v in items)
    return f'<div class="table-wrap"><table class="kv">{rows}</table></div>'


def framing_panel(title: str, items: list[str]) -> str:
    if not items:
        return ""
    return (
        f'<div class="frame-card"><h3>{html.escape(title)}</h3>{bullets(items)}</div>'
    )


def section_frame(meta: dict[str, object]) -> str:
    return (
        '<div class="frame-grid">'
        + framing_panel("Assumptions", meta.get("assumptions", []))
        + framing_panel("Key Takeaways", meta.get("key_takeaways", meta.get("director_readout", [])))
        + framing_panel("Next Steps", meta.get("next_steps", []))
        + '</div>'
    )


def evidence_section(title: str, intro: str, df: pd.DataFrame, max_rows: int = 20, trunc_chars: int = 140) -> str:
    body = ""
    if intro:
        body += para(intro)
    body += table_html(df, max_rows=max_rows, trunc_chars=trunc_chars)
    return section(title, body)


def naics_appendix_html(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return section("Internal Extract - NAICS Diagnostic", '<div class="empty">No data available for NAICS diagnostics.</div>')

    block_col = find_col(df, "block")
    row_type_col = find_col(df, "row_type")
    item_col = find_col(df, "item")
    subitem_col = find_col(df, "subitem")
    count_col = find_col(df, "count")
    pct_col = find_col(df, "pct")
    value_col = find_col(df, "value")
    if not all([block_col, row_type_col, item_col, subitem_col, count_col, pct_col]):
        return evidence_section(
            "Internal Extract - NAICS Diagnostic",
            "NAICS is treated as a data-quality diagnostic first. This fallback view shows the raw diagnostic output because the expected long-format columns were not found.",
            df,
            max_rows=40,
        )

    parts: list[str] = [
        para("NAICS is a reliability check, not a primary segmentation lens. The table below is intentionally separated into summary, distribution, and company-type blocks so the discussion stays focused on whether the field is analytically usable.")
    ]

    def pick(block_name: str, row_types: list[str] | None = None) -> pd.DataFrame:
        sub = df[df[block_col].astype(str).eq(block_name)].copy()
        if row_types:
            sub = sub[sub[row_type_col].astype(str).isin(row_types)]
        cols = [c for c in [item_col, subitem_col, count_col, pct_col, value_col] if c]
        sub = sub[cols].copy()
        rename = {
            item_col: "item",
            subitem_col: "subitem",
            count_col: "count",
            pct_col: "pct",
            value_col: "value",
        }
        sub.rename(columns=rename, inplace=True)
        if "subitem" in sub.columns and sub["subitem"].replace("", pd.NA).isna().all():
            sub.drop(columns=["subitem"], inplace=True)
        if "value" in sub.columns and sub["value"].replace("", pd.NA).isna().all():
            sub.drop(columns=["value"], inplace=True)
        return sub

    parts.append(evidence_section(
        "PMC NAICS Summary",
        "This table answers a narrow question: is PMC NAICS clean enough to support segmentation, or is it mainly a CRM hygiene finding?",
        pick("PMC NAICS Summary"),
        max_rows=10,
    ))
    parts.append(evidence_section(
        "PMC Top NAICS",
        "This distribution shows whether residential property management dominates as expected or whether a long tail of implausible codes weakens the field.",
        pick("PMC Top 10 NAICS"),
        max_rows=12,
    ))
    parts.append(evidence_section(
        "PMC Company Type x NAICS",
        "This is the decisive diagnostic. If Management Company spreads across too many unrelated NAICS values, the field should stay in the data-quality lane.",
        pick("PMC CompanyType x NAICS", ["cross_tab_total", "cross_tab_detail"]),
        max_rows=30,
    ))
    parts.append(evidence_section(
        "HOA NAICS Summary",
        "HOA NAICS is expected to be far more uniform. The value of this table is confirming that expectation and quantifying the blank-rate.",
        pick("HOA NAICS Summary"),
        max_rows=10,
    ))
    parts.append(evidence_section(
        "HOA Top NAICS",
        "This distribution is mainly a confirmation check: one dominant HOA code is normal; a fragmented profile would be a surprise worth investigating.",
        pick("HOA Top 5 NAICS"),
        max_rows=10,
    ))
    return section("Internal Extract - NAICS Diagnostic", ''.join(parts))


def story_template(sections_html: str, nav: list[tuple[str, str]], intro_takeaway: str) -> str:
    nav_html = '<div class="nav-bar"><div class="nav-shell"><nav class="top-nav">' + ''.join(
        f'<a href="#section-{slug}">{html.escape(label)}</a>'
        for slug, label in nav
    ) + '</nav></div></div>'
    return f'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(SITE_TITLE)}</title>
<style>
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;margin:0;background:#f8fafc;color:#1e293b;line-height:1.5}}
.top{{background:#0f172a;color:#fff;border-bottom:4px solid #2563eb}}
.top-shell{{max-width:1600px;margin:0 auto;padding:18px 28px}}
h1{{margin:0;font-size:30px}}
.sub{{margin-top:6px;font-size:14px;opacity:.85}}
.nav-bar{{position:sticky;top:0;z-index:10;background:#f8fafc;border-bottom:1px solid #e2e8f0}}
.nav-shell{{max-width:1600px;margin:0 auto;padding:12px 28px}}
.top-nav{{display:flex;gap:8px;overflow-x:auto;padding-bottom:2px}}
.top-nav a{{flex:0 0 auto;text-decoration:none;color:#0f172a;background:#fff;border:1px solid #cbd5e1;border-radius:999px;padding:8px 12px;font-size:14px;white-space:nowrap}}
.top-nav a:hover{{background:#eff6ff;border-color:#93c5fd}}
.content{{max-width:1600px;margin:0 auto;padding:24px 28px 48px;min-width:0}}
.hero,.section,.story-section,.metric,.frame-card{{background:#fff;border:1px solid #e2e8f0;border-radius:16px}}
.hero{{padding:22px 24px;margin-bottom:20px}}
.hero h2{{margin:0 0 10px;font-size:26px}}
.hero p{{margin:0;color:#334155;font-size:16px}}
.story-section{{padding:22px 24px;margin-bottom:24px;scroll-margin-top:72px}}
.story-section-header{{margin-bottom:18px;padding-bottom:14px;border-bottom:1px solid #e2e8f0}}
.story-section-header h2{{margin:0 0 8px;font-size:26px}}
.story-section-header p{{margin:0;color:#475569;font-size:15px}}
.frame-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:12px;margin:18px 0 20px}}
.frame-card{{padding:14px 16px}}
.frame-card h3{{margin:0 0 10px;font-size:15px;color:#0f172a}}
.metrics{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin:18px 0 8px}}
.metric{{padding:14px 16px}}
.metric-label{{color:#475569;font-size:13px;margin-bottom:8px}}
.metric-value{{font-size:24px;font-weight:700;color:#0f172a}}
.metric-note{{margin-top:6px;font-size:12px;color:#64748b}}
.section{{padding:18px 20px;margin-bottom:18px}}
.section h2{{margin:0 0 12px;font-size:20px}}
.section h3{{margin:18px 0 10px;font-size:16px}}
.table-wrap{{overflow-x:auto;margin-top:10px}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th,td{{border:1px solid #e2e8f0;padding:8px 10px;text-align:left;vertical-align:top}}
th{{background:#f8fafc;font-weight:700}}
table.kv th{{width:280px}}
ul{{margin:8px 0 0 18px;padding:0}}
li{{margin:6px 0}}
.empty{{color:#64748b;font-style:italic;padding:6px 0}}
details{{margin-top:8px}}
summary{{cursor:pointer;font-weight:600;color:#0f172a}}
.appendix-body{{margin-top:14px}}
</style>
</head>
<body>
<div class="top"><div class="top-shell"><h1>{html.escape(SITE_TITLE)}</h1><div class="sub">{html.escape(RUN_LABEL)}</div></div></div>
{nav_html}
<main class="content">
  <div class="hero">
    <h2>Internal Discussion Story Pack</h2>
    <p>{html.escape(intro_takeaway)}</p>
  </div>
  {sections_html}
</main>
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
        evidence_section(
            "Weekly Client Case Trend",
            "This is the operating heartbeat of the business. It shows whether inflow, throughput, and queue growth are moving together or diverging over the 3-month window.",
            ctx.cases.get("D02_ClientWeekly"),
            max_rows=20,
        ),
        section("What This Means", bullets([
            "Weekly inflow is steady enough to support operational analysis and targeted intervention.",
            "The key question is not whether cases are being worked; it is whether the queue is being fully cleared.",
            "The backlog proxy and still-open pattern are the core leverage signals for GenAI productivity use cases.",
        ])),
    ])


def build_subjects(ctx: Ctx) -> str:
    return ''.join([
        evidence_section(
            "Subject-Level Friction Profile",
            "This table is here to separate high-volume repetitive work from genuinely slow or tail-heavy processes. It is the core prioritization lens for AI-assisted operations.",
            ctx.cases.get("D03_SubjectDeep"),
            max_rows=18,
        ),
        evidence_section(
            "Origin by Subject",
            "This cut shows whether subject complexity is tied to intake channel. That matters because routing and intake automation are only useful if the upstream channel carries signal.",
            ctx.cases.get("D10_OriginXSubject"),
            max_rows=18,
        ),
        section("Interpretation Framework", bullets([
            "Fast and repetitive: good candidates for rules, routing, and limited copilot support.",
            "Fast median with a fat tail: best candidates for escalation prediction, summarization, and missing-info detection.",
            "Structurally slow: process redesign candidates first, AI support second.",
        ])),
    ])


def build_sla(ctx: Ctx) -> str:
    return ''.join([
        evidence_section(
            "SLA Breach Profile",
            "This table is not about average performance. It isolates the subjects and queues that are breaking expected service thresholds often enough to matter.",
            ctx.cases.get("D06_SLA_Breach"),
            max_rows=20,
        ),
        evidence_section(
            "Backlog Aging Detail",
            "This view shows whether unresolved work is merely present or meaningfully stale. Aging tails are the cleanest evidence that intervention is needed, not just more descriptive reporting.",
            ctx.cases.get("D07_BacklogDetail"),
            max_rows=25,
        ),
        section("Interpretation", '<p>Subjects with high breach rates and large aged unresolved populations are the clearest evidence of operational pain. These are the best places to test escalation alerts, proactive nudging, and workflow support.</p>'),
    ])


def build_workload(ctx: Ctx) -> str:
    return ''.join([
        evidence_section(
            "Day-of-Week Pattern",
            "This table shows whether the work is evenly distributed across the week or whether operational pressure is concentrated into specific days.",
            ctx.cases.get("D04_DayOfWeek"),
            max_rows=10,
        ),
        evidence_section(
            "Hourly Pattern",
            "This view matters for staffing and real-time assistance. A sharp morning peak suggests where queue-building starts and where routing support may matter most.",
            ctx.cases.get("D05_HourlyPattern"),
            max_rows=24,
        ),
        evidence_section(
            "Owner and Pod Workload",
            "This table identifies who is carrying disproportionate operational load. It is here to separate case-mix issues from workload-distribution issues.",
            ctx.cases.get("D09_OwnerWorkload"),
            max_rows=20,
        ),
        evidence_section(
            "Retouch / Rework Signal",
            "Retouch is a useful secondary diagnostic. If it is low, the main issue is throughput. If it is high, the process may be creating rework loops that a copilot can address.",
            ctx.cases.get("D08_Retouch"),
            max_rows=15,
        ),
    ])


def build_email(ctx: Ctx) -> str:
    text_stats = ctx.internal.get("11_TextFieldStats")
    return ''.join([
        evidence_section(
            "Email Overview",
            "This table establishes the size and character of the one-day communication sample. It is here to anchor how much email-linked operating burden is visible at all.",
            ctx.cases.get("D11_EmailOverview"),
            max_rows=20,
        ),
        evidence_section(
            "Email by Case Subject",
            "This cut shows which operational case types generate disproportionate communication load. Those are the best candidates for summarization and response support.",
            ctx.cases.get("D12_EmailCaseSubjects"),
            max_rows=20,
        ),
        evidence_section(
            "Email Burden per Case",
            "This table is here to identify whether a small set of cases generates outsized communication churn. Those queues are often better AI targets than the highest-volume queues.",
            ctx.cases.get("D13_EmailBurden"),
            max_rows=20,
        ),
        evidence_section(
            "Text Field Stats",
            "These coverage and length metrics answer a gating question: is there enough language in the system to support summarization, classification, and missing-information detection?",
            text_stats,
            max_rows=20,
        ),
        evidence_section(
            "Stripped Email / Text Samples",
            "Sample text is included because feasibility cannot be judged from counts alone. The question is whether the cleaned text contains actionable prose or only metadata and fragments.",
            ctx.cases.get("D14_EmailTextSamples"),
            max_rows=22,
            trunc_chars=180,
        ),
        section("Interpretation", bullets([
            "Email text is a stronger GenAI substrate than raw case Description.",
            "HTML stripping and security-banner removal are prerequisites, not optional enhancements.",
            "Activity Subject is the strongest case-side text field for operational copilot use cases.",
        ])),
    ])


def build_deposits(ctx: Ctx) -> str:
    return ''.join([
        evidence_section(
            "Deposit Concentration",
            "This table quantifies how top-heavy the book is. It is here because operational friction has different strategic importance depending on where value sits.",
            ctx.entity.get("E01_DepositConcentr"),
            max_rows=25,
        ),
        evidence_section(
            "Top PMCs",
            "This view names the clients carrying that concentration. It should be read as a prioritization table, not just a leaderboard.",
            ctx.entity.get("E02_TopPMCs"),
            max_rows=25,
        ),
        section("Interpretation", '<p>The deposit book should shape how friction is interpreted. A small number of clients carry a disproportionate share of economic value, so the same operational issue does not have the same strategic importance across all PMCs.</p>'),
    ])


def build_friction_value(ctx: Ctx) -> str:
    return ''.join([
        evidence_section(
            "Friction vs Value",
            "This is the pilot-selection table. It brings operational burden and client value together so the discussion stays anchored on where intervention matters most.",
            ctx.entity.get("E03_FrictionValue"),
            max_rows=30,
        ),
        evidence_section(
            "Hierarchy Depth",
            "Hierarchy depth matters because operational complexity rises when a single PMC controls many HOAs. This helps explain why some clients are inherently harder to serve.",
            ctx.entity.get("E07_HierarchyDepth"),
            max_rows=20,
        ),
        section("Interpretation", bullets([
            "High-value / high-friction PMCs are the clearest pilot targets.",
            "Low-value / high-friction PMCs may indicate operational drag or relationship-review candidates.",
            "HOA-per-PMC concentration helps explain where operational complexity compounds.",
        ])),
    ])


def build_geo_rm(ctx: Ctx) -> str:
    return ''.join([
        evidence_section(
            "State Profile",
            "This table is the internal operating footprint by geography. It helps connect external market attractiveness with where WAB already carries entity concentration.",
            ctx.entity.get("E04_StateProfile"),
            max_rows=20,
        ),
        evidence_section(
            "RM Coverage",
            "Coverage discipline matters because stale relationship contact can compound service friction and client risk. This is a management signal as much as an analytics signal.",
            ctx.entity.get("E05_RM_Coverage"),
            max_rows=20,
        ),
        evidence_section(
            "Platform Mix",
            "Platform heterogeneity often creates workflow variation. This table helps determine whether a single copilot pattern is realistic or whether solutions need platform-specific tailoring.",
            ctx.entity.get("E08_PlatformMix"),
            max_rows=20,
        ),
        evidence_section(
            "Pod Geography",
            "This view links service structure to geography. It is here to help explain whether workload differences are tied to market concentration, organizational design, or both.",
            ctx.entity.get("E09_PodGeography"),
            max_rows=20,
        ),
    ])


def build_usecases(ctx: Ctx) -> str:
    top20, longlist = usecase_tables(ctx.usecase)
    parts = [evidence_section(
        "Evidence from Current Data",
        "This table is the evidence ledger. It connects the observed data signals to specific opportunity types so the use-case conversation stays grounded.",
        ctx.cases.get("D15_GenAI_Evidence"),
        max_rows=20,
    )]
    parts.append(section("How to Read the Opportunity Map", bullets([
        "Strongly evidenced now: directly supported by the current 4-file scope.",
        "Partially evidenced: directionally supported but still missing history, labels, or process context.",
        "Not evidenced by these files: important ideas, but they depend on other systems or documents.",
    ])))
    if not top20.empty:
        parts.append(evidence_section(
            "Top 20 Use Cases",
            "This optional table aligns the current evidence with the broader use-case inventory so the discussion can use the same labels and taxonomy as earlier work.",
            top20,
            max_rows=20,
        ))
    else:
        parts.append(section("Use Case Mapping", bullets([
            "If the optional use case workbook is not available, use this page as a direct evidence-to-opportunity map from D15_GenAI_Evidence.",
            "Strongest current candidates remain classification/routing, summarization, missing-info detection, escalation prediction, and RM briefing.",
        ])))
    if not longlist.empty:
        parts.append(evidence_section(
            "Expanded Longlist Snapshot",
            "The longlist is included only as a reference layer. It should not dilute focus on the small number of use cases already supported by the current data.",
            longlist.head(20),
            max_rows=20,
        ))
    return ''.join(parts)


def build_appendix_tables(ctx: Ctx) -> str:
    return ''.join([
        evidence_section(
            "Internal Extract - PMC Concentration",
            "This supporting table shows which company names drive the highest visible case volume in the raw internal extract.",
            ctx.internal.get("8_PMC_Concentration"),
            max_rows=30,
        ),
        naics_appendix_html(ctx.internal.get("9_NAICS_Diagnostic")),
        evidence_section(
            "Entity Story Numbers",
            "These are the headline entity-side numbers used across the story pack.",
            ctx.entity.get("E11_StoryNumbers"),
            max_rows=30,
        ),
        evidence_section(
            "Entity Completeness",
            "This table is here for transparency on entity-side data quality and coverage.",
            ctx.entity.get("E10_Completeness"),
            max_rows=30,
        ),
    ])


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
def render_story_section(ctx: Ctx, meta: dict[str, object]) -> str:
    slug = str(meta["slug"])
    title = str(meta["title"])
    takeaway = str(meta["takeaway"])
    body = meta["renderer"](ctx)
    if slug.startswith("appendix_"):
        body = f'<details><summary>Open appendix content</summary><div class="appendix-body">{body}</div></details>'
    return (
        f'<section id="section-{html.escape(slug)}" class="story-section">'
        f'<div class="story-section-header"><h2>{html.escape(title)}</h2><p>{html.escape(takeaway)}</p></div>'
        f'{section_frame(meta)}'
        f'{body}'
        '</section>'
    )


SECTION_REGISTRY: list[dict[str, object]] = [
    {
        "slug": "executive_summary",
        "title": "Executive Summary",
        "nav_label": "Executive Summary",
        "takeaway": "A full-data, internal discussion view of what the four files support and where GenAI opportunities are strongest.",
        "assumptions": [
            "This section is a synthesis layer, not a replacement for the detailed evidence sections.",
            "The current 4-file scope is strongest for operations and service insights, not full commercial relationship analytics.",
        ],
        "inputs": ["E11_StoryNumbers", "D15_GenAI_Evidence", "4_JoinScorecard"],
        "logic": [
            "Pull the smallest set of headline metrics that anchor scale, operational pain, and evidence-backed opportunity.",
            "Use the join scorecard and evidence table to avoid overclaiming from weakly-linked or thin signals.",
        ],
        "director_readout": [
            "The story is strong enough for internal alignment because the operational pain is real, concentrated, and measurable.",
            "Cases are the core signal, email text is usable after preprocessing, and the deposit book is concentrated enough to support targeted pilots.",
        ],
        "next_steps": [
            "Use this section as the starting point for internal walk-throughs, then move immediately into the cases and friction sections.",
            "Avoid debating individual use cases here; use this section to agree on the broad shape of the opportunity.",
        ],
        "renderer": build_executive,
    },
    {
        "slug": "scope_reliability",
        "title": "Data Scope, Caveats, and Join Reliability",
        "nav_label": "Scope and Reliability",
        "takeaway": "This section defines what the data is, what it is not, and how much trust to place in the main join paths.",
        "assumptions": [
            "This is the single source of truth for data scope, reliability, and known caveats.",
            "Later sections should not repeat these caveats unless a section-specific warning materially changes interpretation.",
        ],
        "inputs": ["1A_PMC_Vitals, 1B_HOA_Vitals, 1C_Case_Vitals, 1D_Email_Vitals", "2_DateCoverage", "4_JoinScorecard", "E10_Completeness"],
        "logic": [
            "Establish what is snapshot vs event data, then show whether the main analytical joins are trustworthy enough to support the story.",
            "Capture reliability issues early so downstream sections can stay insight-heavy rather than caveat-heavy.",
        ],
        "director_readout": [
            "The dataset is strong enough for internal use-case discovery, but it is not a universal source for every commercial or operational claim.",
            "The main join graph is usable; the main limitations are time-window scope, text preprocessing needs, and a few entity/deposit anomalies.",
        ],
        "next_steps": [
            "Use this section to settle any questions about scope before debating insights or use cases.",
            "Keep a running punch list of data cleanup items, but do not let that punch list stall the operational story.",
        ],
        "renderer": build_scope,
    },
    {
        "slug": "case_population",
        "title": "Case Population and Operating Baseline",
        "nav_label": "Case Population",
        "takeaway": "Client-only cases should anchor the operational story; internal/system cases distort the raw view.",
        "assumptions": ["Internal/system cases should not be mixed into the core client-facing operating narrative."],
        "inputs": ["D01_PopulationSplit"],
        "logic": ["Separate client workload from internal/system-generated workload before interpreting backlog, subject mix, or GenAI relevance."],
        "director_readout": [
            "Client-only cases are slower, operationally richer, and more relevant to business value than internal/system cases.",
            "If later sections are not client-filtered, they will overstate some pain points and mis-prioritize interventions.",
        ],
        "next_steps": [
            "Use the client-only lens as the default for all operational interpretation.",
            "Keep internal/system cases as a separate stream for automation and process noise analysis.",
        ],
        "renderer": build_population,
    },
    {
        "slug": "weekly_backlog",
        "title": "Weekly Volume and Backlog",
        "nav_label": "Weekly Backlog",
        "takeaway": "The operational problem is not lack of activity; it is persistent backlog accumulation against otherwise steady inflow.",
        "assumptions": ["The weekly view is the cleanest way to show pace, queue growth, and whether the operation is truly catching up."],
        "inputs": ["D02_ClientWeekly"],
        "logic": ["Compare created, resolved, backlog proxy, and still-open patterns over time rather than relying on point-in-time case counts."],
        "director_readout": [
            "The team is working volume consistently, but the queue is not being fully cleared.",
            "This is the operational-leverage section: even modest cycle-time improvements would compound against backlog growth.",
        ],
        "next_steps": [
            "Use this section to justify why productivity-focused GenAI use cases matter even if median resolution times look reasonable.",
            "Track whether the backlog growth is driven by specific subjects, pods, or clients in later sections.",
        ],
        "renderer": build_weekly,
    },
    {
        "slug": "subject_friction",
        "title": "Subject-Level Friction",
        "nav_label": "Subject Friction",
        "takeaway": "Subjects split into fast/repetitive, fat-tail, and structurally slow groups; GenAI should target the right group.",
        "assumptions": [
            "Subject is the cleanest operational segmentation axis in the current data.",
            "GenAI suitability is driven by both workflow shape and text coverage, not by subject counts alone.",
        ],
        "inputs": ["D03_SubjectDeep", "D10_OriginXSubject"],
        "logic": ["Group work by subject, compare volume to cycle-time tail behavior, then layer in text-field coverage and origin mix."],
        "director_readout": [
            "The operation is not uniformly broken; it splits into fast/repetitive, fat-tail, and structurally slow process families.",
            "GenAI should target the fat-tail and repetitive classes first, while the structurally slow classes need process redesign support.",
        ],
        "next_steps": [
            "Use this section to shortlist 3-5 subjects for pilot design.",
            "Do not greenlight a single generic copilot across all subjects; keep the pilot subject-specific.",
        ],
        "renderer": build_subjects,
    },
    {
        "slug": "sla_aging",
        "title": "SLA Breach and Aging",
        "nav_label": "SLA and Aging",
        "takeaway": "Aged unresolved work and subject-specific breach patterns are the clearest pain signals in the dataset.",
        "assumptions": ["SLA breach patterns and backlog aging are stronger pain indicators than average resolution time."],
        "inputs": ["D06_SLA_Breach", "D07_BacklogDetail"],
        "logic": ["Show where the queue is not just active, but persistently late and aging beyond reasonable operating thresholds."],
        "director_readout": [
            "This section converts friction into urgency.",
            "Subjects with high breach rates and old unresolved tails are the clearest targets for escalation prediction, queue nudging, and workflow intervention.",
        ],
        "next_steps": [
            "Anchor escalation and SLA-breach use-case discussions here.",
            "Use the oldest unresolved subject/client combinations as candidate pilot queues.",
        ],
        "renderer": build_sla,
    },
    {
        "slug": "workload_capacity",
        "title": "Workload Patterns and Capacity Signals",
        "nav_label": "Workload Patterns",
        "takeaway": "When work arrives and who carries it helps explain where service strain accumulates.",
        "assumptions": ["Workload shape matters because a useful copilot must land where the operation actually experiences volume concentration."],
        "inputs": ["D04_DayOfWeek", "D05_HourlyPattern", "D09_OwnerWorkload", "D08_Retouch"],
        "logic": ["Use temporal concentration and owner/pod skew to show where staffing or AI support would actually be felt."],
        "director_readout": [
            "This is a weekday, morning-heavy operation with meaningful owner/pod concentration.",
            "The core problem looks like throughput and distribution, not poor resolution quality or extensive rework.",
        ],
        "next_steps": [
            "Use this section to think about where real-time triage or summarization would have the most operational impact.",
            "Investigate whether slow pods are slow because of case mix, geography, or staffing/capacity.",
        ],
        "renderer": build_workload,
    },
    {
        "slug": "email_text",
        "title": "Email and Text Feasibility",
        "nav_label": "Email and Text",
        "takeaway": "Email text is the richest current language asset, but preprocessing is required before it becomes a GenAI input.",
        "assumptions": ["The email file is a one-day communication sample, so this section is about text feasibility and communication burden, not long-run trend."],
        "inputs": ["D11_EmailOverview", "D12_EmailCaseSubjects", "D13_EmailBurden", "D14_EmailTextSamples", "11_TextFieldStats"],
        "logic": ["Evaluate whether the text exists, whether it is rich enough after preprocessing, and which case types generate the most communication load."],
        "director_readout": [
            "Email is the strongest current language source for GenAI in this dataset.",
            "The main gating requirement is not data existence; it is reliable HTML stripping and banner removal.",
        ],
        "next_steps": [
            "Use this section to justify summarization and draft-assist experiments.",
            "Do not frame email findings as time-series results; keep them in the communication-feasibility lane.",
        ],
        "renderer": build_email,
    },
    {
        "slug": "deposit_clients",
        "title": "Deposit and Client Concentration",
        "nav_label": "Deposit and Clients",
        "takeaway": "Economic concentration changes how operational friction should be prioritized across the client base.",
        "assumptions": ["Economic concentration changes the priority of operational pain; identical friction has different meaning on a $300M client vs a $50K client."],
        "inputs": ["E01_DepositConcentr", "E02_TopPMCs"],
        "logic": ["First quantify how concentrated the book is, then name the clients carrying that concentration and show their operating footprint."],
        "director_readout": [
            "This is not a flat book. A small number of PMCs carry a disproportionate share of value.",
            "That makes client selection critical for any GenAI pilot or RM intervention story.",
        ],
        "next_steps": [
            "Use this section to separate strategically important clients from the long tail before deciding where to pilot.",
            "Flag any large anomalous entities for cleanup so they do not distort leadership discussions.",
        ],
        "renderer": build_deposits,
    },
    {
        "slug": "friction_value",
        "title": "Friction vs Value Prioritization",
        "nav_label": "Friction vs Value",
        "takeaway": "Not all friction deserves the same response; the key is where friction intersects relationship value.",
        "assumptions": ["Operational friction should be evaluated relative to relationship value, not in isolation."],
        "inputs": ["E03_FrictionValue", "E07_HierarchyDepth"],
        "logic": ["Combine cases, deposits, and entity complexity so prioritization reflects both pain and business importance."],
        "director_readout": [
            "This section is where the operational story meets the client/economic story.",
            "The right pilot candidates are not merely noisy clients; they are valuable clients with meaningful, repeated friction.",
        ],
        "next_steps": [
            "Use this section to shortlist pilot PMCs and to identify low-value/high-friction relationships for review.",
            "Check whether high-friction PMCs cluster in particular subjects, platforms, or states.",
        ],
        "renderer": build_friction_value,
    },
    {
        "slug": "geography_rm_platform",
        "title": "Geography, RM Coverage, and Platform Signals",
        "nav_label": "Geo RM Platform",
        "takeaway": "Geography, relationship coverage, and platform mix provide context for where pilots should land first.",
        "assumptions": ["Geography, RM coverage, and platform mix are context layers that explain where friction happens and how interventions should land."],
        "inputs": ["E04_StateProfile", "E05_RM_Coverage", "E08_PlatformMix", "E09_PodGeography"],
        "logic": ["Bridge the entity/economic picture to market footprint, RM practice, platform heterogeneity, and pod geography."],
        "director_readout": [
            "This section turns the client list into an operating map.",
            "It also surfaces a second story beyond GenAI: RM coverage discipline and geographic concentration are themselves material management issues.",
        ],
        "next_steps": [
            "Use this section to connect internal operational burden to your external market story.",
            "Treat stale RM coverage and concentrated high-friction states as candidate management actions, not just analytics findings.",
        ],
        "renderer": build_geo_rm,
    },
    {
        "slug": "genai_opportunity_map",
        "title": "GenAI Opportunity Map",
        "nav_label": "GenAI Map",
        "takeaway": "Current data strongly supports some use cases, partially supports others, and leaves some outside the current evidence boundary.",
        "assumptions": [
            "This section is a translation layer from observed data signals to practical GenAI opportunities.",
            "The goal is not to prove every use case, but to sort which ones are strongly evidenced, partially evidenced, or still outside the current data boundary.",
        ],
        "inputs": ["D15_GenAI_Evidence", "Top 20 v2 (optional)", "Expanded Longlist v2 (optional)"],
        "logic": ["Start with what the current data directly supports, then map that evidence onto the existing use-case taxonomy if available."],
        "director_readout": [
            "A small number of use cases are strongly supported already: routing/classification, summarization, missing-info detection, escalation prediction, and RM briefing.",
            "The rest should be treated as adjacent hypotheses or future-state ideas until supported by broader data access.",
        ],
        "next_steps": [
            "Use this section to converge on 3-5 pilots, not to reopen the full longlist.",
            "Keep a clear line between evidence-backed near-term candidates and strategically important but under-evidenced use cases.",
        ],
        "renderer": build_usecases,
    },
    {
        "slug": "appendix_tables",
        "title": "Appendix: Detailed Tables",
        "nav_label": "Appendix Tables",
        "takeaway": "Supporting tables for readers who want more of the underlying output without leaving the HTML pack.",
        "assumptions": ["This appendix is for depth and transparency, not for first-pass storytelling."],
        "inputs": ["8_PMC_Concentration", "9_NAICS_Diagnostic", "E11_StoryNumbers", "E10_Completeness"],
        "logic": ["Keep supporting evidence accessible without forcing readers back into the Excel workbooks."],
        "director_readout": [
            "This section supports discussion when someone wants to go deeper on specific supporting tables.",
            "It should not lead the meeting; it should support it.",
        ],
        "next_steps": [
            "Use this appendix to answer detail questions without cluttering the core flow.",
            "Keep adding only supporting evidence here, not new narrative layers.",
        ],
        "renderer": build_appendix_tables,
    },
    {
        "slug": "appendix_methods",
        "title": "Appendix: Methods and Definitions",
        "nav_label": "Appendix Methods",
        "takeaway": "Definitions, input assumptions, and generator-level expectations.",
        "assumptions": ["Method notes belong in one place so the main sections stay readable."],
        "inputs": ["Generator configuration", "Input workbook paths", "Section rendering rules"],
        "logic": ["Keep portability, assumptions, and interpretation rules transparent in one methods appendix."],
        "director_readout": [
            "This section exists so the rest of the artifact can stay focused on insight rather than mechanics.",
        ],
        "next_steps": [
            "Use this appendix when questions arise about provenance, method, or expected workbook inputs.",
        ],
        "renderer": build_appendix_methods,
    },
]


def main() -> None:
    ctx = Ctx()
    nav = [(str(section["slug"]), str(section["nav_label"])) for section in SECTION_REGISTRY]
    sections_html = ''.join(render_story_section(ctx, section) for section in SECTION_REGISTRY)
    intro = "One self-contained discussion artifact that connects data scope, operational pain, client concentration, and GenAI opportunity in a single shareable file."
    story_html = story_template(sections_html, nav, intro)
    (ctx.outdir / "story.html").write_text(story_html, encoding="utf-8")
    (ctx.outdir / "index.html").write_text('<!doctype html><html><head><meta http-equiv="refresh" content="0; url=story.html"></head><body><a href="story.html">Open story</a></body></html>', encoding="utf-8")
    print(f"HTML story written to: {ctx.outdir / 'story.html'}")


if __name__ == "__main__":
    main()
