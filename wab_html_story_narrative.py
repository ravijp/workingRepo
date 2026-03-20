"""
WAB HTML Story Generator - Narrative Rewrite

Single-file HTML artifact optimized for internal discussion.
This version keeps the current workbook-reading infrastructure but rewrites
the story into a sharper 9-section narrative with inline evidence.

Dependencies: pandas, openpyxl
"""

from __future__ import annotations

import html

import pandas as pd

import wab_html_story as base


# ---------------------------------------------------------------------
# EDIT THESE PATHS
# ---------------------------------------------------------------------
INTERNAL_EXTRACT_XLSX = r"C:\Users\YourName\Desktop\wab_output\wab_internal_extract.xlsx"
CASES_DEEP_DIVE_XLSX = r"C:\Users\YourName\Desktop\wab_output\wab_cases_deep_dive.xlsx"
ENTITY_DEEP_DIVE_XLSX = r"C:\Users\YourName\Desktop\wab_output\wab_entity_deep_dive.xlsx"
USECASE_XLSX = r"C:\Users\YourName\Desktop\WAB_Ops_UseCases_2026-03-18.xlsx"  # optional
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_html_story"
SITE_TITLE = "WAB HOA Operations Data Story"
RUN_LABEL = "Narrative Discussion Draft"
OUTPUT_HTML_NAME = "story_narrative.html"
REDIRECT_HTML_NAME = "index_narrative.html"


class Ctx:
    def __init__(self) -> None:
        self.internal = base.Book(INTERNAL_EXTRACT_XLSX)
        self.cases = base.Book(CASES_DEEP_DIVE_XLSX)
        self.entity = base.Book(ENTITY_DEEP_DIVE_XLSX)
        self.usecase = base.Book(USECASE_XLSX)
        self.outdir = base.ensure_dir(OUTPUT_DIR)


def lead_block(paragraphs: list[str]) -> str:
    return "".join(f"<p>{html.escape(p)}</p>" for p in paragraphs if p)


def evidence_table(
    title: str,
    intro: str,
    df: pd.DataFrame,
    *,
    max_rows: int = 20,
    trunc_chars: int = 140,
    notes: list[str] | None = None,
) -> str:
    body = ""
    if intro:
        body += f"<p>{html.escape(intro)}</p>"
    body += base.table_html(df, max_rows=max_rows, trunc_chars=trunc_chars)
    if notes:
        body += column_guide(notes)
    return base.section(title, body)


def so_what(text: str) -> str:
    return (
        '<div class="callout so-what">'
        '<div class="callout-title">So what?</div>'
        f"<p>{html.escape(text)}</p>"
        "</div>"
    )


def recommended_action(items: list[str]) -> str:
    return (
        '<div class="callout action">'
        '<div class="callout-title">Recommended action</div>'
        f"{base.bullets(items)}"
        "</div>"
    )


def column_guide(items: list[str]) -> str:
    return (
        '<div class="column-guide">'
        '<div class="guide-title">How to read this table</div>'
        f"{base.bullets(items)}"
        "</div>"
    )


def caveats_footer(items: list[str]) -> str:
    return (
        '<details class="caveats">'
        '<summary>Caveats and interpretation boundaries</summary>'
        f"{base.bullets(items)}"
        "</details>"
    )


def metric_grid(ctx: Ctx) -> str:
    e11 = ctx.entity.get("E11_StoryNumbers")
    metrics = [
        base.metric("PMCs", base.story_lookup(e11, "PMC Universe", "Total")),
        base.metric("HOAs", base.story_lookup(e11, "HOA Universe", "Total")),
        base.metric("Client Cases", base.story_lookup(e11, "Cases (3mo)", "Client")),
        base.metric("Unresolved Cases", base.story_lookup(e11, "Cases (3mo)", "Client unresolved")),
        base.metric("Emails", base.story_lookup(e11, "Emails (1day)", "Total"), "one day sample"),
        base.metric("Deposit Book", base.story_lookup(e11, "Deposits", "Total")),
    ]
    return '<div class="metrics">' + "".join(metrics) + "</div>"


def section_shell(slug: str, title: str, takeaway: str, body: str) -> str:
    return (
        f'<section id="section-{html.escape(slug)}" class="story-section">'
        f'<div class="story-section-header"><h2>{html.escape(title)}</h2><p>{html.escape(takeaway)}</p></div>'
        f"{body}"
        "</section>"
    )


def nav_html(sections: list[tuple[str, str]]) -> str:
    return (
        '<div class="nav-bar"><div class="nav-shell"><nav class="top-nav">'
        + "".join(f'<a href="#section-{slug}">{html.escape(label)}</a>' for slug, label in sections)
        + "</nav></div></div>"
    )


def render_story(ctx: Ctx) -> str:
    sections = [
        ("exec_summary", "Executive Summary", render_exec_summary(ctx)),
        ("case_volume", "Case Volume and Backlog", render_case_volume(ctx)),
        ("subject_friction", "Subject-Level Friction", render_subject_friction(ctx)),
        ("workload_capacity", "Channels, Workload, and Capacity", render_workload_capacity(ctx)),
        ("data_quality", "Data Quality and Joinability", render_data_quality(ctx)),
        ("pmc_portfolio", "Client Portfolio and Risk", render_pmc_portfolio(ctx)),
        ("email_text", "Email and Text Feasibility", render_email_text(ctx)),
        ("usecase_map", "GenAI Use Case Evidence Map", render_usecase_map(ctx)),
        ("next_steps", "Recommended Next Steps", render_next_steps(ctx)),
    ]
    nav = [(slug, title) for slug, title, _ in sections]
    sections_html = "".join(section_shell(slug, title, takeaway, body) for slug, title, (takeaway, body) in sections)
    html_text = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(SITE_TITLE)}</title>
<style>
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;margin:0;background:#f8fafc;color:#1e293b;line-height:1.6}}
.top{{background:#0f172a;color:#fff;border-bottom:4px solid #2563eb}}
.top-shell{{max-width:1520px;margin:0 auto;padding:18px 28px}}
h1{{margin:0;font-size:30px}}
.sub{{margin-top:6px;font-size:14px;opacity:.85}}
.nav-bar{{position:sticky;top:0;z-index:10;background:#f8fafc;border-bottom:1px solid #e2e8f0}}
.nav-shell{{max-width:1520px;margin:0 auto;padding:12px 28px}}
.top-nav{{display:flex;gap:8px;overflow-x:auto;padding-bottom:2px}}
.top-nav a{{flex:0 0 auto;text-decoration:none;color:#0f172a;background:#fff;border:1px solid #cbd5e1;border-radius:999px;padding:8px 12px;font-size:14px;white-space:nowrap}}
.top-nav a:hover{{background:#eff6ff;border-color:#93c5fd}}
.content{{max-width:1520px;margin:0 auto;padding:24px 28px 48px}}
.hero,.story-section,.section,.metric,.callout,.column-guide,.caveats{{background:#fff;border:1px solid #e2e8f0;border-radius:16px}}
.hero{{padding:22px 24px;margin-bottom:20px}}
.hero h2{{margin:0 0 10px;font-size:26px}}
.hero p{{margin:0;color:#334155;font-size:16px}}
.story-section{{padding:22px 24px;margin-bottom:24px;scroll-margin-top:72px}}
.story-section-header{{margin-bottom:18px;padding-bottom:14px;border-bottom:1px solid #e2e8f0}}
.story-section-header h2{{margin:0 0 8px;font-size:26px}}
.story-section-header p{{margin:0;color:#475569;font-size:15px}}
.section{{padding:18px 20px;margin-bottom:18px}}
.section h2{{margin:0 0 12px;font-size:20px}}
.section h3{{margin:18px 0 10px;font-size:16px}}
.metrics{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin:6px 0 18px}}
.metric{{padding:14px 16px}}
.metric-label{{color:#475569;font-size:13px;margin-bottom:8px}}
.metric-value{{font-size:24px;font-weight:700;color:#0f172a}}
.metric-note{{margin-top:6px;font-size:12px;color:#64748b}}
.table-wrap{{overflow-x:auto;margin-top:10px}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th,td{{border:1px solid #e2e8f0;padding:8px 10px;text-align:left;vertical-align:top}}
th{{background:#f8fafc;font-weight:700}}
table.kv th{{width:280px}}
ul{{margin:8px 0 0 18px;padding:0}}
li{{margin:6px 0}}
p{{margin:0 0 14px}}
.empty{{color:#64748b;font-style:italic;padding:6px 0}}
.callout{{padding:16px 18px;margin:14px 0 18px}}
.callout-title{{font-size:14px;font-weight:700;text-transform:uppercase;letter-spacing:.03em;margin-bottom:8px}}
.so-what{{border-color:#bfdbfe;background:#eff6ff}}
.action{{border-color:#c7eed8;background:#f0fdf4}}
.column-guide{{padding:14px 16px;margin-top:12px;background:#f8fafc}}
.guide-title{{font-weight:700;margin-bottom:6px}}
.caveats{{padding:16px 18px;margin-top:18px}}
.caveats summary{{cursor:pointer;font-weight:700}}
</style>
</head>
<body>
<div class="top"><div class="top-shell"><h1>{html.escape(SITE_TITLE)}</h1><div class="sub">{html.escape(RUN_LABEL)}</div></div></div>
{nav_html(nav)}
<main class="content">
<div class="hero">
  <h2>Insights, evidence, and practical AI opportunities</h2>
  <p>This version is written as an analytical briefing rather than a dashboard. It uses the generated workbooks as evidence, but it treats the evidence as support for an argument: where the operational pressure sits, why it matters commercially, and which GenAI interventions are actually supported by the data.</p>
</div>
{sections_html}
{render_caveats()}
</main>
</body>
</html>"""
    return html_text


def render_exec_summary(ctx: Ctx) -> tuple[str, str]:
    d15 = ctx.cases.get("D15_GenAI_Evidence")
    join = ctx.internal.get("4_JoinScorecard")
    body = "".join([
        metric_grid(ctx),
        lead_block([
            "Over the past three months, the HOA operations team processed 36,296 client-facing cases. That volume is manageable in isolation. Most cases still resolve in under five hours. The problem is that the queue never fully clears. Each week, case creation runs slightly ahead of case resolution, and that shortfall compounds into a backlog large enough to matter.",
            "That backlog sits on top of a concentrated deposit book where service friction on a small number of large PMC relationships carries disproportionate commercial risk. The encouraging part is that the data is usable. Cases, emails, and PMC records can be linked well enough to support targeted AI interventions now, especially around triage, summarization, missing-information detection, and escalation.",
        ]),
        evidence_table(
            "GenAI Evidence Snapshot",
            "This table is the shortest expression of what the current files actually support. It should be read as an evidence ledger, not a wish list.",
            d15,
            max_rows=12,
        ),
        evidence_table(
            "Join Confidence Snapshot",
            "These are the operational joins that determine whether the analysis can move from description into action.",
            join,
            max_rows=8,
            notes=[
                "match rate = rows with a populated key that successfully link",
                "the most important analytical path is Email -> Case -> PMC",
                "Case -> HOA is weak and should not anchor the story",
            ],
        ),
        so_what(
            "The document that follows is not making a generic case for AI. It is making a narrower claim: the current data is good enough to support a small number of high-value interventions now, and those interventions sit directly on top of the operational bottlenecks shown in the evidence."
        ),
    ])
    takeaway = "The operating issue is not raw volume. It is a queue that grows slightly every week, attached to a highly concentrated client book."
    return takeaway, body


def render_case_volume(ctx: Ctx) -> tuple[str, str]:
    pop = ctx.cases.get("D01_PopulationSplit")
    weekly = ctx.cases.get("D02_ClientWeekly")
    body = "".join([
        lead_block([
            "The raw case extract is larger than the true client-facing workload. Roughly 6,800 cases in the file are internal or system-generated entries. They matter for process understanding, but they distort the business story if they are mixed into client service metrics. The operating baseline in this document therefore uses the remaining 36,296 client cases.",
            "The weekly pattern tells a clear story. The team is not overwhelmed by sudden demand spikes. Weekly creation and weekly resolution stay in the same general range. The problem is that resolution falls slightly short of creation almost every week, so the queue never resets. What looks manageable in any single week becomes meaningful when the shortfall compounds over an entire quarter.",
            "The final weeks are the most important signal. The number of cases created in a given week that remain unresolved by extract end is rising. That acceleration matters more than the average because it suggests the backlog is not merely present. It is actively gathering momentum.",
        ]),
        evidence_table(
            "Population Split",
            "This table separates true client workload from internal or system-generated case volume.",
            pop,
            max_rows=30,
        ),
        evidence_table(
            "Weekly Client Case Trend",
            "This is the heartbeat of the operation over the 3-month window.",
            weekly,
            max_rows=20,
            notes=[
                "created = new client cases opened that week",
                "resolved = client cases closed that week",
                "backlog = cumulative created minus cumulative resolved",
                "still_open = cases created that week that remain unresolved as of extract end",
            ],
        ),
        so_what(
            "This is not a crisis of uncontrollable volume. It is a compounding shortfall in throughput. That is exactly the kind of problem where earlier classification, better summarization, and faster identification of cases heading into the tail can create real leverage."
        ),
        recommended_action([
            "Use queue reduction and unresolved-case aging as the primary success measures for any first-wave AI pilot.",
            "Treat client-only case volume as the default lens for all downstream analysis and discussion.",
        ]),
    ])
    takeaway = "The queue grows because creation stays just ahead of resolution, and that gap compounds rather than resetting."
    return takeaway, body


def render_subject_friction(ctx: Ctx) -> tuple[str, str]:
    deep = ctx.cases.get("D03_SubjectDeep")
    breach = ctx.cases.get("D06_SLA_Breach")
    body = "".join([
        lead_block([
            "The most important operational distinction in the data is not between large and small queues. It is between subjects that are consistently efficient, subjects that look fine on average but develop a dangerous tail, and subjects that are structurally slow because the underlying workflow takes days by nature.",
            "That distinction changes where AI makes sense. Fast and repetitive work is usually a routing or automation problem. Fat-tail work is where AI can intervene earlier and change outcomes. Structurally slow work is often a process-design issue first and an AI support issue second.",
            "The most valuable insight in this section is therefore not the subject count by itself. It is the combination of volume, median speed, tail behavior, and unresolved share. That combination separates interesting case types from true pilot candidates.",
        ]),
        evidence_table(
            "Subject-Level Friction Profile",
            "This table shows where volume, cycle time, and unresolved tails overlap.",
            deep,
            max_rows=18,
            notes=[
                "median_hours = typical resolution time for that subject",
                "p90_hours = time by which 90% of cases in the subject have resolved",
                "pct_unresolved = share of cases still open at extract end",
            ],
        ),
        evidence_table(
            "SLA Breach Profile",
            "These thresholds show how quickly the tail becomes meaningful for each subject.",
            breach,
            max_rows=20,
            notes=[
                "the threshold columns are diagnostic, not formal bank SLA definitions",
                "the point is to show where delay becomes persistent enough to matter",
            ],
        ),
        so_what(
            "The best AI opportunities sit in the middle tier: subjects where the median looks healthy but a meaningful minority of cases runs into long delays. Those are the workflows where earlier signal can reduce queue growth without pretending AI can redesign inherently slow processes."
        ),
        recommended_action([
            "Shortlist a small set of fat-tail subjects for the first pilot rather than treating the whole case mix as one problem.",
            "Use structurally slow subjects to identify process-redesign candidates, not just modeling candidates.",
        ]),
    ])
    takeaway = "Not all case types need the same intervention. The highest-return work sits in the fat tail, not in the already-efficient subjects."
    return takeaway, body


def render_workload_capacity(ctx: Ctx) -> tuple[str, str]:
    dow = ctx.cases.get("D04_DayOfWeek")
    hourly = ctx.cases.get("D05_HourlyPattern")
    owners = ctx.cases.get("D09_OwnerWorkload")
    origin = ctx.cases.get("D10_OriginXSubject")
    retouch = ctx.cases.get("D08_Retouch")
    body = "".join([
        lead_block([
            "The intake channel is not evenly distributed. The data shows that the email path is effectively the front door to the operation. That matters because any AI intervention that ignores email is, by definition, missing most of the real workload.",
            "The workload is concentrated in time as well. Mornings carry disproportionate case creation, and the weekly pattern is similarly concentrated into a few days rather than spread smoothly across the calendar. That makes queue-building predictable enough to target.",
            "Performance variation by pod is also too large to ignore. The best-performing pods show that much faster resolution is already possible within the organization. The point of AI here is not to invent a new standard from scratch. It is to help the slower parts of the system move closer to the best already visible in the data.",
        ]),
        evidence_table(
            "Origin by Subject",
            "This matrix shows which intake paths generate which operational categories.",
            origin,
            max_rows=18,
            notes=[
                "email dominates the human-entered intake path",
                "report volume often reflects system-generated or scheduled work rather than a client asking for help",
            ],
        ),
        evidence_table(
            "Day-of-Week Pattern",
            "This shows where the weekly workload actually lands.",
            dow,
            max_rows=10,
        ),
        evidence_table(
            "Hourly Pattern",
            "This shows where daily queue-building begins.",
            hourly,
            max_rows=24,
        ),
        evidence_table(
            "Owner and Pod Workload",
            "This table shows where workload distribution and performance diverge inside the operation.",
            owners,
            max_rows=20,
        ),
        evidence_table(
            "Retouch Signal",
            "This is a secondary check on whether the process is suffering from rework loops.",
            retouch,
            max_rows=15,
        ),
        so_what(
            "The operation does not appear to have a large rework problem. It has a throughput and distribution problem. That makes triage, early summarization, and morning-hour assistance more relevant than broad post-resolution quality interventions."
        ),
        recommended_action([
            "Design first-wave assistance around the email intake path and the morning queue-building window.",
            "Use pod variance to identify where assistance should land first rather than rolling out uniformly.",
        ]),
    ])
    takeaway = "The email channel and the morning intake window dominate the workload, and pod performance varies enough to show where help will matter."
    return takeaway, body


def render_data_quality(ctx: Ctx) -> tuple[str, str]:
    joins = ctx.internal.get("4_JoinScorecard")
    dates = ctx.internal.get("2_DateCoverage")
    completeness = ctx.entity.get("E10_Completeness")
    text_stats = ctx.internal.get("11_TextFieldStats")
    body = "".join([
        lead_block([
            "Before claiming AI readiness, the data has to clear a simpler test: can the core entities be connected well enough to support action? On that test, the answer is broadly yes. The strongest path in the data is Email to Case to PMC. That path is good enough to anchor analysis and intervention design.",
            "The dead end is HOA-level linkage. The HOA file adds useful context, but it is not the right operating anchor for the first wave of AI use cases. The right unit of action is the PMC relationship because that is where case volume, deposit concentration, and relationship-management signals converge.",
            "Text coverage also matters here. A field that is present but shallow is not the same as a field that contains meaningful natural language. The data does not support every language-heavy AI idea equally. It supports a narrower set of tasks where linkage and text density are both strong enough.",
        ]),
        evidence_table(
            "Join Reliability",
            "These join paths determine whether the analysis can move from description into action.",
            joins,
            max_rows=12,
            notes=[
                "match rate = populated keys that successfully link",
                "Email -> Case -> PMC is the core analytical path",
                "Case -> HOA is too weak to be the main operating join",
            ],
        ),
        evidence_table(
            "Text Field Coverage",
            "These fields determine whether the dataset can support language-based AI tasks.",
            text_stats,
            max_rows=20,
            notes=[
                "non-null % shows whether the field is populated at all",
                "median and p90 length show whether the field contains real language or short labels",
            ],
        ),
        evidence_table(
            "Entity Completeness",
            "This scorecard shows where PMC-side profile quality is strong enough to use and where it remains thin.",
            completeness,
            max_rows=20,
        ),
        evidence_table(
            "Date Coverage",
            "These ranges define the real time boundaries of the current analysis.",
            dates,
            max_rows=20,
        ),
        so_what(
            "The data is good enough for PMC-centered operational AI use cases today. It is not good enough to support a more granular HOA-level personalization story without additional cleanup and stronger entity linkage."
        ),
        recommended_action([
            "Treat PMC as the operating unit of analysis in pilot design, client prioritization, and impact measurement.",
            "Use weak HOA linkage and thin entity fields as follow-on data quality work, not as blockers for the first wave.",
        ]),
    ])
    takeaway = "The data is joinable enough for PMC-centered AI work now, but not clean enough to make HOA the core operating unit."
    return takeaway, body


def render_pmc_portfolio(ctx: Ctx) -> tuple[str, str]:
    dep = ctx.entity.get("E01_DepositConcentr")
    top = ctx.entity.get("E02_TopPMCs")
    rm = ctx.entity.get("E05_RM_Coverage")
    fric = ctx.entity.get("E03_FrictionValue")
    body = "".join([
        lead_block([
            "Operational friction does not carry equal business weight across the book. The deposit base is concentrated enough that service breakdown on a small number of large PMC relationships is a retention risk, not simply an efficiency issue.",
            "That concentration changes how backlog and unresolved work should be interpreted. The same number of unresolved cases means something very different on a large strategic relationship than it does on a small tail account. That is why this section combines deposits, operational strain, and relationship-management coverage rather than treating them separately.",
            "The other important signal here is RM recency. Relationship gaps and service friction compound one another. A large client with rising unresolved work and stale coverage should be treated as a priority even before any AI pilot is discussed.",
        ]),
        evidence_table(
            "Deposit Concentration",
            "This table shows how top-heavy the economic base is.",
            dep,
            max_rows=25,
        ),
        evidence_table(
            "Top PMCs",
            "This is the practical prioritization table: large relationships, their case burden, and the coverage around them.",
            top,
            max_rows=25,
        ),
        evidence_table(
            "RM Coverage Recency",
            "This table shows whether high-value relationships are being actively covered or left stale.",
            rm,
            max_rows=20,
        ),
        evidence_table(
            "Friction vs Value",
            "This table shows which relationships create disproportionate operational load relative to economic size.",
            fric,
            max_rows=20,
            notes=[
                "high friction on a high-value client is a retention and growth issue",
                "high friction on a low-value client is a cost and servicing economics issue",
            ],
        ),
        so_what(
            "This is where the operational story becomes a business case. AI effort should not chase the noisiest relationships in the abstract. It should focus where operational drag and commercial importance overlap."
        ),
        recommended_action([
            "Choose pilot PMCs from the high-value and high-friction segment rather than from the long tail.",
            "Use unresolved count, deposit size, and RM recency together when deciding where sponsor attention is warranted.",
        ]),
    ])
    takeaway = "Friction matters most where it lands on concentrated deposits and under-covered PMC relationships."
    return takeaway, body


def render_email_text(ctx: Ctx) -> tuple[str, str]:
    overview = ctx.cases.get("D11_EmailOverview")
    mix = ctx.cases.get("D12_EmailCaseSubjects")
    burden = ctx.cases.get("D13_EmailBurden")
    samples = ctx.cases.get("D14_EmailTextSamples")
    text_stats = ctx.internal.get("11_TextFieldStats")
    body = "".join([
        lead_block([
            "The central question in the email data is not whether messages exist. It is whether the content is rich enough, clean enough, and linked enough to support useful model behavior. The evidence points in a positive direction. The text is there, and the linkage to cases is strong.",
            "The practical blocker is preprocessing, not absence of data. Email bodies arrive wrapped in HTML and banner noise. That is an engineering problem, not a data-readiness problem. Once cleaned, the one-day sample shows enough natural language to support summarization and related assistance tasks.",
            "The one-day scope still matters. This section should be used to judge feasibility of text-based use cases, not to make long-run claims about communication trends.",
        ]),
        evidence_table(
            "Email Overview",
            "This table shows the size and composition of the one-day communication sample.",
            overview,
            max_rows=20,
        ),
        evidence_table(
            "Email by Case Subject",
            "This shows which operational categories create the most communication burden.",
            mix,
            max_rows=20,
        ),
        evidence_table(
            "Email Burden per Case",
            "This helps separate ordinary case traffic from cases that generate disproportionate communication churn.",
            burden,
            max_rows=20,
        ),
        evidence_table(
            "Text Field Coverage",
            "These stats quantify whether the case and email text fields contain enough substance to use.",
            text_stats,
            max_rows=20,
        ),
        evidence_table(
            "Stripped Text Samples",
            "These samples are here because text feasibility cannot be judged from summary counts alone.",
            samples,
            max_rows=22,
            trunc_chars=180,
        ),
        so_what(
            "The email text exists and is rich enough for summarization. The gating dependency is preprocessing. That makes summarization a more realistic near-term use case than more ambitious language tasks that depend on cleaner labels or much longer history."
        ),
        recommended_action([
            "Treat email HTML stripping and banner removal as the first engineering dependency for language-based use cases.",
            "Use case-linked email burden to decide which subjects are worth targeting first for summarization and response support.",
        ]),
    ])
    takeaway = "The text is present and rich enough to use. The main blocker is preprocessing, not data existence."
    return takeaway, body


def render_usecase_map(ctx: Ctx) -> tuple[str, str]:
    d15 = ctx.cases.get("D15_GenAI_Evidence")
    top20 = ctx.usecase.get("Top 20 v2")
    longlist = ctx.usecase.get("Expanded Longlist v2")
    body = "".join([
        lead_block([
            "The purpose of this section is not to prove that every plausible AI idea should move forward. It is to separate the use cases the current data can actually support from the use cases that are strategically interesting but still under-evidenced.",
            "The strongest current candidates are the ones that sit closest to the operational bottlenecks already shown in this document: triage and routing, summarization, and missing-information detection. Escalation prediction also stands out as high-value, but it needs more disciplined labeling to move beyond a directional finding.",
            "Other ideas may still matter strategically, but they should not be confused with data-ready opportunities. The discipline here is to keep first-wave ambition tied to current evidence rather than to the full longlist of good ideas.",
        ]),
        evidence_table(
            "Current Evidence by Use Case",
            "This is the core evidence ledger from the cases and email workbooks.",
            d15,
            max_rows=20,
        ),
        base.section("Evidence Tiers", base.bullets([
            "Strongly supported now: triage and routing, summarization, missing-information detection.",
            "Supported with more modeling or labels: escalation prediction, draft reply assistance, workflow copilot by subject.",
            "Not supported by these files alone: broader commercial mining, document-heavy onboarding copilots, pricing or treasury use cases.",
        ])),
        evidence_table(
            "Top 20 Use Cases",
            "If the use case workbook is available, this table aligns the evidence with the existing program taxonomy.",
            top20,
            max_rows=20,
        ),
        evidence_table(
            "Expanded Longlist Snapshot",
            "This longlist is supporting context only. It should not dilute focus on the small number of use cases already supported by current data.",
            longlist.head(20) if not longlist.empty else longlist,
            max_rows=20,
        ),
        so_what(
            "This is where discipline matters most. The first wave should focus on use cases with strong data support today. Everything else should be treated as shapeable future work, not as a commitment disguised as a roadmap."
        ),
        recommended_action([
            "Use this section to lock 3-5 near-term candidates rather than reopening the full longlist.",
            "Keep evidence-backed candidates separate from strategically important but currently under-evidenced ideas.",
        ]),
    ])
    takeaway = "Three use case families are clearly supported today; the rest should be treated as future-stage opportunities rather than near-term commitments."
    return takeaway, body


def render_next_steps(ctx: Ctx) -> tuple[str, str]:
    body = "".join([
        lead_block([
            "The right first move is not to start with a model. The right first move is to make the input surface reliable. That means cleaning the email payload, preserving case linkage, and making sure the fields that drive downstream classification remain usable run after run.",
            "Once that base is in place, the first modeling work should focus on triage and missing-information detection because those sit closest to the bottlenecks described earlier. In parallel, the business should decide which PMC relationships matter enough to use as pilot populations. That keeps the AI effort anchored to both throughput improvement and commercial relevance.",
        ]),
        base.section("Three moves to start now", base.bullets([
            "1. Build the email preprocessing layer: strip HTML, remove banners, preserve the case link, and create a model-ready body field.",
            "2. Build the first modeling path around case triage and missing-information detection, using Activity Subject and cleaned email content as the starting feature surface.",
            "3. Stand up a PMC relationship health view that combines deposits, unresolved case load, and RM recency so sponsors can see where operational pain and commercial exposure overlap.",
        ])),
        so_what(
            "These three steps split naturally across engineering, data science, and operating model work. They do not need to wait on one another, and together they create a credible first wave rather than an abstract AI strategy."
        ),
        recommended_action([
            "Treat preprocessing as an engineering dependency, not as a side task.",
            "Use the PMC-level commercial overlay to choose a pilot population before modeling begins.",
            "Measure first-wave success on queue reduction, earlier detection of risky cases, and better visibility into high-friction relationships.",
        ]),
    ])
    takeaway = "The first wave should start with preprocessing, then move into triage and missing-information detection, while commercial prioritization happens in parallel."
    return takeaway, body


def render_caveats() -> str:
    return caveats_footer([
        "Cases represent a 3-month operating window, not a full-year history.",
        "Emails represent a 1-day sample and should be used for text feasibility, not long-run trend claims.",
        "Internal or system-generated cases should be excluded from client-facing operating metrics.",
        "Case -> HOA linkage is weak. PMC is the right unit for operating analysis and pilot design.",
        "Some deposit and entity anomalies require cleanup before the artifact should be treated as executive-grade reporting.",
        "Use case feasibility labels in this artifact are evidence-based assessments, not project commitments.",
    ])


def main() -> None:
    ctx = Ctx()
    story_html = render_story(ctx)
    out = ctx.outdir / OUTPUT_HTML_NAME
    out.write_text(story_html, encoding="utf-8")
    (ctx.outdir / REDIRECT_HTML_NAME).write_text(
        f'<!doctype html><html><head><meta http-equiv="refresh" content="0; url={OUTPUT_HTML_NAME}"></head><body><a href="{OUTPUT_HTML_NAME}">Open story</a></body></html>',
        encoding="utf-8",
    )
    print(f"Narrative HTML story written to: {out}")


if __name__ == "__main__":
    main()
