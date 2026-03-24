# -*- coding: utf-8 -*-
"""
WAB HOA Operations - Data-Driven GenAI Opportunity Assessment
==============================================================
Standalone script that reads the three generated analysis workbooks
and builds a single-page HTML story for internal + leadership discussion.

Optional input: WAB_Ops_UseCases_2026-03-18.xlsx

Dependencies: pandas, openpyxl
"""

from __future__ import annotations

import html as _html
import re
from pathlib import Path

import pandas as pd
import numpy as np

# ---------------------------------------------------------------------
#  EDIT THESE PATHS BEFORE RUNNING
# ---------------------------------------------------------------------
INTERNAL_EXTRACT_XLSX = r"C:\Users\YourName\Desktop\wab_output\wab_internal_extract.xlsx"
CASES_DEEP_DIVE_XLSX  = r"C:\Users\YourName\Desktop\wab_output\wab_cases_deep_dive.xlsx"
ENTITY_DEEP_DIVE_XLSX = r"C:\Users\YourName\Desktop\wab_output\wab_entity_deep_dive.xlsx"
USECASE_XLSX          = r"C:\Users\YourName\Desktop\WAB_Ops_UseCases_2026-03-18.xlsx"  # optional
EMAIL_INSIGHTS_XLSX   = r"C:\Users\YourName\Desktop\wab_output\wab_email_deep_insights.xlsx"  # optional
OUTPUT_DIR            = r"C:\Users\YourName\Desktop\wab_html_story"

SITE_TITLE = "WAB HOA Operations - GenAI Opportunity Assessment"
RUN_LABEL  = "Phase 1 | Internal Discussion Draft | Data as of March 2026"


# ---------------------------------------------------------------------
#  UTILITIES  (kept from prior version - these are solid)
# ---------------------------------------------------------------------

def ensure_dir(p):
    Path(p).mkdir(parents=True, exist_ok=True)
    return Path(p)

def norm(s):
    if s is None: return ""
    return re.sub(r"\s+", " ", str(s).strip().lower())

def clean(s):
    if s is None: return ""
    return re.sub(r"\s+", " ", str(s).replace("\r"," ").replace("\n"," ")).strip()

def trunc(s, n=140):
    t = clean(s)
    return t if len(t) <= n else t[:n].rstrip() + "..."

def esc(s):
    return _html.escape(str(s)) if s else ""

def fmt(v):
    try:
        if pd.isna(v): return ""
    except Exception:
        pass
    if isinstance(v, bool): return str(v)
    if isinstance(v, int):  return f"{v:,}"
    if isinstance(v, float):
        if abs(v) >= 1000 and float(v).is_integer(): return f"{int(v):,}"
        return f"{v:,.1f}"
    return clean(v)

def find_col(df, *cands):
    lookup = {norm(c): c for c in df.columns}
    for cand in cands:
        n = norm(cand)
        if n in lookup: return lookup[n]
    for cand in cands:
        n = norm(cand)
        for k, v in lookup.items():
            if n in k or k in n: return v
    return None

def story_lookup(df, category, metric_name):
    if df.empty: return None
    c1, c2, c3 = find_col(df,"category"), find_col(df,"metric"), find_col(df,"value")
    if not (c1 and c2 and c3): return None
    mask = df[c1].astype(str).str.strip().str.lower().eq(category.lower()) & \
           df[c2].astype(str).str.strip().str.lower().eq(metric_name.lower())
    return df.loc[mask, c3].iloc[0] if mask.any() else None


class Book:
    def __init__(self, path):
        self.path = Path(path)
        self.exists = self.path.is_file()
        self._xls = None
        self.sheet_names = []
        if self.exists:
            try:
                self._xls = pd.ExcelFile(self.path, engine="openpyxl")
                self.sheet_names = list(self._xls.sheet_names)
            except Exception:
                self.exists = False

    def get(self, sheet):
        if not self.exists or self._xls is None or sheet not in self.sheet_names:
            return pd.DataFrame()
        try:
            return pd.read_excel(self._xls, sheet_name=sheet)
        except Exception:
            return pd.DataFrame()


class Ctx:
    def __init__(self):
        self.internal = Book(INTERNAL_EXTRACT_XLSX)
        self.cases    = Book(CASES_DEEP_DIVE_XLSX)
        self.entity   = Book(ENTITY_DEEP_DIVE_XLSX)
        self.usecase  = Book(USECASE_XLSX)
        self.insights = Book(EMAIL_INSIGHTS_XLSX)
        self.outdir   = ensure_dir(OUTPUT_DIR)


# ---------------------------------------------------------------------
#  HTML BUILDING BLOCKS
# ---------------------------------------------------------------------

def h(tag, text, cls="", **attrs):
    a = f' class="{cls}"' if cls else ""
    for k, v in attrs.items():
        a += f' {k.rstrip("_")}="{esc(v)}"'
    return f"<{tag}{a}>{text}</{tag}>"

def p(text):
    """Paragraph. Accepts raw HTML - caller is responsible for escaping user data."""
    return f"<p>{text}</p>"

def prose(*paragraphs):
    """Multiple paragraphs of narrative text. Each string is wrapped in <p>."""
    return "".join(f"<p>{t}</p>" for t in paragraphs if t)

def metric_card(label, value, note=""):
    display = fmt(value) if value is not None else "N/A"
    n = f'<div class="mn">{esc(note)}</div>' if note else ""
    return f'<div class="mc"><div class="ml">{esc(label)}</div><div class="mv">{esc(display)}</div>{n}</div>'

def metric_grid(cards):
    return '<div class="mg">' + "".join(cards) + "</div>"

def table(df, max_rows=20, trunc_len=140, note=""):
    if df is None or df.empty:
        return '<div class="empty">Data not available.</div>'
    d = df.head(max_rows).copy()
    for c in d.columns:
        d[c] = d[c].map(lambda x: trunc(fmt(x), trunc_len))
    out = ['<div class="tw"><table><thead><tr>']
    out.extend(f"<th>{esc(str(c))}</th>" for c in d.columns)
    out.append("</tr></thead><tbody>")
    for _, row in d.iterrows():
        out.append("<tr>")
        out.extend(f"<td>{esc(str(v))}</td>" for v in row.tolist())
        out.append("</tr>")
    out.append("</tbody></table></div>")
    if note:
        out.append(f'<div class="tn">{esc(note)}</div>')
    return "".join(out)

def so_what(text):
    return f'<div class="sw"><strong>What this means:</strong> {text}</div>'

def callout(title, text, style="info"):
    return f'<div class="co co-{style}"><strong>{esc(title)}</strong><br>{text}</div>'

def sub_section(title, body):
    return f'<div class="ss"><h3>{esc(title)}</h3>{body}</div>'

def bullets(items):
    return "<ul>" + "".join(f"<li>{esc(x)}</li>" for x in items if x) + "</ul>"


# ---------------------------------------------------------------------
#  SECTION RENDERERS
# ---------------------------------------------------------------------

def render_executive(ctx):
    e11  = ctx.entity.get("E11_StoryNumbers")
    d15  = ctx.cases.get("D15_GenAI_Evidence")
    join = ctx.internal.get("4_JoinScorecard")

    cards = []
    if not e11.empty:
        cards = [
            metric_card("Client Cases (3 months)", story_lookup(e11, "Cases (3mo)", "client_cases"), "Dec 2025 - Mar 2026"),
            metric_card("Currently Unresolved", story_lookup(e11, "Cases (3mo)", "client_unresolved"), "5% of client cases"),
            metric_card("Net Backlog Growth", "~130 / week", "Accelerating in recent weeks"),
            metric_card("Total Deposits", story_lookup(e11, "PMC Universe", "total_deposits"), "Across 1,658 PMCs (Deposits Rollup, includes HOA deposits)"),
            metric_card("Deposit Concentration", "Moderate", "Top 5 PMCs ~14%; see E01 for full distribution"),
            metric_card("Emails Analyzed (1 day)", story_lookup(e11, "Emails (1day)", "total_emails"), "100% linked to cases"),
        ]

    narrative = prose(
        "Over the past three months, the HOA operations team processed approximately 38,500 client-facing cases "
        "(36,296 with named companies plus ~2,200 blank-company cases confirmed as genuine client emails). "
        "That volume is manageable on a daily basis — most cases resolve in under five hours. "
        "But the team never fully clears the queue. Each week, roughly 130 more cases are created than resolved, "
        "and that gap is accelerating.",

        "This operations load sits on top of a moderately concentrated deposit book. The top five PMC "
        "relationships hold roughly 14% of total deposits, and the top twenty account for about a third. "
        "No single client dominates, but service friction on any of the larger relationships still "
        "carries retention risk given the cumulative deposit exposure.",

        "The good news: the underlying data — cases, emails, PMC records, and HOA linkages — is joinable "
        "at 82–99% match rates, and the email corpus contains enough natural language text to support "
        "AI-assisted triage, summarization, and escalation detection today. This document walks through "
        "the evidence behind these claims and concludes with a recommended set of near-term GenAI interventions "
        "ranked by data readiness.",
    )

    parts = [metric_grid(cards), narrative]

    if not d15.empty:
        parts.append(sub_section("GenAI Opportunity Snapshot",
            p("The table below summarizes what the current data directly supports for each candidate GenAI use case. "
              "The 'signal' column describes what the data tells us; the 'value' column quantifies it.") +
            table(d15, max_rows=25, note="Source: D15_GenAI_Evidence from Cases deep-dive workbook.")
        ))

    if not join.empty:
        parts.append(sub_section("Data Joinability",
            p("For any cross-file analysis to be trustworthy, the linkages between files must work. "
              "The table below shows match rates for the four main join paths. "
              "A match rate above 80% is generally usable; below 50% is a dead end.") +
            table(join, max_rows=8,
                  note="raw_exact_pct = percentage of left-side keys that find an exact match on the right side. "
                       "norm_exact_pct = same test after normalizing casing, whitespace, and punctuation.")
        ))

    return "".join(parts)


def render_case_volume(ctx):
    d01 = ctx.cases.get("D01_PopulationSplit")
    d02 = ctx.cases.get("D02_ClientWeekly")

    narrative = prose(
        "The raw case file contains 43,113 records, but not all of them represent client-facing work. "
        "Roughly 6,800 cases — about 16% — are internal system entries generated by automated processes "
        "such as AAB ADMIN batch jobs or blank-company placeholders. These inflate raw numbers and distort "
        "resolution-time averages if left in. Every metric in this document uses the remaining 36,296 "
        "client cases as the baseline.",

        "The weekly trend tells a clear story. The team creates roughly 2,700 to 3,400 client cases per week "
        "and resolves a similar number. On the surface, that looks manageable. But look at the gap: every single week, "
        "resolution falls slightly short of creation. The shortfall is small — around 130 cases per week — "
        "but it never reverses. Over 14 weeks, that compounds into a cumulative backlog of over 2,100 cases.",

        "More concerning, the pace is accelerating. In the final three weeks of the extract, the number of cases "
        "created in a given week that remain unresolved as of the extract date jumped from 228 to 382 to 455. "
        "That acceleration is the signal worth paying attention to.",
    )

    parts = [narrative]

    if not d01.empty:
        parts.append(sub_section("Client vs. Internal Population",
            p("This table shows the full population split. The 'ALL CASES' rows include everything; "
              "'CLIENT ONLY' excludes internal accounts; 'INTERNAL / SYSTEM' isolates the automated entries. "
              "All subsequent sections in this document use the client-only population.") +
            table(d01, max_rows=30)
        ))

    if not d02.empty:
        parts.append(sub_section("Weekly Client Case Trend",
            p("Each row represents one ISO week. 'created' is the number of new client cases opened that week. "
              "'resolved' is the number closed. 'backlog' is the running total of created minus resolved — "
              "a proxy for how many cases are accumulating over time. 'still_open' counts cases created in that "
              "specific week that have not yet been resolved as of the extract date.") +
            table(d02, max_rows=20)
        ))

    parts.append(so_what(
        "At the current trajectory, the unresolved backlog would reach approximately 5,500 cases within six months "
        "without intervention. Even a modest improvement in resolution speed — shaving one or two hours off the "
        "median — compounds against a growing queue. This is the core operational-leverage argument for AI-assisted "
        "case handling."
    ))

    return "".join(parts)


def render_subject_friction(ctx):
    d03 = ctx.cases.get("D03_SubjectDeep")
    d06 = ctx.cases.get("D06_SLA_Breach")
    d17 = ctx.cases.get("D17_BankerHoursBudget")
    d18 = ctx.cases.get("D18_ClaimCheck")
    d19 = ctx.cases.get("D19_ResearchBreakdown")
    d20 = ctx.cases.get("D20_KeyPersonRisk")
    d21 = ctx.cases.get("D21_MixedTypeFlags")

    # ── Population callout ──
    pop_note = callout("Population Note",
        "This analysis uses an inclusive client population: named-company cases plus blank-company cases. "
        "Validation confirmed that blank-company cases are genuine client emails where CRM did not assign "
        "a company name (99.8% email-originated, 100% banker/pod overlap with named clients, similar "
        "resolution profile). Only AAB ADMIN / WAB ADMIN cases are excluded as admin/multi-client.", "info")

    # ── Tier narrative ──
    narrative = prose(
        "Not all case types behave the same way, and this distinction matters for deciding where to intervene. "
        "The Subject field is a banker-selected pick list (214 distinct values) — a managed operational taxonomy. "
        "When we break cases down by subject, three distinct tiers emerge, each requiring a different intervention.",

        "The first tier is fast and clean. NSF and Non-Post cases resolve in a median of 1.5 hours "
        "with virtually no unresolved tail (0.2%). Fraud Alert and Transfer follow a similar pattern. These "
        "are already efficient — rules-based automation is the right tool, not AI.",

        "The second tier is where the GenAI opportunity concentrates. Research cases (4,073 in three months) "
        "resolve in a median of 3.8 hours, but the 90th percentile stretches to 107 hours. Account Maintenance "
        "shows a similar pattern: 3.6-hour median but a 317-hour P90. These 'fat-tail' subjects are where early "
        "identification of complexity saves the most hours.",

        "The third tier is structurally slow. Signature Card cases take a median of 155 hours — over six business "
        "days — and 49% exceed one full week. CD Maintenance sits at 98 hours, but stakeholder confirmation (Chris, "
        "March 23) clarified this is by design: bankers wait for CD maturity dates, not processing friction. "
        "These need process redesign, not speed optimization.",
    )

    tiers = (
        callout("AUTOMATE — Fast / Clean",
                "NSF and Non-Post (1.5h median), Fraud Alert (2.6h), Transfer (1.3h). "
                "Already efficient. Rules-based automation target.", "info") +
        callout("AI-ASSIST — Fat Tail (GenAI sweet spot)",
                "Research (3.8h median, P90 107h), New Account Request (25h, P90 190h), "
                "Account Maintenance (3.6h, P90 317h). AI triage, draft reply, and escalation detection.", "warn") +
        callout("REDESIGN — Structurally Slow",
                "Signature Card (155h median, 49% &gt; 1 week), Close Account (23h, P90 338h). "
                "Process redesign first. AI support second.", "err") +
        callout("MONITOR — By Design",
                "CD Maintenance (98h — maturity wait, not friction), IntraFi Maintenance (91h). "
                "Hours consumed are wait time, not work time. Exclude from automation ROI.", "info")
    )

    parts = [pop_note, narrative, tiers]

    # ── Banker-Hours Budget (D17) — the ROI table ──
    if d17 is not None and not d17.empty:
        parts.append(sub_section("Banker-Hours Budget by Subject",
            p("This table shows estimated banker-hours consumed per week by each major subject. "
              "'hours_type' distinguishes WORK TIME (recoverable through automation) from WAIT TIME "
              "(calendar wait, e.g. CD maturity). 'est_recoverable_hrs_per_week' applies a conservative "
              "60% factor to work-time subjects — the hours that AI or automation could realistically return.") +
            table(d17, max_rows=27)
        ))

    # ── Claim Check (D18) — Chris's assertions tested ──
    if d18 is not None and not d18.empty:
        parts.append(sub_section("Operational Context: Stakeholder Claims vs Data",
            p("During the March 23 working session, operational leadership provided context on several "
              "key subjects. The table below tests each claim against the data. Verdicts: CONFIRMED means "
              "the data strongly supports the claim; PARTIALLY CONFIRMED means directionally correct but "
              "with nuance.") +
            table(d18, max_rows=10)
        ))

    # ── Research Breakdown (D19) ──
    if d19 is not None and not d19.empty:
        parts.append(sub_section("Research: Not One Workflow, but Five",
            p("Research is the largest single case subject at 4,073 cases, and stakeholders confirmed it is "
              "a catch-all. Keyword classification of Description and Activity Subject text reveals five "
              "distinct sub-types. Payment Research dominates at roughly 40%. The large 'Other/Uncategorized' "
              "share reflects low Description fill (37%) — automated sub-routing will require email body text, "
              "not just CRM fields. This is itself a finding: the current taxonomy obscures workload composition.") +
            table(d19, max_rows=8)
        ))

    # ── Key-Person Risk (D20) ──
    if d20 is not None and not d20.empty:
        # Filter to only YES risks for the narrative
        spof_yes = d20[d20.get("risk_level", pd.Series(dtype=str)) == "YES"] if "risk_level" in d20.columns else d20
        n_spof = len(spof_yes) if not spof_yes.empty else 0
        parts.append(sub_section("Key-Person Risk",
            p(f"{n_spof} subjects have a single banker handling more than 50% of all cases. "
              "If any of these individuals were unavailable — vacation, attrition, reassignment — "
              "hundreds to thousands of cases would need redistribution with no established backup. "
              "This table shows all subjects where a single owner exceeds 35% of volume.") +
            table(d20, max_rows=20)
        ))

    # ── Mixed-Type Flags (D21) ──
    if d21 is not None and not d21.empty:
        parts.append(sub_section("Mixed-Type Subjects (Hidden Sub-Workflows)",
            p("These subjects show P90/median resolution time ratios above 5×, indicating that multiple "
              "distinct workflows are hiding under a single label. For example, a subject with a 3-hour median "
              "but a 300-hour P90 likely contains both quick-resolve and multi-day cases. "
              "These are candidates for subject taxonomy revision or sub-segmentation.") +
            table(d21, max_rows=15)
        ))

    # ── Original detail tables ──
    if d03 is not None and not d03.empty:
        parts.append(sub_section("Subject-Level Detail (Reference)",
            p("Top 15 case subjects ranked by volume with resolution profile, text fill rates, "
              "and unresolved percentage.") +
            table(d03, max_rows=18)
        ))

    if d06 is not None and not d06.empty:
        parts.append(sub_section("SLA Breach Profile (Reference)",
            p("Percentage of cases exceeding 24h, 72h, and 168h thresholds by subject.") +
            table(d06, max_rows=15)
        ))

    parts.append(so_what(
        "The banker-hours budget is the investment case. Signature Card alone consumes the most hours but is "
        "a process redesign target, not an AI target. The AI sweet spot is the fat-tail tier (Research, New Account, "
        "Account Maintenance) where early complexity detection saves the most recoverable hours. "
        "Key-person risk on 9+ subjects demands cross-training regardless of AI deployment. "
        "Research sub-segmentation shows the catch-all can be broken into 5 routable types — "
        "but requires email body text, not just CRM fields."
    ))

    return "".join(parts)


def render_workload(ctx):
    d04 = ctx.cases.get("D04_DayOfWeek")
    d05 = ctx.cases.get("D05_HourlyPattern")
    d09 = ctx.cases.get("D09_OwnerWorkload")
    d10 = ctx.cases.get("D10_OriginXSubject")
    d08 = ctx.cases.get("D08_Retouch")
    d16 = ctx.cases.get("D16_TriageDelay")

    narrative = prose(
        "Understanding when and how work arrives is essential for deciding where to deploy AI assistance.",

        "Eighty-four percent of client cases originate from email. The next largest channel — Report — "
        "accounts for 15%, and nearly all of that is automated NSF/Non-Post notifications generated by the system, "
        "not human requests. Phone, portal, and other channels combined are under 1%. This means the email channel "
        "is effectively the entire AI opportunity surface: any model that processes inbound email covers the vast "
        "majority of case creation.",

        "The work is concentrated in time as well. Nearly half of all cases (48%) are created between 8:00 and "
        "11:00 AM. Tuesday is the peak day at 24% of weekly volume; Saturday and Sunday are essentially zero. "
        "An AI triage system deployed during the 7:00–12:00 window would intercept over 70% of daily inflow "
        "at the point where queue-building begins.",

        "Pod-level performance reveals a wide variance that deserves investigation. The fastest pod resolves cases "
        "in a median of 1.8 hours. The slowest takes 20.7 hours — more than ten times longer. This could reflect "
        "differences in case mix, geographic complexity, or staffing capacity. Regardless of cause, the gap confirms "
        "that best-practice resolution patterns exist within the organization. The question is whether AI can help "
        "the slower pods close the gap.",

        "One reassuring signal: the retouch rate is effectively zero (3 cases out of 12,189 with tracking data). "
        "Cases that are resolved stay resolved. The team does not have a rework problem — it has a throughput "
        "and distribution problem.",
    )

    parts = [narrative]

    if not d10.empty:
        parts.append(sub_section("Origin × Subject",
            p("This cross-tab shows case counts and median resolution hours by intake channel and case subject. "
              "It reveals, for instance, that NSF cases arrive primarily through the Report channel (automated), "
              "while most other subjects are email-originated.") +
            table(d10, max_rows=15)
        ))

    if not d04.empty:
        parts.append(sub_section("Day-of-Week Pattern",
            p("'cases' = total client cases created on that day of the week across all 14 weeks. "
              "'avg_per_day' = cases divided by number of weeks, giving the typical daily volume.") +
            table(d04, max_rows=8)
        ))

    if not d05.empty:
        parts.append(sub_section("Hourly Pattern",
            p("Hour of day when client cases are created. The peak window (8–11 AM) accounts for 48% of daily volume.") +
            table(d05, max_rows=24)
        ))

    if not d09.empty:
        parts.append(sub_section("Pod and Owner Workload",
            p("This table shows case volume, median resolution hours, and unresolved counts by pod and by individual owner. "
              "Pods are the organizational units that handle case work; owners are individual case handlers. "
              "Wide variance in median_hrs may reflect case-mix differences rather than performance, "
              "but it identifies where to investigate.") +
            table(d09, max_rows=25)
        ))

    if not d16.empty:
        parts.append(sub_section("Email-to-Case Triage Delay",
            prose(
                "One of the most actionable metrics in this analysis comes from a simple timestamp comparison. "
                "The SLA Start field records when the originating email was received by the system. The Created On "
                "field records when the banker created the case. The gap between these two timestamps is the "
                "triage delay -- the time an email sits in the queue before a human reads it, decides it should "
                "become a case, and creates that case.",

                "Of the 36,296 client cases, 25,059 (69%) have a measurable triage gap -- meaning the email arrived "
                "before the banker created the case. The remaining 31% have zero or near-zero gap, indicating "
                "auto-created cases or non-email origins.",

                "The median triage delay is 58.6 minutes -- nearly an hour. One in four cases (P75) waits over "
                "3.4 hours. One in ten (P90) waits over 18 hours -- overnight. The single largest bucket in the "
                "distribution is the 1-2 hour range (3,694 cases), but there is also a significant overnight cluster: "
                "3,308 cases sit for 8-24 hours, meaning they arrived after hours and were not triaged until the "
                "next business morning.",

                "The subject-level breakdown reveals that Fraud Alert is the only category with fast triage "
                "(13-minute median) -- likely because it triggers immediate attention. Every other major subject "
                "sits for 42-79 minutes at the median, with Research (77 min), New Account Request (76 min), and "
                "Close Account (79 min) being the slowest. These are exactly the high-volume, fat-tail subjects "
                "identified earlier as the primary GenAI targets.",

                "Pod-level variance in triage speed mirrors the resolution-time variance seen earlier. The fastest pod "
                "triages in a median of 32 minutes; the slowest takes 82 minutes. Pods that start fast also resolve fast.",

                "The hour-of-day pattern is striking. Cases arriving at 7:00 AM have a 190-minute median triage delay -- "
                "over three hours -- because these are overnight emails waiting for bankers to log in. By 9:00 AM the "
                "median drops to 45 minutes as the team actively works the queue. After 5:00 PM, delays spike again.",

                "This metric converts the abstract claim 'AI could help with email triage' into a specific operational "
                "cost. Across 25,059 cases with triage gaps, the total time spent in the queue before case creation "
                "is approximately 24,500 banker-hours over three months -- roughly 40 FTE-hours per week. If AI triage "
                "reduced the median from 59 minutes to 5 minutes by automatically classifying emails and creating cases, "
                "nearly all of that time would be returned to case resolution. The overnight queue at 7:00 AM is the "
                "single highest-value deployment window: emails pre-classified before bankers log in means they start "
                "the day with an organized queue instead of a pile of unread emails.",
            ) +
            table(d16, max_rows=35,
                  note="'triage_minutes' = gap between SLA Start (email received) and Created On (case created). "
                       "Cases with zero or negative gap are auto-created or non-email and excluded from the distribution. "
                       "The BY SUBJECT and BY POD sections show where triage delays are longest. "
                       "BY HOUR OF DAY shows when emails queue up -- 7:00 AM has a 190-minute median delay (overnight backlog).")
        ))

    parts.append(so_what(
        "This is a weekday, morning-heavy, single-channel (email) operation. AI triage deployed in the "
        "7:00-12:00 window covers 70%+ of inflow. The 10x pod variance suggests best-practice patterns "
        "exist to learn from. The near-zero retouch rate confirms the problem is speed, not quality. "
        "The triage delay metric provides the direct baseline for measuring AI triage impact."
    ))

    return "".join(parts)


def render_data_quality(ctx):
    join   = ctx.internal.get("4_JoinScorecard")
    e10    = ctx.entity.get("E10_Completeness")
    stats  = ctx.internal.get("11_TextFieldStats")
    dates  = ctx.internal.get("2_DateCoverage")

    narrative = prose(
        "Before discussing what AI can do with this data, it is important to understand what the data actually "
        "looks like — what fields are populated, which joins are reliable, and where the gaps are.",

        "The most important question is whether the four files can be connected into a single analytical graph. "
        "The answer is yes, with one dead-end path. HOAs link to PMCs at 94% match rate. Emails link to Cases at 99%. "
        "Cases link to PMCs at 83% -- and that rises to roughly 93% when internal/system cases (which have no "
        "company name) are excluded. The one weak link is Cases to HOAs directly, at just 8%. Stakeholder "
        "confirmation explains why: cases are always tracked against the management company (PMC), not individual "
        "HOAs. This is by operating policy, not a data quality defect. The operational unit of analysis is the PMC. "
        "To reach HOA-level detail, the path goes Case to PMC to HOA, not Case to HOA directly.",

        "Two additional field definitions were confirmed in stakeholder discussions. The SLA Start timestamp "
        "records when the originating email was received by the system. The Created On timestamp records when "
        "the banker created the case. The gap between these two values represents the email-to-case conversion "
        "delay -- the time spent reading, deciding, and creating a case from an inbound email. This delay is a "
        "directly measurable triage opportunity for GenAI: any reduction in that gap accelerates case creation. "
        "Additionally, some HOAs are known to be orphaned (not linked to any PMC). The approximately 6% "
        "HOA-to-PMC join gap is a recognized data quality issue, not a systemic linkage failure.",

        "On the text side, the picture is mixed. Case Description fields are only 50% populated, with a median "
        "length of 28 characters when present — quite short. Activity Subject is stronger at 75% fill rate and "
        "54-character median — these contain forwarded email subject lines with entity names, transaction types, "
        "and dates. Email bodies are 100% populated and rich (median 17,767 characters raw, 2,177 after HTML stripping), "
        "but they require preprocessing to be usable.",

        "Entity data completeness is a constraint. Only 30% of PMCs have all key fields populated (name, deposits, "
        "state, HOAs, cases, RM, pod, recent check-in). The main bottleneck is RM check-in, which is only 42% filled. "
        "This does not block AI use cases on the case/email side, but it limits the PMC-level relationship analytics.",
    )

    parts = [narrative]

    if not join.empty:
        parts.append(sub_section("Join Match Rates",
            p("Each row represents an attempted join between two files. 'left_non_null' is how many records on the left "
              "side have a non-empty key. 'norm_exact_pct' is the match rate after normalizing casing and punctuation. "
              "The 'top_unmatched_keys' column shows the most common keys that failed to match — useful for diagnosing "
              "whether failures are systematic (e.g., internal accounts) or random.") +
            table(join, max_rows=8)
        ))

    if not stats.empty:
        parts.append(sub_section("Text Field Coverage",
            p("This table answers a gating question for GenAI: is there enough natural language in the system to support "
              "summarization, classification, or missing-information detection? 'non_null_pct' is the share of records "
              "where the field contains any text. 'median_chars' is the typical length of that text when present. "
              "'p90_chars' is the length at the 90th percentile — the longest 10% of entries.") +
            table(stats, max_rows=10)
        ))

    if not e10.empty:
        parts.append(sub_section("PMC Entity Completeness",
            p("This scorecard shows what percentage of the 1,658 PMC records have each key field populated. "
              "The 'ALL FIELDS COMPLETE' row at the bottom is the share that passes every check. "
              "At 30%, this means roughly 500 PMCs are fully analysis-ready.") +
            table(e10, max_rows=15)
        ))

    if not dates.empty:
        parts.append(sub_section("Date Coverage",
            p("This table shows the time span of every date-like column across all four files. "
              "'distinct_days' tells you how many unique calendar days appear. "
              "Cases span 85 distinct days over 14 weeks. Emails span 1 day (March 11, 2026).") +
            table(dates, max_rows=20)
        ))

    return "".join(parts)


def _build_quadrant_summary(e03):
    """Compute a 4-row quadrant summary from the E03 friction-value table."""
    qcol = find_col(e03, "quadrant")
    dep_col = find_col(e03, "deposits_fmt", "deposits")
    case_col = find_col(e03, "case_count")
    if not qcol:
        return pd.DataFrame()

    # Try to parse deposit values back to numeric for summing
    dep_numeric = None
    if dep_col:
        raw = e03[dep_col].astype(str).str.replace(r"[,$]", "", regex=True).str.replace("B","e9").str.replace("M","e6").str.replace("K","e3")
        dep_numeric = pd.to_numeric(raw, errors="coerce")

    rows = []
    quad_order = [
        ("High Value / Low Friction",  "Protect these. Well-served, high-value relationships."),
        ("High Value / High Friction", "Priority AI targets. Operational investment justified by economic exposure."),
        ("Low Value / Low Friction",   "Steady state. Monitor but do not over-invest."),
        ("Low Value / High Friction",  "Review relationship economics. Disproportionate cost to serve."),
    ]
    for q_name, q_interp in quad_order:
        mask = e03[qcol].astype(str).str.strip().eq(q_name)
        n = int(mask.sum())
        total_cases = int(e03.loc[mask, case_col].sum()) if case_col and mask.any() else 0
        total_dep = ""
        if dep_numeric is not None and mask.any():
            s = dep_numeric[mask].sum()
            if abs(s) >= 1e9: total_dep = f"${s/1e9:.1f}B"
            elif abs(s) >= 1e6: total_dep = f"${s/1e6:.0f}M"
            elif abs(s) >= 1e3: total_dep = f"${s/1e3:.0f}K"
            elif s > 0: total_dep = f"${s:.0f}"
        rows.append({
            "Quadrant": q_name,
            "PMC Count": n,
            "Total Deposits": total_dep,
            "Total Cases": total_cases,
            "Interpretation": q_interp,
        })
    return pd.DataFrame(rows)


def render_pmc_portfolio(ctx):
    e01 = ctx.entity.get("E01_DepositConcentr")
    e02 = ctx.entity.get("E02_TopPMCs")
    e05 = ctx.entity.get("E05_RM_Coverage")
    e03 = ctx.entity.get("E03_FrictionValue")
    e07 = ctx.entity.get("E07_HierarchyDepth")

    narrative = prose(
        "The operational friction described in previous sections does not carry equal business weight across all "
        "clients. The deposit book shows moderate concentration: the top five PMCs hold roughly 14% of "
        "total deposits, while the top twenty hold about a third. No single PMC dominates — the largest "
        "(The Management Trust) holds approximately 3.4%. Meanwhile, over 500 PMCs carry less than $1 million "
        "each, forming a long tail of small accounts that generate operational load without proportional "
        "economic return.",

        "Even at moderate concentration, friction still carries weight. A cluster of unresolved cases at a "
        "multi-hundred-million-dollar client is a different problem than the same volume at a sub-million "
        "account. The former is a retention risk; the latter is a cost question.",

        "Several findings from the data deserve specific attention. Among the top PMCs by deposit size, some show "
        "signs of operational strain combined with relationship-management gaps. The RM check-in field -- which records "
        "the last time a relationship manager contacted a PMC -- is only populated for 42% of PMCs. Of those with "
        "a recorded check-in, 407 (58%) have not been checked in over a year. Among high-deposit PMCs, some of the "
        "largest relationships show check-in gaps exceeding 1,000 days.",
    )

    parts = [narrative]

    # Deposit rollup caveat
    parts.append(callout("Important: How Deposits Are Counted",
        "The Deposits Rollup field at the PMC level is a consolidated figure that includes deposits from all "
        "underlying HOAs managed by that PMC. It updates weekly. PMC deposits and HOA deposits should not be "
        "summed independently -- doing so would double-count. All deposit figures in this document use the "
        "PMC-level rollup as the single source of truth. Additionally, the rollup includes IntraFi Cash Sweep "
        "(ICS) and CDARS balances, though these may not appear in individual sub-account detail.",
        "warn"))

    if not e01.empty:
        parts.append(sub_section("Deposit Distribution",
            p("This table shows the shape of the deposit book using the Deposits Rollup field — the confirmed "
              "single source of truth that consolidates all underlying HOA deposits weekly. "
              "'pmcs_with_deposits' is how many PMCs have a non-zero deposit value recorded. "
              "The percentile rows (p25, median, p75, p90) show the spread. "
              "The 'TOP N CONCENTRATION' rows show what share of total deposits the largest N PMCs hold.") +
            table(e01, max_rows=25)
        ))

    if not e02.empty:
        parts.append(sub_section("Top PMCs by Deposits",
            p("Each row is a named PMC ranked by deposit size. Key columns: 'hoa_count' = number of HOAs managed by "
              "this PMC. 'case_count' = total client cases in 3 months. 'unresolved' = cases still open. "
              "'median_hrs' = typical resolution time. 'top_subject' = the most common case type for this PMC. "
              "'pod' = the operational team handling this client. 'relationship_manager' = the assigned RM.") +
            table(e02, max_rows=25, trunc_len=100)
        ))

    if not e05.empty:
        parts.append(sub_section("RM Coverage and Recency",
            p("This table shows how recently relationship managers have contacted PMCs. "
              "'RECENCY BUCKETS' counts how many PMCs fall into each time-since-last-check-in range. "
              "'HIGH-DEPOSIT STALE' lists specific large-deposit PMCs that have not been checked in over 180 days.") +
            table(e05, max_rows=20)
        ))

    # Friction-Value: quadrant summary first, then detail
    if not e03.empty:
        quad_summary = _build_quadrant_summary(e03)

        parts.append(sub_section("Friction vs. Value -- Quadrant Overview",
            prose(
                "To understand the full picture of how operational friction relates to client value, every PMC with "
                "both deposits and case data is classified into one of four quadrants. The classification uses the "
                "median deposit amount and the median cases-per-million-dollars-on-deposit as the dividing lines. "
                "The table below summarizes each quadrant.",

                "The metric 'cases_per_1M_deposits' measures how many cases a PMC generates for every $1 million "
                "in deposits they hold. A higher number means more operational load per dollar of relationship value. "
                "This metric allows comparison across PMCs of very different sizes -- a $300M client with 600 cases "
                "and a $3M client with 6 cases would have the same ratio.",
            ) +
            table(quad_summary, max_rows=4, trunc_len=200,
                  note="Quadrant assignment uses the median deposit and median cases_per_1M_deposits as thresholds. "
                       "PMCs above the median on both dimensions are 'High Value / High Friction.'")
        ))

        parts.append(sub_section("Friction vs. Value -- Detail",
            p("This table is sorted by cases_per_1M_deposits (highest friction-to-value ratio first), which is why "
              "the visible rows are predominantly Low Value / High Friction. To identify the recommended AI pilot "
              "candidates, look for rows where quadrant = 'High Value / High Friction' -- these are high-deposit "
              "clients with disproportionate operational burden.") +
            table(e03, max_rows=30, trunc_len=100)
        ))

    if not e07.empty:
        parts.append(sub_section("Entity Hierarchy",
            p("This table shows how many HOAs each PMC manages, and which PMCs operate across multiple states. "
              "Hierarchy depth matters because operational complexity rises when a single PMC controls "
              "hundreds or thousands of HOAs across different geographies.") +
            table(e07, max_rows=25)
        ))

    parts.append(so_what(
        "The deposit book is moderately concentrated with a long tail of small accounts. Among the larger PMC "
        "relationships, several show signs of neglect -- stale RM check-ins, high unresolved case counts, or both. "
        "AI-powered relationship health scoring is a retention play, not just an efficiency play. "
        "The friction-value quadrant identifies which clients deserve proactive AI investment (High Value / High Friction) "
        "and which may warrant a relationship economics review (Low Value / High Friction)."
    ))

    return "".join(parts)


def render_email_text(ctx):
    d11   = ctx.cases.get("D11_EmailOverview")
    d12   = ctx.cases.get("D12_EmailCaseSubjects")
    d13   = ctx.cases.get("D13_EmailBurden")
    d14   = ctx.cases.get("D14_EmailTextSamples")
    stats = ctx.internal.get("11_TextFieldStats")

    narrative = prose(
        "The email file contains 2,423 records from a single day -- March 11, 2026. While this is too narrow "
        "for trend analysis, it is sufficient to assess whether the email corpus is rich enough to serve "
        "as input for GenAI applications like summarization, draft reply assistance, and sentiment detection.",

        "Every email in the sample links to a case (100% match rate). The split is 77% outbound "
        "(Sent or Completed status) and 23% inbound (Received). The high outbound ratio means that the team's "
        "own reply patterns are well-represented -- useful for training draft-reply models.",

        "Stakeholder confirmation adds an important detail about the Activity Subject field on case records. "
        "This field contains the original email subject line that triggered the case -- not a CRM-generated label. "
        "At 75% fill rate and a median of 54 characters, it preserves entity names, account numbers, transaction "
        "types, and forwarded-chain context in natural language. This makes Activity Subject the strongest "
        "candidate input for NLP-based triage, classification, or summarization on the case side of the data.",

        "Email bodies are stored as raw HTML with a median length of 17,767 characters. After stripping HTML tags, "
        "style blocks, and decoding entities, the usable text reduces to a median of 2,177 characters — roughly "
        "400 words. That is a 6.4× compression ratio. Importantly, 82% of stripped emails exceed 500 characters, "
        "which is generally sufficient for meaningful summarization.",

        "Every external email carries a standard security banner prefix ('ATTENTION: This email originated from "
        "outside of the WAB Network...'). This is systematic and can be removed with a single regex pattern "
        "as a preprocessing step. After banner removal and HTML stripping, the remaining content includes "
        "real conversational prose — questions, follow-ups, transaction references, and entity names.",
    )

    parts = [narrative]

    if not d11.empty:
        parts.append(sub_section("Email Overview",
            p("Summary statistics for the one-day email sample. 'linked_to_case' shows what percentage of emails "
              "connect to an existing case record. 'median_body_text_chars' is the typical email length after "
              "HTML stripping. 'html_to_text_ratio' shows how much of the raw content is markup versus actual text.") +
            table(d11, max_rows=25)
        ))

    if not d12.empty:
        parts.append(sub_section("Email Volume by Case Subject",
            p("Which case types generate the most email communication? 'emails_per_case' is the average number "
              "of emails per distinct case on the sampled day. 'median_body_chars' shows which subjects produce "
              "the longest email bodies — those are the richest targets for summarization.") +
            table(d12, max_rows=18)
        ))

    if not d13.empty:
        parts.append(sub_section("Email Burden per Case",
            p("How many emails does a typical case generate in one day? The distribution is light-tailed — "
              "most cases get 1–3 emails. The 'HIGHEST EMAIL CASES' rows identify the specific cases that "
              "generated the most communication.") +
            table(d13, max_rows=20)
        ))

    if not stats.empty:
        parts.append(sub_section("Text Field Coverage Across Files",
            p("This table compares text-field quality across both Cases and Emails. "
              "For GenAI feasibility, the key question is whether fields contain enough natural language "
              "to support classification, summarization, or missing-info detection.") +
            table(stats, max_rows=10)
        ))

    if not d14.empty:
        parts.append(sub_section("Sample Email Text (HTML Stripped)",
            p("These are actual email samples with HTML tags removed. They are included because GenAI feasibility "
              "cannot be judged from counts alone — the question is whether the cleaned text contains actionable "
              "prose or only metadata fragments. Rows are stratified: linked emails (connected to a case), "
              "unlinked emails, and the longest email bodies.") +
            table(d14, max_rows=20, trunc_len=180)
        ))

    parts.append(so_what(
        "The email text is there and it is rich enough. After HTML stripping and banner removal, 82% of emails "
        "have over 500 characters of real content — sufficient for summarization. The blocker is preprocessing "
        "(a standard engineering task), not data availability."
    ))

    return "".join(parts)


def render_email_intelligence(ctx):
    """Email deep-insights section — NLP topic modeling, content structure,
    template coverage, missing-info detection, conversation threads, and
    signal-by-subject dashboard.  Reads from the email_deep_insights workbook."""

    i01 = ctx.insights.get("I01_DataScope")
    i02 = ctx.insights.get("I02_ContentStructure")
    i03 = ctx.insights.get("I03_TopicDiscovery")
    i04 = ctx.insights.get("I04_MissingInfo")
    i05 = ctx.insights.get("I05_OutboundTemplates")
    i06 = ctx.insights.get("I06_ConversationThreads")
    i07 = ctx.insights.get("I07_SignalBySubject")
    i08 = ctx.insights.get("I08_TriageByIntent")

    narrative = prose(
        "The previous section established that the email corpus is rich enough to support GenAI applications. "
        "This section goes deeper — applying NLP topic modeling, template-coverage analysis, missing-information "
        "detection, and signal extraction to the 2,423 emails from the single-day sample. The goal is to move "
        "from 'the data exists' to 'here is exactly what an AI model would see and what it could do.'",

        "Several findings reshape the use-case picture. Outbound banker emails are far more repetitive than "
        "expected — 69-87% of responses within major subjects cluster into near-duplicate templates. This "
        "upgrades draft-reply assistance from MEDIUM to HIGH feasibility. Missing-information language appears "
        "in 15% of inbound emails, with TIN (Taxpayer Identification Number) dominating at 258 mentions — a "
        "clear automation target for onboarding nudges. And the median email is 71% quoted history, meaning "
        "summarization value lies in extracting the new content from the noise, not in condensing long prose.",
    )

    parts = [narrative]

    # --- I01: Data Scope ---
    if not i01.empty:
        parts.append(sub_section("Email Corpus Profile",
            prose(
                "The one-day sample contains 2,423 emails: 547 inbound (22.6%) and 1,857 outbound (76.6%). "
                "All 2,423 link to cases (100% match rate), and 2,212 (91.3%) connect to client cases "
                "specifically. After HTML stripping, the median clean body is 1,693 characters. The median "
                "new-content portion — text that is not quoted or forwarded — is 526 characters (29% of "
                "the total body). Over half of emails (53.7%) have more than 500 characters of genuinely "
                "new content, which is the threshold for meaningful NLP input.",

                "Signal detection across the full corpus found missing-information cues in 271 emails (11.2%), "
                "follow-up cues in 184 (7.6%), and urgency cues in 387 (16.0%). These are keyword-based "
                "detections — an LLM-based approach would likely surface more nuanced patterns.",
            ) +
            table(i01, max_rows=15, note="Source: I01_DataScope from email deep-insights workbook.")
        ))

    # --- I02: Content Structure by Subject ---
    if not i02.empty:
        # I02 has mixed layout: rows 2-11 are overall/inbound/outbound stats (metric rows with
        # percentile columns), rows 12+ are BY SUBJECT with different columns (subject, emails,
        # median_new_chars, median_new_ratio, median_body_chars).  Show them as separate tables.
        subj_col = find_col(i02, "subject")
        metric_col = find_col(i02, "metric")

        # Overall / Inbound / Outbound summary rows (percentile distribution)
        if metric_col:
            summary_rows = i02[i02[metric_col].notna() & (i02[metric_col].astype(str).str.strip() != "")].copy()
            # Keep only relevant columns (metric, p25, median, p75, mean) — drop subject-specific cols
            summary_cols = [c for c in summary_rows.columns
                           if find_col(pd.DataFrame(columns=[c]), "metric","p25","median","p75","mean") or c == metric_col]
            if summary_cols:
                summary_rows = summary_rows[summary_cols].dropna(how="all", axis=1)
        else:
            summary_rows = pd.DataFrame()

        # By-subject rows (have a subject value)
        if subj_col:
            by_subj = i02[i02[subj_col].notna() & (i02[subj_col].astype(str).str.strip() != "")].copy()
            # Keep only subject-specific columns, drop metric/percentile cols
            subj_cols = [c for c in by_subj.columns
                         if c == subj_col or find_col(pd.DataFrame(columns=[c]),
                            "emails","median_new_chars","median_new_ratio","median_body_chars")]
            if subj_cols:
                by_subj = by_subj[subj_cols].dropna(how="all", axis=1)
        else:
            by_subj = pd.DataFrame()

        parts.append(sub_section("New Content vs. Quoted History",
            prose(
                "Not all emails carry the same information density. The median email is 29% new content "
                "and 71% quoted or forwarded history. But this ratio varies dramatically by subject and "
                "direction.",

                "Inbound emails have a 15% median new-content ratio (230 new chars at P25, 400 at median). "
                "Outbound emails are higher at 36% (382 new chars at P25, 552 at median) — bankers write "
                "more original text per message than clients do, because clients are replying into threads.",

                "By subject, the variation is stark. Statements emails are 100% new content but only 205 "
                "characters long — these are short, self-contained messages. Fraud Alert emails have a 30% "
                "new-content ratio but 1,589 median body characters — long threads where the new part is "
                "buried. NSF and Non-Post emails carry 22% new content with 1,638 median body chars. "
                "New Account Request has only 13% new content but 2,123 median body chars — the longest "
                "emails with the most quoted noise. This is where summarization adds the most value: "
                "extracting the 13% that is new from a 2,000-character thread.",
            ) +
            (table(summary_rows, max_rows=12,
                   note="Content structure distribution. Rows prefixed '--- OVERALL/INBOUND/OUTBOUND ---' "
                        "show percentile distributions of new-content and quoted-content character counts.")
             if not summary_rows.empty else "") +
            (table(by_subj, max_rows=15,
                   note="By-subject breakdown: emails = inbound client emails for that subject. "
                        "median_new_chars = typical new content length. median_new_ratio = share that is new vs quoted.")
             if not by_subj.empty else "")
        ))

    # --- I03: Topic Discovery (NMF) ---
    if not i03.empty:
        parts.append(sub_section("Hidden Intent Families (NMF Topic Modeling)",
            prose(
                "Within each case subject, emails are not monolithic. Non-negative Matrix Factorization (NMF) "
                "topic modeling reveals distinct intent families hiding inside broad subject categories.",

                "Account Maintenance, for example, splits into three topics: a dominant banking-operations "
                "cluster (56% of emails, terms: bank, alliance, banker, loan), an encrypted-message/Proofpoint "
                "cluster (34%, terms: intended, email, message, information), and a small CD-specific cluster "
                "(10%, terms: ana, cd, reynolds, cd maturity). The Proofpoint cluster is pure noise — these "
                "are encrypted-message read receipts, not actionable client communication.",

                "This pattern repeats across subjects. Research splits into a general banking topic (72.9%) "
                "and a location-specific topic (20.8%, terms: office, email, dallas, rose city). NSF and "
                "Non-Post splits into a banker-response topic (68.8%), a transaction-detail topic (16.7%), "
                "and a Proofpoint-noise topic (14.6%). The Proofpoint noise is a consistent finding: 58.6% "
                "of CD Maintenance emails and 52.4% of Statements emails are encrypted-message receipts "
                "that an AI preprocessing pipeline should filter before any classification or summarization.",
            ) +
            table(i03, max_rows=25,
                  note="NMF topics per case subject. 'pct_of_subject' = share of emails in that subject "
                       "assigned to this topic. 'top_terms' = the most distinctive words for this topic cluster.")
        ))

    # --- I04: Missing-Info Detection ---
    if not i04.empty:
        # Split into by-subject, top-doc-terms, and examples sections — drop the section column
        sec_col = find_col(i04, "section")
        if sec_col:
            by_subj_mi = i04[i04[sec_col].astype(str).str.strip().eq("BY SUBJECT")].drop(columns=[sec_col]).dropna(how="all", axis=1).copy()
            top_docs = i04[i04[sec_col].astype(str).str.strip().eq("TOP DOC TERMS")].drop(columns=[sec_col]).dropna(how="all", axis=1).copy()
            examples = i04[i04[sec_col].astype(str).str.strip().eq("EXAMPLES")].drop(columns=[sec_col]).dropna(how="all", axis=1).copy()
        else:
            by_subj_mi, top_docs, examples = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        parts.append(sub_section("Missing-Information Detection",
            prose(
                "A keyword-based scan of inbound client emails identifies messages that reference missing "
                "documents, pending items, or incomplete information. Across 547 inbound emails, 271 (11.2% "
                "of all emails, higher among inbound) contain at least one missing-information cue.",

                "The document terms that appear most frequently tell a clear story. TIN (Taxpayer "
                "Identification Number) dominates with 258 mentions — nearly half of all document references. "
                "EIN follows at 61, then CD maturity (23), signature card (13), certificate of deposit (7), "
                "W9 (4), management agreement (4), articles of incorporation (2), and CDARS agreement (2). "
                "The concentration on TIN aligns with onboarding workflows where new accounts cannot be "
                "fully opened without a taxpayer ID.",

                "By subject, NSF and Non-Post leads with a 25% missing-cue rate (12 of 48 inbound emails). "
                "New Account Request shows 2.7% — low in percentage but important because these emails have "
                "the highest doc_terms_mentioned count (31 total) and the longest case hours (median 29.2h). "
                "Research emails carry a 16.7% missing-info rate with 31 doc terms and 25.8-hour median "
                "case duration. The combination of missing-info signals and long resolution times in these "
                "subjects makes them prime candidates for automated 'missing document' nudge replies.",
            ) +
            (table(by_subj_mi, max_rows=15,
                   note="Missing-info cue rates by case subject. 'missing_cue_count' = inbound emails with at least one cue. "
                        "'doc_terms_mentioned' = total document-term references across all emails in that subject.")
             if not by_subj_mi.empty else "") +
            (sub_section("Top Document Terms Referenced",
                table(top_docs, max_rows=12,
                      note="Frequency of specific document terms in inbound emails. TIN at 258 mentions is the dominant signal."))
             if not top_docs.empty else "") +
            (sub_section("Example Emails with Missing-Info Cues",
                p("These are real inbound emails where missing-information language was detected. "
                  "The 'doc_terms' column shows which document types were mentioned. "
                  "The 'sample_new_content' column shows the extracted new-content portion of the email.") +
                table(examples, max_rows=10, trunc_len=200))
             if not examples.empty else "")
        ))

    # --- I05: Outbound Template Coverage ---
    if not i05.empty:
        parts.append(sub_section("Outbound Response Templates",
            prose(
                "One of the most impactful findings from the email deep-dive is the degree of repetition "
                "in outbound banker responses. Using TF-IDF cosine similarity clustering, outbound emails "
                "within each case subject were grouped into near-duplicate response families.",

                "The results are striking. New Account Request outbound emails show 85.9% template coverage — "
                "meaning 85.9% of the 198 outbound emails in that subject cluster into one of 5 recognizable "
                "response patterns. CD Maintenance reaches 90.3% (134 outbound, 5 clusters). Account "
                "Maintenance hits 82.4% (170 outbound). NSF and Non-Post is at 92.6% (108 outbound). "
                "Even the lowest-coverage subjects exceed 69%.",

                "This changes the draft-reply feasibility assessment fundamentally. If nearly 9 out of 10 "
                "responses in a subject follow one of a handful of templates, an AI system does not need "
                "to generate creative prose — it needs to select the right template and fill in the "
                "case-specific details (names, account numbers, dates). This is a retrieval-augmented "
                "generation (RAG) pattern, not a freeform generation problem, and it is significantly "
                "easier to build, validate, and deploy.",
            ) +
            table(i05, max_rows=40, trunc_len=200,
                  note="Each row is one cluster within a subject. 'cluster_size' = emails in this pattern. "
                       "'template_coverage_pct' = share of outbound emails in the subject covered by the top-5 clusters. "
                       "'sample_text' = representative email from the cluster (truncated).")
        ))

    # --- I06: Conversation Threads ---
    if not i06.empty:
        sec_col = find_col(i06, "section")
        if sec_col:
            thread_dist = i06[i06[sec_col].astype(str).str.strip().eq("THREAD DISTRIBUTION")].drop(columns=[sec_col]).dropna(how="all", axis=1).copy()
            by_subj_th = i06[i06[sec_col].astype(str).str.strip().eq("BY SUBJECT")].drop(columns=[sec_col]).dropna(how="all", axis=1).copy()
            heaviest = i06[i06[sec_col].astype(str).str.strip().eq("HEAVIEST THREADS")].drop(columns=[sec_col]).dropna(how="all", axis=1).copy()
        else:
            thread_dist, by_subj_th, heaviest = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        parts.append(sub_section("Conversation Thread Depth",
            prose(
                "Email threads reveal case complexity. Of 917 cases with email activity on the sampled day, "
                "the median case had 2.0 emails and the P90 had 4.0. The maximum was 16 emails on a single "
                "case. 88 cases (9.6%) had 5 or more emails — these are the 'heavy threads' where a banker "
                "spends the most time re-reading quoted history to find the latest update.",

                "By subject, New Account Request generates the most threads (130 cases, median 2 emails, "
                "P90 of 4) with 8.5% having missing-info cues. Research cases show 25.0% with follow-up "
                "cues and median case hours of 19.9. CD Maintenance has only 4.5% follow-up cues but a "
                "staggering 125.4 median case hours, driven by structurally slow workflows.",

                "The heaviest individual threads tell the story of where summarization would save the most "
                "time. The top thread (NSF and Non-Post, 16 emails, 3 inbound, 13 outbound) spans multiple "
                "actors and 12 hours. A banker picking up this case would need to read through the full "
                "chain to understand the current state — exactly the task an AI summarizer should handle.",
            ) +
            (table(thread_dist, max_rows=6,
                   note="Thread distribution summary: how many cases have emails, median/P90 thread depth, and heavy-thread count.")
             if not thread_dist.empty else "") +
            (table(by_subj_th, max_rows=18,
                   note="Thread metrics by case subject. 'median_emails' = typical email count per case. "
                        "'pct_with_missing_cue' and 'pct_with_followup' = share of cases in that subject with detected signals.")
             if not by_subj_th.empty else "") +
            (table(heaviest, max_rows=15,
                   note="Heaviest individual threads on the sampled day. 'email_count' = total emails on that case. "
                        "'actors' = distinct senders/recipients. 'case_hours' = total case resolution time.")
             if not heaviest.empty else "")
        ))

    # --- I07: Signal Dashboard by Subject ---
    if not i07.empty:
        parts.append(sub_section("Signal Dashboard by Subject",
            prose(
                "This dashboard consolidates all email intelligence signals — missing-info, follow-up, urgency — "
                "alongside case resolution metrics for each major subject. It is the decision-support view for "
                "prioritizing which subjects to target first with AI.",

                "NSF and Non-Post stands out: 25.0% missing-info rate, 37.5% urgency rate, and fast resolution "
                "(1.4h median, 2.1% unresolved). The high signal rates combined with fast resolution suggest "
                "that bankers already recognize and act on these signals quickly — AI could automate the "
                "recognition step. CD Maintenance shows a different pattern: 44.8% unresolved, 138.2h median "
                "case hours, but only 6.9% missing-info and 27.6% urgency. The problem here is structural "
                "slowness, not missing signals.",

                "The most actionable finding: subjects with high missing-info rates AND long resolution times "
                "are the prime AI targets. Research (16.7% missing-info, 18.8% follow-up, 25.8h median, "
                "8.3% unresolved) and Close Account (15.8% missing-info, 5.3% follow-up, 26.2h median, "
                "0.0% unresolved) fit this profile. New Account Request (2.7% missing-info but 29.2h median, "
                "8.0% unresolved) benefits more from draft-reply than from missing-info detection.",
            ) +
            table(i07, max_rows=15,
                  note="Combined signal dashboard. 'missing_info_pct' = share of inbound emails with missing-info cues. "
                       "'urgency_pct' = share with urgency language. 'pct_unresolved' = cases still open at extract date.")
        ))

    # --- I08: Triage by Intent ---
    if not i08.empty:
        sec_col = find_col(i08, "section")
        if sec_col:
            by_signal = i08[i08[sec_col].astype(str).str.strip().eq("TRIAGE BY SIGNAL")].drop(columns=[sec_col]).dropna(how="all", axis=1).copy()
            by_length = i08[i08[sec_col].astype(str).str.strip().eq("TRIAGE BY CONTENT LENGTH")].drop(columns=[sec_col]).dropna(how="all", axis=1).copy()
        else:
            by_signal, by_length = i08, pd.DataFrame()

        parts.append(sub_section("Do Email Signals Predict Triage Speed?",
            prose(
                "A natural question is whether emails with detectable signals (missing-info, follow-up, urgency) "
                "get triaged faster or slower than emails without them. The answer is counter-intuitive.",

                "Emails flagged with missing-info cues have a median triage time of 41.5 minutes — lower than "
                "the 55.1 minutes for unflagged emails. Follow-up cues show the opposite: 63.3 minutes flagged "
                "vs 52.2 minutes unflagged. Urgency cues are faster: 45.4 minutes flagged vs 54.1 unflagged.",

                "The interpretation: bankers appear to naturally triage signal-rich emails faster (missing-info "
                "and urgency cues catch their eye), while follow-up requests — which often sit in longer "
                "threads — take more time to parse. An AI system that pre-classifies emails by signal type "
                "would formalize and accelerate the pattern bankers are already following intuitively.",

                "Content length also matters. Short emails (<200 characters of new content) have a 60.3-minute "
                "median triage time, likely because they lack context and force the banker to investigate. "
                "Mid-length emails (200-500 chars) triage fastest at 48.9 minutes. The longest emails (1500+ "
                "chars) take 53.5 minutes — more text to read. An AI summarizer that reduces long emails to "
                "their essential content would flatten this curve.",
            ) +
            (table(by_signal, max_rows=5,
                   note="Triage time by signal type. 'flagged_median_triage_min' = median triage delay for emails "
                        "WITH the signal. 'unflagged_median_triage_min' = median for emails WITHOUT it.")
             if not by_signal.empty else "") +
            (table(by_length, max_rows=5,
                   note="Triage time by new-content length bucket. Shorter emails take longer to triage (less context); "
                        "mid-length emails (200-500 chars) are triaged fastest.")
             if not by_length.empty else "")
        ))

    parts.append(so_what(
        "The email deep-dive upgrades three use cases. Draft Reply Assistance moves from MEDIUM to HIGH "
        "feasibility — 69-87% of outbound emails follow repeatable templates, making this a retrieval problem, "
        "not a generation problem. Missing-Info Detection is confirmed with TIN as the dominant signal (258 "
        "mentions) — onboarding nudges are the clear first target. Email Summarization is reframed: the value "
        "is not condensing long prose but extracting the 29% of new content from 71% quoted noise, especially "
        "in heavy threads and subjects like New Account Request (13% new-content ratio, 2,123 chars). "
        "Proofpoint encrypted-message receipts (52-67% of some subjects) should be filtered in preprocessing."
    ))

    return "".join(parts)


def render_geo_rm_platform(ctx):
    e04 = ctx.entity.get("E04_StateProfile")
    e08 = ctx.entity.get("E08_PlatformMix")
    e09 = ctx.entity.get("E09_PodGeography")
    e06 = ctx.entity.get("E06_CompanyType")

    narrative = prose(
        "Geography, accounting platform, and organizational structure provide context for where AI interventions "
        "should land first and whether a single solution pattern can work across the portfolio.",

        "The top three states by PMC count — California (306), Florida (189), and Texas (159) — also lead in "
        "deposits and case volume. These three states alone account for a significant share of the operational "
        "footprint, which means any AI pilot deployed in CA, FL, or TX would cover the largest concentration of "
        "PMCs, deposits, and cases simultaneously.",

        "Platform mix reveals an important heterogeneity. Vantaca is the dominant accounting platform (239 PMCs, "
        "$4.1B deposits, 10,643 cases), followed by VMS (120 PMCs, $1.6B, 3,716 cases) and Caliber (137 PMCs, "
        "$1.5B, 3,611 cases). Different platforms may produce different case patterns, data structures, and "
        "integration requirements. A single AI copilot design may need platform-specific adaptations.",

        "Pod geography confirms that service teams are regionally aligned. WEST01 and WEST02 serve Arizona and "
        "California respectively; EAST03 covers Florida; Central01 spans Michigan, Ohio, and Texas. The pods "
        "with wider geographic spread appear in some cases to have longer median resolution times, though "
        "this pattern may also reflect case-mix or staffing differences rather than geography alone.",
    )

    parts = [narrative]

    if not e04.empty:
        parts.append(sub_section("State Profile",
            p("Internal operating footprint by geography. 'pmc_count' = management companies in that state. "
              "'pmc_deposits' = total deposits. 'hoa_count' = HOAs. 'case_count' = client cases in 3 months.") +
            table(e04, max_rows=20)
        ))

    if not e08.empty:
        parts.append(sub_section("Accounting Platform Mix",
            p("PMC distribution by accounting platform. 'case_count' shows operational volume by platform. "
              "Platforms with high case counts and high deposits are the ones where AI solutions would have "
              "the largest operational surface.") +
            table(e08, max_rows=18)
        ))

    if not e06.empty:
        parts.append(sub_section("Company Type Profile",
            p("PMC distribution by Company Type. Unlike NAICS (which is unreliable in this dataset), "
              "Company Type is a clean field that distinguishes Management Companies from Associations "
              "and other entity types. 85% of PMCs and 97% of cases come from Management Companies.") +
            table(e06, max_rows=10)
        ))

    if not e09.empty:
        parts.append(sub_section("Pod × State Matrix",
            p("This cross-tab shows which pods serve which states, inferred from PMC addresses. "
              "Pods with concentration in a single state tend to resolve cases faster than those spread "
              "across multiple geographies.") +
            table(e09, max_rows=15, trunc_len=60)
        ))

    return "".join(parts)


def render_usecase_map(ctx):
    d15 = ctx.cases.get("D15_GenAI_Evidence")
    i09 = ctx.insights.get("I09_EvidenceScorecard")
    top20 = ctx.usecase.get("Top 20 v2")
    longlist = ctx.usecase.get("Expanded Longlist v2")

    narrative = prose(
        "This section maps the data evidence from previous sections to specific GenAI use case opportunities. "
        "The goal is not to prove every use case, but to sort which ones are strongly supported by the current "
        "data, which are partially supported, and which remain outside the current evidence boundary.",

        "Based on the four files analyzed, the following use cases have the clearest data support today:",
    )

    evidence_prose = (
        callout("Triage and Routing — HIGH feasibility",
                "214 distinct case subjects exist, with 91% of cases originating from email. The top three subjects "
                "account for 29% of volume. For high-volume subjects, simple rules may be sufficient. For the long tail, "
                "a classifier trained on Activity Subject text (75% fill rate, median 54 characters) could route cases "
                "to the right pod and priority level. This is a high-value, moderate-complexity use case.", "info") +
        callout("Email Summarization — HIGH feasibility",
                "82% of emails have over 500 characters of usable text after HTML stripping. The median stripped body "
                "is 2,177 characters (roughly 400 words). Case subjects like IntraFi Maintenance (3,575 median chars) "
                "and CD Maintenance (3,059 chars) produce the richest email bodies. A summarizer that reduces a 400-word "
                "email to a 2-sentence brief would save meaningful read time across 2,500+ daily emails.", "info") +
        callout("Missing-Information Detection — LOW-HANGING FRUIT",
                "545 currently unresolved cases have no Description text at all. This is the simplest possible AI "
                "intervention: a rule or model that flags 'this case has been open X hours with no description — "
                "please add context.' No NLP required for the base version; a simple field-completeness check suffices. "
                "This is the recommended first deployment.", "info") +
        callout("Escalation Prediction — HIGH value, needs labels",
                "The global P90 resolution time is 237.6 hours. 3,412 cases exceeded this threshold in three months. "
                "1,809 cases are currently unresolved, 540 of them over 30 days old. An escalation model that flags cases "
                "drifting toward the P90 boundary — even one business day earlier — would prevent backlog accumulation. "
                "The gap: we do not currently have labeled escalation outcomes (was this case actually escalated?). "
                "Building that label set is the prerequisite.", "warn") +
        callout("Draft Reply Assistance — HIGH feasibility (upgraded)",
                "Email deep-dive reveals 69-87% of outbound banker responses within major subjects cluster into "
                "near-duplicate templates. New Account Request: 85.9% template coverage (198 outbound, 5 clusters). "
                "CD Maintenance: 90.3% (134 outbound). NSF and Non-Post: 92.6% (108 outbound). Account Maintenance: "
                "82.4% (170 outbound). This means an AI system needs to select the right template and fill in "
                "case-specific details (RAG pattern), not generate creative prose — significantly easier to "
                "build, validate, and deploy.", "info") +
        callout("Workflow Copilot — MEDIUM, subject-specific",
                "Among the top-15 subjects by volume, Signature Card (155h median), CD Maintenance (98h), and "
                "IntraFi Maintenance (91h) are structurally slow. A handful of low-volume niche subjects (ePay, "
                "In House ACH migrations) have even longer medians exceeding 1,000 hours, but their case counts are "
                "too small to justify dedicated AI investment. A copilot for the high-volume subjects (Research, "
                "New Account Request, Account Maintenance) that provides step-by-step guidance based on historical "
                "resolution patterns is more practical than attempting to accelerate structurally slow workflows.", "warn")
    )

    parts = [narrative, evidence_prose]

    if not i09.empty:
        parts.append(sub_section("Email Deep-Dive Evidence Scorecard",
            p("This scorecard synthesizes the email deep-insights analysis into a use-case-by-use-case verdict. "
              "Each row evaluates one GenAI use case against the email data, with a metric, threshold, and "
              "pass/fail verdict. The 'implication' column describes what the verdict means for implementation.") +
            table(i09, max_rows=10, trunc_len=200,
                  note="Verdicts: PASS = data clearly supports the use case. MARGINAL = partially supported. "
                       "CHECK TABLE / CHECK THREAD SHEET = evidence exists but requires deeper examination. "
                       "WEAK = current data does not support the use case as hypothesized.")
        ))

    if not d15.empty:
        parts.append(sub_section("Data Evidence Summary",
            p("This table consolidates the evidence signals for each use case category. "
              "The 'signal' column describes what the observation means for GenAI feasibility.") +
            table(d15, max_rows=25)
        ))

    if not top20.empty:
        parts.append(sub_section("Previously Defined Top 20 Use Cases",
            p("This table comes from the stakeholder session use-case workbook. It represents the use cases "
              "identified through interviews and workshops, ranked by strategic value and feasibility. "
              "The data evidence in this document either confirms, partially supports, or does not yet address "
              "each of these use cases.") +
            table(top20, max_rows=20, trunc_len=120)
        ))

    if not longlist.empty:
        parts.append(sub_section("Expanded Use Case Longlist",
            p("The full canonical list of 56 identified use cases with their AI type, 90-day feasibility, "
              "and strategic value ratings. This is included as reference — the near-term focus should remain "
              "on the 3–5 use cases with the strongest current data support.") +
            table(longlist.head(25), max_rows=25, trunc_len=100)
        ))

    # Stakeholder-informed analytical opportunities
    parts.append(sub_section("Stakeholder-Informed Opportunities",
        prose(
            "Conversations with the operations team surfaced additional context that shapes the GenAI "
            "opportunity beyond what the data files alone reveal.",

            "A pre-case email queue exists in the CRM where bankers read incoming emails and decide whether "
            "to convert them into cases. This queue sits upstream of everything analyzed in this document. "
            "Access to this queue data would enable measurement of: the email-to-case conversion rate, the "
            "decision latency (how long an email sits before becoming a case), and the volume of emails that "
            "do not become cases at all. This is the natural upstream point for AI-assisted triage -- an AI "
            "system that reads incoming emails, determines whether each should become a case, creates the case "
            "if appropriate, and drafts an initial reply for the banker to review and action.",

            "This workflow -- classify, create, draft, review -- was described by the operations team as the "
            "ideal target state. It aligns directly with the triage and draft reply use case combination "
            "identified in the data evidence above.",

            "One important scoping constraint: fully automated responses are not currently acceptable for most "
            "case types. However, specific scenarios were identified as candidates for automated nudge-style "
            "replies -- for example, when a new account onboarding package is missing a required document, an "
            "auto-generated response requesting the missing item would be acceptable. This shapes the pilot "
            "design: start with human-in-the-loop (AI drafts, banker reviews and sends), with missing-document "
            "auto-nudge as the first exception where full automation is permitted.",

            "The SLA Start and Created On timestamps in the case data provide a way to measure the triage "
            "delay today -- the gap between when an email arrives and when the banker creates a case from it. "
            "Computing this gap across the 36,296 client cases would establish the baseline that any AI triage "
            "system would need to improve against.",
        )
    ))

    parts.append(so_what(
        "Four use cases now have strong data support: triage/routing, email summarization, "
        "missing-information detection, and draft reply assistance (upgraded to HIGH after template "
        "coverage analysis showed 69-87% of outbound emails follow repeatable patterns). "
        "Escalation prediction remains high-value but needs labeled outcomes. "
        "The recommended starting point is missing-info detection -- it requires no NLP, delivers immediate "
        "visibility, and builds trust in the AI-assisted workflow. The pre-case email queue is the next "
        "data asset to pursue -- it would unlock the full classify-create-draft-review workflow."
    ))

    return "".join(parts)


def render_next_steps(ctx):
    narrative = prose(
        "Based on the evidence presented in this document, three near-term actions are recommended. "
        "They are sequenced by data readiness, not by ambition. All three can begin in parallel.",
    )

    steps = (
        callout("1. Deploy email preprocessing pipeline",
                "Strip HTML tags (6.4× bloat reduction), remove the standard security banner, and link "
                "cleaned text back to case records. This is an engineering task, not a data science task. "
                "It unblocks all email-based GenAI use cases — summarization, draft reply, and sentiment. "
                "Estimated effort: one engineering sprint.", "info") +
        callout("2. Implement missing-information flagging",
                "545 unresolved cases currently have no Description text. A simple automated check — "
                "'this case has been open X hours with no description' — would immediately improve case "
                "completeness and create a feedback loop. No model training required for the first version. "
                "This is the lowest-risk, fastest-to-deploy GenAI-adjacent intervention.", "info") +
        callout("3. Build escalation prediction for Tier 2 subjects",
                "Research (4,073 cases, P90 107h), New Account Request (3,472 cases, P90 190h), and "
                "Account Maintenance (3,018 cases, P90 317h) are the three subjects where a classifier "
                "trained on Activity Subject text could identify cases drifting toward the tail. "
                "Prerequisite: a labeled training set from the operations team indicating which cases "
                "were actually escalated. Estimated effort: 2–3 weeks for labeling, 1–2 weeks for model.", "warn")
    )

    additional = prose(
        "In parallel, several data and infrastructure actions would strengthen the foundation for future use cases:",
    )

    data_actions = bullets([
        "Request access to the pre-case email queue data. This queue -- where bankers read emails and decide whether to create cases -- is the upstream triage decision point. Access would enable measurement of conversion rates, decision latency, and the full classify-create-draft-review AI workflow.",
        "Compute the SLA Start to Created On gap across all email-originated client cases. This gap measures the current triage delay and establishes the baseline that any AI triage system must improve against.",
        "Request a longer email extract (1-2 weeks minimum) to validate email-based use cases beyond a single-day sample.",
        "Investigate the Arizona negative-deposit data quality issue before presenting deposit figures to leadership.",
        "Lift PMC entity completeness from 30% to 60%+ by prioritizing RM check-in field population and deposit data cleanup.",
    ])

    return narrative + steps + additional + data_actions


def render_caveats(ctx):
    return (
        '<details class="caveats"><summary>Data Scope, Caveats, and Known Limitations</summary>'
        '<div class="caveats-body">' +
        prose(
            "This analysis is based on four Excel files extracted from WAB's D365 CRM environment. "
            "It represents a point-in-time view, not a continuously updated data pipeline.",
        ) +
        bullets([
            "Cases: 3-month operating window (December 18, 2025 - March 19, 2026). Not a full annual history. Seasonal patterns cannot be assessed.",
            "Emails: 1-day sample only (March 11, 2026). Communication patterns observed may not generalize across weeks or months.",
            "PMCs and HOAs: Current-state snapshots, not longitudinal entity history. Changes over time cannot be tracked.",
            "All deposit figures use Deposits Rollup — a consolidated weekly figure that includes deposits from all underlying HOAs plus ICS/CDARS balances. PMC and HOA deposits must not be summed separately. Blank-name PMC rows (aggregation records) are excluded from all analysis.",
            "Arizona shows negative total deposits (-$53.8M). This is a data quality issue, not an economic signal.",
            "HOA Company Type is 99.1% null -- this field is unusable for analysis.",
            "NAICS codes are mono-valued in both files (73% = 531311 in PMCs, 87% = 813990 in HOAs). NAICS provides no segmentation power and should be treated as a data-quality finding, not an analytical variable.",
            "Cases are tracked against PMCs by operating policy, not HOAs. The Case to HOA join rate of 7.6% reflects this design, not a data quality defect. The valid analytical path is Case to PMC to HOA.",
            "Approximately 6% of HOAs are known to be orphaned (not linked to any PMC). This is a recognized data quality issue confirmed by stakeholders.",
            "Internal/system cases (15.8% of total) are excluded from all client-facing metrics. They are identified by company names AAB ADMIN, WAB ADMIN, and blank entries.",
            "Use case feasibility ratings are evidence-based assessments from the available data, not implementation commitments.",
        ]) +
        '</div></details>'
    )


# ---------------------------------------------------------------------
#  SECTION REGISTRY & PAGE ASSEMBLY
# ---------------------------------------------------------------------

SECTIONS = [
    ("exec_summary",     "Executive Summary",                     render_executive),
    ("case_volume",      "Case Volume and Backlog Trajectory",    render_case_volume),
    ("subject_friction", "Where Cases Get Stuck",                 render_subject_friction),
    ("workload",         "Channels, Workload, and Capacity",      render_workload),
    ("data_quality",     "Data Quality and Joinability",          render_data_quality),
    ("pmc_portfolio",    "Client Portfolio, Deposits, and Risk",  render_pmc_portfolio),
    ("email_text",       "Email and Text Feasibility",            render_email_text),
    ("email_intel",      "Email Intelligence Deep-Dive",          render_email_intelligence),
    ("geo_platform",     "Geography, Platforms, and Structure",   render_geo_rm_platform),
    ("usecase_map",      "GenAI Use Case Evidence Map",           render_usecase_map),
    ("next_steps",       "Recommended Next Steps",                render_next_steps),
]

CSS = """
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:0;background:#f8fafc;color:#1e293b;line-height:1.65;font-size:15px}
.hdr{background:#0f172a;color:#fff;border-bottom:4px solid #2563eb}
.hdr-in{max-width:1200px;margin:0 auto;padding:20px 32px}
h1{margin:0;font-size:28px;font-weight:700}
.sub{margin-top:4px;font-size:13px;opacity:.8}
.nav{position:sticky;top:0;z-index:10;background:#fff;border-bottom:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,.05)}
.nav-in{max-width:1200px;margin:0 auto;padding:10px 32px;display:flex;gap:6px;overflow-x:auto}
.nav a{flex:0 0 auto;text-decoration:none;color:#334155;border:1px solid #cbd5e1;border-radius:6px;padding:6px 14px;font-size:13px;white-space:nowrap;transition:all .15s}
.nav a:hover{background:#eff6ff;border-color:#93c5fd;color:#1e40af}
.main{max-width:1200px;margin:0 auto;padding:28px 32px 60px}
.sec{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:28px 32px;margin-bottom:24px;scroll-margin-top:64px}
.sec h2{margin:0 0 16px;font-size:24px;font-weight:700;color:#0f172a;border-bottom:2px solid #e2e8f0;padding-bottom:12px}
.ss{margin:20px 0}
.ss h3{font-size:17px;font-weight:600;color:#1e293b;margin:0 0 8px;padding-top:12px;border-top:1px solid #f1f5f9}
p{margin:0 0 12px;color:#334155}
.mg{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin:20px 0}
.mc{background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:16px}
.ml{font-size:12px;color:#64748b;margin-bottom:6px;text-transform:uppercase;letter-spacing:.5px}
.mv{font-size:26px;font-weight:700;color:#0f172a}
.mn{font-size:11px;color:#94a3b8;margin-top:4px}
.tw{overflow-x:auto;margin:12px 0}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{border:1px solid #e2e8f0;padding:7px 10px;text-align:left;vertical-align:top}
th{background:#f8fafc;font-weight:600;color:#475569;font-size:12px}
td{color:#334155}
.tn{font-size:12px;color:#64748b;font-style:italic;margin-top:6px;padding:4px 0}
.sw{background:#f0fdf4;border-left:4px solid #22c55e;padding:14px 18px;margin:18px 0;border-radius:0 8px 8px 0;font-size:14px;color:#166534}
.co{padding:14px 18px;margin:10px 0;border-radius:8px;font-size:14px;line-height:1.55}
.co-info{background:#eff6ff;border:1px solid #bfdbfe;color:#1e3a5f}
.co-warn{background:#fffbeb;border:1px solid #fde68a;color:#78350f}
.co-err{background:#fef2f2;border:1px solid #fecaca;color:#7f1d1d}
ul{margin:8px 0 12px 20px;padding:0}
li{margin:5px 0;color:#334155}
.empty{color:#94a3b8;font-style:italic;padding:8px 0}
.caveats{margin-top:32px;background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:4px 28px}
.caveats summary{padding:16px 0;font-weight:600;color:#475569;cursor:pointer;font-size:15px}
.caveats-body{padding:0 0 20px}
"""


def build_page(ctx):
    nav_html = '<div class="nav"><div class="nav-in">' + "".join(
        f'<a href="#s-{slug}">{esc(label)}</a>' for slug, label, _ in SECTIONS
    ) + "</div></div>"

    sections_html = ""
    for slug, title, renderer in SECTIONS:
        body = renderer(ctx)
        sections_html += (
            f'<section id="s-{esc(slug)}" class="sec">'
            f'<h2>{esc(title)}</h2>'
            f'{body}'
            f'</section>'
        )

    caveats_html = render_caveats(ctx)

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{esc(SITE_TITLE)}</title>
<style>{CSS}</style>
</head>
<body>
<div class="hdr"><div class="hdr-in"><h1>{esc(SITE_TITLE)}</h1><div class="sub">{esc(RUN_LABEL)}</div></div></div>
{nav_html}
<main class="main">
{sections_html}
{caveats_html}
</main>
</body>
</html>"""


def main():
    ctx = Ctx()
    html_content = build_page(ctx)
    out = ctx.outdir / "story.html"
    out.write_text(html_content, encoding="utf-8")
    # redirect page
    (ctx.outdir / "index.html").write_text(
        '<!doctype html><html><head><meta http-equiv="refresh" content="0;url=story.html"></head>'
        '<body><a href="story.html">Open story</a></body></html>',
        encoding="utf-8"
    )
    print(f"HTML story written to: {out}")
    print(f"Sections rendered: {len(SECTIONS)}")
    for slug, title, _ in SECTIONS:
        print(f"  [{slug}] {title}")


if __name__ == "__main__":
    main()
