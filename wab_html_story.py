"""
WAB HOA Operations — Data-Driven GenAI Opportunity Assessment
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

# ┌─────────────────────────────────────────────────────────┐
# │  EDIT THESE PATHS BEFORE RUNNING                        │
# └─────────────────────────────────────────────────────────┘
INTERNAL_EXTRACT_XLSX = r"C:\Users\YourName\Desktop\wab_output\wab_internal_extract.xlsx"
CASES_DEEP_DIVE_XLSX  = r"C:\Users\YourName\Desktop\wab_output\wab_cases_deep_dive.xlsx"
ENTITY_DEEP_DIVE_XLSX = r"C:\Users\YourName\Desktop\wab_output\wab_entity_deep_dive.xlsx"
USECASE_XLSX          = r"C:\Users\YourName\Desktop\WAB_Ops_UseCases_2026-03-18.xlsx"  # optional
OUTPUT_DIR            = r"C:\Users\YourName\Desktop\wab_html_story"

SITE_TITLE = "WAB HOA Operations — GenAI Opportunity Assessment"
RUN_LABEL  = "Phase 1 · Internal Discussion Draft · Data as of March 2026"


# ═══════════════════════════════════════════════════════════
#  UTILITIES  (kept from prior version — these are solid)
# ═══════════════════════════════════════════════════════════

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
        self.outdir   = ensure_dir(OUTPUT_DIR)


# ═══════════════════════════════════════════════════════════
#  HTML BUILDING BLOCKS
# ═══════════════════════════════════════════════════════════

def h(tag, text, cls="", **attrs):
    a = f' class="{cls}"' if cls else ""
    for k, v in attrs.items():
        a += f' {k.rstrip("_")}="{esc(v)}"'
    return f"<{tag}{a}>{text}</{tag}>"

def p(text):
    """Paragraph. Accepts raw HTML — caller is responsible for escaping user data."""
    return f"<p>{text}</p>"

def prose(*paragraphs):
    """Multiple paragraphs of narrative text. Each string is wrapped in <p>."""
    return "".join(f"<p>{t}</p>" for t in paragraphs if t)

def metric_card(label, value, note=""):
    n = f'<div class="mn">{esc(note)}</div>' if note else ""
    return f'<div class="mc"><div class="ml">{esc(label)}</div><div class="mv">{esc(fmt(value))}</div>{n}</div>'

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


# ═══════════════════════════════════════════════════════════
#  SECTION RENDERERS
# ═══════════════════════════════════════════════════════════

def render_executive(ctx):
    e11  = ctx.entity.get("E11_StoryNumbers")
    d15  = ctx.cases.get("D15_GenAI_Evidence")
    join = ctx.internal.get("4_JoinScorecard")

    cards = []
    if not e11.empty:
        cards = [
            metric_card("Client Cases (3 months)", story_lookup(e11, "Cases (3mo)", "client_cases"), "Dec 2025 – Mar 2026"),
            metric_card("Currently Unresolved", story_lookup(e11, "Cases (3mo)", "client_unresolved"), "5% of client cases"),
            metric_card("Net Backlog Growth", "~130 / week", "Accelerating in recent weeks"),
            metric_card("Total Deposits", story_lookup(e11, "PMC Universe", "total_deposits"), "Across 1,658 PMCs"),
            metric_card("Deposit Concentration", "Top 5 = 56%", "Top 50 PMCs hold 77%"),
            metric_card("Emails Analyzed (1 day)", story_lookup(e11, "Emails (1day)", "total_emails"), "100% linked to cases"),
        ]

    narrative = prose(
        "Over the past three months, the HOA operations team processed 36,296 client-facing cases. "
        "That volume is manageable on a daily basis — most cases resolve in under five hours. "
        "But the team never fully clears the queue. Each week, roughly 130 more cases are created than resolved, "
        "and that gap is accelerating. Today, 1,809 cases remain unresolved, 540 of them older than 30 days.",

        "This operations load sits on top of a $24 billion deposit book where the top five PMC relationships "
        "alone account for 56% of total deposits. The concentration means that service friction on even one "
        "major client is a retention risk, not just an efficiency problem.",

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

    narrative = prose(
        "Not all case types behave the same way, and this distinction matters for deciding where AI intervention "
        "would have the highest return. When we break the 36,296 client cases down by subject — the operational "
        "category assigned to each case — three distinct tiers emerge.",

        "The first tier is fast and clean. NSF and Non-Post cases, for example, resolve in a median of 1.5 hours "
        "with virtually no unresolved tail (0.2%). Fraud Alert and Transfer follow a similar pattern. These "
        "categories are already well-machined; AI adds marginal value here beyond possible straight-through automation.",

        "The second tier is where the GenAI opportunity concentrates. Research cases (4,073 in three months) "
        "resolve in a median of 3.8 hours — which looks fast — but the 90th percentile stretches to 107 hours "
        "and the maximum reaches 7,629 hours. Account Maintenance shows a similar pattern: 3.6-hour median "
        "but a 317-hour P90. These are 'fat-tail' subjects where most cases resolve quickly, but a meaningful "
        "minority drags for days or weeks. An AI model that identifies which cases are about to enter the tail — "
        "even a few hours earlier than a human would notice — could redirect attention before the backlog compounds.",

        "The third tier is structurally slow. Signature Card cases take a median of 155 hours — over six business "
        "days — and 49% exceed one full week. CD Maintenance and IntraFi Maintenance follow at 98 and 91 hours "
        "respectively. These are not speed problems that AI can solve. They are process design problems where "
        "the underlying workflow takes days by nature. AI can support these categories (missing-info detection, "
        "workflow nudging), but the primary intervention is process redesign.",
    )

    tiers = (
        callout("Tier 1 — Fast / Clean",
                "NSF and Non-Post (1.5h median, 0.2% unresolved), Fraud Alert (2.6h), Transfer (1.3h), "
                "Statements (1.5h). Already efficient. Automation target, not AI target.", "info") +
        callout("Tier 2 — Fat Tail (GenAI sweet spot)",
                "Research (3.8h median, P90 107h), New Account Request (25h median, P90 190h), "
                "Account Maintenance (3.6h median, P90 317h). Fast on average, but a significant minority "
                "of cases drag for days or weeks. Escalation detection and triage add the most value here.", "warn") +
        callout("Tier 3 — Structurally Slow",
                "Signature Card (155h median, 49% exceed 1 week), CD Maintenance (98h, 10.2% unresolved), "
                "IntraFi Maintenance (91h, 8.8%). These need process redesign first, AI support second.", "err")
    )

    parts = [narrative, tiers]

    if not d03.empty:
        parts.append(sub_section("Subject-Level Detail",
            p("This table shows the top 15 case subjects ranked by volume. Key columns: "
              "'median_hrs' is the time by which half of cases are resolved. "
              "'p90_hrs' is the time by which 90% are resolved — the remaining 10% take longer than this. "
              "'pct_unresolved' is the share of cases still open as of the extract date. "
              "'desc_fill_pct' and 'act_subj_fill_pct' show what percentage of cases in that subject "
              "have text in the Description and Activity Subject fields — critical for GenAI feasibility.") +
            table(d03, max_rows=18)
        ))

    if not d06.empty:
        parts.append(sub_section("SLA Breach Profile",
            p("This table shows what percentage of cases in each subject exceed various time thresholds. "
              "For example, '>24h_pct' is the share of cases that took more than 24 hours to resolve. "
              "Subjects where a large fraction exceeds 72 or 168 hours (one week) are the ones with "
              "structural process delays, not just occasional outliers.") +
            table(d06, max_rows=15)
        ))

    parts.append(so_what(
        "AI should target the fat-tail tier (Tier 2), where early identification of complexity saves the most hours. "
        "Tier 1 is already fast — rules and automation are the right tool. "
        "Tier 3 needs process redesign; deploying AI against a fundamentally slow workflow "
        "would optimize the wrong thing."
    ))

    return "".join(parts)


def render_workload(ctx):
    d04 = ctx.cases.get("D04_DayOfWeek")
    d05 = ctx.cases.get("D05_HourlyPattern")
    d09 = ctx.cases.get("D09_OwnerWorkload")
    d10 = ctx.cases.get("D10_OriginXSubject")
    d08 = ctx.cases.get("D08_Retouch")

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

    parts.append(so_what(
        "This is a weekday, morning-heavy, single-channel (email) operation. AI triage deployed in the "
        "7:00–12:00 window covers 70%+ of inflow. The 10× pod variance suggests best-practice patterns "
        "exist to learn from. The near-zero retouch rate confirms the problem is speed, not quality."
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
        "Cases link to PMCs at 83% — and that rises to roughly 93% when internal/system cases (which have no "
        "company name) are excluded. The one weak link is Cases to HOAs directly, at just 8%. This means the "
        "operational unit of analysis should be the PMC (the management company), not the individual HOA. "
        "To reach HOA-level detail, the path goes Case → PMC → HOA, not Case → HOA.",

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


def render_pmc_portfolio(ctx):
    e01 = ctx.entity.get("E01_DepositConcentr")
    e02 = ctx.entity.get("E02_TopPMCs")
    e05 = ctx.entity.get("E05_RM_Coverage")
    e03 = ctx.entity.get("E03_FrictionValue")
    e07 = ctx.entity.get("E07_HierarchyDepth")

    narrative = prose(
        "The operational friction described in previous sections does not carry equal business weight across all "
        "clients. The HOA deposit book totals approximately $24 billion across 1,658 PMCs, but that value is "
        "radically concentrated. The top five PMCs alone hold 56% of total deposits. The top fifty hold 77%. "
        "Meanwhile, 550 PMCs carry less than $1 million each — a long tail of small accounts that generate "
        "operational load without proportional economic return.",

        "This concentration changes how friction should be interpreted. A 126-case unresolved backlog at a "
        "$300-million client is not the same operational problem as 126 unresolved cases at a $50,000 client. "
        "The former is a retention risk; the latter is a cost question.",

        "Several findings from the data deserve specific attention. Among the top PMCs by deposit size, some show "
        "signs of operational strain combined with relationship-management gaps. The RM check-in field — which records "
        "the last time a relationship manager contacted a PMC — is only populated for 42% of PMCs. Of those with "
        "a recorded check-in, 407 (58%) have not been checked in over a year. Among high-deposit PMCs, some of the "
        "largest relationships show check-in gaps exceeding 1,000 days.",

        "The friction-versus-value analysis identifies which PMCs sit in the 'High Value / High Friction' quadrant — "
        "large deposit relationships generating disproportionate case volume per dollar on deposit. These are "
        "the clearest candidates for AI-assisted intervention because the operational investment is justified by "
        "the economic exposure. Conversely, 'Low Value / High Friction' PMCs — small-deposit clients generating "
        "disproportionate case volume — may warrant a different conversation about relationship economics.",
    )

    parts = [narrative]

    if not e01.empty:
        parts.append(sub_section("Deposit Distribution",
            p("This table shows the shape of the deposit book. 'pmcs_with_deposits' is how many PMCs have a non-zero "
              "deposit value recorded. The percentile rows (p25, median, p75, p90) show the spread. "
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

    if not e03.empty:
        parts.append(sub_section("Friction vs. Value",
            p("This table ranks PMCs by 'cases_per_1M_deposits' — a measure of how much operational load a client "
              "generates relative to the deposits they hold. Higher numbers mean more friction per dollar of value. "
              "The 'quadrant' column classifies each PMC as High or Low on both value and friction dimensions.") +
            table(e03, max_rows=25, trunc_len=100)
        ))

    if not e07.empty:
        parts.append(sub_section("Entity Hierarchy",
            p("This table shows how many HOAs each PMC manages, and which PMCs operate across multiple states. "
              "Hierarchy depth matters because operational complexity rises when a single PMC controls "
              "hundreds or thousands of HOAs across different geographies.") +
            table(e07, max_rows=25)
        ))

    parts.append(so_what(
        "The deposit book is a concentration risk. A small number of PMC relationships carry disproportionate "
        "economic value, and several of the largest show signs of neglect — stale RM check-ins, high unresolved "
        "case counts, or both. AI-powered relationship health scoring is a retention play, not just an efficiency play."
    ))

    return "".join(parts)


def render_email_text(ctx):
    d11   = ctx.cases.get("D11_EmailOverview")
    d12   = ctx.cases.get("D12_EmailCaseSubjects")
    d13   = ctx.cases.get("D13_EmailBurden")
    d14   = ctx.cases.get("D14_EmailTextSamples")
    stats = ctx.internal.get("11_TextFieldStats")

    narrative = prose(
        "The email file contains 2,423 records from a single day — March 11, 2026. While this is too narrow "
        "for trend analysis, it is sufficient to assess whether the email corpus is rich enough to serve "
        "as input for GenAI applications like summarization, draft reply assistance, and sentiment detection.",

        "Every email in the sample links to a case (100% match rate). The split is 77% outbound "
        "(Sent or Completed status) and 23% inbound (Received). The high outbound ratio means that the team's "
        "own reply patterns are well-represented — useful for training draft-reply models.",

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


def render_geo_rm_platform(ctx):
    e04 = ctx.entity.get("E04_StateProfile")
    e08 = ctx.entity.get("E08_PlatformMix")
    e09 = ctx.entity.get("E09_PodGeography")
    e06 = ctx.entity.get("E06_CompanyType")

    narrative = prose(
        "Geography, accounting platform, and organizational structure provide context for where AI interventions "
        "should land first and whether a single solution pattern can work across the portfolio.",

        "The top three states by PMC count — California (306), Florida (189), and Texas (159) — also lead in "
        "deposits and case volume. These are the same states that rank Tier 1 in external market analysis, "
        "confirming that WAB's internal operational burden aligns with where the HOA market is largest.",

        "Platform mix reveals an important heterogeneity. Vantaca is the dominant accounting platform (239 PMCs, "
        "$4.1B deposits, 10,643 cases), followed by VMS (120 PMCs, $1.6B, 3,716 cases) and Caliber (137 PMCs, "
        "$1.5B, 3,611 cases). Different platforms may produce different case patterns, data structures, and "
        "integration requirements. A single AI copilot design may need platform-specific adaptations.",

        "Pod geography confirms that service teams are regionally aligned. WEST01 and WEST02 serve Arizona and "
        "California respectively; EAST03 covers Florida; Central01 spans Michigan, Ohio, and Texas. The pods "
        "with the widest geographic spread also tend to have the slowest median resolution times — suggesting "
        "that geographic dispersion may contribute to context-switching overhead.",
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
        callout("Draft Reply Assistance — MEDIUM feasibility",
                "2,423 emails from one day, 77% outbound. The top email-generating subjects (New Account Request 275, "
                "Research 253, Account Maintenance 235) show repetitive patterns. A draft-reply model trained on sent "
                "emails for the top 3–5 subjects could generate first-pass responses. The constraint is that we only "
                "have one day of email data to assess template coverage.", "warn") +
        callout("Workflow Copilot — MEDIUM, subject-specific",
                "The five slowest case subjects have median resolution times exceeding 1,000 hours — these are "
                "effectively projects, not cases. A copilot for the high-volume subjects (Research, New Account Request, "
                "Account Maintenance) that provides step-by-step guidance based on historical resolution patterns "
                "is more practical than attempting to accelerate the structurally slow subjects.", "warn")
    )

    parts = [narrative, evidence_prose]

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

    parts.append(so_what(
        "Three use cases have strong data support today: triage/routing, email summarization, and "
        "missing-information detection. Escalation prediction is high-value but needs labeled outcomes. "
        "Draft reply assistance is feasible but needs more email history to validate. "
        "The recommended starting point is missing-info detection — it requires no NLP, delivers immediate "
        "visibility, and builds trust in the AI-assisted workflow."
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
        "In parallel, two data-quality actions would strengthen the foundation for future use cases:",
    )

    data_actions = bullets([
        "Request a longer email extract (1–2 weeks minimum) to validate email-based use cases beyond a single-day sample.",
        "Investigate the $11.71B blank-entity anomaly and the Arizona negative-deposit data quality issue before presenting deposit figures to leadership.",
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
            "Cases: 3-month operating window (December 18, 2025 – March 19, 2026). Not a full annual history. Seasonal patterns cannot be assessed.",
            "Emails: 1-day sample only (March 11, 2026). Communication patterns observed may not generalize across weeks or months.",
            "PMCs and HOAs: Current-state snapshots, not longitudinal entity history. Changes over time cannot be tracked.",
            "An anomalous entity with $11.71 billion in deposits and no company name appears in the PMC file. It is likely a parent/system record and has been excluded from concentration narratives where noted.",
            "Arizona shows negative total deposits (-$53.8M). This is a data quality issue, not an economic signal.",
            "HOA Company Type is 99.1% null — this field is unusable for analysis.",
            "NAICS codes are mono-valued in both files (73% = 531311 in PMCs, 87% = 813990 in HOAs). NAICS provides no segmentation power and should be treated as a data-quality finding, not an analytical variable.",
            "The Case → HOA join rate is 7.6%. Direct case-to-HOA analysis is not supported. The valid path is Case → PMC → HOA.",
            "Internal/system cases (15.8% of total) are excluded from all client-facing metrics. They are identified by company names AAB ADMIN, WAB ADMIN, and blank entries.",
            "Use case feasibility ratings are evidence-based assessments from the available data, not implementation commitments.",
        ]) +
        '</div></details>'
    )


# ═══════════════════════════════════════════════════════════
#  SECTION REGISTRY & PAGE ASSEMBLY
# ═══════════════════════════════════════════════════════════

SECTIONS = [
    ("exec_summary",     "Executive Summary",                     render_executive),
    ("case_volume",      "Case Volume and Backlog Trajectory",    render_case_volume),
    ("subject_friction", "Where Cases Get Stuck",                 render_subject_friction),
    ("workload",         "Channels, Workload, and Capacity",      render_workload),
    ("data_quality",     "Data Quality and Joinability",          render_data_quality),
    ("pmc_portfolio",    "Client Portfolio, Deposits, and Risk",  render_pmc_portfolio),
    ("email_text",       "Email and Text Feasibility",            render_email_text),
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
