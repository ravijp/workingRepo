# -*- coding: utf-8 -*-
"""
WAB Phase 1 — Run All Modules and Generate HTML Story
======================================================
One script to go from raw Excel files to the final HTML story.

BEFORE RUNNING:
  1. Edit the paths below to match your VDI environment
  2. Ensure all 4 WAB Excel files are in DATA_DIR
  3. Ensure the use-case workbook is at USECASE_FILE (or set to "" to skip)

WHAT THIS DOES (in order):
  Step 1: wab_internal_extract.py   → wab_internal_extract.xlsx   (16 sheets)
  Step 2: wab_cases_deep_dive.py    → wab_cases_deep_dive.xlsx    (16 sheets)
  Step 3: wab_entity_deep_dive.py   → wab_entity_deep_dive.xlsx   (11 sheets)
  Step 4: wab_email_deep_insights.py→ wab_email_deep_insights.xlsx (9 sheets)
  Step 5: wab_html_story.py         → story.html                  (11 sections)

RUNTIME: ~7 minutes total on VDI (Steps 1-4 are independent, Step 5 reads their outputs)

Dependencies: pandas, openpyxl, numpy, bs4, scikit-learn
Optional:     rapidfuzz (improves template detection in Step 4)
"""

import os
import sys
import time
import importlib

# ─────────────────────────────────────────────────────────
#  EDIT THESE PATHS BEFORE RUNNING
# ─────────────────────────────────────────────────────────

# Where the 4 WAB Excel source files live
DATA_DIR = r"C:\Users\YourName\Desktop"

# Where all output workbooks and HTML go
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_output"

# Use-case workbook (set to "" if not available)
USECASE_FILE = r"C:\Users\YourName\Desktop\WAB_Ops_UseCases_2026-03-18.xlsx"

# HTML story output directory
HTML_DIR = r"C:\Users\YourName\Desktop\wab_html_story"

# ─────────────────────────────────────────────────────────
#  SOURCE FILES (should not need editing unless filenames differ)
# ─────────────────────────────────────────────────────────

PMC_FILE   = os.path.join(DATA_DIR, "AAB - ALL PMCs.xlsx")
HOA_FILE   = os.path.join(DATA_DIR, "AAB - All HOAs.xlsx")
CASE_FILE  = os.path.join(DATA_DIR, "AAB All Cases.xlsx")
EMAIL_FILE = os.path.join(DATA_DIR, "ALL EMAIL Files.xlsx")

# ─────────────────────────────────────────────────────────
#  OUTPUT FILES (auto-derived)
# ─────────────────────────────────────────────────────────

INTERNAL_XLSX = os.path.join(OUTPUT_DIR, "wab_internal_extract.xlsx")
CASES_XLSX    = os.path.join(OUTPUT_DIR, "wab_cases_deep_dive.xlsx")
ENTITY_XLSX   = os.path.join(OUTPUT_DIR, "wab_entity_deep_dive.xlsx")
INSIGHTS_XLSX = os.path.join(OUTPUT_DIR, "wab_email_deep_insights.xlsx")


def check_inputs():
    """Verify all source files exist before running."""
    missing = []
    for label, path in [("PMC", PMC_FILE), ("HOA", HOA_FILE),
                         ("Cases", CASE_FILE), ("Emails", EMAIL_FILE)]:
        if not os.path.isfile(path):
            missing.append(f"  {label}: {path}")
    if missing:
        print("ERROR: Source files not found:")
        print("\n".join(missing))
        print("\nEdit DATA_DIR in this script to point to the folder with your WAB Excel files.")
        sys.exit(1)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(HTML_DIR, exist_ok=True)
    print(f"Source files OK. Output → {OUTPUT_DIR}")


def patch_and_run(module_name, overrides):
    """Import a module, patch its path variables, and run main()."""
    mod = importlib.import_module(module_name)
    for attr, val in overrides.items():
        if hasattr(mod, attr):
            setattr(mod, attr, val)
    mod.main()


def run_step(step_num, label, module_name, overrides):
    """Run one step with timing and error handling."""
    print(f"\n{'='*60}")
    print(f"  Step {step_num}/5: {label}")
    print(f"{'='*60}")
    t0 = time.time()
    try:
        patch_and_run(module_name, overrides)
        elapsed = time.time() - t0
        print(f"  Done in {elapsed:.1f}s")
        return True
    except Exception as e:
        elapsed = time.time() - t0
        print(f"  FAILED after {elapsed:.1f}s: {e}")
        return False


def main():
    print("WAB Phase 1 — Full Pipeline Run")
    print(f"Data:   {DATA_DIR}")
    print(f"Output: {OUTPUT_DIR}")
    print(f"HTML:   {HTML_DIR}")
    print()

    check_inputs()

    t_start = time.time()
    results = {}

    # Step 1: Internal Extract
    results["internal"] = run_step(1, "Internal Extract (file profiling, joins, text stats)",
        "wab_internal_extract", {
            "PMC_FILE": PMC_FILE, "HOA_FILE": HOA_FILE,
            "CASE_FILE": CASE_FILE, "EMAIL_FILE": EMAIL_FILE,
            "OUTPUT_DIR": OUTPUT_DIR,
        })

    # Step 2: Cases Deep Dive
    results["cases"] = run_step(2, "Cases Deep Dive (operations, triage delay, GenAI evidence)",
        "wab_cases_deep_dive", {
            "CASE_FILE": CASE_FILE, "EMAIL_FILE": EMAIL_FILE,
            "OUTPUT_DIR": OUTPUT_DIR,
        })

    # Step 3: Entity Deep Dive
    results["entity"] = run_step(3, "Entity Deep Dive (deposits, friction-value, RM coverage)",
        "wab_entity_deep_dive", {
            "PMC_FILE": PMC_FILE, "HOA_FILE": HOA_FILE,
            "CASE_FILE": CASE_FILE, "EMAIL_FILE": EMAIL_FILE,
            "OUTPUT_DIR": OUTPUT_DIR,
        })

    # Step 4: Email Deep Insights
    results["insights"] = run_step(4, "Email Deep Insights (NMF topics, templates, missing-info)",
        "wab_email_deep_insights", {
            "CASE_FILE": CASE_FILE, "EMAIL_FILE": EMAIL_FILE,
            "OUTPUT_DIR": OUTPUT_DIR,
        })

    # Step 5: HTML Story Generator
    html_overrides = {
        "INTERNAL_EXTRACT_XLSX": INTERNAL_XLSX,
        "CASES_DEEP_DIVE_XLSX": CASES_XLSX,
        "ENTITY_DEEP_DIVE_XLSX": ENTITY_XLSX,
        "EMAIL_INSIGHTS_XLSX": INSIGHTS_XLSX,
        "USECASE_XLSX": USECASE_FILE if USECASE_FILE and os.path.isfile(USECASE_FILE) else "",
        "OUTPUT_DIR": HTML_DIR,
    }
    results["html"] = run_step(5, "HTML Story Generator (11 sections)",
        "wab_html_story", html_overrides)

    # Summary
    elapsed_total = time.time() - t_start
    print(f"\n{'='*60}")
    print(f"  PIPELINE COMPLETE — {elapsed_total:.1f}s total")
    print(f"{'='*60}")
    for name, ok in results.items():
        status = "OK" if ok else "FAILED"
        print(f"  {name:12s}  {status}")

    if results.get("html"):
        story_path = os.path.join(HTML_DIR, "story.html")
        print(f"\n  HTML story: {story_path}")
        print(f"  Open in browser to view.")

    failed = [k for k, v in results.items() if not v]
    if failed:
        print(f"\n  WARNING: {len(failed)} step(s) failed: {', '.join(failed)}")
        print(f"  Check error messages above. The HTML story may be incomplete.")
        sys.exit(1)


if __name__ == "__main__":
    main()
