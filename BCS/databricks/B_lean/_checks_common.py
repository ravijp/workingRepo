# Databricks notebook source
# Shared check helpers for the B_lean *_checks siblings.
# Paste this block at the top of a checks file, or import it as a notebook.
# It is the ONLY place chk()/fmt() live now (the core files carry no asserts).
#
# A miss raises AssertionError and STOPS the run. That stop rule is the whole
# point: it is what caught the silent errors in the Ishant extract (see
# COMPARISON_REPORT.md Tier-4). The checks re-read the tables the core file
# just built; they never rebuild logic.

def fmt(v):
    if v is None:
        return "NULL"
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return f"{v:,}"
    if isinstance(v, float):
        return f"{v:,.0f}"
    return str(v)


def chk(name, actual, expected, tol=0, ctx=None):
    """Raising anchor check; a miss STOPS. expected=None -> just print (measure).
    ctx: optional DataFrame shown before raising, so a failure carries its own
    diagnosis. 5-arg signature kept identical to the original core chk()."""
    if expected is None:
        print(f"MEASURED  {name} = {fmt(actual)}")
        return
    ok = (abs(actual - expected) <= tol) if tol else (actual == expected)
    if not ok:
        if ctx is not None:
            print(f"CONTEXT for the failing check '{name}':")
            ctx.show(500, truncate=False)
        raise AssertionError(f"ANCHOR MISS {name}: got {fmt(actual)}, expected {fmt(expected)}"
                             + (f" (tol {tol})" if tol else ""))
    print(f"PASS  {name} = {fmt(actual)}")
