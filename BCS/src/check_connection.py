"""Step 0 — prove the AWS/Athena path works before running the story.

Checks, in order:
  1. STS identity (are the pasted keys valid?)
  2. SHOW DATABASES (is Athena reachable, are the two schemas visible?)
  3. DESCRIBE each of the three tables (column inventory)
  4. 3 sample rows from each table (values truncated)

Writes output/00_connection.json and prints a human summary.
"""

import json
import os
import sys
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import settings
from src import athena


def _truncate_row(row, width=60):
    out = {}
    for k, v in row.items():
        s = "" if v is None else str(v)
        out[k] = (s[: width - 1] + "…") if len(s) > width else s
    return out


def _parse_describe(rows):
    """DESCRIBE returns one varchar column with 'name \\t type \\t comment' lines."""
    cols = []
    for r in rows:
        line = next(iter(r.values()), "") or ""
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = [p.strip() for p in line.replace("\t", "|").split("|") if p.strip()]
        if parts:
            cols.append({"name": parts[0], "type": parts[1] if len(parts) > 1 else ""})
    return cols


def main():
    settings.ensure_output_dir()
    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "region": settings.AWS_REGION,
        "workgroup": settings.ATHENA_WORKGROUP,
        "checks": [],
    }
    ok = True

    # 1. Identity ---------------------------------------------------------
    print("[1/4] STS identity ...", flush=True)
    try:
        account, arn = athena.sts_identity()
        masked = account[:4] + "*" * (len(account) - 4)
        print(f"      OK  account {masked}  role {arn.rsplit('/', 1)[-1]}")
        report["checks"].append({"check": "sts", "status": "ok", "account_masked": masked, "arn": arn})
    except Exception as exc:  # noqa: BLE001
        hint = athena.friendly_credential_error(exc) or str(exc)
        print(f"      FAIL  {hint}")
        report["checks"].append({"check": "sts", "status": "fail", "error": hint})
        _write(report)
        return 1

    # 2. Databases --------------------------------------------------------
    print("[2/4] SHOW DATABASES ...", flush=True)
    try:
        res = athena.run("SHOW DATABASES")
        dbs = sorted(str(next(iter(r.values()))) for r in res["rows"])
        need = {settings.ACCT_DB, settings.CC_DB}
        missing = sorted(need - set(dbs))
        status = "ok" if not missing else "warn"
        print(f"      {status.upper()}  {len(dbs)} database(s) visible" + (f"; MISSING: {missing}" if missing else ""))
        report["checks"].append({"check": "databases", "status": status, "visible": dbs, "missing": missing})
        ok = ok and not missing
    except Exception as exc:  # noqa: BLE001
        hint = athena.friendly_credential_error(exc) or str(exc)
        print(f"      FAIL  {hint}")
        print("      If the error mentions an output location, set ATHENA_OUTPUT_S3 (see README).")
        report["checks"].append({"check": "databases", "status": "fail", "error": hint})
        _write(report)
        return 1

    # 3 + 4. Per-table describe + sample ----------------------------------
    print("[3/4] DESCRIBE tables + [4/4] sample rows ...", flush=True)
    for label, table in settings.TABLES.items():
        entry = {"check": f"table:{label}", "table": table}
        try:
            desc = athena.run(f"DESCRIBE {table}")
            cols = _parse_describe(desc["rows"])
            sample = athena.run(f"SELECT * FROM {table} LIMIT 3")
            entry.update(
                status="ok",
                column_count=len(cols),
                columns=cols,
                sample_rows=[_truncate_row(r) for r in sample["rows"]],
                sample_elapsed_s=sample["elapsed_s"],
            )
            print(f"      OK    {label}: {len(cols)} columns, {sample['row_count']} sample row(s)")
        except Exception as exc:  # noqa: BLE001
            entry.update(status="fail", error=str(exc))
            print(f"      FAIL  {label}: {exc}")
            ok = False
        report["checks"].append(entry)

    report["overall"] = "ok" if ok else "issues"
    _write(report)
    print(f"\nConnection check: {report['overall'].upper()}  ->  {settings.CONNECTION_JSON}")
    return 0 if ok else 2


def _write(report):
    settings.ensure_output_dir()
    with open(settings.CONNECTION_JSON, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=1, default=str)


if __name__ == "__main__":
    sys.exit(main())
