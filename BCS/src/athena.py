"""Minimal Athena query runner on plain boto3 (no pandas/pyathena needed).

Credentials come from the standard AWS environment variables
(AWS_ACCESS_KEY_ID / AWS_SECRET_ACCESS_KEY / AWS_SESSION_TOKEN).
Stable layer: copy once, should not need frequent changes.
"""

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import boto3
from botocore.exceptions import ClientError

from config import settings

_INT_TYPES = {"tinyint", "smallint", "int", "integer", "bigint"}
_FLOAT_TYPES = {"double", "float", "real", "decimal"}


class AthenaError(RuntimeError):
    pass


def session():
    return boto3.session.Session(region_name=settings.AWS_REGION)


def sts_identity():
    """Return (account, arn) for the current credentials. Raises on bad/expired keys."""
    ident = session().client("sts").get_caller_identity()
    return ident["Account"], ident["Arn"]


def friendly_credential_error(exc):
    """Map common credential failures to a plain instruction."""
    msg = str(exc)
    if "ExpiredToken" in msg or "InvalidClientTokenId" in msg or "SignatureDoesNotMatch" in msg:
        return (
            "AWS credentials are missing, wrong, or expired. Refresh them: copy the "
            "three 'set AWS_...' lines from your credentials page and paste them into "
            "creds.local.cmd (or straight into this cmd window), then rerun."
        )
    if "Unable to locate credentials" in msg or "NoCredentialsError" in type(exc).__name__:
        return (
            "No AWS credentials found in this shell. Paste your 'set AWS_...' block "
            "into creds.local.cmd (or this cmd window) and rerun."
        )
    return None


def _convert(value, athena_type):
    if value is None:
        return None
    t = athena_type.lower()
    try:
        if t in _INT_TYPES:
            return int(value)
        if t in _FLOAT_TYPES:
            return float(value)
        if t == "boolean":
            return value.lower() == "true"
    except (ValueError, AttributeError):
        pass
    return value


def run(sql, timeout_s=None, max_rows=None):
    """Execute one Athena query. Returns:

    {columns, rows (list of dicts), row_count, elapsed_s, scanned_mb, query_id}
    Raises AthenaError with Athena's reason string on failure.
    """
    timeout_s = timeout_s or settings.QUERY_TIMEOUT_S
    max_rows = max_rows or settings.MAX_RESULT_ROWS
    client = session().client("athena")

    start_args = {
        "QueryString": sql,
        "WorkGroup": settings.ATHENA_WORKGROUP,
        "QueryExecutionContext": {"Catalog": settings.ATHENA_CATALOG},
    }
    if settings.ATHENA_OUTPUT_S3:
        start_args["ResultConfiguration"] = {"OutputLocation": settings.ATHENA_OUTPUT_S3}

    t0 = time.time()
    try:
        qid = client.start_query_execution(**start_args)["QueryExecutionId"]
    except ClientError as exc:
        hint = friendly_credential_error(exc)
        raise AthenaError(hint or f"start_query_execution failed: {exc}") from exc

    # Poll until done.
    delay = 0.8
    while True:
        state_resp = client.get_query_execution(QueryExecutionId=qid)["QueryExecution"]
        state = state_resp["Status"]["State"]
        if state in ("SUCCEEDED", "FAILED", "CANCELLED"):
            break
        if time.time() - t0 > timeout_s:
            client.stop_query_execution(QueryExecutionId=qid)
            raise AthenaError(f"query timed out after {timeout_s}s (id {qid})")
        time.sleep(delay)
        delay = min(delay * 1.5, 4.0)

    if state != "SUCCEEDED":
        reason = state_resp["Status"].get("StateChangeReason", "no reason given")
        raise AthenaError(f"{state}: {reason}")

    stats = state_resp.get("Statistics", {})
    scanned_mb = round(stats.get("DataScannedInBytes", 0) / (1024 * 1024), 2)

    # Fetch results (paginated). First row of a SELECT result is the header.
    columns, types, rows = [], [], []
    paginator = client.get_paginator("get_query_results")
    first_page = True
    for page in paginator.paginate(QueryExecutionId=qid):
        if first_page:
            meta = page["ResultSet"]["ResultSetMetadata"]["ColumnInfo"]
            columns = [c["Name"] for c in meta]
            types = [c["Type"] for c in meta]
        for i, raw in enumerate(page["ResultSet"]["Rows"]):
            if first_page and i == 0:
                vals = [d.get("VarCharValue") for d in raw["Data"]]
                if vals == columns:  # header row
                    continue
            cells = [d.get("VarCharValue") for d in raw["Data"]]
            cells += [None] * (len(columns) - len(cells))
            rows.append(
                {col: _convert(cells[j], types[j]) for j, col in enumerate(columns)}
            )
            if len(rows) >= max_rows:
                break
        first_page = False
        if len(rows) >= max_rows:
            break

    return {
        "columns": columns,
        "rows": rows,
        "row_count": len(rows),
        "elapsed_s": round(time.time() - t0, 1),
        "scanned_mb": scanned_mb,
        "query_id": qid,
    }
