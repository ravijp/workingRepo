"""Runtime settings for the BCS Athena data story.

Stable layer: this file rarely changes. Anything environment-specific is
read from environment variables so nothing sensitive lives in the repo.
Table names here are used by the connection check; the analysis SQL in
sql/ is fully literal and self-contained.
"""

import os

# ---------------------------------------------------------------- AWS / Athena
AWS_REGION = (
    os.environ.get("AWS_DEFAULT_REGION")
    or os.environ.get("AWS_REGION")
    or "us-east-1"
)
ATHENA_WORKGROUP = os.environ.get("ATHENA_WORKGROUP", "primary")
# Only needed if the workgroup has no default query-result location.
ATHENA_OUTPUT_S3 = os.environ.get("ATHENA_OUTPUT_S3", "")
ATHENA_CATALOG = os.environ.get("ATHENA_CATALOG", "AwsDataCatalog")

# ---------------------------------------------------------------- Tables (connection check)
ACCT_DB = "fmt_acct_dba"
CC_DB = "contactcenter_bdp_db"
TABLES = {
    "fmt_acct_c": f'"{ACCT_DB}"."fmt_acct_c"',
    "call": f'"{CC_DB}"."call"',
    "transcript": f'"{CC_DB}"."transcript"',
}

# ---------------------------------------------------------------- Runtime knobs
QUERY_TIMEOUT_S = int(os.environ.get("QUERY_TIMEOUT_S", "600"))
MAX_RESULT_ROWS = int(os.environ.get("MAX_RESULT_ROWS", "2000"))

# ---------------------------------------------------------------- Paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SQL_DIR = os.path.join(BASE_DIR, "sql")
MANIFEST = os.path.join(SQL_DIR, "manifest.json")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
CONNECTION_JSON = os.path.join(OUTPUT_DIR, "00_connection.json")
STORY_JSON = os.path.join(OUTPUT_DIR, "story_data.json")
REPORT_HTML = os.path.join(OUTPUT_DIR, "bcs_story.html")


def ensure_output_dir():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
