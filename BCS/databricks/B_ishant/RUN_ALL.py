# Databricks notebook source
# =====================================================================
# B_ishant / RUN_ALL.py
# One notebook to run the whole B_ishant pipeline (Ishant's client-blessed
# methodology). No copy-paste - each step %run's a sibling in this session.
#
# HOW TO RUN
#   Upload the B_ishant/ folder into Databricks (Workspace > Import > folder, or
#   Repos), open THIS notebook from inside that folder, attach a cluster, Run All.
#   `%run ./X` needs X to be a sibling notebook in the SAME folder; the
#   `# Databricks notebook source` header makes each .py import as one notebook.
#
# ORDER IS LOAD-BEARING
#   01 builds the account layer + SAS ledger funnel.
#   02 reads 01, applies the single-max statement anchor, builds the window pop.
#   03 reads 02, computes the DQ1->DQ2 roll + impairment headline.
#   04 reads 02 + 03, builds the transcript sampling frame.
#   Do NOT reorder.
#
# THIS PIPELINE HAS NO CHECK / CERTIFICATION MODULES BY DESIGN. Each build step
# PRINTS its own key counts inline (filename-prefixed) so a VDI screenshot
# captures the numbers. Read the printed "actual vs expected" lines to verify.
# =====================================================================

# COMMAND ----------

# Config - set once here; every sibling reads these widgets (same defaults inline).
dbutils.widgets.text("CATALOG", "cda_model_shared")
dbutils.widgets.text("SCHEMA", "ecm_cld_model")
dbutils.widgets.text("ANCHOR_YM", "202501")
print("[B_ishant/RUN_ALL.py] Config:",
      dbutils.widgets.get("CATALOG"),
      dbutils.widgets.get("SCHEMA"),
      dbutils.widgets.get("ANCHOR_YM"))

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 1 - accounts + SAS funnel (610,183 -> 186,013 = in_sas_ledger)

# COMMAND ----------

# MAGIC %run ./01_accounts

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 2 - statement-window classification (single-max anchor; 19,025 post-due)

# COMMAND ----------

# MAGIC %run ./02_windows

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 3 - roll + impairment (DQ1->DQ2 ~2,879 vs 3,306; ECL / CO / gross NCL; STG_CD_M2)

# COMMAND ----------

# MAGIC %run ./03_roll_impairment

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 4 - transcript sampling frame (feeds Copilot discovery)

# COMMAND ----------

# MAGIC %run ./04_transcript_frame

# COMMAND ----------

print("[B_ishant/RUN_ALL.py] RUN_ALL complete. Screenshot the filename-prefixed "
      "count blocks from each step. Key headline: post-due 19,025 -> roll (2,879 "
      "vs 3,306 UNRECONCILED) -> gross 12M NCL ~$7.599M. Open items: HRAM-exclusion "
      "split, 25-vs-28 day edge, 2,879-vs-3,306 roll, transfer/callback acctid.")
