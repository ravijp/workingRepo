# Databricks notebook source
# =====================================================================
# B_ishant / RUN_ALL.py
# One notebook to run the whole B_ishant pipeline (Ishant's client-blessed
# methodology). No copy-paste - each step %run's a sibling in this session.
#
# !!! TABLE COLLISION NOTICE (owner's explicit intent, 2026-07-22) !!!
#   This pipeline reuses B_lean's table names (uc2_t16_*) and every CREATE is
#   CREATE OR REPLACE TABLE. Running it OVERWRITES the live B_lean tables
#   uc2_t16_00n_acct_monthly and uc2_t16_01s_populations_<vintage>. That is the
#   owner's intent - one t16 table set, no duplicate uc2_ish_* tables. If you need
#   B_lean's exact 01s shape preserved, snapshot it before Run All.
#   Tables written (all CREATE OR REPLACE):
#     uc2_t16_00n_acct_monthly              (01; overwrites B_lean)
#     uc2_t16_01s_populations_<vintage>     (01 builds, 02 folds; overwrites B_lean)
#     uc2_t16_02n_episodes                  (02; call-grain window table)
#     uc2_t16_03r_roll_<vintage>            (03; derived roll cut, t16-style name)
#     uc2_t16_04t_frame_<vintage>           (04; transcript-review frame)
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
# MAGIC ## Step 1 - accounts + SAS funnel (610,183 -> 186,013 = in_sas_ledger) -> uc2_t16_00n_acct_monthly, uc2_t16_01s_populations_<vintage>

# COMMAND ----------

# MAGIC %run ./01_accounts

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 2 - statement-window classification (single-max anchor; 19,025 post-due) -> uc2_t16_02n_episodes + folds uc2_t16_01s_populations_<vintage>

# COMMAND ----------

# MAGIC %run ./02_windows

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 3 - roll + impairment (DQ1->DQ2 ~2,879 vs 3,306; ECL / CO / gross NCL; STG_CD_M2) -> uc2_t16_03r_roll_<vintage>

# COMMAND ----------

# MAGIC %run ./03_roll_impairment

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 4 - transcript sampling frame (feeds Copilot discovery) -> uc2_t16_04t_frame_<vintage>

# COMMAND ----------

# MAGIC %run ./04_transcript_frame

# COMMAND ----------

print("[B_ishant/RUN_ALL.py] RUN_ALL complete. Screenshot the filename-prefixed "
      "count blocks from each step. Key headline: post-due 19,025 -> roll (2,879 "
      "vs 3,306 UNRECONCILED) -> gross 12M NCL ~$7.599M. Open items: HRAM-exclusion "
      "split, 25-vs-28 day edge, 2,879-vs-3,306 roll, transfer/callback acctid.")
