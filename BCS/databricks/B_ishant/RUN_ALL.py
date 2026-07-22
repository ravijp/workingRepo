# Databricks notebook source
# =====================================================================
# RUN_ALL.py
# One notebook to run the whole pipeline. Each step %run's a sibling in this
# session, so there is no copy-paste between steps.
#
#   Tables written (all CREATE OR REPLACE):
#     uc2_t16_00n_acct_monthly              (01; monthly FMT account layer)
#     uc2_t16_01s_populations_<vintage>     (01 builds, 02 folds)
#     uc2_t16_02n_episodes                  (02; call-grain window table)
#     uc2_t16_03r_roll_<vintage>            (03; derived roll cut)
#     uc2_t16_04t_frame_<vintage>           (04; transcript-review frame)
#
# HOW TO RUN
#   Upload the folder into Databricks (Workspace > Import > folder, or Repos),
#   open THIS notebook from inside that folder, attach a cluster, Run All.
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
# Each build step PRINTS its own key counts inline (filename-prefixed) so a VDI
# screenshot captures the numbers.
# =====================================================================

# COMMAND ----------

# Config - set once here; every sibling reads these widgets (same defaults inline).
dbutils.widgets.text("CATALOG", "cda_model_shared")
dbutils.widgets.text("SCHEMA", "ecm_cld_model")
dbutils.widgets.text("ANCHOR_YM", "202501")
print("[RUN_ALL.py] Config:",
      dbutils.widgets.get("CATALOG"),
      dbutils.widgets.get("SCHEMA"),
      dbutils.widgets.get("ANCHOR_YM"))

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 1 - accounts + SAS funnel -> uc2_t16_00n_acct_monthly, uc2_t16_01s_populations_<vintage>

# COMMAND ----------

# MAGIC %run ./01_accounts

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 2 - statement-window classification (single-max anchor) -> uc2_t16_02n_episodes + folds uc2_t16_01s_populations_<vintage>

# COMMAND ----------

# MAGIC %run ./02_windows

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 3 - roll + impairment (DQ1->DQ2; ECL / CO / gross NCL; STG_CD_M2) -> uc2_t16_03r_roll_<vintage>

# COMMAND ----------

# MAGIC %run ./03_roll_impairment

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 4 - transcript sampling frame (feeds Copilot discovery) -> uc2_t16_04t_frame_<vintage>

# COMMAND ----------

# MAGIC %run ./04_transcript_frame

# COMMAND ----------

print("[RUN_ALL.py] RUN_ALL complete. Screenshot the filename-prefixed count "
      "blocks from each step. Open items: roll-cohort definition, HRAM-exclusion "
      "split, 25-vs-28 day edge, transfer/callback acctid.")
