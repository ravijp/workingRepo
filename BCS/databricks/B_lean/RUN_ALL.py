# Databricks notebook source
# MAGIC %md
# MAGIC # RUN_ALL - one notebook to run the whole B_lean pipeline (no copy-paste)
# MAGIC
# MAGIC Upload the `B_lean/` folder into your Databricks workspace (Repos, or
# MAGIC Workspace > Import > folder), open THIS notebook from inside that folder,
# MAGIC attach a cluster, and Run All. Every step below `%run`s a sibling notebook
# MAGIC in THIS SAME session, so the tables each step builds persist for the next
# MAGIC one. No pasting, no re-pasting SETUP.
# MAGIC
# MAGIC `%run ./X` requires X to sit in the SAME workspace folder as this notebook
# MAGIC and to be a notebook (the `# Databricks notebook source` header makes each
# MAGIC .py import as one). If a `%run` says "notebook not found", the folder
# MAGIC upload did not preserve the files as notebooks - re-import the folder (not
# MAGIC individual files) so Databricks registers them as notebooks.
# MAGIC
# MAGIC ORDER IS LOAD-BEARING: B02 builds the 00n-04n layers, B01 reads them and
# MAGIC builds 01s, B02b reads 01s. Do NOT reorder. A_recon_lock is standalone.

# COMMAND ----------

# MAGIC %md
# MAGIC ## Config - set once here; every step reads these widgets
# MAGIC (Each sibling notebook has the same defaults inline, so these only need
# MAGIC changing for a different vintage. Leave as-is for 202501.)

# COMMAND ----------

dbutils.widgets.text("CATALOG", "cda_model_shared")
dbutils.widgets.text("SCHEMA", "ecm_cld_model")
dbutils.widgets.text("ANCHOR_YM", "202501")
# The sibling notebooks read these widget names; %run shares the widget context.
print("Config:",
      dbutils.widgets.get("CATALOG"),
      dbutils.widgets.get("SCHEMA"),
      dbutils.widgets.get("ANCHOR_YM"))

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step A (standalone) - reconciliation lock
# MAGIC Reads the SAS csv + the round-10 tables; builds the uc2_sas_* recon tables.
# MAGIC Independent of the B chain - could run any time. If it hits the transient
# MAGIC AWS Glue "Unable to execute HTTP request" error, just re-run this cell.

# COMMAND ----------

# MAGIC %run ./A_recon_lock_202501

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 1 - B02 (builds 00n/01n/02n/03n/04n; the statement re-anchor lives here)

# COMMAND ----------

# MAGIC %run ./B02_keyfix_aws_layers

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 2 - B01 (reads 00n-04n, builds the 01s spine)

# COMMAND ----------

# MAGIC %run ./B01_sas_spine

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 3 - B02b (reads 01s + 02n + 04n, builds the 04s outcomes table)

# COMMAND ----------

# MAGIC %run ./B02b_outcomes_sas

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 4 - B03 (insight blocks; reads 01s + 04s)

# COMMAND ----------

# MAGIC %run ./B03_insights_sas

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 5 - the confirmation probe (run this after the window fix)
# MAGIC P1 is decisive: in_window_calls vs std_episodes. After the fix they should
# MAGIC be equal at ~19-22k (Ishant's verified ~19,025), not a 22k-vs-25 split.

# COMMAND ----------

# MAGIC %run ./B_window_probe

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 6 - the Story-B distribution (the deliverable: bucket distribution + old/new shift)

# COMMAND ----------

# MAGIC %run ./B_stmt_distribution

# COMMAND ----------

# MAGIC %md
# MAGIC ## Step 7 - the B04 sampler (per-wave; set WAVE/QUOTAS widgets if sampling)

# COMMAND ----------

# MAGIC %run ./B04_stmt_sampler

# COMMAND ----------

# MAGIC %md
# MAGIC ## Verification checks (run once to certify; safe to run every time)
# MAGIC Each _checks sibling re-reads the tables its core built and asserts the
# MAGIC frame-INDEPENDENT anchors (a miss STOPS) while reporting the frame-DEPENDENT
# MAGIC counts in measure-mode (they move under the re-anchor by design). If a
# MAGIC raising anchor STOPS, that is a real defect; read the message.

# COMMAND ----------

# MAGIC %run ./A_recon_lock_checks

# COMMAND ----------

# MAGIC %run ./B02_checks

# COMMAND ----------

# MAGIC %run ./B01_checks

# COMMAND ----------

# MAGIC %run ./B02b_checks

# COMMAND ----------

# MAGIC %run ./B03_checks

# COMMAND ----------

# MAGIC %run ./B04_checks

# COMMAND ----------

print("RUN_ALL complete. Read Step 5 (probe P1) and Step 6 (distribution) for "
      "the Story-B numbers; the _checks output for certification. Any raising "
      "ANCHOR MISS above is a real defect; measure-mode 'MEASURED ... (Jan ref "
      "...)' lines are the intended re-anchor movement.")
