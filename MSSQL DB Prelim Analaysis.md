# Fast Analysis of All Views in DATALAKE

Here's a comprehensive script to get quick stats on all views without actually scanning the full data (which would be prohibitively slow on large views).

## Step 1: Get Metadata for All Views (Instant)

```sql
-- Get list of all views with their schemas and columns
SELECT 
    s.name AS SchemaName,
    v.name AS ViewName,
    s.name + '.' + v.name AS FullViewName,
    v.create_date,
    v.modify_date,
    COUNT(c.column_id) AS ColumnCount,
    STRING_AGG(c.name + ' (' + t.name + 
        CASE 
            WHEN t.name IN ('varchar','nvarchar','char','nchar') THEN '(' + 
                CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length AS VARCHAR) END + ')'
            WHEN t.name IN ('decimal','numeric') THEN '(' + CAST(c.precision AS VARCHAR) + ',' + CAST(c.scale AS VARCHAR) + ')'
            ELSE ''
        END + ')', ', ') WITHIN GROUP (ORDER BY c.column_id) AS Columns
FROM sys.views v
JOIN sys.schemas s ON v.schema_id = s.schema_id
JOIN sys.columns c ON v.object_id = c.object_id
JOIN sys.types t ON c.user_type_id = t.user_type_id
GROUP BY s.name, v.name, v.create_date, v.modify_date
ORDER BY s.name, v.name;
```

## Step 2: Fast Row Counts Using `TOP 0` + Approximate Methods

```sql
-- Quick approximate row counts using various fast methods
-- This uses TOP 1 to check if view is accessible/non-empty without full scan

SET NOCOUNT ON;

DECLARE @results TABLE (
    SchemaName NVARCHAR(128),
    ViewName NVARCHAR(128),
    HasData BIT,
    ErrorMsg NVARCHAR(500)
);

DECLARE @schema NVARCHAR(128), @view NVARCHAR(128), @sql NVARCHAR(MAX);
DECLARE @hasData BIT, @err NVARCHAR(500);

DECLARE view_cursor CURSOR FAST_FORWARD FOR
    SELECT s.name, v.name
    FROM sys.views v
    JOIN sys.schemas s ON v.schema_id = s.schema_id
    ORDER BY s.name, v.name;

OPEN view_cursor;
FETCH NEXT FROM view_cursor INTO @schema, @view;

WHILE @@FETCH_STATUS = 0
BEGIN
    SET @hasData = 0;
    SET @err = NULL;
    
    BEGIN TRY
        SET @sql = 'SELECT @out = CASE WHEN EXISTS (SELECT 1 FROM ' 
                   + QUOTENAME(@schema) + '.' + QUOTENAME(@view) + ') THEN 1 ELSE 0 END';
        EXEC sp_executesql @sql, N'@out BIT OUTPUT', @out = @hasData OUTPUT;
    END TRY
    BEGIN CATCH
        SET @err = LEFT(ERROR_MESSAGE(), 500);
    END CATCH

    INSERT INTO @results VALUES (@schema, @view, @hasData, @err);
    
    FETCH NEXT FROM view_cursor INTO @schema, @view;
END

CLOSE view_cursor;
DEALLOCATE view_cursor;

SELECT * FROM @results ORDER BY SchemaName, ViewName;
```

## Step 3: Sampled Profiling of Each View (The Main Script)

This samples **only 1000 rows** per view using `TOP` — keeps it fast:

```sql
SET NOCOUNT ON;

-- ============================================================
-- FAST VIEW PROFILER: Samples TOP(1000) rows from each view
-- Gives: row sample, NULLs, distinct approx, min, max per column
-- ============================================================

DECLARE @SampleSize INT = 1000;  -- Adjust if needed

-- Results table
IF OBJECT_ID('tempdb..#ViewProfile') IS NOT NULL DROP TABLE #ViewProfile;
CREATE TABLE #ViewProfile (
    SchemaName      NVARCHAR(128),
    ViewName        NVARCHAR(128),
    ColumnName      NVARCHAR(128),
    DataType        NVARCHAR(128),
    SampleRows      INT,
    NullCount       INT,
    NullPct         DECIMAL(5,2),
    DistinctCount   INT,
    MinValue        NVARCHAR(500),
    MaxValue        NVARCHAR(500),
    BlankOrEmptyCount INT,
    ErrorMsg        NVARCHAR(1000)
);

DECLARE @schema NVARCHAR(128), @view NVARCHAR(128);
DECLARE @sql NVARCHAR(MAX), @colSql NVARCHAR(MAX);

DECLARE view_cursor CURSOR FAST_FORWARD FOR
    SELECT s.name, v.name
    FROM sys.views v
    JOIN sys.schemas s ON v.schema_id = s.schema_id
    ORDER BY s.name, v.name;

OPEN view_cursor;
FETCH NEXT FROM view_cursor INTO @schema, @view;

WHILE @@FETCH_STATUS = 0
BEGIN
    BEGIN TRY
        -- Build per-column analysis dynamically
        SET @colSql = '';
        
        SELECT @colSql = @colSql + 
            'SELECT '
            + '''' + REPLACE(@schema, '''', '''''') + ''', '
            + '''' + REPLACE(@view, '''', '''''') + ''', '
            + '''' + REPLACE(c.name, '''', '''''') + ''', '
            + '''' + t.name + CASE 
                    WHEN t.name IN ('varchar','nvarchar','char','nchar') 
                        THEN '(' + CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length AS VARCHAR) END + ')'
                    WHEN t.name IN ('decimal','numeric') 
                        THEN '(' + CAST(c.precision AS VARCHAR) + ',' + CAST(c.scale AS VARCHAR) + ')'
                    ELSE '' END + ''', '
            + 'COUNT(*), '
            + 'SUM(CASE WHEN ' + QUOTENAME(c.name) + ' IS NULL THEN 1 ELSE 0 END), '
            + 'CAST(100.0 * SUM(CASE WHEN ' + QUOTENAME(c.name) + ' IS NULL THEN 1 ELSE 0 END) / NULLIF(COUNT(*),0) AS DECIMAL(5,2)), '
            + 'COUNT(DISTINCT ' + QUOTENAME(c.name) + '), '
            + CASE 
                WHEN t.name IN ('text','ntext','image','xml','geography','geometry','hierarchyid','varbinary','binary') 
                    THEN 'NULL, NULL, '
                WHEN t.name IN ('varchar','nvarchar','char','nchar') 
                    THEN 'MIN(LEFT(CAST(' + QUOTENAME(c.name) + ' AS NVARCHAR(500)),500)), '
                       + 'MAX(LEFT(CAST(' + QUOTENAME(c.name) + ' AS NVARCHAR(500)),500)), '
                ELSE 'CAST(MIN(' + QUOTENAME(c.name) + ') AS NVARCHAR(500)), '
                   + 'CAST(MAX(' + QUOTENAME(c.name) + ') AS NVARCHAR(500)), '
              END
            + CASE 
                WHEN t.name IN ('varchar','nvarchar','char','nchar') 
                    THEN 'SUM(CASE WHEN ' + QUOTENAME(c.name) + ' IS NOT NULL AND LTRIM(RTRIM(CAST(' + QUOTENAME(c.name) + ' AS NVARCHAR(500)))) = '''' THEN 1 ELSE 0 END)'
                ELSE 'NULL'
              END
            + ', NULL'  -- ErrorMsg
            + ' FROM #SampleData'
            + ' UNION ALL '
        FROM sys.columns c
        JOIN sys.types t ON c.user_type_id = t.user_type_id
        WHERE c.object_id = OBJECT_ID(QUOTENAME(@schema) + '.' + QUOTENAME(@view))
        ORDER BY c.column_id;

        -- Remove trailing UNION ALL
        IF LEN(@colSql) > 0
        BEGIN
            SET @colSql = LEFT(@colSql, LEN(@colSql) - 10); -- remove last ' UNION ALL'

            SET @sql = '
                SELECT TOP(' + CAST(@SampleSize AS VARCHAR) + ') * 
                INTO #SampleData 
                FROM ' + QUOTENAME(@schema) + '.' + QUOTENAME(@view) + ';
                
                INSERT INTO #ViewProfile 
                ' + @colSql + ';
                
                DROP TABLE #SampleData;';

            EXEC sp_executesql @sql;
        END
    END TRY
    BEGIN CATCH
        INSERT INTO #ViewProfile (SchemaName, ViewName, ColumnName, ErrorMsg)
        VALUES (@schema, @view, '*ALL*', LEFT(ERROR_MESSAGE(), 1000));
    END CATCH

    FETCH NEXT FROM view_cursor INTO @schema, @view;
END

CLOSE view_cursor;
DEALLOCATE view_cursor;

-- ============================================================
-- OUTPUT RESULTS
-- ============================================================

-- Summary per view
SELECT 
    SchemaName,
    ViewName,
    MAX(SampleRows) AS SampledRows,
    COUNT(DISTINCT ColumnName) AS ColumnsProfiled,
    MAX(ErrorMsg) AS Error
FROM #ViewProfile
GROUP BY SchemaName, ViewName
ORDER BY SchemaName, ViewName;

-- Detailed column-level profiling
SELECT 
    SchemaName,
    ViewName,
    ColumnName,
    DataType,
    SampleRows,
    NullCount,
    NullPct,
    DistinctCount,
    MinValue,
    MaxValue,
    BlankOrEmptyCount,
    ErrorMsg
FROM #ViewProfile
ORDER BY SchemaName, ViewName, ColumnName;

-- Potential issues: columns that are always NULL in sample
SELECT 
    SchemaName, ViewName, ColumnName, DataType
FROM #ViewProfile
WHERE NullPct = 100.00 AND ErrorMsg IS NULL
ORDER BY SchemaName, ViewName;

-- Low cardinality columns (potential categoricals/flags)
SELECT 
    SchemaName, ViewName, ColumnName, DataType, DistinctCount, SampleRows
FROM #ViewProfile
WHERE DistinctCount <= 20 AND DistinctCount > 0 AND ErrorMsg IS NULL
ORDER BY DistinctCount, SchemaName, ViewName;
```

## Step 4 (Optional): Quick Estimated Row Counts via TABLESAMPLE

If the views are backed by tables and you want approximate total row counts:

```sql
-- Check if underlying tables have row count stats (works for base tables referenced)
SELECT 
    s.name AS SchemaName,
    t.name AS TableName,
    SUM(p.rows) AS ApproxRowCount
FROM sys.tables t
JOIN sys.schemas s ON t.schema_id = s.schema_id
JOIN sys.partitions p ON t.object_id = p.object_id AND p.index_id IN (0,1)
GROUP BY s.name, t.name
ORDER BY ApproxRowCount DESC;
```

## Key Design Decisions for Speed

| Technique | Why |
|---|---|
| `TOP(1000)` sampling | Avoids full table scans on billion-row views |
| `SELECT INTO #temp` then analyze | Single pass over the view, then stats computed from temp table in memory |
| `EXISTS` check first | Catches broken views or permission issues without hanging |
| `FAST_FORWARD` cursor | Minimal overhead cursor type |
| Error handling per view | One broken view doesn't kill the entire run |
| `STRING_AGG` for column listing | Quick schema overview in one row per view |

**Expected runtime**: For ~50 views, typically **1–3 minutes** depending on view complexity and source latency. If any single view takes too long (e.g., complex joins without pushdown), you can add a timeout using `EXEC sp_executesql` wrapped with a `WAITFOR` check or simply reduce `@SampleSize` to 100.

## Smart sampling 

```sql
------------------------------------------------------------
-- THIS TOUCHES NO DATA - PURE METADATA - RUNS IN SECONDS
------------------------------------------------------------

-- 1. ALL VIEWS WITH THEIR COLUMNS AND TYPES
SELECT
    s.name AS SchemaName,
    v.name AS ViewName,
    c.column_id AS ColPosition,
    c.name AS ColumnName,
    t.name AS DataType,
    CASE
        WHEN t.name IN ('varchar','char')
            THEN CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length AS VARCHAR(10)) END
        WHEN t.name IN ('nvarchar','nchar')
            THEN CASE WHEN c.max_length = -1 THEN 'MAX' ELSE CAST(c.max_length / 2 AS VARCHAR(10)) END
        WHEN t.name IN ('decimal','numeric')
            THEN CAST(c.precision AS VARCHAR(10)) + ',' + CAST(c.scale AS VARCHAR(10))
        ELSE NULL
    END AS TypeDetail,
    c.is_nullable,
    c.is_identity,
    c.is_computed
FROM sys.views v
JOIN sys.schemas s ON v.schema_id = s.schema_id
JOIN sys.columns c ON v.object_id = c.object_id
JOIN sys.types t ON c.user_type_id = t.user_type_id
ORDER BY s.name, v.name, c.column_id;

-- 2. SUMMARY PER VIEW
SELECT
    s.name AS SchemaName,
    v.name AS ViewName,
    COUNT(*) AS ColumnCount,
    SUM(CASE WHEN t.name IN ('int','bigint','smallint','tinyint','decimal',
                              'numeric','float','real','money','smallmoney') THEN 1 ELSE 0 END) AS NumericCols,
    SUM(CASE WHEN t.name IN ('varchar','nvarchar','char','nchar','text','ntext') THEN 1 ELSE 0 END) AS StringCols,
    SUM(CASE WHEN t.name IN ('date','datetime','datetime2','smalldatetime',
                              'datetimeoffset','time') THEN 1 ELSE 0 END) AS DateCols,
    SUM(CASE WHEN t.name = 'bit' THEN 1 ELSE 0 END) AS BitCols,
    SUM(CASE WHEN t.name = 'uniqueidentifier' THEN 1 ELSE 0 END) AS GuidCols,
    v.create_date,
    v.modify_date
FROM sys.views v
JOIN sys.schemas s ON v.schema_id = s.schema_id
JOIN sys.columns c ON v.object_id = c.object_id
JOIN sys.types t ON c.user_type_id = t.user_type_id
GROUP BY s.name, v.name, v.create_date, v.modify_date
ORDER BY s.name, v.name;

-- 3. VIEW DEFINITIONS (the actual SQL behind each view)
-- THIS IS THE MOST VALUABLE - shows you what each view does
SELECT
    s.name AS SchemaName,
    v.name AS ViewName,
    m.definition AS ViewSQL
FROM sys.views v
JOIN sys.schemas s ON v.schema_id = s.schema_id
JOIN sys.sql_modules m ON v.object_id = m.object_id
ORDER BY s.name, v.name;

-- 4. FIND COMMON COLUMN NAMES ACROSS VIEWS (potential join keys)
SELECT
    c.name AS ColumnName,
    t.name AS DataType,
    COUNT(DISTINCT v.object_id) AS UsedInViewCount,
    LEFT(
        STRING_AGG(CAST(s.name + '.' + v.name AS VARCHAR(MAX)), ', ') 
        WITHIN GROUP (ORDER BY s.name, v.name),
        4000
    ) AS Views_Truncated
FROM sys.views v
JOIN sys.schemas s ON v.schema_id = s.schema_id
JOIN sys.columns c ON v.object_id = c.object_id
JOIN sys.types t ON c.user_type_id = t.user_type_id
GROUP BY c.name, t.name
HAVING COUNT(DISTINCT v.object_id) > 1
ORDER BY COUNT(DISTINCT v.object_id) DESC, c.name;

-- 5. COLUMN NAME PATTERN ANALYSIS (guess meaning from names)
SELECT
    s.name AS SchemaName,
    v.name AS ViewName,
    c.name AS ColumnName,
    t.name AS DataType,
    CASE
        WHEN c.name LIKE '%[_]ID' OR c.name LIKE '%[_]KEY' OR c.name LIKE '%[_]CD'
            THEN 'IDENTIFIER/KEY'
        WHEN c.name LIKE '%DATE%' OR c.name LIKE '%[_]DT' OR c.name LIKE '%[_]DTM'
            THEN 'DATE/TIMESTAMP'
        WHEN c.name LIKE '%AMT%' OR c.name LIKE '%AMOUNT%' OR c.name LIKE '%BAL%'
            OR c.name LIKE '%BALANCE%' OR c.name LIKE '%PRICE%' OR c.name LIKE '%RATE%'
            THEN 'MONETARY/AMOUNT'
        WHEN c.name LIKE '%CNT%' OR c.name LIKE '%COUNT%' OR c.name LIKE '%QTY%'
            OR c.name LIKE '%NUM%' OR c.name LIKE '%NBR%'
            THEN 'COUNT/QUANTITY'
        WHEN c.name LIKE '%FLAG%' OR c.name LIKE '%[_]FLG' OR c.name LIKE '%[_]IND'
            OR c.name LIKE '%INDICATOR%'
            THEN 'FLAG/INDICATOR'
        WHEN c.name LIKE '%NAME%' OR c.name LIKE '%[_]NM'
            THEN 'NAME'
        WHEN c.name LIKE '%DESC%' OR c.name LIKE '%DESCRIPTION%'
            THEN 'DESCRIPTION'
        WHEN c.name LIKE '%ADDR%' OR c.name LIKE '%ADDRESS%'
            THEN 'ADDRESS'
        WHEN c.name LIKE '%STATUS%' OR c.name LIKE '%[_]STS' OR c.name LIKE '%STATE%'
            THEN 'STATUS/STATE'
        WHEN c.name LIKE '%TYPE%' OR c.name LIKE '%[_]TYP' OR c.name LIKE '%CATEGORY%'
            THEN 'TYPE/CATEGORY'
        WHEN c.name LIKE '%CODE%' OR c.name LIKE '%[_]CD'
            THEN 'CODE'
        WHEN c.name LIKE '%PHONE%' OR c.name LIKE '%EMAIL%' OR c.name LIKE '%FAX%'
            THEN 'CONTACT'
        WHEN c.name LIKE '%SSN%' OR c.name LIKE '%TIN%' OR c.name LIKE '%EIN%'
            THEN 'PII-SENSITIVE'
        WHEN c.name LIKE '%ACCT%' OR c.name LIKE '%ACCOUNT%'
            THEN 'ACCOUNT'
        WHEN c.name LIKE '%CUST%' OR c.name LIKE '%CLIENT%'
            THEN 'CUSTOMER'
        WHEN c.name LIKE '%CREATE%' OR c.name LIKE '%INSERT%' OR c.name LIKE '%ADD%'
            THEN 'AUDIT-CREATED'
        WHEN c.name LIKE '%UPDATE%' OR c.name LIKE '%MODIFY%' OR c.name LIKE '%CHANGE%'
            THEN 'AUDIT-MODIFIED'
        ELSE 'OTHER'
    END AS GuessedMeaning
FROM sys.views v
JOIN sys.schemas s ON v.schema_id = s.schema_id
JOIN sys.columns c ON v.object_id = c.object_id
JOIN sys.types t ON c.user_type_id = t.user_type_id
ORDER BY s.name, v.name, c.column_id;

-- 6. IDENTIFY LIKELY PII COLUMNS
SELECT
    s.name AS SchemaName,
    v.name AS ViewName,
    c.name AS ColumnName,
    t.name AS DataType,
    'POTENTIAL PII' AS Warning
FROM sys.views v
JOIN sys.schemas s ON v.schema_id = s.schema_id
JOIN sys.columns c ON v.object_id = c.object_id
JOIN sys.types t ON c.user_type_id = t.user_type_id
WHERE c.name LIKE '%SSN%'
   OR c.name LIKE '%TIN%'
   OR c.name LIKE '%EIN%'
   OR c.name LIKE '%DOB%'
   OR c.name LIKE '%BIRTH%'
   OR c.name LIKE '%PHONE%'
   OR c.name LIKE '%EMAIL%'
   OR c.name LIKE '%ADDRESS%'
   OR c.name LIKE '%SALARY%'
   OR c.name LIKE '%ACCT%NUM%'
   OR c.name LIKE '%ROUTING%'
   OR c.name LIKE '%LICENSE%'
   OR c.name LIKE '%PASSPORT%'
ORDER BY s.name, v.name;

-- Pick the smallest-looking view and test manually
SELECT TOP(10) * FROM FISBIC_CURATED.VW_<smallest_view_name>;
```
