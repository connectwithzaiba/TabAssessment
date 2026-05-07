-- ================================================================
-- TABLEAU → POWER BI MIGRATION ASSESSMENT
-- PostgreSQL Extraction Queries (Tableau Repository)
-- ================================================================
-- Database : workgroup
-- Port      : 8060  (Tableau Server internal Postgres)
-- Access    : psql -h localhost -p 8060 -U readonly -d workgroup
--
-- SECTIONS
--   0  · Diagnostics (run first — validates your environment)
--   1  · Workbooks Inventory
--   2  · Calculated Fields  ← main extraction (XPath)
--   3  · LOD Expressions with DAX Guidance
--   4  · Table Calculations
--   5  · Workbook Migration Complexity Summary
--   6  · Datasources Inventory
--   7  · Assessment Dashboard (counts & summary)
--   8  · CREATE VIEW Statements (save results for reuse)
-- ================================================================


-- ================================================================
-- SECTION 0 · DIAGNOSTICS
-- Run these first to validate your connection and schema.
-- ================================================================

-- 0a. Confirm the workbook XML column name and storage type
--     data_type = 'text'  → proceed with all queries below
--     data_type = 'bytea' → XML is gzip-compressed; see note at bottom
SELECT
    column_name,
    data_type,
    udt_name,
    character_maximum_length
FROM information_schema.columns
WHERE table_schema = 'public'
  AND table_name   = 'workbooks'
  AND column_name  IN ('workbook', 'workbook_xml', 'xml_data')
ORDER BY column_name;


-- 0b. Row counts by site — confirm scope before full extraction
SELECT
    s.name          AS site_name,
    COUNT(w.id)     AS total_workbooks,
    COUNT(CASE WHEN w.workbook IS NOT NULL THEN 1 END) AS with_xml
FROM workbooks w
JOIN sites s ON s.id = w.site_id
GROUP BY s.name
ORDER BY total_workbooks DESC;


-- 0c. Confirm XML is parseable — sample the first workbook
--     Should return a non-empty snippet of the workbook tag
SELECT
    name                                        AS workbook_name,
    left(workbook::text, 200)                   AS xml_preview,
    length(workbook::text)                      AS xml_chars,
    pg_size_pretty(length(workbook::text))      AS xml_size
FROM workbooks
WHERE workbook IS NOT NULL
ORDER BY length(workbook::text) DESC
LIMIT 5;


-- 0d. Quick sanity: how many workbooks contain calculated fields?
SELECT
    COUNT(DISTINCT w.id)    AS total_workbooks,
    COUNT(DISTINCT CASE
        WHEN w.workbook::text ~ '<calculation[^>]*formula=' THEN w.id
    END)                    AS workbooks_with_calcs,
    COUNT(DISTINCT CASE
        WHEN w.workbook::text ~ '\{(FIXED|INCLUDE|EXCLUDE)\b' THEN w.id
    END)                    AS workbooks_with_lod
FROM workbooks w
WHERE w.workbook IS NOT NULL;


-- ================================================================
-- SECTION 1 · WORKBOOKS INVENTORY
-- All workbooks across all sites and projects.
-- ================================================================

SELECT
    s.name                              AS site_name,
    COALESCE(p.name, 'Default')         AS project_name,
    w.name                              AS workbook_name,
    COALESCE(w.owner_name, '')          AS owner_name,
    COALESCE(w.repository_url, '')      AS repository_url,
    w.created_at::date                  AS created_date,
    w.updated_at::date                  AS updated_date,
    pg_size_pretty(length(w.workbook::text)) AS xml_size
FROM workbooks   w
JOIN sites       s ON s.id = w.site_id
LEFT JOIN projects p ON p.id = w.project_id
WHERE w.workbook IS NOT NULL
ORDER BY s.name, p.name, w.name;


-- ================================================================
-- SECTION 2 · CALCULATED FIELDS  (XPath extraction)
-- Parses every workbook's XML to extract calculated field names,
-- formulas, data types, LOD expressions, and table calcs.
-- ================================================================

WITH

-- 2-A: Base workbooks joined to site and project
workbooks_base AS (
    SELECT
        w.id                                AS workbook_id,
        s.name                              AS site_name,
        COALESCE(p.name, 'Default')         AS project_name,
        w.name                              AS workbook_name,
        COALESCE(w.owner_name, '')          AS owner_name,
        w.created_at::date                  AS created_date,
        w.updated_at::date                  AS updated_date,
        w.workbook::xml                     AS workbook_xml
        -- NOTE: if the cast above fails (data_type = bytea), see Section 0a note.
    FROM workbooks w
    JOIN sites     s ON s.id = w.site_id
    LEFT JOIN projects p ON p.id = w.project_id
    WHERE w.workbook IS NOT NULL
),

-- 2-B: Expand to datasource-level XML nodes (one row per datasource per workbook)
ds_nodes AS (
    SELECT
        wb.workbook_id,
        wb.site_name,
        wb.project_name,
        wb.workbook_name,
        wb.owner_name,
        wb.created_date,
        wb.updated_date,
        ds_node
    FROM workbooks_base wb
    CROSS JOIN LATERAL
        unnest(xpath(
            -- Select non-internal datasources only
            '//datasource[not(@name="Parameters") and not(starts-with(@name,"_")) and @name!=""]',
            wb.workbook_xml
        )) AS ds_node
),

-- 2-C: Expand to column-level XML nodes that have a <calculation> child
col_nodes AS (
    SELECT
        dsn.workbook_id,
        dsn.site_name,
        dsn.project_name,
        dsn.workbook_name,
        dsn.owner_name,
        dsn.created_date,
        dsn.updated_date,
        COALESCE(
            (xpath('@caption', dsn.ds_node))[1]::text,
            (xpath('@name',    dsn.ds_node))[1]::text,
            'Unknown'
        )                                   AS datasource_caption,
        (xpath('@name', dsn.ds_node))[1]::text AS datasource_name,
        col_node
    FROM ds_nodes dsn
    CROSS JOIN LATERAL
        unnest(xpath(
            -- Columns that have a calculation child and are not hidden
            'column[calculation and not(@hidden="true")]',
            dsn.ds_node
        )) AS col_node
),

-- 2-D: Pull individual attributes from each column node
fields_raw AS (
    SELECT
        cn.site_name,
        cn.project_name,
        cn.workbook_name,
        cn.owner_name,
        cn.created_date,
        cn.updated_date,
        cn.datasource_caption,
        cn.datasource_name,

        -- Human-readable caption (display name in Tableau)
        COALESCE(
            (xpath('@caption', cn.col_node))[1]::text,
            TRIM(BOTH '[]' FROM (xpath('@name', cn.col_node))[1]::text)
        )                                   AS caption,

        -- Internal field name  (e.g. "[My Calculated Field]")
        (xpath('@name', cn.col_node))[1]::text AS field_name,

        -- Tableau data type
        LOWER(COALESCE(
            (xpath('@datatype', cn.col_node))[1]::text,
            'string'
        ))                                  AS datatype,

        -- Dimension or measure
        LOWER(COALESCE(
            (xpath('@role', cn.col_node))[1]::text,
            ''
        ))                                  AS role,

        -- The calculation formula
        COALESCE(
            (xpath('calculation/@formula', cn.col_node))[1]::text,
            ''
        )                                   AS formula,

        -- calculation class: 'tableau' (standard) or 'table-calculation'
        LOWER(COALESCE(
            (xpath('calculation/@class', cn.col_node))[1]::text,
            'tableau'
        ))                                  AS calc_class

    FROM col_nodes cn
)

-- 2-E: Final SELECT with derived columns
SELECT
    site_name,
    project_name,
    workbook_name,
    owner_name,
    datasource_caption                                  AS datasource,
    caption,
    field_name,
    datatype                                            AS tableau_type,

    -- Power BI data type equivalent
    CASE datatype
        WHEN 'string'   THEN 'Text'
        WHEN 'integer'  THEN 'Whole Number'
        WHEN 'real'     THEN 'Decimal Number'
        WHEN 'boolean'  THEN 'True/False'
        WHEN 'date'     THEN 'Date'
        WHEN 'datetime' THEN 'Date/Time'
        WHEN 'spatial'  THEN 'Geography'
        ELSE            datatype
    END                                                 AS pbi_type,

    role,
    formula,

    -- ── LOD Detection ─────────────────────────────────────────
    CASE WHEN formula ~ '\{(FIXED|INCLUDE|EXCLUDE)\b'
         THEN true ELSE false
    END                                                 AS has_lod,

    -- All LOD types present in this formula (handles multiple)
    COALESCE(
        ARRAY_TO_STRING(
            ARRAY(
                SELECT DISTINCT UPPER(m[1])
                FROM regexp_matches(formula, '\{(FIXED|INCLUDE|EXCLUDE)\b', 'gi') m
            ),
            ', '
        ),
        ''
    )                                                   AS lod_types,

    -- ── Table Calculation Detection ───────────────────────────
    CASE WHEN formula ~ '\m(RUNNING_SUM|RUNNING_AVG|RUNNING_COUNT'
                      '|RUNNING_MAX|RUNNING_MIN'
                      '|WINDOW_SUM|WINDOW_AVG|WINDOW_COUNT'
                      '|WINDOW_MAX|WINDOW_MIN'
                      '|LOOKUP|FIRST|LAST|INDEX|SIZE|TOTAL'
                      '|RANK|RANK_DENSE|RANK_MODIFIED'
                      '|RANK_PERCENTILE|RANK_UNIQUE'
                      '|PREVIOUS_VALUE)\s*\('
         THEN true ELSE false
    END                                                 AS has_table_calc,

    -- ── Migration Complexity ──────────────────────────────────
    CASE
        WHEN formula ~ '\{(FIXED|INCLUDE|EXCLUDE)\b'
          OR formula ~ '\m(RUNNING_SUM|RUNNING_AVG|WINDOW_SUM|WINDOW_AVG'
                       '|LOOKUP|INDEX|RANK|SIZE|TOTAL|PREVIOUS_VALUE)\s*\('
            THEN 'HIGH'
        WHEN formula ~ '\m(IF|CASE|IIF|ELSEIF)\M'
          OR formula ~ '\mDATE(DIFF|PART|ADD|NAME|TRUNC)?\M'
            THEN 'MEDIUM'
        WHEN formula IS NOT NULL AND formula != ''
            THEN 'LOW'
        ELSE 'NONE'
    END                                                 AS migration_complexity,

    created_date,
    updated_date

FROM fields_raw
WHERE formula != ''
ORDER BY
    migration_complexity,          -- HIGH first
    site_name,
    project_name,
    workbook_name,
    caption;


-- ================================================================
-- SECTION 3 · LOD EXPRESSIONS WITH DAX GUIDANCE
-- Filters to LOD-only fields; adds Power BI / DAX migration notes.
-- ================================================================

WITH

workbooks_base AS (
    SELECT
        w.id,
        s.name                              AS site_name,
        COALESCE(p.name, 'Default')         AS project_name,
        w.name                              AS workbook_name,
        w.workbook::xml                     AS workbook_xml
    FROM workbooks w
    JOIN sites     s ON s.id = w.site_id
    LEFT JOIN projects p ON p.id = w.project_id
    WHERE w.workbook IS NOT NULL
),

ds_nodes AS (
    SELECT wb.id, wb.site_name, wb.project_name, wb.workbook_name, ds_node
    FROM workbooks_base wb
    CROSS JOIN LATERAL
        unnest(xpath(
            '//datasource[not(@name="Parameters") and not(starts-with(@name,"_")) and @name!=""]',
            wb.workbook_xml
        )) AS ds_node
),

col_nodes AS (
    SELECT
        dsn.site_name,
        dsn.project_name,
        dsn.workbook_name,
        COALESCE(
            (xpath('@caption', dsn.ds_node))[1]::text,
            (xpath('@name',    dsn.ds_node))[1]::text
        )                                       AS datasource_caption,
        col_node
    FROM ds_nodes dsn
    CROSS JOIN LATERAL
        unnest(xpath(
            'column[calculation and not(@hidden="true")]',
            dsn.ds_node
        )) AS col_node
),

lod_fields AS (
    SELECT
        cn.site_name,
        cn.project_name,
        cn.workbook_name,
        cn.datasource_caption,
        COALESCE(
            (xpath('@caption', cn.col_node))[1]::text,
            TRIM(BOTH '[]' FROM (xpath('@name', cn.col_node))[1]::text)
        )                                       AS caption,
        (xpath('@name', cn.col_node))[1]::text  AS field_name,
        COALESCE(
            (xpath('calculation/@formula', cn.col_node))[1]::text,
            ''
        )                                       AS formula
    FROM col_nodes cn
)

SELECT
    site_name,
    project_name,
    workbook_name,
    datasource_caption                          AS datasource,
    caption,
    field_name,

    -- Extract each LOD instance from the formula
    ARRAY_TO_STRING(
        ARRAY(
            SELECT UPPER(m[1])
            FROM regexp_matches(formula, '\{(FIXED|INCLUDE|EXCLUDE)\b[^}]*\}', 'gi') m
        ),
        ' | '
    )                                           AS lod_expressions_found,

    -- LOD type(s) present
    ARRAY_TO_STRING(
        ARRAY(
            SELECT DISTINCT UPPER(m[1])
            FROM regexp_matches(formula, '\{(FIXED|INCLUDE|EXCLUDE)\b', 'gi') m
        ),
        ', '
    )                                           AS lod_types,

    formula,

    -- DAX migration guidance per LOD type
    CASE
        WHEN formula ~ '\{FIXED\b' AND formula ~ '\{INCLUDE\b'
            THEN 'Multiple LOD types — decompose into separate DAX measures. FIXED → CALCULATE with REMOVEFILTERS(); INCLUDE → CALCULATE with additional FILTER()'
        WHEN formula ~ '\{FIXED\b' AND formula ~ '\{EXCLUDE\b'
            THEN 'Multiple LOD types — FIXED → CALCULATE with REMOVEFILTERS(); EXCLUDE → CALCULATE with REMOVEFILTERS(<specific column>)'
        WHEN formula ~ '\{FIXED\b'
            THEN 'FIXED: Use CALCULATE(<expr>, ALL(<table>)) or REMOVEFILTERS(<col>) to override filter context. Example: CALCULATE(SUM(Sales[Amount]), ALL(Geography[Region]))'
        WHEN formula ~ '\{INCLUDE\b'
            THEN 'INCLUDE: Add granularity using SUMMARIZE or CALCULATE with FILTER. Often maps to a helper measure or AVERAGEX over a virtual table.'
        WHEN formula ~ '\{EXCLUDE\b'
            THEN 'EXCLUDE: Use CALCULATE(<expr>, REMOVEFILTERS(<specific column>)) to remove a specific dimension from context.'
        ELSE 'Review formula — multiple LOD patterns detected'
    END                                         AS dax_migration_guidance,

    'HIGH'                                      AS migration_complexity

FROM lod_fields
WHERE formula ~ '\{(FIXED|INCLUDE|EXCLUDE)\b'
ORDER BY
    CASE
        WHEN formula ~ '\{FIXED\b'   THEN 1
        WHEN formula ~ '\{INCLUDE\b' THEN 2
        WHEN formula ~ '\{EXCLUDE\b' THEN 3
        ELSE 4
    END,
    site_name, project_name, workbook_name, caption;


-- ================================================================
-- SECTION 4 · TABLE CALCULATIONS
-- ================================================================

WITH

workbooks_base AS (
    SELECT
        w.id, s.name AS site_name,
        COALESCE(p.name, 'Default') AS project_name,
        w.name AS workbook_name,
        w.workbook::xml AS workbook_xml
    FROM workbooks w
    JOIN sites s ON s.id = w.site_id
    LEFT JOIN projects p ON p.id = w.project_id
    WHERE w.workbook IS NOT NULL
),

ds_nodes AS (
    SELECT wb.id, wb.site_name, wb.project_name, wb.workbook_name, ds_node
    FROM workbooks_base wb
    CROSS JOIN LATERAL
        unnest(xpath(
            '//datasource[not(@name="Parameters") and not(starts-with(@name,"_"))]',
            wb.workbook_xml
        )) AS ds_node
),

col_nodes AS (
    SELECT
        dsn.site_name, dsn.project_name, dsn.workbook_name,
        COALESCE((xpath('@caption', dsn.ds_node))[1]::text,
                 (xpath('@name',    dsn.ds_node))[1]::text) AS datasource_caption,
        col_node
    FROM ds_nodes dsn
    CROSS JOIN LATERAL
        unnest(xpath('column[calculation and not(@hidden="true")]', dsn.ds_node)) AS col_node
),

tc_fields AS (
    SELECT
        cn.site_name, cn.project_name, cn.workbook_name, cn.datasource_caption,
        COALESCE(
            (xpath('@caption', cn.col_node))[1]::text,
            TRIM(BOTH '[]' FROM (xpath('@name', cn.col_node))[1]::text)
        ) AS caption,
        (xpath('@name', cn.col_node))[1]::text AS field_name,
        COALESCE((xpath('calculation/@formula', cn.col_node))[1]::text, '') AS formula
    FROM col_nodes cn
)

SELECT
    site_name,
    project_name,
    workbook_name,
    datasource_caption          AS datasource,
    caption,
    field_name,
    formula,

    -- Identify which table calc function(s) are used
    ARRAY_TO_STRING(
        ARRAY(
            SELECT DISTINCT UPPER(m[1])
            FROM regexp_matches(
                formula,
                '\m(RUNNING_SUM|RUNNING_AVG|RUNNING_COUNT|RUNNING_MAX|RUNNING_MIN'
                '|WINDOW_SUM|WINDOW_AVG|WINDOW_COUNT|WINDOW_MAX|WINDOW_MIN'
                '|LOOKUP|FIRST|LAST|INDEX|SIZE|TOTAL'
                '|RANK|RANK_DENSE|RANK_MODIFIED|RANK_PERCENTILE|RANK_UNIQUE'
                '|PREVIOUS_VALUE)\M',
                'gi'
            ) m
        ),
        ', '
    )                           AS table_calc_functions,

    -- DAX equivalent guidance
    CASE
        WHEN formula ~ '\mRUNNING_SUM\M'
            THEN 'Running total → CALCULATE(SUM(...), FILTER(ALL(Dates), Dates[Date] <= MAX(Dates[Date])))'
        WHEN formula ~ '\mWINDOW_SUM\M'
            THEN 'Window sum → Consider CALCULATE with DATESINPERIOD or a RANKX-based sliding window measure'
        WHEN formula ~ '\mRANK\M'
            THEN 'Rank → RANKX(ALL(<table>[<col>]), <measure>)'
        WHEN formula ~ '\mLOOKUP\M'
            THEN 'Lookup (relative) → CALCULATE(<measure>, DATEADD or OFFSET logic in DAX)'
        WHEN formula ~ '\mINDEX\M'
            THEN 'Index → RANKX(ALLSELECTED(), <measure>,, ASC, Dense)'
        ELSE 'Review table calc pattern — manual DAX translation required'
    END                         AS dax_migration_guidance,

    'HIGH'                      AS migration_complexity

FROM tc_fields
WHERE formula ~ '\m(RUNNING_SUM|RUNNING_AVG|RUNNING_COUNT|RUNNING_MAX|RUNNING_MIN'
                '|WINDOW_SUM|WINDOW_AVG|WINDOW_COUNT|WINDOW_MAX|WINDOW_MIN'
                '|LOOKUP|FIRST|LAST|INDEX|SIZE|TOTAL'
                '|RANK|RANK_DENSE|RANK_MODIFIED|RANK_PERCENTILE|RANK_UNIQUE'
                '|PREVIOUS_VALUE)\s*\('
ORDER BY site_name, project_name, workbook_name, caption;


-- ================================================================
-- SECTION 5 · WORKBOOK MIGRATION COMPLEXITY SUMMARY
-- Per-workbook rollup: field counts, LOD counts, complexity score.
-- ================================================================

WITH

workbooks_base AS (
    SELECT
        w.id, s.name AS site_name,
        COALESCE(p.name, 'Default') AS project_name,
        w.name AS workbook_name,
        COALESCE(w.owner_name, '') AS owner_name,
        w.created_at::date AS created_date,
        w.updated_at::date AS updated_date,
        w.workbook::xml AS workbook_xml
    FROM workbooks w
    JOIN sites s ON s.id = w.site_id
    LEFT JOIN projects p ON p.id = w.project_id
    WHERE w.workbook IS NOT NULL
),

ds_nodes AS (
    SELECT wb.id, wb.site_name, wb.project_name, wb.workbook_name,
           wb.owner_name, wb.created_date, wb.updated_date, ds_node
    FROM workbooks_base wb
    CROSS JOIN LATERAL
        unnest(xpath(
            '//datasource[not(@name="Parameters") and not(starts-with(@name,"_")) and @name!=""]',
            wb.workbook_xml
        )) AS ds_node
),

-- All columns (calculated + non-calculated)
all_cols AS (
    SELECT
        dsn.id, dsn.site_name, dsn.project_name, dsn.workbook_name,
        dsn.owner_name, dsn.created_date, dsn.updated_date,
        col_node
    FROM ds_nodes dsn
    CROSS JOIN LATERAL
        unnest(xpath('column[not(@hidden="true")]', dsn.ds_node)) AS col_node
),

-- Aggregate per workbook
workbook_stats AS (
    SELECT
        id, site_name, project_name, workbook_name, owner_name,
        created_date, updated_date,

        COUNT(*)                                                AS total_fields,

        COUNT(CASE
            WHEN (xpath('calculation/@formula', col_node))[1]::text IS NOT NULL
             AND (xpath('calculation/@formula', col_node))[1]::text != ''
            THEN 1 END)                                         AS calc_fields,

        COUNT(CASE
            WHEN (xpath('calculation/@formula', col_node))[1]::text
                 ~ '\{(FIXED|INCLUDE|EXCLUDE)\b'
            THEN 1 END)                                         AS lod_fields,

        COUNT(CASE
            WHEN (xpath('calculation/@formula', col_node))[1]::text
                 ~ '\m(RUNNING_SUM|RUNNING_AVG|WINDOW_SUM|WINDOW_AVG'
                   '|LOOKUP|INDEX|RANK|SIZE|TOTAL|PREVIOUS_VALUE)\s*\('
            THEN 1 END)                                         AS table_calc_fields,

        COUNT(CASE
            WHEN (xpath('calculation/@formula', col_node))[1]::text
                 IS NOT NULL
             AND (xpath('calculation/@formula', col_node))[1]::text
                 ~ '\m(IF|CASE|IIF|ELSEIF)\M'
            THEN 1 END)                                         AS conditional_fields

    FROM all_cols
    GROUP BY id, site_name, project_name, workbook_name,
             owner_name, created_date, updated_date
)

SELECT
    site_name,
    project_name,
    workbook_name,
    owner_name,
    total_fields,
    calc_fields,
    lod_fields,
    table_calc_fields,
    conditional_fields,

    -- Migration complexity score
    CASE
        WHEN lod_fields > 5 OR table_calc_fields > 5
            THEN 'HIGH'
        WHEN lod_fields > 0 OR table_calc_fields > 0 OR calc_fields > 20
            THEN 'MEDIUM'
        WHEN calc_fields > 0
            THEN 'LOW'
        ELSE 'MINIMAL'
    END                                     AS migration_complexity,

    -- Estimated DAX re-engineering effort (rough guide)
    CASE
        WHEN lod_fields > 5 OR table_calc_fields > 5
            THEN '> 3 days'
        WHEN lod_fields > 0 OR table_calc_fields > 0
            THEN '1–3 days'
        WHEN calc_fields > 10
            THEN '< 1 day'
        ELSE 'Hours'
    END                                     AS estimated_effort,

    created_date,
    updated_date

FROM workbook_stats
ORDER BY
    CASE
        WHEN lod_fields > 5 OR table_calc_fields > 5  THEN 1
        WHEN lod_fields > 0 OR table_calc_fields > 0  THEN 2
        ELSE 3
    END,
    site_name, project_name, workbook_name;


-- ================================================================
-- SECTION 6 · DATASOURCES INVENTORY
-- Connection type, server, and database for each datasource.
-- ================================================================

WITH

workbooks_base AS (
    SELECT
        w.id, s.name AS site_name,
        COALESCE(p.name, 'Default') AS project_name,
        w.name AS workbook_name,
        w.workbook::xml AS workbook_xml
    FROM workbooks w
    JOIN sites s ON s.id = w.site_id
    LEFT JOIN projects p ON p.id = w.project_id
    WHERE w.workbook IS NOT NULL
),

ds_nodes AS (
    SELECT
        wb.site_name, wb.project_name, wb.workbook_name, ds_node
    FROM workbooks_base wb
    CROSS JOIN LATERAL
        unnest(xpath(
            '//datasource[not(@name="Parameters") and not(starts-with(@name,"_")) and @name!=""]',
            wb.workbook_xml
        )) AS ds_node
),

-- Primary connection (direct connection)
direct_conn AS (
    SELECT
        dsn.site_name, dsn.project_name, dsn.workbook_name,
        (xpath('@caption', dsn.ds_node))[1]::text  AS datasource_caption,
        (xpath('@name',    dsn.ds_node))[1]::text  AS datasource_name,
        conn_node
    FROM ds_nodes dsn
    CROSS JOIN LATERAL
        unnest(xpath('connection', dsn.ds_node)) AS conn_node
),

-- Named connections (federated / multi-source)
named_conn AS (
    SELECT
        dsn.site_name, dsn.project_name, dsn.workbook_name,
        (xpath('@caption', dsn.ds_node))[1]::text  AS datasource_caption,
        (xpath('@name',    dsn.ds_node))[1]::text  AS datasource_name,
        conn_node
    FROM ds_nodes dsn
    CROSS JOIN LATERAL
        unnest(xpath('named-connections/named-connection/connection', dsn.ds_node))
            AS conn_node
),

-- Union direct + named, deduplicate
all_conn AS (
    SELECT * FROM direct_conn
    UNION ALL
    SELECT * FROM named_conn
)

SELECT
    site_name,
    project_name,
    workbook_name,
    COALESCE(datasource_caption, datasource_name)       AS datasource,
    COALESCE(
        (xpath('@class',  conn_node))[1]::text, ''
    )                                                   AS connection_type,
    COALESCE(
        (xpath('@server', conn_node))[1]::text, ''
    )                                                   AS server,
    COALESCE(
        (xpath('@dbname', conn_node))[1]::text,
        (xpath('@database', conn_node))[1]::text, ''
    )                                                   AS database_name,
    COALESCE(
        (xpath('@schema', conn_node))[1]::text, ''
    )                                                   AS schema_name,
    COALESCE(
        (xpath('@username', conn_node))[1]::text, ''
    )                                                   AS db_username

FROM all_conn
WHERE COALESCE((xpath('@class', conn_node))[1]::text, '') != ''
ORDER BY connection_type, site_name, project_name, workbook_name;


-- ================================================================
-- SECTION 7 · ASSESSMENT DASHBOARD
-- Overall counts for the migration assessment report.
-- Run last — gives the executive summary numbers.
-- ================================================================

WITH

workbooks_base AS (
    SELECT
        w.id, s.name AS site_name,
        w.name AS workbook_name,
        w.workbook::xml AS workbook_xml
    FROM workbooks w
    JOIN sites s ON s.id = w.site_id
    WHERE w.workbook IS NOT NULL
),

-- Flatten all non-hidden calculated fields across all workbooks
all_formulas AS (
    SELECT
        wb.id AS workbook_id,
        wb.site_name,
        (xpath('calculation/@formula', col_node))[1]::text AS formula
    FROM workbooks_base wb
    CROSS JOIN LATERAL
        unnest(xpath(
            '//datasource[not(@name="Parameters") and not(starts-with(@name,"_"))]'
            '/column[calculation and not(@hidden="true")]',
            wb.workbook_xml
        )) AS col_node
    WHERE (xpath('calculation/@formula', col_node))[1]::text IS NOT NULL
      AND (xpath('calculation/@formula', col_node))[1]::text != ''
)

SELECT
    -- Workbook totals
    (SELECT COUNT(*) FROM workbooks WHERE workbook IS NOT NULL)
                                                AS total_workbooks,
    (SELECT COUNT(*) FROM sites)                AS total_sites,
    (SELECT COUNT(*) FROM projects)             AS total_projects,

    -- Field totals
    COUNT(*)                                    AS total_calc_fields,
    COUNT(DISTINCT workbook_id)                 AS workbooks_with_calcs,

    -- LOD
    SUM(CASE WHEN formula ~ '\{(FIXED|INCLUDE|EXCLUDE)\b'
             THEN 1 ELSE 0 END)                AS total_lod_fields,
    COUNT(DISTINCT CASE
        WHEN formula ~ '\{FIXED\b'   THEN workbook_id END) AS wbs_with_fixed,
    COUNT(DISTINCT CASE
        WHEN formula ~ '\{INCLUDE\b' THEN workbook_id END) AS wbs_with_include,
    COUNT(DISTINCT CASE
        WHEN formula ~ '\{EXCLUDE\b' THEN workbook_id END) AS wbs_with_exclude,

    -- Table calcs
    SUM(CASE WHEN formula ~ '\m(RUNNING_SUM|RUNNING_AVG|WINDOW_SUM|WINDOW_AVG'
                             '|LOOKUP|INDEX|RANK|PREVIOUS_VALUE)\s*\('
             THEN 1 ELSE 0 END)                AS total_table_calc_fields,

    -- Conditional logic
    SUM(CASE WHEN formula ~ '\m(IF|CASE|IIF|ELSEIF)\M'
             THEN 1 ELSE 0 END)                AS total_conditional_fields,

    -- Date calculations
    SUM(CASE WHEN formula ~ '\mDATE(DIFF|PART|ADD|NAME|TRUNC)?\M'
             THEN 1 ELSE 0 END)                AS total_date_calc_fields,

    -- Complexity breakdown (workbook-level, approximate from formula counts)
    COUNT(DISTINCT CASE
        WHEN formula ~ '\{(FIXED|INCLUDE|EXCLUDE)\b'
          OR formula ~ '\m(RUNNING_SUM|WINDOW_SUM|LOOKUP|RANK)\s*\('
        THEN workbook_id END)                  AS workbooks_high_complexity,

    COUNT(DISTINCT CASE
        WHEN formula !~ '\{(FIXED|INCLUDE|EXCLUDE)\b'
         AND formula !~ '\m(RUNNING_SUM|WINDOW_SUM|LOOKUP|RANK)\s*\('
         AND (  formula ~ '\m(IF|CASE|IIF|ELSEIF)\M'
             OR formula ~ '\mDATE(DIFF|PART|ADD|NAME|TRUNC)?\M')
        THEN workbook_id END)                  AS workbooks_medium_complexity

FROM all_formulas;


-- Connection type summary (for datasource compatibility planning)
SELECT
    COALESCE(
        (xpath('@class', conn_node))[1]::text,
        'unknown'
    )                   AS connection_type,
    COUNT(*)            AS datasource_count

FROM workbooks w
JOIN sites s ON s.id = w.site_id
CROSS JOIN LATERAL
    unnest(xpath('//connection', w.workbook::xml)) AS conn_node
WHERE w.workbook IS NOT NULL
  AND (xpath('@class', conn_node))[1]::text IS NOT NULL
  AND (xpath('@class', conn_node))[1]::text != ''
GROUP BY connection_type
ORDER BY datasource_count DESC;


-- ================================================================
-- SECTION 8 · CREATE VIEW STATEMENTS
-- Save as reusable views for ongoing analysis or BI tool connection.
-- Run once; then query the views like regular tables.
-- ================================================================

-- View 1: All calculated fields (Section 2 as a view)
CREATE OR REPLACE VIEW vw_tableau_calc_fields AS
WITH
workbooks_base AS (
    SELECT w.id, s.name AS site_name,
           COALESCE(p.name, 'Default') AS project_name,
           w.name AS workbook_name,
           COALESCE(w.owner_name, '') AS owner_name,
           w.created_at::date AS created_date,
           w.updated_at::date AS updated_date,
           w.workbook::xml AS workbook_xml
    FROM workbooks w
    JOIN sites s ON s.id = w.site_id
    LEFT JOIN projects p ON p.id = w.project_id
    WHERE w.workbook IS NOT NULL
),
ds_nodes AS (
    SELECT wb.*, ds_node
    FROM workbooks_base wb
    CROSS JOIN LATERAL
        unnest(xpath('//datasource[not(@name="Parameters") and not(starts-with(@name,"_")) and @name!=""]',
                     wb.workbook_xml)) AS ds_node
),
col_nodes AS (
    SELECT dsn.site_name, dsn.project_name, dsn.workbook_name, dsn.owner_name,
           dsn.created_date, dsn.updated_date,
           COALESCE((xpath('@caption', dsn.ds_node))[1]::text,
                    (xpath('@name',    dsn.ds_node))[1]::text) AS datasource_caption,
           col_node
    FROM ds_nodes dsn
    CROSS JOIN LATERAL
        unnest(xpath('column[calculation and not(@hidden="true")]', dsn.ds_node)) AS col_node
)
SELECT
    site_name, project_name, workbook_name, owner_name, created_date, updated_date,
    datasource_caption AS datasource,
    COALESCE((xpath('@caption', col_node))[1]::text,
             TRIM(BOTH '[]' FROM (xpath('@name', col_node))[1]::text)) AS caption,
    (xpath('@name', col_node))[1]::text AS field_name,
    LOWER(COALESCE((xpath('@datatype', col_node))[1]::text, 'string')) AS datatype,
    LOWER(COALESCE((xpath('@role',     col_node))[1]::text, ''))       AS role,
    COALESCE((xpath('calculation/@formula', col_node))[1]::text, '')   AS formula,
    CASE WHEN COALESCE((xpath('calculation/@formula', col_node))[1]::text, '')
              ~ '\{(FIXED|INCLUDE|EXCLUDE)\b' THEN true ELSE false END AS has_lod,
    CASE WHEN COALESCE((xpath('calculation/@formula', col_node))[1]::text, '')
              ~ '\m(RUNNING_SUM|RUNNING_AVG|WINDOW_SUM|WINDOW_AVG|LOOKUP|INDEX|RANK)\s*\('
              THEN true ELSE false END                                  AS has_table_calc,
    CASE
        WHEN COALESCE((xpath('calculation/@formula', col_node))[1]::text, '')
             ~ '\{(FIXED|INCLUDE|EXCLUDE)\b'
          OR COALESCE((xpath('calculation/@formula', col_node))[1]::text, '')
             ~ '\m(RUNNING_SUM|RUNNING_AVG|WINDOW_SUM|WINDOW_AVG|LOOKUP|INDEX|RANK)\s*\('
            THEN 'HIGH'
        WHEN COALESCE((xpath('calculation/@formula', col_node))[1]::text, '')
             ~ '\m(IF|CASE|IIF|ELSEIF)\M'
          OR COALESCE((xpath('calculation/@formula', col_node))[1]::text, '')
             ~ '\mDATE(DIFF|PART|ADD|NAME|TRUNC)?\M'
            THEN 'MEDIUM'
        WHEN COALESCE((xpath('calculation/@formula', col_node))[1]::text, '') != ''
            THEN 'LOW'
        ELSE 'NONE'
    END AS migration_complexity
FROM col_nodes
WHERE COALESCE((xpath('calculation/@formula', col_node))[1]::text, '') != '';


-- View 2: LOD expressions only
CREATE OR REPLACE VIEW vw_tableau_lod_expressions AS
SELECT * FROM vw_tableau_calc_fields WHERE has_lod = true;


-- View 3: Table calculations only
CREATE OR REPLACE VIEW vw_tableau_table_calcs AS
SELECT * FROM vw_tableau_calc_fields WHERE has_table_calc = true;


-- View 4: Per-workbook complexity summary
CREATE OR REPLACE VIEW vw_tableau_workbook_complexity AS
SELECT
    site_name, project_name, workbook_name, owner_name,
    created_date, updated_date,
    COUNT(*)                                        AS total_calc_fields,
    SUM(CASE WHEN has_lod        THEN 1 ELSE 0 END) AS lod_fields,
    SUM(CASE WHEN has_table_calc THEN 1 ELSE 0 END) AS table_calc_fields,
    MAX(migration_complexity)                        AS worst_complexity,
    CASE
        WHEN SUM(CASE WHEN has_lod THEN 1 ELSE 0 END)        > 5
          OR SUM(CASE WHEN has_table_calc THEN 1 ELSE 0 END) > 5 THEN 'HIGH'
        WHEN SUM(CASE WHEN has_lod THEN 1 ELSE 0 END)        > 0
          OR SUM(CASE WHEN has_table_calc THEN 1 ELSE 0 END) > 0 THEN 'MEDIUM'
        ELSE 'LOW'
    END                                              AS migration_complexity
FROM vw_tableau_calc_fields
GROUP BY site_name, project_name, workbook_name, owner_name, created_date, updated_date
ORDER BY migration_complexity, site_name, project_name, workbook_name;


-- ================================================================
-- QUICK REFERENCE: Useful queries against the views
-- ================================================================

-- All HIGH complexity workbooks:
--   SELECT * FROM vw_tableau_workbook_complexity WHERE migration_complexity = 'HIGH';

-- All FIXED LOD expressions across all workbooks:
--   SELECT * FROM vw_tableau_lod_expressions WHERE formula ~ '\{FIXED\b' ORDER BY workbook_name;

-- Top 10 workbooks by number of calculated fields:
--   SELECT * FROM vw_tableau_workbook_complexity ORDER BY total_calc_fields DESC LIMIT 10;

-- Connection type breakdown:
--   SELECT connection_type, COUNT(*) AS n FROM vw_tableau_calc_fields
--   GROUP BY connection_type ORDER BY n DESC;


-- ================================================================
-- NOTE: HANDLING GZIP-COMPRESSED WORKBOOK XML (bytea storage)
-- ================================================================
-- If Section 0a shows data_type = 'bytea', the workbook XML is
-- stored gzip-compressed. PostgreSQL has no built-in gzip function.
-- Options:
--
--   A) Use the Python script (tableau_assessment_extractor.py) —
--      it handles both text and bytea automatically.
--
--   B) Install the pgcrypto extension (if available):
--        SELECT convert_from(
--                 pgp_sym_decrypt(workbook, ''),
--                 'UTF8'
--               ) FROM workbooks LIMIT 1;
--      (This works only if stored with pgcrypto, not standard gzip.)
--
--   C) Create a PL/Python function that calls Python's gzip module:
--        CREATE OR REPLACE FUNCTION gunzip(data bytea) RETURNS text
--        LANGUAGE plpython3u AS $$
--            import gzip
--            return gzip.decompress(bytes(data)).decode('utf-8', errors='replace')
--        $$;
--        -- Then replace w.workbook::xml with gunzip(w.workbook)::xml
--
--   D) SSH-tunnel to Tableau Server, export XML via tsm/tabcmd,
--      then run the queries on the exported files.
-- ================================================================
