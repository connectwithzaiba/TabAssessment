#!/usr/bin/env python3
"""
tableau_assessment_extractor.py
========================================================
Production script: Tableau PostgreSQL Repository → Power BI Assessment
Connects to Tableau Server's internal PostgreSQL repository,
extracts all workbook metadata (fields, formulas, LOD expressions,
data types, datasources) across all sites and projects, and writes:
  - CSV files  (workbooks / fields / lod_fields / datasources)
  - Multi-tab  Excel assessment workbook

Requirements:
  pip install psycopg2-binary pandas openpyxl

Usage:
  python tableau_assessment_extractor.py
  python tableau_assessment_extractor.py --config my_config.json
  TABLEAU_PG_PASSWORD=secret python tableau_assessment_extractor.py

PostgreSQL connection defaults (Tableau Server internal Postgres):
  Host:     localhost   (run on the Tableau Server machine, or SSH-tunnel)
  Port:     8060        (Tableau's dedicated Postgres port)
  Database: workgroup
  User:     readonly    (or tblwgadmin on older Tableau versions)

See config.example.json for all configuration options.
========================================================
"""

import argparse
import csv
import gzip
import json
import logging
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import asdict, dataclass
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Dict, Generator, List, Optional, Tuple

import psycopg2
import psycopg2.extras
from openpyxl import Workbook as XLWorkbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# DEFAULT CONFIGURATION
# Override via --config JSON file or environment variables (see ENV_MAP)
# ─────────────────────────────────────────────────────────────────────────────

DEFAULT_CONFIG: Dict = {
    "pg_host":             "localhost",
    "pg_port":             <<ur port>>,
    "pg_database":         "<<ur db name>>",
    "pg_user":             "readonly",
    "pg_password":         "",          # Prefer env var TABLEAU_PG_PASSWORD
    "pg_connect_timeout":  30,
    "pg_application_name": "tableau_assessment_extractor",
    "output_dir":          "./tableau_assessment_output",
    "batch_size":          50,
    "log_level":           "INFO",
    "log_file":            "tableau_extractor.log",
}

ENV_MAP = {
    "TABLEAU_PG_HOST":     "pg_host",
    "TABLEAU_PG_PORT":     "pg_port",
    "TABLEAU_PG_DATABASE": "pg_database",
    "TABLEAU_PG_USER":     "pg_user",
    "TABLEAU_PG_PASSWORD": "pg_password",
    "TABLEAU_OUTPUT_DIR":  "output_dir",
    "TABLEAU_LOG_LEVEL":   "log_level",
}

# Tableau datatype → Power BI / DAX equivalent
DTYPE_MAP = {
    "string":   "Text",
    "integer":  "Whole Number",
    "real":     "Decimal Number",
    "boolean":  "True/False",
    "date":     "Date",
    "datetime": "Date/Time",
    "spatial":  "Geography",
    "unknown":  "Any",
}

# Regex: LOD expressions
LOD_PATTERN = re.compile(
    r'\{(FIXED|INCLUDE|EXCLUDE)\b[^}]*\}',
    re.IGNORECASE | re.DOTALL
)

# Regex: Table calculation functions
TABLE_CALC_FN = re.compile(
    r'\b(RUNNING_SUM|RUNNING_AVG|RUNNING_COUNT|RUNNING_MAX|RUNNING_MIN'
    r'|WINDOW_SUM|WINDOW_AVG|WINDOW_COUNT|WINDOW_MAX|WINDOW_MIN'
    r'|LOOKUP|FIRST|LAST|INDEX|RANK|RANK_DENSE|RANK_MODIFIED'
    r'|RANK_PERCENTILE|RANK_UNIQUE|PREVIOUS_VALUE|SIZE|TOTAL)\s*\(',
    re.IGNORECASE
)

# DAX migration guidance per LOD type
LOD_DAX_NOTES = {
    "FIXED":   "Use CALCULATE() with ALL() or REMOVEFILTERS() to ignore filter context",
    "INCLUDE": "Use CALCULATE() with additional FILTER() to add granularity",
    "EXCLUDE": "Use CALCULATE() with REMOVEFILTERS(<column>) to drop specific filters",
}


# ─────────────────────────────────────────────────────────────────────────────
# DATA MODELS
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class WorkbookRecord:
    site_name:            str
    project_name:         str
    workbook_name:        str
    owner_name:           str
    repository_url:       str
    created_at:           str
    updated_at:           str
    total_fields:         int = 0
    calc_fields:          int = 0
    lod_fields:           int = 0
    table_calc_fields:    int = 0
    migration_complexity: str = "LOW"


@dataclass
class FieldRecord:
    site_name:            str
    project_name:         str
    workbook_name:        str
    datasource_name:      str
    field_name:           str
    caption:              str
    datatype:             str
    pbi_datatype:         str
    role:                 str
    is_calculated:        bool
    formula:              str
    has_lod:              bool
    lod_type:             str    # "FIXED" / "INCLUDE" / "EXCLUDE" / combined
    has_table_calc:       bool
    migration_complexity: str
    notes:                str = ""


@dataclass
class DatasourceRecord:
    site_name:          str
    project_name:       str
    workbook_name:      str
    datasource_name:    str
    datasource_caption: str
    connection_type:    str
    server:             str
    database_name:      str
    field_count:        int = 0
    calc_count:         int = 0


# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────

def setup_logging(cfg: Dict, output_dir: Path) -> logging.Logger:
    logger = logging.getLogger("tableau_extractor")
    logger.setLevel(cfg.get("log_level", "INFO").upper())

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    log_path = output_dir / cfg.get("log_file", "tableau_extractor.log")
    fh = RotatingFileHandler(
        log_path, maxBytes=10_000_000, backupCount=3, encoding="utf-8"
    )
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION LOADER
# ─────────────────────────────────────────────────────────────────────────────

def load_config(config_path: Optional[str] = None) -> Dict:
    cfg = DEFAULT_CONFIG.copy()

    if config_path and Path(config_path).exists():
        with open(config_path, encoding="utf-8") as f:
            cfg.update(json.load(f))
        print(f"Config loaded from: {config_path}")

    for env_key, cfg_key in ENV_MAP.items():
        val = os.environ.get(env_key)
        if val is not None:
            cfg[cfg_key] = int(val) if cfg_key == "pg_port" else val

    return cfg


# ─────────────────────────────────────────────────────────────────────────────
# POSTGRESQL CONNECTION
# ─────────────────────────────────────────────────────────────────────────────

def get_connection(cfg: Dict, logger: logging.Logger, retries: int = 3):
    """Connect to Tableau's PostgreSQL repo with retry logic."""
    dsn = dict(
        host=cfg["pg_host"],
        port=int(cfg["pg_port"]),
        dbname=cfg["pg_database"],
        user=cfg["pg_user"],
        password=cfg["pg_password"],
        connect_timeout=int(cfg.get("pg_connect_timeout", 30)),
        application_name=cfg.get("pg_application_name", "tableau_extractor"),
        options="-c default_transaction_read_only=on",   # Safety: read-only session
    )

    for attempt in range(1, retries + 1):
        try:
            conn = psycopg2.connect(**dsn)
            logger.info(
                "Connected to PostgreSQL at %s:%s / database: %s",
                cfg["pg_host"], cfg["pg_port"], cfg["pg_database"]
            )
            return conn
        except psycopg2.OperationalError as exc:
            logger.warning(
                "Connection attempt %d/%d failed: %s", attempt, retries, exc
            )
            if attempt < retries:
                wait = 5 * attempt
                logger.info("Retrying in %d seconds…", wait)
                time.sleep(wait)
            else:
                logger.error("All connection attempts exhausted. Check config and Postgres access.")
                raise


# ─────────────────────────────────────────────────────────────────────────────
# SCHEMA INTROSPECTION
# ─────────────────────────────────────────────────────────────────────────────

def detect_workbook_xml_column(cursor, logger: logging.Logger) -> str:
    """
    Detect the actual column that stores workbook XML.
    Varies slightly across Tableau Server versions.
    """
    cursor.execute("""
        SELECT column_name
        FROM   information_schema.columns
        WHERE  table_schema = 'public'
          AND  table_name   = 'workbooks'
          AND  column_name  IN ('workbook', 'workbook_xml', 'xml_data')
        ORDER BY
            CASE column_name
                WHEN 'workbook'     THEN 1
                WHEN 'workbook_xml' THEN 2
                WHEN 'xml_data'     THEN 3
            END
        LIMIT 1
    """)
    row = cursor.fetchone()
    col = row[0] if row else "workbook"
    logger.info("Workbook XML column: '%s'", col)
    return col


def detect_owner_column(cursor, logger: logging.Logger) -> str:
    """owner_name may be stored differently across Tableau versions."""
    cursor.execute("""
        SELECT column_name
        FROM   information_schema.columns
        WHERE  table_schema = 'public'
          AND  table_name   = 'workbooks'
          AND  column_name  IN ('owner_name', 'system_user_name', 'owner_id')
        LIMIT 1
    """)
    row = cursor.fetchone()
    col = row[0] if row else "owner_name"
    logger.debug("Owner column: '%s'", col)
    return col


# ─────────────────────────────────────────────────────────────────────────────
# XML DECODE
# ─────────────────────────────────────────────────────────────────────────────

def decode_xml(raw) -> Optional[str]:
    """
    Handle both raw XML text (str) and gzip-compressed bytea (bytes/memoryview)
    as stored in Tableau's PostgreSQL repository.
    """
    if raw is None:
        return None
    if isinstance(raw, memoryview):
        raw = bytes(raw)
    if isinstance(raw, bytes):
        # Try gzip decompression first
        try:
            return gzip.decompress(raw).decode("utf-8", errors="replace")
        except (OSError, Exception):
            # Plain text stored as bytes
            try:
                return raw.decode("utf-8", errors="replace")
            except Exception:
                return None
    return str(raw)


# ─────────────────────────────────────────────────────────────────────────────
# XML PARSING
# ─────────────────────────────────────────────────────────────────────────────

def parse_workbook_xml(
    xml_text: str,
    wb: WorkbookRecord,
    logger: logging.Logger
) -> Tuple[List[FieldRecord], List[DatasourceRecord]]:
    """
    Parse a Tableau workbook XML string.
    Returns extracted field records and datasource records.
    """
    fields: List[FieldRecord] = []
    datasources: List[DatasourceRecord] = []

    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        logger.warning("XML parse error in '%s/%s': %s",
                       wb.project_name, wb.workbook_name, exc)
        return fields, datasources

    ds_elements = (
        root.findall(".//datasource")
        if root.tag != "datasource"
        else [root]
    )

    for ds_el in ds_elements:
        ds_name    = ds_el.get("name", "")
        ds_caption = ds_el.get("caption", ds_name)

        # Skip Tableau internal / empty datasources
        if not ds_name or ds_name == "Parameters" or ds_name.startswith("_"):
            continue

        # --- Connection metadata ---
        conn_type = db_server = db_name = ""
        conn_el = ds_el.find(".//connection")
        if conn_el is not None:
            conn_type = conn_el.get("class", "")
            db_server = conn_el.get("server", "")
            db_name   = conn_el.get("dbname", "") or conn_el.get("database", "")

        # Federated datasources use named-connections
        if not conn_type:
            nc = ds_el.find(".//named-connections/named-connection/connection")
            if nc is not None:
                conn_type = nc.get("class", "")
                db_server = nc.get("server", "")
                db_name   = nc.get("dbname", "") or nc.get("database", "")

        ds_rec = DatasourceRecord(
            site_name=wb.site_name,
            project_name=wb.project_name,
            workbook_name=wb.workbook_name,
            datasource_name=ds_name,
            datasource_caption=ds_caption or ds_name,
            connection_type=conn_type,
            server=db_server,
            database_name=db_name,
        )

        # --- Column / field metadata ---
        for col_el in ds_el.findall("column"):
            frec = _extract_field(col_el, wb, ds_caption or ds_name)
            if frec is None:
                continue
            fields.append(frec)
            ds_rec.field_count += 1
            if frec.is_calculated:
                ds_rec.calc_count += 1

        # Also check column-instances (some Tableau versions use these)
        for col_el in ds_el.findall(".//column-instance/.."):
            pass   # column-instances reference columns; already captured above

        wb.total_fields += ds_rec.field_count
        wb.calc_fields  += ds_rec.calc_count
        datasources.append(ds_rec)

    # Roll-up LOD / table calc counts to workbook
    wb.lod_fields         = sum(1 for f in fields if f.has_lod)
    wb.table_calc_fields  = sum(1 for f in fields if f.has_table_calc)
    wb.migration_complexity = _score_workbook(wb)

    return fields, datasources


def _extract_field(
    col_el: ET.Element,
    wb: WorkbookRecord,
    ds_caption: str
) -> Optional[FieldRecord]:
    """Extract a single FieldRecord from a <column> XML element."""
    name     = col_el.get("name", "")
    caption  = col_el.get("caption", name.strip("[]"))
    datatype = col_el.get("datatype", "string").lower()
    role     = col_el.get("role", "").lower()

    # Skip hidden and synthetic fields
    if col_el.get("hidden") == "true":
        return None
    if not name or name.startswith("[Number of Records]"):
        return None

    # Calculated field?
    calc_el = col_el.find("calculation")
    formula = ""
    is_calc = False
    if calc_el is not None:
        formula = (calc_el.get("formula") or "").strip()
        is_calc = bool(formula)

    # LOD detection
    lod_matches = LOD_PATTERN.findall(formula) if formula else []
    has_lod  = bool(lod_matches)
    lod_type = ""
    if has_lod:
        unique_types = list(dict.fromkeys(m.upper() for m in lod_matches))
        lod_type = ", ".join(unique_types)

    # Table calc detection
    has_tc = bool(TABLE_CALC_FN.search(formula)) if formula else False

    return FieldRecord(
        site_name=wb.site_name,
        project_name=wb.project_name,
        workbook_name=wb.workbook_name,
        datasource_name=ds_caption,
        field_name=name,
        caption=caption,
        datatype=datatype,
        pbi_datatype=DTYPE_MAP.get(datatype, datatype),
        role=role,
        is_calculated=is_calc,
        formula=formula,
        has_lod=has_lod,
        lod_type=lod_type,
        has_table_calc=has_tc,
        migration_complexity=_score_field(formula, has_lod, has_tc),
    )


def _score_field(formula: str, has_lod: bool, has_tc: bool) -> str:
    if not formula:
        return "NONE"
    if has_lod or has_tc:
        return "HIGH"
    if re.search(r'\b(IF|CASE|IIF|ELSEIF)\b', formula, re.I):
        return "MEDIUM"
    if re.search(r'\bDATE(DIFF|PART|ADD|NAME|TRUNC)?\b', formula, re.I):
        return "MEDIUM"
    return "LOW"


def _score_workbook(wb: WorkbookRecord) -> str:
    if wb.lod_fields > 5 or wb.table_calc_fields > 5:
        return "HIGH"
    if wb.lod_fields > 0 or wb.table_calc_fields > 0 or wb.calc_fields > 20:
        return "MEDIUM"
    if wb.calc_fields > 0:
        return "LOW"
    return "LOW"


# ─────────────────────────────────────────────────────────────────────────────
# DATABASE QUERIES
# ─────────────────────────────────────────────────────────────────────────────

def _build_workbook_query(xml_col: str, owner_col: str) -> str:
    return f"""
        SELECT
            s.name                              AS site_name,
            COALESCE(p.name, 'Default')         AS project_name,
            w.name                              AS workbook_name,
            COALESCE(w.{owner_col}, '')         AS owner_name,
            COALESCE(w.repository_url, '')      AS repository_url,
            w.created_at,
            w.updated_at,
            w.{xml_col}                         AS workbook_xml
        FROM  workbooks  w
        JOIN  sites      s ON s.id = w.site_id
        LEFT JOIN projects p ON p.id = w.project_id
        WHERE w.{xml_col} IS NOT NULL
        ORDER BY s.name, p.name, w.name
        LIMIT  %(limit)s
        OFFSET %(offset)s
    """


def iter_workbooks(
    conn,
    xml_col: str,
    owner_col: str,
    batch_size: int,
    logger: logging.Logger
) -> Generator:
    """Batch-fetch all workbooks from the Tableau repository."""
    with conn.cursor() as cur:
        cur.execute(
            f"SELECT COUNT(*) FROM workbooks WHERE {xml_col} IS NOT NULL"
        )
        total = cur.fetchone()[0]

    logger.info("Total workbooks with XML: %d", total)
    query = _build_workbook_query(xml_col, owner_col)
    offset = 0

    while offset < total:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(query, {"limit": batch_size, "offset": offset})
            rows = cur.fetchall()

        if not rows:
            break

        logger.info(
            "  Batch %d–%d of %d",
            offset + 1, min(offset + len(rows), total), total
        )
        for row in rows:
            yield row

        offset += batch_size


# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT — CSV
# ─────────────────────────────────────────────────────────────────────────────

def write_csvs(
    output_dir:  Path,
    workbooks:   List[WorkbookRecord],
    fields:      List[FieldRecord],
    datasources: List[DatasourceRecord],
    logger:      logging.Logger,
    ts:          str
) -> List[Path]:
    specs = {
        f"workbooks_{ts}.csv": (workbooks, [
            "site_name", "project_name", "workbook_name", "owner_name",
            "created_at", "updated_at", "total_fields", "calc_fields",
            "lod_fields", "table_calc_fields", "migration_complexity",
        ]),
        f"fields_{ts}.csv": (fields, [
            "site_name", "project_name", "workbook_name", "datasource_name",
            "caption", "field_name", "datatype", "pbi_datatype", "role",
            "is_calculated", "formula", "has_lod", "lod_type",
            "has_table_calc", "migration_complexity", "notes",
        ]),
        f"lod_expressions_{ts}.csv": (
            [f for f in fields if f.has_lod],
            [
                "site_name", "project_name", "workbook_name", "datasource_name",
                "caption", "field_name", "lod_type", "formula", "migration_complexity",
            ]
        ),
        f"datasources_{ts}.csv": (datasources, [
            "site_name", "project_name", "workbook_name",
            "datasource_caption", "connection_type", "server",
            "database_name", "field_count", "calc_count",
        ]),
    }

    paths = []
    for filename, (records, columns) in specs.items():
        path = output_dir / filename
        with open(path, "w", newline="", encoding="utf-8") as fh:
            writer = csv.DictWriter(fh, fieldnames=columns, extrasaction="ignore")
            writer.writeheader()
            for rec in records:
                writer.writerow(asdict(rec))
        logger.info("CSV: %-45s  %d rows", filename, len(records))
        paths.append(path)

    return paths


# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT — EXCEL
# ─────────────────────────────────────────────────────────────────────────────

# ── Style constants ──────────────────────────────────────────────────────────
_NAVY   = "1F3864"
_WHITE  = "FFFFFF"
_LGREY  = "EEF2F7"
_AMBER  = "FFF8E1"
_RED_BG = "FDECEA"
_GREEN  = "E8F5E9"
_ORANGE = "FFF3E0"
_DBLUE  = "D6E4F0"

HEADER_FILL = PatternFill("solid", start_color=_NAVY)
ALT_FILL    = PatternFill("solid", start_color=_LGREY)
HIGH_FILL   = PatternFill("solid", start_color=_RED_BG)
MED_FILL    = PatternFill("solid", start_color=_AMBER)
LOW_FILL    = PatternFill("solid", start_color=_GREEN)
NONE_FILL   = None

HEADER_FONT  = Font(bold=True, color=_WHITE, name="Arial", size=10)
BODY_FONT    = Font(name="Arial", size=9)
SECTION_FONT = Font(bold=True, name="Arial", size=9, color=_NAVY)
TITLE_FONT   = Font(bold=True, size=14, name="Arial", color=_NAVY)

THIN = Side(style="thin", color="C0C0C0")
CELL_BORDER = Border(bottom=THIN, right=THIN)

COMPLEXITY_FILL = {
    "HIGH": HIGH_FILL, "MEDIUM": MED_FILL,
    "LOW": LOW_FILL,   "NONE": NONE_FILL,
}


def _auto_width(ws, min_w: int = 10, max_w: int = 65):
    for col in ws.columns:
        best = min_w
        for cell in col:
            try:
                best = max(best, len(str(cell.value or "")) + 2)
            except Exception:
                pass
        ws.column_dimensions[
            get_column_letter(col[0].column)
        ].width = min(best, max_w)


def _write_header(ws, headers: List[str], row: int = 1):
    ws.row_dimensions[row].height = 28
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.font      = HEADER_FONT
        c.fill      = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border    = CELL_BORDER


def _write_rows(
    ws,
    data:         List[List],
    start_row:    int,
    complexity_col: Optional[int] = None
):
    for ri, row_data in enumerate(data):
        rn   = start_row + ri
        base = ALT_FILL if ri % 2 else None

        # Determine row fill from complexity column
        row_fill = base
        if complexity_col:
            try:
                cval = str(row_data[complexity_col - 1])
                row_fill = COMPLEXITY_FILL.get(cval, base)
            except (IndexError, TypeError):
                pass

        for ci, val in enumerate(row_data, 1):
            c = ws.cell(row=rn, column=ci, value=val)
            c.font      = BODY_FONT
            c.border    = CELL_BORDER
            c.alignment = Alignment(vertical="center", wrap_text=False)
            if row_fill:
                c.fill = row_fill


def write_excel(
    output_dir:  Path,
    workbooks:   List[WorkbookRecord],
    fields:      List[FieldRecord],
    datasources: List[DatasourceRecord],
    logger:      logging.Logger,
    ts:          str
) -> Path:
    wb = XLWorkbook()
    wb.remove(wb.active)

    # ── Summary ───────────────────────────────────────────────────────────
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 18

    total_wb     = len(workbooks)
    total_fields = sum(w.total_fields       for w in workbooks)
    total_calc   = sum(w.calc_fields        for w in workbooks)
    total_lod    = sum(w.lod_fields         for w in workbooks)
    total_tc     = sum(w.table_calc_fields  for w in workbooks)
    high_wbs     = sum(1 for w in workbooks if w.migration_complexity == "HIGH")
    med_wbs      = sum(1 for w in workbooks if w.migration_complexity == "MEDIUM")
    low_wbs      = sum(1 for w in workbooks if w.migration_complexity == "LOW")

    lod_counter: Counter = Counter()
    for f in fields:
        if f.has_lod:
            for t in f.lod_type.split(","):
                lod_counter[t.strip()] += 1

    ds_types: Counter = Counter(d.connection_type for d in datasources if d.connection_type)

    def _sec(label, row):
        c = ws.cell(row=row, column=1, value=label)
        c.font = SECTION_FONT
        c.fill = PatternFill("solid", start_color=_DBLUE)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        c.alignment = Alignment(horizontal="left")

    def _row(label, value, row):
        ws.cell(row=row, column=1, value=label).font = BODY_FONT
        ws.cell(row=row, column=2, value=value).font = BODY_FONT

    ws["A1"] = "Tableau → Power BI Migration Assessment"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws["A2"].font = BODY_FONT

    r = 4
    _sec("WORKBOOKS & DATASOURCES", r);  r += 1
    _row("Total Workbooks",    total_wb,          r);  r += 1
    _row("Total Datasources",  len(datasources),  r);  r += 1
    r += 1
    _sec("FIELD COUNTS", r);             r += 1
    _row("Total Fields",            total_fields, r);  r += 1
    _row("Calculated Fields",       total_calc,   r);  r += 1
    _row("LOD Expression Fields",   total_lod,    r);  r += 1
    _row("Table Calc Fields",       total_tc,     r);  r += 1
    r += 1
    _sec("MIGRATION COMPLEXITY", r);     r += 1
    _row("HIGH (requires careful DAX re-engineering)",   high_wbs, r); r += 1
    _row("MEDIUM (conditional logic / date transforms)", med_wbs,  r); r += 1
    _row("LOW (simple aggregations)",                    low_wbs,  r); r += 1
    r += 1
    _sec("LOD TYPE BREAKDOWN", r);       r += 1
    for lod_type, cnt in sorted(lod_counter.items()):
        _row(f"  {lod_type}", cnt, r);   r += 1
    r += 1
    _sec("CONNECTION TYPES DETECTED", r); r += 1
    for conn_type, cnt in ds_types.most_common():
        _row(f"  {conn_type}", cnt, r);  r += 1

    # ── Workbooks ─────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Workbooks")
    ws2.sheet_view.showGridLines = False
    ws2.freeze_panes = "A2"

    _write_header(ws2, [
        "Site", "Project", "Workbook Name", "Owner",
        "Total Fields", "Calc Fields", "LOD Fields", "Table Calc Fields",
        "Migration Complexity", "Created", "Updated"
    ])
    _write_rows(ws2, [[
        w.site_name, w.project_name, w.workbook_name, w.owner_name,
        w.total_fields, w.calc_fields, w.lod_fields, w.table_calc_fields,
        w.migration_complexity,
        str(w.created_at)[:10], str(w.updated_at)[:10]
    ] for w in workbooks], 2, complexity_col=9)
    _auto_width(ws2)

    # ── Calculated Fields ─────────────────────────────────────────────────
    ws3 = wb.create_sheet("Calculated Fields")
    ws3.sheet_view.showGridLines = False
    ws3.freeze_panes = "A2"

    _write_header(ws3, [
        "Site", "Project", "Workbook", "Datasource",
        "Caption", "Field Name", "Tableau Type", "PBI Type",
        "Role", "Formula", "Has LOD", "LOD Type",
        "Has Table Calc", "Migration Complexity"
    ])
    _write_rows(ws3, [[
        f.site_name, f.project_name, f.workbook_name, f.datasource_name,
        f.caption, f.field_name, f.datatype, f.pbi_datatype,
        f.role, f.formula,
        "Yes" if f.has_lod       else "No",
        f.lod_type,
        "Yes" if f.has_table_calc else "No",
        f.migration_complexity
    ] for f in fields if f.is_calculated], 2, complexity_col=14)
    _auto_width(ws3)

    # ── LOD Expressions ───────────────────────────────────────────────────
    ws4 = wb.create_sheet("LOD Expressions")
    ws4.sheet_view.showGridLines = False
    ws4.freeze_panes = "A2"

    _write_header(ws4, [
        "Site", "Project", "Workbook", "Datasource",
        "Caption", "Field Name", "LOD Type", "Formula",
        "DAX Migration Guidance"
    ])
    lod_rows = []
    for f in fields:
        if not f.has_lod:
            continue
        guidance = " | ".join(
            LOD_DAX_NOTES.get(t.strip(), "")
            for t in f.lod_type.split(",") if t.strip()
        )
        lod_rows.append([
            f.site_name, f.project_name, f.workbook_name, f.datasource_name,
            f.caption, f.field_name, f.lod_type, f.formula, guidance
        ])
    _write_rows(ws4, lod_rows, 2)

    # Highlight formula cells in LOD sheet
    for rn in range(2, len(lod_rows) + 2):
        ws4.cell(row=rn, column=8).fill = PatternFill(
            "solid", start_color=_ORANGE
        )

    _auto_width(ws4)

    # ── All Fields ────────────────────────────────────────────────────────
    ws5 = wb.create_sheet("All Fields")
    ws5.sheet_view.showGridLines = False
    ws5.freeze_panes = "A2"

    _write_header(ws5, [
        "Site", "Project", "Workbook", "Datasource",
        "Caption", "Field Name", "Tableau Type", "PBI Type",
        "Role", "Is Calculated", "Has LOD", "Has Table Calc",
        "Migration Complexity"
    ])
    _write_rows(ws5, [[
        f.site_name, f.project_name, f.workbook_name, f.datasource_name,
        f.caption, f.field_name, f.datatype, f.pbi_datatype,
        f.role,
        "Yes" if f.is_calculated  else "No",
        "Yes" if f.has_lod        else "No",
        "Yes" if f.has_table_calc else "No",
        f.migration_complexity
    ] for f in fields], 2, complexity_col=13)
    _auto_width(ws5)

    # ── Datasources ───────────────────────────────────────────────────────
    ws6 = wb.create_sheet("Datasources")
    ws6.sheet_view.showGridLines = False
    ws6.freeze_panes = "A2"

    _write_header(ws6, [
        "Site", "Project", "Workbook", "Datasource Name",
        "Connection Type", "Server", "Database",
        "Total Fields", "Calc Fields"
    ])
    _write_rows(ws6, [[
        d.site_name, d.project_name, d.workbook_name, d.datasource_caption,
        d.connection_type, d.server, d.database_name,
        d.field_count, d.calc_count
    ] for d in datasources], 2)
    _auto_width(ws6)

    out_path = output_dir / f"tableau_assessment_{ts}.xlsx"
    wb.save(out_path)
    logger.info("Excel report: %s", out_path.name)
    return out_path


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Tableau PostgreSQL Repository → Power BI Assessment Extractor",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument(
        "--config", metavar="CONFIG_JSON",
        help="Path to JSON config file (optional). See config.example.json."
    )
    args = parser.parse_args()

    cfg = load_config(args.config)

    output_dir = Path(cfg["output_dir"])
    output_dir.mkdir(parents=True, exist_ok=True)

    logger = setup_logging(cfg, output_dir)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    logger.info("=" * 65)
    logger.info("Tableau Assessment Extractor  |  %s", ts)
    logger.info("Output: %s", output_dir.resolve())
    logger.info("=" * 65)

    conn = get_connection(cfg, logger)

    try:
        with conn.cursor() as cur:
            xml_col   = detect_workbook_xml_column(cur, logger)
            owner_col = detect_owner_column(cur, logger)

        all_workbooks:   List[WorkbookRecord]   = []
        all_fields:      List[FieldRecord]      = []
        all_datasources: List[DatasourceRecord] = []
        parse_errors = 0

        for row in iter_workbooks(
            conn, xml_col, owner_col, int(cfg["batch_size"]), logger
        ):
            wb_rec = WorkbookRecord(
                site_name=row["site_name"],
                project_name=row["project_name"],
                workbook_name=row["workbook_name"],
                owner_name=row["owner_name"],
                repository_url=row["repository_url"],
                created_at=str(row["created_at"] or ""),
                updated_at=str(row["updated_at"] or ""),
            )

            xml_text = decode_xml(row["workbook_xml"])
            if not xml_text:
                logger.debug("Empty XML — skipping '%s'", wb_rec.workbook_name)
                parse_errors += 1
                all_workbooks.append(wb_rec)
                continue

            wb_fields, wb_ds = parse_workbook_xml(xml_text, wb_rec, logger)
            all_workbooks.append(wb_rec)
            all_fields.extend(wb_fields)
            all_datasources.extend(wb_ds)

        # ── Summary ───────────────────────────────────────────────────────
        logger.info("─" * 65)
        logger.info("Extraction complete:")
        logger.info("  Workbooks processed : %d", len(all_workbooks))
        logger.info("  Fields extracted    : %d", len(all_fields))
        logger.info("  Datasources found   : %d", len(all_datasources))
        logger.info("  LOD expression flds : %d",
                    sum(1 for f in all_fields if f.has_lod))
        logger.info("  Table calc fields   : %d",
                    sum(1 for f in all_fields if f.has_table_calc))
        logger.info("  Parse errors        : %d", parse_errors)

        # HIGH / MEDIUM / LOW workbook counts
        for label in ("HIGH", "MEDIUM", "LOW"):
            cnt = sum(1 for w in all_workbooks if w.migration_complexity == label)
            logger.info("  %s complexity wbs  : %d", label.ljust(6), cnt)

        # ── Write outputs ─────────────────────────────────────────────────
        csv_paths = write_csvs(
            output_dir, all_workbooks, all_fields, all_datasources, logger, ts
        )
        xlsx_path = write_excel(
            output_dir, all_workbooks, all_fields, all_datasources, logger, ts
        )

        logger.info("─" * 65)
        logger.info("All outputs saved:")
        logger.info("  Excel  → %s", xlsx_path)
        for p in csv_paths:
            logger.info("  CSV    → %s", p)

    finally:
        conn.close()
        logger.info("PostgreSQL connection closed. Done.")


if __name__ == "__main__":
    main()
