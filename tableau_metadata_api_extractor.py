#!/usr/bin/env python3
"""
tableau_metadata_api_extractor.py
========================================================
Production script: Tableau Metadata API (GraphQL) + REST API → Power BI Assessment
Compatible with Tableau Server 2022.x and 2023.x.

WHAT IT EXTRACTS:
  - Calculated fields (formulas, LOD, table calcs, DAX guidance)
  - Workbooks (complete inventory via GQL_WORKBOOKS + enriched with usage/owner)
  - Embedded datasources (connection types)
  - Custom SQL blocks (structural analysis)
  - Parameters (per-workbook Tableau parameters)
  - Published datasources (shared datasource map + downstream workbook count)
  - Tableau Prep flows (for dataflow migration planning)
  - Owner activity (REST API: lastLogin, active/inactive flag)
  - Usage stats (REST API: totalViewCount, never-viewed flag)

HOW TO RUN:
  pip install requests openpyxl
  cp config_metadata.example.json config_metadata.json  # fill in your values
  python tableau_metadata_api_extractor.py --config config_metadata.json

  # Incremental re-run (only workbooks updated in last 7 days):
  python tableau_metadata_api_extractor.py --config config_metadata.json --since 2025-04-01

  # Password via env var (recommended):
  TABLEAU_PASSWORD=secret python tableau_metadata_api_extractor.py --config config_metadata.json

CLAUDE.md: See CLAUDE.md in this directory for architecture notes and known gaps.
========================================================
"""

import argparse
import concurrent.futures
import csv
import json
import logging
import getpass
import os
import re
import textwrap
import sys
import threading
import time
from collections import Counter
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Dict, Generator, List, Optional, Tuple

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from openpyxl import Workbook as XLWorkbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# Module-level logger — configured properly by setup_logging() in main()
logger = logging.getLogger('tableau_extractor')

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

DEFAULT_CONFIG: Dict = {
    "server_url":            "https://your-tableau-server",
    "api_version":           "3.21",          # Tableau 2023.3 = 3.21
    "username":              "",
    "password":              "",              # Prefer TABLEAU_PASSWORD env var
    "site_content_url":      "",              # "" = Default site; "all" = every site
    "verify_ssl":            True,
    "page_size":             100,             # GraphQL max = 100
    "output_dir":            "./tableau_assessment_output",
    "log_level":             "INFO",
    "log_file":              "tableau_metadata_extractor.log",
    "request_timeout":       60,
    "max_retries":           3,
    "parallel_sites":        4,              # Max concurrent site threads
    "inactive_owner_days":   90,
    "unused_view_threshold": 0,
}

ENV_MAP = {
    "TABLEAU_SERVER_URL":   "server_url",
    "TABLEAU_API_VERSION":  "api_version",
    "TABLEAU_USERNAME":     "username",
    "TABLEAU_PASSWORD":     "password",
    "TABLEAU_SITE":         "site_content_url",
    "TABLEAU_OUTPUT_DIR":   "output_dir",
    "TABLEAU_LOG_LEVEL":    "log_level",
    "TABLEAU_VERIFY_SSL":   "verify_ssl",
}

DTYPE_MAP = {
    "STRING":   "Text",
    "INTEGER":  "Whole Number",
    "REAL":     "Decimal Number",
    "BOOLEAN":  "True/False",
    "DATE":     "Date",
    "DATETIME": "Date/Time",
    "SPATIAL":  "Geography",
    "TABLE":    "Table",
}

LOD_PATTERN    = re.compile(r'\{(FIXED|INCLUDE|EXCLUDE)\b[^}]*\}', re.I | re.S)
TABLE_CALC_PAT = re.compile(
    r'\b(RUNNING_SUM|RUNNING_AVG|RUNNING_COUNT|RUNNING_MAX|RUNNING_MIN'
    r'|WINDOW_SUM|WINDOW_AVG|WINDOW_COUNT|WINDOW_MAX|WINDOW_MIN'
    r'|LOOKUP|FIRST|LAST|INDEX|SIZE|TOTAL'
    r'|RANK|RANK_DENSE|RANK_MODIFIED|RANK_PERCENTILE|RANK_UNIQUE'
    r'|PREVIOUS_VALUE)\s*\(',
    re.I
)
WINDOW_FUNC_PAT = re.compile(
    r'\b(ROW_NUMBER|RANK|DENSE_RANK|NTILE|LAG|LEAD|'
    r'FIRST_VALUE|LAST_VALUE|NTH_VALUE)\s*\('
    r'|(?:SUM|AVG|COUNT|MIN|MAX)\s*\([^)]*\)\s*OVER\s*\(',
    re.I | re.S
)

LOD_DAX = {
    "FIXED":   "CALCULATE(<expr>, ALL(<table>)) — removes all filter context",
    "INCLUDE": "CALCULATE(<expr>, FILTER(...)) — adds granularity via virtual table",
    "EXCLUDE": "CALCULATE(<expr>, REMOVEFILTERS(<col>)) — drops specific column filter",
}

PARAM_PBI_MAP = {
    "string":  "Power BI Text parameter (Manage Parameters > New)",
    "integer": "Power BI Whole Number parameter",
    "real":    "Power BI Decimal Number parameter",
    "boolean": "What-If parameter or report-level slicer",
    "date":    "Power BI Date parameter or date slicer",
    "datetime":"Power BI Date/Time parameter",
}

# ─────────────────────────────────────────────────────────────────────────────
# DATA MODELS
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class FieldRecord:
    site_name:            str
    project_name:         str
    workbook_name:        str
    workbook_owner:       str
    datasource_name:      str
    field_name:           str
    datatype:             str
    pbi_datatype:         str
    role:                 str
    formula:              str
    has_lod:              bool
    lod_types:            str
    has_table_calc:       bool
    table_calc_functions: str
    migration_complexity: str
    dax_guidance:         str = ""


@dataclass
class WorkbookRecord:
    site_name:            str
    project_name:         str
    workbook_name:        str
    owner_name:           str
    created_at:           str
    updated_at:           str
    calc_fields:          int  = 0
    lod_fields:           int  = 0
    table_calc_fields:    int  = 0
    sheet_count:          int  = 0
    dashboard_count:      int  = 0
    migration_complexity: str  = "LOW"
    complexity_score:     int  = 0
    # Usage & ownership (populated by REST API)
    total_views:          int  = 0
    never_viewed:         bool = False
    unused:               bool = False
    owner_is_active:      str  = "Unknown"   # "Yes" / "No" / "Unknown"
    owner_last_login:     str  = "Unknown"


@dataclass
class DatasourceRecord:
    site_name:        str
    project_name:     str
    workbook_name:    str
    datasource_name:  str
    connection_type:  str
    is_published:     bool
    field_count:      int = 0
    calc_count:       int = 0


@dataclass
class CustomSQLRecord:
    site_name:        str
    project_name:     str
    workbook_name:    str
    datasource_name:  str
    sql_name:         str
    connection_type:  str
    sql_query:        str
    line_count:       int
    has_subquery:     bool
    has_joins:        bool
    has_unions:       bool
    has_where:        bool
    has_group_by:     bool
    has_window_func:  bool
    migration_notes:  str


@dataclass
class OwnerRecord:
    site_name:        str
    username:         str
    display_name:     str
    email:            str
    site_role:        str
    last_login:       str
    days_since_login: int    # -1 = never
    is_active:        bool
    owned_workbooks:  int  = 0
    owned_high:       int  = 0
    owned_medium:     int  = 0


@dataclass
class ParameterRecord:
    site_name:          str
    project_name:       str
    workbook_name:      str
    parameter_name:     str
    data_type:          str
    default_value:      str
    allowable_values:   str   # ALL | LIST | RANGE
    pbi_equivalent:     str


@dataclass
class PublishedDatasourceRecord:
    site_name:             str
    project_name:          str
    datasource_name:       str
    owner_name:            str
    connection_type:       str
    has_extracts:          bool
    downstream_workbooks:  int
    created_at:            str
    updated_at:            str
    migration_note:        str


@dataclass
class FlowRecord:
    site_name:      str
    project_name:   str
    flow_name:      str
    owner_name:     str
    created_at:     str
    updated_at:     str
    output_steps:   int
    migration_note: str


# ─────────────────────────────────────────────────────────────────────────────
# GRAPHQL QUERIES
# ─────────────────────────────────────────────────────────────────────────────

GQL_WORKBOOKS = """
query GetWorkbooks($first: Int, $after: String, $filter: WorkbookFilter) {
  workbooksConnection(first: $first, after: $after, filter: $filter) {
    pageInfo { hasNextPage endCursor }
    nodes {
      id name projectName createdAt updatedAt
      owner { name }
      site { name }
      sheetsConnection    { totalCount }
      dashboardsConnection { totalCount }
      embeddedDatasourcesConnection {
        nodes {
          name
          upstreamDatasourcesConnection { nodes { name projectName } }
          upstreamTablesConnection      { nodes { connectionType    } }
        }
      }
    }
  }
}
"""

GQL_CALCULATED_FIELDS = """
query GetCalculatedFields($first: Int, $after: String) {
  calculatedFieldsConnection(first: $first, after: $after) {
    pageInfo { hasNextPage endCursor }
    nodes {
      id name formula dataType role isHidden description
      embeddedInDatasource {
        id name
        workbook {
          id name projectName createdAt updatedAt
          owner { name }
          site  { name }
        }
      }
    }
  }
}
"""

GQL_DATASOURCES = """
query GetDatasources($first: Int, $after: String) {
  embeddedDatasourcesConnection(first: $first, after: $after) {
    pageInfo { hasNextPage endCursor }
    nodes {
      id name
      workbook { name projectName site { name } }
      upstreamTablesConnection {
        nodes { connectionType database { connectionType } }
      }
      upstreamDatasourcesConnection { nodes { name connectionType } }
    }
  }
}
"""

GQL_CUSTOM_SQL = """
query GetCustomSQL($first: Int, $after: String) {
  customSQLTablesConnection(first: $first, after: $after) {
    pageInfo { hasNextPage endCursor }
    nodes {
      id name query connectionType
      downstreamDatasourcesConnection {
        nodes {
          name
          workbook { name projectName site { name } }
        }
      }
    }
  }
}
"""

GQL_PARAMETERS = """
query GetParameters($first: Int, $after: String) {
  parametersConnection(first: $first, after: $after) {
    pageInfo { hasNextPage endCursor }
    nodes {
      id name dataType
      defaultValue
      allowableValuesType
      workbook { name projectName site { name } }
    }
  }
}
"""

GQL_PUBLISHED_DATASOURCES = """
query GetPublishedDatasources($first: Int, $after: String) {
  publishedDatasourcesConnection(first: $first, after: $after) {
    pageInfo { hasNextPage endCursor }
    nodes {
      id name projectName createdAt updatedAt
      owner { name }
      site  { name }
      hasExtracts
      downstreamWorkbooksConnection {
        totalCount
        nodes { name projectName }
      }
      upstreamTablesConnection {
        nodes { connectionType database { connectionType } }
      }
    }
  }
}
"""

GQL_FLOWS = """
query GetFlows($first: Int, $after: String) {
  flowsConnection(first: $first, after: $after) {
    pageInfo { hasNextPage endCursor }
    nodes {
      id name projectName createdAt updatedAt
      owner { name }
      site  { name }
      outputSteps { name type }
    }
  }
}
"""


# ─────────────────────────────────────────────────────────────────────────────
# LOGGING + CONFIG
# ─────────────────────────────────────────────────────────────────────────────

def setup_logging(cfg: Dict, output_dir: Path) -> logging.Logger:
    level = getattr(logging, cfg.get("log_level", "INFO").upper(), logging.INFO)
    logger = logging.getLogger("tableau_extractor")
    logger.setLevel(level)
    if logger.handlers:
        return logger
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(ch)
    log_path = output_dir / cfg.get("log_file", "tableau_metadata_extractor.log")
    fh = RotatingFileHandler(log_path, maxBytes=10 * 1024 * 1024, backupCount=3)
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    return logger


def load_config(config_path: Optional[str] = None) -> Dict:
    cfg = dict(DEFAULT_CONFIG)
    if config_path:
        p = Path(config_path)
        if not p.exists():
            print(f"ERROR: Config file not found: {config_path}", file=sys.stderr)
            sys.exit(1)
        with open(p) as f:
            loaded = json.load(f)
        # Strip comment keys
        cfg.update({k: v for k, v in loaded.items() if not k.startswith("_")})
    for env_key, cfg_key in ENV_MAP.items():
        if env_key in os.environ:
            val = os.environ[env_key]
            if cfg_key == "verify_ssl":
                val = val.lower() not in ("false", "0", "no")
            cfg[cfg_key] = val
    return cfg


# ─────────────────────────────────────────────────────────────────────────────
# HTTP SESSION
# ─────────────────────────────────────────────────────────────────────────────

def build_session(cfg: Dict) -> requests.Session:
    session = requests.Session()
    session.verify = cfg.get("verify_ssl", True)
    retry = Retry(
        total=int(cfg.get("max_retries", 3)),
        backoff_factor=1.5,
        status_forcelist={429, 500, 502, 503, 504},
        allowed_methods={"GET", "POST"},
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://",  adapter)
    return session


# ─────────────────────────────────────────────────────────────────────────────
# AUTH
# ─────────────────────────────────────────────────────────────────────────────

def signin(
    session: requests.Session,
    cfg: Dict,
    logger: logging.Logger,
    site_content_url: str = "",
) -> Tuple[str, str]:
    """Returns (auth_token, site_id)."""
    url = f"{cfg['server_url'].rstrip('/')}/api/{cfg['api_version']}/auth/signin"
    payload = {"credentials": {
        "name": cfg["username"], "password": cfg["password"],
        "site": {"contentUrl": site_content_url},
    }}
    resp = session.post(url, json=payload,
                        headers={"Content-Type": "application/json", "Accept": "application/json"},
                        timeout=int(cfg.get("request_timeout", 60)))
    resp.raise_for_status()
    data      = resp.json()["credentials"]
    token     = data["token"]
    site_id   = data["site"]["id"]
    site_disp = data["site"].get("contentUrl", site_content_url) or "Default"
    logger.info("Signed in → site: '%s'  (id: %s)", site_disp, site_id)
    return token, site_id


def signout(session: requests.Session, cfg: Dict, token: str, logger: logging.Logger):
    url = f"{cfg['server_url'].rstrip('/')}/api/{cfg['api_version']}/auth/signout"
    try:
        session.post(url, headers={"x-tableau-auth": token}, timeout=30)
    except Exception:
        pass
    logger.debug("Signed out.")


def get_all_sites(
    session: requests.Session, cfg: Dict, token: str, logger: logging.Logger
) -> List[Dict]:
    """
    Return ALL sites on the server, paginating through them automatically.

    Tableau's /sites endpoint is paginated exactly like other REST list endpoints.
    A single un-paginated call only returns the first page (server default is
    often 100, but can be lower), which would silently miss sites on servers
    with 90+ sites. This function walks all pages and returns the complete list.

    Each site dict contains at minimum:
      id          — internal LUID  (used for scoped API calls)
      name        — human-readable display name  ← used as the "Site" column value
      contentUrl  — URL slug (empty string for the Default site)
    """
    server    = cfg["server_url"].rstrip("/")
    api       = cfg["api_version"]
    page_size = int(cfg.get("page_size", 100))
    timeout   = int(cfg.get("request_timeout", 60))
    headers   = {"x-tableau-auth": token, "Accept": "application/json"}

    all_sites: List[Dict] = []
    page = 1

    while True:
        resp = session.get(
            f"{server}/api/{api}/sites",
            headers=headers,
            params={"pageSize": page_size, "pageNumber": page},
            timeout=timeout,
        )
        resp.raise_for_status()
        data       = resp.json()
        pagination = data.get("pagination", {})
        total      = int(pagination.get("totalAvailable", 0))
        raw        = data.get("sites", {}).get("site", [])
        if isinstance(raw, dict):          # single-site server wraps as object
            raw = [raw]

        all_sites.extend(raw)
        logger.debug("  Sites page %d: got %d (total available: %d)",
                     page, len(raw), total)

        if len(all_sites) >= total or not raw:
            break
        page += 1

    logger.info(
        "Found %d site(s): %s%s",
        len(all_sites),
        [s.get("name", s.get("contentUrl", "Default")) for s in all_sites[:10]],
        " …" if len(all_sites) > 10 else "",
    )
    return all_sites


# ─────────────────────────────────────────────────────────────────────────────
# REST API HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _rest_paginate(
    session:      requests.Session,
    cfg:          Dict,
    token:        str,
    endpoint:     str,
    outer_key:    str,
    inner_key:    str,
    extra_params: Dict = None,
    logger:       logging.Logger = None,
) -> List[Dict]:
    """Generic paginator for Tableau REST API list endpoints with retry backoff."""
    server    = cfg["server_url"].rstrip("/")
    api       = cfg["api_version"]
    page_size = int(cfg.get("page_size", 100))
    timeout   = int(cfg.get("request_timeout", 60))
    max_ret   = int(cfg.get("max_retries", 3))
    headers   = {"x-tableau-auth": token, "Accept": "application/json"}
    items: List[Dict] = []
    page = 1

    while True:
        params = {"pageSize": page_size, "pageNumber": page}
        if extra_params:
            params.update(extra_params)
        url = f"{server}/api/{api}/{endpoint}"

        for attempt in range(max_ret):
            try:
                resp = session.get(url, headers=headers, params=params, timeout=timeout)
                resp.raise_for_status()
                break
            except requests.RequestException as exc:
                if attempt == max_ret - 1:
                    raise
                wait = 2 ** attempt
                if logger:
                    logger.warning("  REST retry %d/%d for %s (%s) — waiting %ds",
                                   attempt + 1, max_ret, endpoint, exc, wait)
                time.sleep(wait)

        data       = resp.json()
        pagination = data.get("pagination", {})
        total      = int(pagination.get("totalAvailable", 0))
        raw        = data.get(outer_key, {}).get(inner_key, [])
        if isinstance(raw, dict):
            raw = [raw]

        items.extend(raw)
        if len(items) >= total or not raw:
            break
        page += 1

    return items


def fetch_users_rest(
    session: requests.Session, cfg: Dict, token: str,
    site_id: str, logger: logging.Logger,
) -> Dict[str, Dict]:
    users = _rest_paginate(session, cfg, token,
                           f"sites/{site_id}/users", "users", "user", logger=logger)
    result = {u["id"]: u for u in users}
    logger.info("  REST: %d user(s) fetched", len(result))
    return result


def fetch_workbooks_rest(
    session: requests.Session, cfg: Dict, token: str,
    site_id: str, logger: logging.Logger,
) -> Dict:
    """Returns {wb_id: wb_dict} plus a '_by_name' lookup keyed by (lower_name, lower_proj)."""
    wbs = _rest_paginate(session, cfg, token,
                         f"sites/{site_id}/workbooks", "workbooks", "workbook", logger=logger)
    by_id:   Dict[str, Dict]   = {}
    by_name: Dict[tuple, str]  = {}
    for w in wbs:
        by_id[w["id"]] = w
        proj = (w.get("project") or {}).get("name", "") or ""
        by_name[(w.get("name", "").lower(), proj.lower())] = w["id"]
    result = dict(by_id)
    result["_by_name"] = by_name  # type: ignore[assignment]
    logger.info("  REST: %d workbook(s) fetched", len(by_id))
    return result


def fetch_view_counts_rest(
    session: requests.Session, cfg: Dict, token: str,
    site_id: str, logger: logging.Logger,
) -> Dict[str, int]:
    """Returns {workbook_id: total_view_count} aggregated from all views."""
    views = _rest_paginate(session, cfg, token,
                           f"sites/{site_id}/views", "views", "view",
                           extra_params={"includeUsageStatistics": "true"},
                           logger=logger)
    wb_views: Dict[str, int] = {}
    for v in views:
        wb_id = (v.get("workbook") or {}).get("id", "")
        count = int((v.get("usage") or {}).get("totalViewCount", 0))
        if wb_id:
            wb_views[wb_id] = wb_views.get(wb_id, 0) + count
    logger.info("  REST: view stats for %d workbook(s) fetched", len(wb_views))
    return wb_views


# ─────────────────────────────────────────────────────────────────────────────
# GRAPHQL CLIENT
# ─────────────────────────────────────────────────────────────────────────────

def gql_query(
    session: requests.Session, cfg: Dict, token: str,
    query: str, variables: Dict, logger: logging.Logger,
) -> Dict:
    url = f"{cfg['server_url'].rstrip('/')}/api/metadata/graphql"
    headers = {"x-tableau-auth": token, "content-type": "application/json",
               "accept": "application/json"}
    resp = session.post(url, json={"query": query, "variables": variables},
                        headers=headers, timeout=int(cfg.get("request_timeout", 60)))
    resp.raise_for_status()
    result = resp.json()
    if "errors" in result:
        for err in result["errors"]:
            logger.error("GraphQL error: %s", err.get("message", err))
    return result.get("data", {})


def paginate(
    session: requests.Session, cfg: Dict, token: str,
    query: str, connection_key: str, logger: logging.Logger,
    extra_vars: Dict = None,
) -> Generator[Dict, None, None]:
    """Cursor-based GraphQL paginator. Yields each node dict."""
    page_size = int(cfg.get("page_size", 100))
    cursor    = None
    total     = 0

    while True:
        variables = {"first": page_size, "after": cursor}
        if extra_vars:
            variables.update(extra_vars)
        data = gql_query(session, cfg, token, query, variables, logger)
        conn = data.get(connection_key, {})
        if not conn:
            break
        page_info = conn.get("pageInfo", {})
        nodes     = conn.get("nodes", [])
        total    += len(nodes)
        for node in nodes:
            yield node
        if not page_info.get("hasNextPage"):
            break
        cursor = page_info.get("endCursor")
        logger.debug("  %s — fetched %d so far…", connection_key, total)


def safe_paginate(
    session: requests.Session, cfg: Dict, token: str,
    query: str, connection_key: str, logger: logging.Logger,
    extra_vars: Dict = None,
) -> List[Dict]:
    """Wrapper around paginate() that returns empty list on any API error."""
    try:
        return list(paginate(session, cfg, token, query, connection_key, logger, extra_vars))
    except Exception as exc:
        logger.warning("  GraphQL query '%s' failed: %s — skipping.", connection_key, exc)
        return []


# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def _days_since(iso_dt: str) -> int:
    """Return whole days since an ISO datetime string; -1 if empty/unparseable."""
    if not iso_dt:
        return -1
    try:
        dt  = datetime.fromisoformat(iso_dt.replace("Z", "+00:00"))
        now = datetime.now(tz=timezone.utc)
        return max((now - dt).days, 0)
    except Exception:
        return -1


def analyse_formula(formula: str) -> Dict:
    """Detect LOD, table calcs, conditional/date logic; return analysis dict."""
    if not formula:
        return {
            "has_lod": False, "lod_types": "",
            "has_table_calc": False, "table_calc_functions": "",
            "migration_complexity": "NONE", "dax_guidance": "",
        }

    lod_matches = LOD_PATTERN.findall(formula)
    has_lod     = bool(lod_matches)
    lod_types   = ", ".join(dict.fromkeys(m.upper() for m in lod_matches))

    tc_matches  = TABLE_CALC_PAT.findall(formula)
    has_tc      = bool(tc_matches)
    tc_funcs    = ", ".join(dict.fromkeys(m.upper() for m in tc_matches))

    if has_lod or has_tc:
        complexity = "HIGH"
    elif re.search(r'\b(IF|CASE|IIF|ELSEIF)\b', formula, re.I):
        complexity = "MEDIUM"
    elif re.search(r'\bDATE(DIFF|PART|ADD|NAME|TRUNC)?\b', formula, re.I):
        complexity = "MEDIUM"
    else:
        complexity = "LOW"

    guidance_parts = []
    for lod_type in dict.fromkeys(m.upper() for m in lod_matches):
        note = LOD_DAX.get(lod_type, "")
        if note:
            guidance_parts.append(f"{lod_type}: {note}")
    if has_tc:
        guidance_parts.append(
            "Table calc: Re-implement using RANKX / CALCULATE with time-intelligence / window functions in DAX"
        )
    return {
        "has_lod": has_lod, "lod_types": lod_types,
        "has_table_calc": has_tc, "table_calc_functions": tc_funcs,
        "migration_complexity": complexity,
        "dax_guidance": " | ".join(guidance_parts),
    }


def analyse_sql(query: str) -> Dict:
    """Analyse a Custom SQL block for structural patterns."""
    if not query or not query.strip():
        return {
            "line_count": 0, "has_subquery": False, "has_joins": False,
            "has_unions": False, "has_where": False, "has_group_by": False,
            "has_window_func": False, "migration_notes": "",
        }

    has_subquery    = bool(re.search(r'\(\s*SELECT\b', query, re.I))
    has_joins       = bool(re.search(r'\b(INNER|LEFT|RIGHT|FULL|CROSS)\s+JOIN\b', query, re.I))
    has_unions      = bool(re.search(r'\bUNION(\s+ALL)?\b', query, re.I))
    has_where       = bool(re.search(r'\bWHERE\b', query, re.I))
    has_group_by    = bool(re.search(r'\bGROUP\s+BY\b', query, re.I))
    has_window_func = bool(WINDOW_FUNC_PAT.search(query))

    notes = []
    if has_window_func:
        notes.append("Window functions — re-implement as RANKX/EARLIER in DAX or index pattern in Power Query")
    if has_subquery:
        notes.append("Subquery — flatten as a separate Power Query step or CTE")
    if has_unions:
        notes.append("UNION — use Append Queries in Power Query or UNION() in DAX")
    if has_joins:
        notes.append("JOINs — replicate as relationships or Merge Queries in Power BI")
    if has_group_by:
        notes.append("GROUP BY — DAX SUMMARIZE or Power Query Table.Group")
    if not notes:
        notes.append("Simple SELECT — straightforward Power Query migration")

    return {
        "line_count":      len(query.splitlines()),
        "has_subquery":    has_subquery,
        "has_joins":       has_joins,
        "has_unions":      has_unions,
        "has_where":       has_where,
        "has_group_by":    has_group_by,
        "has_window_func": has_window_func,
        "migration_notes": "; ".join(notes),
    }


def _weighted_complexity(wbr: "WorkbookRecord", custom_sql_count: int) -> Tuple[str, int]:
    """
    Return (complexity_label, numeric_score) using a weighted model:
      LOD field        = 10 pts
      Table calc field =  8 pts
      Custom SQL table =  5 pts
      Dashboard        =  3 pts
      Sheet            =  1 pt
      Calc fields > 20 = +10 pts bonus
    HIGH >= 20 | MEDIUM >= 5 | LOW < 5
    """
    score = (
        wbr.lod_fields        * 10
        + wbr.table_calc_fields * 8
        + custom_sql_count      * 5
        + wbr.dashboard_count   * 3
        + wbr.sheet_count       * 1
        + (10 if wbr.calc_fields > 20 else 0)
    )
    if score >= 20:
        label = "HIGH"
    elif score >= 5:
        label = "MEDIUM"
    else:
        label = "LOW"
    return label, score


# ─────────────────────────────────────────────────────────────────────────────
# SITE EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────

def extract_site(
    session:   requests.Session,
    cfg:       Dict,
    token:     str,
    site_id:   str,
    site_name: str,
    logger:    logging.Logger,
    since_dt:  Optional[datetime] = None,
) -> Tuple[
    List[FieldRecord], List[WorkbookRecord], List[DatasourceRecord],
    List[CustomSQLRecord], List[OwnerRecord], List[ParameterRecord],
    List[PublishedDatasourceRecord], List[FlowRecord],
]:
    fields:       List[FieldRecord]               = []
    workbooks:    List[WorkbookRecord]             = []
    datasources:  List[DatasourceRecord]           = []
    custom_sql:   List[CustomSQLRecord]            = []
    owners:       List[OwnerRecord]                = []
    parameters:   List[ParameterRecord]            = []
    pub_ds:       List[PublishedDatasourceRecord]  = []
    flows:        List[FlowRecord]                 = []

    wb_index: Dict[str, WorkbookRecord] = {}

    inactive_days    = int(cfg.get("inactive_owner_days",    90))
    unused_threshold = int(cfg.get("unused_view_threshold",   0))

    logger.info("── Extracting site: '%s'", site_name)

    # ── 0. REST API: users, workbook list, view counts (parallel) ────────
    logger.info("  Fetching REST API data (users / workbooks / views)…")
    users_dict:  Dict[str, Dict] = {}
    wb_rest:     Dict            = {"_by_name": {}}
    view_counts: Dict[str, int]  = {}

    def _fetch_users():
        return fetch_users_rest(session, cfg, token, site_id, logger)

    def _fetch_wb():
        return fetch_workbooks_rest(session, cfg, token, site_id, logger)

    def _fetch_views():
        return fetch_view_counts_rest(session, cfg, token, site_id, logger)

    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as ex:
            fu = ex.submit(_fetch_users)
            fw = ex.submit(_fetch_wb)
            fv = ex.submit(_fetch_views)
            users_dict  = fu.result()
            wb_rest     = fw.result()
            view_counts = fv.result()
    except Exception as exc:
        logger.warning("  REST enrichment failed (%s) — usage/owner data will be empty.", exc)

    user_by_name: Dict[str, Dict] = {
        u.get("name", "").lower(): u for u in users_dict.values()
    }

    # Build since_filter for GraphQL if --since was supplied
    gql_since_vars: Dict = {}
    if since_dt:
        since_str = since_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        gql_since_vars = {"filter": {"updatedAtGt": since_str}}
        logger.info("  Incremental mode: only workbooks updated after %s", since_str)

    # ── 1. Workbook inventory (seeds wb_index for ALL workbooks) ─────────
    logger.info("  Querying workbook inventory (GQL_WORKBOOKS)…")
    for node in safe_paginate(session, cfg, token, GQL_WORKBOOKS,
                              "workbooksConnection", logger, extra_vars=gql_since_vars):
        wb_name = node.get("name", "Unknown")
        proj    = node.get("projectName", "Default") or "Default"
        owner   = (node.get("owner") or {}).get("name", "")
        created = str(node.get("createdAt", ""))[:10]
        updated = str(node.get("updatedAt", ""))[:10]
        sheets  = (node.get("sheetsConnection")    or {}).get("totalCount", 0)
        dashes  = (node.get("dashboardsConnection") or {}).get("totalCount", 0)

        wb_key = f"{site_name}|{proj}|{wb_name}"
        if wb_key not in wb_index:
            wb_index[wb_key] = WorkbookRecord(
                site_name=site_name, project_name=proj,
                workbook_name=wb_name, owner_name=owner,
                created_at=created, updated_at=updated,
                sheet_count=int(sheets or 0),
                dashboard_count=int(dashes or 0),
            )

    # ── 2. Calculated fields ──────────────────────────────────────────────
    logger.info("  Querying calculated fields…")
    for node in safe_paginate(session, cfg, token, GQL_CALCULATED_FIELDS,
                              "calculatedFieldsConnection", logger):
        if node.get("isHidden"):
            continue
        formula = (node.get("formula") or "").strip()
        if not formula:
            continue

        ds      = node.get("embeddedInDatasource") or {}
        wb      = ds.get("workbook") or {}
        wb_name = wb.get("name", "Unknown")
        proj    = wb.get("projectName", "Default") or "Default"
        owner   = (wb.get("owner") or {}).get("name", "")
        created = str(wb.get("createdAt", ""))[:10]
        updated = str(wb.get("updatedAt", ""))[:10]
        dtype   = (node.get("dataType") or "STRING").upper()
        role    = (node.get("role") or "").lower()
        ds_name = ds.get("name", "")

        analysis = analyse_formula(formula)
        fields.append(FieldRecord(
            site_name=site_name, project_name=proj,
            workbook_name=wb_name, workbook_owner=owner,
            datasource_name=ds_name,
            field_name=node.get("name", ""),
            datatype=dtype, pbi_datatype=DTYPE_MAP.get(dtype, dtype),
            role=role, formula=formula, **analysis,
        ))

        # Ensure workbook exists in index (calc-field pass can surface new ones)
        wb_key = f"{site_name}|{proj}|{wb_name}"
        if wb_key not in wb_index:
            wb_index[wb_key] = WorkbookRecord(
                site_name=site_name, project_name=proj,
                workbook_name=wb_name, owner_name=owner,
                created_at=created, updated_at=updated,
            )
        wbr = wb_index[wb_key]
        wbr.calc_fields += 1
        if analysis["has_lod"]:        wbr.lod_fields += 1
        if analysis["has_table_calc"]: wbr.table_calc_fields += 1

    # ── 3. Datasources ────────────────────────────────────────────────────
    logger.info("  Querying embedded datasources…")
    ds_seen = set()
    for node in safe_paginate(session, cfg, token, GQL_DATASOURCES,
                              "embeddedDatasourcesConnection", logger):
        wb      = node.get("workbook") or {}
        wb_name = wb.get("name", "Unknown")
        proj    = wb.get("projectName", "Default") or "Default"
        ds_name = node.get("name", "")
        ds_key  = f"{site_name}|{proj}|{wb_name}|{ds_name}"
        if ds_key in ds_seen:
            continue
        ds_seen.add(ds_key)

        conn_type = ""
        for t in (node.get("upstreamTablesConnection") or {}).get("nodes", []):
            ct = t.get("connectionType") or (t.get("database") or {}).get("connectionType", "")
            if ct:
                conn_type = ct; break
        if not conn_type:
            pub_nodes = (node.get("upstreamDatasourcesConnection") or {}).get("nodes", [])
            if pub_nodes:
                conn_type = pub_nodes[0].get("connectionType", "")

        ds_fields = [f for f in fields
                     if f.workbook_name == wb_name and f.datasource_name == ds_name]
        datasources.append(DatasourceRecord(
            site_name=site_name, project_name=proj,
            workbook_name=wb_name, datasource_name=ds_name,
            connection_type=conn_type or "embedded",
            is_published=bool((node.get("upstreamDatasourcesConnection") or {}).get("nodes")),
            field_count=len(ds_fields),
            calc_count=sum(1 for f in ds_fields if f.formula),
        ))

    # ── 4. Custom SQL ─────────────────────────────────────────────────────
    logger.info("  Querying custom SQL tables…")
    sql_seen = set()
    # Build workbook→custom SQL count for weighted scoring
    wb_sql_count: Dict[str, int] = {}
    for node in safe_paginate(session, cfg, token, GQL_CUSTOM_SQL,
                              "customSQLTablesConnection", logger):
        sql_name  = node.get("name", "")
        query_txt = (node.get("query") or "").strip()
        conn_type = node.get("connectionType", "")
        if not query_txt:
            continue

        ds_nodes = (node.get("downstreamDatasourcesConnection") or {}).get("nodes", [])
        if not ds_nodes:
            ds_nodes = [{"name": "", "workbook": None}]

        for ds in ds_nodes:
            wb      = ds.get("workbook") or {}
            wb_name = wb.get("name", "Unknown")
            proj    = wb.get("projectName", "Default") or "Default"
            ds_name = ds.get("name", "")
            sql_key = f"{site_name}|{proj}|{wb_name}|{sql_name}"
            if sql_key in sql_seen:
                continue
            sql_seen.add(sql_key)

            wb_key = f"{site_name}|{proj}|{wb_name}"
            wb_sql_count[wb_key] = wb_sql_count.get(wb_key, 0) + 1

            analysis = analyse_sql(query_txt)
            custom_sql.append(CustomSQLRecord(
                site_name=site_name, project_name=proj,
                workbook_name=wb_name, datasource_name=ds_name,
                sql_name=sql_name, connection_type=conn_type,
                sql_query=query_txt, **analysis,
            ))

    # ── 5. Parameters ─────────────────────────────────────────────────────
    logger.info("  Querying parameters…")
    for node in safe_paginate(session, cfg, token, GQL_PARAMETERS,
                              "parametersConnection", logger):
        wb      = node.get("workbook") or {}
        wb_name = wb.get("name", "Unknown")
        proj    = wb.get("projectName", "Default") or "Default"
        dtype   = (node.get("dataType") or "string").lower()
        parameters.append(ParameterRecord(
            site_name=site_name, project_name=proj, workbook_name=wb_name,
            parameter_name=node.get("name", ""),
            data_type=dtype,
            default_value=str(node.get("defaultValue") or ""),
            allowable_values=(node.get("allowableValuesType") or "ALL").upper(),
            pbi_equivalent=PARAM_PBI_MAP.get(dtype, "Power BI parameter"),
        ))

    # ── 6. Published datasources ──────────────────────────────────────────
    logger.info("  Querying published datasources…")
    for node in safe_paginate(session, cfg, token, GQL_PUBLISHED_DATASOURCES,
                              "publishedDatasourcesConnection", logger):
        proj    = node.get("projectName", "Default") or "Default"
        owner   = (node.get("owner") or {}).get("name", "")
        has_ext = bool(node.get("hasExtracts"))
        dw_wb   = (node.get("downstreamWorkbooksConnection") or {})
        dw_cnt  = int(dw_wb.get("totalCount", 0))

        conn_type = ""
        for t in (node.get("upstreamTablesConnection") or {}).get("nodes", []):
            ct = t.get("connectionType") or (t.get("database") or {}).get("connectionType", "")
            if ct:
                conn_type = ct; break

        note = "Migrate to shared Power BI Dataset or Dataflow"
        if has_ext:
            note += " (has extract — consider scheduled refresh in Power BI Service)"
        if dw_cnt > 5:
            note += f" — {dw_cnt} downstream workbooks; high-impact shared dataset"

        pub_ds.append(PublishedDatasourceRecord(
            site_name=site_name, project_name=proj,
            datasource_name=node.get("name", ""),
            owner_name=owner,
            connection_type=conn_type or "unknown",
            has_extracts=has_ext,
            downstream_workbooks=dw_cnt,
            created_at=str(node.get("createdAt", ""))[:10],
            updated_at=str(node.get("updatedAt", ""))[:10],
            migration_note=note,
        ))

    # ── 7. Flows ──────────────────────────────────────────────────────────
    logger.info("  Querying Tableau Prep flows…")
    for node in safe_paginate(session, cfg, token, GQL_FLOWS,
                              "flowsConnection", logger):
        proj    = node.get("projectName", "Default") or "Default"
        owner   = (node.get("owner") or {}).get("name", "")
        steps   = len(node.get("outputSteps") or [])
        flows.append(FlowRecord(
            site_name=site_name, project_name=proj,
            flow_name=node.get("name", ""),
            owner_name=owner,
            created_at=str(node.get("createdAt", ""))[:10],
            updated_at=str(node.get("updatedAt", ""))[:10],
            output_steps=steps,
            migration_note="Migrate to Power Query M or Microsoft Fabric Dataflow Gen2",
        ))

    # ── 8. Complexity scoring + REST enrichment ───────────────────────────
    wb_by_name = wb_rest.get("_by_name", {})

    for wb_key, wbr in wb_index.items():
        sql_count = wb_sql_count.get(wb_key, 0)
        label, score = _weighted_complexity(wbr, sql_count)
        wbr.migration_complexity = label
        wbr.complexity_score     = score

        # View count & usage flags
        rest_wb_id = wb_by_name.get(
            (wbr.workbook_name.lower(), wbr.project_name.lower()), ""
        )
        if rest_wb_id:
            wbr.total_views = view_counts.get(rest_wb_id, 0)
        wbr.never_viewed = (wbr.total_views == 0)
        wbr.unused       = (wbr.total_views <= unused_threshold)

        # Owner activity
        owner_user = user_by_name.get(wbr.owner_name.lower(), {})
        if owner_user:
            last_login_iso      = owner_user.get("lastLogin", "")
            days                = _days_since(last_login_iso)
            wbr.owner_last_login = last_login_iso[:10] if last_login_iso else "Never"
            wbr.owner_is_active = "Yes" if (days != -1 and days <= inactive_days) else "No"
        else:
            wbr.owner_last_login = "Unknown"
            wbr.owner_is_active  = "Unknown"

        workbooks.append(wbr)

    # Sort workbooks: HIGH first, then by score desc
    workbooks.sort(key=lambda w: (-w.complexity_score,
                                  w.site_name, w.project_name, w.workbook_name))

    # ── 9. Build OwnerRecord list ─────────────────────────────────────────
    for user in users_dict.values():
        uname      = user.get("name", "")
        last_login = user.get("lastLogin", "")
        days       = _days_since(last_login)
        is_active  = (days != -1 and days <= inactive_days)
        owned      = [w for w in workbooks if w.owner_name.lower() == uname.lower()]
        owners.append(OwnerRecord(
            site_name=site_name,
            username=uname,
            display_name=user.get("fullName", uname),
            email=user.get("email", ""),
            site_role=user.get("siteRole", ""),
            last_login=last_login[:10] if last_login else "Never",
            days_since_login=days,
            is_active=is_active,
            owned_workbooks=len(owned),
            owned_high=sum(1 for w in owned if w.migration_complexity == "HIGH"),
            owned_medium=sum(1 for w in owned if w.migration_complexity == "MEDIUM"),
        ))

    logger.info(
        "  Site '%s': %d fields | %d workbooks | %d datasources | %d custom SQL "
        "| %d params | %d pub-ds | %d flows | %d owners",
        site_name, len(fields), len(workbooks), len(datasources),
        len(custom_sql), len(parameters), len(pub_ds), len(flows), len(owners),
    )
    return fields, workbooks, datasources, custom_sql, owners, parameters, pub_ds, flows



# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — CSV WRITER
# ═══════════════════════════════════════════════════════════════════════════════

def write_csvs(
    out_dir: Path,
    fields:      List[FieldRecord],
    workbooks:   List[WorkbookRecord],
    datasources: List[DatasourceRecord],
    custom_sql:  List[CustomSQLRecord],
    owners:      List[OwnerRecord],
    parameters:  List[ParameterRecord],
    pub_ds:      List[PublishedDatasourceRecord],
    flows:       List[FlowRecord],
) -> None:
    """Write one CSV per entity type into *out_dir*."""
    out_dir.mkdir(parents=True, exist_ok=True)

    def _write(name: str, rows: list, header: List[str]) -> None:
        p = out_dir / name
        with p.open("w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(header)
            w.writerows(rows)
        logger.info("  CSV: %s (%d rows)", p.name, len(rows))

    # ── Calculated fields ──
    _write("calculated_fields.csv", [
        (r.site_name, r.project_name, r.workbook_name, r.datasource_name,
         r.field_name, r.datatype, r.pbi_datatype, r.role,
         r.has_lod, r.lod_types, r.has_table_calc, r.table_calc_functions,
         r.migration_complexity, r.formula, r.dax_guidance)
        for r in fields
    ], ["site","project","workbook","datasource","field_name",
        "datatype","pbi_datatype","role",
        "has_lod","lod_types","has_table_calc","table_calc_functions",
        "migration_complexity","formula","dax_guidance"])

    # ── Workbooks ──
    _write("workbooks.csv", [
        (r.site_name, r.project_name, r.workbook_name, r.owner_name,
         r.owner_is_active, r.total_views, r.never_viewed,
         r.sheet_count, r.dashboard_count,
         r.calc_fields, r.lod_fields, r.table_calc_fields,
         r.migration_complexity, r.complexity_score,
         r.created_at, r.updated_at)
        for r in workbooks
    ], ["site","project","workbook","owner","owner_active",
        "total_views","never_viewed","sheets","dashboards",
        "calc_fields","lod_fields","table_calc_fields",
        "complexity","complexity_score","created_at","updated_at"])

    # ── Embedded datasources ──
    _write("datasources.csv", [
        (r.site_name, r.project_name, r.workbook_name, r.datasource_name,
         r.connection_type, r.is_published, r.field_count, r.calc_count)
        for r in datasources
    ], ["site","project","workbook","datasource","connection_type",
        "is_published","field_count","calc_count"])

    # ── Custom SQL ──
    _write("custom_sql.csv", [
        (r.site_name, r.project_name, r.workbook_name, r.datasource_name,
         r.sql_name, r.connection_type, r.line_count,
         r.has_subquery, r.has_joins, r.has_unions,
         r.has_where, r.has_group_by, r.has_window_func,
         r.migration_notes, r.sql_query)
        for r in custom_sql
    ], ["site","project","workbook","datasource","sql_name","connection_type",
        "line_count","has_subquery","has_joins","has_unions",
        "has_where","has_group_by","has_window_func",
        "migration_notes","sql_query"])

    # ── Owner activity ──
    _write("owner_activity.csv", [
        (r.site_name, r.username, r.display_name, r.email, r.site_role,
         r.last_login, r.days_since_login,
         "Yes" if r.is_active else "No",
         r.owned_workbooks, r.owned_high, r.owned_medium)
        for r in owners
    ], ["site","username","display_name","email","site_role",
        "last_login","days_since_login","is_active",
        "owned_workbooks","owned_high","owned_medium"])

    # ── Parameters ──
    _write("parameters.csv", [
        (r.site_name, r.project_name, r.workbook_name,
         r.parameter_name, r.data_type, r.default_value,
         r.allowable_values, r.pbi_equivalent)
        for r in parameters
    ], ["site","project","workbook","parameter_name","data_type",
        "default_value","allowable_values","pbi_equivalent"])

    # ── Published datasources ──
    _write("published_datasources.csv", [
        (r.site_name, r.project_name, r.datasource_name, r.owner_name,
         r.connection_type, r.has_extracts, r.downstream_workbooks,
         r.created_at, r.updated_at, r.migration_note)
        for r in pub_ds
    ], ["site","project","datasource_name","owner","connection_type",
        "has_extracts","downstream_workbooks",
        "created_at","updated_at","migration_note"])

    # ── Flows ──
    _write("flows.csv", [
        (r.site_name, r.project_name, r.flow_name, r.owner_name,
         r.created_at, r.updated_at, r.output_steps, r.migration_note)
        for r in flows
    ], ["site","project","flow_name","owner",
        "created_at","updated_at","output_steps","migration_note"])


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — EXCEL WRITER
# ═══════════════════════════════════════════════════════════════════════════════

def _header(ws, cols: List[str], hdr_fill, hdr_font, hdr_border) -> None:
    for ci, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.border    = hdr_border
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def _autofit(ws) -> None:
    for col_cells in ws.columns:
        max_len = max(
            (len(str(c.value)) if c.value is not None else 0) for c in col_cells
        )
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 60)


def write_excel(
    out_path: Path,
    fields:      List[FieldRecord],
    workbooks:   List[WorkbookRecord],
    datasources: List[DatasourceRecord],
    custom_sql:  List[CustomSQLRecord],
    owners:      List[OwnerRecord],
    parameters:  List[ParameterRecord],
    pub_ds:      List[PublishedDatasourceRecord],
    flows:       List[FlowRecord],
) -> None:
    """Write a multi-tab Excel assessment workbook."""
    wb_xl = XLWorkbook()
    wb_xl.remove(wb_xl.active)

    HDR_FILL   = PatternFill("solid", fgColor="1F497D")
    HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
    THIN       = Side(style="thin", color="AAAAAA")
    HDR_BORDER = Border(left=THIN, right=THIN, bottom=THIN)
    RED_FILL   = PatternFill("solid", fgColor="FFCCCC")
    YEL_FILL   = PatternFill("solid", fgColor="FFFF99")
    GRN_FILL   = PatternFill("solid", fgColor="CCFFCC")
    WRAP       = Alignment(wrap_text=True, vertical="top")

    def _add(title: str):
        ws = wb_xl.create_sheet(title=title)
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 28
        return ws

    def _row_fill(ws, row_idx: int, fill) -> None:
        for cell in ws[row_idx]:
            cell.fill = fill

    # ── Tab 1: Summary ─────────────────────────────────────────────────────
    ws1 = _add("Summary")
    _header(ws1, ["Metric", "Value"], HDR_FILL, HDR_FONT, HDR_BORDER)
    summary_rows = [
        ("Total Sites",             len({w.site_name for w in workbooks})),
        ("Total Workbooks",         len(workbooks)),
        ("Total Published DS",      len(pub_ds)),
        ("Total Flows",             len(flows)),
        ("Total Calc Fields",       len(fields)),
        ("  — LOD Fields",          sum(1 for f in fields if f.has_lod)),
        ("  — Table Calc Fields",   sum(1 for f in fields if f.has_table_calc)),
        ("Total Custom SQL Blocks",  len(custom_sql)),
        ("HIGH Complexity WBs",     sum(1 for w in workbooks if w.migration_complexity == "HIGH")),
        ("MEDIUM Complexity WBs",   sum(1 for w in workbooks if w.migration_complexity == "MEDIUM")),
        ("LOW Complexity WBs",      sum(1 for w in workbooks if w.migration_complexity == "LOW")),
        ("Never-Viewed Workbooks",  sum(1 for w in workbooks if w.never_viewed or w.total_views == 0)),
        ("Active Owners",           sum(1 for o in owners if o.is_active)),
        ("Inactive Owners",         sum(1 for o in owners if not o.is_active)),
        ("Total Parameters",        len(parameters)),
    ]
    for ri, (metric, val) in enumerate(summary_rows, 2):
        ws1.cell(row=ri, column=1, value=metric)
        c = ws1.cell(row=ri, column=2, value=val)
    _autofit(ws1)

    # ── Tab 2: Workbooks (sorted HIGH → MEDIUM → LOW) ─────────────────────
    ws2 = _add("Workbooks")
    WB_COLS = ["Site", "Project", "Workbook", "Owner", "Owner Active",
               "Total Views", "Never Viewed",
               "Sheets", "Dashboards",
               "Calc Fields", "LOD Fields", "Table Calcs",
               "Complexity", "Score", "Created", "Updated"]
    _header(ws2, WB_COLS, HDR_FILL, HDR_FONT, HDR_BORDER)
    sorted_wbs = sorted(workbooks, key=lambda x: (-x.complexity_score, x.workbook_name))
    for ri, r in enumerate(sorted_wbs, 2):
        vals = [r.site_name, r.project_name, r.workbook_name, r.owner_name,
                r.owner_is_active,
                r.total_views, "Yes" if (r.never_viewed or r.total_views == 0) else "No",
                r.sheet_count, r.dashboard_count,
                r.calc_fields, r.lod_fields, r.table_calc_fields,
                r.migration_complexity, r.complexity_score,
                r.created_at[:10] if r.created_at else "",
                r.updated_at[:10] if r.updated_at else ""]
        for ci, v in enumerate(vals, 1):
            ws2.cell(row=ri, column=ci, value=v).alignment = WRAP
        fill = (RED_FILL if r.migration_complexity == "HIGH"
                else YEL_FILL if r.migration_complexity == "MEDIUM"
                else GRN_FILL)
        _row_fill(ws2, ri, fill)
    _autofit(ws2)

    # ── Tab 3: Calculated Fields ───────────────────────────────────────────
    ws3 = _add("Calculated Fields")
    CF_COLS = ["Site", "Project", "Workbook", "Datasource", "Field Name",
               "Data Type", "PBI Type", "Role",
               "Has LOD?", "LOD Types", "Has Table Calc?", "Table Calc Functions",
               "Complexity", "Formula", "DAX Guidance"]
    _header(ws3, CF_COLS, HDR_FILL, HDR_FONT, HDR_BORDER)
    for ri, r in enumerate(fields, 2):
        vals = [r.site_name, r.project_name, r.workbook_name, r.datasource_name,
                r.field_name, r.datatype, r.pbi_datatype, r.role,
                "Yes" if r.has_lod else "No",
                r.lod_types,
                "Yes" if r.has_table_calc else "No",
                r.table_calc_functions,
                r.migration_complexity, r.formula, r.dax_guidance]
        for ci, v in enumerate(vals, 1):
            ws3.cell(row=ri, column=ci, value=v).alignment = WRAP
        if r.has_lod:
            _row_fill(ws3, ri, YEL_FILL)
        elif r.has_table_calc:
            _row_fill(ws3, ri, PatternFill("solid", fgColor="E2EFDA"))
    _autofit(ws3)

    # ── Tab 4: LOD Fields only ─────────────────────────────────────────────
    ws4 = _add("LOD Fields")
    _header(ws4, CF_COLS, HDR_FILL, HDR_FONT, HDR_BORDER)
    for ri, r in enumerate((r for r in fields if r.has_lod), 2):
        vals = [r.site_name, r.project_name, r.workbook_name, r.datasource_name,
                r.field_name, r.datatype, r.pbi_datatype, r.role,
                "Yes", r.lod_types,
                "Yes" if r.has_table_calc else "No",
                r.table_calc_functions,
                r.migration_complexity, r.formula, r.dax_guidance]
        for ci, v in enumerate(vals, 1):
            ws4.cell(row=ri, column=ci, value=v).alignment = WRAP
    _autofit(ws4)

    # ── Tab 5: Table Calcs only ────────────────────────────────────────────
    ws5 = _add("Table Calcs")
    _header(ws5, CF_COLS, HDR_FILL, HDR_FONT, HDR_BORDER)
    for ri, r in enumerate((r for r in fields if r.has_table_calc), 2):
        vals = [r.site_name, r.project_name, r.workbook_name, r.datasource_name,
                r.field_name, r.datatype, r.pbi_datatype, r.role,
                "Yes" if r.has_lod else "No",
                r.lod_types,
                "Yes", r.table_calc_functions,
                r.migration_complexity, r.formula, r.dax_guidance]
        for ci, v in enumerate(vals, 1):
            ws5.cell(row=ri, column=ci, value=v).alignment = WRAP
    _autofit(ws5)

    # ── Tab 6: Embedded Datasources ────────────────────────────────────────
    ws6 = _add("Datasources")
    DS_COLS = ["Site", "Project", "Workbook", "Datasource",
               "Connection Type", "Published?", "Field Count", "Calc Count"]
    _header(ws6, DS_COLS, HDR_FILL, HDR_FONT, HDR_BORDER)
    for ri, r in enumerate(datasources, 2):
        vals = [r.site_name, r.project_name, r.workbook_name, r.datasource_name,
                r.connection_type,
                "Yes" if r.is_published else "No",
                r.field_count, r.calc_count]
        for ci, v in enumerate(vals, 1):
            ws6.cell(row=ri, column=ci, value=v).alignment = WRAP
    _autofit(ws6)

    # ── Tab 7: Custom SQL ──────────────────────────────────────────────────
    ws7 = _add("Custom SQL")
    SQL_COLS = ["Site", "Project", "Workbook", "Datasource", "SQL Name",
                "Connection Type", "Lines",
                "Subquery?", "Joins?", "Unions?",
                "WHERE?", "GROUP BY?", "Window Func?",
                "Migration Notes", "SQL Query"]
    _header(ws7, SQL_COLS, HDR_FILL, HDR_FONT, HDR_BORDER)
    for ri, r in enumerate(custom_sql, 2):
        vals = [r.site_name, r.project_name, r.workbook_name, r.datasource_name,
                r.sql_name, r.connection_type, r.line_count,
                "Yes" if r.has_subquery  else "No",
                "Yes" if r.has_joins     else "No",
                "Yes" if r.has_unions    else "No",
                "Yes" if r.has_where     else "No",
                "Yes" if r.has_group_by  else "No",
                "Yes" if r.has_window_func else "No",
                r.migration_notes, r.sql_query]
        for ci, v in enumerate(vals, 1):
            ws7.cell(row=ri, column=ci, value=v).alignment = WRAP
        if r.has_window_func:
            _row_fill(ws7, ri, RED_FILL)
        elif r.has_subquery:
            _row_fill(ws7, ri, YEL_FILL)
    _autofit(ws7)

    # ── Tab 8: Owner Activity ──────────────────────────────────────────────
    ws8 = _add("Owner Activity")
    OW_COLS = ["Site", "Username", "Display Name", "Email", "Site Role",
               "Last Login", "Days Since Login", "Active?",
               "Owned WBs", "HIGH WBs", "MEDIUM WBs"]
    _header(ws8, OW_COLS, HDR_FILL, HDR_FONT, HDR_BORDER)
    for ri, r in enumerate(sorted(owners, key=lambda x: (x.is_active, x.username)), 2):
        vals = [r.site_name, r.username, r.display_name, r.email, r.site_role,
                r.last_login, r.days_since_login,
                "Yes" if r.is_active else "No",
                r.owned_workbooks, r.owned_high, r.owned_medium]
        for ci, v in enumerate(vals, 1):
            ws8.cell(row=ri, column=ci, value=v).alignment = WRAP
        if not r.is_active:
            _row_fill(ws8, ri, RED_FILL)
    _autofit(ws8)

    # ── Tab 9: Parameters ──────────────────────────────────────────────────
    ws9 = _add("Parameters")
    PA_COLS = ["Site", "Project", "Workbook", "Parameter Name", "Data Type",
               "Default Value", "Allowable Values", "Power BI Equivalent"]
    _header(ws9, PA_COLS, HDR_FILL, HDR_FONT, HDR_BORDER)
    for ri, r in enumerate(parameters, 2):
        vals = [r.site_name, r.project_name, r.workbook_name,
                r.parameter_name, r.data_type, r.default_value,
                r.allowable_values, r.pbi_equivalent]
        for ci, v in enumerate(vals, 1):
            ws9.cell(row=ri, column=ci, value=v).alignment = WRAP
    _autofit(ws9)

    # ── Tab 10: Published Datasources ──────────────────────────────────────
    ws10 = _add("Published Datasources")
    PD_COLS = ["Site", "Project", "Datasource", "Owner", "Connection Type",
               "Has Extracts", "Downstream WBs",
               "Created", "Updated", "Migration Note"]
    _header(ws10, PD_COLS, HDR_FILL, HDR_FONT, HDR_BORDER)
    for ri, r in enumerate(
        sorted(pub_ds, key=lambda x: -x.downstream_workbooks), 2
    ):
        vals = [r.site_name, r.project_name, r.datasource_name, r.owner_name,
                r.connection_type,
                "Yes" if r.has_extracts else "No",
                r.downstream_workbooks,
                r.created_at[:10] if r.created_at else "",
                r.updated_at[:10] if r.updated_at else "",
                r.migration_note]
        for ci, v in enumerate(vals, 1):
            ws10.cell(row=ri, column=ci, value=v).alignment = WRAP
        if r.downstream_workbooks >= 5:
            _row_fill(ws10, ri, YEL_FILL)
    _autofit(ws10)

    # ── Tab 11: Flows (Tableau Prep) ───────────────────────────────────────
    ws11 = _add("Flows")
    FL_COLS = ["Site", "Project", "Flow Name", "Owner",
               "Created", "Updated", "Output Steps", "Migration Note"]
    _header(ws11, FL_COLS, HDR_FILL, HDR_FONT, HDR_BORDER)
    for ri, r in enumerate(flows, 2):
        vals = [r.site_name, r.project_name, r.flow_name, r.owner_name,
                r.created_at[:10] if r.created_at else "",
                r.updated_at[:10] if r.updated_at else "",
                r.output_steps, r.migration_note]
        for ci, v in enumerate(vals, 1):
            ws11.cell(row=ri, column=ci, value=v).alignment = WRAP
    _autofit(ws11)

    wb_xl.save(str(out_path))
    logger.info("Excel workbook saved: %s", out_path)




# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — MAIN ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Tableau → Power BI Migration Assessment Extractor",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""
            ── USAGE EXAMPLES ────────────────────────────────────────────────────

            List every site on the server (no extraction):
              python tableau_metadata_api_extractor.py --config config_metadata.json --list-sites

            Extract specific sites (comma-separated display names or contentUrls):
              python tableau_metadata_api_extractor.py --config config_metadata.json \\
                --sites "Finance Analytics,HR Portal,Marketing,Sales,IT Operations"

            Extract all sites automatically in batches of 5:
              python tableau_metadata_api_extractor.py --config config_metadata.json \\
                --batch-size 5

            Run only batch 2 of a multi-batch job (0-indexed):
              python tableau_metadata_api_extractor.py --config config_metadata.json \\
                --batch-size 5 --batch 2

            Incremental update (only workbooks changed since date):
              python tableau_metadata_api_extractor.py --config config_metadata.json \\
                --sites "Finance Analytics,HR Portal" --since 2025-01-01

            Use environment variable for password (recommended):
              TABLEAU_PASSWORD=secret python tableau_metadata_api_extractor.py \\
                --config config_metadata.json --sites "Finance Analytics"

            ── OUTPUT STRUCTURE ──────────────────────────────────────────────────

            tableau_assessment_output/
              tableau_assessment_batch_01_Finance-HR-Marketing-IT-Sales.xlsx
              tableau_assessment_batch_02_Ops-Legal-Support-Exec-Dev.xlsx
              ...
              tableau_assessment_COMBINED.xlsx   ← all batches merged
              csv/
                batch_01/  ← per-batch CSVs (written as each batch completes)
                batch_02/
                ...
                combined/  ← merged CSVs across all batches
              tableau_metadata_extractor.log
        """),
    )
    p.add_argument("--config", required=True,
                   help="Path to JSON config file (see config_metadata.example.json)")

    # ── Site selection ──────────────────────────────────────────────────────
    site_grp = p.add_mutually_exclusive_group()
    site_grp.add_argument(
        "--sites", metavar="SITE1,SITE2,...",
        help=(
            "Comma-separated list of sites to extract. Each entry is matched "
            "case-insensitively against the site's display name OR its contentUrl. "
            "Example: --sites \"Finance Analytics,HR Portal,Marketing\""
        ),
    )
    site_grp.add_argument(
        "--list-sites", action="store_true",
        help="Print all sites on the server (name + contentUrl + id) and exit.",
    )

    # ── Batch control ───────────────────────────────────────────────────────
    p.add_argument(
        "--batch-size", type=int, default=5, metavar="N",
        help=(
            "How many sites to process per batch (default: 5). "
            "Each batch produces its own Excel file; a combined Excel is written "
            "after the final batch. Use --batch to run a single specific batch."
        ),
    )
    p.add_argument(
        "--batch", type=int, default=None, metavar="N",
        help=(
            "Run only batch N (0-indexed). Useful for resuming or parallelising "
            "across machines. E.g. --batch 0 runs the first 5 sites, "
            "--batch 1 the next 5, etc. Omit to run ALL batches sequentially."
        ),
    )

    # ── Incremental mode ────────────────────────────────────────────────────
    p.add_argument("--since", metavar="YYYY-MM-DD",
                   help="Only extract workbooks updated after this date.")

    # ── Output control ──────────────────────────────────────────────────────
    p.add_argument("--no-excel", action="store_true",
                   help="Skip Excel generation (write CSVs only).")
    p.add_argument("--no-csv", action="store_true",
                   help="Skip per-batch CSV generation (write Excel only).")
    return p.parse_args()


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS FOR BATCHING
# ─────────────────────────────────────────────────────────────────────────────

def _match_sites(all_sites: List[Dict], names: List[str], log: logging.Logger) -> List[Dict]:
    """
    Return the subset of *all_sites* whose display name or contentUrl matches
    any entry in *names* (case-insensitive).  Warns if any requested name is
    not found.
    """
    lower_names = {n.strip().lower() for n in names}
    matched, matched_lower = [], set()

    for s in all_sites:
        key_name = (s.get("name") or "").lower()
        key_url  = (s.get("contentUrl") or "").lower()
        if key_name in lower_names or key_url in lower_names:
            matched.append(s)
            matched_lower.add(key_name)
            matched_lower.add(key_url)

    not_found = [n for n in names if n.strip().lower() not in matched_lower]
    if not_found:
        log.warning(
            "The following requested sites were NOT found on the server: %s",
            not_found,
        )
    return matched


def _batch_label(batch_sites: List[Dict], batch_num: int) -> str:
    """
    Build a short file-safe label for a batch, e.g.
    'batch_01_Finance-HR-Marketing-IT-Sales'
    """
    safe_names = [
        re.sub(r"[^A-Za-z0-9]", "-", s.get("name") or s.get("contentUrl") or "site")[:20]
        for s in batch_sites
    ]
    return f"batch_{batch_num:02d}_{'_'.join(safe_names)}"


def _run_batch(
    batch_sites:  List[Dict],
    batch_label:  str,
    cfg:          Dict,
    out_dir:      Path,
    since_dt:     Optional[datetime],
    args,
    log:          logging.Logger,
) -> Tuple[List, List, List, List, List, List, List, List]:
    """
    Process all sites in *batch_sites* in parallel (up to parallel_sites workers).
    Returns 8-tuple of accumulated lists for this batch.
    Writes per-batch CSVs immediately so partial results are saved even if a
    later batch fails.
    """
    b_fields:  List[FieldRecord]               = []
    b_wbs:     List[WorkbookRecord]             = []
    b_ds:      List[DatasourceRecord]           = []
    b_sql:     List[CustomSQLRecord]            = []
    b_owners:  List[OwnerRecord]                = []
    b_params:  List[ParameterRecord]            = []
    b_pub_ds:  List[PublishedDatasourceRecord]  = []
    b_flows:   List[FlowRecord]                 = []
    lock = threading.Lock()

    def _process_site(site: Dict) -> None:
        site_id   = site["id"]
        site_name = site.get("name") or site.get("contentUrl") or "Default"
        log.info("  [%s] → %s", batch_label, site_name)
        s = build_session(cfg)
        token, real_site_id = signin(s, cfg, log,
                                     site_content_url=site.get("contentUrl", ""))
        try:
            result = extract_site(s, cfg, token, real_site_id, site_name, log, since_dt)
        finally:
            signout(s, cfg, token, log)
        f, w, d, cs, o, pa, pd, fl = result
        with lock:
            b_fields.extend(f);   b_wbs.extend(w)
            b_ds.extend(d);       b_sql.extend(cs)
            b_owners.extend(o);   b_params.extend(pa)
            b_pub_ds.extend(pd);  b_flows.extend(fl)

    max_workers = max(1, min(int(cfg.get("parallel_sites", 4)), len(batch_sites)))
    if max_workers > 1:
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = {ex.submit(_process_site, s): s for s in batch_sites}
            for fut in concurrent.futures.as_completed(futures):
                site = futures[fut]
                try:
                    fut.result()
                except Exception as exc:
                    log.error("  Site '%s' failed: %s",
                              site.get("name", site.get("id")), exc, exc_info=True)
    else:
        for s in batch_sites:
            try:
                _process_site(s)
            except Exception as exc:
                log.error("  Site '%s' failed: %s",
                          s.get("name", s.get("id")), exc, exc_info=True)

    # ── Write per-batch CSVs immediately ─────────────────────────────────
    if not args.no_csv:
        csv_dir = out_dir / "csv" / batch_label
        write_csvs(csv_dir, b_fields, b_wbs, b_ds, b_sql,
                   b_owners, b_params, b_pub_ds, b_flows)
        log.info("  Batch CSVs written → %s", csv_dir)

    # ── Write per-batch Excel ─────────────────────────────────────────────
    if not args.no_excel:
        xl_path = out_dir / f"tableau_assessment_{batch_label}.xlsx"
        write_excel(xl_path, b_fields, b_wbs, b_ds, b_sql,
                    b_owners, b_params, b_pub_ds, b_flows)
        log.info("  Batch Excel written → %s", xl_path.name)

    return b_fields, b_wbs, b_ds, b_sql, b_owners, b_params, b_pub_ds, b_flows


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    args = _parse_args()
    cfg  = load_config(args.config)

    out_dir = Path(cfg.get("output_dir", "./tableau_assessment_output"))
    out_dir.mkdir(parents=True, exist_ok=True)

    log = setup_logging(cfg, out_dir)
    log.info("═" * 70)
    log.info("Tableau Metadata API Extractor")
    log.info("Server     : %s", cfg["server_url"])
    log.info("Batch size : %d", args.batch_size)
    if args.since:
        log.info("Since      : %s", args.since)
    log.info("═" * 70)

    # ── Parse --since ──────────────────────────────────────────────────────
    since_dt: Optional[datetime] = None
    if args.since:
        try:
            since_dt = datetime.strptime(args.since, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        except ValueError:
            log.error("--since must be YYYY-MM-DD, got: %s", args.since)
            raise SystemExit(1)

    # ── Prompt for credentials if not supplied in config / env vars ────────
    # Priority order (highest → lowest):
    #   1. TABLEAU_USERNAME / TABLEAU_PASSWORD environment variables
    #   2. username / password keys in config JSON
    #   3. Interactive prompt at runtime  ← this block
    #
    # The password prompt uses getpass so it is never echoed to the terminal
    # and never appears in shell history or log files.
    if not cfg.get("username"):
        try:
            cfg["username"] = input("Tableau username: ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            log.error("Username input cancelled.")
            raise SystemExit(1)
        if not cfg["username"]:
            log.error("Username cannot be empty.")
            raise SystemExit(1)

    if not cfg.get("password"):
        try:
            cfg["password"] = getpass.getpass(
                f"Password for {cfg['username']}@{cfg['server_url']}: "
            )
        except (EOFError, KeyboardInterrupt):
            print()
            log.error("Password input cancelled.")
            raise SystemExit(1)
        if not cfg["password"]:
            log.error("Password cannot be empty.")
            raise SystemExit(1)

    log.info("Credentials : %s  (password supplied via %s)",
             cfg["username"],
             "env var" if os.environ.get("TABLEAU_PASSWORD") else
             "config file" if load_config(args.config).get("password") else
             "interactive prompt")

    # ── Sign in to Default site to enumerate all sites ─────────────────────
    init_session = build_session(cfg)
    init_token, init_site_id = signin(init_session, cfg, log, site_content_url="")
    all_server_sites = get_all_sites(init_session, cfg, init_token, log)
    signout(init_session, cfg, init_token, log)

    # ── --list-sites: print table and exit ─────────────────────────────────
    if args.list_sites:
        log.info("═" * 70)
        log.info("%-4s  %-40s  %-30s  %s", "#", "Display Name", "contentUrl", "ID")
        log.info("%-4s  %-40s  %-30s  %s", "─"*4, "─"*40, "─"*30, "─"*36)
        for i, s in enumerate(all_server_sites, 1):
            log.info("%-4d  %-40s  %-30s  %s",
                     i,
                     (s.get("name") or "")[:40],
                     (s.get("contentUrl") or "Default")[:30],
                     s.get("id", ""))
        log.info("═" * 70)
        log.info("Total: %d site(s). Use --sites \"Name1,Name2,...\" to extract specific ones.",
                 len(all_server_sites))
        return

    # ── Resolve which sites to process ────────────────────────────────────
    if args.sites:
        requested = [n.strip() for n in args.sites.split(",") if n.strip()]
        sites_to_process = _match_sites(all_server_sites, requested, log)
        if not sites_to_process:
            log.error("No matching sites found. Run --list-sites to see available sites.")
            raise SystemExit(1)
        log.info("Matched %d/%d requested site(s).", len(sites_to_process), len(requested))
    else:
        sites_to_process = all_server_sites
        log.info("No --sites filter — processing ALL %d site(s).", len(sites_to_process))

    # ── Split into batches ─────────────────────────────────────────────────
    bs      = max(1, args.batch_size)
    batches = [sites_to_process[i:i + bs] for i in range(0, len(sites_to_process), bs)]
    total_batches = len(batches)

    if args.batch is not None:
        if args.batch < 0 or args.batch >= total_batches:
            log.error("--batch %d is out of range. There are %d batch(es) (0-%d).",
                      args.batch, total_batches, total_batches - 1)
            log.error("Batches available:")
            for bi, b in enumerate(batches):
                names = [s.get("name") or s.get("contentUrl") for s in b]
                log.error("  Batch %d: %s", bi, names)
            raise SystemExit(1)
        batches_to_run = [(args.batch, batches[args.batch])]
        log.info("Running single batch %d of %d.", args.batch, total_batches)
    else:
        batches_to_run = list(enumerate(batches))
        log.info("Running all %d batch(es), %d site(s) per batch.",
                 total_batches, bs)

    # ── Print the batch plan ───────────────────────────────────────────────
    log.info("─" * 70)
    log.info("BATCH PLAN:")
    for bi, batch in enumerate(batches):
        names  = [s.get("name") or s.get("contentUrl") for s in batch]
        marker = " ← this run" if (args.batch is not None and bi == args.batch) else ""
        log.info("  Batch %02d / %02d : %s%s", bi, total_batches - 1, names, marker)
    log.info("─" * 70)

    # ── Run batches, accumulate combined results ───────────────────────────
    combined_fields:  List[FieldRecord]               = []
    combined_wbs:     List[WorkbookRecord]             = []
    combined_ds:      List[DatasourceRecord]           = []
    combined_sql:     List[CustomSQLRecord]            = []
    combined_owners:  List[OwnerRecord]                = []
    combined_params:  List[ParameterRecord]            = []
    combined_pub_ds:  List[PublishedDatasourceRecord]  = []
    combined_flows:   List[FlowRecord]                 = []

    for bi, batch_sites in batches_to_run:
        label = _batch_label(batch_sites, bi)
        site_names = [s.get("name") or s.get("contentUrl") for s in batch_sites]
        log.info("▶ Starting batch %02d/%02d: %s", bi, total_batches - 1, site_names)

        f, w, d, cs, o, pa, pd, fl = _run_batch(
            batch_sites, label, cfg, out_dir, since_dt, args, log
        )

        combined_fields.extend(f);   combined_wbs.extend(w)
        combined_ds.extend(d);       combined_sql.extend(cs)
        combined_owners.extend(o);   combined_params.extend(pa)
        combined_pub_ds.extend(pd);  combined_flows.extend(fl)

        log.info("✔ Batch %02d complete — %d workbooks so far (cumulative).",
                 bi, len(combined_wbs))

    # ── Write combined outputs (only when running all batches) ─────────────
    running_all = (args.batch is None)
    if running_all and len(batches_to_run) > 1:
        log.info("─" * 70)
        log.info("Writing COMBINED outputs (%d sites, %d workbooks)…",
                 len(sites_to_process), len(combined_wbs))

        if not args.no_csv:
            csv_combined = out_dir / "csv" / "combined"
            write_csvs(csv_combined,
                       combined_fields, combined_wbs, combined_ds, combined_sql,
                       combined_owners, combined_params, combined_pub_ds, combined_flows)
            log.info("Combined CSVs → %s", csv_combined)

        if not args.no_excel:
            since_tag 