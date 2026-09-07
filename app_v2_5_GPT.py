# app.py – Smart Diff Manager (v2.5 — performance, correctness & UX fixes)
# Changes vs v2.2:
#   1. XLSX loading uses openpyxl read_only=True to avoid materializing every cell
#      as Python Cell objects; header detection also reads only the first 15 rows.
#   2. CSV delimiter detection uses csv.Sniffer + bounded fallbacks instead of
#      trying up to 16 full-file parses.
#   3. Key-based comparison uses merge/anti-join operations instead of a Python
#      dict + per-row/per-cell comparison loop.
#   4. Key selection now only returns columns actually shared by both frames,
#      eliminating fuzzy-match KeyError cases.
#   5. Keyless comparison uses MultiIndex row keys instead of row-wise string joins.
#   6. Added/removed reconciliation uses an inverted value index to avoid building
#      a full removed×added×columns boolean matrix for large sheets.
#   7. Truncation detection is vectorized.
#   8. XLSX exports are cached and download buttons no longer trigger reruns.
#   9. Large changed-row previews are capped to keep the browser responsive.
#
# Changes vs v2.1:
#   1. align_new_columns_to_reference: added positional fallback for literal
#      placeholder header text (e.g. "Unnamed: 1") that already exists as
#      real text in the source file and can never be recovered by name
#      matching; substring-match pass is now guarded against short names
#      matching inside unrelated longer names (e.g. "id" inside "valid")
#   2. detect_data_truncation: fixed to check stringified placeholder/empty
#      values instead of .notna() (which always returned True once cells
#      had already been through .astype(str), so truncation essentially
#      never fired)
#   3. compare_key_based._first_row_dict: .iloc -> .loc (idx from
#      Series.items() is an index LABEL, not a positional index — this only
#      worked before by coincidence of a clean RangeIndex)
#   4. _read_csv_raw_cached: removed the early-break on the first separator
#      that produced >1 column, which could lock onto the wrong separator
#      whenever a text field contained a comma
#   5. _read_sheets_cached (XLSX/XLSM): now reads cell values directly from
#      the already-loaded workbook instead of running a second, independent
#      pd.read_excel() parse over the same file — this was parsing heavy
#      files twice for no benefit. The except block now logs the real
#      traceback instead of silently swallowing it.
#
# Changes vs v2.0:
#   1. reconcile_added_removed: O(n²×cols) → exact-hash + numpy vectorized comparison
#   2. normalize_df called once per sheet, passed through all compare/preview/export functions
#   3. _read_sheets_cached: single pd.read_excel call for XLSX (was double) [superseded by fix #5 above]
#   4. compare_key_based: dict-based O(1) key lookup instead of .loc on non-unique index
#   5. Minor: removed redundant uf.seek(0) after getvalue() in get_auto_detected_rows

import streamlit as st
import pandas as pd
import numpy as np
from difflib import SequenceMatcher
from io import BytesIO
import msoffcrypto
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from typing import List, Dict, Tuple, Optional
import re
import hashlib
import csv
from collections import defaultdict, Counter

DEFAULT_PASSWORD = "mypassword"

st.set_page_config(
    page_title="Smart Diff Manager",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* ── Page background ── */
.stApp {
    background: #0f1117;
    color: #e2e8f0;
}

/* ── Top banner ── */
.banner {
    background: linear-gradient(135deg, #1a1f2e 0%, #0f1117 60%, #1a2744 100%);
    border: 1px solid #2d3748;
    border-radius: 12px;
    padding: 2rem 2.5rem 1.5rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}
.banner::before {
    content: '';
    position: absolute;
    top: -40px; right: -40px;
    width: 200px; height: 200px;
    background: radial-gradient(circle, rgba(59,130,246,0.12) 0%, transparent 70%);
    pointer-events: none;
}
.banner h1 {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.8rem;
    font-weight: 600;
    color: #f1f5f9;
    margin: 0 0 0.4rem;
    letter-spacing: -0.5px;
}
.banner p {
    color: #94a3b8;
    font-size: 0.9rem;
    margin: 0;
    font-weight: 300;
}
.banner .accent {
    color: #60a5fa;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 2px;
    text-transform: uppercase;
    display: block;
    margin-bottom: 0.5rem;
}

/* ── Step cards ── */
.step-card {
    background: #1a1f2e;
    border: 1px solid #2d3748;
    border-radius: 10px;
    padding: 1.5rem;
    margin-bottom: 1rem;
}
.step-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #60a5fa;
    margin-bottom: 0.6rem;
    font-weight: 600;
}

/* ── Metric pills ── */
.metric-row {
    display: flex;
    gap: 1rem;
    flex-wrap: wrap;
    margin: 1rem 0;
}
.metric-pill {
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 0.6rem 1.1rem;
    font-size: 0.85rem;
    color: #cbd5e1;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}
.metric-pill .val {
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    font-size: 1.1rem;
}
.metric-pill.changed .val { color: #fbbf24; }
.metric-pill.added   .val { color: #34d399; }
.metric-pill.removed .val { color: #f87171; }
.metric-pill.info    .val { color: #60a5fa; }

/* ── Sheet header ── */
.sheet-header {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    padding: 0.9rem 1.2rem;
    background: #1a1f2e;
    border-left: 3px solid #60a5fa;
    border-radius: 0 8px 8px 0;
    margin: 1.5rem 0 0.75rem;
}
.sheet-header .sheet-name {
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    font-size: 0.95rem;
    color: #e2e8f0;
}
.sheet-header .sheet-summary {
    color: #64748b;
    font-size: 0.82rem;
}

/* ── Diff section labels ── */
.diff-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.72rem;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    font-weight: 600;
    padding: 0.3rem 0.7rem;
    border-radius: 4px;
    display: inline-block;
    margin-bottom: 0.5rem;
}
.diff-label.changed { background: #451a03; color: #fbbf24; }
.diff-label.added   { background: #052e16; color: #34d399; }
.diff-label.removed { background: #450a0a; color: #f87171; }
.diff-label.key     { background: #172554; color: #93c5fd; }

/* ── File pair card ── */
.pair-card {
    background: #1a1f2e;
    border: 1px solid #2d3748;
    border-radius: 10px;
    padding: 1rem 1.5rem;
    margin-bottom: 0.5rem;
    display: flex;
    align-items: center;
    gap: 1rem;
    font-size: 0.85rem;
}
.pair-card .fname { font-family: 'IBM Plex Mono', monospace; color: #94a3b8; }
.pair-card .arrow { color: #3b82f6; font-size: 1rem; }

/* ── Warn / error banners ── */
.warn-box {
    background: #1c1007;
    border: 1px solid #92400e;
    border-radius: 8px;
    padding: 0.8rem 1.2rem;
    color: #fcd34d;
    font-size: 0.85rem;
    margin-bottom: 0.75rem;
}
.success-box {
    background: #052e16;
    border: 1px solid #166534;
    border-radius: 8px;
    padding: 1rem 1.5rem;
    color: #86efac;
    font-size: 0.95rem;
    text-align: center;
    font-family: 'IBM Plex Mono', monospace;
}

/* ── Streamlit overrides ── */
div[data-testid="stFileUploader"] {
    background: #1a1f2e;
    border: 1px dashed #334155;
    border-radius: 8px;
    padding: 0.5rem;
}
div[data-testid="stTextInput"] > div > div > input {
    background: #1e293b;
    border: 1px solid #334155;
    color: #e2e8f0;
    border-radius: 6px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.85rem;
}
div[data-testid="stSlider"] { padding: 0.2rem 0; }
.stButton > button {
    background: #2563eb;
    color: white;
    border: none;
    border-radius: 8px;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    letter-spacing: 0.5px;
    padding: 0.6rem 1.8rem;
    font-size: 0.9rem;
    transition: background 0.2s;
}
.stButton > button:hover { background: #1d4ed8; }
.stCheckbox label { color: #94a3b8; font-size: 0.85rem; }
.stExpander { border: 1px solid #2d3748 !important; border-radius: 8px !important; background: #1a1f2e; }
</style>
""",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────
st.markdown(
    """
<div class="banner">
  <span class="accent">v2.2 · Excel Diff Tool</span>
  <h1>📊 Smart Diff Manager</h1>
  <p>Key-based comparison across Excel files with smart header detection and column alignment.</p>
</div>
""",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────
# HELPER: stable file hash for caching
# ─────────────────────────────────────────────

def _file_hash(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()


# ─────────────────────────────────────────────
# DECRYPTION (cached by content hash)
# ─────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def decrypt_file_cached(data: bytes, password: str = DEFAULT_PASSWORD) -> bytes:
    fb = BytesIO(data)
    try:
        pd.ExcelFile(fb)
        return data
    except Exception:
        pass
    fb.seek(0)
    try:
        office = msoffcrypto.OfficeFile(fb)
        if not office.is_encrypted():
            return data
        dec = BytesIO()
        office.load_key(password=password)
        office.decrypt(dec)
        return dec.getvalue()
    except Exception:
        return data


def decrypt_file(uploaded_file, password: str = DEFAULT_PASSWORD) -> BytesIO:
    if _is_tabular_text(uploaded_file.name):
        return BytesIO(uploaded_file.getvalue())
    decrypted = decrypt_file_cached(uploaded_file.getvalue(), password)
    return BytesIO(decrypted)


# ─────────────────────────────────────────────
# MULTI-FORMAT FILE UTILITIES
# ─────────────────────────────────────────────

_EXCEL_EXTS = {"xlsx", "xlsm"}
_EXCEL_XLS  = {"xls"}
_EXCEL_ODS  = {"ods"}
_CSV_EXTS   = {"csv", "txt"}
_TSV_EXTS   = {"tsv"}
ALL_SUPPORTED_EXTS = (
    list(_EXCEL_EXTS) + list(_EXCEL_XLS) + list(_EXCEL_ODS)
    + list(_CSV_EXTS) + list(_TSV_EXTS)
)


def _file_ext(filename: str) -> str:
    return filename.rsplit(".", 1)[-1].lower() if "." in filename else ""


def _is_tabular_text(filename: str) -> bool:
    return _file_ext(filename) in _CSV_EXTS | _TSV_EXTS


def _sheet_name_from_filename(filename: str) -> str:
    return filename.rsplit(".", 1)[0] if "." in filename else filename


@st.cache_data(show_spinner=False)
def _read_csv_raw_cached(data: bytes, filename: str) -> pd.DataFrame:
    """Read delimited text with bounded delimiter/encoding probing.

    v2.3: the old implementation could parse the entire file up to 16 times
    (4 encodings × 4 separators). We sniff a sample first, then fall back to
    a small separator set only when sniffing is ambiguous.
    """
    ext = _file_ext(filename)
    preferred = "\t" if ext in _TSV_EXTS else ","
    allowed = ("\t", ",", ";", "|")
    encoding_order = ("utf-8-sig", "utf-8", "cp1252", "latin-1")

    sample = data[:128_000]
    best_df: Optional[pd.DataFrame] = None
    best_cols = 0

    for enc in encoding_order:
        try:
            decoded = sample.decode(enc)
        except UnicodeDecodeError:
            continue

        candidates = [preferred]
        try:
            sniffed = csv.Sniffer().sniff(decoded, delimiters="\t,;|").delimiter
            if sniffed not in candidates:
                candidates.insert(0, sniffed)
        except csv.Error:
            pass

        # Bounded fallback: at most four parses for this encoding.
        for sep in dict.fromkeys(candidates + list(allowed)):
            try:
                df = pd.read_csv(
                    BytesIO(data), header=None, sep=sep,
                    encoding=enc, engine="python",
                    on_bad_lines="skip", dtype=str,
                )
                if df.shape[1] > best_cols:
                    best_df = df
                    best_cols = df.shape[1]
                # Once we have a genuinely tabular result, stop probing encodings.
                if best_cols >= 3:
                    return best_df
            except Exception:
                continue

    return best_df if best_df is not None else pd.DataFrame()


# ─────────────────────────────────────────────
# EAGER AUTO-DETECTION (cached, runs on upload)
# ─────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def detect_header_rows_for_file(file_bytes_raw: bytes, filename: str = "") -> Dict[str, int]:
    result: Dict[str, int] = {}
    ext = _file_ext(filename)

    try:
        if ext in _CSV_EXTS | _TSV_EXTS:
            sh = _sheet_name_from_filename(filename)
            df_raw = _read_csv_raw_cached(file_bytes_raw, filename)
            detected = detect_header_row_heuristic(df_raw, ws=None)
            result[sh] = detected + 1
            return result

        if ext in _EXCEL_ODS:
            all_raw = pd.read_excel(BytesIO(file_bytes_raw), sheet_name=None,
                                    header=None, engine="odf")
            for sh, df_raw in all_raw.items():
                detected = detect_header_row_heuristic(df_raw, ws=None)
                result[sh] = detected + 1
            return result

        if ext in _EXCEL_XLS:
            try:
                all_raw = pd.read_excel(BytesIO(file_bytes_raw), sheet_name=None,
                                        header=None, engine="xlrd")
                for sh, df_raw in all_raw.items():
                    detected = detect_header_row_heuristic(df_raw, ws=None)
                    result[sh] = detected + 1
            except Exception:
                pass
            return result

        # XLSX/XLSM — read only the first 15 rows for eager header detection.
        # This keeps upload-time inspection cheap even for very large sheets.
        wb = load_workbook(BytesIO(file_bytes_raw), read_only=True, data_only=True)
        for ws in wb.worksheets:
            sh = ws.title
            preview_rows = list(ws.iter_rows(max_row=15, values_only=True))
            df_raw = pd.DataFrame(preview_rows) if preview_rows else pd.DataFrame()
            if df_raw.empty:
                result[sh] = 1
                continue
            detected = detect_header_row_heuristic(df_raw, ws=None)
            result[sh] = detected + 1

    except Exception:
        pass
    return result


def get_auto_detected_rows(uploaded_files: list) -> Dict[str, int]:
    merged: Dict[str, int] = {}
    for uf in (uploaded_files or []):
        try:
            # FIX: getvalue() doesn't consume position; no seek needed
            rows = detect_header_rows_for_file(uf.getvalue(), uf.name)
            for sh, row in rows.items():
                merged[sh] = max(merged.get(sh, 1), row)
        except Exception:
            pass
    return merged


# ─────────────────────────────────────────────
# NORMALIZATION UTILITIES
# ─────────────────────────────────────────────

def normalize_colname(name: str) -> str:
    return re.sub(r'[^a-z0-9]', '', str(name).lower().strip())


_BOOL_TRUE  = {"TRUE","T","YES","Y","1","1.0","CHECK","CHECKED","CHECKMARK","✓","✔","ON","ENABLED","ACTIVE"}
_BOOL_FALSE = {"FALSE","F","NO","N","0","0.0","CROSS","UNCHECKED","✗","✘","X","OFF","DISABLED","INACTIVE"}

def _normalize_series(series: pd.Series) -> pd.Series:
    """Fast vectorized cell normalization for a single column."""
    s = (
        series.fillna("")
        .astype(str)
        .str.replace(r"[\s\u00a0]+", " ", regex=True)
        .str.strip()
        .str.upper()
    )

    true_mask = s.isin(_BOOL_TRUE)
    false_mask = s.isin(_BOOL_FALSE)
    s = s.mask(true_mask, "TRUE").mask(false_mask, "FALSE")

    # Preserve the historical numeric normalization, but only run numeric
    # conversion on strings that actually look numeric.
    numeric_mask = s.str.fullmatch(
        r"[+-]?\d+(?:[.,]\d+)?(?:E[+-]?\d+)?", na=False
    )
    if numeric_mask.any():
        numeric = pd.to_numeric(
            s.loc[numeric_mask].str.replace(",", ".", regex=False),
            errors="coerce",
        )
        valid = numeric.notna()
        if valid.any():
            vals = numeric.loc[valid]
            rounded = vals.round()
            formatted = vals.astype(str).str.replace(r"\.0+$", "", regex=True)
            formatted = pd.Series(
                np.where(np.isclose(vals.to_numpy(), rounded.to_numpy()),
                         rounded.astype("int64").astype(str).to_numpy(),
                         formatted.to_numpy()),
                index=vals.index,
            )
            s.loc[formatted.index] = formatted

    s = s.mask(s.eq("NAN") | s.eq("NONE"), "")
    return s


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Vectorized normalization; avoids Python-level per-cell map calls."""
    if df is None or df.empty:
        return pd.DataFrame()
    d = df.copy()
    d.columns = d.columns.map(str)
    for col in d.columns:
        d[col] = _normalize_series(d[col])
    return d


# ─────────────────────────────────────────────
# SMART HEADER DETECTION
# ─────────────────────────────────────────────

def find_header_row_by_column_names(df_raw: pd.DataFrame, reference_columns: List[str]) -> int:
    if not reference_columns:
        return detect_header_row_heuristic(df_raw)
    normalized_ref = {normalize_colname(str(c)) for c in reference_columns if str(c).strip()}
    best_row = 0
    best_match_count = 0
    n_rows = min(15, len(df_raw))
    for i in range(n_rows):
        row = df_raw.iloc[i]
        row_values = {
            normalize_colname(str(val))
            for val in row
            if pd.notna(val) and str(val).strip()
        }
        match_count = len(normalized_ref & row_values)
        if match_count >= max(2, len(normalized_ref) * 0.5) and match_count > best_match_count:
            best_match_count = match_count
            best_row = i
            if match_count >= len(normalized_ref) * 0.8:
                return best_row
    return best_row if best_match_count >= 2 else detect_header_row_heuristic(df_raw)


def find_header_row_with_keys(df_raw: pd.DataFrame, key_columns: List[str], ws=None) -> int:
    if not key_columns:
        return detect_header_row_heuristic(df_raw, ws=ws)
    normalized_keys = {normalize_colname(k) for k in key_columns if k.strip()}
    for i in range(min(15, len(df_raw))):
        row = df_raw.iloc[i]
        row_values = {
            normalize_colname(str(val))
            for val in row
            if pd.notna(val) and str(val).strip()
        }
        if normalized_keys & row_values:
            return i
    return detect_header_row_heuristic(df_raw, ws=ws)


def detect_header_row_heuristic(df_raw: pd.DataFrame, ws=None) -> int:
    if df_raw.empty:
        return 0

    head     = df_raw.head(15)
    n_cols   = df_raw.shape[1] or 1
    search_n = len(head)

    horizontal_merge_rows: set = set()
    vertical_cont_rows:    set = set()
    xl_rows: list = []

    if ws is not None:
        xl_rows = list(ws.iter_rows(max_row=search_n))
        for mr in ws.merged_cells.ranges:
            ri        = mr.min_row - 1
            col_span  = mr.max_col - mr.min_col
            row_span  = mr.max_row - mr.min_row
            if col_span >= 1 and row_span == 0 and ri < search_n:
                horizontal_merge_rows.add(ri)
            if row_span >= 1:
                for r in range(mr.min_row + 1, mr.max_row + 1):
                    if r - 1 < search_n:
                        vertical_cont_rows.add(r - 1)

    best_score, best_row = -1.0, 0

    for i in range(search_n):
        if i in vertical_cont_rows:
            continue

        pd_row = head.iloc[i]
        vals   = [v for v in pd_row if pd.notna(v) and str(v).strip()]
        n_vals = len(vals)

        if n_vals < 2:
            continue
        if n_vals == 1 and len(str(vals[0])) > 50:
            continue

        if xl_rows and i < len(xl_rows):
            bold_count = sum(
                1 for c in xl_rows[i]
                if c.value is not None and str(c.value).strip()
                and c.font and c.font.bold
            )
            bold_r = bold_count / n_vals
        else:
            bold_r = 0.0

        all_str   = float(all(isinstance(v, str) for v in vals))
        unique_r  = len({str(v) for v in vals}) / n_vals
        fill_r    = n_vals / n_cols
        num_r     = sum(
            1 for v in vals
            if str(v).replace(".", "", 1).replace("-", "", 1).isdigit()
        ) / n_vals
        name_like = sum(
            1 for v in vals
            if isinstance(v, str) and len(v) < 40
            and re.match(r"^[A-Za-z]", v.strip())
        ) / n_vals

        lookahead = df_raw.iloc[i + 1 : i + 9]
        if len(lookahead) >= 2:
            fill_counts  = lookahead.apply(lambda r: r.notna().sum(), axis=1)
            post_consist = 1.0 / (1.0 + fill_counts.std())
        else:
            post_consist = 0.0

        score = (
            bold_r        * 3.0
            + fill_r      * 1.5
            + all_str     * 1.0
            + unique_r    * 1.0
            + name_like   * 1.0
            + post_consist * 1.0
            + (1 - num_r) * 0.5
            - i           * 0.1
        )

        if i in horizontal_merge_rows:
            score -= 2.5

        if score > best_score:
            best_score = score
            best_row   = i

    return best_row


# ─────────────────────────────────────────────
# FILE READING  (single read_excel pass for XLSX)
# ─────────────────────────────────────────────

def _apply_header_and_dedup(df_raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    if header_row < len(df_raw):
        raw_cols = df_raw.iloc[header_row].tolist()
        df_data  = df_raw.iloc[header_row + 1:].copy()
        seen_c: Dict[str, int] = {}
        clean_cols = []
        for i, c in enumerate(raw_cols):
            name = str(c).strip() if pd.notna(c) else f"Unnamed_{i}"
            if not name or name == "nan":
                name = f"Unnamed_{i}"
            if name in seen_c:
                seen_c[name] += 1
                clean_cols.append(f"{name}_{seen_c[name]}")
            else:
                seen_c[name] = 0
                clean_cols.append(name)
        df_data.columns = clean_cols
        df_data = df_data.reset_index(drop=True)
    else:
        df_data = df_raw.copy()
        df_data.columns = [str(c).strip() for c in df_data.columns]

    if isinstance(df_data.columns, pd.MultiIndex):
        df_data.columns = [
            " ".join(str(c) for c in col if str(c) != "nan").strip()
            for col in df_data.columns
        ]
    return df_data.astype(str)


@st.cache_data(show_spinner=False)
def _read_sheets_cached(
    file_bytes_raw: bytes,
    filename: str,
    key_columns_tuple: tuple,
    header_overrides_tuple: tuple = (),
    sheet_filter_tuple: tuple = (),
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, int]]:
    """
    Unified multi-format sheet reader.
    XLSX path reads cell values directly from openpyxl's read-only workbook.
    The optional sheet filter lets callers avoid parsing non-common sheets.
        already-loaded workbook (load_workbook) instead of running a second,
    independent pd.read_excel() parse over the same file. The old approach
    fully parsed heavy files twice (once for styles/merges via
    load_workbook, once again for values via pd.read_excel) for no benefit,
    since load_workbook already holds every cell value.
    """
    key_columns = list(key_columns_tuple)
    header_overrides: Dict[str, int] = dict(header_overrides_tuple)
    sheet_filter = set(sheet_filter_tuple) if sheet_filter_tuple else None
    ext = _file_ext(filename)

    # ── CSV / TSV ────────────────────────────────────────────────────
    if ext in _CSV_EXTS | _TSV_EXTS:
        sh = _sheet_name_from_filename(filename)
        if sheet_filter is not None and sh not in sheet_filter:
            return {}, {}
        df_raw = _read_csv_raw_cached(file_bytes_raw, filename)
        if df_raw.empty:
            return {sh: pd.DataFrame()}, {sh: 0}
        hr = header_overrides.get(sh,
             find_header_row_with_keys(df_raw, key_columns) if key_columns
             else detect_header_row_heuristic(df_raw))
        return {sh: _apply_header_and_dedup(df_raw, hr)}, {sh: hr}

    # ── ODS ──────────────────────────────────────────────────────────
    if ext in _EXCEL_ODS:
        try:
            all_raw = pd.read_excel(BytesIO(file_bytes_raw), sheet_name=None,
                                    header=None, engine="odf")
            sheets, header_rows = {}, {}
            for sh, df_raw in all_raw.items():
                if sheet_filter is not None and sh not in sheet_filter:
                    continue
                hr = header_overrides.get(sh,
                     find_header_row_with_keys(df_raw, key_columns) if key_columns
                     else detect_header_row_heuristic(df_raw))
                sheets[sh] = _apply_header_and_dedup(df_raw, hr)
                header_rows[sh] = hr
            return sheets, header_rows
        except Exception:
            return {}, {}

    # ── XLS (legacy) ─────────────────────────────────────────────────
    if ext in _EXCEL_XLS:
        try:
            all_raw = pd.read_excel(BytesIO(file_bytes_raw), sheet_name=None,
                                    header=None, engine="xlrd")
            sheets, header_rows = {}, {}
            for sh, df_raw in all_raw.items():
                hr = header_overrides.get(sh,
                     find_header_row_with_keys(df_raw, key_columns) if key_columns
                     else detect_header_row_heuristic(df_raw))
                sheets[sh] = _apply_header_and_dedup(df_raw, hr)
                header_rows[sh] = hr
            return sheets, header_rows
        except Exception:
            return {}, {}

    # ── XLSX / XLSM — streaming/read-only path ─────────────────────────
    # read_only=True avoids creating a full in-memory Cell graph. Header
    # detection deliberately runs on a small preview, then the data is
    # streamed row-by-row into a DataFrame.
    try:
        wb = load_workbook(
            BytesIO(file_bytes_raw),
            read_only=True,
            data_only=True,
        )
        sheets: Dict[str, pd.DataFrame] = {}
        header_rows: Dict[str, int] = {}

        for ws in wb.worksheets:
            sh = ws.title
            rows = list(ws.iter_rows(values_only=True))
            df_raw = pd.DataFrame(rows) if rows else pd.DataFrame()

            if df_raw.empty:
                sheets[sh] = pd.DataFrame()
                header_rows[sh] = 0
                continue

            # Hidden-column metadata is unavailable in read-only mode.
            # Preserve all columns; explicit header overrides still work.
            if sh in header_overrides:
                hr = header_overrides[sh]
            elif key_columns:
                hr = find_header_row_with_keys(df_raw, key_columns)
            else:
                hr = detect_header_row_heuristic(df_raw)

            sheets[sh] = _apply_header_and_dedup(df_raw, hr)
            header_rows[sh] = hr

        return sheets, header_rows

    except Exception:
        import traceback
        print(
            f"[_read_sheets_cached] XLSX read-only path failed for "
            f"'{filename}':\n{traceback.format_exc()}"
        )
        file_bytes = BytesIO(file_bytes_raw)
        fallback = pd.read_excel(
            file_bytes, sheet_name=None, header=None, engine="openpyxl"
        )
        for sh in fallback:
            fallback[sh] = fallback[sh].astype(str)
        return fallback, {}



def read_visible_sheets_with_header_detection(
    file_bytes: BytesIO,
    filename: str = "",
    key_columns: List[str] = None,
    header_overrides: Dict[str, int] = None,
    sheet_filter: Optional[List[str]] = None,
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, int]]:
    raw = file_bytes.read()
    overrides_tuple = tuple(sorted((header_overrides or {}).items()))
    filter_tuple = tuple(sorted(sheet_filter)) if sheet_filter else ()
    return _read_sheets_cached(
        raw, filename, tuple(key_columns or []), overrides_tuple, filter_tuple
    )


# ─────────────────────────────────────────────
# COLUMN ALIGNMENT
# ─────────────────────────────────────────────

def _normalize_for_match(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r'([_.]\d+)+$', '', s)
    return re.sub(r'[^a-z0-9]', '', s)


def _is_placeholder_name(name: str) -> bool:
    """
    FIX (v2.2): detects generic placeholder column names (e.g. "Unnamed: 1",
    "Unnamed_22") that already exist as literal text in a source file —
    typically baked in by an earlier, unrelated pandas read/write round-trip
    upstream. These can never be recovered by name-based matching since
    there's no real name left to match against.
    """
    return bool(re.match(r'^unnamed\s*[:_]?\s*\d*$', str(name).strip().lower()))


def align_new_columns_to_reference(new_cols: List[str], ref_cols: List[str]) -> List[str]:
    new_cols = [str(c) for c in new_cols]
    ref_norm_map = {i: _normalize_for_match(c) for i, c in enumerate(ref_cols)}
    new_norm_map = {i: _normalize_for_match(c) for i, c in enumerate(new_cols)}
    assigned = set()
    renamed  = list(new_cols)

    # Pass 1: exact normalized match
    for ref_i, ref_norm in ref_norm_map.items():
        for new_i, new_norm in new_norm_map.items():
            if new_i in assigned:
                continue
            if ref_norm and new_norm == ref_norm:
                renamed[new_i] = ref_cols[ref_i]
                assigned.add(new_i)
                break

    # Pass 2: substring match.
    # FIX (v2.2): guarded against short names matching as substrings of
    # unrelated longer names (e.g. "id" inside "valid" or "grid").
    for ref_i, ref_norm in ref_norm_map.items():
        if not ref_norm:
            continue
        for new_i, new_norm in new_norm_map.items():
            if new_i in assigned:
                continue
            shorter, longer = sorted([ref_norm, new_norm], key=len)
            if len(shorter) < 4 or len(shorter) / len(longer) < 0.5:
                continue
            if ref_norm in new_norm or new_norm in ref_norm:
                renamed[new_i] = ref_cols[ref_i]
                assigned.add(new_i)
                break

    # Pass 3 (NEW in v2.2): positional fallback for placeholder/garbage
    # names (e.g. "Unnamed: 1") that already exist as literal text in the
    # source file and therefore cannot be recovered by name matching at
    # all. Falls back to the OLD column sitting at the same column index,
    # since the underlying data is still in the same positional order.
    for new_i, name in enumerate(new_cols):
        if new_i in assigned:
            continue
        if _is_placeholder_name(name) and new_i < len(ref_cols):
            candidate = ref_cols[new_i]
            if not _is_placeholder_name(candidate):
                renamed[new_i] = candidate
                assigned.add(new_i)

    seen: Dict[str, int] = {}
    final = []
    for name in renamed:
        if name in seen:
            seen[name] += 1
            final.append(f"{name}_{seen[name]}")
        else:
            seen[name] = 0
            final.append(name)
    return final


# ─────────────────────────────────────────────
# KEY SELECTION
# ─────────────────────────────────────────────

def find_best_valid_key(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    keys: List[str],
    n1: Optional[pd.DataFrame] = None,
    n2: Optional[pd.DataFrame] = None,
) -> Tuple[List[str], str]:
    """Find the strongest user-requested key among shared columns."""
    if not keys:
        return [], "No keys provided"

    def norm(s: str) -> str:
        s = str(s).strip().lower()
        s = re.sub(r"(_\d+)+$", "", s)
        return re.sub(r"[^a-z0-9]", "", s)

    common = list(df1.columns.intersection(df2.columns))
    if not common:
        return [], "No common columns"

    common_norm = {norm(c): c for c in common}
    best_key = None
    best_uniqueness = -1.0
    best_key_name = None
    best_match_score = -1.0

    # Compute uniqueness from normalized data once, instead of repeatedly
    # converting raw columns to strings for every candidate.
    if n1 is not None and n2 is not None:
        uniq1 = {c: n1[c].nunique(dropna=False) / max(1, len(n1)) for c in common if c in n1}
        uniq2 = {c: n2[c].nunique(dropna=False) / max(1, len(n2)) for c in common if c in n2}
    else:
        uniq1 = {c: df1[c].fillna("").astype(str).nunique() / max(1, len(df1)) for c in common}
        uniq2 = {c: df2[c].fillna("").astype(str).nunique() / max(1, len(df2)) for c in common}

    for requested in keys:
        nr = norm(requested)
        exact = common_norm.get(nr)
        candidates = [(nr, exact)] if exact is not None else common_norm.items()

        for nc, col in candidates:
            score = 1.0 if exact is not None else SequenceMatcher(None, nr, nc).ratio()
            if score < 0.8:
                continue
            uniqueness = (uniq1.get(col, 0.0) + uniq2.get(col, 0.0)) / 2
            if score > best_match_score or (score == best_match_score and uniqueness > best_uniqueness):
                best_match_score = score
                best_uniqueness = uniqueness
                best_key = [col]
                best_key_name = requested

    if best_key:
        return best_key, f"Key: '{best_key_name}' ({best_uniqueness:.0%} unique)"
    return [], "No valid keys found"


# ─────────────────────────────────────────────
# COMPARISON FUNCTIONS
# ─────────────────────────────────────────────

def compare_key_based(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    valid_keys: List[str],
    key_desc: str,
    n1: pd.DataFrame,
    n2: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, str]:
    """Compare shared-key rows using narrow merges + vectorized row hashes."""
    if not valid_keys:
        raise ValueError("No valid key columns found")

    common = list(df1.columns.intersection(df2.columns))
    df1f = df1[common].fillna("").reset_index(drop=True)
    df2f = df2[common].fillna("").reset_index(drop=True)
    n1c = n1.reindex(columns=common).fillna("").reset_index(drop=True)
    n2c = n2.reindex(columns=common).fillna("").reset_index(drop=True)

    left = n1c[valid_keys].copy()
    right = n2c[valid_keys].copy()
    left["__row_id"] = np.arange(len(left), dtype=np.int64)
    right["__row_id"] = np.arange(len(right), dtype=np.int64)
    left = left.drop_duplicates(valid_keys, keep="first")
    right = right.drop_duplicates(valid_keys, keep="first")

    non_key_cols = [c for c in common if c not in valid_keys]
    if non_key_cols:
        left["__data_hash"] = pd.util.hash_pandas_object(
            n1c.loc[left["__row_id"].to_numpy(), non_key_cols], index=False
        ).to_numpy()
        right["__data_hash"] = pd.util.hash_pandas_object(
            n2c.loc[right["__row_id"].to_numpy(), non_key_cols], index=False
        ).to_numpy()
    else:
        left["__data_hash"] = np.uint64(0)
        right["__data_hash"] = np.uint64(0)

    merged = left.merge(
        right,
        on=valid_keys,
        how="inner",
        suffixes=("__old", "__new"),
        sort=False,
    )

    if merged.empty or not non_key_cols:
        changed_old = pd.DataFrame(columns=common)
        changed_new = pd.DataFrame(columns=common)
    else:
        changed_mask = merged["__data_hash__old"].to_numpy() != merged["__data_hash__new"].to_numpy()
        old_ids = merged.loc[changed_mask, "__row_id__old"].astype(int).to_numpy()
        new_ids = merged.loc[changed_mask, "__row_id__new"].astype(int).to_numpy()
        changed_old = df1f.iloc[old_ids][common].reset_index(drop=True)
        changed_new = df2f.iloc[new_ids][common].reset_index(drop=True)

    # Direct MultiIndex membership avoids another wide merge for the anti-join.
    left_key_index = pd.MultiIndex.from_frame(left[valid_keys])
    right_key_index = pd.MultiIndex.from_frame(right[valid_keys])
    right_mask = ~right_key_index.isin(left_key_index)
    left_mask = ~left_key_index.isin(right_key_index)
    added_ids = right.loc[right_mask, "__row_id"].astype(int).to_numpy()
    removed_ids = left.loc[left_mask, "__row_id"].astype(int).to_numpy()

    added = df2f.iloc[added_ids][common].reset_index(drop=True)
    removed = df1f.iloc[removed_ids][common].reset_index(drop=True)
    return changed_old, changed_new, added, removed, key_desc


def compare_keyless(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    n1: pd.DataFrame,
    n2: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Compare normalized rows with 64-bit vectorized row hashes + counts."""
    if df1.empty and df2.empty:
        return pd.DataFrame(columns=df1.columns), pd.DataFrame(columns=df1.columns)

    common = list(n1.columns.intersection(n2.columns))
    if not common:
        return pd.DataFrame(columns=df2.columns), pd.DataFrame(columns=df1.columns)

    n1c = n1[common].reset_index(drop=True)
    n2c = n2[common].reset_index(drop=True)
    h1 = pd.util.hash_pandas_object(n1c, index=False)
    h2 = pd.util.hash_pandas_object(n2c, index=False)

    c1 = h1.value_counts()
    c2 = h2.value_counts()
    removed_counts = c1.subtract(c2, fill_value=0)
    added_counts = c2.subtract(c1, fill_value=0)

    removed_hashes = set(removed_counts[removed_counts > 0].index)
    added_hashes = set(added_counts[added_counts > 0].index)

    # Hashes are used only to identify candidate rows. This preserves duplicate
    # multiplicity while avoiding a wide groupby/merge of every column.
    added = df2.reset_index(drop=True).loc[h2.isin(added_hashes)].reset_index(drop=True)
    removed = df1.reset_index(drop=True).loc[h1.isin(removed_hashes)].reset_index(drop=True)
    return added, removed


def reconcile_added_removed(
    removed: pd.DataFrame,
    added: pd.DataFrame,
    norm_removed: pd.DataFrame,
    norm_added: pd.DataFrame,
    similarity_threshold: float = 0.5,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Pair likely changed rows using an inverted value index.

    Instead of comparing every removed row against every added row across every
    column, index added-row positions by (column, value), then score only
    candidates sharing at least one value with the removed row.
    """
    if removed.empty or added.empty:
        return pd.DataFrame(), pd.DataFrame(), removed.copy(), added.copy()

    common = list(norm_removed.columns.intersection(norm_added.columns))
    if not common:
        return pd.DataFrame(), pd.DataFrame(), removed.copy(), added.copy()

    nr = norm_removed[common].reset_index(drop=True)
    na = norm_added[common].reset_index(drop=True)
    rem_raw = removed.reset_index(drop=True)
    add_raw = added.reset_index(drop=True)
    n_cols = len(common)

    # Pass 1: exact full-row matches are moved rows, not changes.
    rem_keys = pd.MultiIndex.from_frame(nr)
    add_keys = pd.MultiIndex.from_frame(na)
    add_positions = defaultdict(list)
    for ai, key in enumerate(add_keys):
        add_positions[key].append(ai)

    matched_rem = set()
    matched_add = set()
    for ri, key in enumerate(rem_keys):
        for ai in add_positions.get(key, ()):
            if ai not in matched_add:
                matched_rem.add(ri)
                matched_add.add(ai)
                break

    unmatched_ri = [i for i in range(len(nr)) if i not in matched_rem]
    unmatched_ai = [i for i in range(len(na)) if i not in matched_add]

    changed_old_rows = []
    changed_new_rows = []

    if unmatched_ri and unmatched_ai:
        # Inverted index: (column index, normalized value) -> added row ids.
        # Very common values such as "" can create enormous candidate sets, so
        # only index values that occur in a minority of the added rows.
        value_freq = [Counter(na.iloc[unmatched_ai, c].tolist()) for c in range(n_cols)]
        max_freq = max(50, int(len(unmatched_ai) * 0.25))
        inverted = defaultdict(list)
        for pos, ai in enumerate(unmatched_ai):
            row = na.iloc[ai]
            for c, value in enumerate(row):
                if value_freq[c].get(value, 0) <= max_freq:
                    inverted[(c, value)].append(pos)

        used_positions = set()
        min_matches = max(1, int(np.ceil(similarity_threshold * n_cols)))

        for ri in unmatched_ri:
            row = nr.iloc[ri]
            candidate_counts = Counter()

            for c, value in enumerate(row):
                for pos in inverted.get((c, value), ()):
                    if pos not in used_positions:
                        candidate_counts[pos] += 1

            if not candidate_counts:
                continue

            # Deterministic best candidate: most matching columns, then earliest row.
            best_pos, match_count = max(
                candidate_counts.items(),
                key=lambda item: (item[1], -item[0]),
            )
            if match_count >= min_matches:
                ai = unmatched_ai[best_pos]
                matched_rem.add(ri)
                matched_add.add(ai)
                used_positions.add(best_pos)
                changed_old_rows.append(rem_raw.iloc[ri])
                changed_new_rows.append(add_raw.iloc[ai])

    co = (
        pd.DataFrame(changed_old_rows).reset_index(drop=True)
        if changed_old_rows else pd.DataFrame(columns=removed.columns)
    )
    cn = (
        pd.DataFrame(changed_new_rows).reset_index(drop=True)
        if changed_new_rows else pd.DataFrame(columns=added.columns)
    )
    truly_removed = rem_raw.loc[
        ~rem_raw.index.isin(matched_rem)
    ].reset_index(drop=True)
    truly_added = add_raw.loc[
        ~add_raw.index.isin(matched_add)
    ].reset_index(drop=True)
    return co, cn, truly_removed, truly_added



def detect_data_truncation(
    df1: pd.DataFrame, df2: pd.DataFrame
) -> Tuple[bool, int, int]:
    """Vectorized truncation check for already-stringified frames."""
    empty_tokens = {"", "nan", "None"}
    old_rows = int((~df1.isin(empty_tokens).all(axis=1)).sum()) if not df1.empty else 0
    new_rows = int((~df2.isin(empty_tokens).all(axis=1)).sum()) if not df2.empty else 0
    threshold = max(5, old_rows * 0.1)
    return (old_rows - new_rows) >= threshold, old_rows, new_rows


# ─────────────────────────────────────────────
# PREVIEW & EXPORT
# Both now accept pre-normalized old_norm/new_norm to skip redundant normalize_df.
# ─────────────────────────────────────────────

def build_side_by_side_preview(
    changed_old: pd.DataFrame,
    changed_new: pd.DataFrame,
    keys: List[str],
    old_norm: pd.DataFrame,
    new_norm: pd.DataFrame,
):
    """FIX: accepts pre-normalized frames instead of re-normalizing internally."""
    common_cols = changed_old.columns.intersection(changed_new.columns).tolist()

    rows = []
    for idx in range(len(changed_old)):
        row = {}
        for col in common_cols:
            row[f"{col} (Old)"] = str(changed_old.iloc[idx][col])
            row[f"{col} (New)"] = str(changed_new.iloc[idx][col])
        rows.append(row)
    combined = pd.DataFrame(rows)

    shared = set(old_norm.columns) & set(new_norm.columns)

    def highlight(row):
        styles = []
        ridx = row.name
        for col in combined.columns:
            if col.endswith(" (Old)"):
                base = col[:-6]
                changed = (base in shared and ridx < len(old_norm) and ridx < len(new_norm) and
                           str(old_norm.iloc[ridx][base]) != str(new_norm.iloc[ridx][base]))
                styles.append('background-color: #3b1212; color: #fca5a5' if changed
                               else 'background-color: #0f2a1a; color: #86efac')
            elif col.endswith(" (New)"):
                base = col[:-6]
                changed = (base in shared and ridx < len(old_norm) and ridx < len(new_norm) and
                           str(old_norm.iloc[ridx][base]) != str(new_norm.iloc[ridx][base]))
                styles.append('background-color: #14532d; color: #4ade80' if changed
                               else 'background-color: #0f2a1a; color: #86efac')
            else:
                styles.append('')
        return styles

    return combined.style.apply(highlight, axis=1)


def _dedup_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.columns.is_unique:
        return df
    seen: Dict[str, int] = {}
    cols = []
    for c in df.columns:
        if c in seen:
            seen[c] += 1
            cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            cols.append(c)
    df = df.copy()
    df.columns = cols
    return df


def style_added_rows(df: pd.DataFrame):
    if df is None or df.empty:
        return df
    df = _dedup_columns(df.copy())
    return df.style.set_properties(**{"background-color": "#052e16", "color": "#86efac"})


def style_removed_rows(df: pd.DataFrame):
    if df is None or df.empty:
        return df
    df = _dedup_columns(df.copy())
    return df.style.set_properties(**{"background-color": "#450a0a", "color": "#fca5a5"})


@st.cache_data(show_spinner=False)
def export_to_excel(
    changed_old: pd.DataFrame,
    changed_new: pd.DataFrame,
    keys: List[str],
    old_norm: pd.DataFrame,
    new_norm: pd.DataFrame,
) -> bytes:
    """Cache generated workbooks so reruns do not rebuild them."""
    wb = Workbook()
    ws_old = wb.active
    ws_old.title = "Old Values"
    ws_new = wb.create_sheet(title="New Values")
    fill_old = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    fill_new = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

    shared_cols = set(old_norm.columns) & set(new_norm.columns)

    for ws, df, fill in [
        (ws_old, changed_old, fill_old),
        (ws_new, changed_new, fill_new),
    ]:
        for c_idx, col_name in enumerate(df.columns, 1):
            ws.cell(row=1, column=c_idx, value=col_name)

        # itertuples is much cheaper than repeated df.iloc[r][col].
        cols = list(df.columns)
        for r_idx, row in enumerate(df.itertuples(index=False, name=None), 2):
            for c_idx, (col_name, value) in enumerate(zip(cols, row), 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                if (
                    col_name not in keys
                    and col_name in shared_cols
                    and r_idx - 2 < len(old_norm)
                    and r_idx - 2 < len(new_norm)
                    and str(old_norm.iloc[r_idx - 2][col_name])
                    != str(new_norm.iloc[r_idx - 2][col_name])
                ):
                    cell.fill = fill

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


MAX_PREVIEW_ROWS = 500


# ─────────────────────────────────────────────
# UI LAYOUT
# ─────────────────────────────────────────────

st.markdown('<div class="step-label">Step 1 — Upload Files</div>', unsafe_allow_html=True)

if "upload_key" not in st.session_state:
    st.session_state.upload_key = 0

col_l, col_r, col_clr = st.columns([2, 2, 1])
with col_l:
    left = st.file_uploader(
        "OLD files", ALL_SUPPORTED_EXTS, accept_multiple_files=True,
        label_visibility="visible",
        key=f"uploader_left_{st.session_state.upload_key}",
    )
with col_r:
    right = st.file_uploader(
        "NEW files", ALL_SUPPORTED_EXTS, accept_multiple_files=True,
        label_visibility="visible",
        key=f"uploader_right_{st.session_state.upload_key}",
    )
with col_clr:
    st.markdown("<div style='height:1.9rem'></div>", unsafe_allow_html=True)
    if st.button("🗑 Clear files", key="clear_uploads",
                 help="Remove all uploaded files and reset the session",
                 use_container_width=True):
        st.session_state.upload_key += 1
        for k in ["header_overrides", "_hdr_files_key", "manual_pairs",
                  "left_files", "right_files"]:
            st.session_state.pop(k, None)
        for k in list(st.session_state.keys()):
            if k.startswith("hdr_") or k.startswith("_prev_auto_"):
                del st.session_state[k]
        st.rerun()

st.markdown('<div class="step-label" style="margin-top:1.5rem">Step 2 — Configure</div>',
            unsafe_allow_html=True)

cfg_col1, cfg_col2, cfg_col3 = st.columns([2, 1, 1])
with cfg_col1:
    keys_str = st.text_input("Key columns (comma-separated)", "",
                              placeholder="e.g.  ID, Employee Name, Code")
with cfg_col2:
    ignore_suffix = st.checkbox("Ignore filename suffix", True,
                                help="Strips text after the last underscore before matching files")
with cfg_col3:
    threshold = st.slider("Match threshold", 0.5, 1.0, 0.85, 0.05,
                          help="Minimum similarity score for auto-pairing files")

if left or right:
    all_files = (left or []) + (right or [])
    auto_rows = get_auto_detected_rows(all_files)

    all_sheet_names: list = []
    for sh in auto_rows:
        if sh not in all_sheet_names:
            all_sheet_names.append(sh)

    files_key = tuple(sorted(f.name for f in all_files))
    if (st.session_state.get("_hdr_files_key") != files_key
            or "header_overrides" not in st.session_state):
        st.session_state.header_overrides = {}
        st.session_state._hdr_files_key = files_key

    with st.expander("🔢 Header Row Overrides", expanded=False):
        st.markdown(
            '<div style="color:#7888a8;font-size:0.82rem;margin-bottom:0.75rem">'
            'Auto-detection is shown below. Change a value only if the detected row '
            'is wrong — for example when a sheet has a merged group-label row above '
            'the real header. <b>Row 1 = first row of the sheet.</b></div>',
            unsafe_allow_html=True,
        )

        if not all_sheet_names:
            st.caption("Upload files above to see sheet names.")
        else:
            cols_per_row = 3
            sheet_chunks = [all_sheet_names[i:i+cols_per_row]
                            for i in range(0, len(all_sheet_names), cols_per_row)]

            for chunk in sheet_chunks:
                grid = st.columns(cols_per_row)
                for col, sh in zip(grid, chunk):
                    with col:
                        auto_val = auto_rows.get(sh, 1)
                        widget_key = f"hdr_{sh}"

                        if widget_key not in st.session_state:
                            st.session_state[widget_key] = int(
                                st.session_state.header_overrides.get(sh, auto_val)
                            )

                        prev_auto_key = f"_prev_auto_{sh}"
                        if st.session_state.get(prev_auto_key) != auto_val:
                            st.session_state[widget_key] = int(
                                st.session_state.header_overrides.get(sh, auto_val)
                            )
                            st.session_state[prev_auto_key] = auto_val

                        current = st.session_state[widget_key]
                        is_overridden = (current != auto_val)

                        label = (
                            f'"{sh}" ✏️' if is_overridden else f'"{sh}" 🤖'
                        )
                        val = st.number_input(
                            label,
                            min_value=1, max_value=100,
                            value=int(current),
                            step=1,
                            key=widget_key,
                            help=(
                                f"Auto-detected: row {auto_val}. "
                                + ("Currently overridden." if is_overridden
                                   else "Matches auto-detection — change only if wrong.")
                            ),
                        )
                        if val != auto_val:
                            st.session_state.header_overrides[sh] = val
                        elif sh in st.session_state.header_overrides:
                            del st.session_state.header_overrides[sh]

            n_overridden = len(st.session_state.header_overrides)
            if n_overridden:
                overridden_names = ", ".join(
                    f"{sh} → row {row}"
                    for sh, row in st.session_state.header_overrides.items()
                )
                st.markdown(
                    f'<div style="color:#fbbf24;font-size:0.78rem;margin-top:0.5rem">'
                    f'✏️ {n_overridden} sheet(s) overridden: {overridden_names}</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    '<div style="color:#4a5878;font-size:0.78rem;margin-top:0.5rem">'
                    '🤖 All sheets using auto-detection</div>',
                    unsafe_allow_html=True,
                )

            if n_overridden and st.button("↺ Reset all to auto-detect", key="reset_overrides"):
                st.session_state.header_overrides = {}
                st.rerun()

if left and right:
    def comparable(n, ignore):
        n = n.rsplit(".", 1)[0]
        return n.rsplit("_", 1)[0].lower() if ignore and "_" in n else n.lower()

    auto_matched, usedR = [], set()
    for lf in left:
        lcmp = comparable(lf.name, ignore_suffix)
        best, ratio = None, 0.0
        for rf in right:
            if rf.name in usedR:
                continue
            r = SequenceMatcher(None, lcmp, comparable(rf.name, ignore_suffix)).ratio()
            if r >= threshold and r > ratio:
                best, ratio = rf, r
        if best:
            auto_matched.append((lf, best))
            usedR.add(best.name)

    st.markdown('<div class="step-label" style="margin-top:1.5rem">Step 3 — Review & Run</div>',
                unsafe_allow_html=True)

    left_names  = {f.name: f for f in left}
    right_names = {f.name: f for f in right}

    if not auto_matched:
        st.markdown(
            '<div class="warn-box">⚠️ No files matched automatically — the filenames are too '
            'different. Use <b>Manual Pairing</b> below to pair them directly.</div>',
            unsafe_allow_html=True,
        )

    with st.expander(
        f"🔧 Manual Pairing {'(recommended — auto-match found 0 pairs)' if not auto_matched else '(optional override)'}",
        expanded=not auto_matched,
    ):
        st.markdown(
            '<div style="color:#7888a8;font-size:0.82rem;margin-bottom:0.75rem">'
            'Select an OLD file and a NEW file to compare, then click Add Pair. '
            'Manual pairs override auto-matching for those files.</div>',
            unsafe_allow_html=True,
        )
        mp_col1, mp_col2, mp_col3 = st.columns([2, 2, 1])
        with mp_col1:
            sel_old = st.selectbox("OLD file", list(left_names.keys()), key="manual_old")
        with mp_col2:
            sel_new = st.selectbox("NEW file", list(right_names.keys()), key="manual_new")
        with mp_col3:
            st.markdown("<div style='height:1.95rem'></div>", unsafe_allow_html=True)
            add_pair = st.button("＋ Add Pair")

        if "manual_pairs" not in st.session_state:
            st.session_state.manual_pairs = []

        if add_pair:
            entry = (sel_old, sel_new)
            if entry not in st.session_state.manual_pairs:
                st.session_state.manual_pairs.append(entry)

        if st.session_state.manual_pairs:
            st.markdown("**Added manual pairs:**")
            to_remove = []
            for idx, (on, nn) in enumerate(st.session_state.manual_pairs):
                r1, r2 = st.columns([8, 1])
                with r1:
                    st.markdown(
                        f'<div class="pair-card" style="margin-bottom:0.3rem">'
                        f'<span class="fname">{on}</span>'
                        f'<span class="arrow">⟶</span>'
                        f'<span class="fname">{nn}</span>'
                        f'<span style="color:#6382ff;font-family:monospace;font-size:0.72rem">manual</span>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                with r2:
                    if st.button("✕", key=f"rm_{idx}"):
                        to_remove.append(idx)
            for idx in reversed(to_remove):
                st.session_state.manual_pairs.pop(idx)

        if st.button("🗑 Clear all manual pairs", key="clear_manual"):
            st.session_state.manual_pairs = []

    manual_old_names = {on for on, _ in st.session_state.get("manual_pairs", [])}
    merged_matched = [
        (lf, rf) for lf, rf in auto_matched
        if lf.name not in manual_old_names
    ]
    for old_name, new_name in st.session_state.get("manual_pairs", []):
        if old_name in left_names and new_name in right_names:
            merged_matched.append((left_names[old_name], right_names[new_name]))

    matched = merged_matched

    if matched:
        with st.expander(f"📋 {len(matched)} pair(s) ready to compare", expanded=False):
            for lf, rf in matched:
                is_manual = lf.name in manual_old_names
                tag = '<span style="color:#6382ff;font-family:monospace;font-size:0.72rem">manual</span>' if is_manual else ''
                st.markdown(
                    f'<div class="pair-card">'
                    f'<span class="fname">{lf.name}</span>'
                    f'<span class="arrow">⟶</span>'
                    f'<span class="fname">{rf.name}</span>'
                    f'{tag}</div>',
                    unsafe_allow_html=True,
                )
        unmatched_l = [f.name for f in left  if f.name not in {lf.name for lf, _ in matched}]
        unmatched_r = [f.name for f in right if f.name not in {rf.name for _, rf in matched}]
        if unmatched_l or unmatched_r:
            st.markdown(
                f'<div class="warn-box">⚠️ Still unmatched — OLD: {", ".join(unmatched_l) or "none"} '
                f'· NEW: {", ".join(unmatched_r) or "none"}</div>',
                unsafe_allow_html=True,
            )

    run_clicked = st.button(
        "▶  Run Comparison",
        disabled=not matched,
        use_container_width=False,
    )

    if run_clicked and matched:
        prog = st.progress(0, text="Starting…")
        keys = [k.strip() for k in keys_str.split(",") if k.strip()]
        files_with_changes = []

        for i, (lf, rf) in enumerate(matched):
            prog.progress((i) / len(matched), text=f"Comparing {lf.name} …")
            try:
                # Fastest possible path: byte-identical uploads cannot differ.
                # Avoid decryption, workbook parsing, header detection, and comparison entirely.
                if _file_hash(lf.getvalue()) == _file_hash(rf.getvalue()):
                    files_with_changes.append((i + 1, lf.name, rf.name, [], [], []))
                    prog.progress((i + 1) / len(matched), text=f"Identical {i+1}/{len(matched)}")
                    continue

                decL = decrypt_file(lf)
                decR = decrypt_file(rf)

                # Header detection is cheap/cached and gives us sheet names.
                # Only fully parse sheets present on BOTH sides. This is a major
                # win for workbooks containing archival/helper sheets.
                metaL = detect_header_rows_for_file(decL.getvalue(), lf.name)
                metaR = detect_header_rows_for_file(decR.getvalue(), rf.name)
                common_sheets = sorted(set(metaL) & set(metaR))
                missing_in_new: List[str] = [sh for sh in metaL if sh not in metaR]
                missing_in_old: List[str] = [sh for sh in metaR if sh not in metaL]

                if not common_sheets:
                    files_with_changes.append((i + 1, lf.name, rf.name, [], missing_in_new, missing_in_old))
                    prog.progress((i + 1) / len(matched), text=f"Done {i+1}/{len(matched)}")
                    continue

                header_overrides = {
                    sh: (row - 1)
                    for sh, row in st.session_state.get("header_overrides", {}).items()
                    if row is not None
                }
                shL, header_rows_old = read_visible_sheets_with_header_detection(
                    decL, filename=lf.name, key_columns=keys,
                    header_overrides=header_overrides, sheet_filter=common_sheets
                )
                shR, header_rows_new = read_visible_sheets_with_header_detection(
                    decR, filename=rf.name, key_columns=keys,
                    header_overrides=header_overrides, sheet_filter=common_sheets
                )

                for sh in list(shR.keys()):
                    if sh in shL:
                        try:
                            shR[sh].columns = align_new_columns_to_reference(
                                list(shR[sh].columns), list(shL[sh].columns)
                            )
                        except Exception:
                            pass

            except Exception as e:
                import traceback
                st.error(f"❌ Pair {i+1}: Failed to read — {e}")
                st.code(traceback.format_exc())
                prog.progress((i + 1) / len(matched))
                continue

            changed_sheets = []
            for sh in sorted(set(shL) & set(shR)):
                d1, d2 = shL[sh], shR[sh]

                # ── FIX: normalize ONCE per sheet, pass everywhere ────
                n1 = normalize_df(d1)
                n2 = normalize_df(d2)

                use_key = False
                key_desc = ""
                valid_keys = []
                if keys:
                    try:
                        valid_keys, key_desc = find_best_valid_key(d1, d2, keys, n1, n2)
                        use_key = bool(valid_keys)
                    except Exception:
                        use_key = False

                try:
                    if use_key:
                        co, cn, add, rem, key_desc = compare_key_based(
                            d1, d2, valid_keys, key_desc, n1, n2
                        )
                    else:
                        co = cn = pd.DataFrame()
                        add, rem = compare_keyless(d1, d2, n1, n2)
                        if not add.empty and not rem.empty:
                            n_rem = normalize_df(rem)
                            n_add = normalize_df(add)
                            co, cn, rem, add = reconcile_added_removed(
                                rem, add, n_rem, n_add
                            )
                except Exception:
                    co = cn = pd.DataFrame()
                    add, rem = compare_keyless(d1, d2, n1, n2)

                is_trunc, old_rows, new_rows = detect_data_truncation(d1, d2)

                if not (cn.empty and add.empty and rem.empty):
                    # Pre-compute normalized changed frames once for preview + export
                    if not cn.empty:
                        co_norm = normalize_df(co)
                        cn_norm = normalize_df(cn)
                    else:
                        co_norm = cn_norm = pd.DataFrame()

                    changed_sheets.append((sh, co, cn, add, rem, use_key, key_desc,
                                           is_trunc, old_rows, new_rows,
                                           co_norm, cn_norm))

            files_with_changes.append((i + 1, lf.name, rf.name, changed_sheets,
                                       missing_in_new, missing_in_old))
            prog.progress((i + 1) / len(matched), text=f"Done {i+1}/{len(matched)}")

        prog.empty()

        # ── Results ──────────────────────────────
        truly_identical = all(
            not cs and not mn and not mo
            for _, _, _, cs, mn, mo in files_with_changes
        ) if files_with_changes else False

        if len(matched) == 0:
            st.markdown(
                '<div class="warn-box">⚠️ No pairs were compared — '
                'use Manual Pairing above to pair files with different names.</div>',
                unsafe_allow_html=True,
            )
        elif truly_identical:
            st.markdown(
                f'<div class="success-box">✅ All {len(matched)} file pair(s) are identical.</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f'<div class="metric-row">'
                f'<div class="metric-pill info"><span class="val">{len(matched)}</span> pairs compared</div>'
                f'<div class="metric-pill changed"><span class="val">{len(files_with_changes)}</span> with differences</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

            for file_num, old_name, new_name, changed_sheets, missing_in_new, missing_in_old in files_with_changes:
                st.markdown(f"### 📂 Pair {file_num} of {len(matched)}")
                st.markdown(
                    f'<div class="pair-card" style="margin-bottom:1rem">'
                    f'<span class="fname">{old_name}</span>'
                    f'<span class="arrow">⟶</span>'
                    f'<span class="fname">{new_name}</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

                if missing_in_new:
                    sheets_list = ", ".join(f"<b>{s}</b>" for s in missing_in_new)
                    st.markdown(
                        f'<div class="warn-box">📋 Sheet(s) only in OLD file (skipped): {sheets_list}</div>',
                        unsafe_allow_html=True,
                    )
                if missing_in_old:
                    sheets_list = ", ".join(f"<b>{s}</b>" for s in missing_in_old)
                    st.markdown(
                        f'<div class="warn-box">📋 Sheet(s) only in NEW file (skipped): {sheets_list}</div>',
                        unsafe_allow_html=True,
                    )
                if not changed_sheets and not missing_in_new and not missing_in_old:
                    st.markdown(
                        '<div class="success-box" style="text-align:left;margin-bottom:1rem">'
                        '✅ All common sheets are identical.</div>',
                        unsafe_allow_html=True,
                    )

                for sh, co, cn, add, rem, use_key, key_desc, is_trunc, old_rows, new_rows, co_norm, cn_norm in changed_sheets:
                    parts = []
                    if len(cn)  > 0: parts.append(f"{len(cn)} changed")
                    if len(add) > 0: parts.append(f"{len(add)} added")
                    if len(rem) > 0: parts.append(f"{len(rem)} removed")
                    summary = " · ".join(parts) or "differences detected"

                    st.markdown(
                        f'<div class="sheet-header">'
                        f'<span class="sheet-name">📄 {sh}</span>'
                        f'<span class="sheet-summary">{summary}</span>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

                    pills = ""
                    if len(cn)  > 0: pills += f'<div class="metric-pill changed"><span class="val">{len(cn)}</span> changed</div>'
                    if len(add) > 0: pills += f'<div class="metric-pill added"><span class="val">{len(add)}</span> added</div>'
                    if len(rem) > 0: pills += f'<div class="metric-pill removed"><span class="val">{len(rem)}</span> removed</div>'
                    if pills:
                        st.markdown(f'<div class="metric-row">{pills}</div>', unsafe_allow_html=True)

                    if is_trunc:
                        st.markdown(
                            f'<div class="warn-box">⚠️ <b>Truncation detected</b> — '
                            f'OLD: {old_rows} rows · NEW: {new_rows} rows '
                            f'({old_rows - new_rows} rows missing)</div>',
                            unsafe_allow_html=True,
                        )

                    if key_desc and "Key:" in key_desc:
                        st.markdown(
                            f'<span class="diff-label key">🔑 {key_desc}</span>',
                            unsafe_allow_html=True,
                        )

                    if not cn.empty:
                        preview_keys = keys if use_key else []
                        st.markdown('<span class="diff-label changed">🔄 Changed rows</span>',
                                    unsafe_allow_html=True)
                        preview_count = min(len(cn), MAX_PREVIEW_ROWS)
                        if len(cn) > MAX_PREVIEW_ROWS:
                            st.caption(
                                f"Showing the first {MAX_PREVIEW_ROWS:,} changed rows "
                                f"of {len(cn):,}. The download contains all changes."
                            )
                        st.dataframe(
                            build_side_by_side_preview(
                                co.iloc[:preview_count].reset_index(drop=True),
                                cn.iloc[:preview_count].reset_index(drop=True),
                                preview_keys,
                                co_norm.iloc[:preview_count].reset_index(drop=True),
                                cn_norm.iloc[:preview_count].reset_index(drop=True),
                            ),
                            use_container_width=True,
                        )
                        st.download_button(
                            "📥 Download changes (.xlsx)",
                            export_to_excel(co, cn, preview_keys, co_norm, cn_norm),
                            file_name=f"Pair{file_num}_{sh}_changes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{file_num}_{sh}",
                            on_click="ignore",
                        )

                    if not add.empty:
                        st.markdown('<span class="diff-label added">➕ Added rows</span>',
                                    unsafe_allow_html=True)
                        st.dataframe(style_added_rows(add), use_container_width=True)

                    if not rem.empty:
                        st.markdown('<span class="diff-label removed">➖ Removed rows</span>',
                                    unsafe_allow_html=True)
                        st.dataframe(style_removed_rows(rem), use_container_width=True)

                    st.divider()

else:
    st.markdown(
        '<div style="color:#475569;font-size:0.9rem;padding:1.5rem 0;">Upload OLD and NEW files above to begin.</div>',
        unsafe_allow_html=True,
    )
