# app.py – Smart Diff Manager (Improved: UI/UX + Performance + Header Detection)

import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
from io import BytesIO
import msoffcrypto
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from typing import List, Dict, Tuple, Optional
import re
import hashlib

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
  <span class="accent">v2.0 · Excel Diff Tool</span>
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
    decrypted = decrypt_file_cached(uploaded_file.getvalue(), password)
    return BytesIO(decrypted)


# ─────────────────────────────────────────────
# NORMALIZATION UTILITIES
# ─────────────────────────────────────────────

def normalize_colname(name: str) -> str:
    return re.sub(r'[^a-z0-9]', '', str(name).lower().strip())


# Boolean lookup sets
_BOOL_TRUE  = {"TRUE","T","YES","Y","1","1.0","CHECK","CHECKED","CHECKMARK","✓","✔","ON","ENABLED","ACTIVE"}
_BOOL_FALSE = {"FALSE","F","NO","N","0","0.0","CROSS","UNCHECKED","✗","✘","X","OFF","DISABLED","INACTIVE"}

def _normalize_cell(v: str) -> str:
    """Normalize a single string cell value."""
    if v == "" or v == "nan":
        return ""
    s = v.strip().upper()
    if s in _BOOL_TRUE:
        return "TRUE"
    if s in _BOOL_FALSE:
        return "FALSE"
    try:
        num = float(s.replace(",", "."))
        return str(int(num)) if abs(num - int(num)) < 1e-9 else re.sub(r'\.0+$', '', str(num))
    except Exception:
        return s


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Vectorized normalization: fill NaN → stringify → normalize booleans/numerics.
    Much faster than per-cell apply() for large frames.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    d = df.copy()
    d.columns = d.columns.map(str)
    # Vectorized fillna + astype in one pass
    d = d.fillna("").astype(str)
    # Apply scalar function column-by-column (still faster than cell-by-cell)
    for col in d.columns:
        d[col] = d[col].map(_normalize_cell)
    return d


# ─────────────────────────────────────────────
# SMART HEADER DETECTION  (fixed index handling)
# ─────────────────────────────────────────────

def find_header_row_by_column_names(df_raw: pd.DataFrame, reference_columns: List[str]) -> int:
    if not reference_columns:
        return detect_header_row_heuristic(df_raw)
    normalized_ref = {normalize_colname(str(c)) for c in reference_columns if str(c).strip()}
    best_row = 0
    best_match_count = 0
    n_rows = min(15, len(df_raw))
    for i in range(n_rows):                          # ← use range, not iterrows index
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


def find_header_row_with_keys(df_raw: pd.DataFrame, key_columns: List[str]) -> int:
    if not key_columns:
        return detect_header_row_heuristic(df_raw)
    normalized_keys = {normalize_colname(k) for k in key_columns if k.strip()}
    for i in range(min(15, len(df_raw))):            # ← range-based, not iterrows
        row = df_raw.iloc[i]
        row_values = {
            normalize_colname(str(val))
            for val in row
            if pd.notna(val) and str(val).strip()
        }
        if normalized_keys & row_values:
            return i
    return detect_header_row_heuristic(df_raw)


def detect_header_row_heuristic(df_raw: pd.DataFrame) -> int:
    """
    Find the most likely header row in the first 15 rows.
    Uses: non-empty count, text ratio, and column contiguity.
    Fixed: uses iloc[i] (positional) so index resets don't cause off-by-one errors.
    """
    head = df_raw.head(15)
    nonempty_counts = head.apply(lambda r: r.notna().sum(), axis=1)
    max_nonempty = nonempty_counts.max() or 1

    for i in range(len(head)):
        row = head.iloc[i]                           # ← iloc, not label-index
        filled = nonempty_counts.iloc[i]
        if filled < 2:
            continue
        non_null_vals = [v for v in row if pd.notna(v) and str(v).strip()]
        if len(non_null_vals) < 2:
            continue
        filled_ratio  = filled / max_nonempty
        text_ratio    = sum(not str(x).replace('.','',1).isdigit() for x in non_null_vals) / len(non_null_vals)
        valid_indices = [j for j, v in enumerate(row) if pd.notna(v) and str(v).strip()]
        span          = max(valid_indices) - min(valid_indices) + 1 if valid_indices else 1
        contiguous    = len(valid_indices) / span
        if filled_ratio >= 0.4 and text_ratio >= 0.5 and contiguous > 0.5:
            return i
    return 0


# ─────────────────────────────────────────────
# FILE READING  (cached + reduced read_excel calls)
# ─────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def _read_sheets_cached(
    file_bytes_raw: bytes,
    key_columns_tuple: tuple,
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, int]]:
    """
    Reads all visible sheets once, detects headers, returns DataFrames.
    Cached by raw file bytes hash – avoids re-reading on every Streamlit rerun.
    Previously each sheet was read 2-3 times; now it's read exactly ONCE.
    """
    file_bytes = BytesIO(file_bytes_raw)
    key_columns = list(key_columns_tuple)

    try:
        wb = load_workbook(file_bytes, data_only=True)
        visible_sheets = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]

        # Read ALL sheets in one pass (no per-sheet re-seek)
        file_bytes.seek(0)
        all_raw: Dict[str, pd.DataFrame] = pd.read_excel(
            file_bytes, sheet_name=None, header=None, engine="openpyxl"
        )

        sheets: Dict[str, pd.DataFrame] = {}
        header_rows: Dict[str, int] = {}

        for sh in visible_sheets:
            if sh not in all_raw:
                continue
            df_raw = all_raw[sh]

            # Detect visible columns via column_dimensions
            ws = wb[sh]
            visible_col_indices = [
                col_idx - 1
                for col_idx in range(1, ws.max_column + 1)
                if not getattr(ws.column_dimensions.get(
                    ws.cell(row=1, column=col_idx).column_letter
                ), 'hidden', False)
            ]
            if not visible_col_indices:
                visible_col_indices = list(range(df_raw.shape[1]))

            if len(visible_col_indices) < df_raw.shape[1]:
                df_raw = df_raw.iloc[:, visible_col_indices].copy()

            # Detect header row
            if key_columns:
                header_row = find_header_row_with_keys(df_raw, key_columns)
            else:
                header_row = detect_header_row_heuristic(df_raw)

            # Slice header + data in-memory (no second read_excel)
            if header_row < len(df_raw):
                new_cols = df_raw.iloc[header_row].tolist()
                df_data  = df_raw.iloc[header_row + 1:].copy()
                df_data.columns = [str(c) if pd.notna(c) else f"Unnamed_{i}"
                                   for i, c in enumerate(new_cols)]
                df_data = df_data.reset_index(drop=True)
            else:
                df_data = df_raw.copy()
                df_data.columns = [str(c) for c in df_data.columns]

            # Flatten MultiIndex columns if present
            if isinstance(df_data.columns, pd.MultiIndex):
                df_data.columns = [
                    " ".join(str(c) for c in col if str(c) != "nan").strip()
                    for col in df_data.columns
                ]

            sheets[sh]      = df_data.astype(str)
            header_rows[sh] = header_row

        return sheets, header_rows

    except Exception:
        # Fallback: simple read_excel with header=0
        file_bytes.seek(0)
        fallback = pd.read_excel(file_bytes, sheet_name=None, engine="openpyxl")
        for sh in fallback:
            fallback[sh] = fallback[sh].astype(str)
        return fallback, {}


def read_visible_sheets_with_header_detection(
    file_bytes: BytesIO,
    key_columns: List[str] = None,
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, int]]:
    raw = file_bytes.read()
    return _read_sheets_cached(raw, tuple(key_columns or []))


# ─────────────────────────────────────────────
# COLUMN ALIGNMENT
# ─────────────────────────────────────────────

def _normalize_for_match(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r'(_\d+)+$', '', s)      # FIX: was r'(_\\d+)+$' (double-escaped)
    return re.sub(r'[^a-z0-9]', '', s)


def align_new_columns_to_reference(new_cols: List[str], ref_cols: List[str]) -> List[str]:
    new_cols = [str(c) for c in new_cols]
    ref_norm_map = {i: _normalize_for_match(c) for i, c in enumerate(ref_cols)}
    new_norm_map = {i: _normalize_for_match(c) for i, c in enumerate(new_cols)}
    assigned = set()
    renamed  = list(new_cols)

    # Exact match first
    for ref_i, ref_norm in ref_norm_map.items():
        for new_i, new_norm in new_norm_map.items():
            if new_i in assigned:
                continue
            if ref_norm and new_norm == ref_norm:
                renamed[new_i] = ref_cols[ref_i]
                assigned.add(new_i)
                break

    # Substring match fallback
    for ref_i, ref_norm in ref_norm_map.items():
        if not ref_norm:
            continue
        for new_i, new_norm in new_norm_map.items():
            if new_i in assigned:
                continue
            if ref_norm in new_norm or new_norm in ref_norm:
                renamed[new_i] = ref_cols[ref_i]
                assigned.add(new_i)
                break

    # Deduplicate
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

def find_best_valid_key(df1: pd.DataFrame, df2: pd.DataFrame, keys: List[str]) -> Tuple[List[str], str]:
    if not keys:
        return [], "No keys provided"

    def norm(s: str) -> str:
        s = str(s).strip().lower()
        s = re.sub(r'(_\d+)+$', '', s)    # FIX: was r'(_\\d+)+$'
        return re.sub(r'[^a-z0-9]', '', s)

    norm1 = {norm(c): c for c in df1.columns}
    norm2 = {norm(c): c for c in df2.columns}

    best_key, best_uniqueness, best_key_name = None, 0.0, None

    for k in keys:
        nk = norm(k)
        matched_col = None
        match_score = 0.0

        for n1, c1 in norm1.items():
            r1 = SequenceMatcher(None, nk, n1).ratio()
            if r1 < 0.8:
                continue
            for n2 in norm2:
                r2 = SequenceMatcher(None, nk, n2).ratio()
                if r2 < 0.8:
                    continue
                avg = (r1 + r2) / 2
                if avg > match_score:
                    match_score = avg
                    matched_col = c1

        if matched_col:
            try:
                u1 = df1[matched_col].fillna("").astype(str).nunique() / max(1, len(df1))
                u2 = df2[matched_col].fillna("").astype(str).nunique() / max(1, len(df2))
                uniqueness = (u1 + u2) / 2
                if uniqueness > best_uniqueness:
                    best_uniqueness = uniqueness
                    best_key = [matched_col]
                    best_key_name = k
            except Exception:
                continue

    if best_key:
        return best_key, f"Key: '{best_key_name}' ({best_uniqueness:.0%} unique)"
    return [], "No valid keys found"


# ─────────────────────────────────────────────
# COMPARISON FUNCTIONS
# ─────────────────────────────────────────────

def compare_key_based(df1, df2, keys):
    valid_keys, key_desc = find_best_valid_key(df1, df2, keys)
    if not valid_keys:
        raise ValueError("No valid key columns found")

    common = list(df1.columns.intersection(df2.columns))
    df1f = df1[common].fillna("").copy()
    df2f = df2[common].fillna("").copy()
    n1, n2 = normalize_df(df1f), normalize_df(df2f)

    key_str_1 = n1[valid_keys].astype(str).agg("||".join, axis=1)
    key_str_2 = n2[valid_keys].astype(str).agg("||".join, axis=1)
    n1["__key__"] = key_str_1
    n2["__key__"] = key_str_2
    df1f["__key__"] = key_str_1
    df2f["__key__"] = key_str_2

    k1, k2 = set(n1["__key__"]), set(n2["__key__"])
    common_keys  = k1 & k2
    added_keys   = k2 - k1
    removed_keys = k1 - k2

    non_key_cols = [c for c in common if c not in valid_keys and c != "__key__"]

    # Build index maps for O(1) lookup instead of repeated .loc[]
    idx1 = n1.set_index("__key__")
    idx2 = n2.set_index("__key__")
    raw1 = df1f.set_index("__key__")
    raw2 = df2f.set_index("__key__")

    changed_old, changed_new = [], []
    for key in common_keys:
        if key not in idx1.index or key not in idx2.index:
            continue
        r1, r2 = idx1.loc[key], idx2.loc[key]
        # Handle duplicate keys → take first row
        if isinstance(r1, pd.DataFrame): r1 = r1.iloc[0]
        if isinstance(r2, pd.DataFrame): r2 = r2.iloc[0]
        if any(str(r1[c]) != str(r2[c]) for c in non_key_cols):
            changed_old.append(raw1.loc[key].iloc[0] if isinstance(raw1.loc[key], pd.DataFrame) else raw1.loc[key])
            changed_new.append(raw2.loc[key].iloc[0] if isinstance(raw2.loc[key], pd.DataFrame) else raw2.loc[key])

    co = pd.DataFrame(changed_old)[common] if changed_old else pd.DataFrame(columns=common)
    cn = pd.DataFrame(changed_new)[common] if changed_new else pd.DataFrame(columns=common)
    added   = df2f.loc[df2f["__key__"].isin(added_keys),   common].reset_index(drop=True)
    removed = df1f.loc[df1f["__key__"].isin(removed_keys), common].reset_index(drop=True)

    return co.reset_index(drop=True), cn.reset_index(drop=True), added, removed, key_desc


def compare_keyless(df1, df2):
    n1 = normalize_df(df1)
    n2 = normalize_df(df2)
    h1 = pd.util.hash_pandas_object(n1, index=False)
    h2 = pd.util.hash_pandas_object(n2, index=False)
    added   = df2.loc[~h2.isin(h1)].reset_index(drop=True)
    removed = df1.loc[~h1.isin(h2)].reset_index(drop=True)
    return added, removed


def detect_data_truncation(df1: pd.DataFrame, df2: pd.DataFrame) -> Tuple[bool, int, int]:
    old_rows = int(df1.apply(lambda r: r.notna().any(), axis=1).sum())
    new_rows = int(df2.apply(lambda r: r.notna().any(), axis=1).sum())
    threshold = max(5, old_rows * 0.1)
    return (old_rows - new_rows) >= threshold, old_rows, new_rows


# ─────────────────────────────────────────────
# PREVIEW & EXPORT
# ─────────────────────────────────────────────

def build_side_by_side_preview(changed_old: pd.DataFrame, changed_new: pd.DataFrame, keys: List[str]):
    common_cols  = changed_old.columns.intersection(changed_new.columns).tolist()
    old_norm     = normalize_df(changed_old[common_cols])
    new_norm     = normalize_df(changed_new[common_cols])

    rows = []
    for idx in range(len(changed_old)):
        row = {}
        for col in common_cols:
            row[f"{col} (Old)"] = str(changed_old.iloc[idx][col])
            row[f"{col} (New)"] = str(changed_new.iloc[idx][col])
        rows.append(row)
    combined = pd.DataFrame(rows)

    def highlight(row):
        styles = []
        ridx = row.name
        for col in combined.columns:
            if col.endswith(" (Old)"):
                base = col[:-6]
                changed = (base in old_norm.columns and
                           str(old_norm.iloc[ridx][base]) != str(new_norm.iloc[ridx][base]))
                styles.append('background-color: #3b1212; color: #fca5a5' if changed
                               else 'background-color: #0f2a1a; color: #86efac')
            elif col.endswith(" (New)"):
                base = col[:-6]
                changed = (base in new_norm.columns and
                           str(old_norm.iloc[ridx][base]) != str(new_norm.iloc[ridx][base]))
                styles.append('background-color: #0f2a1a; color: #86efac' if changed
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
    return df.style.apply(lambda r: ['background-color: #052e16; color: #86efac' for _ in r], axis=1)


def style_removed_rows(df: pd.DataFrame):
    if df is None or df.empty:
        return df
    df = _dedup_columns(df.copy())
    return df.style.apply(lambda r: ['background-color: #450a0a; color: #fca5a5' for _ in r], axis=1)


def export_to_excel(changed_old: pd.DataFrame, changed_new: pd.DataFrame, keys: List[str]) -> BytesIO:
    wb = Workbook()
    ws_old = wb.active
    ws_old.title = "Old Values"
    ws_new = wb.create_sheet(title="New Values")
    fill_old = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    fill_new = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

    old_norm = normalize_df(changed_old)
    new_norm = normalize_df(changed_new)

    for ws, df, fill, norm_other in [
        (ws_old, changed_old, fill_old, new_norm),
        (ws_new, changed_new, fill_new, old_norm),
    ]:
        for c, col_name in enumerate(df.columns, 1):
            ws.cell(row=1, column=c, value=col_name)
        for r_idx in range(len(df)):
            for c_idx, col_name in enumerate(df.columns, 1):
                cell = ws.cell(row=r_idx + 2, column=c_idx, value=df.iloc[r_idx][col_name])
                if (col_name not in keys and col_name in old_norm.columns and
                        str(old_norm.iloc[r_idx][col_name]) != str(new_norm.iloc[r_idx][col_name])):
                    cell.fill = fill

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ─────────────────────────────────────────────
# UI LAYOUT
# ─────────────────────────────────────────────

# ── Step 1: Upload
st.markdown('<div class="step-label">Step 1 — Upload Files</div>', unsafe_allow_html=True)
col_l, col_r = st.columns(2)
with col_l:
    left = st.file_uploader("OLD files", ["xlsx", "xls"], accept_multiple_files=True,
                             label_visibility="visible")
with col_r:
    right = st.file_uploader("NEW files", ["xlsx", "xls"], accept_multiple_files=True,
                              label_visibility="visible")

# ── Step 2: Config
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

# ── Step 3: Match & Run
if left and right:
    def comparable(n, ignore):
        n = n.rsplit(".", 1)[0]
        return n.rsplit("_", 1)[0].lower() if ignore and "_" in n else n.lower()

    matched, usedR = [], set()
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
            matched.append((lf, best))
            usedR.add(best.name)

    # Matched pairs summary
    st.markdown(f'<div class="step-label" style="margin-top:1.5rem">Step 3 — Review & Run</div>',
                unsafe_allow_html=True)

    if matched:
        with st.expander(f"📋 {len(matched)} matched pair(s)", expanded=False):
            for lf, rf in matched:
                st.markdown(
                    f'<div class="pair-card">'
                    f'<span class="fname">{lf.name}</span>'
                    f'<span class="arrow">⟶</span>'
                    f'<span class="fname">{rf.name}</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
        unmatched_l = [f.name for f in left  if f.name not in {lf.name for lf, _ in matched}]
        unmatched_r = [f.name for f in right if f.name not in {rf.name for _, rf in matched}]
        if unmatched_l or unmatched_r:
            st.markdown(
                f'<div class="warn-box">⚠️ Unmatched — OLD: {", ".join(unmatched_l) or "none"} '
                f'· NEW: {", ".join(unmatched_r) or "none"}</div>',
                unsafe_allow_html=True,
            )
    else:
        st.warning("No file pairs matched. Try lowering the match threshold.")

    run_clicked = st.button("▶  Run Comparison", use_container_width=False)

    if run_clicked and matched:
        prog = st.progress(0, text="Starting…")
        keys = [k.strip() for k in keys_str.split(",") if k.strip()]
        files_with_changes = []

        for i, (lf, rf) in enumerate(matched):
            prog.progress((i) / len(matched), text=f"Comparing {lf.name} …")
            try:
                decL = decrypt_file(lf)
                decR = decrypt_file(rf)

                shL, header_rows_old = read_visible_sheets_with_header_detection(decL, key_columns=keys)

                # Align NEW sheets to OLD header rows & column names
                shR: Dict[str, pd.DataFrame] = {}
                for sh in shL:
                    hr_old = header_rows_old.get(sh, 0)
                    decR.seek(0)
                    df_new_raw = pd.read_excel(decR, sheet_name=sh, header=None, engine="openpyxl")
                    if hr_old < len(df_new_raw):
                        new_cols = df_new_raw.iloc[hr_old].tolist()
                        df_new   = df_new_raw.iloc[hr_old + 1:].reset_index(drop=True).copy()
                        df_new.columns = new_cols
                    else:
                        df_new = pd.read_excel(decR, sheet_name=sh, header=0, engine="openpyxl")
                    try:
                        df_new.columns = align_new_columns_to_reference(
                            list(df_new.columns), list(shL[sh].columns)
                        )
                    except Exception:
                        df_new.columns = [str(c) for c in df_new.columns]
                    shR[sh] = df_new.astype(str)

            except Exception as e:
                import traceback
                st.error(f"❌ Pair {i+1}: Failed to read — {e}")
                st.code(traceback.format_exc())
                prog.progress((i + 1) / len(matched))
                continue

            changed_sheets = []
            for sh in sorted(set(shL) & set(shR)):
                d1, d2 = shL[sh], shR[sh]
                use_key = False
                key_desc = ""
                if keys:
                    try:
                        valid_keys, key_desc = find_best_valid_key(d1, d2, keys)
                        use_key = bool(valid_keys)
                    except Exception:
                        use_key = False

                try:
                    if use_key:
                        co, cn, add, rem, key_desc = compare_key_based(d1, d2, keys)
                    else:
                        co = cn = pd.DataFrame()
                        add, rem = compare_keyless(d1, d2)
                except Exception:
                    co = cn = pd.DataFrame()
                    add, rem = compare_keyless(d1, d2)

                is_trunc, old_rows, new_rows = detect_data_truncation(d1, d2)

                if not (cn.empty and add.empty and rem.empty):
                    changed_sheets.append((sh, co, cn, add, rem, use_key, key_desc,
                                           is_trunc, old_rows, new_rows))

            if changed_sheets:
                files_with_changes.append((i + 1, lf.name, rf.name, changed_sheets))

            prog.progress((i + 1) / len(matched), text=f"Done {i+1}/{len(matched)}")

        prog.empty()

        # ── Results ──────────────────────────────
        if not files_with_changes:
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

            for file_num, old_name, new_name, changed_sheets in files_with_changes:
                st.markdown(f"### 📂 Pair {file_num} of {len(matched)}")
                st.markdown(
                    f'<div class="pair-card" style="margin-bottom:1rem">'
                    f'<span class="fname">{old_name}</span>'
                    f'<span class="arrow">⟶</span>'
                    f'<span class="fname">{new_name}</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

                for sh, co, cn, add, rem, use_key, key_desc, is_trunc, old_rows, new_rows in changed_sheets:
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

                    # Metric pills
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

                    if not cn.empty and use_key:
                        st.markdown('<span class="diff-label changed">🔄 Changed rows</span>',
                                    unsafe_allow_html=True)
                        st.dataframe(build_side_by_side_preview(co, cn, keys), use_container_width=True)
                        st.download_button(
                            "📥 Download changes (.xlsx)",
                            export_to_excel(co, cn, keys),
                            file_name=f"Pair{file_num}_{sh}_changes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{file_num}_{sh}",
                        )
                    elif not cn.empty:
                        st.markdown('<span class="diff-label changed">🔄 Changed rows</span>',
                                    unsafe_allow_html=True)
                        st.dataframe(cn, use_container_width=True)

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
