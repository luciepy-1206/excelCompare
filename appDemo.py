# app.py – Smart Diff Manager (Enhanced Version)

import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
from io import BytesIO
import msoffcrypto
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from typing import List, Dict, Tuple, Optional, Set
import re
from dataclasses import dataclass
from enum import Enum

DEFAULT_PASSWORD = "mypassword"

st.set_page_config(
    page_title="Smart Diff Manager",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------------------------------------------------
# DATA CLASSES & ENUMS
# ----------------------------------------------------------------------

class ComparisonMode(Enum):
    KEY_BASED = "key_based"
    KEYLESS = "keyless"

@dataclass
class SheetComparison:
    """Container for sheet comparison results"""
    sheet_name: str
    changed_old: pd.DataFrame
    changed_new: pd.DataFrame
    added: pd.DataFrame
    removed: pd.DataFrame
    is_key_based: bool
    key_description: str
    is_truncated: bool = False
    old_row_count: int = 0
    new_row_count: int = 0

@dataclass
class FileComparisonResult:
    """Container for file pair comparison results"""
    file_index: int
    old_filename: str
    new_filename: str
    sheet_comparisons: List[SheetComparison]

# ----------------------------------------------------------------------
# STYLES
# ----------------------------------------------------------------------
st.markdown(
    """
<style>
.dataframe-container { font-size: 14px; }
.small-text { font-size: 13px; opacity: 0.9; margin-left: 1rem; }
.success-box { padding: 1rem; background-color: #d4edda; border-radius: 0.5rem; }
.warning-box { padding: 1rem; background-color: #fff3cd; border-radius: 0.5rem; }
</style>
""",
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# HELPER FUNCTIONS
# ----------------------------------------------------------------------

def decrypt_file(uploaded_file, password: Optional[str] = DEFAULT_PASSWORD) -> BytesIO:
    """Decrypt Excel file if encrypted, otherwise return as-is"""
    fb = BytesIO(uploaded_file.getvalue())
    try:
        pd.ExcelFile(fb)
        fb.seek(0)
        return fb
    except Exception:
        fb.seek(0)
        try:
            office = msoffcrypto.OfficeFile(fb)
            if office.is_encrypted():
                dec = BytesIO()
                office.load_key(password=password)
                office.decrypt(dec)
                dec.seek(0)
                return dec
        except Exception:
            pass
        fb.seek(0)
        return fb

# ----------------------------------------------------------------------
# NORMALIZATION UTILITIES
# ----------------------------------------------------------------------

def normalize_colname(name: str) -> str:
    """Normalize column names for comparison - remove special chars, lowercase"""
    return re.sub(r'[^a-z0-9]', '', str(name).lower().strip())

def normalize_logical(v) -> str:
    """
    Normalize various representations of boolean and numeric values for comparison.
    Handles: TRUE/FALSE, numeric values, and standard strings
    """
    if pd.isna(v) or v == "":
        return ""

    s = str(v).strip().upper()

    # Boolean normalization
    TRUE_VALUES = {"TRUE", "T", "YES", "Y", "1", "1.0", "CHECK", "CHECKED", 
                   "CHECKMARK", "✓", "✔", "ON", "ENABLED", "ACTIVE"}
    FALSE_VALUES = {"FALSE", "F", "NO", "N", "0", "0.0", "CROSS", "UNCHECKED", 
                    "✗", "✘", "X", "OFF", "DISABLED", "INACTIVE"}
    
    if s in TRUE_VALUES:
        return "TRUE"
    if s in FALSE_VALUES:
        return "FALSE"

    # Numeric normalization
    try:
        sn = s.replace(",", ".")
        num = float(sn)
        # If integer-valued, cast to int string
        if abs(num - int(num)) < 1e-9:
            return str(int(num))
        # Remove trailing zeros from decimals
        return re.sub(r'\.0+$', '', str(num))
    except (ValueError, OverflowError):
        pass

    return s

def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize dataframe - fill NaN, convert to string, and normalize values"""
    if df is None or df.empty:
        return pd.DataFrame()
    
    d = df.copy().fillna("")
    d.columns = d.columns.map(str)
    
    for col in d.columns:
        d[col] = d[col].astype(str).apply(normalize_logical)
    
    return d

# ----------------------------------------------------------------------
# SMART HEADER DETECTION
# ----------------------------------------------------------------------

def calculate_header_score(row: pd.Series, max_nonempty: int) -> float:
    """Calculate a score for how likely a row is to be a header"""
    filled = row.notna().sum()
    
    if filled < 2:
        return 0.0
    
    non_null_values = [str(val) for val in row if pd.notna(val) and str(val).strip()]
    if len(non_null_values) < 2:
        return 0.0
    
    # Calculate various metrics
    filled_ratio = filled / max(1, max_nonempty)
    text_ratio = sum(isinstance(x, str) and str(x).strip() for x in row) / max(1, filled)
    
    # Calculate contiguity
    valid_indices = [idx for idx, val in enumerate(row) if pd.notna(val) and str(val).strip()]
    if valid_indices:
        span = max(valid_indices) - min(valid_indices) + 1
        contiguity = len(valid_indices) / span
    else:
        contiguity = 0
    
    # Weighted score
    score = (filled_ratio * 0.4) + (text_ratio * 0.3) + (contiguity * 0.3)
    return score

def detect_header_row_heuristic(df_raw: pd.DataFrame, max_rows: int = 15) -> int:
    """Detect header row using heuristic scoring"""
    max_nonempty = df_raw.head(max_rows).apply(lambda r: r.notna().sum(), axis=1).max()
    
    best_row = 0
    best_score = 0.0
    
    for i, row in df_raw.head(max_rows).iterrows():
        score = calculate_header_score(row, max_nonempty)
        if score > best_score:
            best_score = score
            best_row = i
    
    return best_row

def find_header_row_by_column_names(
    df_raw: pd.DataFrame, 
    reference_columns: List[str], 
    max_rows: int = 15
) -> int:
    """Find header row by matching reference column names"""
    if not reference_columns:
        return detect_header_row_heuristic(df_raw, max_rows)
    
    normalized_ref = {normalize_colname(str(c)) for c in reference_columns if str(c).strip()}
    
    best_row = 0
    best_match_count = 0
    min_matches = max(2, len(normalized_ref) * 0.5)
    
    for i in range(min(max_rows, len(df_raw))):
        row = df_raw.iloc[i]
        row_values = {
            normalize_colname(str(val))
            for val in row
            if pd.notna(val) and str(val).strip()
        }
        
        match_count = len(normalized_ref & row_values)
        
        if match_count >= min_matches and match_count > best_match_count:
            best_match_count = match_count
            best_row = i
            
            # Early exit if we have a strong match
            if match_count >= len(normalized_ref) * 0.8:
                return best_row
    
    if best_match_count >= 2:
        return best_row
    
    return detect_header_row_heuristic(df_raw, max_rows)

def find_header_row_with_keys(
    df_raw: pd.DataFrame, 
    key_columns: List[str],
    max_rows: int = 15
) -> int:
    """Find header row by locating key columns"""
    if not key_columns:
        return detect_header_row_heuristic(df_raw, max_rows)
    
    normalized_keys = {normalize_colname(k) for k in key_columns if k.strip()}
    
    for i in range(min(max_rows, len(df_raw))):
        row = df_raw.iloc[i]
        row_values = [
            normalize_colname(str(val))
            for val in row
            if pd.notna(val) and str(val).strip()
        ]
        
        if normalized_keys & set(row_values):
            return i
    
    return detect_header_row_heuristic(df_raw, max_rows)

# ----------------------------------------------------------------------
# EXCEL READING
# ----------------------------------------------------------------------

def read_visible_sheets_with_header_detection(
    file_bytes: BytesIO,
    key_columns: Optional[List[str]] = None,
    reference_headers: Optional[Dict[str, List[str]]] = None
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, int]]:
    """Read visible sheets from Excel with smart header detection"""
    try:
        wb = load_workbook(file_bytes, data_only=True)
        visible_sheets = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
        
        sheets = {}
        header_rows = {}
        
        for sheet_name in visible_sheets:
            ws = wb[sheet_name]
            
            # Get visible column indices
            visible_col_indices = []
            for col_idx in range(1, ws.max_column + 1):
                col_letter = ws.cell(row=1, column=col_idx).column_letter
                col_dim = ws.column_dimensions.get(col_letter)
                is_visible = (col_dim is None) or (not col_dim.hidden)
                if is_visible:
                    visible_col_indices.append(col_idx - 1)
            
            if not visible_col_indices:
                visible_col_indices = list(range(ws.max_column))
            
            # Read raw data to detect header
            file_bytes.seek(0)
            df_raw = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None, engine="openpyxl")
            
            if visible_col_indices and len(visible_col_indices) < len(df_raw.columns):
                df_raw = df_raw.iloc[:, visible_col_indices]
            
            # Detect header row
            if reference_headers and sheet_name in reference_headers:
                header_row = find_header_row_by_column_names(df_raw, reference_headers[sheet_name])
            elif key_columns:
                header_row = find_header_row_with_keys(df_raw, key_columns)
            else:
                header_row = detect_header_row_heuristic(df_raw)
            
            # Read with detected header
            file_bytes.seek(0)
            df_full = pd.read_excel(file_bytes, sheet_name=sheet_name, header=header_row, engine="openpyxl")
            
            if visible_col_indices and len(visible_col_indices) < len(df_full.columns):
                df_main = df_full.iloc[:, visible_col_indices]
            else:
                df_main = df_full
            
            # Handle MultiIndex columns
            if isinstance(df_main.columns, pd.MultiIndex):
                df_main.columns = [
                    " ".join([str(c) for c in col if str(c) != "nan"]).strip()
                    for col in df_main.columns
                ]
            
            df_main = df_main.astype(str)
            sheets[sheet_name] = df_main
            header_rows[sheet_name] = header_row
        
        return sheets, header_rows
        
    except Exception:
        # Fallback: read all sheets
        file_bytes.seek(0)
        all_sheets_raw = pd.read_excel(file_bytes, sheet_name=None, engine="openpyxl", header=None)
        
        sheets = {}
        header_rows = {}
        
        for sheet_name, df_raw in all_sheets_raw.items():
            if reference_headers and sheet_name in reference_headers:
                header_row = find_header_row_by_column_names(df_raw, reference_headers[sheet_name])
            elif key_columns:
                header_row = find_header_row_with_keys(df_raw, key_columns)
            else:
                header_row = detect_header_row_heuristic(df_raw)
            
            file_bytes.seek(0)
            df = pd.read_excel(file_bytes, sheet_name=sheet_name, header=header_row, engine="openpyxl")
            df = df.astype(str)
            sheets[sheet_name] = df
            header_rows[sheet_name] = header_row
        
        return sheets, header_rows

# ----------------------------------------------------------------------
# KEY SELECTION
# ----------------------------------------------------------------------

def find_best_valid_key(
    df1: pd.DataFrame, 
    df2: pd.DataFrame, 
    keys: List[str],
    similarity_threshold: float = 0.8
) -> Tuple[List[str], str]:
    """
    Find the best valid key column that exists in both dataframes.
    Returns the key column name and a description of its quality.
    """
    if not keys:
        return [], "No keys provided"

    def normalize_key(s: str) -> str:
        """Normalize key string for matching"""
        s = str(s).strip().lower()
        s = re.sub(r'(_\d+)+$', '', s)  # Remove trailing _0/_1
        return re.sub(r'[^a-z0-9]', '', s)

    # Create normalized column mappings
    norm1 = {normalize_key(c): c for c in df1.columns}
    norm2 = {normalize_key(c): c for c in df2.columns}

    best_key = None
    best_uniqueness = 0
    best_key_name = None

    for key in keys:
        normalized_key = normalize_key(key)
        matched_col = None
        match_score = 0

        # Find best match in both dataframes
        for norm_col1, actual_col1 in norm1.items():
            ratio1 = SequenceMatcher(None, normalized_key, norm_col1).ratio()
            if ratio1 < similarity_threshold:
                continue
            
            for norm_col2, actual_col2 in norm2.items():
                ratio2 = SequenceMatcher(None, normalized_key, norm_col2).ratio()
                if ratio2 < similarity_threshold:
                    continue
                
                avg_ratio = (ratio1 + ratio2) / 2
                if avg_ratio > match_score:
                    match_score = avg_ratio
                    matched_col = actual_col1

        if matched_col:
            try:
                # Calculate uniqueness
                unique1 = df1[matched_col].fillna("").astype(str).nunique()
                unique2 = df2[matched_col].fillna("").astype(str).nunique()
                uniqueness = (unique1 / max(1, len(df1)) + unique2 / max(1, len(df2))) / 2
                
                if uniqueness > best_uniqueness:
                    best_uniqueness = uniqueness
                    best_key = [matched_col]
                    best_key_name = key
            except Exception:
                continue

    if best_key:
        return best_key, f"Key: '{best_key_name}' ({best_uniqueness:.0%} unique)"
    
    return [], "No valid keys found"

# ----------------------------------------------------------------------
# COMPARISON FUNCTIONS
# ----------------------------------------------------------------------

def compare_key_based(
    df1: pd.DataFrame, 
    df2: pd.DataFrame, 
    keys: List[str]
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, str]:
    """Compare dataframes using key-based matching"""
    valid_keys, key_desc = find_best_valid_key(df1, df2, keys)
    
    if not valid_keys:
        raise ValueError("No valid key columns found")
    
    # Get common columns
    common = list(df1.columns.intersection(df2.columns))
    df1_filtered = df1[common].copy().fillna("")
    df2_filtered = df2[common].copy().fillna("")
    
    # Normalize for comparison
    norm1 = normalize_df(df1_filtered)
    norm2 = normalize_df(df2_filtered)
    
    # Create composite keys
    norm1["__key__"] = norm1[valid_keys].astype(str).agg("||".join, axis=1)
    norm2["__key__"] = norm2[valid_keys].astype(str).agg("||".join, axis=1)
    df1_filtered["__key__"] = norm1["__key__"]
    df2_filtered["__key__"] = norm2["__key__"]
    
    # Find differences
    keys1 = set(norm1["__key__"])
    keys2 = set(norm2["__key__"])
    
    common_keys = keys1 & keys2
    added_keys = keys2 - keys1
    removed_keys = keys1 - keys2
    
    # Find changed rows
    changed_old = []
    changed_new = []
    
    for key in common_keys:
        row1 = norm1[norm1["__key__"] == key]
        row2 = norm2[norm2["__key__"] == key]
        
        if row1.empty or row2.empty:
            continue
        
        row1 = row1.iloc[0]
        row2 = row2.iloc[0]
        
        # Check if any non-key column changed
        is_changed = any(
            str(row1[col]) != str(row2[col]) 
            for col in common 
            if col not in valid_keys and col != "__key__"
        )
        
        if is_changed:
            changed_old.append(df1_filtered[df1_filtered["__key__"] == key].iloc[0])
            changed_new.append(df2_filtered[df2_filtered["__key__"] == key].iloc[0])
    
    # Create result dataframes
    changed_old_df = pd.DataFrame(changed_old)[common] if changed_old else pd.DataFrame(columns=common)
    changed_new_df = pd.DataFrame(changed_new)[common] if changed_new else pd.DataFrame(columns=common)
    
    added_df = (df2_filtered[df2_filtered["__key__"].isin(added_keys)][common].reset_index(drop=True) 
                if added_keys else pd.DataFrame(columns=common))
    removed_df = (df1_filtered[df1_filtered["__key__"].isin(removed_keys)][common].reset_index(drop=True) 
                  if removed_keys else pd.DataFrame(columns=common))
    
    return (changed_old_df.reset_index(drop=True), 
            changed_new_df.reset_index(drop=True), 
            added_df, 
            removed_df, 
            key_desc)

def compare_keyless(
    df1: pd.DataFrame, 
    df2: pd.DataFrame
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Compare dataframes using row-based hashing (no keys)"""
    norm1 = normalize_df(df1)
    norm2 = normalize_df(df2)
    
    hash1 = pd.util.hash_pandas_object(norm1, index=False)
    hash2 = pd.util.hash_pandas_object(norm2, index=False)
    
    added = df2[~hash2.isin(hash1)].copy().reset_index(drop=True)
    removed = df1[~hash1.isin(hash2)].copy().reset_index(drop=True)
    
    return added, removed

def detect_data_truncation(
    df1: pd.DataFrame, 
    df2: pd.DataFrame,
    threshold_ratio: float = 0.1
) -> Tuple[bool, int, int]:
    """Detect if data has been truncated between old and new versions"""
    old_non_empty = df1.apply(lambda row: row.notna().any(), axis=1).sum()
    new_non_empty = df2.apply(lambda row: row.notna().any(), axis=1).sum()
    
    threshold = max(5, old_non_empty * threshold_ratio)
    is_truncated = (old_non_empty - new_non_empty) >= threshold
    
    return is_truncated, old_non_empty, new_non_empty

# ----------------------------------------------------------------------
# COLUMN ALIGNMENT
# ----------------------------------------------------------------------

def normalize_for_column_match(s: str) -> str:
    """Normalize column name for matching"""
    s = str(s).strip().lower()
    s = re.sub(r'(_\d+)+$', '', s)  # Remove trailing _0/_1 suffixes
    return re.sub(r'[^a-z0-9]', '', s)

def align_new_columns_to_reference(
    new_cols: List[str], 
    ref_cols: List[str]
) -> List[str]:
    """Align new file column names to match reference file"""
    new_cols = [str(c) for c in new_cols]
    
    # Create normalized mappings
    ref_norm_map = {i: normalize_for_column_match(c) for i, c in enumerate(ref_cols)}
    new_norm_map = {i: normalize_for_column_match(c) for i, c in enumerate(new_cols)}
    
    assigned_new_idx = set()
    renamed = list(new_cols)
    
    # First pass: exact normalized matches
    for ref_i, ref_norm in ref_norm_map.items():
        for new_i, new_norm in new_norm_map.items():
            if new_i in assigned_new_idx:
                continue
            if ref_norm and new_norm == ref_norm:
                renamed[new_i] = ref_cols[ref_i]
                assigned_new_idx.add(new_i)
                break
    
    # Second pass: partial matches
    for ref_i, ref_norm in ref_norm_map.items():
        if not ref_norm:
            continue
        for new_i, new_norm in new_norm_map.items():
            if new_i in assigned_new_idx:
                continue
            if ref_norm in new_norm or new_norm in ref_norm:
                renamed[new_i] = ref_cols[ref_i]
                assigned_new_idx.add(new_i)
                break
    
    # Handle duplicates
    final = []
    seen = {}
    for name in renamed:
        base = str(name)
        if base in seen:
            seen[base] += 1
            final_name = f"{base}_{seen[base]}"
        else:
            seen[base] = 0
            final_name = base
        final.append(final_name)
    
    return final

# ----------------------------------------------------------------------
# PREVIEW & EXPORT
# ----------------------------------------------------------------------

def build_side_by_side_preview(
    changed_old: pd.DataFrame, 
    changed_new: pd.DataFrame, 
    keys: List[str]
) -> pd.io.formats.style.Styler:
    """Build side-by-side comparison view with highlighting"""
    common_cols = changed_old.columns.intersection(changed_new.columns).tolist()
    
    old_df = changed_old[common_cols].copy()
    new_df = changed_new[common_cols].copy()
    old_df_norm = normalize_df(old_df)
    new_df_norm = normalize_df(new_df)
    
    # Build combined dataframe
    combined_data = []
    for idx in range(len(old_df)):
        row_data = {}
        for col in common_cols:
            row_data[f"{col} (Old)"] = str(old_df.iloc[idx][col])
            row_data[f"{col} (New)"] = str(new_df.iloc[idx][col])
        combined_data.append(row_data)
    
    combined_df = pd.DataFrame(combined_data)
    
    def highlight_changes(row):
        """Apply highlighting based on changes"""
        styles = []
        row_idx = row.name
        
        for col in combined_df.columns:
            if col.endswith(" (Old)"):
                base_col = col[:-6]
                if base_col in old_df_norm.columns:
                    old_val = str(old_df_norm.iloc[row_idx][base_col])
                    new_val = str(new_df_norm.iloc[row_idx][base_col])
                    
                    if old_val != new_val:
                        styles.append('background-color: #ffcccc; color: #000000')
                    else:
                        styles.append('background-color: #e8f5e9; color: #000000')
                else:
                    styles.append('')
            elif col.endswith(" (New)"):
                base_col = col[:-6]
                if base_col in new_df_norm.columns:
                    old_val = str(old_df_norm.iloc[row_idx][base_col])
                    new_val = str(new_df_norm.iloc[row_idx][base_col])
                    
                    if old_val != new_val:
                        styles.append('background-color: #ccffcc; color: #000000')
                    else:
                        styles.append('background-color: #e8f5e9; color: #000000')
                else:
                    styles.append('')
            else:
                styles.append('')
        
        return styles
    
    return combined_df.style.apply(highlight_changes, axis=1)

def style_added_rows(df: pd.DataFrame):
    """Style added rows with blue background"""
    if df is None or df.empty:
        return df
    
    df = df.copy()
    if not df.columns.is_unique:
        df.columns = [f"{col}_{i}" for i, col in enumerate(df.columns)]
    
    return df.style.apply(
        lambda row: ['background-color: #e3f2fd; color: #000000' for _ in row], 
        axis=1
    )

def style_removed_rows(df: pd.DataFrame):
    """Style removed rows with red background"""
    if df is None or df.empty:
        return df
    
    df = df.copy()
    if not df.columns.is_unique:
        df.columns = [f"{col}_{i}" for i, col in enumerate(df.columns)]
    
    return df.style.apply(
        lambda row: ['background-color: #ffebee; color: #000000' for _ in row], 
        axis=1
    )

def export_to_excel(
    changed_old: pd.DataFrame, 
    changed_new: pd.DataFrame, 
    keys: List[str]
) -> BytesIO:
    """Export comparison to Excel with highlighting"""
    wb = Workbook()
    ws_old = wb.active
    ws_old.title = "Old Values"
    ws_new = wb.create_sheet(title="New Values")
    
    old_highlight = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    new_highlight = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    
    changed_old_norm = normalize_df(changed_old)
    changed_new_norm = normalize_df(changed_new)
    
    # Write old values
    for c, col_name in enumerate(changed_old.columns, 1):
        ws_old.cell(row=1, column=c, value=col_name)
    
    for r_idx in range(len(changed_old)):
        for c_idx, col_name in enumerate(changed_old.columns, 1):
            cell = ws_old.cell(row=r_idx + 2, column=c_idx, 
                              value=changed_old.iloc[r_idx][col_name])
            
            if (col_name not in keys and 
                str(changed_old_norm.iloc[r_idx][col_name]) != 
                str(changed_new_norm.iloc[r_idx][col_name])):
                cell.fill = old_highlight
    
    # Write new values
    for c, col_name in enumerate(changed_new.columns, 1):
        ws_new.cell(row=1, column=c, value=col_name)
    
    for r_idx in range(len(changed_new)):
        for c_idx, col_name in enumerate(changed_new.columns, 1):
            cell = ws_new.cell(row=r_idx + 2, column=c_idx, 
                              value=changed_new.iloc[r_idx][col_name])
            
            if (col_name not in keys and 
                str(changed_old_norm.iloc[r_idx][col_name]) != 
                str(changed_new_norm.iloc[r_idx][col_name])):
                cell.fill = new_highlight
    
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ----------------------------------------------------------------------
# FILE MATCHING
# ----------------------------------------------------------------------

def match_file_pairs(
    left_files: List, 
    right_files: List,
    ignore_suffix: bool = True,
    threshold: float = 0.85
) -> List[Tuple]:
    """Match file pairs based on name similarity"""
    def normalize_filename(name: str, ignore_suffix: bool) -> str:
        name = name.rsplit(".", 1)[0]  # Remove extension
        if ignore_suffix and "_" in name:
            name = name.rsplit("_", 1)[0]  # Remove suffix after last underscore
        return name.lower()
    
    matched = []
    used_right = set()
    
    for left_file in left_files:
        left_normalized = normalize_filename(left_file.name, ignore_suffix)
        
        best_match = None
        best_ratio = 0
        
        for right_file in right_files:
            if right_file.name in used_right:
                continue
            
            right_normalized = normalize_filename(right_file.name, ignore_suffix)
            ratio = SequenceMatcher(None, left_normalized, right_normalized).ratio()
            
            if ratio >= threshold and ratio > best_ratio:
                best_match = right_file
                best_ratio = ratio
        
        if best_match:
            matched.append((left_file, best_match))
            used_right.add(best_match.name)
    
    return matched

# ----------------------------------------------------------------------
# MAIN COMPARISON LOGIC
# ----------------------------------------------------------------------

def process_file_pair(
    old_file,
    new_file,
    keys: List[str],
    file_index: int
) -> Optional[FileComparisonResult]:
    """Process a single file pair and return comparison results"""
    try:
        # Decrypt files
        dec_old = decrypt_file(old_file)
        dec_new = decrypt_file(new_file)
        
        # Read old file with header detection
        sheets_old, header_rows_old = read_visible_sheets_with_header_detection(
            dec_old, key_columns=keys
        )
        
        # Align new file to old file structure
        sheets_new = {}
        header_rows_new = {}
        
        for sheet_name in sheets_old.keys():
            header_row_old = header_rows_old.get(sheet_name, 0)
            
            # Read new file raw to get same header row
            dec_new.seek(0)
            df_new_raw = pd.read_excel(dec_new, sheet_name=sheet_name, header=None, engine="openpyxl")
            
            if header_row_old < len(df_new_raw):
                new_header = df_new_raw.iloc[header_row_old].tolist()
                df_new = df_new_raw.iloc[header_row_old + 1:].copy()
                df_new.columns = new_header
            else:
                df_new = pd.read_excel(dec_new, sheet_name=sheet_name, header=0, engine="openpyxl")
            
            # Align column names to old file
            try:
                ref_cols = list(sheets_old[sheet_name].columns)
                aligned_cols = align_new_columns_to_reference(list(df_new.columns), ref_cols)
                df_new.columns = aligned_cols
            except Exception:
                df_new.columns = [str(c) for c in df_new.columns]
            
            sheets_new[sheet_name] = df_new.astype(str)
            header_rows_new[sheet_name] = header_row_old
        
        # Compare sheets
        sheet_comparisons = []
        
        for sheet_name in sorted(set(sheets_old) & set(sheets_new)):
            df_old = sheets_old[sheet_name]
            df_new = sheets_new[sheet_name]
            
            # Determine comparison mode
            use_key_based = False
            key_desc = ""
            
            if keys:
                try:
                    valid_keys, key_desc = find_best_valid_key(df_old, df_new, keys)
                    use_key_based = len(valid_keys) > 0
                except Exception:
                    use_key_based = False
            
            # Perform comparison
            try:
                if use_key_based:
                    changed_old, changed_new, added, removed, key_desc = compare_key_based(
                        df_old, df_new, keys
                    )
                else:
                    changed_old = changed_new = pd.DataFrame()
                    added, removed = compare_keyless(df_old, df_new)
                    key_desc = "Row-based comparison (no key columns)"
            except Exception:
                changed_old = changed_new = pd.DataFrame()
                added, removed = compare_keyless(df_old, df_new)
                key_desc = "Row-based comparison (fallback)"
            
            # Check for truncation
            is_truncated, old_rows, new_rows = detect_data_truncation(df_old, df_new)
            
            # Only add if there are differences
            if not (changed_new.empty and added.empty and removed.empty):
                sheet_comparisons.append(
                    SheetComparison(
                        sheet_name=sheet_name,
                        changed_old=changed_old,
                        changed_new=changed_new,
                        added=added,
                        removed=removed,
                        is_key_based=use_key_based,
                        key_description=key_desc,
                        is_truncated=is_truncated,
                        old_row_count=old_rows,
                        new_row_count=new_rows
                    )
                )
        
        if sheet_comparisons:
            return FileComparisonResult(
                file_index=file_index,
                old_filename=old_file.name,
                new_filename=new_file.name,
                sheet_comparisons=sheet_comparisons
            )
        
        return None
        
    except Exception as e:
        st.error(f"❌ Error processing file pair: {e}")
        import traceback
        with st.expander("Show error details"):
            st.code(traceback.format_exc())
        return None

# ----------------------------------------------------------------------
# UI COMPONENTS
# ----------------------------------------------------------------------

def render_comparison_summary(result: FileComparisonResult, total_pairs: int):
    """Render summary for a single file pair comparison"""
    st.markdown(f"## 📂 File Pair {result.file_index}/{total_pairs}")
    st.markdown(f"**OLD:** `{result.old_filename}`  \n**NEW:** `{result.new_filename}`")
    
    for sheet_comp in result.sheet_comparisons:
        # Build summary text
        summary_parts = []
        if len(sheet_comp.changed_new) > 0:
            summary_parts.append(f"{len(sheet_comp.changed_new)} changed")
        if len(sheet_comp.added) > 0:
            summary_parts.append(f"{len(sheet_comp.added)} added")
        if len(sheet_comp.removed) > 0:
            summary_parts.append(f"{len(sheet_comp.removed)} removed")
        
        summary = ", ".join(summary_parts) if summary_parts else "differences detected"
        
        with st.expander(f"📄 `{sheet_comp.sheet_name}` - {summary}", expanded=True):
            # Show truncation warning
            if sheet_comp.is_truncated:
                st.error("⚠️ **Data Truncation Detected!**")
                st.write(
                    f"OLD file has **{sheet_comp.old_row_count} rows** but "
                    f"NEW file only has **{sheet_comp.new_row_count} rows**"
                )
                st.write(
                    f"**{sheet_comp.old_row_count - sheet_comp.new_row_count} rows** "
                    "of data are missing in the NEW file!"
                )
            
            # Show key information
            if sheet_comp.is_key_based and "Key:" in sheet_comp.key_description:
                st.caption(f"🔑 {sheet_comp.key_description}")
            elif "Row-based" in sheet_comp.key_description:
                st.caption(f"ℹ️ {sheet_comp.key_description}")
            
            # Show changed rows
            if not sheet_comp.changed_new.empty and sheet_comp.is_key_based:
                st.write("**🔄 Changed Rows**")
                preview = build_side_by_side_preview(
                    sheet_comp.changed_old, 
                    sheet_comp.changed_new, 
                    []  # keys already validated
                )
                st.dataframe(preview, use_container_width=True)
                
                # Download button
                excel_file = export_to_excel(
                    sheet_comp.changed_old, 
                    sheet_comp.changed_new, 
                    []
                )
                st.download_button(
                    "📥 Download Excel",
                    excel_file,
                    file_name=f"Pair{result.file_index}_{sheet_comp.sheet_name}_changes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{result.file_index}_{sheet_comp.sheet_name}"
                )
            elif not sheet_comp.changed_new.empty:
                st.write("**🔄 Changed Rows**")
                st.dataframe(sheet_comp.changed_new, use_container_width=True)
            
            # Show added rows
            if not sheet_comp.added.empty:
                st.write("**➕ Added Rows**")
                st.dataframe(style_added_rows(sheet_comp.added), use_container_width=True)
            
            # Show removed rows
            if not sheet_comp.removed.empty:
                st.write("**➖ Removed Rows**")
                st.dataframe(style_removed_rows(sheet_comp.removed), use_container_width=True)

# ----------------------------------------------------------------------
# MAIN APPLICATION
# ----------------------------------------------------------------------

def main():
    """Main application logic"""
    st.title("📊 Smart Diff Manager")
    st.markdown("Compare Excel files with smart header detection and key-based matching.")
    
    # File uploaders
    left_files = st.file_uploader("Upload **OLD** files", ["xlsx", "xls"], accept_multiple_files=True)
    right_files = st.file_uploader("Upload **NEW** files", ["xlsx", "xls"], accept_multiple_files=True)
    
    # Key column input
    keys_str = st.text_input(
        "Key Columns (comma-separated, e.g. ID, Name)", 
        "",
        help="Columns used to match rows between old and new files"
    )
    
    if not (left_files and right_files):
        st.info("👆 Please upload files to start comparison.")
        return
    
    # Matching settings
    col1, col2 = st.columns(2)
    with col1:
        ignore_suffix = st.checkbox(
            "Ignore suffix after last underscore", 
            True,
            help="e.g., 'report_v1.xlsx' matches 'report_v2.xlsx'"
        )
    with col2:
        threshold = st.slider(
            "Auto-match threshold", 
            0.5, 1.0, 0.85, 0.05,
            help="Minimum similarity score to match file pairs"
        )
    
    # Match file pairs
    matched_pairs = match_file_pairs(left_files, right_files, ignore_suffix, threshold)
    
    st.write(f"**{len(matched_pairs)}** file pair(s) matched")
    
    if matched_pairs:
        with st.expander("📋 View matched pairs"):
            for left_file, right_file in matched_pairs:
                st.write(f"• `{left_file.name}` ↔️ `{right_file.name}`")
    
    # Run comparison button
    if not st.button("▶️ Run Comparison", type="primary"):
        return
    
    # Process files
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    keys = [k.strip() for k in keys_str.split(",") if k.strip()]
    results = []
    
    for i, (old_file, new_file) in enumerate(matched_pairs):
        status_text.text(f"Processing pair {i+1}/{len(matched_pairs)}: {old_file.name}")
        
        result = process_file_pair(old_file, new_file, keys, i + 1)
        
        if result:
            results.append(result)
        
        progress_bar.progress((i + 1) / len(matched_pairs))
    
    status_text.empty()
    progress_bar.empty()
    
    # Display results
    if not results:
        st.success(f"✅ All {len(matched_pairs)} file pair(s) are identical!")
    else:
        st.info(f"📊 Found differences in {len(results)} out of {len(matched_pairs)} file pair(s)")
        
        for result in results:
            render_comparison_summary(result, len(matched_pairs))

# ----------------------------------------------------------------------
# RUN APPLICATION
# ----------------------------------------------------------------------

if __name__ == "__main__":
    main()
