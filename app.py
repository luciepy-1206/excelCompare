# app.py – Smart Diff Manager (CLEAN UI + Fixed Multi-Key Detection)

import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
from io import BytesIO
import msoffcrypto
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from typing import List, Dict, Tuple, Optional
import re

DEFAULT_PASSWORD = "mypassword"

st.set_page_config(
    page_title="Smart Diff Manager",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------------------------------------------------
# STYLES
# ----------------------------------------------------------------------
st.markdown(
    """
<style>
.dataframe-container { font-size: 14px; }
.small-text { font-size: 13px; opacity: 0.9; margin-left: 1rem; }
</style>
""",
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# Helper functions
# ----------------------------------------------------------------------
def decrypt_file(uploaded_file, password: Optional[str] = DEFAULT_PASSWORD) -> BytesIO:
    fb = BytesIO(uploaded_file.getvalue())
    fb.seek(0)  # Reset before checks
    
    # Quick check if file is Excel openable without decryption
    try:
        pd.ExcelFile(fb)
        fb.seek(0)
        return fb
    except Exception:
        fb.seek(0)
    
    try:
        office = msoffcrypto.OfficeFile(fb)
        if not office.is_encrypted():
            fb.seek(0)
            return fb
        dec = BytesIO()
        office.load_key(password=password)
        office.decrypt(dec)
        dec.seek(0)
        return dec
    except Exception:
        fb.seek(0)
        return fb



# ----------------------------------------------------------------------
# NORMALIZATION UTILITIES
# ----------------------------------------------------------------------
def normalize_colname(name: str) -> str:
    """Normalize column names for comparison - remove special chars, lowercase"""
    return re.sub(r'[^a-z0-9]', '', str(name).lower().strip())


def normalize_logical(v):
    """Normalize various representations of boolean values to TRUE/FALSE"""
    t = {"TRUE","T","YES","Y","1","1.0","CHECK","CHECKED","CHECKMARK","✓","✔","ON","ENABLED","ACTIVE"}
    f = {"FALSE","F","NO","N","0","0.0","CROSS","UNCHECKED","✗","✘","X","OFF","DISABLED","INACTIVE"}
    s = str(v).strip().upper()
    if s in t: return "TRUE"
    if s in f: return "FALSE"
    return s


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize dataframe - fill NaN, convert to string, and normalize boolean values"""
    d = df.copy().fillna("")
    for c in d.columns:
        d[c] = d[c].astype(str).apply(normalize_logical)
    return d


# ----------------------------------------------------------------------
# SMART HEADER DETECTION - Find key columns ANYWHERE in the row
# ----------------------------------------------------------------------
def find_header_row_by_column_names(df_raw: pd.DataFrame, reference_columns: List[str]) -> int:
    """
    Find the header row by matching column names from a reference (OLD file).
    Searches for a row that has the most matching column names.
    """
    if not reference_columns:
        return detect_header_row_heuristic(df_raw)
    
    # Normalize reference column names
    normalized_ref = {normalize_colname(str(c)) for c in reference_columns if str(c).strip()}
    
    best_row = 0
    best_match_count = 0
    
    # Search first 15 rows
    for i in range(min(15, len(df_raw))):
        row = df_raw.iloc[i]
        
        # Get normalized values from this row
        row_values = {
            normalize_colname(str(val)) 
            for val in row 
            if pd.notna(val) and str(val).strip()
        }
        
        # Count how many reference columns match
        match_count = len(normalized_ref & row_values)
        
        # Need at least 50% match and at least 2 matches
        if match_count >= max(2, len(normalized_ref) * 0.5):
            if match_count > best_match_count:
                best_match_count = match_count
                best_row = i
                
                # If we found 80%+ match, use it immediately
                if match_count >= len(normalized_ref) * 0.8:
                    return best_row
    
    # If we found at least some matches, use best match
    if best_match_count >= 2:
        return best_row
    
    # Otherwise fall back to heuristic
    return detect_header_row_heuristic(df_raw)


def find_header_row_with_keys(df_raw: pd.DataFrame, key_columns: List[str]) -> int:
    """
    Search for header row by finding ANY of the key columns ANYWHERE in the row.
    If no key columns found, uses heuristic detection.
    Returns the first row where at least one key column name is found.
    """
    if not key_columns:
        return detect_header_row_heuristic(df_raw)
    
    normalized_keys = {normalize_colname(k) for k in key_columns if k.strip()}
    
    # Search first 15 rows for key columns
    for i in range(min(15, len(df_raw))):
        row = df_raw.iloc[i]
        
        # Check ALL cells in the row
        row_values = []
        for val in row:
            if pd.notna(val) and str(val).strip():
                row_values.append(normalize_colname(str(val)))
        
        # If we find ANY key column name in this row, it's the header
        if normalized_keys & set(row_values):
            return i
    
    # Fallback: No key columns found, use heuristic
    return detect_header_row_heuristic(df_raw)


def detect_header_row_heuristic(df_raw: pd.DataFrame) -> int:
    """
    Detect header row using heuristics when key columns aren't found.
    Looks for a row with mostly text values and good contiguity.
    """
    max_nonempty = df_raw.head(15).apply(lambda r: r.notna().sum(), axis=1).max()
    
    for i, row in df_raw.head(15).iterrows():
        filled = row.notna().sum()
        if filled < 2:  # Need at least 2 values
            continue
        
        # Skip leading empty cells
        non_null_values = [str(val) for val in row if pd.notna(val) and str(val).strip()]
        if len(non_null_values) < 2:
            continue
            
        filled_ratio = filled / max_nonempty if max_nonempty else 0
        text_ratio = sum(isinstance(x, str) and str(x).strip() for x in row) / max(1, filled)
        
        # Calculate contiguous ratio
        valid_indices = [idx for idx, val in enumerate(row) if pd.notna(val) and str(val).strip()]
        if valid_indices:
            span = max(valid_indices) - min(valid_indices) + 1
            contiguous = len(valid_indices) / span
        else:
            contiguous = 0
            
        # Good header: mostly filled, mostly text, mostly contiguous
        if filled_ratio >= 0.4 and text_ratio >= 0.5 and contiguous > 0.5:
            return i
    
    return 0


def read_visible_sheets_with_header_detection(
    file_bytes: BytesIO, 
    key_columns: List[str] = None,
    reference_headers: Dict[str, List[str]] = None
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, int]]:
    """
    Read visible sheets with smart header detection using key columns.
    Only reads VISIBLE columns (ignores hidden columns).
    
    If reference_headers provided (from OLD file), searches for matching column names in NEW file.
    
    Returns: (sheets_dict, header_rows_dict)
    """
    try:
        # Use openpyxl to detect hidden columns
        wb = load_workbook(file_bytes, data_only=True)
        visible_sheets = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
        
        sheets = {}
        header_rows = {}

        for sh in visible_sheets:
            ws = wb[sh]
            
            # Identify visible columns (not hidden)
            visible_col_indices = []
            col_idx = 0
            for col_cells in ws.iter_cols(min_row=1, max_row=1):
                col_letter = col_cells[0].column_letter
                # Check if column is NOT hidden
                col_dim = ws.column_dimensions.get(col_letter)
                if col_dim is None or not col_dim.hidden:
                    visible_col_indices.append(col_idx)
                col_idx += 1
            
            # If no visible columns detected, assume all are visible
            if not visible_col_indices:
                visible_col_indices = list(range(ws.max_column))
            
            # Read the sheet data with pandas (ignores images/charts/textboxes automatically)
            file_bytes.seek(0)
            df_raw = pd.read_excel(file_bytes, sheet_name=sh, header=None, engine="openpyxl")
            
            # Filter to only visible columns
            if visible_col_indices and len(visible_col_indices) < len(df_raw.columns):
                df_raw = df_raw.iloc[:, visible_col_indices]
            
            # Find header row with multiple strategies
            header_row = 0
            
            # Strategy 1: If we have reference headers from OLD file, find matching row
            if reference_headers and sh in reference_headers:
                header_row = find_header_row_by_column_names(df_raw, reference_headers[sh])
            # Strategy 2: Use key columns if provided
            elif key_columns:
                header_row = find_header_row_with_keys(df_raw, key_columns)
            # Strategy 3: Heuristic detection
            else:
                header_row = detect_header_row_heuristic(df_raw)
            
            # Read again with proper header row
            file_bytes.seek(0)
            df_full = pd.read_excel(file_bytes, sheet_name=sh, header=header_row, engine="openpyxl")
            
            # Filter to only visible columns
            if visible_col_indices and len(visible_col_indices) < len(df_full.columns):
                df_main = df_full.iloc[:, visible_col_indices]
            else:
                df_main = df_full

            # Flatten multi-index headers if present
            if isinstance(df_main.columns, pd.MultiIndex):
                df_main.columns = [
                    " ".join([str(c) for c in col if str(c) != "nan"]).strip() 
                    for col in df_main.columns
                ]
            
            # Convert all columns to string to avoid Arrow serialization errors
            df_main = df_main.astype(str)

            sheets[sh] = df_main
            header_rows[sh] = header_row

        return sheets, header_rows

    except Exception as e:
        # Fallback: Try to read with pandas only
        try:
            file_bytes.seek(0)
            all_sheets_raw = pd.read_excel(file_bytes, sheet_name=None, engine="openpyxl", header=None)
            
            sheets = {}
            header_rows = {}
            
            for sh, df_raw in all_sheets_raw.items():
                # Find header row
                if reference_headers and sh in reference_headers:
                    header_row = find_header_row_by_column_names(df_raw, reference_headers[sh])
                elif key_columns:
                    header_row = find_header_row_with_keys(df_raw, key_columns)
                else:
                    header_row = detect_header_row_heuristic(df_raw)
                
                # Read with proper header
                file_bytes.seek(0)
                df = pd.read_excel(file_bytes, sheet_name=sh, header=header_row, engine="openpyxl")
                
                # Convert to string to avoid Arrow errors
                df = df.astype(str)
                
                sheets[sh] = df
                header_rows[sh] = header_row
            
            return sheets, header_rows
            
        except Exception:
            file_bytes.seek(0)
            all_sheets = pd.read_excel(file_bytes, sheet_name=None, engine="openpyxl")
            # Convert all to string
            for sh in all_sheets:
                all_sheets[sh] = all_sheets[sh].astype(str)
            return all_sheets, {}


# ----------------------------------------------------------------------
# KEY SELECTION - Find best key that exists in both files
# ----------------------------------------------------------------------
def find_best_valid_key(df1: pd.DataFrame, df2: pd.DataFrame, keys: List[str]) -> Tuple[List[str], str]:
    """
    Find the best valid key column that exists in both dataframes.
    Checks each key individually and picks the one with highest uniqueness.
    """
    if not keys:
        return [], "No keys provided"
    
    # Build normalized column mapping
    norm1 = {normalize_colname(c): c for c in df1.columns}
    norm2 = {normalize_colname(c): c for c in df2.columns}
    
    best_key = None
    best_uniqueness = 0
    best_key_name = None
    
    # Try each key individually
    for k in keys:
        k_stripped = k.strip()
        if not k_stripped:
            continue
            
        nk = normalize_colname(k_stripped)
        
        # Check if this key exists in BOTH files
        if nk in norm1 and nk in norm2:
            actual_col = norm1[nk]
            
            try:
                # Calculate uniqueness score
                unique1 = df1[actual_col].fillna("").astype(str).nunique()
                unique2 = df2[actual_col].fillna("").astype(str).nunique()
                
                uniqueness1 = unique1 / max(1, len(df1))
                uniqueness2 = unique2 / max(1, len(df2))
                avg_uniqueness = (uniqueness1 + uniqueness2) / 2
                
                if avg_uniqueness > best_uniqueness:
                    best_uniqueness = avg_uniqueness
                    best_key = [actual_col]
                    best_key_name = k_stripped
            except Exception:
                continue
    
    if best_key:
        return best_key, f"Key: '{best_key_name}' ({best_uniqueness:.0%} unique)"
    
    return [], "No valid keys found"


# ----------------------------------------------------------------------
# COMPARISON FUNCTIONS
# ----------------------------------------------------------------------
def compare_key_based(df1, df2, keys):
    """Compare two dataframes using key columns"""
    valid_keys, key_desc = find_best_valid_key(df1, df2, keys)
    if not valid_keys:
        raise ValueError("No valid key columns found")

    common = list(df1.columns.intersection(df2.columns))
    df1f, df2f = df1[common].copy(), df2[common].copy()

    # Fill NaN with empty string
    df1f = df1f.fillna("")
    df2f = df2f.fillna("")

    n1, n2 = normalize_df(df1f), normalize_df(df2f)
    n1["__key__"] = n1[valid_keys].astype(str).agg("||".join, axis=1)
    n2["__key__"] = n2[valid_keys].astype(str).agg("||".join, axis=1)
    df1f["__key__"], df2f["__key__"] = n1["__key__"], n2["__key__"]

    k1, k2 = set(n1["__key__"]), set(n2["__key__"])
    common_keys, added_keys, removed_keys = k1 & k2, k2 - k1, k1 - k2

    changed_old, changed_new = [], []
    for key in common_keys:
        r1 = n1.loc[n1["__key__"] == key]
        r2 = n2.loc[n2["__key__"] == key]
        if r1.empty or r2.empty:
            continue
        r1, r2 = r1.iloc[0], r2.iloc[0]
        if any(str(r1[c]) != str(r2[c]) for c in common if c not in valid_keys and c != "__key__"):
            changed_old.append(df1f.loc[df1f["__key__"] == key].iloc[0])
            changed_new.append(df2f.loc[df2f["__key__"] == key].iloc[0])

    co = pd.DataFrame(changed_old)[common] if changed_old else pd.DataFrame(columns=common)
    cn = pd.DataFrame(changed_new)[common] if changed_new else pd.DataFrame(columns=common)
    added = df2f.loc[df2f["__key__"].isin(added_keys), common].reset_index(drop=True) if added_keys else pd.DataFrame(columns=common)
    removed = df1f.loc[df1f["__key__"].isin(removed_keys), common].reset_index(drop=True) if removed_keys else pd.DataFrame(columns=common)
    
    return co.reset_index(drop=True), cn.reset_index(drop=True), added, removed, key_desc


def compare_keyless(df1, df2):
    """Compare two dataframes without key columns - detects added and removed rows"""
    n1, n2 = normalize_df(df1), normalize_df(df2)
    h1, h2 = pd.util.hash_pandas_object(n1, index=False), pd.util.hash_pandas_object(n2, index=False)
    
    added = df2.loc[~h2.isin(h1)].copy().reset_index(drop=True)
    removed = df1.loc[~h1.isin(h2)].copy().reset_index(drop=True)
    
    return added, removed


def detect_data_truncation(df1: pd.DataFrame, df2: pd.DataFrame) -> Tuple[bool, int, int]:
    """
    Detect if NEW file has fewer non-empty rows than OLD file (data truncation).
    Returns: (is_truncated, old_row_count, new_row_count)
    """
    # Count non-empty rows (rows with at least one non-null value)
    old_non_empty = df1.apply(lambda row: row.notna().any(), axis=1).sum()
    new_non_empty = df2.apply(lambda row: row.notna().any(), axis=1).sum()
    
    # Consider truncated if NEW has significantly fewer rows (at least 5 rows or 10% less)
    threshold = max(5, old_non_empty * 0.1)
    is_truncated = (old_non_empty - new_non_empty) >= threshold
    
    return is_truncated, old_non_empty, new_non_empty


# ----------------------------------------------------------------------
# PREVIEW & EXPORT WITH HIGHLIGHTING
# ----------------------------------------------------------------------
def build_side_by_side_preview(changed_old: pd.DataFrame, changed_new: pd.DataFrame, keys: List[str]):
    """Creates a side-by-side comparison with color-coded cells"""
    common_cols = changed_old.columns.intersection(changed_new.columns).tolist()
    old_df = changed_old[common_cols].copy()
    new_df = changed_new[common_cols].copy()
    
    old_df_norm = normalize_df(old_df)
    new_df_norm = normalize_df(new_df)
    
    combined_data = []
    for idx in range(len(old_df)):
        row_data = {}
        for col in common_cols:
            old_val = str(old_df.iloc[idx][col])
            new_val = str(new_df.iloc[idx][col])
            row_data[f"{col} (Old)"] = old_val
            row_data[f"{col} (New)"] = new_val
        combined_data.append(row_data)
    
    combined_df = pd.DataFrame(combined_data)
    
    def highlight_changes(row):
        styles = []
        row_idx = row.name
        for i, col in enumerate(combined_df.columns):
            if col.endswith(" (Old)"):
                base_col = col[:-6]
                if base_col in old_df_norm.columns:
                    old_val_norm = str(old_df_norm.iloc[row_idx][base_col])
                    new_val_norm = str(new_df_norm.iloc[row_idx][base_col])
                    
                    if old_val_norm != new_val_norm:
                        styles.append('background-color: #ffcccc; color: #000000')
                    else:
                        styles.append('background-color: #e8f5e9; color: #000000')
                else:
                    styles.append('')
            elif col.endswith(" (New)"):
                base_col = col[:-6]
                if base_col in new_df_norm.columns:
                    old_val_norm = str(old_df_norm.iloc[row_idx][base_col])
                    new_val_norm = str(new_df_norm.iloc[row_idx][base_col])
                    
                    if old_val_norm != new_val_norm:
                        styles.append('background-color: #ccffcc; color: #000000')
                    else:
                        styles.append('background-color: #e8f5e9; color: #000000')
                else:
                    styles.append('')
            else:
                styles.append('')
        return styles
    
    styled_df = combined_df.style.apply(highlight_changes, axis=1)
    return styled_df


def style_added_rows(df: pd.DataFrame):
    def highlight(row):
        return ['background-color: #e3f2fd; color: #000000' for _ in row]
    return df.style.apply(highlight, axis=1)


def style_removed_rows(df: pd.DataFrame):
    def highlight(row):
        return ['background-color: #ffebee; color: #000000' for _ in row]
    return df.style.apply(highlight, axis=1)


def export_to_excel(changed_old: pd.DataFrame, changed_new: pd.DataFrame, keys: List[str]) -> BytesIO:
    wb = Workbook()
    ws_old = wb.active
    ws_old.title = "Old Values"
    ws_new = wb.create_sheet(title="New Values")
    
    old_highlight = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    new_highlight = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

    changed_old_norm = normalize_df(changed_old)
    changed_new_norm = normalize_df(changed_new)

    for c, col_name in enumerate(changed_old.columns, 1):
        ws_old.cell(row=1, column=c, value=col_name)
    for r_idx in range(len(changed_old)):
        for c_idx, col_name in enumerate(changed_old.columns, 1):
            cell = ws_old.cell(row=r_idx + 2, column=c_idx, value=changed_old.iloc[r_idx][col_name])
            if (col_name not in keys and 
                str(changed_old_norm.iloc[r_idx][col_name]) != str(changed_new_norm.iloc[r_idx][col_name])):
                cell.fill = old_highlight

    for c, col_name in enumerate(changed_new.columns, 1):
        ws_new.cell(row=1, column=c, value=col_name)
    for r_idx in range(len(changed_new)):
        for c_idx, col_name in enumerate(changed_new.columns, 1):
            cell = ws_new.cell(row=r_idx + 2, column=c_idx, value=changed_new.iloc[r_idx][col_name])
            if (col_name not in keys and 
                str(changed_old_norm.iloc[r_idx][col_name]) != str(changed_new_norm.iloc[r_idx][col_name])):
                cell.fill = new_highlight

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def unpack_sheetdata(sheetdata):
    if len(sheetdata) == 10:
        return sheetdata
    else:
        # Pad missing elements for backwards compatibility
        return (*sheetdata, False, 0, 0)
# ----------------------------------------------------------------------
# UI AND EXECUTION
# ----------------------------------------------------------------------
st.title("📊 Smart Diff Manager")
st.markdown("Compare Excel files with smart header detection and key-based matching.")

left = st.file_uploader("Upload **OLD** files", ["xlsx", "xls"], accept_multiple_files=True)
right = st.file_uploader("Upload **NEW** files", ["xlsx", "xls"], accept_multiple_files=True)
keys_str = st.text_input("Key Columns (comma-separated, e.g. ID, Name)", "")

if left and right:
    ignore_suffix = st.checkbox("Ignore suffix after last underscore", True)
    threshold = st.slider("Auto-match threshold", 0.5, 1.0, 0.85, 0.05)

    def comparable(n, ignore):
        n = n.rsplit(".", 1)[0]
        return n.rsplit("_", 1)[0].lower() if ignore and ("_" in n) else n.lower()

    matched = []
    usedL, usedR = set(), set()
    for lf in left:
        lcmp = comparable(lf.name, ignore_suffix)
        best, ratio = None, 0
        for rf in right:
            if rf.name in usedR:
                continue
            rcmp = comparable(rf.name, ignore_suffix)
            r = SequenceMatcher(None, lcmp, rcmp).ratio()
            if r >= threshold and r > ratio:
                best, ratio = rf, r
        if best:
            matched.append((lf, best))
            usedL.add(lf.name)
            usedR.add(best.name)

    st.write(f"**{len(matched)}** file pair(s) matched")

    if matched:
        with st.expander("📋 View matched pairs"):
            for lf, rf in matched:
                st.write(f"• `{lf.name}` ↔️ `{rf.name}`")

        if st.button("▶️ Run Comparison"):
            prog = st.progress(0)
            keys = [k.strip() for k in keys_str.split(",") if k.strip()]
            files_with_changes = []
            
            for i, (lf, rf) in enumerate(matched):
                try:
                    decL, decR = decrypt_file(lf), decrypt_file(rf)
                    shL, header_rows_old = read_visible_sheets_with_header_detection(decL, key_columns=keys)
                    reference_headers = {sheet_name: list(df.columns) for sheet_name, df in shL.items()}
                    shR, header_rows_new = read_visible_sheets_with_header_detection(
                        decR,
                        key_columns=keys,
                        reference_headers=reference_headers
                    )

                    # Debug header detection issues here as needed
                except Exception as e:
                    st.error(f"❌ File pair {i+1}/{len(matched)}: Failed to read files: {e}")
                    import traceback
                    st.code(traceback.format_exc())
                    prog.progress((i + 1) / len(matched))
                    continue

                changed_sheets = []
                for sh in sorted(set(shL) & set(shR)):
                    d1, d2 = shL[sh], shR[sh]
                    use_key_based = False
                    key_desc = ""
                    if keys:
                        try:
                            valid_keys, key_desc = find_best_valid_key(d1, d2, keys)
                            use_key_based = len(valid_keys) > 0
                        except Exception:
                            use_key_based = False

                    try:
                        if use_key_based:
                            co, cn, add, rem, key_desc = compare_key_based(d1, d2, keys)
                        else:
                            co = cn = pd.DataFrame()
                            add, rem = compare_keyless(d1, d2)
                    except Exception:
                        co = cn = pd.DataFrame()
                        add, rem = compare_keyless(d1, d2)

                    if not (cn.empty and add.empty and rem.empty):
                        changed_sheets.append((sh, co, cn, add, rem, use_key_based, key_desc))

                if changed_sheets:
                    files_with_changes.append((i + 1, lf.name, rf.name, changed_sheets))
                prog.progress((i + 1) / len(matched))

            if not files_with_changes:
                st.success(f"✅ All {len(matched)} file pair(s) are identical!")
            else:
                st.info(f"📊 Found differences in {len(files_with_changes)} out of {len(matched)} file pair(s)")
                for file_num, old_name, new_name, changed_sheets in files_with_changes:
                    st.markdown(f"## 📂 File Pair {file_num}/{len(matched)}")
                    st.markdown(f"**OLD:** `{old_name}`  \n**NEW:** `{new_name}`")
                    
                    for sheet_data in changed_sheets:
                        # Safe unpack of sheet_data
                        sheet_data = unpack_sheetdata(sheet_data)
                        (
                            sh, co, cn, add, rem,
                            is_key_based, key_desc,
                            is_truncated, old_rows, new_rows
                        ) = sheet_data

                        summary_parts = []
                        if len(cn) > 0:
                            summary_parts.append(f"{len(cn)} changed")
                        if len(add) > 0:
                            summary_parts.append(f"{len(add)} added")
                        if len(rem) > 0:
                            summary_parts.append(f"{len(rem)} removed")

                        summary = ", ".join(summary_parts) if summary_parts else "differences detected"

                        with st.expander(f"📄 `{sh}` - {summary}", expanded=True):
                            if is_truncated:
                                st.error(f"⚠️ **Data Truncation Detected!**")
                                st.write(f"OLD file has **{old_rows} rows** but NEW file only has **{new_rows} rows**")
                                st.write(f"**{old_rows - new_rows} rows** of data are missing in the NEW file!")

                            if is_key_based and key_desc and "Key:" in key_desc:
                                st.caption(f"🔑 {key_desc}")
                            elif key_desc and "Row-based" in key_desc:
                                st.caption(f"ℹ️ {key_desc} (no matching key column found)")

                            if not cn.empty and is_key_based:
                                st.write("**🔄 Changed Rows**")
                                preview = build_side_by_side_preview(co, cn, keys)
                                st.dataframe(preview)  # Removed width=None
                                excel_file = export_to_excel(co, cn, keys)
                                st.download_button(
                                    "📥 Download Excel",
                                    excel_file,
                                    file_name=f"Pair{file_num}_{sh}_changes.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"download_{file_num}_{sh}"
                                )
                            elif not cn.empty:
                                st.write("**🔄 Changed Rows**")
                                st.dataframe(cn)  # Removed width=None
                            if not add.empty:
                                st.write("**➕ Added Rows**")
                                st.dataframe(style_added_rows(add))  # Removed width=None
                            if not rem.empty:
                                st.write("**➖ Removed Rows**")
                                st.dataframe(style_removed_rows(rem))  # Removed width=None
else:
    st.info("Please upload files to start comparison.")
