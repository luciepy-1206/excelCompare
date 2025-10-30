import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
from io import BytesIO
from streamlit_sortables import sort_items
import msoffcrypto
import io
from openpyxl.styles import PatternFill

# --- Page Config ---
st.set_page_config(page_title="📊🚀 Smart Diff Manager", layout="wide")
st.markdown("""
<style>
    [data-theme="light"] { --highlight-background-color: #d4edda; }
    [data-theme="dark"] { --highlight-background-color: #334155; }
</style>
""", unsafe_allow_html=True)

# --- Helper Functions ---
#<editor-fold desc="Helper Functions">
def auto_match(left_files, right_files, threshold):
    potential_matches = []
    for lf in left_files:
        lf_norm = lf.name.lower()
        for rf in right_files:
            rf_norm = rf.name.lower()
            score = SequenceMatcher(None, lf_norm, rf_norm).ratio()
            potential_matches.append((score, lf, rf))
    potential_matches.sort(key=lambda x: x[0], reverse=True)
    
    matched, used_left, used_right = [], set(), set()
    for score, lf, rf in potential_matches:
        if score >= threshold and lf not in used_left and rf not in used_right:
            matched.append((lf, rf, score))
            used_left.add(lf); used_right.add(rf)
    unmatched_left = [f for f in left_files if f not in used_left]
    unmatched_right = [f for f in right_files if f not in used_right]
    return matched, unmatched_left, unmatched_right

def manual_pairing_unmatched(unmatched_left, unmatched_right):
    st.subheader("2. Manually Pair Unmatched Files")
    st.markdown("Drag files to align them. Files left opposite '---' will be ignored.")
    
    left_dict = {f.name: f for f in unmatched_left}
    right_dict = {f.name: f for f in unmatched_right}
    initial_left_names = [f.name for f in unmatched_left]
    initial_right_names = [f.name for f in unmatched_right]

    len_diff = len(initial_left_names) - len(initial_right_names)
    if len_diff > 0: initial_right_names.extend(["---"] * len_diff)
    elif len_diff < 0: initial_left_names.extend(["---"] * abs(len_diff))
        
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"#### 📂 Unmatched OLD Files"); sorted_left_names = sort_items(initial_left_names, key="unmatched_old")
    with c2:
        st.markdown(f"#### 𗂂 Unmatched NEW Files"); sorted_right_names = sort_items(initial_right_names, key="unmatched_new")
        
    pairs = [(left_dict[l], right_dict[r]) for l, r in zip(sorted_left_names, sorted_right_names) if l != "---" and r != "---"]
    if pairs: st.success(f"✅ {len(pairs)} manual pair(s) formed.")
    return pairs

def decrypt_file_bytes(uploaded_file, password=None):
    file_bytes = io.BytesIO(uploaded_file.getvalue())
    try:
        pd.ExcelFile(file_bytes); file_bytes.seek(0); return file_bytes
    except Exception:
        file_bytes.seek(0)
        try:
            office_file = msoffcrypto.OfficeFile(file_bytes)
            if not office_file.is_encrypted(): file_bytes.seek(0); return file_bytes
            if not password: raise ValueError("PASSWORD_REQUIRED")
            decrypted_bytes = io.BytesIO()
            office_file.load_key(password=password); office_file.decrypt(decrypted_bytes)
            decrypted_bytes.seek(0); return decrypted_bytes
        except msoffcrypto.exceptions.InvalidKeyError: raise ValueError("BAD_PASSWORD")
        except Exception: file_bytes.seek(0); return file_bytes

def normalize_df_advanced(df):
    df_norm = df.copy()
    for col in df_norm.columns:
        numeric_col = pd.to_numeric(df_norm[col], errors='coerce')
        numeric_mask = numeric_col.notna()
        if numeric_mask.any():
            df_norm.loc[numeric_mask, col] = numeric_col[numeric_mask].round(9)

    df_norm = df_norm.fillna('').astype(str)
    true_map = {'TRUE', 'T', 'YES', 'Y', '1', '1.0', '✓'}
    false_map = {'FALSE', 'F', 'NO', 'N', '0', '0.0', '✗'}

    for col in df_norm.columns:
        s = df_norm[col].str.strip().str.upper()
        is_true, is_false = s.isin(true_map), s.isin(false_map)
        df_norm.loc[is_true, col] = 'TRUE'
        df_norm.loc[is_false, col] = 'FALSE'
    return df_norm

def compare_sheets_key_based(df1, df2, key_columns):
    """High-performance comparison that identifies Added, Removed, and Changed rows."""
    cols1, cols2 = set(df1.columns), set(df2.columns)
    added_cols, removed_cols = sorted(list(cols2 - cols1)), sorted(list(cols1 - cols2))
    
    for key in key_columns:
        if key not in cols1 or key not in cols2:
            raise ValueError(f"Key column '{key}' not found in both files.")

    df1_norm = normalize_df_advanced(df1)
    df2_norm = normalize_df_advanced(df2)

    merged = pd.merge(df1_norm, df2_norm, on=key_columns, how='outer', suffixes=('_old', '_new'), indicator=True)

    # Use original dataframes for the final output
    added_rows = df2[merged['_merge'] == 'right_only']
    removed_rows = df1[merged['_merge'] == 'left_only']
    
    common_rows = merged[merged['_merge'] == 'both']
    changed_old, changed_new = pd.DataFrame(columns=df1.columns), pd.DataFrame(columns=df2.columns)
    
    if not common_rows.empty:
        diff_mask = pd.Series(False, index=common_rows.index)
        compare_cols = [c for c in df1.columns if c not in key_columns]
        for col in compare_cols:
            if f"{col}_old" in common_rows.columns and f"{col}_new" in common_rows.columns:
                diff_mask |= (common_rows[f'{col}_old'] != common_rows[f'{col}_new'])
        
        if diff_mask.any():
            changed_keys = common_rows[diff_mask][key_columns]
            changed_old = pd.merge(df1, changed_keys, on=key_columns, how='inner')
            changed_new = pd.merge(df2, changed_keys, on=key_columns, how='inner')

    return added_rows, removed_rows, changed_old, changed_new, added_cols, removed_cols
#</editor-fold>

# --- State Initialization ---
if 'view_mode' not in st.session_state: st.session_state.view_mode = 'setup'
# ... (rest of state initialization is the same)
if 'comparison_results' not in st.session_state: st.session_state.comparison_results = None
if "file_passwords" not in st.session_state: st.session_state.file_passwords = {}
if 'pairs' not in st.session_state: st.session_state.pairs = []
if 'report_buffer' not in st.session_state: st.session_state.report_buffer = None
if 'key_columns_str' not in st.session_state: st.session_state.key_columns_str = ""

# --- Core Functions ---
def run_comparison_computation():
    key_columns = [key.strip() for key in st.session_state.key_columns_str.split(',') if key.strip()]
    if not key_columns:
        st.error("Please provide at least one Key Column to identify rows."); return

    status = st.status("Starting comparison...", expanded=True)
    all_results = []
    
    try:
        if not st.session_state.pairs:
            status.update(label="No file pairs formed.", state="error"); return

        for i, (lf, rf) in enumerate(st.session_state.pairs, 1):
            status.write(f"**Pair {i}: `{lf.name}` vs `{rf.name}`**")
            pair_result = {"pair_index": i, "lf_name": lf.name, "rf_name": rf.name, "sheets": [], "error": None}
            
            try:
                status.write("Reading and preparing files...")
                def read_data(file_obj, password):
                    if file_obj.name.lower().endswith('.csv'):
                        return pd.read_csv(file_obj)
                    else:
                        return pd.read_excel(decrypt_file_bytes(file_obj, password))

                df1, df2 = read_data(lf, st.session_state.file_passwords.get(lf.name)), read_data(rf, st.session_state.file_passwords.get(rf.name))
                
                status.write("Comparing content...")
                added, removed, changed_old, changed_new, add_cols, rem_cols = compare_sheets_key_based(df1, df2, key_columns)
                
                pair_result["sheets"].append({
                    "name": "File Content", "added": added, "removed": removed,
                    "changed_old": changed_old, "changed_new": changed_new,
                    "add_cols": add_cols, "rem_cols": rem_cols
                })
            except Exception as e:
                pair_result["error"] = str(e)
            
            all_results.append(pair_result)
        
        status.update(label="Generating Excel report...", state="running")
        st.session_state.report_buffer = generate_excel_report(all_results)
        st.session_state.comparison_results = all_results
        st.session_state.view_mode = 'results'
        status.update(label="Comparison complete!", state="complete", expanded=False)

    except Exception as e:
        status.update(label=f"An error occurred: {e}", state="error")

def generate_excel_report(all_results):
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        any_diffs_found = False
        for result in all_results:
            if result["error"]: continue
            i = result["pair_index"]
            for res in result["sheets"]:
                if not res['added'].empty: res['added'].to_excel(writer, sheet_name=f"P{i}_Added"[:31], index=False); any_diffs_found = True
                if not res['removed'].empty: res['removed'].to_excel(writer, sheet_name=f"P{i}_Removed"[:31], index=False); any_diffs_found = True
                if not res['changed_new'].empty: res['changed_new'].to_excel(writer, sheet_name=f"P{i}_Changed"[:31], index=False); any_diffs_found = True
        
        if not any_diffs_found:
            pd.DataFrame({"Status": ["No differences found."]}) \
              .to_excel(writer, sheet_name="Summary", index=False)
    return output_buffer.getvalue()

def reset_view():
    st.session_state.view_mode = 'setup'

# --- UI: PHASE 2 (RESULTS) ---
if st.session_state.view_mode == 'results':
    st.title("📊🚀 Comparison Results")
    st.button("⬅️ Start New Comparison", on_click=reset_view)
    
    for result in st.session_state.comparison_results:
        st.markdown(f"--- \n ### {result['pair_index']}. `{result['lf_name']}` vs `{result['rf_name']}`")
        if result["error"]:
            st.error(f"❌ Error: {result['error']}"); continue
        
        for res in result["sheets"]:
            is_diff = bool(not res['added'].empty or not res['removed'].empty or not res['changed_new'].empty or res['add_cols'] or res['rem_cols'])
            with st.expander(f"▸ File Content Comparison {'(Differences Found)' if is_diff else '(No Differences)'}", expanded=is_diff):
                if not is_diff:
                    st.success("✅ No differences found."); continue

                if res['add_cols']: st.info(f"🟢 Added columns: {', '.join(res['add_cols'])}")
                if res['rem_cols']: st.warning(f"🔴 Removed columns: {', '.join(res['rem_cols'])}")
                if not res['added'].empty: st.markdown(f"🟢 **{len(res['added'])} Added Rows:**"); st.dataframe(res['added'])
                if not res['removed'].empty: st.markdown(f"🔴 **{len(res['removed'])} Removed Rows:**"); st.dataframe(res['removed'])
                if not res['changed_new'].empty:
                    st.markdown(f"🟡 **{len(res['changed_new'])} Changed Rows:**")
                    
                    key_cols = st.session_state.key_columns_str.split(',')
                    old_df_indexed = res['changed_old'].set_index(key_cols)
                    new_df_styled = res['changed_new'].set_index(key_cols)

                    def highlight_diffs(row):
                        old_row = old_df_indexed.loc[row.name]
                        # Normalize values before final comparison for highlighting
                        old_row_norm = normalize_df_advanced(old_row.to_frame().T).iloc[0]
                        new_row_norm = normalize_df_advanced(row.to_frame().T).iloc[0]
                        return [f'background-color: var(--highlight-background-color)' if str(new_row_norm[col]) != str(old_row_norm[col]) else '' for col in row.index]

                    st.dataframe(new_df_styled.style.apply(highlight_diffs, axis=1))

    if st.session_state.report_buffer:
        st.download_button("📥 Download Full Report", st.session_state.report_buffer, "comparison_report.xlsx")

# --- UI: PHASE 1 (SETUP) ---
else:
    st.title("📊🚀 Smart Diff Manager")
    st.markdown("Compare files by providing a **Key Column** (like an ID) to find changes, additions, and deletions.")
    
    file_types = ["xlsx", "xls", "csv"]
    c1, c2 = st.columns(2)
    with c1: left_files = st.file_uploader("📂 Upload OLD files", type=file_types, accept_multiple_files=True)
    with c2: right_files = st.file_uploader("📂 Upload NEW files", type=file_types, accept_multiple_files=True)
    
    if left_files and right_files:
        st.subheader("⚙️ Configuration (Required)")
        st.text_input(
            "**Key Columns** (comma-separated)", 
            placeholder="e.g., emplid, Transaction ID",
            help="Enter the column name(s) that uniquely identify a row.",
            key='key_columns_str'
        )

        st.subheader("1. Auto-Match Files")
        threshold = st.slider("File name similarity", 0.1, 1.0, 0.8, step=0.05)
        matched, unmatched_left, unmatched_right = auto_match(left_files, right_files, threshold)
        if matched:
            st.success(f"Auto-matched {len(matched)} pair(s).")

        manually_formed_pairs = []
        if unmatched_left or unmatched_right:
            manually_formed_pairs = manual_pairing_unmatched(unmatched_left, unmatched_right)
        
        auto_pairs = [(lf, rf) for lf, rf, _ in matched]
        st.session_state.pairs = auto_pairs + manually_formed_pairs

        st.subheader("3. Final Comparison Queue")
        if st.session_state.pairs:
            st.info(f"{len(st.session_state.pairs)} pair(s) will be compared:")
        else:
            st.warning("No pairs formed.")

        with st.expander("🔑 Enter Passwords (if needed for .xlsx/.xls)"):
            st.markdown("###### OLD Files")
            for f in left_files:
                st.session_state.file_passwords[f.name] = st.text_input(f"Password for **{f.name}**", type="password", key=f"pwd_old_{f.name}")
            st.markdown("###### NEW Files")
            for f in right_files:
                st.session_state.file_passwords[f.name] = st.text_input(f"Password for **{f.name}**", type="password", key=f"pwd_new_{f.name}")
        
        st.button("🚀 Run Comparison", on_click=run_comparison_computation, type="primary", disabled=(not st.session_state.pairs or not st.session_state.key_columns_str))
    else:
        st.info("👆 Upload files to both groups to begin.")
