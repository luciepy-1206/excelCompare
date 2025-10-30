import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
from io import BytesIO
from streamlit_sortables import sort_items
import msoffcrypto
import io

# --- Page Config ---
st.set_page_config(page_title="📊🚀 Smart Diff Manager", layout="wide")
st.markdown("""
<style>
    [data-theme="light"] { --highlight-background-color: #d4edda; }
    [data-theme="dark"] { --highlight-background-color: #334155; }
</style>
""", unsafe_allow_html=True)

# --- Helper Functions (Normalization is the key change) ---
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

def manual_pairing(matched, unmatched_left, unmatched_right, left_files, right_files):
    st.subheader(f"🎛️ Manually Align Files (unpaired files will be ignored)")
    initial_left_names = [lf.name for lf, _, _ in matched] + [f.name for f in unmatched_left]
    initial_right_names = [rf.name for _, rf, _ in matched] + [f.name for f in unmatched_right]

    len_diff = len(initial_left_names) - len(initial_right_names)
    if len_diff > 0: initial_right_names.extend(["---"] * len_diff)
    elif len_diff < 0: initial_left_names.extend(["---"] * abs(len_diff))
        
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"#### 📂 OLD Files"); sorted_left_names = sort_items(initial_left_names, key="old_files")
    with c2:
        st.markdown(f"#### 𗂂 NEW Files"); sorted_right_names = sort_items(initial_right_names, key="new_files")
        
    left_dict = {f.name: f for f in left_files}
    right_dict = {f.name: f for f in right_files}
    
    pairs = [(left_dict[l], right_dict[r]) for l, r in zip(sorted_left_names, sorted_right_names) if l != "---" and r != "---"]
    st.session_state.pairs = pairs
    st.success(f"✅ {len(pairs)} pairs formed for comparison.")
    return pairs

def decrypt_file_bytes(uploaded_file, password=None):
    file_bytes = io.BytesIO(uploaded_file.getvalue())
    try:
        pd.ExcelFile(file_bytes); file_bytes.seek(0); return file_bytes
    except Exception:
        file_bytes.seek(0)
        try:
            office_file = msoffcrypto.OfficeFile(file_bytes)
            if not office_file.is_encrypted():
                file_bytes.seek(0); return file_bytes
            if not password: raise ValueError("PASSWORD_REQUIRED")
            decrypted_bytes = io.BytesIO()
            office_file.load_key(password=password); office_file.decrypt(decrypted_bytes)
            decrypted_bytes.seek(0); return decrypted_bytes
        except msoffcrypto.exceptions.InvalidKeyError: raise ValueError("BAD_PASSWORD")
        except Exception:
            file_bytes.seek(0); return file_bytes

# --- THIS IS THE ROBUST, CORRECTED NORMALIZATION FUNCTION ---
def sanitize_value(x):
    """A robust function to clean and standardize a single cell value."""
    # Define boolean maps
    true_map = {'TRUE', 'T', 'YES', 'Y', '1', '1.0', '✓'}
    false_map = {'FALSE', 'F', 'NO', 'N', '0', '0.0', '✗'}

    # If it's already a number (int, float), round it
    if isinstance(x, (int, float)):
        return round(x, 9)

    # Convert to string for all other operations
    s = str(x).strip()
    s_upper = s.upper()

    # Check for boolean representation
    if s_upper in true_map:
        return "TRUE"
    if s_upper in false_map:
        return "FALSE"

    # Try to convert to a number and round it (catches numbers stored as text)
    try:
        return round(float(s), 9)
    except (ValueError, TypeError):
        # If it's not a number, it's a string. Clean it.
        # This removes encoded characters like the `_x000d_` you found
        if isinstance(s, str):
            return s.replace('_x000d_', '').replace('\n', ' ').replace('\r', ' ')
        return s

def normalize_df_advanced(df):
    """Applies the robust sanitizer to every cell in the DataFrame."""
    return df.applymap(sanitize_value)

def compare_sheets_keyless(df1, df2):
    """High-performance comparison using row hashing after advanced normalization."""
    cols1, cols2 = set(df1.columns), set(df2.columns)
    added_cols, removed_cols = sorted(list(cols2 - cols1)), sorted(list(cols1 - cols2))
    
    if df1.empty: return df2, pd.DataFrame(columns=df1.columns), added_cols, removed_cols
    if df2.empty: return pd.DataFrame(columns=df2.columns), df1, added_cols, removed_cols

    df1_norm = normalize_df_advanced(df1)
    df2_norm = normalize_df_advanced(df2)

    df1_hashes = pd.util.hash_pandas_object(df1_norm, index=False)
    df2_hashes = pd.util.hash_pandas_object(df2_norm, index=False)

    added_mask = ~df2_hashes.isin(df1_hashes)
    removed_mask = ~df1_hashes.isin(df2_hashes)

    return df2[added_mask], df1[removed_mask], added_cols, removed_cols
#</editor-fold>

# --- State Initialization ---
if 'view_mode' not in st.session_state: st.session_state.view_mode = 'setup'
if 'comparison_results' not in st.session_state: st.session_state.comparison_results = None
if "file_passwords" not in st.session_state: st.session_state.file_passwords = {}
if 'pairs' not in st.session_state: st.session_state.pairs = []
if 'report_buffer' not in st.session_state: st.session_state.report_buffer = None

# --- Core Functions ---
def run_comparison_computation():
    status = st.status("Starting comparison...", expanded=True)
    all_results = []
    
    try:
        if not st.session_state.pairs:
            status.update(label="No file pairs were formed. Please align files to compare.", state="error"); return

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

                df1 = read_data(lf, st.session_state.file_passwords.get(lf.name))
                df2 = read_data(rf, st.session_state.file_passwords.get(rf.name))
                
                status.write("Comparing content...")
                added, removed, add_cols, rem_cols = compare_sheets_keyless(df1, df2)
                
                pair_result["sheets"].append({
                    "name": "File Content", "added": added, "removed": removed,
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
                if not res['added'].empty:
                    res['added'].to_excel(writer, sheet_name=f"P{i}_Added"[:31], index=False); any_diffs_found = True
                if not res['removed'].empty:
                    res['removed'].to_excel(writer, sheet_name=f"P{i}_Removed"[:31], index=False); any_diffs_found = True
        
        if not any_diffs_found:
            pd.DataFrame({"Status": ["No differences found across all compared files."]}) \
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
            is_diff = bool(not res['added'].empty or not res['removed'].empty or res['add_cols'] or res['rem_cols'])
            with st.expander(f"▸ File Content Comparison {'(Differences Found)' if is_diff else '(No Differences)'}", expanded=is_diff):
                if not is_diff:
                    st.success("✅ No differences found."); continue

                if res['add_cols']: st.info(f"🟢 Added columns: {', '.join(res['add_cols'])}")
                if res['rem_cols']: st.warning(f"🔴 Removed columns: {', '.join(res['rem_cols'])}")
                if not res['added'].empty: st.markdown(f"🟢 **{len(res['added'])} Added Rows:**"); st.dataframe(res['added'])
                if not res['removed'].empty: st.markdown(f"🔴 **{len(res['removed'])} Removed Rows:**"); st.dataframe(res['removed'])

    if st.session_state.report_buffer:
        st.download_button("📥 Download Full Report", st.session_state.report_buffer, "comparison_report.xlsx")

# --- UI: PHASE 1 (SETUP) ---
else:
    st.title("📊🚀 Smart Diff Manager")
    st.markdown("Handles boolean formats (✓, 1, TRUE) and minor numeric differences automatically.")
    
    file_types = ["xlsx", "xls", "csv"]
    c1, c2 = st.columns(2)
    with c1: left_files = st.file_uploader("📂 Upload OLD files", type=file_types, accept_multiple_files=True)
    with c2: right_files = st.file_uploader("📂 Upload NEW files", type=file_types, accept_multiple_files=True)
    
    if left_files and right_files:
        st.subheader("1. Auto-Match Files")
        threshold = st.slider("File name similarity", 0.1, 1.0, 0.8, step=0.05)
        matched, unmatched_left, unmatched_right = auto_match(left_files, right_files, threshold)
        if matched:
            st.success(f"Auto-matched {len(matched)} pair(s).")

        pairs_formed = manual_pairing(matched, unmatched_left, unmatched_right, left_files, right_files)

        with st.expander("🔑 Enter Passwords (if needed for .xlsx/.xls)"):
            st.markdown("###### OLD Files")
            for f in left_files:
                st.session_state.file_passwords[f.name] = st.text_input(f"Password for **{f.name}**", type="password", key=f"pwd_old_{f.name}")
            st.markdown("###### NEW Files")
            for f in right_files:
                st.session_state.file_passwords[f.name] = st.text_input(f"Password for **{f.name}**", type="password", key=f"pwd_new_{f.name}")
        
        st.button("🚀 Run Comparison", on_click=run_comparison_computation, type="primary", disabled=(not pairs_formed))
    else:
        st.info("👆 Upload Excel or CSV files to both groups to begin.")
