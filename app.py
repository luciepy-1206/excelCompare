import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher
from io import BytesIO
from streamlit_sortables import sort_items
import msoffcrypto
import io
from openpyxl.styles import PatternFill
import fitz  # PyMuPDF
import gc   # --- MEMORY MANAGEMENT --- Import Garbage Collector

# --- Page Config ---
st.set_page_config(page_title="📊📄 Unified Diff Manager", layout="wide")
st.markdown("<style>section.main > div {padding-top: 1rem;}</style>", unsafe_allow_html=True)

#<editor-fold desc="Helper Functions">
def normalize_filename(name: str):
    name = name.lower()
    name = re.sub(r'\.[^.]+$', '', name)
    name = re.sub(r'[\s_\-]+', ' ', name)
    return name.strip()

def auto_match(left_files, right_files, threshold):
    potential_matches = []
    for lf in left_files:
        lf_norm = normalize_filename(lf.name)
        for rf in right_files:
            rf_norm = normalize_filename(rf.name)
            score = SequenceMatcher(None, lf_norm, rf_norm).ratio()
            potential_matches.append((score, lf, rf))
    
    potential_matches.sort(reverse=True)
    matched, used_left, used_right = [], set(), set()
    for score, lf, rf in potential_matches:
        if lf not in used_left and rf not in used_right and score >= threshold:
            matched.append((lf, rf, score))
            used_left.add(lf)
            used_right.add(rf)
    unmatched_left = [lf for lf in left_files if lf not in used_left]
    unmatched_right = [rf for rf in right_files if rf not in used_right]
    return matched, unmatched_left, unmatched_right

def manual_pairing(matched, unmatched_left, unmatched_right, left_files, right_files, group_a="OLD", group_b="NEW"):
    st.subheader(f"🎛️ Manual Pairing (Drag to align {group_a} ↔ {group_b})")
    initial_left_names = [lf.name for lf, _, _ in matched] + [lf.name for lf in unmatched_left]
    initial_right_names = [rf.name for _, rf, _ in matched] + [rf.name for rf in unmatched_right]
    
    len_diff = len(initial_left_names) - len(initial_right_names)
    if len_diff > 0:
        initial_right_names.extend(["---"] * len_diff)
    elif len_diff < 0:
        initial_left_names.extend(["---"] * abs(len_diff))
        
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"### 📂 {group_a} Files"); sorted_left_names = sort_items(initial_left_names, key=f"{group_a}_files")
    with c2:
        st.markdown(f"### 🗂️ {group_b} Files"); sorted_right_names = sort_items(initial_right_names, key=f"{group_b}_files")
        
    left_dict, right_dict = {f.name: f for f in left_files}, {f.name: f for f in right_files}
    pairs = [(left_dict[l], right_dict[r]) for l, r in zip(sorted_left_names, sorted_right_names) if l != "---" and r != "---"]
    st.session_state.pairs = pairs
    st.success(f"✅ {len(pairs)} pairs formed for comparison.")
    return pairs

def normalize_value(value):
    if value is None: return ""
    if isinstance(value, str):
        v_upper = value.strip().upper()
        if v_upper in ["TRUE", "T", "YES", "Y", "1"]: return "TRUE"
        if v_upper in ["FALSE", "F", "NO", "N", "0"]: return "FALSE"
        return value.strip()
    return str(value).strip()

def normalize_df(df):
    df = df.fillna("").astype(str)
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
        df[col] = df[col].apply(normalize_value)
    return df

def compare_sheets(df1, df2):
    df1, df2 = normalize_df(df1), normalize_df(df2)
    common_cols = sorted(list(set(df1.columns) & set(df2.columns)))
    added_cols = sorted(list(set(df2.columns) - set(df1.columns)))
    removed_cols = sorted(list(set(df1.columns) - set(df2.columns)))
    
    if not common_cols: return pd.DataFrame(), pd.DataFrame(), [], added_cols, removed_cols
    
    df1['_key'] = df1[common_cols].apply(lambda r: "|".join(r.values.astype(str)), axis=1)
    df2['_key'] = df2[common_cols].apply(lambda r: "|".join(r.values.astype(str)), axis=1)
    
    added_rows = df2[~df2['_key'].isin(df1['_key'])].drop(columns=['_key'])
    removed_rows = df1[~df1['_key'].isin(df2['_key'])].drop(columns=['_key'])

    potential_changes, rem_indices_to_drop, add_indices_to_drop = [], set(), set()
    for rem_idx, rem_row in removed_rows.iterrows():
        best_sim, best_add_idx = 0.7, None
        rem_str = "|".join(rem_row[common_cols].astype(str))
        for add_idx, add_row in added_rows.iterrows():
            if add_idx in add_indices_to_drop: continue
            add_str = "|".join(add_row[common_cols].astype(str))
            sim = SequenceMatcher(None, rem_str, add_str).ratio()
            if sim > best_sim:
                best_sim, best_add_idx = sim, add_idx
        if best_add_idx is not None:
            potential_changes.append({'old_row': rem_row, 'new_row': added_rows.loc[best_add_idx]})
            rem_indices_to_drop.add(rem_idx)
            add_indices_to_drop.add(best_add_idx)

    added_rows = added_rows.drop(index=list(add_indices_to_drop))
    removed_rows = removed_rows.drop(index=list(rem_indices_to_drop))
    return added_rows, removed_rows, potential_changes, added_cols, removed_cols

def decrypt_file_bytes(uploaded_file, password=None):
    file_bytes = io.BytesIO(uploaded_file.getvalue())
    try:
        pd.ExcelFile(file_bytes); file_bytes.seek(0); return file_bytes
    except Exception:
        file_bytes.seek(0)
        try:
            office_file = msoffcrypto.OfficeFile(file_bytes)
            if not office_file.is_encrypted(): return file_bytes
            if not password: raise ValueError("PASSWORD_REQUIRED")
            decrypted_bytes = io.BytesIO()
            office_file.load_key(password=password); office_file.decrypt(decrypted_bytes)
            decrypted_bytes.seek(0); return decrypted_bytes
        except msoffcrypto.exceptions.InvalidKeyError: raise ValueError("BAD_PASSWORD")
        except Exception as e: raise RuntimeError(f"File decryption failed: {e}")
#</editor-fold>

# --- Initialize State ---
if 'view_mode' not in st.session_state:
    st.session_state.view_mode = 'setup'
if 'comparison_results' not in st.session_state:
    st.session_state.comparison_results = None
if "file_passwords" not in st.session_state:
    st.session_state.file_passwords = {}
if 'pairs' not in st.session_state:
    st.session_state.pairs = []
if 'report_buffer' not in st.session_state:
    st.session_state.report_buffer = None

# --- Computation Logic ---
def run_comparison():
    with st.spinner("Comparing files and generating report..."):
        all_results = []
        output_buffer = BytesIO()
        summary_data = []

        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            # --- ROBUSTNESS FIX ---
            # Create a placeholder sheet first. If the app crashes, this
            # prevents the "At least one sheet must be visible" error.
            pd.DataFrame([{"Status": "Starting..."}]).to_excel(writer, sheet_name="Summary", index=False)

            for i, (lf, rf) in enumerate(st.session_state.pairs, 1):
                pair_result = {"pair_index": i, "lf_name": lf.name, "rf_name": rf.name, "sheets": [], "error": None}
                
                try:
                    # ... (rest of the comparison logic is the same)
                    if lf.type != rf.type: raise ValueError("Mismatched file types")
                    
                    pwd_l = st.session_state.file_passwords.get(lf.name)
                    pwd_r = st.session_state.file_passwords.get(rf.name)
                    
                    xls1 = pd.ExcelFile(decrypt_file_bytes(lf, pwd_l))
                    xls2 = pd.ExcelFile(decrypt_file_bytes(rf, pwd_r))
                    common_sheets = sorted(list(set(xls1.sheet_names) & set(xls2.sheet_names)))
                    
                    pair_had_diffs = False
                    for sheet in common_sheets:
                        df1 = pd.read_excel(xls1, sheet_name=sheet)
                        df2 = pd.read_excel(xls2, sheet_name=sheet)
                        added, removed, changes, add_cols, rem_cols = compare_sheets(df1.copy(), df2.copy())
                        
                        if changes or not added.empty or not removed.empty or add_cols or rem_cols:
                            pair_had_diffs = True

                        pair_result["sheets"].append({
                            "name": sheet, "added": added, "removed": removed, 
                            "changes": changes, "add_cols": add_cols, "rem_cols": rem_cols
                        })

                        # Write to Excel report
                        if not added.empty: added.to_excel(writer, sheet_name=f"P{i}_{sheet}_Added"[:31], index=False)
                        if not removed.empty: removed.to_excel(writer, sheet_name=f"P{i}_{sheet}_Removed"[:31], index=False)
                        if changes:
                            old_df = pd.DataFrame([c['old_row'] for c in changes]).reset_index(drop=True)
                            new_df = pd.DataFrame([c['new_row'] for c in changes]).reset_index(drop=True)
                            all_cols = sorted(list(set(old_df.columns) | set(new_df.columns)))
                            old_df, new_df = old_df.reindex(columns=all_cols, fill_value=""), new_df.reindex(columns=all_cols, fill_value="")
                            sheet_name = f"P{i}_{sheet}_Changed"[:31]
                            new_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            ws = writer.sheets[sheet_name]
                            highlight_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                            for r_idx, row in new_df.iterrows():
                                for c_idx, col_name in enumerate(new_df.columns, 1):
                                    if col_name not in rem_cols and row[col_name] != old_df.iloc[r_idx][col_name]:
                                        ws.cell(row=r_idx + 2, column=c_idx).fill = highlight_fill
                    
                    summary_data.append([lf.name, rf.name, "Differences Found" if pair_had_diffs else "No Differences"])

                except Exception as e:
                    pair_result["error"] = str(e)
                    summary_data.append([lf.name, rf.name, f"Error: {e}"])
                
                all_results.append(pair_result)

                # --- MEMORY MANAGEMENT ---
                # Explicitly call the garbage collector to free up memory after processing each large pair.
                gc.collect()

            # Overwrite the placeholder Summary sheet with the final summary
            summary_df = pd.DataFrame(summary_data, columns=["Old File", "New File", "Status"])
            summary_df.to_excel(writer, sheet_name="Summary", index=False)


        st.session_state.comparison_results = all_results
        st.session_state.report_buffer = output_buffer.getvalue()
        st.session_state.view_mode = 'results'

def reset_view():
    st.session_state.view_mode = 'setup'
    st.session_state.comparison_results = None
    st.session_state.pairs = []
    st.session_state.report_buffer = None


# --- PHASE 2: Display Results View ---
if st.session_state.view_mode == 'results':
    st.title("📊📄 Comparison Results")
    st.button("⬅️ Start New Comparison", on_click=reset_view)
    
    for result in st.session_state.comparison_results:
        st.markdown(f"--- \n ### {result['pair_index']}. Comparing `{result['lf_name']}` ↔ `{result['rf_name']}`")
        if result["error"]:
            st.error(f"❌ Error during comparison: {result['error']}")
            continue
        for sheet_result in result["sheets"]:
            is_diff = bool(not sheet_result['added'].empty or not sheet_result['removed'].empty or sheet_result['changes'] or sheet_result['add_cols'] or sheet_result['rem_cols'])
            with st.expander(f"▸ Sheet: `{sheet_result['name']}` {'(Has Differences)' if is_diff else '(No Differences)'}", expanded=is_diff):
                if not is_diff:
                    st.success("✅ No differences found.")
                    continue
                if sheet_result['add_cols']: st.info(f"🟢 Added columns: {', '.join(sheet_result['add_cols'])}")
                if sheet_result['rem_cols']: st.warning(f"🔴 Removed columns: {', '.join(sheet_result['rem_cols'])}")
                if not sheet_result['added'].empty:
                    st.markdown(f"🟢 **{len(sheet_result['added'])} added row(s):**"); st.dataframe(sheet_result['added'])
                if not sheet_result['removed'].empty:
                    st.markdown(f"🔴 **{len(sheet_result['removed'])} removed row(s):**"); st.dataframe(sheet_result['removed'])
                if sheet_result['changes']:
                    st.markdown(f"🟡 **{len(sheet_result['changes'])} changed row(s):**")
                    old_df = pd.DataFrame([c['old_row'] for c in sheet_result['changes']]).reset_index(drop=True)
                    new_df = pd.DataFrame([c['new_row'] for c in sheet_result['changes']]).reset_index(drop=True)
                    all_cols = sorted(list(set(old_df.columns) | set(new_df.columns)))
                    old_df, new_df = old_df.reindex(columns=all_cols, fill_value=""), new_df.reindex(columns=all_cols, fill_value="")
                    def highlight_diffs(new_row):
                        old_row = old_df.iloc[new_row.name]
                        highlight_color = '#334155' 
                        return [f'background-color: {highlight_color}' if new_row[col] != old_row[col] else '' for col in new_row.index]
                    st.dataframe(new_df.style.apply(highlight_diffs, axis=1))

    if st.session_state.report_buffer:
        st.download_button(label="📥 Download Full Report as Excel", data=st.session_state.report_buffer, file_name="comparison_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- PHASE 1: Setup View ---
else:
    st.title("📊📄 Unified Diff & Highlight Manager")
    st.markdown("Upload files to compare. This tool automatically matches them and highlights all differences.")
    col1, col2 = st.columns(2)
    with col1:
        left_files = st.file_uploader("📂 Upload OLD files", type=["xlsx", "xls"], accept_multiple_files=True, key="old_files")
    with col2:
        right_files = st.file_uploader("📂 Upload NEW files", type=["xlsx", "xls"], accept_multiple_files=True, key="new_files")
    if left_files and right_files:
        threshold = st.slider("Auto-match similarity threshold", 0.5, 1.0, 0.8, step=0.05)
        with st.expander("🔑 Enter Passwords (if needed)"):
            for f in left_files + right_files:
                pwd = st.text_input(f"Password for **{f.name}**", value=st.session_state.file_passwords.get(f.name, ""), type="password", key=f"pwd_{f.name}")
                st.session_state.file_passwords[f.name] = pwd
        matched, unmatched_left, unmatched_right = auto_match(left_files, right_files, threshold)
        if matched:
            st.success(f"Auto-matched {len(matched)} pair(s):")
            for lf, rf, sc in matched: st.write(f"› {lf.name} ↔️ {rf.name} ({sc:.1%})")
        manual_pairing(matched, unmatched_left, unmatched_right, left_files, right_files)
        st.button("🚀 Run Comparison", on_click=run_comparison, type="primary", disabled=(not st.session_state.pairs))
    else:
        st.info("👆 Upload files to both OLD and NEW groups to begin.")
