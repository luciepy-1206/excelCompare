# excel_diff_manager.py
import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher
from io import BytesIO
from streamlit_sortables import sort_items
import msoffcrypto
import io

st.set_page_config(page_title="📊 Excel Smart Diff Manager", layout="wide")
st.markdown("<style>section.main > div {padding-top: 1rem;}</style>", unsafe_allow_html=True)

st.title("📊 Excel Smart Diff Manager (Auto Match + Manual Order)")
st.markdown("""
Upload multiple Excel files on both sides — automatically matches and compares content,  
even when rows are shuffled or TRUE/FALSE become 1/0/Y/N.  
Supports password-protected Excel files (enter password below when required).
""")

# ---------- Helper: Normalize filenames ----------
def normalize_filename(name: str):
    name = name.lower()
    name = re.sub(r'[\s_\-+]+', ' ', name)
    name = re.sub(r'\.[^.]+$', '', name)
    return name.strip()

# ---------- Helper: Auto match files by name ----------
def auto_match(left_files, right_files):
    matched = []
    unmatched_left = []
    unmatched_right = list(right_files)

    for lf in left_files:
        best_score = 0
        best_match = None
        lf_norm = normalize_filename(lf.name)
        for rf in unmatched_right:
            rf_norm = normalize_filename(rf.name)
            score = SequenceMatcher(None, lf_norm, rf_norm).ratio()
            if score > best_score:
                best_score = score
                best_match = rf
        if best_score >= 0.5 and best_match:
            matched.append((lf, best_match, best_score))
            unmatched_right.remove(best_match)
        else:
            unmatched_left.append(lf)
    return matched, unmatched_left, unmatched_right

# ---------- Helper: Normalize boolean-like values ----------
def normalize_bool_like(value):
    if value is None:
        return ""
    if isinstance(value, str):
        v = value.strip().upper()
        if v in ["TRUE", "T", "YES", "Y", "✓", "1"]:
            return "TRUE"
        elif v in ["FALSE", "F", "NO", "N", "✗", "0"]:
            return "FALSE"
        else:
            return value.strip()
    elif isinstance(value, (int, float)):
        if value == 1:
            return "TRUE"
        elif value == 0:
            return "FALSE"
        else:
            return str(value)
    return str(value).strip()

# ---------- Normalize dataframe ----------
def normalize_df(df):
    df = df.fillna("").astype(str)
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
        df[col] = df[col].apply(normalize_bool_like)
    return df

# ---------- Compare sheets smartly ----------
def compare_sheets(df1, df2):
    df1 = normalize_df(df1)
    df2 = normalize_df(df2)
    df1["_key"] = df1.apply(lambda row: "|".join(row.values.astype(str)), axis=1)
    df2["_key"] = df2.apply(lambda row: "|".join(row.values.astype(str)), axis=1)
    added_rows = df2.loc[~df2["_key"].isin(df1["_key"])].drop(columns=["_key"])
    removed_rows = df1.loc[~df1["_key"].isin(df2["_key"])].drop(columns=["_key"])
    return added_rows, removed_rows

# ---------- Helper: Decrypt if password-protected ----------
def decrypt_if_needed_and_return_bytes(uploaded_file, password=None):
    raw = uploaded_file.read()
    try:
        _ = pd.ExcelFile(io.BytesIO(raw))
        return io.BytesIO(raw)  # unencrypted
    except Exception:
        bio = io.BytesIO(raw)
        office = msoffcrypto.OfficeFile(bio)
        if not office.is_encrypted():
            return io.BytesIO(raw)
        if not password:
            raise ValueError("PASSWORD_REQUIRED")
        decrypted = io.BytesIO()
        office.load_key(password=password)
        try:
            office.decrypt(decrypted)
        except Exception:
            raise ValueError("BAD_PASSWORD")
        decrypted.seek(0)
        return decrypted

# ---------- Manual pairing UI ----------
def manual_pairing(left_files, right_files):
    st.subheader("🎛️ Manual Pairing (Drag to align OLD ↔ NEW)")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 📂 OLD Files")
        sorted_left = sort_items([f.name for f in left_files], key="old_files")
    with c2:
        st.markdown("### 🗂️ NEW Files")
        sorted_right = sort_items([f.name for f in right_files], key="new_files")
    left_sorted = [next(x for x in left_files if x.name == name) for name in sorted_left]
    right_sorted = [next(x for x in right_files if x.name == name) for name in sorted_right]
    pairs = list(zip(left_sorted, right_sorted))
    st.success(f"✅ {len(pairs)} manual pairs formed.")
    return pairs

# ---------- UI: Upload files ----------
col1, col2 = st.columns(2)
with col1:
    left_files = st.file_uploader("📂 Upload OLD Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
with col2:
    right_files = st.file_uploader("📂 Upload NEW Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

# session storage for passwords (ephemeral)
if "file_passwords" not in st.session_state:
    st.session_state["file_passwords"] = {}  # key: filename -> password

# Show password inputs for files
def show_password_inputs(files):
    if not files:
        return
    for f in files:
        key = f"pwd_{f.name}"
        existing = st.session_state["file_passwords"].get(f.name, "")
        pwd = st.text_input(f"Password for {f.name} (leave empty if none)", value=existing, type="password", key=key)
        st.session_state["file_passwords"][f.name] = pwd

show_password_inputs(left_files)
show_password_inputs(right_files)

# ---------- Main flow ----------
if left_files and right_files:

    matched, unmatched_left, unmatched_right = auto_match(left_files, right_files)

    if matched:
        st.success(f"Auto matched {len(matched)} pair(s):")
        for lf, rf, sc in matched:
            st.text(f"{lf.name} ↔ {rf.name} ({sc:.0%})")
    else:
        st.warning("No automatic matches found.")

    pairs = manual_pairing(left_files, right_files)

    if st.button("🚀 Run Smart Comparison (supports password-protected files)"):
        progress = st.progress(0)
        total = len(pairs) if pairs else 1
        output_buffer = BytesIO()
        any_diff_overall = False

        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            # placeholder sheet to avoid "At least one sheet must be visible"
            writer.book.create_sheet("Summary")

            for i, (lf, rf) in enumerate(pairs, 1):
                progress.progress(i / total)
                st.write(f"### {i}. Comparing {lf.name} ↔ {rf.name}")
                pwd_l = st.session_state["file_passwords"].get(lf.name) or None
                pwd_r = st.session_state["file_passwords"].get(rf.name) or None

                try:
                    bf_l = decrypt_if_needed_and_return_bytes(lf, password=pwd_l)
                    bf_r = decrypt_if_needed_and_return_bytes(rf, password=pwd_r)
                except ValueError as ve:
                    if str(ve) == "PASSWORD_REQUIRED":
                        st.error(f"🔒 {lf.name if 'lf' in locals() else rf.name} is password-protected — please enter its password and run again.")
                        continue
                    if str(ve) == "BAD_PASSWORD":
                        st.error(f"❌ Wrong password for {lf.name if 'lf' in locals() else rf.name}. Please correct it and run again.")
                        continue
                    continue
                except Exception as e:
                    st.error(f"❌ Unexpected error opening {lf.name if 'lf' in locals() else rf.name}: {e}")
                    continue

                try:
                    xls1 = pd.ExcelFile(bf_l)
                    xls2 = pd.ExcelFile(bf_r)
                except Exception as e:
                    st.error(f"❌ Failed to parse Excel structure for the pair: {e}")
                    continue

                common_sheets = set(xls1.sheet_names).intersection(xls2.sheet_names)
                any_diff_for_pair = False

                for sheet in common_sheets:
                    st.markdown(f"📄 **Sheet:** `{sheet}`")
                    try:
                        df1 = pd.read_excel(bf_l, sheet_name=sheet)
                        df2 = pd.read_excel(bf_r, sheet_name=sheet)
                    except Exception as e:
                        st.error(f"❌ Error reading sheet `{sheet}`: {e}")
                        continue

                    added, removed = compare_sheets(df1, df2)

                    if added.empty and removed.empty:
                        st.success("✅ No differences in this sheet.")
                    else:
                        any_diff_for_pair = True
                        any_diff_overall = True
                        if not added.empty:
                            st.info(f"🟢 {len(added)} added row(s) in NEW file:")
                            st.dataframe(added)
                            safe_name = f"P{i}_{rf.name}_ADDED_{sheet}"[:31]
                            try:
                                added.to_excel(writer, sheet_name=safe_name, index=False)
                            except Exception:
                                added.to_excel(writer, sheet_name=f"P{i}_ADDED_{sheet}"[:31], index=False)
                        if not removed.empty:
                            st.warning(f"🔴 {len(removed)} removed row(s) from OLD file:")
                            st.dataframe(removed)
                            safe_name = f"P{i}_{lf.name}_REMOVED_{sheet}"[:31]
                            try:
                                removed.to_excel(writer, sheet_name=safe_name, index=False)
                            except Exception:
                                removed.to_excel(writer, sheet_name=f"P{i}_REMOVED_{sheet}"[:31], index=False)

                if not any_diff_for_pair:
                    st.success(f"🎉 {lf.name} identical to {rf.name}")

            # finalize writer
            if any_diff_overall:
                if "Summary" in writer.book.sheetnames and len(writer.book.sheetnames) > 1:
                    del writer.book["Summary"]
            else:
                pd.DataFrame([["No differences found across compared files."]]).to_excel(
                    writer, sheet_name="Summary", index=False, header=False
                )

        progress.empty()

        if any_diff_overall:
            st.download_button(
                label="📥 Download Differences as Excel",
                data=output_buffer.getvalue(),
                file_name="all_differences.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.success("✅ No differences found across all files.")
else:
    st.info("👆 Upload both OLD and NEW Excel files to start.")
