import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher
from io import BytesIO
from streamlit_sortables import sort_items
import msoffcrypto
import io

st.set_page_config(page_title="Excel Smart Diff Manager", layout="wide")
st.title("📊 Excel Smart Diff Manager (Auto Match + Manual Order)")
st.markdown("""
Upload multiple Excel files on both sides — automatically matches and compares content,  
even when rows are shuffled or TRUE/FALSE become 1/0/Y/N.  
""")

def decrypt_if_needed_and_return_bytes(uploaded_file, password=None):
    """
    Returns a BytesIO object of the Excel file.
    Raises ValueError("PASSWORD_REQUIRED") if password is needed but not provided.
    Raises ValueError("BAD_PASSWORD") if the password is wrong.
    """
    raw = uploaded_file.read()
    try:
        _ = pd.ExcelFile(io.BytesIO(raw))
        return io.BytesIO(raw)  # unencrypted
    except Exception:
        bio = io.BytesIO(raw)
        office = msoffcrypto.OfficeFile(bio)
        if not office.is_encrypted():
            return io.BytesIO(raw)  # some other read error
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

# ---------- Main ----------
left_files = st.file_uploader("📂 Upload OLD Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
right_files = st.file_uploader("📂 Upload NEW Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

if left_files and right_files:
    matched, unmatched_left, unmatched_right = auto_match(left_files, right_files)

    if matched:
        st.success(f"Auto matched {len(matched)} pair(s):")
        for lf, rf, sc in matched:
            st.text(f"{lf.name} ↔ {rf.name} ({sc:.0%})")
    else:
        st.warning("No automatic matches found.")

    pairs = manual_pairing(left_files, right_files)

    if st.button("🚀 Run Smart Comparison"):
        progress = st.progress(0)
        total = len(pairs)
        output_buffer = BytesIO()

        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            for i, (lf, rf) in enumerate(pairs, 1):
                progress.progress(i / total)
                st.write(f"### {i}. Comparing {lf.name} ↔ {rf.name}")
                try:
                    xls1 = pd.ExcelFile(lf)
                    xls2 = pd.ExcelFile(rf)
                    common_sheets = set(xls1.sheet_names).intersection(xls2.sheet_names)
                    any_diff = False

                    for sheet in common_sheets:
                        st.subheader(f"📄 Sheet: `{sheet}`")
                        df1 = pd.read_excel(lf, sheet_name=sheet)
                        df2 = pd.read_excel(rf, sheet_name=sheet)
                        added, removed = compare_sheets(df1, df2)

                        if added.empty and removed.empty:
                            st.success("✅ No differences found.")
                        else:
                            any_diff = True
                            if not added.empty:
                                st.info(f"🟢 Added {len(added)} row(s):")
                                st.dataframe(added)
                                added.to_excel(writer, sheet_name=f"{lf.name}_ADDED", index=False)
                            if not removed.empty:
                                st.warning(f"🔴 Removed {len(removed)} row(s):")
                                st.dataframe(removed)
                                removed.to_excel(writer, sheet_name=f"{lf.name}_REMOVED", index=False)

                    if not any_diff:
                        st.success(f"🎉 {lf.name} identical to {rf.name}")

                except Exception as e:
                    st.error(f"❌ Error comparing {lf.name} ↔ {rf.name}: {e}")

        progress.empty()

        st.download_button(
            label="📥 Download Differences as Excel",
            data=output_buffer.getvalue(),
            file_name="all_differences.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("👆 Upload both OLD and NEW Excel files to start.")

