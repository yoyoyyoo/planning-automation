import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re

st.set_page_config(page_title="Transcript → EE/CE Updater", layout="centered")

def extract_courses_from_pdf(pdf_file):
    """Extract course codes from transcript PDF."""
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    course_codes = []
    for page in doc:
        text = page.get_text()
        matches = re.findall(r"\b([A-Z]{4}\s\d{3})\b", text)
        course_codes.extend(matches)
    doc.close()
    return sorted(set(course_codes))

def dedup_columns(cols):
    seen = {}
    result = []
    for col in cols:
        if col not in seen:
            seen[col] = 1
            result.append(col)
        else:
            count = seen[col]
            new_col = f"{col}.{count}"
            while new_col in seen:
                count += 1
                new_col = f"{col}.{count}"
            seen[col] += 1
            seen[new_col] = 1
            result.append(new_col)
    return result

# === MAIN APP ===
def main():
    st.title("📘 Transcript → EE/CE Progress Updater")
    st.write("Upload transcript and progress file — auto update Flag column")

    # STEP 1 — Upload Transcript PDF
    transcript_pdf = st.file_uploader("Step 1: Upload Student Transcript (PDF)", type=["pdf"])

    if transcript_pdf:
        with st.spinner("Extracting courses from transcript..."):
            course_codes = extract_courses_from_pdf(transcript_pdf)

        if not course_codes:
            st.error("❌ No courses detected — check transcript format.")
            return

        st.success(f"✅ Detected {len(course_codes)} courses from transcript.")
        with st.expander("View detected courses"):
            st.write(", ".join(course_codes))

        # STEP 2 — Select Program
        program_type = st.radio("Step 2: Select Program Type:", options=["EE", "CE"], horizontal=True)

        # STEP 3 — Upload EE/CE Progress File
        progress_file = st.file_uploader(f"Step 3: Upload {program_type} Progress File (Excel)", type=["xlsx"])

        if progress_file:
            with st.spinner("Reading progress file..."):
                progress_sheets = pd.ExcelFile(progress_file).sheet_names
                sheet_mapping = {
                    'eleceng': 'EE', 'ee': 'EE', 'electrical': 'EE',
                    'compeng': 'CE', 'ce': 'CE', 'computer': 'CE'
                }
                sheet_name = next((name for name in progress_sheets if name.strip().lower() in sheet_mapping.keys()), None)

                if not sheet_name:
                    st.error(f"❌ Could not detect EE/CE sheet — available: {progress_sheets}")
                    return

                raw_df = pd.read_excel(progress_file, sheet_name=sheet_name, header=None)
                section_titles = raw_df.iloc[:, 0].astype(str)

                # === AUTO-DETECT HEADER ROW ===
                header_row_idx = None
                for i in range(0, 15):  # check first 15 rows
                    row_values = raw_df.iloc[i].astype(str).str.lower().tolist()
                    if any("flag" in val for val in row_values):
                        header_row_idx = i
                        break

                if header_row_idx is None:
                    st.error("❌ Could not find header row with 'Flag' column!")
                    return

                df = pd.read_excel(progress_file, sheet_name=sheet_name, header=header_row_idx)
                df.columns = dedup_columns(df.columns)

                flag_columns = df.columns[df.columns.str.contains("flag", case=False, na=False)]
                if len(flag_columns) == 0:
                    st.error("❌ Could not find any 'Flag' columns after detecting header!")
                    return

                flag_col = flag_columns[0]

                updated_df = df.copy()
                updated_df[flag_col] = pd.to_numeric(updated_df[flag_col], errors="coerce").fillna(0).astype(int)

                # STEP 4 — Auto Update Flags
                st.header("Step 4: Auto Update Progress File")

                course_code_pattern = re.compile(r"\b([A-Z]{4}\s?\d{3})\b")

                match_count = 0
                for i, row in updated_df.iterrows():
                    course_match = course_code_pattern.search(str(row[0]))
                    if course_match:
                        course_code = course_match.group(1).strip()
                        if course_code in course_codes:
                            if updated_df.at[i, flag_col] != 1:
                                updated_df.at[i, flag_col] = 1
                                match_count += 1

                st.success(f"✅ {match_count} courses matched and Flag updated.")

                with st.expander("View updated Progress File preview"):
                    st.dataframe(updated_df.head(20))

                # === Download updated file ===
                out_filename = f"Updated_{program_type}_Progress.xlsx"
                out_excel = pd.ExcelWriter(out_filename, engine='xlsxwriter')
                updated_df.to_excel(out_excel, sheet_name=sheet_name, index=False)
                out_excel.close()

                # Save to memory for download
                with open(out_filename, "rb") as f:
                    st.download_button(
                        "📥 Download Updated Progress File",
                        f,
                        file_name=out_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

if __name__ == "__main__":
    main()
