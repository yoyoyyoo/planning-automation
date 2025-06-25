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

if __name__ == "__main__":
    main()
