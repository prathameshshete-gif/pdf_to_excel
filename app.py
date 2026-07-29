"""
Streamlit UI for the ECI election-result PDF extractor.

Upload one or more "DETAILED RESULTS" PDFs, watch the pages being parsed, review
a short summary of what came out, and download the combined result as CSV/Excel.

Run with:
    streamlit run app.py
"""

import io
from pathlib import Path

import pandas as pd
import streamlit as st

from pdf_to_excel import extract_records, to_dataframe, validate

SAMPLE_FORMAT_IMAGE = Path(__file__).with_name("sample_format.png")

st.set_page_config(page_title="Election PDF to Excel", page_icon="📊", layout="wide")


def for_display(df):
    """Make a copy Streamlit can serialise without complaint.

    Columns that mix numbers with the "-" placeholder are of object dtype, which
    Arrow cannot type; Streamlit would fall back to a string cast anyway and log a
    traceback each rerun. Casting here keeps "-" visible and the log quiet. Only
    the preview is affected -- the CSV/Excel downloads keep their real types.
    """
    display = df.copy()
    for column in display.columns:
        if display[column].dtype == object:
            display[column] = display[column].astype(str)
    return display


def show_required_format(expanded):
    """Show what a valid input PDF looks like, so nobody guesses at the format."""
    with st.expander("📄 Required PDF format — click to see a sample page", expanded=expanded):
        left, right = st.columns([3, 2])

        with left:
            if SAMPLE_FORMAT_IMAGE.exists():
                st.image(
                    str(SAMPLE_FORMAT_IMAGE),
                    caption="A valid page: Election Commission of India — DETAILED RESULTS",
                    width="stretch",
                )
            else:
                st.warning(f"Sample image not found at {SAMPLE_FORMAT_IMAGE.name}")

        with right:
            st.markdown(
                """
**Your PDF must be an ECI *DETAILED RESULTS* report**, laid out exactly like the
sample on the left. The parser reads each column by its position on the page, so
the layout matters more than the year or the state.

**Each page needs**

- The column header row: `CANDIDATE NAME · SEX · AGE · CATEGORY · PARTY · SYMBOL`
  then `GENERAL · POSTAL · TOTAL` under **VALID VOTES POLLED**, then `% VOTES POLLED`
- A `Constituency  <n>. <NAME>   TOTAL ELECTORS : <number>` line starting each seat
- One numbered row per candidate, ending with a `TURNOUT TOTAL:` line

**Also required**

- A **text-based** PDF, not a scan or photo. If you cannot select text in a PDF
  reader, it will not work — it needs OCR first.
- No password protection.

**Good to know**

- Wrapped names and symbols spanning two or three lines are stitched back together.
- Blank cells (Sex/Age/Category on *None of the Above* rows) come out as `-`.
- Every extraction is checked against the `TURNOUT TOTAL` and `GRAND TOTAL`
  figures printed in the PDF itself, and anything that disagrees is reported.
"""
            )


def footer():
    """Credit line pinned to the bottom of every screen the app can end on."""
    st.divider()
    st.markdown(
        "<p style='text-align:center; opacity:0.65; font-size:0.9rem;'>"
        "Created by the <strong>AIML Team</strong> with <span style='color:#e25555;'>&hearts;</span>"
        "</p>",
        unsafe_allow_html=True,
    )


st.title("📊 Election PDF → Excel")
st.caption(
    "Upload ECI *DETAILED RESULTS* PDFs. Every candidate row is extracted by column "
    "position, cross-checked against the totals printed in the PDF itself, and any "
    "value the PDF leaves blank is written as `-`."
)

# Shown above the uploader so the format is clear before anyone picks a file, and
# open by default until something is uploaded -- then it folds away to leave room
# for the results.
show_required_format(expanded="uploads" not in st.session_state or not st.session_state.get("uploads"))

uploads = st.file_uploader(
    "PDF files", type="pdf", accept_multiple_files=True, key="uploads",
    help="You can select several PDFs at once; results are combined into one table.",
)

if not uploads:
    st.info("👆 Choose one or more PDF files to begin.")
    footer()
    st.stop()

# ---- extract -------------------------------------------------------------
progress = st.progress(0.0, text="Starting…")
frames, summaries = [], []

for file_index, upload in enumerate(uploads):
    def report(page_no, total_pages, _i=file_index, _name=upload.name):
        # Each file gets an equal slice of the overall bar.
        done = (_i + page_no / total_pages) / len(uploads)
        progress.progress(done, text=f"{_name} — page {page_no} of {total_pages}")

    try:
        records, warnings = extract_records(io.BytesIO(upload.getvalue()), on_page=report)
    except Exception as exc:                      # a non-matching or damaged PDF
        st.error(f"**{upload.name}** could not be read: {exc}")
        continue

    if not records:
        st.warning(f"**{upload.name}** — no candidate rows found. Is this an ECI detailed-results PDF?")
        continue

    df = to_dataframe(records)
    if len(uploads) > 1:
        # Provenance only when it is actually ambiguous -- appended last so the
        # fixed 11-column layout stays intact.
        df["Source File"] = upload.name
    frames.append(df)

    problems, n_constituencies = validate(records)
    summaries.append({
        "File": upload.name,
        "Constituencies": n_constituencies,
        "Candidates": len(df),
        "Parties": df["Party"].nunique(),
        "Total Votes": int(pd.to_numeric(df["Total Valid Votes"], errors="coerce").sum()),
        "Checks": "✅ passed" if not (problems or warnings) else f"⚠️ {len(problems) + len(warnings)} issue(s)",
        "_issues": problems + warnings,
    })

progress.empty()

if not frames:
    footer()
    st.stop()

data = pd.concat(frames, ignore_index=True)

# ---- summary -------------------------------------------------------------
st.subheader("Extraction summary")
st.dataframe(
    for_display(pd.DataFrame(summaries).drop(columns="_issues")),
    hide_index=True, width="stretch",
)

for summary in summaries:
    if summary["_issues"]:
        with st.expander(f"⚠️ Issues in {summary['File']}"):
            for issue in summary["_issues"]:
                st.text(issue)

left, mid, right = st.columns(3)
left.metric("Rows extracted", f"{len(data):,}")
mid.metric("Constituencies", f"{data['Constituency'].nunique():,}")
right.metric("Total valid votes", f"{pd.to_numeric(data['Total Valid Votes'], errors='coerce').sum():,.0f}")

# ---- preview -------------------------------------------------------------
st.subheader("Extracted data")

constituencies = ["All"] + sorted(data["Constituency"].unique())
choice = st.selectbox("Filter by constituency", constituencies)
view = data if choice == "All" else data[data["Constituency"] == choice]

st.dataframe(for_display(view), hide_index=True, width="stretch", height=420)
st.caption(f"Showing {len(view):,} of {len(data):,} rows.")

# ---- download ------------------------------------------------------------
st.subheader("Download")

csv_bytes = data.to_csv(index=False).encode("utf-8")

excel_buffer = io.BytesIO()
with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    data.to_excel(writer, index=False, sheet_name="Results")

col_csv, col_xlsx = st.columns(2)
col_csv.download_button(
    "⬇️ Download CSV", csv_bytes,
    file_name="election_results.csv", mime="text/csv", width="stretch",
)
col_xlsx.download_button(
    "⬇️ Download Excel", excel_buffer.getvalue(),
    file_name="election_results.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    width="stretch",
)

footer()
