"""
Streamlit UI for the ECI election-result PDF extractor.

Upload one or more "DETAILED RESULTS" PDFs, watch the pages being parsed, review
a short summary of what came out, and download one Excel workbook per PDF -- as a
zip when several were uploaded.

Run with:
    streamlit run app.py
"""

import io
import re
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from pdf_to_excel import extract_records, to_dataframe, validate

SAMPLE_FORMAT_IMAGE = Path(__file__).with_name("sample_format.png")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="Election PDF to Excel", page_icon="📊", layout="wide")


def excel_name(pdf_name):
    """'Punjab 2012 AE Result.pdf' -> 'Punjab 2012 AE Result.xlsx'."""
    stem = Path(pdf_name).stem.strip().lstrip(".")
    stem = re.sub(r'[\\/:*?"<>|]', "_", stem)      # illegal on Windows / in zips
    return f"{stem or 'results'}.xlsx"


def to_excel_bytes(df):
    """Render one table as a .xlsx in memory."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")
    return buffer.getvalue()


def zip_of(workbooks):
    """Bundle (filename, bytes) pairs into a single zip, keeping names unique."""
    buffer = io.BytesIO()
    used = set()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, blob in workbooks:
            # Two uploads can share a name; don't let one silently overwrite the other.
            unique, counter = name, 2
            while unique in used:
                unique = f"{Path(name).stem} ({counter}).xlsx"
                counter += 1
            used.add(unique)
            archive.writestr(unique, blob)
    return buffer.getvalue()


def for_display(df):
    """Make a copy Streamlit can serialise without complaint.

    Columns that mix numbers with the "-" placeholder are of object dtype, which
    Arrow cannot type; Streamlit would fall back to a string cast anyway and log a
    traceback each rerun. Casting here keeps "-" visible and the log quiet. Only
    the preview is affected -- the downloaded workbooks keep their real types.
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
**Your PDF must be an ECI *DETAILED RESULTS* report**, laid out like the sample on
the left. Page size and margins do not matter — the column positions are measured
from each file you upload.

**Each page needs**

- The column header row: `CANDIDATE NAME · SEX · AGE · CATEGORY · PARTY`,
  optionally `SYMBOL`, then `GENERAL · POSTAL · TOTAL` under **VALID VOTES
  POLLED**, then `% VOTES POLLED`
- A `Constituency  <n>. <NAME>   TOTAL ELECTORS : <number>` line starting each seat
- One numbered row per candidate, ending with a `TURNOUT TOTAL:` line

**Also required**

- A **text-based** PDF, not a scan or photo. If you cannot select text in a PDF
  reader, it will not work — it needs OCR first.
- No password protection.

**Good to know**

- Reports without a `SYMBOL` column (e.g. Punjab 2012) work; those rows get `-`.
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

    # Each PDF keeps its own table: it is downloaded as its own workbook, so the
    # fixed 11 columns stay exactly as they are, with no provenance column added.
    df = to_dataframe(records)
    frames.append((upload.name, df))

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

# One combined table purely for the on-screen preview and headline numbers; the
# downloads below stay per-file.
data = pd.concat(
    [df.assign(**{"Source File": name}) if len(frames) > 1 else df for name, df in frames],
    ignore_index=True,
)

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

view = data
if "Source File" in data.columns:
    # Constituency names can repeat across states, so narrow by file first.
    picked = st.selectbox("Filter by file", ["All"] + [name for name, _ in frames])
    if picked != "All":
        view = view[view["Source File"] == picked]

choice = st.selectbox("Filter by constituency", ["All"] + sorted(view["Constituency"].unique()))
if choice != "All":
    view = view[view["Constituency"] == choice]

st.dataframe(for_display(view), hide_index=True, width="stretch", height=420)
st.caption(f"Showing {len(view):,} of {len(data):,} rows.")

# ---- download ------------------------------------------------------------
st.subheader("Download")

workbooks = [(excel_name(source), to_excel_bytes(df)) for source, df in frames]

if len(workbooks) == 1:
    name, blob = workbooks[0]
    st.download_button(
        f"⬇️ Download {name}", blob, file_name=name, mime=XLSX_MIME, width="stretch",
    )
else:
    st.download_button(
        f"⬇️ Download all {len(workbooks)} Excel files (.zip)",
        zip_of(workbooks), file_name="election_results.zip",
        mime="application/zip", width="stretch",
    )
    st.caption("One workbook per PDF, named after the source file:")
    for name, blob in workbooks:
        st.download_button(
            f"⬇️ {name}", blob, file_name=name, mime=XLSX_MIME, key=f"dl-{name}",
        )

footer()
