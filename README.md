# Election PDF → Excel

Extracts candidate-level results from Election Commission of India **DETAILED RESULTS**
PDFs into a clean Excel/CSV table — via a Streamlit UI or the command line.

Verified against two full reports: Madhya Pradesh **2018** (230 constituencies, 3,129
candidates) and **2013** (230 constituencies, 2,813 candidates). Every extracted figure
reconciles exactly with the totals the PDFs print for themselves.

---

## Output format

| Constituency | Candidate Name | Sex | Age | Category | Party | Postal Votes | Total Valid Votes | % Votes Polled | Symbol | Total Electors |
|---|---|---|---|---|---|---|---|---|---|---|
| SHEOPUR | BABU JANDEL | M | 47 | GEN | INC | 712 | 98580 | 55.17 | Hand | 224700 |
| SHEOPUR | None of the Above | - | - | - | NOTA | 29 | 1794 | 1.00 | NOTA | 224700 |

Any cell the PDF leaves blank — Sex, Age and Category on *None of the Above* rows — is
written as `-` rather than left empty.

---

## Install

```bash
pip install -r requirements.txt
```

## Run the web UI

```bash
streamlit run app.py
```

Upload one or more PDFs, watch the per-page progress bar, review the summary, then
download CSV or Excel.

## Run on the command line

```bash
python pdf_to_excel.py                          # defaults to the bundled filenames
python pdf_to_excel.py input.pdf output.xlsx    # writes output.xlsx and output.csv
```

```
constituencies : 230
candidate rows : 3129
written        : MP_Election_2018.xlsx
validation     : all rows consistent (general+postal=total, serials 1..n)
```

---

## Which PDFs work

An ECI **DETAILED RESULTS** report, laid out like `sample_format.png`. The parser reads
columns by their position on the page, so the layout matters more than the year or state.
Each page must carry:

- the header row `CANDIDATE NAME · SEX · AGE · CATEGORY · PARTY · SYMBOL`, then
  `GENERAL · POSTAL · TOTAL` under **VALID VOTES POLLED**, then `% VOTES POLLED`
- a `Constituency <n>. <NAME>   TOTAL ELECTORS : <number>` line opening each seat
- numbered candidate rows, closed by a `TURNOUT TOTAL:` line

The PDF must be **text-based**. If you cannot select text in a PDF reader, it is a scan
and needs OCR first. Password-protected files are not supported.

---

## How it works

Plain text extraction breaks on these reports: long candidate names and multi-word
symbols wrap onto extra lines that read as separate records. So the parser works
geometrically instead.

1. **Columns by x-position.** Every column occupies a stable horizontal band on every
   page (`COLUMNS` in `pdf_to_excel.py`). Each word is assigned to the band containing
   its centre.
2. **Rows by y-position.** Words within 3pt vertically form one visual line.
3. **Records vs. continuations.** A line carrying a serial number starts a new candidate;
   any other line is wrapped text stitched back into the record above it.
4. **Mid-word wraps.** Usually a cell breaks between words (`Pressure` / `Cooker`), but
   when a single token is wider than its column the PDF splits it with no hyphen
   (`COCONU` / `T FARM`, party `NINSHA` / `D`). Geometry cannot distinguish the two, so
   the real cases are listed explicitly in `FRAGMENT_JOINS`.

### Adapting to a new PDF

If a report contains mid-word splits not yet listed, find them with:

```bash
python pdf_to_excel.py new_report.pdf --audit
```

This prints every wrap point landing within one character of a column edge — the only
place a mid-word break can occur — alongside how it is currently handled:

```
  symbol     39.8 (edge 43.3)  'COCONU' + 'T'        -> JOINED
  symbol     39.8 (edge 43.3)  'Ploughing' + 'within' -> spaced
```

Anything mis-handled goes into `FRAGMENT_JOINS`.

---

## Accuracy checks

Every run validates itself, and reports anything that fails:

- `general + postal = total` on each candidate row
- serial numbers running `1..n` within each constituency, no gaps
- no empty candidate names

Both reference PDFs also reconcile to the printed grand totals exactly:

| Report | Constituencies | Candidates | General | Postal | Total |
|---|---|---|---|---|---|
| MP 2018 | 230 | 3,129 | 37,850,051 | 287,477 | 38,137,528 |
| MP 2013 | 230 | 2,813 | 33,604,006 | 248,498 | 33,852,504 |

---

## Files

| File | Purpose |
|---|---|
| `pdf_to_excel.py` | Parser and CLI. Import `extract_records`, `to_dataframe`, `validate` to use it as a library. |
| `app.py` | Streamlit UI: multi-file upload, progress, preview, downloads. |
| `sample_format.png` | Reference page shown on the upload screen. |

Source PDFs and generated CSV/Excel files are gitignored — they are large and
regenerable.

---

<p align="center">Created by the <strong>AIML Team</strong> with ♥</p>
