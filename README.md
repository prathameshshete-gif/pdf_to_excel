# Election PDF → Excel

Extracts candidate-level results from Election Commission of India **DETAILED RESULTS**
PDFs into a clean Excel/CSV table — via a Streamlit UI or the command line.

The column layout is measured from each PDF rather than assumed, so reports with
different page sizes and different sets of columns all work. Verified against three full
reports — Madhya Pradesh **2018** and **2013**, and Punjab **2012** (A4, and with no
SYMBOL column at all). Every extracted figure reconciles exactly with the totals the
PDFs print for themselves.

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

An ECI **DETAILED RESULTS** report, laid out like `sample_format.png`. Page size and
margins do not matter — they are measured per file. Each page must carry:

- the header row `CANDIDATE NAME · SEX · AGE · CATEGORY · PARTY`, optionally `SYMBOL`,
  then `GENERAL · POSTAL · TOTAL` under **VALID VOTES POLLED**, then `% VOTES POLLED`
- a `Constituency <n>. <NAME>   TOTAL ELECTORS : <number>` line opening each seat
- numbered candidate rows, closed by a `TURNOUT TOTAL:` line

`SYMBOL` is optional: Punjab 2012 omits it, and those rows come out with `Symbol` set
to `-`. The 11 output columns stay the same either way.

The PDF must be **text-based**. If you cannot select text in a PDF reader, it is a scan
and needs OCR first. Password-protected files are not supported.

---

## How it works

Plain text extraction breaks on these reports: long candidate names and multi-word
symbols wrap onto extra lines that read as separate records. So the parser works
geometrically instead.

1. **Columns discovered per PDF.** The page header names which columns exist (so a
   missing `SYMBOL` is detected rather than assumed), and the bands themselves are
   measured by projecting every candidate-row word onto the x-axis: the stripes that
   stay empty across the whole document are the gutters between columns. Each word is
   then assigned to the band containing its centre. Nothing about page size, margins or
   column count is hardcoded — which is what lets one parser handle Letter-size MP
   reports and A4 Punjab ones alike.
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
| Punjab 2012 | 117 | 1,078 | 13,892,711 | 8,713 | 13,901,424 |

Per-constituency `TURNOUT TOTAL` lines match on all 577 constituencies across the three.

A PDF whose layout cannot be read raises a `LayoutError` naming the problem — a missing
header row, or a column count that disagrees with the header — instead of silently
producing shifted data.

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
