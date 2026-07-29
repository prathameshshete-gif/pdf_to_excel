"""
Madhya Pradesh Legislative Election 2018 -- PDF -> Excel/CSV extractor.

The ECI "DETAILED RESULTS" PDF is a fixed-layout report: every column sits in a
stable band of x-coordinates on every page. So instead of parsing the flat text
(where wrapped candidate names and symbols become ambiguous extra lines), we read
the words with their coordinates, bucket them into columns by x-position, and
bucket them into rows by y-position. Wrapped text is stitched back into the row
that started the record.

Output columns are fixed:
    Constituency, Candidate Name, Sex, Age, Category, Party, Postal Votes,
    Total Valid Votes, % Votes Polled, Symbol, Total Electors

Any column with no value in the PDF (e.g. Sex/Age/Category on "None of the Above"
rows) is written as "-".

Usage:
    python pdf_to_excel.py                      # uses the defaults below
    python pdf_to_excel.py input.pdf out.xlsx
"""

import re
import sys
from collections import defaultdict

import pandas as pd
import pdfplumber

DEFAULT_PDF = "Madhya Pradesh Legislative Election 2018.pdf"
DEFAULT_XLSX = "MP_Election_2018.xlsx"

MISSING = "-"

# The data columns, in the order they appear, keyed by the word that labels them
# in the page header. Column positions are NOT hardcoded: page size and the set of
# columns vary between reports (Punjab 2012 has no SYMBOL column and is A4 rather
# than Letter), so the layout is measured from each PDF -- see detect_layout.
HEADER_ANCHORS = [
    ("name",     "CANDIDATE"),
    ("sex",      "SEX"),
    ("age",      "AGE"),
    ("category", "CATEGORY"),
    ("party",    "PARTY"),
    ("symbol",   "SYMBOL"),     # absent from some reports
    ("general",  "GENERAL"),
    ("postal",   "POSTAL"),
    ("total",    "TOTAL"),
    ("percent",  "POLLED"),
]

# Whitespace this wide between two columns counts as the gutter separating them.
# Measured range that works across the reference reports is 1.5-3.0pt: below that
# the gaps between words inside a cell start to register, above it the narrowest
# real gutter (party/symbol in MP 2018, 3.4pt) is missed.
MIN_GUTTER = 2.5

# Two words are on the same visual line if their tops differ by less than this.
ROW_TOLERANCE = 3.0

HEADER_WORDS = {
    "Election", "DETAILED", "RESULTS", "VALID", "VOTES", "POLLED",
    "CANDIDATE", "NAME", "SEX", "AGE", "CATEGORY", "PARTY", "SYMBOL",
    "GENERAL", "POSTAL", "TOTAL", "%",
}

CONSTITUENCY_RE = re.compile(r"Constituency\s+(\d+)\.\s*(.+?)\s+TOTAL\s+ELECTORS\s*:?\s*(\d+)")

# Lines that are structure rather than candidate data, and so must be kept out of
# the column measurement -- they span several columns at once.
NON_DATA_RE = re.compile(r"^(Constituency\s|TURNOUT|GRAND TOTAL|Page \d+ of \d+)")


class LayoutError(Exception):
    """The PDF does not look like an ECI detailed-results report."""

# A cell that is too narrow for its content wraps onto the next line. Usually it
# breaks between words ("Pressure" / "Cooker" -> "Pressure Cooker"), but when a
# single token is wider than the column the PDF breaks it mid-word with no hyphen
# ("COCONU" / "T FARM" -> "COCONUT FARM"). Geometry cannot tell the two apart --
# both leave the first line filled to the edge -- so the mid-word breaks in this
# document are listed explicitly. Run with --audit to re-derive this list for a
# different PDF: it prints every wrap point that lands within one character of the
# column edge, which is the only place a mid-word break can occur.
FRAGMENT_JOINS = {
    ("SHARPNE", "R"),        # PENCIL SHARPNER
    ("Conditione", "r"),     # Air Conditioner
    ("Gramopho", "ne"),      # Gramophone
    ("Cauliflowe", "r"),     # Cauliflower
    ("CALCULA", "TOR"),      # CALCULATOR
    ("COCONU", "T"),         # COCONUT FARM
    ("Stethosco", "pe"),     # Stethoscope
    ("Refrigerat", "or"),    # Refrigerator
    ("PINEAPP", "LE"),       # PINEAPPLE
    ("Harmoniu", "m"),       # Harmonium
    ("SAMSA", "MPA"),        # party SAMSAMPA
    ("NINSHA", "D"),         # party NINSHAD
    ("CPI(ML)", "(L)"),      # party CPI(ML)(L)
    ("V.S.AWASTHY(VIDHASHAN", "KAR"),
    ("ADVOCATE-RAGHUNANDA", "N"),
}

# Columns whose text can wrap onto following lines.
WRAPPING_COLUMNS = ("name", "symbol", "party", "category")


def column_of(word, columns):
    """Return the column name whose band contains this word's centre."""
    centre = (word["x0"] + word["x1"]) / 2.0
    for name, x0, x1 in columns:
        if x0 <= centre < x1:
            return name
    return None


def fields_from_header(pages_lines):
    """Read the page header to learn which columns this report actually has.

    Reports differ: Punjab 2012 carries no SYMBOL column. Taking the field list
    from the header rather than assuming it keeps the two cases apart.
    """
    for lines in pages_lines:
        for line in lines:
            texts = {word["text"] for word in line["words"]}
            if "CANDIDATE" not in texts:
                continue
            fields = ["sl"]  # the serial-number column is unlabelled
            fields += [field for field, label in HEADER_ANCHORS if label in texts]
            if len(fields) < 6:
                continue
            return fields
    raise LayoutError(
        "No 'CANDIDATE NAME ... GENERAL POSTAL TOTAL' header row found. "
        "This does not look like an ECI detailed-results PDF, or the text layer "
        "is missing (a scanned PDF needs OCR first)."
    )


def find_column_bands(pages_lines, page_width, bin_size=0.25):
    """Locate the columns by finding the vertical whitespace gutters between them.

    Every word on every candidate row is projected onto the x-axis; the stripes
    that stay empty across the whole document are the gutters. This measures the
    real layout instead of assuming a page size or margin.
    """
    occupied = [False] * (int(page_width / bin_size) + 2)
    for lines in pages_lines:
        for line in lines:
            if is_page_furniture(line) or NON_DATA_RE.match(line["text"]):
                continue
            for word in line["words"]:
                for index in range(int(word["x0"] / bin_size), int(word["x1"] / bin_size) + 1):
                    if 0 <= index < len(occupied):
                        occupied[index] = True

    spans, start, empty_run = [], None, 0
    for index, filled in enumerate(occupied):
        if filled:
            if start is None:
                start = index
            empty_run = 0
        elif start is not None:
            empty_run += 1
            if empty_run * bin_size >= MIN_GUTTER:
                spans.append((start * bin_size, (index - empty_run) * bin_size))
                start = None
    if start is not None:
        spans.append((start * bin_size, page_width))
    return spans


def detect_layout(pages_lines, page_width):
    """Work out this PDF's column bands: [(field, x0, x1), ...]."""
    fields = fields_from_header(pages_lines)
    spans = find_column_bands(pages_lines, page_width)

    if len(spans) != len(fields):
        raise LayoutError(
            f"Found {len(spans)} columns of data but the header describes "
            f"{len(fields)} ({', '.join(fields)}). The page layout is not one this "
            f"parser recognises. Column edges measured: "
            f"{['%.0f-%.0f' % s for s in spans]}"
        )

    # Widen each column to meet its neighbours halfway, so a word sitting slightly
    # outside the measured extent still lands in the right column.
    columns = []
    for index, (field, (left, right)) in enumerate(zip(fields, spans)):
        low = 0.0 if index == 0 else (spans[index - 1][1] + left) / 2.0
        high = page_width if index == len(spans) - 1 else (right + spans[index + 1][0]) / 2.0
        columns.append((field, low, high))
    return columns


def read_pages(pdf):
    """Extract every page's words once, grouped into visual lines."""
    return [group_into_lines(page.extract_words(keep_blank_chars=False, use_text_flow=False))
            for page in pdf.pages]


def group_into_lines(words):
    """Group words into visual lines, ordered top-to-bottom, left-to-right."""
    lines = []
    for word in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(word["top"] - lines[-1]["top"]) <= ROW_TOLERANCE:
            lines[-1]["words"].append(word)
        else:
            lines.append({"top": word["top"], "words": [word]})
    for line in lines:
        line["words"].sort(key=lambda w: w["x0"])
        line["text"] = " ".join(w["text"] for w in line["words"])
    return lines


def is_page_furniture(line):
    """Repeated page header / footer / running-total lines carry no candidate data."""
    text = line["text"]
    if line["top"] < 115:  # title + column-header block at the top of every page
        return True
    if text.startswith(("TURNOUT TOTAL", "GRAND TOTAL")):
        return True
    if re.fullmatch(r"Page \d+ of \d+", text):  # footer, e.g. Punjab 2012
        return True
    first = line["words"][0]["text"]
    return first in HEADER_WORDS and "Constituency" not in text


def cells_of(line, columns):
    """Split one visual line into {column: text} using the x-bands."""
    cells = defaultdict(list)
    for word in line["words"]:
        col = column_of(word, columns)
        if col:
            cells[col].append(word["text"])
    return {col: " ".join(parts) for col, parts in cells.items()}


def join_wrapped(base, extra):
    """Attach a wrapped continuation chunk to the text already collected for a cell."""
    if not base:
        return extra
    tail = base.split()[-1]
    head = extra.split()[0]
    # A trailing hyphen ("Auto-" + "Rickshaw") or a known mid-word split means the
    # two fragments are one token; everything else is a plain word break.
    glued = (len(tail) > 1 and tail.endswith("-")) or (tail, head) in FRAGMENT_JOINS
    return base + extra if glued else base + " " + extra


def clean_number(value):
    """'1,234' / '1234' -> int; anything unusable -> None."""
    if value is None:
        return None
    if isinstance(value, int):
        return value
    digits = value.replace(",", "").strip()
    return int(digits) if digits.isdigit() else None


def extract_records(pdf_path, on_page=None):
    """Walk the PDF and return one dict per candidate row.

    `pdf_path` may be a path or any file-like object. `on_page`, if given, is
    called as on_page(page_no, total_pages) after each page so a caller can
    drive a progress bar.
    """
    records = []
    current = None          # constituency context
    record = None           # record currently open (may still gain wrapped text)
    warnings = []

    with pdfplumber.open(pdf_path) as pdf:
        pages_lines = read_pages(pdf)
        page_width = max(page.width for page in pdf.pages)
        columns = detect_layout(pages_lines, page_width)
        total_pages = len(pages_lines)

        for page_no, lines in enumerate(pages_lines, start=1):
            for line in lines:
                if is_page_furniture(line):
                    continue

                header = CONSTITUENCY_RE.search(line["text"])
                if header:
                    current = {
                        "no": int(header.group(1)),
                        "name": header.group(2).strip(),
                        "electors": clean_number(header.group(3)),
                    }
                    record = None
                    continue

                cells = cells_of(line, columns)

                # A row starts a new candidate when it has a serial number in the
                # left-most band. Everything else is wrapped text belonging to the
                # record above it (long names, multi-word symbols).
                if cells.get("sl", "").isdigit():
                    if current is None:
                        warnings.append(f"page {page_no}: candidate row before any constituency header")
                        continue
                    record = {
                        "constituency_no": current["no"],
                        "constituency": current["name"],
                        "electors": current["electors"],
                        "sl": int(cells["sl"]),
                        "name": cells.get("name", ""),
                        "sex": cells.get("sex", ""),
                        "age": cells.get("age", ""),
                        "category": cells.get("category", ""),
                        "party": cells.get("party", ""),
                        "symbol": cells.get("symbol", ""),
                        "general": cells.get("general", ""),
                        "postal": cells.get("postal", ""),
                        "total": cells.get("total", ""),
                        "percent": cells.get("percent", ""),
                        "page": page_no,
                    }
                    records.append(record)
                elif record is not None:
                    for col in WRAPPING_COLUMNS:
                        extra = cells.get(col)
                        if extra:
                            record[col] = join_wrapped(record[col], extra)
                else:
                    warnings.append(f"page {page_no}: unattached line -> {line['text'][:70]!r}")

            if on_page:
                on_page(page_no, total_pages)

    return records, warnings


def to_dataframe(records):
    """Shape the raw records into the fixed output columns, filling gaps with '-'."""
    rows = []
    for r in records:
        def text(value):
            value = (value or "").strip()
            return value if value else MISSING

        def number(value):
            n = clean_number(value)
            return n if n is not None else MISSING

        percent = (r["percent"] or "").strip()
        try:
            percent = float(percent)
        except ValueError:
            percent = MISSING

        rows.append({
            "Constituency": text(r["constituency"]),
            "Candidate Name": text(re.sub(r"\s+", " ", r["name"])),
            "Sex": text(r["sex"]),
            "Age": number(r["age"]),
            "Category": text(r["category"]),
            "Party": text(r["party"]),
            "Postal Votes": number(r["postal"]),
            "Total Valid Votes": number(r["total"]),
            "% Votes Polled": percent,
            "Symbol": text(re.sub(r"\s+", " ", r["symbol"])),
            "Total Electors": number(r["electors"]),
        })
    return pd.DataFrame(rows)


def validate(records):
    """Cross-check the parse against arithmetic the PDF itself asserts."""
    problems = []
    for r in records:
        general, postal, total = (clean_number(r[k]) for k in ("general", "postal", "total"))
        where = f"{r['constituency']} / {r['name']} (page {r['page']})"
        if None in (general, postal, total):
            problems.append(f"non-numeric vote cell: {where}")
        elif general + postal != total:
            problems.append(f"general+postal != total ({general}+{postal}!={total}): {where}")
        if not r["name"].strip():
            problems.append(f"empty candidate name: {where}")

    # Serial numbers inside each constituency must run 1..n without gaps.
    by_ac = defaultdict(list)
    for r in records:
        by_ac[r["constituency_no"]].append(r["sl"])
    for ac_no, serials in sorted(by_ac.items()):
        if serials != list(range(1, len(serials) + 1)):
            problems.append(f"serial numbers out of sequence in constituency {ac_no}: {serials}")

    return problems, len(by_ac)


def audit_wrap_points(pdf_path):
    """Print every wrap point that could be a mid-word break, for review.

    A cell only breaks mid-word when the first line is filled to the column edge,
    so anything more than one character short of the widest word seen in that
    column is certainly a plain word break and is not reported.
    """
    seen = defaultdict(set)
    widest = defaultdict(float)

    with pdfplumber.open(pdf_path) as pdf:
        pages_lines = read_pages(pdf)
        columns = detect_layout(pages_lines, max(page.width for page in pdf.pages))

        for lines in pages_lines:
            previous = {}
            for line in lines:
                if is_page_furniture(line):
                    continue
                by_col = defaultdict(list)
                for word in line["words"]:
                    col = column_of(word, columns)
                    if col:
                        by_col[col].append(word)
                        widest[col] = max(widest[col], word["x1"] - word["x0"])
                if by_col.get("sl"):
                    previous = dict(by_col)
                    continue
                for col, chunk in by_col.items():
                    if col in previous:
                        tail = previous[col][-1]
                        seen[col].add((round(tail["x1"] - tail["x0"], 1), tail["text"], chunk[0]["text"]))
                    previous[col] = previous.get(col, []) + chunk

    print("Wrap points within one character of the column edge -- check each one:")
    print("(listed as: column, first-line width, last fragment, next fragment, current handling)\n")
    for col in sorted(seen):
        limit = widest[col]
        margin = 1.5 * (limit / 10.0)  # ~10 characters fill a column at this font size
        for width, tail, head in sorted(seen[col], key=lambda z: -z[0]):
            if width < limit - margin:
                continue
            joined = "JOINED" if (tail, head) in FRAGMENT_JOINS or (len(tail) > 1 and tail.endswith("-")) else "spaced"
            print(f"  {col:8} {width:6.1f} (edge {limit:.1f})  {tail!r} + {head!r}  -> {joined}")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    pdf_path = args[0] if args else DEFAULT_PDF
    xlsx_path = args[1] if len(args) > 1 else DEFAULT_XLSX

    if "--audit" in sys.argv:
        audit_wrap_points(pdf_path)
        return

    records, warnings = extract_records(pdf_path)
    problems, n_constituencies = validate(records)

    df = to_dataframe(records)
    df.to_excel(xlsx_path, index=False, sheet_name="Results")
    df.to_csv(xlsx_path.rsplit(".", 1)[0] + ".csv", index=False)

    print(f"constituencies : {n_constituencies}")
    print(f"candidate rows : {len(df)}")
    print(f"written        : {xlsx_path}")
    for w in warnings:
        print("WARN :", w)
    for p in problems:
        print("CHECK:", p)
    if not warnings and not problems:
        print("validation     : all rows consistent (general+postal=total, serials 1..n)")


if __name__ == "__main__":
    main()
