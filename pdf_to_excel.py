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

# Column bands (x0, x1) measured from the PDF layout. A word belongs to the
# column whose band contains the word's horizontal centre.
COLUMNS = [
    ("sl",       35.0, 50.0),
    ("name",     50.0, 186.0),
    ("sex",     186.0, 202.0),
    ("age",     202.0, 229.0),
    ("category", 229.0, 284.0),
    ("party",   284.0, 320.0),
    ("symbol",  320.0, 380.0),
    ("general", 380.0, 440.0),
    ("postal",  440.0, 484.0),
    ("total",   484.0, 525.0),
    ("percent", 525.0, 600.0),
]

# Two words are on the same visual line if their tops differ by less than this.
ROW_TOLERANCE = 3.0

HEADER_WORDS = {
    "Election", "DETAILED", "RESULTS", "VALID", "VOTES", "POLLED",
    "CANDIDATE", "NAME", "SEX", "AGE", "CATEGORY", "PARTY", "SYMBOL",
    "GENERAL", "POSTAL", "TOTAL", "%",
}

CONSTITUENCY_RE = re.compile(r"Constituency\s+(\d+)\.\s*(.+?)\s+TOTAL\s+ELECTORS\s*:?\s*(\d+)")

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


def column_of(word):
    """Return the column name whose band contains this word's centre."""
    centre = (word["x0"] + word["x1"]) / 2.0
    for name, x0, x1 in COLUMNS:
        if x0 <= centre < x1:
            return name
    return None


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
    first = line["words"][0]["text"]
    return first in HEADER_WORDS and "Constituency" not in text


def cells_of(line):
    """Split one visual line into {column: text} using the x-bands."""
    cells = defaultdict(list)
    for word in line["words"]:
        col = column_of(word)
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
        total_pages = len(pdf.pages)
        for page_no, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
            for line in group_into_lines(words):
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

                cells = cells_of(line)

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
        for page in pdf.pages:
            previous = {}
            for line in group_into_lines(page.extract_words()):
                if is_page_furniture(line):
                    continue
                by_col = defaultdict(list)
                for word in line["words"]:
                    col = column_of(word)
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
