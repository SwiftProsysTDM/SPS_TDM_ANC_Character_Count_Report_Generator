"""
report_core.py
----------------
Core logic to read Content Capture Studio (.db3) SQLite batch files and
build the "DocType" billing report exactly like 62884.IDX.001_Project DocType
Billing_DataExport.xlsx

Output workbook has 3 sheets:
    1. Manual Key   -> Tab 1 (empty placeholder, user fills manually / extends later)
    2. DocType      -> the report grid (Project ID, Doc Type Name, Batch Name,
                        Image Count, Header Count, Record Count, Character Count)
    3. RawData      -> (hidden) supporting sheet holding every *_orig field value,
                        grouped in blocks per Batch+DocType, so that the
                        Character Count column on the DocType sheet can use a
                        real Excel formula:
                            =SUMPRODUCT(LEN(TRIM(RawData!<range>)))
                        instead of a hard-coded number.

Verified against the sample 62884.IDX.001 batches (content/folder/Probate rows
matched the provided screenshot exactly): Image Count, Header Count, Record
Count = simple COUNT()s; Character Count = SUM of LEN(TRIM(x)) over every
*_orig column of the Record table, for the records that belong to that
Batch + DocType.
"""

import os
import sqlite3
import difflib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

# ---------------------------------------------------------------------------
# Fixed display order for known doc types (Folder pages are physically first
# in a batch, then Content, then case-specific types). Anything not listed
# here is appended alphabetically at the end.
# ---------------------------------------------------------------------------
DOCTYPE_DISPLAY_ORDER = [
    "Folder", "Content", "Probate", "Criminal", "License", "Mixed", "Name Change",
]


def _doctype_sort_key(doctype: str):
    if doctype in DOCTYPE_DISPLAY_ORDER:
        return (0, DOCTYPE_DISPLAY_ORDER.index(doctype))
    return (1, doctype)


def _get_indexed_doctypes(cur):
    """Return {DocTypeName: [TableName, ...]} for every DocType flagged Indexed=1.

    A single DocType can map to MORE THAN ONE table in DocTypeDataTable
    (e.g. 'Card' -> 'Header' AND 'Record'). Character Count must include
    the *_orig fields from every mapped table, not just one.
    """
    cur.execute("SELECT Type FROM DocType WHERE Indexed=1")
    indexed = [r[0] for r in cur.fetchall()]

    cur.execute("SELECT DocType, TableName FROM DocTypeDataTable")
    mapping = {}
    for dt, table in cur.fetchall():
        mapping.setdefault(dt, []).append(table)

    # default to 'Record' table if a doctype has no explicit mapping
    return {dt: mapping.get(dt, ["Record"]) for dt in indexed}


def _orig_columns(cur, table_name):
    cur.execute(f"PRAGMA table_info({table_name})")
    cols = [r[1] for r in cur.fetchall()]
    return [c for c in cols if c.endswith("_orig")]


def _set_safe(cell, value):
    """Write a value to a cell WITHOUT letting openpyxl/Excel misinterpret
    a string that happens to start with '=' (or '+'/'-'/'@') as a formula —
    this can corrupt the file if keyed text or a note began with those
    characters. Forces the cell's type back to plain string afterwards."""
    cell.value = value
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        cell.data_type = "s"
    return cell


def _table_raw_rows(cur, table_name, doctype):
    """Fetch *_orig column values for every row of `table_name` belonging to
    `doctype`. Handles both the Header-level table and the Record-level
    table (Record needs an extra join through Header)."""
    orig_cols = _orig_columns(cur, table_name)
    if not orig_cols:
        return orig_cols, []

    col_list = ",".join(orig_cols)
    if table_name == "Header":
        query = (
            f"SELECT {col_list} FROM Header h "
            f"JOIN Image i ON h.ImageID=i.ImageID WHERE i.DocType=?"
        )
    else:
        query = (
            f"SELECT {col_list} FROM {table_name} r "
            f"JOIN Header h ON r.HeaderID=h.HeaderID "
            f"JOIN Image i ON h.ImageID=i.ImageID WHERE i.DocType=?"
        )
    cur.execute(query, (doctype,))
    return orig_cols, cur.fetchall()


def _process_db3(path, project_id):
    """Return list of dicts, one per (DocType present) in this batch file."""
    conn = sqlite3.connect(path)
    cur = conn.cursor()

    cur.execute("SELECT DISTINCT BatchName FROM Image")
    rows = cur.fetchall()
    batch_name = rows[0][0] if rows else os.path.splitext(os.path.basename(path))[0]

    doctype_tables = _get_indexed_doctypes(cur)

    results = []
    for doctype, table_names in doctype_tables.items():
        cur.execute("SELECT COUNT(*) FROM Image WHERE DocType=?", (doctype,))
        image_count = cur.fetchone()[0]
        if image_count == 0:
            continue  # this doctype not present in this batch, skip the row

        cur.execute(
            "SELECT COUNT(*) FROM Header h JOIN Image i ON h.ImageID=i.ImageID "
            "WHERE i.DocType=?",
            (doctype,),
        )
        header_count = cur.fetchone()[0]

        cur.execute(
            f"SELECT COUNT(*) FROM Record r "
            f"JOIN Header h ON r.HeaderID=h.HeaderID "
            f"JOIN Image i ON h.ImageID=i.ImageID WHERE i.DocType=?",
            (doctype,),
        )
        record_count = cur.fetchone()[0]

        # one raw-data table_block per mapped table (Header and/or Record, ...)
        table_blocks = []
        for table_name in table_names:
            orig_cols, raw_rows = _table_raw_rows(cur, table_name, doctype)
            table_blocks.append(
                {"table_name": table_name, "orig_cols": orig_cols, "raw_rows": raw_rows}
            )

        results.append(
            {
                "project_id": project_id,
                "doctype": doctype,
                "batch_name": batch_name,
                "image_count": image_count,
                "header_count": header_count,
                "record_count": record_count,
                "table_blocks": table_blocks,
            }
        )

    conn.close()

    results.sort(key=lambda r: _doctype_sort_key(r["doctype"]))
    return results


def _build_workbook(all_blocks, project_id, output_path):
    """Shared workbook builder used by both generate_report() and
    generate_partial_link_report(). `all_blocks` is a list of per
    Batch+DocType dicts, each already carrying its final `table_blocks`
    (i.e. any field exclusions must already be applied by the caller)."""
    wb = Workbook()

    # ---------------- RawData (support sheet for formulas) -----------------
    ws_raw = wb.active
    ws_raw.title = "RawData"
    ws_raw.sheet_state = "hidden"
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")

    raw_cursor_row = 1
    block_ranges = []  # parallel to all_blocks: list of (start_row, end_row, start_col, end_col) per table_block

    for block in all_blocks:
        ranges_for_row = []
        for tb in block["table_blocks"]:
            orig_cols = tb["orig_cols"]
            raw_rows = tb["raw_rows"]

            # block label row
            ws_raw.cell(
                row=raw_cursor_row, column=1,
                value=f'{block["batch_name"]} | {block["doctype"]} | {tb["table_name"]}',
            ).font = Font(bold=True)
            raw_cursor_row += 1

            # column header row (starting at column A)
            start_col = 1
            for i, col_name in enumerate(orig_cols):
                c = ws_raw.cell(row=raw_cursor_row, column=start_col + i, value=col_name)
                c.font = header_font
                c.fill = header_fill
            raw_cursor_row += 1

            data_start_row = raw_cursor_row
            for row_vals in raw_rows:
                for i, v in enumerate(row_vals):
                    _set_safe(ws_raw.cell(row=raw_cursor_row, column=start_col + i), v)
                raw_cursor_row += 1
            data_end_row = raw_cursor_row - 1

            if not raw_rows:
                # no records -> empty range, formula will just be 0
                data_start_row = raw_cursor_row
                data_end_row = raw_cursor_row
                raw_cursor_row += 1

            end_col = start_col + max(len(orig_cols), 1) - 1
            ranges_for_row.append((data_start_row, data_end_row, start_col, end_col))

            raw_cursor_row += 1  # blank spacer row between table blocks

        block_ranges.append(ranges_for_row)

    # ---------------- Sheet 1: DocType (the report) -------------------------
    ws = wb.create_sheet("DocType")
    headers = [
        "Project ID:ProjectID", "Doc Type Name", "Batch Name:BatchName",
        "Image Count", "Header Count", "Record Count", "Character Count",
    ]
    red_bold = Font(bold=True, color="C00000")
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = red_bold
    ws.freeze_panes = "A2"

    for r_idx, (block, ranges_for_row) in enumerate(zip(all_blocks, block_ranges), start=2):
        ws.cell(row=r_idx, column=1, value=block["project_id"])
        ws.cell(row=r_idx, column=2, value=block["doctype"])
        ws.cell(row=r_idx, column=3, value=block["batch_name"])
        ws.cell(row=r_idx, column=4, value=block["image_count"])
        ws.cell(row=r_idx, column=5, value=block["header_count"])
        ws.cell(row=r_idx, column=6, value=block["record_count"])

        # Character Count = sum of SUMPRODUCT(LEN(TRIM(range))) across every
        # mapped table's raw-data block (a DocType can map to more than one
        # table, e.g. Card -> Header + Record).
        parts = []
        for (start_row, end_row, start_col, end_col) in ranges_for_row:
            start_letter = get_column_letter(start_col)
            end_letter = get_column_letter(end_col)
            parts.append(
                f"SUMPRODUCT(LEN(TRIM(RawData!${start_letter}${start_row}:"
                f"${end_letter}${end_row})))"
            )
        formula = "=" + "+".join(parts)
        ws.cell(row=r_idx, column=7, value=formula)

    widths = [20, 16, 42, 12, 13, 13, 15]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---------------- Sheet 2: Summary (per-DocType totals, whole project) --
    ws_sum = wb.create_sheet("Summary")
    sum_headers = [
        "Project ID:ProjectID", "Doc Type Name", "Image Count",
        "Header Count", "Record Count", "Character Count",
    ]
    for col_idx, h in enumerate(sum_headers, start=1):
        cell = ws_sum.cell(row=1, column=col_idx, value=h)
        cell.font = red_bold
    ws_sum.freeze_panes = "A2"

    detail_last_row = len(all_blocks) + 1  # header row + one row per block
    unique_doctypes = sorted({b["doctype"] for b in all_blocks})

    for r_idx, doctype in enumerate(unique_doctypes, start=2):
        ws_sum.cell(row=r_idx, column=1, value=project_id)
        ws_sum.cell(row=r_idx, column=2, value=doctype)
        for col_idx, detail_col in zip((3, 4, 5, 6), ("D", "E", "F", "G")):
            formula = (
                f"=SUMIF(DocType!$B$2:$B${detail_last_row},$B{r_idx},"
                f"DocType!${detail_col}$2:${detail_col}${detail_last_row})"
            )
            ws_sum.cell(row=r_idx, column=col_idx, value=formula)

    sum_widths = [20, 16, 12, 13, 13, 15]
    for i, w in enumerate(sum_widths, start=1):
        ws_sum.column_dimensions[get_column_letter(i)].width = w

    # order tabs: DocType (detail), Summary, RawData(hidden)
    wb._sheets = [ws, ws_sum, ws_raw]

    wb.save(output_path)
    return output_path, len(all_blocks)


def generate_report(folder_path: str, output_path: str, project_id: str = None, progress_callback=None):
    """
    folder_path : directory that directly contains the .db3 batch files
    output_path : full path (.xlsx) to write
    project_id  : if None, taken automatically from folder_path's own name
    """
    folder_path = os.path.abspath(folder_path)
    if project_id is None:
        project_id = os.path.basename(folder_path.rstrip("/\\"))

    db3_files = sorted(
        f for f in os.listdir(folder_path) if f.lower().endswith(".db3")
    )
    if not db3_files:
        raise FileNotFoundError(f"No .db3 files found in: {folder_path}")

    all_blocks = []
    total = len(db3_files)
    for idx, fname in enumerate(db3_files, start=1):
        if progress_callback:
            progress_callback(idx - 1, total, f"Processing batch {idx}/{total}: {fname}")
        full = os.path.join(folder_path, fname)
        all_blocks.extend(_process_db3(full, project_id))
        if progress_callback:
            progress_callback(idx, total, f"Completed batch {idx}/{total}: {fname}")

    return _build_workbook(all_blocks, project_id, output_path)


# ---------------------------------------------------------------------------
# "Keying Partially Linked" (Tab 2) — Old + New two-pass billing
# ---------------------------------------------------------------------------
#
# Some projects are keyed in TWO passes:
#   Pass 1 ("Partial Link" / Old zip)  — operator only enters the person's
#       Given Name + Surname, just enough to LINK the record to an existing
#       tree profile. This gets billed on its own, separately.
#   Pass 2 (Old batch promoted -> New zip) — a second operator fully keys
#       every remaining field (dates, ages, places, relationships, etc.)
#       for that same image/record.
#
# If Pass 2's Character Count simply summed *every* *_orig field, the
# Given Name / Surname captured in Pass 1 would be billed a SECOND time.
# So Pass 2's Character Count must EXCLUDE whichever *_orig field(s) were
# used for the Pass-1 linking - determined dynamically by inspecting what
# was actually populated in the "linking" DocType's records in the Old zip
# (verified against project 62105.IDX.008: the only fields ever populated
# under 'Partial Link' were SelfGivenName_orig and SelfSurname_orig, and
# excluding exactly those two fields from the New-zip totals reproduced the
# ground-truth SWF billing sheet exactly, for all 12 sample batches).


def _detect_linking_doctype(cur):
    """Find the DocType used purely for name-linking in the Old zip. We
    look for an Indexed=1 DocType whose name contains both 'partial' and
    'link' (case-insensitive) — e.g. 'Partial Link'."""
    cur.execute("SELECT Type FROM DocType WHERE Indexed=1")
    for (t,) in cur.fetchall():
        low = t.lower()
        if "partial" in low and "link" in low:
            return t
    return None


def _detect_linking_fields(old_folder, linking_doctype):
    """Scan every batch in old_folder and return the set of *_orig column
    names that were EVER populated for `linking_doctype` records. These are
    the fields already billed in the Pass-1 (linking) invoice."""
    fields = set()
    db3_files = sorted(f for f in os.listdir(old_folder) if f.lower().endswith(".db3"))
    for fname in db3_files:
        conn = sqlite3.connect(os.path.join(old_folder, fname))
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM Image WHERE DocType=?", (linking_doctype,))
        if cur.fetchone()[0] == 0:
            conn.close()
            continue

        cur.execute(
            "SELECT TableName FROM DocTypeDataTable WHERE DocType=?", (linking_doctype,)
        )
        tables = [r[0] for r in cur.fetchall()] or ["Record"]

        for table_name in tables:
            orig_cols = _orig_columns(cur, table_name)
            if not orig_cols:
                continue
            col_list = ",".join(orig_cols)
            if table_name == "Header":
                q = (
                    f"SELECT {col_list} FROM Header h "
                    f"JOIN Image i ON h.ImageID=i.ImageID WHERE i.DocType=?"
                )
            else:
                q = (
                    f"SELECT {col_list} FROM {table_name} r "
                    f"JOIN Header h ON r.HeaderID=h.HeaderID "
                    f"JOIN Image i ON h.ImageID=i.ImageID WHERE i.DocType=?"
                )
            cur.execute(q, (linking_doctype,))
            for row in cur.fetchall():
                for col, v in zip(orig_cols, row):
                    if v not in (None, ""):
                        fields.add(col)
        conn.close()
    return fields


def _filter_table_blocks(table_blocks, exclude_fields):
    """Return a copy of table_blocks with `exclude_fields` columns removed
    from every orig_cols/raw_rows pair."""
    if not exclude_fields:
        return table_blocks
    filtered = []
    for tb in table_blocks:
        keep_idx = [i for i, c in enumerate(tb["orig_cols"]) if c not in exclude_fields]
        new_cols = [tb["orig_cols"][i] for i in keep_idx]
        new_rows = [tuple(row[i] for i in keep_idx) for row in tb["raw_rows"]]
        filtered.append({
            "table_name": tb["table_name"],
            "orig_cols": new_cols,
            "raw_rows": new_rows,
        })
    return filtered


def generate_partial_link_report(
    old_folder: str,
    new_folder: str,
    output_path: str,
    project_id: str = None,
    progress_callback=None,
):
    """
    old_folder : folder with the Pass-1 (.db3) batches — only linking fields
                 (e.g. Given Name + Surname) were keyed here.
    new_folder : folder with the SAME batches after Pass-2 full keying.
    output_path: full path (.xlsx) to write.

    Character Count for every fully-keyed DocType row = every *_orig field
    EXCEPT whichever field(s) were used for name-linking in old_folder (auto
    -detected), so the linking characters aren't billed twice.
    """
    old_folder = os.path.abspath(old_folder)
    new_folder = os.path.abspath(new_folder)
    if project_id is None:
        project_id = os.path.basename(new_folder.rstrip("/\\"))

    old_db3_files = sorted(f for f in os.listdir(old_folder) if f.lower().endswith(".db3"))
    if not old_db3_files:
        raise FileNotFoundError(f"No .db3 files found in: {old_folder}")

    # 1) detect the linking doctype + its fields from the Old zip
    sample_conn = sqlite3.connect(os.path.join(old_folder, old_db3_files[0]))
    linking_doctype = _detect_linking_doctype(sample_conn.cursor())
    sample_conn.close()

    linking_fields = set()
    if linking_doctype:
        linking_fields = _detect_linking_fields(old_folder, linking_doctype)

    # 2) process the New (fully-keyed) batches, excluding the linking fields
    db3_files = sorted(f for f in os.listdir(new_folder) if f.lower().endswith(".db3"))
    if not db3_files:
        raise FileNotFoundError(f"No .db3 files found in: {new_folder}")

    all_blocks = []
    total = len(db3_files)
    for idx, fname in enumerate(db3_files, start=1):
        if progress_callback:
            progress_callback(idx - 1, total, f"Processing batch {idx}/{total}: {fname}")
        full = os.path.join(new_folder, fname)
        for block in _process_db3(full, project_id):
            if linking_doctype and block["doctype"] == linking_doctype:
                # still awaiting Pass-2 promotion — not billed in this
                # invoice; it stays on its own Pass-1 linking invoice.
                continue
            block["table_blocks"] = _filter_table_blocks(block["table_blocks"], linking_fields)
            all_blocks.append(block)
        if progress_callback:
            progress_callback(idx, total, f"Completed batch {idx}/{total}: {fname}")

    return _build_workbook(all_blocks, project_id, output_path)


if __name__ == "__main__":
    import sys
    folder = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "DocType_Billing_Report.xlsx"
    path, n = generate_report(folder, out)
    print(f"Wrote {n} rows to {path}")


# ---------------------------------------------------------------------------
# "XLCompare-style" per-batch audit report — visually explains exactly which
# characters were counted towards Character Count.
# ---------------------------------------------------------------------------
#
# For every batch, produces ONE .xlsx with 2 sheets:
#   "Comparison"      -> every Header/Record field: Old value (blue font),
#                        New value with per-character rich-text coloring:
#                            black  = unchanged / excluded from billing
#                                     (i.e. was already present in Old, or is
#                                     a linking field such as Given/Surname)
#                            red+bold = NEW characters that ARE counted in
#                                     Character Count
#   "Character Count" -> Old Total / New Total / Billable (Difference)
#                        Character Count for the batch — same number the
#                        DocType/Summary sheets would show for that batch.
#
# Record alignment: RecordIDs can be RENUMBERED between Old and New when a
# record is inserted mid-list (SQLite re-sequences them) — so records are
# matched by IDENTITY (the linking-field values, e.g. Given+Surname), using
# a sequence alignment (difflib), not by raw RecordID equality.


def _diff_runs(old_val, new_val):
    """Return a list of (text, is_added) runs reconstructing new_val,
    marking which substrings are NOT present in old_val (i.e. new/changed
    characters, using a character-level diff)."""
    old_val = old_val or ""
    new_val = new_val or ""
    sm = difflib.SequenceMatcher(None, old_val, new_val, autojunk=False)
    runs = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal" and j2 > j1:
            runs.append((new_val[j1:j2], False))
        elif tag in ("insert", "replace") and j2 > j1:
            runs.append((new_val[j1:j2], True))
        # 'delete' removes old characters — contributes nothing to new_val
    return runs


def _added_length(old_val, new_val):
    return sum(len(t) for t, is_added in _diff_runs(old_val, new_val) if is_added)


def _rich_new_value(old_val, new_val):
    """Build a CellRichText for new_val: black for unchanged/excluded parts,
    red+bold for the parts that are NEW (billable) characters."""
    runs = _diff_runs(old_val, new_val)
    blocks = [
        TextBlock(InlineFont(color="C00000", b=True) if added else InlineFont(color="000000"), text)
        for text, added in runs if text
    ]
    if not blocks:
        return ""
    return CellRichText(*blocks)


def _align_records(old_by_id, new_by_id, old_order, new_order, key_idx):
    """Match New RecordIDs to Old RecordIDs by IDENTITY (the columns at
    key_idx, e.g. Given+Surname), not by raw ID equality — RecordIDs can be
    renumbered when a record is inserted mid-list. Returns {new_id: old_id
    or None}."""
    def key(row):
        return tuple((row[i] or "").strip().lower() for i in key_idx)

    old_keys = [key(old_by_id[i]) for i in old_order]
    new_keys = [key(new_by_id[i]) for i in new_order]

    sm = difflib.SequenceMatcher(None, old_keys, new_keys, autojunk=False)
    mapping = {nid: None for nid in new_order}
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ("equal", "replace"):
            for oi, ni in zip(range(i1, i2), range(j1, j2)):
                mapping[new_order[ni]] = old_order[oi]
        # 'insert' -> stays None (brand-new record, no Old counterpart)
        # 'delete' -> old-only record, dropped, nothing on the New side
    return mapping


def _batch_comparison_data(old_path, new_path, project_id, linking_doctype, linking_fields):
    """Build the row-level comparison data for one batch (Old vs New),
    already excluding linking fields from the billable count exactly like
    generate_partial_link_report() does. Returns (batch_name, rows, totals)
    where rows is a list of dicts with old/new values + billable flag, and
    totals = {"old_total", "new_total", "billable_total"}."""
    old_conn = sqlite3.connect(old_path)
    new_conn = sqlite3.connect(new_path)
    old_cur, new_cur = old_conn.cursor(), new_conn.cursor()

    new_cur.execute("SELECT DISTINCT BatchName FROM Image")
    br = new_cur.fetchall()
    batch_name = br[0][0] if br else os.path.splitext(os.path.basename(new_path))[0]

    doctype_tables = _get_indexed_doctypes(new_cur)

    rows = []
    old_total = new_total = billable_total = 0

    for doctype, table_names in doctype_tables.items():
        if linking_doctype and doctype == linking_doctype:
            continue
        new_cur.execute("SELECT COUNT(*) FROM Image WHERE DocType=?", (doctype,))
        if new_cur.fetchone()[0] == 0:
            continue

        for table_name in table_names:
            orig_cols = _orig_columns(new_cur, table_name)
            if not orig_cols:
                continue

            if table_name == "Header":
                new_cur.execute(
                    f"SELECT h.HeaderID,{','.join(orig_cols)} FROM Header h "
                    f"JOIN Image i ON h.ImageID=i.ImageID WHERE i.DocType=?", (doctype,)
                )
                new_hdrs = new_cur.fetchall()
                for hrow in new_hdrs:
                    hid, *new_vals = hrow
                    old_cur.execute(
                        f"SELECT {','.join(orig_cols)} FROM Header WHERE HeaderID=?", (hid,)
                    )
                    orow = old_cur.fetchone()
                    old_vals = orow if orow else tuple(None for _ in orig_cols)
                    for field, ov, nv in zip(orig_cols, old_vals, new_vals):
                        if not ov and not nv:
                            continue
                        billable = field not in linking_fields
                        added = _added_length(ov, nv) if billable else 0
                        old_total += len((ov or "").strip())
                        new_total += len((nv or "").strip())
                        billable_total += added
                        rows.append({
                            "doctype": doctype, "table": "Header", "id": hid,
                            "field": field, "old": ov, "new": nv,
                            "billable": billable, "added": added,
                        })
            else:
                new_cur.execute(
                    "SELECT HeaderID FROM Header h JOIN Image i ON h.ImageID=i.ImageID "
                    "WHERE i.DocType=? ORDER BY HeaderID", (doctype,)
                )
                header_ids = [r[0] for r in new_cur.fetchall()]

                key_idx = [i for i, c in enumerate(orig_cols) if c in linking_fields] or [0]

                for hid in header_ids:
                    new_cur.execute(
                        f"SELECT RecordID,{','.join(orig_cols)} FROM {table_name} "
                        f"WHERE HeaderID=? ORDER BY RecordID", (hid,)
                    )
                    new_recs = new_cur.fetchall()
                    old_cur.execute(
                        f"SELECT RecordID,{','.join(orig_cols)} FROM Record "
                        f"WHERE HeaderID=? ORDER BY RecordID", (hid,)
                    )
                    old_recs = old_cur.fetchall()

                    new_order = [r[0] for r in new_recs]
                    new_by_id = {r[0]: r[1:] for r in new_recs}
                    old_order = [r[0] for r in old_recs]
                    old_by_id = {r[0]: r[1:] for r in old_recs}

                    mapping = _align_records(old_by_id, new_by_id, old_order, new_order, key_idx) if old_order else {n: None for n in new_order}

                    for nid in new_order:
                        new_vals = new_by_id[nid]
                        oid = mapping.get(nid)
                        old_vals = old_by_id[oid] if oid is not None else tuple(None for _ in orig_cols)
                        for field, ov, nv in zip(orig_cols, old_vals, new_vals):
                            if not ov and not nv:
                                continue
                            billable = field not in linking_fields
                            added = _added_length(ov, nv) if billable else 0
                            old_total += len((ov or "").strip())
                            new_total += len((nv or "").strip())
                            billable_total += added
                            rows.append({
                                "doctype": doctype, "table": table_name, "id": nid,
                                "field": field, "old": ov, "new": nv,
                                "billable": billable, "added": added,
                            })

    old_conn.close()
    new_conn.close()
    return batch_name, rows, {
        "old_total": old_total, "new_total": new_total, "billable_total": billable_total,
    }


def _build_comparison_workbook(batch_name, project_id, rows, totals, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison"

    hdr_font = Font(bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor="4472C4")
    headers = ["Doc Type", "Table", "ID", "Field", "Old Value", "New Value", "Added (Billable) Chars"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
    ws.freeze_panes = "A2"

    for r_idx, row in enumerate(rows, start=2):
        ws.cell(row=r_idx, column=1, value=row["doctype"])
        ws.cell(row=r_idx, column=2, value=row["table"])
        ws.cell(row=r_idx, column=3, value=row["id"])
        ws.cell(row=r_idx, column=4, value=row["field"])

        old_cell = ws.cell(row=r_idx, column=5)
        _set_safe(old_cell, row["old"] or "")
        old_cell.font = Font(color="0070C0")  # blue = Old value

        new_cell = ws.cell(row=r_idx, column=6)
        rich = _rich_new_value(row["old"], row["new"])
        if rich:
            new_cell.value = rich  # CellRichText isn't string-formula-detected
        else:
            _set_safe(new_cell, row["new"] or "")

        added_cell = ws.cell(row=r_idx, column=7, value=row["added"] if row["billable"] else "excl.")
        if not row["billable"]:
            added_cell.font = Font(color="808080", italic=True)

    widths = [12, 10, 8, 24, 26, 34, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws2 = wb.create_sheet("Character Count")
    labels = [
        ("Project ID", project_id),
        ("Batch Name", batch_name),
        ("Old Total Characters (all fields)", totals["old_total"]),
        ("New Total Characters (all fields)", totals["new_total"]),
        ("Billable (Difference) Character Count", totals["billable_total"]),
    ]
    for i, (label, val) in enumerate(labels, start=1):
        c1 = ws2.cell(row=i, column=1, value=label)
        c1.font = Font(bold=True)
        ws2.cell(row=i, column=2, value=val)
    ws2.column_dimensions["A"].width = 38
    ws2.column_dimensions["B"].width = 30
    ws2.cell(row=6, column=1, value="Legend (Comparison sheet):").font = Font(bold=True, italic=True)
    ws2.cell(row=7, column=1, value="Old Value column").font = Font(color="0070C0")
    ws2.cell(row=7, column=2, value="value captured in Pass-1 (Old)")
    ws2.cell(row=8, column=1, value="New Value: black text").font = Font(color="000000")
    ws2.cell(row=8, column=2, value="unchanged from Old, or excluded (linking field) — NOT billed")
    ws2.cell(row=9, column=1, value="New Value: red bold text").font = Font(color="C00000", bold=True)
    ws2.cell(row=9, column=2, value="new/changed characters — THIS is what's billed")

    wb._sheets = [ws, ws2]
    wb.save(out_path)


def generate_diff_reports(old_folder, new_folder, output_folder, project_id=None, progress_callback=None):
    """
    Generates ONE audit .xlsx PER BATCH into output_folder, each showing Old
    vs New values with color-coded highlighting of exactly which characters
    were counted towards billing (like an XLCompare-style diff tool).
    """
    old_folder = os.path.abspath(old_folder)
    new_folder = os.path.abspath(new_folder)
    os.makedirs(output_folder, exist_ok=True)
    if project_id is None:
        project_id = os.path.basename(new_folder.rstrip("/\\"))

    new_files = sorted(f for f in os.listdir(new_folder) if f.lower().endswith(".db3"))
    if not new_files:
        raise FileNotFoundError(f"No .db3 files found in: {new_folder}")

    # detect the linking doctype + fields ONCE (project-wide scan), instead
    # of per-batch, since it's the same for every batch in the project.
    sample_conn = sqlite3.connect(os.path.join(old_folder, new_files[0]) if os.path.exists(os.path.join(old_folder, new_files[0])) else os.path.join(old_folder, os.listdir(old_folder)[0]))
    linking_doctype = _detect_linking_doctype(sample_conn.cursor())
    sample_conn.close()
    linking_fields = _detect_linking_fields(old_folder, linking_doctype) if linking_doctype else set()

    written = []
    total = len(new_files)
    for idx, fname in enumerate(new_files, start=1):
        if progress_callback:
            progress_callback(idx - 1, total, f"Comparing batch {idx}/{total}: {fname}")

        new_path = os.path.join(new_folder, fname)
        old_path = os.path.join(old_folder, fname)
        if not os.path.exists(old_path):
            if progress_callback:
                progress_callback(idx, total, f"Skipped (no Old match): {fname}")
            continue

        batch_name, rows, totals = _batch_comparison_data(
            old_path, new_path, project_id, linking_doctype, linking_fields
        )
        if not rows:
            if progress_callback:
                progress_callback(idx, total, f"Skipped (no billable doctype): {fname}")
            continue

        out_path = os.path.join(output_folder, f"{batch_name}_Comparison.xlsx")
        _build_comparison_workbook(batch_name, project_id, rows, totals, out_path)
        written.append(out_path)

        if progress_callback:
            progress_callback(idx, total, f"Completed batch {idx}/{total}: {fname}")

    return output_folder, written