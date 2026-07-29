"""Table tools: locate tables, read/write cells, insert and delete rows.

Word tables cannot be edited through paragraph indexes alone - each cell is its
own paragraph, so inserting a paragraph "between rows" corrupts the layout.
These tools operate on the Tables collection instead.

Two gotchas these tools work around:
- A cell's Range ends with the end-of-cell marker (\\r\\x07). Writing to the
  full Range breaks the table, so writes shrink the Range by one character.
- A single cell may hold several paragraphs (e.g. text plus a stray empty
  paragraph). Paragraph indexes therefore do not map 1:1 to cells - use
  find_table_at_paragraph to translate an index into row/column coordinates.
  Rewriting the cell with set_table_cell collapses it back to one paragraph.
"""


import json

from .word import (
    WD_ACTIVE_END_PAGE_NUMBER,
    check_paragraph_index,
    word_session,
    word_write_session,
)


def _cell_text(cell) -> str:
    # Word terminates cell text with \r\x07 (end-of-cell marker).
    return cell.Range.Text.rstrip("\r\x07").rstrip("\r\n")


def _check_table_index(doc, index: int) -> None:
    total = doc.Tables.Count
    if index < 1 or index > total:
        raise ValueError(
            f"Table index {index} out of range (document has {total} tables)."
        )


def _check_row_index(table, row: int, allow_append: bool = False) -> None:
    total = table.Rows.Count
    limit = total + 1 if allow_append else total
    if row < 1 or row > limit:
        raise ValueError(
            f"Row index {row} out of range (table has {total} rows)."
        )


def register(mcp):
    @mcp.tool()
    def list_tables() -> str:
        """List all tables in the document: index, page, size and the text of
        the first (header) row. Use this first to find the table you need."""
        with word_session() as (word, doc):
            tables = []
            for i in range(1, doc.Tables.Count + 1):
                t = doc.Tables(i)
                header = []
                try:
                    for c in range(1, t.Columns.Count + 1):
                        header.append(_cell_text(t.Cell(1, c)))
                except Exception:
                    # merged cells in the header row
                    header = ["<merged cells>"]
                tables.append(
                    {
                        "table_index": i,
                        "page": int(t.Range.Information(WD_ACTIVE_END_PAGE_NUMBER)),
                        "rows": int(t.Rows.Count),
                        "columns": int(t.Columns.Count),
                        "header": header,
                    }
                )
            if not tables:
                return "Document contains no tables."
            return json.dumps(tables, ensure_ascii=False, indent=2)

    @mcp.tool()
    def find_table_at_paragraph(index: int) -> str:
        """Find which table (if any) contains the paragraph at the given index.
        Use after find_text/get_document_outline located a cell, to translate a
        paragraph index into table/row/column coordinates."""
        with word_session() as (word, doc):
            check_paragraph_index(doc, index)
            rng = doc.Paragraphs(index).Range
            if not rng.Information(12):  # wdWithInTable
                return f"Paragraph {index} is not inside a table."
            cell = rng.Cells(1)
            table = rng.Tables(1)
            table_index = None
            for i in range(1, doc.Tables.Count + 1):
                if doc.Tables(i).Range.Start == table.Range.Start:
                    table_index = i
                    break
            return json.dumps(
                {
                    "paragraph_index": index,
                    "table_index": table_index,
                    "row": int(cell.RowIndex),
                    "column": int(cell.ColumnIndex),
                    "rows_total": int(table.Rows.Count),
                    "columns_total": int(table.Columns.Count),
                    "cell_text": _cell_text(cell),
                },
                ensure_ascii=False,
                indent=2,
            )

    @mcp.tool()
    def read_table(table_index: int, start_row: int = 1, count: int = 20) -> str:
        """Read rows of a table as a list of cell-text lists. Paginated - call
        again with a higher start_row for large tables."""
        with word_session() as (word, doc):
            _check_table_index(doc, table_index)
            t = doc.Tables(table_index)
            total = int(t.Rows.Count)
            _check_row_index(t, start_row)
            end = min(start_row + count - 1, total)
            rows = []
            for r in range(start_row, end + 1):
                cells = []
                try:
                    for c in range(1, t.Columns.Count + 1):
                        cells.append(_cell_text(t.Cell(r, c)))
                except Exception:
                    # row with merged/missing cells - walk its actual cells
                    cells = [_cell_text(c) for c in t.Rows(r).Cells]
                rows.append({"row": r, "cells": cells})
            result = {
                "table_index": table_index,
                "rows_total": total,
                "returned": f"{start_row}..{end}",
                "data": rows,
            }
            return json.dumps(result, ensure_ascii=False, indent=2)

    @mcp.tool()
    def goto_table_row(table_index: int, row: int) -> str:
        """Select and scroll to a table row in Word so the user can see it.
        Use before editing to show exactly which row will be changed."""
        with word_session() as (word, doc):
            _check_table_index(doc, table_index)
            t = doc.Tables(table_index)
            _check_row_index(t, row)
            rng = t.Rows(row).Range
            rng.Select()
            try:
                word.ActiveWindow.ScrollIntoView(rng, True)
            except Exception:
                pass
            cells = [_cell_text(c) for c in t.Rows(row).Cells]
            return json.dumps(
                {
                    "table_index": table_index,
                    "row": row,
                    "page": int(rng.Information(WD_ACTIVE_END_PAGE_NUMBER)),
                    "cells": cells,
                },
                ensure_ascii=False,
                indent=2,
            )

    @mcp.tool()
    def set_table_cell(table_index: int, row: int, column: int, text: str) -> str:
        """Set the text of ONE table cell. Keeps the cell's formatting and the
        end-of-cell marker intact."""
        with word_write_session() as (word, doc):
            _check_table_index(doc, table_index)
            t = doc.Tables(table_index)
            _check_row_index(t, row)
            if column < 1 or column > t.Columns.Count:
                raise ValueError(
                    f"Column {column} out of range (table has {t.Columns.Count} columns)."
                )
            cell = t.Cell(row, column)
            rng = cell.Range
            # shrink the range by the end-of-cell marker before overwriting
            rng.End = rng.End - 1
            rng.Text = text
            return f"Table {table_index} cell ({row},{column}) set."

    @mcp.tool()
    def insert_table_row(
        table_index: int, after_row: int, cells: list[str] | None = None
    ) -> str:
        """Insert a new row into a table AFTER the given row, optionally filling
        its cells. Use after_row=0 to insert before the first row. The new row
        inherits formatting from the row it is inserted after."""
        with word_write_session() as (word, doc):
            _check_table_index(doc, table_index)
            t = doc.Tables(table_index)
            total = int(t.Rows.Count)
            if after_row < 0 or after_row > total:
                raise ValueError(
                    f"after_row {after_row} out of range (table has {total} rows; "
                    "use 0 to insert before the first row)."
                )
            if after_row == 0:
                new_row = t.Rows.Add(t.Rows(1))  # BeforeRow -> insert at top
                new_index = 1
            elif after_row == total:
                new_row = t.Rows.Add()  # append at the end
                new_index = total + 1
            else:
                new_row = t.Rows.Add(t.Rows(after_row + 1))
                new_index = after_row + 1
            filled = 0
            if cells:
                for i, value in enumerate(cells, start=1):
                    if i > t.Columns.Count:
                        break
                    rng = new_row.Cells(i).Range
                    rng.End = rng.End - 1
                    rng.Text = value
                    filled += 1
            return (
                f"Row inserted into table {table_index} at position {new_index}"
                f" ({filled} cell(s) filled)."
            )

    @mcp.tool()
    def delete_table_row(table_index: int, row: int) -> str:
        """Delete one row from a table. Verify the row content with read_table
        or goto_table_row first - this cannot be undone through MCP."""
        with word_write_session() as (word, doc):
            _check_table_index(doc, table_index)
            t = doc.Tables(table_index)
            _check_row_index(t, row)
            t.Rows(row).Delete()
            return f"Row {row} deleted from table {table_index}."
