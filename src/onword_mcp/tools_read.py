"""Read-only tools: document info, outline, paragraph reading, search, pages."""

import json

from .word import (
    WD_NUMBER_OF_PAGES,
    check_paragraph_index,
    paragraph_brief,
    paragraph_page,
    word_session,
)


def register(mcp):
    @mcp.tool()
    def get_document_info() -> str:
        """Get basic info about the document currently open in Word:
        name, path, page count, paragraph count, track changes state."""
        with word_session() as (word, doc):
            info = {
                "name": doc.Name,
                "path": doc.FullName,
                "paragraphs": doc.Paragraphs.Count,
                "pages": int(doc.Content.Information(WD_NUMBER_OF_PAGES)),
                "track_changes": bool(doc.TrackRevisions),
                "words": doc.Words.Count,
            }
            return json.dumps(info, ensure_ascii=False, indent=2)

    @mcp.tool()
    def get_document_outline(start_index: int = 1, limit: int = 200) -> str:
        """Get document structure: paragraph indexes, styles, list levels and
        text previews. Paginated for large documents - call again with a higher
        start_index if truncated. Use this first to orient in the document."""
        with word_session() as (word, doc):
            total = doc.Paragraphs.Count
            lines = []
            i = max(1, start_index)
            emitted = 0
            while i <= total and emitted < limit:
                info = paragraph_brief(doc, i)
                if info["preview"]:
                    extra = ""
                    if "list" in info:
                        extra += f" [list:{info['list']} lvl:{info['list_level']}]"
                    if "alignment" in info:
                        extra += f" [{info['alignment']}]"
                    lines.append(
                        f"[{info['index']}] ({info['style']}){extra}: {info['preview']}"
                    )
                    emitted += 1
                i += 1
            header = f"Paragraphs {start_index}..{i - 1} of {total} total"
            if i <= total:
                header += f" (more available, continue with start_index={i})"
            return header + "\n" + "\n".join(lines)

    @mcp.tool()
    def read_paragraphs(start_index: int, count: int = 5) -> str:
        """Read the full text of a block of paragraphs starting at start_index
        (1-based). Use after get_document_outline to read a specific section."""
        with word_session() as (word, doc):
            check_paragraph_index(doc, start_index)
            end = min(start_index + count - 1, doc.Paragraphs.Count)
            parts = []
            for i in range(start_index, end + 1):
                text = doc.Paragraphs(i).Range.Text.rstrip("\r\n\x07")
                parts.append(f"=== [{i}] ===\n{text}")
            return "\n".join(parts)

    @mcp.tool()
    def find_text(query: str, max_results: int = 20) -> str:
        """Find text in the document. Returns paragraph indexes, page numbers
        and previews of matches. Uses Word's native Find (fast on large docs)."""
        with word_session() as (word, doc):
            rng = doc.Content
            find = rng.Find
            find.ClearFormatting()
            find.Text = query
            find.Forward = True
            find.Wrap = 0  # wdFindStop
            results = []
            while len(results) < max_results and find.Execute():
                para_range = rng.Paragraphs(1).Range
                # paragraph index = paragraphs from doc start to range start
                idx = doc.Range(0, rng.Start).Paragraphs.Count
                page = int(rng.Information(3))  # wdActiveEndPageNumber
                preview = para_range.Text.rstrip("\r\n\x07")[:100]
                results.append(f"[{idx}] (page {page}): {preview}")
                rng.Collapse(0)  # collapse to end, continue searching
            if not results:
                return f"No matches for '{query}'."
            return f"{len(results)} match(es) for '{query}':\n" + "\n".join(results)

    @mcp.tool()
    def get_page_paragraphs(page_number: int) -> str:
        """List paragraphs located on a given page (index, style, list info,
        preview). Use when the user refers to content by page number,
        e.g. 'the bullet on page 10'."""
        with word_session() as (word, doc):
            total = doc.Paragraphs.Count
            lines = []
            # binary search for the first paragraph on the page
            lo, hi, first = 1, total, None
            while lo <= hi:
                mid = (lo + hi) // 2
                pg = paragraph_page(doc.Paragraphs(mid))
                if pg < page_number:
                    lo = mid + 1
                else:
                    if pg == page_number:
                        first = mid
                    hi = mid - 1
            if first is None:
                pages = int(doc.Content.Information(WD_NUMBER_OF_PAGES))
                return f"No paragraphs found on page {page_number} (document has {pages} pages)."
            i = first
            while i <= total:
                if paragraph_page(doc.Paragraphs(i)) != page_number:
                    break
                info = paragraph_brief(doc, i)
                extra = ""
                if "list" in info:
                    extra += f" [list:{info['list']} lvl:{info['list_level']}]"
                lines.append(f"[{info['index']}] ({info['style']}){extra}: {info['preview']}")
                i += 1
            return f"Page {page_number} paragraphs:\n" + "\n".join(lines)

    @mcp.tool()
    def get_selection() -> str:
        """Get the text and location of the user's current selection/cursor
        in Word."""
        with word_session() as (word, doc):
            sel = word.Selection
            text = sel.Text.rstrip("\r\n\x07")
            idx = doc.Range(0, sel.Range.Start).Paragraphs.Count if sel.Range.Start > 0 else 1
            page = int(sel.Range.Information(3))
            return json.dumps(
                {"paragraph_index": idx, "page": page, "text": text},
                ensure_ascii=False,
                indent=2,
            )