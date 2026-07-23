"""Text-editing tools: replace, insert, append, delete, selection, track changes.

All edits are surgical (single Range operations), never rewriting the whole
document body - safe for large documents and Teams/SharePoint co-authoring.
"""

from .word import (
    WD_COLLAPSE_END,
    WD_COLLAPSE_START,
    check_paragraph_index,
    word_write_session,
)


def register(mcp):
    @mcp.tool()
    def replace_paragraph(index: int, new_text: str) -> str:
        """Replace the text of ONE paragraph at the given 1-based index.
        Keeps the paragraph mark so styles and numbering are preserved.
        Safe for Teams co-authoring (only local Range is touched)."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            p = doc.Paragraphs(index)
            p.Range.Text = new_text.rstrip("\r\n") + "\r"
            return f"Paragraph {index} replaced."

    @mcp.tool()
    def insert_after_paragraph(index: int, text: str) -> str:
        """Insert new paragraph(s) AFTER the paragraph at the given index.
        Use '\\n' in text to create multiple paragraphs."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            rng = doc.Paragraphs(index).Range
            rng.Collapse(WD_COLLAPSE_END)
            rng.Text = text.replace("\n", "\r").rstrip("\r") + "\r"
            return f"Text inserted after paragraph {index}."

    @mcp.tool()
    def insert_before_paragraph(index: int, text: str) -> str:
        """Insert new paragraph(s) BEFORE the paragraph at the given index."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            rng = doc.Paragraphs(index).Range
            rng.Collapse(WD_COLLAPSE_START)
            rng.Text = text.replace("\n", "\r").rstrip("\r") + "\r"
            return f"Text inserted before paragraph {index}."

    @mcp.tool()
    def append_to_document(text: str) -> str:
        """Append new paragraph(s) to the very end of the document without
        loading or rewriting the rest of it."""
        with word_write_session() as (word, doc):
            rng = doc.Content
            rng.Collapse(WD_COLLAPSE_END)
            rng.Text = "\r" + text.replace("\n", "\r").rstrip("\r")
            return "Text appended to end of document."

    @mcp.tool()
    def delete_paragraphs(start_index: int, count: int = 1) -> str:
        """Delete 'count' paragraphs starting at start_index (1-based)."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, start_index)
            end = min(start_index + count - 1, doc.Paragraphs.Count)
            # delete from the end so indexes stay valid
            for i in range(end, start_index - 1, -1):
                doc.Paragraphs(i).Range.Delete()
            return f"Deleted paragraphs {start_index}..{end}."

    @mcp.tool()
    def replace_text_in_paragraph(index: int, find: str, replace: str) -> str:
        """Find and replace a substring within ONE paragraph only.
        Preserves all other formatting in the paragraph."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            rng = doc.Paragraphs(index).Range
            f = rng.Find
            f.ClearFormatting()
            f.Text = find
            f.Replacement.ClearFormatting()
            f.Replacement.Text = replace
            f.Wrap = 0  # wdFindStop
            replaced = f.Execute(Replace=2)  # wdReplaceAll (within range)
            if replaced:
                return f"Replaced '{find}' in paragraph {index}."
            return f"'{find}' not found in paragraph {index}."

    @mcp.tool()
    def insert_at_selection(text: str) -> str:
        """Insert text at the user's current cursor position or replace the
        current selection in Word."""
        with word_write_session() as (word, doc):
            word.Selection.Text = text
            return "Text inserted at selection."

    @mcp.tool()
    def set_track_changes(enabled: bool) -> str:
        """Enable or disable Word's Track Changes for the active document.
        When enabled, all edits made by these tools appear as revisions the
        user can accept or reject."""
        with word_write_session() as (word, doc):
            doc.TrackRevisions = enabled
            state = "enabled" if enabled else "disabled"
            return f"Track changes {state}."

    @mcp.tool()
    def save_document() -> str:
        """Save the active document (for cloud/Teams documents this triggers
        the normal co-authoring sync)."""
        with word_write_session() as (word, doc):
            doc.Save()
            return f"Document '{doc.Name}' saved."