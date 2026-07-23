"""COM helpers for talking to the live, running Microsoft Word instance.

Design notes:
- Never cache COM objects between tool calls. FastMCP may run tools on
  different threads; COM interface pointers must not cross apartment
  boundaries. Re-acquiring via GetActiveObject on every call is cheap
  (it is a running-object-table lookup) and avoids marshalling entirely.
- Every tool call does CoInitialize()/CoUninitialize() for its thread.
- Writes are wrapped so ScreenUpdating is disabled during the change and
  always restored, keeping large documents fast and the UI responsive.
"""

import sys
from contextlib import contextmanager

IS_WINDOWS = sys.platform == "win32"

if IS_WINDOWS:
    import pythoncom
    import win32com.client

# Word enum constants (subset used by the tools)
WD_COLLAPSE_END = 0
WD_COLLAPSE_START = 1
WD_ACTIVE_END_PAGE_NUMBER = 3  # wdActiveEndPageNumber
WD_NUMBER_OF_PAGES = 4  # wdNumberOfPagesInDocument
WD_STORY = 6  # wdStory

WD_ALIGN = {
    "left": 0,
    "center": 1,
    "right": 2,
    "justify": 3,
}
WD_ALIGN_REVERSE = {v: k for k, v in WD_ALIGN.items()}

# ListFormat.ListType values
WD_LIST_TYPE = {
    0: "none",
    1: "list_num",
    2: "bullet",
    3: "simple_numbering",
    4: "outline_numbering",
    5: "mixed_numbering",
    6: "picture_bullet",
}


class WordNotAvailableError(RuntimeError):
    """Raised when the running Word instance cannot be reached."""


@contextmanager
def word_session():
    """Yield (word_app, active_document) connected to the running Word.

    Initializes COM for the current thread and cleans up afterwards.
    """
    if not IS_WINDOWS:
        raise WordNotAvailableError(
            "onword-mcp requires Microsoft Windows with Microsoft Word "
            "running (pywin32 COM automation)."
        )
    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception as exc:
            raise WordNotAvailableError(
                f"Cannot connect to a running Microsoft Word instance: {exc}. "
                "Make sure Word is running with a document open."
            ) from exc
        if word.Documents.Count == 0:
            raise WordNotAvailableError(
                "Word is running but no document is open."
            )
        yield word, word.ActiveDocument
    finally:
        pythoncom.CoUninitialize()


@contextmanager
def word_write_session():
    """Like word_session, but disables screen updating during the edit."""
    with word_session() as (word, doc):
        try:
            word.ScreenUpdating = False
        except Exception:
            pass
        try:
            yield word, doc
        finally:
            try:
                word.ScreenUpdating = True
            except Exception:
                pass


def check_paragraph_index(doc, index: int) -> None:
    """Validate a 1-based paragraph index, raise ValueError if invalid."""
    total = doc.Paragraphs.Count
    if index < 1 or index > total:
        raise ValueError(
            f"Paragraph index {index} out of range (document has {total} paragraphs)."
        )


def paragraph_page(paragraph) -> int:
    """Page number where the paragraph starts."""
    return int(paragraph.Range.Information(WD_ACTIVE_END_PAGE_NUMBER))


def paragraph_brief(doc, index: int, preview_len: int = 80) -> dict:
    """Compact info about one paragraph (index, style, list info, preview)."""
    p = doc.Paragraphs(index)
    rng = p.Range
    text = rng.Text.rstrip("\r\n\x07")
    info = {
        "index": index,
        "style": str(p.Style),
        "preview": text[:preview_len] + ("..." if len(text) > preview_len else ""),
    }
    try:
        lf = rng.ListFormat
        if lf.ListType != 0:
            info["list"] = WD_LIST_TYPE.get(int(lf.ListType), str(lf.ListType))
            info["list_level"] = int(lf.ListLevelNumber)
    except Exception:
        pass
    try:
        align = int(p.Format.Alignment)
        if align != 0:
            info["alignment"] = WD_ALIGN_REVERSE.get(align, str(align))
    except Exception:
        pass
    return info


def cm_to_points(cm: float) -> float:
    return cm * 28.35


def points_to_cm(points: float) -> float:
    return round(points / 28.35, 2)