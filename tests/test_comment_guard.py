"""Self-check for the comment thread guard.

The guard exists because of a real incident: a reply meant for one reviewer's
thread was written onto another's, twice in a row, because comment indexes had
shifted in between. The write succeeded both times and nothing complained.

Run without a test framework (and without Word):

    python tests/test_comment_guard.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from onword_mcp.tools_comments import _guard  # noqa: E402


class _Range:
    def __init__(self, text):
        self.Text = text


class _Comment:
    """Minimal stand-in for a Word Comment COM object."""

    def __init__(self, author, text):
        self.Author = author
        self.Range = _Range(text)
        self.Ancestor = None


class _Doc:
    def __init__(self, comment):
        self._comment = comment

    def Comments(self, index):
        return self._comment


HAVRANEK = "Tohle je na prodiskutovani na schuzce. 5 znaku je min. u Accountu."
doc = _Doc(_Comment("Havranek Zdenek", HAVRANEK))


def _refuses(**kwargs):
    """True when the guard blocks the write."""
    try:
        _guard(doc, 37, kwargs.get("expect_author"),
               kwargs.get("expect_contains"), "reply to")
    except ValueError:
        return True
    return False


# the intended thread is accepted, on full name and on a fragment
assert not _refuses(expect_author="Havranek Zdenek")
assert not _refuses(expect_author="Havranek")
# author matching ignores case
assert not _refuses(expect_author="havranek zdenek")
# text matching works and is also case-insensitive
assert not _refuses(expect_contains="5 znaku")
assert not _refuses(expect_contains="5 ZNAKU")
# no expectation given: nothing to check, write proceeds
assert not _refuses()

# the actual incident: index now points at a different reviewer's thread
assert _refuses(expect_author="Chloupek Petr")
# right author, wrong thread of theirs - caught by the text expectation
assert _refuses(expect_contains="Remediace tagu")
# both must hold, not just one
assert _refuses(expect_author="Havranek", expect_contains="Remediace tagu")

# the refusal has to explain itself, otherwise the caller cannot recover
try:
    _guard(doc, 37, "Chloupek Petr", None, "reply to")
except ValueError as exc:
    message = str(exc)
    assert "Havranek Zdenek" in message, "must name the actual author"
    assert "Chloupek Petr" in message, "must name the expected author"
    assert "5 znaku" in message, "must quote the actual thread text"
    assert "list_comments" in message, "must say how to recover"

print("comment guard: all checks passed")
