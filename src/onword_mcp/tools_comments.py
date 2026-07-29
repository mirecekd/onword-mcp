"""Comment tools: list, read, navigate, reply, resolve and delete comments.

Word exposes comments as a flat Document.Comments collection, but modern Word
groups them into threads: a reply has .Ancestor pointing at the thread's first
comment, and the resolved flag (.Done) lives on that first comment only. These
tools therefore always report both the thread root and the resolved state of
the thread, so "open comment" means "comment whose thread is not resolved".

Gotchas worked around here:
- .Done raises on very old documents / pre-2013 comment format; treated as
  unresolved instead of failing the whole listing.
- .Scope is the highlighted anchor text, .Range is the comment body. Reporting
  the body as the anchor is a common mistake, so both are returned separately.
- Replies must be added through Replies.Add on the thread root; adding to a
  reply raises. resolve/reply therefore redirect to .Ancestor automatically.
"""

import json

from .word import WD_ACTIVE_END_PAGE_NUMBER, word_session, word_write_session


def _check_comment_index(doc, index: int) -> None:
    total = doc.Comments.Count
    if index < 1 or index > total:
        raise ValueError(
            f"Comment index {index} out of range (document has {total} comments)."
        )


def _root(doc, comment):
    """Return the thread root for a comment (itself when it is not a reply)."""
    try:
        anc = comment.Ancestor
    except Exception:
        return comment
    return anc if anc is not None else comment


def _root_index(doc, comment) -> int | None:
    root = _root(doc, comment)
    try:
        start = root.Reference.Start
    except Exception:
        return None
    for i in range(1, doc.Comments.Count + 1):
        try:
            if doc.Comments(i).Reference.Start == start:
                return i
        except Exception:
            continue
    return None


def _is_done(comment) -> bool:
    try:
        return bool(comment.Done)
    except Exception:
        return False


def _text(rng, limit: int | None = None) -> str:
    try:
        t = rng.Text or ""
    except Exception:
        return ""
    t = t.rstrip("\r\n\x07")
    if limit and len(t) > limit:
        t = t[:limit] + "..."
    return t


def _para_index(doc, comment) -> int | None:
    try:
        return doc.Range(0, comment.Scope.Start).Paragraphs.Count
    except Exception:
        return None


def _brief(doc, index: int, preview: int = 200) -> dict:
    c = doc.Comments(index)
    root = _root(doc, c)
    is_reply = False
    try:
        is_reply = c.Ancestor is not None
    except Exception:
        pass
    info = {
        "comment_index": index,
        "author": str(c.Author),
        "date": str(c.Date)[:19],
        "resolved": _is_done(root),
        "is_reply": is_reply,
        "paragraph_index": _para_index(doc, c),
        "anchor": _text(c.Scope, 120),
        "text": _text(c.Range, preview),
    }
    if is_reply:
        info["thread_root"] = _root_index(doc, c)
    try:
        info["page"] = int(c.Scope.Information(WD_ACTIVE_END_PAGE_NUMBER))
    except Exception:
        pass
    return info


def register(mcp):
    @mcp.tool()
    def list_comments(
        only_open: bool = True,
        start_index: int = 1,
        limit: int = 30,
        from_paragraph: int = 0,
        to_paragraph: int = 0,
    ) -> str:
        """List comments: index, author, date, resolved state, anchor text and
        comment text. By default only unresolved threads. Paginated - call again
        with a higher start_index if truncated. Use from_paragraph/to_paragraph
        to restrict the listing to one chapter or section."""
        with word_session() as (word, doc):
            total = doc.Comments.Count
            if total == 0:
                return "Document contains no comments."
            items = []
            i = max(1, start_index)
            while i <= total and len(items) < limit:
                info = _brief(doc, i)
                keep = not (only_open and info["resolved"])
                pi = info["paragraph_index"]
                if keep and from_paragraph and (pi is None or pi < from_paragraph):
                    keep = False
                if keep and to_paragraph and (pi is None or pi > to_paragraph):
                    keep = False
                if keep:
                    items.append(info)
                i += 1
            result = {
                "comments_total": total,
                "filter": {
                    "only_open": only_open,
                    "from_paragraph": from_paragraph or None,
                    "to_paragraph": to_paragraph or None,
                },
                "scanned": f"{start_index}..{i - 1}",
                "returned": len(items),
                "comments": items,
            }
            if i <= total:
                result["continue_with_start_index"] = i
            return json.dumps(result, ensure_ascii=False, indent=2)

    @mcp.tool()
    def get_comment(comment_index: int) -> str:
        """Read one comment in full, together with its whole thread (the root
        comment and all replies) and the text it is anchored to."""
        with word_session() as (word, doc):
            _check_comment_index(doc, comment_index)
            c = doc.Comments(comment_index)
            root = _root(doc, c)
            root_index = _root_index(doc, c) or comment_index
            thread = []
            try:
                replies = root.Replies
                for r in range(1, replies.Count + 1):
                    rep = replies(r)
                    thread.append(
                        {
                            "author": str(rep.Author),
                            "date": str(rep.Date)[:19],
                            "text": _text(rep.Range),
                        }
                    )
            except Exception:
                pass
            out = {
                "comment_index": comment_index,
                "thread_root_index": root_index,
                "resolved": _is_done(root),
                "paragraph_index": _para_index(doc, c),
                "anchor": _text(c.Scope),
                "root": {
                    "author": str(root.Author),
                    "date": str(root.Date)[:19],
                    "text": _text(root.Range),
                },
                "replies": thread,
            }
            return json.dumps(out, ensure_ascii=False, indent=2)

    @mcp.tool()
    def goto_comment(comment_index: int) -> str:
        """Select and scroll to the text a comment is anchored to, so the user
        can see it in Word. Use before acting on a comment."""
        with word_session() as (word, doc):
            _check_comment_index(doc, comment_index)
            c = doc.Comments(comment_index)
            scope = c.Scope
            scope.Select()
            try:
                word.ActiveWindow.ScrollIntoView(scope, True)
            except Exception:
                pass
            return json.dumps(_brief(doc, comment_index, preview=2000),
                              ensure_ascii=False, indent=2)

    @mcp.tool()
    def reply_to_comment(comment_index: int, text: str) -> str:
        """Add a reply to a comment thread, e.g. to record how a review remark
        was addressed. The reply is always attached to the thread root."""
        with word_write_session() as (word, doc):
            _check_comment_index(doc, comment_index)
            root = _root(doc, doc.Comments(comment_index))
            root.Replies.Add(root.Scope, text)
            return f"Reply added to comment thread of comment {comment_index}."

    @mcp.tool()
    def resolve_comment(comment_index: int, resolved: bool = True) -> str:
        """Mark a comment thread as resolved (or reopen it with resolved=False).
        Resolving always applies to the whole thread."""
        with word_write_session() as (word, doc):
            _check_comment_index(doc, comment_index)
            root = _root(doc, doc.Comments(comment_index))
            try:
                root.Done = resolved
            except Exception as exc:
                raise RuntimeError(
                    f"Cannot set the resolved state of comment {comment_index}: {exc}. "
                    "The document may use the legacy comment format."
                ) from exc
            state = "resolved" if resolved else "reopened"
            return f"Comment thread of comment {comment_index} {state}."

    @mcp.tool()
    def delete_comment(comment_index: int) -> str:
        """Delete a comment (a thread root deletes its replies too). Verify with
        get_comment first - this cannot be undone through MCP. Prefer
        resolve_comment; reviewers usually want the history kept."""
        with word_write_session() as (word, doc):
            _check_comment_index(doc, comment_index)
            doc.Comments(comment_index).Delete()
            return f"Comment {comment_index} deleted."
