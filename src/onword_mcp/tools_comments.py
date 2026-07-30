"""Comment tools: list, read, navigate, add, reply, resolve and delete comments.

Word exposes comments as a flat Document.Comments collection, but modern Word
groups them into threads: a reply has .Ancestor pointing at the thread's first
comment, and the resolved flag (.Done) lives on that first comment only. These
tools therefore always report both the thread root and the resolved state of
the thread, so "open comment" means "comment whose thread is not resolved".

THE BIG TRAP: comment indexes are positional, not stable IDs. They shift after
every comment operation (add/reply/delete) AND after text edits that move or
delete the anchor paragraphs. An index read one operation ago may already point
at a different thread - and writing a reply onto a stranger's thread is easy to
miss, because the write itself succeeds.

Two defences are built in:
- every write returns the thread root's author and text, so a mistake is
  visible in the tool result instead of surfacing days later;
- the optional expect_author / expect_contains guards abort the write when the
  target thread is not the one the caller believes it is. They are mandatory
  for delete_comment, which cannot be undone.

Other gotchas worked around here:
- .Done raises on the pre-2013 comment format; treated as unresolved rather
  than failing the whole listing.
- .Scope is the highlighted anchor text, .Range is the comment body. Both are
  reported separately - confusing them is a common mistake.
- Replies must be added through Replies.Add on the thread root; adding to a
  reply raises. reply/resolve therefore redirect to .Ancestor automatically.
- Comments.Add over a whole paragraph Range would swallow the paragraph mark
  and anchor the comment across the break, so add_comment trims it.
- Writes can silently roll back (observed on this codebase's main use case),
  so writes save by default and re-read the live state afterwards.
"""

import json

from .word import (
    WD_ACTIVE_END_PAGE_NUMBER,
    check_paragraph_index,
    word_session,
    word_write_session,
)


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


SHIFT_HINT = (
    "Comment indexes shift after every comment operation and after text edits. "
    "Re-run list_comments (filtered with from_paragraph/to_paragraph) and "
    "confirm with get_comment immediately before acting."
)


def _guard(doc, comment_index, expect_author, expect_contains, operation):
    """Abort when the target thread is not the one the caller expects.

    Returns the thread root so callers do not have to look it up twice.
    """
    comment = doc.Comments(comment_index)
    root = _root(doc, comment)
    author = str(root.Author)
    body = _text(root.Range)
    if expect_author and expect_author.casefold() not in author.casefold():
        raise ValueError(
            f"Refusing to {operation} comment {comment_index}: thread root is by "
            f"{author!r}, but expect_author was {expect_author!r}. "
            f"Root text: {body[:200]!r}. {SHIFT_HINT}"
        )
    if expect_contains and expect_contains.casefold() not in body.casefold():
        raise ValueError(
            f"Refusing to {operation} comment {comment_index}: thread root text "
            f"does not contain {expect_contains!r}. Root is by {author!r} and "
            f"reads {body[:200]!r}. {SHIFT_HINT}"
        )
    return root


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


def _thread_state(doc, root, action: str, saved: bool) -> str:
    """Re-read the live thread after a write and report what is really there."""
    replies = []
    try:
        rep = root.Replies
        for r in range(1, rep.Count + 1):
            replies.append(
                {
                    "author": str(rep(r).Author),
                    "date": str(rep(r).Date)[:19],
                    "text": _text(rep(r).Range, 200),
                }
            )
    except Exception:
        pass
    return json.dumps(
        {
            "action": action,
            "saved": saved,
            "verified_after_write": {
                "thread_root_author": str(root.Author),
                "thread_root_text": _text(root.Range, 200),
                "anchor": _text(root.Scope, 120),
                "resolved": _is_done(root),
                "reply_count": len(replies),
                "replies": replies,
            },
            "note": "State above was re-read from Word after the write. "
            "If it is not the thread you meant, the index had shifted - "
            "use expect_author/expect_contains to make such a mistake fail.",
        },
        ensure_ascii=False,
        indent=2,
    )


def register(mcp):
    @mcp.tool()
    def add_comment(
        paragraph_index: int,
        text: str,
        anchor: str | None = None,
        save: bool = True,
    ) -> str:
        """Add a NEW comment thread anchored to a paragraph, or to a substring
        inside it when 'anchor' is given. Use reply_to_comment to append to an
        existing thread instead.

        Note: Word deletes a comment together with the paragraph it is anchored
        to, so a remark on a block you are about to move must be re-created here
        afterwards."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, paragraph_index)
            rng = doc.Paragraphs(paragraph_index).Range
            if anchor:
                body = rng.Text
                pos = body.find(anchor)
                if pos == -1:
                    raise ValueError(
                        f"Anchor text {anchor!r} not found in paragraph "
                        f"{paragraph_index}."
                    )
                rng = doc.Range(rng.Start + pos, rng.Start + pos + len(anchor))
            else:
                # drop the paragraph mark so the comment does not span the break
                if rng.End > rng.Start:
                    rng = doc.Range(rng.Start, rng.End - 1)
            comment = doc.Comments.Add(rng, text)
            if save:
                doc.Save()
            return _thread_state(
                doc, comment, f"comment added to paragraph {paragraph_index}", save
            )

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
        to restrict the listing to one chapter or section.

        Indexes are positional and shift after every comment operation and after
        text edits, so call this immediately before acting on a comment."""
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
        comment and all replies) and the text it is anchored to. Use this to
        confirm an index really points at the thread you mean before writing."""
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
    def reply_to_comment(
        comment_index: int,
        text: str,
        expect_author: str | None = None,
        expect_contains: str | None = None,
        save: bool = True,
    ) -> str:
        """Add a reply to a comment thread, e.g. to record how a review remark
        was addressed. The reply is always attached to the thread root.

        Pass expect_author (and/or expect_contains) to have the write REFUSED
        when the index no longer points at that thread - strongly recommended,
        because indexes shift after every comment operation and after text
        edits. The result echoes the thread root so a mismatch is obvious."""
        with word_write_session() as (word, doc):
            _check_comment_index(doc, comment_index)
            root = _guard(doc, comment_index, expect_author, expect_contains, "reply to")
            root.Replies.Add(root.Scope, text)
            if save:
                doc.Save()
            return _thread_state(
                doc, root, f"reply added to thread of comment {comment_index}", save
            )

    @mcp.tool()
    def resolve_comment(
        comment_index: int,
        resolved: bool = True,
        expect_author: str | None = None,
        expect_contains: str | None = None,
        save: bool = True,
    ) -> str:
        """Mark a comment thread as resolved (or reopen it with resolved=False).
        Resolving always applies to the whole thread.

        Pass expect_author / expect_contains to guard against a shifted index.
        The resolved flag is re-read from Word after the write, because this
        flag has been observed to silently roll back."""
        with word_write_session() as (word, doc):
            _check_comment_index(doc, comment_index)
            action = "resolve" if resolved else "reopen"
            root = _guard(doc, comment_index, expect_author, expect_contains, action)
            try:
                root.Done = resolved
            except Exception as exc:
                raise RuntimeError(
                    f"Cannot set the resolved state of comment {comment_index}: {exc}. "
                    "The document may use the legacy comment format."
                ) from exc
            if save:
                doc.Save()
            if _is_done(root) != resolved:
                raise RuntimeError(
                    f"Word did not keep the resolved state of comment "
                    f"{comment_index} (wanted {resolved}, still "
                    f"{_is_done(root)}). Retry and verify again."
                )
            state = "resolved" if resolved else "reopened"
            return _thread_state(
                doc, root, f"thread of comment {comment_index} {state}", save
            )

    @mcp.tool()
    def delete_comment(
        comment_index: int,
        expect_author: str,
        expect_contains: str | None = None,
        save: bool = True,
    ) -> str:
        """Delete a comment (a thread root deletes its replies too).

        expect_author is REQUIRED: deleting through MCP cannot be undone, and a
        shifted index would otherwise destroy a reviewer's remark. The deletion
        is refused unless the comment really is by that author. Prefer
        resolve_comment - reviewers usually want the history kept."""
        with word_write_session() as (word, doc):
            _check_comment_index(doc, comment_index)
            comment = doc.Comments(comment_index)
            author = str(comment.Author)
            body = _text(comment.Range)
            # guard against the comment itself, not the thread root: deleting a
            # reply must check that reply, not the thread it hangs under
            if expect_author.casefold() not in author.casefold():
                raise ValueError(
                    f"Refusing to delete comment {comment_index}: it is by "
                    f"{author!r}, but expect_author was {expect_author!r}. "
                    f"Comment text: {body[:200]!r}. {SHIFT_HINT}"
                )
            if expect_contains and expect_contains.casefold() not in body.casefold():
                raise ValueError(
                    f"Refusing to delete comment {comment_index}: its text does "
                    f"not contain {expect_contains!r}. It is by {author!r} and "
                    f"reads {body[:200]!r}. {SHIFT_HINT}"
                )
            comment.Delete()
            if save:
                doc.Save()
            return json.dumps(
                {
                    "action": f"comment {comment_index} deleted",
                    "saved": save,
                    "deleted": {"author": author, "text": body[:200]},
                    "comments_total_now": doc.Comments.Count,
                    "note": "Remaining comment indexes have shifted. Re-run "
                    "list_comments before the next comment operation.",
                },
                ensure_ascii=False,
                indent=2,
            )
