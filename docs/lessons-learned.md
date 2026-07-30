# Lessons learned: driving Word through COM

Everything here was hit in practice while an LLM edited a 88-page, 67-table
review document (with 150+ comments) live on SharePoint. Each item cost real
debugging time or a real mistake in a customer deliverable, so the list is
worth reading before automating Word.

Where a lesson could be turned into a guard rail, it was - see
[What the server now enforces](#what-the-server-now-enforces).

## 1. Indexes are positions, not identities

Paragraph, table and comment indexes are ordinals in a collection. Insert,
move or delete anything and every later index shifts.

This is obvious for paragraphs but easy to forget for **comments**, where two
separate causes stack up:

- every comment operation (`add`, `reply`, `delete`) inserts or removes an
  entry in `Document.Comments`;
- **editing the text shifts them too**, because comments are ordered by the
  position of their anchor.

The failure mode is nasty: a reply written to a stale index lands on somebody
else's thread, and **the write succeeds**. Nothing complains. In the review
document this happened twice in a row - the second time while fixing the
first, because the corrective `delete_comment` had shifted the indexes again.

**Rule:** re-locate the target immediately before every write. Never reuse an
index from earlier in the conversation, from a log, or from notes.

## 2. Writes can silently roll back

A tool reports success, verification right afterwards passes, and in the next
session the change is simply gone. Observed on:

- plain text (a typo fix reverted twice, days apart),
- paragraph styles,
- comment replies,
- the comment `Done` (resolved) flag.

The likely culprit is SharePoint/Teams co-authoring sync losing a change that
was never flushed, but the practical answer is the same regardless of cause:

**Rule:** `save_document` after each individual write, not once at the end of a
batch. Then read the state back. After adopting this, nothing was lost again.
And re-verify in the *next* session too - that is when rollbacks surfaced.

## 3. Word deletes a comment together with its anchor paragraph

Moving a block means insert-elsewhere + delete-original. Any comment anchored
to the original paragraphs is **destroyed, unrecoverably**. A reviewer's remark
vanishes with no warning.

**Rule:** before moving a block, list the comments on it. Re-create them
afterwards with `add_comment`.

## 4. `Range.Find.Execute(Replace=...)` lies

Over COM it reports success without changing anything. This server never uses
it: `replace_text_in_paragraph` locates the substring in Python and assigns to
a character sub-Range instead.

## 5. Find reports an index one too low

`find_text` and `goto_text` frequently return an index one lower than the real
paragraph - for prose, for headings, and for paragraphs inside table cells.

**Rule:** read both `index` and `index + 1` before editing.

Also: **Find is case-insensitive.** `find_text("Quicksight")` matches a
correctly spelled `QuickSight`, and `find_text("TODO")` matches "todo" inside
ordinary words. Differences in capitalisation must be confirmed by reading the
text, not by counting matches.

## 6. All Caps headings must be searched in capitals

When a heading's *style* applies All Caps, the underlying text is mixed case
but the search matches what is displayed. Searching `Landing zone` failed on
such a heading; `LANDING ZONE` worked. `read_paragraphs` and `goto_text` find
the text either way, which makes the failure confusing.

## 7. `replace_paragraph` resets the style to Normal

Replacing a heading's whole text drops its style, which kills its numbering and
its entry in the table of contents.

**Rule:** for headings prefer `replace_text_in_paragraph`. If you must use
`replace_paragraph`, re-apply `set_paragraph_style` immediately and verify
`list_string`.

## 8. Inserted paragraphs inherit the target's formatting - all of them

`insert_before_paragraph` / `insert_after_paragraph` give **every** inserted
paragraph the style of the reference paragraph. Insert a five-paragraph block
at a heading and you get five numbered headings.

Two more quirks:

- inserting after the **last item of a list** produces `Normal`, not a bullet;
  fixing it needs `set_paragraph_style("List Paragraph")` **and**
  `set_list_level(n)` - the style alone gives level 1;
- inserting **into the middle** of a list inherits correctly.

When inserting at several places in one list, work **backwards**, otherwise
earlier insertions invalidate the later targets.

## 9. Word renumbers headings; it does not fix cross-references

Move a section and the heading numbers update automatically. Every reference
like "see section 3.3.1.5" written as literal text stays wrong.

**Rule:** after moving a section, hunt down the references (`find_text` on
"section 3.3." and similar). And remember the table of contents needs the user
to press **F9** - no API call does it.

## 10. A table cell is its own paragraph, and may hold several

Paragraph-level inserts inside a table corrupt the layout, so tables have their
own tools. Because a cell can contain more than one paragraph, indexes do not
map 1:1 to cells - use `find_table_at_paragraph` to translate. Rewriting a cell
with `set_table_cell` flattens it back to one paragraph, which is also the
cleanest way to remove stray empty paragraphs inside cells.

Related traps:

- what looks like "an empty paragraph after the table caption" may be a
  paragraph **inside the header cell**; text inserted there wrecks the header;
- `insert_table_row` fills cells only via `set_table_cell` afterwards;
- multi-paragraph cell content uses `\r` as the separator;
- **the text is written literally** - passing `&gt;= 1.6.0` puts `&gt;` in the
  document. Use real characters or wording ("1.6.0 and newer").

Word offers no way here to create or delete a whole table; that must be done by
hand.

## 11. Check table *contents*, not just structure

In the review document, two tables consisted entirely of template placeholders
(`[Example:]IBM E870 Power8`, `Weblogic`, `Oracle DB`) in a deliverable about
to be signed off. **Four review passes missed it**, because
`get_document_outline` only shows a preview of each cell and `find_text` cannot
find a placeholder nobody thought to search for.

**Rule:** reviewing a document means three passes - `get_document_outline` for
structure and styles, `read_table` for table **contents**, and `find_text` for
specific strings. The last one alone finds nothing structural.

## 12. Structure hides in styles, not in text

Two serious defects were invisible to any text search:

- a subsection was **pasted into the middle of another one**, so readers met
  the first section's rules under the second one's heading;
- a whole block had become **items 8-15 of a numbered list**, making the
  document state that "IP-allocation rule 15" was "disabling CloudTrail is
  forbidden".

Both were only visible by walking paragraphs and watching `style`, `list_type`
and `list_string`.

Also useful: `get_paragraph_formatting` reporting `bold: "mixed"` means the
paragraph's runs disagree - handy for spotting hand-formatted text.

Note that `remove_list` + `set_paragraph_style` is not enough to turn a bullet
into a sub-heading: Word keeps the bullet's indentation, so
`set_paragraph_indent` is needed too. Copy the values from a comparable
existing heading. And un-numbering one item mid-list does **not** renumber the
rest.

## 13. `insert_at_selection` can write somewhere else entirely

The worst trap found. It can write into a completely different paragraph than
`goto_text` just selected, splitting a word in the process.

**Rule:** prefer `replace_text_in_paragraph` and `set_table_cell` - they take
an explicit index and have never written out of bounds. If a write appears not
to have happened, **do not immediately retry with another method**; first
`find_text` a fragment of the new text across the whole document to check
whether it landed somewhere unexpected.

Note that `insert_at_selection` replaces the **selected range**, so a selection
covering a hyperlink or URL will eat it.

## 14. COM errors that are not bugs

| Error | Cause | Way around |
|---|---|---|
| `The range cannot be deleted` (-2147352567) | a field, comment or revision object shifts the Range offsets | search a shorter fragment, or `goto_text` + `insert_at_selection` |
| `The TrackRevisions method or property is not available because comment card is selected in pane` | the user has a comment card selected in Word | `list_tables` / `read_table` still work; use them to check structure |
| `Call was rejected by callee` (-2147418111) | Word is busy | retry with a smaller `limit` |

## 15. Code blocks use soft line breaks

A "code block" is often **one paragraph containing `\x0b`** (vertical tab,
Shift+Enter), not several paragraphs. Searching with `\n` finds nothing; in
Word's Find syntax a soft break is `^l`.

## What the server now enforces

Documentation nobody reads prevents nothing, so the painful lessons became
behaviour:

- **`expect_author` / `expect_contains` guards** on `reply_to_comment` and
  `resolve_comment` - the write is refused, with an explanatory error, when the
  index no longer points at that thread. **Required** on `delete_comment`,
  since deletion cannot be undone.
- **Writes echo the live thread back** (root author, root text, anchor,
  resolved state, replies) so a wrong target is visible in the result rather
  than weeks later.
- **`save=True` by default** on comment writes, because saving per operation is
  what stopped the rollbacks.
- **`resolve_comment` re-reads the flag** after saving and raises if Word did
  not keep it.
- **`add_comment` trims the paragraph mark** so a comment never spans a
  paragraph break.
- `replace_text_in_paragraph` avoids `Find.Execute(Replace=...)` entirely.

The remaining lessons cannot be enforced in code - they are judgement calls
about how to review a document - which is why they are written down here.
