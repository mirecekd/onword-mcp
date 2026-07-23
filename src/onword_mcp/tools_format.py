"""Formatting tools: styles, lists/bullets, indentation, alignment, fonts."""

import json

from .word import (
    WD_ALIGN,
    WD_ALIGN_REVERSE,
    WD_LIST_TYPE,
    check_paragraph_index,
    cm_to_points,
    points_to_cm,
    word_session,
    word_write_session,
)


def register(mcp):
    @mcp.tool()
    def get_paragraph_formatting(index: int) -> str:
        """Get detailed formatting of a paragraph: style, alignment, indents,
        list/bullet info (type and level), and font of the first run."""
        with word_session() as (word, doc):
            check_paragraph_index(doc, index)
            p = doc.Paragraphs(index)
            rng = p.Range
            fmt = p.Format
            lf = rng.ListFormat
            info = {
                "index": index,
                "style": str(p.Style),
                "alignment": WD_ALIGN_REVERSE.get(int(fmt.Alignment), str(fmt.Alignment)),
                "left_indent_cm": points_to_cm(fmt.LeftIndent),
                "first_line_indent_cm": points_to_cm(fmt.FirstLineIndent),
                "list_type": WD_LIST_TYPE.get(int(lf.ListType), str(lf.ListType)),
            }
            if int(lf.ListType) != 0:
                info["list_level"] = int(lf.ListLevelNumber)
                info["list_string"] = str(lf.ListString)
            font = rng.Font
            info["font"] = {
                "name": str(font.Name),
                "size": float(font.Size) if font.Size > 0 else None,
                "bold": bool(font.Bold) if font.Bold in (0, 1, True, False) else "mixed",
                "italic": bool(font.Italic) if font.Italic in (0, 1, True, False) else "mixed",
            }
            return json.dumps(info, ensure_ascii=False, indent=2)

    @mcp.tool()
    def list_document_styles(styles_in_use_only: bool = True) -> str:
        """List paragraph style names available in the document. Use these
        names with set_paragraph_style."""
        with word_session() as (word, doc):
            names = []
            for s in doc.Styles:
                try:
                    if s.Type == 1:  # wdStyleTypeParagraph
                        if styles_in_use_only and not s.InUse:
                            continue
                        names.append(str(s.NameLocal))
                except Exception:
                    continue
            return "\n".join(sorted(names))

    @mcp.tool()
    def set_paragraph_style(index: int, style_name: str) -> str:
        """Apply a paragraph style (e.g. 'Heading 1', 'Normal', 'List Bullet')
        to the paragraph at the given index. Use list_document_styles to see
        available style names."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            doc.Paragraphs(index).Style = style_name
            return f"Style '{style_name}' applied to paragraph {index}."

    @mcp.tool()
    def list_indent(index: int, count: int = 1) -> str:
        """Demote (move RIGHT) list item(s)/bullet(s) by one level, starting
        at paragraph index. Equivalent to pressing Tab in a Word list."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            end = min(index + count - 1, doc.Paragraphs.Count)
            for i in range(index, end + 1):
                doc.Paragraphs(i).Range.ListFormat.ListIndent()
            return f"Paragraph(s) {index}..{end} indented one list level right."

    @mcp.tool()
    def list_outdent(index: int, count: int = 1) -> str:
        """Promote (move LEFT) list item(s)/bullet(s) by one level, starting
        at paragraph index. Equivalent to pressing Shift+Tab in a Word list."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            end = min(index + count - 1, doc.Paragraphs.Count)
            for i in range(index, end + 1):
                doc.Paragraphs(i).Range.ListFormat.ListOutdent()
            return f"Paragraph(s) {index}..{end} outdented one list level left."

    @mcp.tool()
    def set_list_level(index: int, level: int) -> str:
        """Set the exact list level (1-9) of a list item/bullet at the given
        paragraph index."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            lf = doc.Paragraphs(index).Range.ListFormat
            if int(lf.ListType) == 0:
                return f"Paragraph {index} is not a list item."
            lf.ListLevelNumber = level
            return f"Paragraph {index} set to list level {level}."

    @mcp.tool()
    def convert_to_list(index: int, count: int = 1, numbered: bool = False) -> str:
        """Convert paragraph(s) to a bulleted list (or numbered list when
        numbered=True), starting at the given index."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            end = min(index + count - 1, doc.Paragraphs.Count)
            start_rng = doc.Paragraphs(index).Range.Start
            end_rng = doc.Paragraphs(end).Range.End
            rng = doc.Range(start_rng, end_rng)
            if numbered:
                rng.ListFormat.ApplyNumberDefault()
            else:
                rng.ListFormat.ApplyBulletDefault()
            kind = "numbered" if numbered else "bulleted"
            return f"Paragraph(s) {index}..{end} converted to {kind} list."

    @mcp.tool()
    def remove_list(index: int, count: int = 1) -> str:
        """Remove bullets/numbering from paragraph(s), starting at the index."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            end = min(index + count - 1, doc.Paragraphs.Count)
            start_rng = doc.Paragraphs(index).Range.Start
            end_rng = doc.Paragraphs(end).Range.End
            doc.Range(start_rng, end_rng).ListFormat.RemoveNumbers()
            return f"List formatting removed from paragraph(s) {index}..{end}."

    @mcp.tool()
    def set_paragraph_indent(
        index: int, left_cm: float | None = None, first_line_cm: float | None = None
    ) -> str:
        """Set left indent and/or first-line indent of a paragraph in
        centimeters (for non-list paragraphs; for lists use list_indent)."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            fmt = doc.Paragraphs(index).Format
            changes = []
            if left_cm is not None:
                fmt.LeftIndent = cm_to_points(left_cm)
                changes.append(f"left={left_cm}cm")
            if first_line_cm is not None:
                fmt.FirstLineIndent = cm_to_points(first_line_cm)
                changes.append(f"first_line={first_line_cm}cm")
            if not changes:
                return "No indent values provided."
            return f"Paragraph {index} indent set: {', '.join(changes)}."

    @mcp.tool()
    def set_paragraph_alignment(index: int, alignment: str) -> str:
        """Set paragraph alignment: 'left', 'center', 'right' or 'justify'."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            key = alignment.lower()
            if key not in WD_ALIGN:
                return f"Unknown alignment '{alignment}'. Use: {', '.join(WD_ALIGN)}."
            doc.Paragraphs(index).Format.Alignment = WD_ALIGN[key]
            return f"Paragraph {index} aligned {key}."

    @mcp.tool()
    def format_text(
        index: int,
        find: str,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        size: float | None = None,
        font_name: str | None = None,
    ) -> str:
        """Apply character formatting (bold/italic/underline/size/font) to the
        first occurrence of 'find' inside the paragraph at the given index."""
        with word_write_session() as (word, doc):
            check_paragraph_index(doc, index)
            rng = doc.Paragraphs(index).Range
            f = rng.Find
            f.ClearFormatting()
            f.Text = find
            f.Wrap = 0  # wdFindStop
            if not f.Execute():
                return f"'{find}' not found in paragraph {index}."
            font = rng.Font
            changes = []
            if bold is not None:
                font.Bold = bold
                changes.append(f"bold={bold}")
            if italic is not None:
                font.Italic = italic
                changes.append(f"italic={italic}")
            if underline is not None:
                font.Underline = 1 if underline else 0
                changes.append(f"underline={underline}")
            if size is not None:
                font.Size = size
                changes.append(f"size={size}")
            if font_name is not None:
                font.Name = font_name
                changes.append(f"font={font_name}")
            if not changes:
                return "No formatting attributes provided."
            return f"Formatted '{find}' in paragraph {index}: {', '.join(changes)}."