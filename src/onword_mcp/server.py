"""onword-mcp server entry point.

Transport selection (CLI args take precedence over env vars):
  --transport / MCP_TRANSPORT : stdio (default), streamable-http, http, sse
  --port      / MCP_PORT      : default 18347
  --host      / MCP_HOST      : default 127.0.0.1

Examples:
  onword-mcp
  onword-mcp --transport streamable-http --port 18347
  cmd /c "set MCP_TRANSPORT=streamable-http&& set MCP_PORT=18347&& onword-mcp"
"""

import argparse
import os
import sys

from fastmcp import FastMCP

from . import (
    __version__,
    tools_comments,
    tools_format,
    tools_read,
    tools_table,
    tools_write,
)

DEFAULT_PORT = 18347
DEFAULT_HOST = "127.0.0.1"

INSTRUCTIONS = """\
Live editor for the document currently open in Microsoft Word on this machine.
Workflow for large documents:
1. get_document_info / get_document_outline to orient (paginated, token-cheap).
2. read_paragraphs / get_page_paragraphs / find_text to locate content.
3. Edit surgically: replace_paragraph, insert_after_paragraph, formatting tools.
4. For tables use the table tools (list_tables, read_table, set_table_cell,
   insert_table_row) - a table cell is its own paragraph, so paragraph-level
   inserts would corrupt the table layout.
5. For review remarks use the comment tools (list_comments shows unresolved
   threads, goto_comment shows the anchor in Word, reply_to_comment records
   how a remark was addressed, resolve_comment closes the thread).
Never ask for the whole document text - always work with paragraph indexes.

ALL INDEXES ARE POSITIONAL, NOT STABLE IDS. Paragraph, table and comment
indexes shift whenever content is inserted, moved or deleted. Re-locate the
target immediately before each write; never reuse an index from earlier in the
conversation or from your notes.

Two habits this server was hardened around, learned the hard way:
- Verify AFTER saving. Writes have been seen to silently roll back - the tool
  reports success and the change is gone. Read the state back.
- On comments, pass expect_author to reply_to_comment / resolve_comment (it is
  required by delete_comment). A shifted index otherwise writes onto a
  stranger's thread and nothing complains.

find_text and goto_text often report an index one lower than the real one, so
read both index and index+1 before editing.
"""

mcp = FastMCP("onword-mcp", instructions=INSTRUCTIONS)

tools_read.register(mcp)
tools_write.register(mcp)
tools_format.register(mcp)
tools_table.register(mcp)
tools_comments.register(mcp)



def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        prog="onword-mcp",
        description="MCP server for live editing of the open Microsoft Word document.",
    )
    parser.add_argument(
        "--transport",
        default=os.environ.get("MCP_TRANSPORT", "stdio"),
        choices=["stdio", "streamable-http", "http", "sse"],
        help="MCP transport (default: stdio, or MCP_TRANSPORT env var)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.environ.get("MCP_PORT", DEFAULT_PORT)),
        help=f"Port for HTTP/SSE transports (default: {DEFAULT_PORT}, or MCP_PORT env var)",
    )
    parser.add_argument(
        "--host",
        default=os.environ.get("MCP_HOST", DEFAULT_HOST),
        help=f"Host to bind for HTTP/SSE transports (default: {DEFAULT_HOST}, or MCP_HOST env var)",
    )
    parser.add_argument("--version", action="version", version=f"onword-mcp {__version__}")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    transport = args.transport
    if transport == "streamable-http":
        transport = "http"
    if transport == "stdio":
        mcp.run(transport="stdio")
    else:
        print(
            f"onword-mcp {__version__} listening on {transport}://{args.host}:{args.port}/",
            file=sys.stderr,
        )
        mcp.run(transport=transport, host=args.host, port=args.port)


if __name__ == "__main__":
    main()