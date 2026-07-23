"""Allow running as a module: python -m onword_mcp

Fallback for Windows environments where AV/AppLocker blocks the
uv-generated console-script trampoline exe (Access is denied, os error 5).
"""

from .server import main

main()