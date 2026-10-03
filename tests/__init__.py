"""Test package: puts src/python (the plugin) and the project root (build_oxt) on sys.path."""
import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (os.path.join(_ROOT, "src", "python"), _ROOT):
    if _p not in sys.path:
        sys.path.insert(0, _p)
