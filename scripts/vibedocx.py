"""Entry point for vibedocx — invoked as `python scripts/vibedocx.py <cmd>`."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from helper.cli import main

if __name__ == "__main__":
    main()
