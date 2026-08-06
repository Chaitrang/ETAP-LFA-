"""
main.py
-------
Application entry point.

    python main.py            # launches the desktop application
    python cli.py report.pdf  # same pipeline without a GUI
"""

from __future__ import annotations

import sys


def main() -> int:
    try:
        import ui
    except ImportError as exc:  # PySide6 missing
        print(
            "PySide6 is required to run the desktop application.\n"
            "Install the dependencies with:  pip install -r requirements.txt\n"
            f"({exc})",
            file=sys.stderr,
        )
        return 1
    return ui.run()


if __name__ == "__main__":
    raise SystemExit(main())
