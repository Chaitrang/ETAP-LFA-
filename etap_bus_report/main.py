"""
main.py
-------
Single entry point for both front ends.

    python main.py                 # desktop application (PySide6)
    streamlit run main.py          # web application  (same pipeline, no Qt)
    python cli.py report.pdf       # batch / CI

When this file is executed by Streamlit - locally or on Streamlit Community
Cloud - it detects the Streamlit runtime and renders the web UI.  PySide6 is
never imported in that case, because a headless server has no display and no
GUI system libraries (libglib).
"""

from __future__ import annotations

import os
import sys

# Make the sibling modules importable no matter which directory the app is
# started from (Streamlit Cloud runs it from the repository root).
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _running_under_streamlit() -> bool:
    """True when this file is being executed by ``streamlit run``."""
    # A script run context exists whenever Streamlit is executing the script.
    for module in (
        "streamlit.runtime.scriptrunner",
        "streamlit.runtime.scriptrunner_utils.script_run_context",
        "streamlit.script_run_context",
    ):
        try:
            get_ctx = __import__(module, fromlist=["get_script_run_ctx"]).get_script_run_ctx
        except Exception:
            continue
        try:
            if get_ctx(suppress_warning=True) is not None:
                return True
        except TypeError:  # older signature without the keyword
            try:
                if get_ctx() is not None:
                    return True
            except Exception:
                continue
        except Exception:
            continue

    # Fallback: a live Streamlit server in this process.
    try:
        from streamlit.runtime import exists

        return bool(exists())
    except Exception:
        return False


def main() -> int:
    if _running_under_streamlit():
        import streamlit_app

        # render() must be called on every script run - Streamlit re-executes
        # this file on each interaction, but the imported module body does not.
        streamlit_app.render()
        return 0

    try:
        import ui
    except ImportError as exc:  # PySide6 missing or unusable
        print(
            "PySide6 is required to run the desktop application.\n"
            "Install the desktop dependencies with:  pip install -r requirements-desktop.txt\n"
            "To run the web version instead:          streamlit run main.py\n"
            f"({exc})",
            file=sys.stderr,
        )
        return 1
    return ui.run()


if _running_under_streamlit():
    # Streamlit executes the module top to bottom, with __name__ == "__main__".
    # It must not see a SystemExit, which would discard the rendered page.
    main()
elif __name__ == "__main__":
    raise SystemExit(main())
