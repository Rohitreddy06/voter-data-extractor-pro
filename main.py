"""
main.py
=======
Entry point for Voter Data Extractor Pro.

Run with:
    python main.py
"""

from __future__ import annotations

import sys
from utils import logger


def verify_translation_dependencies() -> None:
    """Check if translator dependencies are ready before launching UI."""
    try:
        import translator

        if not getattr(translator, "_DEEP_TRANSLATION_AVAILABLE", False):
            logger.warning(
                "Translation dependencies (pandas/deep-translator) missing. "
                "Translation feature will be disabled."
            )
    except Exception as exc:
        logger.warning("Could not verify translator module: %s", exc)


def main() -> int:
    # 1. Pre-check translator setup
    verify_translation_dependencies()

    # 2. Load GUI dependencies
    try:
        from gui import VoterExtractorApp
    except ImportError as exc:
        print(
            "Failed to import GUI dependencies. Make sure you have installed "
            "everything in requirements.txt (pip install -r requirements.txt).\n"
            f"Original error: {exc}"
        )
        return 1

    # 3. Initialize and run application
    app = VoterExtractorApp()
    app.protocol("WM_DELETE_WINDOW", app.on_close)

    try:
        app.mainloop()
    except Exception as exc:  # pragma: no cover
        logger.error("Fatal GUI error: %s", exc)
        raise

    return 0


if __name__ == "__main__":
    sys.exit(main())