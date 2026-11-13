"""Allow direct module execution via `python -m excel_reconciliation`."""

from __future__ import annotations

from excel_reconciliation.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
