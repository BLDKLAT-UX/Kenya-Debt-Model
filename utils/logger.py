# -*- coding: utf-8 -*-
"""
utils/logger.py
===============
Centralised logging for the Kenya Debt Model.

- Logs to both terminal (coloured) and logs/model_run_YYYYMMDD_HHMMSS.log
- Log files persist as an audit trail of every model run
- Import and use anywhere in the project:

    from utils.logger import get_logger
    log = get_logger(__name__)
    log.info("Building COVER sheet ...")
    log.warning("GDP growth below 3% — check assumptions")
    log.error("Failed to write Excel file")
"""

import logging
import os
import sys
from datetime import datetime
from pathlib import Path

# ── Paths
ROOT_DIR = Path(__file__).resolve().parent.parent
LOGS_DIR = ROOT_DIR / "logs"
LOGS_DIR.mkdir(exist_ok=True)

# ── Log file name stamped with run datetime
_LOG_FILE = LOGS_DIR / f"model_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

# ── Colour codes for terminal output (Windows-safe)
_COLOURS = {
    "DEBUG":    "\033[36m",   # cyan
    "INFO":     "\033[92m",   # green
    "WARNING":  "\033[93m",   # yellow
    "ERROR":    "\033[91m",   # red
    "CRITICAL": "\033[95m",   # magenta
    "RESET":    "\033[0m",
    "DIM":      "\033[2m",
    "BOLD":     "\033[1m",
}

_COLOURS_ENABLED = sys.stdout.isatty() or os.environ.get("FORCE_COLOR", "0") == "1"


class _ColouredFormatter(logging.Formatter):
    """Terminal formatter with colour-coded level names and timestamps."""

    FMT = "{dim}{time}{reset}  {colour}{bold}{level:<8}{reset}  {name}  {msg}"

    def format(self, record: logging.LogRecord) -> str:  # noqa: A003
        if _COLOURS_ENABLED:
            c = _COLOURS.get(record.levelname, "")
            reset = _COLOURS["RESET"]
            dim = _COLOURS["DIM"]
            bold = _COLOURS["BOLD"]
        else:
            c = reset = dim = bold = ""

        time_str = datetime.fromtimestamp(record.created).strftime("%H:%M:%S")
        return self.FMT.format(
            dim=dim, time=time_str, reset=reset,
            colour=c, bold=bold,
            level=record.levelname,
            name=f"{record.name}",
            msg=record.getMessage(),
        )


class _FileFormatter(logging.Formatter):
    """Plain formatter for log files — no colour codes."""

    FMT = "%(asctime)s  %(levelname)-8s  %(name)s  %(message)s"
    DATEFMT = "%Y-%m-%d %H:%M:%S"

    def __init__(self):
        super().__init__(fmt=self.FMT, datefmt=self.DATEFMT)


# ── Root logger setup (only configure once)
_configured = False


def _configure_root() -> None:
    global _configured
    if _configured:
        return

    root = logging.getLogger("kenya_debt")
    root.setLevel(logging.DEBUG)

    # Terminal handler
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(_ColouredFormatter())
    root.addHandler(ch)

    # File handler
    fh = logging.FileHandler(_LOG_FILE, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(_FileFormatter())
    root.addHandler(fh)

    _configured = True


def get_logger(name: str) -> logging.Logger:
    """
    Return a logger namespaced under 'kenya_debt.<name>'.

    Usage::

        from utils.logger import get_logger
        log = get_logger(__name__)
        log.info("Starting model run")
    """
    _configure_root()
    return logging.getLogger(f"kenya_debt.{name}")


def log_section(log: logging.Logger, title: str) -> None:
    """Print a prominent section divider in the logs."""
    bar = "─" * 55
    log.info(bar)
    log.info(f"  {title}")
    log.info(bar)


def log_run_summary(log: logging.Logger, outputs: list[tuple[str, bool]]) -> None:
    """
    Print a final run summary table.

    outputs: list of (filename, success) tuples
    """
    log.info("=" * 55)
    log.info("  RUN SUMMARY")
    log.info("=" * 55)
    all_ok = True
    for fname, ok in outputs:
        status = "OK " if ok else "FAIL"
        log.info(f"  [{status}]  {fname}")
        if not ok:
            all_ok = False
    log.info("─" * 55)
    if all_ok:
        log.info(f"  All outputs generated. Log saved → {_LOG_FILE.name}")
    else:
        log.error("  One or more steps FAILED. Check log above.")
    log.info("=" * 55)
