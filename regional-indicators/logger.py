"""
logger.py
---------
Centralised logging for all fetchers and the orchestrator.
Each run gets its own timestamped log file in logs/.
Console output is also shown, colour-coded by level.
"""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

# ANSI colour codes for console
_COLOURS = {
    "DEBUG":    "\033[36m",   # cyan
    "INFO":     "\033[32m",   # green
    "WARNING":  "\033[33m",   # yellow
    "ERROR":    "\033[31m",   # red
    "CRITICAL": "\033[35m",   # magenta
    "RESET":    "\033[0m",
}


class ColouredFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        colour = _COLOURS.get(record.levelname, "")
        reset  = _COLOURS["RESET"]
        record.levelname = f"{colour}{record.levelname:<8}{reset}"
        return super().format(record)


def get_logger(name: str = "regional-indicators",
               run_timestamp: str | None = None) -> logging.Logger:
    """
    Returns a logger that writes to both the console and a timestamped
    log file.  Call once from run_update.py; subsequent get_logger() calls
    with the same name return the same logger instance.
    """
    logger = logging.getLogger(name)

    if logger.handlers:
        # Already configured — return existing instance
        return logger

    logger.setLevel(logging.DEBUG)

    ts = run_timestamp or datetime.now().strftime("%Y-%m-%d_%H%M%S")

    # ── File handler (plain text, DEBUG+) ─────────────────────────────────────
    log_path = LOG_DIR / f"{ts}.log"
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))

    # ── Console handler (coloured, INFO+) ─────────────────────────────────────
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(ColouredFormatter(
        "%(asctime)s  %(levelname)s  %(message)s",
        datefmt="%H:%M:%S",
    ))

    logger.addHandler(fh)
    logger.addHandler(ch)

    logger.info(f"Log file: {log_path}")
    return logger


def get_child_logger(parent_name: str, child_name: str) -> logging.Logger:
    """
    Returns a child logger, e.g. get_child_logger('regional-indicators', 'fetch_income')
    inherits handlers from parent so everything goes to the same log file.
    """
    return logging.getLogger(f"{parent_name}.{child_name}")
