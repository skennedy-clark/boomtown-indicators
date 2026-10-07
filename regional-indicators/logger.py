"""
regional-indicators/logger.py

Logging for the fetchers and run_update.py.

Each run writes a timestamped log file to logs/ and prints colour-coded
messages to the console.
"""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

# ANSI colour codes for console output
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
    """Return the root logger, writing to the console and to a timestamped
    log file. Handlers are attached once; later calls return the same
    logger.
    """
    logger = logging.getLogger(name)

    if logger.handlers:
        # Already configured.
        return logger

    logger.setLevel(logging.DEBUG)

    ts = run_timestamp or datetime.now().strftime("%Y-%m-%d_%H%M%S")

    # ── File handler: plain text, DEBUG and above ─────────────────────────────
    log_path = LOG_DIR / f"{ts}.log"
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))

    # ── Console handler: coloured, INFO and above ─────────────────────────────
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
    """Return a child logger, for example
    get_child_logger('regional-indicators', 'fetch_income'). It uses the
    parent's handlers, so all output goes to the same log file.
    """
    return logging.getLogger(f"{parent_name}.{child_name}")
