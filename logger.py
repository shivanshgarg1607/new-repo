"""
logger.py
----------
Central logging utility for the Hospital Review Automation Bot.
"""

import logging
from pathlib import Path

from config import AUTOMATION_LOG, VERBOSE


def get_logger(name: str) -> logging.Logger:
    """
    Returns a configured logger instance.
    """

    logger = logging.getLogger(name)

    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    formatter = logging.Formatter(
        "[%(asctime)s] %(levelname)s | %(name)s | %(message)s",
        "%Y-%m-%d %H:%M:%S"
    )

    # Log to file
    file_handler = logging.FileHandler(AUTOMATION_LOG, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Log to console
    if VERBOSE:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    return logger