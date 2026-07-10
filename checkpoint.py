"""
checkpoint.py
--------------
Handles saving and loading automation checkpoints.
"""

import json
from datetime import datetime
from pathlib import Path

from config import CHECKPOINT_FILE
from logger import get_logger

log = get_logger("CHECKPOINT")


def save_checkpoint(
    phase: str,
    hospital_index: int,
    hospital_name: str,
    status: str = "running"
):
    """
    Save the current automation state.
    """

    data = {
        "phase": phase,
        "hospital_index": hospital_index,
        "hospital_name": hospital_name,
        "status": status,
        "last_updated": datetime.now().isoformat()
    }

    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)

    log.info(f"Checkpoint saved -> {hospital_name}")


def load_checkpoint():
    """
    Load checkpoint if it exists.
    """

    if not CHECKPOINT_FILE.exists():
        return None

    with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def clear_checkpoint():
    """
    Delete checkpoint after successful completion.
    """

    if CHECKPOINT_FILE.exists():
        CHECKPOINT_FILE.unlink()
        log.info("Checkpoint cleared.")