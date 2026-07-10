from logger import get_logger

log = get_logger("TEST")

log.info("Logger is working.")

try:
    x = 5 / 0
except Exception:
    log.exception("Example exception")