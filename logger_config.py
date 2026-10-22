import logging
import sys


LOG_FILE = "metal_rates_automator.log"


def setup_logging():
    logger = logging.getLogger()

    # Prevent duplicate handlers if setup_logging()
    # is called more than once.
    if logger.handlers:
        return

    logger.setLevel(logging.INFO)

    # --------------------------------------------------
    # Terminal
    # Only ERROR and CRITICAL appear in PowerShell
    # --------------------------------------------------
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.ERROR)

    console_formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    )

    console_handler.setFormatter(console_formatter)

    # --------------------------------------------------
    # Log file
    # INFO + WARNING + ERROR
    # --------------------------------------------------
    file_handler = logging.FileHandler(
        LOG_FILE,
        encoding="utf-8"
    )

    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(console_formatter)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)