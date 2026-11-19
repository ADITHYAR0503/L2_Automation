import logging
import os
from datetime import datetime


def setup_logger() -> logging.Logger:
    os.makedirs("logs", exist_ok=True)
    log_file = os.path.join("logs", f"l2_automation_{datetime.now().strftime('%Y_%m_%d')}.log")

    logger = logging.getLogger("L2_Automation")
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if not logger.handlers:
        formatter = logging.Formatter(
            "%(asctime)s | %(levelname)-7s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )

        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(formatter)

        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        logger.addHandler(stream_handler)

    return logger