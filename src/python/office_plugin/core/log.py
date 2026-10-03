"""File logging into the LibreOffice user profile (never raises)."""

import logging
import logging.handlers
import os

_NAME = "office_plugin"


def setup(base_dir=None, debug=False):
    """Configure and return the plugin logger. `base_dir=None` -> no file output."""
    logger = logging.getLogger(_NAME)
    logger.setLevel(logging.DEBUG if debug else logging.INFO)
    for h in list(logger.handlers):
        logger.removeHandler(h)
        h.close()
    handler = None
    if base_dir:
        try:
            os.makedirs(base_dir, exist_ok=True)
            handler = logging.handlers.RotatingFileHandler(
                os.path.join(base_dir, "office_plugin.log"),
                maxBytes=512 * 1024, backupCount=2, encoding="utf-8")
            handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        except OSError:
            handler = None
    logger.addHandler(handler or logging.NullHandler())
    logger.propagate = False
    return logger


def get():
    return logging.getLogger(_NAME)
