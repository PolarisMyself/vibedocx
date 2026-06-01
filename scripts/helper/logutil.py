"""Logging setup for vibedocx.

Usage:
    from helper.logutil import logger, setup_logging
    setup_logging(verbose=True)
    logger.debug("Detailed operation info")
    logger.info("Operation completed")
"""

import logging
import sys

logger = logging.getLogger('vibedocx')

_initialized = False


def setup_logging(verbose=False, quiet=False):
    """Initialize vibedocx logging. Call once at CLI startup.

    Args:
        verbose: Enable DEBUG level output.
        quiet: Suppress all output except errors.
    """
    global _initialized
    if _initialized:
        return

    handler = logging.StreamHandler(sys.stderr)
    if quiet:
        handler.setLevel(logging.ERROR)
    elif verbose:
        handler.setLevel(logging.DEBUG)
    else:
        handler.setLevel(logging.INFO)

    handler.setFormatter(logging.Formatter(
        '%(levelname)s [%(name)s] %(message)s'
    ))
    logger.addHandler(handler)
    logger.setLevel(logging.DEBUG)  # Logger allows all; handler filters

    _initialized = True
