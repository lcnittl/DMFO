import logging
from pathlib import Path
from typing import Tuple

import pywintypes  # win32com.client.pywintypes
import win32com.client
import win32con
import win32ui
from dmfo.classes import VCSFileData

logger = logging.getLogger(__name__)


def ask_resolved() -> bool:
    ret = win32ui.MessageBox(
        "Confirm conflict resolution?",
        "Merge Complete?",
        win32con.MB_YESNO | win32con.MB_ICONQUESTION,
    )

    # win32ui.MessageBox returns 6 for "Yes" and 7 for "No"
    return ret == 6


def init_com_obj(app_name: str) -> Tuple[int, object]:
    logger.debug("Initializing COM object...")
    # TODO: progressbar
    try:
        com_obj = win32com.client.DispatchEx(f"{app_name}.Application")
        com_obj.Visible = False
        logger.debug("Done")
        # TODO: progressbar
    except pywintypes.com_error as exc:
        logger.critical(
            "You must have Microsoft %s installed to perform this operation.", app_name
        )
        logger.debug("COM Error: '%s'", exc)
        return (3, None)
    return (0, com_obj)
