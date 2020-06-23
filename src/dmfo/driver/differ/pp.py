import logging
from pathlib import Path
from typing import Dict

import pywintypes  # win32com.client.pywintypes
import win32com.client
import win32con
import win32ui
from dmfo.classes import VCSFileData
from dmfo.constants.mso.mso import MsoTriState
from dmfo.constants.mso.pp import PpWindowState
from dmfo.driver.common import init_com_obj

logger = logging.getLogger(__name__)


def pp() -> int:
    pass
