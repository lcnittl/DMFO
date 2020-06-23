import logging
from pathlib import Path
from typing import Dict

import pywintypes  # win32com.client.pywintypes
import win32com.client
import win32con
import win32ui
from dmfo.classes import VCSFileData
from dmfo.constants.mso.wd import (
    WdCompareDestination,
    WdSaveOptions,
    WdUseFormattingFrom,
    WdWindowState,
)
from dmfo.driver.common import init_com_obj

logger = logging.getLogger(__name__)


def wd(filedata_map: Dict[str, object]) -> int:
    ret, COMObj = init_com_obj("Word")  # noqa: N806
    if ret:
        return ret

    try:
        for alias in ["LOCAL", "REMOTE"]:
            filename = filedata_map[alias].get_name()
            logger.debug("Opening '%s' ('%s')", alias, filename)
            # TODO: progressbar
            filedata_map[alias].fileobj = COMObj.Documents.Open(
                FileName=str(filename),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            logger.debug("Done")
            # TODO: progressbar

        logger.debug("Diffing 'REMOTE' vs 'LOCAL'")
        # TODO: progressbar
        filedata_map["DIFF"].fileobj = COMObj.CompareDocuments(
            OriginalDocument=filedata_map["LOCAL"].fileobj,
            RevisedDocument=filedata_map["REMOTE"].fileobj,
            Destination=WdCompareDestination.wdCompareDestinationNew,
            CompareFormatting=True,
            CompareCaseChanges=True,
            CompareWhitespace=True,
            CompareTables=True,
            CompareHeaders=True,
            CompareFootnotes=True,
            CompareTextboxes=True,
            CompareFields=True,
            CompareComments=True,
            RevisedAuthor="REMOTE",
            IgnoreAllComparisonWarnings=True,
        )
        logger.debug("Done")
        # TODO: progressbar

        for alias in ["LOCAL", "REMOTE"]:
            logger.debug("Closing '%s'", alias)
            # TODO: progressbar
            filedata_map[alias].fileobj.Close(
                SaveChanges=WdSaveOptions.wdDoNotSaveChanges
            )
            # filedata_map[alias].pop("Object")
            logger.debug("Done")
            # TODO: progressbar

        logger.debug("Setting 'DIFF' to unsaved")
        # TODO: progressbar
        filedata_map["DIFF"].fileobj.Saved = 1
        logger.debug("Done")
        # TODO: progressbar

        logger.debug("Bringing to foreground")
        # TODO: progressbar
        COMObj.Visible = True
        COMObj.Activate()
        COMObj.WindowState = WdWindowState.wdWindowStateMinimize
        COMObj.WindowState = WdWindowState.wdWindowStateMaximize
        logger.debug("Done")
        # TODO: progressbar
    except pywintypes.com_error as exc:
        logger.error("COM Error: '%s'", exc.args[1])
        logger.debug("COM Error: '%s'", exc)
        return 6
    return 0
