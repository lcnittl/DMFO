import logging
from pathlib import Path
from typing import Dict, List, Tuple, Union

import pywintypes  # win32com.client.pywintypes
import win32com.client
import win32con
import win32ui
from classes import VCSFileData
from constants.mso.mso import MsoTriState
from constants.mso.pp import PpWindowState
from constants.mso.wd import (
    WdCompareDestination,
    WdSaveOptions,
    WdUseFormattingFrom,
    WdWindowState,
)
from driver import init_com_obj

logger = logging.getLogger(__name__)


def differ(filedata_map: Dict[str, object]) -> int:
    filedata_map["DIFF"] = VCSFileData(Path())

    extension = VCSFileData.target_ext
    if extension in [".doc", ".docx"]:
        ret = diff_wd(filedata_map=filedata_map)
    elif extension in [".ppt", ".pptx"]:
        ret = diff_pp(filedata_map=filedata_map)
    else:
        logger.critical(
            "DMFO-Diff does not know what to do with '%s' files.", extension
        )
        ret = 2

    filedata_map.pop("DIFF")
    return ret


def diff_pp() -> int:
    pass


def diff_wd(filedata_map: Dict[str, object]) -> int:
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
