import logging
from pathlib import Path
from typing import Dict, List, Tuple, Union

import pywintypes  # win32com.client.pywintypes
import win32com.client
import win32con
import win32ui
from dmfo.constants.mso.wd import (
    WdCompareDestination,
    WdSaveOptions,
    WdUseFormattingFrom,
    WdWindowState,
)
from dmfo.driver.common import ask_resolved, init_com_obj

logger = logging.getLogger(__name__)


def wd(filedata_map: Dict[str, object]) -> int:
    ret, COMObj = init_com_obj("Word")  # noqa: N806
    if ret:
        return ret

    try:
        for alias in ["BASE", "LOCAL", "REMOTE"]:
            filename = filedata_map[alias].get_name()
            # TODO: progressbar
            logger.debug("Opening '%s' ('%s')", alias, filename)
            filedata_map[alias].fileobj = COMObj.Documents.Open(  # noqa: N806
                FileName=str(filename),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            logger.debug("Done")
            # TODO: progressbar

        for alias in ["LOCAL", "REMOTE"]:
            logger.debug("Diffing '%s' vs 'BASE'", alias)
            # TODO: progressbar
            filedata_map[alias].fileobj = COMObj.CompareDocuments(
                OriginalDocument=filedata_map["BASE"].fileobj,
                RevisedDocument=filedata_map[alias].fileobj,
                Destination=WdCompareDestination.wdCompareDestinationRevised,
                CompareFormatting=True,
                CompareCaseChanges=True,
                CompareWhitespace=True,
                CompareTables=True,
                CompareHeaders=True,
                CompareFootnotes=True,
                CompareTextboxes=True,
                CompareFields=True,
                CompareComments=True,
                RevisedAuthor=alias,
                IgnoreAllComparisonWarnings=True,
            )
            logger.debug("Done")
            # TODO: progressbar

        logger.debug("Merging changes")
        # TODO: progressbar
        filedata_map["MERGE"].fileobj = COMObj.MergeDocuments(
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
            OriginalAuthor="Merge LOCAL",
            RevisedAuthor="Merge REMOTE",
            FormatFrom=WdUseFormattingFrom.wdFormattingFromPrompt,
        )
        logger.debug("Done")
        filename = filedata_map["LOCAL"].get_name()
        # TODO: progressbar

        for alias in ["BASE", "LOCAL", "REMOTE"]:
            logger.debug("Closing '%s'", alias)
            # TODO: progressbar
            filedata_map[alias].fileobj.Close(
                SaveChanges=WdSaveOptions.wdDoNotSaveChanges
            )
            # filedata_map[alias].pop("Object")
            logger.debug("Done")
            # TODO: progressbar

        logger.debug("Saving 'MERGE'")
        # TODO: progressbar
        filedata_map["MERGE"].fileobj.SaveAs(
            FileName=str(filename),
            AddToRecentFiles=False,
        )
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

    logger.debug("Asking for merge confirmation...")
    is_resolved = ask_resolved()
    logger.debug("Reply: '%s'", is_resolved)
    # TODO: progressbar

    try:
        logger.debug("Checking if 'MERGE' is still open...")
        COMObj.Documents.Item(str(filename))
        logger.debug("'MERGE' is still open.")
    except pywintypes.com_error as exc:
        reopen = True
        if exc.args[0] in [-2147352567]:
            logger.debug("'MERGE' has been closed, reopening...")
        elif exc.args[0] in [-2147023174, -2147023179]:
            logger.debug("COMObj has been closed, reinitializing...")
            ret, COMObj = init_com_obj("Word")  # noqa: N806
            if ret:
                return ret
        else:
            logger.error("COM Error: '%s'", exc.args[1])
            logger.debug("COM Error: '%s'", exc)
    else:
        reopen = False

    COMObj.Visible = False

    if reopen:
        logger.debug("Opening '%s' ('%s')", "MERGE", filedata_map["LOCAL"].name)
        filedata_map["MERGE"].fileobj = COMObj.Documents.Open(  # noqa: N806
            FileName=str(filename),
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )
        logger.debug("Done")
        # TODO: progressbar

    if filedata_map["MERGE"].fileobj.TrackRevisions:
        logger.warning("Warning: Track Changes is active. Please deactivate!")
        # Deactivate track changes?
    if is_resolved:
        if filedata_map["MERGE"].fileobj.Revisions.Count > 0:
            is_resolved = False  # TODO: Transfer to ps1 script
            logger.warning(
                "Warning: Unresolved revisions in the document. "
                + "Will exit as 'unresolved'."
            )
    # TODO: progressbar

    # TODO: progressbar
    filedata_map["MERGE"].fileobj.Close()
    if COMObj.Documents.Count == 0:
        logger.debug("No more open documents in COMObj, closing...")
        COMObj.Quit()
        logger.debug("Done")
    # filedata_map["MERGE"].pop("Object")
    # TODO: progressbar

    # Return as return code: is_resolved -> True: 0, False: 1
    return int(not is_resolved)
