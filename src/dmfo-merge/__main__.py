#!/usr/bin/env python3
"""DMFO-Merge

Merge driver for Office documents. Opens files in MSO to merge them.
"""
import argparse
import logging
import logging.handlers
import shlex
import shutil
import subprocess  # nosec
import sys
import tkinter as tk
import tkinter.messagebox
from pathlib import Path

import colorlog
import pywintypes
import win32com.client

LOG_LVLS = {
    # "NOTSET": logging.NOTSET,  # 0
    "DEBUG": logging.DEBUG,  # 10
    "INFO": logging.INFO,  # 20
    "WARNING": logging.WARNING,  # 30
    "ERROR": logging.ERROR,  # 40
    "CRITICAL": logging.CRITICAL,  # 50
}


class MsoTriState:
    # Not supported.
    # msoTriStateToggle = -3
    # Not supported.
    # msoTriStateMixed = -2
    # True.
    msoTrue = -1  # noqa: N815
    # False.
    msoFalse = 0  # noqa: N815
    # Not supported.
    # msoCTrue = 1


class WdSaveOptions:
    # > Prompt the user to save pending changes.
    # wdPromptToSaveChanges = -2
    # > Save pending changes automatically without prompting the user.
    # wdSaveChanges = -1
    # > Do not save pending changes.
    wdDoNotSaveChanges = 0  # noqa: N815


class WdUseFormattingFrom:
    # > Copy source formatting from the current item.
    # wdFormattingFromCurrent = 0
    # > Copy source formatting from the current selection.
    # wdFormattingFromSelected = 1
    # > Prompt the user for formatting to use.
    wdFormattingFromPrompt = 2  # noqa: N815


class WdWindowState:
    # > Normal.
    # wdWindowStateNormal = 0
    # > Maximized.
    wdWindowStateMaximize = 1  # noqa: N815
    # > Minimized.
    wdWindowStateMinimize = 2  # noqa: N815


class WdCompareDestination:
    # > Tracks the differences between the two files using tracked changes in the
    # > original document.
    # wdCompareDestinationOriginal = 0
    # > Tracks the differences between the two files using tracked changes in the
    # > revised document.
    wdCompareDestinationRevised = 1  # noqa: N815
    # > Creates a new file and tracks the differences between the original document
    # > and the revised document using tracked changes.
    wdCompareDestinationNew = 2  # noqa: N815


def parse_args():
    parser = argparse.ArgumentParser(
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )

    parser.add_argument(
        "BaseFileName", type=Path, help="$BASE (%O)", metavar="BFName",
    )
    parser.add_argument(
        "LocalFileName", type=Path, help="$LOCAL (%A)", metavar="LFName",
    )
    parser.add_argument(
        "RemoteFileName", type=Path, help="$REMOTE (%B)", metavar="RFName",
    )
    parser.add_argument(
        "ConflictMarkerSize",
        type=str,
        nargs="?",
        default=None,
        help="conflict-marker-size (%L)",
        metavar="CMS",
    )
    parser.add_argument(
        "MergeDest",
        type=Path,
        nargs="?",
        default=None,
        help="$MERGED (%P)",
        metavar="MDest",
    )

    logging_grp = parser.add_argument_group(title="Logging")
    logging_grp.add_argument(
        "-v",
        "--verbosity",
        default="INFO",
        type=str.upper,
        choices=list(LOG_LVLS.keys()),
        help="Console log level",
    )
    logging_grp.add_argument(
        "-l",
        "--log",
        default="DEBUG",
        type=str.upper,
        choices=list(LOG_LVLS.keys()),
        help="File log level",
    )

    return parser.parse_args()


def setup_root_logger() -> logging.Logger:
    logger = logging.getLogger()
    logger.setLevel(logging.NOTSET)

    module_loglevel_map = {
        "pywin32": logging.WARNING,
    }
    for module, loglevel in module_loglevel_map.items():
        logging.getLogger(module).setLevel(loglevel)

    log_filename = Path(f"{Path(__file__).stem}.log")
    log_roll = log_filename.is_file()
    file_handler = logging.handlers.RotatingFileHandler(
        filename=log_filename, mode="a", backupCount=9, encoding="utf-8",
    )
    if log_roll:
        file_handler.doRollover()
    file_handler.setLevel(LOG_LVLS[args.log])
    file_handler.setFormatter(
        logging.Formatter(
            fmt="[%(asctime)s.%(msecs)03d][%(name)s:%(levelname).4s] %(message)s",
            datefmt="%Y-%m-%dT%H:%M:%S",
        )
    )
    logger.addHandler(file_handler)

    console_handler = colorlog.StreamHandler()
    console_handler.setLevel(LOG_LVLS[args.verbosity])
    console_handler.setFormatter(
        colorlog.ColoredFormatter(
            fmt="[%(bold_blue)s%(name)s%(reset)s:%(log_color)s%(levelname).4s%(reset)s] %(msg_log_color)s%(message)s",
            log_colors={
                "DEBUG": "fg_bold_cyan",
                "INFO": "fg_bold_green",
                "WARNING": "fg_bold_yellow",
                "ERROR": "fg_bold_red",
                "CRITICAL": "fg_bold_purple",
            },
            secondary_log_colors={
                "msg": {
                    "DEBUG": "fg_white",
                    "INFO": "fg_bold_white",
                    "WARNING": "fg_yellow",
                    "ERROR": "fg_thin_red",
                    "CRITICAL": "fg_bold_red",
                },
            },
        )
    )
    logger.addHandler(console_handler)

    if False:
        # List all log levels with their respective coloring
        for log_lvl_name, log_lvl in LOG_LVLS.items():
            logger.log(log_lvl, "This is test message for %s", log_lvl_name)

    return logger


def ask_resolved() -> bool:
    ret = tk.messagebox.askyesno(
        title="Merge Complete?",
        message=f"Confirm conflict resolution for {args.MergeDest}?",
        icon="question",
    )
    # if not ret:
    #     tk.messagebox.showinfo(
    #         "Return", "You will now return to the application screen"
    #     )
    return ret


def init_files() -> int:
    global is_lfs
    is_lfs = False
    for alias in ["BASE", "LOCAL", "REMOTE"]:
        FileName = FileNameMap[alias].resolve(strict=True)  # noqa: N806
        FileNameMap[alias] = FileName
        FileNameExt = Path(str(FileName) + extension)  # noqa: N806
        # Write-Progress -Activity $activity -Status "Preparing $key" -PercentComplete $complete
        logging.debug("Processing '%s' ('%s')", alias, FileName)
        logging.debug("Checking if is Git LFS pointer...")
        cmd = f"git lfs pointer --check --file '{FileName}'"
        ret = subprocess.run(  # nosec
            shlex.split(cmd), stdin=sys.stdin, stdout=sys.stdout, stderr=sys.stderr,
        ).returncode
        if ret == 0:
            logging.debug("Is LFS pointer")
            is_lfs = True
            logging.info("Converting LFS pointer to blob...")
            cmd = (
                "cmd.exe /c 'type "
                + str(FileName)
                + " | git-lfs smudge > "
                + str(FileNameExt)
                + "'"
            )
            subprocess.run(  # nosec
                shlex.split(cmd), stdin=sys.stdin, stdout=sys.stdout, stderr=sys.stderr,
            )
            logging.debug("Done")
        elif ret == 1:
            logging.debug("Not LFS pointer")
            shutil.copy(FileName, FileNameExt)
        elif ret == 2:
            logging.critical("File not found")
            return 1
        else:
            logging.critical("Unknown return code '%s'!", ret)
            return 1

        FileMode = FileName.stat().st_mode  # noqa: N806
        logging.debug("File has %s mode", oct(FileMode))
        if FileMode == 0o100444:
            logging.debug("Removing read-only flag...")
            FileName.chmod(0o666)
            logging.debug("Done")
        # $complete += 30
    FileNameMap["MERGE"] = FileNameMap["LOCAL"]  # noqa: N806
    # $complete = 100
    # Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
    # sleep 1
    return 0


def dmfo_merge_wd() -> int:
    # . $PSScriptRoot\..\constants\const_wd.ps1

    # $activity = "Compiling 3-way-merge with MS Word. This may take a while... "
    # $complete = 0
    # Write-Progress -Activity $activity -Status "Initializing COM object" -PercentComplete $complete
    logging.debug("Initializing COM object...")
    try:
        COMObj = win32com.client.DispatchEx("Word.Application")  # noqa: N806
        COMObj.Visible = False
        # $complete += 20
        logging.debug("Done")
    except pywintypes.com_error as exc:
        logging.critical(
            "You must have Microsoft Word installed to perform this operation."
        )
        logging.debug("COM Error: '%s'", exc)
        return 1

    try:
        for alias in ["BASE", "LOCAL", "REMOTE"]:
            FileNameExt = Path(str(FileNameMap[alias]) + extension)  # noqa: N806
            # Write-Progress -Activity $activity -Status "Opening $key" -PercentComplete $complete
            logging.debug("Opening '%s' ('%s')", alias, FileNameExt)
            FileObjMap[alias] = COMObj.Documents.Open(  # noqa: N806
                FileName=str(FileNameExt),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            # $complete += 10
            logging.debug("Done")

        for alias in ["LOCAL", "REMOTE"]:
            # Write-Progress -Activity $activity -Status "Diffing $key vs BASE" -PercentComplete $complete
            logging.debug("Diffing '%s' vs 'BASE'", alias)
            FileObjMap[alias] = COMObj.CompareDocuments(
                OriginalDocument=FileObjMap["BASE"],
                RevisedDocument=FileObjMap[alias],
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
            logging.debug("Done")
            # $complete += 5

        logging.debug("Merging changes")
        # Write-Progress -Activity $activity -Status "Merging changes" -PercentComplete $complete
        FileObjMap["MERGE"] = COMObj.MergeDocuments(
            FileObjMap["LOCAL"],  # OriginalDocument
            FileObjMap["REMOTE"],  # RevisedDocument
            WdCompareDestination.wdCompareDestinationNew,  # Destination
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
        logging.debug("Done")
        # $complete += 10

        for alias in ["BASE", "LOCAL", "REMOTE"]:
            logging.debug("Closing '%s'", alias)
            # Write-Progress -Activity $activity -Status "Closing $key" -PercentComplete $complete
            FileObjMap[alias].Close(SaveChanges=WdSaveOptions.wdDoNotSaveChanges)
            FileObjMap.pop(alias)
            logging.debug("Done")
            # $complete += 5

        logging.debug("Saving 'MERGE'")
        # Write-Progress -Activity $activity -Status "Saving MERGE" -PercentComplete $complete
        FileObjMap["MERGE"].SaveAs(
            FileName=str(FileNameMap["MERGE"]) + extension, AddToRecentFiles=False,
        )
        logging.debug("Done")
        # $complete += 5

        logging.debug("Bringing to foreground")
        # Write-Progress -Activity $activity -Status "Bringing to foreground" -PercentComplete $complete
        COMObj.Visible = True
        COMObj.Activate()
        COMObj.WindowState = WdWindowState.wdWindowStateMinimize
        COMObj.WindowState = WdWindowState.wdWindowStateMaximize
        logging.debug("Done")
        # $complete = 100

        # Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
        # sleep 1
    except pywintypes.com_error as exc:
        logging.error("COM Error: '%s'", exc.args[1])
        logging.debug("COM Error: '%s'", exc)
        return 1

    logging.debug("Asking for merge confirmation...")
    is_resolved = ask_resolved()
    logging.debug("Reply: '%s'", is_resolved)

    # $activity = "Cleaning up... "
    # $complete = 0
    # Write-Progress -Activity $activity -Status "Checking merged file" -PercentComplete $complete
    try:
        logging.debug("Checking if 'MERGE' is still open...")
        COMObj.Documents.Item(str(FileNameMap["MERGE"]) + extension)
        logging.debug("'MERGE' is still open.")
    except pywintypes.com_error as exc:
        reopen = True
        if exc.args[0] == -2147352567:
            logging.debug("'MERGE' has been closed, reopening...")
        elif exc.args[0] == -2147023174:
            logging.debug("COMObj has been closed, reinitializing...")
            COMObj = win32com.client.DispatchEx("Word.Application")  # noqa: N806
            logging.debug("Done")
        else:
            logging.error("COM Error: '%s'", exc.args[1])
            logging.debug("COM Error: '%s'", exc)
    else:
        reopen = False
    finally:
        COMObj.Visible = False

        if reopen:
            logging.debug("Opening '%s' ('%s')", "MERGE", FileNameMap["MERGE"])
            FileObjMap["MERGE"] = COMObj.Documents.Open(  # noqa: N806
                FileName=str(FileNameMap["MERGE"]) + extension,
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            # $complete += 10
            logging.debug("Done")

    # FileObjMap["MERGE"].Activate()  # TODO: Needed?
    if FileObjMap["MERGE"].TrackRevisions:
        logging.warning("Warning: Track Changes is active. Please deactivate!")
        # Deactivate track changes?
    if is_resolved:
        if FileObjMap["MERGE"].Revisions.Count > 0:
            is_resolved = False  # TODO: Transfer to ps1 script
            logging.warning(
                "Warning: Unresolved revisions in the document. "
                + "Will exit as 'unresolved'."
            )
    # $complete += 70

    # Write-Progress -Activity $activity -Status "Closing COM Object" -PercentComplete $complete
    FileObjMap["MERGE"].Close()
    if COMObj.Documents.Count == 0:
        logging.debug("No more open documents in COMObj, closing...")
        COMObj.Quit()
        logging.debug("Done")
    # $complete += 10

    # Write-Progress -Activity $activity -Status "Copying merged file" -PercentComplete $complete
    if is_lfs:
        logging.info("Converting LFS blob to pointer...")
        cmd = (
            "cmd.exe /c 'type "
            + str(FileNameMap["MERGE"])
            + extension
            + " | git-lfs clean > "
            + str(FileNameMap["MERGE"])
            + "'"
        )
        subprocess.run(  # nosec
            shlex.split(cmd), stdin=sys.stdin, stdout=sys.stdout, stderr=sys.stderr,
        )
    else:
        logging.debug("Copying merged file...")
        shutil.copy(Path(str(FileNameMap["MERGE"]) + extension), FileNameMap["MERGE"])
    logging.debug("Done")
    # $complete += 10

    # Write-Progress -Activity $activity -Status "Removing aux files" -PercentComplete $complete
    for alias in ["BASE", "LOCAL", "REMOTE"]:
        FileNameExt = Path(str(FileNameMap[alias]) + extension)  # noqa: N806
        logging.debug("Removing aux file '%s' ('%s')...", alias, FileNameExt)
        FileNameExt.unlink()
        logging.debug("Done")
    # $complete = 100

    # Write-Progress -Activity $activity -Status "Done" -PercentComplete 100
    # sleep 1

    # Return as return code is_resolved -> True: 0, False: 1
    return int(not is_resolved)


args = parse_args()
root_logger = setup_root_logger()

extension = args.MergeDest.suffix
logging.debug("Merging '%s' file.", extension)

root = tk.Tk()
root.withdraw()


# $activity = "Preparing files... "
# $complete = 0
# Write-Progress -Activity $activity -Status "Initializing" -PercentComplete $complete
FileNameMap = {
    "BASE": args.BaseFileName,
    "LOCAL": args.LocalFileName,
    "REMOTE": args.RemoteFileName,
}
FileObjMap = {}
# $complete += 10

ret = init_files()
if ret:
    sys.exit(ret)

if extension in [".doc", ".docx"]:
    ret = dmfo_merge_wd()
else:
    logging.critical("DMFO-Merge does not know what to do with '%s' files.", extension)
    ret = 1
sys.exit(ret)
