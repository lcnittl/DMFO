#!/usr/bin/env python3
"""DMFO

Diff and merge driver for Office documents. Opens files in MSO to compare them.
"""
import argparse
import logging
import logging.handlers
import shlex
import shutil
import subprocess  # nosec
import sys
import tempfile
import tkinter as tk
import tkinter.messagebox
from pathlib import Path
from typing import Dict, List, Tuple, Union

import colorlog
import pywintypes  # win32com.client.pywintypes
import win32com.client
import win32con
import win32ui
from constants.mso.mso import MsoTriState
from constants.mso.pp import PpWindowState
from constants.mso.wd import (
    WdCompareDestination,
    WdSaveOptions,
    WdUseFormattingFrom,
    WdWindowState,
)

logger = logging.getLogger(__name__)


DEFAULT_LOG_PATH = Path(".")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )

    logging_grp = parser.add_argument_group(title="Logging")
    logging_grp.add_argument(
        "-v",
        "--verbosity",
        default="INFO",
        type=str.upper,
        choices=list(logging._nameToLevel.keys()),
        help="Console log level",
    )
    logging_grp.add_argument(
        "-l",
        "--log",
        default="DEBUG",
        type=str.upper,
        choices=list(logging._nameToLevel.keys()),
        help="File log level",
    )

    subparser = parser.add_subparsers(dest="mode", help="Choose driver to run:")

    diff_parser = subparser.add_parser("diff", help="diff driver")
    diff_parser.add_argument(
        "DiffPath",
        # dest="TargetPath",
        type=Path,
        help="path",
        metavar="DPath",
    )
    diff_parser.add_argument(
        "LocalFileName", type=Path, help="old-file ($LOCAL)", metavar="LFName",
    )
    diff_parser.add_argument(
        "LocalFileHex", type=str, help="old-hex", metavar="LFHex",
    )
    diff_parser.add_argument(
        "LocalFileMode", type=str, help="old-mode", metavar="LFMode",
    )
    diff_parser.add_argument(
        "RemoteFileName", type=Path, help="new-file ($REMOTE)", metavar="RFName",
    )
    diff_parser.add_argument(
        "RemoteFileHex", type=str, help="new-hex", metavar="RFHex",
    )
    diff_parser.add_argument(
        "RemoteFileMode", type=str, help="new-mode", metavar="RFMode",
    )

    merge_parser = subparser.add_parser("merge", help="merge driver")
    merge_parser.add_argument(
        "BaseFileName", type=Path, help="$BASE (%%O)", metavar="BFName",
    )
    merge_parser.add_argument(
        "LocalFileName", type=Path, help="$LOCAL (%%A)", metavar="LFName",
    )
    merge_parser.add_argument(
        "RemoteFileName", type=Path, help="$REMOTE (%%B)", metavar="RFName",
    )
    merge_parser.add_argument(
        "ConflictMarkerSize",
        type=str,
        nargs="?",
        default=None,
        help="conflict-marker-size (%%L)",
        metavar="CMS",
    )
    merge_parser.add_argument(
        "MergeDest",
        # dest="TargetPath",
        type=Path,
        nargs="?",
        default=None,
        help="$MERGED (%%P)",
        metavar="MDest",
    )

    return parser.parse_args()


def setup_root_logger(path: Path = DEFAULT_LOG_PATH) -> logging.Logger:
    global log_filename

    logger = logging.getLogger()
    logger.setLevel(logging.NOTSET)

    """
    module_loglevel_map = {
        "pywin32": logger.WARNING,
    }
    for module, loglevel in module_loglevel_map.items():
        logging.getLogger(module).setLevel(loglevel)
    """

    log_filename = Path(f"{Path(path) / Path(__file__).stem}.log")
    log_roll = log_filename.is_file()
    file_handler = logging.handlers.RotatingFileHandler(
        filename=log_filename, mode="a", backupCount=9, encoding="utf-8",
    )
    if log_roll:
        file_handler.doRollover()
    file_handler.setLevel(args.log)
    file_handler.setFormatter(
        logging.Formatter(
            fmt="[%(asctime)s.%(msecs)03d][%(name)s:%(levelname).4s] %(message)s",
            datefmt="%Y-%m-%dT%H:%M:%S",
        )
    )
    logger.addHandler(file_handler)

    console_handler = colorlog.StreamHandler()
    console_handler.setLevel(args.verbosity)
    console_handler.setFormatter(
        colorlog.ColoredFormatter(
            fmt="[%(bold_blue)s%(name)s%(reset)s:%(log_color)s%(levelname).4s%(reset)s] %(msg_log_color)s%(message)s",
            log_colors={
                "DEBUG": "fg_bold_cyan",
                "INFO": "fg_bold_green",
                "WARNING": "fg_bold_yellow",
                "ERROR": "fg_bold_red",
                "CRITICAL": "fg_thin_red",
            },
            secondary_log_colors={
                "msg": {
                    "DEBUG": "fg_white",
                    "INFO": "fg_bold_white",
                    "WARNING": "fg_bold_yellow",
                    "ERROR": "fg_bold_red",
                    "CRITICAL": "fg_thin_red",
                },
            },
        )
    )
    logger.addHandler(console_handler)

    if False:
        # List all log levels with their respective coloring
        for log_lvl_name, log_lvl in logging._nameToLevel.items():
            logger.log(log_lvl, "This is test message for %s", log_lvl_name)

    return logger


def preproc_files(file_map: Dict[str, Dict[str, Union[bool, object, Path]]]) -> int:
    # TODO: progressbar
    for alias in file_map.keys():
        filename = file_map[alias]["Name"].resolve(strict=True)
        file_map[alias]["Name"] = filename
        logger.debug("Processing '%s' ('%s')", alias, filename)
        # TODO: progressbar

        if filename.suffix == extension:
            has_extension = True
            fices = ["_", ""]
        else:
            has_extension = False
            fices = ["", extension]

        logger.debug("Checking if is Git LFS pointer...")
        cmd = f"git lfs pointer --check --file '{filename}'"
        ret = subprocess.run(  # nosec
            shlex.split(cmd), stdout=sys.stdout, stderr=sys.stderr,
        ).returncode
        if ret == 0:
            logger.debug("Yes, is LFS pointer")
            is_lfs = True
            logger.info("Converting LFS pointer to blob...")
            aux_filename = Path(str(filename).join(fices))
            cmd = (
                "cmd.exe /c 'type "
                + str(filename)
                + " | git-lfs smudge > "
                + str(aux_filename)
                + "'"
            )
            subprocess.run(  # nosec
                shlex.split(cmd), stdout=sys.stdout, stderr=sys.stderr,
            )
            if has_extension:
                shutil.move(aux_filename, filename)
            logger.debug("Done")
        elif ret == 1:
            logger.debug("No, is not LFS pointer")
            is_lfs = False
            if not has_extension:
                shutil.copy(filename, aux_filename)
        elif ret == 2:
            logger.critical("File not found")
            return 4
        else:
            logger.critical("Unknown return code '%s'!", ret)
            return 5
        file_map[alias]["isLFS"] = is_lfs

        if not has_extension:
            filename = aux_filename
            file_map[alias]["NameExt"] = aux_filename

        filemode = filename.stat().st_mode
        logger.debug("File has %s mode", oct(filemode))
        if filemode == 0o100444:
            logger.debug("Removing read-only flag...")
            filename.chmod(0o666)
            logger.debug("Done")
        # TODO: progressbar
    # TODO: progressbar
    return 0


def postproc_files(file_map: Dict[str, Dict[str, Union[bool, object, Path]]]) -> None:
    # TODO: progressbar
    if args.mode == "merge":
        # Convert to LFS pointer only if one of the decendants is managed by LFS
        if any(file_map[alias]["isLFS"] for alias in ["LOCAL", "REMOTE"]):
            logger.info("Converting LFS blob to pointer...")
            cmd = (
                "cmd.exe /c 'type "
                + str(file_map["LOCAL"]["NameExt"])
                + " | git-lfs clean > "
                + str(file_map["LOCAL"]["Name"])
                + "'"
            )
            subprocess.run(  # nosec
                shlex.split(cmd), stdout=sys.stdout, stderr=sys.stderr,
            )
        else:
            logger.debug("Copying merged file...")
            shutil.copy(file_map["LOCAL"]["NameExt"], file_map["LOCAL"]["Name"])
        logger.debug("Done")
        # TODO: progressbar

    # Delete generated aux files
    for alias in filter(lambda alias: "NameExt" in file_map[alias], file_map.keys()):
        filename = file_map[alias]["NameExt"]
        logger.debug("Removing aux file '%s' ('%s')...", alias, filename)
        filename.unlink()
        logger.debug("Done")
    # TODO: progressbar


def ask_resolved() -> bool:
    ret = win32ui.MessageBox(
        f"Confirm conflict resolution for {22}?",
        "Merge Complete?",
        win32con.MB_YESNO | win32con.MB_ICONQUESTION,
    )

    # win32ui.MessageBox returns 6 for "Yes" and 7 for "No"
    return ret == 6


def dmfo_diff_pp() -> int:
    pass


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


def dmfo_diff(file_map: Dict[str, Dict[str, Union[bool, object, Path]]]) -> int:
    file_map["DIFF"] = {}

    if extension in [".doc", ".docx"]:
        ret = dmfo_diff_wd(file_map=file_map)
    elif extension in [".ppt", ".pptx"]:
        ret = dmfo_diff_pp(file_map=file_map)
    else:
        logger.critical(
            "DMFO-Diff does not know what to do with '%s' files.", extension
        )
        ret = 2
    return ret


def dmfo_merge(file_map: Dict[str, Dict[str, Union[bool, object, Path]]]) -> int:
    file_map["MERGE"] = {}

    if extension in [".doc", ".docx"]:
        ret = dmfo_merge_wd(file_map=file_map)
    else:
        logger.critical(
            "DMFO-Merge does not know what to do with '%s' files.", extension
        )
        ret = 2
    return ret


def dmfo_diff_wd(file_map: Dict[str, Dict[str, Union[bool, object, Path]]]) -> int:
    ret, COMObj = init_com_obj("Word")  # noqa: N806
    if ret:
        return ret

    try:
        for alias in ["LOCAL", "REMOTE"]:
            filename = (
                file_map[alias]["Name"]
                if "NameExt" not in file_map[alias]
                else file_map[alias]["NameExt"]
            )
            logger.debug("Opening '%s' ('%s')", alias, filename)
            # TODO: progressbar
            file_map[alias]["Object"] = COMObj.Documents.Open(
                FileName=str(filename),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            logger.debug("Done")
            # TODO: progressbar

        logger.debug("Diffing 'REMOTE' vs 'LOCAL'")
        # TODO: progressbar
        file_map["DIFF"]["Object"] = COMObj.CompareDocuments(
            OriginalDocument=file_map["LOCAL"]["Object"],
            RevisedDocument=file_map["REMOTE"]["Object"],
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
            file_map[alias]["Object"].Close(
                SaveChanges=WdSaveOptions.wdDoNotSaveChanges
            )
            file_map[alias].pop("Object")
            logger.debug("Done")
            # TODO: progressbar

        logger.debug("Setting 'DIFF' to unsaved")
        # TODO: progressbar
        file_map["DIFF"]["Object"].Saved = 1
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


def dmfo_merge_wd(file_map: Dict[str, Dict[str, Union[bool, object, Path]]]) -> int:
    ret, COMObj = init_com_obj("Word")  # noqa: N806
    if ret:
        return ret

    try:
        for alias in ["BASE", "LOCAL", "REMOTE"]:
            aux_filename = Path(str(file_map[alias]["Name"]) + extension)  # noqa: N806
            # TODO: progressbar
            logger.debug("Opening '%s' ('%s')", alias, aux_filename)
            file_map[alias]["Object"] = COMObj.Documents.Open(  # noqa: N806
                FileName=str(aux_filename),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            logger.debug("Done")
            # TODO: progressbar

        for alias in ["LOCAL", "REMOTE"]:
            logger.debug("Diffing '%s' vs 'BASE'", alias)
            # TODO: progressbar
            file_map[alias]["Object"] = COMObj.CompareDocuments(
                OriginalDocument=file_map["BASE"]["Object"],
                RevisedDocument=file_map[alias]["Object"],
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
        file_map["MERGE"]["Object"] = COMObj.MergeDocuments(
            OriginalDocument=file_map["LOCAL"]["Object"],
            RevisedDocument=file_map["REMOTE"]["Object"],
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
        # TODO: progressbar

        for alias in ["BASE", "LOCAL", "REMOTE"]:
            logger.debug("Closing '%s'", alias)
            # TODO: progressbar
            file_map[alias]["Object"].Close(
                SaveChanges=WdSaveOptions.wdDoNotSaveChanges
            )
            file_map[alias].pop("Object")
            logger.debug("Done")
            # TODO: progressbar

        logger.debug("Saving 'MERGE'")
        # TODO: progressbar
        file_map["MERGE"]["Object"].SaveAs(
            FileName=str(file_map["LOCAL"]["Name"]) + extension, AddToRecentFiles=False,
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
        COMObj.Documents.Item(str(file_map["LOCAL"]["Name"]) + extension)
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
        logger.debug("Opening '%s' ('%s')", "MERGE", file_map["LOCAL"]["Name"])
        file_map["MERGE"]["Object"] = COMObj.Documents.Open(  # noqa: N806
            FileName=str(file_map["LOCAL"]["Name"]) + extension,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )
        logger.debug("Done")
        # TODO: progressbar

    if file_map["MERGE"]["Object"].TrackRevisions:
        logger.warning("Warning: Track Changes is active. Please deactivate!")
        # Deactivate track changes?
    if is_resolved:
        if file_map["MERGE"]["Object"].Revisions.Count > 0:
            is_resolved = False  # TODO: Transfer to ps1 script
            logger.warning(
                "Warning: Unresolved revisions in the document. "
                + "Will exit as 'unresolved'."
            )
    # TODO: progressbar

    # TODO: progressbar
    file_map["MERGE"]["Object"].Close()
    if COMObj.Documents.Count == 0:
        logger.debug("No more open documents in COMObj, closing...")
        COMObj.Quit()
        logger.debug("Done")
    file_map["MERGE"].pop("Object")
    # TODO: progressbar

    # Return as return code: is_resolved -> True: 0, False: 1
    return int(not is_resolved)


args = parse_args()
TempDir = Path(tempfile.mkdtemp(prefix=f"dmfo_{args.mode}_"))
root_logger = setup_root_logger(path=TempDir)

FileMap = {
    "LOCAL": {"Name": args.LocalFileName},
    "REMOTE": {"Name": args.RemoteFileName},
}

# extension = args.TargetPath.suffix
if args.mode == "diff":
    extension = args.DiffPath.suffix
    logger.debug("Diffing '%s' file.", extension)
elif args.mode == "merge":
    extension = args.MergeDest.suffix
    logger.debug("Merging '%s' file.", extension)

    FileMap.update({"BASE": {"Name": args.BaseFileName}})

ret = preproc_files(file_map=FileMap)
if ret:
    sys.exit(ret)

if args.mode == "diff":
    ret = dmfo_diff(file_map=FileMap)
elif args.mode == "merge":
    ret = dmfo_merge(file_map=FileMap)

postproc_files(file_map=FileMap)

if ret > 1:
    logger.critical(
        "DMFO %s exited with return code %s, check log for details (%s)",
        args.mode,
        ret,
        log_filename.resolve(),
    )
sys.exit(ret)

"""Return codes:
1: Not auto-merged
2: Unknown file extension
3: COM Application (Word, PowerPoint) not installed
4: File not found
5: Unknown git lfs pointer --check return code
6: Unexpected pywin32 com_error
"""
