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
from typing import Dict, List, Tuple

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
        "pywin32": logging.WARNING,
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


def preproc_files(filename_map: Dict[str, Path], target_extension: str) -> int:
    global has_extension, is_lfs
    # TODO: progressbar
    has_extension = False  # TODO: make has_extension a per-file-attribute
    for alias in filename_map.keys():
        filename = filename_map[alias].resolve(strict=True)
        filename_map[alias] = filename
        logging.debug("Processing '%s' ('%s')", alias, filename)
        # TODO: progressbar

        if filename.suffix == target_extension:
            has_extension |= True
            fices = ["_", ""]
        else:
            fices = ["", target_extension]

        logging.debug("Checking if is Git LFS pointer...")
        cmd = f"git lfs pointer --check --file '{filename}'"
        ret = subprocess.run(  # nosec
            shlex.split(cmd), stdout=sys.stdout, stderr=sys.stderr,
        ).returncode
        if ret == 0:
            logging.debug("Yes, is LFS pointer")
            is_lfs = True
            logging.info("Converting LFS pointer to blob...")
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
            logging.debug("Done")
        elif ret == 1:
            logging.debug("No, is not LFS pointer")
            if not has_extension:
                shutil.copy(filename, aux_filename)
        elif ret == 2:
            logging.critical("File not found")
            return 4
        else:
            logging.critical("Unknown return code '%s'!", ret)
            return 5

        if not has_extension:
            filename = aux_filename

        filemode = filename.stat().st_mode
        logging.debug("File has %s mode", oct(filemode))
        if filemode == 0o100444:
            logging.debug("Removing read-only flag...")
            filename.chmod(0o666)
            logging.debug("Done")
        # TODO: progressbar
    # TODO: progressbar
    return 0


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
    logging.debug("Initializing COM object...")
    # TODO: progressbar
    try:
        com_obj = win32com.client.DispatchEx(f"{app_name}.Application")
        com_obj.Visible = False
        logging.debug("Done")
        # TODO: progressbar
    except pywintypes.com_error as exc:
        logging.critical(
            "You must have Microsoft %s installed to perform this operation.", app_name
        )
        logging.debug("COM Error: '%s'", exc)
        return (3, None)
    return (0, com_obj)


def dmfo_diff(filename_map: Dict[str, Path], fileobj_map: Dict[str, object]) -> int:
    if extension in [".doc", ".docx"]:
        ret = dmfo_diff_wd(filename_map=filename_map, fileobj_map=fileobj_map)
    elif extension in [".ppt", ".pptx"]:
        ret = dmfo_diff_pp(filename_map=filename_map, fileobj_map=fileobj_map)
    else:
        logging.critical(
            "DMFO-Diff does not know what to do with '%s' files.", extension
        )
        ret = 2
    return ret


def dmfo_merge(filename_map: Dict[str, Path], fileobj_map: Dict[str, object]) -> int:
    if extension in [".doc", ".docx"]:
        ret = dmfo_merge_wd(filename_map=FileNameMap, fileobj_map=FileObjMap)
    else:
        logging.critical(
            "DMFO-Merge does not know what to do with '%s' files.", extension
        )
        ret = 2
    return ret


def dmfo_diff_wd(filename_map: Dict[str, Path], fileobj_map: Dict[str, object]) -> int:
    ret, COMObj = init_com_obj("Word")  # noqa: N806
    if ret:
        return ret

    try:
        for alias in ["LOCAL", "REMOTE"]:
            filename = filename_map[alias]  # noqa: N806
            logging.debug("Opening '%s' ('%s')", alias, filename)
            # TODO: progressbar
            fileobj_map[alias] = COMObj.Documents.Open(  # noqa: N806
                FileName=str(filename),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            logging.debug("Done")
            # TODO: progressbar

        logging.debug("Diffing 'REMOTE' vs 'LOCAL'")
        # TODO: progressbar
        fileobj_map["DIFF"] = COMObj.CompareDocuments(
            OriginalDocument=fileobj_map["LOCAL"],
            RevisedDocument=fileobj_map["REMOTE"],
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
        logging.debug("Done")
        # TODO: progressbar

        for alias in ["LOCAL", "REMOTE"]:
            logging.debug("Closing '%s'", alias)
            # TODO: progressbar
            fileobj_map[alias].Close(SaveChanges=WdSaveOptions.wdDoNotSaveChanges)
            fileobj_map.pop(alias)
            logging.debug("Done")
            # TODO: progressbar

        logging.debug("Setting 'DIFF' to unsaved")
        # TODO: progressbar
        fileobj_map["DIFF"].Saved = 1
        logging.debug("Done")
        # TODO: progressbar

        logging.debug("Bringing to foreground")
        # TODO: progressbar
        COMObj.Visible = True
        COMObj.Activate()
        COMObj.WindowState = WdWindowState.wdWindowStateMinimize
        COMObj.WindowState = WdWindowState.wdWindowStateMaximize
        logging.debug("Done")
        # TODO: progressbar
    except pywintypes.com_error as exc:
        logging.error("COM Error: '%s'", exc.args[1])
        logging.debug("COM Error: '%s'", exc)
        return 6
    return 0


def dmfo_merge_wd(filename_map: Dict[str, Path], fileobj_map: Dict[str, object]) -> int:
    ret, COMObj = init_com_obj("Word")  # noqa: N806
    if ret:
        return ret

    try:
        for alias in ["BASE", "LOCAL", "REMOTE"]:
            aux_filename = Path(str(filename_map[alias]) + extension)  # noqa: N806
            # TODO: progressbar
            logging.debug("Opening '%s' ('%s')", alias, aux_filename)
            fileobj_map[alias] = COMObj.Documents.Open(  # noqa: N806
                FileName=str(aux_filename),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            logging.debug("Done")
            # TODO: progressbar

        for alias in ["LOCAL", "REMOTE"]:
            logging.debug("Diffing '%s' vs 'BASE'", alias)
            # TODO: progressbar
            fileobj_map[alias] = COMObj.CompareDocuments(
                OriginalDocument=fileobj_map["BASE"],
                RevisedDocument=fileobj_map[alias],
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
            # TODO: progressbar

        logging.debug("Merging changes")
        # TODO: progressbar
        fileobj_map["MERGE"] = COMObj.MergeDocuments(
            OriginalDocument=fileobj_map["LOCAL"],
            RevisedDocument=fileobj_map["REMOTE"],
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
        logging.debug("Done")
        # TODO: progressbar

        for alias in ["BASE", "LOCAL", "REMOTE"]:
            logging.debug("Closing '%s'", alias)
            # TODO: progressbar
            fileobj_map[alias].Close(SaveChanges=WdSaveOptions.wdDoNotSaveChanges)
            fileobj_map.pop(alias)
            logging.debug("Done")
            # TODO: progressbar

        filename_map["MERGE"] = filename_map["LOCAL"]

        logging.debug("Saving 'MERGE'")
        # TODO: progressbar
        fileobj_map["MERGE"].SaveAs(
            FileName=str(filename_map["MERGE"]) + extension, AddToRecentFiles=False,
        )
        logging.debug("Done")
        # TODO: progressbar

        logging.debug("Bringing to foreground")
        # TODO: progressbar
        COMObj.Visible = True
        COMObj.Activate()
        COMObj.WindowState = WdWindowState.wdWindowStateMinimize
        COMObj.WindowState = WdWindowState.wdWindowStateMaximize
        logging.debug("Done")
        # TODO: progressbar
    except pywintypes.com_error as exc:
        logging.error("COM Error: '%s'", exc.args[1])
        logging.debug("COM Error: '%s'", exc)
        return 6

    logging.debug("Asking for merge confirmation...")
    is_resolved = ask_resolved()
    logging.debug("Reply: '%s'", is_resolved)
    # TODO: progressbar

    try:
        logging.debug("Checking if 'MERGE' is still open...")
        COMObj.Documents.Item(str(filename_map["MERGE"]) + extension)
        logging.debug("'MERGE' is still open.")
    except pywintypes.com_error as exc:
        reopen = True
        if exc.args[0] in [-2147352567]:
            logging.debug("'MERGE' has been closed, reopening...")
        elif exc.args[0] in [-2147023174, -2147023179]:
            logging.debug("COMObj has been closed, reinitializing...")
            ret, COMObj = init_com_obj("Word")  # noqa: N806
            if ret:
                return ret
        else:
            logging.error("COM Error: '%s'", exc.args[1])
            logging.debug("COM Error: '%s'", exc)
    else:
        reopen = False

    COMObj.Visible = False

    if reopen:
        logging.debug("Opening '%s' ('%s')", "MERGE", filename_map["MERGE"])
        fileobj_map["MERGE"] = COMObj.Documents.Open(  # noqa: N806
            FileName=str(filename_map["MERGE"]) + extension,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )
        logging.debug("Done")
        # TODO: progressbar

    if fileobj_map["MERGE"].TrackRevisions:
        logging.warning("Warning: Track Changes is active. Please deactivate!")
        # Deactivate track changes?
    if is_resolved:
        if fileobj_map["MERGE"].Revisions.Count > 0:
            is_resolved = False  # TODO: Transfer to ps1 script
            logging.warning(
                "Warning: Unresolved revisions in the document. "
                + "Will exit as 'unresolved'."
            )
    # TODO: progressbar

    # TODO: progressbar
    fileobj_map["MERGE"].Close()
    if COMObj.Documents.Count == 0:
        logging.debug("No more open documents in COMObj, closing...")
        COMObj.Quit()
        logging.debug("Done")
    # TODO: progressbar

    # TODO: progressbar
    if is_lfs:
        logging.info("Converting LFS blob to pointer...")
        cmd = (
            "cmd.exe /c 'type "
            + (str(filename_map["MERGE"]) + extension)
            + " | git-lfs clean > "
            + str(filename_map["MERGE"])
            + "'"
        )
        subprocess.run(  # nosec
            shlex.split(cmd), stdout=sys.stdout, stderr=sys.stderr,
        )
    else:
        logging.debug("Copying merged file...")
        shutil.copy(Path(str(filename_map["MERGE"]) + extension), filename_map["MERGE"])
    logging.debug("Done")
    # TODO: progressbar

    filename_map.pop("MERGE")

    # Return as return code: is_resolved -> True: 0, False: 1
    return int(not is_resolved)


args = parse_args()
TempDir = Path(tempfile.mkdtemp(prefix=f"dmfo_{args.mode}_"))
root_logger = setup_root_logger(path=TempDir)

FileNameMap = {
    "LOCAL": args.LocalFileName,
    "REMOTE": args.RemoteFileName,
}
FileObjMap = {}

# extension = args.TargetPath.suffix
if args.mode == "diff":
    extension = args.DiffPath.suffix
    logging.debug("Diffing '%s' file.", extension)
elif args.mode == "merge":
    extension = args.MergeDest.suffix
    logging.debug("Merging '%s' file.", extension)

    FileNameMap.update({"BASE": args.BaseFileName})

ret = preproc_files(filename_map=FileNameMap, target_extension=extension)
if ret:
    sys.exit(ret)

if args.mode == "diff":
    ret = dmfo_diff(filename_map=FileNameMap, fileobj_map=FileObjMap)
elif args.mode == "merge":
    ret = dmfo_merge(filename_map=FileNameMap, fileobj_map=FileObjMap)

if not has_extension:
    # TODO: progressbar
    for alias in FileNameMap.keys():
        FileNameExt = Path(str(FileNameMap[alias]) + extension)  # noqa: N806
        logging.debug("Removing aux file '%s' ('%s')...", alias, FileNameExt)
        FileNameExt.unlink()
        logging.debug("Done")
    # TODO: progressbar

if ret > 1:
    logging.critical(
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
