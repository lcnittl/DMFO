#!/usr/bin/env python3
"""DMFO-Diff

Diff driver for Office documents. Opens files in MSO to compare them.
"""
import argparse
import logging
import logging.handlers
import shlex
import shutil
import subprocess  # nosec
import sys
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


class WdSaveOptions:
    # Prompt the user to save pending changes.
    # wdPromptToSaveChanges = -2
    # Save pending changes automatically without prompting the user.
    # wdSaveChanges = -1
    # Do not save pending changes.
    wdDoNotSaveChanges = 0  # noqa: N815


class WdUseFormattingFrom:
    # Copy source formatting from the current item.
    # wdFormattingFromCurrent = 0
    # Copy source formatting from the current selection.
    # wdFormattingFromSelected = 1
    # Prompt the user for formatting to use.
    wdFormattingFromPrompt = 2  # noqa: N815


class WdWindowState:
    # Normal.
    # wdWindowStateNormal = 0
    # Maximized.
    wdWindowStateMaximize = 1  # noqa: N815
    # Minimized.
    wdWindowStateMinimize = 2  # noqa: N815


class WdCompareDestination:
    # Tracks the differences between the two files using tracked changes in the
    # original document.
    # wdCompareDestinationOriginal = 0
    # Tracks the differences between the two files using tracked changes in the
    # revised document.
    wdCompareDestinationRevised = 1  # noqa: N815
    # Creates a new file and tracks the differences between the original document
    # and the revised document using tracked changes.
    wdCompareDestinationNew = 2  # noqa: N815


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )

    parser.add_argument(
        "DiffPath", type=Path, help="path", metavar="DPath",
    )
    parser.add_argument(
        "LocalFileName", type=Path, help="old-file ($LOCAL)", metavar="LFName",
    )
    parser.add_argument(
        "LocalFileHex", type=str, help="old-hex", metavar="LFHex",
    )
    parser.add_argument(
        "LocalFileMode", type=str, help="old-mode", metavar="LFMode",
    )
    parser.add_argument(
        "RemoteFileName", type=Path, help="new-file ($REMOTE)", metavar="RFName",
    )
    parser.add_argument(
        "RemoteFileHex", type=str, help="new-hex", metavar="RFHex",
    )
    parser.add_argument(
        "RemoteFileMode", type=str, help="new-mode", metavar="RFMode",
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


def init_files() -> None:
    for alias, FileName in FileNameMap.items():  # noqa: N806
        # Write-Progress -Activity $activity -Status "Preparing $key" -PercentComplete $complete
        FileName.resolve(strict=True)  # noqa: N806
        FileNameMap[alias] = FileName
        logging.debug("Processing '%s' ('%s')", alias, FileName)
        logging.debug("Checking if is Git LFS pointer...")
        cmd = f"git lfs pointer --check --file '{FileName}'"
        ret = subprocess.run(  # nosec
            shlex.split(cmd), stdin=sys.stdin, stdout=sys.stdout, stderr=sys.stderr,
        ).returncode
        if ret == 0:
            logging.debug("Is LFS pointer")
            # is_lfs = True
            logging.info("Converting LFS pointer to blob...")
            AuxFileName = Path(str(FileName) + "_")  # noqa: N806
            cmd = (
                "cmd.exe /c 'type "
                + str(FileName)
                + " | git-lfs smudge > "
                + str(AuxFileName)
                + "'"
            )
            subprocess.run(  # nosec
                shlex.split(cmd), stdin=sys.stdin, stdout=sys.stdout, stderr=sys.stderr,
            )
            shutil.move(AuxFileName, FileName)
            logging.debug("Done")
        elif ret == 1:
            logging.debug("Not LFS pointer")
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
        # $complete += 40
    # $complete = 100
    # Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
    # sleep 1


def dmfo_diff_wd() -> int:
    # . $PSScriptRoot\..\constants\const_wd.ps1

    # $activity = "Compiling diff of '$DiffPath' with MS Word. This may take a while... "
    # $complete = 0
    # Write-Progress -Activity $activity -Status "Initializing COM object" -PercentComplete $complete
    try:
        logging.debug("Initializing COM object...")
        COMObj = win32com.client.DispatchEx("Word.Application")  # noqa: N806
        COMObj.Visible = False
        # $complete += 20
        logging.debug("Done")
    except pywintypes.com_error:
        logging.critical(
            "You must have Microsoft Word installed to perform this operation."
        )
        return 1

    try:
        for alias, FileName in FileNameMap.items():  # noqa: N806
            logging.debug("Processing '%s' ('%s')", alias, FileName)
            # Write-Progress -Activity $activity -Status "Opening $key" -PercentComplete $complete
            logging.debug("Opening '%s'", alias)
            FileObjMap[alias] = COMObj.Documents.Open(  # noqa: N806
                FileName=str(FileName),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            # $complete += 15
            logging.debug("Done")

        logging.debug("Diffing REMOTE vs LOCAL")
        # Write-Progress -Activity $activity -Status "Diffing REMOTE vs LOCAL" -PercentComplete $complete
        FileObjMap["DIFF"] = COMObj.CompareDocuments(
            OriginalDocument=FileObjMap["LOCAL"],
            RevisedDocument=FileObjMap["REMOTE"],
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
        # $complete += 10

        for alias in ["LOCAL", "REMOTE"]:
            logging.debug("Closing '%s'", alias)
            # Write-Progress -Activity $activity -Status "Closing $key" -PercentComplete $complete
            FileObjMap[alias].Close(SaveChanges=WdSaveOptions.wdDoNotSaveChanges)
            FileObjMap.pop(alias)
            # $complete += 5

        logging.debug("Setting 'DIFF' to unsaved")
        # Write-Progress -Activity $activity -Status "Setting DIFF to unsaved" -PercentComplete $complete
        FileObjMap["DIFF"].Saved = 1
        # $complete += 10

        logging.debug("Bringing to foreground")
        # Write-Progress -Activity $activity -Status "Bringing to foreground" -PercentComplete $complete
        COMObj.Visible = True
        COMObj.Activate()
        COMObj.WindowState = WdWindowState.wdWindowStateMinimize
        COMObj.WindowState = WdWindowState.wdWindowStateMaximize
        # $complete = 100

        # Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
        # sleep 1
    except pywintypes.com_error as exc:
        logging.error("COM-Critical: '%s'", exc.args[1])
        return 1
    return 0


def dmfo_diff_pp() -> None:
    pass


args = parse_args()
root_logger = setup_root_logger()

extension = args.DiffPath.suffix
logging.debug("Diffing '%s' file.", extension)


# $activity = "Preparing files... "
# $complete = 0
# Write-Progress -Activity $activity -Status "Initializing" -PercentComplete $complete
FileNameMap = {
    "LOCAL": args.LocalFileName,
    "REMOTE": args.RemoteFileName,
}
FileObjMap = {}
# $complete += 20

ret = init_files()
if ret:
    sys.exit(ret)


if extension in [".doc", ".docx"]:
    ret = dmfo_diff_wd()
elif extension in [".ppt", ".pptx"]:
    ret = dmfo_diff_pp()
else:
    logging.critical("DMFO-Diff does not know what to do with '%s' files.", extension)
    ret = 1
sys.exit(ret)
