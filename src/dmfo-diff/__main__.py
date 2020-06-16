#!/usr/bin/env python3
"""DMFO-Diff

Diff driver for Office documents. Opens files in MSO to compare them.
"""
import argparse
import logging
import logging.handlers
import sys
from pathlib import Path

import pywin

LOG_LVLS = {
    # "NOTSET": logging.NOTSET,  # 0
    "DEBUG": logging.DEBUG,  # 10
    "INFO": logging.INFO,  # 20
    "WARNING": logging.WARNING,  # 30
    "ERROR": logging.ERROR,  # 40
    "CRITICAL": logging.CRITICAL,  # 50
}


def parse_args():
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


args = parse_args()
root_logger = setup_root_logger()

ext = args.DiffPath.suffix
logging.DEBUG("Diffing '%s' file.", ext)


FileNames = {
    "LOCAL": args.LocalFileName,
    "REMOTE": args.RemoteFileName,
}

for name, FilePath in FileNames.items():
    print(name, FilePath)
"""
$extension = [System.IO.Path]::GetExtension($DiffPath)


$activity = "Preparing files... "
$complete = 0
Write-Progress -Activity $activity -Status "Initializing" -PercentComplete $complete
$FileNames = @{
    LOCAL = $LocalFileName;
    REMOTE = $RemoteFileName
}
$Files = @{}
$complete += 20

foreach ($key in @($FileNames.Keys)) {
    Write-Progress -Activity $activity -Status "Preparing $key" -PercentComplete $complete
    $FileName = (Resolve-Path $FileNames[$key]).Path
    $FileNames[$key] = $FileName
    git lfs pointer --check --file $FileName
    if ($?) {
        $LFS = $true
        Write-Host Converting LFS pointer to blob.
        cmd.exe /c "type $($FileName) | git-lfs smudge > $($FileName + "_")"
        mv -Force $($FileName + "_") $FileName
    }
    $File = Get-ChildItem $FileNameExt
    if ($File.IsReadOnly) {
        $File.IsReadOnly = $false
    }
    $complete += 40
}
$complete = 100

Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
sleep 1


if (@(".doc", ".docx") -contains $extension) {
    . $PSScriptRoot\dmfo-diff\dmfo-diff_wd.ps1
} elseif (@(".ppt", ".pptx") -contains $extension) {
    . $PSScriptRoot\dmfo-diff\dmfo-diff_pp.ps1
} else {
    Write-Host "DMFO-Diff does not know what to do with '$extension' files."
    exit(1)
}
exit($LastExitCode)
"""
