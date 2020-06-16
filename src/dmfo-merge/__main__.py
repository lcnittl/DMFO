#!/usr/bin/env python3
"""DMFO-Merge

Merge driver for Office documents. Opens files in MSO to merge them.
"""
import argparse
import logging
import logging.handlers
import sys

import pywin32

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
        "BaseFileName", type=str, help="$BASE", metavar="BFName",
    )
    parser.add_argument(
        "LocalFileName", type=str, help="$LOCAL", metavar="LFName",
    )
    parser.add_argument(
        "RemoteFileName", type=str, help="$REMOTE", metavar="RFName",
    )
    parser.add_argument(
        "ConflictMarkerSize",
        type=str,
        nargs="?",
        default=None,
        help="conflict-marker-size",
        metavar="CMS",
    )
    parser.add_argument(
        "MergeDest", type=str, nargs="?", default=None, help="$MERGED", metavar="MDest",
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


args = parse_args()


def ExitApplication():
    MsgBox = tk.messagebox.askquestion(
        "Exit Application",
        "Are you sure you want to exit the application",
        icon="warning",
    )
    if MsgBox == "yes":
        root.destroy()
    else:
        tk.messagebox.showinfo(
            "Return", "You will now return to the application screen"
        )


"""
param(
    [parameter(Mandatory=$true)][string] $BaseFileName,  # $BASE
    [parameter(Mandatory=$true)][string] $LocalFileName,  # $LOCAL
    [parameter(Mandatory=$true)][string] $RemoteFileName,  # $REMOTE
    [parameter(Mandatory=$false)][string] $ConflictMarkerSize,  # conflict-marker-size
    [parameter(Mandatory=$false)][string] $MergeDest  # $MERGED
)
if ($PSVersionTable.PSVersion.Major -lt 6) {
    $PSDefaultParameterValues["Out-File:Encoding"] = "utf8"
}
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms

$extension = [System.IO.Path]::GetExtension($MergeDest)


$activity = "Preparing files... "
$complete = 0
Write-Progress -Activity $activity -Status "Initializing" -PercentComplete $complete
$FileNames = @{
    BASE = $BaseFileName;
    LOCAL = $LocalFileName;
    REMOTE = $RemoteFileName
}
$FileNamesExt = @{}
$Files = @{}
$complete += 10

foreach ($key in @($FileNames.Keys)) {
    Write-Progress -Activity $activity -Status "Preparing $key" -PercentComplete $complete
    $FileName = (Resolve-Path $FileNames[$key]).Path
    $FileNames[$key] = $FileName
    $FileNameExt = $FileName + $extension
    $FileNamesExt += @{$key = "$FileNameExt"}
    git lfs pointer --check --file $FileName
    if ($?) {
        $LFS = $true
        Write-Host Converting LFS pointer to blob.
        cmd.exe /c "type $($FileName) | git-lfs smudge > $($FileNameExt)"
    } else {
        cp $FileName $FileNameExt
    }
    $File = Get-ChildItem $FileNameExt
    if ($File.IsReadOnly) {
        $File.IsReadOnly = $false
    }
    $complete += 30
}
$complete = 100

Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
sleep 1


if (@(".doc", ".docx") -contains $extension) {
    . $PSScriptRoot\dmfo-merge\dmfo-merge_wd.ps1
} else {
    Write-Host "DMFO-Merge does not know what to do with '$extension' files."
    exit(1)
}
exit($LastExitCode)
"""
