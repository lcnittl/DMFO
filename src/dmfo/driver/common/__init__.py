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
from typing import Dict, Tuple

import pywintypes  # win32com.client.pywintypes
import win32com.client
import win32con
import win32ui
from classes import VCSFileData

logger = logging.getLogger(__name__)


def preproc_files(filedata_map: Dict[str, object]) -> int:
    # TODO: progressbar
    for alias in filedata_map.keys():
        filename = filedata_map[alias].name.resolve(strict=True)
        filedata_map[alias].name = filename
        logger.debug("Processing '%s' ('%s')", alias, filename)
        # TODO: progressbar

        extension = VCSFileData.target_ext
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
        filedata_map[alias].is_lfs = is_lfs

        if not has_extension:
            filename = aux_filename

        filemode = filename.stat().st_mode
        logger.debug("File has %s mode", oct(filemode))
        if filemode == 0o100444:
            logger.debug("Removing read-only flag...")
            filename.chmod(0o666)
            logger.debug("Done")
        # TODO: progressbar
    # TODO: progressbar
    return 0


def postproc_files(filedata_map: Dict[str, object], mode: str) -> None:
    # TODO: progressbar
    if mode == "merge":
        # Convert to LFS pointer only if one of the decendants is managed by LFS
        if any(filedata_map[alias].is_lfs for alias in ["LOCAL", "REMOTE"]):
            logger.info("Converting LFS blob to pointer...")
            cmd = (
                "cmd.exe /c 'type "
                + str(filedata_map["LOCAL"].get_name())
                + " | git-lfs clean > "
                + str(filedata_map["LOCAL"].name)
                + "'"
            )
            subprocess.run(  # nosec
                shlex.split(cmd), stdout=sys.stdout, stderr=sys.stderr,
            )
        else:
            logger.debug("Copying merged file...")
            shutil.copy(filedata_map["LOCAL"].get_name(), filedata_map["LOCAL"].name)
        logger.debug("Done")
        # TODO: progressbar

    # Delete generated aux files
    for alias in filter(
        lambda alias: not filedata_map[alias].has_ext(), filedata_map.keys()
    ):
        filename = filedata_map[alias].get_name()
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
