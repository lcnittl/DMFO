import argparse
import logging
from pathlib import Path
from typing import Dict

import driver.differ
import driver.merger
from classes import VCSFileData

logger = logging.getLogger(__name__)


def diff(filedata_map: Dict[str, object]) -> int:
    filedata_map["DIFF"] = VCSFileData(Path())

    extension = VCSFileData.target_ext
    if extension in [".doc", ".docx"]:
        ret = driver.differ.wd(filedata_map=filedata_map)
    elif extension in [".ppt", ".pptx"]:
        ret = driver.differ.pp(filedata_map=filedata_map)
    else:
        logger.critical(
            "DMFO-Diff does not know what to do with '%s' files.", extension
        )
        ret = 2

    filedata_map.pop("DIFF")
    return ret


def merge(filedata_map: Dict[str, object]) -> int:
    filedata_map["MERGE"] = VCSFileData(Path())

    extension = VCSFileData.target_ext
    if extension in [".doc", ".docx"]:
        ret = driver.merger.wd(filedata_map=filedata_map)
    else:
        logger.critical(
            "DMFO-Merge does not know what to do with '%s' files.", extension
        )
        ret = 2

    filedata_map.pop("MERGE")
    return ret
