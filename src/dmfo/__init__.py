import argparse
import logging
import logging.handlers
import shutil
import sys
import tempfile
from importlib import metadata
from pathlib import Path

import colorlog
import dmfo.driver
import dmfo.files
import dmfo.installer
from dmfo.classes import VCSFileData

try:
    __version__ = metadata.version("DMFO")
except metadata.PackageNotFoundError:
    __version__ = None

logger = logging.getLogger(__name__)

DEFAULT_LOG_PATH = Path(".")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )

    parser.add_argument(
        "--version",
        action="version",
        version=str(__version__),
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

    subparser = parser.add_subparsers(
        dest="mode", title="Mode", required=True, help="Choose mode of operation:"
    )

    diff_parser = subparser.add_parser("diff", help="Run diff driver")
    diff_parser.add_argument(
        "DiffPath",
        # dest="TargetPath",
        type=Path,
        help="path",
        metavar="DPath",
    )
    diff_parser.add_argument(
        "LocalFileName",
        type=Path,
        help="old-file ($LOCAL)",
        metavar="LFName",
    )
    diff_parser.add_argument(
        "LocalFileHex",
        type=str,
        help="old-hex",
        metavar="LFHex",
    )
    diff_parser.add_argument(
        "LocalFileMode",
        type=str,
        help="old-mode",
        metavar="LFMode",
    )
    diff_parser.add_argument(
        "RemoteFileName",
        type=Path,
        help="new-file ($REMOTE)",
        metavar="RFName",
    )
    diff_parser.add_argument(
        "RemoteFileHex",
        type=str,
        help="new-hex",
        metavar="RFHex",
    )
    diff_parser.add_argument(
        "RemoteFileMode",
        type=str,
        help="new-mode",
        metavar="RFMode",
    )

    merge_parser = subparser.add_parser("merge", help="Run merge driver")
    merge_parser.add_argument(
        "BaseFileName",
        type=Path,
        help="$BASE (%%O)",
        metavar="BFName",
    )
    merge_parser.add_argument(
        "LocalFileName",
        type=Path,
        help="$LOCAL (%%A)",
        metavar="LFName",
    )
    merge_parser.add_argument(
        "RemoteFileName",
        type=Path,
        help="$REMOTE (%%B)",
        metavar="RFName",
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

    install_parser = subparser.add_parser("install", help="Add DMFO to Git config")
    install_parser.add_argument(
        "scope",
        default="global",
        type=str.lower,
        nargs="?",
        choices=["system", "global", "local", "worktree"],
        help="Scope",
    )

    return parser.parse_args()


def setup_root_logger(path: Path = DEFAULT_LOG_PATH) -> logging.Logger:
    global logfile_path

    logger = logging.getLogger()
    logger.setLevel(logging.NOTSET)

    """
    module_loglevel_map = {
        "pywin32": logger.WARNING,
    }
    for module, loglevel in module_loglevel_map.items():
        logging.getLogger(module).setLevel(loglevel)
    """

    logfile_path = Path(f"{Path(path) / Path(__file__).stem}.log")
    log_roll = logfile_path.is_file()
    file_handler = logging.handlers.RotatingFileHandler(
        filename=logfile_path,
        mode="a",
        backupCount=9,
        encoding="utf-8",
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


args = parse_args()
TempDir = Path(tempfile.mkdtemp(prefix=f"dmfo_{args.mode}_"))
root_logger = setup_root_logger(path=TempDir)


def main():
    """Return codes:
    1: Not auto-merged
    2: Unknown file extension
    3: COM Application (Word, PowerPoint) not installed
    4: File not found
    5: Unknown git lfs pointer --check return code
    6: Unexpected pywin32 com_error
    """
    logger.debug(
        "DMFO is logging to '%s'",
        logfile_path,
    )

    if args.mode == "install":
        ret = dmfo.installer.install(scope=args.scope)
    else:
        filedatamap = {
            "LOCAL": VCSFileData(args.LocalFileName),
            "REMOTE": VCSFileData(args.RemoteFileName),
        }

        # extension = args.TargetPath.suffix
        if args.mode == "diff":
            extension = args.DiffPath.suffix
            logger.debug("Diffing '%s' file.", extension)
        elif args.mode == "merge":
            extension = args.MergeDest.suffix
            logger.debug("Merging '%s' file.", extension)

            filedatamap["BASE"] = VCSFileData(args.BaseFileName)
        VCSFileData.target_ext = extension

        ret = dmfo.files.preproc(filedata_map=filedatamap)
        if ret:
            sys.exit(ret)

        if args.mode == "diff":
            ret = dmfo.driver.diff(filedata_map=filedatamap)
        elif args.mode == "merge":
            ret = dmfo.driver.merge(filedata_map=filedatamap)

        dmfo.files.postproc(filedata_map=filedatamap, mode=args.mode)

    if ret <= 1:
        logger.debug("Removing log temp dir (%s)", TempDir)

        # Remove all file handlers
        for handler in filter(
            lambda handler: isinstance(
                handler, (logging.FileHandler, logging.handlers.RotatingFileHandler)
            ),
            root_logger.handlers,
        ):
            handler.close()
            root_logger.removeHandler(handler)

        shutil.rmtree(TempDir)
    else:
        logger.critical(
            "DMFO %s exited with return code %s, check log for details (%s)",
            args.mode,
            ret,
            logfile_path,
        )
    sys.exit(ret)
