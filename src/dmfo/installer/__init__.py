import logging
import shlex
import subprocess  # nosec
import sys

logger = logging.getLogger(__name__)


def install(scope: str) -> int:
    logger.info("Registering DMFO in %s git config...", scope)

    scope = f"--{scope}"
    dmfo_exec = "dmfo.exe"

    cmds = []
    # Register Diff Driver
    cmds += [f"git config {scope} diff.dmfo.name 'DMFO diff driver'"]
    cmds += [f"git config {scope} diff.dmfo.command '{dmfo_exec} diff'"]
    cmds += [f"git config {scope} diff.dmfo.binary 'true'"]
    # Register Merge Driver
    cmds += [f"git config {scope} merge.dmfo.name 'DMFO merge driver'"]
    cmds += [f"git config {scope} merge.dmfo.driver '{dmfo_exec} merge %O %A %B %L %P'"]
    cmds += [f"git config {scope} merge.dmfo.binary 'true'"]

    for cmd in cmds:
        logger.debug("Executing: %s", cmd)
        ret = subprocess.run(  # nosec
            shlex.split(cmd),
            stdout=sys.stdout,
            stderr=sys.stderr,
        ).returncode
        if ret:
            logger.error("Returned: %s", ret)
            return ret
        else:
            logger.debug("Done.")

    logger.info("Done.")
    return 0
