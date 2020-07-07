# DMFO -- Diff and Merge For Office

This is a set of scripts that enable convenient diff and merge of Office-related file
types (currently Word and PowerPoint (diff only)). The office application will be
started using COM automation, thus an Office installation is required.

DMFO is LFS compatible.

**Important:** Legacy PowerShell scripts are located in [ps1][ps1] and may still be
used. However, not all new features will be ported back to the ps1 scripts.

## Usage

### Installation

Installable with `pip` or [`pipx`][pipx] (recommended).

```cmd
pipx install git+https://github.com/lcnittl/DMFO.git
```

**Note:** Adding [PyPI][pypi] support is planned.

### Git Integration

These tools are intended to be used with Git, so that `git diff` and `git merge` will
use Office applications to compare and "merge" files. Simply configure `.gitattributes`
and `.gitconfig` to support the DMFO diff and merge driver. Enjoy to diff and merge
Office documents by simply running:

```cmd
> git diff
> git merge
```

as usual and with any paramter they accept.

#### `.gitconfig`

Simply register the diff and merge drivers by running `dmfo install` (scope can be given
by appending `system`, `global`, `local` or `worktree`, default is `global`).

Alternatively, add the entries manually to your git config:

```ini
[diff "dmfo"]
	name = DMFO diff driver
	command = dmfo diff
	binary = true
[merge "dmfo"]
	name = DMFO merge driver
	driver = dmfo merge %O %A %B %L %P
	binary = true
```

Make sure that `dmfo`'s path is in your path variable, otherwise prepand `dmfo` with the
executable's path.

#### `.gitattributes`

Specify the following drivers in your `.gitattributes` file (currently DMFO is only
supporting Word files):

```
*.doc diff=dmfo merge=dmfo
*.docx diff=dmfo merge=dmfo
*.ppt diff=dmfo
*.pptx diff=dmfo
```

### CLI

This option might be added at a later time.

## Reqirements

- Git (for Windows)
- Git LFS
- Powershell (>=5.1)
- Microsoft Office (\[and/or\]: Word, Powerpoint)

## Platform

In its current implementation, DMFO is suited for Windows 10. Not tested on other
platforms.

## License

GNU General Public License v3.0 or later

See [LICENSE][license] for the full text.

[license]: LICENSE
[extdiff]: https://github.com/ForNeVeR/ExtDiff
[pipx]: https://pypi.org/project/pipx/
[ps1]: ps1/
[pypi]: https://pypi.org/
