# DMFO -- Diff and Merge For Office

This is a set of scripts that enable convenient diff and merge of Office-related file
types (currently Word and PowerPoint (diff only)). The office application will be
started using COM automation, thus an Office installation is required.

DMFO is LFS compatible.

## Usage

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

Simply register the diff and merge drivers by running [`reg_dmfo_drivers.ps1`][register]
(scope can be given as parameter `--local`, `--global` or `--system`, default is
`--global`).

Alternatively, add the entries manually to your git config:

```ini
[diff "dmfo"]
	name = DMFO diff driver
	command = "powershell.exe -File <pathToDMFO>/dmfo-diff.ps1"
	binary = true
[merge "dmfo"]
	name = DMFO merge driver
	driver = "powershell.exe -File <pathToDMFO>/dmfo-merge.ps1 %O %A %B %L %P"
	binary = true
```

Replace `<pathToDMFO>` with the repo's path.

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

See [COPYING][copying] for the full text.

---

Inspired by [ExtDiff][extdiff].

[copying]: COPYING
[extdiff]: https://github.com/ForNeVeR/ExtDiff
[register]: reg_dmfo_drivers.ps1
