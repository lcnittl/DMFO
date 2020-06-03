# DMFO -- Diff and Merge For Office

This is a set of scripts that enable convenient diff and merge of Office-related file
types (currently only Word, future support for Excel and PowerPoint is planned). The
office application will be started using COM automation.

DMFO is LFS compatible.

## Usage

### Git Integration

These tools are intended to use with git, so that `git diff` and `git merge` will use
Office applications compare and "merge" files.

Usage requires the configuration of `.gitattributes` and `.gitconfig` to support a
custom diff and merge driver.

#### `.gitconfig`

Add the following drivers to your git config file:

```ini
[diff "dmfo"]
	name = DMFO diff driver
	command = "powershell.exe -File <pathToDMFO>/dmfo-diff.ps1"
	binary = true
[merge "dmfo"]
	name = DMFO merge driver
	driver = "powershell.exe -File ~<pathToDMFO>/dmfo-merge.ps1 %O %A %B %L %P"
	binary = true
```

Replace `<pathToDMFO>` with the repo's path.

#### `.gitattributes`

Specify the following drivers in your `.gitattributes` file (currently DMFO only
supports Word files):

```
*.docx diff=dmfo merge=dmfo
*.pptx diff=dmfo merge=dmfo
*.xlsx diff=dmfo
```

### CLI

This option might be added at a later time.

---

Inspired by [ExtDiff][extdiff].

[extdiff]: https://github.com/ForNeVeR/ExtDiff
