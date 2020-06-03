# DMFO -- Diff and Merge For Office

This is a set of scripts that enable convenient diff and merge of Office-related file
types (currently only Word, future support for Excel and PowerPoint is planned). The
office application will be started using COM automation.

DMFO is LFS compatible.

## Usage

### Git Integration

You can also use this tool with git, so that `git diff` will use Microsoft Word to diff
`*.docx` files.

To do this, you must configure your `.gitattributes` and `.gitconfig` to support a
custom diff tool.

#### `.gitattributes`

Specify the following drivers in your `.gitattributes` file (currently DMFO only
supports Word files):

```
*.docx diff=dmfo merge=dmfo
*.pptx diff=dmfo merge=dmfo
*.xlsx diff=dmfo
```

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

### CLI

This option might be added at a late time.

---

Inspired by [ExtDiff][extdiff].

[extdiff]: https://github.com/ForNeVeR/ExtDiff
