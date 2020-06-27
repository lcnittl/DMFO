# DMFO -- Diff and Merge For Office (Legacy PowerShell Edition)

**Important:** This folder contains the legacy PowerShell scripts. This readme solely
contains the differences in the setup of DMFO legacy PowerShell scripts. Please
reregister DMFO in your Git config if you are updating from a pre-Python version (_vide
infra_) to continue using DMFO PowerShell edition, or switch to DMFO python edition.

## Usage

### Git Integration

#### `.gitconfig`

Simply register the diff and merge drivers by running [`reg_dmfo_drivers.ps1`][register]
(scope can be given as parameter `--local`, `--global` or `--system`, default is
`--global`).

Alternatively, add the entries manually to your git config:

```ini
[diff "dmfo"]
	name = DMFO diff driver
	command = "powershell.exe -File <pathToDMFO>/ps1/dmfo-diff.ps1"
	binary = true
[merge "dmfo"]
	name = DMFO merge driver
	driver = "powershell.exe -File <pathToDMFO>/ps1/dmfo-merge.ps1 %O %A %B %L %P"
	binary = true
```

Replace `<pathToDMFO>` with the repo's path.

---

Inspired by [ExtDiff][extdiff].

[extdiff]: https://github.com/ForNeVeR/ExtDiff
[register]: /ps1/reg_dmfo_drivers.ps1
