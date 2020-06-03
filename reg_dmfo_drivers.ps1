param(
    [parameter(Mandatory=$false)][ValidateSet(
        "--local",
        "--global",
        "--system"
    )][string] $scope = "--global"  # Scope
)
$ErrorActionPreference = "Stop"

$InstallPath = $PSScriptRoot -replace "\\","/"

# Register Diff Driver
git config $scope diff.dmfo.name "DMFO diff driver"
git config $scope diff.dmfo.command "powershell.exe -File $InstallPath/dmfo-diff.ps1"
git config $scope diff.dmfo.binary "true"

# Register Merge Driver
git config $scope merge.dmfo.name "DMFO merge driver"
git config $scope merge.dmfo.driver "powershell.exe -File $InstallPath/dmfo-merge.ps1 %O %A %B %L %P"
git config $scope merge.dmfo.binary "true"
