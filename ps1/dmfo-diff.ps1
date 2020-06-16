param(
    [string] $DiffPath,  # path
    [string] $LocalFileName,  # old-file  $LOCAL
    [string] $LocalFileHex,  # old-hex
    [string] $LocalFileMode,  # old-mode
    [string] $RemoteFileName,  # new-file  $REMOTE
    [string] $RemoteFileHex,  # new-hex
    [string] $RemoteFileMode  #  new-mode
)
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms

$extension = [System.IO.Path]::GetExtension($DiffPath)


$activity = "Preparing files... "
$complete = 0
Write-Progress -Activity $activity -Status "Initializing" -PercentComplete $complete
$FileNames = @{
    LOCAL = $LocalFileName;
    REMOTE = $RemoteFileName
}
$Files = @{}
$complete += 20

foreach ($key in @($FileNames.Keys)) {
    Write-Progress -Activity $activity -Status "Preparing $key" -PercentComplete $complete
    $FileName = (Resolve-Path $FileNames[$key]).Path
    $FileNames[$key] = $FileName
    git lfs pointer --check --file $FileName
    if ($?) {
        $LFS = $true
        Write-Host Converting LFS pointer to blob.
        cmd.exe /c "type $($FileName) | git-lfs smudge > $($FileName + "_")"
        mv -Force $($FileName + "_") $FileName
    }
    $File = Get-ChildItem $FileName
    if ($File.IsReadOnly) {
        $File.IsReadOnly = $false
    }
    $complete += 40
}
$complete = 100

Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
sleep 1


if (@(".doc", ".docx") -contains $extension) {
    . $PSScriptRoot\dmfo-diff\dmfo-diff_wd.ps1
} elseif (@(".ppt", ".pptx") -contains $extension) {
    . $PSScriptRoot\dmfo-diff\dmfo-diff_pp.ps1
} else {
    Write-Host "DMFO-Diff does not know what to do with '$extension' files."
    exit(1)
}
exit($LastExitCode)
