param(
    [parameter(Mandatory=$true)][string] $BaseFileName,  # $BASE
    [parameter(Mandatory=$true)][string] $LocalFileName,  # $LOCAL
    [parameter(Mandatory=$true)][string] $RemoteFileName,  # $REMOTE
    [parameter(Mandatory=$false)][string] $ConflictMarkerSize,  # conflict-marker-size
    [parameter(Mandatory=$false)][string] $MergeDest  # $MERGED
)
if ($PSVersionTable.PSVersion.Major -lt 6) {
    $PSDefaultParameterValues["Out-File:Encoding"] = "utf8"
}
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms

$extension = [System.IO.Path]::GetExtension($MergeDest)


$activity = "Preparing files... "
$complete = 0
Write-Progress -Activity $activity -Status "Initializing" -PercentComplete $complete
$FileNames = @{
    BASE = $BaseFileName;
    LOCAL = $LocalFileName;
    REMOTE = $RemoteFileName
}
$FileNamesExt = @{}
$Files = @{}
$complete += 10

foreach ($key in @($FileNames.Keys)) {
    Write-Progress -Activity $activity -Status "Preparing $key" -PercentComplete $complete
    $FileName = (Resolve-Path $FileNames[$key]).Path
    $FileNames[$key] = $FileName
    $FileNameExt = $FileName + $extension
    $FileNamesExt += @{$key = "$FileNameExt"}
    git lfs pointer --check --file $FileName
    if ($?) {
        $LFS = $true
        Write-Host Converting LFS pointer to blob.
        cmd.exe /c "type $($FileName) | git-lfs smudge > $($FileNameExt)"
    } else {
        cp $FileName $FileNameExt
    }
    $File = Get-ChildItem $FileNameExt
    if ($File.IsReadOnly) {
        $File.IsReadOnly = $false
    }
    $complete += 30
}
$complete = 100

Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
sleep 1


if (@(".doc", ".docx") -contains $extension) {
    . $PSScriptRoot\dmfo-merge\dmfo-merge_wd.ps1
} else {
    Write-Host "DMFO-Merge does not know what to do with '$extension' files."
    exit(1)
}
exit($LastExitCode)
