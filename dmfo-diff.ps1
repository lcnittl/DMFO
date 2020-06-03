#  path old-file old-hex old-mode new-file new-hex new-mode
param(
    [string] $DiffPath,
    [string] $LocalFileName,  # $LOCAL
    [string] $LocalFileHex,
    [string] $LocalFileMode,
    [string] $RemoteFileName,  # $REMOTE
    [string] $RemoteFileHex,
    [string] $RemoteFileMode
)
$ErrorActionPreference = "Stop"

# Constants
enum WdCompareTarget {
    # wdCompareTargetSelected = 0  # Places comparison differences in the target document.
    # wdCompareTargetCurrent = 1  # Places comparison differences in the current document. Default.
    wdCompareTargetNew = 2  # Places comparison differences in a new document.
}
enum WdSaveOptions {
    # wdPromptToSaveChanges = -2  # Prompt the user to save pending changes.
    # wdSaveChanges = -1 # Save pending changes automatically without prompting the user.
    wdDoNotSaveChanges = 0  # Do not save pending changes.
}
enum WdWindowState {
    # wdWindowStateNormal = 0  # Normal.
    wdWindowStateMaximize = 1  # Maximized.
    wdWindowStateMinimize = 2  # Minimized.
}


$activity = "Preparing files... "
$complete = 0
Write-Progress -Activity $activity -Status "Initializing" -PercentComplete $complete
$FileNames = @{
    Local = $LocalFileName;
    Remote = $RemoteFileName
}
$complete += 20

foreach ($key in @($FileNames.Keys)) {
    Write-Progress -Activity $activity -Status "Preparing $FileNames[$key]" -PercentComplete $complete
    $FileName = (Resolve-Path $FileNames[$key]).Path
    $FileNames[$key] = $FileName
    git lfs pointer --check --file $FileName
    if ($?) {
        $LFS = $true
        Write-Host Converting LFS pointer to blob.
        cmd.exe /c "type $($FileName) | git-lfs smudge > $($FileName + "_")"
        mv -Force $($FileName + "_") $FileName
    }
    $File = Get-ChildItem $FileNameExt
    if ($File.IsReadOnly) {
        $File.IsReadOnly = $false
    }
    $complete += 40
}
$complete = 100

Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
sleep 1


$activity = "Compiling diff of '$DiffPath' with MS Word. This may take a while... "
$complete = 0
Write-Progress -Activity $activity -Status "Initializing COM object" -PercentComplete $complete
try {
    $COMObj = New-Object -ComObject "Word.Application"
    $COMObj.Visible = $false
    $complete += 40
} catch [Runtime.Interopservices.COMException] {
    Write-Host "You must have Microsoft Word installed to perform this operation."
    exit(1)
}
try {
    Write-Progress -Activity $activity -Status "Opening LOCAL" -PercentComplete $complete
    $LocalFile = $COMObj.Documents.Open(
        [ref]$FileNames["Local"],  # FileName
        [ref]$false,  # ConfirmConversions
        [ref]$false,  # ReadOnly
        [ref]$false  # AddToRecentFiles
    )
    $complete += 20

    Write-Progress -Activity $activity -Status "Diffing REMOTE vs LOCAL" -PercentComplete $complete
    $LocalFile.Compare(
        [ref]$FileNames["Remote"],  # Name
        [ref]"REMOTE",  # AuthorName
        [ref][WdCompareTarget]::wdCompareTargetNew,  # CompareTarget
        [ref]$true,  # DetectFormatChanges
        [ref]$true,  # IgnoreAllComparisonWarnings
        [ref]$false,  # AddToRecentFiles
        [ref]$false,  # RemovePersonalInformation
        [ref]$true  # RemoveDateAndTime
    )
    $complete += 10

    Write-Progress -Activity $activity -Status "Closing LOCAL" -PercentComplete $complete
    $LocalFile.Close(
        [ref][WdSaveOptions]::wdDoNotSaveChanges  # SaveChanges
    )
    $complete += 10

    Write-Progress -Activity $activity -Status "Setting DIFF to unsaved" -PercentComplete $complete
    $COMObj.ActiveDocument.Saved = 1
    $complete += 10

    Write-Progress -Activity $activity -Status "Bringing to foreground" -PercentComplete $complete
    $COMObj.Visible = $true
    $COMObj.Activate()
    $COMObj.WindowState = [WdWindowState]::wdWindowStateMinimize
    $COMObj.WindowState = [WdWindowState]::wdWindowStateMaximize
    $complete = 100

    Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
    sleep 1
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($_.Exception)
}
