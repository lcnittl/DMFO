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

$ErrorActionPreference = 'Stop'

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


$FileNames = @{
    Local = $LocalFileName;
    Remote = $RemoteFileName
}

foreach ($key in @($FileNames.Keys)) {
    $FileName = (Resolve-Path $FileNames[$key]).Path
    $FileNames[$key] = $FileName
    git lfs pointer --check --file $FileName
    if ($?) {
         Write-Host Converting LFS pointer to blob.
        cmd.exe /c "type $($FileName) | git-lfs smudge > $($FileName)"
    }
    $File = Get-ChildItem $FileNameExt
    if ($File.IsReadOnly) {
        $File.IsReadOnly = $false
    }
}


try {
    $activity = "Compiling diff of '$DiffPath' with MS Word. This may take a while... "
    Write-Progress -Activity $activity -Status "Initializing COM object" -PercentComplete 0
    $COMObj = New-Object -ComObject Word.Application
    $COMObj.Visible = $false

    Write-Progress -Activity $activity -Status "Opening LOCAL" -PercentComplete 40
    $LocalFile = $COMObj.Documents.Open(
        [ref]$FileNames["Local"],  # FileName
        [ref]$false,  # ConfirmConversions
        [ref]$false,  # ReadOnly
        [ref]$false  # AddToRecentFiles
    )

    Write-Progress -Activity $activity -Status "Diffing REMOTE vs LOCAL" -PercentComplete 60
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

    Write-Progress -Activity $activity -Status "Closing LOCAL" -PercentComplete 70
    $LocalFile.Close(
        [ref][WdSaveOptions]::wdDoNotSaveChanges  # SaveChanges
    )

    Write-Progress -Activity $activity -Status "Setting DIFF to unsaved" -PercentComplete 80
    $COMObj.ActiveDocument.Saved = 1

    Write-Progress -Activity $activity -Status "Bringing to foreground" -PercentComplete 90
    $COMObj.Visible = $true
    $COMObj.Activate()
    $COMObj.WindowState = [WdWindowState]::wdWindowStateMinimize
    $COMObj.WindowState = [WdWindowState]::wdWindowStateMaximize
    Write-Progress -Activity $activity -Status "Done" -PercentComplete 100
    sleep 1
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($_.Exception)
}
