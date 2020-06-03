#  path old-file old-hex old-mode new-file new-hex new-mode
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


. $PSScriptRoot\constants.ps1


$extension = ".docx"


$activity = "Preparing files... "
$complete = 0
Write-Progress -Activity $activity -Status "Initializing" -PercentComplete $complete
$FileNames = @{
    Base = $BaseFileName;
    Local = $LocalFileName;
    Remote = $RemoteFileName
}
$FileNamesExt = @{}
$complete += 10

foreach ($key in @($FileNames.Keys)) {
    Write-Progress -Activity $activity -Status "Preparing $FileNames[$key]" -PercentComplete $complete
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


$activity = "Compiling 3-way-merge with MS Word. This may take a while... "
$complete = 0
Write-Progress -Activity $activity -Status "Initializing COM object" -PercentComplete $complete
try {
    $COMObj = New-Object -ComObject "Word.Application"
    $COMObj.Visible = $false
    $complete += 20
} catch [Runtime.Interopservices.COMException] {
    Write-Host "You must have Microsoft Word installed to perform this operation."
    exit(1)
}
try {
    Write-Progress -Activity $activity -Status "Opening BASE" -PercentComplete $complete
    $BaseFile = $COMObj.Documents.Open(
        [ref]$FileNamesExt["Base"],  # FileName
        [ref]$false,  # ConfirmConversions
        [ref]$false,  # ReadOnly
        [ref]$false  # AddToRecentFiles
    )
    $complete += 10

    Write-Progress -Activity $activity -Status "Opening LOCAL" -PercentComplete $complete
    $LocalFile = $COMObj.Documents.Open(
        [ref]$FileNamesExt["Local"],  # FileName
        [ref]$false,  # ConfirmConversions
        [ref]$false,  # ReadOnly
        [ref]$false  # AddToRecentFiles
    )
    $complete += 10

    Write-Progress -Activity $activity -Status "Opening REMOTE" -PercentComplete $complete
    $RemoteFile = $COMObj.Documents.Open(
        [ref]$FileNamesExt["Remote"],  # FileName
        [ref]$false,  # ConfirmConversions
        [ref]$false,  # ReadOnly
        [ref]$false  # AddToRecentFiles
    )
    $complete += 10

    Write-Progress -Activity $activity -Status "Diffing LOCAL vs BASE" -PercentComplete $complete
    $BaseFile.Activate()
    $BaseFile.Compare(
        [ref]$FileNamesExt["Local"],  # Name
        [ref]"LOCAL",  # AuthorName
        [ref][WdCompareTarget]::wdCompareTargetSelected,  # CompareTarget
        [ref]$true,  # DetectFormatChanges
        [ref]$true,  # IgnoreAllComparisonWarnings
        [ref]$false,  # AddToRecentFiles
        [ref]$false,  # RemovePersonalInformation
        [ref]$true  # RemoveDateAndTime
    )
    $complete += 5

    Write-Progress -Activity $activity -Status "Diffing REMOTE vs BASE" -PercentComplete $complete

    $BaseFile.Activate()
    $BaseFile.Compare(
        [ref]$FileNamesExt["Remote"],  # Name
        [ref]"REMOTE",  # AuthorName
        [ref][WdCompareTarget]::wdCompareTargetSelected,  # CompareTarget
        [ref]$true,  # DetectFormatChanges
        [ref]$true,  # IgnoreAllComparisonWarnings
        [ref]$false,  # AddToRecentFiles
        [ref]$false,  # RemovePersonalInformation
        [ref]$true  # RemoveDateAndTime
    )
    $complete += 5

    Write-Progress -Activity $activity -Status "Closing BASE" -PercentComplete $complete
    $BaseFile.Close(
        [ref][WdSaveOptions]::wdDoNotSaveChanges
    )
    $complete += 5

    Write-Progress -Activity $activity -Status "Merging changes" -PercentComplete $complete
    # Although the filename and not the object is speciefied in Merge
    # it takes the content of the document in the active session.
    $LocalFile.Activate()
    $LocalFile.Merge(
        [ref]$FileNamesExt["Remote"],  # Name
        [ref][WdMergeTarget]::wdMergeTargetNew,  # MergeTarget
        [ref]$true,  # DetectFormatChanges
        [ref][WdUseFormattingFrom]::wdFormattingFromPrompt,  # UseFormattingFrom
        [ref]$false  # AddToRecentFile
    )
    $MergedFile = $COMObj.ActiveDocument
    $complete += 10

    Write-Progress -Activity $activity -Status "Closing LOCAL and REMOTE" -PercentComplete $complete
    $LocalFile.Close(
        [ref][WdSaveOptions]::wdDoNotSaveChanges  # SaveChanges
    )
    $RemoteFile.Close(
        [ref][WdSaveOptions]::wdDoNotSaveChanges  # SaveChanges
    )
    $complete += 10

    Write-Progress -Activity $activity -Status "Saving MERGED" -PercentComplete $complete
    $MergedFile.SaveAs(
        [ref]$FileNamesExt["Local"],  # FileName
        [type]::missing,  # FileFormat  #[ref]$saveFormat::wdFormatDocument
        [type]::missing,  # LockComments
        [type]::missing,  # Password
        [ref]$false  # AddToRecentFiles
    )
    $complete += 5

    Write-Progress -Activity $activity -Status "Bringing to foreground" -PercentComplete $complete
    $COMObj.Visible = $true
    $COMObj.Activate()
    $COMObj.WindowState = [WdWindowState]::wdWindowStateMinimize
    $COMObj.WindowState = [WdWindowState]::wdWindowStateMaximize
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($_.Exception)
}
$complete = 100

Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
sleep 1

$resolved = $Host.UI.PromptForChoice(
    "Merge Complete?",
    "Confirm conflict resolution for ${MergeDest}: ",
    @("&Yes"; "&No"),
    1
)

$activity = "Cleaning up... "
$complete = 0
Write-Progress -Activity $activity -Status "Checking merged file" -PercentComplete $complete
try {
    # Check if file is still open
    $null = $COMObj.Documents.Item($FileNamesExt["Local"])
} catch [Runtime.Interopservices.COMException] {
    # Document was closed already
    $reopen = $true
} catch [Management.Automation.RuntimeException] {
    # ComObj was closed already
    $COMObj = New-Object -ComObject "Word.Application"
    $COMObj.Visible = $false
    $reopen = $true
}
if ($reopen) {
    $MergedFile = $COMObj.Documents.Open(
        [ref]$FileNamesExt["Local"],  # FileName
        [ref]$false,  # ConfirmConversions
        [ref]$false,  # ReadOnly
        [ref]$false  # AddToRecentFiles
    )
}

$MergedFile.Activate()
if ($MergedFile.TrackRevisions) {
    Write-Host "Warning: Track Changes is active. Please deactivate!"
}
if ($resolved -eq 0) {
    if ($MergedFile.Revisions.count -gt 0) {
        Write-Host "Warning: Unresolved revisions in the document. Please resolve!"
    }
}
$complete += 70

Write-Progress -Activity $activity -Status "Closing COM Object" -PercentComplete $complete
$MergedFile.Close()
if ($COMObj.Documents.Count -eq 0) {
    $COMObj.Quit()
}
$complete += 10

Write-Progress -Activity $activity -Status "Copying merged file" -PercentComplete $complete
if ($LFS) {
    Write-Host Converting LFS blob to pointer.
    cmd.exe /c "type $($FileNamesExt['Local']) | git-lfs clean > $($FileNames['Local'])"
} else {
    cp $FileNamesExt["Local"] $FileNames["Local"]
}
$complete += 10

Write-Progress -Activity $activity -Status "Removing aux files" -PercentComplete $complete
foreach ($key in $FileNamesExt.Keys) {
    rm $FileNamesExt[$key]
}
$complete = 100

Write-Progress -Activity $activity -Status "Done" -PercentComplete 100
sleep 1

exit($resolved)
