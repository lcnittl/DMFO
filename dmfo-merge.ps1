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
    foreach ($key in $FileNamesExt.Keys) {
        Write-Progress -Activity $activity -Status "Opening $key" -PercentComplete $complete
        $File = $COMObj.Documents.Open(
            [ref]$FileNamesExt[$key],  # FileName
            [ref]$false,  # ConfirmConversions
            [ref]$false,  # ReadOnly
            [ref]$false  # AddToRecentFiles
        )
        $Files += @{$key = $File}
        $complete += 10
    }

    foreach ($key in @("LOCAL", "REMOTE")) {
        Write-Progress -Activity $activity -Status "Diffing $key vs BASE" -PercentComplete $complete
        $Files[$key] = $COMObj.CompareDocuments(
            [ref]$Files["BASE"],  # OriginalDocument
            [ref]$Files[$key],  # RevisedDocument
            [ref][WdCompareDestination]::wdCompareDestinationRevised,  # Destination
            [type]::missing,  # Granularity
            [ref]$true,  # CompareFormatting
            [ref]$true,  # CompareCaseChanges
            [ref]$true,  # CompareWhitespace
            [ref]$true,  # CompareTables
            [ref]$true,  # CompareHeaders
            [ref]$true,  # CompareFootnotes
            [ref]$true,  # CompareTextboxes
            [ref]$true,  # CompareFields
            [ref]$true,  # CompareComments
            [type]::missing,
            [ref]$key,  # RevisedAuthor
            [ref]$true  # IgnoreAllComparisonWarnings
        )
        $complete += 5
        Write-Host Opened $key
    }

    Write-Progress -Activity $activity -Status "Merging changes" -PercentComplete $complete
    $MergedFile = $COMObj.MergeDocuments(
        [ref]$Files["LOCAL"],  # OriginalDocument
        [ref]$Files["REMOTE"],  # RevisedDocument
        [ref][WdCompareDestination]::wdCompareDestinationNew,  # Destination
        [type]::missing,  # Granularity
        [ref]$true,  # CompareFormatting
        [ref]$true,  # CompareCaseChanges
        [ref]$true,  # CompareWhitespace
        [ref]$true,  # CompareTables
        [ref]$true,  # CompareHeaders
        [ref]$true,  # CompareFootnotes
        [ref]$true,  # CompareTextboxes
        [ref]$true,  # CompareFields
        [ref]$true,  # CompareComments
        [type]::missing,
        [ref]"Merge LOCAL",  # OriginalAuthor
        [ref]"Merge REMOTE",  # RevisedAuthor
        [ref][WdUseFormattingFrom]::wdFormattingFromPrompt  # FormatFrom
    )
    $complete += 10

    foreach ($key in $Files.Keys) {
        Write-Progress -Activity $activity -Status "Closing $key" -PercentComplete $complete
        $Files[$key].Close(
            [ref][WdSaveOptions]::wdDoNotSaveChanges  # SaveChanges
        )
        $complete += 5
    }

    Write-Progress -Activity $activity -Status "Saving MERGED" -PercentComplete $complete
    $MergedFile.SaveAs(
        [ref]$FileNamesExt["LOCAL"],  # FileName
        [type]::missing,  # FileFormat
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
    $reopen = $true
}
$COMObj.Visible = $false
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
