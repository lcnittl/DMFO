. $PSScriptRoot\..\constants\const_wd.ps1


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
    [System.Windows.Forms.MessageBox]::Show($_.Exception)
}
$complete = 100

Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
sleep 1

$resolved = [System.Windows.Forms.MessageBox]::Show(
    "Confirm conflict resolution for ${MergeDest}?",
    "Merge Complete?",
    "YesNo",
    "Question")

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
if ($resolved -eq "Yes") {
    if ($MergedFile.Revisions.Count -gt 0) {
        Write-Host "Warning: Unresolved revisions in the document. Will exit as 'unresolved'."
        $resolved = "No"
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

if ($resolved -eq "No") {
    exit(1)
}
exit(0)
