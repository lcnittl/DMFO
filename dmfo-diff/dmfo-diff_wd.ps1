. $PSScriptRoot\..\constants\const_wd.ps1


$activity = "Compiling diff of '$DiffPath' with MS Word. This may take a while... "
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
    foreach ($key in $FileNames.Keys) {
        Write-Progress -Activity $activity -Status "Opening $key" -PercentComplete $complete
        $File = $COMObj.Documents.Open(
            [ref]$FileNames[$key],  # FileName
            [ref]$false,  # ConfirmConversions
            [ref]$false,  # ReadOnly
            [ref]$false  # AddToRecentFiles
        )
        $Files += @{$key = $File}
        $complete += 15
    }

    Write-Progress -Activity $activity -Status "Diffing REMOTE vs LOCAL" -PercentComplete $complete
    $DiffFile = $COMObj.CompareDocuments(
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
        [ref]"REMOTE",  # RevisedAuthor
        [ref]$true  # IgnoreAllComparisonWarnings
    )
    $complete += 10

    foreach ($key in $Files.Keys) {
        Write-Progress -Activity $activity -Status "Closing $key" -PercentComplete $complete
        $Files[$key].Close(
            [ref][WdSaveOptions]::wdDoNotSaveChanges  # SaveChanges
        )
        $complete += 5
    }

    Write-Progress -Activity $activity -Status "Setting DIFF to unsaved" -PercentComplete $complete
    $DiffFile.Saved = 1
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
    [System.Windows.Forms.MessageBox]::Show($_.Exception)
}
exit(0)
