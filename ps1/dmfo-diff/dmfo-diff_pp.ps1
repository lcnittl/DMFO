. $PSScriptRoot\..\constants\const_mso.ps1
. $PSScriptRoot\..\constants\const_pp.ps1


$activity = "Compiling diff of '$DiffPath' with MS PowerPoint. This may take a while... "
$complete = 0
Write-Progress -Activity $activity -Status "Initializing COM object" -PercentComplete $complete
try {
    $COMObj = New-Object -ComObject "PowerPoint.Application"
    # $COMObj.Visible = [msoTriState]::msoFalse  # Not allowed in PowerPoint
    $complete += 20
} catch [Runtime.Interopservices.COMException] {
    Write-Host "You must have Microsoft PowerPoint installed to perform this operation."
    exit(1)
}
try {
    foreach ($key in @("LOCAL")) {
        Write-Progress -Activity $activity -Status "Opening $key" -PercentComplete $complete
        $File = $COMObj.Presentations.Open(
            [ref]$FileNames[$key],  # FileName
            [ref][msoTriState]::msoFalse,  # ReadOnly
            [ref][msoTriState]::msoTrue,  # Untitled
            [ref][msoTriState]::msoTrue  # WithWindow  # False would be nice, but then Merge fails
        )
        $Files += @{$key = $File}
        $complete += 40
    }

    Write-Progress -Activity $activity -Status "Diffing REMOTE vs LOCAL" -PercentComplete $complete
    $Files["LOCAL"].Merge(
        [ref]$FileNames["REMOTE"]  # Path
    )

    $DiffFile = $COMObj.ActivePresentation
    $complete += 20

    Write-Progress -Activity $activity -Status "Setting DIFF to unsaved" -PercentComplete $complete
    $DiffFile.Saved = 1
    $complete += 10

    Write-Progress -Activity $activity -Status "Bringing to foreground" -PercentComplete $complete
    $COMObj.Activate()
    $COMObj.ActiveWindow.WindowState = [PpWindowState]::ppWindowMinimized
    $COMObj.ActiveWindow.WindowState = [PpWindowState]::ppWindowMaximized
    $complete = 100

    Write-Progress -Activity $activity -Status "Done" -PercentComplete $complete
    sleep 1
} catch {
    [System.Windows.Forms.MessageBox]::Show($_.Exception)
}
exit(0)
