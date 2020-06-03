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
#$ErrorActionPreference = "Stop"


$extension = ".docx"

enum WdCompareTarget {
    wdCompareTargetSelected = 0  # Places comparison differences in the target document.
    # wdCompareTargetCurrent = 1  # Places comparison differences in the current document. Default.
    # wdCompareTargetNew = 2  # Places comparison differences in a new document.
}
enum WdMergeTarget {
    # wdMergeTargetSelected = 0  # Merge into selected document.
    # wdMergeTargetCurrent = 1  # Merge into current document.
    wdMergeTargetNew = 2  # Merge into new document.
}
enum WdOpenFormat {
    wdOpenFormatAuto = 0  # The existing format.
    # wdOpenFormatDocument = 1  # Word format.
    # wdOpenFormatDocument97 = 1  # Microsoft Word 97 document format.
    # wdOpenFormatTemplate = 2  # As a Word template.
    # wdOpenFormatTemplate97 = 2  # Word 97 template format.
    # wdOpenFormatRTF = 3  # Rich text format (RTF).
    # wdOpenFormatText = 4  # Unencoded text format.
    # wdOpenFormatEncodedText = 5  # Encoded text format.
    # wdOpenFormatUnicodeText = 5  # Unicode text format.
    # wdOpenFormatAllWord = 6  # A Microsoft Word format that is backward compatible with earlier versions of Word.
    # wdOpenFormatWebPages = 7  # HTML format.
    # wdOpenFormatXML = 8  # XML format.
    # wdOpenFormatXMLDocument = 9  # XML document format.
    # wdOpenFormatXMLDocumentMacroEnabled = 10  # XML document format with macros enabled.
    # wdOpenFormatXMLTemplate = 11  # XML template format.
    # wdOpenFormatXMLTemplateMacroEnabled = 12  # XML template format with macros enabled.
    # wdOpenFormatAllWordTemplates = 13  # Word template format.
    # wdOpenFormatXMLDocumentSerialized = 14  # Open XML file format saved as a single XML file.
    # wdOpenFormatXMLDocumentMacroEnabledSerialized = 15  # Open XML file format with macros enabled saved as a single XML file.
    # wdOpenFormatXMLTemplateSerialized = 16  # Open XML template format saved as a XML single file.
    # wdOpenFormatXMLTemplateMacroEnabledSerialized = 17  # Open XML template format with macros enabled saved as a single XML file.
    # wdOpenFormatOpenDocumentText = 18  # OpenDocument Text format.
}
enum WdSaveOptions {
    # wdPromptToSaveChanges = -2  # Prompt the user to save pending changes.
    # wdSaveChanges = -1 # Save pending changes automatically without prompting the user.
    wdDoNotSaveChanges = 0  # Do not save pending changes.
}
enum WdUseFormattingFrom {
    # wdFormattingFromCurrent = 0  # Copy source formatting from the current item.
    # wdFormattingFromSelected = 1  # Copy source formatting from the current selection.
    wdFormattingFromPrompt = 2  # Prompt the user for formatting to use.
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
        [ref]$false,  # AddToRecentFiles
        [type]::missing,  # PasswordDocument
        [type]::missing,  # PasswordTemplate
        [type]::missing,  # Revert
        [type]::missing,  # WritePasswordDocument
        [type]::missing,  # WritePasswordTemplate
        [ref][wdOpenFormat]::wdOpenFormatAuto  # Format
    )
    $complete += 10

    Write-Progress -Activity $activity -Status "Opening LOCAL" -PercentComplete $complete
    $LocalFile = $COMObj.Documents.Open(
        [ref]$FileNamesExt["Local"],  # FileName
        [ref]$false,  # ConfirmConversions
        [ref]$false,  # ReadOnly
        [ref]$false,  # AddToRecentFiles
        [type]::missing,  # PasswordDocument
        [type]::missing,  # PasswordTemplate
        [type]::missing,  # Revert
        [type]::missing,  # WritePasswordDocument
        [type]::missing,  # WritePasswordTemplate
        [ref][wdOpenFormat]::wdOpenFormatAuto  # Format
    )
    $complete += 10

    Write-Progress -Activity $activity -Status "Opening REMOTE" -PercentComplete $complete
    $RemoteFile = $COMObj.Documents.Open(
        [ref]$FileNamesExt["Remote"],  # FileName
        [ref]$false,  # ConfirmConversions
        [ref]$false,  # ReadOnly
        [ref]$false,  # AddToRecentFiles
        [type]::missing,  # PasswordDocument
        [type]::missing,  # PasswordTemplate
        [type]::missing,  # Revert
        [type]::missing,  # WritePasswordDocument
        [type]::missing,  # WritePasswordTemplate
        [ref][wdOpenFormat]::wdOpenFormatAuto  # Format
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
    #$RemoteFile.Activate()
    #$RemoteFile.Merge(
    #    [ref]$FileNamesExt["Local"],  # Name
    #    [ref][WdMergeTarget]::wdMergeTargetNew,  # MergeTarget
    #    [ref]$true,  # DetectFormatChanges
    #    [ref][WdUseFormattingFrom]::wdFormattingFromPrompt,  # UseFormattingFrom
    #    [ref]$false  # AddToRecentFile
    #)
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
    # $COMObj = [Runtime.Interopservices.Marshal]::GetActiveObject("Word.Application")
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
        [ref]$false,  # AddToRecentFiles
        [type]::missing,  # PasswordDocument
        [type]::missing,  # PasswordTemplate
        [type]::missing,  # Revert
        [type]::missing,  # WritePasswordDocument
        [type]::missing,  # WritePasswordTemplate
        [ref][wdOpenFormat]::wdOpenFormatAuto  # Format
    )
}

$MergedFile.Activate()
if ($MergedFile.TrackRevisions) {
    Write-Host "Warning: Track Changes is active. Please deactivate!"
    # $MergedFile.TrackRevisions = 0
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
