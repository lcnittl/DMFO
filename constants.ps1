# Constants
enum WdCompareTarget {
    wdCompareTargetSelected = 0  # Places comparison differences in the target document.
    # wdCompareTargetCurrent = 1  # Places comparison differences in the current document. Default.
    wdCompareTargetNew = 2  # Places comparison differences in a new document.
}
enum WdMergeTarget {
    # wdMergeTargetSelected = 0  # Merge into selected document.
    # wdMergeTargetCurrent = 1  # Merge into current document.
    wdMergeTargetNew = 2  # Merge into new document.
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
enum WdCompareDestination {
    # wdCompareDestinationOriginal = 0  #Tracks the differences between the two files using tracked changes in the original document.
    wdCompareDestinationRevised = 1  #Tracks the differences between the two files using tracked changes in the revised document.
    wdCompareDestinationNew = 2  #Creates a new file and tracks the differences between the original document and the revised document using tracked changes.
}
