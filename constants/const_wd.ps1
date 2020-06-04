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
