class WdSaveOptions:
    # > Prompt the user to save pending changes.
    # wdPromptToSaveChanges = -2
    # > Save pending changes automatically without prompting the user.
    # wdSaveChanges = -1
    # > Do not save pending changes.
    wdDoNotSaveChanges = 0  # noqa: N815


class WdUseFormattingFrom:
    # > Copy source formatting from the current item.
    # wdFormattingFromCurrent = 0
    # > Copy source formatting from the current selection.
    # wdFormattingFromSelected = 1
    # > Prompt the user for formatting to use.
    wdFormattingFromPrompt = 2  # noqa: N815


class WdWindowState:
    # > Normal.
    # wdWindowStateNormal = 0
    # > Maximized.
    wdWindowStateMaximize = 1  # noqa: N815
    # > Minimized.
    wdWindowStateMinimize = 2  # noqa: N815


class WdCompareDestination:
    # > Tracks the differences between the two files using tracked changes in the
    # > original document.
    # wdCompareDestinationOriginal = 0
    # > Tracks the differences between the two files using tracked changes in the
    # > revised document.
    wdCompareDestinationRevised = 1  # noqa: N815
    # > Creates a new file and tracks the differences between the original document
    # > and the revised document using tracked changes.
    wdCompareDestinationNew = 2  # noqa: N815
