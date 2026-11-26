namespace XlClean.Core;

/// <summary>
/// Identifies the current phase of a batch.
/// </summary>
public enum WorkbookStage
{
    /// <summary>A workbook is about to be copied and opened.</summary>
    StartingFile,

    /// <summary>A worksheet is being transformed.</summary>
    ProcessingWorksheet,

    /// <summary>A workbook attempt has finished.</summary>
    CompletedFile,
}

/// <summary>
/// Reports coarse, single-pass batch progress without pre-scanning workbooks.
/// </summary>
public sealed record WorkbookProgress(
    int CompletedFiles,
    int TotalFiles,
    string InputPath,
    string? WorksheetName,
    WorkbookStage Stage);
