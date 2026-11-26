namespace XlClean.Core;

using System.Diagnostics.CodeAnalysis;

/// <summary>
/// Categorizes a workbook that could not be converted.
/// </summary>
public enum FormulaRemovalFailureKind
{
    /// <summary>The workbook was converted successfully.</summary>
    None,

    /// <summary>The input path, extension, or request item is invalid.</summary>
    InvalidInput,

    /// <summary>A target formula has no cached value that can replace it.</summary>
    MissingCachedValue,

    /// <summary>A multi-cell or table formula definition crosses the requested start row.</summary>
    FormulaRangeCrossesBoundary,

    /// <summary>The input is not a readable OpenXML spreadsheet package.</summary>
    InvalidWorkbook,

    /// <summary>The input or output could not be accessed.</summary>
    FileAccess,

    /// <summary>An unexpected conversion error occurred.</summary>
    Unexpected,
}

/// <summary>
/// Identifies the current phase of a batch conversion.
/// </summary>
public enum FormulaRemovalStage
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
public sealed record FormulaRemovalProgress(
    int CompletedFiles,
    int TotalFiles,
    string InputPath,
    string? WorksheetName,
    FormulaRemovalStage Stage);

/// <summary>
/// Describes the observable outcome for one input workbook.
/// </summary>
public sealed record FormulaRemovalFileResult(
    string InputPath,
    string? OutputPath,
    FormulaRemovalFailureKind FailureKind,
    string? ErrorMessage,
    int RemovedFormulaCount,
    int ProcessedSheetCount,
    int ExcludedSheetCount,
    int RemovedTableFormulaDefinitionCount,
    IReadOnlyList<string> Warnings)
{
    /// <summary>
    /// Gets a value indicating whether a complete output was published.
    /// </summary>
    public bool Succeeded => FailureKind == FormulaRemovalFailureKind.None;
}

/// <summary>
/// Summarizes a completed batch while retaining every per-file result.
/// </summary>
public sealed class FormulaRemovalBatchResult
{
    /// <summary>
    /// Initializes a batch result.
    /// </summary>
    /// <param name="files">Results in the same order as the request.</param>
    public FormulaRemovalBatchResult(IEnumerable<FormulaRemovalFileResult> files)
    {
        ArgumentNullException.ThrowIfNull(files);
        Files = Array.AsReadOnly(files.ToArray());
    }

    /// <summary>
    /// Gets all per-file outcomes in request order.
    /// </summary>
    public IReadOnlyList<FormulaRemovalFileResult> Files { get; }

    /// <summary>
    /// Gets the number of published outputs.
    /// </summary>
    public int SucceededCount => Files.Count(file => file.Succeeded);

    /// <summary>
    /// Gets the number of failed inputs.
    /// </summary>
    public int FailedCount => Files.Count - SucceededCount;
}

/// <summary>
/// Reports cancellation while retaining results for files that finished before it.
/// </summary>
[SuppressMessage(
    "Design",
    "CA1032:Implement standard exception constructors",
    Justification = "Only the engine creates this result-carrying cancellation exception.")]
public sealed class FormulaRemovalCanceledException : OperationCanceledException
{
    /// <summary>
    /// Initializes a cancellation result for a partially completed batch.
    /// </summary>
    /// <param name="completedResult">Files completed before cancellation was observed.</param>
    /// <param name="cancellationToken">The token that canceled the operation.</param>
    /// <param name="innerException">The cancellation exception raised inside the engine.</param>
    public FormulaRemovalCanceledException(
        FormulaRemovalBatchResult completedResult,
        CancellationToken cancellationToken,
        Exception? innerException = null)
        : base("The formula-removal batch was canceled.", innerException, cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(completedResult);
        CompletedResult = completedResult;
    }

    /// <summary>
    /// Gets the immutable results published before cancellation.
    /// </summary>
    public FormulaRemovalBatchResult CompletedResult { get; }
}
