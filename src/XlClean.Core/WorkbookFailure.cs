namespace XlClean.Core;

using System.Diagnostics.CodeAnalysis;

/// <summary>
/// Categorizes a workbook that could not be processed. The CLI reports each member as a stable
/// identifier, so members are added, never renamed.
/// </summary>
public enum WorkbookFailureKind
{
    /// <summary>The workbook was processed successfully.</summary>
    None,

    /// <summary>The input path, extension, or request item is invalid.</summary>
    InvalidInput,

    /// <summary>A target formula has no cached value that can replace it.</summary>
    MissingCachedValue,

    /// <summary>A multi-cell or table formula definition crosses the requested start row.</summary>
    FormulaRangeCrossesBoundary,

    /// <summary>The input is not a readable spreadsheet package.</summary>
    InvalidWorkbook,

    /// <summary>The input or output could not be accessed.</summary>
    FileAccess,

    /// <summary>An unexpected error occurred.</summary>
    Unexpected,

    /// <summary>
    /// The workbook holds content that may reference what the operation rewrites, in a form the
    /// operation does not rewrite, so it cannot prove the result equivalent.
    /// </summary>
    UnsupportedWorkbookContent,
}

/// <summary>
/// A file-level failure an operation detected itself, carrying the kind to report.
/// </summary>
/// <remarks>
/// The pipeline turns every other exception into a kind by its type. Operations throw a subclass
/// when the type alone would misreport the cause.
/// </remarks>
[SuppressMessage(
    "Design",
    "CA1032:Implement standard exception constructors",
    Justification = "Every failure is created with the kind it reports.")]
internal abstract class WorkbookFailureException(WorkbookFailureKind kind, string message)
    : Exception(message)
{
    public WorkbookFailureKind Kind { get; } = kind;
}
