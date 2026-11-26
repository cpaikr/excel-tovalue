namespace XlClean.Core;

using System.Diagnostics.CodeAnalysis;
using System.IO.Packaging;
using XlClean.Core.Styles;

/// <summary>
/// Describes one immutable batch of style cleanups.
/// </summary>
public sealed class StyleCleanupRequest
{
    /// <summary>Initializes a new style-cleanup request.</summary>
    /// <param name="inputPaths">Workbook paths to process in order.</param>
    /// <param name="outputDirectory">Directory in which outputs are created.</param>
    /// <exception cref="ArgumentException">
    /// An input path is blank, the input sequence is empty, or the output directory is invalid.
    /// </exception>
    public StyleCleanupRequest(IEnumerable<string> inputPaths, string outputDirectory)
    {
        ArgumentNullException.ThrowIfNull(inputPaths);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);

        string[] inputs = inputPaths.ToArray();
        if (inputs.Length == 0)
        {
            throw new ArgumentException("At least one input workbook is required.", nameof(inputPaths));
        }

        if (inputs.Any(string.IsNullOrWhiteSpace))
        {
            throw new ArgumentException("Input workbook paths cannot be blank.", nameof(inputPaths));
        }

        InputPaths = Array.AsReadOnly(inputs);
        OutputDirectory = Path.GetFullPath(outputDirectory);
    }

    /// <summary>Gets the workbook paths to process in order.</summary>
    public IReadOnlyList<string> InputPaths { get; }

    /// <summary>Gets the destination directory.</summary>
    public string OutputDirectory { get; }
}

/// <summary>What a style cleanup changed in one published workbook.</summary>
/// <param name="CellStylesBefore">Named styles (<c>cellStyles</c>) in the source.</param>
/// <param name="CellStylesAfter">Named styles in the output.</param>
/// <param name="CellFormatsBefore">Cell formats (<c>cellXfs</c>) in the source.</param>
/// <param name="CellFormatsAfter">Cell formats in the output.</param>
/// <param name="MergedCellFormats">Used cell formats folded into an identical one.</param>
public sealed record StyleCleanupChanges(
    int CellStylesBefore,
    int CellStylesAfter,
    int CellFormatsBefore,
    int CellFormatsAfter,
    int MergedCellFormats);

/// <summary>The observable outcome for one input workbook.</summary>
/// <param name="Changes">What changed; null exactly when the workbook failed.</param>
public sealed record StyleCleanupFileResult(
    string InputPath,
    string? OutputPath,
    WorkbookFailureKind FailureKind,
    string? ErrorMessage,
    StyleCleanupChanges? Changes,
    IReadOnlyList<string> Warnings)
{
    /// <summary>Gets a value indicating whether a complete output was published.</summary>
    public bool Succeeded => FailureKind == WorkbookFailureKind.None;
}

/// <summary>Summarizes a completed batch while retaining every per-file result.</summary>
public sealed class StyleCleanupBatchResult
{
    /// <summary>Initializes a batch result.</summary>
    /// <param name="files">Results in the same order as the request.</param>
    public StyleCleanupBatchResult(IEnumerable<StyleCleanupFileResult> files)
    {
        ArgumentNullException.ThrowIfNull(files);
        Files = Array.AsReadOnly(files.ToArray());
    }

    /// <summary>Gets all per-file outcomes in request order.</summary>
    public IReadOnlyList<StyleCleanupFileResult> Files { get; }

    /// <summary>Gets the number of published outputs.</summary>
    public int SucceededCount => Files.Count(file => file.Succeeded);

    /// <summary>Gets the number of failed inputs.</summary>
    public int FailedCount => Files.Count - SucceededCount;
}

/// <summary>Reports cancellation while retaining results for files that finished before it.</summary>
[SuppressMessage(
    "Design",
    "CA1032:Implement standard exception constructors",
    Justification = "Only the engine creates this result-carrying cancellation exception.")]
public sealed class StyleCleanupCanceledException : OperationCanceledException
{
    /// <summary>Initializes a cancellation result for a partially completed batch.</summary>
    /// <param name="completedResult">Files completed before cancellation was observed.</param>
    /// <param name="cancellationToken">The token that canceled the operation.</param>
    /// <param name="innerException">The cancellation exception raised inside the engine.</param>
    public StyleCleanupCanceledException(
        StyleCleanupBatchResult completedResult,
        CancellationToken cancellationToken,
        Exception? innerException = null)
        : base("The style-cleanup batch was canceled.", innerException, cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(completedResult);
        CompletedResult = completedResult;
    }

    /// <summary>Gets the immutable results published before cancellation.</summary>
    public StyleCleanupBatchResult CompletedResult { get; }
}

/// <summary>
/// Removes unused custom named styles and unused cell formats, and merges identical cell formats,
/// the way XLStylesTool does, without changing how any cell looks.
/// </summary>
/// <remarks>
/// The rules are in plans/style-cleanup.md. The safety guarantees belong to
/// <see cref="WorkbookBatchPipeline"/>; every output is also checked by
/// <see cref="OpenXmlStyleVerifier"/>, which is independent of the planner, before it is published.
/// </remarks>
public static class StyleCleanupEngine
{
    /// <summary>
    /// Non-built-in cell and style format records Excel accepts together ([MS-XLSB] 2.4.878).
    /// </summary>
    public const int FormatRecordLimit = 0xFF96;

    private const string VbaProjectRelationship =
        "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

    /// <summary>Cleans all requested workbooks sequentially on a worker thread.</summary>
    /// <param name="request">Immutable inputs.</param>
    /// <param name="progress">Optional coarse progress sink.</param>
    /// <param name="cancellationToken">Stops before publication and removes temporary files.</param>
    /// <returns>One explicit result for every requested input path.</returns>
    /// <exception cref="IOException">The output directory cannot be created or resolved.</exception>
    /// <exception cref="UnauthorizedAccessException">The output directory cannot be accessed.</exception>
    /// <exception cref="StyleCleanupCanceledException">
    /// Cancellation was observed after the batch started. The exception retains completed results.
    /// </exception>
    /// <exception cref="OperationCanceledException">Cancellation prevented the worker from starting.</exception>
    public static Task<StyleCleanupBatchResult> CleanAsync(
        StyleCleanupRequest request,
        IProgress<WorkbookProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        return Task.Run(
            () =>
            {
                try
                {
                    return ToBatchResult(WorkbookBatchPipeline.Run(
                        request.InputPaths,
                        request.OutputDirectory,
                        new StyleCleanupOperation(),
                        progress,
                        cancellationToken));
                }
                catch (WorkbookBatchCanceledException<StyleCleanupReport> exception)
                {
                    throw new StyleCleanupCanceledException(
                        ToBatchResult(exception.Completed),
                        exception.CancellationToken,
                        exception.InnerException);
                }
            },
            cancellationToken);
    }

    private static StyleCleanupBatchResult ToBatchResult(IReadOnlyList<WorkbookFileOutcome<StyleCleanupReport>> outcomes) =>
        new(outcomes.Select(outcome => new StyleCleanupFileResult(
            outcome.InputPath,
            outcome.OutputPath,
            outcome.FailureKind,
            outcome.ErrorMessage,
            outcome.Changes?.Changes,
            outcome.Changes?.Warnings ?? Array.Empty<string>())));

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> Warnings(string workbookPath, StyleSheetModel model, StylePlan plan)
    {
        List<string> warnings = [];
        int removedCustomStyles = model.NamedStyles.Count - plan.KeptNamedStyles.Count;
        if (removedCustomStyles > 0 && HasVbaProject(workbookPath))
        {
            warnings.Add(
                $"The workbook has macros and {removedCustomStyles} unused custom cell styles were removed; " +
                "a macro that applies a style by name may no longer find it.");
        }

        int formatRecords = plan.KeptCellFormatCount + plan.CellStyleFormatMap.Count(index => index is not null);
        if (formatRecords > FormatRecordLimit)
        {
            warnings.Add(
                $"The cleaned workbook still has {formatRecords} cell and style formats, more than the " +
                $"{FormatRecordLimit} Excel supports; it may still report too many different cell formats.");
        }

        return warnings.AsReadOnly();
    }

    private static bool HasVbaProject(string workbookPath)
    {
        using Package package = Package.Open(workbookPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        return OpenXmlStyleReader.GetWorkbookPart(package).GetRelationshipsByType(VbaProjectRelationship).Any();
    }

    /// <summary>What the operation reports for a published workbook.</summary>
    internal sealed record StyleCleanupReport(StyleCleanupChanges Changes, IReadOnlyList<string> Warnings);

    private sealed class StyleCleanupOperation : IWorkbookOperation<StyleCleanupReport>
    {
        public string OutputNameSuffix => "_스타일정리";

        public StyleCleanupReport Transform(WorkbookFile file, CancellationToken cancellationToken)
        {
            bool binary = file.Format == WorkbookFormat.Binary;
            (StyleSheetModel model, StyleUsage usage) = binary
                ? XlsbStyleCleanup.Read(file.Path, cancellationToken)
                : OpenXmlStyleReader.Read(file.Path, cancellationToken);
            StylePlan plan = StylePlanner.Plan(model, usage);
            IReadOnlyList<string> warnings = Warnings(file.Path, model, plan);
            if (binary)
            {
                XlsbStyleCleanup.Apply(file.Path, plan, cancellationToken);
            }
            else
            {
                OpenXmlStyleRewriter.Apply(file.Path, plan, cancellationToken);
            }

            return new StyleCleanupReport(
                new StyleCleanupChanges(
                    model.NamedStyles.Count,
                    plan.KeptNamedStyles.Count,
                    model.CellFormats.Count,
                    plan.KeptCellFormatCount,
                    plan.MergedCellFormats),
                warnings);
        }

        public void Verify(WorkbookFile file, StyleCleanupReport changes, CancellationToken cancellationToken)
        {
            if (file.Format == WorkbookFormat.Binary)
            {
                XlsbStyleVerifier.Verify(file.InputPath, file.Path, cancellationToken);
            }
            else
            {
                OpenXmlStyleVerifier.Verify(file.InputPath, file.Path, cancellationToken);
            }
        }
    }
}
