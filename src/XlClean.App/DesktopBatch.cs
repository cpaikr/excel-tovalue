namespace XlClean.App;

using System.Diagnostics.CodeAnalysis;
using XlClean.Core;

/// <summary>The cleanup the desktop app runs on every file in a batch.</summary>
internal enum DesktopOperation
{
    /// <summary>Replace formulas with their last saved values (<c>xlclean values</c>).</summary>
    Values,

    /// <summary>Remove unused styles and merge identical cell formats (<c>xlclean styles</c>).</summary>
    Styles,
}

/// <summary>One input's outcome, in the shape the form and summary need for either operation.</summary>
/// <param name="Changes">Style counts for a published style cleanup; null otherwise.</param>
internal sealed record DesktopFileResult(
    string InputPath,
    string? OutputPath,
    WorkbookFailureKind FailureKind,
    string? ErrorMessage,
    StyleCleanupChanges? Changes,
    IReadOnlyList<string> Warnings)
{
    public bool Succeeded => FailureKind == WorkbookFailureKind.None;
}

/// <summary>Per-file outcomes of one desktop batch in request order.</summary>
internal sealed class DesktopBatchResult(IEnumerable<DesktopFileResult> files)
{
    public static DesktopBatchResult Empty { get; } = new([]);

    public IReadOnlyList<DesktopFileResult> Files { get; } = Array.AsReadOnly(files.ToArray());

    public int SucceededCount => Files.Count(file => file.Succeeded);

    public int FailedCount => Files.Count - SucceededCount;
}

/// <summary>Cancellation of a desktop batch that keeps the files completed before it.</summary>
[SuppressMessage(
    "Design",
    "CA1032:Implement standard exception constructors",
    Justification = "Only the desktop batch adapters create this result-carrying exception.")]
internal sealed class DesktopBatchCanceledException(
    DesktopBatchResult completedResult,
    OperationCanceledException innerException)
    : OperationCanceledException(innerException.Message, innerException, innerException.CancellationToken)
{
    public DesktopBatchResult CompletedResult { get; } = completedResult;
}

/// <summary>Starts one batch; the request is already bound.</summary>
internal delegate Task<DesktopBatchResult> DesktopBatchRunner(
    IProgress<WorkbookProgress>? progress,
    CancellationToken cancellationToken);

/// <summary>Binds each engine to <see cref="DesktopBatchRunner"/> and maps its results.</summary>
internal static class DesktopBatchJobs
{
    public static DesktopBatchRunner RemoveFormulas(
        FormulaRemovalRequest request,
        Func<FormulaRemovalRequest, IProgress<WorkbookProgress>?, CancellationToken, Task<FormulaRemovalBatchResult>> convert) =>
        async (progress, cancellationToken) =>
        {
            try
            {
                return From(await convert(request, progress, cancellationToken));
            }
            catch (FormulaRemovalCanceledException exception)
            {
                throw new DesktopBatchCanceledException(From(exception.CompletedResult), exception);
            }
        };

    public static DesktopBatchRunner CleanStyles(
        StyleCleanupRequest request,
        Func<StyleCleanupRequest, IProgress<WorkbookProgress>?, CancellationToken, Task<StyleCleanupBatchResult>> clean) =>
        async (progress, cancellationToken) =>
        {
            try
            {
                return From(await clean(request, progress, cancellationToken));
            }
            catch (StyleCleanupCanceledException exception)
            {
                throw new DesktopBatchCanceledException(From(exception.CompletedResult), exception);
            }
        };

    internal static DesktopBatchResult From(FormulaRemovalBatchResult result) =>
        new(result.Files.Select(file => new DesktopFileResult(
            file.InputPath,
            file.OutputPath,
            file.FailureKind,
            file.ErrorMessage,
            Changes: null,
            file.Warnings ?? Array.Empty<string>())));

    internal static DesktopBatchResult From(StyleCleanupBatchResult result) =>
        new(result.Files.Select(file => new DesktopFileResult(
            file.InputPath,
            file.OutputPath,
            file.FailureKind,
            file.ErrorMessage,
            file.Changes,
            file.Warnings)));
}
