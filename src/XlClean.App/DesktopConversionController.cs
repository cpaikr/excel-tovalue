namespace XlClean.App;

using System.Diagnostics.CodeAnalysis;
using XlClean.Core;

internal delegate Task<FormulaRemovalBatchResult> DesktopFormulaRemovalConverter(
    FormulaRemovalRequest request,
    IProgress<WorkbookProgress>? progress,
    CancellationToken cancellationToken);

/// <summary>
/// Owns the lifetime and cancellation state of one desktop conversion.
/// </summary>
internal sealed class DesktopConversionController(
    DesktopFormulaRemovalConverter converter) : IDisposable
{
    private readonly DesktopFormulaRemovalConverter converter =
        converter ?? throw new ArgumentNullException(nameof(converter));
    private CancellationTokenSource? cancellation;

    /// <summary>
    /// Gets a value indicating whether a conversion has not yet returned.
    /// </summary>
    public bool IsRunning => cancellation is not null;

    /// <summary>
    /// Gets a value indicating whether the active conversion was asked to stop.
    /// </summary>
    public bool IsCancellationRequested => cancellation?.IsCancellationRequested == true;

    /// <summary>
    /// Runs one conversion and turns every engine terminal state into data for the form.
    /// </summary>
    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "The desktop event loop must render unexpected converter failures instead of terminating.")]
    public async Task<DesktopConversionOutcome> RunAsync(
        FormulaRemovalRequest request,
        IProgress<WorkbookProgress>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(request);
        if (cancellation is not null)
        {
            throw new InvalidOperationException("A desktop conversion is already running.");
        }

        CancellationTokenSource currentCancellation = new();
        cancellation = currentCancellation;

        try
        {
            FormulaRemovalBatchResult result = await converter(
                request,
                progress,
                currentCancellation.Token);
            return DesktopConversionOutcome.Completed(result);
        }
        catch (FormulaRemovalCanceledException exception)
        {
            return DesktopConversionOutcome.Canceled(exception.CompletedResult);
        }
        catch (OperationCanceledException) when (currentCancellation.IsCancellationRequested)
        {
            return DesktopConversionOutcome.Canceled(new FormulaRemovalBatchResult([]));
        }
        catch (Exception exception)
        {
            return DesktopConversionOutcome.Failed(exception);
        }
        finally
        {
            currentCancellation.Dispose();
            if (ReferenceEquals(cancellation, currentCancellation))
            {
                cancellation = null;
            }
        }
    }

    /// <summary>
    /// Requests cooperative cancellation once.
    /// </summary>
    /// <returns><see langword="true"/> only for the first active request.</returns>
    public bool RequestCancellation()
    {
        if (cancellation is null || cancellation.IsCancellationRequested)
        {
            return false;
        }

        cancellation.Cancel();
        return true;
    }

    public void Dispose()
    {
        if (cancellation is null)
        {
            return;
        }

        cancellation.Cancel();
        cancellation.Dispose();
        cancellation = null;
    }
}

/// <summary>
/// Describes how one desktop conversion stopped.
/// </summary>
internal sealed record DesktopConversionOutcome(
    FormulaRemovalBatchResult Result,
    bool WasCanceled,
    Exception? Failure)
{
    public static DesktopConversionOutcome Completed(FormulaRemovalBatchResult result) =>
        new(result, WasCanceled: false, Failure: null);

    public static DesktopConversionOutcome Canceled(FormulaRemovalBatchResult result) =>
        new(result, WasCanceled: true, Failure: null);

    public static DesktopConversionOutcome Failed(Exception exception) =>
        new(new FormulaRemovalBatchResult([]), WasCanceled: false, exception);
}
