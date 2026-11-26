namespace XlClean.App;

using System.Diagnostics.CodeAnalysis;
using XlClean.Core;

/// <summary>
/// Owns the lifetime and cancellation state of one desktop batch, whichever operation it runs.
/// </summary>
internal sealed class DesktopConversionController : IDisposable
{
    private CancellationTokenSource? cancellation;

    /// <summary>
    /// Gets a value indicating whether a batch has not yet returned.
    /// </summary>
    public bool IsRunning => cancellation is not null;

    /// <summary>
    /// Gets a value indicating whether the active batch was asked to stop.
    /// </summary>
    public bool IsCancellationRequested => cancellation?.IsCancellationRequested == true;

    /// <summary>
    /// Runs one batch and turns every engine terminal state into data for the form.
    /// </summary>
    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "The desktop event loop must render unexpected engine failures instead of terminating.")]
    public async Task<DesktopConversionOutcome> RunAsync(
        DesktopBatchRunner runner,
        IProgress<WorkbookProgress>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(runner);
        if (cancellation is not null)
        {
            throw new InvalidOperationException("A desktop batch is already running.");
        }

        CancellationTokenSource currentCancellation = new();
        cancellation = currentCancellation;

        try
        {
            DesktopBatchResult result = await runner(progress, currentCancellation.Token);
            return DesktopConversionOutcome.Completed(result);
        }
        catch (DesktopBatchCanceledException exception)
        {
            return DesktopConversionOutcome.Canceled(exception.CompletedResult);
        }
        catch (OperationCanceledException) when (currentCancellation.IsCancellationRequested)
        {
            return DesktopConversionOutcome.Canceled(DesktopBatchResult.Empty);
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
/// Describes how one desktop batch stopped.
/// </summary>
internal sealed record DesktopConversionOutcome(
    DesktopBatchResult Result,
    bool WasCanceled,
    Exception? Failure)
{
    public static DesktopConversionOutcome Completed(DesktopBatchResult result) =>
        new(result, WasCanceled: false, Failure: null);

    public static DesktopConversionOutcome Canceled(DesktopBatchResult result) =>
        new(result, WasCanceled: true, Failure: null);

    public static DesktopConversionOutcome Failed(Exception exception) =>
        new(DesktopBatchResult.Empty, WasCanceled: false, exception);
}
