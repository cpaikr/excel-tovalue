namespace XlClean.App.Tests;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlClean.App;
using XlClean.Core;

[TestClass]
public sealed class DesktopConversionControllerTests
{
    [TestMethod]
    public async Task CompletedBatchReturnsResultAndResetsState()
    {
        DesktopBatchResult expected = new([Success("Book.xlsx", "Output.xlsx")]);
        using DesktopConversionController controller = new();

        DesktopConversionOutcome outcome = await controller.RunAsync((progress, cancellationToken) =>
        {
            Assert.IsNull(progress);
            Assert.IsFalse(cancellationToken.IsCancellationRequested);
            return Task.FromResult(expected);
        });

        Assert.AreSame(expected, outcome.Result);
        Assert.IsFalse(outcome.WasCanceled);
        Assert.IsNull(outcome.Failure);
        Assert.IsFalse(controller.IsRunning);
        Assert.IsFalse(controller.IsCancellationRequested);
    }

    [TestMethod]
    public async Task CancellationRequestIsIdempotentAndReturnsPartialResult()
    {
        TaskCompletionSource<bool> started = new(TaskCreationOptions.RunContinuationsAsynchronously);
        DesktopBatchResult partial = new([Success("First.xlsx", "First-output.xlsx")]);
        using DesktopConversionController controller = new();

        Task<DesktopConversionOutcome> pending = controller.RunAsync(async (_, cancellationToken) =>
        {
            started.SetResult(true);
            try
            {
                await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
                return DesktopBatchResult.Empty;
            }
            catch (OperationCanceledException exception)
            {
                throw new DesktopBatchCanceledException(partial, exception);
            }
        });
        await started.Task;

        Assert.IsTrue(controller.IsRunning);
        Assert.IsTrue(controller.RequestCancellation());
        Assert.IsTrue(controller.IsCancellationRequested);
        Assert.IsFalse(controller.RequestCancellation());

        DesktopConversionOutcome outcome = await pending;

        Assert.AreSame(partial, outcome.Result);
        Assert.IsTrue(outcome.WasCanceled);
        Assert.IsNull(outcome.Failure);
        Assert.IsFalse(controller.IsRunning);
    }

    [TestMethod]
    public async Task PlainCancellationProducesAnEmptyCanceledResult()
    {
        TaskCompletionSource<bool> started = new(TaskCreationOptions.RunContinuationsAsynchronously);
        using DesktopConversionController controller = new();

        Task<DesktopConversionOutcome> pending = controller.RunAsync(async (_, cancellationToken) =>
        {
            started.SetResult(true);
            await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            return DesktopBatchResult.Empty;
        });
        await started.Task;
        controller.RequestCancellation();

        DesktopConversionOutcome outcome = await pending;

        Assert.IsTrue(outcome.WasCanceled);
        Assert.IsEmpty(outcome.Result.Files);
        Assert.IsNull(outcome.Failure);
    }

    [TestMethod]
    public async Task UnexpectedFailureBecomesARecoverableOutcome()
    {
        InvalidOperationException expected = new("broken engine");
        using DesktopConversionController controller = new();

        DesktopConversionOutcome outcome = await controller.RunAsync((_, _) =>
            Task.FromException<DesktopBatchResult>(expected));

        Assert.AreSame(expected, outcome.Failure);
        Assert.IsFalse(outcome.WasCanceled);
        Assert.IsEmpty(outcome.Result.Files);
        Assert.IsFalse(controller.IsRunning);
    }

    [TestMethod]
    public async Task ConcurrentBatchIsRejectedWithoutReplacingTheActiveOne()
    {
        TaskCompletionSource<bool> started = new(TaskCreationOptions.RunContinuationsAsynchronously);
        TaskCompletionSource<DesktopBatchResult> finish = new(TaskCreationOptions.RunContinuationsAsynchronously);
        using DesktopConversionController controller = new();

        Task<DesktopConversionOutcome> first = controller.RunAsync(async (_, _) =>
        {
            started.SetResult(true);
            return await finish.Task;
        });
        await started.Task;

        await Assert.ThrowsExactlyAsync<InvalidOperationException>(() =>
            controller.RunAsync((_, _) => Task.FromResult(DesktopBatchResult.Empty)));

        DesktopBatchResult expected = new([Success("Book.xlsx", "Output.xlsx")]);
        finish.SetResult(expected);
        DesktopConversionOutcome outcome = await first;

        Assert.AreSame(expected, outcome.Result);
        Assert.IsFalse(controller.IsRunning);
    }

    [TestMethod]
    public async Task FormulaRemovalResultsAndCancellationMapToTheDesktopShape()
    {
        FormulaRemovalRequest request = new(["Book.xlsx"], Path.GetTempPath());
        FormulaRemovalBatchResult partial = new(
        [
            new FormulaRemovalFileResult(
                "Book.xlsx",
                "Book_수식제거.xlsx",
                WorkbookFailureKind.None,
                null,
                3,
                1,
                0,
                0,
                ["manual calculation"]),
        ]);
        using DesktopConversionController controller = new();

        DesktopConversionOutcome outcome = await controller.RunAsync(DesktopBatchJobs.RemoveFormulas(
            request,
            (actual, _, cancellationToken) =>
            {
                Assert.AreSame(request, actual);
                throw new FormulaRemovalCanceledException(partial, cancellationToken);
            }));

        Assert.IsTrue(outcome.WasCanceled);
        DesktopFileResult file = outcome.Result.Files.Single();
        Assert.AreEqual("Book_수식제거.xlsx", file.OutputPath);
        Assert.IsNull(file.Changes);
        Assert.AreEqual("manual calculation", file.Warnings.Single());
    }

    [TestMethod]
    public async Task StyleCleanupResultsKeepTheirChanges()
    {
        StyleCleanupRequest request = new(["Book.xlsx", "Bad.xlsx"], Path.GetTempPath());
        StyleCleanupChanges changes = new(10, 2, 8, 3, 1);
        StyleCleanupBatchResult result = new(
        [
            new StyleCleanupFileResult(
                "Book.xlsx",
                "Book_스타일정리.xlsx",
                WorkbookFailureKind.None,
                null,
                changes,
                Array.Empty<string>()),
            new StyleCleanupFileResult(
                "Bad.xlsx",
                null,
                WorkbookFailureKind.UnsupportedWorkbookContent,
                "macro sheet",
                null,
                Array.Empty<string>()),
        ]);
        using DesktopConversionController controller = new();

        DesktopConversionOutcome outcome = await controller.RunAsync(DesktopBatchJobs.CleanStyles(
            request,
            (actual, _, _) =>
            {
                Assert.AreSame(request, actual);
                return Task.FromResult(result);
            }));

        Assert.IsFalse(outcome.WasCanceled);
        Assert.AreEqual(1, outcome.Result.SucceededCount);
        Assert.AreEqual(1, outcome.Result.FailedCount);
        Assert.AreSame(changes, outcome.Result.Files[0].Changes);
        Assert.AreEqual(WorkbookFailureKind.UnsupportedWorkbookContent, outcome.Result.Files[1].FailureKind);
        Assert.AreEqual("macro sheet", outcome.Result.Files[1].ErrorMessage);
    }

    [TestMethod]
    public async Task StyleCleanupCancellationKeepsCompletedFiles()
    {
        StyleCleanupRequest request = new(["Book.xlsx"], Path.GetTempPath());
        StyleCleanupBatchResult partial = new(
        [
            new StyleCleanupFileResult(
                "Book.xlsx",
                "Book_스타일정리.xlsx",
                WorkbookFailureKind.None,
                null,
                new StyleCleanupChanges(3, 1, 3, 1, 0),
                Array.Empty<string>()),
        ]);
        using DesktopConversionController controller = new();

        DesktopConversionOutcome outcome = await controller.RunAsync(DesktopBatchJobs.CleanStyles(
            request,
            (_, _, cancellationToken) => throw new StyleCleanupCanceledException(partial, cancellationToken)));

        Assert.IsTrue(outcome.WasCanceled);
        Assert.AreEqual("Book_스타일정리.xlsx", outcome.Result.Files.Single().OutputPath);
    }

    private static DesktopFileResult Success(string inputPath, string outputPath) =>
        new(inputPath, outputPath, WorkbookFailureKind.None, null, null, Array.Empty<string>());
}
