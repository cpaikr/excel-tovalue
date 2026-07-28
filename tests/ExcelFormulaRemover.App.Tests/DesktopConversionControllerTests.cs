namespace ExcelFormulaRemover.App.Tests;

using ExcelFormulaRemover.App;
using ExcelFormulaRemover.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

[TestClass]
public sealed class DesktopConversionControllerTests
{
    [TestMethod]
    public async Task CompletedConversionReturnsResultAndResetsState()
    {
        FormulaRemovalBatchResult expected = new([Success("Book.xlsx", "Output.xlsx")]);
        using DesktopConversionController controller = new((_, progress, cancellationToken) =>
        {
            Assert.IsNull(progress);
            Assert.IsFalse(cancellationToken.IsCancellationRequested);
            return Task.FromResult(expected);
        });

        DesktopConversionOutcome outcome = await controller.RunAsync(CreateRequest());

        Assert.AreSame(expected, outcome.Result);
        Assert.IsFalse(outcome.WasCanceled);
        Assert.IsNull(outcome.Failure);
        Assert.IsFalse(controller.IsRunning);
        Assert.IsFalse(controller.IsCancellationRequested);
    }

    [TestMethod]
    public async Task CancellationRequestIsIdempotentAndReturnsPartialResult()
    {
        TaskCompletionSource<bool> started = new(
            TaskCreationOptions.RunContinuationsAsynchronously);
        FormulaRemovalBatchResult partial = new([Success("First.xlsx", "First-output.xlsx")]);
        using DesktopConversionController controller = new(async (_, _, cancellationToken) =>
        {
            started.SetResult(true);
            try
            {
                await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
                return new FormulaRemovalBatchResult([]);
            }
            catch (OperationCanceledException exception)
            {
                throw new FormulaRemovalCanceledException(partial, cancellationToken, exception);
            }
        });

        Task<DesktopConversionOutcome> pending = controller.RunAsync(CreateRequest());
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
        TaskCompletionSource<bool> started = new(
            TaskCreationOptions.RunContinuationsAsynchronously);
        using DesktopConversionController controller = new(async (_, _, cancellationToken) =>
        {
            started.SetResult(true);
            await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            return new FormulaRemovalBatchResult([]);
        });

        Task<DesktopConversionOutcome> pending = controller.RunAsync(CreateRequest());
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
        InvalidOperationException expected = new("broken converter");
        using DesktopConversionController controller = new((_, _, _) =>
            Task.FromException<FormulaRemovalBatchResult>(expected));

        DesktopConversionOutcome outcome = await controller.RunAsync(CreateRequest());

        Assert.AreSame(expected, outcome.Failure);
        Assert.IsFalse(outcome.WasCanceled);
        Assert.IsEmpty(outcome.Result.Files);
        Assert.IsFalse(controller.IsRunning);
    }

    [TestMethod]
    public async Task ConcurrentConversionIsRejectedWithoutReplacingTheActiveOne()
    {
        TaskCompletionSource<bool> started = new(
            TaskCreationOptions.RunContinuationsAsynchronously);
        TaskCompletionSource<FormulaRemovalBatchResult> finish = new(
            TaskCreationOptions.RunContinuationsAsynchronously);
        using DesktopConversionController controller = new(async (_, _, _) =>
        {
            started.SetResult(true);
            return await finish.Task;
        });

        Task<DesktopConversionOutcome> first = controller.RunAsync(CreateRequest());
        await started.Task;

        await Assert.ThrowsExactlyAsync<InvalidOperationException>(() =>
            controller.RunAsync(CreateRequest()));

        FormulaRemovalBatchResult expected = new([Success("Book.xlsx", "Output.xlsx")]);
        finish.SetResult(expected);
        DesktopConversionOutcome outcome = await first;

        Assert.AreSame(expected, outcome.Result);
        Assert.IsFalse(controller.IsRunning);
    }

    private static FormulaRemovalRequest CreateRequest() =>
        new(["Book.xlsx"], Path.GetTempPath());

    private static FormulaRemovalFileResult Success(string inputPath, string outputPath) =>
        new(
            inputPath,
            outputPath,
            FormulaRemovalFailureKind.None,
            null,
            1,
            1,
            0,
            0,
            Array.Empty<string>());
}
