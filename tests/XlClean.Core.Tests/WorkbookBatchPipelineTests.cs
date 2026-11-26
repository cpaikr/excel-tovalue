namespace XlClean.Core.Tests;

using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;

/// <summary>
/// The safety contract every operation inherits, exercised through an operation that knows
/// nothing about formulas.
/// </summary>
[TestClass]
public sealed class WorkbookBatchPipelineTests
{
    private string testDirectory = null!;

    [TestInitialize]
    public void Initialize()
    {
        testDirectory = Path.Combine(
            Path.GetTempPath(),
            "XlClean.Tests",
            Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        Directory.CreateDirectory(testDirectory);
    }

    [TestCleanup]
    public void Cleanup()
    {
        if (Directory.Exists(testDirectory))
        {
            Directory.Delete(testDirectory, recursive: true);
        }
    }

    [TestMethod]
    public void PublishesUnderTheOperationSuffixWithoutTouchingTheInput()
    {
        string input = CreateWorkbook("Book.xlsx");
        byte[] inputHash = Hash(input);
        string output = Path.Combine(testDirectory, "out");
        File.WriteAllText(Path.Combine(output.EnsureDirectory(), "Book_test.xlsx"), "existing");

        IReadOnlyList<WorkbookFileOutcome<Report>> outcomes = Run(new FakeOperation(), output, input, input);

        WorkbookFileOutcome<Report> first = outcomes[0];
        Assert.AreEqual(WorkbookFailureKind.None, first.FailureKind);
        Assert.AreEqual(Path.Combine(output, "Book_test (2).xlsx"), first.OutputPath);
        Assert.AreEqual("transformed", first.Changes?.Note);
        Assert.AreEqual(WorkbookFailureKind.InvalidInput, outcomes[1].FailureKind);
        Assert.IsNull(outcomes[1].Changes);
        Assert.AreEqual("existing", File.ReadAllText(Path.Combine(output, "Book_test.xlsx")));
        CollectionAssert.AreEqual(inputHash, Hash(input));
        AssertNoStagingFiles(output);
    }

    [TestMethod]
    public void OperationFailureKindIsReportedAndTheBatchContinues()
    {
        string failing = CreateWorkbook("Failing.xlsx");
        string passing = CreateWorkbook("Passing.xlsx");
        string output = Path.Combine(testDirectory, "out");
        FakeOperation operation = new()
        {
            TransformFailure = path => path.EndsWith("Failing.xlsx", StringComparison.Ordinal)
                ? new TestFailure(WorkbookFailureKind.MissingCachedValue, "No cache.")
                : null,
        };

        IReadOnlyList<WorkbookFileOutcome<Report>> outcomes = Run(operation, output, failing, passing);

        Assert.AreEqual(WorkbookFailureKind.MissingCachedValue, outcomes[0].FailureKind);
        Assert.AreEqual("No cache.", outcomes[0].ErrorMessage);
        Assert.IsNull(outcomes[0].OutputPath);
        Assert.IsNull(outcomes[0].Changes);
        Assert.AreEqual(WorkbookFailureKind.None, outcomes[1].FailureKind);
        Assert.AreEqual("Passing_test.xlsx", Path.GetFileName(Directory.GetFiles(output).Single()));
    }

    [TestMethod]
    public void FailedOperationVerificationPublishesNothing()
    {
        string input = CreateWorkbook("Book.xlsx");
        string output = Path.Combine(testDirectory, "out");
        FakeOperation operation = new()
        {
            VerifyFailure = new InvalidDataException("Verifier found a difference."),
        };

        WorkbookFileOutcome<Report> outcome = Run(operation, output, input).Single();

        Assert.AreEqual(WorkbookFailureKind.InvalidWorkbook, outcome.FailureKind);
        Assert.AreEqual("Verifier found a difference.", outcome.ErrorMessage);
        Assert.IsEmpty(Directory.GetFileSystemEntries(output));
        Assert.AreEqual(1, operation.VerifyCalls);
    }

    [TestMethod]
    public void ReopenCheckRejectsACorruptedTransformBeforeTheOperationVerifies()
    {
        string input = CreateWorkbook("Book.xlsx");
        string output = Path.Combine(testDirectory, "out");
        FakeOperation operation = new()
        {
            Corrupt = true,
        };

        WorkbookFileOutcome<Report> outcome = Run(operation, output, input).Single();

        Assert.AreEqual(WorkbookFailureKind.InvalidWorkbook, outcome.FailureKind);
        Assert.AreEqual(0, operation.VerifyCalls);
        Assert.IsEmpty(Directory.GetFileSystemEntries(output));
    }

    [TestMethod]
    [DataRow("Book.xls")]
    [DataRow("Book.csv")]
    public void UnsupportedExtensionIsInvalidInputWithoutTransforming(string name)
    {
        string input = Path.Combine(testDirectory, name);
        File.WriteAllText(input, "not a workbook");
        FakeOperation operation = new();

        WorkbookFileOutcome<Report> outcome = Run(operation, Path.Combine(testDirectory, "out"), input).Single();

        Assert.AreEqual(WorkbookFailureKind.InvalidInput, outcome.FailureKind);
        Assert.AreEqual(0, operation.TransformCalls);
    }

    [TestMethod]
    public void TransformSeesTheFormatAndOnlyTheStagedCopy()
    {
        string input = CreateWorkbook("Book.xlsm");
        string output = Path.Combine(testDirectory, "out");
        FakeOperation operation = new();

        Run(operation, output, input);

        Assert.AreEqual(WorkbookFormat.OpenXml, operation.LastFormat);
        Assert.AreEqual(output, Path.GetDirectoryName(operation.LastPath));
        Assert.AreNotEqual(input, operation.LastPath);
        Assert.StartsWith(".xlclean-", Path.GetFileName(operation.LastPath)!);
    }

    [TestMethod]
    public void CancellationKeepsPublishedOutcomesAndRemovesTheStagedCopy()
    {
        string first = CreateWorkbook("First.xlsx");
        string second = CreateWorkbook("Second.xlsx");
        string output = Path.Combine(testDirectory, "out");
        using CancellationTokenSource cancellation = new();
        FakeOperation operation = new()
        {
            OnTransform = path =>
            {
                if (path.Contains("Second", StringComparison.Ordinal))
                {
                    cancellation.Cancel();
                }
            },
        };

        WorkbookBatchCanceledException<Report> exception = Assert.ThrowsExactly<WorkbookBatchCanceledException<Report>>(
            () => WorkbookBatchPipeline.Run([first, second], output, operation, null, cancellation.Token));

        Assert.HasCount(1, exception.Completed);
        Assert.AreEqual(WorkbookFailureKind.None, exception.Completed[0].FailureKind);
        Assert.AreEqual("First_test.xlsx", Path.GetFileName(Directory.GetFiles(output).Single()));
        AssertNoStagingFiles(output);
    }

    [TestMethod]
    public void BinaryPackageArgumentExceptionIsAnInvalidWorkbook()
    {
        string input = Path.Combine(testDirectory, "Book.xlsb");
        File.WriteAllText(input, "staged bytes are never parsed by this operation");
        FakeOperation operation = new()
        {
            TransformFailure = _ => new ArgumentException("Bad part name."),
        };

        WorkbookFileOutcome<Report> outcome = Run(operation, Path.Combine(testDirectory, "out"), input).Single();

        Assert.AreEqual(WorkbookFormat.Binary, operation.LastFormat);
        Assert.AreEqual(WorkbookFailureKind.InvalidWorkbook, outcome.FailureKind);
        StringAssert.Contains(outcome.ErrorMessage, "Bad part name.");
    }

    [TestMethod]
    public void OpenXmlArgumentExceptionStaysInvalidInput()
    {
        string input = CreateWorkbook("Book.xlsx");
        FakeOperation operation = new()
        {
            TransformFailure = _ => new ArgumentException("Bad request."),
        };

        WorkbookFileOutcome<Report> outcome = Run(operation, Path.Combine(testDirectory, "out"), input).Single();

        Assert.AreEqual(WorkbookFailureKind.InvalidInput, outcome.FailureKind);
    }

    [TestMethod]
    public void MissingInputIsInvalidInputWithoutTransforming()
    {
        FakeOperation operation = new();

        WorkbookFileOutcome<Report> outcome = Run(
            operation,
            Path.Combine(testDirectory, "out"),
            Path.Combine(testDirectory, "Missing.xlsx")).Single();

        Assert.AreEqual(WorkbookFailureKind.InvalidInput, outcome.FailureKind);
        Assert.AreEqual(0, operation.TransformCalls);
    }

    [TestMethod]
    public void PreCanceledBatchThrowsPlainCancellationAndCreatesNothing()
    {
        string input = CreateWorkbook("Book.xlsx");
        string output = Path.Combine(testDirectory, "out");
        using CancellationTokenSource cancellation = new();
        cancellation.Cancel();

        OperationCanceledException exception = Assert.ThrowsExactly<OperationCanceledException>(
            () => WorkbookBatchPipeline.Run([input], output, new FakeOperation(), null, cancellation.Token));

        Assert.AreEqual(cancellation.Token, exception.CancellationToken);
        Assert.IsFalse(Directory.Exists(output));
    }

    [TestMethod]
    public void ReportsStartAndCompletionForEveryFileIncludingFailures()
    {
        string failing = CreateWorkbook("Failing.xlsx");
        string passing = CreateWorkbook("Passing.xlsx");
        List<WorkbookProgress> events = [];
        FakeOperation operation = new()
        {
            TransformFailure = path => path == failing
                ? new TestFailure(WorkbookFailureKind.Unexpected, "Failed.")
                : null,
            ReportSheet = true,
        };

        WorkbookBatchPipeline.Run(
            [failing, passing],
            Path.Combine(testDirectory, "out"),
            operation,
            new SynchronousProgress(events.Add),
            CancellationToken.None);

        CollectionAssert.AreEqual(
            new[]
            {
                new WorkbookProgress(0, 2, failing, null, WorkbookStage.StartingFile),
                new WorkbookProgress(0, 2, failing, "Sheet1", WorkbookStage.ProcessingWorksheet),
                new WorkbookProgress(1, 2, failing, null, WorkbookStage.CompletedFile),
                new WorkbookProgress(1, 2, passing, null, WorkbookStage.StartingFile),
                new WorkbookProgress(1, 2, passing, "Sheet1", WorkbookStage.ProcessingWorksheet),
                new WorkbookProgress(2, 2, passing, null, WorkbookStage.CompletedFile),
            },
            events);
    }

    [TestMethod]
    public void TransformReceivesTheResolvedInputPath()
    {
        string input = CreateWorkbook("Book.xlsx");
        string relative = Path.GetRelativePath(Environment.CurrentDirectory, input);
        FakeOperation operation = new();

        WorkbookFileOutcome<Report> outcome = Run(operation, Path.Combine(testDirectory, "out"), relative).Single();

        Assert.AreEqual(relative, outcome.InputPath);
        Assert.AreEqual(input, operation.LastInputPath);
    }

    private static IReadOnlyList<WorkbookFileOutcome<Report>> Run(
        FakeOperation operation,
        string outputDirectory,
        params string[] inputs) =>
        WorkbookBatchPipeline.Run(inputs, outputDirectory, operation, null, CancellationToken.None);

    private static byte[] Hash(string path) => SHA256.HashData(File.ReadAllBytes(path));

    private static void AssertNoStagingFiles(string directory) =>
        Assert.IsEmpty(Directory.GetFiles(directory, ".xlclean-*"));

    private string CreateWorkbook(string name)
    {
        string path = Path.Combine(testDirectory, name);
        SpreadsheetDocumentType type = name.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)
            ? SpreadsheetDocumentType.MacroEnabledWorkbook
            : SpreadsheetDocumentType.Workbook;
        using SpreadsheetDocument document = SpreadsheetDocument.Create(path, type);
        WorkbookPart workbookPart = document.AddWorkbookPart();
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        workbookPart.Workbook = new Workbook(new Sheets(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1U,
            Name = "Sheet1",
        }));
        return path;
    }

    private sealed record Report(string Note);

    private sealed class TestFailure(WorkbookFailureKind kind, string message)
        : WorkbookFailureException(kind, message);

    private sealed class FakeOperation : IWorkbookOperation<Report>
    {
        public Func<string, Exception?> TransformFailure { get; init; } = _ => null;

        public Action<string> OnTransform { get; init; } = _ => { };

        public Exception? VerifyFailure { get; init; }

        public bool Corrupt { get; init; }

        public bool ReportSheet { get; init; }

        public string? LastInputPath { get; private set; }

        public int TransformCalls { get; private set; }

        public int VerifyCalls { get; private set; }

        public WorkbookFormat? LastFormat { get; private set; }

        public string? LastPath { get; private set; }

        public string OutputNameSuffix => "_test";

        public Report Transform(WorkbookFile file, CancellationToken cancellationToken)
        {
            TransformCalls++;
            LastFormat = file.Format;
            LastPath = file.Path;
            LastInputPath = file.InputPath;
            if (ReportSheet)
            {
                file.ReportWorksheet("Sheet1");
            }

            OnTransform(file.RequestedInputPath);
            cancellationToken.ThrowIfCancellationRequested();
            if (TransformFailure(file.RequestedInputPath) is { } failure)
            {
                throw failure;
            }

            if (Corrupt)
            {
                File.WriteAllText(file.Path, "not a package any more");
            }

            return new Report("transformed");
        }

        public void Verify(WorkbookFile file, Report changes, CancellationToken cancellationToken)
        {
            VerifyCalls++;
            if (VerifyFailure is { } failure)
            {
                throw failure;
            }
        }
    }
}

/// <summary>Records progress on the reporting thread; <see cref="Progress{T}"/> posts asynchronously.</summary>
internal sealed class SynchronousProgress(Action<WorkbookProgress> report) : IProgress<WorkbookProgress>
{
    public void Report(WorkbookProgress value) => report(value);
}

internal static class DirectoryTestExtensions
{
    public static string EnsureDirectory(this string path)
    {
        Directory.CreateDirectory(path);
        return path;
    }
}
