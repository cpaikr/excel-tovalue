namespace ExcelFormulaRemover.Cli.Tests;

using System.Globalization;
using System.Text.Json;
using ExcelFormulaRemover.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

[TestClass]
public sealed class CliApplicationTests
{
    [TestMethod]
    public async Task HelpReturnsSuccessWithoutConverting()
    {
        Invocation invocation = await Invoke(["--help"], UnexpectedConversion);

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        StringAssert.Contains(invocation.StandardOutput, "Usage:");
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    public async Task HelpBeforeTerminatorWinsOverMalformedArguments()
    {
        Invocation invocation = await Invoke(["--unknown", "--help"], UnexpectedConversion);

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        StringAssert.Contains(invocation.StandardOutput, "Usage:");
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    [DataRow(new string[] { "Book.xlsx" }, "--output")]
    [DataRow(new string[] { "--output", "out" }, "input workbook")]
    [DataRow(new string[] { "--wat" }, "Unknown option")]
    [DataRow(new string[] { "--output", "out", "--start-row", "0", "Book.xlsx" }, "Start row")]
    [DataRow(new string[] { "--output", "out", "--output", "other", "Book.xlsx" }, "more than once")]
    public async Task InvalidArgumentsReturnUsageError(string[] arguments, string expectedMessage)
    {
        Invocation invocation = await Invoke(arguments, UnexpectedConversion);

        Assert.AreEqual(CliApplication.UsageErrorExitCode, invocation.ExitCode);
        StringAssert.Contains(invocation.StandardOutput, expectedMessage);
        StringAssert.Contains(invocation.StandardOutput, "kind: usage");
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    public async Task BuildsRequestFromLongAndShortOptions()
    {
        FormulaRemovalRequest? capturedRequest = null;
        FormulaRemovalConverter converter = (request, progress, cancellationToken) =>
        {
            capturedRequest = request;
            Assert.IsNull(progress);
            Assert.IsFalse(cancellationToken.IsCancellationRequested);
            return Task.FromResult(SuccessResult(request.InputPaths[0], "published.xlsx"));
        };

        Invocation invocation = await Invoke(
            [
                "--output=out",
                "--start-row",
                "42",
                "-x",
                "TF",
                "--exclude=Archive",
                "First.xlsx",
                "Second.xlsm",
            ],
            converter);

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        Assert.IsNotNull(capturedRequest);
        Assert.HasCount(2, capturedRequest.InputPaths);
        Assert.AreEqual("First.xlsx", capturedRequest.InputPaths[0]);
        Assert.AreEqual("Second.xlsm", capturedRequest.InputPaths[1]);
        Assert.AreEqual(42, capturedRequest.StartRow);
        Assert.HasCount(2, capturedRequest.ExcludedSheetNamePatterns);
        Assert.AreEqual("TF", capturedRequest.ExcludedSheetNamePatterns[0]);
        Assert.AreEqual("Archive", capturedRequest.ExcludedSheetNamePatterns[1]);
        Assert.AreEqual(Path.GetFullPath("out"), capturedRequest.OutputDirectory);
    }

    [TestMethod]
    public async Task DoubleDashAllowsInputBeginningWithHyphen()
    {
        FormulaRemovalRequest? capturedRequest = null;
        FormulaRemovalConverter converter = (request, _, _) =>
        {
            capturedRequest = request;
            return Task.FromResult(SuccessResult(request.InputPaths[0], "published.xlsx"));
        };

        Invocation invocation = await Invoke(
            ["--output", "out", "--", "-Book.xlsx"],
            converter);

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        Assert.IsNotNull(capturedRequest);
        Assert.AreEqual("-Book.xlsx", capturedRequest.InputPaths.Single());
    }

    [TestMethod]
    public async Task HelpAfterDoubleDashIsAnInputPath()
    {
        FormulaRemovalRequest? capturedRequest = null;
        FormulaRemovalConverter converter = (request, _, _) =>
        {
            capturedRequest = request;
            return Task.FromResult(SuccessResult(request.InputPaths[0], "published.xlsx"));
        };

        Invocation invocation = await Invoke(
            ["--output", "out", "--", "--help"],
            converter);

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        Assert.IsNotNull(capturedRequest);
        Assert.AreEqual("--help", capturedRequest.InputPaths.Single());
    }

    [TestMethod]
    [DataRow("1", 1)]
    [DataRow("1048576", FormulaRemovalRequest.MaximumRowNumber)]
    public async Task AcceptsExcelStartRowBoundaries(string value, int expected)
    {
        FormulaRemovalRequest? capturedRequest = null;

        Invocation invocation = await Invoke(
            ["--output", "out", "--start-row", value, "Book.xlsx"],
            (request, _, _) =>
            {
                capturedRequest = request;
                return Task.FromResult(SuccessResult(request.InputPaths[0], "published.xlsx"));
            });

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        Assert.IsNotNull(capturedRequest);
        Assert.AreEqual(expected, capturedRequest.StartRow);
    }

    [TestMethod]
    [DataRow("+1")]
    [DataRow("1,000")]
    [DataRow("١")]
    [DataRow("1048577")]
    public async Task RejectsNonCanonicalStartRows(string value)
    {
        Invocation invocation = await Invoke(
            ["--output", "out", $"--start-row={value}", "Book.xlsx"],
            UnexpectedConversion);

        Assert.AreEqual(CliApplication.UsageErrorExitCode, invocation.ExitCode);
        StringAssert.Contains(invocation.StandardOutput, "Start row");
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    public async Task ConversionFailuresReturnOneAndKeepRequestOrder()
    {
        FormulaRemovalBatchResult result = new(
        [
            SuccessfulFile("First.xlsx", "First_수식제거.xlsx"),
            FailedFile("Second.xlsx", FormulaRemovalFailureKind.MissingCachedValue, "No cache."),
        ]);

        Invocation invocation = await Invoke(
            ["--output", "out", "First.xlsx", "Second.xlsx"],
            (_, _, _) => Task.FromResult(result));

        Assert.AreEqual(CliApplication.ConversionFailureExitCode, invocation.ExitCode);
        Assert.AreEqual(
            """
            files[2]{index,input,status,output,failure}:
              1,"First.xlsx",ok,"First_수식제거.xlsx",null
              2,"Second.xlsx",failed,null,missing-cached-value
            messages[1]{index,level,text}:
              2,error,"No cache."
            summary:
              total: 2
              succeeded: 1
              failed: 1
              canceled: false
            """.ReplaceLineEndings("\n"),
            invocation.StandardOutput);
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    public async Task JsonOutputIsValidNdjsonWithSummaryLast()
    {
        FormulaRemovalBatchResult result = new(
            [SuccessfulFile("Book.xlsx", "Book_수식제거.xlsx")]);

        Invocation invocation = await Invoke(
            ["--json", "--output", "out", "Book.xlsx"],
            (_, _, _) => Task.FromResult(result));

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        string[] lines = invocation.StandardOutput
            .Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.HasCount(2, lines);
        using JsonDocument file = JsonDocument.Parse(lines[0]);
        using JsonDocument summary = JsonDocument.Parse(lines[1]);
        Assert.AreEqual("result", file.RootElement.GetProperty("type").GetString());
        Assert.IsTrue(file.RootElement.GetProperty("succeeded").GetBoolean());
        Assert.AreEqual("summary", summary.RootElement.GetProperty("type").GetString());
        Assert.AreEqual(1, summary.RootElement.GetProperty("succeeded").GetInt32());
        Assert.IsFalse(summary.RootElement.GetProperty("canceled").GetBoolean());
    }

    [TestMethod]
    public async Task ToonOutputIsInvariantAndEscapesDynamicStrings()
    {
        CultureInfo previousCulture = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
        try
        {
            FormulaRemovalBatchResult result = new(
            [
                new FormulaRemovalFileResult(
                    "A,\n\"\\😀한글.xlsx",
                    "out,\tBook.xlsx",
                    FormulaRemovalFailureKind.None,
                    null,
                    1,
                    1,
                    0,
                    0,
                    ["line\r\n\u0001"]),
            ]);

            Invocation invocation = await Invoke(
                ["--output", "out", "Book.xlsx"],
                (_, _, _) => Task.FromResult(result));

            Assert.AreEqual(
                """
                files[1]{index,input,status,output,failure}:
                  1,"A,\n\"\\😀한글.xlsx",ok,"out,\tBook.xlsx",null
                messages[1]{index,level,text}:
                  1,warning,"line\r\n\u0001"
                summary:
                  total: 1
                  succeeded: 1
                  failed: 0
                  canceled: false
                """.ReplaceLineEndings("\n"),
                invocation.StandardOutput);
            Assert.DoesNotEndWith(invocation.StandardOutput, "\n");
        }
        finally
        {
            CultureInfo.CurrentCulture = previousCulture;
        }
    }

    [TestMethod]
    public async Task CancellationReturnsConventionalExitCode()
    {
        using CancellationTokenSource cancellation = new();
        cancellation.Cancel();

        Invocation invocation = await Invoke(
            ["--output", "out", "Book.xlsx"],
            (_, _, token) => Task.FromCanceled<FormulaRemovalBatchResult>(token),
            cancellation.Token);

        Assert.AreEqual(CliApplication.CanceledExitCode, invocation.ExitCode);
        StringAssert.Contains(invocation.StandardOutput, "files: []");
        StringAssert.Contains(invocation.StandardOutput, "canceled: true");
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    public async Task CancellationReportsFilesCompletedBeforeIt()
    {
        using CancellationTokenSource cancellation = new();
        FormulaRemovalBatchResult completed = new(
            [SuccessfulFile("First.xlsx", "First_수식제거.xlsx")]);

        Invocation invocation = await Invoke(
            ["--output", "out", "First.xlsx", "Second.xlsx"],
            (_, _, token) =>
            {
                cancellation.Cancel();
                throw new FormulaRemovalCanceledException(completed, token);
            },
            cancellation.Token);

        Assert.AreEqual(CliApplication.CanceledExitCode, invocation.ExitCode);
        StringAssert.Contains(invocation.StandardOutput, "\"First.xlsx\",ok");
        StringAssert.Contains(invocation.StandardOutput, "canceled: true");
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    public async Task JsonModeFormatsUsageErrorsAsJson()
    {
        foreach (string[] arguments in new[]
        {
            new[] { "--json", "--unknown" },
            new[] { "--json", "--output", "bad\0directory", "Book.xlsx" },
        })
        {
            Invocation invocation = await Invoke(arguments, UnexpectedConversion);

            Assert.AreEqual(CliApplication.UsageErrorExitCode, invocation.ExitCode);
            using JsonDocument document = JsonDocument.Parse(invocation.StandardOutput);
            Assert.AreEqual("error", document.RootElement.GetProperty("type").GetString());
            Assert.AreEqual("usage", document.RootElement.GetProperty("kind").GetString());
            Assert.AreEqual(string.Empty, invocation.StandardError);
        }
    }

    [TestMethod]
    public async Task FatalConversionErrorIsReportedWithoutStackTrace()
    {
        Invocation invocation = await Invoke(
            ["--output", "out", "Book.xlsx"],
            (_, _, _) => throw new IOException("Disk unavailable."));

        Assert.AreEqual(CliApplication.ConversionFailureExitCode, invocation.ExitCode);
        StringAssert.Contains(invocation.StandardOutput, "kind: conversion");
        StringAssert.Contains(invocation.StandardOutput, "type: \"IOException\"");
        StringAssert.Contains(invocation.StandardOutput, "message: \"Disk unavailable.\"");
        Assert.DoesNotContain(invocation.StandardOutput, " at ");
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    private static Task<FormulaRemovalBatchResult> UnexpectedConversion(
        FormulaRemovalRequest request,
        IProgress<FormulaRemovalProgress>? progress,
        CancellationToken cancellationToken) =>
        throw new AssertFailedException(
            $"Conversion was unexpectedly invoked for '{request.InputPaths[0]}'.");

    private static FormulaRemovalBatchResult SuccessResult(string inputPath, string outputPath) =>
        new([SuccessfulFile(inputPath, outputPath)]);

    private static FormulaRemovalFileResult SuccessfulFile(string inputPath, string outputPath) =>
        new(
            inputPath,
            outputPath,
            FormulaRemovalFailureKind.None,
            null,
            3,
            2,
            1,
            0,
            Array.Empty<string>());

    private static FormulaRemovalFileResult FailedFile(
        string inputPath,
        FormulaRemovalFailureKind failureKind,
        string message) =>
        new(
            inputPath,
            null,
            failureKind,
            message,
            0,
            0,
            0,
            0,
            Array.Empty<string>());

    private static async Task<Invocation> Invoke(
        IReadOnlyList<string> arguments,
        FormulaRemovalConverter converter,
        CancellationToken cancellationToken = default)
    {
        using StringWriter standardOutput = new();
        using StringWriter standardError = new();
        int exitCode = await CliApplication.RunAsync(
            arguments,
            standardOutput,
            standardError,
            converter,
            cancellationToken);
        return new Invocation(exitCode, standardOutput.ToString(), standardError.ToString());
    }

    private sealed record Invocation(
        int ExitCode,
        string StandardOutput,
        string StandardError);
}
