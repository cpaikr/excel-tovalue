namespace XlClean.Cli.Tests;

using System.Globalization;
using System.Text.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlClean.Core;

[TestClass]
public sealed class CliApplicationTests
{
    [TestMethod]
    public async Task HelpReturnsSuccessWithoutConverting()
    {
        Invocation invocation = await Invoke(["--help"], UnexpectedConversion);

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        StringAssert.Contains(invocation.StandardOutput, "Commands:");
        StringAssert.Contains(invocation.StandardOutput, "  values  ");
        Assert.EndsWith("\n", invocation.StandardOutput);
        Assert.DoesNotContain("\r", invocation.StandardOutput);
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    public async Task BareInvocationPrintsExactlyTheRootHelp()
    {
        Invocation help = await Invoke(["--help"], UnexpectedConversion);
        Invocation bare = await Invoke([], UnexpectedConversion);

        Assert.AreEqual(help, bare);
        Assert.AreEqual(CliApplication.SuccessExitCode, bare.ExitCode);
    }

    [TestMethod]
    [DataRow(new string[] { "values", "--help" })]
    [DataRow(new string[] { "values", "-h" })]
    [DataRow(new string[] { "values", "--unknown", "--version", "--help" })]
    public async Task CommandHelpDescribesTheCommand(string[] arguments)
    {
        Invocation invocation = await Invoke(arguments, UnexpectedConversion);
        Invocation root = await Invoke(["--help"], UnexpectedConversion);

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        StringAssert.Contains(invocation.StandardOutput, "xlclean values --output <directory>");
        StringAssert.Contains(invocation.StandardOutput, "removedFormulas");
        Assert.AreNotEqual(root.StandardOutput, invocation.StandardOutput);
        Assert.EndsWith("\n", invocation.StandardOutput);
        Assert.DoesNotContain("\r", invocation.StandardOutput);
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    [DataRow(new string[] { "--version" })]
    [DataRow(new string[] { "--unknown", "--version" })]
    [DataRow(new string[] { "values", "--version" })]
    public async Task VersionPrintsNameAndReleaseIdentityWithoutConverting(string[] arguments)
    {
        Invocation invocation = await Invoke(arguments, UnexpectedConversion);

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        Assert.AreEqual(
            $"xlclean {CliApplication.Version}\n",
            invocation.StandardOutput);
        Assert.DoesNotContain("+", CliApplication.Version);
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    public async Task CommandMustComeFirstForCommandHelp()
    {
        Invocation invocation = await Invoke(["--version", "values", "--help"], UnexpectedConversion);
        Invocation root = await Invoke(["--help"], UnexpectedConversion);

        Assert.AreEqual(root, invocation);
    }

    [TestMethod]
    public async Task CommandNameAfterTerminatorIsAnInputPath()
    {
        FormulaRemovalRequest? capturedRequest = null;
        Invocation invocation = await Invoke(
            ["values", "--output", "out", "--", "values"],
            (request, _, _) =>
            {
                capturedRequest = request;
                return Task.FromResult(SuccessResult(request.InputPaths[0], "published.xlsx"));
            });

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        Assert.IsNotNull(capturedRequest);
        Assert.AreEqual("values", capturedRequest.InputPaths.Single());
    }

    [TestMethod]
    public async Task HelpWinsOverVersion()
    {
        Invocation invocation = await Invoke(["--version", "-h"], UnexpectedConversion);

        StringAssert.Contains(invocation.StandardOutput, "Commands:");
    }

    [TestMethod]
    public async Task HelpBeforeTerminatorWinsOverMalformedArguments()
    {
        Invocation invocation = await Invoke(["--unknown", "--help"], UnexpectedConversion);

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        StringAssert.Contains(invocation.StandardOutput, "Commands:");
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    [DataRow(new string[] { "values", "Book.xlsx" }, "--output")]
    [DataRow(new string[] { "values", "--output", "out" }, "input workbook")]
    [DataRow(new string[] { "values" }, "--output")]
    [DataRow(new string[] { "values", "--wat" }, "Unknown option")]
    [DataRow(new string[] { "values", "--output", "out", "--start-row", "0", "Book.xlsx" }, "Start row")]
    [DataRow(new string[] { "values", "--output", "out", "--output", "other", "Book.xlsx" }, "more than once")]
    [DataRow(new string[] { "values", "--json", "--output", "out", "Book.xlsx" }, "Unknown option '--json'")]
    [DataRow(new string[] { "values", "--output", "bad\0directory", "Book.xlsx" }, "")]
    public async Task InvalidArgumentsReturnUsageError(string[] arguments, string expectedMessage)
    {
        Invocation invocation = await Invoke(arguments, UnexpectedConversion);

        Assert.AreEqual(CliApplication.UsageErrorExitCode, invocation.ExitCode);
        using JsonDocument document = ParseSingleDocument(invocation.StandardOutput);
        JsonElement root = document.RootElement;
        Assert.IsFalse(root.GetProperty("ok").GetBoolean());
        Assert.AreEqual(0, root.GetProperty("files").GetArrayLength());
        Assert.AreEqual(0, root.GetProperty("summary").GetProperty("total").GetInt32());
        JsonElement error = root.GetProperty("error");
        Assert.AreEqual("usage", error.GetProperty("kind").GetString());
        StringAssert.Contains(error.GetProperty("message").GetString(), expectedMessage);
        Assert.AreEqual("Run xlclean values --help", error.GetProperty("help").GetString());
        Assert.AreEqual("values", root.GetProperty("command").GetString());
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    [DataRow(new string[] { "--output", "out", "Book.xlsx" }, "A command is required before options; for example 'xlclean values --output ...'.")]
    [DataRow(new string[] { "-o", "out", "Book.xlsx" }, "A command is required before options; for example 'xlclean values -o ...'.")]
    [DataRow(new string[] { "convert", "--output", "out", "Book.xlsx" }, "Unknown command 'convert'.")]
    [DataRow(new string[] { "Values", "--output", "out", "Book.xlsx" }, "Unknown command 'Values'.")]
    [DataRow(new string[] { "Book.xlsx" }, "Unknown command 'Book.xlsx'.")]
    [DataRow(new string[] { "--", "values" }, "A command is required; for example 'xlclean values --output <directory> <workbook>...'.")]
    [DataRow(new string[] { "--output", "out", "values", "Book.xlsx" }, "A command is required before options; for example 'xlclean values --output ...'.")]
    public async Task MissingOrUnknownCommandIsAUsageErrorWithoutCommand(
        string[] arguments,
        string expectedMessage)
    {
        Invocation invocation = await Invoke(arguments, UnexpectedConversion);

        Assert.AreEqual(CliApplication.UsageErrorExitCode, invocation.ExitCode);
        using JsonDocument document = ParseSingleDocument(invocation.StandardOutput);
        JsonElement root = document.RootElement;
        Assert.AreEqual(JsonValueKind.Null, root.GetProperty("command").ValueKind);
        Assert.AreEqual(0, root.GetProperty("files").GetArrayLength());
        JsonElement error = root.GetProperty("error");
        Assert.AreEqual("usage", error.GetProperty("kind").GetString());
        Assert.AreEqual(expectedMessage, error.GetProperty("message").GetString());
        Assert.AreEqual("Run xlclean --help", error.GetProperty("help").GetString());
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
                "values",
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
            ["values", "--output", "out", "--", "-Book.xlsx"],
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
            ["values", "--output", "out", "--", "--help"],
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
            ["values", "--output", "out", "--start-row", value, "Book.xlsx"],
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
            ["values", "--output", "out", $"--start-row={value}", "Book.xlsx"],
            UnexpectedConversion);

        Assert.AreEqual(CliApplication.UsageErrorExitCode, invocation.ExitCode);
        using JsonDocument document = ParseSingleDocument(invocation.StandardOutput);
        StringAssert.Contains(
            document.RootElement.GetProperty("error").GetProperty("message").GetString(),
            "Start row");
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
            ["values", "--output", "out", "First.xlsx", "Second.xlsx"],
            (_, _, _) => Task.FromResult(result));

        Assert.AreEqual(CliApplication.ConversionFailureExitCode, invocation.ExitCode);
        Assert.AreEqual(
            "{\"ok\":false,\"version\":\"" + CliApplication.Version + "\",\"command\":\"values\",\"files\":[" +
            "{\"input\":\"First.xlsx\",\"status\":\"converted\",\"output\":\"First_수식제거.xlsx\"," +
            "\"failure\":null,\"changes\":{\"removedFormulas\":3,\"removedTableFormulaDefinitions\":0," +
            "\"processedSheets\":2,\"excludedSheets\":1},\"warnings\":[]}," +
            "{\"input\":\"Second.xlsx\",\"status\":\"failed\",\"output\":null," +
            "\"failure\":{\"kind\":\"missing-cached-value\",\"message\":\"No cache.\"}," +
            "\"changes\":null,\"warnings\":[]}]," +
            "\"summary\":{\"total\":2,\"succeeded\":1,\"failed\":1,\"canceled\":false}," +
            "\"error\":null}\n",
            invocation.StandardOutput);
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    [DataRow(FormulaRemovalFailureKind.InvalidInput, "invalid-input")]
    [DataRow(FormulaRemovalFailureKind.MissingCachedValue, "missing-cached-value")]
    [DataRow(FormulaRemovalFailureKind.FormulaRangeCrossesBoundary, "formula-range-crosses-boundary")]
    [DataRow(FormulaRemovalFailureKind.InvalidWorkbook, "invalid-workbook")]
    [DataRow(FormulaRemovalFailureKind.FileAccess, "file-access")]
    [DataRow(FormulaRemovalFailureKind.Unexpected, "unexpected")]
    public async Task FailureKindsUseStableIdentifiers(FormulaRemovalFailureKind kind, string expected)
    {
        Invocation invocation = await Invoke(
            ["values", "--output", "out", "Book.xlsx"],
            (_, _, _) => Task.FromResult(
                new FormulaRemovalBatchResult([FailedFile("Book.xlsx", kind, "Failed.")])));

        using JsonDocument document = ParseSingleDocument(invocation.StandardOutput);
        Assert.AreEqual(
            expected,
            document.RootElement.GetProperty("files")[0]
                .GetProperty("failure").GetProperty("kind").GetString());
    }

    [TestMethod]
    public async Task SuccessfulBatchReportsOkWithCountsAndWarnings()
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
                    1234,
                    5,
                    6,
                    7,
                    ["line\r\n\u0001"]),
            ]);

            Invocation invocation = await Invoke(
                ["values", "--output", "out", "Book.xlsx"],
                (_, _, _) => Task.FromResult(result));

            Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
            StringAssert.Contains(invocation.StandardOutput, "한글");
            using JsonDocument document = ParseSingleDocument(invocation.StandardOutput);
            JsonElement root = document.RootElement;
            Assert.IsTrue(root.GetProperty("ok").GetBoolean());
            Assert.AreEqual(CliApplication.Version, root.GetProperty("version").GetString());
            Assert.AreEqual(JsonValueKind.Null, root.GetProperty("error").ValueKind);
            JsonElement file = root.GetProperty("files")[0];
            Assert.AreEqual("A,\n\"\\😀한글.xlsx", file.GetProperty("input").GetString());
            Assert.AreEqual("out,\tBook.xlsx", file.GetProperty("output").GetString());
            Assert.AreEqual("converted", file.GetProperty("status").GetString());
            Assert.AreEqual(JsonValueKind.Null, file.GetProperty("failure").ValueKind);
            Assert.AreEqual("values", root.GetProperty("command").GetString());
            JsonElement changes = file.GetProperty("changes");
            Assert.AreEqual(1234, changes.GetProperty("removedFormulas").GetInt32());
            Assert.AreEqual(5, changes.GetProperty("processedSheets").GetInt32());
            Assert.AreEqual(6, changes.GetProperty("excludedSheets").GetInt32());
            Assert.AreEqual(7, changes.GetProperty("removedTableFormulaDefinitions").GetInt32());
            Assert.AreEqual("line\r\n\u0001", file.GetProperty("warnings")[0].GetString());
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
            ["values", "--output", "out", "Book.xlsx"],
            (_, _, token) => Task.FromCanceled<FormulaRemovalBatchResult>(token),
            cancellation.Token);

        Assert.AreEqual(CliApplication.CanceledExitCode, invocation.ExitCode);
        using JsonDocument document = ParseSingleDocument(invocation.StandardOutput);
        JsonElement root = document.RootElement;
        Assert.IsFalse(root.GetProperty("ok").GetBoolean());
        Assert.AreEqual("values", root.GetProperty("command").GetString());
        Assert.AreEqual(0, root.GetProperty("files").GetArrayLength());
        Assert.IsTrue(root.GetProperty("summary").GetProperty("canceled").GetBoolean());
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    public async Task CancellationReportsFilesCompletedBeforeIt()
    {
        using CancellationTokenSource cancellation = new();
        FormulaRemovalBatchResult completed = new(
            [SuccessfulFile("First.xlsx", "First_수식제거.xlsx")]);

        Invocation invocation = await Invoke(
            ["values", "--output", "out", "First.xlsx", "Second.xlsx"],
            (_, _, token) =>
            {
                cancellation.Cancel();
                throw new FormulaRemovalCanceledException(completed, token);
            },
            cancellation.Token);

        Assert.AreEqual(CliApplication.CanceledExitCode, invocation.ExitCode);
        using JsonDocument document = ParseSingleDocument(invocation.StandardOutput);
        JsonElement root = document.RootElement;
        Assert.IsFalse(root.GetProperty("ok").GetBoolean());
        Assert.AreEqual("values", root.GetProperty("command").GetString());
        Assert.AreEqual("converted", root.GetProperty("files")[0].GetProperty("status").GetString());
        Assert.AreEqual(1, root.GetProperty("summary").GetProperty("succeeded").GetInt32());
        Assert.IsTrue(root.GetProperty("summary").GetProperty("canceled").GetBoolean());
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    public async Task FatalConversionErrorIsReportedWithoutStackTrace()
    {
        Invocation invocation = await Invoke(
            ["values", "--output", "out", "Book.xlsx"],
            (_, _, _) => throw new IOException("Disk unavailable."));

        Assert.AreEqual(CliApplication.ConversionFailureExitCode, invocation.ExitCode);
        using JsonDocument document = ParseSingleDocument(invocation.StandardOutput);
        Assert.AreEqual("values", document.RootElement.GetProperty("command").GetString());
        JsonElement error = document.RootElement.GetProperty("error");
        Assert.AreEqual("unexpected", error.GetProperty("kind").GetString());
        Assert.AreEqual("Disk unavailable.", error.GetProperty("message").GetString());
        Assert.DoesNotContain(" at ", invocation.StandardOutput);
        Assert.DoesNotContain("IOException", invocation.StandardOutput);
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    // The process contract is one JSON document followed by exactly one newline.
    private static JsonDocument ParseSingleDocument(string standardOutput)
    {
        Assert.EndsWith("\n", standardOutput);
        string body = standardOutput[..^1];
        Assert.DoesNotContain("\n", body);
        Assert.DoesNotContain("\r", body);
        return JsonDocument.Parse(body);
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
