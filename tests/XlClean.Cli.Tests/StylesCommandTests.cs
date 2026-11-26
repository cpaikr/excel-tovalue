namespace XlClean.Cli.Tests;

using System.Text.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlClean.Core;

[TestClass]
public sealed class StylesCommandTests
{
    [TestMethod]
    public async Task RootHelpListsTheStylesCommand()
    {
        Invocation invocation = await Invoke(["--help"], UnexpectedCleanup);

        StringAssert.Contains(invocation.StandardOutput, "  styles  ");
    }

    [TestMethod]
    public async Task StylesHelpDescribesTheCommand()
    {
        Invocation invocation = await Invoke(["styles", "--help"], UnexpectedCleanup);

        Assert.AreEqual(CliApplication.SuccessExitCode, invocation.ExitCode);
        StringAssert.Contains(invocation.StandardOutput, "xlclean styles --output <directory> <workbook>...");
        StringAssert.Contains(invocation.StandardOutput, "mergedCellFormats");
        StringAssert.Contains(invocation.StandardOutput, "unsupported-workbook-content");
        Assert.DoesNotContain("--start-row", invocation.StandardOutput);
        Assert.DoesNotContain("\r", invocation.StandardOutput);
    }

    [TestMethod]
    [DataRow(new string[] { "styles", "--output", "out", "--start-row", "2", "Book.xlsx" }, "Option '--start-row' applies only to 'xlclean values'.")]
    [DataRow(new string[] { "styles", "--output", "out", "-x", "TF", "Book.xlsx" }, "Option '-x' applies only to 'xlclean values'.")]
    [DataRow(new string[] { "styles", "--exclude=TF", "--output", "out", "Book.xlsx" }, "Option '--exclude' applies only to 'xlclean values'.")]
    [DataRow(new string[] { "styles", "--output", "out", "--json", "Book.xlsx" }, "Unknown option '--json'.")]
    [DataRow(new string[] { "styles", "Book.xlsx" }, "Required option '--output' was not provided.")]
    [DataRow(new string[] { "styles", "--output", "out" }, "At least one input workbook is required.")]
    public async Task ScopeOptionsAndMissingArgumentsAreUsageErrors(string[] arguments, string expectedMessage)
    {
        Invocation invocation = await Invoke(arguments, UnexpectedCleanup);

        Assert.AreEqual(CliApplication.UsageErrorExitCode, invocation.ExitCode);
        using JsonDocument document = ParseSingleDocument(invocation.StandardOutput);
        JsonElement root = document.RootElement;
        Assert.AreEqual("styles", root.GetProperty("command").GetString());
        Assert.AreEqual(expectedMessage, root.GetProperty("error").GetProperty("message").GetString());
        Assert.AreEqual("Run xlclean styles --help", root.GetProperty("error").GetProperty("help").GetString());
    }

    [TestMethod]
    public async Task BuildsTheRequestAndReportsChangesExactly()
    {
        StyleCleanupRequest? captured = null;
        StyleCleanupBatchResult result = new(
        [
            new StyleCleanupFileResult(
                "First.xlsx",
                "First_스타일정리.xlsx",
                WorkbookFailureKind.None,
                null,
                new StyleCleanupChanges(5003, 4, 812, 37, 120),
                ["warning"]),
            new StyleCleanupFileResult(
                "Second.xlsm",
                null,
                WorkbookFailureKind.UnsupportedWorkbookContent,
                "Macro sheet.",
                null,
                []),
        ]);

        Invocation invocation = await Invoke(
            ["styles", "-o", "out", "First.xlsx", "--", "Second.xlsm"],
            (request, progress, _) =>
            {
                captured = request;
                Assert.IsNull(progress);
                return Task.FromResult(result);
            });

        Assert.AreEqual(CliApplication.ConversionFailureExitCode, invocation.ExitCode);
        Assert.IsNotNull(captured);
        Assert.AreEqual(Path.GetFullPath("out"), captured.OutputDirectory);
        Assert.AreEqual("First.xlsx,Second.xlsm", string.Join(',', captured.InputPaths));
        Assert.AreEqual(
            "{\"ok\":false,\"version\":\"" + CliApplication.Version + "\",\"command\":\"styles\",\"files\":[" +
            "{\"input\":\"First.xlsx\",\"status\":\"converted\",\"output\":\"First_스타일정리.xlsx\",\"failure\":null," +
            "\"changes\":{\"cellStyles\":{\"before\":5003,\"after\":4},\"cellFormats\":{\"before\":812,\"after\":37}," +
            "\"mergedCellFormats\":120},\"warnings\":[\"warning\"]}," +
            "{\"input\":\"Second.xlsm\",\"status\":\"failed\",\"output\":null," +
            "\"failure\":{\"kind\":\"unsupported-workbook-content\",\"message\":\"Macro sheet.\"}," +
            "\"changes\":null,\"warnings\":[]}]," +
            "\"summary\":{\"total\":2,\"succeeded\":1,\"failed\":1,\"canceled\":false},\"error\":null}\n",
            invocation.StandardOutput);
        Assert.AreEqual(string.Empty, invocation.StandardError);
    }

    [TestMethod]
    public async Task CancellationReportsFilesCompletedBeforeIt()
    {
        using CancellationTokenSource cancellation = new();
        StyleCleanupBatchResult completed = new(
        [
            new StyleCleanupFileResult(
                "First.xlsx",
                "First_스타일정리.xlsx",
                WorkbookFailureKind.None,
                null,
                new StyleCleanupChanges(1, 1, 1, 1, 0),
                []),
        ]);

        Invocation invocation = await Invoke(
            ["styles", "--output", "out", "First.xlsx", "Second.xlsx"],
            (_, _, token) =>
            {
                cancellation.Cancel();
                throw new StyleCleanupCanceledException(completed, token);
            },
            cancellation.Token);

        Assert.AreEqual(CliApplication.CanceledExitCode, invocation.ExitCode);
        using JsonDocument document = ParseSingleDocument(invocation.StandardOutput);
        JsonElement root = document.RootElement;
        Assert.AreEqual("styles", root.GetProperty("command").GetString());
        Assert.AreEqual(1, root.GetProperty("files").GetArrayLength());
        Assert.IsTrue(root.GetProperty("summary").GetProperty("canceled").GetBoolean());
    }

    [TestMethod]
    public async Task UnexpectedCleanupErrorIsReportedWithoutStackTrace()
    {
        Invocation invocation = await Invoke(
            ["styles", "--output", "out", "Book.xlsx"],
            (_, _, _) => throw new IOException("Disk unavailable."));

        Assert.AreEqual(CliApplication.ConversionFailureExitCode, invocation.ExitCode);
        using JsonDocument document = ParseSingleDocument(invocation.StandardOutput);
        Assert.AreEqual("unexpected", document.RootElement.GetProperty("error").GetProperty("kind").GetString());
        Assert.AreEqual("styles", document.RootElement.GetProperty("command").GetString());
    }

    private static JsonDocument ParseSingleDocument(string standardOutput)
    {
        Assert.EndsWith("\n", standardOutput);
        string body = standardOutput[..^1];
        Assert.DoesNotContain("\n", body);
        return JsonDocument.Parse(body);
    }

    private static Task<StyleCleanupBatchResult> UnexpectedCleanup(
        StyleCleanupRequest request,
        IProgress<WorkbookProgress>? progress,
        CancellationToken cancellationToken) =>
        throw new AssertFailedException($"Cleanup was unexpectedly invoked for '{request.InputPaths[0]}'.");

    private static Task<FormulaRemovalBatchResult> UnexpectedConversion(
        FormulaRemovalRequest request,
        IProgress<WorkbookProgress>? progress,
        CancellationToken cancellationToken) =>
        throw new AssertFailedException($"Conversion was unexpectedly invoked for '{request.InputPaths[0]}'.");

    private static async Task<Invocation> Invoke(
        IReadOnlyList<string> arguments,
        StyleCleanupConverter cleaner,
        CancellationToken cancellationToken = default)
    {
        using StringWriter standardOutput = new();
        using StringWriter standardError = new();
        int exitCode = await CliApplication.RunAsync(
            arguments,
            standardOutput,
            standardError,
            UnexpectedConversion,
            cleaner,
            cancellationToken);
        return new Invocation(exitCode, standardOutput.ToString(), standardError.ToString());
    }

    private sealed record Invocation(int ExitCode, string StandardOutput, string StandardError);
}
