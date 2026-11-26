namespace XlClean.Cli.Tests;

using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;

/// <summary>
/// Exercises the command line as an external process, which is the boundary
/// agents and scripts actually use. By default the tests launch the built
/// assembly; <c>XLCLEAN_EXE</c> points them at an installed
/// release executable so the same contract certifies the published artifact.
/// </summary>
[TestClass]
public sealed class ProcessContractTests
{
    private const string ExecutableVariable = "XLCLEAN_EXE";
    private const string ExpectedVersionVariable = "XLCLEAN_EXPECTED_VERSION";

    private string workingDirectory = string.Empty;

    [TestInitialize]
    public void CreateWorkingDirectory()
    {
        workingDirectory = Path.Combine(
            Path.GetTempPath(),
            "xlclean-tests",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workingDirectory);
    }

    [TestCleanup]
    public void DeleteWorkingDirectory()
    {
        if (Directory.Exists(workingDirectory))
        {
            Directory.Delete(workingDirectory, recursive: true);
        }
    }

    [TestMethod]
    public async Task BareInvocationIsByteIdenticalToHelp()
    {
        ProcessResult help = await Run("--help");
        ProcessResult bare = await Run();

        Assert.AreEqual(0, help.ExitCode);
        Assert.AreEqual(0, bare.ExitCode);
        CollectionAssert.AreEqual(help.StandardOutput, bare.StandardOutput);
        StringAssert.Contains(help.StandardOutputText, "Commands:");
        Assert.EndsWith("\n", help.StandardOutputText);
        Assert.DoesNotContain("\r", help.StandardOutputText);
        Assert.IsEmpty(help.StandardError);
        Assert.IsEmpty(bare.StandardError);
    }

    [TestMethod]
    public async Task VersionReportsTheReleaseIdentity()
    {
        string expectedVersion =
            Environment.GetEnvironmentVariable(ExpectedVersionVariable) is { Length: > 0 } configured
                ? configured
                : CliApplication.Version;

        ProcessResult result = await Run("--version");

        Assert.AreEqual(0, result.ExitCode);
        Assert.AreEqual($"xlclean {expectedVersion}\n", result.StandardOutputText);
        Assert.IsEmpty(result.StandardError);
    }

    [TestMethod]
    public async Task ConvertsWithoutTouchingTheSourceOrEarlierOutputs()
    {
        string input = Path.Combine(workingDirectory, "한글 Book.xlsx");
        string outputDirectory = Path.Combine(workingDirectory, "out");
        CreateWorkbook(input, cachedValue: "42");
        byte[] sourceHash = SHA256.HashData(await File.ReadAllBytesAsync(input));

        ProcessResult first = await Run("values", "--output", outputDirectory, input);
        ProcessResult second = await Run("values", "--output", outputDirectory, input);

        Assert.AreEqual(0, first.ExitCode);
        Assert.AreEqual(0, second.ExitCode);
        Assert.IsEmpty(first.StandardError);
        using JsonDocument firstDocument = ParseSingleDocument(first);
        using JsonDocument secondDocument = ParseSingleDocument(second);
        JsonElement root = firstDocument.RootElement;
        Assert.IsTrue(root.GetProperty("ok").GetBoolean());
        Assert.AreEqual(JsonValueKind.Null, root.GetProperty("error").ValueKind);
        Assert.AreEqual(1, root.GetProperty("summary").GetProperty("succeeded").GetInt32());

        JsonElement file = root.GetProperty("files")[0];
        Assert.AreEqual(input, file.GetProperty("input").GetString());
        Assert.AreEqual("converted", file.GetProperty("status").GetString());
        Assert.AreEqual("values", root.GetProperty("command").GetString());
        Assert.AreEqual(1, file.GetProperty("changes").GetProperty("removedFormulas").GetInt32());
        string firstOutput = file.GetProperty("output").GetString()!;
        string secondOutput = secondDocument.RootElement
            .GetProperty("files")[0].GetProperty("output").GetString()!;

        Assert.AreNotEqual(firstOutput, secondOutput);
        Assert.IsTrue(File.Exists(firstOutput));
        Assert.IsTrue(File.Exists(secondOutput));
        Assert.AreEqual(outputDirectory, Path.GetDirectoryName(firstOutput));
        Assert.IsNull(ReadFormula(firstOutput));
        Assert.AreEqual("42", ReadValue(firstOutput));
        CollectionAssert.AreEqual(sourceHash, SHA256.HashData(await File.ReadAllBytesAsync(input)));
        Assert.AreEqual("6*7", ReadFormula(input));
    }

    [TestMethod]
    public async Task MissingCachedValueFailsTheFileWithoutPublishingOutput()
    {
        string input = Path.Combine(workingDirectory, "Uncalculated.xlsx");
        string outputDirectory = Path.Combine(workingDirectory, "out");
        CreateWorkbook(input, cachedValue: null);

        ProcessResult result = await Run("values", "--output", outputDirectory, input);

        Assert.AreEqual(1, result.ExitCode);
        Assert.IsEmpty(result.StandardError);
        using JsonDocument document = ParseSingleDocument(result);
        JsonElement root = document.RootElement;
        Assert.IsFalse(root.GetProperty("ok").GetBoolean());
        JsonElement file = root.GetProperty("files")[0];
        Assert.AreEqual("failed", file.GetProperty("status").GetString());
        Assert.AreEqual(JsonValueKind.Null, file.GetProperty("output").ValueKind);
        Assert.AreEqual(
            "missing-cached-value",
            file.GetProperty("failure").GetProperty("kind").GetString());
        Assert.AreEqual(JsonValueKind.Null, file.GetProperty("changes").ValueKind);
        Assert.IsFalse(
            Directory.Exists(outputDirectory) && Directory.EnumerateFileSystemEntries(outputDirectory).Any());
    }

    [TestMethod]
    public async Task UsageErrorsExitTwoWithAStructuredError()
    {
        ProcessResult result = await Run("--no-such-option");

        Assert.AreEqual(2, result.ExitCode);
        Assert.IsEmpty(result.StandardError);
        using JsonDocument document = ParseSingleDocument(result);
        Assert.AreEqual(
            "usage",
            document.RootElement.GetProperty("error").GetProperty("kind").GetString());
    }

    [TestMethod]
    public async Task FlatThreeXFormIsAUsageErrorNamingTheCommand()
    {
        ProcessResult result = await Run("--output", "out", "Book.xlsx");

        Assert.AreEqual(2, result.ExitCode);
        Assert.IsEmpty(result.StandardError);
        using JsonDocument document = ParseSingleDocument(result);
        JsonElement root = document.RootElement;
        Assert.AreEqual(JsonValueKind.Null, root.GetProperty("command").ValueKind);
        StringAssert.Contains(
            root.GetProperty("error").GetProperty("message").GetString(),
            "xlclean values");
    }

    private static JsonDocument ParseSingleDocument(ProcessResult result)
    {
        string text = result.StandardOutputText;
        Assert.EndsWith("\n", text);
        string body = text[..^1];
        Assert.DoesNotContain("\n", body);
        Assert.DoesNotContain("\r", body);
        return JsonDocument.Parse(body);
    }

    private static async Task<ProcessResult> Run(params string[] arguments)
    {
        ProcessStartInfo startInfo = new()
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            RedirectStandardInput = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };

        if (Environment.GetEnvironmentVariable(ExecutableVariable) is { Length: > 0 } executable)
        {
            startInfo.FileName = executable;
        }
        else
        {
            startInfo.FileName = "dotnet";
            startInfo.ArgumentList.Add(Path.Combine(AppContext.BaseDirectory, "xlclean.dll"));
        }

        foreach (string argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using Process process = Process.Start(startInfo) ??
            throw new InvalidOperationException($"Could not start '{startInfo.FileName}'.");
        process.StandardInput.Close();

        // Read raw bytes: the contract fixes the encoding and line endings, so
        // the test must not let a decoder or the console normalize them.
        using MemoryStream standardOutput = new();
        using MemoryStream standardError = new();
        using CancellationTokenSource timeout = new(TimeSpan.FromMinutes(2));
        try
        {
            await Task.WhenAll(
                process.StandardOutput.BaseStream.CopyToAsync(standardOutput, timeout.Token),
                process.StandardError.BaseStream.CopyToAsync(standardError, timeout.Token));
            await process.WaitForExitAsync(timeout.Token);
        }
        catch (OperationCanceledException)
        {
            // A hung child would otherwise keep the working directory locked
            // and hide the timeout behind a cleanup failure.
            process.Kill(entireProcessTree: true);
            throw;
        }

        return new ProcessResult(process.ExitCode, standardOutput.ToArray(), standardError.ToArray());
    }

    private static void CreateWorkbook(string path, string? cachedValue)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Create(
            path,
            SpreadsheetDocumentType.Workbook);
        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        Cell cell = new()
        {
            CellReference = "A1",
            CellFormula = new CellFormula("6*7"),
        };
        if (cachedValue is not null)
        {
            cell.CellValue = new CellValue(cachedValue);
        }

        Row row = new() { RowIndex = 1 };
        row.Append(cell);
        worksheetPart.Worksheet = new Worksheet(new SheetData(row));
        workbookPart.Workbook.AppendChild(new Sheets(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Data",
        }));
    }

    private static string? ReadFormula(string path) => ReadFirstCell(path).CellFormula?.Text;

    private static string? ReadValue(string path) => ReadFirstCell(path).CellValue?.Text;

    private static Cell ReadFirstCell(string path)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(path, isEditable: false);
        Cell cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!
            .Descendants<Cell>().Single();
        return (Cell)cell.CloneNode(deep: true);
    }

    private sealed record ProcessResult(int ExitCode, byte[] StandardOutput, byte[] StandardError)
    {
        public string StandardOutputText { get; } =
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true)
                .GetString(StandardOutput);
    }
}
