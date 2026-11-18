namespace ExcelFormulaRemover.Cli;

using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Security;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using ExcelFormulaRemover.Core;

internal delegate Task<FormulaRemovalBatchResult> FormulaRemovalConverter(
    FormulaRemovalRequest request,
    IProgress<FormulaRemovalProgress>? progress,
    CancellationToken cancellationToken);

internal static class CliApplication
{
    public const int SuccessExitCode = 0;
    public const int ConversionFailureExitCode = 1;
    public const int UsageErrorExitCode = 2;
    public const int CanceledExitCode = 130;

    private const string CommandName = "excel-formula-remover";

    private const string HelpText =
        """
        Remove Excel formulas by replacing them with their last saved cached values.

        Usage:
          excel-formula-remover --output <directory> [options] <workbook>...

        Options:
          -o, --output <directory>  Required destination directory.
          -r, --start-row <number>  First row to convert (default: 1).
          -x, --exclude <pattern>   Keep sheets whose name contains this text; repeatable.
          -h, --help                Show this help.
              --version             Show the version.
          --                        Treat every later argument as a workbook path.

        Input: .xlsx, .xlsm, and .xlsb workbooks. Each output keeps its input format.

        Output: one JSON document on stdout, followed by a newline; stderr stays empty.
          ok       true only when every workbook converted (exit code 0).
          version  The version of this executable.
          files    One entry per attempted workbook, in request order: input, status
                   ("converted" or "failed"), output (the published path), failure
                   (kind and message), removed-formula and sheet counts, warnings.
          summary  total, succeeded, failed, canceled.
          error    Set when the invocation itself failed (kind "usage" or
                   "unexpected"); files is then empty.

        Failure kinds: invalid-input, missing-cached-value,
          formula-range-crosses-boundary, invalid-workbook, file-access, unexpected.

        Exit codes:
          0  Every workbook converted, or help or version was printed.
          1  One or more conversions failed, or the run failed (see error).
          2  Command-line usage error.
        130  Canceled; files completed before cancellation are still reported.

        The tool does not calculate formulas. Recalculate and save workbooks in
        Excel before use. Sources and existing outputs are never overwritten, so
        read each output path from the result instead of predicting its name.

        """;

    // Paths and messages are routinely Korean and quote option names. The
    // relaxed encoder leaves both unescaped, which keeps the document readable
    // to people and agents; it is only unsafe when embedded in HTML, and stdout
    // is not.
    private static readonly JsonWriterOptions WriterOptions = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static string Version { get; } = ReadVersion();

    public static Task<int> RunAsync(
        IReadOnlyList<string> arguments,
        TextWriter standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken) =>
        RunAsync(
            arguments,
            standardOutput,
            standardError,
            FormulaRemovalEngine.ConvertAsync,
            cancellationToken);

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "The process boundary reports unexpected failures without emitting unstable stack traces.")]
    internal static async Task<int> RunAsync(
        IReadOnlyList<string> arguments,
        TextWriter standardOutput,
        TextWriter standardError,
        FormulaRemovalConverter converter,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentNullException.ThrowIfNull(standardOutput);
        ArgumentNullException.ThrowIfNull(standardError);
        ArgumentNullException.ThrowIfNull(converter);

        CommandLineParseResult parseResult = CommandLineParser.Parse(arguments);
        if (parseResult.Kind == CommandLineParseKind.Help)
        {
            await standardOutput.WriteAsync(HelpText.ReplaceLineEndings("\n")).ConfigureAwait(false);
            return SuccessExitCode;
        }

        if (parseResult.Kind == CommandLineParseKind.Version)
        {
            await standardOutput.WriteAsync($"{CommandName} {Version}\n").ConfigureAwait(false);
            return SuccessExitCode;
        }

        if (parseResult.Kind == CommandLineParseKind.Failure)
        {
            await WriteInvocationError(
                standardOutput,
                "usage",
                parseResult.ErrorMessage!,
                $"Run {CommandName} --help").ConfigureAwait(false);
            return UsageErrorExitCode;
        }

        CommandLineOptions options = parseResult.Options!;
        FormulaRemovalRequest request;
        try
        {
            request = new FormulaRemovalRequest(
                options.InputPaths,
                options.OutputDirectory,
                options.StartRow,
                options.ExcludedSheetNamePatterns);
        }
        catch (Exception exception) when (
            exception is ArgumentException or NotSupportedException or PathTooLongException or SecurityException)
        {
            await WriteInvocationError(
                standardOutput,
                "usage",
                exception.Message,
                $"Run {CommandName} --help").ConfigureAwait(false);
            return UsageErrorExitCode;
        }

        try
        {
            FormulaRemovalBatchResult result = await converter(
                request,
                null,
                cancellationToken).ConfigureAwait(false);

            await WriteBatchResult(standardOutput, result, canceled: false).ConfigureAwait(false);
            return result.FailedCount == 0 ? SuccessExitCode : ConversionFailureExitCode;
        }
        catch (FormulaRemovalCanceledException exception) when (cancellationToken.IsCancellationRequested)
        {
            await WriteBatchResult(
                standardOutput,
                exception.CompletedResult,
                canceled: true).ConfigureAwait(false);
            return CanceledExitCode;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            await WriteBatchResult(standardOutput, new([]), canceled: true).ConfigureAwait(false);
            return CanceledExitCode;
        }
        catch (Exception exception)
        {
            await WriteInvocationError(
                standardOutput,
                "unexpected",
                exception.Message,
                "Check paths and permissions, then retry.").ConfigureAwait(false);
            return ConversionFailureExitCode;
        }
    }

    private static Task WriteBatchResult(
        TextWriter output,
        FormulaRemovalBatchResult result,
        bool canceled) =>
        WriteDocument(output, writer =>
        {
            writer.WriteBoolean("ok", !canceled && result.FailedCount == 0);
            writer.WriteString("version", Version);
            writer.WriteStartArray("files");
            foreach (FormulaRemovalFileResult file in result.Files)
            {
                writer.WriteStartObject();
                writer.WriteString("input", file.InputPath);
                writer.WriteString("status", file.Succeeded ? "converted" : "failed");
                writer.WriteString("output", file.OutputPath);
                if (file.Succeeded)
                {
                    writer.WriteNull("failure");
                }
                else
                {
                    writer.WriteStartObject("failure");
                    writer.WriteString("kind", FormatFailureKind(file.FailureKind));
                    writer.WriteString("message", file.ErrorMessage ?? string.Empty);
                    writer.WriteEndObject();
                }

                writer.WriteNumber("removedFormulas", file.RemovedFormulaCount);
                writer.WriteNumber(
                    "removedTableFormulaDefinitions",
                    file.RemovedTableFormulaDefinitionCount);
                writer.WriteNumber("processedSheets", file.ProcessedSheetCount);
                writer.WriteNumber("excludedSheets", file.ExcludedSheetCount);
                writer.WriteStartArray("warnings");
                foreach (string warning in file.Warnings)
                {
                    writer.WriteStringValue(warning);
                }

                writer.WriteEndArray();
                writer.WriteEndObject();
            }

            writer.WriteEndArray();
            WriteSummary(writer, result.Files.Count, result.SucceededCount, result.FailedCount, canceled);
            writer.WriteNull("error");
        });

    private static Task WriteInvocationError(
        TextWriter output,
        string kind,
        string message,
        string help) =>
        WriteDocument(output, writer =>
        {
            writer.WriteBoolean("ok", false);
            writer.WriteString("version", Version);
            writer.WriteStartArray("files");
            writer.WriteEndArray();
            WriteSummary(writer, total: 0, succeeded: 0, failed: 0, canceled: false);
            writer.WriteStartObject("error");
            writer.WriteString("kind", kind);
            writer.WriteString("message", message);
            writer.WriteString("help", help);
            writer.WriteEndObject();
        });

    private static void WriteSummary(
        Utf8JsonWriter writer,
        int total,
        int succeeded,
        int failed,
        bool canceled)
    {
        writer.WriteStartObject("summary");
        writer.WriteNumber("total", total);
        writer.WriteNumber("succeeded", succeeded);
        writer.WriteNumber("failed", failed);
        writer.WriteBoolean("canceled", canceled);
        writer.WriteEndObject();
    }

    // Every structured run emits exactly one document and one trailing newline,
    // built completely before the first byte is written so a failure while
    // encoding cannot leave a truncated document on stdout.
    private static async Task WriteDocument(TextWriter output, Action<Utf8JsonWriter> writeProperties)
    {
        using MemoryStream buffer = new();
        using (Utf8JsonWriter writer = new(buffer, WriterOptions))
        {
            writer.WriteStartObject();
            writeProperties(writer);
            writer.WriteEndObject();
        }

        string document = Encoding.UTF8.GetString(buffer.GetBuffer(), 0, checked((int)buffer.Length));
        await output.WriteAsync(document + "\n").ConfigureAwait(false);
    }

    private static string FormatFailureKind(FormulaRemovalFailureKind failureKind) =>
        failureKind switch
        {
            FormulaRemovalFailureKind.InvalidInput => "invalid-input",
            FormulaRemovalFailureKind.MissingCachedValue => "missing-cached-value",
            FormulaRemovalFailureKind.FormulaRangeCrossesBoundary => "formula-range-crosses-boundary",
            FormulaRemovalFailureKind.InvalidWorkbook => "invalid-workbook",
            FormulaRemovalFailureKind.FileAccess => "file-access",
            FormulaRemovalFailureKind.Unexpected => "unexpected",
            _ => throw new ArgumentOutOfRangeException(nameof(failureKind), failureKind, null),
        };

    // The SDK appends "+<commit>" to the informational version; the release
    // identity callers compare against is the part before it.
    private static string ReadVersion()
    {
        string? informationalVersion = typeof(CliApplication).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        if (string.IsNullOrWhiteSpace(informationalVersion))
        {
            return "0.0.0";
        }

        int metadataIndex = informationalVersion.IndexOf('+', StringComparison.Ordinal);
        return metadataIndex < 0 ? informationalVersion : informationalVersion[..metadataIndex];
    }
}
