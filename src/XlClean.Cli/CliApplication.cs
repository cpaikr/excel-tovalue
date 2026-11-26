namespace XlClean.Cli;

using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Security;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using XlClean.Core;

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

    private const string CommandName = "xlclean";

    private const string RootHelpText =
        """
        Clean Excel workbooks into new files without modifying the originals.

        Usage:
          xlclean <command> [options] <workbook>...
          xlclean <command> --help

        Commands:
          values  Replace formulas with their last saved cached values.

        Options:
          -h, --help     Show this help, or a command's help after its name.
              --version  Show the version.

        Output: every command run writes one JSON document on stdout, followed by
        a newline; stderr stays empty.
          ok       true only when every workbook succeeded (exit code 0).
          version  The version of this executable.
          command  The command that ran, or null when no valid command was given.
          files    One entry per attempted workbook, in request order: input, status
                   ("converted" or "failed"), output (the published path), failure
                   (kind and message), changes (counts specific to the command,
                   see its help; null for a failed workbook), warnings.
          summary  total, succeeded, failed, canceled.
          error    Set when the invocation itself failed (kind "usage" or
                   "unexpected"); files is then empty.

        Exit codes:
          0  Every workbook succeeded, or help or version was printed.
          1  One or more workbooks failed, or the run failed (see error).
          2  Command-line usage error.
        130  Canceled; files completed before cancellation are still reported.

        Sources and existing outputs are never overwritten, so read each output
        path from the result instead of predicting its name.

        """;

    private const string ValuesHelpText =
        """
        Replace Excel formulas with their last saved cached values.

        Usage:
          xlclean values --output <directory> [options] <workbook>...

        Options:
          -o, --output <directory>  Required destination directory.
          -r, --start-row <number>  First row to convert (default: 1).
          -x, --exclude <pattern>   Keep sheets whose name contains this text; repeatable.
          -h, --help                Show this help.
              --version             Show the version.
          --                        Treat every later argument as a workbook path.

        Input: .xlsx, .xlsm, and .xlsb workbooks. Each output keeps its input format.

        changes: removedFormulas, removedTableFormulaDefinitions, processedSheets,
          excludedSheets.

        Failure kinds: invalid-input, missing-cached-value,
          formula-range-crosses-boundary, invalid-workbook, file-access, unexpected.

        The tool does not calculate formulas. Recalculate and save workbooks in
        Excel before use.

        Run 'xlclean --help' for the output document and exit codes.

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
            string help = parseResult.Command switch
            {
                null => RootHelpText,
                CliCommand.Values => ValuesHelpText,
                _ => throw new UnreachableException(),
            };
            await standardOutput.WriteAsync(help.ReplaceLineEndings("\n")).ConfigureAwait(false);
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
                parseResult.Command,
                "usage",
                parseResult.ErrorMessage!,
                HelpHint(parseResult.Command)).ConfigureAwait(false);
            return UsageErrorExitCode;
        }

        CliCommand command = parseResult.Command!.Value;
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
                command,
                "usage",
                exception.Message,
                HelpHint(command)).ConfigureAwait(false);
            return UsageErrorExitCode;
        }

        try
        {
            FormulaRemovalBatchResult result = await converter(
                request,
                null,
                cancellationToken).ConfigureAwait(false);

            await WriteBatchResult(standardOutput, command, result, canceled: false).ConfigureAwait(false);
            return result.FailedCount == 0 ? SuccessExitCode : ConversionFailureExitCode;
        }
        catch (FormulaRemovalCanceledException exception) when (cancellationToken.IsCancellationRequested)
        {
            await WriteBatchResult(
                standardOutput,
                command,
                exception.CompletedResult,
                canceled: true).ConfigureAwait(false);
            return CanceledExitCode;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            await WriteBatchResult(standardOutput, command, new([]), canceled: true).ConfigureAwait(false);
            return CanceledExitCode;
        }
        catch (Exception exception)
        {
            await WriteInvocationError(
                standardOutput,
                command,
                "unexpected",
                exception.Message,
                "Check paths and permissions, then retry.").ConfigureAwait(false);
            return ConversionFailureExitCode;
        }
    }

    private static Task WriteBatchResult(
        TextWriter output,
        CliCommand command,
        FormulaRemovalBatchResult result,
        bool canceled) =>
        WriteDocument(output, writer =>
        {
            writer.WriteBoolean("ok", !canceled && result.FailedCount == 0);
            writer.WriteString("version", Version);
            writer.WriteString("command", FormatCommand(command));
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

                // A failed file published nothing, so it has no changes to report.
                if (file.Succeeded)
                {
                    writer.WriteStartObject("changes");
                    writer.WriteNumber("removedFormulas", file.RemovedFormulaCount);
                    writer.WriteNumber(
                        "removedTableFormulaDefinitions",
                        file.RemovedTableFormulaDefinitionCount);
                    writer.WriteNumber("processedSheets", file.ProcessedSheetCount);
                    writer.WriteNumber("excludedSheets", file.ExcludedSheetCount);
                    writer.WriteEndObject();
                }
                else
                {
                    writer.WriteNull("changes");
                }
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
        CliCommand? command,
        string kind,
        string message,
        string help) =>
        WriteDocument(output, writer =>
        {
            writer.WriteBoolean("ok", false);
            writer.WriteString("version", Version);
            if (command is { } known)
            {
                writer.WriteString("command", FormatCommand(known));
            }
            else
            {
                writer.WriteNull("command");
            }

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

    private static string HelpHint(CliCommand? command) =>
        command is { } known
            ? $"Run {CommandName} {FormatCommand(known)} --help"
            : $"Run {CommandName} --help";

    private static string FormatCommand(CliCommand command) => CommandLineParser.NameOf(command);

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
