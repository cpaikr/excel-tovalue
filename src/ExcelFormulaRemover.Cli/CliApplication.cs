namespace ExcelFormulaRemover.Cli;

using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Security;
using System.Text;
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

    private const string HelpText =
        """
        Remove Excel formulas by replacing them with their last saved cached values.

        Usage:
          excel-formula-remover --output <directory> [options] <workbook>...

        Options:
          -o, --output <directory>  Required destination directory.
          -r, --start-row <number>  First row to convert (default: 1).
          -x, --exclude <pattern>    Keep sheets containing this text; repeatable.
              --json                 Write detailed NDJSON instead of compact TOON.
          -h, --help                 Show this help.

        Exit codes:
          0  Every workbook converted successfully.
          1  One or more conversions failed.
          2  Command-line usage error.
        130  Canceled by the user.

        The tool does not calculate formulas. Save recalculated workbooks before use.
        Existing outputs are never overwritten.
        Default structured output follows the TOON 4.1 specification.
        """;

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
        standardOutput.NewLine = "\n";
        standardError.NewLine = "\n";

        bool jsonRequested = CommandLineParser.RequestsJson(arguments);
        CommandLineParseResult parseResult = CommandLineParser.Parse(arguments);
        if (parseResult.Kind == CommandLineParseKind.Help)
        {
            await standardOutput.WriteAsync(HelpText.ReplaceLineEndings("\n")).ConfigureAwait(false);
            return SuccessExitCode;
        }

        if (parseResult.Kind == CommandLineParseKind.Failure)
        {
            await WriteUsageError(
                standardOutput,
                parseResult.ErrorMessage!,
                jsonRequested).ConfigureAwait(false);
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
            await WriteUsageError(
                standardOutput,
                exception.Message,
                options.UseJson).ConfigureAwait(false);
            return UsageErrorExitCode;
        }

        try
        {
            FormulaRemovalBatchResult result = await converter(
                request,
                null,
                cancellationToken).ConfigureAwait(false);

            if (options.UseJson)
            {
                await WriteJsonResult(standardOutput, result, canceled: false).ConfigureAwait(false);
            }
            else
            {
                await WriteToonResult(standardOutput, result, canceled: false).ConfigureAwait(false);
            }

            return result.FailedCount == 0 ? SuccessExitCode : ConversionFailureExitCode;
        }
        catch (FormulaRemovalCanceledException exception) when (cancellationToken.IsCancellationRequested)
        {
            if (options.UseJson)
            {
                await WriteJsonResult(
                    standardOutput,
                    exception.CompletedResult,
                    canceled: true).ConfigureAwait(false);
            }
            else
            {
                await WriteToonResult(
                    standardOutput,
                    exception.CompletedResult,
                    canceled: true).ConfigureAwait(false);
            }

            return CanceledExitCode;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            FormulaRemovalBatchResult emptyResult = new([]);
            if (options.UseJson)
            {
                await WriteJsonResult(standardOutput, emptyResult, canceled: true).ConfigureAwait(false);
            }
            else
            {
                await WriteToonResult(standardOutput, emptyResult, canceled: true).ConfigureAwait(false);
            }

            return CanceledExitCode;
        }
        catch (Exception exception)
        {
            if (options.UseJson)
            {
                await standardOutput.WriteAsync(JsonSerializer.Serialize(new
                {
                    type = "fatal",
                    errorType = exception.GetType().Name,
                    message = exception.Message,
                })).ConfigureAwait(false);
            }
            else
            {
                await standardOutput.WriteAsync(string.Join(
                    '\n',
                    "error:",
                    "  kind: conversion",
                    $"  type: {QuoteToonString(exception.GetType().Name)}",
                    $"  message: {QuoteToonString(exception.Message)}",
                    $"help: {QuoteToonString("Check paths and permissions, then retry.")}"))
                    .ConfigureAwait(false);
            }

            return ConversionFailureExitCode;
        }
    }

    private static async Task WriteToonResult(
        TextWriter output,
        FormulaRemovalBatchResult result,
        bool canceled)
    {
        List<string> lines = [];
        if (result.Files.Count == 0)
        {
            lines.Add("files: []");
        }
        else
        {
            lines.Add($"files[{Invariant(result.Files.Count)}]{{index,input,status,output,failure}}:");
            for (int index = 0; index < result.Files.Count; index++)
            {
                FormulaRemovalFileResult file = result.Files[index];
                string status = file.Succeeded ? "ok" : "failed";
                string outputPath = file.OutputPath is null ? "null" : QuoteToonString(file.OutputPath);
                string failure = file.Succeeded ? "null" : FormatFailureKind(file.FailureKind);
                lines.Add(
                    $"  {Invariant(index + 1)},{QuoteToonString(file.InputPath)}," +
                    $"{status},{outputPath},{failure}");
            }
        }

        int messageCount = result.Files.Sum(file =>
            file.Warnings.Count + (file.Succeeded ? 0 : 1));
        if (messageCount == 0)
        {
            lines.Add("messages: []");
        }
        else
        {
            lines.Add($"messages[{Invariant(messageCount)}]{{index,level,text}}:");
            for (int index = 0; index < result.Files.Count; index++)
            {
                FormulaRemovalFileResult file = result.Files[index];
                if (!file.Succeeded)
                {
                    lines.Add(
                        $"  {Invariant(index + 1)},error," +
                        QuoteToonString(file.ErrorMessage ?? string.Empty));
                }

                lines.AddRange(file.Warnings.Select(warning =>
                    $"  {Invariant(index + 1)},warning,{QuoteToonString(warning)}"));
            }
        }

        lines.Add("summary:");
        lines.Add($"  total: {Invariant(result.Files.Count)}");
        lines.Add($"  succeeded: {Invariant(result.SucceededCount)}");
        lines.Add($"  failed: {Invariant(result.FailedCount)}");
        lines.Add($"  canceled: {canceled.ToString().ToLowerInvariant()}");
        await output.WriteAsync(string.Join('\n', lines)).ConfigureAwait(false);
    }

    private static async Task WriteJsonResult(
        TextWriter output,
        FormulaRemovalBatchResult result,
        bool canceled)
    {
        List<string> lines = [];
        foreach (FormulaRemovalFileResult file in result.Files)
        {
            lines.Add(JsonSerializer.Serialize(new
            {
                type = "result",
                succeeded = file.Succeeded,
                inputPath = file.InputPath,
                outputPath = file.OutputPath,
                failureKind = FormatFailureKind(file.FailureKind),
                errorMessage = file.ErrorMessage,
                removedFormulaCount = file.RemovedFormulaCount,
                removedTableFormulaDefinitionCount = file.RemovedTableFormulaDefinitionCount,
                processedSheetCount = file.ProcessedSheetCount,
                excludedSheetCount = file.ExcludedSheetCount,
                warnings = file.Warnings,
            }));
        }

        lines.Add(JsonSerializer.Serialize(new
        {
            type = "summary",
            total = result.Files.Count,
            succeeded = result.SucceededCount,
            failed = result.FailedCount,
            canceled,
        }));
        await output.WriteAsync(string.Join('\n', lines)).ConfigureAwait(false);
    }

    private static async Task WriteUsageError(TextWriter output, string message, bool useJson)
    {
        if (useJson)
        {
            await output.WriteAsync(JsonSerializer.Serialize(new
            {
                type = "error",
                kind = "usage",
                message,
                help = "Run excel-formula-remover --help",
            })).ConfigureAwait(false);
            return;
        }

        await output.WriteAsync(string.Join(
            '\n',
            "error:",
            "  kind: usage",
            $"  message: {QuoteToonString(message)}",
            $"help: {QuoteToonString("Run excel-formula-remover --help")}"))
            .ConfigureAwait(false);
    }

    private static string FormatFailureKind(FormulaRemovalFailureKind failureKind) =>
        failureKind switch
        {
            FormulaRemovalFailureKind.None => "none",
            FormulaRemovalFailureKind.InvalidInput => "invalid-input",
            FormulaRemovalFailureKind.MissingCachedValue => "missing-cached-value",
            FormulaRemovalFailureKind.FormulaRangeCrossesBoundary => "formula-range-crosses-boundary",
            FormulaRemovalFailureKind.InvalidWorkbook => "invalid-workbook",
            FormulaRemovalFailureKind.FileAccess => "file-access",
            FormulaRemovalFailureKind.Unexpected => "unexpected",
            _ => throw new ArgumentOutOfRangeException(nameof(failureKind), failureKind, null),
        };

    private static string Invariant(int value) =>
        value.ToString(CultureInfo.InvariantCulture);

    private static string QuoteToonString(string value)
    {
        StringBuilder builder = new(value.Length + 2);
        builder.Append('"');

        for (int index = 0; index < value.Length; index++)
        {
            char character = value[index];
            switch (character)
            {
                case '\\':
                    builder.Append("\\\\");
                    break;
                case '"':
                    builder.Append("\\\"");
                    break;
                case '\n':
                    builder.Append("\\n");
                    break;
                case '\r':
                    builder.Append("\\r");
                    break;
                case '\t':
                    builder.Append("\\t");
                    break;
                default:
                    if (character < ' ')
                    {
                        builder.Append("\\u");
                        builder.Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                    }
                    else if (char.IsHighSurrogate(character))
                    {
                        if (index + 1 >= value.Length || !char.IsLowSurrogate(value[index + 1]))
                        {
                            throw new ArgumentException(
                                "TOON strings cannot contain unpaired UTF-16 surrogates.",
                                nameof(value));
                        }

                        builder.Append(character);
                        builder.Append(value[++index]);
                    }
                    else if (char.IsLowSurrogate(character))
                    {
                        throw new ArgumentException(
                            "TOON strings cannot contain unpaired UTF-16 surrogates.",
                            nameof(value));
                    }
                    else
                    {
                        builder.Append(character);
                    }

                    break;
            }
        }

        builder.Append('"');
        return builder.ToString();
    }
}
