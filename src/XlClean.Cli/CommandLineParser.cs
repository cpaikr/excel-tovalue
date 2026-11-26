namespace XlClean.Cli;

using System.Diagnostics;
using System.Globalization;
using XlClean.Core;

internal static class CommandLineParser
{
    public static CommandLineParseResult Parse(IReadOnlyList<string> arguments)
    {
        ArgumentNullException.ThrowIfNull(arguments);

        // A bare invocation is the root help, so there is one command inventory.
        if (arguments.Count == 0)
        {
            return CommandLineParseResult.Help();
        }

        // Help wins over version, and both win over malformed arguments, so
        // discovery never depends on the rest of the command line being valid.
        bool versionRequested = false;
        foreach (string argument in arguments)
        {
            if (argument == "--")
            {
                break;
            }

            if (argument is "--help" or "-h")
            {
                return CommandLineParseResult.Help();
            }

            versionRequested |= argument == "--version";
        }

        if (versionRequested)
        {
            return CommandLineParseResult.Version();
        }

        string? outputDirectory = null;
        int startRow = 1;
        bool startRowWasSet = false;
        List<string> excludedPatterns = [];
        List<string> inputPaths = [];
        bool optionsEnded = false;

        for (int index = 0; index < arguments.Count; index++)
        {
            string argument = arguments[index];
            if (!optionsEnded && argument == "--")
            {
                optionsEnded = true;
                continue;
            }

            if (!optionsEnded && TrySplitLongOption(argument, out string option, out string? inlineValue))
            {
                CommandLineParseResult? failure = ParseLongOption(
                    option,
                    inlineValue,
                    arguments,
                    ref index,
                    ref outputDirectory,
                    ref startRow,
                    ref startRowWasSet,
                    excludedPatterns);
                if (failure is not null)
                {
                    return failure;
                }

                continue;
            }

            if (!optionsEnded && argument is "-o" or "-r" or "-x")
            {
                string? value = ReadFollowingValue(arguments, ref index);
                if (value is null)
                {
                    return CommandLineParseResult.Failure($"Option '{argument}' requires a value.");
                }

                CommandLineParseResult? failure = argument switch
                {
                    "-o" => SetOutputDirectory(value, ref outputDirectory),
                    "-r" => SetStartRow(value, ref startRow, ref startRowWasSet),
                    "-x" => AddExcludedPattern(value, excludedPatterns),
                    _ => throw new UnreachableException(),
                };
                if (failure is not null)
                {
                    return failure;
                }

                continue;
            }

            if (!optionsEnded && argument.StartsWith('-'))
            {
                return CommandLineParseResult.Failure($"Unknown option '{argument}'.");
            }

            inputPaths.Add(argument);
        }

        if (outputDirectory is null)
        {
            return CommandLineParseResult.Failure("Required option '--output' was not provided.");
        }

        if (inputPaths.Count == 0)
        {
            return CommandLineParseResult.Failure("At least one input workbook is required.");
        }

        return CommandLineParseResult.Success(new CommandLineOptions(
            inputPaths.AsReadOnly(),
            outputDirectory,
            startRow,
            excludedPatterns.AsReadOnly()));
    }

    private static CommandLineParseResult? ParseLongOption(
        string option,
        string? inlineValue,
        IReadOnlyList<string> arguments,
        ref int index,
        ref string? outputDirectory,
        ref int startRow,
        ref bool startRowWasSet,
        List<string> excludedPatterns)
    {
        if (option is not ("--output" or "--start-row" or "--exclude"))
        {
            return CommandLineParseResult.Failure($"Unknown option '{option}'.");
        }

        string? value = inlineValue ?? ReadFollowingValue(arguments, ref index);
        if (string.IsNullOrWhiteSpace(value))
        {
            return CommandLineParseResult.Failure($"Option '{option}' requires a value.");
        }

        return option switch
        {
            "--output" => SetOutputDirectory(value, ref outputDirectory),
            "--start-row" => SetStartRow(value, ref startRow, ref startRowWasSet),
            "--exclude" => AddExcludedPattern(value, excludedPatterns),
            _ => throw new UnreachableException(),
        };
    }

    private static CommandLineParseResult? SetOutputDirectory(
        string value,
        ref string? outputDirectory)
    {
        if (outputDirectory is not null)
        {
            return CommandLineParseResult.Failure("Option '--output' was provided more than once.");
        }

        outputDirectory = value;
        return null;
    }

    private static CommandLineParseResult? SetStartRow(
        string value,
        ref int startRow,
        ref bool startRowWasSet)
    {
        if (startRowWasSet)
        {
            return CommandLineParseResult.Failure("Option '--start-row' was provided more than once.");
        }

        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) ||
            parsed is < 1 or > FormulaRemovalRequest.MaximumRowNumber)
        {
            return CommandLineParseResult.Failure(
                $"Start row must be an integer from 1 to {FormulaRemovalRequest.MaximumRowNumber}.");
        }

        startRow = parsed;
        startRowWasSet = true;
        return null;
    }

    private static CommandLineParseResult? AddExcludedPattern(
        string value,
        List<string> excludedPatterns)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return CommandLineParseResult.Failure("Excluded sheet patterns cannot be blank.");
        }

        excludedPatterns.Add(value);
        return null;
    }

    private static bool TrySplitLongOption(
        string argument,
        out string option,
        out string? inlineValue)
    {
        option = argument;
        inlineValue = null;

        if (!argument.StartsWith("--", StringComparison.Ordinal))
        {
            return false;
        }

        int equalsIndex = argument.IndexOf('=', StringComparison.Ordinal);
        if (equalsIndex >= 0)
        {
            option = argument[..equalsIndex];
            inlineValue = argument[(equalsIndex + 1)..];
        }

        return true;
    }

    private static string? ReadFollowingValue(IReadOnlyList<string> arguments, ref int index)
    {
        if (index + 1 >= arguments.Count)
        {
            return null;
        }

        string value = arguments[index + 1];
        if (value.StartsWith('-'))
        {
            return null;
        }

        index++;
        return value;
    }
}

internal sealed record CommandLineOptions(
    IReadOnlyList<string> InputPaths,
    string OutputDirectory,
    int StartRow,
    IReadOnlyList<string> ExcludedSheetNamePatterns);

internal enum CommandLineParseKind
{
    Success,
    Help,
    Version,
    Failure,
}

internal sealed record CommandLineParseResult(
    CommandLineParseKind Kind,
    CommandLineOptions? Options,
    string? ErrorMessage)
{
    public static CommandLineParseResult Success(CommandLineOptions options) =>
        new(CommandLineParseKind.Success, options, null);

    public static CommandLineParseResult Help() =>
        new(CommandLineParseKind.Help, null, null);

    public static CommandLineParseResult Version() =>
        new(CommandLineParseKind.Version, null, null);

    public static CommandLineParseResult Failure(string message) =>
        new(CommandLineParseKind.Failure, null, message);
}
