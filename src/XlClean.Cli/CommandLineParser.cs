namespace XlClean.Cli;

using System.Diagnostics;
using System.Globalization;
using XlClean.Core;

internal static class CommandLineParser
{
    private static readonly Dictionary<string, CliCommand> Commands =
        Enum.GetValues<CliCommand>().ToDictionary(NameOf, StringComparer.Ordinal);

    /// <summary>The command name as typed and as reported in the JSON result.</summary>
    public static string NameOf(CliCommand command) =>
        command switch
        {
            CliCommand.Values => "values",
            CliCommand.Styles => "styles",
            _ => throw new ArgumentOutOfRangeException(nameof(command), command, null),
        };

    public static CommandLineParseResult Parse(IReadOnlyList<string> arguments)
    {
        ArgumentNullException.ThrowIfNull(arguments);

        // A bare invocation is the root help, so there is one command inventory.
        if (arguments.Count == 0)
        {
            return CommandLineParseResult.Help(null);
        }

        if (!Commands.TryGetValue(arguments[0], out CliCommand command))
        {
            return ParseRoot(arguments);
        }

        IReadOnlyList<string> commandArguments = [.. arguments.Skip(1)];
        CommandLineParseResult? discovery = ParseDiscovery(commandArguments, command);
        return discovery ?? ParseCommand(command, commandArguments);
    }

    // Without a command only discovery works. An option first is the 3.x
    // flat form, so the error names the command that replaced it.
    private static CommandLineParseResult ParseRoot(IReadOnlyList<string> arguments)
    {
        CommandLineParseResult? discovery = ParseDiscovery(arguments, null);
        if (discovery is not null)
        {
            return discovery;
        }

        string first = arguments[0];
        if (first == "--")
        {
            return CommandLineParseResult.Failure(
                "A command is required; for example 'xlclean values --output <directory> <workbook>...'.",
                null);
        }

        return first.StartsWith('-')
            ? CommandLineParseResult.Failure(
                $"A command is required before options; for example 'xlclean values {first} ...'.",
                null)
            : CommandLineParseResult.Failure($"Unknown command '{first}'.", null);
    }

    // Help wins over version, and both win over malformed arguments, so
    // discovery never depends on the rest of the command line being valid.
    private static CommandLineParseResult? ParseDiscovery(
        IReadOnlyList<string> arguments,
        CliCommand? command)
    {
        bool versionRequested = false;
        foreach (string argument in arguments)
        {
            if (argument == "--")
            {
                break;
            }

            if (argument is "--help" or "-h")
            {
                return CommandLineParseResult.Help(command);
            }

            versionRequested |= argument == "--version";
        }

        return versionRequested ? CommandLineParseResult.Version() : null;
    }

    // Every command takes --output and workbooks; only values takes a scope.
    private static CommandLineParseResult ParseCommand(CliCommand command, IReadOnlyList<string> arguments)
    {
        bool allowsScope = command == CliCommand.Values;
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
                    command,
                    allowsScope,
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

            if (!optionsEnded && (argument == "-o" || (allowsScope && argument is "-r" or "-x")))
            {
                string? value = ReadFollowingValue(arguments, ref index);
                if (value is null)
                {
                    return CommandLineParseResult.Failure(command, $"Option '{argument}' requires a value.");
                }

                string? error = argument switch
                {
                    "-o" => SetOutputDirectory(value, ref outputDirectory),
                    "-r" => SetStartRow(value, ref startRow, ref startRowWasSet),
                    "-x" => AddExcludedPattern(value, excludedPatterns),
                    _ => throw new UnreachableException(),
                };
                if (error is not null)
                {
                    return CommandLineParseResult.Failure(command, error);
                }

                continue;
            }

            if (!optionsEnded && !allowsScope && argument is "-r" or "-x")
            {
                return CommandLineParseResult.Failure(command, $"Option '{argument}' applies only to 'xlclean values'.");
            }

            if (!optionsEnded && argument.StartsWith('-'))
            {
                return CommandLineParseResult.Failure(command, $"Unknown option '{argument}'.");
            }

            inputPaths.Add(argument);
        }

        if (outputDirectory is null)
        {
            return CommandLineParseResult.Failure(command, "Required option '--output' was not provided.");
        }

        if (inputPaths.Count == 0)
        {
            return CommandLineParseResult.Failure(command, "At least one input workbook is required.");
        }

        return CommandLineParseResult.Success(command, new CommandLineOptions(
            inputPaths.AsReadOnly(),
            outputDirectory,
            startRow,
            excludedPatterns.AsReadOnly()));
    }

    private static CommandLineParseResult? ParseLongOption(
        CliCommand command,
        bool allowsScope,
        string option,
        string? inlineValue,
        IReadOnlyList<string> arguments,
        ref int index,
        ref string? outputDirectory,
        ref int startRow,
        ref bool startRowWasSet,
        List<string> excludedPatterns)
    {
        if (!allowsScope && option is "--start-row" or "--exclude")
        {
            return CommandLineParseResult.Failure(command, $"Option '{option}' applies only to 'xlclean values'.");
        }

        if (option != "--output" && !(allowsScope && option is "--start-row" or "--exclude"))
        {
            return CommandLineParseResult.Failure(command, $"Unknown option '{option}'.");
        }

        string? value = inlineValue ?? ReadFollowingValue(arguments, ref index);
        if (string.IsNullOrWhiteSpace(value))
        {
            return CommandLineParseResult.Failure(command, $"Option '{option}' requires a value.");
        }

        string? error = option switch
        {
            "--output" => SetOutputDirectory(value, ref outputDirectory),
            "--start-row" => SetStartRow(value, ref startRow, ref startRowWasSet),
            "--exclude" => AddExcludedPattern(value, excludedPatterns),
            _ => throw new UnreachableException(),
        };
        return error is null ? null : CommandLineParseResult.Failure(command, error);
    }

    private static string? SetOutputDirectory(
        string value,
        ref string? outputDirectory)
    {
        if (outputDirectory is not null)
        {
            return "Option '--output' was provided more than once.";
        }

        outputDirectory = value;
        return null;
    }

    private static string? SetStartRow(
        string value,
        ref int startRow,
        ref bool startRowWasSet)
    {
        if (startRowWasSet)
        {
            return "Option '--start-row' was provided more than once.";
        }

        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) ||
            parsed is < 1 or > FormulaRemovalRequest.MaximumRowNumber)
        {
            return $"Start row must be an integer from 1 to {FormulaRemovalRequest.MaximumRowNumber}.";
        }

        startRow = parsed;
        startRowWasSet = true;
        return null;
    }

    private static string? AddExcludedPattern(
        string value,
        List<string> excludedPatterns)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Excluded sheet patterns cannot be blank.";
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

internal enum CliCommand
{
    Values,
    Styles,
}

internal enum CommandLineParseKind
{
    Success,
    Help,
    Version,
    Failure,
}

/// <param name="Command">
/// The command the arguments selected. For help it chooses between root and
/// command help; for a failure it is null when no valid command was given.
/// </param>
internal sealed record CommandLineParseResult(
    CommandLineParseKind Kind,
    CliCommand? Command,
    CommandLineOptions? Options,
    string? ErrorMessage)
{
    public static CommandLineParseResult Success(CliCommand command, CommandLineOptions options) =>
        new(CommandLineParseKind.Success, command, options, null);

    public static CommandLineParseResult Help(CliCommand? command) =>
        new(CommandLineParseKind.Help, command, null, null);

    public static CommandLineParseResult Version() =>
        new(CommandLineParseKind.Version, null, null, null);

    public static CommandLineParseResult Failure(string message, CliCommand? command) =>
        new(CommandLineParseKind.Failure, command, null, message);

    public static CommandLineParseResult Failure(CliCommand command, string message) =>
        Failure(message, command);
}
