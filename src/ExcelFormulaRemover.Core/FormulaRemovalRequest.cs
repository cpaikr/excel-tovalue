namespace ExcelFormulaRemover.Core;

/// <summary>
/// Describes one immutable batch of workbook conversions.
/// </summary>
public sealed class FormulaRemovalRequest
{
    /// <summary>
    /// The largest row number supported by the Excel worksheet format.
    /// </summary>
    public const int MaximumRowNumber = 1_048_576;

    /// <summary>
    /// Initializes a new formula-removal request.
    /// </summary>
    /// <param name="inputPaths">Workbook paths to process in order.</param>
    /// <param name="outputDirectory">Directory in which outputs are created.</param>
    /// <param name="startRow">First worksheet row whose cell formulas are converted.</param>
    /// <param name="excludedSheetNamePatterns">
    /// Case-insensitive substrings identifying sheets that remain unchanged.
    /// </param>
    /// <exception cref="ArgumentException">
    /// An input path is blank, the input sequence is empty, or the output directory is invalid.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="startRow"/> is outside the Excel worksheet range.
    /// </exception>
    public FormulaRemovalRequest(
        IEnumerable<string> inputPaths,
        string outputDirectory,
        int startRow = 1,
        IEnumerable<string>? excludedSheetNamePatterns = null)
    {
        ArgumentNullException.ThrowIfNull(inputPaths);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);

        string[] inputs = inputPaths.ToArray();
        if (inputs.Length == 0)
        {
            throw new ArgumentException("At least one input workbook is required.", nameof(inputPaths));
        }

        if (inputs.Any(string.IsNullOrWhiteSpace))
        {
            throw new ArgumentException("Input workbook paths cannot be blank.", nameof(inputPaths));
        }

        if (startRow is < 1 or > MaximumRowNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(startRow),
                startRow,
                $"The start row must be between 1 and {MaximumRowNumber:N0}.");
        }

        string[] patterns = (excludedSheetNamePatterns ?? [])
            .Select(pattern => pattern?.Trim())
            .Where(pattern => !string.IsNullOrWhiteSpace(pattern))
            .Select(pattern => pattern!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        InputPaths = Array.AsReadOnly(inputs);
        OutputDirectory = Path.GetFullPath(outputDirectory);
        StartRow = startRow;
        ExcludedSheetNamePatterns = Array.AsReadOnly(patterns);
    }

    /// <summary>
    /// Gets the workbook paths to process in order.
    /// </summary>
    public IReadOnlyList<string> InputPaths { get; }

    /// <summary>
    /// Gets the destination directory.
    /// </summary>
    public string OutputDirectory { get; }

    /// <summary>
    /// Gets the first row whose cell formulas are converted.
    /// </summary>
    public int StartRow { get; }

    /// <summary>
    /// Gets the case-insensitive sheet-name substrings that remain unchanged.
    /// </summary>
    public IReadOnlyList<string> ExcludedSheetNamePatterns { get; }
}
