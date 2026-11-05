namespace ExcelFormulaRemover.Core;

/// <summary>
/// Format-independent scope rules shared by the OpenXML and binary workbook rewriters.
/// </summary>
internal static class FormulaScope
{
    public static bool IsExcluded(string sheetName, IReadOnlyList<string> patterns)
    {
        return patterns.Any(pattern =>
            sheetName.Contains(pattern, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Decides whether a multi-row formula definition is converted.
    /// </summary>
    /// <returns><see langword="true"/> when every row of the range is at or after the start row.</returns>
    /// <exception cref="FormulaRangeBoundaryException">The range crosses the start row.</exception>
    public static bool RangeIsInScope(
        string formulaKind,
        string sheetName,
        string owner,
        string reference,
        int firstRow,
        int lastRow,
        int startRow)
    {
        if (lastRow < startRow)
        {
            return false;
        }

        if (firstRow >= startRow)
        {
            return true;
        }

        throw new FormulaRangeBoundaryException(
            formulaKind,
            sheetName,
            owner,
            reference,
            startRow);
    }
}

/// <summary>
/// A multi-cell or table formula definition cannot be split safely at the start row.
/// </summary>
internal sealed class FormulaRangeBoundaryException(
    string formulaKind,
    string sheetName,
    string owner,
    string formulaRange,
    int startRow)
    : Exception(
        $"Worksheet '{sheetName}', '{owner}' owns a {formulaKind} over range '{formulaRange}', " +
        $"which crosses start row {startRow}. Choose row 1 or a row " +
        "after the formula range.");

/// <summary>
/// Per-workbook counters returned by a format-specific rewriter.
/// </summary>
internal sealed record ConversionStatistics(
    int RemovedFormulaCount,
    int ProcessedSheetCount,
    int ExcludedSheetCount,
    int RemovedTableFormulaDefinitionCount,
    IReadOnlyList<string> Warnings);
