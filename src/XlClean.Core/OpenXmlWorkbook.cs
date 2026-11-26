namespace XlClean.Core;

using DocumentFormat.OpenXml.Packaging;

/// <summary>
/// Opening rules shared by every operation that reads or rewrites an OpenXML workbook.
/// </summary>
internal static class OpenXmlWorkbook
{
    // The SDK's default part-size cap rejects large but legitimate worksheets.
    private const long MaximumCharactersInPart = 1_073_741_824;

    public static OpenSettings CreateOpenSettings() =>
        new()
        {
            AutoSave = false,
            MaxCharactersInPart = MaximumCharactersInPart,
        };

    /// <summary>Reopens a transformed package read-only and requires a readable sheets collection.</summary>
    public static void VerifyCanReopen(string workbookPath)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(
            workbookPath,
            false,
            CreateOpenSettings());

        if (document.WorkbookPart?.Workbook?.Sheets is null)
        {
            throw new InvalidDataException("The converted package has no readable sheets collection.");
        }
    }
}
