namespace XlClean.Core.Tests;

using System.Diagnostics;
using System.Globalization;
using System.IO.Packaging;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static XlClean.Core.Tests.StyleTestWorkbook;

[TestClass]
public sealed class StyleCleanupEngineTests
{
    private string testDirectory = null!;
    private string outputDirectory = null!;

    [TestInitialize]
    public void Initialize()
    {
        testDirectory = Path.Combine(
            Path.GetTempPath(),
            "XlClean.Tests",
            Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        outputDirectory = Path.Combine(testDirectory, "out");
        Directory.CreateDirectory(testDirectory);
    }

    [TestCleanup]
    public void Cleanup()
    {
        if (Directory.Exists(testDirectory))
        {
            Directory.Delete(testDirectory, recursive: true);
        }
    }

    [TestMethod]
    public async Task CleansABloatedWorkbookWithoutChangingTheSource()
    {
        string input = Path.Combine(testDirectory, "Bloated.xlsx");
        CreateBloatedWorkbook(input, junkStyles: 200, duplicateFormats: 30, cells: 40);
        byte[] sourceHash = SHA256.HashData(await File.ReadAllBytesAsync(input));

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        Assert.AreEqual(Path.Combine(outputDirectory, "Bloated_스타일정리.xlsx"), result.OutputPath);
        StyleCleanupChanges changes = result.Changes!;
        Assert.AreEqual(203, changes.CellStylesBefore);
        Assert.AreEqual(3, changes.CellStylesAfter);
        Assert.AreEqual(33, changes.CellFormatsBefore);
        Assert.AreEqual(2, changes.CellFormatsAfter);
        Assert.AreEqual(30, changes.MergedCellFormats);
        Assert.IsEmpty(result.Warnings);
        CollectionAssert.AreEqual(sourceHash, SHA256.HashData(await File.ReadAllBytesAsync(input)));

        string styles = ReadPart(result.OutputPath!, StylesUri);
        Assert.DoesNotContain("Junk", styles);
        StringAssert.Contains(styles, "Head");
        StringAssert.Contains(styles, """<cellXfs count="2">""");
        StringAssert.StartsWith(ReadPart(result.OutputPath!, WorksheetUri), """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        Assert.IsEmpty(SchemaErrors(result.OutputPath!));
    }

    [TestMethod]
    public async Task CleaningTheOutputAgainChangesNothing()
    {
        string input = Path.Combine(testDirectory, "Bloated.xlsx");
        CreateBloatedWorkbook(input, junkStyles: 20, duplicateFormats: 5);
        string first = (await CleanSingle(input)).OutputPath!;

        StyleCleanupFileResult second = await CleanSingle(first);

        Assert.AreEqual(WorkbookFailureKind.None, second.FailureKind, second.ErrorMessage);
        Assert.AreEqual(second.Changes!.CellStylesBefore, second.Changes.CellStylesAfter);
        Assert.AreEqual(second.Changes.CellFormatsBefore, second.Changes.CellFormatsAfter);
        Assert.AreEqual(0, second.Changes.MergedCellFormats);
        Assert.AreEqual(ReadPart(first, StylesUri), ReadPart(second.OutputPath!, StylesUri));
    }

    [TestMethod]
    public async Task CleansAWorkbookWrittenByTheSdk()
    {
        string input = Path.Combine(testDirectory, "Sdk.xlsx");
        CreateSdkWorkbook(input);

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        Assert.AreEqual(52, result.Changes!.CellStylesBefore);
        Assert.AreEqual(2, result.Changes.CellStylesAfter);
        Assert.AreEqual(3, result.Changes.CellFormatsAfter);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(result.OutputPath!, false);
        Stylesheet stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        Assert.AreEqual(3, stylesheet.CellFormats!.Elements<CellFormat>().Count());
        Assert.AreEqual(
            "Normal,Kept",
            string.Join(',', stylesheet.CellStyles!.Elements<CellStyle>().Select(style => style.Name!.Value)));
        Assert.IsEmpty(SchemaErrors(result.OutputPath!));
    }

    [TestMethod]
    public async Task CleanWorkbookIsPublishedUnchanged()
    {
        string input = Path.Combine(testDirectory, "Clean.xlsx");
        Create(
            input,
            Styles(
                """<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>""",
                """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>""",
                """<cellStyle name="Normal" xfId="0" builtinId="0"/>"""),
            """<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>""");

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        Assert.AreEqual(new StyleCleanupChanges(1, 1, 1, 1, 0), result.Changes);
        CollectionAssert.AreEqual(await File.ReadAllBytesAsync(input), await File.ReadAllBytesAsync(result.OutputPath!));
    }

    [TestMethod]
    public async Task UnreadableBinaryWorkbookFailsWithoutOutput()
    {
        string input = Path.Combine(testDirectory, "Book.xlsb");
        await File.WriteAllTextAsync(input, "not a package");

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.InvalidWorkbook, result.FailureKind);
        Assert.IsNull(result.Changes);
        Assert.IsFalse(Directory.Exists(outputDirectory) && Directory.EnumerateFileSystemEntries(outputDirectory).Any());
    }

    [TestMethod]
    public async Task MacroSheetFailsClosedWithoutOutput()
    {
        string input = Path.Combine(testDirectory, "Macro.xlsx");
        Create(
            input,
            Styles(
                """<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>""",
                """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>""",
                """<cellStyle name="Normal" xfId="0" builtinId="0"/>"""),
            "<sheetData/>",
            extraWorkbookPart: (
                "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
                "application/vnd.ms-excel.macrosheet+xml",
                "/xl/macrosheets/sheet1.xml"));

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.UnsupportedWorkbookContent, result.FailureKind);
        Assert.IsEmpty(Directory.GetFiles(outputDirectory));
    }

    [TestMethod]
    public async Task WarnsWhenAMacroWorkbookLosesCustomStyles()
    {
        string input = Path.Combine(testDirectory, "Macro.xlsx");
        CreateBloatedWorkbook(input, junkStyles: 3, duplicateFormats: 0);
        using (Package package = Package.Open(input, FileMode.Open, FileAccess.ReadWrite))
        {
            PackagePart vba = package.CreatePart(
                new Uri("/xl/vbaProject.bin", UriKind.Relative),
                "application/vnd.ms-office.vbaProject");
            using (Stream stream = vba.GetStream(FileMode.Create))
            {
                stream.Write([1, 2, 3]);
            }

            package.GetPart(new Uri("/xl/workbook.xml", UriKind.Relative)).CreateRelationship(
                vba.Uri,
                TargetMode.Internal,
                "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
                "rId8");
        }

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        StringAssert.Contains(result.Warnings.Single(), "3 unused custom cell styles");
    }

    [TestMethod]
    [Timeout(120_000, CooperativeCancellation = true)]
    public async Task LargeStyleTablesFinishQuickly()
    {
        string input = Path.Combine(testDirectory, "Large.xlsx");
        CreateBloatedWorkbook(input, junkStyles: 20_000, duplicateFormats: 20_000, cells: 20_000);
        Stopwatch timer = Stopwatch.StartNew();

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        Assert.AreEqual(2, result.Changes!.CellFormatsAfter);
        Assert.IsLessThan(60, timer.Elapsed.TotalSeconds);
    }

    [TestMethod]
    public async Task RenumberingPreservesEveryOtherWorksheetDetail()
    {
        // Formats 1 and 2 merge, so every index changes; everything else must survive the copy.
        const string sheetBody =
            "<x:cols><x:col min=\"1\" max=\"1\" width=\"9\" style=\"2\" customWidth=\"1\"/></x:cols>" +
            "<x:sheetData><x:row r=\"1\" s=\"2\" customFormat=\"1\" x14ac:dyDescent=\"0.25\">" +
            "<x:c r=\"A1\" s=\"2\" t=\"inlineStr\"><x:is><x:t xml:space=\"preserve\">  two&#13;&#10;lines  </x:t></x:is></x:c>" +
            "<x:c r=\"B1\" s=\"1\" t=\"str\"><x:f><![CDATA[\"a\"&\"b\"]]></x:f><x:v>ab</x:v></x:c>" +
            "<x:c r=\"C1\"><x:v>3</x:v></x:c></x:row></x:sheetData>" +
            "<!-- kept? -->" +
            "<x:dataValidations count=\"1\"><x:dataValidation type=\"whole\" allowBlank=\"1\" showInputMessage=\"1\" prompt=\"Line1&#10;Line2&#9;tab&#13;cr\" sqref=\"A1\"/></x:dataValidations>" +
            "<x:headerFooter><x:oddHeader>&amp;L&#13;&#10;Header</x:oddHeader></x:headerFooter>";
        string input = Path.Combine(testDirectory, "Detailed.xlsx");
        Create(
            input,
            Styles(
                """<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>""",
                """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>""",
                """<cellStyle name="Normal" xfId="0" builtinId="0"/>"""),
            "<sheetData/>");
        ReplacePart(
            input,
            WorksheetUri,
            $"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<x:worksheet xmlns:x=\"{StyleTestWorkbook.Main}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" mc:Ignorable=\"x14ac\">{sheetBody}</x:worksheet>");

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        Assert.AreEqual(1, result.Changes!.MergedCellFormats);
        string worksheet = ReadPart(result.OutputPath!, WorksheetUri);
        StringAssert.StartsWith(worksheet, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        StringAssert.Contains(worksheet, "<x:col min=\"1\" max=\"1\" width=\"9\" style=\"1\" customWidth=\"1\"");
        StringAssert.Contains(worksheet, "<x:row r=\"1\" s=\"1\" customFormat=\"1\" x14ac:dyDescent=\"0.25\">");
        StringAssert.Contains(worksheet, "<![CDATA[\"a\"&\"b\"]]>");
        StringAssert.Contains(worksheet, "<!-- kept? -->");
        StringAssert.Contains(worksheet, "mc:Ignorable=\"x14ac\"");
    }

    [TestMethod]
    public async Task RemovingOnlyStylesLeavesWorksheetsUntouched()
    {
        string input = Path.Combine(testDirectory, "StylesOnly.xlsx");
        CreateBloatedWorkbook(input, junkStyles: 5, duplicateFormats: 0);
        string sourceSheet = ReadPart(input, WorksheetUri);

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        Assert.AreEqual(3, result.Changes!.CellStylesAfter);
        Assert.AreEqual(sourceSheet, ReadPart(result.OutputPath!, WorksheetUri));
    }

    [TestMethod]
    [Timeout(180_000, CooperativeCancellation = true)]
    public async Task WarnsWhenTheCleanedWorkbookStillExceedsExcelsFormatLimit()
    {
        // 65,500 distinct used formats cannot be merged or removed.
        const int formats = 65_500;
        StringBuilder cellFormats = new("""<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>""");
        StringBuilder sheet = new("<sheetData><row r=\"1\">");
        for (int index = 1; index < formats; index++)
        {
            cellFormats.Append(CultureInfo.InvariantCulture, $"""<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment indent="{index}"/></xf>""");
        }

        sheet.Append("</row>");
        for (int index = 1; index < formats; index++)
        {
            sheet.Append(CultureInfo.InvariantCulture, $"""<row r="{index + 1}"><c r="A{index + 1}" s="{index}"/></row>""");
        }

        sheet.Append("</sheetData>");
        string input = Path.Combine(testDirectory, "Huge.xlsx");
        Create(
            input,
            Styles(
                """<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>""",
                cellFormats.ToString(),
                """<cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Junk" xfId="0"/>"""),
            sheet.ToString());

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        StringAssert.Contains(result.Warnings.Single(), "more than the 65430 Excel supports");
    }

    [TestMethod]
    public async Task OutputNamesNeverOverwrite()
    {
        string input = Path.Combine(testDirectory, "Book.xlsx");
        CreateBloatedWorkbook(input, junkStyles: 2, duplicateFormats: 1);
        Directory.CreateDirectory(outputDirectory);
        await File.WriteAllTextAsync(Path.Combine(outputDirectory, "Book_스타일정리.xlsx"), "existing");

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(Path.Combine(outputDirectory, "Book_스타일정리 (2).xlsx"), result.OutputPath);
        Assert.AreEqual("existing", await File.ReadAllTextAsync(Path.Combine(outputDirectory, "Book_스타일정리.xlsx")));
    }

    /// <summary>
    /// Normal, a table-named "Head", and <paramref name="junkStyles"/> unused custom styles; a
    /// default format, one bold format, and <paramref name="duplicateFormats"/> copies of it that
    /// use an identical second bold font.
    /// </summary>
    private static void CreateBloatedWorkbook(string path, int junkStyles, int duplicateFormats, int cells = 4)
    {
        StringBuilder styleFormats = new("""<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>""");
        StringBuilder namedStyles = new("""<cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Good" xfId="0" builtinId="26"/><cellStyle name="Head" xfId="1"/>""");
        for (int index = 0; index < junkStyles; index++)
        {
            styleFormats.Append("""<xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>""");
            namedStyles.Append(CultureInfo.InvariantCulture, $"""<cellStyle name="Junk {index}" xfId="{index + 2}"/>""");
        }

        StringBuilder cellFormats = new("""<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>""");
        for (int index = 0; index < duplicateFormats; index++)
        {
            cellFormats.Append("""<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>""");
        }

        cellFormats.Append("""<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>""");

        StringBuilder sheet = new("<sheetData>");
        for (int row = 1; row <= cells; row++)
        {
            int format = duplicateFormats == 0 ? 1 : 1 + (row % (duplicateFormats + 1));
            sheet.Append(CultureInfo.InvariantCulture, $"""<row r="{row}"><c r="A{row}" s="{format}"><v>{row}</v></c><c r="B{row}" t="s"><v>0</v></c></row>""");
        }

        sheet.Append("</sheetData>");
        Create(
            path,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                Styles(styleFormats.ToString(), cellFormats.ToString(), namedStyles.ToString()),
            sheet.ToString(),
            table: $"""<table xmlns="{StyleTestWorkbook.Main}" id="1" name="T" displayName="T" ref="A1:B2" headerRowCellStyle="Head"><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>""");
        ReplacePart(
            path,
            WorksheetUri,
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" + Worksheet(sheet.ToString(), hasTable: true));
    }

    /// <summary>
    /// A workbook built and saved through the SDK's document model, with 50 unused custom styles
    /// and two used formats that share content.
    /// </summary>
    private static void CreateSdkWorkbook(string path)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        WorkbookPart workbookPart = document.AddWorkbookPart();
        WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        CellStyleFormats styleFormats = new(new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U });
        CellStyles cellStyles = new(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U });
        styleFormats.Append(new CellFormat { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U });
        cellStyles.Append(new CellStyle { Name = "Kept", FormatId = 1U });
        for (uint index = 0; index < 50; index++)
        {
            styleFormats.Append(new CellFormat { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U });
            cellStyles.Append(new CellStyle { Name = $"Copied {index}", FormatId = index + 2 });
        }

        stylesPart.Stylesheet = new Stylesheet(
            new Fonts(new Font(new FontSize { Val = 11D }), new Font(new Bold(), new FontSize { Val = 11D })),
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
            new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())),
            styleFormats,
            new CellFormats(
                new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U },
                new CellFormat { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U, FormatId = 1U, ApplyFont = true },
                new CellFormat { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U, FormatId = 1U, ApplyFont = true },
                new CellFormat { NumberFormatId = 14U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyNumberFormat = true }),
            cellStyles);

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData(
            new Row(
                new Cell { CellReference = "A1", StyleIndex = 1U, CellValue = new CellValue(1) },
                new Cell { CellReference = "B1", StyleIndex = 2U, CellValue = new CellValue(2) },
                new Cell { CellReference = "C1", StyleIndex = 3U, CellValue = new CellValue(45000) })
            { RowIndex = 1U }));
        workbookPart.Workbook = new Workbook(new Sheets(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1U,
            Name = "Data",
        }));
    }

    private static string ReadPart(string path, string uri)
    {
        using Package package = Package.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using StreamReader reader = new(package.GetPart(new Uri(uri, UriKind.Relative)).GetStream(FileMode.Open, FileAccess.Read));
        return reader.ReadToEnd();
    }

    private static string[] SchemaErrors(string path)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(path, false);
        OpenXmlValidator validator = new(FileFormatVersions.Microsoft365);
        return [.. validator.Validate(document.WorkbookPart!.WorkbookStylesPart!).Select(error => error.Description ?? error.Id)];
    }

    private async Task<StyleCleanupFileResult> CleanSingle(string input) =>
        (await StyleCleanupEngine.CleanAsync(new StyleCleanupRequest([input], outputDirectory))).Files.Single();
}
