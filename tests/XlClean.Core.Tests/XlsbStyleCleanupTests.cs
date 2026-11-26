namespace XlClean.Core.Tests;

using System.Buffers.Binary;
using System.Globalization;
using System.IO.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlClean.Core.Styles;

/// <summary>
/// .xlsb style cleanup against workbooks saved by Excel, and edits of them that exercise merging,
/// verification, and the fail-closed rules.
/// </summary>
[TestClass]
public sealed class XlsbStyleCleanupTests
{
    private const string StylesUri = "/xl/styles.bin";
    private const string WorksheetUri = "/xl/worksheets/sheet1.bin";

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
    public async Task RemovesUnusedCustomStylesFromAnExcelWorkbook()
    {
        string input = Fixture("ExcelStyles.xlsb");

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        Assert.AreEqual(new StyleCleanupChanges(8, 3, 11, 11, 0), result.Changes);
        string[] names = StyleNames(result.OutputPath!);
        Assert.AreEqual("Good,Normal,Used custom", string.Join(',', names.Order(StringComparer.Ordinal)));

        // Cell indices did not move, so the worksheet and its binary index are untouched.
        CollectionAssert.AreEqual(PartBytes(input, WorksheetUri), PartBytes(result.OutputPath!, WorksheetUri));
    }

    [TestMethod]
    public async Task CleaningTheOutputAgainChangesNothing()
    {
        string first = (await CleanSingle(Fixture("ExcelStyles.xlsb"))).OutputPath!;

        StyleCleanupFileResult second = await CleanSingle(first);

        Assert.AreEqual(WorkbookFailureKind.None, second.FailureKind, second.ErrorMessage);
        Assert.AreEqual(new StyleCleanupChanges(3, 3, 11, 11, 0), second.Changes);
        CollectionAssert.AreEqual(PartBytes(first, StylesUri), PartBytes(second.OutputPath!, StylesUri));
    }

    [TestMethod]
    public async Task KeepsStylesThatTablesName()
    {
        // Point every cell at the default format so only the table records still name the styles.
        string input = Fixture("ExcelTableStyles.xlsb");
        EditWorksheet(input, records => records.Select(record =>
            XlsbStyleRecords.CellRecords.Contains(record.Type) ? WithCellStyle(record, 0) : record));

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        string[] names = StyleNames(result.OutputPath!);
        CollectionAssert.Contains(names, "Head style");
        CollectionAssert.Contains(names, "Body style");
    }

    [TestMethod]
    public async Task MergesADuplicateFormatAndRenumbersCells()
    {
        // Append a copy of the default format and point column B's default cells (B5, B6) at it,
        // while column A's keep the original, so both copies are in use and must merge.
        string input = Fixture("ExcelStyles.xlsb");
        int duplicate = AppendCellFormatCopy(input, sourceIndex: 0);
        EditWorksheet(input, records => records.Select(record =>
            IsCell(record, column: 1, style: 0) ? WithCellStyle(record, duplicate) : record));

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        Assert.AreEqual(1, result.Changes!.MergedCellFormats);
        Assert.AreEqual(11, result.Changes.CellFormatsAfter);
        Assert.IsFalse(Records(result.OutputPath!, WorksheetUri).Any(record => IsCell(record, column: 1, style: duplicate)));
        Assert.AreEqual(2, Records(result.OutputPath!, WorksheetUri).Count(record => IsCell(record, column: 1, style: 0)));
        Assert.IsFalse(PartExists(result.OutputPath!, "/xl/worksheets/binaryIndex1.bin"));
    }

    [TestMethod]
    public async Task VerifierRejectsACellThatNowLooksDifferent()
    {
        string input = Fixture("ExcelStyles.xlsb");
        string output = (await CleanSingle(input)).OutputPath!;
        string tampered = Path.Combine(testDirectory, "Tampered.xlsb");
        File.Copy(output, tampered);

        // B1 (bold, format 2) now uses format 3 (yellow fill).
        EditWorksheet(tampered, records => records.Select(record =>
            IsCell(record, column: 1, style: 2) ? WithCellStyle(record, 3) : record));

        StyleVerificationException exception = Assert.ThrowsExactly<StyleVerificationException>(
            () => XlsbStyleVerifier.Verify(input, tampered, CancellationToken.None));
        StringAssert.Contains(exception.Message, "looks different");
    }

    [TestMethod]
    public async Task VerifierRejectsARemovedBuiltInStyle()
    {
        string input = Fixture("ExcelStyles.xlsb");
        string output = (await CleanSingle(input)).OutputPath!;
        string tampered = Path.Combine(testDirectory, "Tampered.xlsb");
        File.Copy(output, tampered);
        EditPart(tampered, StylesUri, records => RemoveStyle(records, "Good"));

        Assert.ThrowsExactly<StyleVerificationException>(
            () => XlsbStyleVerifier.Verify(input, tampered, CancellationToken.None));
    }

    [TestMethod]
    public async Task VerifierRejectsARemovedTableStyle()
    {
        string input = Fixture("ExcelTableStyles.xlsb");
        string output = (await CleanSingle(input)).OutputPath!;
        string tampered = Path.Combine(testDirectory, "Tampered.xlsb");
        File.Copy(output, tampered);
        EditPart(tampered, StylesUri, records => RemoveStyle(records, "Head style"));

        StyleVerificationException exception = Assert.ThrowsExactly<StyleVerificationException>(
            () => XlsbStyleVerifier.Verify(input, tampered, CancellationToken.None));
        StringAssert.Contains(exception.Message, "Head style");
    }

    [TestMethod]
    public async Task VerifierRejectsAStaleSectionCount()
    {
        string input = Fixture("ExcelStyles.xlsb");
        string output = (await CleanSingle(input)).OutputPath!;
        string tampered = Path.Combine(testDirectory, "Tampered.xlsb");
        File.Copy(output, tampered);
        EditPart(tampered, StylesUri, records => records.Select(record =>
        {
            if (record.Type != XlsbStyleRecords.BeginStyles)
            {
                return record;
            }

            byte[] data = (byte[])record.Data.Clone();
            BinaryPrimitives.WriteUInt32LittleEndian(data, 99);
            return new Biff12Record(record.Type, data);
        }));

        StyleVerificationException exception = Assert.ThrowsExactly<StyleVerificationException>(
            () => XlsbStyleVerifier.Verify(input, tampered, CancellationToken.None));
        StringAssert.Contains(exception.Message, "declares 99");
    }

    [TestMethod]
    public async Task RenumbersRowAndColumnFormats()
    {
        // Copies of the row 3 format (8) and the column J format (9), and the row and column moved
        // onto them, so the originals become unused and the copies are renumbered.
        string input = Fixture("ExcelStyles.xlsb");
        int rowCopy = AppendCellFormatCopy(input, sourceIndex: 8);
        int columnCopy = AppendCellFormatCopy(input, sourceIndex: 9);
        EditWorksheet(input, records => records.Select(record => record.Type switch
        {
            XlsbStyleRecords.RowHeader when XlsbStyleRecords.RowStyleApplies(record.Data) =>
                WithUInt32(record, XlsbStyleRecords.RowStyleOffset, rowCopy),
            XlsbStyleRecords.ColumnInfo => WithUInt32(record, XlsbStyleRecords.ColumnStyleOffset, columnCopy),
            _ => record,
        }));

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
        Assert.AreEqual(11, result.Changes!.CellFormatsAfter);
        List<Biff12Record> sheet = Records(result.OutputPath!, WorksheetUri);
        uint row = sheet.Single(record => record.Type == XlsbStyleRecords.RowHeader && XlsbStyleRecords.RowStyleApplies(record.Data))
            .ReadUInt32(XlsbStyleRecords.RowStyleOffset);
        uint column = sheet.Single(record => record.Type == XlsbStyleRecords.ColumnInfo).ReadUInt32(XlsbStyleRecords.ColumnStyleOffset);
        Assert.IsLessThan(11U, row);
        Assert.IsLessThan(11U, column);
    }

    [TestMethod]
    public async Task VerifierRejectsAKeptIndexRelationshipWithoutItsPart()
    {
        string input = Fixture("ExcelStyles.xlsb");
        int duplicate = AppendCellFormatCopy(input, sourceIndex: 0);
        EditWorksheet(input, records => records.Select(record =>
            IsCell(record, column: 1, style: 0) ? WithCellStyle(record, duplicate) : record));
        string output = (await CleanSingle(input)).OutputPath!;
        string tampered = Path.Combine(testDirectory, "Tampered.xlsb");
        File.Copy(output, tampered);
        using (Package package = Package.Open(tampered, FileMode.Open, FileAccess.ReadWrite))
        {
            package.GetPart(new Uri(WorksheetUri, UriKind.Relative)).CreateRelationship(
                new Uri("binaryIndex1.bin", UriKind.Relative),
                TargetMode.Internal,
                XlsbStyleCleanup.BinaryIndexRelationship,
                "rIdIndex");
        }

        Assert.ThrowsExactly<StyleVerificationException>(
            () => XlsbStyleVerifier.Verify(input, tampered, CancellationToken.None));
    }

    [TestMethod]
    public async Task UnknownCellDataRecordFailsClosed()
    {
        // A BrtShortBlank (12), whose layout this tool does not rely on, inside sheet data.
        string input = Fixture("ExcelStyles.xlsb");
        EditWorksheet(input, records => records.SelectMany(record =>
            record.Type == 1 ? [record, new Biff12Record(12, [0, 0, 0, 0, 0, 0])] : new[] { record }));

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.UnsupportedWorkbookContent, result.FailureKind);
        StringAssert.Contains(result.ErrorMessage, "record 12");
    }

    [TestMethod]
    public async Task UnappliedRowStyleIndexDoesNotCountAsUse()
    {
        // Row 1 has no row format (fGhostDirty clear); a stale index there is not a reference,
        // and with nothing to renumber the worksheet is left as it is.
        string input = Fixture("ExcelStyles.xlsb");
        EditWorksheet(input, records => records.Select(record =>
        {
            if (record.Type != 0 || BinaryPrimitives.ReadUInt32LittleEndian(record.Data) != 0)
            {
                return record;
            }

            byte[] data = (byte[])record.Data.Clone();
            BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(4), 999_999);
            return new Biff12Record(record.Type, data);
        }));

        StyleCleanupFileResult result = await CleanSingle(input);

        Assert.AreEqual(WorkbookFailureKind.None, result.FailureKind, result.ErrorMessage);
    }

    private static Biff12Record WithUInt32(Biff12Record record, int offset, int value)
    {
        byte[] data = (byte[])record.Data.Clone();
        BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(offset), (uint)value);
        return new Biff12Record(record.Type, data);
    }

    private static bool IsCell(Biff12Record record, int column, int style) =>
        XlsbStyleRecords.CellRecords.Contains(record.Type) &&
        BinaryPrimitives.ReadInt32LittleEndian(record.Data) == column &&
        XlsbStyleRecords.ReadCellStyle(record.Data) == style;

    private static Biff12Record WithCellStyle(Biff12Record record, int style)
    {
        byte[] data = (byte[])record.Data.Clone();
        XlsbStyleRecords.WriteCellStyle(data, style);
        return new Biff12Record(record.Type, data);
    }

    private static int AppendCellFormatCopy(string path, int sourceIndex)
    {
        int added = -1;
        EditPart(path, StylesUri, records =>
        {
            List<Biff12Record> result = [];
            List<Biff12Record> formats = [];
            bool inside = false;
            foreach (Biff12Record record in records)
            {
                if (record.Type == XlsbStyleRecords.BeginCellFormats)
                {
                    inside = true;
                    byte[] count = (byte[])record.Data.Clone();
                    BinaryPrimitives.WriteUInt32LittleEndian(count, BinaryPrimitives.ReadUInt32LittleEndian(count) + 1);
                    result.Add(new Biff12Record(record.Type, count));
                    continue;
                }

                if (inside && record.Type == XlsbStyleRecords.EndCellFormats)
                {
                    inside = false;
                    added = formats.Count;
                    result.Add(new Biff12Record(XlsbStyleRecords.Format, (byte[])formats[sourceIndex].Data.Clone()));
                }
                else if (inside && record.Type == XlsbStyleRecords.Format)
                {
                    formats.Add(record);
                }

                result.Add(record);
            }

            return result;
        });
        return added;
    }

    private static List<Biff12Record> RemoveStyle(IEnumerable<Biff12Record> records, string name)
    {
        List<Biff12Record> list = [.. records];
        int style = list.FindIndex(record => record.Type == XlsbStyleRecords.Style && record.ReadWideString(8, out _) == name);
        Assert.IsGreaterThanOrEqualTo(3, style);
        list.RemoveRange(style - 3, 4);
        int begin = list.FindIndex(record => record.Type == XlsbStyleRecords.BeginStyles);
        byte[] count = (byte[])list[begin].Data.Clone();
        BinaryPrimitives.WriteUInt32LittleEndian(count, BinaryPrimitives.ReadUInt32LittleEndian(count) - 1);
        list[begin] = new Biff12Record(list[begin].Type, count);
        return list;
    }

    private static void EditWorksheet(string path, Func<IEnumerable<Biff12Record>, IEnumerable<Biff12Record>> edit) =>
        EditPart(path, WorksheetUri, edit);

    private static void EditPart(string path, string uri, Func<IEnumerable<Biff12Record>, IEnumerable<Biff12Record>> edit)
    {
        List<Biff12Record> edited = [.. edit(Records(path, uri))];
        using Package package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);
        using Stream stream = package.GetPart(new Uri(uri, UriKind.Relative)).GetStream(FileMode.Create, FileAccess.Write);
        foreach (Biff12Record record in edited)
        {
            Biff12Writer.Write(stream, record);
        }
    }

    private static List<Biff12Record> Records(string path, string uri)
    {
        using Package package = Package.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return XlsbPackage.ReadAllRecords(package.GetPart(new Uri(uri, UriKind.Relative)), CancellationToken.None);
    }

    private static string[] StyleNames(string path) =>
        [.. Records(path, StylesUri)
            .Where(record => record.Type == XlsbStyleRecords.Style)
            .Select(record => record.ReadWideString(8, out _)!)];

    private static byte[] PartBytes(string path, string uri)
    {
        using Package package = Package.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using Stream stream = package.GetPart(new Uri(uri, UriKind.Relative)).GetStream(FileMode.Open, FileAccess.Read);
        using MemoryStream memory = new();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static bool PartExists(string path, string uri)
    {
        using Package package = Package.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return package.PartExists(new Uri(uri, UriKind.Relative));
    }

    private string Fixture(string name)
    {
        string copy = Path.Combine(testDirectory, name);
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", name), copy);
        return copy;
    }

    private async Task<StyleCleanupFileResult> CleanSingle(string input) =>
        (await StyleCleanupEngine.CleanAsync(new StyleCleanupRequest([input], outputDirectory))).Files.Single();
}
