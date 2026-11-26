namespace XlClean.Core.Tests;

using System.Buffers.Binary;
using System.Globalization;
using System.IO.Compression;
using System.IO.Packaging;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

/// <summary>
/// Exercises .xlsb conversion with packages whose record layouts match workbooks saved by Excel.
/// </summary>
[TestClass]
public sealed class XlsbFormulaRemovalTests
{
    private const string WorkbookUri = "/xl/workbook.bin";
    private const string CalcChainUri = "/xl/calcChain.bin";
    private const string RelationshipRoot = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string BinaryIndexRelationship = "http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex";

    private const int RowHeader = 0;
    private const int CellError = 3;
    private const int CellBoolean = 4;
    private const int CellReal = 5;
    private const int CellString = 6;
    private const int FormulaString = 8;
    private const int FormulaNumber = 9;
    private const int FormulaBoolean = 10;
    private const int FormulaError = 11;
    private const int CellMetadata = 49;
    private const int CalcProperties = 157;
    private const int BeginListColumn = 347;
    private const int ListCalculatedColumnFormula = 351;
    private const int ListTotalsFormula = 352;
    private const int ArrayFormula = 426;
    private const int SharedFormula = 427;
    private const int DataTable = 428;

    private string testDirectory = null!;

    [TestInitialize]
    public void Initialize()
    {
        testDirectory = Path.Combine(
            Path.GetTempPath(),
            "XlClean.Tests",
            Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
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
    public async Task ConvertsEveryFormulaResultTypeInScope()
    {
        string inputPath = Path.Combine(testDirectory, "Book.xlsb");
        CreateWorkbook(
            inputPath,
            manualCalculation: false,
            new BinarySheet(
                "Data",
                [
                    Row(0), FormulaNumberCell(0, 2),
                    Row(1), CellMeta(), FormulaNumberCell(1, 4), FormulaStringCell(2, "합계"),
                    FormulaBooleanCell(3, true), FormulaErrorCell(4, 0x07), RealCell(5, 9),
                ]),
            new BinarySheet("Keep_TF", [Row(1), FormulaNumberCell(0, 6)]));
        byte[] sourceBefore = File.ReadAllBytes(inputPath);
        byte[] excludedBefore = ReadPart(inputPath, "/xl/worksheets/sheet2.bin");

        FormulaRemovalFileResult result = await ConvertSingleAsync(
            inputPath,
            startRow: 2,
            excluded: ["tf"]);

        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(4, result.RemovedFormulaCount);
        Assert.AreEqual(1, result.ProcessedSheetCount);
        Assert.AreEqual(1, result.ExcludedSheetCount);
        Assert.IsEmpty(result.Warnings);
        Assert.AreEqual(".xlsb", Path.GetExtension(result.OutputPath));
        CollectionAssert.AreEqual(sourceBefore, File.ReadAllBytes(inputPath));

        List<Record> data = ReadRecords(result.OutputPath!, "/xl/worksheets/sheet1.bin");
        CollectionAssert.AreEqual(
            new[] { 129, 145, RowHeader, FormulaNumber, RowHeader, CellReal, CellString, CellBoolean, CellError, CellReal, 146, 130 },
            data.Select(record => record.Type).ToArray());
        Assert.AreEqual(4d, BinaryPrimitives.ReadDoubleLittleEndian(data[5].Data.AsSpan(8)));
        Assert.HasCount(16, data[5].Data);
        Assert.AreEqual("합계", ReadWideString(data[6].Data, 8));
        Assert.HasCount(8 + 4 + 4, data[6].Data);
        CollectionAssert.AreEqual(new byte[] { 1 }, data[7].Data[8..]);
        CollectionAssert.AreEqual(new byte[] { 0x07 }, data[8].Data[8..]);
        CollectionAssert.AreEqual(Cell(1), data[5].Data[..8]);

        CollectionAssert.AreEqual(excludedBefore, ReadPart(result.OutputPath!, "/xl/worksheets/sheet2.bin"));
        Assert.IsFalse(PartExists(result.OutputPath!, CalcChainUri));
        Assert.IsFalse(PartExists(result.OutputPath!, "/xl/worksheets/binaryIndex1.bin"));
        Assert.IsTrue(PartExists(result.OutputPath!, "/xl/worksheets/binaryIndex2.bin"));

        Record calcProperties = ReadRecords(result.OutputPath!, WorkbookUri)
            .Single(record => record.Type == CalcProperties);
        Assert.AreEqual(0U, BinaryPrimitives.ReadUInt32LittleEndian(calcProperties.Data));
        // Excel's saved 0x006A plus fFullCalcOnLoad, fSomeUncalced, and fNoDeps.
        Assert.AreEqual(0x017B, BinaryPrimitives.ReadUInt16LittleEndian(calcProperties.Data.AsSpan(24)));
    }

    [TestMethod]
    public async Task WorkbookWithoutFormulasKeepsWorksheetAndCalculationChain()
    {
        string inputPath = Path.Combine(testDirectory, "Values.xlsb");
        CreateWorkbook(inputPath, false, new BinarySheet("Data", [Row(0), RealCell(0, 1)]));
        byte[] sheetBefore = ReadPart(inputPath, "/xl/worksheets/sheet1.bin");

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath);

        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(0, result.RemovedFormulaCount);
        CollectionAssert.AreEqual(sheetBefore, ReadPart(result.OutputPath!, "/xl/worksheets/sheet1.bin"));
        Assert.IsTrue(PartExists(result.OutputPath!, CalcChainUri));
        Assert.IsTrue(PartExists(result.OutputPath!, "/xl/worksheets/binaryIndex1.bin"));
    }

    [TestMethod]
    public async Task ShrinksSharedFormulaAcrossStartRowAndDropsItWhenFullyInScope()
    {
        string inputPath = Path.Combine(testDirectory, "Shared.xlsb");
        CreateWorkbook(
            inputPath,
            false,
            new BinarySheet(
                "Data",
                [
                    Row(0), FormulaNumberCell(1, 1), RangeRecord(SharedFormula, 0, 3, 1, 1),
                    Row(1), FormulaNumberCell(1, 2),
                    Row(2), FormulaNumberCell(1, 3), FormulaNumberCell(2, 3), RangeRecord(SharedFormula, 2, 3, 2, 2),
                    Row(3), FormulaNumberCell(1, 4), FormulaNumberCell(2, 4),
                ]));

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath, startRow: 3);

        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(4, result.RemovedFormulaCount);
        List<Record> shared = ReadRecords(result.OutputPath!, "/xl/worksheets/sheet1.bin")
            .Where(record => record.Type == SharedFormula)
            .ToList();
        Assert.HasCount(1, shared);
        Assert.AreEqual(0U, BinaryPrimitives.ReadUInt32LittleEndian(shared[0].Data));
        Assert.AreEqual(1U, BinaryPrimitives.ReadUInt32LittleEndian(shared[0].Data.AsSpan(4)));
    }

    [TestMethod]
    [DataRow(ArrayFormula, "array formula")]
    [DataRow(DataTable, "data-table formula")]
    public async Task MultiCellDefinitionCrossingStartRowFailsWithoutOutput(int recordType, string kind)
    {
        string inputPath = Path.Combine(testDirectory, "Array.xlsb");
        string outputDirectory = Path.Combine(testDirectory, "output");
        CreateWorkbook(
            inputPath,
            false,
            new BinarySheet(
                "Data",
                [
                    Row(1), FormulaNumberCell(2, 10), RangeRecord(recordType, 1, 3, 2, 2),
                    Row(2), FormulaNumberCell(2, 10),
                    Row(3), FormulaNumberCell(2, 10),
                ]));

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath, startRow: 3, outputDirectory);

        Assert.AreEqual(WorkbookFailureKind.FormulaRangeCrossesBoundary, result.FailureKind);
        StringAssert.Contains(result.ErrorMessage, kind);
        StringAssert.Contains(result.ErrorMessage, "C2:C4");
        Assert.IsFalse(Directory.EnumerateFiles(outputDirectory).Any());
    }

    [TestMethod]
    public async Task DropsArrayAndDataTableDefinitionsFullyInScope()
    {
        string inputPath = Path.Combine(testDirectory, "Definitions.xlsb");
        CreateWorkbook(
            inputPath,
            false,
            new BinarySheet(
                "Data",
                [
                    Row(0), CellMeta(), FormulaNumberCell(0, 1), RangeRecord(ArrayFormula, 0, 1, 0, 0),
                    RealCell(3, 1), RangeRecord(DataTable, 0, 1, 4, 4), RealCell(4, 5),
                    Row(1), FormulaNumberCell(0, 2), RealCell(4, 6),
                ]));

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath);

        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        CollectionAssert.AreEqual(
            new[] { 129, 145, RowHeader, CellReal, CellReal, CellReal, RowHeader, CellReal, CellReal, 146, 130 },
            ReadRecords(result.OutputPath!, "/xl/worksheets/sheet1.bin").Select(record => record.Type).ToArray());
    }

    [TestMethod]
    public async Task DataTableAloneInScopeCountsAsRemovedFormulaAndDeletesCalcChain()
    {
        string inputPath = Path.Combine(testDirectory, "DataTableOnly.xlsb");
        CreateWorkbook(
            inputPath,
            false,
            new BinarySheet(
                "Data",
                [
                    Row(0), FormulaNumberCell(1, 2),
                    Row(1), RealCell(0, 1), RangeRecord(DataTable, 1, 2, 1, 1), RealCell(1, 4),
                    Row(2), RealCell(0, 2), RealCell(1, 8),
                ]));

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath, startRow: 2);

        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(1, result.RemovedFormulaCount);
        Assert.IsFalse(ReadRecords(result.OutputPath!, "/xl/worksheets/sheet1.bin")
            .Any(record => record.Type == DataTable));
        Assert.IsFalse(PartExists(result.OutputPath!, CalcChainUri));
    }

    [TestMethod]
    public async Task OversizedStringLengthIsInvalidWorkbook()
    {
        string inputPath = Path.Combine(testDirectory, "Oversized.xlsb");
        Record formula = new(
            FormulaString,
            Concat(Cell(0), UInt32(0x3FFF_FFFF), [0, 0], ParsedFormula()));
        CreateWorkbook(inputPath, false, new BinarySheet("Data", [Row(0), formula]));

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath);

        Assert.AreEqual(WorkbookFailureKind.InvalidWorkbook, result.FailureKind, result.ErrorMessage);
    }

    [TestMethod]
    public async Task MalformedRelationshipTargetIsInvalidWorkbook()
    {
        string inputPath = Path.Combine(testDirectory, "BadTarget.xlsb");
        CreateWorkbook(inputPath, false, new BinarySheet("Data", [Row(0), FormulaNumberCell(0, 1)]));
        using (ZipArchive archive = ZipFile.Open(inputPath, ZipArchiveMode.Update))
        {
            ZipArchiveEntry relationships = archive.GetEntry("xl/_rels/workbook.bin.rels")!;
            string xml;
            using (StreamReader reader = new(relationships.Open()))
            {
                xml = reader.ReadToEnd();
            }

            relationships.Delete();
            using StreamWriter writer = new(archive.CreateEntry("xl/_rels/workbook.bin.rels").Open());
            writer.Write(xml.Replace("worksheets/sheet1.bin", "worksheets/", StringComparison.Ordinal));
        }

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath);

        Assert.AreEqual(WorkbookFailureKind.InvalidWorkbook, result.FailureKind, result.ErrorMessage);
    }

    [TestMethod]
    public async Task RemovesTableFormulaDefinitionsInScope()
    {
        string inputPath = Path.Combine(testDirectory, "Table.xlsb");
        CreateWorkbook(
            inputPath,
            false,
            new BinarySheet(
                "Data",
                [Row(1), FormulaNumberCell(1, 2), Row(3), FormulaNumberCell(1, 9)],
                Table: TableRecords(firstRow: 0, lastRow: 3, totalsRows: 1)));

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath);

        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(2, result.RemovedFormulaCount);
        Assert.AreEqual(2, result.RemovedTableFormulaDefinitionCount);
        List<Record> table = ReadRecords(result.OutputPath!, "/xl/tables/table1.bin");
        Assert.IsFalse(table.Any(record => record.Type is ListCalculatedColumnFormula or ListTotalsFormula));

        // Excel rejects a totals cell value whose column lost its totals function.
        CollectionAssert.AreEqual(
            new uint[] { 9, 6 },
            table
                .Where(record => record.Type == BeginListColumn)
                .Select(record => BinaryPrimitives.ReadUInt32LittleEndian(record.Data.AsSpan(4)))
                .ToArray());
    }

    [TestMethod]
    public async Task PreservesTableDefinitionsAboveStartRowAndRejectsCrossingThem()
    {
        string aboveInput = Path.Combine(testDirectory, "Above.xlsb");
        string crossingInput = Path.Combine(testDirectory, "Crossing.xlsb");
        string outputDirectory = Path.Combine(testDirectory, "output");
        CreateWorkbook(aboveInput, false, new BinarySheet("Data", [], TableRecords(0, 3, 1)));
        CreateWorkbook(crossingInput, false, new BinarySheet("Data", [], TableRecords(0, 3, 1)));
        byte[] tableBefore = ReadPart(aboveInput, "/xl/tables/table1.bin");

        FormulaRemovalFileResult above = await ConvertSingleAsync(aboveInput, startRow: 10, outputDirectory);
        FormulaRemovalFileResult crossing = await ConvertSingleAsync(crossingInput, startRow: 3, outputDirectory);

        Assert.IsTrue(above.Succeeded, above.ErrorMessage);
        CollectionAssert.AreEqual(tableBefore, ReadPart(above.OutputPath!, "/xl/tables/table1.bin"));
        Assert.AreEqual(WorkbookFailureKind.FormulaRangeCrossesBoundary, crossing.FailureKind);
        StringAssert.Contains(crossing.ErrorMessage, "Sales");
        StringAssert.Contains(crossing.ErrorMessage, "A1:B4");
    }

    [TestMethod]
    public async Task ReportsManualCalculationMode()
    {
        string inputPath = Path.Combine(testDirectory, "Manual.xlsb");
        CreateWorkbook(inputPath, true, new BinarySheet("Data", [Row(0), FormulaNumberCell(0, 1)]));

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath);

        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        StringAssert.Contains(result.Warnings.Single(), "manual calculation");
    }

    [TestMethod]
    public async Task OpenXmlPackageWithBinaryExtensionIsInvalidWorkbook()
    {
        string inputPath = Path.Combine(testDirectory, "Renamed.xlsb");
        using (Package package = Package.Open(inputPath, FileMode.CreateNew))
        {
            PackagePart workbook = package.CreatePart(
                new Uri("/xl/workbook.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            package.CreateRelationship(workbook.Uri, TargetMode.Internal, RelationshipRoot + "officeDocument");
        }

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath);

        Assert.AreEqual(WorkbookFailureKind.InvalidWorkbook, result.FailureKind);
    }

    [TestMethod]
    public async Task TruncatedWorksheetRecordIsInvalidWorkbook()
    {
        string inputPath = Path.Combine(testDirectory, "Truncated.xlsb");
        string outputDirectory = Path.Combine(testDirectory, "output");
        byte[] truncated = [FormulaNumber, 37, 0x00, 0x00];
        CreateWorkbook(inputPath, false, new BinarySheet("Data", [], RawSheet: truncated));

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath, outputDirectory: outputDirectory);

        Assert.AreEqual(WorkbookFailureKind.InvalidWorkbook, result.FailureKind);
        Assert.IsFalse(Directory.EnumerateFiles(outputDirectory).Any());
    }

    [TestMethod]
    public async Task OversizedRecordHeaderFailsWithoutReadingPastEndOfStream()
    {
        string inputPath = Path.Combine(testDirectory, "OversizedHeader.xlsb");
        string outputDirectory = Path.Combine(testDirectory, "output");

        // A FormulaNumber header that declares a 200,000,000-byte payload followed by 4 bytes.
        byte[] oversized = [FormulaNumber, 0x80, 0x84, 0xAF, 0x5F, 0x00, 0x00, 0x00, 0x00];
        CreateWorkbook(inputPath, false, new BinarySheet("Data", [], RawSheet: oversized));

        FormulaRemovalFileResult result = await ConvertSingleAsync(inputPath, outputDirectory: outputDirectory);

        Assert.AreEqual(WorkbookFailureKind.InvalidWorkbook, result.FailureKind, result.ErrorMessage);
        Assert.IsFalse(Directory.EnumerateFiles(outputDirectory).Any());
    }

    private async Task<FormulaRemovalFileResult> ConvertSingleAsync(
        string inputPath,
        int startRow = 1,
        string? outputDirectory = null,
        string[]? excluded = null)
    {
        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(new FormulaRemovalRequest(
            [inputPath],
            outputDirectory ?? Path.Combine(testDirectory, "output"),
            startRow,
            excluded));
        return batch.Files.Single();
    }

    private static void CreateWorkbook(string path, bool manualCalculation, params BinarySheet[] sheets)
    {
        using Package package = Package.Open(path, FileMode.CreateNew);
        PackagePart workbook = package.CreatePart(
            new Uri(WorkbookUri, UriKind.Relative),
            "application/vnd.ms-excel.sheet.binary.macroEnabled.main");
        package.CreateRelationship(workbook.Uri, TargetMode.Internal, RelationshipRoot + "officeDocument");

        List<Record> workbookRecords = [new(131, []), new(143, [])];
        for (int index = 0; index < sheets.Length; index++)
        {
            int number = index + 1;
            BinarySheet sheet = sheets[index];
            PackagePart sheetPart = package.CreatePart(
                new Uri($"/xl/worksheets/sheet{number}.bin", UriKind.Relative),
                "application/vnd.ms-excel.worksheet");
            string relationshipId = $"rId{number}";
            workbook.CreateRelationship(
                new Uri($"worksheets/sheet{number}.bin", UriKind.Relative),
                TargetMode.Internal,
                RelationshipRoot + "worksheet",
                relationshipId);
            WritePart(
                sheetPart,
                sheet.RawSheet ?? Serialize([new(129, []), new(145, []), .. sheet.Records, new(146, []), new(130, [])]));

            PackagePart indexPart = package.CreatePart(
                new Uri($"/xl/worksheets/binaryIndex{number}.bin", UriKind.Relative),
                "application/vnd.ms-excel.binIndexWs");
            WritePart(indexPart, Serialize([new(42, new byte[24])]));
            sheetPart.CreateRelationship(
                new Uri($"binaryIndex{number}.bin", UriKind.Relative),
                TargetMode.Internal,
                BinaryIndexRelationship);

            if (sheet.Table is { } table)
            {
                PackagePart tablePart = package.CreatePart(
                    new Uri($"/xl/tables/table{number}.bin", UriKind.Relative),
                    "application/vnd.ms-excel.table");
                WritePart(tablePart, Serialize(table));
                sheetPart.CreateRelationship(
                    new Uri($"../tables/table{number}.bin", UriKind.Relative),
                    TargetMode.Internal,
                    RelationshipRoot + "table");
            }

            workbookRecords.Add(new(156, Concat(
                UInt32(0),
                UInt32((uint)number),
                WideString(relationshipId),
                WideString(sheet.Name))));
        }

        workbookRecords.Add(new(144, []));
        workbookRecords.Add(new(CalcProperties, Concat(
            UInt32(191_029),
            UInt32(manualCalculation ? 0U : 1U),
            UInt32(100),
            BitConverter.GetBytes(0.001),
            UInt32(1),
            [0x6A, 0x00])));
        workbookRecords.Add(new(132, []));
        WritePart(workbook, Serialize(workbookRecords));

        PackagePart calcChain = package.CreatePart(
            new Uri(CalcChainUri, UriKind.Relative),
            "application/vnd.ms-excel.calcChain");
        WritePart(calcChain, Serialize([new(107, new byte[12])]));
        workbook.CreateRelationship(
            new Uri("calcChain.bin", UriKind.Relative),
            TargetMode.Internal,
            RelationshipRoot + "calcChain");
    }

    private static List<Record> TableRecords(uint firstRow, uint lastRow, uint totalsRows)
    {
        byte[] list = Concat(
            Range(firstRow, lastRow, 0, 1),
            UInt32(0),
            UInt32(1),
            UInt32(1),
            UInt32(totalsRows),
            UInt32(1),
            Enumerable.Repeat((byte)0xFF, 24).ToArray(),
            UInt32(0),
            WideString("Sales"),
            WideString("Sales"),
            UInt32(0));
        byte[] Column(uint id, uint totalsFunction) => Concat(
            UInt32(id),
            UInt32(totalsFunction),
            Enumerable.Repeat((byte)0xFF, 12).ToArray(),
            UInt32(0),
            UInt32(uint.MaxValue),
            WideString(id == 1 ? "N" : "D"));
        byte[] formula = Concat([0x00], ParsedFormula());

        return
        [
            new(343, list),
            new(345, UInt32(2)),
            new(BeginListColumn, Column(1, 9)),
            new(ListTotalsFormula, formula),
            new(348, []),
            new(BeginListColumn, Column(2, 6)),
            new(ListCalculatedColumnFormula, formula),
            new(348, []),
            new(346, []),
            new(344, []),
        ];
    }

    private static Record Row(uint zeroBasedRow) => new(RowHeader, Concat(UInt32(zeroBasedRow), new byte[21]));

    private static Record CellMeta() => new(CellMetadata, UInt32(1));

    private static Record RealCell(uint column, double value) =>
        new(CellReal, Concat(Cell(column), BitConverter.GetBytes(value)));

    private static Record FormulaNumberCell(uint column, double value) =>
        new(FormulaNumber, Concat(Cell(column), BitConverter.GetBytes(value), [0, 0], ParsedFormula()));

    private static Record FormulaStringCell(uint column, string value) =>
        new(FormulaString, Concat(Cell(column), WideString(value), [0, 0], ParsedFormula()));

    private static Record FormulaBooleanCell(uint column, bool value) =>
        new(FormulaBoolean, Concat(Cell(column), [value ? (byte)1 : (byte)0], [0, 0], ParsedFormula()));

    private static Record FormulaErrorCell(uint column, byte error) =>
        new(FormulaError, Concat(Cell(column), [error], [0, 0], ParsedFormula()));

    private static Record RangeRecord(int type, uint firstRow, uint lastRow, uint firstColumn, uint lastColumn) =>
        new(type, Concat(Range(firstRow, lastRow, firstColumn, lastColumn), [0x08], ParsedFormula()));

    /// <summary>A CellParsedFormula for <c>=1</c>: cce, PtgInt, and an empty cb.</summary>
    private static byte[] ParsedFormula() => Concat(UInt32(3), [0x1E, 0x01, 0x00], UInt32(0));

    private static byte[] Cell(uint column) => Concat(UInt32(column), [0x00, 0x00, 0x00, 0x00]);

    private static byte[] Range(uint firstRow, uint lastRow, uint firstColumn, uint lastColumn) =>
        Concat(UInt32(firstRow), UInt32(lastRow), UInt32(firstColumn), UInt32(lastColumn));

    private static byte[] UInt32(uint value)
    {
        byte[] bytes = new byte[4];
        BinaryPrimitives.WriteUInt32LittleEndian(bytes, value);
        return bytes;
    }

    private static byte[] WideString(string value) =>
        Concat(UInt32((uint)value.Length), Encoding.Unicode.GetBytes(value));

    private static string ReadWideString(byte[] data, int offset)
    {
        int length = checked((int)BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(offset)));
        return Encoding.Unicode.GetString(data, offset + 4, length * 2);
    }

    private static byte[] Concat(params byte[][] parts) => parts.SelectMany(part => part).ToArray();

    private static byte[] Serialize(IEnumerable<Record> records)
    {
        using MemoryStream stream = new();
        foreach (Record record in records)
        {
            if (record.Type < 0x80)
            {
                stream.WriteByte((byte)record.Type);
            }
            else
            {
                stream.WriteByte((byte)((record.Type & 0x7F) | 0x80));
                stream.WriteByte((byte)(record.Type >> 7));
            }

            int remaining = record.Data.Length;
            do
            {
                byte value = (byte)(remaining & 0x7F);
                remaining >>= 7;
                stream.WriteByte(remaining == 0 ? value : (byte)(value | 0x80));
            }
            while (remaining != 0);

            stream.Write(record.Data);
        }

        return stream.ToArray();
    }

    private static List<Record> ReadRecords(string packagePath, string partUri)
    {
        byte[] bytes = ReadPart(packagePath, partUri);
        List<Record> records = [];
        int position = 0;
        while (position < bytes.Length)
        {
            int type = bytes[position++];
            if ((type & 0x80) != 0)
            {
                type = (type & 0x7F) | ((bytes[position++] & 0x7F) << 7);
            }

            int length = 0;
            for (int shift = 0; ; shift += 7)
            {
                byte value = bytes[position++];
                length |= (value & 0x7F) << shift;
                if ((value & 0x80) == 0)
                {
                    break;
                }
            }

            records.Add(new Record(type, bytes[position..(position + length)]));
            position += length;
        }

        return records;
    }

    private static byte[] ReadPart(string packagePath, string partUri)
    {
        using Package package = Package.Open(packagePath, FileMode.Open, FileAccess.Read);
        using Stream stream = package.GetPart(new Uri(partUri, UriKind.Relative)).GetStream();
        using MemoryStream copy = new();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    private static bool PartExists(string packagePath, string partUri)
    {
        using Package package = Package.Open(packagePath, FileMode.Open, FileAccess.Read);
        return package.PartExists(new Uri(partUri, UriKind.Relative));
    }

    private static void WritePart(PackagePart part, byte[] bytes)
    {
        using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
        stream.Write(bytes);
    }

    private sealed record Record(int Type, byte[] Data);

    private sealed record BinarySheet(
        string Name,
        List<Record> Records,
        List<Record>? Table = null,
        byte[]? RawSheet = null);
}
