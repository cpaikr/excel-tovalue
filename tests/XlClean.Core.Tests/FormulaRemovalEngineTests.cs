namespace XlClean.Core.Tests;

using System.Globalization;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

[TestClass]
public sealed class FormulaRemovalEngineTests
{
    private static readonly byte[] MacroPayload = [0xDE, 0xAD, 0xBE, 0xEF, 0x01];
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
    public async Task ConvertsOnlyRowsAndWorksheetsInScope()
    {
        string inputPath = Path.Combine(testDirectory, "Book.xlsx");
        string outputDirectory = Path.Combine(testDirectory, "output");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec(
                "Data",
                new CellSpec("A1", "1+1", "2"),
                new CellSpec("A2", "2+2", "4"),
                new CellSpec("A3", null, "7")),
            new SheetSpec(
                "Keep_TF",
                new CellSpec("A2", "3+3", "6")));
        byte[] sourceBefore = File.ReadAllBytes(inputPath);

        FormulaRemovalRequest request = new(
            [inputPath],
            outputDirectory,
            startRow: 2,
            excludedSheetNamePatterns: ["tf"]);
        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(request);

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(1, result.RemovedFormulaCount);
        Assert.AreEqual(1, result.ProcessedSheetCount);
        Assert.AreEqual(1, result.ExcludedSheetCount);
        Assert.IsNotNull(result.OutputPath);
        CollectionAssert.AreEqual(sourceBefore, File.ReadAllBytes(inputPath));

        CellState rowOne = ReadCell(result.OutputPath, "Data", "A1");
        CellState rowTwo = ReadCell(result.OutputPath, "Data", "A2");
        CellState excluded = ReadCell(result.OutputPath, "Keep_TF", "A2");
        Assert.AreEqual("1+1", rowOne.Formula);
        Assert.IsNull(rowTwo.Formula);
        Assert.AreEqual("4", rowTwo.Value);
        Assert.AreEqual("3+3", excluded.Formula);
        AssertPackageIsValid(result.OutputPath);
    }

    [TestMethod]
    public async Task MissingCachedValueFailsWithoutPublishingOutput()
    {
        string inputPath = Path.Combine(testDirectory, "MissingCache.xlsx");
        string outputDirectory = Path.Combine(testDirectory, "output");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", null)));

        FormulaRemovalRequest request = new([inputPath], outputDirectory);
        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(request);

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsFalse(result.Succeeded);
        Assert.AreEqual(WorkbookFailureKind.MissingCachedValue, result.FailureKind);
        StringAssert.Contains(result.ErrorMessage, "Data");
        StringAssert.Contains(result.ErrorMessage, "A1");
        Assert.IsFalse(Directory.EnumerateFiles(outputDirectory).Any());
    }

    [TestMethod]
    public async Task OneFailureDoesNotAbortRemainingWorkbooks()
    {
        string badInput = Path.Combine(testDirectory, "Bad.xlsx");
        string goodInput = Path.Combine(testDirectory, "Good.xlsx");
        string outputDirectory = Path.Combine(testDirectory, "output");
        CreateWorkbook(
            badInput,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", null)));
        CreateWorkbook(
            goodInput,
            false,
            new SheetSpec("Data", new CellSpec("A1", "2+2", "4")));

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([badInput, goodInput], outputDirectory));

        Assert.AreEqual(1, batch.FailedCount);
        Assert.AreEqual(1, batch.SucceededCount);
        Assert.IsFalse(batch.Files[0].Succeeded);
        Assert.IsTrue(batch.Files[1].Succeeded, batch.Files[1].ErrorMessage);
    }

    [TestMethod]
    public async Task NeverOverwritesQueuedInputsOrEarlierOutputs()
    {
        string firstInput = Path.Combine(testDirectory, "Report.xlsx");
        string secondInput = Path.Combine(testDirectory, "Report_수식제거.xlsx");
        CreateWorkbook(
            firstInput,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));
        CreateWorkbook(
            secondInput,
            false,
            new SheetSpec("Data", new CellSpec("A1", "2+2", "4")));
        byte[] firstBefore = File.ReadAllBytes(firstInput);
        byte[] secondBefore = File.ReadAllBytes(secondInput);

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([firstInput, secondInput], testDirectory));

        Assert.AreEqual(2, batch.SucceededCount);
        Assert.AreEqual(
            "Report_수식제거 (2).xlsx",
            Path.GetFileName(batch.Files[0].OutputPath));
        Assert.AreEqual(
            "Report_수식제거_수식제거.xlsx",
            Path.GetFileName(batch.Files[1].OutputPath));
        CollectionAssert.AreEqual(firstBefore, File.ReadAllBytes(firstInput));
        CollectionAssert.AreEqual(secondBefore, File.ReadAllBytes(secondInput));

        FormulaRemovalBatchResult repeated = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([firstInput], testDirectory));
        Assert.AreEqual(
            "Report_수식제거 (3).xlsx",
            Path.GetFileName(repeated.Files.Single().OutputPath));
    }

    [TestMethod]
    public async Task SkipsDirectoryThatUsesPreferredOutputName()
    {
        string inputPath = Path.Combine(testDirectory, "Report.xlsx");
        string outputDirectory = Path.Combine(testDirectory, "output");
        string collidingDirectory = Path.Combine(outputDirectory, "Report_수식제거.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));
        Directory.CreateDirectory(collidingDirectory);

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], outputDirectory));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual("Report_수식제거 (2).xlsx", Path.GetFileName(result.OutputPath));
        Assert.IsTrue(Directory.Exists(collidingDirectory));
    }

    [TestMethod]
    public async Task RetriesWithNumberedNameWhenOutputAppearsDuringConversion()
    {
        string inputPath = Path.Combine(testDirectory, "Race.xlsx");
        string outputDirectory = Path.Combine(testDirectory, "output");
        string competingPath = Path.Combine(outputDirectory, "Race_수식제거.xlsx");
        byte[] competingBytes = [0x01, 0x02, 0x03];
        bool createdCompetingFile = false;
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));

        InlineProgress<WorkbookProgress> progress = new(value =>
        {
            if (!createdCompetingFile && value.Stage == WorkbookStage.ProcessingWorksheet)
            {
                File.WriteAllBytes(competingPath, competingBytes);
                createdCompetingFile = true;
            }
        });

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], outputDirectory),
            progress);

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual("Race_수식제거 (2).xlsx", Path.GetFileName(result.OutputPath));
        CollectionAssert.AreEqual(competingBytes, File.ReadAllBytes(competingPath));
    }

    [TestMethod]
    public async Task PreservesMacroPayload()
    {
        string inputPath = Path.Combine(testDirectory, "MacroBook.xlsm");
        CreateWorkbook(
            inputPath,
            true,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], Path.Combine(testDirectory, "output")));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.IsNotNull(result.OutputPath);
        CollectionAssert.AreEqual(MacroPayload, ReadMacroPayload(result.OutputPath));
    }

    [TestMethod]
    public async Task PreservesCachedValueTypes()
    {
        string inputPath = Path.Combine(testDirectory, "ValueTypes.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec(
                "Data",
                new CellSpec("A1", "TRUE()", "1", CellValues.Boolean),
                new CellSpec("A2", "NA()", "#N/A", CellValues.Error),
                new CellSpec("A3", "\"text\"", "text", CellValues.String),
                new CellSpec("A4", "5/4", "1.25", CellValues.Number)));

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], Path.Combine(testDirectory, "output")));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.IsNotNull(result.OutputPath);
        Assert.AreEqual(CellValues.Boolean, ReadCell(result.OutputPath, "Data", "A1").DataType);
        Assert.AreEqual(CellValues.Error, ReadCell(result.OutputPath, "Data", "A2").DataType);
        Assert.AreEqual(CellValues.String, ReadCell(result.OutputPath, "Data", "A3").DataType);
        Assert.AreEqual(CellValues.Number, ReadCell(result.OutputPath, "Data", "A4").DataType);
    }

    [TestMethod]
    public async Task RemovesTableFormulaDefinitionsForFullSheetConversion()
    {
        string inputPath = Path.Combine(testDirectory, "TableBook.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec(
                "Data",
                new CellSpec("A1", null, "Header", CellValues.String),
                new CellSpec("A2", "ROW()", "2")));
        AddCalculatedTable(inputPath);

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], Path.Combine(testDirectory, "output")));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(1, result.RemovedTableFormulaDefinitionCount);
        Assert.IsNotNull(result.OutputPath);
        Assert.IsNull(ReadTableFormula(result.OutputPath));
    }

    [TestMethod]
    public async Task RemovesTableDefinitionForPartialConversion()
    {
        string inputPath = Path.Combine(testDirectory, "PartialTable.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec(
                "Data",
                new CellSpec("A1", null, "Header", CellValues.String),
                new CellSpec("A2", "ROW()", "2")));
        AddCalculatedTable(inputPath);

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest(
                [inputPath],
                Path.Combine(testDirectory, "output"),
                startRow: 2));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(1, result.RemovedTableFormulaDefinitionCount);
        Assert.IsEmpty(result.Warnings);
        Assert.IsNotNull(result.OutputPath);
        Assert.IsNull(ReadTableFormula(result.OutputPath));
    }

    [TestMethod]
    public async Task PreservesTableFormulaDefinitionWhollyAboveStartRow()
    {
        string inputPath = Path.Combine(testDirectory, "TableAboveStart.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec(
                "Data",
                new CellSpec("A1", null, "Header", CellValues.String),
                new CellSpec("A2", "ROW()", "2")));
        AddCalculatedTable(inputPath);

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest(
                [inputPath],
                Path.Combine(testDirectory, "output"),
                startRow: 3));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(0, result.RemovedFormulaCount);
        Assert.AreEqual(0, result.RemovedTableFormulaDefinitionCount);
        Assert.IsNotNull(result.OutputPath);
        Assert.AreEqual("ROW()", ReadTableFormula(result.OutputPath));
        Assert.AreEqual("ROW()", ReadCell(result.OutputPath, "Data", "A2").Formula);
    }

    [TestMethod]
    public async Task TableFormulaCrossingStartRowFailsWithoutOutput()
    {
        string inputPath = Path.Combine(testDirectory, "CrossingTable.xlsx");
        string outputDirectory = Path.Combine(testDirectory, "output");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec(
                "Data",
                new CellSpec("A1", null, "Header", CellValues.String),
                new CellSpec("A2", "ROW()", "2"),
                new CellSpec("A3", "ROW()", "3")));
        AddCalculatedTable(inputPath, "A1:A3");

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], outputDirectory, startRow: 3));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.AreEqual(
            WorkbookFailureKind.FormulaRangeCrossesBoundary,
            result.FailureKind);
        Assert.IsFalse(Directory.EnumerateFiles(outputDirectory).Any());
    }

    [TestMethod]
    public async Task HistoricalTotalsFlagDoesNotHideTableBoundaryCrossing()
    {
        string inputPath = Path.Combine(testDirectory, "HistoricalTotals.xlsx");
        string outputDirectory = Path.Combine(testDirectory, "output");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec(
                "Data",
                new CellSpec("A1", null, "Header", CellValues.String),
                new CellSpec("A2", "ROW()", "2"),
                new CellSpec("A3", "ROW()", "3")));
        AddCalculatedTable(inputPath, "A1:A3", totalsRowShown: true);

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], outputDirectory, startRow: 3));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.AreEqual(
            WorkbookFailureKind.FormulaRangeCrossesBoundary,
            result.FailureKind);
        Assert.IsFalse(Directory.EnumerateFiles(outputDirectory).Any());
    }

    [TestMethod]
    public async Task PreservesTotalsRowFunctionWhenTotalsRowIsInScope()
    {
        string inputPath = Path.Combine(testDirectory, "TotalsTable.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec(
                "Data",
                new CellSpec("A1", null, "Header", CellValues.String),
                new CellSpec("A2", "ROW()", "2"),
                new CellSpec("A3", "SUBTOTAL(109,A2:A2)", "2")));
        AddCalculatedTable(
            inputPath,
            "A1:A3",
            totalsRowCount: 1,
            totalsRowFunction: TotalsRowFunctionValues.Sum);

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest(
                [inputPath],
                Path.Combine(testDirectory, "output"),
                startRow: 3));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(1, result.RemovedFormulaCount);
        Assert.AreEqual(0, result.RemovedTableFormulaDefinitionCount);
        Assert.IsNotNull(result.OutputPath);
        Assert.AreEqual(TotalsRowFunctionValues.Sum, ReadTotalsRowFunction(result.OutputPath));
        Assert.AreEqual("ROW()", ReadTableFormula(result.OutputPath));
        Assert.IsNull(ReadCell(result.OutputPath, "Data", "A3").Formula);
    }

    [TestMethod]
    public async Task RemovesCustomTotalsRowFormulaButKeepsItsFunction()
    {
        string inputPath = Path.Combine(testDirectory, "CustomTotals.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec(
                "Data",
                new CellSpec("A1", null, "Header", CellValues.String),
                new CellSpec("A2", "ROW()", "2"),
                new CellSpec("A3", "SUM(A2:A2)*2", "4")));
        AddCalculatedTable(
            inputPath,
            "A1:A3",
            totalsRowCount: 1,
            totalsRowFunction: TotalsRowFunctionValues.Custom,
            totalsRowFormula: "SUM(DataTable[Value])*2");

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], Path.Combine(testDirectory, "output")));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(2, result.RemovedTableFormulaDefinitionCount);
        Assert.IsNotNull(result.OutputPath);
        Assert.AreEqual(TotalsRowFunctionValues.Custom, ReadTotalsRowFunction(result.OutputPath));
        Assert.IsNull(ReadTableFormula(result.OutputPath));
        using SpreadsheetDocument document = SpreadsheetDocument.Open(result.OutputPath, false);
        TableColumn column = document.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single()
            .Table!.Descendants<TableColumn>().Single();
        Assert.IsNull(column.TotalsRowFormula);
    }

    [TestMethod]
    public async Task ShrinksSharedFormulaRangeAcrossStartRow()
    {
        string inputPath = Path.Combine(testDirectory, "Shared.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec(
                "Data",
                new CellSpec("A1", "ROW()", "1"),
                new CellSpec("A2", "ROW()", "2"),
                new CellSpec("A3", "ROW()", "3")));
        SetSharedFormulaRange(inputPath, "A1:A3");

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest(
                [inputPath],
                Path.Combine(testDirectory, "output"),
                startRow: 3));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.IsNotNull(result.OutputPath);
        Assert.AreEqual("A1:A2", ReadFormulaReference(result.OutputPath, "A1"));
        Assert.IsNotNull(ReadCell(result.OutputPath, "Data", "A2").Formula);
        Assert.IsNull(ReadCell(result.OutputPath, "Data", "A3").Formula);
        AssertPackageIsValid(result.OutputPath);
    }

    [TestMethod]
    public async Task ArrayFormulaCrossingStartRowFailsWithoutOutput()
    {
        string inputPath = Path.Combine(testDirectory, "Array.xlsx");
        string outputDirectory = Path.Combine(testDirectory, "output");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec(
                "Data",
                new CellSpec("A1", "ROW(A1:A2)", "1"),
                new CellSpec("A2", null, "2")));
        SetArrayFormulaRange(inputPath, "A1:A2");

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], outputDirectory, startRow: 2));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsFalse(result.Succeeded);
        Assert.AreEqual(
            WorkbookFailureKind.FormulaRangeCrossesBoundary,
            result.FailureKind);
        Assert.IsFalse(Directory.EnumerateFiles(outputDirectory).Any());
    }

    [TestMethod]
    public async Task ReportsManualCalculationModeAndPreservesHeaderFooterText()
    {
        string inputPath = Path.Combine(testDirectory, "Manual.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));
        SetManualCalculationAndHeader(inputPath, "&CQuarterly report");

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], Path.Combine(testDirectory, "output")));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.HasCount(1, result.Warnings);
        Assert.IsNotNull(result.OutputPath);
        Assert.AreEqual("&CQuarterly report", ReadOddHeader(result.OutputPath));
    }

    [TestMethod]
    public async Task InsertsCalculationPropertiesInSchemaOrder()
    {
        string inputPath = Path.Combine(testDirectory, "SchemaOrder.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));
        AddPostCalculationWorkbookChild(inputPath);

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], Path.Combine(testDirectory, "output")));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.IsNotNull(result.OutputPath);
        AssertPackageIsValid(result.OutputPath);
    }

    [TestMethod]
    public async Task PreservesWorksheetMiscellaneousXmlNodes()
    {
        string inputPath = Path.Combine(testDirectory, "MiscXml.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));
        AddWorksheetProcessingInstruction(inputPath);

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], Path.Combine(testDirectory, "output")));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.IsNotNull(result.OutputPath);
        StringAssert.Contains(ReadWorksheetXml(result.OutputPath), "<?preserve marker?>");
    }

    [TestMethod]
    public async Task InvalidPackageReturnsFailure()
    {
        string inputPath = Path.Combine(testDirectory, "Invalid.xlsx");
        await File.WriteAllTextAsync(inputPath, "not an OpenXML package");

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], Path.Combine(testDirectory, "output")));

        Assert.AreEqual(WorkbookFailureKind.InvalidWorkbook, batch.Files.Single().FailureKind);
    }

    [TestMethod]
    public async Task MissingInputReturnsInvalidInput()
    {
        string inputPath = Path.Combine(testDirectory, "Missing.xlsx");

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], Path.Combine(testDirectory, "output")));

        Assert.AreEqual(WorkbookFailureKind.InvalidInput, batch.Files.Single().FailureKind);
    }

    [TestMethod]
    public async Task MissingWorksheetRelationshipReturnsInvalidWorkbook()
    {
        string inputPath = Path.Combine(testDirectory, "BrokenRelationship.xlsx");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(inputPath, true))
        {
            Workbook workbook = document.WorkbookPart!.Workbook ??
                throw new InvalidDataException("Test workbook is missing.");
            workbook.Sheets!.Elements<Sheet>().Single().Id = "rIdMissing";
            workbook.Save();
        }

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest([inputPath], Path.Combine(testDirectory, "output")));

        Assert.AreEqual(WorkbookFailureKind.InvalidWorkbook, batch.Files.Single().FailureKind);
    }

    [TestMethod]
    public void InvalidOutputDirectoryIsRejectedByRequest()
    {
        Assert.ThrowsExactly<ArgumentException>(() =>
            new FormulaRemovalRequest(["Book.xlsx"], "bad\0directory"));
    }

    [TestMethod]
    public async Task RetainsRequestedRelativePathInResultsAndProgress()
    {
        string inputPath = Path.Combine(testDirectory, "Relative.xlsx");
        string requestedPath = Path.GetRelativePath(Directory.GetCurrentDirectory(), inputPath);
        List<WorkbookProgress> reports = [];
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));

        FormulaRemovalBatchResult batch = await FormulaRemovalEngine.ConvertAsync(
            new FormulaRemovalRequest(
                [requestedPath],
                Path.Combine(testDirectory, "output")),
            new InlineProgress<WorkbookProgress>(reports.Add));

        FormulaRemovalFileResult result = batch.Files.Single();
        Assert.IsTrue(result.Succeeded, result.ErrorMessage);
        Assert.AreEqual(requestedPath, result.InputPath);
        Assert.IsNotEmpty(reports);
        Assert.IsTrue(reports.All(report => report.InputPath == requestedPath));
    }

    [TestMethod]
    public async Task PreCanceledRequestPublishesNothing()
    {
        string inputPath = Path.Combine(testDirectory, "Book.xlsx");
        string outputDirectory = Path.Combine(testDirectory, "output");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));
        using CancellationTokenSource cancellation = new();
        cancellation.Cancel();

        bool canceled = false;
        try
        {
            await FormulaRemovalEngine.ConvertAsync(
                new FormulaRemovalRequest([inputPath], outputDirectory),
                cancellationToken: cancellation.Token);
        }
        catch (OperationCanceledException)
        {
            canceled = true;
        }

        Assert.IsTrue(canceled);
        Assert.IsFalse(Directory.Exists(outputDirectory));
    }

    [TestMethod]
    public async Task CancellationDuringWorksheetRemovesOnlyStagingFiles()
    {
        string inputPath = Path.Combine(testDirectory, "Book.xlsx");
        string outputDirectory = Path.Combine(testDirectory, "output");
        CreateWorkbook(
            inputPath,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));
        byte[] sourceBefore = File.ReadAllBytes(inputPath);
        using CancellationTokenSource cancellation = new();
        string? reportedInputPath = null;
        InlineProgress<WorkbookProgress> progress = new(update =>
        {
            if (update.Stage == WorkbookStage.ProcessingWorksheet)
            {
                reportedInputPath = update.InputPath;
                cancellation.Cancel();
            }
        });

        bool canceled = false;
        try
        {
            await FormulaRemovalEngine.ConvertAsync(
                new FormulaRemovalRequest([inputPath], outputDirectory),
                progress,
                cancellation.Token);
        }
        catch (OperationCanceledException)
        {
            canceled = true;
        }

        Assert.IsTrue(canceled);
        Assert.AreEqual(Path.GetFullPath(inputPath), reportedInputPath);
        CollectionAssert.AreEqual(sourceBefore, File.ReadAllBytes(inputPath));
        Assert.IsFalse(Directory.EnumerateFiles(outputDirectory).Any());
    }

    [TestMethod]
    public async Task CancellationRetainsResultsForAlreadyPublishedFiles()
    {
        string firstInput = Path.Combine(testDirectory, "First.xlsx");
        string secondInput = Path.Combine(testDirectory, "Second.xlsx");
        string outputDirectory = Path.Combine(testDirectory, "output");
        CreateWorkbook(
            firstInput,
            false,
            new SheetSpec("Data", new CellSpec("A1", "1+1", "2")));
        CreateWorkbook(
            secondInput,
            false,
            new SheetSpec("Data", new CellSpec("A1", "2+2", "4")));
        using CancellationTokenSource cancellation = new();
        InlineProgress<WorkbookProgress> progress = new(value =>
        {
            if (value.InputPath == secondInput &&
                value.Stage == WorkbookStage.ProcessingWorksheet)
            {
                cancellation.Cancel();
            }
        });

        FormulaRemovalCanceledException? canceled = null;
        try
        {
            await FormulaRemovalEngine.ConvertAsync(
                new FormulaRemovalRequest([firstInput, secondInput], outputDirectory),
                progress,
                cancellation.Token);
        }
        catch (FormulaRemovalCanceledException exception)
        {
            canceled = exception;
        }

        Assert.IsNotNull(canceled);
        Assert.AreEqual(1, canceled.CompletedResult.SucceededCount);
        FormulaRemovalFileResult published = canceled.CompletedResult.Files.Single();
        Assert.AreEqual(firstInput, published.InputPath);
        Assert.IsNotNull(published.OutputPath);
        Assert.IsTrue(File.Exists(published.OutputPath));
        Assert.HasCount(1, Directory.EnumerateFiles(outputDirectory).ToArray());
    }

    private static void CreateWorkbook(
        string path,
        bool macroEnabled,
        params SheetSpec[] sheetSpecs)
    {
        SpreadsheetDocumentType documentType = macroEnabled
            ? SpreadsheetDocumentType.MacroEnabledWorkbook
            : SpreadsheetDocumentType.Workbook;
        using SpreadsheetDocument document = SpreadsheetDocument.Create(path, documentType);
        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

        uint sheetId = 1;
        foreach (SheetSpec sheetSpec in sheetSpecs)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            SheetData sheetData = new();

            foreach (IGrouping<uint, CellSpec> rowCells in sheetSpec.Cells.GroupBy(cell => GetRow(cell.Reference)))
            {
                Row row = new() { RowIndex = rowCells.Key };
                foreach (CellSpec cellSpec in rowCells)
                {
                    Cell cell = new()
                    {
                        CellReference = cellSpec.Reference,
                        DataType = cellSpec.DataType ?? CellValues.Number,
                    };

                    if (cellSpec.Formula is not null)
                    {
                        cell.CellFormula = new CellFormula(cellSpec.Formula);
                    }

                    if (cellSpec.CachedValue is not null)
                    {
                        cell.CellValue = new CellValue(cellSpec.CachedValue);
                    }

                    row.Append(cell);
                }

                sheetData.Append(row);
            }

            worksheetPart.Worksheet = new Worksheet(sheetData);
            worksheetPart.Worksheet.Save();
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId++,
                Name = sheetSpec.Name,
            });
        }

        if (macroEnabled)
        {
            VbaProjectPart vbaProjectPart = workbookPart.AddNewPart<VbaProjectPart>();
            using MemoryStream payload = new(MacroPayload, writable: false);
            vbaProjectPart.FeedData(payload);
        }

        workbookPart.Workbook.Save();
    }

    private static void AddCalculatedTable(
        string workbookPath,
        string reference = "A1:A2",
        uint totalsRowCount = 0,
        TotalsRowFunctionValues? totalsRowFunction = null,
        bool? totalsRowShown = null,
        string? totalsRowFormula = null)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, true);
        WorkbookPart workbookPart = document.WorkbookPart!;
        WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
        TableDefinitionPart tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
        TableColumn column = new()
        {
            Id = 1U,
            Name = "Value",
            CalculatedColumnFormula = new CalculatedColumnFormula("ROW()"),
            TotalsRowFunction = totalsRowFunction,
            TotalsRowFormula = totalsRowFormula is null ? null : new TotalsRowFormula(totalsRowFormula),
        };
        TableColumns columns = new(column) { Count = 1U };
        tablePart.Table = new Table(
            new AutoFilter { Reference = reference },
            columns,
            new TableStyleInfo
            {
                Name = "TableStyleMedium2",
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false,
            })
        {
            Id = 1U,
            Name = "DataTable",
            DisplayName = "DataTable",
            Reference = reference,
            TotalsRowCount = totalsRowCount,
            TotalsRowShown = totalsRowShown ?? totalsRowCount > 0,
        };
        tablePart.Table.Save();

        TableParts tableParts = new(
            new TablePart { Id = worksheetPart.GetIdOfPart(tablePart) })
        {
            Count = 1U,
        };
        worksheetPart.Worksheet!.Append(tableParts);
        worksheetPart.Worksheet.Save();
    }

    private static void AddPostCalculationWorkbookChild(string workbookPath)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, true);
        Workbook workbook = document.WorkbookPart!.Workbook ??
            throw new InvalidDataException("Test workbook is missing.");
        workbook.CalculationProperties?.Remove();
        workbook.Append(new WebPublishing());
        workbook.Save();
    }

    private static void AddWorksheetProcessingInstruction(string workbookPath)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, true);
        Worksheet worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        worksheet.InsertAt(
            new OpenXmlMiscNode(XmlNodeType.ProcessingInstruction, "<?preserve marker?>"),
            0);
        worksheet.Save();
    }

    private static void SetManualCalculationAndHeader(string workbookPath, string headerText)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, true);
        WorkbookPart workbookPart = document.WorkbookPart!;
        Workbook workbook = workbookPart.Workbook ??
            throw new InvalidDataException("Test workbook is missing.");
        workbook.CalculationProperties = new CalculationProperties
        {
            CalculationMode = CalculateModeValues.Manual,
        };
        workbook.Save();

        Worksheet worksheet = workbookPart.WorksheetParts.Single().Worksheet ??
            throw new InvalidDataException("Test worksheet is missing.");
        worksheet.Append(new HeaderFooter(new OddHeader(headerText)));
        worksheet.Save();
    }

    private static void SetSharedFormulaRange(string workbookPath, string reference)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, true);
        Worksheet worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Cell[] cells = worksheet.Descendants<Cell>().ToArray();
        cells[0].CellFormula = new CellFormula("ROW()")
        {
            FormulaType = CellFormulaValues.Shared,
            Reference = reference,
            SharedIndex = 0U,
        };

        foreach (Cell cell in cells.Skip(1))
        {
            cell.CellFormula = new CellFormula
            {
                FormulaType = CellFormulaValues.Shared,
                SharedIndex = 0U,
            };
        }

        worksheet.Save();
    }

    private static void SetArrayFormulaRange(string workbookPath, string reference)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, true);
        Worksheet worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Cell master = worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
        master.CellFormula!.FormulaType = CellFormulaValues.Array;
        master.CellFormula.Reference = reference;
        worksheet.Save();
    }

    private static CellState ReadCell(string workbookPath, string sheetName, string reference)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, false);
        WorkbookPart workbookPart = document.WorkbookPart!;
        Sheet sheet = workbookPart.Workbook!.Sheets!
            .Elements<Sheet>()
            .Single(candidate => candidate.Name?.Value == sheetName);
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        Cell cell = worksheetPart.Worksheet!
            .Descendants<Cell>()
            .Single(candidate => candidate.CellReference?.Value == reference);
        return new CellState(cell.CellFormula?.Text, cell.CellValue?.Text, cell.DataType?.Value);
    }

    private static byte[] ReadMacroPayload(string workbookPath)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, false);
        using Stream stream = document.WorkbookPart!.VbaProjectPart!.GetStream(FileMode.Open, FileAccess.Read);
        using MemoryStream payload = new();
        stream.CopyTo(payload);
        return payload.ToArray();
    }

    private static string? ReadTableFormula(string workbookPath)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, false);
        return document.WorkbookPart!
            .WorksheetParts
            .Single()
            .TableDefinitionParts
            .Single()
            .Table!
            .Descendants<TableColumn>()
            .Single()
            .CalculatedColumnFormula?
            .Text;
    }

    private static TotalsRowFunctionValues? ReadTotalsRowFunction(string workbookPath)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, false);
        return document.WorkbookPart!
            .WorksheetParts
            .Single()
            .TableDefinitionParts
            .Single()
            .Table!
            .Descendants<TableColumn>()
            .Single()
            .TotalsRowFunction?
            .Value;
    }

    private static string? ReadOddHeader(string workbookPath)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, false);
        return document.WorkbookPart!
            .WorksheetParts
            .Single()
            .Worksheet!
            .GetFirstChild<HeaderFooter>()?
            .OddHeader?
            .Text;
    }

    private static string? ReadFormulaReference(string workbookPath, string cellReference)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, false);
        return document.WorkbookPart!
            .WorksheetParts
            .Single()
            .Worksheet!
            .Descendants<Cell>()
            .Single(cell => cell.CellReference?.Value == cellReference)
            .CellFormula?
            .Reference?
            .Value;
    }

    private static string ReadWorksheetXml(string workbookPath)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, false);
        using StreamReader reader = new(
            document.WorkbookPart!.WorksheetParts.Single().GetStream(FileMode.Open, FileAccess.Read));
        return reader.ReadToEnd();
    }

    private static void AssertPackageIsValid(string workbookPath)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, false);
        List<ValidationErrorInfo> errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.AreEqual(
            0,
            errors.Count,
            string.Join(Environment.NewLine, errors.Select(error => error.Description)));
    }

    private static uint GetRow(string cellReference)
    {
        string row = new(cellReference.Where(char.IsAsciiDigit).ToArray());
        return uint.Parse(row, CultureInfo.InvariantCulture);
    }

    private sealed record CellSpec(
        string Reference,
        string? Formula,
        string? CachedValue,
        CellValues? DataType = null);

    private sealed record SheetSpec(string Name, params CellSpec[] Cells);

    private sealed record CellState(string? Formula, string? Value, CellValues? DataType);

    private sealed class InlineProgress<T>(Action<T> report) : IProgress<T>
    {
        public void Report(T value)
        {
            report(value);
        }
    }
}
