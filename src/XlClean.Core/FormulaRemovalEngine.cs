namespace XlClean.Core;

using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

/// <summary>
/// Converts worksheet cell formulas to their last saved cached values.
/// </summary>
/// <remarks>
/// The safety guarantees (no modified input, no overwritten output, verified atomic publication,
/// no partial output on failure) belong to <see cref="WorkbookBatchPipeline"/>. This module
/// supplies the formula rewrite for each format.
/// </remarks>
public static class FormulaRemovalEngine
{
    private const int CopyBufferSize = 1024 * 1024;

    /// <summary>
    /// Converts all requested workbooks sequentially on a worker thread.
    /// </summary>
    /// <param name="request">Immutable conversion inputs and options.</param>
    /// <param name="progress">Optional coarse progress sink.</param>
    /// <param name="cancellationToken">Stops before publication and removes temporary files.</param>
    /// <returns>One explicit result for every requested input path.</returns>
    /// <exception cref="IOException">The output directory cannot be created or resolved.</exception>
    /// <exception cref="UnauthorizedAccessException">The output directory cannot be accessed.</exception>
    /// <exception cref="FormulaRemovalCanceledException">
    /// Cancellation was observed after the batch started. The exception retains completed results.
    /// </exception>
    /// <exception cref="OperationCanceledException">
    /// Cancellation prevented the worker from starting.
    /// </exception>
    public static Task<FormulaRemovalBatchResult> ConvertAsync(
        FormulaRemovalRequest request,
        IProgress<WorkbookProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        return Task.Run(
            () => ConvertBatch(request, progress, cancellationToken),
            cancellationToken);
    }

    private static FormulaRemovalBatchResult ConvertBatch(
        FormulaRemovalRequest request,
        IProgress<WorkbookProgress>? progress,
        CancellationToken cancellationToken)
    {
        try
        {
            return ToBatchResult(WorkbookBatchPipeline.Run(
                request.InputPaths,
                request.OutputDirectory,
                new FormulaRemovalOperation(request),
                progress,
                cancellationToken));
        }
        catch (WorkbookBatchCanceledException<ConversionStatistics> exception)
        {
            throw new FormulaRemovalCanceledException(
                ToBatchResult(exception.Completed),
                exception.CancellationToken,
                exception.InnerException);
        }
    }

    private static FormulaRemovalBatchResult ToBatchResult(
        IReadOnlyList<WorkbookFileOutcome<ConversionStatistics>> outcomes) =>
        new(outcomes.Select(outcome => outcome.Changes is { } statistics
            ? new FormulaRemovalFileResult(
                outcome.InputPath,
                outcome.OutputPath,
                outcome.FailureKind,
                outcome.ErrorMessage,
                statistics.RemovedFormulaCount,
                statistics.ProcessedSheetCount,
                statistics.ExcludedSheetCount,
                statistics.RemovedTableFormulaDefinitionCount,
                Array.AsReadOnly(statistics.Warnings.ToArray()))
            : new FormulaRemovalFileResult(
                outcome.InputPath,
                outcome.OutputPath,
                outcome.FailureKind,
                outcome.ErrorMessage,
                0,
                0,
                0,
                0,
                Array.Empty<string>())));

    private static ConversionStatistics TransformWorkbook(
        WorkbookFile file,
        FormulaRemovalRequest request,
        CancellationToken cancellationToken)
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(
            file.Path,
            true,
            OpenXmlWorkbook.CreateOpenSettings());

        WorkbookPart workbookPart = document.WorkbookPart ??
            throw new InvalidDataException("The package has no workbook part.");
        Workbook workbook = workbookPart.Workbook ??
            throw new InvalidDataException("The package has no workbook root element.");
        Sheets sheets = workbook.Sheets ??
            throw new InvalidDataException("The workbook has no sheets collection.");

        List<string> warnings = [];
        if (workbook.CalculationProperties?.CalculationMode?.Value == CalculateModeValues.Manual)
        {
            warnings.Add(
                "The workbook uses manual calculation; cached formula values may be out of date.");
        }

        int removedFormulas = 0;
        int processedSheets = 0;
        int excludedSheets = 0;
        int removedTableFormulaDefinitions = 0;

        foreach (Sheet sheet in sheets.Elements<Sheet>())
        {
            cancellationToken.ThrowIfCancellationRequested();

            string relationshipId = sheet.Id?.Value ??
                throw new InvalidDataException("A workbook sheet has no relationship id.");
            OpenXmlPart sheetPart;
            try
            {
                sheetPart = workbookPart.GetPartById(relationshipId);
            }
            catch (ArgumentOutOfRangeException exception)
            {
                throw new InvalidDataException(
                    $"Worksheet relationship '{relationshipId}' does not exist.",
                    exception);
            }
            if (sheetPart is not WorksheetPart worksheetPart)
            {
                continue;
            }

            string sheetName = sheet.Name?.Value ??
                throw new InvalidDataException("A worksheet has no name.");
            if (FormulaScope.IsExcluded(sheetName, request.ExcludedSheetNamePatterns))
            {
                excludedSheets++;
                continue;
            }

            file.ReportWorksheet(sheetName);

            removedTableFormulaDefinitions += RewriteTableFormulaDefinitions(
                worksheetPart,
                sheetName,
                request.StartRow);
            removedFormulas += RewriteWorksheet(
                worksheetPart,
                sheetName,
                request.StartRow,
                file.Path,
                cancellationToken);
            processedSheets++;
        }

        if (removedFormulas > 0 || removedTableFormulaDefinitions > 0)
        {
            if (workbookPart.CalculationChainPart is { } calculationChainPart)
            {
                workbookPart.DeletePart(calculationChainPart);
            }

            CalculationProperties? calculationProperties = workbook.CalculationProperties;
            if (calculationProperties is null)
            {
                calculationProperties = new CalculationProperties();
                workbook.AddChild(calculationProperties, throwOnError: true);
            }

            calculationProperties.CalculationId = 0U;
            calculationProperties.CalculationCompleted = false;
            calculationProperties.ForceFullCalculation = true;
            calculationProperties.FullCalculationOnLoad = true;
            workbook.Save();
        }

        return new ConversionStatistics(
            removedFormulas,
            processedSheets,
            excludedSheets,
            removedTableFormulaDefinitions,
            warnings);
    }

    private static int RewriteWorksheet(
        WorksheetPart worksheetPart,
        string sheetName,
        int startRow,
        string workbookPath,
        CancellationToken cancellationToken)
    {
        string stagingPath = Path.Combine(
            Path.GetDirectoryName(workbookPath)!,
            $".xlclean-{Guid.NewGuid():N}.worksheet.tmp");

        int removedFormulas = 0;
        uint currentRow = 0;

        try
        {
            using (OpenXmlReader reader = OpenXmlReader.Create(worksheetPart, readMiscNodes: true))
            using (FileStream stagingStream = new(
                stagingPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                CopyBufferSize,
                FileOptions.SequentialScan))
            using (OpenXmlWriter writer = OpenXmlWriter.Create(stagingStream))
            {
                if (reader.StandaloneXml is bool standalone)
                {
                    writer.WriteStartDocument(standalone);
                }
                else
                {
                    writer.WriteStartDocument();
                }

                bool hasNode = reader.Read();
                while (hasNode)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (reader.IsMiscNode)
                    {
                        OpenXmlElement node = reader.LoadCurrentElement() ??
                            throw new InvalidDataException("A worksheet XML node could not be read.");
                        writer.WriteElement(node);

                        // Loading a miscellaneous node advances this reader to
                        // the following node, unlike loading a regular element.
                        hasNode = !reader.EOF;
                        continue;
                    }

                    if (reader.IsStartElement)
                    {
                        if (reader.ElementType == typeof(Row))
                        {
                            currentRow = GetRowIndex(reader);
                            writer.WriteStartElement(reader);
                        }
                        else if (reader.ElementType == typeof(Cell))
                        {
                            Cell cell = reader.LoadCurrentElement() as Cell ??
                                throw new InvalidDataException("A worksheet cell could not be read.");
                            int rowNumber = GetCellRowNumber(cell.CellReference?.Value, currentRow);

                            if (rowNumber < startRow && cell.CellFormula is not null)
                            {
                                AdjustFormulaRangeAboveBoundary(
                                    cell.CellFormula,
                                    sheetName,
                                    cell.CellReference?.Value ?? $"row {rowNumber}",
                                    startRow);
                            }
                            else if (rowNumber >= startRow && cell.CellFormula is not null)
                            {
                                if (cell.CellValue is null)
                                {
                                    string cellReference = cell.CellReference?.Value ?? $"row {rowNumber}";
                                    throw new MissingCachedValueException(sheetName, cellReference);
                                }

                                cell.CellFormula.Remove();
                                removedFormulas++;
                            }

                            writer.WriteElement(cell);
                        }
                        else if (
                            typeof(OpenXmlLeafTextElement).IsAssignableFrom(reader.ElementType) ||
                            reader.ElementType == typeof(OpenXmlUnknownElement))
                        {
                            OpenXmlElement element = reader.LoadCurrentElement() ??
                                throw new InvalidDataException("A worksheet element could not be read.");
                            writer.WriteElement(element);
                        }
                        else
                        {
                            writer.WriteStartElement(reader);
                        }
                    }
                    else if (reader.IsEndElement)
                    {
                        writer.WriteEndElement();
                        if (reader.ElementType == typeof(Row))
                        {
                            currentRow = 0;
                        }
                    }

                    hasNode = reader.Read();
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            using FileStream transformedWorksheet = new(
                stagingPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                CopyBufferSize,
                FileOptions.SequentialScan);
            worksheetPart.FeedData(transformedWorksheet);

            return removedFormulas;
        }
        finally
        {
            WorkbookBatchPipeline.TryDeleteStagingFile(stagingPath);
        }
    }

    private static void AdjustFormulaRangeAboveBoundary(
        CellFormula formula,
        string sheetName,
        string cellReference,
        int startRow)
    {
        string? reference = formula.Reference?.Value;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return;
        }

        (string firstCell, string lastCell, _, int lastRow) = ParseRange(reference);

        if (lastRow < startRow)
        {
            return;
        }

        CellFormulaValues formulaType = formula.FormulaType?.Value ?? CellFormulaValues.Normal;
        if (formulaType == CellFormulaValues.Shared)
        {
            string shortenedLastCell = ReplaceRowNumber(lastCell, startRow - 1);
            formula.Reference = $"{firstCell}:{shortenedLastCell}";
            return;
        }

        if (formulaType == CellFormulaValues.Array || formulaType == CellFormulaValues.DataTable)
        {
            throw new FormulaRangeBoundaryException(
                "array or data-table formula",
                sheetName,
                cellReference,
                reference,
                startRow);
        }
    }

    private static string ReplaceRowNumber(string cellReference, int rowNumber)
    {
        int digitStart = cellReference.Length;
        while (digitStart > 0 && char.IsAsciiDigit(cellReference[digitStart - 1]))
        {
            digitStart--;
        }

        if (digitStart == cellReference.Length)
        {
            throw new InvalidDataException($"Cell reference '{cellReference}' has no row number.");
        }

        return cellReference[..digitStart] + rowNumber.ToString(CultureInfo.InvariantCulture);
    }

    private static int RewriteTableFormulaDefinitions(
        WorksheetPart worksheetPart,
        string sheetName,
        int startRow)
    {
        int removed = 0;

        foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts)
        {
            Table table = tablePart.Table ??
                throw new InvalidDataException($"Worksheet '{sheetName}' has an unreadable table.");
            string tableName = table.DisplayName?.Value ?? table.Name?.Value ?? "unnamed table";
            string tableReference = table.Reference?.Value ??
                throw new InvalidDataException($"Table '{tableName}' has no cell range.");
            (_, _, int firstRow, int lastRow) = ParseRange(tableReference);

            int headerRowCount = checked((int)(table.HeaderRowCount?.Value ?? 1U));
            int totalsRowCount = checked((int)(table.TotalsRowCount?.Value ?? 0U));
            int dataFirstRow = firstRow + headerRowCount;
            int dataLastRow = lastRow - totalsRowCount;

            int removedFromTable = 0;
            foreach (TableColumn column in table.Descendants<TableColumn>())
            {
                if (column.CalculatedColumnFormula is not null &&
                    FormulaScope.RangeIsInScope(
                        "table calculated-column formula",
                        sheetName,
                        tableName,
                        tableReference,
                        dataFirstRow,
                        Math.Max(dataFirstRow, dataLastRow),
                        startRow))
                {
                    column.CalculatedColumnFormula.Remove();
                    removedFromTable++;
                }

                // The totals-row function is kept: Excel refuses to open a table whose
                // totals cell holds a value while its column has no totals function.
                // A function without a custom formula shows the stored cell value.
                if (column.TotalsRowFormula is not null &&
                    (totalsRowCount == 0
                        ? startRow == 1
                        : FormulaScope.RangeIsInScope(
                            "table totals-row formula",
                            sheetName,
                            tableName,
                            tableReference,
                            lastRow - totalsRowCount + 1,
                            lastRow,
                            startRow)))
                {
                    column.TotalsRowFormula.Remove();
                    removedFromTable++;
                }
            }

            if (removedFromTable > 0)
            {
                table.Save();
                removed += removedFromTable;
            }
        }

        return removed;
    }

    private static (string FirstCell, string LastCell, int FirstRow, int LastRow) ParseRange(
        string reference)
    {
        int separator = reference.IndexOf(':', StringComparison.Ordinal);
        string firstCell = separator >= 0 ? reference[..separator] : reference;
        string lastCell = separator >= 0 ? reference[(separator + 1)..] : reference;
        int firstRow = GetCellRowNumber(firstCell, 0);
        int lastRow = GetCellRowNumber(lastCell, 0);
        if (lastRow < firstRow)
        {
            throw new InvalidDataException($"Formula range '{reference}' is reversed.");
        }

        return (firstCell, lastCell, firstRow, lastRow);
    }

    private static uint GetRowIndex(OpenXmlReader reader)
    {
        OpenXmlAttribute rowIndex = reader.Attributes.FirstOrDefault(attribute =>
            attribute.LocalName == "r" && string.IsNullOrEmpty(attribute.NamespaceUri));

        return uint.TryParse(rowIndex.Value, NumberStyles.None, CultureInfo.InvariantCulture, out uint value)
            ? value
            : 0;
    }

    private static int GetCellRowNumber(string? cellReference, uint currentRow)
    {
        if (!string.IsNullOrWhiteSpace(cellReference))
        {
            ReadOnlySpan<char> reference = cellReference.AsSpan();
            int digitStart = reference.Length;
            while (digitStart > 0 && char.IsAsciiDigit(reference[digitStart - 1]))
            {
                digitStart--;
            }

            if (digitStart < reference.Length &&
                int.TryParse(
                    reference[digitStart..],
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out int rowNumber) &&
                rowNumber is >= 1 and <= FormulaRemovalRequest.MaximumRowNumber)
            {
                return rowNumber;
            }

            throw new InvalidDataException($"Cell reference '{cellReference}' has an invalid row number.");
        }

        if (currentRow is >= 1 and <= FormulaRemovalRequest.MaximumRowNumber)
        {
            return (int)currentRow;
        }

        throw new InvalidDataException("A worksheet cell has no usable row reference.");
    }

    /// <summary>Formula removal as a pipeline operation, dispatching on the workbook format.</summary>
    private sealed class FormulaRemovalOperation(FormulaRemovalRequest request)
        : IWorkbookOperation<ConversionStatistics>
    {
        public string OutputNameSuffix => "_수식제거";

        public ConversionStatistics Transform(WorkbookFile file, CancellationToken cancellationToken) =>
            file.Format == WorkbookFormat.Binary
                ? XlsbFormulaRemover.Transform(file, request, cancellationToken)
                : TransformWorkbook(file, request, cancellationToken);

        // The pipeline's reopen check covers what formula removal needs to prove.
        public void Verify(WorkbookFile file, ConversionStatistics changes, CancellationToken cancellationToken)
        {
        }
    }

    private sealed class MissingCachedValueException(string sheetName, string cellReference)
        : WorkbookFailureException(
            WorkbookFailureKind.MissingCachedValue,
            $"Worksheet '{sheetName}', cell '{cellReference}' has a formula but no cached value. " +
            "Open and recalculate the workbook in Excel, save it, and try again.");
}
