namespace ExcelFormulaRemover.Core;

using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Packaging;
using System.Security;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

/// <summary>
/// Converts worksheet cell formulas to their last saved cached values.
/// </summary>
/// <remarks>
/// The module never modifies an input workbook or overwrites an existing
/// output. Each file is transformed in a sibling temporary file and published
/// with a same-volume move only after the package can be reopened.
/// </remarks>
public static class FormulaRemovalEngine
{
    private const long MaximumCharactersInPart = 1_073_741_824;
    private const int CopyBufferSize = 1024 * 1024;
    private static readonly string[] SupportedExtensions = [".xlsx", ".xlsm"];
    private static readonly StringComparer PathComparer =
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparer.OrdinalIgnoreCase
            : StringComparer.Ordinal;

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
        IProgress<FormulaRemovalProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        return Task.Run(
            () => ConvertBatch(request, progress, cancellationToken),
            cancellationToken);
    }

    private static FormulaRemovalBatchResult ConvertBatch(
        FormulaRemovalRequest request,
        IProgress<FormulaRemovalProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        string outputDirectory = request.OutputDirectory;
        Directory.CreateDirectory(outputDirectory);

        HashSet<string> inputPaths = GetNormalizedInputPaths(request.InputPaths);
        HashSet<string> seenInputs = new(PathComparer);
        OutputPathAllocator outputPaths = new(outputDirectory, inputPaths);
        List<FormulaRemovalFileResult> results = new(request.InputPaths.Count);

        try
        {
            for (int index = 0; index < request.InputPaths.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                string inputPath = request.InputPaths[index];
                progress?.Report(new FormulaRemovalProgress(
                    index,
                    request.InputPaths.Count,
                    inputPath,
                    null,
                    FormulaRemovalStage.StartingFile));

                FormulaRemovalFileResult result = ConvertFile(
                    inputPath,
                    request,
                    outputPaths,
                    seenInputs,
                    index,
                    progress,
                    cancellationToken);

                results.Add(result);
                progress?.Report(new FormulaRemovalProgress(
                    index + 1,
                    request.InputPaths.Count,
                    inputPath,
                    null,
                    FormulaRemovalStage.CompletedFile));
            }
        }
        catch (OperationCanceledException exception) when (cancellationToken.IsCancellationRequested)
        {
            throw new FormulaRemovalCanceledException(
                new FormulaRemovalBatchResult(results),
                cancellationToken,
                exception);
        }

        return new FormulaRemovalBatchResult(results);
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "A batch must isolate unexpected failures to the affected workbook and continue safely.")]
    private static FormulaRemovalFileResult ConvertFile(
        string requestedInputPath,
        FormulaRemovalRequest request,
        OutputPathAllocator outputPaths,
        HashSet<string> seenInputs,
        int completedFiles,
        IProgress<FormulaRemovalProgress>? progress,
        CancellationToken cancellationToken)
    {
        string? temporaryPath = null;

        try
        {
            string inputPath = Path.GetFullPath(requestedInputPath);
            ValidateInput(inputPath);

            if (!seenInputs.Add(inputPath))
            {
                return Failure(
                    requestedInputPath,
                    FormulaRemovalFailureKind.InvalidInput,
                    "The same input workbook was requested more than once.");
            }

            string outputPath = outputPaths.Allocate(inputPath);
            temporaryPath = CreateTemporaryPath(outputPath);

            CopyWithCancellation(inputPath, temporaryPath, cancellationToken);
            ConversionStatistics statistics = TransformWorkbook(
                temporaryPath,
                requestedInputPath,
                request,
                completedFiles,
                progress,
                cancellationToken);

            cancellationToken.ThrowIfCancellationRequested();
            VerifyPackageCanReopen(temporaryPath);
            cancellationToken.ThrowIfCancellationRequested();

            outputPath = PublishWithUniqueName(
                temporaryPath,
                outputPath,
                inputPath,
                outputPaths,
                cancellationToken);
            temporaryPath = null;

            return new FormulaRemovalFileResult(
                requestedInputPath,
                outputPath,
                FormulaRemovalFailureKind.None,
                null,
                statistics.RemovedFormulaCount,
                statistics.ProcessedSheetCount,
                statistics.ExcludedSheetCount,
                statistics.RemovedTableFormulaDefinitionCount,
                Array.AsReadOnly(statistics.Warnings.ToArray()));
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (MissingCachedValueException exception)
        {
            return Failure(
                requestedInputPath,
                FormulaRemovalFailureKind.MissingCachedValue,
                exception.Message);
        }
        catch (FormulaRangeBoundaryException exception)
        {
            return Failure(
                requestedInputPath,
                FormulaRemovalFailureKind.FormulaRangeCrossesBoundary,
                exception.Message);
        }
        catch (Exception exception) when (
            exception is ArgumentException or FileNotFoundException or NotSupportedException or
                PathTooLongException or SecurityException)
        {
            return Failure(
                requestedInputPath,
                FormulaRemovalFailureKind.InvalidInput,
                exception.Message);
        }
        catch (Exception exception) when (
            exception is FileFormatException or InvalidDataException or OpenXmlPackageException or
                XmlException or FormatException or InvalidOperationException)
        {
            return Failure(
                requestedInputPath,
                FormulaRemovalFailureKind.InvalidWorkbook,
                exception.Message);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return Failure(
                requestedInputPath,
                FormulaRemovalFailureKind.FileAccess,
                exception.Message);
        }
        catch (Exception exception)
        {
            return Failure(
                requestedInputPath,
                FormulaRemovalFailureKind.Unexpected,
                exception.Message);
        }
        finally
        {
            TryDeleteTemporaryFile(temporaryPath);
        }
    }

    private static ConversionStatistics TransformWorkbook(
        string workbookPath,
        string inputPath,
        FormulaRemovalRequest request,
        int completedFiles,
        IProgress<FormulaRemovalProgress>? progress,
        CancellationToken cancellationToken)
    {
        OpenSettings settings = CreateOpenSettings();
        using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookPath, true, settings);

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
            if (IsExcluded(sheetName, request.ExcludedSheetNamePatterns))
            {
                excludedSheets++;
                continue;
            }

            progress?.Report(new FormulaRemovalProgress(
                completedFiles,
                request.InputPaths.Count,
                inputPath,
                sheetName,
                FormulaRemovalStage.ProcessingWorksheet));

            removedTableFormulaDefinitions += RewriteTableFormulaDefinitions(
                worksheetPart,
                sheetName,
                request.StartRow);
            removedFormulas += RewriteWorksheet(
                worksheetPart,
                sheetName,
                request.StartRow,
                workbookPath,
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
            $".excel-formula-remover-{Guid.NewGuid():N}.worksheet.tmp");

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
            TryDeleteTemporaryFile(stagingPath);
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
                    FormulaRangeIsInScope(
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

                bool hasTotalsFormula = column.TotalsRowFormula is not null;
                bool hasTotalsFunction = column.TotalsRowFunction?.Value is { } totalsFunction &&
                    totalsFunction != TotalsRowFunctionValues.None;
                bool totalsAreInScope = false;
                if (hasTotalsFormula || hasTotalsFunction)
                {
                    totalsAreInScope = totalsRowCount == 0
                        ? startRow == 1
                        : FormulaRangeIsInScope(
                            "table totals-row formula",
                            sheetName,
                            tableName,
                            tableReference,
                            lastRow - totalsRowCount + 1,
                            lastRow,
                            startRow);
                }

                if (column.TotalsRowFormula is not null && totalsAreInScope)
                {
                    column.TotalsRowFormula.Remove();
                    removedFromTable++;
                }

                if (hasTotalsFunction && totalsAreInScope)
                {
                    column.TotalsRowFunction = null;
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

    private static bool FormulaRangeIsInScope(
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

    private static bool IsExcluded(string sheetName, IReadOnlyList<string> patterns)
    {
        return patterns.Any(pattern =>
            sheetName.Contains(pattern, StringComparison.OrdinalIgnoreCase));
    }

    private static void ValidateInput(string inputPath)
    {
        string extension = Path.GetExtension(inputPath);
        if (!SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"Unsupported workbook extension '{extension}'. Use .xlsx or .xlsm.",
                nameof(inputPath));
        }

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("The input workbook does not exist.", inputPath);
        }
    }

    private static HashSet<string> GetNormalizedInputPaths(IReadOnlyList<string> requestedPaths)
    {
        HashSet<string> paths = new(PathComparer);
        foreach (string requestedPath in requestedPaths)
        {
            try
            {
                paths.Add(Path.GetFullPath(requestedPath));
            }
            catch (Exception exception) when (
                exception is ArgumentException or NotSupportedException or PathTooLongException or SecurityException)
            {
                // The per-file conversion returns a precise invalid-input result.
            }
        }

        return paths;
    }

    private static void CopyWithCancellation(
        string inputPath,
        string outputPath,
        CancellationToken cancellationToken)
    {
        using FileStream input = new(
            inputPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            CopyBufferSize,
            FileOptions.SequentialScan);
        using FileStream output = new(
            outputPath,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            CopyBufferSize,
            FileOptions.SequentialScan);

        byte[] buffer = new byte[CopyBufferSize];
        int bytesRead;
        while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            output.Write(buffer, 0, bytesRead);
        }

        output.Flush(flushToDisk: true);
    }

    private static void VerifyPackageCanReopen(string workbookPath)
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

    private static OpenSettings CreateOpenSettings()
    {
        return new OpenSettings
        {
            AutoSave = false,
            MaxCharactersInPart = MaximumCharactersInPart,
        };
    }

    private static string CreateTemporaryPath(string outputPath)
    {
        string directory = Path.GetDirectoryName(outputPath)!;
        string fileName = Path.GetFileNameWithoutExtension(outputPath);
        string extension = Path.GetExtension(outputPath);
        return Path.Combine(
            directory,
            $".excel-formula-remover-{fileName}-{Guid.NewGuid():N}.tmp{extension}");
    }

    private static string PublishWithUniqueName(
        string temporaryPath,
        string initialOutputPath,
        string inputPath,
        OutputPathAllocator outputPaths,
        CancellationToken cancellationToken)
    {
        string outputPath = initialOutputPath;

        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                File.Move(temporaryPath, outputPath, overwrite: false);
                return outputPath;
            }
            catch (IOException) when (Path.Exists(outputPath))
            {
                outputPath = outputPaths.Allocate(inputPath);
            }
        }
    }

    private static FormulaRemovalFileResult Failure(
        string inputPath,
        FormulaRemovalFailureKind failureKind,
        string message)
    {
        return new FormulaRemovalFileResult(
            inputPath,
            null,
            failureKind,
            message,
            0,
            0,
            0,
            0,
            Array.Empty<string>());
    }

    private static void TryDeleteTemporaryFile(string? path)
    {
        if (string.IsNullOrEmpty(path))
        {
            return;
        }

        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
            // The uniquely named staging file is never a user file. A later
            // process or antivirus lock may delay its removal.
        }
        catch (UnauthorizedAccessException)
        {
            // See the IOException rationale above.
        }
    }

    private sealed class OutputPathAllocator(
        string outputDirectory,
        IEnumerable<string> inputPaths)
    {
        private readonly HashSet<string> reservedPaths = new(inputPaths, PathComparer);

        public string Allocate(string inputPath)
        {
            string baseName = $"{Path.GetFileNameWithoutExtension(inputPath)}_수식제거";
            string extension = Path.GetExtension(inputPath);

            for (int sequence = 1; sequence < int.MaxValue; sequence++)
            {
                string suffix = sequence == 1 ? string.Empty : $" ({sequence})";
                string candidate = Path.Combine(outputDirectory, $"{baseName}{suffix}{extension}");

                if (!Path.Exists(candidate) && reservedPaths.Add(candidate))
                {
                    return candidate;
                }
            }

            throw new IOException($"No available output name could be allocated for '{inputPath}'.");
        }
    }

    private sealed class MissingCachedValueException(string sheetName, string cellReference)
        : Exception(
            $"Worksheet '{sheetName}', cell '{cellReference}' has a formula but no cached value. " +
            "Open and recalculate the workbook in Excel, save it, and try again.");

    private sealed class FormulaRangeBoundaryException(
        string formulaKind,
        string sheetName,
        string owner,
        string formulaRange,
        int startRow)
        : Exception(
            $"Worksheet '{sheetName}', '{owner}' owns a {formulaKind} over range '{formulaRange}', " +
            $"which crosses start row {startRow}. Choose row 1 or a row " +
            "after the formula range.");

    private sealed record ConversionStatistics(
        int RemovedFormulaCount,
        int ProcessedSheetCount,
        int ExcludedSheetCount,
        int RemovedTableFormulaDefinitionCount,
        IReadOnlyList<string> Warnings);
}
