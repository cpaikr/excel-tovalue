namespace XlClean.Core;

using System.Buffers.Binary;
using System.Globalization;
using System.IO.Packaging;

/// <summary>
/// Rewrites Excel binary (.xlsb, [MS-XLSB]) worksheets so formula cells keep only their cached values.
/// </summary>
/// <remarks>
/// Records are streamed and copied unchanged except for the formula-bearing records listed in
/// <see cref="RecordType"/>. The rules mirror the OpenXML path: formulas at or after the start row
/// become value cells, multi-cell definitions are shrunk or rejected at the boundary, and table
/// column formulas in scope are removed. A BIFF12 formula record always stores its last result,
/// so binary workbooks cannot have a missing cached value.
/// </remarks>
internal static class XlsbFormulaRemover
{
    private const int CopyBufferSize = 1024 * 1024;
    private const int CancellationCheckInterval = 4096;
    private const int CellLength = 8;
    private const int RowRangeLength = 16;

    private const string CalculationChainRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
    private const string BinaryIndexRelationship =
        "http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex";

    /// <remarks>
    /// The pipeline reports package <see cref="ArgumentException"/>s from this method as an
    /// invalid workbook.
    /// </remarks>
    public static ConversionStatistics Transform(
        WorkbookFile file,
        FormulaRemovalRequest request,
        CancellationToken cancellationToken)
    {
        using Package package = Package.Open(
            file.Path,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.None);

        PackagePart workbookPart = XlsbPackage.GetWorkbookPart(package);
        WorkbookInfo workbook = ReadWorkbook(workbookPart);

        List<string> warnings = [];
        if (workbook.CalculationProperties is { } calculationProperties &&
            calculationProperties.ReadUInt32(CalculationPropertiesLayout.AutoRecalcOffset) == 0)
        {
            warnings.Add(
                "The workbook uses manual calculation; cached formula values may be out of date.");
        }

        int removedFormulas = 0;
        int processedSheets = 0;
        int excludedSheets = 0;
        int removedTableFormulaDefinitions = 0;

        foreach ((string relationshipId, string sheetName) in workbook.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();

            PackagePart? sheetPart = XlsbPackage.GetRelatedPart(
                workbookPart,
                relationshipId,
                XlsbPackage.WorksheetRelationship);
            if (sheetPart is null)
            {
                continue;
            }

            if (FormulaScope.IsExcluded(sheetName, request.ExcludedSheetNamePatterns))
            {
                excludedSheets++;
                continue;
            }

            file.ReportWorksheet(sheetName);

            foreach (PackagePart tablePart in XlsbPackage.GetRelatedParts(sheetPart, XlsbPackage.TableRelationship))
            {
                removedTableFormulaDefinitions += RewriteTable(tablePart, sheetName, request.StartRow);
            }

            removedFormulas += RewriteWorksheet(
                package,
                sheetPart,
                sheetName,
                request.StartRow,
                file.Path,
                cancellationToken);
            processedSheets++;
        }

        if (removedFormulas > 0 || removedTableFormulaDefinitions > 0)
        {
            XlsbPackage.DeleteRelatedParts(package, workbookPart, CalculationChainRelationship);
            RequestFullCalculationOnLoad(workbookPart);
        }

        return new ConversionStatistics(
            removedFormulas,
            processedSheets,
            excludedSheets,
            removedTableFormulaDefinitions,
            warnings);
    }

    private static int RewriteWorksheet(
        Package package,
        PackagePart sheetPart,
        string sheetName,
        int startRow,
        string workbookPath,
        CancellationToken cancellationToken)
    {
        string stagingPath = Path.Combine(
            Path.GetDirectoryName(workbookPath)!,
            $".xlclean-{Guid.NewGuid():N}.worksheet.tmp");

        int removedFormulas = 0;
        bool changed = false;

        try
        {
            using (Stream source = sheetPart.GetStream(FileMode.Open, FileAccess.Read))
            using (FileStream staging = new(
                stagingPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                CopyBufferSize,
                FileOptions.SequentialScan))
            {
                Biff12Reader reader = new(new BufferedStream(source, CopyBufferSize));
                long currentRow = -1;
                Biff12Record? pendingCellMetadata = null;
                long recordCount = 0;

                while (reader.TryRead(out Biff12Record record))
                {
                    if (++recordCount % CancellationCheckInterval == 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                    }

                    switch (record.Type)
                    {
                        case RecordType.RowHeader:
                            currentRow = record.ReadUInt32(0);
                            break;

                        case RecordType.CellMetadata:
                            // Cell metadata describes the following cell (for example a dynamic
                            // array anchor). Hold it until that cell decides whether it survives.
                            FlushPending(staging, ref pendingCellMetadata);
                            pendingCellMetadata = record;
                            continue;

                        case RecordType.FormulaString:
                        case RecordType.FormulaNumber:
                        case RecordType.FormulaBoolean:
                        case RecordType.FormulaError:
                            if (GetOneBasedRow(currentRow) >= startRow)
                            {
                                pendingCellMetadata = null;
                                Biff12Writer.Write(staging, ToValueRecord(record));
                                removedFormulas++;
                                changed = true;
                                continue;
                            }

                            break;

                        case RecordType.ArrayFormula:
                        case RecordType.SharedFormula:
                        case RecordType.DataTable:
                            FlushPending(staging, ref pendingCellMetadata);
                            Biff12Record? retained = AdjustRangeDefinition(record, sheetName, startRow);
                            if (retained is { } kept)
                            {
                                Biff12Writer.Write(staging, kept);
                            }
                            else if (record.Type == RecordType.DataTable)
                            {
                                // Data-table interior cells are already plain values; the
                                // BrtTable record is the only formula, as <f t="dataTable"> in .xlsx.
                                removedFormulas++;
                            }

                            changed |= retained is null || !ReferenceEquals(retained.Value.Data, record.Data);
                            continue;
                    }

                    FlushPending(staging, ref pendingCellMetadata);
                    Biff12Writer.Write(staging, record);
                }

                FlushPending(staging, ref pendingCellMetadata);
                staging.Flush();
            }

            if (!changed)
            {
                return 0;
            }

            cancellationToken.ThrowIfCancellationRequested();
            using (FileStream transformed = new(
                stagingPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                CopyBufferSize,
                FileOptions.SequentialScan))
            using (Stream destination = sheetPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                CopyWithCancellation(transformed, destination, cancellationToken);
            }

            // The binary index stores byte offsets into the original worksheet stream. It is an
            // optional load accelerator, so remove it rather than leave stale offsets.
            XlsbPackage.DeleteRelatedParts(package, sheetPart, BinaryIndexRelationship);
            return removedFormulas;
        }
        finally
        {
            WorkbookBatchPipeline.TryDeleteStagingFile(stagingPath);
        }
    }

    private static void CopyWithCancellation(
        Stream source,
        Stream destination,
        CancellationToken cancellationToken)
    {
        byte[] buffer = new byte[CopyBufferSize];
        int bytesRead;
        while ((bytesRead = source.Read(buffer, 0, buffer.Length)) > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            destination.Write(buffer, 0, bytesRead);
        }
    }

    private static Biff12Record ToValueRecord(Biff12Record formula)
    {
        byte[] data = formula.Data;
        switch (formula.Type)
        {
            case RecordType.FormulaNumber:
                formula.RequireLength(CellLength + sizeof(double) + sizeof(ushort));
                return new Biff12Record(RecordType.CellReal, data[..(CellLength + sizeof(double))]);

            case RecordType.FormulaBoolean:
                formula.RequireLength(CellLength + 1 + sizeof(ushort));
                return new Biff12Record(RecordType.CellBoolean, data[..(CellLength + 1)]);

            case RecordType.FormulaError:
                formula.RequireLength(CellLength + 1 + sizeof(ushort));
                return new Biff12Record(RecordType.CellError, data[..(CellLength + 1)]);

            case RecordType.FormulaString:
                _ = formula.ReadWideString(CellLength, out int valueEnd);
                formula.RequireLength(valueEnd + sizeof(ushort));
                return new Biff12Record(RecordType.CellString, data[..valueEnd]);

            default:
                throw new InvalidOperationException($"Record {formula.Type} is not a formula cell.");
        }
    }

    /// <returns>The record to keep, or <see langword="null"/> when every covered cell was converted.</returns>
    private static Biff12Record? AdjustRangeDefinition(Biff12Record record, string sheetName, int startRow)
    {
        record.RequireLength(RowRangeLength);
        RowRange range = RowRange.Read(record, 0);
        string reference = range.ToA1Reference();

        if (record.Type == RecordType.SharedFormula)
        {
            if (range.LastRow < startRow)
            {
                return record;
            }

            if (range.FirstRow >= startRow)
            {
                return null;
            }

            // The anchor is the first cell, which stays a formula above the boundary.
            byte[] shortened = (byte[])record.Data.Clone();
            BinaryPrimitives.WriteUInt32LittleEndian(shortened.AsSpan(4), checked((uint)(startRow - 2)));
            return new Biff12Record(record.Type, shortened);
        }

        string formulaKind = record.Type == RecordType.ArrayFormula
            ? "array formula"
            : "data-table formula";
        return FormulaScope.RangeIsInScope(
            formulaKind,
            sheetName,
            range.FirstCellReference(),
            reference,
            range.FirstRow,
            range.LastRow,
            startRow)
            ? null
            : record;
    }

    private static int RewriteTable(PackagePart tablePart, string sheetName, int startRow)
    {
        List<Biff12Record> records = XlsbPackage.ReadAllRecords(tablePart, CancellationToken.None);
        int listIndex = records.FindIndex(record => record.Type == RecordType.BeginList);
        if (listIndex < 0)
        {
            throw new InvalidDataException($"Worksheet '{sheetName}' has an unreadable table.");
        }

        Biff12Record list = records[listIndex];
        list.RequireLength(TableLayout.DisplayNameOffset);
        RowRange range = RowRange.Read(list, 0);
        string tableReference = range.ToA1Reference();
        string tableName = ReadTableName(list);
        int headerRowCount = checked((int)list.ReadUInt32(TableLayout.HeaderRowCountOffset));
        int totalsRowCount = checked((int)list.ReadUInt32(TableLayout.TotalsRowCountOffset));
        int dataFirstRow = range.FirstRow + headerRowCount;
        int dataLastRow = range.LastRow - totalsRowCount;

        int removed = 0;
        List<Biff12Record> output = new(records.Count);
        for (int index = 0; index < records.Count; index++)
        {
            Biff12Record record = records[index];
            if (record.Type != RecordType.BeginListColumn)
            {
                output.Add(record);
                continue;
            }

            int end = records.FindIndex(index, item => item.Type == RecordType.EndListColumn);
            if (end < 0)
            {
                throw new InvalidDataException($"Table '{tableName}' has an unterminated column.");
            }

            List<Biff12Record> column = records.GetRange(index, end - index + 1);
            removed += RewriteTableColumn(
                column,
                sheetName,
                tableName,
                tableReference,
                range,
                totalsRowCount,
                dataFirstRow,
                dataLastRow,
                startRow);
            output.AddRange(column);
            index = end;
        }

        if (removed > 0)
        {
            using Stream destination = tablePart.GetStream(FileMode.Create, FileAccess.Write);
            foreach (Biff12Record record in output)
            {
                Biff12Writer.Write(destination, record);
            }
        }

        return removed;
    }

    private static int RewriteTableColumn(
        List<Biff12Record> column,
        string sheetName,
        string tableName,
        string tableReference,
        RowRange tableRange,
        int totalsRowCount,
        int dataFirstRow,
        int dataLastRow,
        int startRow)
    {
        int removed = 0;

        if (column.Exists(record => record.Type == RecordType.ListCalculatedColumnFormula) &&
            FormulaScope.RangeIsInScope(
                "table calculated-column formula",
                sheetName,
                tableName,
                tableReference,
                dataFirstRow,
                Math.Max(dataFirstRow, dataLastRow),
                startRow))
        {
            removed += column.RemoveAll(record => record.Type == RecordType.ListCalculatedColumnFormula);
        }

        // The totals-row function (BrtBeginListCol.ilta) is kept: Excel refuses to open a
        // table whose totals cell holds a value while its column has no totals function.
        if (column.Exists(record => record.Type == RecordType.ListTotalsFormula) &&
            (totalsRowCount == 0
                ? startRow == 1
                : FormulaScope.RangeIsInScope(
                    "table totals-row formula",
                    sheetName,
                    tableName,
                    tableReference,
                    tableRange.LastRow - totalsRowCount + 1,
                    tableRange.LastRow,
                    startRow)))
        {
            removed += column.RemoveAll(record => record.Type == RecordType.ListTotalsFormula);
        }

        return removed;
    }

    private static string ReadTableName(Biff12Record list)
    {
        string? name = list.ReadWideString(TableLayout.NameOffset, out int displayNameOffset);
        string? displayName = list.ReadWideString(displayNameOffset, out _);
        return displayName ?? name ?? "unnamed table";
    }

    private static void RequestFullCalculationOnLoad(PackagePart workbookPart)
    {
        List<Biff12Record> records = XlsbPackage.ReadAllRecords(workbookPart, CancellationToken.None);
        int index = records.FindIndex(record => record.Type == RecordType.CalculationProperties);
        if (index < 0)
        {
            // Excel writes BrtCalcProp in every workbook it saves. Without it the deleted
            // calculation chain still makes Excel rebuild dependencies on load.
            return;
        }

        Biff12Record properties = records[index];
        properties.RequireLength(CalculationPropertiesLayout.FlagsOffset + sizeof(ushort));
        byte[] data = (byte[])properties.Data.Clone();
        BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(CalculationPropertiesLayout.RecalcIdOffset), 0);
        Span<byte> flags = data.AsSpan(CalculationPropertiesLayout.FlagsOffset, sizeof(ushort));
        BinaryPrimitives.WriteUInt16LittleEndian(
            flags,
            (ushort)(BinaryPrimitives.ReadUInt16LittleEndian(flags) | CalculationPropertiesLayout.ForceRecalculationFlags));
        records[index] = new Biff12Record(properties.Type, data);

        using Stream destination = workbookPart.GetStream(FileMode.Create, FileAccess.Write);
        foreach (Biff12Record record in records)
        {
            Biff12Writer.Write(destination, record);
        }
    }

    private static WorkbookInfo ReadWorkbook(PackagePart workbookPart)
    {
        List<Biff12Record> records = XlsbPackage.ReadAllRecords(workbookPart, CancellationToken.None);
        Biff12Record? calculationProperties = null;
        foreach (Biff12Record record in records)
        {
            if (record.Type == RecordType.CalculationProperties)
            {
                record.RequireLength(CalculationPropertiesLayout.FlagsOffset + sizeof(ushort));
                calculationProperties = record;
            }
        }

        return new WorkbookInfo(XlsbPackage.ReadSheets(records), calculationProperties);
    }

    private static int GetOneBasedRow(long zeroBasedRow)
    {
        return zeroBasedRow is >= 0 and < FormulaRemovalRequest.MaximumRowNumber
            ? (int)zeroBasedRow + 1
            : throw new InvalidDataException("A worksheet cell has no usable row reference.");
    }

    private static void FlushPending(Stream destination, ref Biff12Record? pending)
    {
        if (pending is { } record)
        {
            Biff12Writer.Write(destination, record);
            pending = null;
        }
    }

    private sealed record WorkbookInfo(
        IReadOnlyList<(string RelationshipId, string Name)> Sheets,
        Biff12Record? CalculationProperties);

    /// <summary>An UncheckedRfX: zero-based first/last row and column, converted to one-based rows.</summary>
    private readonly record struct RowRange(int FirstRow, int LastRow, uint FirstColumn, uint LastColumn)
    {
        private const uint MaximumColumnCount = 16_384;

        public static RowRange Read(Biff12Record record, int offset)
        {
            uint firstRow = record.ReadUInt32(offset);
            uint lastRow = record.ReadUInt32(offset + 4);
            uint firstColumn = record.ReadUInt32(offset + 8);
            uint lastColumn = record.ReadUInt32(offset + 12);
            if (lastRow < firstRow || firstRow >= FormulaRemovalRequest.MaximumRowNumber ||
                lastRow >= FormulaRemovalRequest.MaximumRowNumber || lastColumn < firstColumn ||
                lastColumn >= MaximumColumnCount)
            {
                throw new InvalidDataException($"Binary workbook record {record.Type} has an invalid cell range.");
            }

            return new RowRange((int)firstRow + 1, (int)lastRow + 1, firstColumn, lastColumn);
        }

        public string FirstCellReference() => CellReference(FirstColumn, FirstRow);

        public string ToA1Reference()
        {
            string first = FirstCellReference();
            string last = CellReference(LastColumn, LastRow);
            return first == last ? first : $"{first}:{last}";
        }

        private static string CellReference(uint zeroBasedColumn, int row)
        {
            Span<char> letters = stackalloc char[3];
            int position = letters.Length;
            uint column = zeroBasedColumn + 1;
            while (column > 0)
            {
                column--;
                letters[--position] = (char)('A' + (int)(column % 26));
                column /= 26;
            }

            return string.Concat(letters[position..], row.ToString(CultureInfo.InvariantCulture));
        }
    }

    /// <summary>[MS-XLSB] record numbers used by the rewriter, confirmed against Excel output.</summary>
    private static class RecordType
    {
        public const int RowHeader = 0;
        public const int CellError = 3;
        public const int CellBoolean = 4;
        public const int CellReal = 5;
        public const int CellString = 6;
        public const int FormulaString = 8;
        public const int FormulaNumber = 9;
        public const int FormulaBoolean = 10;
        public const int FormulaError = 11;
        public const int CellMetadata = 49;
        public const int CalculationProperties = 157;
        public const int BeginList = 343;
        public const int BeginListColumn = 347;
        public const int EndListColumn = 348;
        public const int ListCalculatedColumnFormula = 351;
        public const int ListTotalsFormula = 352;
        public const int ArrayFormula = 426;
        public const int SharedFormula = 427;
        public const int DataTable = 428;
    }

    /// <summary>BrtCalcProp: recalcID, fAutoRecalc, cCalcCount, xnumDelta, cUserThreadCount, flags.</summary>
    private static class CalculationPropertiesLayout
    {
        public const int RecalcIdOffset = 0;
        public const int AutoRecalcOffset = 4;
        public const int FlagsOffset = 24;

        /// <summary>
        /// fFullCalcOnLoad (bit 0), fSomeUncalced (bit 4), and fNoDeps (bit 8): the binary
        /// forms of .xlsx fullCalcOnLoad, calcCompleted="0", and forceFullCalc, confirmed by
        /// saving Excel workbooks with each setting.
        /// </summary>
        public const ushort ForceRecalculationFlags = 0x0001 | 0x0010 | 0x0100;
    }

    /// <summary>BrtBeginList fixed-field offsets.</summary>
    private static class TableLayout
    {
        public const int HeaderRowCountOffset = 24;
        public const int TotalsRowCountOffset = 28;
        public const int NameOffset = 64;
        public const int DisplayNameOffset = NameOffset + sizeof(uint);
    }
}
