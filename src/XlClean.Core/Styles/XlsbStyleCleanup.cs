namespace XlClean.Core.Styles;

using System.Buffers.Binary;
using System.IO.Packaging;

/// <summary>
/// Reads and rewrites the style references of an .xlsb package: styles.bin, cell, row, and column
/// records in worksheets, and the cell style names tables use.
/// </summary>
/// <remarks>
/// Only styles.bin and, when a referenced cell format moves, the worksheets are written. A
/// rewritten worksheet loses its binary index, an optional load accelerator that stores byte
/// offsets into the old stream; Excel rebuilds it.
/// </remarks>
internal static class XlsbStyleCleanup
{
    public const string BinaryIndexRelationship =
        "http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex";

    private const int BufferSize = 1024 * 1024;
    private const int CancellationCheckInterval = 4096;

    public static (StyleSheetModel Model, StyleUsage Usage) Read(string workbookPath, CancellationToken cancellationToken)
    {
        using Package package = Package.Open(workbookPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        try
        {
            PackagePart workbookPart = XlsbPackage.GetWorkbookPart(package);
            OpenXmlStyleReader.RejectUnsupportedParts(workbookPart);
            PackagePart? stylesPart = StylesPart(workbookPart);
            StyleSheetModel model = stylesPart is null
                ? StyleSheetModel.Empty
                : XlsbStyleSheet.Read(stylesPart, cancellationToken).Model;

            HashSet<int> cellFormats = [0];
            HashSet<string> tableStyleNames = new(StringComparer.Ordinal);
            foreach (PackagePart worksheet in Worksheets(workbookPart, cancellationToken))
            {
                ScanWorksheet(worksheet, record => cellFormats.Add(record), cancellationToken);
                foreach (PackagePart table in XlsbPackage.GetRelatedParts(worksheet, XlsbPackage.TableRelationship))
                {
                    tableStyleNames.UnionWith(TableStyleNames(table));
                }
            }

            int highest = cellFormats.Max();
            if (highest > 0 && highest >= model.CellFormats.Count)
            {
                throw new InvalidDataException(
                    $"A cell references cell format {highest}, but the workbook defines only {model.CellFormats.Count}.");
            }

            return (model, new StyleUsage(cellFormats, tableStyleNames));
        }
        catch (ArgumentException exception)
        {
            throw XlsbPackage.MalformedPackage(exception);
        }
    }

    public static void Apply(string workbookPath, StylePlan plan, CancellationToken cancellationToken)
    {
        if (plan.IsIdentity)
        {
            return;
        }

        using Package package = Package.Open(workbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        PackagePart workbookPart = XlsbPackage.GetWorkbookPart(package);
        PackagePart stylesPart = StylesPart(workbookPart) ??
            throw new InvalidDataException("The workbook has no styles part to clean.");
        RewriteStyleSheet(stylesPart, plan, cancellationToken);

        if (plan.RenumbersCellFormats)
        {
            foreach (PackagePart worksheet in Worksheets(workbookPart, cancellationToken))
            {
                RewriteWorksheet(package, worksheet, plan.CellFormatMap, workbookPath, cancellationToken);
            }
        }
    }

    public static PackagePart? StylesPart(PackagePart workbookPart) =>
        XlsbPackage.GetRelatedParts(workbookPart, OpenXmlStyleReader.StylesRelationship).SingleOrDefault();

    public static IEnumerable<PackagePart> Worksheets(PackagePart workbookPart, CancellationToken cancellationToken)
    {
        foreach ((string relationshipId, _) in XlsbPackage.ReadSheets(XlsbPackage.ReadAllRecords(workbookPart, cancellationToken)))
        {
            if (XlsbPackage.GetRelatedPart(workbookPart, relationshipId, XlsbPackage.WorksheetRelationship) is { } worksheet)
            {
                yield return worksheet;
            }
        }
    }

    /// <summary>
    /// Reports every cell format a worksheet references and refuses records in sheet data that
    /// could hold a reference the cleanup does not know.
    /// </summary>
    private static void ScanWorksheet(PackagePart worksheet, Action<int> reference, CancellationToken cancellationToken)
    {
        using Stream stream = worksheet.GetStream(FileMode.Open, FileAccess.Read);
        Biff12Reader reader = new(new BufferedStream(stream, BufferSize));
        bool inSheetData = false;
        long count = 0;
        while (reader.TryRead(out Biff12Record record))
        {
            if (++count % CancellationCheckInterval == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            inSheetData = Track(record, inSheetData, worksheet);
            if (StyleIndex(record) is { } index)
            {
                reference(index);
            }
        }
    }

    private static bool Track(Biff12Record record, bool inSheetData, PackagePart worksheet)
    {
        if (record.Type == XlsbStyleRecords.BeginSheetData)
        {
            return true;
        }

        if (record.Type == XlsbStyleRecords.EndSheetData)
        {
            return false;
        }

        if (inSheetData && !XlsbStyleRecords.SheetDataRecords.Contains(record.Type))
        {
            throw new UnsupportedStyleContentException(
                $"Worksheet '{worksheet.Uri}' contains record {record.Type} in its cell data, which this tool does not rewrite.");
        }

        return inSheetData;
    }

    /// <returns>The cell format a record applies, or null when it applies none.</returns>
    private static int? StyleIndex(Biff12Record record)
    {
        if (XlsbStyleRecords.CellRecords.Contains(record.Type))
        {
            record.RequireLength(8);
            return XlsbStyleRecords.ReadCellStyle(record.Data);
        }

        if (record.Type == XlsbStyleRecords.RowHeader)
        {
            record.RequireLength(XlsbStyleRecords.RowFlagsOffset + 1);
            return XlsbStyleRecords.RowStyleApplies(record.Data)
                ? ToIndex(record.ReadUInt32(XlsbStyleRecords.RowStyleOffset))
                : null;
        }

        if (record.Type == XlsbStyleRecords.ColumnInfo)
        {
            return ToIndex(record.ReadUInt32(XlsbStyleRecords.ColumnStyleOffset));
        }

        return null;
    }

    private static int ToIndex(uint value) =>
        value <= int.MaxValue ? (int)value : throw new InvalidDataException($"The style index {value} is out of range.");

    // BrtBeginList: name and display name at 64, then comment, header, data, and totals styles.
    // BrtBeginListCol: 24 fixed bytes, then name, caption, totals text, and the three styles.
    private static List<string> TableStyleNames(PackagePart table)
    {
        List<string> names = [];
        foreach (Biff12Record record in XlsbPackage.ReadAllRecords(table, CancellationToken.None))
        {
            (int offset, int skip) = record.Type switch
            {
                XlsbStyleRecords.BeginList => (64, 3),
                XlsbStyleRecords.BeginListColumn => (24, 3),
                _ => (-1, 0),
            };
            if (offset < 0)
            {
                continue;
            }

            for (int index = 0; index < skip + 3; index++)
            {
                string? value = record.ReadWideString(offset, out offset);
                if (index >= skip && value is not null)
                {
                    names.Add(value);
                }
            }
        }

        return names;
    }

    private static void RewriteStyleSheet(PackagePart stylesPart, StylePlan plan, CancellationToken cancellationToken)
    {
        XlsbStyleSheet sheet = XlsbStyleSheet.Read(stylesPart, cancellationToken);
        List<Biff12Record> records = sheet.Records;

        HashSet<int> seenTargets = [];
        bool[] representative = [.. plan.CellFormatMap.Select(target => target is { } value && seenTargets.Add(value))];
        HashSet<int> keptStyles = [.. plan.KeptNamedStyles];

        // Each record index maps to its replacement: the same record, a patched one, or nothing.
        Biff12Record?[] output = [.. records.Select(record => (Biff12Record?)record)];
        Rewrite(sheet.StyleFormats, plan.CellStyleFormatMap.Select(target => target is not null).ToArray(), patch: null);
        Rewrite(sheet.CellFormats, representative, data => PatchParent(data, plan.CellStyleFormatMap));
        Rewrite(sheet.Styles, Enumerable.Range(0, sheet.Styles.Groups.Count).Select(keptStyles.Contains).ToArray(), data =>
            BinaryPrimitives.WriteUInt32LittleEndian(
                data,
                (uint)Target(plan.CellStyleFormatMap, (int)BinaryPrimitives.ReadUInt32LittleEndian(data), "cell style format")));

        SetCount(sheet.StyleFormats, plan.CellStyleFormatMap.Count(target => target is not null));
        SetCount(sheet.CellFormats, representative.Count(kept => kept));
        SetCount(sheet.Styles, keptStyles.Count);

        using Stream destination = stylesPart.GetStream(FileMode.Create, FileAccess.Write);
        foreach (Biff12Record? record in output)
        {
            if (record is { } kept)
            {
                Biff12Writer.Write(destination, kept);
            }
        }

        void Rewrite(XlsbStyleSheet.Section section, bool[] keep, Action<byte[]>? patch)
        {
            for (int index = 0; index < section.Groups.Count; index++)
            {
                XlsbStyleSheet.Group group = section.Groups[index];
                if (!keep[index])
                {
                    for (int record = group.Start; record < group.End; record++)
                    {
                        output[record] = null;
                    }
                }
                else if (patch is not null)
                {
                    byte[] data = (byte[])records[group.Item].Data.Clone();
                    patch(data);
                    output[group.Item] = new Biff12Record(records[group.Item].Type, data);
                }
            }
        }

        void SetCount(XlsbStyleSheet.Section section, int count)
        {
            if (section.Begin < 0)
            {
                return;
            }

            Biff12Record begin = records[section.Begin];
            begin.RequireLength(4);
            byte[] data = (byte[])begin.Data.Clone();
            BinaryPrimitives.WriteUInt32LittleEndian(data, (uint)count);
            output[section.Begin] = new Biff12Record(begin.Type, data);
        }
    }

    private static void PatchParent(byte[] data, IReadOnlyList<int?> styleFormatMap)
    {
        int parent = XlsbStyleRecords.ReadUInt16(data, 0);
        if (parent != XlsbStyleRecords.NoParent)
        {
            int target = Target(styleFormatMap, parent, "cell style format");
            BinaryPrimitives.WriteUInt16LittleEndian(data, checked((ushort)target));
        }
    }

    private static void RewriteWorksheet(
        Package package,
        PackagePart worksheet,
        IReadOnlyList<int?> cellFormatMap,
        string workbookPath,
        CancellationToken cancellationToken)
    {
        string stagingPath = Path.Combine(
            Path.GetDirectoryName(workbookPath)!,
            $".xlclean-{Guid.NewGuid():N}.worksheet.tmp");
        try
        {
            using (Stream source = worksheet.GetStream(FileMode.Open, FileAccess.Read))
            using (FileStream staging = new(stagingPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, BufferSize))
            {
                Biff12Reader reader = new(new BufferedStream(source, BufferSize));
                bool inSheetData = false;
                long count = 0;
                while (reader.TryRead(out Biff12Record record))
                {
                    if (++count % CancellationCheckInterval == 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                    }

                    inSheetData = Track(record, inSheetData, worksheet);
                    Biff12Writer.Write(staging, Renumber(record, cellFormatMap));
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            using (FileStream transformed = new(stagingPath, FileMode.Open, FileAccess.Read, FileShare.Read, BufferSize))
            using (Stream destination = worksheet.GetStream(FileMode.Create, FileAccess.Write))
            {
                transformed.CopyTo(destination, BufferSize);
            }

            XlsbPackage.DeleteRelatedParts(package, worksheet, BinaryIndexRelationship);
        }
        finally
        {
            WorkbookBatchPipeline.TryDeleteStagingFile(stagingPath);
        }
    }

    private static Biff12Record Renumber(Biff12Record record, IReadOnlyList<int?> cellFormatMap)
    {
        if (StyleIndex(record) is not { } index)
        {
            if (record.Type == XlsbStyleRecords.RowHeader)
            {
                return RenumberUnappliedRowStyle(record, cellFormatMap);
            }

            return record;
        }

        int target = Target(cellFormatMap, index, "cell format");
        if (target == index)
        {
            return record;
        }

        byte[] data = (byte[])record.Data.Clone();
        if (XlsbStyleRecords.CellRecords.Contains(record.Type))
        {
            XlsbStyleRecords.WriteCellStyle(data, target);
        }
        else
        {
            int offset = record.Type == XlsbStyleRecords.RowHeader
                ? XlsbStyleRecords.RowStyleOffset
                : XlsbStyleRecords.ColumnStyleOffset;
            BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(offset), (uint)target);
        }

        return new Biff12Record(record.Type, data);
    }

    // Excel writes 0 in a row's index when fGhostDirty is clear, and ignores the field. A non-zero
    // value is kept pointing at an equivalent format when there is one; otherwise it could fall
    // outside the shrunk table, so the file fails rather than guess.
    private static Biff12Record RenumberUnappliedRowStyle(Biff12Record record, IReadOnlyList<int?> cellFormatMap)
    {
        uint stale = record.ReadUInt32(XlsbStyleRecords.RowStyleOffset);
        if (stale == 0)
        {
            return record;
        }

        if (stale >= cellFormatMap.Count || cellFormatMap[(int)stale] is not { } target)
        {
            throw new UnsupportedStyleContentException(
                $"A row carries an unused style index ({stale}) that this tool cannot keep valid after cleanup.");
        }

        if (target == stale)
        {
            return record;
        }

        byte[] data = (byte[])record.Data.Clone();
        BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(XlsbStyleRecords.RowStyleOffset), (uint)target);
        return new Biff12Record(record.Type, data);
    }

    private static int Target(IReadOnlyList<int?> map, int index, string kind) =>
        index < map.Count && map[index] is { } target
            ? target
            : throw new InvalidDataException($"The workbook references {kind} {index}, which the cleanup plan does not cover.");
}
