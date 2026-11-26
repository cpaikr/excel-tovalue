namespace XlClean.Core.Styles;

using System.IO.Packaging;
using System.Text;

/// <summary>
/// Proves that a style-cleaned .xlsb package looks the same as its source.
/// </summary>
/// <remarks>
/// The .xlsb counterpart of <see cref="OpenXmlStyleVerifier"/>, with the same independence: it
/// does not use the planner or the rewriter's section model. Formats are resolved from raw
/// records on each side and compared as text. What may differ: the three rewritten style sections,
/// the style index fields of cell, row, and column records, and the binary index of a rewritten
/// worksheet, which is removed. Everything else must be byte-identical.
/// </remarks>
internal static class XlsbStyleVerifier
{
    private const int FirstCustomNumberFormatId = 164;

    private static readonly HashSet<int> RewrittenSectionBegins =
    [
        XlsbStyleRecords.BeginCellStyleFormats,
        XlsbStyleRecords.BeginCellFormats,
        XlsbStyleRecords.BeginStyles,
    ];

    public static void Verify(string sourcePath, string outputPath, CancellationToken cancellationToken)
    {
        using Package source = Package.Open(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using Package output = Package.Open(outputPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        try
        {
            PackagePart sourceWorkbook = XlsbPackage.GetWorkbookPart(source);
            PackagePart? sourceStyles = XlsbStyleCleanup.StylesPart(sourceWorkbook);
            List<PackagePart> worksheets = [.. XlsbStyleCleanup.Worksheets(sourceWorkbook, cancellationToken)];
            HashSet<Uri> rewritable = [.. worksheets.Select(part => part.Uri)];
            if (sourceStyles is not null)
            {
                rewritable.Add(sourceStyles.Uri);
            }

            HashSet<Uri> removableIndexes = [.. worksheets
                .SelectMany(worksheet => XlsbPackage.GetRelatedParts(worksheet, XlsbStyleCleanup.BinaryIndexRelationship))
                .Select(part => part.Uri)];
            ComparePackageStructure(source, output, rewritable, removableIndexes);

            Sheet sourceSheet = Sheet.Load(sourceStyles, cancellationToken);
            Sheet outputSheet = Sheet.Load(sourceStyles is null ? null : output.GetPart(sourceStyles.Uri), cancellationToken);
            CompareUnrewrittenSections(sourceSheet, outputSheet);
            CompareNamedStyles(sourceSheet, outputSheet, worksheets);
            outputSheet.CheckEveryFormatResolves();
            if (sourceSheet.ResolveCellFormat(0) != outputSheet.ResolveCellFormat(0))
            {
                throw new StyleVerificationException("The default cell format changed.");
            }

            if (sourceSheet.ResolveDefaultStyleFormat() != outputSheet.ResolveDefaultStyleFormat())
            {
                throw new StyleVerificationException("The default cell style format changed.");
            }

            Dictionary<(int, int), bool> equivalent = [];
            foreach (PackagePart worksheet in worksheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                CompareWorksheet(worksheet, output.GetPart(worksheet.Uri), sourceSheet, outputSheet, equivalent, cancellationToken);
            }
        }
        catch (ArgumentException exception)
        {
            throw XlsbPackage.MalformedPackage(exception);
        }
    }

    private static void ComparePackageStructure(Package source, Package output, HashSet<Uri> rewritable, HashSet<Uri> removableIndexes)
    {
        Dictionary<Uri, PackagePart> outputParts = output.GetParts()
            .Where(part => !PackUriHelper.IsRelationshipPartUri(part.Uri))
            .ToDictionary(part => part.Uri);
        CompareRelationships("the package", source.GetRelationships(), output.GetRelationships(), removableIndexes, output.PartExists);
        foreach (PackagePart sourcePart in source.GetParts().Where(part => !PackUriHelper.IsRelationshipPartUri(part.Uri)))
        {
            if (!outputParts.Remove(sourcePart.Uri, out PackagePart? outputPart))
            {
                if (removableIndexes.Contains(sourcePart.Uri) && !IsStillReferenced(output, sourcePart.Uri))
                {
                    continue;
                }

                throw new StyleVerificationException($"Part '{sourcePart.Uri}' is missing.");
            }

            if (!string.Equals(sourcePart.ContentType, outputPart.ContentType, StringComparison.Ordinal))
            {
                throw new StyleVerificationException($"Part '{sourcePart.Uri}' changed content type.");
            }

            CompareRelationships($"part '{sourcePart.Uri}'", sourcePart.GetRelationships(), outputPart.GetRelationships(), removableIndexes, output.PartExists);
            if (!rewritable.Contains(sourcePart.Uri) && !SameBytes(sourcePart, outputPart))
            {
                throw new StyleVerificationException($"Part '{sourcePart.Uri}' changed.");
            }
        }

        if (outputParts.Count > 0)
        {
            throw new StyleVerificationException($"Part '{outputParts.Keys.First()}' was added.");
        }
    }

    private static bool IsStillReferenced(Package output, Uri part) =>
        output.GetParts()
            .Where(candidate => !PackUriHelper.IsRelationshipPartUri(candidate.Uri))
            .SelectMany(candidate => candidate.GetRelationships())
            .Any(relationship => relationship.TargetMode == TargetMode.Internal &&
                PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri) == part);

    // A rewritten worksheet may drop the relationship to its binary index, but only together with
    // the index part itself; nothing else may change.
    private static void CompareRelationships(
        string owner,
        PackageRelationshipCollection source,
        PackageRelationshipCollection output,
        HashSet<Uri> removableIndexes,
        Func<Uri, bool> outputPartExists)
    {
        string[] expected = [.. source.Select(Describe).Order(StringComparer.Ordinal)];
        string[] actual = [.. output.Select(Describe).Order(StringComparer.Ordinal)];
        string[] missing = [.. expected.Except(actual, StringComparer.Ordinal)];
        bool onlyIndexesRemoved = actual.All(expected.Contains) && source
            .Where(relationship => missing.Contains(Describe(relationship)))
            .All(relationship =>
                relationship.RelationshipType == XlsbStyleCleanup.BinaryIndexRelationship &&
                removableIndexes.Contains(PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri)) &&
                !outputPartExists(PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri)));
        if (!onlyIndexesRemoved)
        {
            throw new StyleVerificationException($"The relationships of {owner} changed.");
        }

        static string Describe(PackageRelationship relationship) =>
            $"{relationship.Id}|{relationship.RelationshipType}|{relationship.TargetUri.OriginalString}|{relationship.TargetMode}";
    }

    private static bool SameBytes(PackagePart source, PackagePart output)
    {
        using Stream first = source.GetStream(FileMode.Open, FileAccess.Read);
        using Stream second = output.GetStream(FileMode.Open, FileAccess.Read);
        byte[] firstBuffer = new byte[81920];
        byte[] secondBuffer = new byte[81920];
        while (true)
        {
            int firstRead = first.ReadAtLeast(firstBuffer, firstBuffer.Length, throwOnEndOfStream: false);
            int secondRead = second.ReadAtLeast(secondBuffer, secondBuffer.Length, throwOnEndOfStream: false);
            if (firstRead != secondRead || !firstBuffer.AsSpan(0, firstRead).SequenceEqual(secondBuffer.AsSpan(0, secondRead)))
            {
                return false;
            }

            if (firstRead == 0)
            {
                return true;
            }
        }
    }

    private static void CompareUnrewrittenSections(Sheet source, Sheet output)
    {
        string[] expected = source.Unrewritten();
        string[] actual = output.Unrewritten();
        if (!expected.SequenceEqual(actual, StringComparer.Ordinal))
        {
            throw new StyleVerificationException("A style sheet section other than formats and named styles changed.");
        }

        output.CheckCounts();
    }

    private static void CompareNamedStyles(Sheet source, Sheet output, List<PackagePart> worksheets)
    {
        Dictionary<string, int> remaining = new(StringComparer.Ordinal);
        foreach (string style in source.DescribeNamedStyles())
        {
            remaining[style] = remaining.GetValueOrDefault(style) + 1;
        }

        foreach (string style in output.DescribeNamedStyles())
        {
            if (remaining.GetValueOrDefault(style) == 0)
            {
                throw new StyleVerificationException("An output cell style is not an unchanged source style.");
            }

            remaining[style]--;
        }

        HashSet<string> outputNames = [.. output.StyleNames()];
        foreach ((string name, bool builtIn) in source.StyleNamesWithBuiltIn())
        {
            if (builtIn && !outputNames.Contains(name))
            {
                throw new StyleVerificationException($"Built-in cell style '{name}' was removed.");
            }
        }

        HashSet<string> sourceNames = [.. source.StyleNames()];
        foreach (string name in TableStyleNames(worksheets))
        {
            if (sourceNames.Contains(name) && !outputNames.Contains(name))
            {
                throw new StyleVerificationException($"Cell style '{name}', which a table uses, was removed.");
            }
        }
    }

    // From the record definitions ([MS-XLSB] 2.4.100, 2.4.101), not the reader: a list record's
    // strings at 64 are name, display name, comment, then the three styles; a list column's at 24
    // are name, caption, totals text, then the three styles.
    private static IEnumerable<string> TableStyleNames(List<PackagePart> worksheets)
    {
        foreach (PackagePart worksheet in worksheets)
        {
            foreach (PackagePart table in XlsbPackage.GetRelatedParts(worksheet, XlsbPackage.TableRelationship))
            {
                foreach (Biff12Record record in XlsbPackage.ReadAllRecords(table, CancellationToken.None))
                {
                    int offset = record.Type switch
                    {
                        XlsbStyleRecords.BeginList => 64,
                        XlsbStyleRecords.BeginListColumn => 24,
                        _ => -1,
                    };
                    for (int index = 0; offset >= 0 && index < 6; index++)
                    {
                        string? value = record.ReadWideString(offset, out offset);
                        if (index >= 3 && value is not null)
                        {
                            yield return value;
                        }
                    }
                }
            }
        }
    }

    private static void CompareWorksheet(
        PackagePart source,
        PackagePart output,
        Sheet sourceSheet,
        Sheet outputSheet,
        Dictionary<(int, int), bool> equivalent,
        CancellationToken cancellationToken)
    {
        using Stream sourceStream = source.GetStream(FileMode.Open, FileAccess.Read);
        using Stream outputStream = output.GetStream(FileMode.Open, FileAccess.Read);
        Biff12Reader sourceReader = new(new BufferedStream(sourceStream, 1024 * 1024));
        Biff12Reader outputReader = new(new BufferedStream(outputStream, 1024 * 1024));
        long count = 0;
        while (true)
        {
            bool hasSource = sourceReader.TryRead(out Biff12Record expected);
            bool hasOutput = outputReader.TryRead(out Biff12Record actual);
            if (hasSource != hasOutput)
            {
                throw new StyleVerificationException($"Worksheet '{source.Uri}' has a different number of records.");
            }

            if (!hasSource)
            {
                return;
            }

            if (++count % 4096 == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            (int Offset, int Length)? style = StyleField(expected);
            if (expected.Type != actual.Type || expected.Data.Length != actual.Data.Length ||
                !SameExcept(expected.Data, actual.Data, style))
            {
                throw new StyleVerificationException($"Worksheet '{source.Uri}' record {count} ({expected.Type}) changed.");
            }

            if (style is { } field &&
                !expected.Data.AsSpan(field.Offset, field.Length).SequenceEqual(actual.Data.AsSpan(field.Offset, field.Length)))
            {
                int sourceIndex = ReadIndex(expected.Data, field);
                int outputIndex = ReadIndex(actual.Data, field);
                if (!equivalent.TryGetValue((sourceIndex, outputIndex), out bool same))
                {
                    same = sourceSheet.ResolveCellFormat(sourceIndex) == outputSheet.ResolveCellFormat(outputIndex);
                    equivalent.Add((sourceIndex, outputIndex), same);
                }

                if (!same)
                {
                    throw new StyleVerificationException(
                        $"Worksheet '{source.Uri}' record {count} looks different: cell format {sourceIndex} became {outputIndex}.");
                }
            }
        }
    }

    /// <summary>The bytes of a record that hold a style index, or null when it holds none in effect.</summary>
    private static (int Offset, int Length)? StyleField(Biff12Record record)
    {
        if (XlsbStyleRecords.CellRecords.Contains(record.Type) && record.Data.Length >= 8)
        {
            return (XlsbStyleRecords.CellStyleOffset, 3);
        }

        // A row's index is compared even when fGhostDirty says it does not apply, because the
        // rewriter keeps it in range by renumbering it too.
        if (record.Type == XlsbStyleRecords.RowHeader && record.Data.Length > XlsbStyleRecords.RowFlagsOffset)
        {
            return (XlsbStyleRecords.RowStyleOffset, 4);
        }

        if (record.Type == XlsbStyleRecords.ColumnInfo && record.Data.Length >= XlsbStyleRecords.ColumnStyleOffset + 4)
        {
            return (XlsbStyleRecords.ColumnStyleOffset, 4);
        }

        return null;
    }

    private static bool SameExcept(byte[] expected, byte[] actual, (int Offset, int Length)? field)
    {
        if (field is not { } skip)
        {
            return expected.AsSpan().SequenceEqual(actual);
        }

        return expected.AsSpan(0, skip.Offset).SequenceEqual(actual.AsSpan(0, skip.Offset)) &&
            expected.AsSpan(skip.Offset + skip.Length).SequenceEqual(actual.AsSpan(skip.Offset + skip.Length));
    }

    private static int ReadIndex(byte[] data, (int Offset, int Length) field)
    {
        int value = 0;
        for (int index = field.Length - 1; index >= 0; index--)
        {
            value = (value << 8) | data[field.Offset + index];
        }

        return value;
    }

    /// <summary>One side's styles.bin, read as a flat record list and resolved to text.</summary>
    private sealed class Sheet
    {
        private readonly List<Biff12Record> records;
        private readonly List<(Biff12Record Format, string Attached)> cellFormats;
        private readonly List<(Biff12Record Format, string Attached)> styleFormats;
        private readonly List<(Biff12Record Style, string Prefix)> styles;
        private readonly Dictionary<int, string> numberFormats = [];
        private readonly List<string> fonts = [];
        private readonly List<string> fills = [];
        private readonly List<string> borders = [];

        private Sheet(List<Biff12Record> records)
        {
            this.records = records;
            styleFormats = Formats(XlsbStyleRecords.BeginCellStyleFormats, XlsbStyleRecords.EndCellStyleFormats);
            cellFormats = Formats(XlsbStyleRecords.BeginCellFormats, XlsbStyleRecords.EndCellFormats);
            styles = NamedStyles();
            foreach (Biff12Record record in records)
            {
                switch (record.Type)
                {
                    case XlsbStyleRecords.NumberFormat:
                        numberFormats[XlsbStyleRecords.ReadUInt16(record.Data, 0)] = Convert.ToHexString(record.Data, 2, record.Data.Length - 2);
                        break;
                    case XlsbStyleRecords.Font:
                        fonts.Add(Convert.ToHexString(record.Data));
                        break;
                    case XlsbStyleRecords.Fill:
                        fills.Add(Convert.ToHexString(record.Data));
                        break;
                    case XlsbStyleRecords.Border:
                        borders.Add(Convert.ToHexString(record.Data));
                        break;
                }
            }
        }

        public static Sheet Load(PackagePart? part, CancellationToken cancellationToken) =>
            new(part is null ? [] : XlsbPackage.ReadAllRecords(part, cancellationToken));

        /// <summary>Every record outside the three rewritten sections, as text, in order.</summary>
        public string[] Unrewritten()
        {
            List<string> result = [];
            int? skippingUntil = null;
            foreach (Biff12Record record in records)
            {
                if (skippingUntil is { } end)
                {
                    if (record.Type == end)
                    {
                        skippingUntil = null;
                        result.Add($"section-end {record.Type}");
                    }

                    continue;
                }

                if (RewrittenSectionBegins.Contains(record.Type))
                {
                    skippingUntil = record.Type switch
                    {
                        XlsbStyleRecords.BeginCellStyleFormats => XlsbStyleRecords.EndCellStyleFormats,
                        XlsbStyleRecords.BeginCellFormats => XlsbStyleRecords.EndCellFormats,
                        _ => XlsbStyleRecords.EndStyles,
                    };
                    result.Add($"section-begin {record.Type}");
                    continue;
                }

                result.Add($"{record.Type}:{Convert.ToHexString(record.Data)}");
            }

            return [.. result];
        }

        public void CheckCounts()
        {
            Check(XlsbStyleRecords.BeginCellStyleFormats, styleFormats.Count);
            Check(XlsbStyleRecords.BeginCellFormats, cellFormats.Count);
            Check(XlsbStyleRecords.BeginStyles, styles.Count);

            void Check(int beginType, int actual)
            {
                if (records.FirstOrDefault(record => record.Type == beginType) is { Data.Length: >= 4 } begin &&
                    begin.ReadUInt32(0) != actual)
                {
                    throw new StyleVerificationException($"Style section {beginType} declares {begin.ReadUInt32(0)} entries but has {actual}.");
                }
            }
        }

        public IEnumerable<string> DescribeNamedStyles() =>
            styles.Select(style => $"{Convert.ToHexString(style.Style.Data, 4, style.Style.Data.Length - 4)}{style.Prefix}=>{ResolveStyleFormat((int)style.Style.ReadUInt32(0), includeNames: false)}");

        public IEnumerable<string> StyleNames() => styles.Select(style => style.Style.ReadWideString(8, out _) ?? string.Empty);

        public IEnumerable<(string Name, bool BuiltIn)> StyleNamesWithBuiltIn() =>
            styles.Select(style => (style.Style.ReadWideString(8, out _) ?? string.Empty, (style.Style.Data[4] & 0x01) != 0));

        public string ResolveCellFormat(int index)
        {
            if (cellFormats.Count == 0)
            {
                return index == 0
                    ? "implicit default"
                    : throw new StyleVerificationException($"Cell format {index} does not exist.");
            }

            if (index >= cellFormats.Count)
            {
                throw new StyleVerificationException($"Cell format {index} does not exist.");
            }

            (Biff12Record format, string attached) = cellFormats[index];
            int parent = XlsbStyleRecords.ReadUInt16(format.Data, 0);
            return Resolve(format, attached) + "|parent=" +
                (parent == XlsbStyleRecords.NoParent ? "none" : ResolveStyleFormat(parent, includeNames: true));
        }

        public string ResolveDefaultStyleFormat() =>
            styleFormats.Count == 0 ? "no style formats" : ResolveStyleFormat(0, includeNames: true);

        public void CheckEveryFormatResolves()
        {
            for (int index = 0; index < cellFormats.Count; index++)
            {
                _ = ResolveCellFormat(index);
            }

            for (int index = 0; index < styleFormats.Count; index++)
            {
                _ = ResolveStyleFormat(index, includeNames: false);
            }
        }

        public string ResolveStyleFormat(int index, bool includeNames)
        {
            if (index >= styleFormats.Count)
            {
                throw new StyleVerificationException($"Cell style format {index} does not exist.");
            }

            (Biff12Record format, string attached) = styleFormats[index];
            StringBuilder resolved = new(Resolve(format, attached));
            if (includeNames)
            {
                foreach (string style in styles
                    .Where(style => style.Style.ReadUInt32(0) == index)
                    .Select(style => Convert.ToHexString(style.Style.Data, 4, style.Style.Data.Length - 4) + style.Prefix)
                    .Order(StringComparer.Ordinal))
                {
                    resolved.Append("|style=").Append(style);
                }
            }

            return resolved.ToString();
        }

        private string Resolve(Biff12Record format, string attached)
        {
            if (format.Data.Length < XlsbStyleRecords.FormatLength)
            {
                throw new StyleVerificationException("A format record is truncated.");
            }

            int numberFormat = XlsbStyleRecords.ReadUInt16(format.Data, 2);
            string numberFormatText = numberFormats.TryGetValue(numberFormat, out string? code)
                ? (numberFormat < FirstCustomNumberFormatId ? $"builtin {numberFormat} redefined {code}" : $"custom {code}")
                : numberFormat < FirstCustomNumberFormatId
                    ? $"builtin {numberFormat}"
                    : throw new StyleVerificationException($"Number format {numberFormat} is not defined.");
            return $"own={Convert.ToHexString(format.Data, XlsbStyleRecords.FormatIndexLength, format.Data.Length - XlsbStyleRecords.FormatIndexLength)}{attached}" +
                $"|numFmt={numberFormatText}" +
                $"|font={Component(fonts, XlsbStyleRecords.ReadUInt16(format.Data, 4), "font")}" +
                $"|fill={Component(fills, XlsbStyleRecords.ReadUInt16(format.Data, 6), "fill")}" +
                $"|border={Component(borders, XlsbStyleRecords.ReadUInt16(format.Data, 8), "border")}";
        }

        private static string Component(List<string> table, int index, string kind) =>
            index < table.Count ? table[index] : throw new StyleVerificationException($"The {kind} {index} does not exist.");

        // A format record owns the records that follow it up to the next format or the section end.
        private List<(Biff12Record, string)> Formats(int beginType, int endType)
        {
            List<(Biff12Record, string)> result = [];
            bool inside = false;
            foreach (Biff12Record record in records)
            {
                if (record.Type == beginType)
                {
                    inside = true;
                }
                else if (record.Type == endType)
                {
                    inside = false;
                }
                else if (inside && record.Type == XlsbStyleRecords.Format)
                {
                    result.Add((record, string.Empty));
                }
                else if (inside && result.Count > 0)
                {
                    (Biff12Record format, string attached) = result[^1];
                    result[^1] = (format, $"{attached}|{record.Type}:{Convert.ToHexString(record.Data)}");
                }
            }

            return result;
        }

        // A named style owns the records that precede it since the previous style or section begin.
        private List<(Biff12Record, string)> NamedStyles()
        {
            List<(Biff12Record, string)> result = [];
            bool inside = false;
            StringBuilder prefix = new();
            foreach (Biff12Record record in records)
            {
                if (record.Type == XlsbStyleRecords.BeginStyles)
                {
                    inside = true;
                }
                else if (record.Type == XlsbStyleRecords.EndStyles)
                {
                    inside = false;
                }
                else if (inside && record.Type == XlsbStyleRecords.Style)
                {
                    result.Add((record, prefix.ToString()));
                    prefix.Clear();
                }
                else if (inside)
                {
                    prefix.Append('|').Append(record.Type).Append(':').Append(Convert.ToHexString(record.Data));
                }
            }

            return result;
        }
    }
}
