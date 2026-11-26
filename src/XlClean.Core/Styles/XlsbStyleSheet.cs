namespace XlClean.Core.Styles;

using System.Buffers.Binary;
using System.IO.Packaging;

/// <summary>
/// [MS-XLSB] record numbers and layouts the .xlsb style cleanup relies on, each confirmed against
/// workbooks saved by Excel 16 (2026-11-26).
/// </summary>
internal static class XlsbStyleRecords
{
    public const int Font = 43;
    public const int NumberFormat = 44;
    public const int Fill = 45;
    public const int Border = 46;
    public const int Format = 47;
    public const int Style = 48;
    public const int BeginFonts = 611;
    public const int EndFonts = 612;
    public const int BeginFills = 603;
    public const int EndFills = 604;
    public const int BeginBorders = 613;
    public const int EndBorders = 614;
    public const int BeginNumberFormats = 615;
    public const int EndNumberFormats = 616;
    public const int BeginCellStyleFormats = 626;
    public const int EndCellStyleFormats = 627;
    public const int BeginCellFormats = 617;
    public const int EndCellFormats = 618;
    public const int BeginStyles = 619;
    public const int EndStyles = 620;

    /// <summary>BrtFRTBegin/End and BrtACBegin/End: future-record and alternate-content wrappers.</summary>
    public const int FutureBegin = 35;
    public const int FutureEnd = 36;
    public const int AlternateBegin = 37;
    public const int AlternateEnd = 38;

    /// <summary>The style uid inside the alternate-content group that precedes each BrtStyle.</summary>
    public const int Uid = 3072;

    public const int RowHeader = 0;
    public const int ColumnInfo = 60;
    public const int BeginSheetData = 145;
    public const int EndSheetData = 146;
    public const int RowDescent = 1024;
    public const int BeginList = 343;
    public const int BeginListColumn = 347;

    /// <summary>Cell records whose payload starts with the Cell structure: column, 24-bit iStyleRef, flags.</summary>
    public static readonly IReadOnlySet<int> CellRecords = new HashSet<int> { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 62 };

    /// <summary>
    /// Every record type the cleanup accepts between BrtBeginSheetData and BrtEndSheetData. A record
    /// outside this set might hold a style index the cleanup does not rewrite, so it fails the file.
    /// The compact BrtShort* cell forms are deliberately absent: their layout is not documented and
    /// Excel did not write them in the confirmed samples.
    /// </summary>
    public static readonly IReadOnlySet<int> SheetDataRecords = new HashSet<int>
    {
        RowHeader, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 62,
        49, // BrtCellMeta
        426, 427, 428, // BrtArrFmla, BrtShrFmla, BrtTable
        FutureBegin, FutureEnd, AlternateBegin, AlternateEnd, RowDescent,
    };

    /// <summary>BrtXF: ixfeParent, iFmt, iFont, iFill, ixBorder (16-bit each), then six content bytes.</summary>
    public const int FormatLength = 16;
    public const int FormatIndexLength = 10;
    public const ushort NoParent = 0xFFFF;

    /// <summary>BrtRowHdr: ixfe at 4; fGhostDirty is bit 6 of byte 11 and says whether ixfe applies.</summary>
    public const int RowStyleOffset = 4;
    public const int RowFlagsOffset = 11;
    public const byte GhostDirty = 0x40;

    /// <summary>BrtColInfo: colFirst, colLast, coldx, then ixfe at 12.</summary>
    public const int ColumnStyleOffset = 12;

    /// <summary>The Cell structure's iStyleRef: 24 bits at offset 4.</summary>
    public const int CellStyleOffset = 4;
    public const int MaximumCellStyle = (1 << 24) - 1;

    public static int ReadCellStyle(byte[] data) => data[4] | (data[5] << 8) | (data[6] << 16);

    public static void WriteCellStyle(byte[] data, int value)
    {
        data[4] = (byte)value;
        data[5] = (byte)(value >> 8);
        data[6] = (byte)(value >> 16);
    }

    public static bool RowStyleApplies(byte[] data) => (data[RowFlagsOffset] & GhostDirty) != 0;

    public static int ReadUInt16(byte[] data, int offset) => BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(offset));
}

/// <summary>
/// The styles.bin part split into the three rewritten sections and everything else.
/// </summary>
/// <remarks>
/// Each format record keeps the future records that follow it, and each named style keeps the
/// alternate-content group (its uid) that precedes it, so the two move or disappear together.
/// </remarks>
internal sealed class XlsbStyleSheet
{
    private XlsbStyleSheet(
        List<Biff12Record> records,
        Section styleFormats,
        Section cellFormats,
        Section styles,
        StyleSheetModel model)
    {
        Records = records;
        StyleFormats = styleFormats;
        CellFormats = cellFormats;
        Styles = styles;
        Model = model;
    }

    public List<Biff12Record> Records { get; }

    public Section StyleFormats { get; }

    public Section CellFormats { get; }

    public Section Styles { get; }

    public StyleSheetModel Model { get; }

    public static XlsbStyleSheet Read(PackagePart stylesPart, CancellationToken cancellationToken)
    {
        List<Biff12Record> records = XlsbPackage.ReadAllRecords(stylesPart, cancellationToken);

        Dictionary<int, string> numberFormats = [];
        foreach (Biff12Record record in RecordsIn(records, XlsbStyleRecords.BeginNumberFormats, XlsbStyleRecords.EndNumberFormats))
        {
            if (record.Type == XlsbStyleRecords.NumberFormat)
            {
                record.RequireLength(2);
                int id = XlsbStyleRecords.ReadUInt16(record.Data, 0);
                string code = record.ReadWideString(2, out _) ??
                    throw new InvalidDataException($"Number format {id} has no format code.");
                if (!numberFormats.TryAdd(id, code))
                {
                    throw new InvalidDataException($"Number format {id} is defined more than once.");
                }
            }
        }

        Section styleFormats = Section.Read(records, XlsbStyleRecords.BeginCellStyleFormats, XlsbStyleRecords.EndCellStyleFormats, XlsbStyleRecords.Format, itemLeadsGroup: true);
        Section cellFormats = Section.Read(records, XlsbStyleRecords.BeginCellFormats, XlsbStyleRecords.EndCellFormats, XlsbStyleRecords.Format, itemLeadsGroup: true);
        Section styles = Section.Read(records, XlsbStyleRecords.BeginStyles, XlsbStyleRecords.EndStyles, XlsbStyleRecords.Style, itemLeadsGroup: false);

        StyleSheetModel model = new(
            numberFormats,
            Components(records, XlsbStyleRecords.BeginFonts, XlsbStyleRecords.EndFonts, XlsbStyleRecords.Font),
            Components(records, XlsbStyleRecords.BeginFills, XlsbStyleRecords.EndFills, XlsbStyleRecords.Fill),
            Components(records, XlsbStyleRecords.BeginBorders, XlsbStyleRecords.EndBorders, XlsbStyleRecords.Border),
            [.. styleFormats.Groups.Select(group => ReadFormat(records, group, isStyleFormat: true))],
            [.. cellFormats.Groups.Select(group => ReadFormat(records, group, isStyleFormat: false))],
            [.. styles.Groups.Select(group => ReadNamedStyle(records, group))]);

        return new XlsbStyleSheet(records, styleFormats, cellFormats, styles, model);
    }

    private static string[] Components(List<Biff12Record> records, int begin, int end, int type) =>
        [.. RecordsIn(records, begin, end).Where(record => record.Type == type).Select(record => Convert.ToHexString(record.Data))];

    private static IEnumerable<Biff12Record> RecordsIn(List<Biff12Record> records, int begin, int end)
    {
        int start = records.FindIndex(record => record.Type == begin);
        if (start < 0)
        {
            yield break;
        }

        for (int index = start + 1; index < records.Count && records[index].Type != end; index++)
        {
            yield return records[index];
        }
    }

    private static StyleFormat ReadFormat(List<Biff12Record> records, Group group, bool isStyleFormat)
    {
        Biff12Record format = records[group.Item];
        format.RequireLength(XlsbStyleRecords.FormatLength);
        byte[] data = format.Data;
        int parent = XlsbStyleRecords.ReadUInt16(data, 0);

        // Content without the five indices, plus any future records attached to this format.
        string own = Convert.ToHexString(data, XlsbStyleRecords.FormatIndexLength, data.Length - XlsbStyleRecords.FormatIndexLength) +
            Attached(records, group);
        return new StyleFormat(
            XlsbStyleRecords.ReadUInt16(data, 2),
            XlsbStyleRecords.ReadUInt16(data, 4),
            XlsbStyleRecords.ReadUInt16(data, 6),
            XlsbStyleRecords.ReadUInt16(data, 8),
            isStyleFormat || parent == XlsbStyleRecords.NoParent ? null : parent,
            own);
    }

    // BrtStyle: ixf (4), grbitObj1 (2: fBuiltIn bit 0), iStyBuiltIn (1), iLevel (1), stName.
    private static NamedStyle ReadNamedStyle(List<Biff12Record> records, Group group)
    {
        Biff12Record style = records[group.Item];
        style.RequireLength(8);
        bool builtIn = (style.Data[4] & 0x01) != 0;
        string name = style.ReadWideString(8, out _) ?? throw new InvalidDataException("A cell style has no name.");
        return new NamedStyle(
            name,
            (int)style.ReadUInt32(0),
            builtIn ? style.Data[6] : null,
            Convert.ToHexString(style.Data, 4, style.Data.Length - 4) + Attached(records, group));
    }

    private static string Attached(List<Biff12Record> records, Group group) =>
        string.Concat(Enumerable.Range(group.Start, group.End - group.Start)
            .Where(index => index != group.Item)
            .Select(index => $"|{records[index].Type}:{Convert.ToHexString(records[index].Data)}"));

    /// <summary>A record range: the item record plus the records that belong to it.</summary>
    internal readonly record struct Group(int Start, int End, int Item);

    /// <summary>A begin/end section and its item groups, as record indices.</summary>
    internal sealed record Section(int Begin, int End, IReadOnlyList<Group> Groups)
    {
        public static Section Read(List<Biff12Record> records, int beginType, int endType, int itemType, bool itemLeadsGroup)
        {
            int begin = records.FindIndex(record => record.Type == beginType);
            if (begin < 0)
            {
                return new Section(-1, -1, []);
            }

            int end = records.FindIndex(begin + 1, record => record.Type == endType);
            if (end < 0)
            {
                throw new InvalidDataException($"Style section {beginType} has no end record.");
            }

            List<Group> groups = [];
            int start = begin + 1;
            int index = start;
            while (index < end)
            {
                int type = records[index].Type;
                if (type == itemType)
                {
                    if (itemLeadsGroup)
                    {
                        // A format owns the future-record groups that follow it.
                        int groupEnd = index + 1;
                        while (groupEnd < end && records[groupEnd].Type == XlsbStyleRecords.FutureBegin)
                        {
                            groupEnd = SkipWrapper(records, groupEnd, end, XlsbStyleRecords.FutureBegin, XlsbStyleRecords.FutureEnd);
                        }

                        groups.Add(new Group(index, groupEnd, index));
                        index = groupEnd;
                    }
                    else
                    {
                        groups.Add(new Group(start, index + 1, index));
                        index++;
                    }

                    start = index;
                    continue;
                }

                // Before a named style only alternate-content groups (Excel writes the style's uid
                // there) may appear; they move with the style.
                if (!itemLeadsGroup && type == XlsbStyleRecords.AlternateBegin)
                {
                    index = SkipWrapper(records, index, end, XlsbStyleRecords.AlternateBegin, XlsbStyleRecords.AlternateEnd);
                    continue;
                }

                throw new UnsupportedStyleContentException(
                    $"Style section {beginType} contains record {type}, which this tool does not rewrite.");
            }

            if (start != end)
            {
                throw new UnsupportedStyleContentException(
                    $"Style section {beginType} ends with records that belong to no entry.");
            }

            return new Section(begin, end, groups);
        }

        private static int SkipWrapper(List<Biff12Record> records, int index, int end, int beginType, int endType)
        {
            int depth = 0;
            for (; index < end; index++)
            {
                if (records[index].Type == beginType)
                {
                    depth++;
                }
                else if (records[index].Type == endType && --depth == 0)
                {
                    return index + 1;
                }
            }

            throw new InvalidDataException($"Record group {beginType} is not closed.");
        }
    }
}
