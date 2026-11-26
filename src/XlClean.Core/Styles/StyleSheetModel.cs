namespace XlClean.Core.Styles;

/// <summary>
/// One <c>xf</c> record from <c>cellStyleXfs</c> or <c>cellXfs</c>, split into the indices it
/// points through and everything else.
/// </summary>
/// <param name="NumberFormatId">The raw <c>numFmtId</c>, or null when absent.</param>
/// <param name="FontId">The raw <c>fontId</c>, or null when absent.</param>
/// <param name="FillId">The raw <c>fillId</c>, or null when absent.</param>
/// <param name="BorderId">The raw <c>borderId</c>, or null when absent.</param>
/// <param name="ParentId">
/// The raw <c>xfId</c> (the named style's format), or null when absent. Excel reads an absent
/// <c>xfId</c> as 0, the Normal style's format.
/// </param>
/// <param name="OwnContent">
/// The canonical form of the record without those five attributes: every other attribute and
/// every child (alignment, protection, extensions), with absent attributes left absent.
/// </param>
internal sealed record StyleFormat(
    int? NumberFormatId,
    int? FontId,
    int? FillId,
    int? BorderId,
    int? ParentId,
    string OwnContent);

/// <summary>One <c>cellStyle</c> entry: a named style.</summary>
/// <param name="FormatId">The <c>xfId</c> into <c>cellStyleXfs</c>.</param>
/// <param name="BuiltinId">The <c>builtinId</c>, or null for a custom style.</param>
/// <param name="Content">
/// The canonical entry without <c>xfId</c>: name, flags such as <c>hidden</c> and
/// <c>customBuiltin</c>, and extensions.
/// </param>
internal sealed record NamedStyle(string Name, int FormatId, int? BuiltinId, string Content)
{
    public bool IsBuiltin => BuiltinId is not null;
}

/// <summary>
/// The parts of a workbook's style sheet that the cleanup reads. Component tables are kept as
/// canonical content because the cleanup never rewrites them; it only compares through them.
/// </summary>
internal sealed record StyleSheetModel(
    IReadOnlyDictionary<int, string> CustomNumberFormats,
    IReadOnlyList<string> Fonts,
    IReadOnlyList<string> Fills,
    IReadOnlyList<string> Borders,
    IReadOnlyList<StyleFormat> CellStyleFormats,
    IReadOnlyList<StyleFormat> CellFormats,
    IReadOnlyList<NamedStyle> NamedStyles)
{
    /// <summary>A workbook without a styles part.</summary>
    public static StyleSheetModel Empty { get; } = new(
        new Dictionary<int, string>(),
        [],
        [],
        [],
        [],
        [],
        []);
}

/// <summary>What the rest of the workbook references in the style sheet.</summary>
/// <param name="CellFormatIndices">
/// Every <c>cellXfs</c> index a cell, row, or column refers to, including the implicit 0.
/// </param>
/// <param name="TableStyleNames">Cell style names that tables and table columns refer to.</param>
internal sealed record StyleUsage(
    IReadOnlySet<int> CellFormatIndices,
    IReadOnlySet<string> TableStyleNames);

/// <summary>
/// A workbook holds something that may reference style indices in a form the cleanup does not
/// rewrite.
/// </summary>
internal sealed class UnsupportedStyleContentException(string message)
    : WorkbookFailureException(WorkbookFailureKind.UnsupportedWorkbookContent, message);
