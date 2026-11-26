namespace XlClean.Core.Styles;

using System.Text;

/// <summary>
/// The renumbering a style cleanup applies. Indices not listed as kept are removed.
/// </summary>
/// <param name="CellFormatMap">
/// For each source <c>cellXfs</c> index, the output index, or null when nothing references it.
/// Merged formats map to the index of the format that represents them.
/// </param>
/// <param name="CellStyleFormatMap">
/// For each source <c>cellStyleXfs</c> index, the output index, or null when it is removed.
/// </param>
/// <param name="KeptNamedStyles">Source <c>cellStyles</c> indices that stay, in source order.</param>
/// <param name="MergedCellFormats">Used cell formats folded into an identical earlier one.</param>
internal sealed record StylePlan(
    IReadOnlyList<int?> CellFormatMap,
    IReadOnlyList<int?> CellStyleFormatMap,
    IReadOnlyList<int> KeptNamedStyles,
    int MergedCellFormats)
{
    /// <summary>Gets the number of cell formats the output keeps.</summary>
    public int KeptCellFormatCount => CellFormatMap.Where(index => index is not null).Distinct().Count();

    /// <summary>Gets a value indicating whether any cell, row, or column index changes.</summary>
    public bool RenumbersCellFormats => CellFormatMap.Where((target, index) => target != index).Any();
}

/// <summary>
/// Decides which styles a workbook can drop or merge without changing how any cell looks.
/// </summary>
/// <remarks>
/// The rules are in plans/style-cleanup.md ("What the cleanup does"). The planner is a pure
/// function of the model and usage; the verifier checks its result independently.
/// </remarks>
internal static class StylePlanner
{
    /// <summary>Number format ids below this are built in and may render differently per locale.</summary>
    private const int FirstCustomNumberFormatId = 164;

    public static StylePlan Plan(StyleSheetModel model, StyleUsage usage)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(usage);

        StyleKeys keys = new(model);

        // Used cell formats, merged by resolved content. The first of each group represents it,
        // so output order follows source order and index 0 stays 0.
        int?[] cellFormatMap = new int?[model.CellFormats.Count];
        Dictionary<string, int> representatives = new(StringComparer.Ordinal);
        List<int> keptCellFormats = [];
        int merged = 0;
        for (int index = 0; index < model.CellFormats.Count; index++)
        {
            if (index != 0 && !usage.CellFormatIndices.Contains(index))
            {
                continue;
            }

            string key = keys.CellFormatKey(index);
            if (representatives.TryGetValue(key, out int representative))
            {
                cellFormatMap[index] = representative;
                merged++;
            }
            else
            {
                representative = keptCellFormats.Count;
                representatives.Add(key, representative);
                keptCellFormats.Add(index);
                cellFormatMap[index] = representative;
            }
        }

        // Named styles and their formats. A style format is reachable from a kept cell format,
        // from index 0, or from a kept named style; a named style is kept when it is built in,
        // when a table names it, or when its format is reachable from a kept cell format.
        HashSet<int> referencedByCells = [0];
        foreach (int index in keptCellFormats)
        {
            referencedByCells.Add(RequireStyleFormat(model, model.CellFormats[index].ParentId ?? 0, $"Cell format {index}"));
        }

        List<int> keptNamedStyles = [];
        HashSet<int> keptStyleFormats = [.. referencedByCells];
        for (int index = 0; index < model.NamedStyles.Count; index++)
        {
            NamedStyle style = model.NamedStyles[index];
            int format = RequireStyleFormat(model, style.FormatId, $"Cell style '{style.Name}'");
            if (style.IsBuiltin || usage.TableStyleNames.Contains(style.Name) || referencedByCells.Contains(format))
            {
                keptNamedStyles.Add(index);
                keptStyleFormats.Add(format);
            }
        }

        int?[] styleFormatMap = new int?[model.CellStyleFormats.Count];
        int next = 0;
        for (int index = 0; index < model.CellStyleFormats.Count; index++)
        {
            if (keptStyleFormats.Contains(index))
            {
                styleFormatMap[index] = next++;
            }
        }

        return new StylePlan(cellFormatMap, styleFormatMap, keptNamedStyles, merged);
    }

    private static int RequireStyleFormat(StyleSheetModel model, int index, string owner) =>
        index < model.CellStyleFormats.Count
            ? index
            : throw new InvalidDataException(
                $"{owner} refers to cell style format {index}, but the workbook defines only {model.CellStyleFormats.Count}.");

    /// <summary>Equality keys for formats, resolved through the component tables.</summary>
    /// <remarks>
    /// Every segment is length-prefixed, so text inside a name, format code, or attribute value
    /// can never imitate a segment boundary.
    /// </remarks>
    private sealed class StyleKeys(StyleSheetModel model)
    {
        private readonly Dictionary<int, string> styleFormatKeys = [];

        private readonly ILookup<int, NamedStyle> stylesByFormat = model.NamedStyles.ToLookup(style => style.FormatId);

        public string CellFormatKey(int index)
        {
            StyleFormat format = model.CellFormats[index];
            StringBuilder key = new();
            AppendContent(key, format, $"Cell format {index}");
            Append(key, "parentPresent", format.ParentId is null ? "no" : "yes");
            Append(key, "parent", StyleFormatKey(format.ParentId ?? 0, $"Cell format {index}"));
            return key.ToString();
        }

        // A named style's identity is its entries as well as its content: two cells that look
        // alike but report different style names, or flags, must not be merged.
        private string StyleFormatKey(int index, string owner)
        {
            if (styleFormatKeys.TryGetValue(index, out string? cached))
            {
                return cached;
            }

            StyleFormat format = model.CellStyleFormats[RequireStyleFormat(model, index, owner)];
            StringBuilder key = new();
            AppendContent(key, format, $"Cell style format {index}");
            Append(key, "parentAttribute", format.ParentId?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "absent");
            foreach (string style in stylesByFormat[index].Select(style => style.Content).Order(StringComparer.Ordinal))
            {
                Append(key, "style", style);
            }

            string result = key.ToString();
            styleFormatKeys.Add(index, result);
            return result;
        }

        private static void Append(StringBuilder key, string name, string value) =>
            key.Append(name).Append('#').Append(value.Length).Append(':').Append(value);

        private void AppendContent(StringBuilder key, StyleFormat format, string owner)
        {
            Append(key, "own", format.OwnContent);
            Append(key, "numFmt", NumberFormatKey(format.NumberFormatId, owner));
            Append(key, "font", Component(model.Fonts, format.FontId, "font", owner));
            Append(key, "fill", Component(model.Fills, format.FillId, "fill", owner));
            Append(key, "border", Component(model.Borders, format.BorderId, "border", owner));
        }

        // Built-in ids render per locale and have no entry, so they compare by id. A custom
        // format compares by code, and never equals a built-in id even when the text matches.
        private string NumberFormatKey(int? id, string owner)
        {
            if (id is not { } value)
            {
                return "absent";
            }

            bool defined = model.CustomNumberFormats.TryGetValue(value, out string? code);
            if (value < FirstCustomNumberFormatId)
            {
                return defined ? $"builtin:{value}:{code}" : $"builtin:{value}";
            }

            return defined
                ? $"custom:{code}"
                : throw new InvalidDataException($"{owner} uses number format {value}, which the workbook does not define.");
        }

        private static string Component(IReadOnlyList<string> table, int? id, string kind, string owner)
        {
            if (id is not { } value)
            {
                return "absent";
            }

            return value < table.Count
                ? table[value]
                : throw new InvalidDataException(
                    $"{owner} uses {kind} {value}, but the workbook defines only {table.Count}.");
        }
    }
}
