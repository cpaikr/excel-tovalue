namespace XlClean.Core.Styles;

using System.Globalization;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

/// <summary>
/// The cleaned workbook differs from its source in a way the cleanup must not cause.
/// </summary>
internal sealed class StyleVerificationException(string message)
    : WorkbookFailureException(
        WorkbookFailureKind.Unexpected,
        $"The cleaned workbook failed verification, so nothing was published: {message}");

/// <summary>
/// Proves that a style-cleaned .xlsx or .xlsm package looks the same as its source.
/// </summary>
/// <remarks>
/// <para>
/// This check is deliberately independent of <see cref="StylePlanner"/>. It does not use the
/// planner's equality keys or model; it re-reads both packages and compares the formatting each
/// cell, row, and column resolves to, element by element. A property the planner forgets would
/// make it merge wrongly, and this comparison would then see the difference.
/// </para>
/// <para>
/// What may differ: the <c>cellStyleXfs</c>, <c>cellXfs</c>, and <c>cellStyles</c> sections of
/// the style sheet, and the style index attributes of cells, rows, and columns. Everything else
/// in the package must be identical.
/// </para>
/// <para>
/// Limit of that independence: which parts count as worksheets and tables, and which parts make
/// a workbook unsupported, come from <see cref="OpenXmlStyleReader"/>. Every other part must be
/// byte-identical, so a part type that held cell style indices and that neither knows about would
/// keep its old indices unnoticed. No such part type exists in transitional SpreadsheetML today.
/// </para>
/// </remarks>
internal static class OpenXmlStyleVerifier
{
    private const int FirstCustomNumberFormatId = 164;

    private static readonly XNamespace Main = OpenXmlStyleReader.MainNamespace;

    private static readonly HashSet<string> RewrittenSections = new(StringComparer.Ordinal)
    {
        "cellStyleXfs",
        "cellXfs",
        "cellStyles",
    };

    private static readonly XName[] FormatIndexAttributes =
        ["numFmtId", "fontId", "fillId", "borderId", "xfId"];

    private const string MarkupCompatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    // Comments and processing instructions carry no workbook content, and the SDK's writer drops
    // them, so comparing them would only fail correct rewrites.
    private static readonly XmlReaderSettings ReaderSettings = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreWhitespace = true,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
    };

    public static void Verify(string sourcePath, string outputPath, CancellationToken cancellationToken)
    {
        using (Package source = Package.Open(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (Package output = Package.Open(outputPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            PackagePart sourceWorkbook = OpenXmlStyleReader.GetWorkbookPart(source);
            PackagePart? sourceStyles =
                OpenXmlStyleReader.GetRelatedParts(sourceWorkbook, OpenXmlStyleReader.StylesRelationship).SingleOrDefault();
            HashSet<Uri> worksheets = [.. OpenXmlStyleReader
                .GetRelatedParts(sourceWorkbook, OpenXmlStyleReader.WorksheetRelationship)
                .Select(part => part.Uri)];

            ComparePackageStructure(source, output, sourceStyles?.Uri, worksheets);

            StyleSheet sourceSheet = StyleSheet.Load(sourceStyles);
            StyleSheet outputSheet = StyleSheet.Load(sourceStyles is null ? null : output.GetPart(sourceStyles.Uri));
            CompareUnrewrittenStyleSections(sourceSheet, outputSheet);
            CompareNamedStyles(sourceSheet, outputSheet, sourceWorkbook);

            outputSheet.CheckEveryFormatResolves();
            FormatComparer formats = new(sourceSheet, outputSheet);

            // The default format and the default named style apply even where nothing names them.
            if (!formats.Equivalent(0, 0))
            {
                throw new StyleVerificationException("The default cell format changed.");
            }

            if (!XNode.DeepEquals(sourceSheet.ResolveDefaultStyleFormat(), outputSheet.ResolveDefaultStyleFormat()))
            {
                throw new StyleVerificationException("The default cell style format changed.");
            }

            foreach (Uri worksheet in worksheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                CompareWorksheet(source.GetPart(worksheet), output.GetPart(worksheet), formats, cancellationToken);
            }
        }

        CompareSchemaFindings(sourcePath, outputPath);
    }

    private static void ComparePackageStructure(Package source, Package output, Uri? styles, HashSet<Uri> worksheets)
    {
        CompareRelationships("the package", source.GetRelationships(), output.GetRelationships());

        Dictionary<Uri, PackagePart> outputParts = output.GetParts()
            .Where(part => !PackUriHelper.IsRelationshipPartUri(part.Uri))
            .ToDictionary(part => part.Uri);
        foreach (PackagePart sourcePart in source.GetParts())
        {
            if (PackUriHelper.IsRelationshipPartUri(sourcePart.Uri))
            {
                continue;
            }

            if (!outputParts.Remove(sourcePart.Uri, out PackagePart? outputPart))
            {
                throw new StyleVerificationException($"Part '{sourcePart.Uri}' is missing.");
            }

            if (!string.Equals(sourcePart.ContentType, outputPart.ContentType, StringComparison.Ordinal))
            {
                throw new StyleVerificationException($"Part '{sourcePart.Uri}' changed content type.");
            }

            CompareRelationships($"part '{sourcePart.Uri}'", sourcePart.GetRelationships(), outputPart.GetRelationships());
            if (sourcePart.Uri != styles && !worksheets.Contains(sourcePart.Uri) && !SameBytes(sourcePart, outputPart))
            {
                throw new StyleVerificationException($"Part '{sourcePart.Uri}' changed.");
            }
        }

        if (outputParts.Count > 0)
        {
            throw new StyleVerificationException($"Part '{outputParts.Keys.First()}' was added.");
        }
    }

    private static void CompareRelationships(
        string owner,
        PackageRelationshipCollection source,
        PackageRelationshipCollection output)
    {
        static string Describe(PackageRelationship relationship) =>
            $"{relationship.Id}|{relationship.RelationshipType}|{relationship.TargetUri.OriginalString}|{relationship.TargetMode}";

        string[] expected = [.. source.Select(Describe).Order(StringComparer.Ordinal)];
        string[] actual = [.. output.Select(Describe).Order(StringComparer.Ordinal)];
        if (!expected.SequenceEqual(actual, StringComparer.Ordinal))
        {
            throw new StyleVerificationException($"The relationships of {owner} changed.");
        }
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

    private static void CompareUnrewrittenStyleSections(StyleSheet source, StyleSheet output)
    {
        if (source.Root is null || output.Root is null)
        {
            if (source.Root is not null || output.Root is not null)
            {
                throw new StyleVerificationException("The styles part appeared or disappeared.");
            }

            return;
        }

        if (!XNode.DeepEquals(Normalize(new XElement(source.Root.Name, source.Root.Attributes())), Normalize(new XElement(output.Root.Name, output.Root.Attributes()))))
        {
            throw new StyleVerificationException("The style sheet's root attributes changed.");
        }

        XElement[] sourceChildren = [.. source.Root.Elements()];
        XElement[] outputChildren = [.. output.Root.Elements()];
        if (!sourceChildren.Select(child => child.Name).SequenceEqual(outputChildren.Select(child => child.Name)))
        {
            throw new StyleVerificationException("The style sheet's sections changed.");
        }

        for (int index = 0; index < sourceChildren.Length; index++)
        {
            XElement outputChild = outputChildren[index];
            if (RewrittenSections.Contains(outputChild.Name.LocalName) && outputChild.Name.Namespace == Main)
            {
                if ((string?)outputChild.Attribute("count") is { } count &&
                    count != outputChild.Elements().Count().ToString(CultureInfo.InvariantCulture))
                {
                    throw new StyleVerificationException(
                        $"The '{outputChild.Name.LocalName}' count does not match its entries.");
                }

                continue;
            }

            if (!XNode.DeepEquals(Normalize(sourceChildren[index]), Normalize(outputChild)))
            {
                throw new StyleVerificationException($"The style sheet section '{outputChild.Name.LocalName}' changed.");
            }
        }
    }

    // Every output named style must be a source named style unchanged, pointing at the same
    // resolved format; built-in styles and styles that tables name must all survive.
    private static void CompareNamedStyles(StyleSheet source, StyleSheet output, PackagePart workbook)
    {
        Dictionary<string, int> remaining = new(StringComparer.Ordinal);
        foreach (XElement style in source.NamedStyles)
        {
            string described = DescribeNamedStyle(source, style);
            remaining[described] = remaining.GetValueOrDefault(described) + 1;
        }

        HashSet<string> outputNames = new(StringComparer.Ordinal);
        foreach (XElement style in output.NamedStyles)
        {
            string described = DescribeNamedStyle(output, style);
            if (remaining.GetValueOrDefault(described) == 0)
            {
                throw new StyleVerificationException($"Cell style '{(string?)style.Attribute("name")}' is not a source style.");
            }

            remaining[described]--;
            outputNames.Add((string?)style.Attribute("name") ?? string.Empty);
        }

        foreach (XElement style in source.NamedStyles.Where(style => style.Attribute("builtinId") is not null))
        {
            if (!outputNames.Contains((string?)style.Attribute("name") ?? string.Empty))
            {
                throw new StyleVerificationException($"Built-in cell style '{(string?)style.Attribute("name")}' was removed.");
            }
        }

        HashSet<string> sourceNames = [.. source.NamedStyles.Select(style => (string?)style.Attribute("name") ?? string.Empty)];
        foreach (string name in TableStyleNames(workbook))
        {
            if (sourceNames.Contains(name) && !outputNames.Contains(name))
            {
                throw new StyleVerificationException($"Cell style '{name}', which a table uses, was removed.");
            }
        }
    }

    private static string DescribeNamedStyle(StyleSheet sheet, XElement style)
    {
        int formatId = ParseIndex((string?)style.Attribute("xfId") ?? "0", "cellStyle");
        return new XElement(
            "namedStyle",
            Normalize(style, "xfId"),
            sheet.ResolveStyleFormat(formatId, includeNames: false)).ToString(SaveOptions.DisableFormatting);
    }

    private static IEnumerable<string> TableStyleNames(PackagePart workbook)
    {
        foreach (PackagePart worksheet in OpenXmlStyleReader.GetRelatedParts(workbook, OpenXmlStyleReader.WorksheetRelationship))
        {
            foreach (PackagePart table in OpenXmlStyleReader.GetRelatedParts(worksheet, OpenXmlStyleReader.TableRelationship))
            {
                XElement root = LoadRoot(table) ?? throw new StyleVerificationException($"Table part '{table.Uri}' is empty.");
                foreach (XElement element in root.DescendantsAndSelf())
                {
                    foreach (string attribute in new[] { "headerRowCellStyle", "dataCellStyle", "totalsRowCellStyle" })
                    {
                        if ((string?)element.Attribute(attribute) is { } name)
                        {
                            yield return name;
                        }
                    }
                }
            }
        }
    }

    private static void CompareWorksheet(
        PackagePart source,
        PackagePart output,
        FormatComparer formats,
        CancellationToken cancellationToken)
    {
        using Stream sourceStream = source.GetStream(FileMode.Open, FileAccess.Read);
        using Stream outputStream = output.GetStream(FileMode.Open, FileAccess.Read);
        using XmlReader sourceReader = XmlReader.Create(sourceStream, ReaderSettings);
        using XmlReader outputReader = XmlReader.Create(outputStream, ReaderSettings);
        using IEnumerator<XmlToken> expected = Tokens(sourceReader).GetEnumerator();
        using IEnumerator<XmlToken> actual = Tokens(outputReader).GetEnumerator();

        long tokens = 0;
        while (true)
        {
            bool hasExpected = expected.MoveNext();
            bool hasActual = actual.MoveNext();
            if (hasExpected != hasActual)
            {
                throw new StyleVerificationException($"Worksheet '{source.Uri}' has a different structure.");
            }

            if (!hasExpected)
            {
                return;
            }

            if (++tokens % 4096 == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            XmlToken sourceToken = expected.Current;
            XmlToken outputToken = actual.Current;
            if (!sourceToken.SameShape(outputToken))
            {
                throw new StyleVerificationException(
                    $"Worksheet '{source.Uri}' changed at {sourceToken.Describe()}.");
            }

            if (sourceToken.Kind == XmlTokenKind.Start && sourceToken.StyleAttribute is { } styleAttribute)
            {
                int sourceIndex = ParseIndex(sourceToken.StyleValue ?? "0", styleAttribute);
                int outputIndex = ParseIndex(outputToken.StyleValue ?? "0", styleAttribute);
                if (!formats.Equivalent(sourceIndex, outputIndex))
                {
                    throw new StyleVerificationException(
                        $"Worksheet '{source.Uri}' {sourceToken.Describe()} looks different: cell format {sourceIndex} became {outputIndex}.");
                }
            }
        }
    }

    private static IEnumerable<XmlToken> Tokens(XmlReader reader)
    {
        while (reader.Read())
        {
            switch (reader.NodeType)
            {
                case XmlNodeType.Element:
                    string? styleAttribute = reader.NamespaceURI == OpenXmlStyleReader.MainNamespace
                        ? reader.LocalName switch
                        {
                            "c" or "row" => "s",
                            "col" => "style",
                            _ => null,
                        }
                        : null;
                    string? styleValue = styleAttribute is null ? null : reader.GetAttribute(styleAttribute);
                    string? reference = reader.GetAttribute("r");
                    List<string> attributes = [];
                    bool empty = reader.IsEmptyElement;
                    if (reader.MoveToFirstAttribute())
                    {
                        do
                        {
                            bool declaration = reader.Prefix == "xmlns" || (reader.Prefix.Length == 0 && reader.LocalName == "xmlns");
                            bool style = styleAttribute is not null && reader.NamespaceURI.Length == 0 && reader.LocalName == styleAttribute;
                            if (!declaration && !style)
                            {
                                attributes.Add($"{{{reader.NamespaceURI}}}{reader.LocalName}={AttributeValue(reader)}");
                            }
                        }
                        while (reader.MoveToNextAttribute());

                        reader.MoveToElement();
                    }

                    attributes.Sort(StringComparer.Ordinal);
                    string name = $"{{{reader.NamespaceURI}}}{reader.LocalName}";
                    yield return new XmlToken(XmlTokenKind.Start, name, string.Join('\u0001', attributes), styleAttribute, styleValue, reference);
                    if (empty)
                    {
                        yield return new XmlToken(XmlTokenKind.End, name, string.Empty, null, null, null);
                    }

                    break;

                case XmlNodeType.EndElement:
                    yield return new XmlToken(
                        XmlTokenKind.End,
                        $"{{{reader.NamespaceURI}}}{reader.LocalName}",
                        string.Empty,
                        null,
                        null,
                        null);
                    break;

                case XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace:
                    yield return new XmlToken(XmlTokenKind.Text, string.Empty, reader.Value, null, null, null);
                    break;

                case XmlNodeType.Comment or XmlNodeType.ProcessingInstruction:
                    yield return new XmlToken(XmlTokenKind.Other, reader.Name, reader.Value, null, null, null);
                    break;
            }
        }
    }

    // Markup-compatibility attributes name namespaces by prefix. Comparing the prefixes alone would
    // miss a rewrite that drops or rebinds a declaration, which Excel reports as damage, so the
    // prefixes are compared by the namespaces they resolve to.
    private static string AttributeValue(XmlReader reader)
    {
        if (reader.NamespaceURI != MarkupCompatibilityNamespace ||
            reader.LocalName is not ("Ignorable" or "ProcessContent" or "MustUnderstand"))
        {
            return reader.Value;
        }

        IEnumerable<string> resolved = reader.Value
            .Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(item =>
            {
                int colon = item.IndexOf(':', StringComparison.Ordinal);
                string prefix = colon < 0 ? item : item[..colon];
                string rest = colon < 0 ? string.Empty : item[colon..];
                string namespaceName = reader.LookupNamespace(prefix) ??
                    throw new StyleVerificationException($"Prefix '{prefix}' in '{reader.LocalName}' is not declared.");
                return $"{{{namespaceName}}}{rest}";
            });
        return string.Join(' ', resolved);
    }

    private static void CompareSchemaFindings(string sourcePath, string outputPath)
    {
        Dictionary<string, int> expected = SchemaFindings(sourcePath);
        foreach ((string finding, int count) in SchemaFindings(outputPath))
        {
            if (count > expected.GetValueOrDefault(finding))
            {
                throw new StyleVerificationException($"The style sheet has a new schema finding: {finding}.");
            }
        }
    }

    // Positional paths shift when sections are compacted, so findings are compared by what they
    // are and where, not by path.
    private static Dictionary<string, int> SchemaFindings(string path)
    {
        Dictionary<string, int> findings = new(StringComparer.Ordinal);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(path, false, OpenXmlWorkbook.CreateOpenSettings());
        if (document.WorkbookPart?.WorkbookStylesPart is not { } stylesPart)
        {
            return findings;
        }

        OpenXmlValidator validator = new(FileFormatVersions.Microsoft365) { MaxNumberOfErrors = 0 };
        foreach (ValidationErrorInfo error in validator.Validate(stylesPart))
        {
            string key = $"{error.Id} on '{error.Node?.LocalName}'";
            findings[key] = findings.GetValueOrDefault(key) + 1;
        }

        return findings;
    }

    private static XElement? LoadRoot(PackagePart part)
    {
        using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using XmlReader reader = XmlReader.Create(stream, ReaderSettings);
        return XDocument.Load(reader).Root;
    }

    private static XElement Normalize(XElement element, params XName[] excluded)
    {
        XElement normalized = new(
            element.Name,
            element.Attributes()
                .Where(attribute => !attribute.IsNamespaceDeclaration && !excluded.Contains(attribute.Name))
                .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal)
                .Select(attribute => new XAttribute(attribute.Name, attribute.Value)));
        foreach (XNode node in element.Nodes())
        {
            if (node is XElement child)
            {
                normalized.Add(Normalize(child));
            }
            else if (node is XText text)
            {
                normalized.Add(new XText(text.Value));
            }
        }

        return normalized;
    }

    private static int ParseIndex(string value, string owner) =>
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int index)
            ? index
            : throw new StyleVerificationException($"A '{owner}' style index '{value}' is not a number.");

    private enum XmlTokenKind
    {
        Start,
        End,
        Text,
        Other,
    }

    private sealed record XmlToken(
        XmlTokenKind Kind,
        string Name,
        string Content,
        string? StyleAttribute,
        string? StyleValue,
        string? Reference)
    {
        public bool SameShape(XmlToken other) =>
            Kind == other.Kind &&
            string.Equals(Name, other.Name, StringComparison.Ordinal) &&
            string.Equals(Content, other.Content, StringComparison.Ordinal) &&
            string.Equals(StyleAttribute, other.StyleAttribute, StringComparison.Ordinal);

        public string Describe() =>
            Reference is null ? $"'{Name}'" : $"'{Name}' {Reference}";
    }

    /// <summary>One side's style sheet, resolving formats to self-contained elements.</summary>
    private sealed class StyleSheet
    {
        private readonly XElement[] cellFormats;
        private readonly XElement[] styleFormats;
        private readonly Dictionary<int, XElement> numberFormats;
        private readonly XElement[] fonts;
        private readonly XElement[] fills;
        private readonly XElement[] borders;

        private readonly ILookup<string, XElement> stylesByFormat;

        private StyleSheet(XElement? root)
        {
            Root = root;
            cellFormats = Items(root, "cellXfs", "xf");
            styleFormats = Items(root, "cellStyleXfs", "xf");
            NamedStyles = Items(root, "cellStyles", "cellStyle");
            stylesByFormat = NamedStyles.ToLookup(style => (string?)style.Attribute("xfId") ?? "0", StringComparer.Ordinal);
            numberFormats = Items(root, "numFmts", "numFmt")
                .ToDictionary(element => ParseIndex((string?)element.Attribute("numFmtId") ?? string.Empty, "numFmt"));
            fonts = Items(root, "fonts", "font");
            fills = Items(root, "fills", "fill");
            borders = Items(root, "borders", "border");
        }

        public XElement? Root { get; }

        public XElement[] NamedStyles { get; }

        public static StyleSheet Load(PackagePart? stylesPart) =>
            new(stylesPart is null ? null : LoadRoot(stylesPart));

        public XElement ResolveCellFormat(int index)
        {
            // Without a style sheet only the implicit default exists.
            if (cellFormats.Length == 0)
            {
                return index == 0
                    ? new XElement("implicitDefaultFormat")
                    : throw new StyleVerificationException($"Cell format {index} does not exist.");
            }

            XElement xf = index < cellFormats.Length
                ? cellFormats[index]
                : throw new StyleVerificationException($"Cell format {index} does not exist.");
            XElement resolved = Resolve(xf);

            // Excel reads an absent xfId as 0, so the parent always resolves; whether the attribute
            // was written is compared too.
            string? parent = (string?)xf.Attribute("xfId");
            resolved.Add(new XElement(
                parent is null ? "implicitParent" : "parent",
                ResolveStyleFormatIfAny(ParseIndex(parent ?? "0", "xf"))));
            return resolved;
        }

        public XElement ResolveDefaultStyleFormat() => ResolveStyleFormatIfAny(0);

        /// <summary>Resolves every output entry so a dangling index cannot hide in an unused one.</summary>
        public void CheckEveryFormatResolves()
        {
            for (int index = 0; index < cellFormats.Length; index++)
            {
                _ = ResolveCellFormat(index);
            }

            for (int index = 0; index < styleFormats.Length; index++)
            {
                _ = ResolveStyleFormat(index, includeNames: false);
            }
        }

        private XElement ResolveStyleFormatIfAny(int index) =>
            styleFormats.Length == 0 && index == 0
                ? new XElement("noStyleFormats")
                : ResolveStyleFormat(index, includeNames: true);

        public XElement ResolveStyleFormat(int index, bool includeNames)
        {
            XElement xf = index < styleFormats.Length
                ? styleFormats[index]
                : throw new StyleVerificationException($"Cell style format {index} does not exist.");
            XElement resolved = Resolve(xf);
            if (includeNames)
            {
                string id = index.ToString(CultureInfo.InvariantCulture);
                foreach (XElement style in stylesByFormat[id]
                    .Select(style => Normalize(style, "xfId"))
                    .OrderBy(style => style.ToString(SaveOptions.DisableFormatting), StringComparer.Ordinal))
                {
                    resolved.Add(style);
                }
            }

            return resolved;
        }

        private static XElement[] Items(XElement? root, string section, string item) =>
            root?.Element(Main + section)?.Elements(Main + item).ToArray() ?? [];

        private XElement Resolve(XElement xf) =>
            new(
                "format",
                Normalize(xf, FormatIndexAttributes),
                NumberFormat((string?)xf.Attribute("numFmtId")),
                Component("font", fonts, (string?)xf.Attribute("fontId")),
                Component("fill", fills, (string?)xf.Attribute("fillId")),
                Component("border", borders, (string?)xf.Attribute("borderId")));

        private XElement NumberFormat(string? id)
        {
            if (id is null)
            {
                return new XElement("noNumberFormat");
            }

            int value = ParseIndex(id, "xf");
            if (numberFormats.TryGetValue(value, out XElement? definition))
            {
                // A custom format is its code; a redefined built-in id stays tied to its id.
                return value < FirstCustomNumberFormatId
                    ? new XElement("builtinNumberFormat", new XAttribute("id", value), Normalize(definition))
                    : new XElement("customNumberFormat", Normalize(definition, "numFmtId"));
            }

            return value < FirstCustomNumberFormatId
                ? new XElement("builtinNumberFormat", new XAttribute("id", value))
                : throw new StyleVerificationException($"Number format {value} is not defined.");
        }

        private static XElement Component(string kind, XElement[] table, string? id)
        {
            if (id is null)
            {
                return new XElement($"no-{kind}");
            }

            int value = ParseIndex(id, "xf");
            return value < table.Length
                ? new XElement(kind, Normalize(table[value]))
                : throw new StyleVerificationException($"The {kind} {value} does not exist.");
        }
    }

    /// <summary>Compares source and output cell formats by what they resolve to, caching results.</summary>
    private sealed class FormatComparer(StyleSheet source, StyleSheet output)
    {
        private readonly Dictionary<(int Source, int Output), bool> results = [];
        private readonly Dictionary<int, XElement> sourceFormats = [];
        private readonly Dictionary<int, XElement> outputFormats = [];

        public bool Equivalent(int sourceIndex, int outputIndex)
        {
            if (results.TryGetValue((sourceIndex, outputIndex), out bool known))
            {
                return known;
            }

            bool result = XNode.DeepEquals(
                Resolve(source, sourceFormats, sourceIndex),
                Resolve(output, outputFormats, outputIndex));
            results.Add((sourceIndex, outputIndex), result);
            return result;
        }

        private static XElement Resolve(StyleSheet sheet, Dictionary<int, XElement> cache, int index)
        {
            if (!cache.TryGetValue(index, out XElement? resolved))
            {
                resolved = sheet.ResolveCellFormat(index);
                cache.Add(index, resolved);
            }

            return resolved;
        }
    }
}
