namespace XlClean.Core.Styles;

using System.Globalization;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;

/// <summary>
/// Reads the style sheet of an .xlsx or .xlsm package and every reference to it.
/// </summary>
/// <remarks>
/// Parts are read as XML straight from the package so that the reader sees exactly what the
/// file says, attribute presence included. Anything that may hold a style index in a form the
/// cleanup does not rewrite fails with <see cref="UnsupportedStyleContentException"/>.
/// </remarks>
internal static class OpenXmlStyleReader
{
    public const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public const string StylesRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

    public const string WorksheetRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    public const string TableRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

    private const string OfficeDocumentRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    private const string StrictOfficeDocumentRelationship =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";

    private static readonly XNamespace Main = MainNamespace;

    /// <summary>Workbook relationships whose parts carry cell style indices the cleanup does not rewrite.</summary>
    private static readonly Dictionary<string, string> UnsupportedWorkbookRelationships = new(StringComparer.Ordinal)
    {
        ["http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet"] = "an Excel 4.0 macro sheet",
        ["http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet"] =
            "an international macro sheet",
        ["http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders"] =
            "shared-workbook revision history",
    };

    private static readonly XName[] IndexAttributes =
        ["numFmtId", "fontId", "fillId", "borderId", "xfId"];

    private static readonly XmlReaderSettings ReaderSettings = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
    };

    public static (StyleSheetModel Model, StyleUsage Usage) Read(string workbookPath, CancellationToken cancellationToken)
    {
        using Package package = Package.Open(workbookPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        PackagePart workbookPart = GetWorkbookPart(package);
        RejectUnsupportedParts(workbookPart);

        PackagePart? stylesPart = GetRelatedParts(workbookPart, StylesRelationship).SingleOrDefault();
        StyleSheetModel model = stylesPart is null ? StyleSheetModel.Empty : ReadStyleSheet(stylesPart);

        HashSet<int> cellFormats = [0];
        HashSet<string> tableStyleNames = new(StringComparer.Ordinal);
        foreach (PackagePart worksheetPart in GetRelatedParts(workbookPart, WorksheetRelationship))
        {
            cancellationToken.ThrowIfCancellationRequested();
            CollectWorksheetReferences(worksheetPart, cellFormats, cancellationToken);
            foreach (PackagePart tablePart in GetRelatedParts(worksheetPart, TableRelationship))
            {
                CollectTableStyleNames(tablePart, tableStyleNames);
            }
        }

        // Index 0 is implicit even without a style sheet; anything else must exist.
        int highest = cellFormats.Max();
        if (highest > 0 && highest >= model.CellFormats.Count)
        {
            throw new InvalidDataException(
                $"A cell references cell format {highest}, but the workbook defines only {model.CellFormats.Count}.");
        }

        return (model, new StyleUsage(cellFormats, tableStyleNames));
    }

    /// <summary>Fails closed on workbook parts that carry cell style indices the cleanup does not rewrite.</summary>
    public static void RejectUnsupportedParts(PackagePart workbookPart)
    {
        foreach (PackageRelationship relationship in workbookPart.GetRelationships())
        {
            if (UnsupportedWorkbookRelationships.TryGetValue(relationship.RelationshipType, out string? description))
            {
                throw new UnsupportedStyleContentException(
                    $"The workbook contains {description}, whose cells reference styles in a form this tool does not rewrite.");
            }
        }
    }

    public static PackagePart GetWorkbookPart(Package package)
    {
        if (package.GetRelationshipsByType(StrictOfficeDocumentRelationship).Any())
        {
            throw new UnsupportedStyleContentException(
                "The workbook is saved as Strict Open XML, which this tool does not rewrite. Save it as a regular Excel workbook.");
        }

        PackageRelationship relationship = package
            .GetRelationshipsByType(OfficeDocumentRelationship)
            .FirstOrDefault() ??
            throw new InvalidDataException("The package has no workbook part.");
        Uri partUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
        return package.PartExists(partUri)
            ? package.GetPart(partUri)
            : throw new InvalidDataException("The package has no workbook part.");
    }

    public static IEnumerable<PackagePart> GetRelatedParts(PackagePart source, string relationshipType)
    {
        foreach (PackageRelationship relationship in source.GetRelationshipsByType(relationshipType))
        {
            if (relationship.TargetMode != TargetMode.Internal)
            {
                continue;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(source.Uri, relationship.TargetUri);
            yield return source.Package.PartExists(partUri)
                ? source.Package.GetPart(partUri)
                : throw new InvalidDataException($"Relationship '{relationship.Id}' has no target part.");
        }
    }

    private static StyleSheetModel ReadStyleSheet(PackagePart stylesPart)
    {
        XElement root = LoadRoot(stylesPart, "styleSheet");

        Dictionary<int, string> numberFormats = [];
        foreach (XElement numberFormat in Section(root, "numFmts", "numFmt"))
        {
            int id = RequiredInt(numberFormat, "numFmtId");
            string code = (string?)numberFormat.Attribute("formatCode") ??
                throw new InvalidDataException($"Number format {id} has no format code.");
            if (!numberFormats.TryAdd(id, code))
            {
                throw new InvalidDataException($"Number format {id} is defined more than once.");
            }
        }

        List<NamedStyle> namedStyles = [];
        foreach (XElement cellStyle in Section(root, "cellStyles", "cellStyle"))
        {
            namedStyles.Add(new NamedStyle(
                (string?)cellStyle.Attribute("name") ?? throw new InvalidDataException("A cell style has no name."),
                RequiredInt(cellStyle, "xfId"),
                OptionalInt(cellStyle, "builtinId"),
                CanonicalXml.Of(cellStyle, "xfId")));
        }

        return new StyleSheetModel(
            numberFormats,
            Section(root, "fonts", "font").Select(font => CanonicalXml.Of(font)).ToArray(),
            Section(root, "fills", "fill").Select(fill => CanonicalXml.Of(fill)).ToArray(),
            Section(root, "borders", "border").Select(border => CanonicalXml.Of(border)).ToArray(),
            Section(root, "cellStyleXfs", "xf").Select(ReadFormat).ToArray(),
            Section(root, "cellXfs", "xf").Select(ReadFormat).ToArray(),
            namedStyles);
    }

    private static StyleFormat ReadFormat(XElement xf) =>
        new(
            OptionalInt(xf, "numFmtId"),
            OptionalInt(xf, "fontId"),
            OptionalInt(xf, "fillId"),
            OptionalInt(xf, "borderId"),
            OptionalInt(xf, "xfId"),
            CanonicalXml.Of(xf, IndexAttributes));

    private static void CollectWorksheetReferences(
        PackagePart worksheetPart,
        HashSet<int> cellFormats,
        CancellationToken cancellationToken)
    {
        using Stream stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read);
        using XmlReader reader = XmlReader.Create(stream, ReaderSettings);
        reader.MoveToContent();
        RequireMainRoot(reader.NamespaceURI, reader.LocalName, "worksheet", worksheetPart);

        int elements = 0;
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != MainNamespace)
            {
                continue;
            }

            if (++elements % 4096 == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            string? index = reader.LocalName switch
            {
                "c" or "row" => reader.GetAttribute("s"),
                "col" => reader.GetAttribute("style"),
                _ => null,
            };
            if (index is not null)
            {
                cellFormats.Add(ParseIndex(index, reader.LocalName));
            }
        }
    }

    private static void CollectTableStyleNames(PackagePart tablePart, HashSet<string> names)
    {
        XElement table = LoadRoot(tablePart, "table");
        IEnumerable<XElement> columns = table.Element(Main + "tableColumns")?.Elements(Main + "tableColumn") ?? [];
        foreach (XElement element in columns.Prepend(table))
        {
            foreach (string attribute in new[] { "headerRowCellStyle", "dataCellStyle", "totalsRowCellStyle" })
            {
                if ((string?)element.Attribute(attribute) is { } name)
                {
                    names.Add(name);
                }
            }
        }
    }

    private static XElement LoadRoot(PackagePart part, string expectedName)
    {
        using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using XmlReader reader = XmlReader.Create(stream, ReaderSettings);
        XElement root = XDocument.Load(reader).Root ??
            throw new InvalidDataException($"Part '{part.Uri}' has no root element.");
        RequireMainRoot(root.Name.NamespaceName, root.Name.LocalName, expectedName, part);
        return root;
    }

    // Strict OOXML and other dialects use other namespaces. Reading them as empty would make
    // every style look unused, so they are refused rather than misread.
    private static void RequireMainRoot(string namespaceName, string localName, string expectedName, PackagePart part)
    {
        if (namespaceName != MainNamespace || localName != expectedName)
        {
            throw new UnsupportedStyleContentException(
                $"Part '{part.Uri}' is '{{{namespaceName}}}{localName}', not a transitional SpreadsheetML {expectedName}.");
        }
    }

    private static IEnumerable<XElement> Section(XElement root, string section, string item) =>
        root.Element(Main + section)?.Elements(Main + item) ?? [];

    private static int RequiredInt(XElement element, string attribute) =>
        OptionalInt(element, attribute) ??
        throw new InvalidDataException($"A '{element.Name.LocalName}' element has no '{attribute}'.");

    private static int? OptionalInt(XElement element, string attribute) =>
        (string?)element.Attribute(attribute) is { } value ? ParseIndex(value, element.Name.LocalName) : null;

    private static int ParseIndex(string value, string owner) =>
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int index)
            ? index
            : throw new InvalidDataException($"A '{owner}' element has the invalid style index '{value}'.");
}
