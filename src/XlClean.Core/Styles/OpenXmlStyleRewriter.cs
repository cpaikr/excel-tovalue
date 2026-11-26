namespace XlClean.Core.Styles;

using System.Globalization;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Xml.Linq;

/// <summary>
/// Applies a <see cref="StylePlan"/> to a staged .xlsx or .xlsm copy.
/// </summary>
/// <remarks>
/// Only the styles part and, when cell format indices change, the worksheets are written. Parts
/// are rewritten through System.IO.Packaging rather than the SDK's document model, so no other
/// part is ever loaded or re-serialized. Worksheets are copied node by node with only the three
/// style index attributes changed; everything else, including prefixes, namespace declarations,
/// and the XML declaration, is written back as read.
/// <para>
/// System.IO.Packaging opens a writable package in update mode, which holds each part it rewrites
/// in memory, so a single worksheet is limited to what a memory stream can hold (about 2 GB
/// uncompressed). Larger parts fail the file safely as unexpected.
/// </para>
/// </remarks>
internal static class OpenXmlStyleRewriter
{
    private const int BufferSize = 1024 * 1024;

    private static readonly XNamespace Main = OpenXmlStyleReader.MainNamespace;

    private static readonly XmlReaderSettings ReaderSettings = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
    };

    // Entitize writes tab, line feed, and carriage return inside attributes, and carriage return
    // inside text, as character references. Written raw, a reader would normalize them to spaces
    // and line feeds, so multi-line validation prompts or header text would change.
    private static readonly XmlWriterSettings WriterSettings = new()
    {
        Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        OmitXmlDeclaration = true,
        ConformanceLevel = ConformanceLevel.Auto,
        NewLineHandling = NewLineHandling.Entitize,
        CheckCharacters = false,
    };

    public static void Apply(string workbookPath, StylePlan plan, CancellationToken cancellationToken)
    {
        if (plan.IsIdentity)
        {
            return;
        }

        cancellationToken.ThrowIfCancellationRequested();
        using Package package = Package.Open(workbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        PackagePart workbookPart = OpenXmlStyleReader.GetWorkbookPart(package);
        PackagePart stylesPart = OpenXmlStyleReader
            .GetRelatedParts(workbookPart, OpenXmlStyleReader.StylesRelationship)
            .Single();
        RewriteStyleSheet(stylesPart, plan);

        if (plan.RenumbersCellFormats)
        {
            foreach (PackagePart worksheet in OpenXmlStyleReader.GetRelatedParts(
                workbookPart,
                OpenXmlStyleReader.WorksheetRelationship))
            {
                cancellationToken.ThrowIfCancellationRequested();
                RewriteWorksheet(worksheet, plan.CellFormatMap, workbookPath, cancellationToken);
            }
        }
    }

    private static void RewriteStyleSheet(PackagePart stylesPart, StylePlan plan)
    {
        XDocument document;
        using (Stream stream = stylesPart.GetStream(FileMode.Open, FileAccess.Read))
        using (XmlReader reader = XmlReader.Create(stream, ReaderSettings))
        {
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }

        XElement root = document.Root ?? throw new InvalidDataException("The styles part has no root element.");

        XElement[] cellFormats = Items(root, "cellXfs", "xf");
        XElement[] styleFormats = Items(root, "cellStyleXfs", "xf");
        XElement[] namedStyles = Items(root, "cellStyles", "cellStyle");

        // A format is written once, at the position of the first source format that maps to its
        // output index; merged and unused formats are dropped.
        HashSet<int> seenTargets = [];
        bool[] representative = [.. plan.CellFormatMap.Select(target => target is { } value && seenTargets.Add(value))];

        // Remove from the end so earlier indices stay valid while removing.
        for (int index = cellFormats.Length - 1; index >= 0; index--)
        {
            if (!representative[index])
            {
                RemoveWithLeadingWhitespace(cellFormats[index]);
            }
            else if (cellFormats[index].Attribute("xfId") is { } parent)
            {
                parent.Value = Target(plan.CellStyleFormatMap, parent.Value, "cell style format").ToString(CultureInfo.InvariantCulture);
            }
        }

        for (int index = styleFormats.Length - 1; index >= 0; index--)
        {
            if (plan.CellStyleFormatMap[index] is null)
            {
                RemoveWithLeadingWhitespace(styleFormats[index]);
            }
        }

        HashSet<int> keptStyles = [.. plan.KeptNamedStyles];
        for (int index = namedStyles.Length - 1; index >= 0; index--)
        {
            XElement style = namedStyles[index];
            if (!keptStyles.Contains(index))
            {
                RemoveWithLeadingWhitespace(style);
            }
            else if (style.Attribute("xfId") is { } parent)
            {
                parent.Value = Target(plan.CellStyleFormatMap, parent.Value, "cell style format").ToString(CultureInfo.InvariantCulture);
            }
        }

        UpdateCount(root, "cellXfs", "xf");
        UpdateCount(root, "cellStyleXfs", "xf");
        UpdateCount(root, "cellStyles", "cellStyle");

        using Stream output = stylesPart.GetStream(FileMode.Create, FileAccess.Write);
        if (document.Declaration is { } declaration)
        {
            // A loaded document drops the whitespace after its declaration; Excel writes a line break.
            WriteDeclaration(output, declaration.Version, declaration.Standalone, lineBreak: "\r\n");
        }

        using XmlWriter writer = XmlWriter.Create(output, WriterSettings);
        root.WriteTo(writer);
    }

    private static void RewriteWorksheet(
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
            using (XmlReader reader = XmlReader.Create(source, ReaderSettings))
            {
                bool hasNode = reader.Read();
                if (hasNode && reader.NodeType == XmlNodeType.XmlDeclaration)
                {
                    // The whitespace after the declaration is a node the copy writes back itself.
                    WriteDeclaration(
                        staging,
                        reader.GetAttribute("version"),
                        reader.GetAttribute("standalone"),
                        lineBreak: string.Empty);
                    hasNode = reader.Read();
                }

                using XmlWriter writer = XmlWriter.Create(staging, WriterSettings);
                if (hasNode)
                {
                    CopyWithRenumbering(reader, writer, cellFormatMap, cancellationToken);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            using FileStream transformed = new(stagingPath, FileMode.Open, FileAccess.Read, FileShare.Read, BufferSize);
            using Stream destination = worksheet.GetStream(FileMode.Create, FileAccess.Write);
            transformed.CopyTo(destination, BufferSize);
        }
        finally
        {
            WorkbookBatchPipeline.TryDeleteStagingFile(stagingPath);
        }
    }

    /// <summary>Copies from the reader's current node to the end of the document.</summary>
    private static void CopyWithRenumbering(
        XmlReader reader,
        XmlWriter writer,
        IReadOnlyList<int?> cellFormatMap,
        CancellationToken cancellationToken)
    {
        long nodes = 0;
        do
        {
            if (++nodes % 4096 == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            switch (reader.NodeType)
            {
                case XmlNodeType.Element:
                    bool empty = reader.IsEmptyElement;
                    string? styleAttribute = reader.NamespaceURI == OpenXmlStyleReader.MainNamespace
                        ? reader.LocalName switch
                        {
                            "c" or "row" => "s",
                            "col" => "style",
                            _ => null,
                        }
                        : null;
                    writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
                    if (reader.MoveToFirstAttribute())
                    {
                        do
                        {
                            string value = styleAttribute is not null &&
                                reader.NamespaceURI.Length == 0 &&
                                reader.LocalName == styleAttribute
                                    ? Target(cellFormatMap, reader.Value, "cell format").ToString(CultureInfo.InvariantCulture)
                                    : reader.Value;
                            writer.WriteAttributeString(reader.Prefix, reader.LocalName, reader.NamespaceURI, value);
                        }
                        while (reader.MoveToNextAttribute());

                        reader.MoveToElement();
                    }

                    if (empty)
                    {
                        writer.WriteEndElement();
                    }

                    break;

                case XmlNodeType.EndElement:
                    writer.WriteFullEndElement();
                    break;

                case XmlNodeType.Text:
                    writer.WriteString(reader.Value);
                    break;

                case XmlNodeType.CDATA:
                    writer.WriteCData(reader.Value);
                    break;

                case XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace:
                    writer.WriteWhitespace(reader.Value);
                    break;

                case XmlNodeType.Comment:
                    writer.WriteComment(reader.Value);
                    break;

                case XmlNodeType.ProcessingInstruction:
                    writer.WriteProcessingInstruction(reader.Name, reader.Value);
                    break;

                case XmlNodeType.EntityReference:
                    writer.WriteEntityRef(reader.Name);
                    break;
            }
        }
        while (reader.Read());
    }

    private static int Target(IReadOnlyList<int?> map, string value, string kind)
    {
        int index = int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : throw new InvalidDataException($"The {kind} index '{value}' is not a number.");
        return index < map.Count && map[index] is { } target
            ? target
            : throw new InvalidDataException($"The workbook references {kind} {index}, which the cleanup plan does not cover.");
    }

    private static XElement[] Items(XElement root, string section, string item) =>
        root.Element(Main + section)?.Elements(Main + item).ToArray() ?? [];

    private static void UpdateCount(XElement root, string section, string item)
    {
        if (root.Element(Main + section) is { } element && element.Attribute("count") is { } count)
        {
            count.Value = element.Elements(Main + item).Count().ToString(CultureInfo.InvariantCulture);
        }
    }

    // Keeps pretty-printed style sheets tidy: an entry's indentation goes with it.
    private static void RemoveWithLeadingWhitespace(XElement element)
    {
        if (element.PreviousNode is XText text && string.IsNullOrWhiteSpace(text.Value))
        {
            text.Remove();
        }

        element.Remove();
    }

    // XmlWriter only writes its own declaration form, so Excel's is written as bytes ahead of it.
    // The writer always emits UTF-8, so the declaration says so whatever the source declared.
    private static void WriteDeclaration(Stream output, string? version, string? standalone, string lineBreak)
    {
        StringBuilder text = new($"<?xml version=\"{version ?? "1.0"}\" encoding=\"UTF-8\"");
        if (standalone is not null)
        {
            text.Append(CultureInfo.InvariantCulture, $" standalone=\"{standalone}\"");
        }

        text.Append("?>").Append(lineBreak);
        output.Write(Encoding.UTF8.GetBytes(text.ToString()));
    }
}
