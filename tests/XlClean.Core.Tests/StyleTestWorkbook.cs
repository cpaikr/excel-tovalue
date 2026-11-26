namespace XlClean.Core.Tests;

using System.IO.Packaging;
using System.Text;

/// <summary>
/// Builds minimal .xlsx packages from raw part XML, with fixed relationship ids, so tests control
/// every byte the style reader and verifier see.
/// </summary>
internal static class StyleTestWorkbook
{
    public const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public const string WorksheetUri = "/xl/worksheets/sheet1.xml";

    public const string StylesUri = "/xl/styles.xml";

    public const string SharedStringsUri = "/xl/sharedStrings.xml";

    public const string TableUri = "/xl/tables/table1.xml";

    private const string RelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static void Create(
        string path,
        string? styles,
        string sheetData,
        string? table = null,
        (string Type, string ContentType, string Uri)? extraWorkbookPart = null)
    {
        using Package package = Package.Open(path, FileMode.CreateNew, FileAccess.ReadWrite);
        PackagePart workbook = AddPart(
            package,
            "/xl/workbook.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            $"""<workbook xmlns="{Main}" xmlns:r="{RelationshipNamespace}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>""");
        package.CreateRelationship(
            workbook.Uri,
            TargetMode.Internal,
            $"{RelationshipNamespace}/officeDocument",
            "rId1");

        PackagePart worksheet = AddPart(
            package,
            WorksheetUri,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            Worksheet(sheetData, table is not null));
        workbook.CreateRelationship(worksheet.Uri, TargetMode.Internal, $"{RelationshipNamespace}/worksheet", "rId1");

        if (styles is not null)
        {
            PackagePart stylesPart = AddPart(
                package,
                StylesUri,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                styles);
            workbook.CreateRelationship(stylesPart.Uri, TargetMode.Internal, $"{RelationshipNamespace}/styles", "rId2");
        }

        PackagePart sharedStrings = AddPart(
            package,
            SharedStringsUri,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            $"""<sst xmlns="{Main}" count="1" uniqueCount="1"><si><t>text</t></si></sst>""");
        workbook.CreateRelationship(
            sharedStrings.Uri,
            TargetMode.Internal,
            $"{RelationshipNamespace}/sharedStrings",
            "rId3");

        if (table is not null)
        {
            PackagePart tablePart = AddPart(
                package,
                TableUri,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
                table);
            worksheet.CreateRelationship(
                new Uri("../tables/table1.xml", UriKind.Relative),
                TargetMode.Internal,
                $"{RelationshipNamespace}/table",
                "rId1");
            _ = tablePart;
        }

        if (extraWorkbookPart is { } extra)
        {
            PackagePart part = AddPart(package, extra.Uri, extra.ContentType, $"""<root xmlns="{Main}"/>""");
            workbook.CreateRelationship(part.Uri, TargetMode.Internal, extra.Type, "rId9");
        }
    }

    public static string Worksheet(string sheetData, bool hasTable = false) =>
        $"""<worksheet xmlns="{Main}" xmlns:r="{RelationshipNamespace}">{sheetData}{(hasTable ? """<tableParts count="1"><tablePart r:id="rId1"/></tableParts>""" : string.Empty)}</worksheet>""";

    /// <summary>A style sheet with one font, the two required fills, and one border.</summary>
    public static string Styles(
        string cellStyleXfs,
        string cellXfs,
        string cellStyles,
        string numFmts = "",
        string fonts = """<font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font>""") =>
        $"""<styleSheet xmlns="{Main}">{(numFmts.Length == 0 ? string.Empty : $"<numFmts>{numFmts}</numFmts>")}<fonts>{fonts}</fonts><fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="{Count(cellStyleXfs, "<xf")}">{cellStyleXfs}</cellStyleXfs><cellXfs count="{Count(cellXfs, "<xf")}">{cellXfs}</cellXfs><cellStyles count="{Count(cellStyles, "<cellStyle ")}">{cellStyles}</cellStyles></styleSheet>""";

    public static void ReplacePart(string path, string uri, string xml)
    {
        using Package package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);
        PackagePart part = package.GetPart(new Uri(uri, UriKind.Relative));
        using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
        byte[] bytes = Encoding.UTF8.GetBytes(xml);
        stream.Write(bytes);
    }

    public static void AddLoosePart(string path, string uri)
    {
        using Package package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);
        PackagePart part = package.CreatePart(new Uri(uri, UriKind.Relative), "application/xml");
        using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
        stream.Write("<extra/>"u8);
    }

    private static int Count(string xml, string marker)
    {
        int count = 0;
        for (int index = xml.IndexOf(marker, StringComparison.Ordinal); index >= 0; index = xml.IndexOf(marker, index + 1, StringComparison.Ordinal))
        {
            count++;
        }

        return count;
    }

    private static PackagePart AddPart(Package package, string uri, string contentType, string xml)
    {
        PackagePart part = package.CreatePart(new Uri(uri, UriKind.Relative), contentType, CompressionOption.Normal);
        using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
        byte[] bytes = Encoding.UTF8.GetBytes(xml);
        stream.Write(bytes);
        return part;
    }
}
