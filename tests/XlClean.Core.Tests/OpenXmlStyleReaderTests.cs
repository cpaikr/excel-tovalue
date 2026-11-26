namespace XlClean.Core.Tests;

using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlClean.Core.Styles;
using static XlClean.Core.Tests.StyleTestWorkbook;

[TestClass]
public sealed class OpenXmlStyleReaderTests
{
    private const string BasicStyles =
        """<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>""";

    private string testDirectory = null!;

    [TestInitialize]
    public void Initialize()
    {
        testDirectory = Path.Combine(
            Path.GetTempPath(),
            "XlClean.Tests",
            Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        Directory.CreateDirectory(testDirectory);
    }

    [TestCleanup]
    public void Cleanup()
    {
        if (Directory.Exists(testDirectory))
        {
            Directory.Delete(testDirectory, recursive: true);
        }
    }

    [TestMethod]
    public void CollectsCellRowColumnAndTableReferences()
    {
        string path = Path.Combine(testDirectory, "Book.xlsx");
        Create(
            path,
            Styles(
                BasicStyles,
                string.Concat(Enumerable.Repeat("""<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>""", 6)),
                """<cellStyle name="Normal" xfId="0" builtinId="0"/>"""),
            """<cols><col min="1" max="1" style="4"/></cols><sheetData><row r="1" s="3" customFormat="1"><c r="A1" s="2"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>""",
            table: $"""<table xmlns="{Main}" id="1" name="T" displayName="T" ref="A1:B2" headerRowCellStyle="Head"><tableColumns count="2"><tableColumn id="1" name="A" dataCellStyle="Body"/><tableColumn id="2" name="B" totalsRowCellStyle="Total"/></tableColumns></table>""");

        (StyleSheetModel model, StyleUsage usage) = OpenXmlStyleReader.Read(path, CancellationToken.None);

        Assert.AreEqual("0,2,3,4", string.Join(',', usage.CellFormatIndices.Order()));
        Assert.AreEqual("Body,Head,Total", string.Join(',', usage.TableStyleNames.Order(StringComparer.Ordinal)));
        Assert.HasCount(6, model.CellFormats);
        Assert.AreEqual(0, model.CellFormats[1].ParentId);
        Assert.AreEqual("Normal", model.NamedStyles.Single().Name);
        Assert.AreEqual(0, model.NamedStyles.Single().BuiltinId);
    }

    [TestMethod]
    public void CanonicalContentIgnoresAttributeOrderButKeepsAbsence()
    {
        string path = Path.Combine(testDirectory, "Book.xlsx");
        Create(
            path,
            Styles(
                BasicStyles,
                """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="0"/>""" +
                """<xf applyFill="0" borderId="0" fillId="0" fontId="0" numFmtId="0" applyFont="1"/>""" +
                """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1"/>""" +
                """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="0"><alignment horizontal="center"/></xf>""",
                """<cellStyle name="Normal" xfId="0" builtinId="0"/>"""),
            "<sheetData/>");

        StyleSheetModel model = OpenXmlStyleReader.Read(path, CancellationToken.None).Model;

        Assert.AreEqual(model.CellFormats[0].OwnContent, model.CellFormats[1].OwnContent);
        Assert.AreNotEqual(model.CellFormats[0].OwnContent, model.CellFormats[2].OwnContent);
        Assert.AreNotEqual(model.CellFormats[0].OwnContent, model.CellFormats[3].OwnContent);
        Assert.DoesNotContain("fontId", model.CellFormats[0].OwnContent);
    }

    [TestMethod]
    [DataRow("http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet", "application/vnd.ms-excel.macrosheet+xml")]
    [DataRow("http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet", "application/vnd.ms-excel.intlmacrosheet+xml")]
    [DataRow("http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders", "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml")]
    public void UnsupportedPartsFailClosed(string relationshipType, string contentType)
    {
        string path = Path.Combine(testDirectory, "Book.xlsx");
        Create(
            path,
            Styles(BasicStyles, BasicStyles, """<cellStyle name="Normal" xfId="0" builtinId="0"/>"""),
            "<sheetData/>",
            extraWorkbookPart: (relationshipType, contentType, "/xl/extra.xml"));

        Assert.ThrowsExactly<UnsupportedStyleContentException>(
            () => OpenXmlStyleReader.Read(path, CancellationToken.None));
    }

    [TestMethod]
    public void StrictNamespaceIsRefusedRatherThanReadAsEmpty()
    {
        string path = Path.Combine(testDirectory, "Book.xlsx");
        Create(
            path,
            """<styleSheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><cellXfs count="1"><xf/></cellXfs></styleSheet>""",
            "<sheetData/>");

        UnsupportedStyleContentException exception = Assert.ThrowsExactly<UnsupportedStyleContentException>(
            () => OpenXmlStyleReader.Read(path, CancellationToken.None));
        Assert.AreEqual(WorkbookFailureKind.UnsupportedWorkbookContent, exception.Kind);
    }

    [TestMethod]
    public void CellFormatsWithoutAStyleSheetAreAnInvalidWorkbook()
    {
        string path = Path.Combine(testDirectory, "Book.xlsx");
        Create(path, styles: null, """<sheetData><row r="1"><c r="A1" s="2"/></row></sheetData>""");

        Assert.ThrowsExactly<InvalidDataException>(() => OpenXmlStyleReader.Read(path, CancellationToken.None));
    }

    [TestMethod]
    public void DefaultOnlyWorkbookWithoutAStyleSheetReads()
    {
        string path = Path.Combine(testDirectory, "Book.xlsx");
        Create(path, styles: null, """<sheetData><row r="1"><c r="A1"/><c r="B1" s="0"/></row></sheetData>""");

        (StyleSheetModel model, StyleUsage usage) = OpenXmlStyleReader.Read(path, CancellationToken.None);

        Assert.IsEmpty(model.CellFormats);
        Assert.AreEqual("0", string.Join(',', usage.CellFormatIndices));
    }

    [TestMethod]
    public void StrictPackageIsUnsupported()
    {
        string path = Path.Combine(testDirectory, "Strict.xlsx");
        using (System.IO.Packaging.Package package = System.IO.Packaging.Package.Open(path, FileMode.CreateNew))
        {
            System.IO.Packaging.PackagePart workbook = package.CreatePart(
                new Uri("/xl/workbook.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            package.CreateRelationship(
                workbook.Uri,
                System.IO.Packaging.TargetMode.Internal,
                "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument");
        }

        Assert.ThrowsExactly<UnsupportedStyleContentException>(() => OpenXmlStyleReader.Read(path, CancellationToken.None));
    }

    [TestMethod]
    [DataRow("""<c r="A1" s="5"/>""")]
    [DataRow("""<c r="A1" s="-1"/>""")]
    [DataRow("""<c r="A1" s="x"/>""")]
    public void InvalidCellStyleIndexIsAnInvalidWorkbook(string cell)
    {
        string path = Path.Combine(testDirectory, "Book.xlsx");
        Create(
            path,
            Styles(BasicStyles, BasicStyles, """<cellStyle name="Normal" xfId="0" builtinId="0"/>"""),
            $"<sheetData><row r=\"1\">{cell}</row></sheetData>");

        Assert.ThrowsExactly<InvalidDataException>(() => OpenXmlStyleReader.Read(path, CancellationToken.None));
    }
}
