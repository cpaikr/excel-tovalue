namespace XlClean.Core.Tests;

using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlClean.Core.Styles;
using static XlClean.Core.Tests.StyleTestWorkbook;

/// <summary>
/// The verifier must accept an equivalent cleanup and reject each kind of change it claims to
/// catch, so a broken verifier cannot pass silently.
/// </summary>
[TestClass]
public sealed class OpenXmlStyleVerifierTests
{
    // Source: a junk custom style, a table-named style, and two bold formats that differ only in
    // which (identical) bold font they use, plus one unused format.
    private const string SourceStyleFormats =
        """<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>""" +
        """<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>""" +
        """<xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>""";

    private const string SourceCellFormats =
        """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>""" +
        """<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>""" +
        """<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>""" +
        """<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>""";

    private const string SourceNamedStyles =
        """<cellStyle name="Normal" xfId="0" builtinId="0"/>""" +
        """<cellStyle name="Junk 1" xfId="1"/>""" +
        """<cellStyle name="Head" xfId="2"/>""";

    private const string SourceSheet =
        """<sheetData><row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="2"><v>2</v></c><c r="C1"><v>3</v></c></row></sheetData>""";

    // The equivalent cleanup: format 2 merged into 1, format 3 removed, "Junk 1" removed.
    private const string CleanStyleFormats =
        """<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>""" +
        """<xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>""";

    private const string CleanCellFormats =
        """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>""" +
        """<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>""";

    private const string CleanNamedStyles =
        """<cellStyle name="Normal" xfId="0" builtinId="0"/>""" +
        """<cellStyle name="Head" xfId="1"/>""";

    private const string CleanSheet =
        """<sheetData><row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1"><v>2</v></c><c r="C1"><v>3</v></c></row></sheetData>""";

    private const string Table =
        $"""<table xmlns="{Main}" id="1" name="T" displayName="T" ref="A1:C2" headerRowCellStyle="Head"><tableColumns count="1"><tableColumn id="1" name="A"/></tableColumns></table>""";

    private string testDirectory = null!;
    private string source = null!;
    private string output = null!;

    [TestInitialize]
    public void Initialize()
    {
        testDirectory = Path.Combine(
            Path.GetTempPath(),
            "XlClean.Tests",
            Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        Directory.CreateDirectory(testDirectory);
        source = Path.Combine(testDirectory, "Source.xlsx");
        output = Path.Combine(testDirectory, "Output.xlsx");
        Create(source, Styles(SourceStyleFormats, SourceCellFormats, SourceNamedStyles), SourceSheet, Table);
        File.Copy(source, output);
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
    public void AcceptsAnUnchangedCopy() =>
        OpenXmlStyleVerifier.Verify(source, output, CancellationToken.None);

    [TestMethod]
    public void AcceptsAnEquivalentCleanup()
    {
        WriteOutput(CleanStyleFormats, CleanCellFormats, CleanNamedStyles, CleanSheet);

        OpenXmlStyleVerifier.Verify(source, output, CancellationToken.None);
    }

    [TestMethod]
    public void AcceptsReserializedWorksheetXml()
    {
        // A prefix instead of the default namespace, reordered attributes, an explicit default
        // style index, an empty element written long-hand, and whitespace between elements.
        ReplacePart(output, StylesUri, Styles(CleanStyleFormats, CleanCellFormats, CleanNamedStyles));
        ReplacePart(
            output,
            WorksheetUri,
            $"""<x:worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="{Main}">  <x:sheetData><x:row r="1"><x:c t="s" s="1" r="A1"><x:v>0</x:v></x:c><x:c r="B1" s="1"><x:v>2</x:v></x:c><x:c r="C1" s="0"><x:v>3</x:v></x:c></x:row></x:sheetData><x:tableParts count="1"><x:tablePart r:id="rId1"></x:tablePart></x:tableParts></x:worksheet>""");

        OpenXmlStyleVerifier.Verify(source, output, CancellationToken.None);
    }

    [TestMethod]
    public void RejectsACellThatNowLooksDifferent()
    {
        WriteOutput(
            CleanStyleFormats,
            CleanCellFormats,
            CleanNamedStyles,
            CleanSheet.Replace("""r="B1" s="1""", """r="B1" s="0""", StringComparison.Ordinal));

        AssertRejected("B1");
    }

    [TestMethod]
    public void RejectsAChangedValue() =>
        RejectSheet(CleanSheet.Replace("<v>3</v>", "<v>4</v>", StringComparison.Ordinal));

    [TestMethod]
    public void RejectsAChangedCellAttribute() =>
        RejectSheet(CleanSheet.Replace("""r="A1" s="1" t="s""", """r="A1" s="1" t="str""", StringComparison.Ordinal));

    [TestMethod]
    public void RejectsADroppedCell() =>
        RejectSheet(CleanSheet.Replace("""<c r="C1"><v>3</v></c>""", string.Empty, StringComparison.Ordinal));

    [TestMethod]
    public void RejectsAPropertyLostInAMergedFormat()
    {
        WriteOutput(
            CleanStyleFormats,
            CleanCellFormats.Replace(" applyFont=\"1\"", string.Empty, StringComparison.Ordinal),
            CleanNamedStyles,
            CleanSheet);

        AssertRejected("looks different");
    }

    [TestMethod]
    public void RejectsARemovedBuiltinStyle()
    {
        WriteOutput(
            CleanStyleFormats,
            CleanCellFormats,
            """<cellStyle name="Head" xfId="1"/>""",
            CleanSheet);

        AssertRejected("Normal");
    }

    [TestMethod]
    public void RejectsARemovedTableStyle()
    {
        WriteOutput(
            CleanStyleFormats[..CleanStyleFormats.IndexOf("/>", StringComparison.Ordinal)] + "/>",
            CleanCellFormats,
            """<cellStyle name="Normal" xfId="0" builtinId="0"/>""",
            CleanSheet);

        AssertRejected("Head");
    }

    [TestMethod]
    public void RejectsARenamedOrRepointedStyle()
    {
        WriteOutput(
            CleanStyleFormats,
            CleanCellFormats,
            CleanNamedStyles.Replace("Head", "Header", StringComparison.Ordinal),
            CleanSheet);

        AssertRejected("not a source style");
    }

    [TestMethod]
    public void RejectsAChangedComponentTable()
    {
        ReplacePart(
            output,
            StylesUri,
            Styles(
                CleanStyleFormats,
                CleanCellFormats,
                CleanNamedStyles,
                fonts: """<font><sz val="11"/><name val="Calibri"/></font><font><i/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font>"""));
        ReplacePart(output, WorksheetUri, Worksheet(CleanSheet, hasTable: true));

        AssertRejected("fonts");
    }

    [TestMethod]
    public void RejectsAStaleSectionCount()
    {
        string styles = Styles(CleanStyleFormats, CleanCellFormats, CleanNamedStyles)
            .Replace("""<cellXfs count="2">""", """<cellXfs count="4">""", StringComparison.Ordinal);
        ReplacePart(output, StylesUri, styles);
        ReplacePart(output, WorksheetUri, Worksheet(CleanSheet, hasTable: true));

        AssertRejected("count");
    }

    [TestMethod]
    public void RejectsAChangeToAnotherPart()
    {
        ReplacePart(output, SharedStringsUri, $"""<sst xmlns="{Main}" count="1" uniqueCount="1"><si><t>other</t></si></sst>""");

        AssertRejected(SharedStringsUri);
    }

    [TestMethod]
    public void RejectsAnAddedPart()
    {
        AddLoosePart(output, "/xl/extra.xml");

        AssertRejected("added");
    }

    [TestMethod]
    public void RejectsANewSchemaFindingInAnUnusedEntry()
    {
        WriteOutput(
            CleanStyleFormats,
            CleanCellFormats + "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"maybe\"/>",
            CleanNamedStyles,
            CleanSheet);

        AssertRejected("schema");
    }

    [TestMethod]
    public void RejectsADanglingIndexInAnUnusedEntry()
    {
        WriteOutput(
            CleanStyleFormats,
            CleanCellFormats + "<xf numFmtId=\"0\" fontId=\"9\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
            CleanNamedStyles,
            CleanSheet);

        AssertRejected("font 9");
    }

    [TestMethod]
    public void RejectsADefaultStyleSwapBehindAnImplicitParent()
    {
        // Cell formats without xfId belong to style format 0. Swapping which named style sits at
        // index 0 changes the style every such cell reports, with no cell index changed.
        const string styleFormats =
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>";
        const string implicitFormats =
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\"/>";
        File.Delete(source);
        File.Delete(output);
        Create(
            source,
            Styles(styleFormats, implicitFormats, "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/><cellStyle name=\"Custom\" xfId=\"1\"/>"),
            "<sheetData><row r=\"1\"><c r=\"A1\" s=\"1\"/></row></sheetData>");
        File.Copy(source, output);
        ReplacePart(
            output,
            StylesUri,
            Styles(styleFormats, implicitFormats, "<cellStyle name=\"Normal\" xfId=\"1\" builtinId=\"0\"/><cellStyle name=\"Custom\" xfId=\"0\"/>"));

        AssertRejected("default");
    }

    [TestMethod]
    public void RejectsADroppedCompatibilityNamespace()
    {
        const string declaration = " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"";
        string declared =
            $"<worksheet xmlns=\"{Main}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"{declaration} mc:Ignorable=\"x14ac\"><sheetData/></worksheet>";
        ReplacePart(source, WorksheetUri, declared);
        ReplacePart(output, WorksheetUri, declared);
        OpenXmlStyleVerifier.Verify(source, output, CancellationToken.None);

        ReplacePart(output, WorksheetUri, declared.Replace(declaration, string.Empty, StringComparison.Ordinal));

        AssertRejected("x14ac");
    }

    private void RejectSheet(string sheet)
    {
        WriteOutput(CleanStyleFormats, CleanCellFormats, CleanNamedStyles, sheet);
        AssertRejected(WorksheetUri);
    }

    private void WriteOutput(string styleFormats, string cellFormats, string namedStyles, string sheet)
    {
        ReplacePart(output, StylesUri, Styles(styleFormats, cellFormats, namedStyles));
        ReplacePart(output, WorksheetUri, Worksheet(sheet, hasTable: true));
    }

    private void AssertRejected(string expectedInMessage)
    {
        StyleVerificationException exception = Assert.ThrowsExactly<StyleVerificationException>(
            () => OpenXmlStyleVerifier.Verify(source, output, CancellationToken.None));
        StringAssert.Contains(exception.Message, expectedInMessage);
        Assert.AreEqual(WorkbookFailureKind.Unexpected, exception.Kind);
    }
}
