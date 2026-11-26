namespace XlClean.Core.Tests;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlClean.Core.Styles;

[TestClass]
public sealed class StylePlannerTests
{
    private const string Plain = "<xf/>";

    [TestMethod]
    public void RemovesUnusedCellFormatsAndKeepsTheDefault()
    {
        StyleSheetModel model = Model(
            cellFormats: [Format(Plain, font: 0), Format("<xf a=\"1\"/>", font: 0), Format("<xf a=\"2\"/>", font: 0)]);

        StylePlan plan = StylePlanner.Plan(model, Usage([2]));

        CollectionAssert.AreEqual(new int?[] { 0, null, 1 }, plan.CellFormatMap.ToArray());
        Assert.AreEqual(2, plan.KeptCellFormatCount);
        Assert.AreEqual(0, plan.MergedCellFormats);
        Assert.IsTrue(plan.RenumbersCellFormats);
    }

    [TestMethod]
    public void NothingToRemoveIsTheIdentity()
    {
        StyleSheetModel model = Model(
            cellFormats: [Format(Plain, font: 0), Format("<xf a=\"1\"/>", font: 0)]);

        StylePlan plan = StylePlanner.Plan(model, Usage([1]));

        CollectionAssert.AreEqual(new int?[] { 0, 1 }, plan.CellFormatMap.ToArray());
        Assert.IsFalse(plan.RenumbersCellFormats);
    }

    [TestMethod]
    public void MergesUsedFormatsWithIdenticalResolvedContent()
    {
        // Fonts 1 and 2 have the same content, so formats 1 and 2 look the same.
        StyleSheetModel model = Model(
            fonts: ["<font/>", "<font><b/></font>", "<font><b/></font>"],
            cellFormats: [Format(Plain, font: 0), Format(Plain, font: 1), Format(Plain, font: 2), Format(Plain, font: 1)]);

        StylePlan plan = StylePlanner.Plan(model, Usage([1, 2, 3]));

        CollectionAssert.AreEqual(new int?[] { 0, 1, 1, 1 }, plan.CellFormatMap.ToArray());
        Assert.AreEqual(2, plan.MergedCellFormats);
        Assert.AreEqual(2, plan.KeptCellFormatCount);
    }

    [TestMethod]
    public void ContentDifferencesPreventMerges()
    {
        StyleSheetModel model = Model(
            fonts: ["<font/>", "<font><b/></font>"],
            cellFormats:
            [
                Format(Plain, font: 0),
                Format(Plain, font: 1),
                Format("<xf applyFont=\"1\"/>", font: 1),
                Format(Plain, font: 1, fill: 1),
                Format(Plain, font: 1, border: null),
            ]);

        StylePlan plan = StylePlanner.Plan(model, Usage([1, 2, 3, 4]));

        Assert.AreEqual(0, plan.MergedCellFormats);
        CollectionAssert.AreEqual(new int?[] { 0, 1, 2, 3, 4 }, plan.CellFormatMap.ToArray());
    }

    [TestMethod]
    public void BuiltInNumberFormatNeverEqualsACustomCodeWithTheSameText()
    {
        StyleSheetModel model = Model(
            numberFormats: new Dictionary<int, string> { [164] = "m/d/yyyy", [165] = "m/d/yyyy" },
            cellFormats:
            [
                Format(Plain, font: 0),
                Format(Plain, font: 0, numberFormat: 14),
                Format(Plain, font: 0, numberFormat: 164),
                Format(Plain, font: 0, numberFormat: 165),
            ]);

        StylePlan plan = StylePlanner.Plan(model, Usage([1, 2, 3]));

        // 164 and 165 share a code and merge; built-in 14 stays apart.
        CollectionAssert.AreEqual(new int?[] { 0, 1, 2, 2 }, plan.CellFormatMap.ToArray());
    }

    [TestMethod]
    public void FormatsUnderDifferentlyNamedStylesStayApart()
    {
        StyleSheetModel model = Model(
            styleFormats: [Format(Plain, font: 0), Format(Plain, font: 0)],
            namedStyles: [new NamedStyle("Normal", 0, 0, Content("Normal")), new NamedStyle("Normal 2", 1, null, Content("Normal 2"))],
            cellFormats: [Format(Plain, font: 0, parent: 0), Format(Plain, font: 0, parent: 1)]);

        StylePlan plan = StylePlanner.Plan(model, Usage([1]));

        CollectionAssert.AreEqual(new int?[] { 0, 1 }, plan.CellFormatMap.ToArray());
        Assert.AreEqual("0,1", string.Join(',', plan.KeptNamedStyles));
    }

    [TestMethod]
    public void StyleFlagsAndAnImplicitParentKeepFormatsApart()
    {
        StyleSheetModel model = Model(
            styleFormats: [Format(Plain, font: 0), Format(Plain, font: 0), Format(Plain, font: 0)],
            namedStyles:
            [
                new NamedStyle("Normal", 0, 0, Content("Normal")),
                new NamedStyle("Note", 1, null, "<cellStyle name=\"Note\"/>"),
                new NamedStyle("Note", 2, null, "<cellStyle name=\"Note\" hidden=\"1\"/>"),
            ],
            cellFormats:
            [
                Format(Plain, font: 0, parent: 0),
                Format(Plain, font: 0, parent: null),
                Format(Plain, font: 0, parent: 1),
                Format(Plain, font: 0, parent: 2),
            ]);

        StylePlan plan = StylePlanner.Plan(model, Usage([1, 2, 3]));

        Assert.AreEqual(0, plan.MergedCellFormats);
    }

    [TestMethod]
    public void NamesThatLookLikeKeySeparatorsCannotCollide()
    {
        // One style whose entry embeds what a pair of styles would produce in the key.
        StyleSheetModel model = Model(
            styleFormats: [Format(Plain, font: 0), Format(Plain, font: 0), Format(Plain, font: 0)],
            namedStyles:
            [
                new NamedStyle("Normal", 0, 0, Content("Normal")),
                new NamedStyle("A", 1, null, "A"),
                new NamedStyle("B", 1, null, "B"),
                new NamedStyle("AB", 2, null, "A" + "style#1:B"),
            ],
            cellFormats: [Format(Plain, font: 0, parent: 0), Format(Plain, font: 0, parent: 1), Format(Plain, font: 0, parent: 2)]);

        StylePlan plan = StylePlanner.Plan(model, Usage([1, 2]));

        Assert.AreEqual(0, plan.MergedCellFormats);
    }

    [TestMethod]
    public void RemovesOnlyUnreferencedCustomNamedStyles()
    {
        StyleSheetModel model = Model(
            styleFormats:
            [
                Format(Plain, font: 0),
                Format("<xf s=\"good\"/>", font: 0),
                Format("<xf s=\"used\"/>", font: 0),
                Format("<xf s=\"junk\"/>", font: 0),
                Format("<xf s=\"table\"/>", font: 0),
                Format("<xf s=\"orphan\"/>", font: 0),
            ],
            namedStyles:
            [
                new NamedStyle("Normal", 0, 0, Content("Normal")),
                new NamedStyle("Good", 1, 26, Content("Good")),
                new NamedStyle("Used custom", 2, null, Content("Used custom")),
                new NamedStyle("Junk 1", 3, null, Content("Junk 1")),
                new NamedStyle("Header", 4, null, Content("Header")),
            ],
            cellFormats: [Format(Plain, font: 0, parent: 0), Format(Plain, font: 0, parent: 2), Format(Plain, font: 0, parent: 3)]);

        StylePlan plan = StylePlanner.Plan(model, Usage([1], tableStyles: ["Header"]));

        Assert.AreEqual("0,1,2,4", string.Join(',', plan.KeptNamedStyles));
        CollectionAssert.AreEqual(new int?[] { 0, 1, 2, null, 3, null }, plan.CellStyleFormatMap.ToArray());
        CollectionAssert.AreEqual(new int?[] { 0, 1, null }, plan.CellFormatMap.ToArray());
    }

    [TestMethod]
    public void UnresolvableReferencesFailInsteadOfGuessing()
    {
        Assert.ThrowsExactly<InvalidDataException>(() => StylePlanner.Plan(
            Model(cellFormats: [Format(Plain, font: 5)]),
            Usage([])));
        Assert.ThrowsExactly<InvalidDataException>(() => StylePlanner.Plan(
            Model(cellFormats: [Format(Plain, font: 0, numberFormat: 170)]),
            Usage([])));
        Assert.ThrowsExactly<InvalidDataException>(() => StylePlanner.Plan(
            Model(cellFormats: [Format(Plain, font: 0, parent: 4)]),
            Usage([])));
        Assert.ThrowsExactly<InvalidDataException>(() => StylePlanner.Plan(
            Model(namedStyles: [new NamedStyle("Broken", 7, null, Content("Broken"))], cellFormats: [Format(Plain, font: 0)]),
            Usage([])));
    }

    [TestMethod]
    public void WorkbookWithoutStylesPlansNothing()
    {
        StylePlan plan = StylePlanner.Plan(StyleSheetModel.Empty, Usage([3]));

        Assert.IsEmpty(plan.CellFormatMap);
        Assert.IsFalse(plan.RenumbersCellFormats);
    }

    private static StyleFormat Format(
        string own,
        int? font,
        int? parent = null,
        int? numberFormat = 0,
        int? fill = 0,
        int? border = 0) =>
        new(numberFormat, font, fill, border, parent, own);

    private static string Content(string name) => $"<cellStyle name=\"{name}\"/>";

    private static StyleUsage Usage(int[] cellFormats, string[]? tableStyles = null) =>
        new(new HashSet<int>(cellFormats.Append(0)), new HashSet<string>(tableStyles ?? [], StringComparer.Ordinal));

    private static StyleSheetModel Model(
        StyleFormat[] cellFormats,
        string[]? fonts = null,
        StyleFormat[]? styleFormats = null,
        NamedStyle[]? namedStyles = null,
        Dictionary<int, string>? numberFormats = null) =>
        new(
            numberFormats ?? [],
            fonts ?? ["<font/>"],
            ["<fill none/>", "<fill gray/>"],
            ["<border/>"],
            styleFormats ?? [Format(Plain, font: 0)],
            cellFormats,
            namedStyles ?? [new NamedStyle("Normal", 0, 0, Content("Normal"))]);
}
