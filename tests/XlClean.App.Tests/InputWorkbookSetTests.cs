namespace XlClean.App.Tests;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlClean.App;

[TestClass]
public sealed class InputWorkbookSetTests
{
    private string testDirectory = null!;

    [TestInitialize]
    public void Initialize()
    {
        testDirectory = Path.Combine(
            Path.GetTempPath(),
            "XlClean.App.Tests",
            Guid.NewGuid().ToString("N"));
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
    public void AddRangePreservesOrderNormalizesPathsAndReportsEveryCandidate()
    {
        string first = CreateFile("First.xlsx");
        string second = CreateFile("Second.XLSM");
        string unsupported = CreateFile("Notes.csv");
        string missing = Path.Combine(testDirectory, "Missing.xlsx");
        string directory = Path.Combine(testDirectory, "Folder.xlsx");
        Directory.CreateDirectory(directory);

        InputWorkbookSet workbooks = new();
        InputAdditionResult result = workbooks.Add(
        [
            Path.Combine(testDirectory, ".", Path.GetFileName(first)),
            second,
            first.ToUpperInvariant(),
            unsupported,
            missing,
            directory,
        ]);

        Assert.AreEqual(2, result.AddedCount);
        Assert.AreEqual(1, result.DuplicateCount);
        Assert.AreEqual(3, result.RejectedCount);
        Assert.AreEqual(6, result.CandidateCount);
        CollectionAssert.AreEqual(
            new[] { Path.GetFullPath(first), Path.GetFullPath(second) },
            workbooks.Items.ToArray());
    }

    [TestMethod]
    public void InvalidCandidatesAreRejectedWithoutThrowing()
    {
        InputWorkbookSet workbooks = new();

        InputAdditionResult result = workbooks.Add(
            new[] { null!, string.Empty, "   ", "bad\0path.xlsx" });

        Assert.AreEqual(0, result.AddedCount);
        Assert.AreEqual(0, result.DuplicateCount);
        Assert.AreEqual(4, result.RejectedCount);
        Assert.AreEqual(0, workbooks.Count);
    }

    [TestMethod]
    public void AcceptsBinaryWorkbooks()
    {
        string binary = CreateFile("Binary.XLSB");
        InputWorkbookSet workbooks = new();

        InputAdditionResult result = workbooks.Add([binary]);

        Assert.AreEqual(1, result.AddedCount);
        CollectionAssert.AreEqual(new[] { Path.GetFullPath(binary) }, workbooks.Items.ToArray());
    }

    [TestMethod]
    public void SnapshotDoesNotChangeAfterRemoveAddOrClear()
    {
        string first = CreateFile("First.xlsx");
        string second = CreateFile("Second.xlsm");
        string third = CreateFile("Third.xlsx");
        InputWorkbookSet workbooks = new();
        workbooks.Add([first, second]);
        IReadOnlyList<string> snapshot = workbooks.Items;

        Assert.AreEqual(1, workbooks.Remove([first]));
        workbooks.Add([third]);
        workbooks.Clear();

        Assert.AreEqual(0, workbooks.Count);
        CollectionAssert.AreEqual(
            new[] { Path.GetFullPath(first), Path.GetFullPath(second) },
            snapshot.ToArray());
        Assert.ThrowsExactly<NotSupportedException>(() =>
            ((IList<string>)snapshot).Add("Other.xlsx"));
    }

    [TestMethod]
    public void RemoveUsesNormalizedIdentityEvenAfterSourceDisappears()
    {
        string input = CreateFile("Book.xlsx");
        InputWorkbookSet workbooks = new();
        workbooks.Add([input]);
        File.Delete(input);

        Assert.AreEqual(1, workbooks.Remove([input.ToUpperInvariant()]));
        Assert.AreEqual(0, workbooks.Remove([input]));
        Assert.AreEqual(0, workbooks.Count);
    }

    [TestMethod]
    public void AddRejectsNullSequence()
    {
        InputWorkbookSet workbooks = new();

        Assert.ThrowsExactly<ArgumentNullException>(() => workbooks.Add(null!));
    }

    private string CreateFile(string fileName)
    {
        string path = Path.Combine(testDirectory, fileName);
        File.WriteAllBytes(path, [0x50, 0x4B]);
        return path;
    }
}
