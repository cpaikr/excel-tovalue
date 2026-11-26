namespace XlClean.App.Tests;

using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlClean.App;

[TestClass]
public sealed class MainFormContractTests
{
    [TestMethod]
    [DataRow(CloseReason.WindowsShutDown, true)]
    [DataRow(CloseReason.ApplicationExitCall, false)]
    public void CloseReasonControlsWhetherPromptingIsAllowed(
        CloseReason closeReason,
        bool expected)
    {
        Assert.AreEqual(expected, MainForm.IsNonInteractiveCloseReason(closeReason));
    }

    [STATestMethod]
    public void FormDefinesDefaultAndCancellationActions()
    {
        using MainForm form = new();

        Assert.IsNotNull(form.AcceptButton);
        Assert.IsNotNull(form.CancelButton);
    }

    [STATestMethod]
    public void OperationSelectorSwitchesTheStartAction()
    {
        using MainForm form = new();
        Button start = (Button)form.AcceptButton!;
        RadioButton values = Find<RadioButton>(form, "수식을 값으로 바꾸기(&V)");
        RadioButton styles = Find<RadioButton>(form, "스타일 정리(&T)");

        Assert.IsTrue(values.Checked);
        Assert.AreEqual("값으로 변환(&C)", start.Text);
        Assert.AreEqual("3. 저장 위치와 범위", Find<GroupBox>(form, "3. 저장 위치와 범위").Text);

        styles.Checked = true;

        Assert.IsFalse(values.Checked);
        Assert.AreEqual("정리 시작(&C)", start.Text);
        Find<GroupBox>(form, "3. 저장 위치");
    }

    [STATestMethod]
    [DataRow(false)]
    [DataRow(true)]
    public void AccessKeysAreUniqueInEitherOperation(bool styles)
    {
        using MainForm form = new();
        Find<RadioButton>(form, styles ? "스타일 정리(&T)" : "수식을 값으로 바꾸기(&V)").Checked = true;

        char[] keys = Descendants(form)
            .Select(control => control.Text)
            .Select(text => text.IndexOf('&', StringComparison.Ordinal) is int index and >= 0 &&
                index + 1 < text.Length
                ? char.ToUpperInvariant(text[index + 1])
                : '\0')
            .Where(key => key != '\0')
            .ToArray();

        Assert.HasCount(keys.Distinct().Count(), keys);
    }

    private static T Find<T>(Control root, string text)
        where T : Control =>
        Descendants(root).OfType<T>().SingleOrDefault(control => control.Text == text) ??
            throw new AssertFailedException($"No {typeof(T).Name} with text '{text}'.");

    private static IEnumerable<Control> Descendants(Control root)
    {
        foreach (Control child in root.Controls)
        {
            yield return child;
            foreach (Control descendant in Descendants(child))
            {
                yield return descendant;
            }
        }
    }
}
