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
}
