namespace XlClean.App.Tests;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlClean.App;
using XlClean.Core;

[TestClass]
public sealed class BatchSummaryFormatterTests
{
    [TestMethod]
    public void CompletionIncludesCountsPublishedOutputsGuidanceAndWarnings()
    {
        DesktopBatchResult result = new(
        [
            Success(
                @"C:\input\Good.xlsx",
                @"C:\output\Good_수식제거.xlsx",
                ["The workbook uses manual calculation; cached values may be out of date."]),
            Success(@"C:\input\Macro.xlsm", @"C:\output\Macro_수식제거.xlsm"),
            Failure(
                @"C:\input\Missing.xlsx",
                WorkbookFailureKind.MissingCachedValue,
                "Worksheet 'Data', cell 'A1' has no cached value."),
        ]);

        string summary = BatchSummaryFormatter.Format(
            result,
            canceled: false,
            @"C:\output");

        Assert.AreEqual(
            """
            처리가 완료되었습니다.
            전체 3개 · 성공 2개 · 실패 1개
            출력 폴더: C:\output

            확인이 필요한 파일
            • C:\input\Missing.xlsx — Excel에서 수식을 다시 계산하고 저장한 뒤 다시 시도하세요.
              세부 정보: Worksheet 'Data', cell 'A1' has no cached value.

            주의 사항
            • C:\input\Good.xlsx — 통합 문서가 수동 계산 모드입니다. 저장된 값이 최신인지 Excel에서 확인하세요.

            저장된 파일
            • C:\output\Good_수식제거.xlsx
            • C:\output\Macro_수식제거.xlsm
            """.ReplaceLineEndings("\n"),
            summary);
    }

    [TestMethod]
    public void CancellationClearlyDescribesCompletedOutputs()
    {
        DesktopBatchResult completed = new(
            [Success("First.xlsx", @"C:\output\First_수식제거.xlsx")]);

        string summary = BatchSummaryFormatter.Format(
            completed,
            canceled: true,
            @"C:\output");

        Assert.AreEqual(
            """
            처리가 중지되었습니다.
            중지 전 완료 1개 · 성공 1개 · 실패 0개
            진행 중이던 파일은 저장되지 않았습니다.
            출력 폴더: C:\output

            저장된 파일
            • C:\output\First_수식제거.xlsx
            """.ReplaceLineEndings("\n"),
            summary);
    }

    [TestMethod]
    public void EmptyPartialResultStatesThatNothingWasPublished()
    {
        string summary = BatchSummaryFormatter.Format(
            new DesktopBatchResult([]),
            canceled: true,
            outputDirectory: string.Empty);

        StringAssert.Contains(summary, "중지 전 완료 0개 · 성공 0개 · 실패 0개");
        StringAssert.Contains(summary, "출력 폴더: (경로 정보 없음)");
        StringAssert.Contains(summary, "저장된 파일이 없습니다.");
    }

    [TestMethod]
    public void DefaultLimitIsGlobalAndPrioritizesRecoveryDetails()
    {
        List<DesktopFileResult> files = [];
        for (int index = 0; index < 10; index++)
        {
            files.Add(Success(
                $"Good{index}.xlsx",
                $"Good{index}_수식제거.xlsx",
                [$"warning {index}"]));
            files.Add(Failure(
                $"Bad{index}.xlsx",
                WorkbookFailureKind.InvalidWorkbook,
                $"error {index}"));
        }

        string summary = BatchSummaryFormatter.Format(
            new DesktopBatchResult(files),
            canceled: false,
            "output");

        Assert.AreEqual(1, CountOccurrences(summary, "• 그 외 2개"));
        Assert.AreEqual(2, CountOccurrences(summary, "• 그 외 10개"));
        Assert.AreEqual(8, CountOccurrences(summary, "• Bad"));
        Assert.DoesNotContain(summary, "Good0_수식제거.xlsx");
        Assert.DoesNotContain(summary, "warning 0");
        Assert.DoesNotContain(summary, "Good8_수식제거.xlsx");
        Assert.DoesNotContain(summary, "Bad8.xlsx");
    }

    [TestMethod]
    [DataRow(WorkbookFailureKind.InvalidInput, "파일 경로와 .xlsx/.xlsm/.xlsb 확장자")]
    [DataRow(WorkbookFailureKind.MissingCachedValue, "수식을 다시 계산")]
    [DataRow(WorkbookFailureKind.FormulaRangeCrossesBoundary, "시작 행을 1로")]
    [DataRow(WorkbookFailureKind.InvalidWorkbook, "Excel에서 정상적으로 열리는지")]
    [DataRow(WorkbookFailureKind.FileAccess, "읽기/쓰기 권한")]
    [DataRow(WorkbookFailureKind.Unexpected, "문제가 계속되면")]
    public void EveryFailureKindHasRecoveryGuidance(
        WorkbookFailureKind failureKind,
        string expectedGuidance)
    {
        DesktopBatchResult result = new(
            [Failure("Book.xlsx", failureKind, "detail")]);

        string summary = BatchSummaryFormatter.Format(result, false, "output");

        StringAssert.Contains(summary, expectedGuidance);
    }

    [TestMethod]
    public void DynamicMessagesAndPathsAreSingleLineAndBounded()
    {
        string longTail = new('x', 5_000);
        string longPath = @"C:\" + new string('d', 500) + @"\Book.xlsx";
        DesktopBatchResult result = new(
        [
            Failure(
                longPath,
                WorkbookFailureKind.Unexpected,
                "first\r\nsecond\t" + longTail + "UNBOUNDED_TAIL"),
        ]);

        string summary = BatchSummaryFormatter.Format(result, false, longPath);

        Assert.DoesNotContain(summary, "UNBOUNDED_TAIL");
        Assert.DoesNotContain(summary, "first\r\nsecond");
        StringAssert.Contains(summary, "first second");
        StringAssert.Contains(summary, "…");
        Assert.IsLessThan(1_500, summary.Length);
    }

    [TestMethod]
    public void StyleCleanupShowsCountsAndExplainsStyleWarnings()
    {
        DesktopBatchResult result = new(
        [
            new DesktopFileResult(
                @"C:\input\Report.xlsm",
                @"C:\output\Report_스타일정리.xlsm",
                WorkbookFailureKind.None,
                null,
                new StyleCleanupChanges(2001, 1, 40, 12, 3),
                [
                    "The workbook has macros and 2000 unused custom cell styles were removed; " +
                        $"{StyleCleanupEngine.MacroStyleWarningPhrase} may no longer find it.",
                    "The cleaned workbook still has 70000 cell and style formats, more than the 65430 " +
                        $"Excel supports; it may still report {StyleCleanupEngine.FormatLimitWarningPhrase}.",
                ]),
            new DesktopFileResult(
                @"C:\input\Clean.xlsx",
                @"C:\output\Clean_스타일정리.xlsx",
                WorkbookFailureKind.None,
                null,
                new StyleCleanupChanges(1, 1, 2, 2, 0),
                Array.Empty<string>()),
        ]);

        string summary = BatchSummaryFormatter.Format(result, canceled: false, @"C:\output");

        Assert.AreEqual(
            """
            처리가 완료되었습니다.
            전체 2개 · 성공 2개 · 실패 0개
            출력 폴더: C:\output

            주의 사항
            • C:\input\Report.xlsm — 매크로가 있는 통합 문서에서 사용하지 않는 사용자 지정 셀 스타일 2000개를 제거했습니다. 스타일 이름으로 서식을 적용하는 매크로가 있다면 동작을 확인하세요.
            • C:\input\Report.xlsm — 정리한 뒤에도 셀 서식과 스타일 서식 70000개가 Excel 한도 65430개를 넘습니다. '셀 서식이 너무 많습니다' 오류가 계속될 수 있습니다.

            저장된 파일
            • C:\output\Report_스타일정리.xlsm
              셀 스타일 2001 → 1개 · 셀 서식 40 → 12개 · 같은 서식 3개 병합
            • C:\output\Clean_스타일정리.xlsx
              셀 스타일 1 → 1개 · 셀 서식 2 → 2개
            """.ReplaceLineEndings("\n"),
            summary);
    }

    [TestMethod]
    public void StyleWarningWithoutACountStaysGrammatical()
    {
        DesktopBatchResult result = new(
            [Success("Book.xlsm", "Out.xlsm", [$"Styles were removed; {StyleCleanupEngine.MacroStyleWarningPhrase}."])]);

        string summary = BatchSummaryFormatter.Format(result, canceled: false, "output");

        StringAssert.Contains(summary, "사용자 지정 셀 스타일을 제거했습니다.");
    }

    [TestMethod]
    public void UnrecognizedWarningIsShownAsIs()
    {
        DesktopBatchResult result = new([Success("Book.xlsx", "Out.xlsx", ["something new"])]);

        string summary = BatchSummaryFormatter.Format(result, canceled: false, "output");

        StringAssert.Contains(summary, "• Book.xlsx — something new");
    }

    [TestMethod]
    [DataRow(0)]
    [DataRow(51)]
    public void DetailLimitHasAValidBoundedRange(int detailLimit)
    {
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() =>
            BatchSummaryFormatter.Format(
                new DesktopBatchResult([]),
                false,
                "output",
                detailLimit));
    }

    private static DesktopFileResult Success(
        string inputPath,
        string outputPath,
        IReadOnlyList<string>? warnings = null) =>
        new(
            inputPath,
            outputPath,
            WorkbookFailureKind.None,
            null,
            null,
            warnings ?? Array.Empty<string>());

    private static DesktopFileResult Failure(
        string inputPath,
        WorkbookFailureKind failureKind,
        string message) =>
        new(
            inputPath,
            null,
            failureKind,
            message,
            null,
            Array.Empty<string>());

    private static int CountOccurrences(string value, string expected)
    {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(expected, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += expected.Length;
        }

        return count;
    }
}
