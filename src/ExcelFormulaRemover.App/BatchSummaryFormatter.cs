namespace ExcelFormulaRemover.App;

using System.Globalization;
using System.Text;
using ExcelFormulaRemover.Core;

/// <summary>
/// Produces bounded, recovery-oriented Korean summaries for the native app.
/// </summary>
internal static class BatchSummaryFormatter
{
    internal const int DefaultDetailLimit = 8;
    private const int MaximumDetailLimit = 50;
    private const int MaximumMessageLength = 160;
    private const int MaximumPathLength = 160;

    /// <summary>
    /// Formats a completed or canceled batch for a native summary dialog.
    /// </summary>
    /// <param name="result">Completed per-file results, including partial cancellation results.</param>
    /// <param name="canceled">Whether the batch ended through cancellation.</param>
    /// <param name="outputDirectory">The selected output directory.</param>
    /// <param name="detailLimit">Maximum detail entries shown across all sections.</param>
    /// <returns>A concise Korean plain-text summary.</returns>
    internal static string Format(
        FormulaRemovalBatchResult result,
        bool canceled,
        string outputDirectory,
        int detailLimit = DefaultDetailLimit)
    {
        ArgumentNullException.ThrowIfNull(result);
        if (detailLimit is < 1 or > MaximumDetailLimit)
        {
            throw new ArgumentOutOfRangeException(
                nameof(detailLimit),
                detailLimit,
                $"Detail limit must be between 1 and {MaximumDetailLimit}.");
        }

        StringBuilder summary = new();
        summary.Append(canceled ? "처리가 중지되었습니다." : "처리가 완료되었습니다.");
        summary.Append('\n');
        if (canceled)
        {
            summary.Append("중지 전 완료 ");
        }
        else
        {
            summary.Append("전체 ");
        }

        summary.Append(Invariant(result.Files.Count));
        summary.Append("개 · 성공 ");
        summary.Append(Invariant(result.SucceededCount));
        summary.Append("개 · 실패 ");
        summary.Append(Invariant(result.FailedCount));
        summary.Append('개');

        if (canceled)
        {
            summary.Append('\n');
            summary.Append("진행 중이던 파일은 저장되지 않았습니다.");
        }

        summary.Append('\n');
        summary.Append("출력 폴더: ");
        summary.Append(FormatOptionalPath(outputDirectory));

        int remainingDetails = detailLimit;
        AppendFailures(summary, result, ref remainingDetails);
        AppendWarnings(summary, result, ref remainingDetails);
        AppendPublishedOutputs(summary, result, ref remainingDetails);

        return summary.ToString();
    }

    private static void AppendPublishedOutputs(
        StringBuilder summary,
        FormulaRemovalBatchResult result,
        ref int remainingDetails)
    {
        if (result.SucceededCount == 0)
        {
            AppendSectionBreak(summary);
            summary.Append("저장된 파일이 없습니다.");
            return;
        }

        AppendSectionBreak(summary);
        summary.Append("저장된 파일");

        int shown = 0;
        foreach (FormulaRemovalFileResult file in result.Files)
        {
            if (!file.Succeeded || remainingDetails == 0)
            {
                continue;
            }

            summary.Append('\n');
            summary.Append("• ");
            summary.Append(FormatOptionalPath(file.OutputPath));
            shown++;
            remainingDetails--;
        }

        AppendOmittedCount(summary, result.SucceededCount - shown);
    }

    private static void AppendFailures(
        StringBuilder summary,
        FormulaRemovalBatchResult result,
        ref int remainingDetails)
    {
        if (result.FailedCount == 0)
        {
            return;
        }

        AppendSectionBreak(summary);
        summary.Append("확인이 필요한 파일");

        int shown = 0;
        foreach (FormulaRemovalFileResult file in result.Files)
        {
            if (file.Succeeded || remainingDetails == 0)
            {
                continue;
            }

            summary.Append('\n');
            summary.Append("• ");
            summary.Append(FormatOptionalPath(file.InputPath));
            summary.Append(" — ");
            summary.Append(GetRecoveryGuidance(file.FailureKind));

            string detail = FormatMessage(file.ErrorMessage);
            if (detail.Length > 0)
            {
                summary.Append('\n');
                summary.Append("  세부 정보: ");
                summary.Append(detail);
            }

            shown++;
            remainingDetails--;
        }

        AppendOmittedCount(summary, result.FailedCount - shown);
    }

    private static void AppendWarnings(
        StringBuilder summary,
        FormulaRemovalBatchResult result,
        ref int remainingDetails)
    {
        long warningCount = 0;
        foreach (FormulaRemovalFileResult file in result.Files)
        {
            warningCount += file.Warnings?.Count ?? 0;
        }

        if (warningCount == 0)
        {
            return;
        }

        AppendSectionBreak(summary);
        summary.Append("주의 사항");

        int shown = 0;
        foreach (FormulaRemovalFileResult file in result.Files)
        {
            if (file.Warnings is null)
            {
                continue;
            }

            foreach (string warning in file.Warnings)
            {
                if (remainingDetails == 0)
                {
                    break;
                }

                summary.Append('\n');
                summary.Append("• ");
                summary.Append(FormatOptionalPath(file.InputPath));
                summary.Append(" — ");
                summary.Append(FormatWarning(warning));
                shown++;
                remainingDetails--;
            }
        }

        AppendOmittedCount(summary, warningCount - shown);
    }

    private static void AppendSectionBreak(StringBuilder summary)
    {
        summary.Append('\n');
        summary.Append('\n');
    }

    private static void AppendOmittedCount(StringBuilder summary, long omittedCount)
    {
        if (omittedCount <= 0)
        {
            return;
        }

        summary.Append('\n');
        summary.Append("• 그 외 ");
        summary.Append(omittedCount.ToString(CultureInfo.InvariantCulture));
        summary.Append('개');
    }

    private static string GetRecoveryGuidance(FormulaRemovalFailureKind failureKind) =>
        failureKind switch
        {
            FormulaRemovalFailureKind.InvalidInput =>
                "파일 경로와 .xlsx/.xlsm/.xlsb 확장자를 확인한 뒤 다시 추가하세요.",
            FormulaRemovalFailureKind.MissingCachedValue =>
                "Excel에서 수식을 다시 계산하고 저장한 뒤 다시 시도하세요.",
            FormulaRemovalFailureKind.FormulaRangeCrossesBoundary =>
                "시작 행을 1로 바꾸거나 수식 범위가 끝난 뒤의 행으로 조정하세요.",
            FormulaRemovalFailureKind.InvalidWorkbook =>
                "파일이 Excel에서 정상적으로 열리는지 확인하고 다시 저장한 뒤 시도하세요.",
            FormulaRemovalFailureKind.FileAccess =>
                "파일이 열려 있지 않은지와 파일·출력 폴더의 읽기/쓰기 권한을 확인하세요.",
            FormulaRemovalFailureKind.Unexpected =>
                "다시 시도하고, 문제가 계속되면 아래 세부 정보를 기록해 주세요.",
            _ => "아래 세부 정보를 확인한 뒤 다시 시도하세요.",
        };

    private static string FormatWarning(string? warning)
    {
        string formatted = FormatMessage(warning);
        if (formatted.Contains("manual calculation", StringComparison.OrdinalIgnoreCase))
        {
            return "통합 문서가 수동 계산 모드입니다. 저장된 값이 최신인지 Excel에서 확인하세요.";
        }

        return formatted.Length == 0 ? "추가 확인이 필요합니다." : formatted;
    }

    private static string FormatOptionalPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return "(경로 정보 없음)";
        }

        string singleLinePath = ReplaceLineBreaks(path.Trim());
        if (singleLinePath.Length <= MaximumPathLength)
        {
            return singleLinePath;
        }

        string fileName;
        try
        {
            fileName = Path.GetFileName(singleLinePath);
        }
        catch (ArgumentException)
        {
            fileName = string.Empty;
        }

        if (fileName.Length == 0 || fileName.Length >= MaximumPathLength - 1)
        {
            return Truncate(singleLinePath, MaximumPathLength);
        }

        int prefixLength = MaximumPathLength - fileName.Length - 1;
        return string.Concat(singleLinePath.AsSpan(0, prefixLength), "…", fileName);
    }

    private static string FormatMessage(string? message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return string.Empty;
        }

        StringBuilder singleLine = new(message.Length);
        bool previousWasWhitespace = false;
        foreach (char character in message.Trim())
        {
            if (char.IsWhiteSpace(character))
            {
                if (!previousWasWhitespace)
                {
                    singleLine.Append(' ');
                    previousWasWhitespace = true;
                }
            }
            else
            {
                singleLine.Append(character);
                previousWasWhitespace = false;
            }
        }

        return Truncate(singleLine.ToString(), MaximumMessageLength);
    }

    private static string ReplaceLineBreaks(string value) =>
        value.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');

    private static string Truncate(string value, int maximumLength)
    {
        if (value.Length <= maximumLength)
        {
            return value;
        }

        int contentLength = maximumLength - 1;
        if (contentLength > 0 && char.IsHighSurrogate(value[contentLength - 1]))
        {
            contentLength--;
        }

        return string.Concat(value.AsSpan(0, contentLength), "…");
    }

    private static string Invariant(int value) =>
        value.ToString(CultureInfo.InvariantCulture);
}
