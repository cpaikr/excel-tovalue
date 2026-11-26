namespace XlClean.Core;

using System.Diagnostics.CodeAnalysis;
using System.IO.Packaging;
using System.Security;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

/// <summary>The container format of a supported workbook.</summary>
internal enum WorkbookFormat
{
    /// <summary>An OpenXML spreadsheet package: .xlsx or .xlsm.</summary>
    OpenXml,

    /// <summary>An Excel binary workbook package: .xlsb.</summary>
    Binary,
}

/// <summary>
/// The staged copy an operation transforms, and the context it reports progress against.
/// </summary>
/// <param name="Path">The temporary copy. Operations change only this file.</param>
/// <param name="InputPath">The resolved source workbook, for read-only comparison.</param>
/// <param name="RequestedInputPath">The input path exactly as requested, for messages and progress.</param>
/// <param name="Format">The container format, derived from the input extension.</param>
/// <param name="FileIndex">Zero-based position of the file in the batch.</param>
/// <param name="TotalFiles">Number of files in the batch.</param>
/// <param name="Progress">Optional progress sink.</param>
internal sealed record WorkbookFile(
    string Path,
    string InputPath,
    string RequestedInputPath,
    WorkbookFormat Format,
    int FileIndex,
    int TotalFiles,
    IProgress<WorkbookProgress>? Progress)
{
    public void ReportWorksheet(string worksheetName) =>
        Progress?.Report(new WorkbookProgress(
            FileIndex,
            TotalFiles,
            RequestedInputPath,
            worksheetName,
            WorkbookStage.ProcessingWorksheet));
}

/// <summary>
/// One workbook transformation run through <see cref="WorkbookBatchPipeline"/>.
/// </summary>
/// <typeparam name="TChanges">What the operation reports for a published workbook.</typeparam>
/// <remarks>
/// An operation reports a file-level failure of its own by throwing a
/// <see cref="WorkbookFailureException"/>. Every other exception is classified by the pipeline;
/// for a binary workbook an <see cref="ArgumentException"/> from System.IO.Packaging describes
/// the package, so the pipeline reports it as an invalid workbook rather than invalid input.
/// </remarks>
internal interface IWorkbookOperation<TChanges>
    where TChanges : class
{
    /// <summary>Gets the text appended to the input file name, such as <c>_수식제거</c>.</summary>
    string OutputNameSuffix { get; }

    /// <summary>Transforms the staged copy in place.</summary>
    TChanges Transform(WorkbookFile file, CancellationToken cancellationToken);

    /// <summary>
    /// Checks the transformed copy beyond the pipeline's reopen check. Throwing fails the file
    /// and nothing is published.
    /// </summary>
    void Verify(WorkbookFile file, TChanges changes, CancellationToken cancellationToken);
}

/// <summary>The outcome for one requested workbook.</summary>
/// <param name="Changes">The operation's report; null exactly when the file failed.</param>
internal sealed record WorkbookFileOutcome<TChanges>(
    string InputPath,
    string? OutputPath,
    WorkbookFailureKind FailureKind,
    string? ErrorMessage,
    TChanges? Changes)
    where TChanges : class;

/// <summary>Cancellation that keeps the outcomes of files published before it.</summary>
[SuppressMessage(
    "Design",
    "CA1032:Implement standard exception constructors",
    Justification = "Only the pipeline creates this result-carrying cancellation exception.")]
internal sealed class WorkbookBatchCanceledException<TChanges>(
    IReadOnlyList<WorkbookFileOutcome<TChanges>> completed,
    Exception innerException,
    CancellationToken cancellationToken)
    : OperationCanceledException("The workbook batch was canceled.", innerException, cancellationToken)
    where TChanges : class
{
    public IReadOnlyList<WorkbookFileOutcome<TChanges>> Completed { get; } = completed;
}

/// <summary>
/// Runs one operation over a batch of workbooks with the safety contract every operation shares.
/// </summary>
/// <remarks>
/// An input workbook is never modified and an existing file is never overwritten. Each file is
/// copied to a uniquely named sibling of its output, transformed and verified there, and published
/// with a same-volume move only after the package reopens and the operation's own check passes.
/// Any failure removes the staged copy, so a failed file leaves no partial output, and the batch
/// continues with the next file.
/// </remarks>
internal static class WorkbookBatchPipeline
{
    private const int CopyBufferSize = 1024 * 1024;

    /// <summary>Gets the comparer that matches how the host file system compares paths.</summary>
    public static StringComparer PathComparer { get; } =
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparer.OrdinalIgnoreCase
            : StringComparer.Ordinal;

    /// <summary>Processes every input in order on the calling thread.</summary>
    /// <returns>One outcome per requested input, in request order.</returns>
    /// <exception cref="IOException">The output directory cannot be created.</exception>
    /// <exception cref="UnauthorizedAccessException">The output directory cannot be accessed.</exception>
    /// <exception cref="WorkbookBatchCanceledException{TChanges}">
    /// Cancellation was observed after the batch started.
    /// </exception>
    /// <exception cref="OperationCanceledException">Cancellation was requested before the batch started.</exception>
    public static IReadOnlyList<WorkbookFileOutcome<TChanges>> Run<TChanges>(
        IReadOnlyList<string> inputPaths,
        string outputDirectory,
        IWorkbookOperation<TChanges> operation,
        IProgress<WorkbookProgress>? progress,
        CancellationToken cancellationToken)
        where TChanges : class
    {
        ArgumentNullException.ThrowIfNull(inputPaths);
        ArgumentNullException.ThrowIfNull(operation);
        cancellationToken.ThrowIfCancellationRequested();

        Directory.CreateDirectory(outputDirectory);

        HashSet<string> seenInputs = new(PathComparer);
        OutputPathAllocator outputPaths = new(
            outputDirectory,
            operation.OutputNameSuffix,
            GetNormalizedInputPaths(inputPaths));
        List<WorkbookFileOutcome<TChanges>> outcomes = new(inputPaths.Count);

        try
        {
            for (int index = 0; index < inputPaths.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                string inputPath = inputPaths[index];
                progress?.Report(new WorkbookProgress(
                    index,
                    inputPaths.Count,
                    inputPath,
                    null,
                    WorkbookStage.StartingFile));

                outcomes.Add(ProcessFile(
                    inputPath,
                    index,
                    inputPaths.Count,
                    operation,
                    outputPaths,
                    seenInputs,
                    progress,
                    cancellationToken));

                progress?.Report(new WorkbookProgress(
                    index + 1,
                    inputPaths.Count,
                    inputPath,
                    null,
                    WorkbookStage.CompletedFile));
            }
        }
        catch (OperationCanceledException exception) when (cancellationToken.IsCancellationRequested)
        {
            throw new WorkbookBatchCanceledException<TChanges>(
                outcomes.AsReadOnly(),
                exception,
                cancellationToken);
        }

        return outcomes.AsReadOnly();
    }

    /// <summary>Deletes a staging file the pipeline or an operation created, ignoring locks.</summary>
    public static void TryDeleteStagingFile(string? path)
    {
        if (string.IsNullOrEmpty(path))
        {
            return;
        }

        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
            // The uniquely named staging file is never a user file. A later
            // process or antivirus lock may delay its removal.
        }
        catch (UnauthorizedAccessException)
        {
            // See the IOException rationale above.
        }
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "A batch must isolate unexpected failures to the affected workbook and continue safely.")]
    private static WorkbookFileOutcome<TChanges> ProcessFile<TChanges>(
        string requestedInputPath,
        int index,
        int totalFiles,
        IWorkbookOperation<TChanges> operation,
        OutputPathAllocator outputPaths,
        HashSet<string> seenInputs,
        IProgress<WorkbookProgress>? progress,
        CancellationToken cancellationToken)
        where TChanges : class
    {
        string? temporaryPath = null;

        try
        {
            string inputPath = Path.GetFullPath(requestedInputPath);
            WorkbookFormat format = GetFormat(inputPath);
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException("The input workbook does not exist.", inputPath);
            }

            if (!seenInputs.Add(inputPath))
            {
                return Failure<TChanges>(
                    requestedInputPath,
                    WorkbookFailureKind.InvalidInput,
                    "The same input workbook was requested more than once.");
            }

            string outputPath = outputPaths.Allocate(inputPath);
            temporaryPath = CreateTemporaryPath(outputPath);
            CopyWithCancellation(inputPath, temporaryPath, cancellationToken);

            WorkbookFile file = new(
                temporaryPath,
                inputPath,
                requestedInputPath,
                format,
                index,
                totalFiles,
                progress);
            TChanges changes = RunPackageStep(format, () => operation.Transform(file, cancellationToken));

            cancellationToken.ThrowIfCancellationRequested();
            VerifyCanReopen(temporaryPath, format, cancellationToken);
            RunPackageStep(format, () =>
            {
                operation.Verify(file, changes, cancellationToken);
                return changes;
            });
            cancellationToken.ThrowIfCancellationRequested();

            outputPath = PublishWithUniqueName(
                temporaryPath,
                outputPath,
                inputPath,
                outputPaths,
                cancellationToken);
            temporaryPath = null;

            return new WorkbookFileOutcome<TChanges>(
                requestedInputPath,
                outputPath,
                WorkbookFailureKind.None,
                null,
                changes);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (WorkbookFailureException exception)
        {
            return Failure<TChanges>(requestedInputPath, exception.Kind, exception.Message);
        }
        catch (Exception exception) when (
            exception is ArgumentException or FileNotFoundException or NotSupportedException or
                PathTooLongException or SecurityException)
        {
            return Failure<TChanges>(requestedInputPath, WorkbookFailureKind.InvalidInput, exception.Message);
        }
        catch (Exception exception) when (
            exception is FileFormatException or InvalidDataException or OpenXmlPackageException or
                XmlException or FormatException or InvalidOperationException or OverflowException)
        {
            return Failure<TChanges>(requestedInputPath, WorkbookFailureKind.InvalidWorkbook, exception.Message);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return Failure<TChanges>(requestedInputPath, WorkbookFailureKind.FileAccess, exception.Message);
        }
        catch (Exception exception)
        {
            return Failure<TChanges>(requestedInputPath, WorkbookFailureKind.Unexpected, exception.Message);
        }
        finally
        {
            TryDeleteStagingFile(temporaryPath);
        }
    }

    private static T RunPackageStep<T>(WorkbookFormat format, Func<T> step)
    {
        try
        {
            return step();
        }
        catch (ArgumentException exception) when (format == WorkbookFormat.Binary)
        {
            throw XlsbPackage.MalformedPackage(exception);
        }
    }

    private static WorkbookFormat GetFormat(string inputPath)
    {
        string extension = Path.GetExtension(inputPath);
        if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            return WorkbookFormat.OpenXml;
        }

        if (extension.Equals(".xlsb", StringComparison.OrdinalIgnoreCase))
        {
            return WorkbookFormat.Binary;
        }

        throw new ArgumentException(
            $"Unsupported workbook extension '{extension}'. Use .xlsx, .xlsm, or .xlsb.",
            nameof(inputPath));
    }

    private static void VerifyCanReopen(string workbookPath, WorkbookFormat format, CancellationToken cancellationToken)
    {
        if (format == WorkbookFormat.Binary)
        {
            XlsbPackage.VerifyCanReopen(workbookPath, cancellationToken);
        }
        else
        {
            OpenXmlWorkbook.VerifyCanReopen(workbookPath);
        }
    }

    private static HashSet<string> GetNormalizedInputPaths(IReadOnlyList<string> requestedPaths)
    {
        HashSet<string> paths = new(PathComparer);
        foreach (string requestedPath in requestedPaths)
        {
            try
            {
                paths.Add(Path.GetFullPath(requestedPath));
            }
            catch (Exception exception) when (
                exception is ArgumentException or NotSupportedException or PathTooLongException or SecurityException)
            {
                // The per-file attempt returns a precise invalid-input result.
            }
        }

        return paths;
    }

    private static void CopyWithCancellation(
        string inputPath,
        string outputPath,
        CancellationToken cancellationToken)
    {
        using FileStream input = new(
            inputPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            CopyBufferSize,
            FileOptions.SequentialScan);
        using FileStream output = new(
            outputPath,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            CopyBufferSize,
            FileOptions.SequentialScan);

        byte[] buffer = new byte[CopyBufferSize];
        int bytesRead;
        while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            output.Write(buffer, 0, bytesRead);
        }

        output.Flush(flushToDisk: true);
    }

    private static string CreateTemporaryPath(string outputPath)
    {
        string directory = Path.GetDirectoryName(outputPath)!;
        string fileName = Path.GetFileNameWithoutExtension(outputPath);
        string extension = Path.GetExtension(outputPath);
        return Path.Combine(
            directory,
            $".xlclean-{fileName}-{Guid.NewGuid():N}.tmp{extension}");
    }

    private static string PublishWithUniqueName(
        string temporaryPath,
        string initialOutputPath,
        string inputPath,
        OutputPathAllocator outputPaths,
        CancellationToken cancellationToken)
    {
        string outputPath = initialOutputPath;

        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                File.Move(temporaryPath, outputPath, overwrite: false);
                return outputPath;
            }
            catch (IOException) when (Path.Exists(outputPath))
            {
                outputPath = outputPaths.Allocate(inputPath);
            }
        }
    }

    private static WorkbookFileOutcome<TChanges> Failure<TChanges>(
        string inputPath,
        WorkbookFailureKind failureKind,
        string message)
        where TChanges : class =>
        new(inputPath, null, failureKind, message, null);

    /// <summary>
    /// Hands out output names that collide with no existing file, no input of the batch, and no
    /// name already handed out: <c>Book_suffix.xlsx</c>, then <c>Book_suffix (2).xlsx</c>, and so on.
    /// </summary>
    private sealed class OutputPathAllocator(
        string outputDirectory,
        string suffix,
        IEnumerable<string> inputPaths)
    {
        private readonly HashSet<string> reservedPaths = new(inputPaths, PathComparer);

        public string Allocate(string inputPath)
        {
            string baseName = $"{Path.GetFileNameWithoutExtension(inputPath)}{suffix}";
            string extension = Path.GetExtension(inputPath);

            for (int sequence = 1; sequence < int.MaxValue; sequence++)
            {
                string sequenceSuffix = sequence == 1 ? string.Empty : $" ({sequence})";
                string candidate = Path.Combine(outputDirectory, $"{baseName}{sequenceSuffix}{extension}");

                if (!Path.Exists(candidate) && reservedPaths.Add(candidate))
                {
                    return candidate;
                }
            }

            throw new IOException($"No available output name could be allocated for '{inputPath}'.");
        }
    }
}
