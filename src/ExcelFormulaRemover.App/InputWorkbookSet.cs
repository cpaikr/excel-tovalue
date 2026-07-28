using System.Collections.ObjectModel;
using System.Security;

namespace ExcelFormulaRemover.App;

/// <summary>
/// Reports how a group of candidate workbook paths changed the input set.
/// </summary>
internal readonly record struct InputAdditionResult(
    int AddedCount,
    int DuplicateCount,
    int RejectedCount)
{
    /// <summary>
    /// Gets the total number of candidates examined.
    /// </summary>
    public int CandidateCount => AddedCount + DuplicateCount + RejectedCount;
}

/// <summary>
/// Maintains the ordered, validated workbook paths selected in the native app.
/// </summary>
internal sealed class InputWorkbookSet
{
    private static readonly StringComparer PathComparer = StringComparer.OrdinalIgnoreCase;
    private readonly List<string> paths = [];
    private readonly HashSet<string> normalizedPaths = new(PathComparer);

    /// <summary>
    /// Gets the number of selected workbooks.
    /// </summary>
    public int Count => paths.Count;

    /// <summary>
    /// Gets a detached, read-only snapshot in insertion order.
    /// </summary>
    public IReadOnlyList<string> Items =>
        new ReadOnlyCollection<string>(paths.ToArray());

    /// <summary>
    /// Adds existing OpenXML workbooks while retaining the first occurrence of each path.
    /// </summary>
    /// <param name="candidatePaths">Paths to classify in insertion order.</param>
    /// <returns>Counts for newly added, duplicate, and rejected candidates.</returns>
    public InputAdditionResult Add(IEnumerable<string> candidatePaths)
    {
        ArgumentNullException.ThrowIfNull(candidatePaths);

        int addedCount = 0;
        int duplicateCount = 0;
        int rejectedCount = 0;

        foreach (string? candidatePath in candidatePaths)
        {
            if (!TryNormalizeWorkbookPath(candidatePath, out string normalizedPath))
            {
                rejectedCount++;
                continue;
            }

            if (normalizedPaths.Contains(normalizedPath))
            {
                duplicateCount++;
                continue;
            }

            if (!File.Exists(normalizedPath))
            {
                rejectedCount++;
                continue;
            }

            normalizedPaths.Add(normalizedPath);
            paths.Add(normalizedPath);
            addedCount++;
        }

        return new InputAdditionResult(addedCount, duplicateCount, rejectedCount);
    }

    /// <summary>
    /// Removes a selected path without requiring the file to still exist on disk.
    /// </summary>
    /// <param name="candidatePaths">Relative or absolute paths identifying selected items.</param>
    /// <returns>The number of items removed.</returns>
    public int Remove(IEnumerable<string> candidatePaths)
    {
        ArgumentNullException.ThrowIfNull(candidatePaths);

        int removedCount = 0;
        foreach (string? candidatePath in candidatePaths)
        {
            if (!TryNormalizePath(candidatePath, out string normalizedPath) ||
                !normalizedPaths.Contains(normalizedPath))
            {
                continue;
            }

            int index = paths.FindIndex(candidate => PathComparer.Equals(candidate, normalizedPath));
            if (index < 0)
            {
                continue;
            }

            paths.RemoveAt(index);
            normalizedPaths.Remove(normalizedPath);
            removedCount++;
        }

        return removedCount;
    }

    /// <summary>
    /// Removes every selected workbook.
    /// </summary>
    public void Clear()
    {
        paths.Clear();
        normalizedPaths.Clear();
    }

    private static bool TryNormalizeWorkbookPath(string? path, out string normalizedPath)
    {
        if (!TryNormalizePath(path, out normalizedPath))
        {
            return false;
        }

        string extension = Path.GetExtension(normalizedPath);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryNormalizePath(string? path, out string normalizedPath)
    {
        normalizedPath = string.Empty;
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        try
        {
            normalizedPath = Path.GetFullPath(path);
            return true;
        }
        catch (Exception exception) when (
            exception is ArgumentException or NotSupportedException or PathTooLongException or
                SecurityException)
        {
            return false;
        }
    }
}
