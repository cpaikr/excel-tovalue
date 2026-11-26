namespace XlClean.Core;

using System.IO.Packaging;

/// <summary>
/// Package navigation and record framing for Excel binary workbooks, shared by every operation.
/// </summary>
/// <remarks>
/// System.IO.Packaging reports malformed content types, part names, and relationship targets with
/// <see cref="ArgumentException"/>. Those describe the workbook, not the request, so callers wrap
/// their package work with <see cref="MalformedPackage"/>.
/// </remarks>
internal static class XlsbPackage
{
    public const string WorksheetRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    public const string TableRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

    private const int BufferSize = 1024 * 1024;
    private const int CancellationCheckInterval = 4096;

    /// <summary>BrtBundleSh, confirmed against Excel output.</summary>
    private const int BundleSheetRecordType = 156;

    private const string BinaryWorkbookContentType =
        "application/vnd.ms-excel.sheet.binary.macroEnabled.main";

    private const string OfficeDocumentRelationship =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    public static InvalidDataException MalformedPackage(ArgumentException exception) =>
        new($"The binary workbook package is malformed: {exception.Message}", exception);

    /// <summary>
    /// Reopens the package and re-reads the framing of every workbook, worksheet, and table record.
    /// </summary>
    public static void VerifyCanReopen(string workbookPath, CancellationToken cancellationToken)
    {
        try
        {
            using Package package = Package.Open(workbookPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            PackagePart workbookPart = GetWorkbookPart(package);
            List<Biff12Record> workbookRecords = ReadAllRecords(workbookPart, cancellationToken);

            foreach ((string relationshipId, _) in ReadSheets(workbookRecords))
            {
                PackagePart? sheetPart = GetRelatedPart(workbookPart, relationshipId, WorksheetRelationship);
                if (sheetPart is null)
                {
                    continue;
                }

                ReadAllRecords(sheetPart, cancellationToken);
                foreach (PackagePart tablePart in GetRelatedParts(sheetPart, TableRelationship))
                {
                    ReadAllRecords(tablePart, cancellationToken);
                }
            }
        }
        catch (ArgumentException exception)
        {
            throw MalformedPackage(exception);
        }
    }

    public static PackagePart GetWorkbookPart(Package package)
    {
        PackageRelationship relationship = package
            .GetRelationshipsByType(OfficeDocumentRelationship)
            .FirstOrDefault() ??
            throw new InvalidDataException("The package has no workbook part.");
        Uri partUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
        if (!package.PartExists(partUri))
        {
            throw new InvalidDataException("The package has no workbook part.");
        }

        PackagePart part = package.GetPart(partUri);
        if (!string.Equals(part.ContentType, BinaryWorkbookContentType, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidDataException(
                $"The .xlsb package contains a '{part.ContentType}' workbook, not an Excel binary workbook.");
        }

        return part;
    }

    /// <summary>Lists the sheets in workbook order from the workbook part's records.</summary>
    public static List<(string RelationshipId, string Name)> ReadSheets(IEnumerable<Biff12Record> workbookRecords)
    {
        List<(string RelationshipId, string Name)> sheets = [];
        foreach (Biff12Record record in workbookRecords)
        {
            if (record.Type == BundleSheetRecordType)
            {
                // BrtBundleSh: hsState (4), iTabID (4), strRelID, strName.
                string relationshipId = record.ReadWideString(8, out int nameOffset) ??
                    throw new InvalidDataException("A workbook sheet has no relationship id.");
                string name = record.ReadWideString(nameOffset, out _) ??
                    throw new InvalidDataException("A worksheet has no name.");
                sheets.Add((relationshipId, name));
            }
        }

        return sheets.Count > 0
            ? sheets
            : throw new InvalidDataException("The workbook has no sheets collection.");
    }

    /// <returns>The target part, or <see langword="null"/> when the relationship has another type.</returns>
    public static PackagePart? GetRelatedPart(PackagePart source, string relationshipId, string relationshipType)
    {
        if (!source.RelationshipExists(relationshipId))
        {
            throw new InvalidDataException($"Worksheet relationship '{relationshipId}' does not exist.");
        }

        PackageRelationship relationship = source.GetRelationship(relationshipId);
        if (relationship.RelationshipType != relationshipType ||
            relationship.TargetMode != TargetMode.Internal)
        {
            return null;
        }

        Uri partUri = PackUriHelper.ResolvePartUri(source.Uri, relationship.TargetUri);
        return source.Package.PartExists(partUri)
            ? source.Package.GetPart(partUri)
            : throw new InvalidDataException($"Worksheet relationship '{relationshipId}' has no target part.");
    }

    public static List<PackagePart> GetRelatedParts(PackagePart source, string relationshipType)
    {
        List<PackagePart> parts = [];
        foreach (PackageRelationship relationship in source.GetRelationshipsByType(relationshipType))
        {
            if (relationship.TargetMode != TargetMode.Internal)
            {
                continue;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(source.Uri, relationship.TargetUri);
            if (!source.Package.PartExists(partUri))
            {
                throw new InvalidDataException($"Relationship '{relationship.Id}' has no target part.");
            }

            parts.Add(source.Package.GetPart(partUri));
        }

        return parts;
    }

    public static void DeleteRelatedParts(Package package, PackagePart source, string relationshipType)
    {
        foreach (PackageRelationship relationship in source.GetRelationshipsByType(relationshipType).ToArray())
        {
            if (relationship.TargetMode == TargetMode.Internal)
            {
                Uri partUri = PackUriHelper.ResolvePartUri(source.Uri, relationship.TargetUri);
                if (package.PartExists(partUri))
                {
                    package.DeletePart(partUri);
                }
            }

            source.DeleteRelationship(relationship.Id);
        }
    }

    public static List<Biff12Record> ReadAllRecords(PackagePart part, CancellationToken cancellationToken)
    {
        List<Biff12Record> records = [];
        using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
        Biff12Reader reader = new(new BufferedStream(stream, BufferSize));
        while (reader.TryRead(out Biff12Record record))
        {
            if (records.Count % CancellationCheckInterval == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            records.Add(record);
        }

        return records;
    }
}
