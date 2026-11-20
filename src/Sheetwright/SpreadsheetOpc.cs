using System.Xml;

namespace Sheetwright;

/// <summary>
/// The names of parts of the package and the relationships between them.
/// </summary>
/// <remarks>
/// A workbook does not say where its sheets are; it says <c>rId3</c>, and a relationships part says what that
/// is. The target is a path, relative to the part that names it unless it starts with a slash, in a file whose
/// author chose it. A path that climbs out of the package is refused rather than looked up under another name.
/// External relationships (a hyperlink, a linked workbook) point away from the package and are never followed.
/// </remarks>
internal static class SpreadsheetOpc
{
    internal const string OfficeDocumentType = "/officeDocument";
    internal const string WorksheetType = "/worksheet";
    internal const string SharedStringsType = "/sharedStrings";
    internal const string StylesType = "/styles";

    internal sealed record Relationship(string Id, string Type, string Target, bool External);

    /// <summary>A part name as the zip spells it, in the form lookups use: forward slashes, no leading slash.</summary>
    internal static string Normalise(string name)
    {
        return name.Replace('\\', '/').TrimStart('/');
    }

    /// <summary>The name of the relationships part that belongs to <paramref name="partName"/>.</summary>
    internal static string RelationshipsPartOf(string partName)
    {
        int slash = partName.LastIndexOf('/');
        return slash < 0
            ? "_rels/" + partName + ".rels"
            : string.Concat(partName.AsSpan(0, slash + 1), "_rels/", partName.AsSpan(slash + 1), ".rels");
    }

    /// <summary>The part a relationship target names, from the part that holds the relationship.</summary>
    /// <exception cref="InvalidDataException">The target leaves the package.</exception>
    internal static string Resolve(string fromPart, string target)
    {
        string path = target.Replace('\\', '/');
        List<string> segments = [];

        if (!path.StartsWith('/'))
        {
            int slash = fromPart.LastIndexOf('/');
            if (slash > 0)
            {
                segments.AddRange(fromPart[..slash].Split('/', StringSplitOptions.RemoveEmptyEntries));
            }
        }

        foreach (string segment in path.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".")
            {
                continue;
            }

            if (segment == "..")
            {
                if (segments.Count == 0)
                {
                    throw new InvalidDataException("A relationship points outside the package.");
                }

                segments.RemoveAt(segments.Count - 1);
                continue;
            }

            segments.Add(segment);
        }

        return string.Join('/', segments);
    }

    /// <summary>The relationships in a relationships part, in order.</summary>
    internal static List<Relationship> Read(Stream part, string partName, SpreadsheetReaderOptions options)
    {
        List<Relationship> relationships = [];
        using XmlReader reader = SpreadsheetXml.Create(part);

        while (SpreadsheetXml.Read(reader, partName))
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship")
            {
                string? id = reader.GetAttribute("Id");
                string? type = reader.GetAttribute("Type");
                string? target = reader.GetAttribute("Target");
                if (id is null || type is null || target is null)
                {
                    continue;
                }

                if (relationships.Count >= options.MaxParts)
                {
                    throw new SpreadsheetReadLimitException(
                        SpreadsheetReadLimit.MaxParts,
                        options.MaxParts,
                        $"part '{partName}' has more relationships than that.");
                }

                bool external = string.Equals(reader.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                relationships.Add(new Relationship(id, type, target, external));
            }
        }

        return relationships;
    }

    /// <summary>Whether the relationship's type ends with the given suffix (transitional and strict differ only before it).</summary>
    internal static bool IsType(Relationship relationship, string suffix)
    {
        return relationship.Type.EndsWith(suffix, StringComparison.Ordinal);
    }
}
