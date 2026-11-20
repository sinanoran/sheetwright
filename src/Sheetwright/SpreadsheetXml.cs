using System.Buffers;
using System.Globalization;
using System.Text;
using System.Xml;

namespace Sheetwright;

/// <summary>
/// How the workbook reader reads XML it did not write.
/// </summary>
/// <remarks>
/// Nothing here resolves anything: no DTD (so no entity expansion), no external resource, no schema. A part
/// that has a document type declaration is refused. Names are matched without their namespace, because a
/// workbook is written in the transitional namespaces or, from some producers, in the strict ones, and
/// everything else about the two is the same to a reader that wants the cells.
/// </remarks>
internal static class SpreadsheetXml
{
    /// <summary>Spreadsheet XML is a few levels deep; this is two orders of magnitude of room.</summary>
    internal const int MaxDepth = 64;

    private static readonly XmlReaderSettings Settings = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        CloseInput = true,
    };

    internal static XmlReader Create(Stream stream) => XmlReader.Create(stream, Settings);

    /// <summary>
    /// <see cref="XmlReader.Read"/>, with malformed XML and runaway nesting reported as the failure they are.
    /// </summary>
    internal static bool Read(XmlReader reader, string partName)
    {
        try
        {
            bool more = reader.Read();
            if (more && reader.Depth > MaxDepth)
            {
                throw new InvalidDataException($"Part '{partName}' is nested deeper than a workbook ever is.");
            }

            return more;
        }
        catch (XmlException exception)
        {
            throw new InvalidDataException($"Part '{partName}' is not well-formed XML.", exception);
        }
    }

    /// <summary>
    /// The text of the element the reader is on, appended to <paramref name="text"/>, read in pieces so that a
    /// cell over the limit is refused before it is held.
    /// </summary>
    /// <remarks>Leaves the reader on the element's end tag (or on the element, when it is empty).</remarks>
    internal static void ReadText(
        XmlReader reader,
        string partName,
        StringBuilder text,
        SpreadsheetReaderOptions options,
        string where)
    {
        if (reader.IsEmptyElement)
        {
            return;
        }

        int depth = reader.Depth;
        char[] chunk = ArrayPool<char>.Shared.Rent(2048);
        try
        {
            while (Read(reader, partName))
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.Whitespace:
                    case XmlNodeType.SignificantWhitespace:
                        int read;
                        while ((read = reader.ReadValueChunk(chunk, 0, chunk.Length)) > 0)
                        {
                            if (text.Length + read > options.MaxCellCharacters)
                            {
                                throw new SpreadsheetReadLimitException(
                                    SpreadsheetReadLimit.MaxCellCharacters,
                                    options.MaxCellCharacters,
                                    $"{where} has more characters than that.");
                            }

                            text.Append(chunk, 0, read);
                        }

                        break;

                    case XmlNodeType.EndElement when reader.Depth == depth:
                        return;
                }
            }
        }
        finally
        {
            ArrayPool<char>.Shared.Return(chunk);
        }
    }

    /// <summary>
    /// The text of a string item (<c>si</c> in the shared strings, <c>is</c> in a cell): the <c>t</c> elements
    /// in order, whether the item is plain or built of runs, without the phonetic guide of an East Asian run.
    /// </summary>
    /// <remarks>Leaves the reader on the item's end tag.</remarks>
    internal static void ReadStringItem(
        XmlReader reader,
        string partName,
        StringBuilder text,
        SpreadsheetReaderOptions options,
        string where)
    {
        if (reader.IsEmptyElement)
        {
            return;
        }

        int depth = reader.Depth;
        int phoneticDepth = -1;

        while (Read(reader, partName))
        {
            if (phoneticDepth >= 0)
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == phoneticDepth)
                {
                    phoneticDepth = -1;
                }

                continue;
            }

            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.LocalName == "rPh")
                {
                    if (!reader.IsEmptyElement)
                    {
                        phoneticDepth = reader.Depth;
                    }
                }
                else if (reader.LocalName == "t")
                {
                    ReadText(reader, partName, text, options, where);
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)
            {
                return;
            }
        }
    }

    /// <summary>An attribute with no namespace, or <see langword="null"/>.</summary>
    internal static string? Attribute(XmlReader reader, string name) => reader.GetAttribute(name);

    /// <summary>
    /// The <c>id</c> attribute that refers to a relationship (<c>r:id</c>), in either the transitional or the
    /// strict namespace.
    /// </summary>
    internal static string? RelationshipId(XmlReader reader)
    {
        if (!reader.HasAttributes)
        {
            return null;
        }

        string? found = null;
        if (reader.MoveToFirstAttribute())
        {
            do
            {
                if (reader.LocalName == "id"
                    && reader.NamespaceURI.EndsWith("/relationships", StringComparison.Ordinal))
                {
                    found = reader.Value;
                    break;
                }
            }
            while (reader.MoveToNextAttribute());

            reader.MoveToElement();
        }

        return found;
    }

    /// <summary>An attribute that is a switch, written 1, 0, true or false.</summary>
    internal static bool Flag(string? value) => value is "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);

    internal static bool TryParseInt(string? value, out int number)
    {
        return int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out number);
    }
}
