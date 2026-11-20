using System.Xml;

namespace Sheetwright;

/// <summary>
/// Which cell styles of a workbook are date styles.
/// </summary>
/// <remarks>
/// A cell names its style by position in the <c>cellXfs</c> list of the style part, and that entry names a
/// number format, a built-in id or a custom one defined in <c>numFmts</c>. The answer is kept as one flag per
/// style so that reading a cell costs an array lookup.
/// </remarks>
internal static class SpreadsheetStyleDates
{
    // Excel itself allows 64,000 cell formats; a file with a million has no business being read.
    private const int MaxFormats = 1_048_576;

    internal static bool[] Load(Stream part, string partName)
    {
        Dictionary<int, bool> custom = [];
        List<int> formatIds = [];

        using XmlReader reader = SpreadsheetXml.Create(part);
        bool inCellXfs = false;
        int cellXfsDepth = -1;

        while (SpreadsheetXml.Read(reader, partName))
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.LocalName == "numFmt" && reader.Depth == 2)
                {
                    if (SpreadsheetXml.TryParseInt(reader.GetAttribute("numFmtId"), out int id)
                        && reader.GetAttribute("formatCode") is string code)
                    {
                        custom[id] = SpreadsheetDateFormats.IsDateFormatCode(code);
                    }
                }
                else if (reader.LocalName == "cellXfs" && reader.Depth == 1)
                {
                    inCellXfs = !reader.IsEmptyElement;
                    cellXfsDepth = reader.Depth;
                }
                else if (inCellXfs && reader.LocalName == "xf" && reader.Depth == cellXfsDepth + 1)
                {
                    if (formatIds.Count >= MaxFormats)
                    {
                        throw new InvalidDataException($"Part '{partName}' has more cell formats than any workbook has.");
                    }

                    formatIds.Add(SpreadsheetXml.TryParseInt(reader.GetAttribute("numFmtId"), out int formatId) ? formatId : 0);
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement && inCellXfs && reader.Depth == cellXfsDepth)
            {
                inCellXfs = false;
            }
        }

        bool[] flags = new bool[formatIds.Count];
        for (int index = 0; index < flags.Length; index++)
        {
            int formatId = formatIds[index];
            flags[index] = custom.TryGetValue(formatId, out bool isDate)
                ? isDate
                : SpreadsheetDateFormats.IsBuiltInDate(formatId);
        }

        return flags;
    }
}
