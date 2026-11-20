using System.Xml;
using System.Text;

namespace Sheetwright;

/// <summary>
/// The shared-string table of a workbook, loaded within the limits.
/// </summary>
/// <remarks>
/// A cell that holds text usually holds a number: the position of the text in this table, which stores each
/// distinct string once. It has to be held to be looked up, so it is the one thing the reader keeps in memory
/// whole, and it is where a file can name millions of strings. Both counts are limited as the table is read,
/// the entries and the characters, so a flood is refused at the entry that crosses the line and not after the
/// memory is spent.
/// </remarks>
internal static class SpreadsheetSharedStrings
{
    internal static string[] Load(Stream part, string partName, SpreadsheetReaderOptions options)
    {
        List<string> strings = [];
        long characters = 0;
        StringBuilder text = new();

        using XmlReader reader = SpreadsheetXml.Create(part);
        while (SpreadsheetXml.Read(reader, partName))
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "si" || reader.Depth != 1)
            {
                continue;
            }

            if (strings.Count >= options.MaxSharedStrings)
            {
                throw new SpreadsheetReadLimitException(
                    SpreadsheetReadLimit.MaxSharedStrings,
                    options.MaxSharedStrings,
                    "the shared-string table has more entries than that.");
            }

            text.Clear();
            SpreadsheetXml.ReadStringItem(reader, partName, text, options, "a shared string");

            characters += text.Length;
            if (characters > options.MaxSharedStringCharacters)
            {
                throw new SpreadsheetReadLimitException(
                    SpreadsheetReadLimit.MaxSharedStringCharacters,
                    options.MaxSharedStringCharacters,
                    "the shared-string table has more characters than that.");
            }

            strings.Add(text.ToString());
        }

        return [.. strings];
    }
}
