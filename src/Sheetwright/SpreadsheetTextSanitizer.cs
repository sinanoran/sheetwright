using System.Text;
using System.Xml;

namespace Sheetwright;

internal static class SpreadsheetTextSanitizer
{
    internal static string? SanitizeNullable(string? value)
    {
        return value is null ? null : Sanitize(value);
    }

    internal static string Sanitize(string value)
    {
        StringBuilder? sanitized = null;

        for (int index = 0; index < value.Length; index++)
        {
            char character = value[index];
            if (char.IsHighSurrogate(character)
                && index + 1 < value.Length
                && char.IsLowSurrogate(value[index + 1]))
            {
                if (sanitized is not null)
                {
                    sanitized.Append(character);
                    sanitized.Append(value[++index]);
                }
                else
                {
                    index++;
                }

                continue;
            }

            if (XmlConvert.IsXmlChar(character))
            {
                sanitized?.Append(character);
                continue;
            }

            if (sanitized is null)
            {
                sanitized = new StringBuilder(value.Length);
                sanitized.Append(value, 0, index);
            }
        }

        return sanitized?.ToString() ?? value;
    }
}
