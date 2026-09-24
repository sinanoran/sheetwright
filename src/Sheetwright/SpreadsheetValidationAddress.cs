using System.Text;

namespace Sheetwright;

/// <summary>
/// The cells a data validation applies to.
/// </summary>
/// <remarks>
/// The address is parsed rather than carried through as text. It used to be
/// copied straight into <c>sqref</c>, which meant a caller who wrote
/// <c>"Sheet1!A1"</c> — the form every other API here accepts — put a sheet
/// qualifier somewhere the schema does not allow one, and a caller who mistyped
/// it entirely found out when Excel refused the file. Both now fail on the line
/// that made the mistake.
/// </remarks>
public sealed class SpreadsheetValidationAddress
{
    internal SpreadsheetValidationAddress(string address)
    {
        IReadOnlyList<string> parts = Split(address);

        if (parts.Count == 0)
        {
            throw new ArgumentException("A validation address is required.", nameof(address));
        }

        Address = string.Join(' ', parts.Select(part => SpreadsheetAddress.Parse(part).Reference));
    }

    public string Address { get; }

    /// <summary>
    /// The rectangles of a <c>sqref</c>, which separates them with spaces.
    /// </summary>
    /// <remarks>
    /// Worth supporting: a list of rectangles is how one dropdown reaches columns
    /// B and E without the validation being declared twice.
    ///
    /// The quote tracking is not decoration. A sheet name may contain a space,
    /// and <c>'Some Sheet'!A1</c> split naively becomes two fragments that are
    /// each an invalid address — so the qualifier this class exists to strip
    /// would have been the one input it could not read.
    /// </remarks>
    private static IReadOnlyList<string> Split(string address)
    {
        List<string> parts = [];
        StringBuilder current = new();
        bool quoted = false;

        foreach (char character in address)
        {
            if (character == '\'')
            {
                quoted = !quoted;
                current.Append(character);
                continue;
            }

            if (!quoted && char.IsWhiteSpace(character))
            {
                if (current.Length > 0)
                {
                    parts.Add(current.ToString());
                    current.Clear();
                }

                continue;
            }

            current.Append(character);
        }

        if (current.Length > 0)
        {
            parts.Add(current.ToString());
        }

        return parts;
    }
}
