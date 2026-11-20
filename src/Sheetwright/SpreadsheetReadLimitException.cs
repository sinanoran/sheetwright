namespace Sheetwright;

/// <summary>
/// A file is bigger than <see cref="SpreadsheetReaderOptions"/> allow. The reader stopped; nothing was
/// truncated.
/// </summary>
/// <remarks>
/// A file that is malformed is an <see cref="InvalidDataException"/>; one that is well formed and too big is
/// this, because a caller says something different to a person in the two cases ("this is not a spreadsheet"
/// and "this spreadsheet is bigger than we accept") and reads <see cref="Limit"/> to say which limit. The
/// message names the limit and its value and never contains a cell of the file.
/// </remarks>
public sealed class SpreadsheetReadLimitException : Exception
{
    /// <summary>Creates the exception for <paramref name="limit"/> set to <paramref name="maximum"/>.</summary>
    public SpreadsheetReadLimitException(SpreadsheetReadLimit limit, long maximum, string detail)
        : base(Describe(limit, maximum, detail))
    {
        Limit = limit;
        Maximum = maximum;
    }

    /// <summary>The limit that was passed.</summary>
    public SpreadsheetReadLimit Limit { get; }

    /// <summary>The value the limit was set to.</summary>
    public long Maximum { get; }

    private static string Describe(SpreadsheetReadLimit limit, long maximum, string detail)
    {
        return $"The file is over the {limit} limit of {maximum.ToString(System.Globalization.CultureInfo.InvariantCulture)}"
            + $" (SpreadsheetReaderOptions.{limit}): {detail}";
    }
}
