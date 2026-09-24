using System.Globalization;
using System.Text.RegularExpressions;

namespace Sheetwright;

internal readonly struct SpreadsheetAddress
{
    /// <summary>The last row of an Excel worksheet.</summary>
    internal const int MaximumRow = 1_048_576;

    /// <summary>The last column of an Excel worksheet, XFD.</summary>
    internal const int MaximumColumn = 16_384;

    private static readonly Regex CellRangePattern = new(
        @"^\$?(?<fromColumn>[A-Z]+)\$?(?<fromRow>\d+)(?::\$?(?<toColumn>[A-Z]+)\$?(?<toRow>\d+))?$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex ColumnRangePattern = new(
        @"^\$?(?<fromColumn>[A-Z]+):\$?(?<toColumn>[A-Z]+)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex RowRangePattern = new(
        @"^\$?(?<fromRow>\d+):\$?(?<toRow>\d+)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <remarks>
    /// The bounds are checked here rather than trusted, because nothing further
    /// down the pipeline checks them and nothing complains. A row index of zero
    /// travels all the way through and is written out as
    /// <c>&lt;row r="0"&gt;&lt;c r="A0"&gt;</c>, which is not a row Excel has: the
    /// caller gets a package that was produced without error and cannot be
    /// opened. The same is true past XFD and past row 1,048,576.
    ///
    /// An inverted range — C1:A1 — is the quieter version of the same thing.
    /// Every loop over it runs zero times, so styling it, merging it or filling
    /// it does nothing whatever and reports success.
    /// </remarks>
    internal SpreadsheetAddress(int fromRow, int fromColumn, int toRow, int toColumn, bool wholeColumn = false, bool wholeRow = false)
    {
        ThrowIfOutsideTheGrid(fromRow, nameof(fromRow), MaximumRow);
        ThrowIfOutsideTheGrid(toRow, nameof(toRow), MaximumRow);
        ThrowIfOutsideTheGrid(fromColumn, nameof(fromColumn), MaximumColumn);
        ThrowIfOutsideTheGrid(toColumn, nameof(toColumn), MaximumColumn);

        if (toRow < fromRow)
        {
            throw new ArgumentOutOfRangeException(
                nameof(toRow),
                toRow,
                $"A range cannot end at row {toRow.ToString(CultureInfo.InvariantCulture)} and begin at row {fromRow.ToString(CultureInfo.InvariantCulture)}.");
        }

        if (toColumn < fromColumn)
        {
            throw new ArgumentOutOfRangeException(
                nameof(toColumn),
                toColumn,
                $"A range cannot end at column {toColumn.ToString(CultureInfo.InvariantCulture)} and begin at column {fromColumn.ToString(CultureInfo.InvariantCulture)}.");
        }

        FromRow = fromRow;
        FromColumn = fromColumn;
        ToRow = toRow;
        ToColumn = toColumn;
        IsWholeColumn = wholeColumn;
        IsWholeRow = wholeRow;
    }

    internal int FromRow { get; }

    internal int FromColumn { get; }

    internal int ToRow { get; }

    internal int ToColumn { get; }

    internal bool IsWholeColumn { get; }

    internal bool IsWholeRow { get; }

    internal string Reference
    {
        get
        {
            if (IsWholeColumn)
            {
                return $"{GetColumnName(FromColumn)}:{GetColumnName(ToColumn)}";
            }

            if (IsWholeRow)
            {
                return $"{FromRow.ToString(CultureInfo.InvariantCulture)}:{ToRow.ToString(CultureInfo.InvariantCulture)}";
            }

            string from = GetCellReference(FromRow, FromColumn);
            string to = GetCellReference(ToRow, ToColumn);
            return string.Equals(from, to, StringComparison.Ordinal) ? from : $"{from}:{to}";
        }
    }

    internal static SpreadsheetAddress Parse(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
        {
            throw new ArgumentException("A spreadsheet address is required.", nameof(address));
        }

        string normalized = address.Trim();
        int bangIndex = normalized.LastIndexOf('!');
        if (bangIndex >= 0)
        {
            normalized = normalized.Substring(bangIndex + 1);
        }

        Match match = CellRangePattern.Match(normalized);
        if (match.Success)
        {
            int fromRow = int.Parse(match.Groups["fromRow"].Value, CultureInfo.InvariantCulture);
            int fromColumn = GetColumnNumber(match.Groups["fromColumn"].Value);
            int toRow = match.Groups["toRow"].Success
                ? int.Parse(match.Groups["toRow"].Value, CultureInfo.InvariantCulture)
                : fromRow;
            int toColumn = match.Groups["toColumn"].Success
                ? GetColumnNumber(match.Groups["toColumn"].Value)
                : fromColumn;

            return new SpreadsheetAddress(fromRow, fromColumn, toRow, toColumn);
        }

        match = ColumnRangePattern.Match(normalized);
        if (match.Success)
        {
            return new SpreadsheetAddress(
                1,
                GetColumnNumber(match.Groups["fromColumn"].Value),
                1_048_576,
                GetColumnNumber(match.Groups["toColumn"].Value),
                wholeColumn: true);
        }

        match = RowRangePattern.Match(normalized);
        if (match.Success)
        {
            return new SpreadsheetAddress(
                int.Parse(match.Groups["fromRow"].Value, CultureInfo.InvariantCulture),
                1,
                int.Parse(match.Groups["toRow"].Value, CultureInfo.InvariantCulture),
                16_384,
                wholeRow: true);
        }

        throw new FormatException($"'{address}' is not a supported spreadsheet address.");
    }

    internal static string GetCellReference(int row, int column)
    {
        return GetColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
    }

    internal static string GetColumnName(int column)
    {
        ThrowIfOutsideTheGrid(column, nameof(column), MaximumColumn);

        string result = string.Empty;
        int value = column;
        while (value > 0)
        {
            int modulo = (value - 1) % 26;
            result = Convert.ToChar('A' + modulo) + result;
            value = (value - modulo) / 26;
        }

        return result;
    }

    internal static int GetColumnNumber(string columnName)
    {
        int result = 0;
        foreach (char character in columnName.ToUpperInvariant())
        {
            if (character < 'A' || character > 'Z')
            {
                throw new FormatException($"'{columnName}' is not a valid column name.");
            }

            result = checked(result * 26 + character - 'A' + 1);

            // Checked before the loop ends rather than after it, so that a long
            // run of letters cannot overflow on the way to being rejected.
            if (result > MaximumColumn)
            {
                throw new FormatException($"'{columnName}' is beyond XFD, the last column of a worksheet.");
            }
        }

        return result;
    }

    private static void ThrowIfOutsideTheGrid(int value, string name, int maximum)
    {
        if (value < 1 || value > maximum)
        {
            string noun = maximum == MaximumRow ? "rows" : "columns";
            throw new ArgumentOutOfRangeException(
                name,
                value,
                $"A worksheet has {noun} 1 to {maximum.ToString(CultureInfo.InvariantCulture)}.");
        }
    }
}
