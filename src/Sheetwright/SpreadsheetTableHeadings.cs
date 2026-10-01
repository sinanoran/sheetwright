using System.Globalization;

namespace Sheetwright;

/// <summary>
/// Table column naming, shared by <see cref="SpreadsheetWorksheet"/> and
/// <see cref="SpreadsheetStreamWriter"/> so both name a table's columns the same way.
/// </summary>
/// <remarks>
/// Excel requires a table's column names to be unique and non-empty, and it checks them against
/// the cells of the header row: a name invented here has to be written back into the cell it came
/// from, or the file is one Excel offers to repair. Both callers do that.
/// </remarks>
internal static class SpreadsheetTableHeadings
{
    internal static string Resolve(string? heading, int column, HashSet<string> taken)
    {
        string name = string.IsNullOrWhiteSpace(heading)
            ? SpreadsheetAddress.GetColumnName(column)
            : heading!;

        if (taken.Add(name))
        {
            return name;
        }

        int suffix = 2;
        string candidate;
        do
        {
            candidate = $"{name}{suffix.ToString(CultureInfo.InvariantCulture)}";
            suffix++;
        }
        while (!taken.Add(candidate));

        return candidate;
    }
}
