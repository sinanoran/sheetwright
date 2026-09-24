namespace Sheetwright;

/// <summary>
/// One column of a grid export: a heading, and how to get the value out of a row.
/// </summary>
/// <param name="Header">The heading, written into the first row.</param>
/// <param name="Value">
/// The cell value. Return the underlying type rather than a string wherever there
/// is one — a <see cref="DateTime"/> or a number written as text sorts
/// alphabetically and cannot be summed, which is most of the reason the person
/// asked for Excel instead of CSV.
/// </param>
public sealed record SpreadsheetGridColumn<T>(string Header, Func<T, object?> Value)
{
    /// <summary>An Excel format code, for the whole column.</summary>
    public string? NumberFormat { get; init; }
}

/// <summary>
/// A list of rows, as a workbook somebody can open.
/// </summary>
/// <remarks>
/// <para>
/// Every grid export wants the same four things — a bold heading row that stays
/// put while you scroll, columns wide enough to read, dates that are dates — and
/// none of them is interesting enough to write twice. This exists because there
/// are two callers today, the member directory and the audit trail, and they had
/// begun to differ in ways nobody had decided on.
/// </para>
/// <para>
/// It holds the whole workbook in memory, so the caller decides how many rows is
/// too many before calling. There is no paging here and there should not be: a
/// row count that needs paging needs a job and a file, not a bigger buffer.
/// </para>
/// </remarks>
public static class SpreadsheetGrid
{
    /// <summary>The media type of an .xlsx file.</summary>
    /// <remarks>
    /// Long, and wrong in a way nothing catches: a browser handed the wrong one
    /// downloads a file Excel then refuses to open, and the report reads as "the
    /// export is corrupt".
    /// </remarks>
    public const string ContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /// <summary>
    /// Writes the grid in whichever format the caller asked for.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Returns false for a format it does not know rather than throwing, because
    /// the right answer to <c>?format=xls</c> is a 400 naming the field, and the
    /// exception type that produces one belongs to the caller's application, not
    /// to a spreadsheet library. The caller owns the error; this owns the list of
    /// formats.
    /// </para>
    /// <para>
    /// Falling back to Excel for an unrecognised format would be the friendlier
    /// spelling and the wrong one: a typo would return a workbook to something
    /// expecting CSV, which fails later and somewhere else.
    /// </para>
    /// </remarks>
    public static bool TryWriteAs<T>(
        string? format,
        string sheetName,
        IReadOnlyList<SpreadsheetGridColumn<T>> columns,
        IEnumerable<T> rows,
        [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out SpreadsheetGridFile? file)
    {
        // Absent means Excel: it is what a person clicking Export wants, and it
        // is what the endpoints returned before there was a choice.
        if (string.IsNullOrWhiteSpace(format) || format.Equals("xlsx", StringComparison.OrdinalIgnoreCase))
        {
            file = new SpreadsheetGridFile(Write(sheetName, columns, rows), ContentType, ".xlsx");
            return true;
        }

        if (format.Equals("csv", StringComparison.OrdinalIgnoreCase))
        {
            file = new SpreadsheetGridFile(
                SpreadsheetGridCsv.Write(columns, rows),
                SpreadsheetGridCsv.ContentType,
                ".csv");
            return true;
        }

        file = null;
        return false;
    }

    public static byte[] Write<T>(
        string sheetName,
        IReadOnlyList<SpreadsheetGridColumn<T>> columns,
        IEnumerable<T> rows)
    {
        ArgumentNullException.ThrowIfNull(sheetName);
        ArgumentNullException.ThrowIfNull(columns);
        ArgumentNullException.ThrowIfNull(rows);

        if (columns.Count == 0)
        {
            throw new ArgumentException("A grid needs at least one column.", nameof(columns));
        }

        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add(sheetName);

        for (int column = 0; column < columns.Count; column++)
        {
            sheet.Cells[1, column + 1].Value = columns[column].Header;

            if (columns[column].NumberFormat is string format)
            {
                sheet.Column(column + 1).Style.NumberFormat.SetFormat(format);
            }
        }

        sheet.Cells[1, 1, 1, columns.Count].Style.Font.SetBold(true);

        int row = 2;
        foreach (T item in rows)
        {
            for (int column = 0; column < columns.Count; column++)
            {
                sheet.Cells[row, column + 1].Value = columns[column].Value(item);
            }

            row++;
        }

        // Before the widths are measured, because the heading is often the
        // longest thing in a column and auto-fit reads what is there now.
        sheet.View.FreezePanes(2, 1);
        sheet.Cells.AutoFitColumns();

        return package.GetAsByteArray();
    }
}

/// <summary>A written grid, and what to call it.</summary>
/// <param name="Content">The bytes of the written file.</param>
/// <param name="ContentType">The MIME type to serve those bytes as.</param>
/// <param name="Extension">
/// Including the dot, so a caller composes a name without knowing which format
/// it asked for.
/// </param>
public sealed record SpreadsheetGridFile(byte[] Content, string ContentType, string Extension);
