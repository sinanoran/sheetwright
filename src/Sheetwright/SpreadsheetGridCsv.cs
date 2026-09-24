using System.Globalization;
using System.Text;

namespace Sheetwright;

/// <summary>
/// The same grid, as CSV, for whatever is going to read it next.
/// </summary>
/// <remarks>
/// <para>
/// A different audience rather than a cheaper workbook. The .xlsx export is for a
/// person: it has widths, a frozen heading and dates that sort. This is for
/// another system — an import, a script, a <c>COPY</c> — which wants none of that
/// and does want a file it can parse without a library.
/// </para>
/// <para>
/// It shares <see cref="SpreadsheetGridColumn{T}"/> with the workbook writer, so
/// a grid describes its columns once and can be handed out in either form.
/// <c>NumberFormat</c> is ignored here, which is the honest thing for a format
/// with no formatting: values are written in a machine-readable spelling and the
/// reader decides how they look.
/// </para>
/// <para>
/// A CSV writer living in a project called Sheetwright is slightly
/// off-name. It is here rather than in a third project because it is forty lines
/// that exist entirely to reuse the column type above, and a project boundary
/// drawn for tidiness is a project boundary somebody has to maintain.
/// </para>
/// </remarks>
public static class SpreadsheetGridCsv
{
    /// <remarks>
    /// The charset is part of it. Without it a reader is entitled to guess, and
    /// the ones that guess wrong turn a Turkish name into mojibake.
    /// </remarks>
    public const string ContentType = "text/csv; charset=utf-8";

    /// <summary>
    /// Characters that make Excel read a cell as a formula rather than as text.
    /// </summary>
    /// <remarks>
    /// This is the CSV-only half of a real vulnerability. In a workbook a string
    /// is written as a string and Excel never evaluates it; in CSV there is no
    /// type, so a display name of <c>=1+1</c> becomes a formula, and one of
    /// <c>=HYPERLINK(...)</c> or a <c>DDE</c> call is an attack that runs on the
    /// machine of whoever opened the export. The data here is user-supplied —
    /// display names, search terms, audit payloads — so this is reachable.
    /// </remarks>
    private static readonly char[] FormulaLeaders = ['=', '+', '-', '@', '\t', '\r'];

    public static byte[] Write<T>(
        IReadOnlyList<SpreadsheetGridColumn<T>> columns,
        IEnumerable<T> rows)
    {
        ArgumentNullException.ThrowIfNull(columns);
        ArgumentNullException.ThrowIfNull(rows);

        if (columns.Count == 0)
        {
            throw new ArgumentException("A grid needs at least one column.", nameof(columns));
        }

        StringBuilder csv = new();

        AppendRow(csv, columns.Select(column => Field(column.Header, quoteLeaders: true)));

        foreach (T row in rows)
        {
            AppendRow(csv, columns.Select(column => Render(column.Value(row))));
        }

        // With a byte-order mark. RFC 4180 says nothing about it and Excel needs
        // it: double-clicked without one, a UTF-8 file is read in the machine's
        // ANSI code page and every accented character is wrong. The cost is that
        // a naive reader sees three bytes it did not expect, which is the lesser
        // failure and the more visible one.
        return [.. Encoding.UTF8.GetPreamble(), .. Encoding.UTF8.GetBytes(csv.ToString())];
    }

    private static void AppendRow(StringBuilder csv, IEnumerable<string> fields)
    {
        csv.AppendJoin(',', fields);

        // CRLF, because RFC 4180 says so and because the readers that care are
        // the older ones.
        csv.Append("\r\n");
    }

    /// <summary>One value, in a spelling something else can parse.</summary>
    /// <remarks>
    /// The type is known here, which is what makes the formula guard safe to
    /// apply. A number is written by us and cannot be hostile, so a negative one
    /// keeps its minus sign; only text a caller supplied is treated as suspect.
    /// </remarks>
    private static string Render(object? value)
    {
        return value switch
        {
            null or DBNull => string.Empty,

            // Round-trip, so a reader gets the instant back rather than a
            // rendering of it. Unlike the workbook, nobody is looking at this.
            DateTime date => Field(date.ToString("O", CultureInfo.InvariantCulture), quoteLeaders: false),
            DateTimeOffset moment => Field(moment.ToString("O", CultureInfo.InvariantCulture), quoteLeaders: false),
            TimeSpan span => Field(span.ToString("c", CultureInfo.InvariantCulture), quoteLeaders: false),

            bool flag => flag ? "true" : "false",

            byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal =>
                Field(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, quoteLeaders: false),

            _ => Field(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, quoteLeaders: true),
        };
    }

    private static string Field(string value, bool quoteLeaders)
    {
        if (quoteLeaders && value.Length > 0 && Array.IndexOf(FormulaLeaders, value[0]) >= 0)
        {
            // A leading apostrophe is what Excel itself writes to mean "this is
            // text". It is visible in a plain-text reader, which is the trade:
            // one stray character beats arbitrary evaluation.
            value = "'" + value;
        }

        bool mustQuote = value.Contains('"', StringComparison.Ordinal)
            || value.Contains(',', StringComparison.Ordinal)
            || value.Contains('\n', StringComparison.Ordinal)
            || value.Contains('\r', StringComparison.Ordinal)
            // Leading or trailing space survives only inside quotes, and a
            // trimmed name is a changed name.
            || (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1])));

        return mustQuote
            ? "\"" + value.Replace("\"", "\"\"", StringComparison.Ordinal) + "\""
            : value;
    }
}
