using System.Text;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// The same grid as CSV, for whatever reads it next.
/// </summary>
/// <remarks>
/// <para>
/// Two things are being defended. The first is escaping: a comma, a quote or a
/// newline inside a value shifts every field after it, and the failure is silent
/// — the file parses, and the columns are wrong from that row on.
/// </para>
/// <para>
/// The second is the formula guard, which is a security control rather than a
/// formatting one. A cell in a workbook is typed and Excel never evaluates a
/// string; CSV has no types, so a display name of <c>=1+1</c> becomes a formula
/// and one calling <c>DDE</c> is code running on the machine of whoever opened
/// the export. Everything exported here is user-supplied.
/// </para>
/// </remarks>
public sealed class SpreadsheetGridCsvTests
{
    private sealed record Row(string Text, int Count, DateTime When);

    private static readonly SpreadsheetGridColumn<Row>[] Columns =
    [
        new("Text", row => row.Text),
        new("Count", row => row.Count),
        new("When", row => row.When)
    ];

    [Fact]
    public void The_headings_come_first_then_a_row_each()
    {
        string csv = Write([new("Ada", 3, new DateTime(2026, 1, 15))]);

        Assert.StartsWith("Text,Count,When\r\n", csv, StringComparison.Ordinal);
        Assert.Contains("Ada,3,", csv, StringComparison.Ordinal);
    }

    [Fact]
    public void Rows_end_with_a_carriage_return_and_a_line_feed()
    {
        // RFC 4180, and the readers that care are the older ones.
        string csv = Write([new("Ada", 1, default)]);

        Assert.EndsWith("\r\n", csv, StringComparison.Ordinal);
        Assert.DoesNotContain("\n\n", csv, StringComparison.Ordinal);
    }

    [Fact]
    public void A_value_containing_a_comma_is_quoted()
    {
        // Without this, every column after it shifts by one and the file still
        // parses. That is the whole reason quoting exists.
        Assert.Contains("\"Ada, Countess\"", Write([new("Ada, Countess", 1, default)]), StringComparison.Ordinal);
    }

    [Fact]
    public void A_quote_inside_a_value_is_doubled()
    {
        Assert.Contains("\"Ada \"\"the Countess\"\"\"", Write([new("Ada \"the Countess\"", 1, default)]), StringComparison.Ordinal);
    }

    [Fact]
    public void A_newline_inside_a_value_is_quoted_rather_than_ending_the_row()
    {
        string csv = Write([new("first\nsecond", 1, default)]);

        Assert.Contains("\"first\nsecond\"", csv, StringComparison.Ordinal);
    }

    [Fact]
    public void Surrounding_space_is_preserved_by_quoting()
    {
        // A trimmed name is a changed name, and the trimming would be done by
        // the reader rather than here.
        Assert.Contains("\"  padded  \"", Write([new("  padded  ", 1, default)]), StringComparison.Ordinal);
    }

    [Fact]
    public void An_ordinary_value_is_not_quoted()
    {
        // The negative control. Quoting everything would pass every test above
        // and make the file harder to read for no benefit.
        Assert.Contains("Ada,1,", Write([new("Ada", 1, default)]), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("=1+1")]
    [InlineData("+1")]
    [InlineData("-1+1")]
    [InlineData("@SUM(A1)")]
    public void Text_that_Excel_would_treat_as_a_formula_is_neutralised(string hostile)
    {
        // Prefixed with an apostrophe, which is what Excel itself writes to mean
        // "this is text".
        string csv = Write([new(hostile, 1, default)]);

        Assert.Contains("'" + hostile, csv, StringComparison.Ordinal);
        Assert.DoesNotContain("\n" + hostile, csv, StringComparison.Ordinal);
    }

    [Fact]
    public void The_formula_guard_and_the_quoting_compose()
    {
        // The case that matters most and is the least like the others: a
        // hyperlink formula carries both quotes and a comma, so it needs the
        // apostrophe, the surrounding quotes and its own quotes doubled — and
        // getting any one of the three wrong leaves it executable or shifts the
        // row.
        string csv = Write([new("=HYPERLINK(\"http://example.test\",\"click\")", 1, default)]);

        Assert.Contains(
            "\"'=HYPERLINK(\"\"http://example.test\"\",\"\"click\"\")\"",
            csv,
            StringComparison.Ordinal);
    }

    [Fact]
    public void A_heading_is_guarded_the_same_way_as_a_value()
    {
        // Headings are as caller-supplied as the rows are.
        byte[] bytes = SpreadsheetGridCsv.Write<Row>(
            [new("=cmd", row => row.Count)],
            [new("x", 1, default)]);

        Assert.Contains("'=cmd", Text(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void A_negative_number_keeps_its_minus_sign()
    {
        // The formula guard applies to text and not to numbers. Applying it
        // everywhere would turn -5 into '-5, which is the sort of fix that
        // silently corrupts a column of amounts.
        string csv = Write([new("x", -5, default)]);

        Assert.Contains("x,-5,", csv, StringComparison.Ordinal);
        Assert.DoesNotContain("'-5", csv, StringComparison.Ordinal);
    }

    [Fact]
    public void A_date_is_written_round_trippable_rather_than_pretty()
    {
        // Nobody is looking at this file; something is parsing it.
        string csv = Write([new("x", 1, new DateTime(2026, 8, 27, 9, 30, 0))]);

        Assert.Contains("2026-08-27T09:30:00", csv, StringComparison.Ordinal);
    }

    [Fact]
    public void A_null_is_an_empty_field_rather_than_the_word()
    {
        byte[] bytes = SpreadsheetGridCsv.Write<Row>(
            [new("Text", _ => null), new("Count", row => row.Count)],
            [new("x", 7, default)]);

        Assert.Contains("\r\n,7\r\n", Text(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void The_file_begins_with_a_byte_order_mark()
    {
        // Without it, Excel opens a double-clicked UTF-8 file in the machine's
        // ANSI code page and every accented character is wrong.
        byte[] bytes = SpreadsheetGridCsv.Write(Columns, [new("Ünal", 1, default)]);

        Assert.Equal<byte[]>([0xEF, 0xBB, 0xBF], bytes[..3]);
        Assert.Contains("Ünal", Text(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void An_empty_grid_is_a_heading_row_and_nothing_else()
    {
        string csv = Write([]);

        Assert.Equal("Text,Count,When\r\n", csv);
    }

    [Fact]
    public void A_grid_with_no_columns_is_refused()
    {
        Assert.Throws<ArgumentException>(
            () => SpreadsheetGridCsv.Write(Array.Empty<SpreadsheetGridColumn<Row>>(), []));
    }

    [Fact]
    public void Nulls_are_refused()
    {
        Assert.Throws<ArgumentNullException>(() => SpreadsheetGridCsv.Write<Row>(null!, []));
        Assert.Throws<ArgumentNullException>(() => SpreadsheetGridCsv.Write(Columns, null!));
    }

    [Fact]
    public void The_media_type_names_the_charset()
    {
        // A reader that has to guess the encoding will sooner or later guess
        // wrong, and the symptom is a mangled name rather than an error.
        Assert.Equal("text/csv; charset=utf-8", SpreadsheetGridCsv.ContentType);
    }

    private static string Write(Row[] rows) => Text(SpreadsheetGridCsv.Write(Columns, rows));

    private static string Text(byte[] bytes) => Encoding.UTF8.GetString(bytes)[1..];
}
