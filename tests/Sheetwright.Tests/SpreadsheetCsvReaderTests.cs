using System.Text;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// The CSV reader: the same file from every machine that writes one.
/// </summary>
/// <remarks>
/// <para>
/// A CSV has no header that says how it was written. A Turkish or German Excel writes semicolons and a comma
/// decimal, an English one commas; a UTF-8 file with a byte-order mark is not the same bytes as one without; a
/// quoted field may hold a line break. Each of those is a way to read a file wrong without an error, so each
/// has a test, and the files here are written byte by byte rather than by Sheetwright's own writer, which
/// would only ever agree with itself.
/// </para>
/// <para>
/// The apostrophe rule is the other half of <see cref="SpreadsheetGridCsv"/>'s formula guard: what the writer
/// adds so that Excel will not run a cell, the reader takes off, so a file Sheetwright wrote reads back as the
/// values that went in.
/// </para>
/// </remarks>
public sealed class SpreadsheetCsvReaderTests
{
    private static readonly UTF8Encoding Utf8NoBom = new(false);

    private static byte[] Utf8(string text) => Utf8NoBom.GetBytes(text);

    private static List<string[]> Read(byte[] bytes, SpreadsheetReaderOptions? options = null)
    {
        using MemoryStream stream = new(bytes);
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream, options);
        return [.. reader.ReadRows().Select(row => row.Cells.Select(cell => cell.ToString()).ToArray())];
    }

    private static List<string[]> Read(string text, SpreadsheetReaderOptions? options = null) => Read(Utf8(text), options);

    [Fact]
    public void Rows_come_back_in_order_with_their_own_row_numbers()
    {
        using MemoryStream stream = new(Utf8("a,b\r\n1,2\r\n3,4\r\n"));
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);

        List<SpreadsheetRow> rows = [.. reader.ReadRows()];

        Assert.Equal([1, 2, 3], rows.Select(row => row.RowNumber));
        Assert.Equal("3", rows[2][0].Text);
        Assert.Equal(SpreadsheetCellKind.Text, rows[2][0].Kind);
    }

    [Fact]
    public void Every_cell_of_a_csv_is_text_and_an_empty_field_is_empty()
    {
        using MemoryStream stream = new(Utf8("12.480,00;;x\n"));
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);

        SpreadsheetRow row = reader.ReadRows().Single();

        Assert.Equal(';', reader.Delimiter);
        Assert.Equal("12.480,00", row[0].Text);
        Assert.True(row[1].IsEmpty);
        Assert.Equal("x", row[2].Text);
        Assert.True(row[9].IsEmpty);
    }

    [Theory]
    [InlineData("name,city,age\nAda,London,36\nGrace,New York,45\n", ',')]
    [InlineData("name;city;age\nAda;London;36\nGrace;New York;45\n", ';')]
    [InlineData("name\tcity\tage\nAda\tLondon\t36\nGrace\tNew York\t45\n", '\t')]
    [InlineData("name|city|age\nAda|London|36\nGrace|New York|45\n", '|')]
    public void The_separator_is_the_one_that_splits_the_rows_the_same_way(string text, char expected)
    {
        using MemoryStream stream = new(Utf8(text));
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);

        List<string[]> rows = [.. reader.ReadRows().Select(row => row.Cells.Select(cell => cell.ToString()).ToArray())];

        Assert.Equal(expected, reader.Delimiter);
        Assert.Equal(3, rows.Count);
        Assert.All(rows, row => Assert.Equal(3, row.Length));
    }

    [Fact]
    public void A_separator_inside_quotes_does_not_decide_the_separator()
    {
        // Semicolons only inside quotes: the file is comma-separated.
        using MemoryStream stream = new(Utf8("a,b\n\"x;y;z\",\"p;q\"\n\"1;2\",3\n"));
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);

        List<SpreadsheetRow> rows = [.. reader.ReadRows()];

        Assert.Equal(',', reader.Delimiter);
        Assert.Equal("x;y;z", rows[1][0].Text);
    }

    [Fact]
    public void One_column_with_no_separator_is_read_as_comma_separated()
    {
        using MemoryStream stream = new(Utf8("name\nAda\nGrace\n"));
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);

        Assert.Equal(3, reader.ReadRows().Count());
        Assert.Equal(',', reader.Delimiter);
    }

    [Fact]
    public void A_sep_line_sets_the_separator_and_is_not_a_row()
    {
        List<string[]> rows = Read("sep=;\r\na;b\r\n1;2\r\n");

        Assert.Equal(2, rows.Count);
        Assert.Equal(["a", "b"], rows[0]);
    }

    [Fact]
    public void The_row_numbers_do_not_count_the_sep_line()
    {
        using MemoryStream stream = new(Utf8("sep=,\na,b\n"));
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);

        Assert.Equal(1, reader.ReadRows().Single().RowNumber);
    }

    [Fact]
    public void A_sep_line_is_data_when_the_directive_is_switched_off()
    {
        List<string[]> rows = Read("sep=;\na;b\n", new SpreadsheetReaderOptions { HonourSepDirective = false });

        // The line is data: it is split on the separator the file is detected to have, like any other row.
        Assert.Equal(2, rows.Count);
        Assert.Equal(["sep="], rows[0]);
    }

    [Fact]
    public void A_separator_given_in_the_options_wins_over_the_file()
    {
        List<string[]> rows = Read("a;b;c\n1;2;3\n", new SpreadsheetReaderOptions { Delimiter = ',' });

        Assert.Equal(["a;b;c"], rows[0]);
    }

    [Fact]
    public void The_separator_cannot_be_a_quote_or_a_line_break()
    {
        using MemoryStream stream = new(Utf8("a,b\n"));

        Assert.Throws<ArgumentException>(() => SpreadsheetCsvReader.Open(stream, new SpreadsheetReaderOptions { Delimiter = '"' }));
        Assert.Throws<ArgumentException>(() => SpreadsheetCsvReader.Open(stream, new SpreadsheetReaderOptions { Delimiter = '\n' }));
    }

    [Fact]
    public void A_quoted_field_holds_separators_doubled_quotes_and_line_breaks()
    {
        List<string[]> rows = Read("a,b\r\n\"one, two\",\"say \"\"hi\"\"\"\r\n\"line1\r\nline2\",\"x\ny\"\r\n");

        Assert.Equal(3, rows.Count);
        Assert.Equal(["one, two", "say \"hi\""], rows[1]);
        Assert.Equal(["line1\r\nline2", "x\ny"], rows[2]);
    }

    [Fact]
    public void A_quoted_line_break_is_one_record_and_numbers_the_rows_after_it_by_record()
    {
        using MemoryStream stream = new(Utf8("a,b\n\"x\ny\",1\nlast,2\n"));
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);

        Assert.Equal([1, 2, 3], reader.ReadRows().Select(row => row.RowNumber));
    }

    [Fact]
    public void A_quote_in_the_middle_of_an_unquoted_field_is_a_character()
    {
        List<string[]> rows = Read("a,b\n5\" pipe,x\n");

        Assert.Equal(["5\" pipe", "x"], rows[1]);
    }

    [Fact]
    public void Text_after_a_closing_quote_is_kept()
    {
        List<string[]> rows = Read("a,b\n\"abc\"def,x\n");

        Assert.Equal(["abcdef", "x"], rows[1]);
    }

    [Theory]
    [InlineData("a,b\r\n1,2\r\n3,4\r\n")]
    [InlineData("a,b\n1,2\n3,4\n")]
    [InlineData("a,b\r1,2\r3,4\r")]
    [InlineData("a,b\r\n1,2\r\n3,4")]
    public void Records_end_at_crlf_lf_or_a_lone_cr_and_the_last_need_not(string text)
    {
        List<string[]> rows = Read(text);

        Assert.Equal(3, rows.Count);
        Assert.Equal(["3", "4"], rows[2]);
    }

    [Fact]
    public void Blank_lines_are_skipped_but_still_use_up_their_row_number()
    {
        using MemoryStream stream = new(Utf8("a,b\n\n1,2\n\n\n3,4\n\n"));
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);

        Assert.Equal([1, 3, 6], reader.ReadRows().Select(row => row.RowNumber));
    }

    [Fact]
    public void Empty_rows_are_returned_when_asked_for()
    {
        List<string[]> rows = Read("a,b\n\n1,2\n", new SpreadsheetReaderOptions { IncludeEmptyRows = true });

        Assert.Equal(3, rows.Count);
        Assert.Empty(rows[1]);
    }

    [Fact]
    public void Rows_may_have_different_lengths_and_a_trailing_separator_is_an_empty_field()
    {
        List<string[]> rows = Read("a,b,c\n1,2\n3,4,5,6\n7,8,\n");

        // A row ends at its last cell with a value, so a trailing separator does not make it longer.
        Assert.Equal([3, 2, 4, 2], rows.Select(row => row.Length));
    }

    [Fact]
    public void A_quote_that_is_never_closed_is_refused_with_the_row()
    {
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => Read("a,b\n1,\"never closed\n2,3\n"));

        Assert.Contains("row 2", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_utf8_file_with_a_byte_order_mark_reads_without_it()
    {
        byte[] bytes = [.. new UTF8Encoding(true).GetPreamble(), .. Utf8("ad;şehir\nAyşe;İstanbul\n")];

        List<string[]> rows = Read(bytes);

        Assert.Equal("ad", rows[0][0]);
        Assert.Equal(["Ayşe", "İstanbul"], rows[1]);
    }

    [Fact]
    public void A_utf8_file_with_no_mark_is_read_as_utf8()
    {
        List<string[]> rows = Read("ad;şehir\nÖzlem;Ağrı\n");

        Assert.Equal(["Özlem", "Ağrı"], rows[1]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void A_utf16_file_is_read_by_its_byte_order_mark(bool bigEndian)
    {
        UnicodeEncoding encoding = new(bigEndian, byteOrderMark: true);
        byte[] bytes = [.. encoding.GetPreamble(), .. encoding.GetBytes("ad;şehir\r\nAyşe;İstanbul\r\n")];

        using MemoryStream stream = new(bytes);
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);
        List<SpreadsheetRow> rows = [.. reader.ReadRows()];

        Assert.Equal(bigEndian ? "utf-16BE" : "utf-16", reader.Encoding.WebName);
        Assert.Equal(';', reader.Delimiter);
        Assert.Equal("İstanbul", rows[1][1].Text);
    }

    [Fact]
    public void A_utf32_file_is_refused_by_name()
    {
        byte[] bytes = [.. new UTF32Encoding(false, true).GetPreamble(), .. new UTF32Encoding(false, true).GetBytes("a,b\n")];
        using MemoryStream stream = new(bytes);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => SpreadsheetCsvReader.Open(stream));

        Assert.Contains("UTF-32", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Bytes_that_are_not_valid_utf8_are_refused_and_not_replaced()
    {
        // 0xFD is a dotless i in Windows-1254 and an invalid byte in UTF-8: a Turkish Excel's "CSV (Comma delimited)".
        byte[] bytes = [.. Utf8("ad;soyad\nAyşe;"), 0x41, 0xFD, 0x6C, (byte)'\n'];

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => Read(bytes));

        Assert.Contains("encoding", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void An_encoding_in_the_options_reads_a_legacy_file()
    {
        byte[] bytes = [.. Encoding.Latin1.GetBytes("nom;ville\nRené;Orléans\n")];

        List<string[]> rows = Read(bytes, new SpreadsheetReaderOptions { Encoding = Encoding.Latin1 });

        Assert.Equal(["René", "Orléans"], rows[1]);
    }

    [Fact]
    public void An_encoding_in_the_options_still_skips_its_own_byte_order_mark()
    {
        byte[] bytes = [.. new UTF8Encoding(true).GetPreamble(), .. Utf8("a;b\n")];

        List<string[]> rows = Read(bytes, new SpreadsheetReaderOptions { Encoding = new UTF8Encoding(false, true) });

        Assert.Equal("a", rows[0][0]);
    }

    [Fact]
    public void A_stream_that_cannot_seek_and_hands_out_a_byte_at_a_time_reads_the_same()
    {
        byte[] bytes = Utf8("a;b\r\n\"x;\r\ny\";2\r\nöz;3\r\n");

        using ByteAtATimeStream stream = new(bytes);
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);
        List<string[]> rows = [.. reader.ReadRows().Select(row => row.Cells.Select(cell => cell.ToString()).ToArray())];

        Assert.Equal(';', reader.Delimiter);
        Assert.Equal(["x;\r\ny", "2"], rows[1]);
        Assert.Equal(["öz", "3"], rows[2]);
    }

    [Fact]
    public void A_file_longer_than_the_start_that_is_looked_at_is_read_whole_across_the_cut()
    {
        // Three-byte characters, so that the 64 KiB looked at to choose the separator ends inside one.
        StringBuilder text = new("head;er\n");
        for (int index = 0; index < 6_000; index++)
        {
            text.Append("€€€€€;ç").Append(index).Append('\n');
        }

        using MemoryStream stream = new(Utf8(text.ToString()));
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);
        List<SpreadsheetRow> rows = [.. reader.ReadRows()];

        Assert.Equal(';', reader.Delimiter);
        Assert.Equal(6_001, rows.Count);
        Assert.Equal("ç5999", rows[^1][1].Text);
        Assert.Equal("€€€€€", rows[^1][0].Text);
    }

    [Fact]
    public void A_file_is_read_once()
    {
        using MemoryStream stream = new(Utf8("a,b\n"));
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);
        _ = reader.ReadRows().ToList();

        Assert.Throws<InvalidOperationException>(() => reader.ReadRows());
    }

    [Fact]
    public void Disposing_closes_the_stream_unless_it_was_left_open()
    {
        MemoryStream closed = new(Utf8("a,b\n"));
        SpreadsheetCsvReader.Open(closed).Dispose();
        Assert.False(closed.CanRead);

        MemoryStream open = new(Utf8("a,b\n"));
        SpreadsheetCsvReader.Open(open, leaveOpen: true).Dispose();
        Assert.True(open.CanRead);
    }

    [Fact]
    public void An_empty_file_has_no_rows()
    {
        Assert.Empty(Read(string.Empty));
    }

    [Fact]
    public void More_rows_than_the_limit_is_refused_by_name_and_exactly_the_limit_is_not()
    {
        SpreadsheetReaderOptions options = new() { MaxRows = 3 };

        Assert.Equal(3, Read("a\nb\nc\n", options).Count);
        SpreadsheetReadLimitException exception = Assert.Throws<SpreadsheetReadLimitException>(() => Read("a\nb\nc\nd\n", options));

        Assert.Equal(SpreadsheetReadLimit.MaxRows, exception.Limit);
        Assert.Equal(3, exception.Maximum);
        Assert.Contains("MaxRows", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Empty_rows_do_not_count_against_the_row_limit()
    {
        Assert.Equal(2, Read("a\n\n\n\n\nb\n", new SpreadsheetReaderOptions { MaxRows = 2 }).Count);
    }

    [Fact]
    public void More_columns_than_the_limit_is_refused_by_name()
    {
        SpreadsheetReaderOptions options = new() { MaxColumns = 3 };

        Assert.Single(Read("a,b,c\n", options));
        SpreadsheetReadLimitException exception = Assert.Throws<SpreadsheetReadLimitException>(() => Read("a,b,c,d\n", options));

        Assert.Equal(SpreadsheetReadLimit.MaxColumns, exception.Limit);
    }

    [Fact]
    public void A_cell_longer_than_the_limit_is_refused_while_it_is_read_and_not_truncated()
    {
        SpreadsheetReaderOptions options = new() { MaxCellCharacters = 10 };

        Assert.Single(Read(new string('x', 10) + "\n", options));
        SpreadsheetReadLimitException exception = Assert.Throws<SpreadsheetReadLimitException>(() => Read(new string('x', 11) + "\n", options));

        Assert.Equal(SpreadsheetReadLimit.MaxCellCharacters, exception.Limit);
        Assert.DoesNotContain("xxx", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_quoted_cell_is_held_to_the_same_limit()
    {
        SpreadsheetReaderOptions options = new() { MaxCellCharacters = 10 };

        Assert.Throws<SpreadsheetReadLimitException>(() => Read("\"" + new string('y', 11) + "\"\n", options));
    }

    [Theory]
    [InlineData("'=1+1", "=1+1")]
    [InlineData("'+1", "+1")]
    [InlineData("'-5text", "-5text")]
    [InlineData("'@SUM(A1)", "@SUM(A1)")]
    [InlineData("'＝1", "＝1")]
    [InlineData("'＋1", "＋1")]
    [InlineData("'－1", "－1")]
    [InlineData("'＠x", "＠x")]
    public void One_apostrophe_in_front_of_a_formula_leader_is_taken_off(string written, string expected)
    {
        Assert.Equal(expected, Read("h\n" + written + "\n")[1][0]);
    }

    [Fact]
    public void An_apostrophe_in_front_of_a_tab_or_a_carriage_return_is_taken_off()
    {
        Assert.Equal("\tx", Read("h\n\"'\tx\"\n")[1][0]);
        Assert.Equal("\rx", Read("h\n\"'\rx\"\n")[1][0]);
    }

    [Theory]
    [InlineData("'abc")]
    [InlineData("'")]
    [InlineData("''=1")]
    [InlineData("it's")]
    [InlineData("'1")]
    public void Any_other_apostrophe_is_data(string written)
    {
        Assert.Equal(written, Read("h\n" + written + "\n")[1][0]);
    }

    [Fact]
    public void The_apostrophe_rule_can_be_switched_off()
    {
        List<string[]> rows = Read("h\n'=1+1\n", new SpreadsheetReaderOptions { UnescapeFormulaGuard = false });

        Assert.Equal("'=1+1", rows[1][0]);
    }

    [Fact]
    public void The_apostrophe_is_taken_off_a_quoted_field_too()
    {
        Assert.Equal("=a,b", Read("h\n\"'=a,b\"\n")[1][0]);
    }

    [Fact]
    public void A_file_the_grid_writer_wrote_reads_back_as_the_values_that_went_in()
    {
        string[] values =
        [
            "plain",
            "=1+1",
            "+31 555 0100",
            "-minus text",
            "@mention",
            "\ttabbed",
            "\rcarriage",
            "＝full width",
            "＋full width",
            "－full width",
            "＠full width",
            "a,b",
            "line\nbreak",
            "say \"hi\"",
            " leading space",
            "trailing space ",
            "ünïcödé İstanbul",
            "'plain apostrophe",
            "1;2;3",
        ];
        SpreadsheetGridColumn<string>[] columns = [new("Text", value => value), new("Length", value => value.Length)];

        byte[] csv = SpreadsheetGridCsv.Write(columns, values);
        List<SpreadsheetRow> rows;
        using (MemoryStream stream = new(csv))
        using (SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream))
        {
            rows = [.. reader.ReadRows()];
        }

        Assert.Equal("Text", rows[0][0].Text);
        Assert.Equal(values, rows.Skip(1).Select(row => row[0].Text));
        Assert.Equal(values.Select(value => value.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)), rows.Skip(1).Select(row => row[1].Text));
    }

    [Fact]
    public void A_guarded_heading_reads_back_too()
    {
        SpreadsheetGridColumn<string>[] columns = [new("=Name", value => value), new("-Size", value => value)];

        byte[] csv = SpreadsheetGridCsv.Write(columns, ["x"]);
        using MemoryStream stream = new(csv);
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);
        SpreadsheetRow header = reader.ReadRows().First();

        Assert.Equal("=Name", header[0].Text);
        Assert.Equal("-Size", header[1].Text);
    }

    [Fact]
    public void A_text_that_itself_starts_with_an_apostrophe_and_a_leader_cannot_be_told_from_a_guarded_one()
    {
        // The writer guards only text that starts with a leader, so "'=x" is written as it is, and the reader
        // cannot know the apostrophe was the author's. This is the one ambiguity of the rule, documented on
        // SpreadsheetGridCsv; the test pins it so a change to either side is a decision and not an accident.
        SpreadsheetGridColumn<string>[] columns = [new("Text", value => value)];

        byte[] csv = SpreadsheetGridCsv.Write(columns, ["'=x"]);
        using MemoryStream stream = new(csv);
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);

        Assert.Equal("=x", reader.ReadRows().Skip(1).Single()[0].Text);
    }

    /// <summary>A stream that cannot seek and returns one byte per read, the worst a network can do.</summary>
    private sealed class ByteAtATimeStream(byte[] bytes) : Stream
    {
        private int _position;

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (_position >= bytes.Length || count == 0)
            {
                return 0;
            }

            buffer[offset] = bytes[_position++];
            return 1;
        }

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
