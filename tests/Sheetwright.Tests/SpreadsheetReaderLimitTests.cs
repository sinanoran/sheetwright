using System.IO.Compression;
using System.Text;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// What the workbook reader refuses, and that it refuses it by name.
/// </summary>
/// <remarks>
/// <para>
/// An uploaded workbook is a zip written by a stranger. A few kilobytes can hold gigabytes of zeros; a
/// shared-string table can name millions of strings; a tag can have a million attributes. Each of those is an
/// attack on the memory of the process that reads it, so each limit has a file that goes past it, written here
/// at test scale by setting the limit low, and each is checked for the same three things: it is refused, the
/// refusal names the limit that was passed, and the file that is exactly at the limit is not refused. A limit
/// that errs by one is a limit somebody's real file will hit.
/// </para>
/// <para>
/// The size limits are checked twice, once against what the zip header declares (to refuse a bomb before reading
/// any of it) and once against the bytes that come out (because a header can lie), and both are tested.
/// </para>
/// </remarks>
public sealed class SpreadsheetReaderLimitTests
{
    private static SpreadsheetReadLimitException Refused(Action action)
    {
        return Assert.Throws<SpreadsheetReadLimitException>(action);
    }

    private static List<SpreadsheetRow> Read(XlsxBuilder workbook, SpreadsheetReaderOptions options)
    {
        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(workbook.ToStream(), options);
        return [.. reader.ReadRows(0)];
    }

    private static List<SpreadsheetRow> Read(MemoryStream stream, SpreadsheetReaderOptions options)
    {
        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(stream, options);
        return [.. reader.ReadRows(0)];
    }

    private static string Rows(int count) => string.Concat(Enumerable.Range(1, count).Select(row => XlsxBuilder.Row(row, XlsxBuilder.Number($"A{row}", "1"))));

    [Fact]
    public void A_zip_bomb_is_refused_by_the_size_its_header_declares_before_any_of_it_is_read()
    {
        // Eight megabytes of one byte, which deflates to a few kilobytes.
        XlsxBuilder builder = XlsxBuilder.Workbook(string.Empty);
        builder.Add("xl/bomb.bin", new byte[8 * 1024 * 1024]);

        SpreadsheetReadLimitException exception = Refused(() =>
            SpreadsheetWorkbookReader.Open(builder.ToStream(), new SpreadsheetReaderOptions { MaxUncompressedBytes = 1024 * 1024, MaxCompressionRatio = 100_000 }));

        Assert.Equal(SpreadsheetReadLimit.MaxUncompressedBytes, exception.Limit);
        Assert.Equal(1024 * 1024, exception.Maximum);
        Assert.Contains("MaxUncompressedBytes", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_ratio_bomb_is_refused_by_the_ratio_and_not_by_the_total_size()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(string.Empty);
        builder.Add("xl/bomb.bin", new byte[8 * 1024 * 1024]);

        SpreadsheetReadLimitException exception = Refused(() =>
            SpreadsheetWorkbookReader.Open(builder.ToStream(), new SpreadsheetReaderOptions { MaxUncompressedBytes = 1024L * 1024 * 1024, MaxCompressionRatio = 100 }));

        Assert.Equal(SpreadsheetReadLimit.MaxCompressionRatio, exception.Limit);
        Assert.Equal(100, exception.Maximum);
        Assert.Contains("bomb.bin", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_workbook_whose_parts_stay_within_the_size_and_the_ratio_is_not_refused()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(Rows(50));

        Assert.Equal(50, Read(builder, new SpreadsheetReaderOptions { MaxUncompressedBytes = 64 * 1024, MaxCompressionRatio = 100 }).Count);
    }

    [Fact]
    public void A_part_that_inflates_past_what_its_own_header_declares_is_stopped_while_it_is_read()
    {
        // The header is the part of the file an attacker controls, so the bytes are counted as they come out.
        SpreadsheetReaderOptions options = new() { MaxUncompressedBytes = long.MaxValue / 4 };
        using MemoryStream inner = new(new byte[10_000]);
        using SpreadsheetBoundedPartStream stream = new(inner, "xl/liar.xml", declaredLength: 1_000, compressedLength: 1_000, options);

        byte[] buffer = new byte[4_096];
        int total = 0;
        SpreadsheetReadLimitException exception = Assert.Throws<SpreadsheetReadLimitException>(() =>
        {
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
            {
                total += read;
            }
        });

        Assert.Equal(SpreadsheetReadLimit.MaxUncompressedBytes, exception.Limit);
        Assert.True(total <= 4_096);
        Assert.Contains("liar.xml", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_part_that_expands_past_the_ratio_is_stopped_while_it_is_read_once_it_is_past_the_first_mebibyte()
    {
        SpreadsheetReaderOptions options = new() { MaxCompressionRatio = 50 };
        using MemoryStream inner = new(new byte[4 * 1024 * 1024]);
        using SpreadsheetBoundedPartStream stream = new(inner, "xl/ratio.xml", declaredLength: 4 * 1024 * 1024, compressedLength: 1_000, options);

        byte[] buffer = new byte[64 * 1024];
        long total = 0;
        SpreadsheetReadLimitException exception = Assert.Throws<SpreadsheetReadLimitException>(() =>
        {
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
            {
                total += read;
            }
        });

        Assert.Equal(SpreadsheetReadLimit.MaxCompressionRatio, exception.Limit);
        Assert.InRange(total, SpreadsheetBoundedPartStream.RatioFloorBytes, SpreadsheetBoundedPartStream.RatioFloorBytes + buffer.Length);
    }

    [Fact]
    public void A_header_that_understates_the_size_never_lets_more_than_it_declares_through()
    {
        // The declared size in the zip's central directory is overwritten with 100, so the part that is read
        // is cut off at 100 bytes (or refused), and never read past what the file said it was.
        XlsxBuilder builder = XlsxBuilder.Workbook(Rows(200));
        byte[] bytes = builder.ToStream().ToArray();
        Patch(bytes, "sheet1.xml", declaredLength: 100);

        Exception exception = Assert.ThrowsAny<Exception>(() => Read(new MemoryStream(bytes), new SpreadsheetReaderOptions()));

        Assert.True(exception is InvalidDataException or SpreadsheetReadLimitException, exception.GetType().Name);
    }

    [Fact]
    public void More_parts_than_the_limit_is_refused_by_name_and_exactly_the_limit_is_not()
    {
        XlsxBuilder Make(int extra)
        {
            XlsxBuilder builder = XlsxBuilder.Workbook(Rows(1));
            for (int index = 0; index < extra; index++)
            {
                builder.Add($"xl/extra{index}.xml", "<x/>");
            }

            return builder;
        }

        using ZipArchive probe = new(Make(0).ToStream());
        int baseParts = probe.Entries.Count;
        SpreadsheetReaderOptions options = new() { MaxParts = baseParts + 3 };

        Assert.Single(Read(Make(3), options));
        SpreadsheetReadLimitException exception = Refused(() => Read(Make(4), options));

        Assert.Equal(SpreadsheetReadLimit.MaxParts, exception.Limit);
    }

    [Fact]
    public void A_shared_string_flood_is_refused_by_the_count_before_any_row_is_read()
    {
        string flood = string.Concat(Enumerable.Repeat("<si><t>x</t></si>", 2_000));
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Shared("A1", 0)), sharedStrings: flood);
        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(
            builder.ToStream(),
            new SpreadsheetReaderOptions { MaxSharedStrings = 1_000 });

        // At the call, not at the first row: the table is loaded when the sheet is asked for.
        SpreadsheetReadLimitException exception = Refused(() => reader.ReadRows(0));

        Assert.Equal(SpreadsheetReadLimit.MaxSharedStrings, exception.Limit);
        Assert.Equal(1_000, exception.Maximum);
    }

    [Fact]
    public void A_shared_string_table_exactly_at_the_limit_is_read()
    {
        string table = string.Concat(Enumerable.Repeat("<si><t>x</t></si>", 1_000));
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Shared("A1", 999)), sharedStrings: table);

        Assert.Equal("x", Read(builder, new SpreadsheetReaderOptions { MaxSharedStrings = 1_000 })[0][0].Text);
    }

    [Fact]
    public void A_shared_string_table_with_too_many_characters_is_refused_by_the_characters()
    {
        string table = string.Concat(Enumerable.Repeat($"<si><t>{new string('y', 1_000)}</t></si>", 10));
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Shared("A1", 0)), sharedStrings: table);

        SpreadsheetReadLimitException exception = Refused(() => Read(builder, new SpreadsheetReaderOptions { MaxSharedStringCharacters = 5_000 }));

        Assert.Equal(SpreadsheetReadLimit.MaxSharedStringCharacters, exception.Limit);
        Assert.Equal(5_000, exception.Maximum);
    }

    [Fact]
    public void A_shared_string_longer_than_a_cell_can_be_is_refused_by_the_cell_limit()
    {
        string table = $"<si><t>{new string('z', 500)}</t></si>";
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Shared("A1", 0)), sharedStrings: table);

        SpreadsheetReadLimitException exception = Refused(() => Read(builder, new SpreadsheetReaderOptions { MaxCellCharacters = 100 }));

        Assert.Equal(SpreadsheetReadLimit.MaxCellCharacters, exception.Limit);
    }

    [Fact]
    public void A_cell_longer_than_the_limit_is_refused_and_a_cell_exactly_at_it_is_not()
    {
        SpreadsheetReaderOptions options = new() { MaxCellCharacters = 100 };

        Assert.Equal(100, Read(XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", new string('a', 100)))), options)[0][0].Text!.Length);

        SpreadsheetReadLimitException exception = Refused(() =>
            Read(XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", new string('a', 101)))), options));

        Assert.Equal(SpreadsheetReadLimit.MaxCellCharacters, exception.Limit);
        Assert.DoesNotContain("aaaa", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_value_element_longer_than_the_limit_is_refused_in_pieces_before_it_is_held()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, $"<c r=\"A1\" t=\"str\"><v>{new string('q', 200_000)}</v></c>"));

        SpreadsheetReadLimitException exception = Refused(() => Read(builder, new SpreadsheetReaderOptions { MaxCellCharacters = 1_000 }));

        Assert.Equal(SpreadsheetReadLimit.MaxCellCharacters, exception.Limit);
    }

    [Fact]
    public void A_tag_with_a_million_attributes_is_refused_before_the_xml_reader_holds_it()
    {
        // 200,000 attributes in one start tag: a single token with no end for megabytes. The ratio is allowed to
        // be anything so that it is the token that is refused and not the way it compresses.
        StringBuilder tag = new("<c r=\"A1\"");
        for (int index = 0; index < 200_000; index++)
        {
            tag.Append(" a").Append(index).Append("=\"1\"");
        }

        tag.Append("><v>1</v></c>");
        XlsxBuilder builder = XlsxBuilder.Workbook(string.Empty);
        builder.Replace("xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{XlsxBuilder.Main}\"><sheetData><row r=\"1\">{tag}</row></sheetData></worksheet>");

        SpreadsheetReadLimitException exception = Refused(() => Read(builder, new SpreadsheetReaderOptions { MaxCompressionRatio = 1_000_000 }));

        Assert.Equal(SpreadsheetReadLimit.MaxCellCharacters, exception.Limit);
        Assert.Contains("longer than any cell", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_comment_that_never_ends_is_refused_the_same_way()
    {
        StringBuilder comment = new("<!--");
        comment.Append(new string('c', 1_500_000)).Append("-->");
        XlsxBuilder builder = XlsxBuilder.Workbook(string.Empty);
        builder.Replace("xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{XlsxBuilder.Main}\">{comment}<sheetData/></worksheet>");

        SpreadsheetReadLimitException exception = Refused(() => Read(builder, new SpreadsheetReaderOptions { MaxCompressionRatio = 1_000_000 }));

        Assert.Equal(SpreadsheetReadLimit.MaxCellCharacters, exception.Limit);
    }

    [Fact]
    public void More_rows_than_the_limit_is_refused_and_exactly_the_limit_is_not()
    {
        SpreadsheetReaderOptions options = new() { MaxRows = 5 };

        Assert.Equal(5, Read(XlsxBuilder.Workbook(Rows(5)), options).Count);
        SpreadsheetReadLimitException exception = Refused(() => Read(XlsxBuilder.Workbook(Rows(6)), options));

        Assert.Equal(SpreadsheetReadLimit.MaxRows, exception.Limit);
        Assert.Contains("row 6", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Formatted_but_empty_rows_do_not_count_against_the_row_limit()
    {
        // A sheet that was formatted down to row 5,000 and holds data in two rows has two rows.
        string empties = string.Concat(Enumerable.Range(1, 5_000).Select(row => $"<row r=\"{row}\"><c r=\"A{row}\" s=\"1\"/></row>"));
        XlsxBuilder builder = XlsxBuilder.Workbook(empties + XlsxBuilder.Row(5_001, XlsxBuilder.Number("A5001", "1")) + XlsxBuilder.Row(5_002, XlsxBuilder.Number("A5002", "2")));

        Assert.Equal(2, Read(builder, new SpreadsheetReaderOptions { MaxRows = 2 }).Count);
    }

    [Fact]
    public void A_cell_past_the_column_limit_is_refused_and_one_at_it_is_not()
    {
        SpreadsheetReaderOptions options = new() { MaxColumns = 3 };

        Assert.Equal(3, Read(XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Number("C1", "1"))), options)[0].Count);
        SpreadsheetReadLimitException exception = Refused(() => Read(XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Number("D1", "1"))), options));

        Assert.Equal(SpreadsheetReadLimit.MaxColumns, exception.Limit);
    }

    [Fact]
    public void The_last_column_excel_has_is_read_and_one_past_it_is_refused_by_default()
    {
        List<SpreadsheetRow> rows = Read(XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Number("XFD1", "1"))), SpreadsheetReaderOptions.Default);

        Assert.Equal(16_384, rows[0].Count);
        Assert.Throws<SpreadsheetReadLimitException>(() => Read(XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Number("XFE1", "1"))), SpreadsheetReaderOptions.Default));
        Assert.Throws<SpreadsheetReadLimitException>(() => Read(XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Number("ZZZZZZZZZ1", "1"))), SpreadsheetReaderOptions.Default));
    }

    [Fact]
    public void The_options_refuse_a_limit_that_is_not_positive()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new SpreadsheetReaderOptions { MaxRows = 0 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new SpreadsheetReaderOptions { MaxUncompressedBytes = -1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new SpreadsheetReaderOptions { MaxCompressionRatio = 0 });
    }

    [Fact]
    public void The_default_limits_are_the_ones_the_documentation_states()
    {
        SpreadsheetReaderOptions defaults = SpreadsheetReaderOptions.Default;

        Assert.Equal(200L * 1024 * 1024, defaults.MaxUncompressedBytes);
        Assert.Equal(100, defaults.MaxCompressionRatio);
        Assert.Equal(1_000_000, defaults.MaxSharedStrings);
        Assert.Equal(50_000_000, defaults.MaxSharedStringCharacters);
        Assert.Equal(1_000, defaults.MaxParts);
        Assert.Equal(1_048_576, defaults.MaxRows);
        Assert.Equal(16_384, defaults.MaxColumns);
        Assert.Equal(32_767, defaults.MaxCellCharacters);
        Assert.False(defaults.IncludeEmptyRows);
        Assert.True(defaults.HonourSepDirective);
        Assert.True(defaults.UnescapeFormulaGuard);
    }

    [Fact]
    public void A_limit_exception_names_its_limit_and_value_and_never_a_cell()
    {
        SpreadsheetReadLimitException exception = new(SpreadsheetReadLimit.MaxRows, 50_000, "the sheet has more rows than that.");

        Assert.Equal("The file is over the MaxRows limit of 50000 (SpreadsheetReaderOptions.MaxRows): the sheet has more rows than that.", exception.Message);
        Assert.Equal(SpreadsheetReadLimit.MaxRows, exception.Limit);
        Assert.Equal(50_000, exception.Maximum);
    }

    /// <summary>
    /// Overwrites the declared uncompressed size of the named entry in the zip's central directory.
    /// </summary>
    private static void Patch(byte[] zip, string nameEndsWith, uint declaredLength)
    {
        // A central directory record is PK\x01\x02, and the uncompressed size sits 24 bytes in; the name follows
        // the 46-byte fixed part.
        for (int index = 0; index < zip.Length - 46; index++)
        {
            if (zip[index] != 0x50 || zip[index + 1] != 0x4B || zip[index + 2] != 0x01 || zip[index + 3] != 0x02)
            {
                continue;
            }

            int nameLength = zip[index + 28] | (zip[index + 29] << 8);
            string name = Encoding.UTF8.GetString(zip, index + 46, nameLength);
            if (name.EndsWith(nameEndsWith, StringComparison.Ordinal))
            {
                BitConverter.GetBytes(declaredLength).CopyTo(zip, index + 24);
                return;
            }
        }

        throw new InvalidOperationException("The entry was not found in the central directory.");
    }
}
