using System.Globalization;
using System.Text;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// The ways a file got past the first version of the readers, each as a test.
/// </summary>
/// <remarks>
/// <para>
/// A review of the readers found a file for each of these: a tag whose attributes carried the character the
/// first guard counted, a sheet whose every row held one cell at the far right, a workbook that named a million
/// sheets and a million relationships to look them up in, a formula that returned the empty string a thousand
/// times, a number too small for a decimal that came back as zero, and a date stored as a time that came back as
/// today. They are kept as tests because each looks fine from the outside: nothing fails, the file is read, and
/// the cost or the value is wrong.
/// </para>
/// </remarks>
[Collection(HeapCollection.Name)]
public sealed class SpreadsheetReaderHardeningTests
{
    private static List<SpreadsheetRow> Read(XlsxBuilder workbook, SpreadsheetReaderOptions? options = null)
    {
        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(workbook.ToStream(), options);
        return [.. reader.ReadRows(0)];
    }

    private static XlsxBuilder WithSheet(string sheetXml)
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(string.Empty);
        builder.Replace("xl/worksheets/sheet1.xml", sheetXml);
        return builder;
    }

    private static SpreadsheetCellValue One(string cell, string? styles = null, int? style = null)
    {
        List<SpreadsheetRow> rows = Read(XlsxBuilder.Workbook(XlsxBuilder.Row(1, cell), styles: styles));
        return rows.Single()[0];
    }

    // --- one token with no end, whatever character it is full of

    [Fact]
    public void A_start_tag_whose_attribute_values_hold_the_closing_bracket_is_still_one_token()
    {
        StringBuilder tag = new("<c r=\"A1\"");
        for (int index = 0; index < 400_000; index++)
        {
            tag.Append(" a").Append(index).Append("=\">\"");
        }

        tag.Append("><v>1</v></c>");
        XlsxBuilder builder = WithSheet($"<worksheet xmlns=\"{XlsxBuilder.Main}\"><sheetData><row r=\"1\">{tag}</row></sheetData></worksheet>");

        SpreadsheetReadLimitException exception = Assert.Throws<SpreadsheetReadLimitException>(
            () => Read(builder, new SpreadsheetReaderOptions { MaxCompressionRatio = 1_000_000 }));

        Assert.Equal(SpreadsheetReadLimit.MaxCellCharacters, exception.Limit);
    }

    [Fact]
    public void A_cdata_section_that_never_ends_is_one_token_even_when_it_is_full_of_markup()
    {
        string cdata = "<![CDATA[" + string.Concat(Enumerable.Repeat("<a></a>", 300_000)) + "]]>";
        XlsxBuilder builder = WithSheet($"<worksheet xmlns=\"{XlsxBuilder.Main}\"><sheetData><row r=\"1\"><c r=\"A1\"><junk>{cdata}</junk></c></row></sheetData></worksheet>");

        SpreadsheetReadLimitException exception = Assert.Throws<SpreadsheetReadLimitException>(
            () => Read(builder, new SpreadsheetReaderOptions { MaxCompressionRatio = 1_000_000 }));

        Assert.Equal(SpreadsheetReadLimit.MaxCellCharacters, exception.Limit);
    }

    [Fact]
    public void A_long_run_of_text_is_one_token_however_often_it_is_broken_by_entities()
    {
        string text = string.Concat(Enumerable.Repeat("&amp;&gt;&lt;", 200_000));
        XlsxBuilder builder = WithSheet($"<worksheet xmlns=\"{XlsxBuilder.Main}\"><sheetData><row r=\"1\"><c r=\"A1\"><junk>{text}</junk></c></row></sheetData></worksheet>");

        Assert.Throws<SpreadsheetReadLimitException>(() => Read(builder, new SpreadsheetReaderOptions { MaxCompressionRatio = 1_000_000 }));
    }

    [Fact]
    public void An_ordinary_sheet_with_a_lot_of_markup_is_not_a_long_token()
    {
        // Many short tokens, a lot of them: a megabyte of ordinary cells.
        string rows = string.Concat(Enumerable.Range(1, 5_000).Select(row => XlsxBuilder.Row(row, XlsxBuilder.Number($"A{row}", "1"), XlsxBuilder.Inline($"B{row}", "text"))));

        Assert.Equal(5_000, Read(XlsxBuilder.Workbook(rows)).Count);
    }

    [Fact]
    public void A_part_written_in_utf16_is_refused_because_its_markup_cannot_be_followed_byte_by_byte()
    {
        XlsxBuilder clean = new();
        clean.Add("_rels/.rels", $"<Relationships xmlns=\"{XlsxBuilder.PackageRel}\"><Relationship Id=\"rId1\" Type=\"{XlsxBuilder.Rel}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
        clean.Add("xl/workbook.xml", $"<workbook xmlns=\"{XlsxBuilder.Main}\" xmlns:r=\"{XlsxBuilder.Rel}\"><sheets><sheet name=\"S\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
        clean.Add("xl/_rels/workbook.xml.rels", $"<Relationships xmlns=\"{XlsxBuilder.PackageRel}\"><Relationship Id=\"rId1\" Type=\"{XlsxBuilder.Rel}/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
        clean.Add("xl/worksheets/sheet1.xml", [.. Encoding.Unicode.GetPreamble(), .. Encoding.Unicode.GetBytes(XlsxBuilder.Sheet(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x"))))]);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => Read(clean));

        Assert.Contains("UTF-16", exception.Message, StringComparison.Ordinal);
    }

    // --- a sheet of rows that each hold one cell at the far right

    [Fact]
    public void A_row_with_one_cell_at_the_far_right_costs_one_cell_and_not_sixteen_thousand()
    {
        string rows = string.Concat(Enumerable.Range(1, 4_000).Select(row => XlsxBuilder.Row(row, $"<c r=\"XFD{row}\" t=\"b\"><v>1</v></c>")));
        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(XlsxBuilder.Workbook(rows).ToStream());

        long before = GC.GetAllocatedBytesForCurrentThread();
        int seen = 0;
        foreach (SpreadsheetRow row in reader.ReadRows(0))
        {
            Assert.Equal(16_384, row.Count);
            Assert.True(row[16_383].Boolean);
            Assert.True(row[0].IsEmpty);
            seen++;
        }

        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.Equal(4_000, seen);
        // Dense rows would allocate about 640 KiB each, 2.5 GB for these.
        Assert.True(allocated < 200L * 1024 * 1024, $"allocated {allocated} bytes");
    }

    [Fact]
    public void A_csv_record_of_empty_fields_with_one_value_is_sparse_too()
    {
        using MemoryStream stream = new(new UTF8Encoding(false).GetBytes(new string(',', 9_000) + "x\n"));
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);

        SpreadsheetRow row = reader.ReadRows().Single();

        Assert.Equal(9_001, row.Count);
        Assert.Equal("x", row[9_000].Text);
        Assert.Equal(9_001, row.Cells.Count);
    }

    [Fact]
    public void A_row_made_by_hand_ends_at_its_last_cell_with_a_value()
    {
        SpreadsheetRow row = new(7, [SpreadsheetCellValue.FromText("a"), SpreadsheetCellValue.Empty, SpreadsheetCellValue.FromNumber(2m), SpreadsheetCellValue.Empty]);

        Assert.Equal(7, row.RowNumber);
        Assert.Equal(3, row.Count);
        Assert.Equal("a", row[0].Text);
        Assert.True(row[1].IsEmpty);
        Assert.Equal(2m, row[2].Number);
        Assert.True(row[3].IsEmpty);
        Assert.False(row.IsEmpty);
        Assert.Equal(3, row.Cells.Count);
        Assert.Throws<ArgumentOutOfRangeException>(() => new SpreadsheetRow(0, []));
        Assert.Throws<ArgumentOutOfRangeException>(() => row[-1]);
    }

    // --- a workbook that names too much

    [Fact]
    public void A_workbook_that_lists_more_sheets_than_the_limit_is_refused_by_name()
    {
        const int count = 30;
        string sheets = string.Concat(Enumerable.Range(1, count).Select(index => $"<sheet name=\"S{index}\" sheetId=\"{index}\" r:id=\"rId{index}\"/>"));
        string rels = string.Concat(Enumerable.Range(1, count).Select(index => $"<Relationship Id=\"rId{index}\" Type=\"{XlsxBuilder.Rel}/worksheet\" Target=\"worksheets/sheet{index}.xml\"/>"));
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x")));
        builder.Replace("xl/workbook.xml", $"<workbook xmlns=\"{XlsxBuilder.Main}\" xmlns:r=\"{XlsxBuilder.Rel}\"><sheets>{sheets}</sheets></workbook>");
        builder.Replace("xl/_rels/workbook.xml.rels", $"<Relationships xmlns=\"{XlsxBuilder.PackageRel}\">{rels}</Relationships>");
        for (int index = 2; index <= count; index++)
        {
            builder.Add($"xl/worksheets/sheet{index}.xml", XlsxBuilder.Sheet(string.Empty));
        }

        SpreadsheetReadLimitException exception = Assert.Throws<SpreadsheetReadLimitException>(
            () => SpreadsheetWorkbookReader.Open(builder.ToStream(), new SpreadsheetReaderOptions { MaxParts = 10 }));

        Assert.Equal(SpreadsheetReadLimit.MaxParts, exception.Limit);
    }

    [Fact]
    public void Two_sheets_that_are_one_part_are_refused_because_reading_both_would_read_it_twice()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x")));
        builder.Replace(
            "xl/workbook.xml",
            $"<workbook xmlns=\"{XlsxBuilder.Main}\" xmlns:r=\"{XlsxBuilder.Rel}\"><sheets><sheet name=\"A\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"B\" sheetId=\"2\" r:id=\"rId1\"/></sheets></workbook>");

        Assert.Throws<InvalidDataException>(() => SpreadsheetWorkbookReader.Open(builder.ToStream()));
    }

    [Fact]
    public void A_relationships_part_with_more_relationships_than_the_limit_is_refused_by_name()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x")));
        string many = string.Concat(Enumerable.Range(100, 50).Select(index => $"<Relationship Id=\"rId{index}\" Type=\"{XlsxBuilder.Rel}/hyperlink\" Target=\"http://example.invalid/{index}\" TargetMode=\"External\"/>"));
        builder.Replace(
            "xl/_rels/workbook.xml.rels",
            $"<Relationships xmlns=\"{XlsxBuilder.PackageRel}\"><Relationship Id=\"rId1\" Type=\"{XlsxBuilder.Rel}/worksheet\" Target=\"worksheets/sheet1.xml\"/>{many}</Relationships>");

        SpreadsheetReadLimitException exception = Assert.Throws<SpreadsheetReadLimitException>(
            () => SpreadsheetWorkbookReader.Open(builder.ToStream(), new SpreadsheetReaderOptions { MaxParts = 10 }));

        Assert.Equal(SpreadsheetReadLimit.MaxParts, exception.Limit);
    }

    [Fact]
    public void A_part_that_is_not_xml_is_not_held_to_the_compression_ratio()
    {
        // A flat-colour picture compresses a thousand to one, and nothing here ever reads it.
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "fine")));
        builder.Add("xl/media/image1.bmp", new byte[4 * 1024 * 1024]);

        Assert.Equal("fine", Read(builder, new SpreadsheetReaderOptions { MaxCompressionRatio = 100 })[0][0].Text);
    }

    [Fact]
    public void The_number_of_columns_allowed_has_a_ceiling()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new SpreadsheetReaderOptions { MaxColumns = 1_000_001 });
        Assert.Equal(1_000_000, new SpreadsheetReaderOptions { MaxColumns = 1_000_000 }.MaxColumns);
    }

    // --- values a real file has

    [Fact]
    public void A_formula_that_returns_the_empty_string_is_an_empty_cell_as_it_is_in_a_csv()
    {
        // =IF(A2="","",...) filled down: Excel stores each as a string result with no characters.
        string rows = string.Concat(Enumerable.Range(1, 500).Select(row => XlsxBuilder.Row(row, $"<c r=\"A{row}\" t=\"str\"><f>IF(1,\"\",\"\")</f><v></v></c>")))
            + XlsxBuilder.Row(501, XlsxBuilder.Number("A501", "1"));

        List<SpreadsheetRow> read = Read(XlsxBuilder.Workbook(rows), new SpreadsheetReaderOptions { MaxRows = 1 });

        Assert.Equal([501], read.Select(row => row.RowNumber));
    }

    [Fact]
    public void An_empty_inline_string_and_an_empty_shared_string_are_empty_cells()
    {
        Assert.Empty(Read(XlsxBuilder.Workbook(XlsxBuilder.Row(1, "<c r=\"A1\" t=\"inlineStr\"><is><t></t></is></c>"))));

        List<SpreadsheetRow> rows = Read(XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Shared("A1", 0)), sharedStrings: "<si><t></t></si>"));
        Assert.Empty(rows);
    }

    [Fact]
    public void A_date_keeps_the_serial_number_the_file_stores_for_the_durations_a_date_cannot_say()
    {
        // [h]:mm holding 36 hours: the serial is 1.5 and the "date" of it means nothing.
        string styles = XlsxBuilder.Styles("<numFmt numFmtId=\"164\" formatCode=\"[h]:mm\"/>", 0, 164);

        SpreadsheetCellValue value = One(XlsxBuilder.Number("A1", "1.5", style: 1), styles);

        Assert.Equal(SpreadsheetCellKind.Date, value.Kind);
        Assert.Equal(1.5m, value.Number);
    }

    [Fact]
    public void An_iso_date_cell_is_read_only_as_an_iso_date()
    {
        Assert.Equal(new DateTime(2026, 11, 20), One("<c r=\"A1\" t=\"d\"><v>2026-11-20</v></c>").Date);
        Assert.Equal(new DateTime(2026, 11, 20, 13, 30, 5), One("<c r=\"A1\" t=\"d\"><v>2026-11-20T13:30:05</v></c>").Date);

        // A time alone has no day, and "today" is a different answer every morning; a US date is not ISO.
        Assert.Equal("10:00:00", One("<c r=\"A1\" t=\"d\"><v>10:00:00</v></c>").Text);
        Assert.Equal("1/2/2020", One("<c r=\"A1\" t=\"d\"><v>1/2/2020</v></c>").Text);
    }

    [Theory]
    [InlineData("1E-30")]
    [InlineData("1.5E-29")]
    [InlineData("-3E-40")]
    [InlineData("0.00000000000000000000000000001")]
    public void A_number_too_small_for_a_decimal_comes_back_as_text_and_not_as_zero(string raw)
    {
        SpreadsheetCellValue value = One(XlsxBuilder.Number("A1", raw));

        Assert.Equal(SpreadsheetCellKind.Text, value.Kind);
        Assert.Equal(raw, value.Text);
    }

    [Theory]
    [InlineData("0")]
    [InlineData("0.0")]
    [InlineData("-0")]
    [InlineData("0E5")]
    [InlineData("0.000E+00")]
    public void A_zero_is_still_a_number(string raw)
    {
        SpreadsheetCellValue value = One(XlsxBuilder.Number("A1", raw));

        Assert.Equal(SpreadsheetCellKind.Number, value.Kind);
        Assert.Equal(0m, value.Number);
    }

    // --- the CSV reader

    [Fact]
    public void An_encoding_passed_in_the_options_is_read_strictly_whatever_its_own_fallback_is()
    {
        // Encoding.UTF8 replaces bytes it cannot decode with U+FFFD. The reader must not let it.
        byte[] bytes = [.. "a;b\nx;"u8.ToArray(), 0xFD, (byte)'\n'];
        using MemoryStream stream = new(bytes);
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream, new SpreadsheetReaderOptions { Encoding = Encoding.UTF8 });

        Assert.Throws<InvalidDataException>(() => reader.ReadRows().ToList());
    }

    [Fact]
    public void A_guarded_cell_exactly_at_the_limit_reads_back_and_one_character_more_is_refused()
    {
        SpreadsheetReaderOptions options = new() { MaxCellCharacters = 20 };
        SpreadsheetGridColumn<string>[] columns = [new("Text", value => value)];

        string atLimit = "=" + new string('a', 19);
        byte[] written = SpreadsheetGridCsv.Write(columns, [atLimit]);
        using (MemoryStream stream = new(written))
        using (SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream, options))
        {
            Assert.Equal(atLimit, reader.ReadRows().Skip(1).Single()[0].Text);
        }

        byte[] tooLong = SpreadsheetGridCsv.Write(columns, [atLimit + "a"]);
        using MemoryStream over = new(tooLong);
        using SpreadsheetCsvReader overReader = SpreadsheetCsvReader.Open(over, options);

        Assert.Equal(SpreadsheetReadLimit.MaxCellCharacters, Assert.Throws<SpreadsheetReadLimitException>(() => overReader.ReadRows().ToList()).Limit);
    }

    [Fact]
    public void An_apostrophe_that_is_data_does_not_get_the_extra_character_the_guard_has()
    {
        SpreadsheetReaderOptions options = new() { MaxCellCharacters = 5 };
        using MemoryStream stream = new("'abcde\n"u8.ToArray());
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream, options);

        Assert.Throws<SpreadsheetReadLimitException>(() => reader.ReadRows().ToList());
    }

    [Fact]
    public void A_row_ends_at_its_last_cell_with_a_value_in_both_readers()
    {
        using MemoryStream csv = new("a,b,,\n"u8.ToArray());
        using SpreadsheetCsvReader csvReader = SpreadsheetCsvReader.Open(csv);

        List<SpreadsheetRow> workbook = Read(XlsxBuilder.Workbook("<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\" s=\"2\"/><c r=\"D1\" s=\"2\"/></row>"));

        Assert.Equal(2, csvReader.ReadRows().Single().Count);
        Assert.Equal(2, workbook.Single().Count);
        Assert.Equal(
            2m.ToString(CultureInfo.InvariantCulture),
            workbook.Single()[1].ToString());
    }
}
